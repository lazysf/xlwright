// Throwaway experiment for the stage 2 open questions in docs/architecture.md "残った専用のExcelの片付け".
// leave starts a dedicated Excel, confirms ownership, records its PID and creation time, and kills its own
// process without releasing COM, as if the server were killed. watch, probe, list, show, and kill act only on the
// recorded Excel, after matching its PID and creation time; they never touch another Excel.
//
// Usage: dotnet run --project spikes\leftover -- leave <state.json> [<workbook> [hidden|protected|busy]]
//        dotnet run --project spikes\leftover -- watch <state.json> <seconds>
//        dotnet run --project spikes\leftover -- probe <state.json> [<seconds>]
//        dotnet run --project spikes\leftover -- list <state.json>
//        dotnet run --project spikes\leftover -- show <state.json>
//        dotnet run --project spikes\leftover -- kill <state.json>

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            return args switch
            {
                ["leave", var state] => Leave(state, null, "open"),
                ["leave", var state, var workbook] => Leave(state, Path.GetFullPath(workbook), "open"),
                ["leave", var state, var workbook, var form] when form is "hidden" or "protected" or "busy" =>
                    Leave(state, Path.GetFullPath(workbook), form),
                ["watch", var state, var seconds] => Watch(Read(state), int.Parse(seconds)),
                ["probe", var state] => Probe(Read(state), 90),
                ["probe", var state, var seconds] => Probe(Read(state), int.Parse(seconds)),
                ["list", var state] => Probe(Read(state), 90, closeAndQuit: false),
                ["show", var state] => Show(Read(state)),
                ["kill", var state] => Kill(Read(state)),
                _ => Usage(),
            };
        }
        catch (NotOwnedException e)
        {
            Log($"STOP: {e.Message}; the Excel is left untouched");
            return 2;
        }
    }

    private static int Usage()
    {
        Console.Error.WriteLine("Usage: Leftover leave <state.json> [<workbook>] | watch <state.json> <seconds> | probe <state.json>");
        return 64;
    }

    private static int Leave(string state, string? workbook, string form)
    {
        var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Log($"before: {before.Count} EXCEL process(es)");
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var requested = DateTime.UtcNow;
                dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
                _ = Native.GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out var pid);
                using var process = Process.GetProcessById((int)pid);
                var created = process.StartTime.ToUniversalTime();
                if (before.Contains((int)pid) || process.ProcessName != "EXCEL" || created <= requested)
                {
                    throw new NotOwnedException($"PID {pid} is not a new EXCEL started after the request");
                }

                if (workbook is not null && form == "protected")
                {
                    app.ProtectedViewWindows.Open(workbook);
                }
                else if (workbook is not null)
                {
                    dynamic book = app.Workbooks.Open(workbook);
                    Log($"opened {Path.GetFileName(workbook)}: ReadOnly {(bool)book.ReadOnly}");
                    if (form == "hidden")
                    {
                        book.Windows[1].Visible = false;
                    }
                }

                File.WriteAllText(state, JsonSerializer.Serialize(new Recorded((int)pid, created.Ticks)));
                Log($"owned PID {pid} created {created:O}; workbook {(workbook is null ? "none" : Path.GetFileName(workbook))} ({form}); Workbooks {(int)app.Workbooks.Count}, ProtectedViewWindows {(int)app.ProtectedViewWindows.Count}; killing this process without releasing COM");
                if (form == "busy")
                {
                    // The macro never returns, so another thread kills this process while the call is pending.
                    new Thread(() => { Thread.Sleep(3_000); Process.GetCurrentProcess().Kill(); }) { IsBackground = true }.Start();
                    app.Run($"'{Path.GetFileName(workbook)}'!Calc.Forever");
                }

                Process.GetCurrentProcess().Kill();
            }
            catch (Exception e)
            {
                failure = e;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (failure is NotOwnedException)
        {
            throw failure;
        }

        Log($"FAILED {failure?.GetType().Name}: {failure?.Message}");
        return 1;
    }

    private static int Watch(Recorded recorded, int seconds)
    {
        var watch = Stopwatch.StartNew();
        while (watch.Elapsed.TotalSeconds < seconds)
        {
            if (!Matches(recorded))
            {
                Log($"PID {recorded.Pid} ended within {watch.Elapsed.TotalSeconds:F0} s of watching");
                return 0;
            }

            Thread.Sleep(5_000);
        }

        Log($"PID {recorded.Pid} still runs after {seconds} s");
        return 1;
    }

    private static int Probe(Recorded recorded, int seconds, bool closeAndQuit = true)
    {
        if (!Matches(recorded))
        {
            Log($"PID {recorded.Pid} with the recorded creation time is not running");
            return 1;
        }

        var (mains, books) = Windows(recorded.Pid);
        Log($"PID {recorded.Pid}: {mains.Count} XLMAIN window(s) (visible {string.Join(",", mains.Select(Native.IsWindowVisible))}), {books.Count} EXCEL7 window(s)");
        foreach (var (kind, handle) in mains.Select(h => ("XLMAIN", h)).Concat(books.Select(h => ("EXCEL7", h))))
        {
            var connected = false;
            if (!OnSta(() => connected = TryConnect(kind, handle, recorded, closeAndQuit), TimeSpan.FromSeconds(seconds)))
            {
                Log($"{kind}: the call did not return within {seconds} s; the thread is left behind");
            }

            if (connected)
            {
                return 0;
            }
        }

        Log("no window gave an Excel object; the Excel is left running");
        return 1;
    }

    // Shows the recorded Excel's main window that holds a workbook window, without COM, as the handoff script would.
    private static int Show(Recorded recorded)
    {
        if (!Matches(recorded))
        {
            Log($"PID {recorded.Pid} with the recorded creation time is not running");
            return 1;
        }

        var (mains, books) = Windows(recorded.Pid);
        var targets = mains.Where(main => books.Any(book => Native.GetAncestor(book, 2) == main)).ToList();
        if (targets.Count == 0)
        {
            Log("no XLMAIN holds an EXCEL7 window; nothing is shown");
            return 1;
        }

        foreach (var main in targets)
        {
            var wasVisible = Native.ShowWindow(main, 5); // SW_SHOW
            var foreground = Native.SetForegroundWindow(main);
            Log($"XLMAIN 0x{main.ToInt64():X}: ShowWindow (was visible {wasVisible}), SetForegroundWindow {foreground}, now visible {Native.IsWindowVisible(main)}");
        }

        _ = Windows(recorded.Pid);
        return 0;
    }

    private static int Kill(Recorded recorded)
    {
        if (!Matches(recorded))
        {
            Log($"PID {recorded.Pid} with the recorded creation time is not running");
            return 1;
        }

        using var process = Process.GetProcessById(recorded.Pid);
        process.Kill();
        var ended = process.WaitForExit(5_000);
        Log($"killed PID {recorded.Pid}: {(ended ? "ended" : "did not end within 5 s")}");
        return ended ? 0 : 1;
    }

    private static bool TryConnect(string kind, IntPtr handle, Recorded recorded, bool closeAndQuit)
    {
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        var watch = Stopwatch.StartNew();
        var hr = Native.AccessibleObjectFromWindow(handle, 0xFFFFFFF0, ref iid, out var accessible);
        Log($"{kind}: AccessibleObjectFromWindow 0x{hr:X8} in {watch.ElapsedMilliseconds} ms");
        if (hr != 0 || accessible is null)
        {
            return false;
        }

        dynamic app = ((dynamic)accessible).Application;
        _ = Native.GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out var pid);
        if (pid != recorded.Pid)
        {
            throw new NotOwnedException($"the window's Application.Hwnd belongs to PID {pid}");
        }

        int protectedViews = app.ProtectedViewWindows.Count;
        var names = new List<string>();
        foreach (dynamic book in app.Workbooks)
        {
            names.Add((string)book.Name);
        }

        Log($"{kind}: connected; Application.Hwnd PID {pid}; Workbooks [{string.Join(", ", names)}]; ProtectedViewWindows {protectedViews}; Visible {(bool)app.Visible}");
        if (!closeAndQuit)
        {
            Marshal.FinalReleaseComObject(app);
            Marshal.FinalReleaseComObject(accessible);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        foreach (dynamic book in app.Workbooks)
        {
            book.Close(false);
        }

        app.Quit();
        Marshal.FinalReleaseComObject(app);
        Marshal.FinalReleaseComObject(accessible);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        var ended = Watch(recorded, 75) == 0;
        Log($"after Close and Quit: {(ended ? "ended" : "still runs")}");
        return true;
    }

    private static (List<IntPtr> Mains, List<IntPtr> Books) Windows(int pid)
    {
        var mains = new List<IntPtr>();
        var books = new List<IntPtr>();
        var classes = new List<string>();
        Native.EnumWindows((window, parameter) =>
        {
            _ = Native.GetWindowThreadProcessId(window, out var owner);
            if (owner != pid)
            {
                return true;
            }

            var top = Native.ClassName(window);
            var children = new List<string>();
            Native.EnumChildWindows(window, (child, parameter) =>
            {
                var name = Native.ClassName(child);
                children.Add(name);
                if (top == "XLMAIN" && name == "EXCEL7")
                {
                    books.Add(child);
                }

                return true;
            }, IntPtr.Zero);
            if (top == "XLMAIN")
            {
                mains.Add(window);
            }

            classes.Add($"{top}{(Native.IsWindowVisible(window) ? "(visible)" : "")}[{string.Join(",", children.Distinct())}]");
            return true;
        }, IntPtr.Zero);
        Log($"PID {pid} top-level windows: {string.Join(" ", classes)}");
        return (mains, books);
    }

    private static bool OnSta(Action action, TimeSpan limit)
    {
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (NotOwnedException e) { Log($"STOP: {e.Message}; the Excel is left untouched"); }
            catch (Exception e) { Log($"FAILED {e.GetType().Name} 0x{e.HResult:X8}: {e.Message}"); }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return thread.Join(limit);
    }

    private static bool Matches(Recorded recorded)
    {
        try
        {
            using var process = Process.GetProcessById(recorded.Pid);
            return !process.HasExited && process.ProcessName == "EXCEL" && process.StartTime.ToUniversalTime().Ticks == recorded.CreatedTicks;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static Recorded Read(string state) => JsonSerializer.Deserialize<Recorded>(File.ReadAllText(state))!;

    private static readonly long s_start = Stopwatch.GetTimestamp();

    private static void Log(string message) =>
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss} +{Stopwatch.GetElapsedTime(s_start).TotalSeconds,7:F1}] {message}");
}

internal sealed record Recorded(int Pid, long CreatedTicks);

internal sealed class NotOwnedException(string message) : Exception(message);

internal static class Native
{
    public delegate bool WindowCallback(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr parameter);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr window, uint flags); // GA_ROOT is 2

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr window, int command);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int length);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object? accessible);

    public static string ClassName(IntPtr window)
    {
        var name = new StringBuilder(256);
        return GetClassName(window, name, name.Capacity) > 0 ? name.ToString() : "";
    }
}
