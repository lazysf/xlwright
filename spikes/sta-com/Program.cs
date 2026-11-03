// Throwaway experiment for the stage 1 open questions in docs/architecture.md that need no workbook.
// It starts dedicated Excel instances from .NET STA threads, confirms ownership, and ends them.
// It only ends an Excel whose ownership it confirmed: a new PID, named EXCEL.EXE, created after the start request.
//
// Usage: dotnet run --project spikes\sta-com -- <start-stop|retry-later|release|pending-kill|all>

using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

internal static class Program
{
    private static Dictionary<int, DateTime> s_before = new();

    private static int Main(string[] args)
    {
        var trial = args.Length == 1 ? args[0] : "";
        var trials = trial switch
        {
            "start-stop" => new[] { "start-stop" },
            "retry-later" => new[] { "retry-later" },
            "release" => new[] { "release" },
            "pending-kill" => new[] { "pending-kill" },
            "all" => new[] { "start-stop", "retry-later", "release", "pending-kill" },
            _ => Array.Empty<string>(),
        };
        if (trials.Length == 0)
        {
            Console.Error.WriteLine("Usage: StaCom <start-stop|retry-later|release|pending-kill|all>");
            return 64;
        }

        s_before = Snapshot();
        Log($"before: {s_before.Count} EXCEL process(es) {Describe(s_before)}");
        Log($"runtime: {RuntimeInformation.FrameworkDescription}, {RuntimeInformation.ProcessArchitecture}");

        var failed = false;
        foreach (var name in trials)
        {
            try
            {
                switch (name)
                {
                    case "start-stop":
                        for (var i = 1; i <= 3; i++) StartStop(i);
                        break;
                    case "retry-later":
                        RetryLater();
                        break;
                    case "release":
                        Release(setUserControl: false);
                        Release(setUserControl: true);
                        UserControlQuit();
                        break;
                    case "pending-kill":
                        PendingKill();
                        break;
                }
            }
            catch (NotOwnedException e)
            {
                // Ownership was not confirmed: the Excel is left untouched.
                Log($"STOP {name}: ownership not confirmed, Excel left untouched: {e.Message}");
                failed = true;
                break;
            }
        }

        var after = Snapshot();
        var unchanged = s_before.All(p => after.TryGetValue(p.Key, out var t) && t == p.Value);
        var leftover = after.Keys.Where(pid => !s_before.ContainsKey(pid)).ToList();
        Log($"after: {after.Count} EXCEL process(es) {Describe(after)}");
        Log($"check: pre-existing unchanged={unchanged}, started by this run and still running={leftover.Count}");
        return failed || !unchanged || leftover.Count > 0 ? 1 : 0;
    }

    // Trial 1: start, confirm ownership, read the command line, Quit, and time the exit.
    private static void StartStop(int round)
    {
        RunSta($"start-stop#{round}", () =>
        {
            var filter = MessageFilter.Register($"start-stop#{round}");
            var (app, owned) = StartAndOwn($"start-stop#{round}");
            QuitAndWait($"start-stop#{round}", app, owned);
            filter.Report();
            MessageFilter.Revoke();
        });
    }

    // Trial 2: while thread A is inside Application.Wait, thread B calls Application.Version,
    // first with a message filter that retries RETRYLATER for up to 10 s, then without a filter.
    private static void RetryLater()
    {
        RunSta("retry-later/A", () =>
        {
            var filter = MessageFilter.Register("retry-later/A");
            var (app, owned) = StartAndOwn("retry-later");
            try
            {
            foreach (var withFilter in new[] { true, false })
            {
                var label = withFilter ? "B(filter)" : "B(no filter)";
                var stream = MarshalToStream(app);
                var callerB = new Thread(() => CallVersionFromOtherSta(label, stream, withFilter, delayMs: 1000));
                callerB.SetApartmentState(ApartmentState.STA);
                callerB.IsBackground = true;
                callerB.Start();

                var watch = Stopwatch.StartNew();
                app.Wait(DateTime.Now.AddSeconds(8));
                Log($"retry-later A: Application.Wait(8 s) returned after {watch.ElapsedMilliseconds} ms");
                if (!callerB.Join(20_000)) Log($"retry-later {label}: did not return within 20 s");
            }

            // Same, while A's call makes Excel recalculate. The new workbook has no file and is closed without saving.
            dynamic workbooks = app.Workbooks;
            dynamic workbook = workbooks.Add();
            dynamic sheets = workbook.Worksheets;
            dynamic sheet = sheets[1];
            dynamic cells = sheet.Range["A1:A100"];
            var round = 0;
            foreach (var withFilter in new[] { true, false })
            {
                var label = withFilter ? "B(filter, recalc)" : "B(no filter, recalc)";
                var stream = MarshalToStream(app);
                var callerB = new Thread(() => CallVersionFromOtherSta(label, stream, withFilter, delayMs: 1000));
                callerB.SetApartmentState(ApartmentState.STA);
                callerB.IsBackground = true;
                callerB.Start();

                var watch = Stopwatch.StartNew();
                cells.Formula = $"=SUMPRODUCT(SQRT(ROW($1:$1048576)+{++round}))";
                Log($"retry-later A: setting 100 heavy formulas returned after {watch.ElapsedMilliseconds} ms");
                if (!callerB.Join(30_000)) Log($"retry-later {label}: did not return within 30 s");
            }

            Marshal.FinalReleaseComObject(cells);
            Marshal.FinalReleaseComObject(sheet);
            Marshal.FinalReleaseComObject(sheets);
            workbook.Close(false);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(workbooks);
            }
            finally
            {
                QuitAndWait("retry-later", app, owned);
            }

            filter.Report();
            MessageFilter.Revoke();
        });
    }

    private static void CallVersionFromOtherSta(string label, IntPtr stream, bool withFilter, int delayMs)
    {
        var filter = withFilter ? MessageFilter.Register($"retry-later {label}") : null;
        dynamic app = UnmarshalFromStream(stream);
        Thread.Sleep(delayMs);
        var watch = Stopwatch.StartNew();
        try
        {
            string version = app.Version;
            Log($"retry-later {label}: Application.Version={version} after {watch.ElapsedMilliseconds} ms");
        }
        catch (COMException e)
        {
            Log($"retry-later {label}: COMException 0x{e.HResult:X8} after {watch.ElapsedMilliseconds} ms");
        }
        catch (Exception e)
        {
            // The dynamic binder wraps a failed GetIDsOfNames in MissingMemberException; the HRESULT is in the message.
            Log($"retry-later {label}: {e.GetType().Name} after {watch.ElapsedMilliseconds} ms: {e.Message}");
        }
        finally
        {
            Marshal.FinalReleaseComObject(app);
            filter?.Report();
            if (filter != null) MessageFilter.Revoke();
        }
    }

    // Trial 3: release every COM reference without Quit and see whether the hidden Excel ends by itself.
    private static void Release(bool setUserControl)
    {
        var label = setUserControl ? "release(UserControl=True)" : "release";
        Owned? owned = null;
        RunSta(label, () =>
        {
            var (app, o) = StartAndOwn(label);
            owned = o;
            bool userControl = app.UserControl;
            Log($"{label}: UserControl before={userControl}");
            if (setUserControl)
            {
                app.UserControl = true;
                Log($"{label}: UserControl set to {(bool)app.UserControl}");
            }

            RequireNoWorkbooks(label, app);
            Marshal.FinalReleaseComObject(app);
        });
        GC.Collect();
        GC.WaitForPendingFinalizers();

        var watch = Stopwatch.StartNew();
        var ended = owned!.WaitForExit(30_000);
        Log($"{label}: ended by itself={ended} ({(ended ? watch.ElapsedMilliseconds + " ms after release" : "not within 30 s")})");
        if (!ended) owned.Terminate(label);
        owned.Dispose();
    }

    // Trial 3b: with UserControl=True, does Quit still end the Excel?
    private static void UserControlQuit()
    {
        RunSta("quit(UserControl=True)", () =>
        {
            var (app, owned) = StartAndOwn("quit(UserControl=True)");
            app.UserControl = true;
            QuitAndWait("quit(UserControl=True)", app, owned);
        });
    }

    // Trial 4: terminate the owned Excel while thread A has a pending call (Application.Wait 60 s).
    private static void PendingKill()
    {
        Owned? owned = null;
        using var ready = new ManualResetEventSlim();
        var returnedAt = new long[1];
        var callResult = new string[1];
        var caller = new Thread(() =>
        {
            var filter = MessageFilter.Register("pending-kill/A");
            dynamic app;
            try
            {
                (app, var o) = StartAndOwn("pending-kill");
                RequireNoWorkbooks("pending-kill", app);
                owned = o;
            }
            catch (NotOwnedException e)
            {
                Log($"STOP pending-kill: ownership not confirmed, Excel left untouched: {e.Message}");
                return;
            }
            finally
            {
                ready.Set();
            }

            try
            {
                app.Wait(DateTime.Now.AddSeconds(60));
                callResult[0] = "returned normally";
            }
            catch (COMException e)
            {
                callResult[0] = $"COMException 0x{e.HResult:X8}";
            }
            catch (Exception e)
            {
                callResult[0] = $"{e.GetType().Name} 0x{e.HResult:X8}";
            }

            returnedAt[0] = Stopwatch.GetTimestamp();
            try { Marshal.FinalReleaseComObject(app); } catch (Exception e) { Log($"pending-kill: release threw {e.GetType().Name}"); }
            filter.Report();
            MessageFilter.Revoke();
        });
        caller.SetApartmentState(ApartmentState.STA);
        caller.IsBackground = true;
        caller.Start();
        if (!ready.Wait(150_000))
        {
            Log("pending-kill: start did not finish within 150 s");
            return;
        }

        if (owned == null) throw new NotOwnedException("see the STOP line above");
        Thread.Sleep(3000);
        var killedAt = Stopwatch.GetTimestamp();
        owned.Terminate("pending-kill");
        var returned = caller.Join(15_000);
        Log(returned
            ? $"pending-kill: pending call {callResult[0]} {Stopwatch.GetElapsedTime(killedAt, returnedAt[0]).TotalMilliseconds:F0} ms after TerminateProcess"
            : "pending-kill: pending call did not return within 15 s after TerminateProcess");
        owned.Dispose();
    }

    private static (dynamic App, Owned Owned) StartAndOwn(string label)
    {
        var requestedUtc = DateTime.UtcNow;
        var watch = Stopwatch.StartNew();
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
        var createMs = watch.ElapsedMilliseconds;
        int hwnd = app.Hwnd;
        _ = Native.GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);

        var handle = Native.OpenProcess(Native.ProcessAccess, false, (int)pid);
        if (handle.IsInvalid) throw new NotOwnedException($"OpenProcess({pid}) failed: {new Win32Exception().Message}");
        var image = Native.ImagePath(handle);
        var created = Native.CreationTimeUtc(handle);
        var commandLine = Native.CommandLine(handle);
        var isNew = !s_before.ContainsKey((int)pid);
        var nameOk = string.Equals(Path.GetFileName(image), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase);
        var afterRequest = created > requestedUtc;
        Log($"{label}: CreateInstance {createMs} ms, Hwnd PID {pid}, image {Path.GetFileName(image)}, created {created:O} (requested {requestedUtc:O}), new PID={isNew}");
        Log($"{label}: command line {Redact(commandLine)}; contains /automation={commandLine?.Contains("/automation", StringComparison.OrdinalIgnoreCase)}");
        Log($"{label}: Visible={(bool)app.Visible}");
        if (!(isNew && nameOk && afterRequest))
        {
            handle.Dispose();
            throw new NotOwnedException($"PID {pid} new={isNew} name={nameOk} createdAfterRequest={afterRequest}");
        }

        return (app, new Owned((int)pid, created, handle));
    }

    private static void RequireNoWorkbooks(string label, dynamic app)
    {
        dynamic workbooks = app.Workbooks;
        dynamic protectedViews = app.ProtectedViewWindows;
        int workbookCount = workbooks.Count;
        int protectedCount = protectedViews.Count;
        Marshal.FinalReleaseComObject(workbooks);
        Marshal.FinalReleaseComObject(protectedViews);
        Log($"{label}: Workbooks={workbookCount}, ProtectedViewWindows={protectedCount}");
        if (workbookCount != 0 || protectedCount != 0) throw new NotOwnedException("the Excel has open workbooks");
    }

    private static void QuitAndWait(string label, dynamic app, Owned owned)
    {
        RequireNoWorkbooks(label, app);
        var watch = Stopwatch.StartNew();
        app.Quit();
        var quitMs = watch.ElapsedMilliseconds;
        Marshal.FinalReleaseComObject(app);
        // Intermediate RCWs (for example Worksheets[1]) are released only by the GC.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        var ended = owned.WaitForExit(75_000);
        Log($"{label}: Quit returned in {quitMs} ms; process {(ended ? $"ended {watch.ElapsedMilliseconds - quitMs} ms after Quit returned, exit code {owned.ExitCode()}" : "did not end within 75 s")}");
        if (!ended) owned.Terminate(label);
        owned.Dispose();
    }

    private static void RunSta(string label, Action action)
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (Exception e) { failure = e; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (failure is NotOwnedException) throw failure;
        if (failure != null) Log($"{label}: FAILED {failure.GetType().Name} 0x{failure.HResult:X8}: {failure.Message}");
    }

    private static readonly Guid IID_IDispatch = new("00020400-0000-0000-C000-000000000046");
    // Hands the interface to another apartment; each stream is unmarshaled once.
    private static IntPtr MarshalToStream(object app)
    {
        var iid = IID_IDispatch;
        Marshal.ThrowExceptionForHR(Native.CoMarshalInterThreadInterfaceInStream(ref iid, app, out var stream));
        return stream;
    }

    private static object UnmarshalFromStream(IntPtr stream)
    {
        var iid = IID_IDispatch;
        Marshal.ThrowExceptionForHR(Native.CoGetInterfaceAndReleaseStream(stream, ref iid, out var app));
        return app;
    }

    private static Dictionary<int, DateTime> Snapshot()
    {
        var result = new Dictionary<int, DateTime>();
        foreach (var p in Process.GetProcessesByName("EXCEL"))
        {
            using (p)
            {
                try { result[p.Id] = p.StartTime.ToUniversalTime(); }
                catch (Exception) { result[p.Id] = DateTime.MinValue; }
            }
        }

        return result;
    }

    private static string Describe(Dictionary<int, DateTime> processes) =>
        string.Join(", ", processes.Select(p => $"{p.Key}@{p.Value:O}"));

    // Keep the executable's file name and the switches; drop the install folder from the log.
    private static string Redact(string? commandLine)
    {
        if (commandLine == null) return "<unreadable>";
        var close = commandLine.StartsWith('"') ? commandLine.IndexOf('"', 1) : -1;
        return close > 0 ? "\"…\\" + Path.GetFileName(commandLine[1..close]) + "\"" + commandLine[(close + 1)..] : commandLine;
    }

    private static readonly long s_start = Stopwatch.GetTimestamp();
    private static readonly object s_logLock = new();

    internal static void Log(string message)
    {
        lock (s_logLock)
        {
            Console.WriteLine($"[{Stopwatch.GetElapsedTime(s_start).TotalSeconds,8:F3}] {message}");
        }
    }
}

internal sealed class NotOwnedException(string message) : Exception(message);

// The owned Excel: the same process handle is used to check identity, wait, and terminate.
internal sealed class Owned(int pid, DateTime createdUtc, SafeProcessHandle handle) : IDisposable
{
    public bool WaitForExit(int milliseconds) => Native.WaitForSingleObject(handle, milliseconds) == 0;

    public int ExitCode() => Native.GetExitCodeProcess(handle, out var code) ? code : -1;

    public void Terminate(string label)
    {
        // Same handle as at ownership: the PID cannot have been reused for another process.
        if (Native.CreationTimeUtc(handle) != createdUtc) throw new InvalidOperationException("identity changed");
        if (WaitForExit(0))
        {
            Program.Log($"{label}: PID {pid} had already ended before TerminateProcess");
            return;
        }

        var ok = Native.TerminateProcess(handle, 1);
        var watch = Stopwatch.StartNew();
        var ended = WaitForExit(5000);
        Program.Log($"{label}: TerminateProcess PID {pid} ok={ok}; ended={ended} in {watch.ElapsedMilliseconds} ms");
    }

    public void Dispose() => handle.Dispose();
}

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

// Retries SERVERCALL_RETRYLATER for up to 10 s (docs/architecture.md); cancels SERVERCALL_REJECTED. Records every callback.
internal sealed class MessageFilter(string label) : IOleMessageFilter
{
    private const int ServerCallRejected = 1;
    private const int ServerCallRetryLater = 2;
    private readonly List<string> _rejections = new();
    private int _messagePending;

    public static MessageFilter Register(string label)
    {
        var filter = new MessageFilter(label);
        Marshal.ThrowExceptionForHR(Native.CoRegisterMessageFilter(filter, out _));
        return filter;
    }

    public static void Revoke() => Marshal.ThrowExceptionForHR(Native.CoRegisterMessageFilter(null, out _));

    public void Report() =>
        Program.Log($"{label} filter: RetryRejectedCall {_rejections.Count} time(s) [{string.Join("; ", _rejections.Take(5))}{(_rejections.Count > 5 ? "; …" : "")}], MessagePending {_messagePending} time(s)");

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        var type = dwRejectType switch { ServerCallRejected => "REJECTED", ServerCallRetryLater => "RETRYLATER", _ => dwRejectType.ToString() };
        _rejections.Add($"{type} at tick {dwTickCount}");
        return dwRejectType == ServerCallRetryLater && dwTickCount < 10_000 ? 100 : -1;
    }

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
    {
        _messagePending++;
        return 2; // PENDINGMSG_WAITDEFPROCESS
    }
}

internal static class Native
{
    public const int ProcessAccess = 0x1000 | 0x00100000 | 0x0001; // QUERY_LIMITED_INFORMATION | SYNCHRONIZE | TERMINATE

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern SafeProcessHandle OpenProcess(int access, bool inherit, int processId);

    [DllImport("kernel32.dll")]
    public static extern int WaitForSingleObject(SafeProcessHandle handle, int milliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool TerminateProcess(SafeProcessHandle handle, int exitCode);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetExitCodeProcess(SafeProcessHandle handle, out int exitCode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetProcessTimes(SafeProcessHandle handle, out long creation, out long exit, out long kernel, out long user);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool QueryFullProcessImageName(SafeProcessHandle handle, int flags, StringBuilder name, ref int size);

    [DllImport("ntdll.dll")]
    private static extern int NtQueryInformationProcess(SafeProcessHandle handle, int infoClass, IntPtr buffer, int length, out int returnLength);

    [DllImport("ole32.dll")]
    public static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    [DllImport("ole32.dll")]
    public static extern int CoMarshalInterThreadInterfaceInStream(ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] object unknown, out IntPtr stream);

    [DllImport("ole32.dll")]
    public static extern int CoGetInterfaceAndReleaseStream(IntPtr stream, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object unknown);

    public static DateTime CreationTimeUtc(SafeProcessHandle handle) =>
        GetProcessTimes(handle, out var creation, out _, out _, out _) ? DateTime.FromFileTimeUtc(creation) : throw new Win32Exception();

    public static string ImagePath(SafeProcessHandle handle)
    {
        var size = 1024;
        var name = new StringBuilder(size);
        return QueryFullProcessImageName(handle, 0, name, ref size) ? name.ToString() : throw new Win32Exception();
    }

    // ProcessCommandLineInformation (60) returns a UNICODE_STRING followed by its characters.
    public static string? CommandLine(SafeProcessHandle handle)
    {
        const int ProcessCommandLineInformation = 60;
        _ = NtQueryInformationProcess(handle, ProcessCommandLineInformation, IntPtr.Zero, 0, out var length);
        if (length <= 0) return null;
        var buffer = Marshal.AllocHGlobal(length);
        try
        {
            if (NtQueryInformationProcess(handle, ProcessCommandLineInformation, buffer, length, out _) != 0) return null;
            var bytes = (ushort)Marshal.ReadInt16(buffer, 0);
            return Marshal.PtrToStringUni(Marshal.ReadIntPtr(buffer, IntPtr.Size), bytes / 2);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }
}
