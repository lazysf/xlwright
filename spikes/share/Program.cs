// Throwaway experiment for the stage 2 open questions in docs/architecture.md "開始の順序" (the sharing check of
// step 4) and "残った専用のExcelの片付け" (opening the main workbook read-only while a dedicated Excel holds it).
// Owned dedicated Excel instances stand in for the daily Excel. It only ends an Excel whose ownership it
// confirmed: a new PID, named EXCEL, created after the start request.
//
// Usage: dotnet run --project spikes\share -- <workbook>

using System.Diagnostics;
using System.Runtime.InteropServices;

internal static class Program
{
    private static HashSet<int> s_before = [];

    private static int Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: Share <workbook>");
            return 64;
        }

        var path = Path.GetFullPath(args[0]);
        s_before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Log($"before: {s_before.Count} EXCEL process(es)");

        Log($"no Excel: {TryWrite(path)}");
        OnSta("edit holder", 120, () =>
        {
            var (holder, holderPid) = StartOwned("holder");
            try
            {
                dynamic book = holder.Workbooks.Open(path);
                Log($"holder opened for edit: ReadOnly {(bool)book.ReadOnly}");
                Log($"holder edits: {TryWrite(path)}");
                OnSta("read-only open", 30, () => OpenInAnother(path, readOnly: true));
                OnSta("plain open", 30, () => OpenInAnother(path, readOnly: false));
                book.Close(false);
            }
            finally
            {
                End(holder, holderPid);
            }
        });
        OnSta("read-only holder", 60, () =>
        {
            var (holder, holderPid) = StartOwned("holder");
            try
            {
                dynamic book = holder.Workbooks.Open(Filename: path, ReadOnly: true);
                Log($"holder opened read-only: ReadOnly {(bool)book.ReadOnly}");
                Log($"holder reads: {TryWrite(path)}");
                book.Close(false);
            }
            finally
            {
                End(holder, holderPid);
            }
        });

        var after = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Log($"after: {after.Count} EXCEL process(es); unchanged {after.SetEquals(s_before)}");
        return 0;
    }

    // Step 4 of the start: open the file for writing for a moment, letting others read and write.
    private static string TryWrite(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            return "opened for writing";
        }
        catch (IOException e)
        {
            return $"IOException 0x{e.HResult:X8}: {e.Message}";
        }
    }

    private static void OpenInAnother(string path, bool readOnly)
    {
        var (other, otherPid) = StartOwned("other");
        try
        {
            other.DisplayAlerts = false;
            var watch = Stopwatch.StartNew();
            dynamic book = readOnly ? other.Workbooks.Open(Filename: path, ReadOnly: true) : other.Workbooks.Open(path);
            Log($"other opened ({(readOnly ? "ReadOnly argument" : "no argument")}) in {watch.ElapsedMilliseconds} ms: ReadOnly {(bool)book.ReadOnly}");
            book.Close(false);
        }
        finally
        {
            End(other, otherPid);
        }
    }

    private static (dynamic App, int Pid) StartOwned(string label)
    {
        var requested = DateTime.UtcNow;
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
        _ = GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out var pid);
        using var process = Process.GetProcessById((int)pid);
        if (s_before.Contains((int)pid) || process.ProcessName != "EXCEL" || process.StartTime.ToUniversalTime() <= requested)
        {
            throw new InvalidOperationException($"{label}: PID {pid} is not a new EXCEL started after the request; it is left untouched");
        }

        s_owned[(int)pid] = process.StartTime.ToUniversalTime();
        Log($"{label}: owned PID {pid}");
        return (app, (int)pid);
    }

    private static readonly Dictionary<int, DateTime> s_owned = [];

    // Quit, then end the owned process if it stays; this experiment does not test the natural end.
    private static void End(dynamic app, int pid)
    {
        try { app.Quit(); } catch (COMException e) { Log($"Quit failed 0x{e.HResult:X8}"); }
        Marshal.FinalReleaseComObject(app);
        Kill(pid);
    }

    private static void Kill(int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            if (!s_owned.TryGetValue(pid, out var created) || process.StartTime.ToUniversalTime() != created)
            {
                Log($"PID {pid} does not match the owned record; it is left untouched");
                return;
            }

            if (!process.WaitForExit(5_000))
            {
                process.Kill();
                process.WaitForExit(5_000);
                Log($"PID {pid} ended by Kill after Quit");
            }
        }
        catch (ArgumentException)
        {
        }
    }

    private static void OnSta(string label, int seconds, Action action)
    {
        var thread = new Thread(() =>
        {
            try { action(); }
            catch (Exception e) { Log($"{label}: FAILED {e.GetType().Name} 0x{e.HResult:X8}: {e.Message}"); }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        var owned = s_owned.Keys.ToHashSet();
        thread.Start();
        if (!thread.Join(TimeSpan.FromSeconds(seconds)))
        {
            Log($"{label}: did not return within {seconds} s; ending the Excel it started");
            foreach (var pid in s_owned.Keys.Except(owned).ToList())
            {
                Kill(pid);
            }
        }
    }

    private static readonly long s_start = Stopwatch.GetTimestamp();

    private static void Log(string message) =>
        Console.WriteLine($"[{Stopwatch.GetElapsedTime(s_start).TotalSeconds,6:F1}] {message}");

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
