using System.Diagnostics;

namespace Spike0;

internal sealed class S11Probe(string scenario, string run, S11Evidence evidence) : IDisposable
{
    private readonly object sync = new();
    private ProcessIdentity? identity;
    private nint hwnd;
    private long lastObservation;
    private readonly ManualResetEventSlim shellReady = new();
    internal string Scenario => scenario;
    internal string OtherPath => Path.Combine(run, "s11-other.xlsx");
    internal S11Evidence Evidence => evidence;
    internal void SignalShell() => shellReady.Set();
    internal void WaitForShell()
    {
        if (!shellReady.Wait(TimeSpan.FromSeconds(300)))
            throw new TimeoutException("S11 shell input was not received; do not retry.");
    }

    internal void ObserveExcelProcesses(string phase)
    {
        var processes = Process.GetProcessesByName("EXCEL");
        try { evidence.Write(phase, "observed", processes.Select(ProcessIdentity.Read).ToArray()); }
        finally { foreach (var process in processes) process.Dispose(); }
    }

    internal static string? Parse(string[] args) => args.Length == 1 ? args[0] switch
    {
        "--excel-s11r2-baseline" => "baseline",
        "--excel-s11r2-protected-view" => "protected-view",
        "--excel-s11r2-wait" => "wait",
        "--excel-s11r2-modal" => "modal",
        "--excel-s11r2-shell2" => "shell2",
        _ => null
    } : null;

    internal void Identify(ProcessIdentity process, nint window)
    {
        lock (sync) { identity = process; hwnd = window; }
    }

    internal void Observe(ExecutionGate gate)
    {
        ProcessIdentity? target;
        nint window;
        lock (sync) { target = identity; window = hwnd; }
        if (target is null || Stopwatch.GetElapsedTime(lastObservation) < TimeSpan.FromSeconds(1)) return;
        lastObservation = Stopwatch.GetTimestamp();
        try { evidence.Write("windows", "observed", new { snapshot = gate.Snapshot, windows = S11Windows.Capture(target, window) }); }
        catch (Exception error)
        {
            gate.RevokeForDiagnosticFailure();
            evidence.Write("windows", "unavailable", new { error = error.GetType().Name, error.HResult });
            throw;
        }
    }

    public void Dispose() => shellReady.Dispose();
}
