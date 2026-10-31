using System.Runtime.InteropServices;
using System.Text.Json;

namespace Spike0;

internal sealed class S9Probe(string scenario, string run, ExecutionGate gate) : IDisposable
{
    private readonly ManualResetEventSlim start = new();
    private int armed;
    private int notified;
    internal RetryFilter? Filter { get; private set; }
    internal static string? Parse(string[] args) => args.Length == 1 ? args[0] switch
    {
        "--excel-s9-release" => "release", "--excel-s9-limit" => "limit",
        "--excel-s9-deadline" => "deadline", _ => null
    } : null;

    internal static void RequirePrevious(string repository, string workspace, string scenario)
    {
        if (scenario == "release") return;
        // B and C require A's actual success. B's failure is never promoted to success.
        string Read(string file)
        {
            string path = Path.Combine(workspace, file);
            RuntimeAssets.RejectReparsePoints(repository, path);
            return File.ReadAllText(path);
        }
        string attempt = Read("s9-release.attempt");
        if (attempt.Length == 0 || attempt != Read("s9-release.success"))
            throw new InvalidDataException("S9 release evidence is incomplete.");
        if (scenario == "deadline" && Read("s9-limit.attempt").Length == 0)
            throw new InvalidDataException("Evaluate B and separately approve recovery before C.");
    }

    internal void SignalStart() { if (Volatile.Read(ref armed) == 1) start.Set(); }
    internal void WaitForStart()
    {
        Volatile.Write(ref armed, 1);
        Console.WriteLine("S9: On the owned copy, edit Lifecycle!D2 with F2 without typing. Enter go here. Keep editing until instructed.");
        while (!start.Wait(50))
            if (gate.Snapshot.StopReason is not null) throw new WorkStoppedException();
        Volatile.Write(ref armed, 0);
    }

    internal TimeSpan Timeout => TimeSpan.FromSeconds(scenario == "deadline" ? 2 : 30);
    internal bool IsRelease => scenario == "release";
    internal bool IsDeadline => scenario == "deadline";
    internal RetryFilter Install() => Filter = new RetryFilter(gate);
    internal void NotifyRelease()
    {
        if (scenario == "release" && Filter?.ReleaseDue == true && Interlocked.Exchange(ref notified, 1) == 0)
            Console.WriteLine("S9: Release editing with Esc in the owned Excel now.");
    }
    internal void SaveEvidence(string outcome)
    {
        string path = Path.Combine(run, "s9-" + outcome + ".json");
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        JsonSerializer.Serialize(stream, new { scenario, outcome, snapshot = gate.Snapshot,
            runtime = Environment.Version.ToString(), build = Convert.ToHexString(
                System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(typeof(S9Probe).Assembly.Location))),
            filter = Filter?.Snapshot() });
        stream.Flush(true);
    }
    public void Dispose() => start.Dispose();
}

[ComVisible(true), Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IS9MessageFilter
{
    [PreserveSig] int HandleInComingCall(int callType, nint caller, int ticks, nint info);
    [PreserveSig] int RetryRejectedCall(nint callee, int ticks, int rejectType);
    [PreserveSig] int MessagePending(nint callee, int ticks, int pendingType);
}

internal sealed record RetryDecision(int Call, int RejectType, int Ticks, long Timestamp, double ElapsedMs, double BudgetMs, int Delay, string Reason);

[ComVisible(true), ClassInterface(ClassInterfaceType.None)]
public sealed class RetryFilter : IS9MessageFilter, IDisposable
{
    private readonly ExecutionGate gate;
    private readonly object sync = new();
    private readonly TimeProvider clock;
    internal RetryFilter(ExecutionGate gate, TimeProvider? time = null)
    {
        this.gate = gate;
        clock = time ?? TimeProvider.System;
    }
    private readonly List<RetryDecision> decisions = [];
    private IS9MessageFilter? previous;
    private bool installed;
    private long callStarted;
    private long? firstRejection;
    private TimeSpan spent;
    private bool active;
    private bool rejected;
    private bool recordingFailed;
    private int calls;
    private int ownerThread;
    private string? apartment;
    private long? returnedAt;
    private int? returnedHresult;
    internal bool Restored { get; private set; }
    internal bool ReleaseDue { get { lock (sync) return firstRejection is long first && clock.GetElapsedTime(first) >= TimeSpan.FromSeconds(2); } }
    internal void Register()
    {
        ownerThread = Environment.CurrentManagedThreadId;
        apartment = Thread.CurrentThread.GetApartmentState().ToString();
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
            CoRegisterMessageFilter(this, out previous) != 0)
            throw new InvalidOperationException("S9 message filter registration failed.");
        installed = true;
    }
    internal void BeginCall()
    {
        lock (sync)
        {
            if (active) throw new InvalidOperationException("S9 call already active.");
            active = true;
            rejected = false;
            callStarted = clock.GetTimestamp();
            calls++;
        }
    }
    internal void EndCall(int hresult = 0)
    {
        lock (sync)
        {
            if (rejected) spent += clock.GetElapsedTime(callStarted);
            active = false;
            returnedAt = clock.GetTimestamp();
            returnedHresult = hresult;
        }
    }
    internal bool ObservedRetry { get { lock (sync) return !recordingFailed && decisions.Any(d => d.Delay == 100); } }
    internal bool DeadlineRefusal { get { lock (sync) return decisions.LastOrDefault()?.Reason == "operation-deadline"; } }
    internal object Snapshot()
    {
        lock (sync) return new { calls, ownerThread, apartment, active, installed, Restored,
            recordingFailed, callStarted, returnedAt, returnedHresult, timestampFrequency = clock.TimestampFrequency,
            spentMs = spent.TotalMilliseconds,
            decisions = decisions.ToArray() };
    }
    public int HandleInComingCall(int callType, nint caller, int ticks, nint info) => 0;
    public int MessagePending(nint callee, int ticks, int pendingType) => 2;
    public int RetryRejectedCall(nint callee, int ticks, int rejectType)
    {
        lock (sync)
        {
            try
            {
                TimeSpan elapsed = clock.GetElapsedTime(callStarted);
                TimeSpan budget = TimeSpan.FromSeconds(10) - spent - elapsed;
                string reason = "retry-later";
                int delay = -1;
                if (!active || recordingFailed) reason = "inactive-or-recording-failed";
                else if (rejectType != 2) reason = "not-retry-later";
                else
                {
                    rejected = true;
                    firstRejection ??= clock.GetTimestamp();
                    try
                    {
                        if (gate.Remaining < TimeSpan.FromMilliseconds(100)) reason = "operation-deadline";
                        else if (budget < TimeSpan.FromMilliseconds(100)) reason = "retry-budget";
                        else delay = 100;
                    }
                    catch (WorkStoppedException) { reason = "permission-revoked"; }
                }
                if (decisions.Count >= 1024) throw new InvalidOperationException("S9 evidence capacity exceeded.");
                decisions.Add(new(calls, rejectType, ticks, clock.GetTimestamp(), elapsed.TotalMilliseconds, budget.TotalMilliseconds, delay, reason));
                return delay;
            }
            catch { recordingFailed = true; return -1; }
        }
    }
    public void Dispose()
    {
        if (!installed) return;
        if (Environment.CurrentManagedThreadId != ownerThread || CoRegisterMessageFilter(previous, out _) != 0)
            throw new InvalidOperationException("S9 message filter restoration failed.");
        installed = false;
        Restored = true;
    }
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IS9MessageFilter? filter, out IS9MessageFilter? previous);
}
