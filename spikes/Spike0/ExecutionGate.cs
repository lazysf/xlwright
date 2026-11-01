namespace Spike0;

internal sealed record ExecutionSnapshot(string State, bool Busy, string Stage, string? StopReason);

// Only the STA executes work. The monitor revokes permission without calling COM.
internal sealed class ExecutionGate(TimeProvider? clock = null)
{
    private readonly object sync = new();
    private readonly TimeProvider clock = clock ?? TimeProvider.System;
    private long started;
    private long shutdownRequestedAt;
    private static readonly TimeSpan ShutdownTimeout = TimeSpan.FromSeconds(120);
    private TimeSpan limit;
    private string state = "closed";
    private string stage = "prepared";
    private string? stopReason;
    private bool busy;
    private bool shutdown;

    internal ExecutionSnapshot Snapshot
    {
        get { lock (sync) return new(state, busy, stage, stopReason); }
    }

    internal void Begin(string operation, TimeSpan timeout, bool forShutdown = false)
    {
        lock (sync)
        {
            Expire();
            if (busy || state == "isolated" && stopReason != "shutdown" ||
                stopReason is not null && !(forShutdown && stopReason == "shutdown") ||
                forShutdown && stopReason != "shutdown")
                throw new WorkStoppedException();
            shutdown = forShutdown;
            started = forShutdown ? shutdownRequestedAt : clock.GetTimestamp();
            limit = timeout;
            stage = operation;
            busy = true;
        }
    }

    internal T Step<T>(string nextStage, Func<T> action)
    {
        lock (sync)
        {
            RequirePermission();
            stage = nextStage;
            // Permission to start is committed here. An active call is never cancelled.
        }
        return action();
    }

    internal void Step(string nextStage, Action action) => Step(nextStage, () =>
    {
        action();
        return true;
    });

    internal void Complete(string nextState)
    {
        lock (sync)
        {
            RequirePermission();
            state = nextState;
            busy = false;
        }
    }

    internal void Fault(bool generationStarted, bool preserveReady = false)
    {
        lock (sync)
        {
            Expire();
            state = generationStarted ?
                (state == "ready" && preserveReady && stopReason is null ? "ready" : "isolated") : "closed";
            busy = false; // The worker has unwound; no COM call is still running here.
        }
    }

    internal bool CheckTimeout()
    {
        lock (sync)
        {
            Expire();
            return stopReason == "timeout";
        }
    }

    internal void RevokeForDiagnosticFailure()
    {
        lock (sync)
        {
            stopReason = "diagnostic-failure";
            state = "isolated";
            // The worker, not the observer, clears busy after its active call unwinds.
        }
    }

    internal void RequestShutdown()
    {
        lock (sync)
        {
            Expire();
            if (stopReason is null)
            {
                stopReason = "shutdown";
                shutdownRequestedAt = clock.GetTimestamp();
            }
        }
    }

    internal TimeSpan Remaining
    {
        get
        {
            lock (sync)
            {
                RequirePermission();
                return limit - clock.GetElapsedTime(started);
            }
        }
    }

    private void RequirePermission()
    {
        Expire();
        if (!busy || stopReason is not null && !(shutdown && stopReason == "shutdown"))
            throw new WorkStoppedException();
    }

    private void Expire()
    {
        if (busy && clock.GetElapsedTime(started) >= limit ||
            stopReason == "shutdown" && clock.GetElapsedTime(shutdownRequestedAt) >= ShutdownTimeout)
        {
            stopReason = "timeout";
            state = "isolated";
        }
    }
}

internal sealed class WorkStoppedException : Exception
{
    internal WorkStoppedException() : base("Operation permission was revoked; no next step is allowed.") { }
}
