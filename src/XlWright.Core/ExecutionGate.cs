namespace XlWright.Core;

internal enum SessionState { Closed, Ready, Isolated }

internal sealed record ExecutionSnapshot(SessionState State, bool Busy, string Stage, string? StopReason);

// Only the STA executes work. The monitor revokes permission without calling COM.
internal sealed class ExecutionGate(TimeProvider? clock = null)
{
    private readonly object sync = new();
    private readonly TimeProvider clock = clock ?? TimeProvider.System;
    private long started;
    private long shutdownRequestedAt;
    private static readonly TimeSpan ShutdownTimeout = TimeSpan.FromSeconds(120);
    private TimeSpan limit;
    private SessionState state = SessionState.Closed;
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
            if (busy || state == SessionState.Isolated && stopReason != "shutdown" ||
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

    // The public acceptance order. It checks and begins under one lock, so two requests cannot both be accepted.
    internal void Accept(string operation, TimeSpan timeout, SessionState required)
    {
        lock (sync)
        {
            Expire();
            if (busy)
                throw new CoreException(ErrorCode.Conflict, ErrorDetail.OperationBusy,
                    "Another operation is running. Check session_status and retry after busy is false.");
            if (state != required)
                throw (required, state) switch
                {
                    (SessionState.Closed, _) => new CoreException(ErrorCode.Conflict, ErrorDetail.SessionActive,
                        "A session is already active. End it before starting another."),
                    (_, SessionState.Closed) => new CoreException(ErrorCode.Conflict, ErrorDetail.SessionNotStarted,
                        "No session is active. Call session_start first."),
                    _ => new CoreException(ErrorCode.Conflict, ErrorDetail.SessionIsolated,
                        "The session is isolated. Do not resend; follow manual recovery.")
                };
            Begin(operation, timeout);
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

    internal void Complete(SessionState nextState)
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
                (state == SessionState.Ready && preserveReady && stopReason is null ? SessionState.Ready : SessionState.Isolated) :
                SessionState.Closed;
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
            state = SessionState.Isolated;
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
            state = SessionState.Isolated;
        }
    }
}

internal sealed class WorkStoppedException : Exception
{
    internal WorkStoppedException() : base("Operation permission was revoked; no next step is allowed.") { }
}
