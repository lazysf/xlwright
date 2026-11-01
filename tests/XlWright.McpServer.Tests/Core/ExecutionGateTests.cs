using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

// Ported from the Spike0 self-checks, plus the L6 and L7 conditions that do not need Excel.
public sealed class ExecutionGateTests
{
    private static readonly TimeSpan Limit = TimeSpan.FromSeconds(120);

    [Fact]
    public void CompletionBeforeDeadlineWins()
    {
        var clock = new ManualClock();
        var gate = new ExecutionGate(clock);
        gate.Begin("normal", Limit);
        clock.Advance(TimeSpan.FromSeconds(119));
        gate.Complete(SessionState.Ready);
        clock.Advance(TimeSpan.FromSeconds(2));

        Assert.False(gate.CheckTimeout());
        Assert.Equal(SessionState.Ready, gate.Snapshot.State);
    }

    [Fact]
    public void DeadlineRejectsCompletionAndLaterOperations()
    {
        var clock = new ManualClock();
        var gate = new ExecutionGate(clock);
        gate.Begin("deadline", Limit);
        clock.Advance(Limit);

        Assert.Throws<WorkStoppedException>(() => gate.Complete(SessionState.Ready));
        gate.Fault(generationStarted: true);
        Assert.Equal(SessionState.Isolated, gate.Snapshot.State);
        Assert.Throws<WorkStoppedException>(() => gate.Begin("new", Limit));
    }

    // L6: the response is decided once. A late return neither completes nor runs the next stage.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LateReturnKeepsIsolationAndSkipsNextSteps(bool fails)
    {
        var clock = new ManualClock();
        var gate = new ExecutionGate(clock);
        gate.Begin("accepted", TimeSpan.FromSeconds(2));
        int calls = 0;
        bool busyAtDeadline = false;
        var accepted = () => gate.Step("accepted", () =>
        {
            calls++;
            clock.Advance(TimeSpan.FromSeconds(2));
            busyAtDeadline = gate.CheckTimeout() && gate.Snapshot.Busy;
            Assert.Throws<WorkStoppedException>(() => gate.Begin("concurrent", Limit));
            if (fails) throw new IOException("Synthetic accepted failure");
        });
        if (fails) Assert.Throws<IOException>(accepted);
        else accepted();

        foreach (string stage in new[] { "write", "save", "close", "record-completed" })
            Assert.Throws<WorkStoppedException>(() => gate.Step(stage, () => calls++));
        Assert.Throws<WorkStoppedException>(() => gate.Complete(SessionState.Ready));
        gate.Fault(generationStarted: true);

        Assert.True(busyAtDeadline);
        Assert.Equal(1, calls);
        Assert.False(gate.Snapshot.Busy);
        Assert.Equal(SessionState.Isolated, gate.Snapshot.State);
    }

    [Fact(Timeout = 10000)]
    public async Task BlockedCallStaysBusyUntilItReturns()
    {
        var clock = new ManualClock();
        var gate = new ExecutionGate(clock);
        gate.Begin("delayed", Limit);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        bool nextRan = false;
        Task worker = Task.Run(() =>
        {
            gate.Step("active-call", () => { entered.Set(); release.Wait(); });
            Assert.Throws<WorkStoppedException>(() => gate.Step("late-next", () => nextRan = true));
            Assert.Throws<WorkStoppedException>(() => gate.Complete(SessionState.Ready));
            gate.Fault(generationStarted: true);
        }, TestContext.Current.CancellationToken);
        Assert.True(entered.Wait(TimeSpan.FromSeconds(5), TestContext.Current.CancellationToken));
        clock.Advance(Limit);

        Assert.True(gate.CheckTimeout());
        Assert.True(gate.Snapshot.Busy);
        Assert.Equal(SessionState.Isolated, gate.Snapshot.State);
        release.Set();
        await worker.WaitAsync(TimeSpan.FromSeconds(5), TestContext.Current.CancellationToken);

        Assert.False(nextRan);
        Assert.False(gate.Snapshot.Busy);
        Assert.Equal(SessionState.Isolated, gate.Snapshot.State);
    }

    // L7: after an EOF only the explicit cleanup operation may run.
    [Fact]
    public void ShutdownPermitsOnlyCleanup()
    {
        var gate = new ExecutionGate();
        gate.Begin("active", Limit);
        gate.RequestShutdown();

        Assert.Throws<WorkStoppedException>(() => gate.Step("normal-next", () => { }));
        gate.Fault(generationStarted: true);
        Assert.Throws<WorkStoppedException>(() => gate.Begin("normal", Limit));
        gate.Begin("shutdown", Limit, forShutdown: true);
        gate.Step("cleanup", () => { });
        gate.Complete(SessionState.Closed);
        Assert.Equal(SessionState.Closed, gate.Snapshot.State);
    }

    [Fact]
    public void CleanupWithoutShutdownRequestIsRejected()
    {
        var gate = new ExecutionGate();
        Assert.Throws<WorkStoppedException>(() => gate.Begin("shutdown", Limit, forShutdown: true));
    }

    [Fact]
    public void ShutdownDeadlineIncludesDrainingTheActiveCall()
    {
        var clock = new ManualClock();
        var gate = new ExecutionGate(clock);
        gate.Begin("active", Limit);
        clock.Advance(TimeSpan.FromSeconds(1));
        gate.RequestShutdown();
        clock.Advance(TimeSpan.FromSeconds(118));
        gate.RequestShutdown(); // Repeated requests must not restart the clock.
        gate.Fault(generationStarted: true);
        gate.Begin("shutdown", Limit, forShutdown: true);
        clock.Advance(TimeSpan.FromSeconds(2));

        Assert.True(gate.CheckTimeout());
        Assert.Throws<WorkStoppedException>(() => gate.Step("late-cleanup", () => { }));
    }

    [Fact]
    public void DiagnosticFailureRevokesTheNextStep()
    {
        var gate = new ExecutionGate();
        gate.Begin("active-call", TimeSpan.FromSeconds(30));
        gate.RevokeForDiagnosticFailure();

        Assert.Equal(new ExecutionSnapshot(SessionState.Isolated, true, "active-call", "diagnostic-failure"), gate.Snapshot);
        bool next = false;
        Assert.Throws<WorkStoppedException>(() => gate.Step("next", () => next = true));
        Assert.False(next);
    }

    [Fact]
    public void FailureBeforeGenerationReturnsToClosed()
    {
        var gate = new ExecutionGate();
        gate.Begin("start", Limit);
        gate.Fault(generationStarted: false);

        Assert.Equal(SessionState.Closed, gate.Snapshot.State);
        Assert.False(gate.Snapshot.Busy);
    }
}
