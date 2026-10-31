using System.Windows.Forms;

namespace Spike0;

internal static class S9Checks
{
    internal static void Run(string directory, Action<string, Action> check)
    {
        string evidence = Path.Combine(directory, "s9");
        Directory.CreateDirectory(evidence);
        int observation = 0;
        void Save(RetryFilter filter) => File.WriteAllText(Path.Combine(evidence, (++observation) + ".json"),
            System.Text.Json.JsonSerializer.Serialize(filter.Snapshot()));
        void Require(bool value) { if (!value) throw new InvalidOperationException("S9 assertion failed."); }
        check("S9 dispatches posted work on its STA and restores its filter", () =>
        {
            Control? control = null;
            int dispatched = 0;
            Require(StaRunner.Start(() =>
            {
                try
                {
                    var gate = new ExecutionGate();
                    using var filter = new RetryFilter(gate);
                    filter.Register();
                    filter.Dispose();
                    return dispatched == Environment.CurrentManagedThreadId && filter.Restored;
                }
                finally { control!.Dispose(); }
            }, () =>
            {
                control = new Control();
                _ = control.Handle;
                control.BeginInvoke((Action)(() => dispatched = Environment.CurrentManagedThreadId));
            }).WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult());
        });
        foreach (int type in new[] { 1, 2, 99 })
            check("S9 rejection type " + type, () =>
            {
                var gate = new ExecutionGate();
                gate.Begin("synthetic", TimeSpan.FromSeconds(30));
                var filter = new RetryFilter(gate);
                filter.BeginCall();
                Require(filter.RetryRejectedCall(0, 0, type) == (type == 2 ? 100 : -1));
                filter.EndCall();
                Save(filter);
            });
        foreach (int milliseconds in new[] { 9900, 9901, 10000 })
            check("S9 retry budget boundary " + milliseconds, () =>
            {
                var clock = new Clock();
                var gate = new ExecutionGate(clock);
                gate.Begin("synthetic", TimeSpan.FromSeconds(30));
                var filter = new RetryFilter(gate, clock);
                filter.BeginCall();
                clock.Advance(milliseconds);
                Require(filter.RetryRejectedCall(0, milliseconds, 2) == (milliseconds == 9900 ? 100 : -1));
                filter.EndCall();
                Save(filter);
            });
        check("S9 multiple calls share the retry budget", () =>
        {
            var clock = new Clock();
            var gate = new ExecutionGate(clock);
            gate.Begin("synthetic", TimeSpan.FromSeconds(30));
            var filter = new RetryFilter(gate, clock);
            filter.BeginCall();
            Require(filter.RetryRejectedCall(0, 0, 2) == 100);
            clock.Advance(6000);
            filter.EndCall();
            Save(filter);
            filter.BeginCall();
            clock.Advance(4000);
            Require(filter.RetryRejectedCall(0, 4000, 2) == -1);
            filter.EndCall();
            Save(filter);
        });
        check("S9 operation deadline takes precedence", () =>
        {
            var clock = new Clock();
            var gate = new ExecutionGate(clock);
            gate.Begin("synthetic", TimeSpan.FromSeconds(2));
            var filter = new RetryFilter(gate, clock);
            filter.BeginCall();
            clock.Advance(1901);
            Require(filter.RetryRejectedCall(0, 1901, 2) == -1 && filter.DeadlineRefusal);
            clock.Advance(99);
            Require(gate.CheckTimeout() && gate.Snapshot.Busy);
            filter.EndCall();
            gate.Fault(true);
        });
        foreach (bool fails in new[] { false, true })
            check("S9 accepted late return is not replayed: exception=" + fails, () =>
            {
                var clock = new Clock();
                var gate = new ExecutionGate(clock);
                gate.Begin("synthetic", TimeSpan.FromSeconds(2));
                int calls = 0;
                try
                {
                    gate.Step("accepted", () =>
                    {
                        calls++;
                        clock.Advance(2000);
                        Require(gate.CheckTimeout() && gate.Snapshot.Busy);
                        try { gate.Begin("concurrent", TimeSpan.FromSeconds(30)); throw new InvalidOperationException(); }
                        catch (WorkStoppedException) { }
                        if (fails) throw new IOException("Synthetic accepted failure");
                    });
                }
                catch (IOException) { }
                foreach (string stage in new[] { "write", "save", "close", "record-completed" })
                {
                    try { gate.Step(stage, () => { calls++; }); throw new InvalidOperationException(); }
                    catch (WorkStoppedException) { }
                }
                try { gate.Complete("ready"); throw new InvalidOperationException(); }
                catch (WorkStoppedException) { }
                gate.Fault(true);
                Require(calls == 1 && !gate.Snapshot.Busy && gate.Snapshot.State == "isolated");
            });
    }
    private sealed class Clock : TimeProvider
    {
        private long ticks;
        public override long TimestampFrequency => TimeSpan.TicksPerSecond;
        public override long GetTimestamp() => ticks;
        internal void Advance(int milliseconds) => ticks += TimeSpan.FromMilliseconds(milliseconds).Ticks;
    }
}
