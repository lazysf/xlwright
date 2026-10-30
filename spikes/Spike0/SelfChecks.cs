using System.Text.Json;

namespace Spike0;

// Focused checks for this experiment. No COM activation or test-framework dependency.
internal static class SelfChecks
{
    internal static async Task<int> Run(string repository)
    {
        int passed = 0;
        void Check(string name, Action action)
        {
            action();
            passed++;
            Console.WriteLine("PASS " + name);
        }
        RuntimeAssets.VerifyPias();
        RuntimeAssets.VerifyFixture(repository);
        Check("S10 fixed case selection", () =>
        {
            Assert(StartupProbe.Parse(["--excel-s10-a2"]) == "A2");
            Assert(StartupProbe.Parse(["--excel-s10-a"]) is null);
            Assert(StartupProbe.Parse(["--excel-s10-b"]) == "B");
            Assert(StartupProbe.Parse(["--excel-s10-c"]) == "C");
            Assert(StartupProbe.Parse(["--excel-s10-d"]) is null);
            Assert(StartupProbe.Parse(["--excel-s10-a", "other.xlsm"]) is null);
            Assert(StartupProbe.Parse(["--excel-s10-z"]) is null);
        });
        Check("S10 positive and negative evidence", () =>
        {
            StartupProbe.RequireTrace("A2", "open", ["open"]);
            StartupProbe.RequireTrace("A2", "closed", ["open", "before-close"]);
            StartupProbe.RequireTrace("B", "open", []);
            StartupProbe.RequireTrace("B", "closed", ["auto-open"]);
            StartupProbe.RequireTrace("C", "closed", []);
            Expect<InvalidDataException>(() => StartupProbe.RequireTrace("A2", "open", []));
            Expect<InvalidDataException>(() => StartupProbe.RequireTrace("B", "closed", ["auto-open", "before-close"]));
            Expect<InvalidDataException>(() => StartupProbe.RequireTrace("C", "open", ["open"]));
            Expect<InvalidDataException>(() => StartupProbe.RequireTrace("D", "closed", []));
        });
        Check("S10 revised case ordering", () =>
        {
            Assert(StartupProbe.PreviousCase("A2") is null);
            Assert(StartupProbe.PreviousCase("B") == "A2");
            Assert(StartupProbe.PreviousCase("C") == "B");
            Expect<InvalidDataException>(() => StartupProbe.PreviousCase("A"));
        });
        Check("S10 unopened calculation never calls the getter", () =>
        {
            bool called = false;
            int? observed = StartupProbe.ReadCalculation(false, () => { called = true; throw new InvalidOperationException(); });
            Assert(observed is null && !called);
        });
        Check("S10 open calculation reads once and propagates failure", () =>
        {
            int calls = 0;
            Assert(StartupProbe.ReadCalculation(true, () => { calls++; return -4105; }) == -4105 && calls == 1);
            Expect<InvalidOperationException>(() => StartupProbe.ReadCalculation(true, () => throw new InvalidOperationException()));
        });
        string directory = RuntimeAssets.CreateLocalDirectory(repository,
            Path.Combine(".xlwright", "spike0-checks", Guid.NewGuid().ToString("N")));
        string workbook = Path.Combine(directory, "lifecycle.xlsx");
        Check("S10 repeated case is rejected before creating a guard", () =>
        {
            string attemptPath = Path.Combine(directory, "repeat.attempt");
            string guardPath = Path.Combine(directory, "repeat.guard");
            StartupProbe.RequireNewAttempt(attemptPath);
            File.WriteAllText(attemptPath, "original evidence");
            Expect<IOException>(() =>
            {
                StartupProbe.RequireNewAttempt(attemptPath);
                File.WriteAllText(guardPath, "must not be created");
            });
            Assert(File.ReadAllText(attemptPath) == "original evidence" && !File.Exists(guardPath));
            Expect<IOException>(() => StartupProbe.RequireNewAttempt(directory));
        });
        Check("S10 fixed hash rejects a modified copy", () =>
        {
            string probe = Path.Combine(directory, "s10-hash-check.bin");
            File.WriteAllText(probe, "reviewed fixture");
            string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(probe)));
            StartupProbe.VerifyHash(probe, hash);
            File.WriteAllText(probe, "modified fixture");
            Expect<InvalidDataException>(() => StartupProbe.VerifyHash(probe, hash));
        });
        var owner = new ProcessIdentity(12345, new DateTime(2026, 10, 29, 0, 0, 0, DateTimeKind.Utc), "EXCEL");
        var owned = new OwnershipRecord(1, "owned", workbook, owner);
        string ownedJson = JsonSerializer.Serialize(owned, OwnershipRecord.JsonOptions);

        Check("record round-trip and UTC identity", () => Assert(OwnershipRecord.Parse(ownedJson) == owned));
        Check("invalid records fail closed", () =>
        {
            foreach (string text in new[]
            {
                "", "{", "[]", "{}", "{\"version\":2,\"kind\":\"completed\"}",
                "{\"version\":1,\"kind\":\"unknown\"}", "{\"version\":1,\"kind\":\"owned\"}",
                "{\"version\":1,\"kind\":\"starting\",\"workbook_path\":\"relative.xlsx\"}",
                "{\"version\":1,\"version\":1,\"kind\":\"completed\"}",
                ownedJson.Replace("12345", "0"), ownedJson.Replace(".0000000Z", ".0000000+09:00")
            }) Expect<SessionRecordInvalidException>(() => OwnershipRecord.Parse(text));
        });
        Check("exclusive handle and verified truncation without BOM", () =>
        {
            string path = Path.Combine(directory, "exclusive.lock");
            using (var file = OwnershipRecordFile.Acquire(path))
            {
                Assert(file.CreatedNew);
                file.Write(owned);
                Expect<IOException>(() => { using var other = OwnershipRecordFile.Acquire(path); });
                file.Write(new(1, "completed"));
                Assert(file.Read() == new OwnershipRecord(1, "completed"));
            }
            byte[] bytes = File.ReadAllBytes(path);
            Assert(bytes[0] == (byte)'{');
            Assert(System.Text.Encoding.UTF8.GetString(bytes) == "{\"version\":1,\"kind\":\"completed\"}");
            using var reopened = OwnershipRecordFile.Acquire(path);
            Assert(!reopened.CreatedNew);
            reopened.CheckPrevious();
        });
        Check("invalid and starting records are preserved", () =>
        {
            foreach (string text in new[] { "", "broken", JsonSerializer.Serialize(
                new OwnershipRecord(1, "starting", workbook), OwnershipRecord.JsonOptions) })
            {
                string path = Path.Combine(directory, Guid.NewGuid() + ".lock");
                File.WriteAllText(path, text);
                using (var file = OwnershipRecordFile.Acquire(path))
                    Expect<SessionRecordInvalidException>(file.CheckPrevious);
                Assert(File.ReadAllText(path) == text);
            }
        });
        Check("live PID is rejected without altering record", () =>
        {
            using var current = System.Diagnostics.Process.GetCurrentProcess();
            var existing = owned with { Process = ProcessIdentity.Read(current) };
            string path = Path.Combine(directory, "live.lock");
            using (var created = OwnershipRecordFile.Acquire(path)) created.Write(existing);
            using var file = OwnershipRecordFile.Acquire(path);
            Expect<InvalidOperationException>(file.CheckPrevious);
            Assert(file.Read() == existing);
            file.Write(existing with { Process = existing.Process! with { Name = "different" } });
            Expect<InvalidOperationException>(file.CheckPrevious);
        });
        Check("I/O failure does not permit the next stage", () =>
        {
            var gate = new ExecutionGate();
            gate.Begin("record", TimeSpan.FromSeconds(120));
            var file = OwnershipRecordFile.Acquire(Path.Combine(directory, "closed-handle.lock"));
            file.Dispose();
            bool nextRan = false;
            try
            {
                gate.Step("write", () => file.Write(owned));
                gate.Step("next", () => nextRan = true);
                throw new InvalidOperationException("Expected a failed write.");
            }
            catch (ObjectDisposedException) { gate.Fault(generationStarted: true); }
            Assert(!nextRan && gate.Snapshot.State == "isolated");
        });
        Check("completion wins before deadline", () =>
        {
            var clock = new ManualClock();
            var gate = new ExecutionGate(clock);
            gate.Begin("normal", TimeSpan.FromSeconds(120));
            clock.Advance(TimeSpan.FromSeconds(119));
            gate.Complete("ready");
            clock.Advance(TimeSpan.FromSeconds(2));
            Assert(!gate.CheckTimeout() && gate.Snapshot.State == "ready");
        });
        Check("deadline rejects completion and a subsequent operation", () =>
        {
            var clock = new ManualClock();
            var gate = new ExecutionGate(clock);
            gate.Begin("deadline", TimeSpan.FromSeconds(120));
            clock.Advance(TimeSpan.FromSeconds(120));
            Expect<WorkStoppedException>(() => gate.Complete("ready"));
            gate.Fault(generationStarted: true);
            Assert(gate.Snapshot.State == "isolated");
            Expect<WorkStoppedException>(() => gate.Begin("new", TimeSpan.FromSeconds(120)));
        });

        var delayedClock = new ManualClock();
        var delayedGate = new ExecutionGate(delayedClock);
        delayedGate.Begin("delayed", TimeSpan.FromSeconds(120));
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        bool delayedNextRan = false;
        Task delayed = Task.Run(() =>
        {
            delayedGate.Step("active-call", () => { entered.Set(); release.Wait(); });
            Expect<WorkStoppedException>(() => delayedGate.Step("late-next", () => delayedNextRan = true));
            Expect<WorkStoppedException>(() => delayedGate.Complete("ready"));
            delayedGate.Fault(generationStarted: true);
        });
        if (!entered.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException("Self-check worker did not start.");
        delayedClock.Advance(TimeSpan.FromSeconds(120));
        Assert(delayedGate.CheckTimeout() && delayedGate.Snapshot.Busy);
        release.Set();
        await delayed.WaitAsync(TimeSpan.FromSeconds(5));
        Check("late return preserves isolation and skips the next step", () =>
            Assert(!delayedNextRan && !delayedGate.Snapshot.Busy && delayedGate.Snapshot.State == "isolated"));

        Check("shutdown only permits an explicit cleanup operation", () =>
        {
            var gate = new ExecutionGate();
            gate.Begin("active", TimeSpan.FromSeconds(120));
            gate.RequestShutdown();
            Expect<WorkStoppedException>(() => gate.Step("normal-next", () => { }));
            gate.Fault(generationStarted: true);
            Expect<WorkStoppedException>(() => gate.Begin("normal", TimeSpan.FromSeconds(120)));
            gate.Begin("shutdown", TimeSpan.FromSeconds(120), forShutdown: true);
            gate.Step("cleanup", () => { });
            gate.Complete("closed");
            Assert(gate.Snapshot.State == "closed");
        });
        Check("STA message loop starts without Excel", () =>
            Assert(StaRunner.Start(() => Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                .WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult()));

        Check("shutdown deadline includes draining the active call", () =>
        {
            var clock = new ManualClock();
            var gate = new ExecutionGate(clock);
            gate.Begin("active", TimeSpan.FromSeconds(120));
            clock.Advance(TimeSpan.FromSeconds(1));
            gate.RequestShutdown();
            clock.Advance(TimeSpan.FromSeconds(118));
            gate.RequestShutdown(); // Repeated requests must not restart the clock.
            gate.Fault(generationStarted: true);
            gate.Begin("shutdown", TimeSpan.FromSeconds(120), forShutdown: true);
            clock.Advance(TimeSpan.FromSeconds(2));
            Assert(gate.CheckTimeout());
            Expect<WorkStoppedException>(() => gate.Step("late-cleanup", () => { }));
        });

        LifecycleChecks.Run(directory, Check);
        Check("retained record verifies exact sharing violation", () =>
        {
            using var file = OwnershipRecordFile.Acquire(Path.Combine(directory, "retained-proof.lock"));
            file.Write(owned);
            file.VerifyRetained(owned);
            Expect<IOException>(() => file.VerifyRetained(new(1, "completed")));
        });
        Console.WriteLine($"{passed} self-checks passed. Excel was not activated.");
        return 0;
    }

    private static void Assert(bool condition)
    {
        if (!condition) throw new InvalidOperationException("Self-check assertion failed.");
    }

    private static void Expect<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException("Expected " + typeof(T).Name);
    }

    private sealed class ManualClock : TimeProvider
    {
        private long timestamp;
        public override long TimestampFrequency => TimeSpan.TicksPerSecond;
        public override long GetTimestamp() => Interlocked.Read(ref timestamp);
        internal void Advance(TimeSpan duration) => Interlocked.Add(ref timestamp, duration.Ticks);
    }
}
