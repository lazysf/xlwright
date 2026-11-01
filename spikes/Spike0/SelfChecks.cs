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
        Check("prepared build must match its manifest", () =>
        {
            string copy = RuntimeAssets.CreateLocalDirectory(repository,
                Path.Combine(".xlwright", "spike0-checks", "prepared-" + Guid.NewGuid().ToString("N")));
            string dll = Path.Combine(copy, "Spike0.dll");
            File.WriteAllText(dll, "build");
            string hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData("build"u8));
            File.WriteAllText(Path.Combine(copy, "manifest.json"), JsonSerializer.Serialize(new
            {
                checkRun = "run", files = new[] { new { path = "Spike0.dll", sha256 = hash } }
            }));
            Assert(PreparedBuild.Verify(copy) == new PreparedBuild.Identity("run", hash));
            Expect<InvalidOperationException>(() => PreparedBuild.Require(repository, copy));
            File.WriteAllText(Path.Combine(copy, "extra.dll"), "extra");
            Expect<InvalidDataException>(() => PreparedBuild.Verify(copy));
            File.Delete(Path.Combine(copy, "extra.dll"));
            File.WriteAllText(dll, "other");
            Expect<InvalidDataException>(() => PreparedBuild.Verify(copy));
            File.Delete(dll);
            Expect<InvalidDataException>(() => PreparedBuild.Verify(copy));
            File.Delete(Path.Combine(copy, "manifest.json"));
            Expect<InvalidDataException>(() => PreparedBuild.Verify(copy));
            Directory.Delete(copy);
        });
        Check("S14 fixed selection and ordering", () =>
        {
            Assert(CalculationProbe.Parse(["--excel-s14-auto"]) == "auto");
            Assert(CalculationProbe.Parse(["--excel-s14-manual"]) == "manual");
            Assert(CalculationProbe.Parse(["--excel-s14-mode-change"]) == "mode-change");
            Assert(CalculationProbe.Parse(["--excel-s14-auto", "other.xlsm"]) is null);
            Assert(CalculationProbe.Parse(["--excel-s14-unknown"]) is null);
            Assert(CalculationProbe.Previous("auto") is null);
            Assert(CalculationProbe.Previous("manual") == "auto");
            Assert(CalculationProbe.Previous("mode-change") == "manual");
            Expect<InvalidDataException>(() => CalculationProbe.Previous("unknown"));
        });
        Check("S14 saved package rejects stale values, mode and external parts", () =>
        {
            string s14Directory = RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/spike0-checks/" + Guid.NewGuid().ToString("N"));
            void Verify(string name, string mode, int cached, bool external, bool unboundPrefix = false)
            {
                string path = Path.Combine(s14Directory, name + ".xlsm");
                using (var archive = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create))
                {
                    void Part(string part, string xml)
                    {
                        using var writer = new StreamWriter(archive.CreateEntry(part).Open());
                        writer.Write(xml);
                    }
                    const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    Part("xl/vbaProject.bin", "synthetic; never execute");
                    string compatibility = unboundPrefix ? " xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='missing'" : "";
                    Part("xl/workbook.xml", $"<workbook xmlns='{ns}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'{compatibility}><sheets><sheet name='Probe' r:id='p'/><sheet name='Result' r:id='r'/></sheets><calcPr calcMode='{mode}'/></workbook>");
                    Part("xl/_rels/workbook.xml.rels", "<Relationships><Relationship Id='p' Target='worksheets/input.xml'/><Relationship Id='r' Target='worksheets/output.xml'/></Relationships>");
                    Part("xl/worksheets/input.xml", $"<worksheet xmlns='{ns}'><sheetData><row><c r='B2'><v>21</v></c><c r='C2'><f>B2*2</f><v>{cached}</v></c></row></sheetData></worksheet>");
                    Part("xl/worksheets/output.xml", $"<worksheet xmlns='{ns}'><sheetData><row><c r='B2'><f>Probe!C2+1</f><v>43</v></c></row></sheetData></worksheet>");
                    if (external) Part("xl/connections.xml", "<connections/>");
                }
                CalculationProbe.VerifyPackage(path, 21, "manual");
            }
            Verify("valid", "manual", 42, false);
            Expect<InvalidDataException>(() => Verify("stale", "manual", 2, false));
            Expect<InvalidDataException>(() => Verify("mode", "auto", 42, false));
            Expect<InvalidDataException>(() => Verify("external", "manual", 42, true));
            Expect<InvalidDataException>(() => Verify("unbound-prefix", "manual", 42, false, true));
            Expect<FileNotFoundException>(() => CalculationProbe.RequirePrevious(repository, s14Directory, "manual"));
            File.WriteAllText(Path.Combine(s14Directory, "s14-auto.attempt"), "run");
            File.WriteAllText(Path.Combine(s14Directory, "s14-auto.success"), "other");
            Expect<InvalidDataException>(() => CalculationProbe.RequirePrevious(repository, s14Directory, "manual"));
            File.WriteAllText(Path.Combine(s14Directory, "s14-auto.success"), "run");
            CalculationProbe.RequirePrevious(repository, s14Directory, "manual");
        });
        Check("saved rich errors resolve metadata references and positional keys", () =>
        {
            var parts = new Dictionary<string, System.Xml.Linq.XDocument>();
            const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string rich = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
            parts["xl/metadata.xml"] = System.Xml.Linq.XDocument.Parse(
                $"<metadata xmlns='{ns}' xmlns:r='{rich}'><metadataTypes><metadataType name='XLRICHVALUE'/></metadataTypes>" +
                "<valueMetadata><bk><rc t='1' v='0'/></bk><bk><rc t='1' v='1'/></bk></valueMetadata>" +
                "<futureMetadata name='XLRICHVALUE'><bk><r:rvb i='0'/></bk><bk><r:rvb i='1'/></bk></futureMetadata></metadata>");
            parts["xl/richData/rdrichvalue.xml"] = System.Xml.Linq.XDocument.Parse(
                $"<rvData xmlns='{rich}'><rv s='0'><v>1</v><v>8</v></rv><rv s='0'><v>3</v><v>13</v></rv></rvData>");
            parts["xl/richData/rdrichvaluestructure.xml"] = System.Xml.Linq.XDocument.Parse(
                $"<rvStructures xmlns='{rich}'><s t='_error'><k n='subType' t='i'/><k n='errorType' t='i'/></s></rvStructures>");
            System.Xml.Linq.XDocument Read(string part) => parts.TryGetValue(part, out var xml) ? xml :
                throw new InvalidDataException("Missing synthetic part.");
            var cell = System.Xml.Linq.XElement.Parse($"<c xmlns='{ns}' t='e' vm='1'><v>#VALUE!</v></c>");
            Assert(SavedCellError.Read(cell, Read) == "#SPILL!");
            cell.SetAttributeValue("vm", "2");
            Assert(SavedCellError.Read(cell, Read) == "#CALC!");
            cell.SetAttributeValue("vm", "0");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            cell.SetAttributeValue("vm", "3");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            cell.SetAttributeValue("vm", "x");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            cell.SetAttributeValue("vm", "1");
            var structure = parts["xl/richData/rdrichvaluestructure.xml"].Root!.Elements().Single();
            structure.SetAttributeValue("t", "unknown");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            structure.SetAttributeValue("t", "_error");
            structure.Elements().Last().SetAttributeValue("n", "wrongKey");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            structure.Elements().Last().SetAttributeValue("n", "errorType");
            var values = parts["xl/richData/rdrichvalue.xml"].Root!.Elements().First();
            values.Elements().Last().Value = "999";
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            values.Elements().Last().Value = "8";
            cell.Element(System.Xml.Linq.XName.Get("v", ns))!.Value = "#SPILL!";
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            cell.Element(System.Xml.Linq.XName.Get("v", ns))!.Value = "#VALUE!";
            values.Elements().Last().Remove();
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            parts.Remove("xl/richData/rdrichvalue.xml");
            Expect<InvalidDataException>(() => SavedCellError.Read(cell, Read));
            cell.SetAttributeValue("vm", null);
            Assert(SavedCellError.Read(cell, Read) == "#VALUE!");
        });
        Check("cell fixed selection and ordering", () =>
        {
            Assert(CellProbe.Parse(["--excel-s1"]) == "s1" && CellProbe.Previous("s1") is null);
            Assert(CellProbe.Parse(["--excel-s4"]) == "s4" && CellProbe.Previous("s4") == "s3");
            Assert(CellProbe.Parse(["--excel-s1", "other.xlsx"]) is null);
            Assert(CellProbe.Parse(["--excel-s5"]) is null);
            Expect<InvalidDataException>(() => CellProbe.Previous("s5"));
        });
        Check("cell errors preserve unknown raw and distinguish numbers and strings", () =>
        {
            int unknown = CellProbe.ErrorRaw(2999);
            string json = JsonSerializer.Serialize(CellProbe.Project(unknown));
            using var document = JsonDocument.Parse(json);
            Assert(document.RootElement.GetProperty("code").ValueKind == JsonValueKind.Null);
            Assert(document.RootElement.GetProperty("raw").GetInt32() == unknown);
            Assert(CellProbe.Project(2007d) is double);
            Assert(CellProbe.Project((double)CellProbe.ErrorRaw(2007)) is double);
            Assert(CellProbe.Project("#DIV/0!") is string);
            using var known = JsonDocument.Parse(JsonSerializer.Serialize(CellProbe.Project(CellProbe.ErrorRaw(2045))));
            Assert(known.RootElement.GetProperty("code").GetString() == "#SPILL!");
        });
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
        Check("cell previous evidence and repeat rejection preserve files", () =>
        {
            Expect<InvalidDataException>(() => CellProbe.RequirePrevious(repository, directory, "s2"));
            string attempt = Path.Combine(directory, "s1.attempt");
            string success = Path.Combine(directory, "s1.success");
            string guard = Path.Combine(directory, "cell.guard");
            File.WriteAllText(attempt, "original run");
            File.WriteAllText(success, "different run");
            File.WriteAllText(guard, "existing guard");
            Expect<IOException>(() => StartupProbe.RequireNewAttempt(guard));
            Expect<InvalidDataException>(() => CellProbe.RequirePrevious(repository, directory, "s2"));
            Expect<IOException>(() => StartupProbe.RequireNewAttempt(attempt));
            Assert(File.ReadAllText(attempt) == "original run" && File.ReadAllText(guard) == "existing guard");
            File.WriteAllText(success, "original run");
            CellProbe.RequirePrevious(repository, directory, "s2");
        });
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
                    Expect<SessionRecordInvalidException>(() => file.CheckPrevious());
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
            Expect<InvalidOperationException>(() => file.CheckPrevious());
            Assert(file.Read() == existing);
            file.Write(existing with { Process = existing.Process! with { Name = "different" } });
            Expect<InvalidOperationException>(() => file.CheckPrevious());
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
        S11Checks.Run(directory, Check);
        S9Checks.Run(directory, Check);
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
