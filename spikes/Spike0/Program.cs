using System.Text.Json;

namespace Spike0;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        try
        {
            string repository = RuntimeAssets.FindRepository();
            if (args is ["--s9-com-check"])
                return await S9ComCheck.Run(repository);
            if (args is ["--s9-com-receiver", var receiverCase, var nonce])
                return await S9ComReceiver.Run(receiverCase, nonce);
            if (args.Length == 0 || args is ["--self-check"])
                return await SelfChecks.Run(repository);
            if (args is ["--verify-s4-saved"])
            {
                CellProbe.RecheckS4(repository);
                return 0;
            }
            bool preflightOnly = args is ["--preflight", _];
            if (preflightOnly) args = args[1..];
            string? startupCase = StartupProbe.Parse(args);
            string? s9Case = S9Probe.Parse(args);
            string? cellCase = CellProbe.Parse(args);
            string? calculationCase = CalculationProbe.Parse(args);
            if (s9Case is null && startupCase is null && cellCase is null && calculationCase is null && args is not ["--excel"] and not ["--excel-end-refusal"] and not ["--excel-record-fault-starting"]
                and not ["--excel-record-fault-owned"] and not ["--excel-record-fault-completed"])
            {
                Console.Error.WriteLine("Usage: Spike0 [--self-check | --excel | --excel-end-refusal | " +
                    "--excel-record-fault-starting | --excel-record-fault-owned | --excel-record-fault-completed]. " +
                    "S10: --excel-s10-a2 | --excel-s10-b | --excel-s10-c. " +
                    "Cells: --excel-s1 | --excel-s2 | --excel-s3 | --excel-s4. " +
                    "Calculation: --excel-s14-auto | --excel-s14-manual | --excel-s14-mode-change. " +
                    "S9: --excel-s9-release | --excel-s9-limit | --excel-s9-deadline. " +
                    "Controlled COM without Excel: --s9-com-check. " +
                    "Offline: --verify-s4-saved. Checks only, without Excel: --preflight <Excel argument>. " +
                    "Excel experiments require prior approval.");
                return 1;
            }
            if (preflightOnly)
            {
                Preflight(repository);
                Console.WriteLine("Preflight passed. Excel was not started.");
                return 0;
            }
            string? recordFault = args[0] switch
            {
                "--excel-record-fault-starting" => "starting",
                "--excel-record-fault-owned" => "owned",
                "--excel-record-fault-completed" => "completed",
                _ => null
            };
            return await RunExcel(repository, args[0] == "--excel-end-refusal", recordFault, startupCase, cellCase, calculationCase, s9Case);
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"{error.GetType().Name}: {error.Message}");
            return 1;
        }
    }

    private static void RequireNoGuard(string guardPath)
    {
        try { _ = File.GetAttributes(guardPath); }
        catch (FileNotFoundException) { return; }
        catch (DirectoryNotFoundException) { return; }
        // Unknown I/O failures propagate; only a confirmed missing guard permits a start.
        throw new InvalidOperationException("An experiment guard remains; manual review is required. No Excel was started.");
    }

    // Read-only checks for every Excel entry. --preflight stops after them, so automated tests cannot reach Excel.
    private static (PreparedBuild.Identity Build, string GuardPath) Preflight(string repository)
    {
        var build = PreparedBuild.Require(repository, AppContext.BaseDirectory);
        string guardPath = Path.Combine(repository, ".xlwright", "mcp-development", "active-smoke.json");
        RuntimeAssets.RejectReparsePoints(repository, guardPath);
        RequireNoGuard(guardPath);
        return (build, guardPath);
    }

    private static async Task<int> RunExcel(string repository, bool endRefusal, string? recordFault, string? startupCase, string? cellCase, string? calculationCase, string? s9Case)
    {
        var (build, guardPath) = Preflight(repository);
        RuntimeAssets.VerifyPias();
        RuntimeAssets.VerifyFixture(repository);
        var startupFixture = startupCase is null ? null : StartupProbe.VerifyFixture(repository);
        var calculationFixture = calculationCase is null ? null : CalculationProbe.VerifyFixture(repository, calculationCase);
        string? experiment = s9Case is not null ? "s9-" + s9Case : calculationCase is null ? cellCase ?? (startupCase is null ? null : "s10-" + startupCase) : "s14-" + calculationCase;
        string scenarioName = experiment ?? (endRefusal ? "end-refusal" : recordFault is null ? "smoke" : "record-fault-" + recordFault);
        string? attemptPath = null;
        if (experiment is not null)
        {
            attemptPath = Path.Combine(repository, ".xlwright", "spike0", experiment + ".attempt");
            RuntimeAssets.RejectReparsePoints(repository, attemptPath);
            StartupProbe.RequireNewAttempt(attemptPath);
        }
        string workspace = RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/spike0");
        string recordPath = Path.Combine(workspace, ".xlwright.lock");
        RuntimeAssets.RejectReparsePoints(repository, recordPath);
        using var record = OwnershipRecordFile.Acquire(recordPath);
        // Recheck under the shared record lock before any record transition or new guard.
        // The guard also keeps a record awaiting manual recovery from the automatic next-start transition.
        RequireNoGuard(guardPath);
        if (attemptPath is not null) StartupProbe.RequireNewAttempt(attemptPath);
        if (s9Case is not null)
        {
            if (record.CreatedNew || record.Read().Kind != "completed")
                throw new InvalidDataException("S9 requires existing completed ownership.");
            S9Probe.RequirePrevious(repository, workspace, s9Case);
        }
        if (calculationCase is not null)
        {
            if (record.CreatedNew || record.Read().Kind != "completed")
                throw new InvalidDataException("T9 requires existing completed ownership.");
            CalculationProbe.RequirePrevious(repository, workspace, calculationCase);
        }
        if (cellCase is not null)
        {
            if (!record.CreatedNew && record.Read().Kind != "completed")
                throw new InvalidDataException("T8 requires completed ownership; preserve the existing record.");
            CellProbe.RequirePrevious(repository, workspace, cellCase);
        }
        var replacedRecord = record.CheckPrevious();
        string? previousStartup = startupCase is null ? null : StartupProbe.PreviousCase(startupCase);
        string? previous = previousStartup is null ? null : "s10-" + previousStartup;
        if (previous is not null)
        {
            string success = Path.Combine(workspace, previous + ".success");
            RuntimeAssets.RejectReparsePoints(repository, success);
            string attempt = Path.Combine(workspace, previous + ".attempt");
            RuntimeAssets.RejectReparsePoints(repository, attempt);
            if (!File.Exists(success) || !File.Exists(attempt) || new FileInfo(success).Length == 0 ||
                File.ReadAllText(success) != File.ReadAllText(attempt))
                throw new InvalidOperationException("Previous S10 case has not succeeded.");
        }
        // Installed after the previous record check so only the run's own transition can fail.
        if (recordFault is not null) record.InjectWriteFailure(recordFault);
        string run = RuntimeAssets.CreateLocalDirectory(repository,
            Path.Combine(".xlwright", "spike0", "runs", Guid.NewGuid().ToString("N")));
        RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/mcp-development");
        using (var guard = new FileStream(guardPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
        {
            JsonSerializer.Serialize(guard, new { scenario = scenarioName, run });
            guard.Flush(true);
        }
        if (experiment is not null)
        {
            using var attempt = new FileStream(attemptPath!, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
            using var writer = new StreamWriter(attempt);
            writer.WriteLine(run);
            writer.Flush();
            attempt.Flush(true);
        }
        string workbook = Path.Combine(run, startupCase is null && calculationCase is null ? "lifecycle.xlsx" : "probe.xlsm");
        File.Copy(calculationFixture?.Path ?? startupFixture?.Path ?? Path.Combine(repository, "tests", "fixtures", "lifecycle.xlsx"), workbook, overwrite: false);
        if (calculationFixture is not null) StartupProbe.VerifyHash(workbook, calculationFixture.Hash);
        if (cellCase is not null) StartupProbe.VerifyHash(workbook, RuntimeAssets.FixtureHash);
        if (startupFixture is not null)
        {
            StartupProbe.VerifyHash(workbook, startupFixture.Hash);
        }
        if (startupCase is not null) File.WriteAllText(Path.Combine(run, "s10-vba.log"), "");
        string? unmanaged = endRefusal ? Path.Combine(run, "unmanaged.xlsx") : null;
        if (unmanaged is not null)
            File.Copy(Path.Combine(repository, "tests", "fixtures", "lifecycle.xlsx"), unmanaged, overwrite: false);
        string diagnosticPath = Path.Combine(run, "diagnostics.jsonl");
        using var log = new DiagnosticLog(diagnosticPath);
        var gate = new ExecutionGate();
        using var s9 = s9Case is null ? null : new S9Probe(s9Case, run, gate);
        log.Write("run.prepared", gate, new { workbook_path = workbook, record_path = recordPath, diagnostic_path = diagnosticPath,
            server_pid = Environment.ProcessId, assembly_hash = build.Assembly, check_run = build.CheckRun,
            replaced_record = replacedRecord });
        ConsoleCancelEventHandler onCancel = (_, e) =>
        {
            e.Cancel = true;
            gate.RequestShutdown();
        };
        Console.CancelKeyPress += onCancel;
        try
        {
            // The console monitor and deadline never own or access a COM object.
            _ = Task.Run(async () =>
            {
                try
                {
                    while (true)
                    {
                        string? line = await Console.In.ReadLineAsync();
                        if (line == "go") s9?.SignalStart();
                        if (line is null || line == "stop")
                        {
                            gate.RequestShutdown();
                            return;
                        }
                        if (line == "status") Console.WriteLine(JsonSerializer.Serialize(gate.Snapshot));
                    }
                }
                catch (Exception error)
                {
                    gate.RequestShutdown();
                    Console.Error.WriteLine("Console monitor failed: " + error.GetType().Name);
                }
            });
            if (cellCase is not null) CellProbe.WriteIdentity(repository, run, cellCase);
            if (calculationCase is not null) CalculationProbe.WriteIdentity(repository, run, calculationFixture!);
            var scenario = new ExcelScenario(workbook, record, gate, log, unmanaged, recordFault, startupCase, cellCase, calculationCase, s9);
            Task<bool> work = StaRunner.Start(scenario.Run);
            while (!work.IsCompleted)
            {
                s9?.NotifyRelease();
                if (gate.CheckTimeout())
                {
                    s9?.SaveEvidence("timeout");
                    log.Write("operation.result_unknown", gate, new
                    {
                        record_path = recordPath,
                        diagnostic_path = diagnosticPath,
                        server_pid = Environment.ProcessId,
                        manual_recovery = true,
                        note = "No further COM operation is allowed. Confirm this process has exited before manual Excel recovery."
                    });
                    return 2;
                }
                await Task.WhenAny(work, Task.Delay(50));
            }
            try
            {
                bool smokePassed = await work;
                if (!smokePassed) return 1;
                if (calculationCase is not null)
                {
                    CalculationProbe.VerifySaved(workbook, calculationCase);
                    StartupProbe.VerifyHash(calculationFixture!.Path, calculationFixture.Hash);
                }
                else if (cellCase is not null) CellProbe.VerifySaved(workbook);
                else if (startupCase is null) RuntimeAssets.VerifyCell(workbook, 1);
                else StartupProbe.VerifyHash(startupFixture!.Path, startupFixture.Hash);
                RuntimeAssets.VerifyFixture(repository);
                log.Write(calculationCase is not null ? "s14.completed" : cellCase is not null ? "cell.completed" : startupCase is null ? "saved-file.verified" : "s10.completed", gate);
                if (experiment is not null)
                {
                    string success = Path.Combine(workspace, experiment + ".success");
                    RuntimeAssets.RejectReparsePoints(repository, success);
                    using (var stream = new FileStream(success, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    {
                        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(run + Environment.NewLine);
                        stream.Write(bytes);
                        stream.Flush(true);
                    }
                    if (File.ReadAllText(success) != File.ReadAllText(attemptPath!))
                        throw new InvalidDataException("S10 success evidence did not read back.");
                }
                RuntimeAssets.RejectReparsePoints(repository, guardPath);
                using (var guard = JsonDocument.Parse(File.ReadAllText(guardPath)))
                    if (guard.RootElement.GetProperty("run").GetString() != run ||
                        guard.RootElement.GetProperty("scenario").GetString() != scenarioName)
                        throw new InvalidDataException("Experiment guard changed; preserve it for review.");
                File.Delete(guardPath);
                return 0;
            }
            catch (Exception error)
            {
                log.Write("run.failed", gate, new { error = error.GetType().Name,
                    error.Message, record_path = recordPath, diagnostic_path = diagnosticPath,
                    server_pid = Environment.ProcessId });
                return gate.Snapshot.StopReason == "timeout" ? 2 : 1;
            }
        }
        finally
        {
            gate.RequestShutdown();
            Console.CancelKeyPress -= onCancel;
        }
    }
}
