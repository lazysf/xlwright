using System.Text.Json;

namespace Spike0;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        try
        {
            string repository = RuntimeAssets.FindRepository();
            if (args.Length == 0 || args is ["--self-check"])
                return await SelfChecks.Run(repository);
            string? startupCase = StartupProbe.Parse(args);
            if (startupCase is null && args is not ["--excel"] and not ["--excel-end-refusal"] and not ["--excel-record-fault-starting"]
                and not ["--excel-record-fault-owned"] and not ["--excel-record-fault-completed"])
            {
                Console.Error.WriteLine("Usage: Spike0 [--self-check | --excel | --excel-end-refusal | " +
                    "--excel-record-fault-starting | --excel-record-fault-owned | --excel-record-fault-completed]. " +
                    "S10: --excel-s10-a2 | --excel-s10-b | --excel-s10-c. " +
                    "Excel experiments require prior approval.");
                return 1;
            }
            string? recordFault = args[0] switch
            {
                "--excel-record-fault-starting" => "starting",
                "--excel-record-fault-owned" => "owned",
                "--excel-record-fault-completed" => "completed",
                _ => null
            };
            return await RunExcel(repository, args[0] == "--excel-end-refusal", recordFault, startupCase);
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"{error.GetType().Name}: {error.Message}");
            return 1;
        }
    }

    private static async Task<int> RunExcel(string repository, bool endRefusal, string? recordFault, string? startupCase)
    {
        RuntimeAssets.VerifyPias();
        RuntimeAssets.VerifyFixture(repository);
        var startupFixture = startupCase is null ? null : StartupProbe.VerifyFixture(repository);
        string guardPath = Path.Combine(repository, ".xlwright", "mcp-development", "active-smoke.json");
        string? attemptPath = null;
        if (startupCase is not null)
        {
            RuntimeAssets.RejectReparsePoints(repository, guardPath);
            if (File.Exists(guardPath)) throw new InvalidOperationException("Existing MCP guard requires review.");
            attemptPath = Path.Combine(repository, ".xlwright", "spike0", "s10-" + startupCase + ".attempt");
            RuntimeAssets.RejectReparsePoints(repository, attemptPath);
            StartupProbe.RequireNewAttempt(attemptPath);
        }
        string workspace = RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/spike0");
        string recordPath = Path.Combine(workspace, ".xlwright.lock");
        RuntimeAssets.RejectReparsePoints(repository, recordPath);
        using var record = OwnershipRecordFile.Acquire(recordPath);
        // Recheck under the shared record lock before any record transition or new guard.
        if (attemptPath is not null) StartupProbe.RequireNewAttempt(attemptPath);
        record.CheckPrevious();
        string? previous = startupCase is null ? null : StartupProbe.PreviousCase(startupCase);
        if (previous is not null)
        {
            string success = Path.Combine(workspace, "s10-" + previous + ".success");
            RuntimeAssets.RejectReparsePoints(repository, success);
            string attempt = Path.Combine(workspace, "s10-" + previous + ".attempt");
            RuntimeAssets.RejectReparsePoints(repository, attempt);
            if (!File.Exists(success) || !File.Exists(attempt) || new FileInfo(success).Length == 0 ||
                File.ReadAllText(success) != File.ReadAllText(attempt))
                throw new InvalidOperationException("Previous S10 case has not succeeded.");
        }
        // Installed after the previous record check so only the run's own transition can fail.
        if (recordFault is not null) record.InjectWriteFailure(recordFault);
        string run = RuntimeAssets.CreateLocalDirectory(repository,
            Path.Combine(".xlwright", "spike0", "runs", Guid.NewGuid().ToString("N")));
        if (startupCase is not null)
        {
            RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/mcp-development");
            using (var guard = new FileStream(guardPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                JsonSerializer.Serialize(guard, new { scenario = "s10-" + startupCase, run });
                guard.Flush(true);
            }
            using var attempt = new FileStream(attemptPath!, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
            using var writer = new StreamWriter(attempt);
            writer.WriteLine(run);
        }
        string workbook = Path.Combine(run, startupCase is null ? "lifecycle.xlsx" : "probe.xlsm");
        File.Copy(startupFixture?.Path ?? Path.Combine(repository, "tests", "fixtures", "lifecycle.xlsx"), workbook, overwrite: false);
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
        log.Write("run.prepared", gate, new { workbook_path = workbook, record_path = recordPath, diagnostic_path = diagnosticPath,
            server_pid = Environment.ProcessId });
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
            var scenario = new ExcelScenario(workbook, record, gate, log, unmanaged, recordFault, startupCase);
            Task<bool> work = StaRunner.Start(scenario.Run);
            while (!work.IsCompleted)
            {
                if (gate.CheckTimeout())
                {
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
                if (startupCase is null) RuntimeAssets.VerifyCell(workbook, 1);
                else StartupProbe.VerifyHash(startupFixture!.Path, startupFixture.Hash);
                RuntimeAssets.VerifyFixture(repository);
                log.Write(startupCase is null ? "saved-file.verified" : "s10.completed", gate);
                if (startupCase is not null)
                {
                    string success = Path.Combine(workspace, "s10-" + startupCase + ".success");
                    RuntimeAssets.RejectReparsePoints(repository, success);
                    using (var stream = new FileStream(success, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    {
                        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(run + Environment.NewLine);
                        stream.Write(bytes);
                        stream.Flush(true);
                    }
                    if (File.ReadAllText(success) != File.ReadAllText(Path.Combine(workspace, "s10-" + startupCase + ".attempt")))
                        throw new InvalidDataException("S10 success evidence did not read back.");
                    RuntimeAssets.RejectReparsePoints(repository, guardPath);
                    using (var guard = JsonDocument.Parse(File.ReadAllText(guardPath)))
                        if (guard.RootElement.GetProperty("run").GetString() != run ||
                            guard.RootElement.GetProperty("scenario").GetString() != "s10-" + startupCase)
                            throw new InvalidDataException("S10 guard changed; preserve it for review.");
                    File.Delete(guardPath);
                }
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
