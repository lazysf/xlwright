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
            if (args is not ["--excel"] and not ["--excel-end-refusal"])
            {
                Console.Error.WriteLine("Usage: Spike0 [--self-check | --excel | --excel-end-refusal]. Excel experiments require prior approval.");
                return 1;
            }
            return await RunExcel(repository, args[0] == "--excel-end-refusal");
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"{error.GetType().Name}: {error.Message}");
            return 1;
        }
    }

    private static async Task<int> RunExcel(string repository, bool endRefusal)
    {
        RuntimeAssets.VerifyPias();
        RuntimeAssets.VerifyFixture(repository);
        string workspace = RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/spike0");
        string recordPath = Path.Combine(workspace, ".xlwright.lock");
        RuntimeAssets.RejectReparsePoints(repository, recordPath);
        using var record = OwnershipRecordFile.Acquire(recordPath);
        record.CheckPrevious();
        string run = RuntimeAssets.CreateLocalDirectory(repository,
            Path.Combine(".xlwright", "spike0", "runs", Guid.NewGuid().ToString("N")));
        string workbook = Path.Combine(run, "lifecycle.xlsx");
        File.Copy(Path.Combine(repository, "tests", "fixtures", "lifecycle.xlsx"), workbook, overwrite: false);
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
            var scenario = new ExcelScenario(workbook, record, gate, log, unmanaged);
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
                RuntimeAssets.VerifyCell(workbook, 1);
                RuntimeAssets.VerifyFixture(repository);
                log.Write("saved-file.verified", gate);
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
