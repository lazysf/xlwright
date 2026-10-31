using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Spike0;

internal static class S9ComCheck
{
    private sealed record ClientObservation(int Calls, int Hresult, bool Continued, ExecutionSnapshot State, object Filter);

    internal static async Task<int> Run(string repository)
    {
        string guardPath = Path.Combine(repository, ".xlwright", "spike0-com-checks", "active.json");
        RuntimeAssets.RejectReparsePoints(repository, guardPath);
        StartupProbe.RequireNewAttempt(guardPath);
        string directory = RuntimeAssets.CreateLocalDirectory(repository,
            ".xlwright/spike0-com-checks/" + Guid.NewGuid().ToString("N"));
        using (var guard = new FileStream(guardPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
        {
            JsonSerializer.Serialize(guard, new { directory, build = S9ComReceiver.Build });
            guard.Flush(true);
        }
        Console.WriteLine("Controlled COM evidence: " + Path.GetRelativePath(repository, directory));
        foreach (string scenario in S9ComReceiver.Cases)
        {
            if (!await RunCase(directory, scenario)) return 1;
            Console.WriteLine("PASS controlled COM " + scenario);
        }
        Write(Path.Combine(directory, "completed.json"), new { cases = S9ComReceiver.Cases, passed = true });
        RuntimeAssets.RejectReparsePoints(repository, guardPath);
        using (var guard = JsonDocument.Parse(File.ReadAllText(guardPath)))
            if (guard.RootElement.GetProperty("directory").GetString() != directory)
                throw new InvalidDataException("Controlled COM guard changed; preserve it.");
        File.Delete(guardPath);
        Console.WriteLine("7 controlled COM cases passed. Excel was not activated.");
        return 0;
    }

    private static async Task<bool> RunCase(string directory, string scenario)
    {
        string nonce = Guid.NewGuid().ToString("N");
        string host = Environment.ProcessPath ?? throw new InvalidOperationException("Missing host.");
        if (!Path.GetFileName(host).Equals("dotnet.exe", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Use dotnet.exe with the built Spike0 DLL.");
        var start = new ProcessStartInfo(host)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        start.ArgumentList.Add(typeof(S9ComCheck).Assembly.Location);
        start.ArgumentList.Add("--s9-com-receiver");
        start.ArgumentList.Add(scenario);
        start.ArgumentList.Add(nonce);
        using var receiver = Process.Start(start) ?? throw new InvalidOperationException("Receiver did not start.");
        string prefix = Path.Combine(directory, scenario);
        Task<string>? stderr = null;
        var gate = new ExecutionGate();
        ClientObservation? client = null;
        ExecutionSnapshot? timeoutObservation = null;
        Task<bool>? work = null;
        try
        {
            stderr = receiver.StandardError.ReadToEndAsync();
            using var controller = Process.GetCurrentProcess();
            Write(prefix + ".process.json", new { nonce, scenario, receiver.Id,
                startedAtUtc = receiver.StartTime.ToUniversalTime(), build = S9ComReceiver.Build,
                runtime = Environment.Version.ToString(), architecture = RuntimeInformation.ProcessArchitecture.ToString(),
                controller = new { controller.Id, startedAtUtc = controller.StartTime.ToUniversalTime() } });
            string readyLine = await receiver.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10))
                ?? throw new InvalidDataException("Receiver did not provide its interface.");
            if (readyLine.Length > 100000) throw new InvalidDataException("Receiver reply too large.");
            var ready = JsonSerializer.Deserialize<S9ComReady>(readyLine)
                ?? throw new InvalidDataException("Invalid receiver reply.");
            if (ready.Nonce != nonce || ready.Scenario != scenario || ready.Pid != receiver.Id || ready.Build != S9ComReceiver.Build)
                throw new InvalidDataException("Receiver identity mismatch.");
            // Import only the reference received directly from our own fixed child, never a file or user argument.
            work = StaRunner.Start(() =>
            {
                var target = S9ComMarshalling.Import(ready.Reference);
                using var filter = new RetryFilter(gate);
                try
                {
                    filter.Register();
                    gate.Begin("controlled-commit", TimeSpan.FromMilliseconds(scenario == "late" ? 100 : scenario == "deadline" ? 2000 : 30000));
                    int calls = 0;
                    int hresult = 0;
                    bool continued = false;
                    filter.BeginCall();
                    try { gate.Step("commit", () => { calls++; target.Commit(0); }); }
                    catch (Exception error) { hresult = error.HResult; }
                    finally { filter.EndCall(hresult); }
                    if (scenario == "deadline" && filter.DeadlineRefusal)
                        while (!gate.CheckTimeout()) Thread.Sleep(1);
                    if (hresult == 0)
                    {
                        try
                        {
                            // A managed marker represents normal continuation; it performs no COM or file I/O.
                            gate.Step("next-stage-marker", () => continued = true);
                            gate.Complete("ready");
                        }
                        catch (WorkStoppedException) { gate.Fault(true); }
                    }
                    else gate.Fault(true);
                    filter.Dispose();
                    client = new(calls, hresult, continued, gate.Snapshot, filter.Snapshot());
                    return true;
                }
                finally { Marshal.FinalReleaseComObject(target); }
            });
            var observationLimit = Stopwatch.StartNew();
            while (!work.IsCompleted)
            {
                if (gate.CheckTimeout()) timeoutObservation ??= gate.Snapshot;
                if (observationLimit.Elapsed >= TimeSpan.FromSeconds(45))
                    throw new TimeoutException("Controlled COM call has not unwound; do not retry.");
                await Task.WhenAny(work, Task.Delay(10));
            }
            await work;
            Write(prefix + ".client.json", new { client, timeoutObservation });
            await receiver.StandardInput.WriteLineAsync("stop");
            receiver.StandardInput.Close();
            string resultLine = await receiver.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10))
                ?? throw new InvalidDataException("Receiver result missing.");
            var result = JsonSerializer.Deserialize<S9ComReceived>(resultLine)
                ?? throw new InvalidDataException("Invalid receiver result.");
            await receiver.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Write(prefix + ".receiver.json", result);
            File.WriteAllText(prefix + ".stderr.log", await stderr);
            if (result.Nonce != nonce || result.Scenario != scenario || !result.Restored || receiver.ExitCode != 0 || client is null)
                throw new InvalidDataException("Incomplete receiver evidence.");
            using var filterJson = JsonDocument.Parse(JsonSerializer.Serialize(client.Filter));
            var decisions = filterJson.RootElement.GetProperty("decisions").EnumerateArray().ToArray();
            bool hadRetry = decisions.Any(d => d.GetProperty("Delay").GetInt32() == 100);
            bool onlyRetryLater = decisions.All(d => d.GetProperty("RejectType").GetInt32() == 2);
            bool expectedReady = scenario is "accept" or "retry";
            bool pass = client.Calls == 1 && filterJson.RootElement.GetProperty("Restored").GetBoolean() &&
                !filterJson.RootElement.GetProperty("recordingFailed").GetBoolean() &&
                !client.State.Busy && client.State.State == (expectedReady ? "ready" : "isolated") &&
                result.Incoming == result.Rejected + result.BodyCalls && (scenario switch
            {
                "accept" => client.Hresult == 0 && client.Continued && result.BodyCalls == 1 && decisions.Length == 0,
                "retry" => client.Hresult == 0 && client.Continued && result.BodyCalls == 1 && result.Rejected == 2 && hadRetry && onlyRetryLater,
                "reject" => client.Hresult != 0 && !client.Continued && result.BodyCalls == 0 && result.Rejected == 1 && !hadRetry &&
                    decisions.Length == 1 && decisions[0].GetProperty("RejectType").GetInt32() == 1,
                "limit" => client.Hresult != 0 && !client.Continued && result.BodyCalls == 0 && hadRetry && onlyRetryLater &&
                    decisions.Any(d => d.GetProperty("Reason").GetString() == "retry-budget") && client.State.StopReason is null,
                "deadline" => client.Hresult != 0 && !client.Continued && result.BodyCalls == 0 && hadRetry && onlyRetryLater && client.State.StopReason == "timeout",
                "failure" => client.Hresult == unchecked((int)0x80004005) && !client.Continued && result.BodyCalls == 1 && decisions.Length == 0,
                "late" => client.Hresult == 0 && !client.Continued && result.BodyCalls == 1 && decisions.Length == 0 &&
                    client.State.State == "isolated" && timeoutObservation is { Busy: true },
                _ => false
            });
            Write(prefix + ".assessment.json", new { scenario, passed = pass, note = "Expected failures are test observations, not successful COM operations." });
            return pass;
        }
        catch (Exception error)
        {
            Write(prefix + ".unknown.json", new { scenario, error = error.GetType().Name, error.Message,
                snapshot = gate.Snapshot, workerUnwound = work?.IsCompleted });
            Console.Error.WriteLine("Controlled COM check stopped: " + scenario + "; preserve " + Path.GetFileName(directory));
            return false;
        }
        finally
        {
            // EOF requests normal helper exit. It neither kills the helper nor cancels a pending COM call.
            receiver.StandardInput.Close();
        }
    }

    private static void Write(string path, object value)
    {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        JsonSerializer.Serialize(stream, value);
        stream.Flush(true);
    }
}
