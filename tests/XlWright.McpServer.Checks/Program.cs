using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using XlWright.McpServer;

if (args is ["--excel"] or ["--excel-end-refusal"] or ["--excel-record-fault-starting"] or ["--excel-record-fault-completed"])
{
    string behavior = File.ReadAllText("fake-behavior.txt");
    if (behavior.StartsWith("fault-", StringComparison.Ordinal))
    {
        if (args[0] != "--excel-record-" + behavior) return 3;
        Console.WriteLine("{\"observation\":\"record-fault.verified\"}");
        Console.WriteLine("{\"observation\":\"run.failed\"}");
        return 1;
    }
    if (behavior.StartsWith("refusal", StringComparison.Ordinal))
    {
        if (args[0] != "--excel-end-refusal") return 3;
        if (behavior is "refusal-eof" or "refusal-result-io")
            while (await Console.In.ReadLineAsync() is not null) { }
        Console.WriteLine("{\"observation\":\"end-refusal.verified\"}");
        Console.WriteLine("{\"observation\":\"run.failed\"}");
        if (behavior == "refusal-zero") Console.WriteLine("{\"observation\":\"saved-file.verified\"}");
        return behavior == "refusal-zero" ? 0 : 1;
    }
    if (behavior is "eof" or "result-io")
    {
        while (await Console.In.ReadLineAsync() is not null) { }
    }
    if (behavior == "fail") return 1;
    Console.WriteLine("{\"observation\":\"saved-file.verified\"}");
    return 0;
}

string repository = Environment.CurrentDirectory;
string dotnet = Environment.ProcessPath!;
string server = Path.Combine(repository, "src/XlWright.McpServer/bin/Release/net10.0-windows/XlWright.McpServer.dll");
string checksDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
string testRoot = Path.Combine(repository, ".xlwright", "mcp-checks", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testRoot);
using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(60));
int passed = 0;
string blockedRepository = PrepareFake("blocked");

foreach (bool development in new[] { false, true })
{
    var eofLaunch = new ProcessStartInfo(dotnet)
    {
        UseShellExecute = false, CreateNoWindow = true,
        WorkingDirectory = development ? blockedRepository : testRoot,
        RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
    };
    eofLaunch.ArgumentList.Add(server);
    if (development) eofLaunch.ArgumentList.Add("--development");
    using (var eofProcess = Process.Start(eofLaunch)!)
    {
        var stdout = eofProcess.StandardOutput.ReadToEndAsync();
        var stderr = eofProcess.StandardError.ReadToEndAsync();
        eofProcess.StandardInput.Close();
        await eofProcess.WaitForExitAsync(deadline.Token).WaitAsync(TimeSpan.FromSeconds(10));
        Require(eofProcess.ExitCode == 0, "Server exits on EOF without forced termination: " + development);
        Require(await stdout == "", "No diagnostics on protocol stdout: " + development);
        await stderr;
    }
    var launch = new StdioClientTransport(new()
    {
        Name = "xlwright-check", Command = dotnet,
        Arguments = development ? [server, "--development"] : [server],
        WorkingDirectory = development ? blockedRepository : testRoot
    });
    await using var client = await McpClient.CreateAsync(launch, cancellationToken: deadline.Token);
    if (!development)
        Require(client.ServerCapabilities.Tools is null, "Normal mode advertises no development tools");
    if (development)
    {
        var tools = await client.ListToolsAsync(cancellationToken: deadline.Token);
        Require(tools.Count == 6 && tools.Any(tool => tool.Name == "dev_excel_end_refusal_start") &&
            tools.Any(tool => tool.Name == "dev_excel_record_fault_starting_start") &&
            tools.Any(tool => tool.Name == "dev_excel_record_fault_completed_start"), "Development tool list");
        var status = await client.CallToolAsync("dev_status", cancellationToken: deadline.Token);
        Require(status.IsError != true, "Status tool responds");
        var blocked = await client.CallToolAsync("dev_excel_smoke_start", cancellationToken: deadline.Token);
        Require(blocked.IsError == true, "Synthetic starting blocks Excel before launch");
        var refusalBlocked = await client.CallToolAsync("dev_excel_end_refusal_start", cancellationToken: deadline.Token);
        Require(refusalBlocked.IsError == true, "Synthetic starting blocks end-refusal before launch");
        foreach (string name in new[] { "dev_excel_record_fault_starting_start", "dev_excel_record_fault_completed_start" })
        {
            var faultBlocked = await client.CallToolAsync(name, cancellationToken: deadline.Token);
            Require(faultBlocked.IsError == true, "Synthetic starting blocks L2 before launch: " + name);
        }
    }
}

// Exercise the public JSON round trip, including the SDK's Guid binding.
var receiptLaunch = new StdioClientTransport(new()
{
    Name = "receipt-check", Command = dotnet, Arguments = [server, "--development"],
    WorkingDirectory = PrepareFake("receipt")
});
await using (var client = await McpClient.CreateAsync(receiptLaunch, cancellationToken: deadline.Token))
{
    var accepted = await client.CallToolAsync("dev_excel_smoke_start", cancellationToken: deadline.Token);
    using var snapshot = JsonDocument.Parse(accepted.Content.OfType<TextContentBlock>().Single().Text);
    string returnedReceipt = snapshot.RootElement.GetProperty("receipt").GetString()!;
    Require(returnedReceipt == Guid.Parse(returnedReceipt).ToString("D"), "Receipt uses canonical GUID form");
    var observed = await client.CallToolAsync("dev_excel_smoke_status",
        new Dictionary<string, object?> { ["receiptId"] = returnedReceipt }, cancellationToken: deadline.Token);
    Require(observed.IsError != true, "Returned receipt is accepted unchanged by the status tool");
}

foreach (string behavior in new[] { "success", "fail", "eof", "result-io", "refusal", "refusal-zero", "refusal-eof", "refusal-result-io",
    "fault-starting", "fault-completed" })
{
    string fake = PrepareFake(behavior);
    Environment.CurrentDirectory = fake;
    var adapter = new DevelopmentTools();
    await using (adapter)
    {
        bool refusal = behavior.StartsWith("refusal", StringComparison.Ordinal);
        bool fault = behavior.StartsWith("fault-", StringComparison.Ordinal);
        using var accepted = JsonDocument.Parse(behavior switch
        {
            "fault-starting" => adapter.StartRecordFaultStarting(),
            "fault-completed" => adapter.StartRecordFaultCompleted(),
            _ => refusal ? adapter.StartEndRefusal() : adapter.Start()
        });
        if (refusal)
        {
            string receipt = accepted.RootElement.GetProperty("receipt").GetString()!;
            using var observed = JsonDocument.Parse(adapter.SmokeStatus(Guid.Parse(receipt)));
            Require(observed.RootElement.GetProperty("scenario").GetString() == "end-refusal", "End-refusal receipt and scenario");
        }
        if (fault)
        {
            string receipt = accepted.RootElement.GetProperty("receipt").GetString()!;
            using var observed = JsonDocument.Parse(adapter.SmokeStatus(Guid.Parse(receipt)));
            Require(observed.RootElement.GetProperty("scenario").GetString() == "record-" + behavior, "L2 receipt and scenario: " + behavior);
        }
        if (behavior is "result-io" or "refusal-result-io")
        {
            string resultPath = Path.Combine(fake, accepted.RootElement.GetProperty("logDirectory").GetString()!,
                accepted.RootElement.GetProperty("receipt").GetString()! + ".result.json");
            Directory.CreateDirectory(resultPath);
        }
        else if (behavior is "eof" or "refusal-eof")
        {
            bool rejected = false;
            try { adapter.StartEndRefusal(); } catch (InvalidOperationException) { rejected = true; }
            Require(rejected, "Concurrent start rejected");
        }
        else
        {
            while (State(adapter) == "running") await Task.Delay(20, deadline.Token);
            Require(State(adapter) == (behavior == "success" ? "completed" : "failed"), "Exit result: " + behavior);
            if (refusal)
            {
                using var status = JsonDocument.Parse(adapter.Status());
                Require(status.RootElement.GetProperty("latestObservation").GetString() ==
                    (behavior == "refusal-zero" ? "saved-file.verified" : "run.failed") &&
                    status.RootElement.GetProperty("manualReviewRequired").GetBoolean(), "Refusal requires review even with exit zero");
                string stdoutPath = Path.Combine(fake, status.RootElement.GetProperty("logDirectory").GetString()!,
                    status.RootElement.GetProperty("receipt").GetString()! + ".stdout.jsonl");
                Require(File.ReadLines(stdoutPath).Any(line => line.Contains("end-refusal.verified", StringComparison.Ordinal)),
                    "Refusal evidence survives the terminal failure event");
                bool rejected = false;
                try { adapter.Start(); } catch (InvalidOperationException) { rejected = true; }
                Require(rejected, "Smoke rejected after end-refusal");
            }
            if (fault)
            {
                using var status = JsonDocument.Parse(adapter.Status());
                Require(status.RootElement.GetProperty("latestObservation").GetString() == "run.failed" &&
                    status.RootElement.GetProperty("manualReviewRequired").GetBoolean(), "L2 fault requires review: " + behavior);
                string stdoutPath = Path.Combine(fake, status.RootElement.GetProperty("logDirectory").GetString()!,
                    status.RootElement.GetProperty("receipt").GetString()! + ".stdout.jsonl");
                Require(File.ReadLines(stdoutPath).Any(line => line.Contains("record-fault.verified", StringComparison.Ordinal)),
                    "L2 fault evidence survives the terminal failure event: " + behavior);
                bool rejected = false;
                try { adapter.StartRecordFaultCompleted(); } catch (InvalidOperationException) { rejected = true; }
                Require(rejected, "L2 start rejected after a failed run: " + behavior);
            }
        }
    }
    if (behavior is "refusal-eof" or "refusal-result-io")
    {
        using var finalStatus = JsonDocument.Parse(adapter.Status());
        Require(State(adapter) == (behavior == "refusal-eof" ? "failed" : "unknown") &&
            finalStatus.RootElement.GetProperty("manualReviewRequired").GetBoolean(),
            "Refusal shutdown or result I/O failure requires review: " + behavior);
    }
    bool guarded = File.Exists(Path.Combine(fake, ".xlwright", "mcp-development", "active-smoke.json"));
    Require(guarded == (behavior is "fail" or "result-io" or "refusal" or "refusal-zero" or "refusal-eof" or "refusal-result-io" or
        "fault-starting" or "fault-completed"), "Guard retained only on failure: " + behavior);
    Environment.CurrentDirectory = repository;
}
Console.WriteLine($"Passed {passed} non-COM checks.");
return 0;

string PrepareFake(string behavior)
{
    string fake = Path.Combine(testRoot, behavior);
    string fakeBin = Path.Combine(fake, "spikes", "Spike0", "bin", "Release", "net10.0-windows");
    Directory.CreateDirectory(fakeBin);
    Directory.CreateDirectory(Path.Combine(fake, ".xlwright", "spike0"));
    File.WriteAllText(Path.Combine(fake, "spikes", "Spike0", "Spike0.csproj"), "<Project />");
    File.WriteAllText(Path.Combine(fake, ".xlwright", "spike0", ".xlwright.lock"), "{\"version\":1,\"kind\":\"completed\"}");
    if (behavior == "blocked")
        File.WriteAllText(Path.Combine(fake, ".xlwright", "spike0", ".xlwright.lock"), "{\"version\":1,\"kind\":\"starting\"}");
    File.WriteAllText(Path.Combine(fake, "fake-behavior.txt"), behavior);
    foreach (string file in Directory.GetFiles(checksDirectory))
        File.Copy(file, Path.Combine(fakeBin, Path.GetFileName(file)));
    foreach (string suffix in new[] { ".dll", ".deps.json", ".runtimeconfig.json" })
        File.Copy(Path.Combine(checksDirectory, "XlWright.McpServer.Checks" + suffix), Path.Combine(fakeBin, "Spike0" + suffix));

    return fake;
}

void Require(bool condition, string name)
{
    if (!condition) throw new InvalidOperationException(name);
    passed++;
    Console.WriteLine("PASS " + name);
}

static string State(DevelopmentTools adapter)
{
    using var document = JsonDocument.Parse(adapter.Status());
    return document.RootElement.GetProperty("state").GetString()!;
}
