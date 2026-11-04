using System.Text.Json.Nodes;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "診断の記録の形式": one JSON Lines file per session_start, the events, and data.log in failures,
// with a fake Excel.
public sealed class DiagnosticLogTests : IDisposable
{
    private static readonly SessionTimeouts Short = new(
        Start: TimeSpan.FromMilliseconds(800),
        Sync: TimeSpan.FromMilliseconds(500),
        Save: TimeSpan.FromMilliseconds(500),
        QuitExit: TimeSpan.FromMilliseconds(300),
        CleanupQuitExit: TimeSpan.FromMilliseconds(300),
        TerminateExit: TimeSpan.FromMilliseconds(200),
        PendingReturn: TimeSpan.FromMilliseconds(300));

    private readonly TempRepository _repo = new();
    private readonly FakeLauncher _launcher = new();
    private readonly FakeProcessQuery _processes = new();
    private readonly FakeConnector _connector = new();
    private readonly ExcelSession _session;
    private readonly AppFolder _app;

    public DiagnosticLogTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, _processes, _connector, Short);
    }

    private string LogFolder => Path.Combine(_app.FullPath, ".xlwright", "logs");

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task WritesTheStartTheCallsAndTheEndOfASession()
    {
        await _session.StartAsync(_app);
        await _session.ReadAsync("Data", Range("A1"), formulas: false);
        await _session.EndAsync(save: false);

        var file = Assert.Single(Directory.GetFiles(LogFolder));
        Assert.Matches(@"^\d{8}T\d{6}Z-[0-9a-f]{8}\.jsonl$", Path.GetFileName(file));
        var lines = Lines();
        Assert.Equal(["start", "owned", "call", "call", "end", "call"], lines.Select(line => line["event"]!.GetValue<string>()));
        Assert.All(lines, line => DateTime.Parse(line["time_utc"]!.GetValue<string>()));
        Assert.Equal("apps/sales", lines[0]["app"]!.GetValue<string>());
        Assert.Equal(Environment.ProcessId, lines[0]["server_pid"]!.GetValue<int>());
        Assert.Equal(4242, lines[1]["pid"]!.GetValue<int>());
        Assert.Equal("session_start", lines[2]["tool"]!.GetValue<string>());
        Assert.True(lines[2]["ok"]!.GetValue<bool>());
        Assert.Equal("range_read", lines[3]["tool"]!.GetValue<string>());
        Assert.Equal("quit", lines[4]["ended_by"]!.GetValue<string>());
        Assert.True(lines[4]["exit_confirmed"]!.GetValue<bool>());
    }

    [Fact]
    public async Task PointsAFailureInTheSessionToTheLog()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.ReadAsync("Missing", Range("A1"), formulas: false);

        var log = outcome.Failure!.Data!["log"]!.GetValue<string>();
        Assert.Matches(@"^apps/sales/\.xlwright/logs/\d{8}T\d{6}Z-[0-9a-f]{8}\.jsonl$", log);
        Assert.True(File.Exists(Path.Combine(_repo.Root, log)));
        var call = Lines().Last();
        Assert.Equal("NOT_FOUND", call["code"]!.GetValue<string>());
        Assert.Equal("SHEET_NOT_FOUND", call["detail"]!.GetValue<string>());
        Assert.Null(call["ok"]);
    }

    [Fact]
    public async Task PointsAFailedStartToItsLog()
    {
        File.SetAttributes(_app.WorkbookPath, FileAttributes.ReadOnly);

        var outcome = await _session.StartAsync(_app);
        File.SetAttributes(_app.WorkbookPath, FileAttributes.Normal);

        Assert.Equal("WORKBOOK_FILE_READ_ONLY", outcome.Failure!.Detail);
        Assert.NotNull(outcome.Failure.Data!["log"]);
        Assert.Equal("WORKBOOK_FILE_READ_ONLY", Lines().Last()["detail"]!.GetValue<string>());
    }

    [Fact]
    public async Task LeavesTheLogOutOfAFailureOutsideASession()
    {
        await _session.StartAsync(_app);
        await _session.EndAsync(save: false);

        var outcome = await _session.SaveAsync();

        Assert.Equal("NO_SESSION", outcome.Failure!.Detail);
        Assert.Null(outcome.Failure.Data?["log"]);
    }

    [Fact]
    public async Task WritesTheWindowsAndTheEndWhenATimeLimitEndsTheExcel()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.RangeBlock = Block.UntilTerminated;

        var outcome = await _session.ReadAsync("Data", Range("A1"), formulas: false);

        Assert.Equal("TIMEOUT", outcome.Failure!.Detail);
        Assert.NotNull(outcome.Failure.Data!["log"]);
        var lines = Lines();
        var windows = lines.Single(line => line["event"]!.GetValue<string>() == "windows");
        Assert.Equal([true], windows["xlmain_enabled"]!.AsArray().Select(node => node!.GetValue<bool>()));
        Assert.False(windows["dialog"]!.GetValue<bool>());
        var end = lines.Single(line => line["event"]!.GetValue<string>() == "end");
        Assert.Equal("terminate", end["ended_by"]!.GetValue<string>());
    }

    [Fact]
    public async Task WritesTheCleanupOfALeftover()
    {
        var recorded = new RecordedProcess(999, DateTime.UtcNow.AddMinutes(-5), "EXCEL");
        var record = Path.Combine(_app.FullPath, ".xlwright", "session.json");
        Directory.CreateDirectory(Path.GetDirectoryName(record)!);
        File.WriteAllText(record, new OwnershipRecord(OwnershipRecord.Owned, recorded.StartedAtUtc.AddSeconds(-1), "Sales.xlsm", recorded).ToJson());
        _processes.Running.Add(recorded);
        _connector.Result = new LeftoverWindow.NoWorkbookWindow();

        await _session.StartAsync(_app);

        var cleanup = Lines().Single(line => line["event"]!.GetValue<string>() == "cleanup");
        Assert.Equal(999, cleanup["pid"]!.GetValue<int>());
        Assert.Equal("terminate", cleanup["ended_by"]!.GetValue<string>());
    }

    [Fact]
    public async Task WritesALeftoverThatWasLeftUntouched()
    {
        var recorded = new RecordedProcess(999, DateTime.UtcNow.AddMinutes(-5), "EXCEL");
        var record = Path.Combine(_app.FullPath, ".xlwright", "session.json");
        Directory.CreateDirectory(Path.GetDirectoryName(record)!);
        File.WriteAllText(record, new OwnershipRecord(OwnershipRecord.Owned, recorded.StartedAtUtc.AddSeconds(-1), "Sales.xlsm", recorded).ToJson());
        _processes.Running.Add(recorded);

        await _session.StartAsync(_app);

        var cleanup = Lines().Single(line => line["event"]!.GetValue<string>() == "cleanup");
        Assert.Equal(999, cleanup["pid"]!.GetValue<int>());
        Assert.Equal("LEFTOVER_EXCEL", cleanup["detail"]!.GetValue<string>());
        Assert.Null(cleanup["ended_by"]);
    }

    [Fact]
    public async Task WritesNoValuesArgumentsReturnsOrMessages()
    {
        await _session.StartAsync(_app);
        await _session.WriteAsync("Data", Range("A1"), RangeWrite.OfValues(new object?[,] { { "secret-cell" } }));
        _launcher.Excel.MacroResult = new MacroResult.Returned("secret-return");
        await _session.RunAsync(new MacroName("Calc", "Twice"), ["secret-argument"], TimeSpan.FromSeconds(5));
        await _session.ReadAsync("Missing", Range("A1"), formulas: false);

        var text = File.ReadAllText(Assert.Single(Directory.GetFiles(LogFolder)));
        Assert.DoesNotContain("secret", text);
        Assert.DoesNotContain("Missing", text);
        Assert.Equal("Calc.Twice", Lines().Single(line => line["tool"]?.GetValue<string>() == "macro_run")["macro"]!.GetValue<string>());
    }

    [Fact]
    public async Task WorksWithoutALogWhenItCannotBeCreated()
    {
        Directory.CreateDirectory(Path.Combine(_app.FullPath, ".xlwright"));
        File.WriteAllText(LogFolder, "not a folder");

        var started = await _session.StartAsync(_app);
        var failed = await _session.ReadAsync("Missing", Range("A1"), formulas: false);

        Assert.Null(started.Failure);
        Assert.Equal("SHEET_NOT_FOUND", failed.Failure!.Detail);
        Assert.Null(failed.Failure.Data?["log"]);
    }

    private List<JsonNode> Lines() =>
        File.ReadAllLines(Directory.GetFiles(LogFolder).Single()).Select(line => JsonNode.Parse(line)!).ToList();

    private static CellRange Range(string text) => CellRange.Parse(text).Range!;
}
