using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The app lock and the ownership record of docs/architecture.md "実行時の状態と所有記録" and "開始の順序" 2-5,
// with a fake Excel and fake process queries.
public sealed class SessionRecordTests : IDisposable
{
    private static readonly SessionTimeouts Short = new(
        Start: TimeSpan.FromMilliseconds(500),
        Sync: TimeSpan.FromMilliseconds(500),
        Save: TimeSpan.FromMilliseconds(500),
        QuitExit: TimeSpan.FromMilliseconds(300),
        TerminateExit: TimeSpan.FromMilliseconds(200),
        PendingReturn: TimeSpan.FromMilliseconds(300));

    private static readonly DateTime Marked = new(2026, 11, 2, 1, 0, 0, DateTimeKind.Utc);

    private readonly TempRepository _repo = new();
    private readonly FakeLauncher _launcher = new();
    private readonly FakeProcessQuery _processes = new();
    private readonly FakeConnector _connector = new();
    private readonly ExcelSession _session;
    private readonly AppFolder _app;

    public SessionRecordTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, _processes, _connector, Short);
    }

    private string RecordPath => Path.Combine(_app.FullPath, ".xlwright", "session.json");

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task StartRecordsTheOwnedExcelAndEndCompletesIt()
    {
        _launcher.Process.Facts = _launcher.Process.Facts with { CreatedUtc = DateTime.UtcNow.AddMinutes(1) };
        await _session.StartAsync(_app);
        Assert.Throws<IOException>(() => File.ReadAllText(RecordPath));

        var outcome = await _session.EndAsync(save: false);

        Assert.Null(outcome.Failure);
        var record = ReadRecord();
        Assert.Equal(OwnershipRecord.Completed, record.State);
        Assert.Equal("Sales.xlsm", record.Workbook);
        Assert.Null(record.Process);
    }

    [Fact]
    public async Task StartWritesTheOwnedRecordBeforeItReturns()
    {
        var created = DateTime.UtcNow.AddMinutes(1);
        _launcher.Process.Facts = _launcher.Process.Facts with { CreatedUtc = created };

        await _session.StartAsync(_app);
        _session.Dispose();

        var record = ReadRecord();
        Assert.Equal(OwnershipRecord.Owned, record.State);
        Assert.Equal(new RecordedProcess(4242, created, "EXCEL"), record.Process);
    }

    [Fact]
    public async Task AnotherSessionCannotStartTheSameApp()
    {
        await _session.StartAsync(_app);
        var other = new FakeLauncher();
        using var second = new ExcelSession(other, new FakeProcessQuery(), new FakeConnector(), Short);

        var outcome = await second.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "SESSION_IN_USE");
        Assert.Equal(0, other.Launches);
    }

    [Fact]
    public async Task TheAppIsFreeAgainAfterTheEnd()
    {
        await _session.StartAsync(_app);
        await _session.EndAsync(save: false);
        using var second = new ExcelSession(new FakeLauncher(), new FakeProcessQuery(), new FakeConnector(), Short);

        var outcome = await second.StartAsync(_app);

        Assert.Null(outcome.Failure);
    }

    [Theory]
    [InlineData("")]
    [InlineData("{\"version\":1")]
    [InlineData("{\"version\":1,\"state\":\"completed\",\"marked_at_utc\":\"2026-11-02T01:00:00.0000000Z\",\"workbook\":\"Sales.xlsm\",\"extra\":0}")]
    public async Task StartStopsAtARecordItCannotRead(string text)
    {
        WriteRecord(text);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "RECORD_INVALID");
        Assert.Equal("apps/sales/.xlwright/session.json", outcome.Failure!.Data!["record"]!.GetValue<string>());
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(text, File.ReadAllText(RecordPath));
    }

    [Fact]
    public async Task StartStopsWhenAnExcelStartedAfterTheMarkIsRunning()
    {
        WriteRecord(new OwnershipRecord(OwnershipRecord.Starting, Marked, "Sales.xlsm", null).ToJson());
        _processes.Automation.Add(new ProcessFacts(777, "EXCEL.EXE", Marked.AddSeconds(1), null));

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "UNKNOWN_EXCEL");
        var process = Assert.Single(outcome.Failure!.Data!["processes"]!.AsArray())!;
        Assert.Equal(777, process["pid"]!.GetValue<int>());
        Assert.Equal(Marked, _processes.AutomationSince);
        Assert.Equal(0, _launcher.Launches);
    }

    [Fact]
    public async Task StartCompletesAStartingMarkWithoutAnExcel()
    {
        WriteRecord(new OwnershipRecord(OwnershipRecord.Starting, Marked, "Sales.xlsm", null).ToJson());

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal(1, _launcher.Launches);
    }

    [Fact]
    public async Task StartStopsAtALeftoverExcelThatStillRuns()
    {
        var leftover = new RecordedProcess(999, Marked.AddSeconds(1), "EXCEL");
        WriteRecord(new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", leftover).ToJson());
        _processes.Running.Add(leftover);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "LEFTOVER_EXCEL");
        Assert.Equal(999, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("apps/sales/.xlwright/session.json", outcome.Failure.Data["record"]!.GetValue<string>());
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task StartStopsWhenTheLeftoverExcelCannotBeChecked()
    {
        var leftover = new RecordedProcess(999, Marked.AddSeconds(1), "EXCEL");
        WriteRecord(new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", leftover).ToJson());
        _processes.Unchecked.Add(leftover);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "LEFTOVER_EXCEL");
        Assert.Equal(999, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task APendingCallAfterOwnershipKeepsTheApp()
    {
        _launcher.Excel.OpenBlock = Block.Forever;

        var outcome = await _session.StartAsync(_app);
        using var second = new ExcelSession(new FakeLauncher(), new FakeProcessQuery(), new FakeConnector(), Short);
        var again = await second.StartAsync(_app);

        Assert.Equal("TIMEOUT", outcome.Failure!.Detail);
        AssertFailure(again, "CONFLICT", "SESSION_IN_USE");
    }

    [Fact]
    public async Task StartCompletesAnOwnedRecordWhoseExcelHasEnded()
    {
        WriteRecord(new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", new RecordedProcess(999, Marked, "EXCEL")).ToJson());

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal(1, _launcher.Launches);
    }

    [Fact]
    public async Task StartRemovesTheResultsOfTheLastSession()
    {
        var results = Path.Combine(_app.FullPath, ".xlwright", "results");
        Directory.CreateDirectory(results);
        File.WriteAllText(Path.Combine(results, "old.json"), "{}");

        await _session.StartAsync(_app);

        Assert.False(Directory.Exists(results));
    }

    [Fact]
    public async Task StartStopsWhenAnotherProgramHasTheWorkbookOpenForEditing()
    {
        using (new FileStream(_app.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            var outcome = await _session.StartAsync(_app);

            AssertFailure(outcome, "CONFLICT", "WORKBOOK_IN_USE");
        }

        Assert.Equal(0, _launcher.Launches);
    }

    [Fact]
    public async Task StartGoesOnWhenAnotherProgramOnlyReadsTheWorkbook()
    {
        using var reader = new FileStream(_app.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
    }

    [Fact]
    public async Task StartStopsAtAReadOnlyWorkbookFile()
    {
        File.SetAttributes(_app.WorkbookPath, FileAttributes.ReadOnly);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "WORKBOOK_FILE_READ_ONLY");
        Assert.Equal(0, _launcher.Launches);
    }

    [Fact]
    public async Task AnExitThatIsNotConfirmedLeavesTheRecordOwnedAndFreesTheApp()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitEndsProcess = false;
        _launcher.Process.TerminateEnds = false;

        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "CONFLICT", "EXIT_NOT_CONFIRMED");
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task AnUnconfirmedOwnerLeavesTheStartingMarkAndFreesTheApp()
    {
        _launcher.Process.Facts = _launcher.Process.Facts with { CommandLine = null };

        var outcome = await _session.StartAsync(_app);

        Assert.Equal("OWNERSHIP_NOT_CONFIRMED", outcome.Failure!.Detail);
        Assert.Equal(OwnershipRecord.Starting, ReadRecord().State);
    }

    [Fact]
    public async Task APendingCallKeepsTheApp()
    {
        _launcher.LaunchBlock = Block.Forever;

        var outcome = await _session.StartAsync(_app);
        using var second = new ExcelSession(new FakeLauncher(), new FakeProcessQuery(), new FakeConnector(), Short);
        var again = await second.StartAsync(_app);

        Assert.Equal("TIMEOUT", outcome.Failure!.Detail);
        AssertFailure(again, "CONFLICT", "SESSION_IN_USE");
    }

    [Fact]
    public async Task ATimeoutThatEndsTheExcelCompletesTheRecord()
    {
        _launcher.Excel.OpenBlock = Block.UntilTerminated;

        var outcome = await _session.StartAsync(_app);

        Assert.Equal("TIMEOUT", outcome.Failure!.Detail);
        Assert.Equal(OwnershipRecord.Completed, ReadRecord().State);
    }

    private OwnershipRecord ReadRecord() => OwnershipRecord.Parse(File.ReadAllText(RecordPath))!;

    private void WriteRecord(string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(RecordPath)!);
        File.WriteAllText(RecordPath, text);
    }

    private static void AssertFailure(SessionOutcome outcome, string code, string detail)
    {
        Assert.Null(outcome.Result);
        Assert.NotNull(outcome.Failure);
        Assert.Equal(code, outcome.Failure.Code);
        Assert.Equal(detail, outcome.Failure.Detail);
        Assert.Equal("closed", outcome.Session.State);
    }
}
