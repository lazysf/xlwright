using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "残った専用のExcelの片付け" 2-5: the next start cleans up the Excel an owned record names,
// with a fake connection to that Excel.
public sealed class SessionCleanupTests : IDisposable
{
    private static readonly SessionTimeouts Short = new(
        Start: TimeSpan.FromMilliseconds(800),
        Sync: TimeSpan.FromMilliseconds(500),
        Save: TimeSpan.FromMilliseconds(500),
        QuitExit: TimeSpan.FromMilliseconds(300),
        CleanupQuitExit: TimeSpan.FromMilliseconds(300),
        TerminateExit: TimeSpan.FromMilliseconds(200),
        PendingReturn: TimeSpan.FromMilliseconds(300));

    private static readonly DateTime Marked = new(2026, 11, 2, 1, 0, 0, DateTimeKind.Utc);
    private static readonly RecordedProcess Recorded = new(999, Marked.AddSeconds(1), "EXCEL");

    private readonly TempRepository _repo = new();
    private readonly FakeLauncher _launcher = new();
    private readonly FakeProcessQuery _processes = new();
    private readonly FakeConnector _connector = new();
    private readonly FakeExcel _leftover = new();
    private readonly ExcelSession _session;
    private readonly AppFolder _app;

    public SessionCleanupTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, _processes, _connector, Short);
        var record = Path.Combine(_app.FullPath, ".xlwright", "session.json");
        Directory.CreateDirectory(Path.GetDirectoryName(record)!);
        File.WriteAllText(record, new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", Recorded).ToJson());
        _processes.Running.Add(Recorded);
        _leftover.Process = _processes.Leftover;
        _leftover.Open(_app.WorkbookPath);
        _connector.Result = new LeftoverWindow.Connected(_leftover);
    }

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task ClosesALeftoverWithOnlyTheAppWorkbooksAndStarts()
    {
        _leftover.ExtraWorkbooks.Add(Path.Combine(_app.FullPath, "data", "Input.xlsx"));

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal(999, _connector.Connects.Single());
        Assert.Equal(["Open " + _app.WorkbookPath, "OpenWorkbookPaths", "CloseAllAndQuit"], _leftover.Calls);
        var cleanup = outcome.Result!["cleanup"]!;
        Assert.Equal(999, cleanup["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("quit", cleanup["ended_by"]!.GetValue<string>());
        Assert.Equal(0, _processes.Leftover.Terminations);
        Assert.Equal(1, _launcher.Launches);
    }

    [Fact]
    public async Task EndsALeftoverThatDoesNotEndAfterQuit()
    {
        _leftover.QuitEndsProcess = false;

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal("terminate", outcome.Result!["cleanup"]!["ended_by"]!.GetValue<string>());
        Assert.Equal(1, _processes.Leftover.Terminations);
    }

    [Fact]
    public async Task WaitsAfterQuitOnlyForTheCleanupWait()
    {
        _leftover.QuitEndsProcess = false;
        using var session = new ExcelSession(_launcher, _processes, _connector,
            Short with { Start = TimeSpan.FromSeconds(30), QuitExit = TimeSpan.FromSeconds(20) });
        var watch = System.Diagnostics.Stopwatch.StartNew();

        var outcome = await session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal("terminate", outcome.Result!["cleanup"]!["ended_by"]!.GetValue<string>());
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), $"took {watch.Elapsed}");
    }

    [Theory]
    [InlineData("open failed")]
    [InlineData("read-only file")]
    public async Task ReportsTheCleanupWhenTheStartFailsAfterIt(string why)
    {
        if (why == "open failed") _launcher.Excel.OpenThrows = new InvalidOperationException("cannot open");
        if (why == "read-only file") File.SetAttributes(_app.WorkbookPath, FileAttributes.ReadOnly);

        var outcome = await _session.StartAsync(_app);
        File.SetAttributes(_app.WorkbookPath, FileAttributes.Normal);

        Assert.NotNull(outcome.Failure);
        var cleanup = outcome.Failure.Data!["cleanup"]!;
        Assert.Equal(999, cleanup["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("quit", cleanup["ended_by"]!.GetValue<string>());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("\"EXCEL.EXE\" C:\\Daily\\Book.xlsx")]
    public async Task LeavesALeftoverNotStartedFromCom(string? commandLine)
    {
        _processes.LeftoverCommandLine = commandLine;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "LEFTOVER_EXCEL");
        Assert.Equal(999, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
        Assert.Empty(_connector.Connects);
        Assert.Equal(0, _processes.Leftover.Terminations);
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task EndsALeftoverWhoseQuitDoesNotReturn()
    {
        _leftover.QuitBlock = Block.UntilTerminated;

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal("terminate", outcome.Result!["cleanup"]!["ended_by"]!.GetValue<string>());
    }

    [Fact]
    public async Task DoesNotStartWhenAQuitOfTheLeftoverIsStillPending()
    {
        _leftover.QuitBlock = Block.Forever;

        var outcome = await _session.StartAsync(_app);
        var again = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT");
        Assert.True(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.False(outcome.Failure.Data["restartable"]!.GetValue<bool>());
        Assert.Equal("terminate", outcome.Failure.Data["cleanup"]!["ended_by"]!.GetValue<string>());
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal("SERVER_UNUSABLE", again.Failure!.Detail);
    }

    [Fact]
    public async Task EndsALeftoverWithoutAWorkbookWindowWithoutConnecting()
    {
        _connector.Result = new LeftoverWindow.NoWorkbookWindow();

        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal("terminate", outcome.Result!["cleanup"]!["ended_by"]!.GetValue<string>());
        Assert.Equal(1, _processes.Leftover.Terminations);
        Assert.Equal(["Open " + _app.WorkbookPath], _leftover.Calls);
    }

    [Theory]
    [InlineData("outside")]
    [InlineData("no path")]
    [InlineData("not reachable")]
    public async Task LeavesALeftoverItCannotCheck(string why)
    {
        if (why == "outside") _leftover.ExtraWorkbooks.Add(Path.Combine(_repo.Root, "Other.xlsx"));
        if (why == "no path") _leftover.ExtraWorkbooks.Add(null);
        if (why == "not reachable") _connector.Result = new LeftoverWindow.NotReachable("no window gave an Excel");

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "LEFTOVER_EXCEL");
        Assert.Equal(999, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("apps/sales/.xlwright/session.json", outcome.Failure.Data["record"]!.GetValue<string>());
        Assert.DoesNotContain("CloseAllAndQuit", _leftover.Calls);
        Assert.Equal(0, _processes.Leftover.Terminations);
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task ReleasesAConnectedLeftoverItLeaves()
    {
        _leftover.ExtraWorkbooks.Add(null);

        await _session.StartAsync(_app);

        Assert.Contains("ReleaseReferences", _leftover.Calls);
    }

    [Fact]
    public async Task ReportsALeftoverWhoseEndIsNotConfirmed()
    {
        _connector.Result = new LeftoverWindow.NoWorkbookWindow();
        _processes.Leftover.TerminateEnds = false;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "EXIT_NOT_CONFIRMED");
        Assert.Equal(999, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("apps/sales/.xlwright/session.json", outcome.Failure.Data["record"]!.GetValue<string>());
        Assert.Equal(0, _launcher.Launches);
        Assert.Equal(OwnershipRecord.Owned, ReadRecord().State);
    }

    [Fact]
    public async Task LeavesALeftoverWhenTheConnectionDoesNotReturnAndRefusesLaterStarts()
    {
        _connector.BlocksForever = true;

        var outcome = await _session.StartAsync(_app);
        var again = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT");
        Assert.False(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.False(outcome.Failure.Data["restartable"]!.GetValue<bool>());
        Assert.Equal(999, outcome.Failure.Data["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("apps/sales/.xlwright/session.json", outcome.Failure.Data["record"]!.GetValue<string>());
        Assert.Equal(0, _processes.Leftover.Terminations);
        Assert.Equal("SERVER_UNUSABLE", again.Failure!.Detail);
    }

    private OwnershipRecord ReadRecord() =>
        OwnershipRecord.Parse(File.ReadAllText(Path.Combine(_app.FullPath, ".xlwright", "session.json")))!;

    private static void AssertFailure(SessionOutcome outcome, string code, string detail)
    {
        Assert.Null(outcome.Result);
        Assert.NotNull(outcome.Failure);
        Assert.Equal(code, outcome.Failure.Code);
        Assert.Equal(detail, outcome.Failure.Detail);
        Assert.Equal("closed", outcome.Session.State);
    }
}
