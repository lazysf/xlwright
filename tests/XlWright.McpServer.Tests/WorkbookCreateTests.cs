using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The session without a main workbook and workbook_create (docs/api.md, stage 4), with a fake Excel behind the Core boundary.
public sealed class WorkbookCreateTests : IDisposable
{
    private static readonly SessionTimeouts Short = new(
        Start: TimeSpan.FromMilliseconds(500),
        Sync: TimeSpan.FromMilliseconds(500),
        Save: TimeSpan.FromMilliseconds(500),
        QuitExit: TimeSpan.FromMilliseconds(300),
        CleanupQuitExit: TimeSpan.FromMilliseconds(300),
        TerminateExit: TimeSpan.FromMilliseconds(200),
        PendingReturn: TimeSpan.FromMilliseconds(300));

    private readonly TempRepository _repo = new();
    private readonly FakeLauncher _launcher = new();
    private readonly ExcelSession _session;
    private readonly AppFolder _empty;

    public WorkbookCreateTests()
    {
        _repo.AddApp("apps/new", "notes.txt");
        _empty = AppFolder.Resolve(_repo.Root, "apps/new").Folder!;
        _session = new ExcelSession(_launcher, new FakeProcessQuery(), new FakeConnector(), Short);
    }

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task StartsWithoutOpeningAWorkbookAndRecordsAnEmptyName()
    {
        var outcome = await _session.StartAsync(_empty);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Null(outcome.Result!["workbook"]);
        Assert.True(outcome.Result.ContainsKey("workbook"));
        Assert.Equal(["TakeControl"], _launcher.Excel.Calls);
        Assert.Null(_session.Status().Workbook);

        await _session.EndAsync(save: false);
        var record = OwnershipRecord.Parse(File.ReadAllText(Path.Combine(_empty.FullPath, ".xlwright", "session.json")))!;
        Assert.Equal("", record.Workbook);
    }

    [Fact]
    public async Task OtherToolsAskForAWorkbookFirst()
    {
        await _session.StartAsync(_empty);

        AssertFailure(await _session.SaveAsync(), "CONFLICT", "NO_WORKBOOK", "ready");
        AssertFailure(await _session.ListModulesAsync(), "CONFLICT", "NO_WORKBOOK", "ready");
        AssertFailure(await _session.ReadAsync("Sheet1", CellRange.Parse("A1").Range!, false), "CONFLICT", "NO_WORKBOOK", "ready");
        AssertFailure(await _session.ExtractAsync(null), "CONFLICT", "NO_WORKBOOK", "ready");
        Assert.Equal(["TakeControl"], _launcher.Excel.Calls);
    }

    [Fact]
    public async Task EndIgnoresSaveWithoutAWorkbook()
    {
        await _session.StartAsync(_empty);

        var outcome = await _session.EndAsync(save: true);

        Assert.Null(outcome.Failure);
        Assert.False(outcome.Result!["saved"]!.GetValue<bool>());
        Assert.DoesNotContain("Save", _launcher.Excel.Calls);
    }

    [Theory]
    [InlineData("New.xlsm", 52)]
    [InlineData("売上.xlsx", 51)]
    public async Task CreatesTheMainWorkbookAndOpensTheOtherTools(string name, int format)
    {
        await _session.StartAsync(_empty);

        var outcome = await _session.CreateWorkbookAsync(name);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal(name, outcome.Result!["workbook"]!.GetValue<string>());
        Assert.Equal($"Create {Path.Combine(_empty.FullPath, name)} {format}", _launcher.Excel.Calls[^1]);
        Assert.Equal(name, _session.Status().Workbook);
        Assert.Null((await _session.SaveAsync()).Failure);
        Assert.Equal("Save", _launcher.Excel.Calls[^1]);
    }

    [Fact]
    public async Task RejectsABadName()
    {
        await _session.StartAsync(_empty);

        AssertFailure(await _session.CreateWorkbookAsync("data/New.xlsm"), "INVALID_REQUEST", "INVALID_WORKBOOK_NAME", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("Create", StringComparison.Ordinal));
    }

    [Fact]
    public async Task RejectsASecondWorkbook()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        await _session.StartAsync(AppFolder.Resolve(_repo.Root, "apps/sales").Folder!);

        AssertFailure(await _session.CreateWorkbookAsync("New.xlsm"), "CONFLICT", "WORKBOOK_EXISTS", "ready");
    }

    [Fact]
    public async Task RejectsAFileWithTheSameName()
    {
        await _session.StartAsync(_empty);
        File.WriteAllText(Path.Combine(_empty.FullPath, "New.xlsm"), "");

        AssertFailure(await _session.CreateWorkbookAsync("New.xlsm"), "CONFLICT", "FILE_EXISTS", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("Create", StringComparison.Ordinal));
    }

    [Fact]
    public async Task AFailedSaveLeavesTheSessionWithoutAWorkbook()
    {
        await _session.StartAsync(_empty);
        _launcher.Excel.CreateThrows = new System.Runtime.InteropServices.COMException("disk full");

        AssertFailure(await _session.CreateWorkbookAsync("New.xlsm"), "EXCEL_ERROR", "SAVE_FAILED", "ready");
        Assert.Null(_session.Status().Workbook);
        AssertFailure(await _session.SaveAsync(), "CONFLICT", "NO_WORKBOOK", "ready");
    }

    [Fact]
    public async Task TimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_empty);
        _launcher.Excel.CreateBlock = Block.UntilTerminated;

        AssertFailure(await _session.CreateWorkbookAsync("New.xlsm"), "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task WithoutASessionIsAConflict()
    {
        AssertFailure(await _session.CreateWorkbookAsync("New.xlsm"), "CONFLICT", "NO_SESSION", "closed");
    }

    private static void AssertFailure(SessionOutcome outcome, string code, string detail, string state)
    {
        Assert.Null(outcome.Result);
        Assert.NotNull(outcome.Failure);
        Assert.Equal(code, outcome.Failure.Code);
        Assert.Equal(detail, outcome.Failure.Detail);
        Assert.Equal(state, outcome.Session.State);
    }
}
