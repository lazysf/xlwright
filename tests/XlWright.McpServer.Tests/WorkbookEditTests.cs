using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// workbook_get and workbook_edit convert_to_xlsm (docs/api.md, stage 4), with a fake Excel behind the Core boundary.
public sealed class WorkbookEditTests : IDisposable
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
    private readonly AppFolder _xlsx;

    public WorkbookEditTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsx");
        _xlsx = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, new FakeProcessQuery(), new FakeConnector(), Short);
    }

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task GetReturnsTheNameTheFormatAndWhetherItIsSaved(bool saved)
    {
        await _session.StartAsync(_xlsx);
        _launcher.Excel.Saved = saved;

        var outcome = await _session.GetWorkbookAsync();

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("Sales.xlsx", outcome.Result!["workbook"]!.GetValue<string>());
        Assert.Equal("xlsx", outcome.Result["format"]!.GetValue<string>());
        Assert.Equal(saved, outcome.Result["saved"]!.GetValue<bool>());
    }

    [Fact]
    public async Task GetTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_xlsx);
        _launcher.Excel.SavedBlock = Block.UntilTerminated;

        AssertFailure(await _session.GetWorkbookAsync(), "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task ConvertSavesAsXlsmAndTheSessionUsesIt()
    {
        await _session.StartAsync(_xlsx);

        var outcome = await _session.ConvertToXlsmAsync();

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("Sales.xlsm", outcome.Result!["workbook"]!.GetValue<string>());
        Assert.Equal("Sales.xlsx", outcome.Result["previous"]!.GetValue<string>());
        Assert.Equal($"SaveAs {Path.Combine(_xlsx.FullPath, "Sales.xlsm")} 52", _launcher.Excel.Calls[^1]);
        Assert.Equal("Sales.xlsm", _session.Status().Workbook);
        Assert.Equal("xlsm", (await _session.GetWorkbookAsync()).Result!["format"]!.GetValue<string>());
        Assert.Null((await _session.ListModulesAsync()).Failure);
    }

    [Fact]
    public async Task ConvertRejectsAMacroEnabledWorkbook()
    {
        _repo.AddApp("apps/macro", "Macro.xlsm");
        await _session.StartAsync(AppFolder.Resolve(_repo.Root, "apps/macro").Folder!);

        AssertFailure(await _session.ConvertToXlsmAsync(), "INVALID_REQUEST", "ALREADY_MACRO_ENABLED", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("SaveAs", StringComparison.Ordinal));
    }

    [Fact]
    public async Task ConvertRejectsAFileWithTheSameName()
    {
        await _session.StartAsync(_xlsx);
        File.WriteAllText(Path.Combine(_xlsx.FullPath, "Sales.xlsm"), "");

        AssertFailure(await _session.ConvertToXlsmAsync(), "CONFLICT", "FILE_EXISTS", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("SaveAs", StringComparison.Ordinal));
    }

    [Fact]
    public async Task AFailedSaveKeepsTheXlsx()
    {
        await _session.StartAsync(_xlsx);
        _launcher.Excel.SaveAsThrows = new System.Runtime.InteropServices.COMException("disk full");

        AssertFailure(await _session.ConvertToXlsmAsync(), "EXCEL_ERROR", "SAVE_FAILED", "ready");
        Assert.Equal("Sales.xlsx", _session.Status().Workbook);
    }

    [Fact]
    public async Task ConvertTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_xlsx);
        _launcher.Excel.SaveAsBlock = Block.UntilTerminated;

        AssertFailure(await _session.ConvertToXlsmAsync(), "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task BothAskForAWorkbookFirst()
    {
        _repo.AddApp("apps/new", "notes.txt");
        await _session.StartAsync(AppFolder.Resolve(_repo.Root, "apps/new").Folder!);

        AssertFailure(await _session.GetWorkbookAsync(), "CONFLICT", "NO_WORKBOOK", "ready");
        AssertFailure(await _session.ConvertToXlsmAsync(), "CONFLICT", "NO_WORKBOOK", "ready");
    }

    [Fact]
    public async Task WithoutASessionIsAConflict()
    {
        AssertFailure(await _session.GetWorkbookAsync(), "CONFLICT", "NO_SESSION", "closed");
        AssertFailure(await _session.ConvertToXlsmAsync(), "CONFLICT", "NO_SESSION", "closed");
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
