using System.Text.Json.Nodes;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The start and end rules of docs/architecture.md, with a fake Excel behind the Core boundary.
public sealed class ExcelSessionTests : IDisposable
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
    private readonly AppFolder _app;

    public ExcelSessionTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, new FakeProcessQuery(), new FakeConnector(), Short);
    }

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task StartTakesControlOpensTheMainWorkbookAndReportsTheExcel()
    {
        var outcome = await _session.StartAsync(_app);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal(["TakeControl", "Open " + _app.WorkbookPath], _launcher.Excel.Calls);
        var result = outcome.Result!;
        Assert.Equal("apps/sales", result["app"]!.GetValue<string>());
        Assert.Equal("Sales.xlsm", result["workbook"]!.GetValue<string>());
        Assert.Equal(4242, result["excel"]!["pid"]!.GetValue<int>());
        Assert.Equal("ready", _session.Status().State);
    }

    [Theory]
    [InlineData("EXCEL.EXE", "\"EXCEL.EXE\" -Embedding", 1)]
    [InlineData("NOTEPAD.EXE", "\"NOTEPAD.EXE\" /automation", 1)]
    [InlineData("EXCEL.EXE", "\"EXCEL.EXE\" /automation -Embedding", -60)]
    [InlineData("EXCEL.EXE", null, 1)]
    public async Task StartLeavesAnExcelWhoseOwnershipIsNotConfirmed(string image, string? commandLine, int createdAfterSeconds)
    {
        _launcher.Process.Facts = new ProcessFacts(4242, image, DateTime.UtcNow.AddSeconds(createdAfterSeconds), commandLine);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "OWNERSHIP_NOT_CONFIRMED", "closed");
        Assert.False(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.False(outcome.Failure.Data["restartable"]!.GetValue<bool>());
        Assert.Empty(_launcher.Excel.Calls);
        Assert.Equal(0, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task StartEndsTheExcelWhenTheWorkbookOpensReadOnly()
    {
        _launcher.Excel.OpensReadOnly = true;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "WORKBOOK_READ_ONLY", "closed");
        Assert.Contains("CloseAllAndQuit", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task StartEndsTheExcelWhenTheWorkbookDoesNotOpen()
    {
        _launcher.Excel.OpenThrows = new InvalidOperationException("cannot open");

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "EXCEL_ERROR", "OPEN_FAILED", "closed");
        Assert.Contains("CloseAllAndQuit", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task StartReportsAnExcelThatDidNotStart()
    {
        _launcher.LaunchThrows = new InvalidOperationException("class not registered");

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "EXCEL_ERROR", "EXCEL_NOT_STARTED", "closed");
        Assert.Contains("class not registered", outcome.Failure!.Message);
        Assert.Empty(_launcher.Excel.Calls);
    }

    [Fact]
    public async Task StartRejectsASecondSession()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "SESSION_ACTIVE", "ready");
        Assert.Equal(1, _launcher.Launches);
    }

    [Fact]
    public async Task StartTimeoutBeforeOwnershipLeavesTheExcelUntouched()
    {
        _launcher.LaunchBlock = Block.UntilTerminated;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.False(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.Equal(0, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task StartTimeoutAfterOwnershipTerminatesTheOwnedExcel()
    {
        _launcher.Excel.OpenBlock = Block.UntilTerminated;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.True(outcome.Failure.Data["restartable"]!.GetValue<bool>());
    }

    [Fact]
    public async Task APendingCallAfterTerminationMakesTheServerUnusable()
    {
        _launcher.Excel.OpenBlock = Block.Forever;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.False(outcome.Failure!.Data!["restartable"]!.GetValue<bool>());
        var again = await _session.StartAsync(_app);
        AssertFailure(again, "CONFLICT", "SERVER_UNUSABLE", "closed");
    }

    [Fact]
    public async Task EndSavesClosesAndWaitsForTheExcelToExit()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.EndAsync(save: true);

        Assert.Null(outcome.Failure);
        Assert.Equal("closed", outcome.Session.State);
        Assert.True(outcome.Result!["saved"]!.GetValue<bool>());
        Assert.Equal(["Save", "OpenWorkbookPaths", "CloseAllAndQuit"], _launcher.Excel.Calls.Skip(2));
        Assert.Equal(0, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task EndWithoutSavingDoesNotSave()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.EndAsync(save: false);

        Assert.False(outcome.Result!["saved"]!.GetValue<bool>());
        Assert.DoesNotContain("Save", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task EndKeepsTheSessionWhenSavingFails()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveThrows = new InvalidOperationException("disk full");

        var outcome = await _session.EndAsync(save: true);

        AssertFailure(outcome, "EXCEL_ERROR", "SAVE_FAILED", "ready");
        Assert.DoesNotContain("CloseAllAndQuit", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task EndLeavesTheExcelWhenAWorkbookOutsideTheAppIsOpen()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExtraWorkbooks.Add(Path.Combine(_repo.Root, "Other.xlsx"));

        var outcome = await _session.EndAsync(save: true);

        AssertFailure(outcome, "CONFLICT", "FOREIGN_WORKBOOK", "closed");
        Assert.DoesNotContain("CloseAllAndQuit", _launcher.Excel.Calls);
        Assert.Equal(0, _launcher.Process.Terminations);
        var data = outcome.Failure!.Data!;
        Assert.True(data["saved"]!.GetValue<bool>());
        Assert.Equal(4242, data["excel"]!["pid"]!.GetValue<int>());
    }

    [Fact]
    public async Task EndClosesWorkbooksWithoutAPathAndInsideTheApp()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExtraWorkbooks.Add(null);
        _launcher.Excel.ExtraWorkbooks.Add(Path.Combine(_app.FullPath, "data", "Input.xlsx"));

        var outcome = await _session.EndAsync(save: false);

        Assert.Null(outcome.Failure);
        Assert.Contains("CloseAllAndQuit", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task EndTerminatesTheOwnedExcelWhenItDoesNotExitAfterQuit()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitEndsProcess = false;

        var outcome = await _session.EndAsync(save: false);

        Assert.Null(outcome.Failure);
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.Equal("closed", outcome.Session.State);
    }

    [Fact]
    public async Task EndReportsWhenTheExitCannotBeConfirmed()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitEndsProcess = false;
        _launcher.Process.TerminateEnds = false;

        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "CONFLICT", "EXIT_NOT_CONFIRMED", "closed");
        Assert.Equal(4242, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
    }

    [Fact]
    public async Task EndSaveTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveBlock = Block.UntilTerminated;

        var outcome = await _session.EndAsync(save: true);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
    }

    [Fact]
    public async Task EndQuitTimeoutTerminatesTheOwnedExcelAndReportsStateUnknown()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitBlock = Block.UntilTerminated;

        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(outcome.Failure!.Data!["restartable"]!.GetValue<bool>());
    }

    [Fact]
    public async Task AQuitThatNeverReturnsMakesTheServerUnusable()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitBlock = Block.Forever;

        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.False(outcome.Failure!.Data!["restartable"]!.GetValue<bool>());
        AssertFailure(await _session.StartAsync(_app), "CONFLICT", "SERVER_UNUSABLE", "closed");
    }

    [Theory]
    [InlineData("https://example.com/Book.xlsx")]
    [InlineData("Book.xlsx")]
    public async Task EndTreatsAPathThatIsNotFullyQualifiedAsOutsideTheApp(string path)
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExtraWorkbooks.Add(path);

        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "CONFLICT", "FOREIGN_WORKBOOK", "closed");
        Assert.DoesNotContain("CloseAllAndQuit", _launcher.Excel.Calls);
    }

    [Fact]
    public async Task EndReportsTheSaveResultWhenTheExitCannotBeConfirmed()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.QuitEndsProcess = false;
        _launcher.Process.TerminateEnds = false;

        var outcome = await _session.EndAsync(save: true);

        Assert.True(outcome.Failure!.Data!["saved"]!.GetValue<bool>());
    }

    [Fact]
    public async Task StartKeepsTheReasonWhenTheExitCannotBeConfirmed()
    {
        _launcher.Excel.OpensReadOnly = true;
        _launcher.Excel.QuitEndsProcess = false;
        _launcher.Process.TerminateEnds = false;

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "CONFLICT", "WORKBOOK_READ_ONLY", "closed");
        Assert.Equal(4242, outcome.Failure!.Data!["excel"]!["pid"]!.GetValue<int>());
    }

    [Fact]
    public async Task StartReportsAnUnidentifiedProcessAsOwnershipNotConfirmed()
    {
        _launcher.Process.Facts = new ProcessFacts(0, "", DateTime.UtcNow.AddSeconds(1), null);

        var outcome = await _session.StartAsync(_app);

        AssertFailure(outcome, "STATE_UNKNOWN", "OWNERSHIP_NOT_CONFIRMED", "closed");
        Assert.Contains("could not be identified", outcome.Failure!.Message);
    }

    [Fact]
    public async Task SaveSavesAndKeepsTheSession()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.SaveAsync();

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("Sales.xlsm", outcome.Result!["workbook"]!.GetValue<string>());
        Assert.Equal(["Save"], _launcher.Excel.Calls.Skip(2));
    }

    [Fact]
    public async Task SaveWithoutASessionIsAConflict()
    {
        var outcome = await _session.SaveAsync();

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    [Fact]
    public async Task SaveKeepsTheSessionWhenSavingFails()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveThrows = new InvalidOperationException("disk full");

        var outcome = await _session.SaveAsync();

        AssertFailure(outcome, "EXCEL_ERROR", "SAVE_FAILED", "ready");
        Assert.Contains("disk full", outcome.Failure!.Message);
        Assert.Equal(0, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task SaveTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveBlock = Block.UntilTerminated;

        var outcome = await _session.SaveAsync();

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.True(outcome.Failure.Data["restartable"]!.GetValue<bool>());
    }

    [Fact]
    public async Task SaveIsRejectedWhileBusy()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveBlock = Block.UntilTerminated;

        var saving = _session.SaveAsync();
        await Task.Delay(100, TestContext.Current.CancellationToken);
        var second = await _session.SaveAsync();
        await saving;

        AssertFailure(second, "CONFLICT", "BUSY", "ready", busy: true);
    }

    [Fact]
    public async Task ReadReturnsValuesWithNullsAndCellErrors()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.Values = new object?[,] { { 1.0, "a", true }, { null, new CellError("#DIV/0!"), 2.5 } };

        var outcome = await _session.ReadAsync("Data", Range("A1:C2"), formulas: false);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("ReadValues Data!A1:C2", _launcher.Excel.Calls[^1]);
        Assert.Equal("Data", outcome.Result!["sheet"]!.GetValue<string>());
        Assert.Equal("A1:C2", outcome.Result["range"]!.GetValue<string>());
        Assert.Equal("""[[1,"a",true],[null,{"error":"#DIV/0!"},2.5]]""", outcome.Result["values"]!.ToJsonString());
        Assert.Null(outcome.Result["formulas"]);
    }

    [Fact]
    public async Task ReadReturnsFormulasInsteadOfValues()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.Formulas = new[,] { { "=SUM(A2:A3)", "text", "" } };

        var outcome = await _session.ReadAsync("Data", Range("A1:C1"), formulas: true);

        Assert.Equal("ReadFormulas Data!A1:C1", _launcher.Excel.Calls[^1]);
        Assert.Equal("""[["=SUM(A2:A3)","text",""]]""", outcome.Result!["formulas"]!.ToJsonString());
        Assert.Null(outcome.Result["values"]);
    }

    [Fact]
    public async Task WriteReportsTheCellsItWrote()
    {
        await _session.StartAsync(_app);
        var write = RangeWrite.OfValues(new object?[,] { { 1.0, "a" }, { null, false } });

        var outcome = await _session.WriteAsync("Data", Range("B2:C3"), write);

        Assert.Null(outcome.Failure);
        Assert.Equal("Write Data!B2:C3", _launcher.Excel.Calls[^1]);
        Assert.Same(write, _launcher.Excel.Written);
        Assert.Equal("B2:C3", outcome.Result!["range"]!.GetValue<string>());
        Assert.Equal(4, outcome.Result["cells"]!.GetValue<long>());
    }

    [Theory]
    [InlineData("read")]
    [InlineData("write")]
    public async Task ReportsAMissingSheet(string operation)
    {
        await _session.StartAsync(_app);

        var outcome = await Range(operation, "Missing");

        AssertFailure(outcome, "NOT_FOUND", "SHEET_NOT_FOUND", "ready");
        Assert.Contains("Missing", outcome.Failure!.Message);
    }

    [Fact]
    public async Task ReadKeepsTheSessionWhenExcelFails()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.RangeThrows = new InvalidOperationException("broken");

        var outcome = await Range("read", "Sheet1");

        Assert.Equal("EXCEL_ERROR", outcome.Failure!.Code);
        Assert.Null(outcome.Failure.Detail);
        Assert.Contains("broken", outcome.Failure.Message);
        Assert.Equal("ready", outcome.Session.State);
    }

    [Fact]
    public async Task WriteFailureAsksToReadTheRangeAgain()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.RangeThrows = new InvalidOperationException("the sheet is protected");

        var outcome = await Range("write", "Sheet1");

        AssertFailure(outcome, "EXCEL_ERROR", "WRITE_FAILED", "ready");
        Assert.Contains("the sheet is protected", outcome.Failure!.Message);
        Assert.Contains("range_read", outcome.Failure.Message);
    }

    [Theory]
    [InlineData("read")]
    [InlineData("write")]
    public async Task RangeTimeoutTerminatesTheOwnedExcel(string operation)
    {
        await _session.StartAsync(_app);
        _launcher.Excel.RangeBlock = Block.UntilTerminated;

        var outcome = await Range(operation, "Sheet1");

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Theory]
    [InlineData("read")]
    [InlineData("write")]
    public async Task RangeWithoutASessionIsAConflict(string operation)
    {
        var outcome = await Range(operation, "Sheet1");

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    [Fact]
    public async Task SyncReportsTheModulesAndTheRemovedOnesAndDeletesTheConvertedFiles()
    {
        _repo.AddApp("apps/sales", "src/modules/Calc.bas", "src/classes/Counter.cls");
        await _session.StartAsync(_app);
        _launcher.Excel.Removed.Add("Old");

        var outcome = await _session.SyncAsync();

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("SyncVba Calc,XlWright,Counter", _launcher.Excel.Calls[^1]);
        Assert.Equal(
            """[{"name":"Calc","type":"standard"},{"name":"XlWright","type":"standard"},{"name":"Counter","type":"class"}]""",
            outcome.Result!["modules"]!.ToJsonString());
        Assert.Equal("""["Old"]""", outcome.Result["removed"]!.ToJsonString());
        Assert.All(_launcher.Excel.Synced!, m => Assert.False(File.Exists(m.ImportPath)));
    }

    [Fact]
    public async Task SyncRejectsAMainWorkbookWithoutMacros()
    {
        _repo.AddApp("apps/plain", "Plain.xlsx");
        await _session.StartAsync(AppFolder.Resolve(_repo.Root, "apps/plain").Folder!);

        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, "INVALID_REQUEST", "NOT_MACRO_ENABLED", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("SyncVba"));
    }

    [Fact]
    public async Task SyncRejectsAnInvalidSourceWithoutTouchingExcel()
    {
        _repo.AddApp("apps/sales", "src/modules/Calc.cls");
        await _session.StartAsync(_app);

        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, "INVALID_REQUEST", "SOURCE_INVALID", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("SyncVba"));
    }

    [Theory]
    [InlineData("VBA_ACCESS_NOT_TRUSTED", "CONFLICT")]
    [InlineData("VBA_PROJECT_LOCKED", "CONFLICT")]
    [InlineData("SYNC_FAILED", "EXCEL_ERROR")]
    public async Task SyncReportsWhatExcelRefused(string detail, string code)
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SyncThrows = new VbaSyncException(detail, ["Calc"], "refused");

        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, code, detail, "ready");
        Assert.Contains("refused", outcome.Failure!.Message);
        if (detail == "SYNC_FAILED")
        {
            Assert.Equal("""["Calc"]""", outcome.Failure.Data!["done"]!.ToJsonString());
        }
    }

    [Fact]
    public async Task SyncReportsAnUnexpectedFailureAsSyncFailed()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SyncThrows = new InvalidOperationException("broken");

        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, "EXCEL_ERROR", "SYNC_FAILED", "ready");
        Assert.Contains("broken", outcome.Failure!.Message);
        Assert.NotNull(outcome.Failure.Data!["done"]);
    }

    [Fact]
    public async Task SyncTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SyncBlock = Block.UntilTerminated;

        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task SyncWithoutASessionIsAConflict()
    {
        var outcome = await _session.SyncAsync();

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    [Fact]
    public async Task RunSyncsThenRunsAndReturnsTheValue()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.RunAsync(new MacroName("Calc", "Twice"), [21.0], TimeSpan.FromSeconds(5));

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal(["SyncVba XlWright", "RunMacro Calc.Twice 21"], _launcher.Excel.Calls.Skip(2));
        var result = outcome.Result!;
        Assert.Equal("Calc.Twice", result["macro"]!.GetValue<string>());
        Assert.Equal(42, result["return"]!.GetValue<double>());
        Assert.Null(result["return_type"]);
        Assert.True(result["elapsed_ms"]!.GetValue<long>() >= 0);
        Assert.Equal("""{"modules":[{"name":"XlWright","type":"standard"}],"removed":[]}""", result["sync"]!.ToJsonString());
    }

    [Theory]
    [InlineData("Collection")]
    [InlineData("Date")]
    public async Task RunNamesAReturnTypeItDoesNotReturn(string typeName)
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.OfType(typeName);

        var outcome = await _session.RunAsync(new MacroName("Calc", "Make"), [], TimeSpan.FromSeconds(5));

        Assert.True(outcome.Result!.ContainsKey("return"));
        Assert.Null(outcome.Result["return"]);
        Assert.Equal(typeName, outcome.Result["return_type"]!.GetValue<string>());
    }

    [Fact]
    public async Task RunNamesADotNetTypeItCannotReturn()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.Returned(new DateTime(2026, 11, 3));

        var outcome = await _session.RunAsync(new MacroName("Calc", "When"), [], TimeSpan.FromSeconds(5));

        Assert.Equal("DateTime", outcome.Result!["return_type"]!.GetValue<string>());
    }

    [Fact]
    public async Task RunWritesALargeReturnValueToAResultFile()
    {
        await _session.StartAsync(_app);
        var large = new string('x', 30_000);
        _launcher.Excel.MacroResult = new MacroResult.Returned(large);

        var outcome = await _session.RunAsync(new MacroName("Calc", "Big"), [], TimeSpan.FromSeconds(5));

        var result = outcome.Result!;
        Assert.Equal(new string('x', 1_000), result["return"]!.GetValue<string>());
        var file = result["result_file"]!.GetValue<string>();
        Assert.StartsWith("apps/sales/.xlwright/results/", file);
        var full = System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(Path.Combine(_repo.Root, file)))!;
        Assert.Equal(large, full["return"]!.GetValue<string>());
        Assert.Equal("Calc.Big", full["macro"]!.GetValue<string>());
    }

    [Fact]
    public async Task RunCountsANonAsciiReturnValueWithoutEscaping()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.Returned(new string('あ', 9_000));

        var outcome = await _session.RunAsync(new MacroName("Calc", "Kana"), [], TimeSpan.FromSeconds(5));

        Assert.Equal(new string('あ', 9_000), outcome.Result!["return"]!.GetValue<string>());
        Assert.False(outcome.Result.ContainsKey("result_file"));
    }

    [Fact]
    public async Task RunWritesANonAsciiResultFileWithoutEscaping()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.Returned(new string('あ', 30_000));

        var outcome = await _session.RunAsync(new MacroName("Calc", "Kana"), [], TimeSpan.FromSeconds(5));

        var text = File.ReadAllText(Path.Combine(_repo.Root, outcome.Result!["result_file"]!.GetValue<string>()));
        Assert.Contains(new string('あ', 30_000), text);
    }

    [Fact]
    public async Task RunDoesNotRunWhenTheSyncFails()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SyncThrows = new VbaSyncException("SYNC_FAILED", [], "refused");

        var outcome = await _session.RunAsync(new MacroName("Calc", "Twice"), [], TimeSpan.FromSeconds(5));

        AssertFailure(outcome, "EXCEL_ERROR", "SYNC_FAILED", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("RunMacro"));
    }

    [Fact]
    public async Task RunRejectsAMainWorkbookWithoutMacros()
    {
        _repo.AddApp("apps/plain", "Plain.xlsx");
        await _session.StartAsync(AppFolder.Resolve(_repo.Root, "apps/plain").Folder!);

        var outcome = await _session.RunAsync(new MacroName("Calc", "Twice"), [], TimeSpan.FromSeconds(5));

        AssertFailure(outcome, "INVALID_REQUEST", "NOT_MACRO_ENABLED", "ready");
    }

    [Fact]
    public async Task RunReportsAMissingMacro()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.NotFound("Calc has no public Sub or Function named Missing.");

        var outcome = await _session.RunAsync(new MacroName("Calc", "Missing"), [], TimeSpan.FromSeconds(5));

        AssertFailure(outcome, "NOT_FOUND", "MACRO_NOT_FOUND", "ready");
        Assert.Contains("named Missing", outcome.Failure!.Message);
    }

    [Fact]
    public async Task RunReportsTheRunTimeErrorTheGeneratedModuleCaught()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroResult = new MacroResult.Failed(11, "0 で除算しました。", "VBAProject");

        var outcome = await _session.RunAsync(new MacroName("Calc", "DivZero"), [], TimeSpan.FromSeconds(5));

        AssertFailure(outcome, "EXCEL_ERROR", "MACRO_ERROR", "ready");
        Assert.Equal(11, outcome.Failure!.Data!["number"]!.GetValue<int>());
        Assert.Equal("0 で除算しました。", outcome.Failure.Data["description"]!.GetValue<string>());
        Assert.Equal("VBAProject", outcome.Failure.Data["source"]!.GetValue<string>());
    }

    [Fact]
    public async Task RunEndsTheExcelWhenADialogStopsTheMacro()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroBlock = Block.UntilTerminated;
        _launcher.Process.Dialog = new DialogText("Microsoft Visual Basic", "実行時エラー '11':\n0 で除算しました。");

        var outcome = await _session.RunAsync(new MacroName("Calc", "DivZero"), [], TimeSpan.FromSeconds(30));

        AssertFailure(outcome, "STATE_UNKNOWN", "DIALOG", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(outcome.Failure!.Data!["excel_ended"]!.GetValue<bool>());
        Assert.Equal("Microsoft Visual Basic", outcome.Failure.Data["dialog"]!["title"]!.GetValue<string>());
        Assert.Equal("実行時エラー '11':\n0 で除算しました。", outcome.Failure.Data["dialog"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public async Task RunEndsTheExcelWhenTheLimitPasses()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroBlock = Block.UntilTerminated;

        var outcome = await _session.RunAsync(new MacroName("Calc", "Forever"), [], TimeSpan.FromMilliseconds(700));

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.True(_launcher.Process.DialogChecks >= 2);
    }

    [Fact]
    public async Task StatusReportsTheRunningMacro()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.MacroBlock = Block.UntilTerminated;

        var running = _session.RunAsync(new MacroName("Calc", "Forever"), [], TimeSpan.FromMilliseconds(600));
        await Task.Delay(300, TestContext.Current.CancellationToken);
        var status = _session.Status();
        await running;

        Assert.True(status.Busy);
        Assert.Equal("macro_run", status.Operation);
    }

    [Fact]
    public async Task RunWithoutASessionIsAConflict()
    {
        var outcome = await _session.RunAsync(new MacroName("Calc", "Twice"), [], TimeSpan.FromSeconds(5));

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    [Fact]
    public async Task EndWithoutASessionIsAConflict()
    {
        var outcome = await _session.EndAsync(save: false);

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    [Fact]
    public async Task StatusReportsTheRunningOperation()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveBlock = Block.UntilTerminated;

        var ending = _session.EndAsync(save: true);
        await Task.Delay(100, TestContext.Current.CancellationToken);
        var status = _session.Status();
        await ending;

        Assert.Equal("ready", status.State);
        Assert.True(status.Busy);
        Assert.Equal("session_end", status.Operation);
        Assert.True(status.Elapsed > TimeSpan.Zero);
        var busy = await _session.EndAsync(save: false);
        Assert.Equal("NO_SESSION", busy.Failure!.Detail);
    }

    [Fact]
    public async Task RejectsAnotherOperationWhileBusy()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.SaveBlock = Block.UntilTerminated;

        var ending = _session.EndAsync(save: true);
        await Task.Delay(100, TestContext.Current.CancellationToken);
        var second = await _session.EndAsync(save: false);
        await ending;

        AssertFailure(second, "CONFLICT", "BUSY", "ready", busy: true);
    }

    private static CellRange Range(string address) => CellRange.Parse(address).Range!;

    private Task<SessionOutcome> Range(string operation, string sheet) => operation == "read"
        ? _session.ReadAsync(sheet, Range("A1"), formulas: false)
        : _session.WriteAsync(sheet, Range("A1"), RangeWrite.Clear);

    private static void AssertFailure(SessionOutcome outcome, string code, string detail, string state, bool busy = false)
    {
        Assert.Null(outcome.Result);
        Assert.NotNull(outcome.Failure);
        Assert.Equal(code, outcome.Failure.Code);
        Assert.Equal(detail, outcome.Failure.Detail);
        Assert.Equal(state, outcome.Session.State);
        Assert.Equal(busy, outcome.Session.Busy);
        if (code == "STATE_UNKNOWN")
        {
            Assert.NotNull(outcome.Failure.Data?["excel_ended"]);
            Assert.NotNull(outcome.Failure.Data?["restartable"]);
        }
    }
}
