using System.Text;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/api.md "test_run" and docs/architecture.md "テスト": the test copy flow, with a fake Excel behind the Core boundary.
public sealed class TestRunTests : IDisposable
{
    private static readonly SessionTimeouts Short = new(
        Start: TimeSpan.FromMilliseconds(500),
        Sync: TimeSpan.FromMilliseconds(500),
        Save: TimeSpan.FromMilliseconds(500),
        QuitExit: TimeSpan.FromMilliseconds(300),
        CleanupQuitExit: TimeSpan.FromMilliseconds(300),
        TerminateExit: TimeSpan.FromMilliseconds(200),
        PendingReturn: TimeSpan.FromMilliseconds(300))
    {
        CompileWatch = TimeSpan.FromMilliseconds(600),
    };

    private static readonly TimeSpan Limit = TimeSpan.FromSeconds(5);

    private readonly TempRepository _repo = new();
    private readonly FakeLauncher _launcher = new();
    private readonly ExcelSession _session;
    private readonly AppFolder _app;

    public TestRunTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm", "data/note.txt", "src/modules/Calc.bas", ".git/HEAD");
        Source("tests/CalcTests.bas", "Public Sub TestAdd()\nEnd Sub\nPublic Sub TestSplit()\nEnd Sub\nPublic Sub TestLoad()\nEnd Sub\n");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _session = new ExcelSession(_launcher, new FakeProcessQuery(), new FakeConnector(), Short);
    }

    private FakeExcel Excel => _launcher.Excel;

    private string TestFolder => Path.Combine(_app.FullPath, ".xlwright", "test");

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    [Fact]
    public async Task RunsEachTestInTheCopyAndReopensTheMainWorkbook()
    {
        await _session.StartAsync(_app);
        Excel.MacroResults["CalcTests.TestSplit"] = new MacroResult.Failed(-2147221503, "AreEqual: expected 3, actual 2. 列の数", VbaSources.AssertName);
        Excel.MacroResults["CalcTests.TestLoad"] = new MacroResult.Failed(53, "ファイルが見つかりません。", "VBAProject");

        var outcome = await _session.TestRunAsync(null, Limit);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        var result = outcome.Result!;
        Assert.Equal("""{"passed":1,"failed":1,"errors":1}""", result["summary"]!.ToJsonString());
        var tests = result["tests"]!.AsArray();
        Assert.Equal(["CalcTests.TestAdd", "CalcTests.TestSplit", "CalcTests.TestLoad"], tests.Select(t => t!["name"]!.GetValue<string>()));
        Assert.Equal(["passed", "failed", "error"], tests.Select(t => t!["result"]!.GetValue<string>()));
        Assert.Equal("AreEqual: expected 3, actual 2. 列の数", tests[1]!["message"]!.GetValue<string>());
        Assert.Equal(53, tests[2]!["number"]!.GetValue<int>());
        Assert.Equal("ファイルが見つかりません。", tests[2]!["description"]!.GetValue<string>());
        Assert.Equal("VBAProject", tests[2]!["source"]!.GetValue<string>());
        Assert.NotNull(result["elapsed_ms"]);
        Assert.NotNull(result["sync"]!["modules"]);
        Assert.Null(result["leftover_copy"]);

        var calls = Excel.Calls.SkipWhile(c => !c.StartsWith("SyncVba")).Skip(1).ToList();
        var copy = calls[2]["Open ".Length..];
        Assert.Equal(Path.Combine(TestFolder, Path.GetFileName(Path.GetDirectoryName(copy)!), "Sales.xlsm"), copy);
        Assert.Equal(
            ["Save", "CloseMain", "Open " + copy, $"ImportModules CalcTests,{VbaSources.AssertName}", "StartCompile", "IsCompiled",
             "RunMacro CalcTests.TestAdd", "RunMacro CalcTests.TestSplit", "RunMacro CalcTests.TestLoad",
             "CloseUnder " + Path.GetDirectoryName(copy), "Open " + _app.WorkbookPath],
            calls);
        Assert.Empty(Directory.EnumerateFileSystemEntries(TestFolder));
        Assert.False(Directory.Exists(Path.Combine(_app.FullPath, ".xlwright", "sync")) &&
            Directory.EnumerateFileSystemEntries(Path.Combine(_app.FullPath, ".xlwright", "sync")).Any());
    }

    [Fact]
    public async Task CopiesTheAppFolderWithoutSourcesTestsAndDotFolders()
    {
        await _session.StartAsync(_app);
        List<string>? copied = null;
        Excel.OnOpen = path =>
        {
            if (path != _app.WorkbookPath)
            {
                var folder = Path.GetDirectoryName(path)!;
                copied = Directory.EnumerateFileSystemEntries(folder, "*", SearchOption.AllDirectories)
                    .Select(p => Path.GetRelativePath(folder, p).Replace('\\', '/')).Order(StringComparer.Ordinal).ToList();
            }
        };

        await _session.TestRunAsync(null, Limit);

        Assert.Equal(["Sales.xlsm", "data", "data/note.txt"], copied);
    }

    [Fact]
    public async Task RunsOnlyTheNamedTests()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.TestRunAsync(["CalcTests.TestLoad"], Limit);

        Assert.Equal(["CalcTests.TestLoad"], outcome.Result!["tests"]!.AsArray().Select(t => t!["name"]!.GetValue<string>()));
        Assert.Single(Excel.Calls, c => c.StartsWith("RunMacro"));
    }

    [Fact]
    public async Task StopsAtAMissingTestWithoutTouchingExcel()
    {
        await _session.StartAsync(_app);
        var before = Excel.Calls.Count;

        var outcome = await _session.TestRunAsync(["CalcTests.TestNone"], Limit);

        AssertFailure(outcome, "NOT_FOUND", "TEST_NOT_FOUND", "ready");
        Assert.Equal(before, Excel.Calls.Count);
    }

    [Fact]
    public async Task ASaveFailureRunsNoTestAndKeepsTheMainWorkbookOpen()
    {
        await _session.StartAsync(_app);
        Excel.SaveThrows = new InvalidOperationException("disk full");

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "EXCEL_ERROR", "SAVE_FAILED", "ready");
        Assert.DoesNotContain("CloseMain", Excel.Calls);
    }

    [Fact]
    public async Task ACopyThatDoesNotOpenReopensTheMainWorkbook()
    {
        await _session.StartAsync(_app);
        Excel.OpenFails = path => path == _app.WorkbookPath ? null : new InvalidOperationException("cannot open the copy");

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "EXCEL_ERROR", "OPEN_FAILED", "ready");
        Assert.Equal("Open " + _app.WorkbookPath, Excel.Calls[^1]);
        Assert.Empty(Directory.EnumerateFileSystemEntries(TestFolder));
    }

    [Fact]
    public async Task AFolderThatCannotBeCopiedReopensTheMainWorkbook()
    {
        await _session.StartAsync(_app);
        using var held = new FileStream(Path.Combine(_app.FullPath, "data", "note.txt"), FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "INTERNAL_ERROR", "COPY_FAILED", "ready");
        Assert.Equal("Open " + _app.WorkbookPath, Excel.Calls[^1]);
        Assert.DoesNotContain(Excel.Calls, c => c.StartsWith("RunMacro"));
    }

    [Fact]
    public async Task AnImportFailureReopensTheMainWorkbook()
    {
        await _session.StartAsync(_app);
        Excel.ImportThrows = new VbaSyncException("SYNC_FAILED", ["CalcTests"], "XlWrightAssert was imported as XlWrightAssert1.");

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "EXCEL_ERROR", "SYNC_FAILED", "ready");
        Assert.Equal("Open " + _app.WorkbookPath, Excel.Calls[^1]);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ACompileErrorRunsNoTestAndReopensTheMainWorkbook(bool placeKnown)
    {
        await _session.StartAsync(_app);
        _launcher.Process.Dialog = new DialogText("Microsoft Visual Basic for Applications", "コンパイル エラー:\nSub または Function が定義されていません。");
        Excel.ErrorPlace = placeKnown ? new CompileErrorPlace("CalcTests", 2) : null;

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "EXCEL_ERROR", "COMPILE_ERROR", "ready");
        var error = outcome.Failure!.Data!["error"]!;
        Assert.Equal("コンパイル エラー: Sub または Function が定義されていません。", error["message"]!.GetValue<string>());
        Assert.Equal(placeKnown ? "CalcTests" : null, error["module"]?.GetValue<string>());
        Assert.Equal(placeKnown ? 2 : null, error["line"]?.GetValue<int>());
        Assert.Equal(1, _launcher.Process.DialogCloses);
        Assert.Contains("HideEditor", Excel.Calls);
        Assert.DoesNotContain(Excel.Calls, c => c.StartsWith("RunMacro", StringComparison.Ordinal));
        Assert.Equal("Open " + _app.WorkbookPath, Excel.Calls[^1]);
        Assert.Empty(Directory.EnumerateFileSystemEntries(TestFolder));
        Assert.Null((await _session.TestRunAsync(null, Limit)).Failure);
    }

    [Fact]
    public async Task AnAlreadyCompiledCopyRunsTheTests()
    {
        await _session.StartAsync(_app);
        Excel.CompileStart = CompileStart.AlreadyCompiled;

        var outcome = await _session.TestRunAsync(null, Limit);

        Assert.Null(outcome.Failure);
        Assert.DoesNotContain("IsCompiled", Excel.Calls);
        Assert.Equal(3, outcome.Result!["tests"]!.AsArray().Count);
    }

    [Fact]
    public async Task ACompileNotConfirmedStillRunsTheTests()
    {
        await _session.StartAsync(_app);
        Excel.CompilesAfter = int.MaxValue;

        var outcome = await _session.TestRunAsync(null, Limit);

        Assert.Null(outcome.Failure);
        Assert.Equal(3, outcome.Result!["tests"]!.AsArray().Count);
    }

    [Fact]
    public async Task ACompileDialogThatStaysEndsTheExcelAndRemovesTheCopy()
    {
        await _session.StartAsync(_app);
        _launcher.Process.Dialog = new DialogText("Microsoft Visual Basic for Applications", "コンパイル エラー:");
        _launcher.Process.CloseLeavesDialog = true;

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "STATE_UNKNOWN", "DIALOG", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.DoesNotContain(Excel.Calls, c => c.StartsWith("RunMacro", StringComparison.Ordinal));
        Assert.Empty(Directory.EnumerateFileSystemEntries(TestFolder));
    }

    [Fact]
    public async Task AMainWorkbookThatReopensReadOnlyEndsTheSession()
    {
        await _session.StartAsync(_app);
        var opens = 0;
        Excel.OpensReadOnlyAt = path => path == _app.WorkbookPath && ++opens > 0;

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "CONFLICT", "WORKBOOK_NOT_REOPENED", "closed");
        Assert.Equal(3, outcome.Failure!.Data!["tests"]!.AsArray().Count);
        Assert.Contains("CloseAllAndQuit", Excel.Calls);
    }

    [Fact]
    public async Task ADialogInATestEndsTheExcelAndReportsTheTestsSoFar()
    {
        await _session.StartAsync(_app);
        Excel.MacroBlocks["CalcTests.TestSplit"] = Block.UntilTerminated;
        _ = Task.Run(async () =>
        {
            while (!Excel.Calls.Contains("RunMacro CalcTests.TestSplit")) await Task.Delay(20);
            _launcher.Process.Dialog = new DialogText("Microsoft Excel", "止まった");
        });

        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "STATE_UNKNOWN", "DIALOG", "closed");
        var data = outcome.Failure!.Data!;
        Assert.Equal("止まった", data["dialog"]!["text"]!.GetValue<string>());
        Assert.Equal("CalcTests.TestSplit", data["test"]!.GetValue<string>());
        Assert.Equal(["CalcTests.TestAdd"], data["tests"]!.AsArray().Select(t => t!["name"]!.GetValue<string>()));
        Assert.Equal(1, _launcher.Process.Terminations);
        Assert.DoesNotContain("RunMacro CalcTests.TestLoad", Excel.Calls);
    }

    [Fact]
    public async Task ATestPastTheLimitEndsTheExcel()
    {
        await _session.StartAsync(_app);
        Excel.MacroBlocks["CalcTests.TestAdd"] = Block.UntilTerminated;

        var outcome = await _session.TestRunAsync(null, TimeSpan.FromMilliseconds(400));

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal("CalcTests.TestAdd", outcome.Failure!.Data!["test"]!.GetValue<string>());
        Assert.Empty(outcome.Failure.Data["tests"]!.AsArray());
    }

    [Fact]
    public async Task StatusShowsTheRunningTest()
    {
        await _session.StartAsync(_app);
        var gate = new ManualResetEventSlim();
        Excel.MacroGate = ("CalcTests.TestSplit", gate);

        var running = _session.TestRunAsync(null, Limit);
        while (!Excel.Calls.Contains("RunMacro CalcTests.TestSplit")) await Task.Delay(20);
        var status = _session.Status();
        gate.Set();
        await running;

        Assert.Equal("test_run", status.Operation);
        Assert.Equal(new TestProgress("CalcTests.TestSplit", 2, 3), status.Test);
        Assert.Null(_session.Status().Test);
    }

    [Fact]
    public async Task ACopyThatCannotBeRemovedIsReported()
    {
        await _session.StartAsync(_app);
        FileStream? held = null;
        Excel.OnOpen = path =>
        {
            if (path != _app.WorkbookPath)
            {
                held = new FileStream(Path.Combine(Path.GetDirectoryName(path)!, "data", "note.txt"), FileMode.Open, FileAccess.Read, FileShare.None);
            }
        };

        var outcome = await _session.TestRunAsync(null, Limit);
        held!.Dispose();

        Assert.Null(outcome.Failure);
        Assert.StartsWith("apps/sales/.xlwright/test/", outcome.Result!["leftover_copy"]!.GetValue<string>());
    }

    [Fact]
    public async Task ManyResultsGoToAResultFileAndTheResponseKeepsTheFailures()
    {
        Source("tests/CalcTests.bas", string.Concat(Enumerable.Range(1, 300).Select(i => $"Public Sub Test{i:000}()\nEnd Sub\n")));
        await _session.StartAsync(_app);
        Excel.MacroResult = new MacroResult.Failed(-2147221503, "AreEqual: " + new string('x', 100), VbaSources.AssertName);
        Excel.MacroResults["CalcTests.Test001"] = new MacroResult.Returned(null);

        var outcome = await _session.TestRunAsync(null, Limit);

        var result = outcome.Result!;
        Assert.Equal("""{"passed":1,"failed":299,"errors":0}""", result["summary"]!.ToJsonString());
        var file = Path.Combine(_repo.Root, result["result_file"]!.GetValue<string>());
        Assert.Equal(300, System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(file))!["tests"]!.AsArray().Count);
        var kept = result["tests"]!.AsArray();
        Assert.InRange(kept.Count, 1, 298);
        Assert.All(kept, t => Assert.Equal("failed", t!["result"]!.GetValue<string>()));
        Assert.True(result.ToJsonString(JsonText.Options).Length < 20_000);
    }

    [Fact]
    public async Task TestRunWithoutASessionIsAConflict()
    {
        var outcome = await _session.TestRunAsync(null, Limit);

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
    }

    private void Source(string file, string text)
    {
        var path = Path.Combine(_repo.Root, "apps", "sales", file);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, text, new UTF8Encoding(false));
    }

    private static void AssertFailure(SessionOutcome outcome, string code, string detail, string state)
    {
        Assert.Null(outcome.Result);
        Assert.NotNull(outcome.Failure);
        Assert.Equal(code, outcome.Failure.Code);
        Assert.Equal(detail, outcome.Failure.Detail);
        Assert.Equal(state, outcome.Session.State);
        Assert.False(outcome.Session.Busy);
    }
}
