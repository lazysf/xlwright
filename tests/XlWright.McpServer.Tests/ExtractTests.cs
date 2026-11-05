using System.IO.Compression;
using System.Text;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The vba_extract flow of docs/api.md "vba_extract", with a fake Excel behind the Core boundary.
public sealed class ExtractTests : IDisposable
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

    public ExtractTests()
    {
        _repo.AddApp("apps/old");
        var main = Path.Combine(_repo.Root, "apps", "old", "Old.xlsm");
        using (var zip = ZipFile.Open(main, ZipArchiveMode.Create))
        {
            zip.CreateEntry("xl/workbook.xml");
        }

        _app = AppFolder.Resolve(_repo.Root, "apps/old").Folder!;
        _session = new ExcelSession(_launcher, new FakeProcessQuery(), new FakeConnector(), Short);
    }

    public void Dispose()
    {
        _session.Dispose();
        _repo.Dispose();
    }

    private string ExtractFolder => Path.Combine(_app.FullPath, ".xlwright", "extract");

    [Fact]
    public async Task OpensARenamedCopyExportsWritesTheSourcesAndRemovesTheCopy()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.CopyModules = [("Calc", "standard", "Attribute VB_Name = \"Calc\"\r\n' 日本語\r\n"), ("Form1", "form", "")];

        var outcome = await _session.ExtractAsync(null);

        Assert.Null(outcome.Failure);
        Assert.Equal("ready", outcome.Session.State);
        Assert.Equal("Old.xlsm", outcome.Result!["workbook"]!.GetValue<string>());
        Assert.Equal("""[{"file":"src/modules/Calc.bas","type":"standard"}]""", outcome.Result["files"]!.ToJsonString());
        Assert.Equal("""[{"name":"Form1","type":"form"}]""", outcome.Result["skipped"]!.ToJsonString());

        var calls = _launcher.Excel.Calls.SkipWhile(c => !c.StartsWith("OpenExtractCopy", StringComparison.Ordinal)).ToList();
        Assert.Equal(2, calls.Count);
        var copy = calls[0]["OpenExtractCopy ".Length..];
        var folder = Path.GetDirectoryName(copy)!;
        Assert.Equal(ExtractFolder, Path.GetDirectoryName(folder));
        Assert.Equal(Path.GetFileName(folder) + ".xlsm", Path.GetFileName(copy));
        Assert.True(_launcher.Excel.ExtractCopyExisted);
        Assert.StartsWith("ExportExtractCopy " + folder, calls[1], StringComparison.Ordinal);
        Assert.False(Directory.Exists(folder));
        Assert.False(Directory.Exists(ExtractFolder));

        Assert.Equal("Attribute VB_Name = \"Calc\"\r\n' 日本語\r\n", File.ReadAllText(Path.Combine(_app.FullPath, "src", "modules", "Calc.bas"), Encoding.UTF8));
    }

    [Fact]
    public async Task StopsBeforeExcelWhenTheCheckFails()
    {
        await _session.StartAsync(_app);

        var outcome = await _session.ExtractAsync("data/None.xlsm");

        AssertFailure(outcome, "NOT_FOUND", "WORKBOOK_NOT_FOUND", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.Contains("ExtractCopy", StringComparison.Ordinal));
        Assert.False(Directory.Exists(ExtractFolder) && Directory.EnumerateFileSystemEntries(ExtractFolder).Any());
    }

    [Fact]
    public async Task ReportsACopyExcelCouldNotOpenAndRemovesIt()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExtractOpenThrows = new System.Runtime.InteropServices.COMException("cannot open");

        var outcome = await _session.ExtractAsync(null);

        AssertFailure(outcome, "EXCEL_ERROR", "OPEN_FAILED", "ready");
        Assert.DoesNotContain(_launcher.Excel.Calls, c => c.StartsWith("ExportExtractCopy", StringComparison.Ordinal));
        Assert.False(Directory.Exists(ExtractFolder));
    }

    [Theory]
    [InlineData("VBA_ACCESS_NOT_TRUSTED")]
    [InlineData("VBA_PROJECT_LOCKED")]
    public async Task ReportsWhatExcelRefusedForTheProject(string detail)
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExportThrows = new VbaSyncException(detail, [], "refused");

        var outcome = await _session.ExtractAsync(null);

        AssertFailure(outcome, "CONFLICT", detail, "ready");
        Assert.False(Directory.Exists(ExtractFolder));
    }

    [Fact]
    public async Task ReportsAFailedExportAsExtractFailedWithNothingDone()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExportThrows = new System.Runtime.InteropServices.COMException("export failed");

        var outcome = await _session.ExtractAsync(null);

        AssertFailure(outcome, "EXCEL_ERROR", "EXTRACT_FAILED", "ready");
        Assert.Equal("[]", outcome.Failure!.Data!["done"]!.ToJsonString());
        Assert.False(Directory.Exists(Path.Combine(_app.FullPath, "src")));
    }

    [Fact]
    public async Task OpenTimeoutTerminatesTheOwnedExcel()
    {
        await _session.StartAsync(_app);
        _launcher.Excel.ExtractOpenBlock = Block.UntilTerminated;

        var outcome = await _session.ExtractAsync(null);

        AssertFailure(outcome, "STATE_UNKNOWN", "TIMEOUT", "closed");
        Assert.Equal(1, _launcher.Process.Terminations);
    }

    [Fact]
    public async Task WithoutASessionIsAConflict()
    {
        var outcome = await _session.ExtractAsync(null);

        AssertFailure(outcome, "CONFLICT", "NO_SESSION", "closed");
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
