using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/api.md "表示要求のコマンド": shows a leftover dedicated Excel to the person and ends it after they close it,
// with fake process queries and windows.
public sealed class ExcelShowTests : IDisposable
{
    private static readonly DateTime Marked = new(2026, 11, 2, 1, 0, 0, DateTimeKind.Utc);
    private static readonly RecordedProcess Recorded = new(999, Marked.AddSeconds(1), "EXCEL");

    private readonly TempRepository _repo = new();
    private readonly FakeProcessQuery _processes = new();
    private readonly FakeWindows _windows = new();
    private readonly StringWriter _output = new();
    private readonly AppFolder _app;
    private readonly string _record;

    public ExcelShowTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
        _record = Path.Combine(_app.FullPath, ".xlwright", "session.json");
    }

    public void Dispose() => _repo.Dispose();

    [Fact]
    public void EndsWithoutCreatingARecordWhenThereIsNone()
    {
        Assert.Equal(0, Run(""));
        Assert.False(File.Exists(_record));
        Assert.Empty(_windows.Shown);
    }

    [Fact]
    public void StopsWhenAnotherConversationHoldsTheApp()
    {
        WriteOwned();
        using var held = new FileStream(_record, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        Assert.Equal(1, Run(""));
        Assert.Contains("in use", _output.ToString());
    }

    [Fact]
    public void StopsAtARecordNotInTheExpectedForm()
    {
        WriteRecord("{\"version\":1,");

        Assert.Equal(1, Run(""));
        Assert.Empty(_windows.Shown);
    }

    [Fact]
    public void EndsAtACompletedRecord()
    {
        WriteRecord(new OwnershipRecord(OwnershipRecord.Completed, Marked, "Sales.xlsm", null).ToJson());

        Assert.Equal(0, Run(""));
        Assert.Empty(_windows.Shown);
    }

    [Fact]
    public void ListsTheExcelStartedAfterAStartingMarkWithoutShowing()
    {
        WriteRecord(new OwnershipRecord(OwnershipRecord.Starting, Marked, "Sales.xlsm", null).ToJson());
        _processes.Automation.Add(new ProcessFacts(1234, "EXCEL.EXE", Marked.AddSeconds(2), "/automation"));

        Assert.Equal(1, Run(""));
        Assert.Contains("1234", _output.ToString());
        Assert.Empty(_windows.Shown);
    }

    [Fact]
    public void EndsWhenTheRecordedExcelIsGone()
    {
        WriteOwned();

        Assert.Equal(0, Run(""));
        Assert.Empty(_windows.Shown);
    }

    [Fact]
    public void StopsWhenTheRecordedExcelCannotBeChecked()
    {
        WriteOwned();
        _processes.Unchecked.Add(Recorded);

        Assert.Equal(1, Run(""));
        Assert.Empty(_windows.Shown);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("\"EXCEL.EXE\" C:\\Daily\\Book.xlsx")]
    public void LeavesAnExcelNotStartedFromCom(string? commandLine)
    {
        WriteOwned();
        _processes.Running.Add(Recorded);
        _processes.LeftoverCommandLine = commandLine;

        Assert.Equal(1, Run("done\n"));
        Assert.Empty(_windows.Shown);
        Assert.Equal(0, _processes.Leftover.Terminations);
    }

    [Fact]
    public void ShowsTheExcelAndEndsWhenThePersonClosedIt()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);
        _windows.OnShow = () => _processes.Leftover.Exit();

        Assert.Equal(0, Run("done\n"));
        Assert.Equal([999], _windows.Shown);
        Assert.Equal(0, _processes.Leftover.Terminations);
        AssertRecordUnchanged();
    }

    [Fact]
    public void EndsAnExcelLeftWithoutAWorkbookWindow()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);

        Assert.Equal(0, Run("done\n"));
        Assert.Equal(1, _processes.Leftover.Terminations);
        AssertRecordUnchanged();
    }

    [Fact]
    public void WaitsAgainWhileAWorkbookWindowIsOpen()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);
        _windows.WorkbookWindowAnswers.Enqueue(true);

        Assert.Equal(0, Run("done\ndone\n"));
        Assert.Contains("still open", _output.ToString());
        Assert.Equal(1, _processes.Leftover.Terminations);
    }

    [Fact]
    public void StopsWhenTheEndIsNotConfirmed()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);
        _processes.Leftover.TerminateEnds = false;

        Assert.Equal(1, Run("done\n"));
        AssertRecordUnchanged();
    }

    [Fact]
    public void LeavesTheExcelWhenInputCloses()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);

        Assert.Equal(1, Run("ok\n"));
        Assert.Equal(0, _processes.Leftover.Terminations);
    }

    [Fact]
    public void ReleasesTheAppAfterwards()
    {
        WriteOwned();
        _processes.Running.Add(Recorded);

        Run("done\n");

        using var again = new FileStream(_record, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    private int Run(string input) =>
        new ExcelShow(_processes, _windows, new StringReader(input), _output, TimeSpan.FromMilliseconds(200)).Run(_app);

    private void WriteOwned() => WriteRecord(new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", Recorded).ToJson());

    private void WriteRecord(string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_record)!);
        File.WriteAllText(_record, text);
    }

    private void AssertRecordUnchanged() =>
        Assert.Equal(new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", Recorded).ToJson(), File.ReadAllText(_record));
}

// Stands in for the workbook windows of a process: Shown holds the PIDs shown, and WorkbookWindowAnswers the
// answers to the checks after the person's input (false when it runs out).
internal sealed class FakeWindows : IWorkbookWindows
{
    public List<int> Shown { get; } = [];

    public Action? OnShow { get; set; }

    public Queue<bool> WorkbookWindowAnswers { get; } = new();

    public int Show(int pid)
    {
        Shown.Add(pid);
        OnShow?.Invoke();
        return 1;
    }

    public bool HasWorkbookWindow(int pid) => WorkbookWindowAnswers.Count > 0 && WorkbookWindowAnswers.Dequeue();
}
