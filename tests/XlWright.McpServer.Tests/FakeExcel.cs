using XlWright.Core;

namespace XlWright.McpServer.Tests;

// Stands in for Excel behind the Core boundary. Each call is recorded; a call can be made to throw or to block
// until the fake process is terminated (or forever, to leave a pending call).
internal sealed class FakeLauncher : IExcelLauncher
{
    public FakeLauncher() => Excel.Process = Process;

    public FakeExcel Excel { get; } = new();

    public FakeProcess Process { get; } = new();

    public Block LaunchBlock { get; set; } = Block.None;

    public int Launches { get; private set; }

    public LaunchedExcel Launch()
    {
        Launches++;
        Process.WaitIf(LaunchBlock);
        return new LaunchedExcel(Excel, Process);
    }
}

internal enum Block
{
    None,
    UntilTerminated,
    Forever,
}

internal sealed class FakeExcel : IExcel
{
    private readonly List<string> _calls = [];

    public FakeProcess Process { get; set; } = null!;

    public IReadOnlyList<string> Calls
    {
        get { lock (_calls) return _calls.ToList(); }
    }

    public bool OpensReadOnly { get; set; }

    public Exception? OpenThrows { get; set; }

    public Exception? SaveThrows { get; set; }

    public Block OpenBlock { get; set; } = Block.None;

    public Block SaveBlock { get; set; } = Block.None;

    public List<string?> ExtraWorkbooks { get; } = [];

    public bool QuitEndsProcess { get; set; } = true;

    private string? _main;

    public void TakeControl() => Record("TakeControl");

    public bool Open(string path)
    {
        Record("Open " + path);
        Process.WaitIf(OpenBlock);
        if (OpenThrows is not null) throw OpenThrows;
        _main = path;
        return OpensReadOnly;
    }

    public void Save()
    {
        Record("Save");
        Process.WaitIf(SaveBlock);
        if (SaveThrows is not null) throw SaveThrows;
    }

    public IReadOnlyList<string?> OpenWorkbookPaths()
    {
        Record("OpenWorkbookPaths");
        return (_main is null ? [] : new[] { _main }).Concat(ExtraWorkbooks).ToList();
    }

    public Block QuitBlock { get; set; } = Block.None;

    public void CloseAllAndQuit()
    {
        Record("CloseAllAndQuit");
        Process.WaitIf(QuitBlock);
        if (QuitEndsProcess) Process.Exit();
    }

    public void ReleaseReferences() => Record("ReleaseReferences");

    public List<string> Sheets { get; } = ["Sheet1", "Data"];

    public object?[,] Values { get; set; } = new object?[1, 1];

    public string[,] Formulas { get; set; } = new string[1, 1];

    public Exception? RangeThrows { get; set; }

    public Block RangeBlock { get; set; } = Block.None;

    public RangeWrite? Written { get; private set; }

    public object?[,]? ReadValues(string sheet, CellRange range) => Range("ReadValues", sheet, range) ? Values : null;

    public string[,]? ReadFormulas(string sheet, CellRange range) => Range("ReadFormulas", sheet, range) ? Formulas : null;

    public bool Write(string sheet, CellRange range, RangeWrite write)
    {
        if (!Range("Write", sheet, range)) return false;
        Written = write;
        return true;
    }

    public List<string> Removed { get; } = [];

    public Exception? SyncThrows { get; set; }

    public Block SyncBlock { get; set; } = Block.None;

    public IReadOnlyList<VbaModule>? Synced { get; private set; }

    public IReadOnlyList<string> SyncVba(IReadOnlyList<VbaModule> modules)
    {
        Record("SyncVba " + string.Join(",", modules.Select(m => m.Name)));
        Synced = modules;
        Process.WaitIf(SyncBlock);
        if (SyncThrows is not null) throw SyncThrows;
        return Removed;
    }

    private bool Range(string call, string sheet, CellRange range)
    {
        Record($"{call} {sheet}!{range.Address}");
        Process.WaitIf(RangeBlock);
        if (RangeThrows is not null) throw RangeThrows;
        return Sheets.Contains(sheet, StringComparer.OrdinalIgnoreCase);
    }

    private void Record(string call)
    {
        lock (_calls) _calls.Add(call);
    }
}

internal sealed class FakeProcess : IOwnedProcess
{
    private readonly ManualResetEventSlim _exited = new();

    public ProcessFacts Facts { get; set; } = new(4242, "EXCEL.EXE", DateTime.MaxValue, "\"EXCEL.EXE\" /automation -Embedding");

    public bool TerminateEnds { get; set; } = true;

    public int Terminations { get; private set; }

    public void Exit() => _exited.Set();

    public bool WaitForExit(TimeSpan timeout) => _exited.Wait(timeout);

    public void Terminate()
    {
        Terminations++;
        if (TerminateEnds) _exited.Set();
    }

    public void WaitIf(Block block)
    {
        if (block == Block.UntilTerminated) _exited.Wait();
        if (block == Block.Forever) Thread.Sleep(Timeout.Infinite);
    }

    public void Dispose()
    {
    }
}
