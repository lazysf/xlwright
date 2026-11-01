using System.Diagnostics;
using XlWright.Core;

namespace XlWright.McpServer.Tests.Core;

// Records the order of driver calls. Before runs first, so a test can make a call fail or hang.
internal sealed class FakeExcelDriver : IExcelDriver
{
    private readonly Dictionary<string, object?> cells = [];
    internal List<string> Calls { get; } = [];
    internal Action<string>? Before { get; set; }
    internal bool ExitConfirmed { get; set; } = true;

    private void Called(string name)
    {
        lock (Calls) Calls.Add(name);
        Before?.Invoke(name);
    }

    public void Generate() => Called(nameof(Generate));

    public ProcessIdentity Identify()
    {
        Called(nameof(Identify));
        using var current = Process.GetCurrentProcess();
        return ProcessIdentity.Read(current);
    }

    public void Open(string workbookPath) => Called(nameof(Open));

    public CellContent Read(string sheet, string address, bool values, bool formulas)
    {
        Called(nameof(Read));
        return new(cells.GetValueOrDefault(address), null);
    }

    public void Write(string sheet, string address, CellWrite write)
    {
        Called(nameof(Write));
        cells[address] = write.Value;
    }

    public void Save() => Called(nameof(Save));

    public void VerifyEndSafety(string workbookPath, bool requireEmpty) => Called(requireEmpty ? "VerifyEmpty" : "Verify");

    public void Close() => Called(nameof(Close));

    public void Quit() => Called(nameof(Quit));

    public bool ReleaseAndWaitForExit(TimeSpan timeout)
    {
        Called("WaitForExit");
        return ExitConfirmed;
    }
}
