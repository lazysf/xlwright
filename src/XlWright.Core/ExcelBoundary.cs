namespace XlWright.Core;

// The boundary between the session rules and Excel (docs/architecture.md "本体の部品の境界").
// Every IExcelLauncher and IExcel call runs on the session's dedicated STA thread.

/// <summary>Starts a dedicated Excel. Returns before ownership is confirmed; the session confirms it.</summary>
public interface IExcelLauncher
{
    LaunchedExcel Launch();
}

/// <summary>The Excel that one session started, and the process that Hwnd points to.</summary>
public sealed record LaunchedExcel(IExcel Excel, IOwnedProcess Process);

/// <summary>Excel operations for one session. Implementations release every COM object they obtain.</summary>
public interface IExcel
{
    /// <summary>Sets UserControl to true, so releasing the references does not end the Excel.</summary>
    void TakeControl();

    /// <summary>Opens the workbook with AutomationSecurity=Low and events off; returns Workbook.ReadOnly.</summary>
    bool Open(string path);

    void Save();

    /// <summary>
    /// Recalculates and waits for CalculationState xlDone, then reads Value2 of the main workbook's range as double,
    /// string, bool, null, or <see cref="CellError"/>; null when no sheet has that name.
    /// </summary>
    object?[,]? ReadValues(string sheet, CellRange range);

    /// <summary>Reads Formula (English function names, A1); null when no sheet has that name.</summary>
    string[,]? ReadFormulas(string sheet, CellRange range);

    /// <summary>Writes Value2, Formula, or ClearContents with events off; false when no sheet has that name.</summary>
    bool Write(string sheet, CellRange range, RangeWrite write);

    /// <summary>The full paths of open workbooks and protected view windows; null for a workbook without a path.</summary>
    IReadOnlyList<string?> OpenWorkbookPaths();

    /// <summary>
    /// Removes every standard and class module of the main workbook, then imports <paramref name="modules"/> in order
    /// and checks each imported name and type. Returns the removed names that are not in <paramref name="modules"/>.
    /// Throws <see cref="VbaSyncException"/> for the failures docs/api.md "vba_sync" names.
    /// </summary>
    IReadOnlyList<string> SyncVba(IReadOnlyList<VbaModule> modules);

    /// <summary>Closes every workbook without saving, asks Excel to quit, and releases the references.</summary>
    void CloseAllAndQuit();

    /// <summary>Releases the references without closing anything, when the session leaves the Excel running.</summary>
    void ReleaseReferences();
}

/// <summary>
/// The process found from Hwnd, held through one handle for identity, waiting, and termination
/// (docs/architecture.md "完了にする条件").
/// </summary>
public interface IOwnedProcess : IDisposable
{
    ProcessFacts Facts { get; }

    bool WaitForExit(TimeSpan timeout);

    /// <summary>Terminates the process through the handle opened when it was found.</summary>
    void Terminate();
}

/// <summary>A vba_sync failure: the detail of docs/api.md "vba_sync" and the modules imported before it.</summary>
public sealed class VbaSyncException(string detail, IReadOnlyList<string> done, string message) : Exception(message)
{
    public string Detail { get; } = detail;

    public IReadOnlyList<string> Done { get; } = done;
}

/// <summary>What ownership is decided from (docs/architecture.md "所有の確定").</summary>
public sealed record ProcessFacts(int Pid, string ImageName, DateTime CreatedUtc, string? CommandLine)
{
    /// <summary>Returns null when the process is the Excel this session started, or the reason it is not.</summary>
    public string? OwnershipProblem(DateTime requestedUtc)
    {
        if (string.IsNullOrEmpty(ImageName))
        {
            return "the process could not be identified";
        }

        if (!string.Equals(Path.GetFileName(ImageName), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
        {
            return $"the process is {Path.GetFileName(ImageName)}, not EXCEL.EXE";
        }

        if (CreatedUtc <= requestedUtc)
        {
            return "the process was created before the start was requested";
        }

        if (CommandLine is null)
        {
            return "the command line could not be read";
        }

        return CommandLine.Contains("/automation", StringComparison.OrdinalIgnoreCase)
            ? null
            : "the command line does not contain /automation";
    }
}

/// <summary>Time limits (docs/architecture.md "期限と大きさの上限").</summary>
public sealed record SessionTimeouts(
    TimeSpan Start,
    TimeSpan Sync,
    TimeSpan Save,
    TimeSpan QuitExit,
    TimeSpan TerminateExit,
    TimeSpan PendingReturn)
{
    public static SessionTimeouts Default { get; } = new(
        Start: TimeSpan.FromSeconds(120),
        Sync: TimeSpan.FromSeconds(30),
        Save: TimeSpan.FromSeconds(120),
        QuitExit: TimeSpan.FromSeconds(75),
        TerminateExit: TimeSpan.FromSeconds(5),
        PendingReturn: TimeSpan.FromSeconds(10));
}
