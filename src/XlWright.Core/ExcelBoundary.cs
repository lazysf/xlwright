namespace XlWright.Core;

// The boundary between the session rules and Excel (docs/architecture.md "本体の部品の境界").
// Every IExcelLauncher and IExcel call runs on the session's dedicated STA thread.

/// <summary>Starts a dedicated Excel. Returns before ownership is confirmed; the session confirms it.</summary>
public interface IExcelLauncher
{
    LaunchedExcel Launch();
}

/// <summary>
/// Reads OS processes for the ownership record (docs/architecture.md "完了にする条件"). It never touches COM.
/// </summary>
public interface IProcessQuery
{
    /// <summary>
    /// Whether the recorded process still runs. Gone only when the OS says so: no process has the PID, its creation
    /// time differs (a reused PID), or it has exited. A process that cannot be read is Unknown, never Gone.
    /// </summary>
    ProcessLookup Find(RecordedProcess process);

    /// <summary>The EXCEL processes created after the time whose command line contains /automation or cannot be read.</summary>
    IReadOnlyList<ProcessFacts> AutomationExcelsSince(DateTime utc);
}

public enum ProcessPresence
{
    Gone,
    Running,
    Unknown,
}

/// <summary>What the OS says about a recorded process; Process is the one held through a handle when it runs.</summary>
public sealed record ProcessLookup(ProcessPresence Presence, IOwnedProcess? Process = null);

/// <summary>
/// Connects to a leftover dedicated Excel through its windows (docs/architecture.md "残った専用のExcelの片付け" 3).
/// </summary>
public interface IExcelConnector
{
    /// <summary>
    /// Runs on the calling STA thread, and the Excel it returns stays on that thread. It reads windows first and
    /// connects only through a workbook window (EXCEL7) of the process.
    /// </summary>
    LeftoverWindow Connect(int pid);
}

/// <summary>What the windows of a leftover Excel gave.</summary>
public abstract record LeftoverWindow
{
    private LeftoverWindow()
    {
    }

    /// <summary>The process has no workbook window, so it holds no workbook to lose (docs/notes/leftover-excel.md).</summary>
    public sealed record NoWorkbookWindow : LeftoverWindow;

    /// <summary>No workbook window gave an Excel whose Hwnd belongs to the process.</summary>
    public sealed record NotReachable(string Reason) : LeftoverWindow;

    /// <summary>The Excel of the process, for listing the open workbooks and closing them.</summary>
    public sealed record Connected(IExcel Excel) : LeftoverWindow;
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
    /// Adds a new workbook and saves it as <paramref name="path"/> in <paramref name="format"/> (52 .xlsm, 51 .xlsx); it
    /// becomes the main workbook (docs/api.md "workbook_create"). When the save fails, it closes the new workbook
    /// without saving and throws.
    /// </summary>
    void Create(string path, int format);

    /// <summary>Workbook.Saved of the main workbook: false when it has changes that are not saved (docs/api.md "workbook_get").</summary>
    bool IsSaved();

    /// <summary>
    /// Saves the main workbook as <paramref name="path"/> in <paramref name="format"/>; the main workbook is that file
    /// from then on (docs/api.md "workbook_edit"). When the save fails, the main workbook stays the file it was, and it throws.
    /// </summary>
    void SaveAs(string path, int format);

    /// <summary>Closes the main workbook without saving, with events off; test_run saves it first.</summary>
    void CloseMain();

    /// <summary>
    /// Closes without saving, with events and alerts off, every open workbook whose full path is inside
    /// <paramref name="folder"/>: the test copy and what the tests opened from it (docs/api.md "test_run").
    /// </summary>
    void CloseUnder(string folder);

    /// <summary>
    /// Imports <paramref name="modules"/> into the open main workbook (the test copy during test_run) and checks each
    /// imported name and type, as <see cref="SyncVba"/> does. Throws <see cref="VbaSyncException"/> with SYNC_FAILED.
    /// </summary>
    void ImportModules(IReadOnlyList<VbaModule> modules);

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
    /// Checks that each of <paramref name="documents"/> is a sheet or workbook module of the main workbook, removes
    /// every standard and class module, imports <paramref name="modules"/> in order and checks each imported name and
    /// type, then replaces the code of each document module. Returns the removed names that are not in
    /// <paramref name="modules"/>. Throws <see cref="VbaSyncException"/> for the failures docs/api.md "vba_sync" names.
    /// </summary>
    IReadOnlyList<string> SyncVba(IReadOnlyList<VbaModule> modules, IReadOnlyList<VbaDocumentModule> documents);

    /// <summary>
    /// Lists the modules of the main workbook in VBComponents order. Throws <see cref="VbaSyncException"/> with
    /// VBA_ACCESS_NOT_TRUSTED or VBA_PROJECT_LOCKED, as vba_sync does.
    /// </summary>
    IReadOnlyList<VbaComponent> ListVba();

    /// <summary>
    /// Opens the copy for vba_extract with AutomationSecurity=ForceDisable and events off, then sets both back. The copy's
    /// macros stay disabled after that (docs/notes/vba-extract.md). The main workbook is not touched.
    /// </summary>
    void OpenExtractCopy(string path);

    /// <summary>
    /// Exports the standard, class, and document modules of the copy into <paramref name="folder"/>, lists the forms
    /// without exporting them, and closes the copy without saving, also when it fails. Throws
    /// <see cref="VbaSyncException"/> with VBA_ACCESS_NOT_TRUSTED or VBA_PROJECT_LOCKED.
    /// </summary>
    IReadOnlyList<ExportedModule> ExportExtractCopy(string folder);

    /// <summary>
    /// Lists the references of the main workbook's VBA project in References order (docs/api.md "vba_getの参照設定").
    /// Throws <see cref="VbaSyncException"/> with VBA_ACCESS_NOT_TRUSTED or VBA_PROJECT_LOCKED, as vba_sync does.
    /// </summary>
    IReadOnlyList<VbaReference> ListReferences();

    /// <summary>
    /// Adds a reference with References.AddFromGuid or removes the one with the GUID (docs/api.md "vba_edit"). Throws
    /// <see cref="VbaSyncException"/> as <see cref="ListReferences"/> does, and the exception Excel gave when it refuses.
    /// </summary>
    ReferenceChange EditReference(ReferenceEdit edit);

    /// <summary>
    /// Makes the main workbook's project the ActiveVBProject and executes the VBE's "Compile VBAProject" control
    /// (docs/architecture.md "テスト", docs/notes/test-copy.md). It returns at once; a compile error shows a dialog later.
    /// </summary>
    CompileStart StartCompile();

    /// <summary>Whether the compile control is disabled, which it is once the project is compiled.</summary>
    bool IsCompiled();

    /// <summary>The module and line the VBE selected for a compile error, or null when it cannot be read.</summary>
    CompileErrorPlace? ReadErrorPlace();

    /// <summary>Hides the VBE window that a compile error shows on the person's screen.</summary>
    void HideEditor();

    /// <summary>
    /// Checks that the macro is a public Sub or Function of a standard module, adds the module that calls it by
    /// name, runs that module's Run with Application.Run, and removes the module (docs/architecture.md
    /// "実行ごとに作るモジュール"). Blocks while the macro runs.
    /// </summary>
    MacroResult RunMacro(MacroName macro, IReadOnlyList<object> args);

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

    /// <summary>
    /// A dialog that blocks the process, with its text, or null (docs/architecture.md "失敗の返し方と記録").
    /// Reads windows only; it never touches COM.
    /// </summary>
    DialogText? FindDialog();

    /// <summary>Posts WM_CLOSE to the dialog <see cref="FindDialog"/> finds, if there is one. Reads windows only.</summary>
    void CloseDialog();

    /// <summary>Whether each XLMAIN window of the process is enabled, for the diagnostic log. Reads windows only.</summary>
    IReadOnlyList<bool> MainWindowsEnabled();
}

/// <summary>
/// One module for vba_get (docs/api.md): Type is standard, class, document, or form; Sheet is the sheet name of a
/// sheet module, and null for the others.
/// </summary>
public sealed record VbaComponent(string Name, string Type, int Lines, string? Sheet);

/// <summary>
/// One reference for vba_get (docs/api.md "vba_getの参照設定"). Name is null when a broken reference does not give it.
/// </summary>
public sealed record VbaReference(string? Name, string Guid, int Major, int Minor, bool BuiltIn, bool Broken);

/// <summary>What vba_edit asks for: add the GUID at the version, or remove it (Major and Minor are 0).</summary>
public sealed record ReferenceEdit(bool Add, string Guid, int Major, int Minor);

/// <summary>What <see cref="IExcel.EditReference"/> did, or why it did nothing.</summary>
public enum ReferenceChange
{
    Added,

    /// <summary>The project already had a reference with the GUID; nothing was added.</summary>
    AlreadyPresent,

    Removed,

    /// <summary>No reference has the GUID; nothing was removed.</summary>
    NotFound,

    /// <summary>The reference with the GUID is built in (VBA, Excel) and cannot be removed; nothing was removed.</summary>
    BuiltIn,
}

public enum CompileStart
{
    /// <summary>The compile control was already disabled: the project is compiled, so nothing was executed.</summary>
    AlreadyCompiled,

    /// <summary>The control was executed.</summary>
    Started,
}

/// <summary>The test that test_run is running: its name, its place from 1, and the number of tests.</summary>
public sealed record TestProgress(string Name, int Index, int Count);

/// <summary>Where the VBE put the selection for a compile error.</summary>
public sealed record CompileErrorPlace(string Module, int Line);

/// <summary>The title and the body text of a dialog. It goes only into the response, never into a log.</summary>
public sealed record DialogText(string Title, string Text);

/// <summary>What one macro_run call came to.</summary>
public abstract record MacroResult
{
    private MacroResult()
    {
    }

    /// <summary>No public Sub or Function of a standard module has that name.</summary>
    public sealed record NotFound(string Reason) : MacroResult;

    /// <summary>A run-time error the generated module caught.</summary>
    public sealed record Failed(int Number, string Description, string Source) : MacroResult;

    /// <summary>A return value; null for a Sub, Empty, or Null.</summary>
    public sealed record Returned(object? Value) : MacroResult;

    /// <summary>A return value that is not returned, by its VBA type name (an object, an array, a CVErr value, a Date).</summary>
    public sealed record OfType(string TypeName) : MacroResult;
}

/// <summary>
/// A vba_sync failure: the detail of docs/api.md "vba_sync", the modules done before it, and the source file it is
/// about when there is one.
/// </summary>
public sealed class VbaSyncException(string detail, IReadOnlyList<string> done, string message, string? file = null) : Exception(message)
{
    public string Detail { get; } = detail;

    public IReadOnlyList<string> Done { get; } = done;

    public string? File { get; } = file;
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
    TimeSpan CleanupQuitExit,
    TimeSpan TerminateExit,
    TimeSpan PendingReturn)
{
    public static SessionTimeouts Default { get; } = new(
        Start: TimeSpan.FromSeconds(120),
        Sync: TimeSpan.FromSeconds(30),
        Save: TimeSpan.FromSeconds(120),
        QuitExit: TimeSpan.FromSeconds(75),
        CleanupQuitExit: TimeSpan.FromSeconds(10),
        TerminateExit: TimeSpan.FromSeconds(5),
        PendingReturn: TimeSpan.FromSeconds(10));

    /// <summary>How long vba_check watches for a dialog after the compile (docs/api.md "vba_check").</summary>
    public TimeSpan CompileWatch { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>How long a dialog may stay after WM_CLOSE before the Excel is ended.</summary>
    public TimeSpan DialogClose { get; init; } = TimeSpan.FromSeconds(1);
}
