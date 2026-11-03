using System.Runtime.InteropServices;

namespace XlWright.Core.Com;

/// <summary>
/// Starts Excel.Application through COM on the session's STA thread. Never connects to a running Excel
/// (no GetActiveObject; docs/architecture.md "日常のExcelとの分離").
/// </summary>
public sealed class ComExcelLauncher : IExcelLauncher
{
    public LaunchedExcel Launch()
    {
        MessageFilter.Register();
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
        dynamic app = Activator.CreateInstance(type)!;
        try
        {
            int hwnd = app.Hwnd;
            _ = NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);
            return new LaunchedExcel(new ComExcel(app), WindowsProcess.Open((int)pid));
        }
        catch (COMException)
        {
            // An Excel is running but its process cannot be found, so ownership cannot be confirmed.
            return new LaunchedExcel(new ComExcel(app), WindowsProcess.Unidentified());
        }
    }
}

/// <summary>
/// The Excel operations of one session, through late binding (docs/architecture.md "本体の部品の境界").
/// Every COM object it obtains is released explicitly: an unreleased intermediate object kept Excel running after Quit.
/// </summary>
internal sealed class ComExcel(dynamic app) : IExcel
{
    private const int AutomationSecurityLow = 1; // msoAutomationSecurityLow
    private const int CalculationDone = 0; // xlDone

    private dynamic? _app = app;
    private dynamic? _main;

    private dynamic App => _app ?? throw new InvalidOperationException("The Excel references were released.");

    private dynamic Main => _main ?? throw new InvalidOperationException("The main workbook is not open.");

    public void TakeControl() => App.UserControl = true;

    public bool Open(string path)
    {
        App.AutomationSecurity = AutomationSecurityLow;
        App.EnableEvents = false;
        dynamic workbooks = App.Workbooks;
        try
        {
            _main = workbooks.Open(path);
        }
        finally
        {
            Release(workbooks);
            App.EnableEvents = true;
        }

        return (bool)_main.ReadOnly;
    }

    public void Save() => Main.Save();

    public object?[,]? ReadValues(string sheet, CellRange range)
    {
        // Calculate and the check run in the same operation as the read (docs/architecture.md "起動時の設定").
        App.Calculate();
        while ((int)App.CalculationState != CalculationDone)
        {
            Thread.Sleep(50);
        }

        return OnRange(sheet, range, target => Cells(range, (object?)target.Value2, value => value switch
        {
            // Value2 returns numbers as double, so an Int32 is a cell error (docs/notes/range-values.md).
            int hresult => CellError.FromHResult(hresult),
            null or double or string or bool => value,
            _ => throw new InvalidOperationException($"Excel returned a cell of type {value.GetType().Name}."),
        }));
    }

    public string[,]? ReadFormulas(string sheet, CellRange range) =>
        OnRange(sheet, range, target => Cells(range, (object?)target.Formula, value => (string)value!));

    public bool Write(string sheet, CellRange range, RangeWrite write)
    {
        bool events = App.EnableEvents;
        App.EnableEvents = false;
        try
        {
            return OnRange(sheet, range, target =>
            {
                if (write.Values is { } values)
                {
                    target.Value2 = values;
                }
                else if (write.Formulas is { } formulas)
                {
                    var cells = new object?[range.Rows, range.Columns];
                    Array.Copy(formulas, cells, formulas.Length);
                    target.Formula = cells;
                }
                else
                {
                    target.ClearContents();
                }

                return true;
            });
        }
        finally
        {
            App.EnableEvents = events;
        }
    }

    public IReadOnlyList<string?> OpenWorkbookPaths()
    {
        var paths = new List<string?>();
        dynamic workbooks = App.Workbooks;
        try
        {
            for (var i = 1; i <= (int)workbooks.Count; i++)
            {
                dynamic workbook = workbooks.Item(i);
                try
                {
                    paths.Add(string.IsNullOrEmpty((string)workbook.Path) ? null : (string)workbook.FullName);
                }
                finally
                {
                    Release(workbook);
                }
            }
        }
        finally
        {
            Release(workbooks);
        }

        // A workbook in Protected View is not in Workbooks (docs/architecture.md "残った専用のExcelの片付け").
        dynamic windows = App.ProtectedViewWindows;
        try
        {
            for (var i = 1; i <= (int)windows.Count; i++)
            {
                dynamic window = windows.Item(i);
                try
                {
                    paths.Add(Path.Combine((string)window.SourcePath, (string)window.SourceName));
                }
                finally
                {
                    Release(window);
                }
            }
        }
        finally
        {
            Release(windows);
        }

        return paths;
    }

    // Events stay off while xlwright closes workbooks, so Workbook_BeforeClose cannot stop or change the close
    // (docs/architecture.md "起動時の設定"). Quit and the release run even when a close fails.
    public void CloseAllAndQuit()
    {
        ReleaseMain();
        try
        {
            App.EnableEvents = false;
            dynamic workbooks = App.Workbooks;
            try
            {
                for (var i = (int)workbooks.Count; i >= 1; i--)
                {
                    dynamic workbook = workbooks.Item(i);
                    try
                    {
                        workbook.Close(false);
                    }
                    finally
                    {
                        Release(workbook);
                    }
                }
            }
            finally
            {
                Release(workbooks);
            }
        }
        finally
        {
            try
            {
                App.Quit();
            }
            finally
            {
                ReleaseReferences();
            }
        }
    }

    public void ReleaseReferences()
    {
        ReleaseMain();
        if (_app is not null)
        {
            Release(_app);
            _app = null;
        }
    }

    private void ReleaseMain()
    {
        if (_main is not null)
        {
            Release(_main);
            _main = null;
        }
    }

    // Runs `use` on the range of the main workbook's sheet; returns default when no sheet has that name.
    // Sheet names are compared without case, as Excel does.
    private T? OnRange<T>(string sheet, CellRange range, Func<dynamic, T> use)
    {
        dynamic worksheets = Main.Worksheets;
        try
        {
            for (var i = 1; i <= (int)worksheets.Count; i++)
            {
                dynamic worksheet = worksheets.Item(i);
                try
                {
                    if (!string.Equals((string)worksheet.Name, sheet, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    dynamic target = worksheet.Range(range.Address);
                    try
                    {
                        return use(target);
                    }
                    finally
                    {
                        Release(target);
                    }
                }
                finally
                {
                    Release(worksheet);
                }
            }

            return default;
        }
        finally
        {
            Release(worksheets);
        }
    }

    // Value2 and Formula return one cell as a value, and more cells as a 1-based object[,].
    private static T[,] Cells<T>(CellRange range, object? returned, Func<object?, T> convert)
    {
        var cells = new T[range.Rows, range.Columns];
        if (returned is object?[,] array)
        {
            for (var row = 0; row < range.Rows; row++)
            {
                for (var column = 0; column < range.Columns; column++)
                {
                    cells[row, column] = convert(array[row + 1, column + 1]);
                }
            }
        }
        else
        {
            cells[0, 0] = convert(returned);
        }

        return cells;
    }

    private static void Release(object comObject) => Marshal.FinalReleaseComObject(comObject);
}
