// Throwaway experiment for stage 4 (docs/architecture.md, "VBAの反映"). It starts one dedicated Excel on an STA thread
// and works only on workbooks in a new temporary folder.
//   1. Make a source workbook with Workbook_Open, Auto_Open, sheet and workbook calculate/change events, a volatile UDF
//      used by a cell, a class module, Japanese text, and a reference to the Scripting Runtime. Each event and the UDF
//      append a line to a marker file next to the workbook.
//   1b. As a control, open the source with AutomationSecurity=Low and events on, and provoke the same as below.
//   2. Copy it into an app folder and open the copy with AutomationSecurity=ForceDisable while EnableEvents is true,
//      so that only ForceDisable stands between the workbook and its code. Recalculate, change a cell, and try
//      Application.Run. Read the modules (CodeModule.Lines), export them, and list the references.
//   3. With the copy still open, set AutomationSecurity back to Low and repeat the recalculation, change and run.
//   4. Make a workbook with an Excel 4.0 macro sheet, save it as .xlsm, and list the parts of both packages.
// A watch prints any dialog of the owned Excel and closes it with WM_CLOSE. A step that does not return within 30
// seconds ends the owned Excel. It only ends the Excel whose ownership it confirmed: a new PID, named EXCEL.EXE,
// created after the start request.
//
// Usage: dotnet run --project spikes\extract

using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    private const uint WmClose = 0x0010;
    private const int FormatXlsm = 52;
    private const int SecurityLow = 1;
    private const int SecurityForceDisable = 3;
    private const int SheetTypeExcel4Macro = 3; // xlExcel4MacroSheet
    private const string ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}";

    private static readonly string Folder = Path.Combine(Path.GetTempPath(), "xlwright-spike-extract-" + Guid.NewGuid().ToString("N"));
    private static Process? s_owned;
    private static readonly HashSet<IntPtr> s_seenDialogs = [];

    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args.Length != 0)
        {
            Console.Error.WriteLine("Usage: Extract");
            return 64;
        }

        var exit = 1;
        var thread = new Thread(() => exit = Run());
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        Directory.Delete(Folder, recursive: true);
        Console.WriteLine($"temporary folder removed: {!Directory.Exists(Folder)}");
        return exit;
    }

    private static int Run()
    {
        Directory.CreateDirectory(Folder);
        var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Console.WriteLine($"before: {before.Count} EXCEL process(es)");
        var requested = DateTime.Now;
        dynamic? app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
        GetWindowThreadProcessId(new IntPtr((int)app!.Hwnd), out var rawPid);
        var process = Process.GetProcessById((int)rawPid);
        var owned = !before.Contains(process.Id) &&
            process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) &&
            process.StartTime > requested.AddSeconds(-1);
        Console.WriteLine($"started PID {process.Id}, owned {owned}, visible {(bool)app.Visible}");
        if (!owned)
        {
            Console.WriteLine("ownership not confirmed; leaving it untouched");
            return 1;
        }

        s_owned = process;
        app.UserControl = true;
        using var watch = new Timer(_ => WatchDialogs(), null, TimeSpan.Zero, TimeSpan.FromMilliseconds(200));
        var ended = false;
        try
        {
            Trials(app);
        }
        catch (Exception e) when (process.HasExited)
        {
            Console.WriteLine($"Excel ended during the trials: {Describe(e)}");
            ended = true;
        }
        catch (Exception e)
        {
            Console.WriteLine($"trial failed: {Describe(e)}");
        }
        finally
        {
            if (!ended)
            {
                app.DisplayAlerts = false;
                dynamic workbooks = app.Workbooks;
                while ((int)workbooks.Count > 0)
                {
                    dynamic open = workbooks.Item(1);
                    open.Close(false);
                    Release(open);
                }

                Release(workbooks);
                app.Quit();
                Release(app);
            }

            app = null;
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        ended = process.WaitForExit(TimeSpan.FromSeconds(20));
        if (!ended)
        {
            process.Kill();
            ended = process.WaitForExit(TimeSpan.FromSeconds(5));
            Console.WriteLine("terminated after Quit");
        }

        var after = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Console.WriteLine($"ended {ended}; after: {after.Count} EXCEL process(es), unchanged from before {after.SetEquals(before)}");
        return 0;
    }

    private static void Trials(dynamic app)
    {
        var sourceFolder = Path.Combine(Folder, "source");
        var appFolder = Path.Combine(Folder, "app");
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(appFolder);
        var source = Path.Combine(sourceFolder, "Existing.xlsm");
        var copy = Path.Combine(appFolder, "Existing.xlsm");
        dynamic workbooks = app.Workbooks;

        Console.WriteLine("== 1. make the source workbook (Low, events off)");
        app.AutomationSecurity = SecurityLow;
        app.EnableEvents = false;
        dynamic book = workbooks.Add();
        dynamic project = book.VBProject;
        dynamic components = project.VBComponents;
        SetCode(components, (string)book.CodeName, BookCode);
        dynamic sheet = book.Worksheets.Item(1);
        SetCode(components, (string)sheet.CodeName, SheetCode);
        Release(AddModule(components, 1, "Udf", UdfModule));
        Release(AddModule(components, 2, "Counter", ClassModule));
        dynamic reference = project.References.AddFromGuid(ScriptingGuid, 1, 0);
        Console.WriteLine($"  added reference {(string)reference.Name}");
        Release(reference);
        sheet.Range("A1").Formula = "=Mark()";
        sheet.Range("A2").Formula = "=NOW()";
        app.DisplayAlerts = false;
        book.SaveAs(source, FormatXlsm);
        app.DisplayAlerts = true;
        Console.WriteLine($"  A1 cached value {Describe(() => Convert.ToString(sheet.Range("A1").Value))}; source markers: {Markers(sourceFolder)}");
        Release(sheet);
        Release(components);
        Release(project);
        book.Close(false);
        Release(book);
        app.EnableEvents = true;

        Console.WriteLine("== 1b. control: open the source with Low, events on");
        book = Guard("open", () => workbooks.Open(source, 0, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false));
        Console.WriteLine($"  opened; markers: {Markers(sourceFolder)}");
        Provoke(app, book, "B1", "Low");
        app.EnableEvents = false;
        book.Close(false);
        app.EnableEvents = true;
        Release(book);

        Console.WriteLine("== 2. copy into the app folder and open with ForceDisable, events on");
        File.Copy(source, copy);
        app.AutomationSecurity = SecurityForceDisable;
        Console.WriteLine($"  AutomationSecurity {(int)app.AutomationSecurity}, EnableEvents {(bool)app.EnableEvents}");
        book = Guard("open", () => workbooks.Open(copy, 0, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false));
        Console.WriteLine($"  opened; ReadOnly {(bool)book.ReadOnly}; markers: {Markers(appFolder)}");
        Provoke(app, book, "B1", "ForceDisable");

        Console.WriteLine("== 2b. read the modules and the references");
        project = book.VBProject;
        components = project.VBComponents;
        var exportFolder = Path.Combine(Folder, "export");
        Directory.CreateDirectory(exportFolder);
        foreach (dynamic component in components)
        {
            string name = component.Name;
            int type = component.Type;
            dynamic code = component.CodeModule;
            int count = code.CountOfLines;
            var text = count == 0 ? "" : (string)code.Lines(1, count);
            Release(code);
            var exported = Path.Combine(exportFolder, name + (type == 1 ? ".bas" : ".cls"));
            var exportResult = Describe(() =>
            {
                component.Export(exported);
                var bytes = File.ReadAllBytes(exported);
                var decoded = Encoding.GetEncoding(932).GetString(bytes);
                return $"exported {bytes.Length} bytes, CP932 has 日本語 {decoded.Contains("日本語")}";
            });
            Console.WriteLine($"  {name}(type {type}): {count} line(s), Lines has 日本語 {text.Contains("日本語")}; {exportResult}");
            Release(component);
        }

        dynamic references = project.References;
        foreach (dynamic item in references)
        {
            Console.WriteLine($"  reference {Describe(() => (string)item.Name)} {Describe(() => (string)item.Guid)} {Describe(() => $"{(int)item.Major}.{(int)item.Minor}")} builtIn {Describe(() => Convert.ToString((bool)item.BuiltIn))} broken {Describe(() => Convert.ToString((bool)item.IsBroken))}");
            Release(item);
        }

        Release(references);
        Release(components);
        Release(project);

        Console.WriteLine("== 3. set AutomationSecurity back to Low with the copy open");
        app.AutomationSecurity = SecurityLow;
        Provoke(app, book, "B2", "Low after open");

        app.EnableEvents = false;
        book.Close(false);
        app.EnableEvents = true;
        Release(book);

        Console.WriteLine("== 4. Excel 4.0 macro sheet and the package parts");
        book = workbooks.Add();
        dynamic sheets = book.Sheets;
        var macroSheet = Describe(() => { sheets.Add(Type.Missing, Type.Missing, 1, SheetTypeExcel4Macro); return "added"; });
        Console.WriteLine($"  macro sheet: {macroSheet}");
        Release(sheets);
        var withMacroSheet = Path.Combine(sourceFolder, "MacroSheet.xlsm");
        app.DisplayAlerts = false;
        book.SaveAs(withMacroSheet, FormatXlsm);
        app.DisplayAlerts = true;
        book.Close(false);
        Release(book);
        Release(workbooks);
        Console.WriteLine($"  Existing.xlsm parts: {Parts(source)}");
        Console.WriteLine($"  MacroSheet.xlsm parts: {Parts(withMacroSheet)}");
    }

    // Recalculates, changes a cell, and tries Application.Run, printing the markers after each.
    private static void Provoke(dynamic app, dynamic book, string cell, string label)
    {
        var folder = (string)book.Path;
        dynamic sheet = book.Worksheets.Item(1);
        Guard("calculate", () => { app.CalculateFull(); return 0; });
        Console.WriteLine($"  [{label}] after CalculateFull: A1 {Describe(() => Convert.ToString(sheet.Range("A1").Value))} text {Describe(() => (string)sheet.Range("A1").Text)}; markers: {Markers(folder)}");
        Guard("change", () => { sheet.Range(cell).Value = "changed"; return 0; });
        Console.WriteLine($"  [{label}] after writing {cell}: markers: {Markers(folder)}");
        var run = Describe(() => Guard("run", () => Convert.ToString(app.Run("'" + (string)book.Name + "'!Udf.Mark"))));
        Console.WriteLine($"  [{label}] Application.Run Udf.Mark: {run}; markers: {Markers(folder)}");
        Release(sheet);
    }

    private static string Markers(string folder)
    {
        var files = Directory.EnumerateFiles(folder, "marker-*.txt")
            .Select(f => $"{Path.GetFileNameWithoutExtension(f)[7..]}x{File.ReadAllLines(f).Length}")
            .Order()
            .ToList();
        return files.Count == 0 ? "none" : string.Join(", ", files);
    }

    private static string Parts(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        var folders = zip.Entries.Select(e => e.FullName).Where(n => n.StartsWith("xl/", StringComparison.Ordinal))
            .Select(n => n.Count(c => c == '/') > 1 ? n[..(n.IndexOf('/', 3) + 1)] : n).Distinct().Order();
        return string.Join(", ", folders);
    }

    // ---- helpers ----

    // Replaces the code of an existing module found by name; for a sheet or workbook module, the name is the CodeName.
    private static void SetCode(dynamic components, string name, string text)
    {
        dynamic component = components.Item(name);
        dynamic code = component.CodeModule;
        int count = code.CountOfLines;
        if (count > 0) code.DeleteLines(1, count);
        code.AddFromString(text.ReplaceLineEndings("\r\n"));
        Release(code);
        Release(component);
    }

    private static dynamic AddModule(dynamic components, int type, string name, string text)
    {
        dynamic added = components.Add(type); // 1 vbext_ct_StdModule, 2 vbext_ct_ClassModule
        added.Name = name;
        dynamic code = added.CodeModule;
        code.AddFromString(text.ReplaceLineEndings("\r\n"));
        Release(code);
        return added;
    }

    private static T Guard<T>(string step, Func<T> action)
    {
        using var timer = new Timer(_ =>
        {
            Console.WriteLine($"  {step}: no return within 30 s; ending the owned Excel");
            s_owned!.Kill();
        }, null, TimeSpan.FromSeconds(30), Timeout.InfiniteTimeSpan);
        return action();
    }

    private static void WatchDialogs()
    {
        if (s_owned is null || s_owned.HasExited) return;
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid != s_owned.Id || !IsWindowVisible(hwnd)) return true;
            var name = new StringBuilder(64);
            GetClassName(hwnd, name, name.Capacity);
            if (name.ToString() != "#32770") return true;
            lock (s_seenDialogs)
            {
                if (!s_seenDialogs.Add(hwnd)) return true;
            }

            var texts = new List<string>();
            EnumChildWindows(hwnd, (child, _) =>
            {
                var childClass = new StringBuilder(64);
                GetClassName(child, childClass, childClass.Capacity);
                if (childClass.ToString() == "Static" && IsWindowVisible(child))
                {
                    var text = new StringBuilder(1024);
                    GetWindowText(child, text, text.Capacity);
                    if (text.Length > 0) texts.Add(text.ToString());
                }

                return true;
            }, IntPtr.Zero);
            var title = new StringBuilder(256);
            GetWindowText(hwnd, title, title.Capacity);
            Console.WriteLine($"  [dialog] \"{title}\": {string.Join(" / ", texts).ReplaceLineEndings(" ")}; close");
            PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
    }

    private static string Describe(Func<string?> read)
    {
        try
        {
            return read() ?? "null";
        }
        catch (Exception e)
        {
            return Describe(e);
        }
    }

    private static string Describe(Exception e)
    {
        var inner = e is System.Reflection.TargetInvocationException { InnerException: { } i } ? i : e;
        return $"{inner.GetType().Name} 0x{inner.HResult:X8} {inner.Message.ReplaceLineEndings(" ")}";
    }

    private static void Release(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject)) Marshal.FinalReleaseComObject(comObject);
    }

    private const string BookCode = """
        Option Explicit

        Private Sub Workbook_Open()
            Udf.Note "open"
        End Sub

        Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
            Udf.Note "bookcalc"
        End Sub
        """;

    private const string SheetCode = """
        Option Explicit

        Private Sub Worksheet_Change(ByVal Target As Range)
            Udf.Note "change"
        End Sub

        Private Sub Worksheet_Calculate()
            Udf.Note "sheetcalc"
        End Sub
        """;

    private const string UdfModule = """
        Option Explicit
        ' 日本語のコメント

        Public Sub Auto_Open()
            Note "autoopen"
        End Sub

        Public Function Mark() As Long
            Application.Volatile
            Note "udf"
            Mark = 7
        End Function

        Public Sub Note(ByVal what As String)
            Dim f As Integer
            f = FreeFile
            Open ThisWorkbook.Path & "\marker-" & what & ".txt" For Append As #f
            Print #f, what
            Close #f
        End Sub
        """;

    private const string ClassModule = """
        Option Explicit
        ' 日本語のクラス

        Public Value As Long
        """;
}
