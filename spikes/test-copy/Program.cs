// Throwaway experiment for stage 3 (docs/architecture.md, "テスト" and "VBAの反映"). It starts one dedicated Excel on
// an STA thread and works only on workbooks in a new temporary folder.
//   flow:     the test copy flow. Save and close the main workbook, copy the app folder to .xlwright/test/<run-id>/,
//             open the copy with events off, import a test module, run it with Application.Run qualified by the main
//             workbook's name, close the copy without saving, and reopen the main workbook.
//   compile:  compile the VBA project from outside with the VBE's "Compile VBAProject" control (Id 578), with good
//             code and with compile errors, while Excel is hidden.
//   document: find the sheet and workbook modules by CodeName and replace their code.
// A watch prints any dialog of the owned Excel and closes it with WM_CLOSE. A step that does not return within 30
// seconds ends the owned Excel. It only ends the Excel whose ownership it confirmed: a new PID, named EXCEL.EXE,
// created after the start request.
//
// The second argument chooses what the watch does with a dialog: close (WM_CLOSE, the default), ok (WM_COMMAND with
// IDOK), or keep (leave it open).
//
// Usage: dotnet run --project spikes\test-copy -- <flow|compile|document> [close|ok|keep]

using System.Diagnostics;
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

    private static readonly string Folder = Path.Combine(Path.GetTempPath(), "xlwright-spike-test-copy-" + Guid.NewGuid().ToString("N"));
    private static Process? s_owned;
    private static readonly HashSet<IntPtr> s_seenDialogs = [];
    private static string s_dialogAction = "close";
    private const uint WmCommand = 0x0111;

    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var exit = 1;
        var mode = args.Length >= 1 ? args[0] : "";
        s_dialogAction = args.Length == 2 ? args[1] : "close";
        if (mode is not ("flow" or "compile" or "document") || args.Length > 2 || s_dialogAction is not ("close" or "ok" or "keep"))
        {
            Console.Error.WriteLine("Usage: TestCopy <flow|compile|document> [close|ok|keep]");
            return 64;
        }

        var thread = new Thread(() => exit = Run(mode));
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        Directory.Delete(Folder, recursive: true);
        Console.WriteLine($"temporary folder removed: {!Directory.Exists(Folder)}");
        return exit;
    }

    private static int Run(string mode)
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
            switch (mode)
            {
                case "flow": Flow(app); break;
                case "compile": Compile(app); break;
                default: Document(app); break;
            }
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

    // ---- flow ----

    private static void Flow(dynamic app)
    {
        var appFolder = Path.Combine(Folder, "app");
        Directory.CreateDirectory(Path.Combine(appFolder, "data"));
        Directory.CreateDirectory(Path.Combine(appFolder, "src", "modules"));
        Directory.CreateDirectory(Path.Combine(appFolder, "tests"));
        Directory.CreateDirectory(Path.Combine(appFolder, ".git"));
        File.WriteAllText(Path.Combine(appFolder, "data", "note.txt"), "original");
        File.WriteAllText(Path.Combine(appFolder, "src", "modules", "Calc.bas"), "source");
        File.WriteAllText(Path.Combine(appFolder, "tests", "Tests.bas"), TestsModule.ReplaceLineEndings("\r\n"), Encoding.GetEncoding(932));
        var main = Path.Combine(appFolder, "Main.xlsm");

        Console.WriteLine("== make the main workbook with Workbook_Open and Auto_Open");
        dynamic workbooks = app.Workbooks;
        dynamic book = workbooks.Add();
        dynamic components = book.VBProject.VBComponents;
        SetCode(components, (string)book.CodeName, OpenEventCode);
        AddStandard(components, "Starter", AutoOpenModule);
        ((dynamic)book.Worksheets.Item(1)).Range("A1").Value = "original";
        app.DisplayAlerts = false;
        book.SaveAs(main, FormatXlsm);
        app.DisplayAlerts = true;
        book.Close(false);
        Release(components);
        Release(book);
        Console.WriteLine($"  saved {Path.GetFileName(main)}; markers: {Markers(appFolder)}");

        Console.WriteLine("== open the main workbook with events off");
        book = Open(app, workbooks, main, events: false);
        Console.WriteLine($"  markers after open with events off: {Markers(appFolder)}");

        Console.WriteLine("== 1. save and close the main workbook");
        app.EnableEvents = false;
        book.Save();
        book.Close(false);
        app.EnableEvents = true;
        Release(book);

        Console.WriteLine("== 2. copy the app folder");
        var runId = Guid.NewGuid().ToString("N")[..8];
        var copyFolder = Path.Combine(appFolder, ".xlwright", "test", runId);
        CopyApp(appFolder, copyFolder);
        Console.WriteLine($"  copied: {string.Join(", ", Directory.EnumerateFileSystemEntries(copyFolder, "*", SearchOption.AllDirectories).Select(p => Path.GetRelativePath(copyFolder, p)))}");

        Console.WriteLine("== 3. open the copy with events off and import the test module");
        var copyMain = Path.Combine(copyFolder, "Main.xlsm");
        dynamic copy = Open(app, workbooks, copyMain, events: false);
        Console.WriteLine($"  markers in the copy after open: {Markers(copyFolder)}");
        Console.WriteLine($"  copy name {(string)copy.Name}, full name {Shorten((string)copy.FullName)}");
        dynamic copyComponents = copy.VBProject.VBComponents;
        app.EnableEvents = false;
        dynamic imported = copyComponents.Import(Path.Combine(copyFolder, "..", "..", "..", "tests", "Tests.bas"));
        app.EnableEvents = true;
        Console.WriteLine($"  imported {(string)imported.Name}");
        Release(imported);

        Console.WriteLine("== the main workbook while the copy is open");
        try
        {
            app.DisplayAlerts = false;
            dynamic second = Guard("open the main", () => workbooks.Open(main));
            Console.WriteLine($"  opened: {Shorten((string)second.FullName)}; open books {(int)workbooks.Count}");
            Release(second);
        }
        catch (Exception e)
        {
            Console.WriteLine($"  open the main: {Describe(e)}; open books {(int)workbooks.Count}");
        }
        finally
        {
            app.DisplayAlerts = true;
        }

        Console.WriteLine("== 4. run the tests, qualified by the main workbook's name, with events on");
        string qualifier = "'Main.xlsm'!";
        object where = Guard("TestWhere", () => app.Run(qualifier + "Tests.TestWhere"));
        Console.WriteLine($"  Tests.TestWhere: {Shorten((string)where)}");
        object byName = Guard("TestByName", () => app.Run(qualifier + "Tests.TestByName"));
        Console.WriteLine($"  Tests.TestByName (Workbooks(\"Main.xlsm\")): {Shorten((string)byName)}");
        Guard("TestWrites", () => app.Run(qualifier + "Tests.TestWrites"));
        Console.WriteLine($"  after Tests.TestWrites: copy A1 \"{(string)((dynamic)copy.Worksheets.Item(1)).Range("A1").Value}\", copy note \"{File.ReadAllText(Path.Combine(copyFolder, "data", "note.txt")).Trim()}\", main note \"{File.ReadAllText(Path.Combine(appFolder, "data", "note.txt")).Trim()}\"");
        Release(copyComponents);

        Console.WriteLine("== 5. close the copy without saving, remove it, and reopen the main workbook");
        app.EnableEvents = false;
        app.DisplayAlerts = false;
        copy.Close(false);
        app.DisplayAlerts = true;
        app.EnableEvents = true;
        Release(copy);
        Directory.Delete(copyFolder, recursive: true);
        Console.WriteLine($"  copy folder removed: {!Directory.Exists(copyFolder)}");
        book = Open(app, workbooks, main, events: false);
        dynamic mainComponents = book.VBProject.VBComponents;
        var names = new List<string>();
        foreach (dynamic component in mainComponents)
        {
            names.Add((string)component.Name);
            Release(component);
        }

        Console.WriteLine($"  main A1 \"{(string)((dynamic)book.Worksheets.Item(1)).Range("A1").Value}\", modules [{string.Join(", ", names)}], markers {Markers(appFolder)}");
        Release(mainComponents);

        Console.WriteLine("== control: reopen the main workbook with events on");
        app.EnableEvents = false;
        book.Close(false);
        app.EnableEvents = true;
        Release(book);
        book = Open(app, workbooks, main, events: true);
        Console.WriteLine($"  markers after open with events on: {Markers(appFolder)}");
        Release(book);
        Release(workbooks);
    }

    private static void CopyApp(string appFolder, string target)
    {
        Directory.CreateDirectory(target);
        foreach (var entry in Directory.EnumerateFileSystemEntries(appFolder))
        {
            var name = Path.GetFileName(entry);
            if (name.StartsWith('.') || name is "src" or "tests") continue;
            if (Directory.Exists(entry))
            {
                CopyApp(entry, Path.Combine(target, name));
            }
            else
            {
                File.Copy(entry, Path.Combine(target, name));
            }
        }
    }

    private static string Markers(string folder) =>
        $"Workbook_Open {File.Exists(Path.Combine(folder, "opened-event.txt"))}, Auto_Open {File.Exists(Path.Combine(folder, "opened-auto.txt"))}";

    private static dynamic Open(dynamic app, dynamic workbooks, string path, bool events)
    {
        dynamic book;
        app.AutomationSecurity = SecurityLow;
        app.EnableEvents = events;
        try
        {
            book = Guard("open", () => workbooks.Open(path));
        }
        finally
        {
            app.EnableEvents = true;
        }

        return book;
    }

    // ---- compile ----

    private static void Compile(dynamic app)
    {
        dynamic workbooks = app.Workbooks;
        dynamic book = workbooks.Add();
        dynamic project = book.VBProject;
        dynamic components = project.VBComponents;
        dynamic vbe = app.VBE;
        dynamic control = vbe.CommandBars.FindControl(1, 578); // msoControlButton, "Compile VBAProject"
        Console.WriteLine($"control: \"{(string)control.Caption}\"; VBE window visible {(bool)vbe.MainWindow.Visible}");
        Console.WriteLine($"active project: {Describe(() => (string)vbe.ActiveVBProject.Name)}; this project {(string)project.Name}");
        vbe.ActiveVBProject = project;

        AddStandard(components, "Good", GoodModule);
        Console.WriteLine($"== good code: enabled before {(bool)control.Enabled}");
        Execute(control);
        Console.WriteLine($"  enabled after {(bool)control.Enabled}; VBE window visible {(bool)vbe.MainWindow.Visible}; Excel visible {(bool)app.Visible}");

        dynamic bad = AddStandard(components, "Undeclared", UndeclaredModule);
        Console.WriteLine($"== undeclared variable: enabled before {(bool)control.Enabled}");
        Execute(control);
        Console.WriteLine($"  enabled after {(bool)control.Enabled}; VBE window visible {(bool)vbe.MainWindow.Visible}; Excel visible {(bool)app.Visible}");
        Console.WriteLine($"  selection: {Describe(() => Selection(vbe))}");
        vbe.MainWindow.Visible = false;
        Console.WriteLine($"  hid the VBE window: visible {(bool)vbe.MainWindow.Visible}");
        components.Remove(bad);
        Release(bad);

        bad = AddStandard(components, "Syntax", SyntaxModule);
        Console.WriteLine($"== syntax error: enabled before {(bool)control.Enabled}");
        Execute(control);
        Console.WriteLine($"  enabled after {(bool)control.Enabled}; VBE window visible {(bool)vbe.MainWindow.Visible}");
        Console.WriteLine($"  selection: {Describe(() => Selection(vbe))}");
        components.Remove(bad);
        Release(bad);

        Console.WriteLine($"== after removing the errors: enabled before {(bool)control.Enabled}");
        Execute(control);
        Console.WriteLine($"  enabled after {(bool)control.Enabled}");

        Console.WriteLine("== a second workbook is the active project, compile the first");
        dynamic other = workbooks.Add();
        dynamic otherComponents = other.VBProject.VBComponents;
        dynamic otherBad = AddStandard(otherComponents, "OtherBad", UndeclaredModule);
        Console.WriteLine($"  active project: {Describe(() => (string)vbe.ActiveVBProject.Name)} of {Describe(() => (string)vbe.ActiveVBProject.FileName)}");
        dynamic touched = AddStandard(components, "Touched", GoodModule);
        vbe.ActiveVBProject = project;
        Console.WriteLine($"  set active to the first; enabled before {(bool)control.Enabled}");
        Execute(control);
        Console.WriteLine($"  enabled after {(bool)control.Enabled}");
        Release(touched);
        Release(otherBad);
        Release(otherComponents);
        Release(other);
        Release(control);
        Release(vbe);
        Release(components);
        Release(project);
        Release(book);
        Release(workbooks);
    }

    private static void Execute(dynamic control)
    {
        var elapsed = Stopwatch.StartNew();
        try
        {
            int seen;
            lock (s_seenDialogs) seen = s_seenDialogs.Count;
            Guard("compile", () => control.Execute());
            Console.WriteLine($"  Execute returned in {elapsed.ElapsedMilliseconds} ms");
            Thread.Sleep(2000); // let a dialog show and the watch handle it before the next call
            lock (s_seenDialogs) Console.WriteLine($"  dialogs within 2 s: {s_seenDialogs.Count - seen}");
        }
        catch (Exception e) when (!s_owned!.HasExited)
        {
            Console.WriteLine($"  Execute: {Describe(e)} in {elapsed.ElapsedMilliseconds} ms");
        }
    }

    private static string Selection(dynamic vbe)
    {
        dynamic pane = vbe.ActiveCodePane;
        if (pane is null) return "no active code pane";
        int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
        pane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
        string module = pane.CodeModule.Parent.Name;
        return $"{module} line {startLine} column {startColumn} to line {endLine} column {endColumn}";
    }

    // ---- document ----

    private static void Document(dynamic app)
    {
        dynamic workbooks = app.Workbooks;
        dynamic book = workbooks.Add();
        dynamic first = book.Worksheets.Item(1);
        dynamic added = book.Worksheets.Add();
        Console.WriteLine($"before VBProject: workbook CodeName \"{(string)book.CodeName}\", first sheet \"{(string)first.Name}\" CodeName \"{(string)first.CodeName}\", added sheet \"{(string)added.Name}\" CodeName \"{(string)added.CodeName}\"");
        first.Name = "データ";
        dynamic project = book.VBProject;
        dynamic components = project.VBComponents;
        Console.WriteLine($"after VBProject: first sheet \"{(string)first.Name}\" CodeName \"{(string)first.CodeName}\", added sheet CodeName \"{(string)added.CodeName}\"");
        Console.WriteLine($"components: {ListComponents(components)}");

        var sheetCodeName = (string)first.CodeName;
        var bookCodeName = (string)book.CodeName;
        SetCode(components, sheetCodeName, SheetCode);
        SetCode(components, bookCodeName, BookCode);
        SetCode(components, sheetCodeName, SheetCode); // a second write replaces, not appends
        Console.WriteLine($"read back sheet: {ReadCode(components, sheetCodeName) == SheetCode.ReplaceLineEndings("\r\n").TrimEnd()}, book: {ReadCode(components, bookCodeName) == BookCode.ReplaceLineEndings("\r\n").TrimEnd()}");
        Console.WriteLine($"components after: {ListComponents(components)}");
        try
        {
            components.Item("NoSuchModule");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Item of a missing name: {Describe(e)}");
        }

        var path = Path.Combine(Folder, "Doc.xlsm");
        app.DisplayAlerts = false;
        book.SaveAs(path, FormatXlsm);
        app.DisplayAlerts = true;
        book.Close(false);
        Release(components);
        Release(project);
        Release(added);
        Release(first);
        Release(book);

        Console.WriteLine("== reopen, then write a cell with events on");
        book = Open(app, workbooks, path, events: false);
        dynamic sheet = book.Worksheets.Item("データ");
        Console.WriteLine($"reopened: sheet \"データ\" CodeName \"{(string)sheet.CodeName}\", code kept {ReadCode(book.VBProject.VBComponents, sheetCodeName) == SheetCode.ReplaceLineEndings("\r\n").TrimEnd()}");
        sheet.Range("A1").Value = 5;
        Console.WriteLine($"Worksheet_Change wrote B1: \"{(string?)sheet.Range("B1").Value}\"");
        object hello = Guard("Hello", () => app.Run($"'{(string)book.Name}'!{bookCodeName}.Hello"));
        Console.WriteLine($"{bookCodeName}.Hello: \"{hello}\"");
        Release(sheet);
        Release(book);
        Release(workbooks);
    }

    private static string ListComponents(dynamic components)
    {
        var list = new List<string>();
        foreach (dynamic component in components)
        {
            list.Add($"{(string)component.Name}({(int)component.Type})");
            Release(component);
        }

        return string.Join(", ", list);
    }

    private static string ReadCode(dynamic components, string name)
    {
        dynamic code = components.Item(name).CodeModule;
        int count = code.CountOfLines;
        return count == 0 ? "" : ((string)code.Lines(1, count)).TrimEnd();
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
        Console.WriteLine($"  set code of {name} (type {(int)component.Type}): {count} line(s) replaced");
        Release(code);
        Release(component);
    }

    private static dynamic AddStandard(dynamic components, string name, string text)
    {
        dynamic added = components.Add(1); // vbext_ct_StdModule
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
            Console.WriteLine($"  [dialog] \"{title}\": {string.Join(" / ", texts).ReplaceLineEndings(" ")}; {s_dialogAction}");
            if (s_dialogAction == "close") PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
            if (s_dialogAction == "ok") PostMessage(hwnd, WmCommand, new IntPtr(1), IntPtr.Zero); // IDOK
            return true;
        }, IntPtr.Zero);
    }

    private static string Shorten(string path) => path.Replace(Folder, "<tmp>", StringComparison.OrdinalIgnoreCase);

    private static string Describe(Func<string> read)
    {
        try
        {
            return read();
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

    private const string OpenEventCode = """
        Option Explicit

        Private Sub Workbook_Open()
            Dim f As Integer
            f = FreeFile
            Open ThisWorkbook.Path & "\opened-event.txt" For Output As #f
            Close #f
        End Sub
        """;

    private const string AutoOpenModule = """
        Option Explicit

        Public Sub Auto_Open()
            Dim f As Integer
            f = FreeFile
            Open ThisWorkbook.Path & "\opened-auto.txt" For Output As #f
            Close #f
        End Sub
        """;

    private const string TestsModule = """
        Attribute VB_Name = "Tests"
        Option Explicit

        Public Function TestWhere() As String
            TestWhere = ThisWorkbook.FullName
        End Function

        Public Function TestByName() As String
            TestByName = Workbooks("Main.xlsm").FullName
        End Function

        Public Sub TestWrites()
            ThisWorkbook.Worksheets(1).Range("A1").Value = "changed by test"
            Dim f As Integer
            f = FreeFile
            Open ThisWorkbook.Path & "\data\note.txt" For Output As #f
            Print #f, "changed by test"
            Close #f
        End Sub
        """;

    private const string GoodModule = """
        Option Explicit

        Public Function Twice(ByVal x As Long) As Long
            Twice = x * 2
        End Function
        """;

    private const string UndeclaredModule = """
        Option Explicit

        Public Sub Broken()
            notDeclared = 1
        End Sub
        """;

    private const string SyntaxModule = """
        Option Explicit

        Public Sub Broken()
            Dim x As Long
            x = = 1
        End Sub
        """;

    private const string SheetCode = """
        Option Explicit

        Private Sub Worksheet_Change(ByVal Target As Range)
            If Target.Address = "$A$1" Then Me.Range("B1").Value = "変更を見た"
        End Sub
        """;

    private const string BookCode = """
        Option Explicit

        Public Function Hello() As String
            Hello = "ブック"
        End Function
        """;
}
