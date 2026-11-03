// Trial 6 (stage 1): import standard and class modules into an .xlsm that the experiment creates under
// .xlwright/spikes (not tracked by Git), replacing them all, and run them with Application.Run.
// Every Excel call that may hang (a macro, a Close after a failed macro) runs under a watchdog that
// terminates the owned Excel through the handle opened at ownership.

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

internal static class VbaTrials
{
    private const int StdModule = 1;
    private const int ClassModule = 2;
    private const string Book = "App.xlsm";
    private const string Qualifier = "'" + Book + "'!";

    private static string s_app = "";
    private static Encoding s_ansi = Encoding.ASCII;

    public static void Run()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        s_ansi = Encoding.GetEncoding(932);
        s_app = Path.Combine(Program.RepositoryRoot(), ".xlwright", "spikes", "vba", "App");
        if (Directory.Exists(s_app)) Directory.Delete(s_app, recursive: true);
        WriteSources();
        var book = Path.Combine(s_app, Book);

        Session("vba/setup", (app, owned) => Setup(app, book));
        Session("vba/run", (app, owned) => RunMacros(app, owned, book));
        Session("vba/runtime-error", (app, owned) => RunOne(app, owned, book, "Calc.RaiseError", seconds: 20));
        Session("vba/compile-error", (app, owned) => CompileError(app, owned, book));
        Session("vba/msgbox", (app, owned) => RunOne(app, owned, book, "Calc.ShowMessage", seconds: 20));
        Session("vba/kill-during-macro", (app, owned) => RunOne(app, owned, book, "Calc.Forever", seconds: 5));

        Directory.Delete(s_app, recursive: true);
        Program.Log($"vba: removed {s_app.Replace(Program.RepositoryRoot(), "<repo>")}");
    }

    private static void Setup(dynamic app, string book)
    {
        Program.Log($"vba/setup: AutomationSecurity at start={(int)app.AutomationSecurity}");
        dynamic workbooks = app.Workbooks;
        dynamic workbook = workbooks.Add();
        workbook.SaveAs(book, 52); // xlOpenXMLWorkbookMacroEnabled
        dynamic project = workbook.VBProject;
        dynamic components = project.VBComponents;

        for (var round = 1; round <= 2; round++)
        {
            var watch = Stopwatch.StartNew();
            ReplaceAll(components, Path.Combine(s_app, "src"));
            Program.Log($"vba/setup: full replace #{round} in {watch.ElapsedMilliseconds} ms -> {ListComponents(components)}");
        }

        foreach (var name in new[] { "Calc", "Counter" })
        {
            dynamic component = components.Item(name);
            var exported = Path.Combine(s_app, name + ".exported");
            component.Export(exported);
            var attributes = File.ReadAllLines(exported, s_ansi).Where(l => l.StartsWith("Attribute ")).ToList();
            Program.Log($"vba/setup: exported {name} attributes: {string.Join(" | ", attributes)}");
            Marshal.FinalReleaseComObject(component);
        }

        // Test setup only: an Open event that marks D1. Sheet and workbook modules are stage 3.
        dynamic thisWorkbook = components.Item("ThisWorkbook");
        dynamic code = thisWorkbook.CodeModule;
        code.AddFromString("Private Sub Workbook_Open()\r\n    Me.Worksheets(1).Range(\"D1\").Value = \"opened\"\r\nEnd Sub");
        Marshal.FinalReleaseComObject(code);
        Marshal.FinalReleaseComObject(thisWorkbook);
        Marshal.FinalReleaseComObject(components);
        Marshal.FinalReleaseComObject(project);

        workbook.Save();
        workbook.Close(false);
        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(workbooks);
    }

    // Removes every standard and class module, then imports every file under src/modules and src/classes.
    private static void ReplaceAll(dynamic components, string src)
    {
        for (var i = (int)components.Count; i >= 1; i--)
        {
            dynamic component = components.Item(i);
            int type = component.Type;
            if (type == StdModule || type == ClassModule) components.Remove(component);
            Marshal.FinalReleaseComObject(component);
        }

        foreach (var file in Directory.GetFiles(Path.Combine(src, "modules"), "*.bas").Concat(Directory.GetFiles(Path.Combine(src, "classes"), "*.cls")))
        {
            dynamic imported = components.Import(file);
            Marshal.FinalReleaseComObject(imported);
        }
    }

    private static string ListComponents(dynamic components)
    {
        var names = new List<string>();
        for (var i = 1; i <= (int)components.Count; i++)
        {
            dynamic component = components.Item(i);
            names.Add($"{(string)component.Name}({(int)component.Type})");
            Marshal.FinalReleaseComObject(component);
        }

        return string.Join(", ", names);
    }

    private static void RunMacros(dynamic app, Owned owned, string book)
    {
        dynamic workbooks = app.Workbooks;
        dynamic workbook = OpenWithEventsOff(app, workbooks, book, "vba/run");

        Call(owned, "Twice(21)", () => app.Run(Qualifier + "Calc.Twice", 21));
        Call(owned, "Echo(\"あいう\")", () => app.Run(Qualifier + "Calc.Echo", "あいう"));
        Call(owned, "Calc.Japanese (CP932 file)", () => app.Run(Qualifier + "Calc.Japanese"), expected: "日本語");
        Call(owned, "CharsetUtf8.Japanese (UTF-8 file)", () => app.Run(Qualifier + "CharsetUtf8.Japanese"), expected: "日本語");
        Call(owned, "UseCounter (PredeclaredId, default member)", () => app.Run(Qualifier + "Calc.UseCounter"));
        Call(owned, "WriteResult", () => app.Run(Qualifier + "Calc.WriteResult"));
        Program.Log($"vba/run: C1 after WriteResult={ReadCell(workbook, "C1")}");
        Call(owned, "'NoSuch.xlsm'!Calc.Twice", () => app.Run("'NoSuch.xlsm'!Calc.Twice", 1));
        Program.Log($"vba/run: Workbooks after the NoSuch call={(int)workbooks.Count}");

        workbook.Close(false);
        Marshal.FinalReleaseComObject(workbook);

        // Control: with events on, the Open event runs.
        app.EnableEvents = true;
        workbook = workbooks.Open(book);
        Program.Log($"vba/run: opened with EnableEvents=True, D1={ReadCell(workbook, "D1")}");
        workbook.Close(false);
        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(workbooks);
    }

    private static void RunOne(dynamic app, Owned owned, string book, string macro, int seconds)
    {
        dynamic workbooks = app.Workbooks;
        dynamic workbook = OpenWithEventsOff(app, workbooks, book, "vba/" + macro);
        Call(owned, macro, () => app.Run(Qualifier + macro), seconds: seconds);
        if (owned.HasExited) return;
        Call(owned, "Close after " + macro, () => { workbook.Close(false); return null; }, seconds: 15);
        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(workbooks);
    }

    private static void CompileError(dynamic app, Owned owned, string book)
    {
        dynamic workbooks = app.Workbooks;
        dynamic workbook = OpenWithEventsOff(app, workbooks, book, "vba/compile-error");
        dynamic project = workbook.VBProject;
        dynamic components = project.VBComponents;
        dynamic imported = components.Import(Path.Combine(s_app, "Broken.bas"));
        foreach (var o in new object[] { imported, components, project }) Marshal.FinalReleaseComObject(o);
        Call(owned, "Broken.Broken (compile error)", () => app.Run(Qualifier + "Broken.Broken"), seconds: 20);
        if (owned.HasExited) return;
        Call(owned, "Calc.Twice(1) after the compile error", () => app.Run(Qualifier + "Calc.Twice", 1), seconds: 20);
        if (owned.HasExited) return;
        Call(owned, "Close after the compile error", () => { workbook.Close(false); return null; }, seconds: 15);
        Marshal.FinalReleaseComObject(workbook);
        Marshal.FinalReleaseComObject(workbooks);
    }

    private static dynamic OpenWithEventsOff(dynamic app, dynamic workbooks, string book, string label)
    {
        app.AutomationSecurity = 1; // msoAutomationSecurityLow
        app.EnableEvents = false;
        dynamic workbook = workbooks.Open(book);
        app.EnableEvents = true;
        Program.Log($"{label}: opened with AutomationSecurity=Low, EnableEvents=False; D1={ReadCell(workbook, "D1")}");
        return workbook;
    }

    private static string ReadCell(dynamic workbook, string address)
    {
        dynamic sheets = workbook.Worksheets;
        dynamic sheet = sheets[1];
        dynamic cell = sheet.Range[address];
        object? value = cell.Value2;
        foreach (var o in new object[] { cell, sheet, sheets }) Marshal.FinalReleaseComObject(o);
        return value == null ? "<empty>" : $"\"{value}\"";
    }

    // Runs one Excel call. A watchdog polls every 250 ms for a dialog (docs/architecture.md: XLMAIN disabled, and a
    // visible enabled window whose owner chain reaches XLMAIN). On a dialog it reads the text and terminates the
    // owned Excel at once; otherwise it terminates the owned Excel when the call does not return in time.
    private static void Call(Owned owned, string label, Func<object?> call, string? expected = null, int seconds = 30)
    {
        using var done = new ManualResetEventSlim();
        var watchdog = new Thread(() =>
        {
            var started = Stopwatch.StartNew();
            while (!done.Wait(250))
            {
                var dialog = DialogProbe.Find(owned.Pid, out var windows, out var rule);
                if (dialog != IntPtr.Zero)
                {
                    var read = Stopwatch.StartNew();
                    var text = DialogProbe.ReadText(dialog, budgetMs: 500);
                    Program.Log($"vba watchdog ({label}): dialog found by the {rule} rule {started.ElapsedMilliseconds} ms after the call; windows: {windows}");
                    Program.Log($"vba watchdog ({label}): dialog text read in {read.ElapsedMilliseconds} ms: {text}");
                    owned.Terminate($"vba watchdog ({label}, dialog)");
                    return;
                }

                if (started.ElapsedMilliseconds >= seconds * 1000)
                {
                    Program.Log($"vba watchdog ({label}): no dialog found within {seconds} s; windows: {windows}");
                    owned.Terminate($"vba watchdog ({label}, {seconds} s)");
                    return;
                }
            }
        }) { IsBackground = true };
        watchdog.Start();
        var watch = Stopwatch.StartNew();
        try
        {
            var result = call();
            var match = expected == null ? "" : $", equals expected={Equals(result, expected)}";
            Program.Log($"vba: {label} -> {Describe(result)} in {watch.ElapsedMilliseconds} ms{match}");
        }
        catch (Exception e)
        {
            Program.Log($"vba: {label} -> {e.GetType().Name} 0x{e.HResult:X8} in {watch.ElapsedMilliseconds} ms: {e.Message.Trim()}");
        }
        finally
        {
            done.Set();
            watchdog.Join();
        }
    }

    private static string Describe(object? value) => value switch
    {
        null => "<null>",
        string s => $"\"{s}\" (UTF-16 {string.Join(" ", s.Select(c => ((int)c).ToString("X4")))})",
        _ => $"{value} ({value.GetType().Name})",
    };

    // One owned Excel per session. Leftover workbooks (all under the experiment's folder) are closed without saving.
    private static void Session(string label, Action<dynamic, Owned> body)
    {
        Program.RunSta(label, () =>
        {
            var filter = MessageFilter.Register(label);
            var (app, owned) = Program.StartAndOwn(label);
            app.UserControl = true;
            try
            {
                body(app, owned);
            }
            finally
            {
                if (!owned.HasExited)
                {
                    Call(owned, label + " close leftovers", () =>
                    {
                        dynamic workbooks = app.Workbooks;
                        for (var i = (int)workbooks.Count; i >= 1; i--)
                        {
                            dynamic workbook = workbooks.Item(i);
                            workbook.Close(false);
                            Marshal.FinalReleaseComObject(workbook);
                        }

                        Marshal.FinalReleaseComObject(workbooks);
                        return null;
                    }, seconds: 15);
                }

                if (owned.HasExited)
                {
                    Program.Log($"{label}: the owned Excel had ended; no Quit");
                    try { Marshal.FinalReleaseComObject(app); } catch (Exception) { }
                    owned.Dispose();
                }
                else
                {
                    Program.QuitAndWait(label, app, owned);
                }

                filter.Report();
                MessageFilter.Revoke();
            }
        });
    }

    private static void WriteSources()
    {
        var modules = Directory.CreateDirectory(Path.Combine(s_app, "src", "modules")).FullName;
        var classes = Directory.CreateDirectory(Path.Combine(s_app, "src", "classes")).FullName;

        WriteAnsi(Path.Combine(modules, "Calc.bas"), """
            Attribute VB_Name = "Calc"
            Option Explicit
            ' 日本語のコメント

            Public Function Twice(ByVal x As Long) As Long
            Attribute Twice.VB_Description = "Doubles x"
                Twice = x * 2
            End Function

            Public Function Echo(ByVal s As String) As String
                Echo = s & "!"
            End Function

            Public Function Japanese() As String
                Japanese = "日本語"
            End Function

            Public Function UseCounter() As Long
                Counter.Increment
                Counter.Increment
                UseCounter = Counter
            End Function

            Public Sub WriteResult()
                ThisWorkbook.Worksheets(1).Range("C1").Value = "written"
            End Sub

            Public Sub RaiseError()
                Dim x As Long
                x = 1 / 0
            End Sub

            Public Sub ShowMessage()
                MsgBox "テストのメッセージ", vbOKOnly, "テストの題名"
            End Sub

            Public Sub Forever()
                Dim i As Long
                Do
                    i = (i + 1) Mod 1000
                Loop
            End Sub
            """);

        File.WriteAllText(Path.Combine(modules, "CharsetUtf8.bas"), Crlf("""
            Attribute VB_Name = "CharsetUtf8"
            Option Explicit

            Public Function Japanese() As String
                Japanese = "日本語"
            End Function
            """), new UTF8Encoding(false));

        WriteAnsi(Path.Combine(classes, "Counter.cls"), """
            VERSION 1.0 CLASS
            BEGIN
              MultiUse = -1  'True
            END
            Attribute VB_Name = "Counter"
            Attribute VB_GlobalNameSpace = False
            Attribute VB_Creatable = False
            Attribute VB_PredeclaredId = True
            Attribute VB_Exposed = False
            Option Explicit

            Private m As Long

            Public Sub Increment()
                m = m + 1
            End Sub

            Public Property Get Value() As Long
            Attribute Value.VB_UserMemId = 0
                Value = m
            End Property
            """);

        // Imported only in the compile-error session, outside src/ so the full replace does not pick it up.
        WriteAnsi(Path.Combine(s_app, "Broken.bas"), """
            Attribute VB_Name = "Broken"
            Option Explicit

            Public Sub Broken()
                undeclared = 1
            End Sub
            """);
    }

    private static void WriteAnsi(string path, string text) => File.WriteAllText(path, Crlf(text), s_ansi);

    private static string Crlf(string text) => text.ReplaceLineEndings("\r\n") + "\r\n";
}
