// Throwaway experiment for macro_run (docs/api.md, stage 1). It starts one dedicated Excel on an STA thread,
// imports a test module and the xlwright helpers into a new workbook that has no file, and calls macros in one of
// two ways. "nested": a Run function in the xlwright module calls the macro with Application.Run. "direct": for each
// call, a generated module calls the macro by name, and is removed after the call. It prints what .NET receives.
// A watch ends the owned Excel if one call takes longer than 20 seconds (a dialog that was not caught).
// It only ends the Excel whose ownership it confirmed: a new PID, named EXCEL.EXE, created after the start request.
//
// Usage: dotnet run --project spikes\macro-run -- <nested|direct>

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    private static readonly string Folder = Path.Combine(Path.GetTempPath(), "xlwright-spike-macro-run-" + Guid.NewGuid().ToString("N"));
    private static Process? s_owned;

    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var exit = 1;
        var mode = args.Length == 1 ? args[0] : "";
        if (mode is not ("nested" or "direct"))
        {
            Console.Error.WriteLine("Usage: MacroRun <nested|direct>");
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
        Console.WriteLine($"started PID {process.Id}, owned {owned}");
        if (!owned)
        {
            Console.WriteLine("ownership not confirmed; leaving it untouched");
            return 1;
        }

        s_owned = process;
        app.UserControl = true;
        dynamic workbooks = app.Workbooks;
        dynamic book = workbooks.Add();
        dynamic project = book.VBProject;
        dynamic components = project.VBComponents;
        var ended = false;
        try
        {
            Import(components, "Calc.bas", CalcModule);
            Import(components, "XlWright.bas", XlWrightModule());
            string qualifier = $"'{(string)book.Name}'!";
            if (mode == "nested")
            {
                Call(app, qualifier, "Calc.Twice", new object[] { 21 });
                Call(app, qualifier, "Calc.DivZero", Array.Empty<object>());
            }
            else
            {
                Direct(app, qualifier, components);
            }
        }
        catch (Exception e) when (process.HasExited)
        {
            Console.WriteLine($"Excel ended during the trials: {e.GetType().Name} 0x{e.HResult:X8}");
            ended = true;
        }
        finally
        {
            if (!ended)
            {
                Release(components);
                Release(project);
                app.DisplayAlerts = false;
                book.Close(false);
                Release(book);
                Release(workbooks);
                app.Quit();
                Release(app);
            }

            app = null;
        }

        ended = process.WaitForExit(TimeSpan.FromSeconds(75));
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

    private static void Direct(dynamic app, string qualifier, dynamic components)
    {
        Console.WriteLine("== through a generated module");
        foreach (var (macro, args) in new (string, object[])[]
        {
            ("Calc.Twice", [21]),
            ("Calc.Greet", ["世界"]),
            ("Calc.Sum30", Enumerable.Range(1, 30).Cast<object>().ToArray()),
            ("Calc.Mixed", [1.5, "a", true]),
            ("Calc.ByRefLong", [41]),
            ("Calc.Optional5", []),
            ("Calc.Many", [1, 2, 3]),
            ("Calc.NoReturn", []),
            ("Calc.DivZero", []),
            ("Calc.RaiseCustom", []),
            ("Calc.RaiseInner", []),
            ("Calc.HandlesItself", []),
            ("Calc.Twice", ["x"]),
            ("Calc.ReturnsCurrency", []),
            ("Calc.ReturnsDate", []),
            ("Calc.ReturnsEmpty", []),
            ("Calc.ReturnsNull", []),
            ("Calc.ReturnsArray", []),
            ("Calc.ReturnsCollection", []),
            ("Calc.ReturnsApplication", []),
            ("Calc.ReturnsCellError", []),
            ("Calc.ReturnsLongString", []),
            ("Calc.Hidden", []),
            ("Calc.Missing", []),
            ("Calc.Counter", []),
            ("NoModule.Twice", []),
            ("Calc.EndsAll", []),
            ("Calc.Twice", [21]),
            ("Calc.Twice", []),
        })
        {
            var dot = macro.IndexOf('.');
            var (module, procedure) = (macro[..dot], macro[(dot + 1)..]);
            string declaration;
            try
            {
                dynamic component = components.Item(module);
                dynamic code = component.CodeModule;
                int body = code.ProcBodyLine(procedure, 0); // vbext_pk_Proc
                declaration = ((string)code.Lines(body, 1)).Trim();
                Console.WriteLine($"  {macro}: module type {(int)component.Type}, declaration \"{declaration}\"");
                Release(code);
                Release(component);
            }
            catch (Exception e)
            {
                Console.WriteLine($"  {macro}: lookup {Describe(e)}");
                continue;
            }

            if (declaration.StartsWith("Private", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"  {macro}: private; not called");
                continue;
            }

            var isFunction = declaration.Contains("Function ", StringComparison.OrdinalIgnoreCase);
            var list = string.Join(", ", Enumerable.Range(0, args.Length).Select(i => $"(args({i}))"));
            var call = isFunction
                ? $"    Run = XlWright.Result({macro}({list}))\n"
                : $"    Call {macro}{(args.Length == 0 ? "" : $"({list})")}\n    Run = Array(0, Empty)\n";
            var generated = "Option Explicit\n\nPublic Function Run(ByVal args As Variant) As Variant\n    On Error GoTo Failed\n" + call +
                "    Exit Function\nFailed:\n    Run = Array(1, Err.Number, Err.Description, Err.Source)\nEnd Function\n";

            var watch = Stopwatch.StartNew();
            dynamic added = components.Add(1); // vbext_ct_StdModule
            added.Name = "XlWrightCall";
            dynamic addedCode = added.CodeModule;
            addedCode.AddFromString(generated.ReplaceLineEndings("\r\n"));
            Release(addedCode);
            var prepared = watch.ElapsedMilliseconds;
            try
            {
                Call(app, qualifier, "XlWrightCall.Run", args, direct: true);
            }
            finally
            {
                if (!s_owned!.HasExited)
                {
                    watch.Restart();
                    components.Remove(added);
                    Console.WriteLine($"    generated module: added in {prepared} ms, removed in {watch.ElapsedMilliseconds} ms");
                }

                Release(added);
            }
        }
    }

    private static void Call(dynamic app, string qualifier, string macro, object[] args, bool direct = false)
    {
        using var watch = new Timer(_ =>
        {
            Console.WriteLine($"  {macro}: no return within 20 s; ending the owned Excel");
            s_owned!.Kill();
        }, null, TimeSpan.FromSeconds(20), Timeout.InfiniteTimeSpan);
        var elapsed = Stopwatch.StartNew();
        try
        {
            object? result = direct ? app.Run(qualifier + macro, args) : app.Run(qualifier + "XlWright.Run", macro, args);
            Console.WriteLine($"  {macro}({args.Length} args): {Show(result)} in {elapsed.ElapsedMilliseconds} ms");
        }
        catch (Exception e)
        {
            Console.WriteLine($"  {macro}: {Describe(e)} in {elapsed.ElapsedMilliseconds} ms");
            if (s_owned!.HasExited) throw;
        }
    }

    private static string Show(object? value) => value switch
    {
        null => "null",
        Array array => "[" + string.Join(", ", array.Cast<object?>().Select(Show)) + "]",
        string text => $"String({text.Length}) {(text.Length > 40 ? text[..40] + "…" : text)}",
        _ => $"{value.GetType().Name} {value}",
    };

    private static string Describe(Exception e)
    {
        var inner = e is System.Reflection.TargetInvocationException { InnerException: { } i } ? i : e;
        return $"{inner.GetType().Name} 0x{inner.HResult:X8} {inner.Message.ReplaceLineEndings(" ")}";
    }

    private static void Import(dynamic components, string file, string text)
    {
        var path = Path.Combine(Folder, file);
        File.WriteAllText(path, text.ReplaceLineEndings("\r\n"), Encoding.GetEncoding(932));
        dynamic imported = components.Import(path);
        Console.WriteLine($"imported {(string)imported.Name}");
        Release(imported);
    }

    private static void Release(object comObject) => Marshal.FinalReleaseComObject(comObject);

    // The candidate: calls the macro with Application.Run, catches a run-time error, and returns
    // Array(0, value), Array(1, number, description, source), or Array(2, type name) for a value COM cannot carry.
    private static string XlWrightModule()
    {
        var cases = new StringBuilder();
        for (var count = 0; count <= 30; count++)
        {
            var list = string.Concat(Enumerable.Range(0, count).Select(i => $", args(first + {i})"));
            cases.Append($"    Case {count}: result = Application.Run(macro{list})\n");
        }

        return $$"""
            Attribute VB_Name = "XlWright"
            Option Explicit

            Public Function Run(ByVal macro As String, ByVal args As Variant) As Variant
                Dim result As Variant
                Dim first As Long
                first = LBound(args)
                On Error GoTo Failed
                Select Case UBound(args) - first + 1
            {{cases}}    End Select
                On Error GoTo 0
                If IsArray(result) Then
                    Run = Array(2, TypeName(result))
                Else
                    Run = Array(0, result)
                End If
                Exit Function
            Failed:
                Run = Array(1, Err.Number, Err.Description, Err.Source)
            End Function

            ' A ByVal Variant keeps an object as an object, so its default member is not read.
            Public Function Result(ByVal value As Variant) As Variant
                If IsObject(value) Or IsArray(value) Then
                    Result = Array(2, TypeName(value))
                Else
                    Result = Array(0, value)
                End If
            End Function
            """;
    }

    private const string CalcModule = """
        Attribute VB_Name = "Calc"
        Option Explicit

        Public Function Twice(ByVal x As Long) As Long
            Twice = x * 2
        End Function

        Public Function Greet(ByVal name As String) As String
            Greet = "こんにちは、" & name
        End Function

        Public Function Sum30(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, _
                              a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29, a30) As Long
            Sum30 = a1 + a2 + a3 + a4 + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14 + a15 + _
                    a16 + a17 + a18 + a19 + a20 + a21 + a22 + a23 + a24 + a25 + a26 + a27 + a28 + a29 + a30
        End Function

        Public Function Mixed(ByVal d As Double, ByVal s As String, ByVal b As Boolean) As String
            Mixed = TypeName(d) & "," & TypeName(s) & "," & TypeName(b)
        End Function

        Public Function ByRefLong(ByRef x As Long) As Long
            x = x + 1
            ByRefLong = x
        End Function

        Public Function Optional5(Optional ByVal x As Long = 5) As Long
            Optional5 = x
        End Function

        Public Function Many(ParamArray xs() As Variant) As Long
            Many = UBound(xs) - LBound(xs) + 1
        End Function

        Public Sub NoReturn()
            ThisWorkbook.Worksheets(1).Range("A1").Value = "written"
        End Sub

        Public Sub DivZero()
            Dim a As Double
            a = 1 / 0
        End Sub

        Public Sub RaiseCustom()
            Err.Raise vbObjectError + 513, "Calc.RaiseCustom", "カスタムのエラー"
        End Sub

        Public Sub RaiseInner()
            Inner
        End Sub

        Private Sub Inner()
            Dim a As Variant
            a = CLng("not a number")
        End Sub

        Public Function HandlesItself() As Long
            On Error Resume Next
            Dim a As Double
            a = 1 / 0
            HandlesItself = Err.Number
        End Function

        Public Function ReturnsCurrency() As Currency
            ReturnsCurrency = 1.25
        End Function

        Public Function ReturnsDate() As Date
            ReturnsDate = DateSerial(2026, 11, 3)
        End Function

        Public Function ReturnsEmpty() As Variant
        End Function

        Public Function ReturnsNull() As Variant
            ReturnsNull = Null
        End Function

        Public Function ReturnsArray() As Variant
            ReturnsArray = Array(1, 2)
        End Function

        Public Function ReturnsCollection() As Object
            Set ReturnsCollection = New Collection
        End Function

        Public Function ReturnsApplication() As Object
            Set ReturnsApplication = Application
        End Function

        Public Function ReturnsCellError() As Variant
            ReturnsCellError = CVErr(2042)
        End Function

        Public Function ReturnsLongString() As String
            ReturnsLongString = String(30000, "x")
        End Function

        Private Function Hidden() As Long
            Hidden = 7
        End Function

        Public Sub EndsAll()
            End
        End Sub
        """;
}
