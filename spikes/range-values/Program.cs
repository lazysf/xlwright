// Throwaway experiment for range_read and range_write (docs/api.md, stage 1).
// It starts one dedicated Excel on an STA thread, works on a new workbook that has no file, and prints
// the CLR types and values that late-bound Value2 and Formula return, and the exceptions writes raise.
// It only ends the Excel whose ownership it confirmed: a new PID, named EXCEL.EXE, created after the start request.
//
// Usage: dotnet run --project spikes\range-values

using System.Diagnostics;
using System.Runtime.InteropServices;

internal static class Program
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    private static int Main()
    {
        var exit = 1;
        var thread = new Thread(() => exit = Run());
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        return exit;
    }

    private static int Run()
    {
        var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Console.WriteLine($"before: {before.Count} EXCEL process(es)");
        var requested = DateTime.Now;
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
        GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out var rawPid);
        var process = Process.GetProcessById((int)rawPid);
        var owned = !before.Contains(process.Id) &&
            process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) &&
            process.StartTime > requested.AddSeconds(-1);
        Console.WriteLine($"started PID {process.Id}, owned {owned}, version {app.Version}");
        if (!owned)
        {
            Console.WriteLine("ownership not confirmed; leaving it untouched");
            return 1;
        }

        app.UserControl = true;
        try
        {
            dynamic book = app.Workbooks.Add();
            dynamic sheet = book.Worksheets.Item(1);
            Trial("values round trip", () => ValuesRoundTrip(sheet));
            Trial("cell errors", () => CellErrors(app, sheet));
            Trial("formulas", () => Formulas(sheet));
            Trial("write failures", () => WriteFailures(sheet));
            Trial("addresses", () => Addresses(book, sheet));
            Trial("calculation", () => Calculation(app, sheet));
            Trial("size", () => Size(sheet));
            app.DisplayAlerts = false;
            book.Close(false);
        }
        finally
        {
            app.Quit();
            app = null!;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        var ended = process.WaitForExit(TimeSpan.FromSeconds(75));
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

    private static void Trial(string name, Action body)
    {
        Console.WriteLine($"== {name}");
        try
        {
            body();
        }
        catch (Exception e)
        {
            Console.WriteLine($"  trial failed: {Describe(e)}");
        }
    }

    private static void ValuesRoundTrip(dynamic sheet)
    {
        var input = new object?[,]
        {
            { 1.5, "abc", true, null, "1/2", "001" },
            { "'x", "=1+1", "日本語", 42, "2026-11-03", "TRUE" },
        };
        dynamic range = sheet.Range("A1:F2");
        range.Value2 = input;
        Print("A1:F2 Value2", range.Value2);
        Print("A1 Value2 (single cell)", sheet.Range("A1").Value2);
        Print("A1:A1 Value2", sheet.Range("A1:A1").Value2);
        Print("D1 (written null) Value2", sheet.Range("D1").Value2);
        Print("B2 Formula (written as value \"=1+1\")", sheet.Range("B2").Formula);

        var zeroBased = new object?[,] { { "z0", "z1" } };
        sheet.Range("H1:I1").Value2 = zeroBased;
        Print("H1:I1 after a zero-based array", sheet.Range("H1:I1").Value2);

        sheet.Range("J1").Value2 = "keep";
        sheet.Range("J1:J1").Value2 = new object?[,] { { null } };
        Print("J1 after writing null over \"keep\"", sheet.Range("J1").Value2);

        sheet.Range("K1").Value2 = 45964.0;
        sheet.Range("K1").NumberFormat = "yyyy-mm-dd";
        Print("K1 date-formatted Value2", sheet.Range("K1").Value2);
        Print("K1 Value (not Value2)", sheet.Range("K1").Value);
        sheet.Range("L1").Value2 = 12.34;
        sheet.Range("L1").NumberFormat = "$#,##0.00";
        Print("L1 currency-formatted Value2", sheet.Range("L1").Value2);
        Print("L1 currency-formatted Value", sheet.Range("L1").Value);
    }

    private static void CellErrors(dynamic app, dynamic sheet)
    {
        sheet.Cells.ClearContents();
        var formulas = new[]
        {
            "=1/0", "=NA()", "=SQRT(-1)", "=NOSUCHFUNC()", "=\"a\"+1", "=A1 B1", "=#REF!", "=SEQUENCE(3)", "=LAMBDA(x,x)",
        };
        for (var i = 0; i < formulas.Length; i++)
        {
            sheet.Cells(i + 1, 1).Formula = formulas[i];
        }

        sheet.Range("H3").Value2 = "blocks the spill";
        sheet.Range("H1").Formula = "=SEQUENCE(5)";
        app.Calculate();
        dynamic values = sheet.Range($"A1:A{formulas.Length}").Value2;
        dynamic functions = app.WorksheetFunction;
        for (var i = 0; i < formulas.Length; i++)
        {
            object? value = values[i + 1, 1];
            dynamic cell = sheet.Cells(i + 1, 1);
            Console.WriteLine($"  {formulas[i],-16} Value2 {Show(value),-34} Text {(string)cell.Text,-10} IsError {(bool)functions.IsError(cell)}");
        }

        object? spill = sheet.Range("H1").Value2;
        Console.WriteLine($"  =SEQUENCE(5) blocked: Value2 {Show(spill)} Text {(string)sheet.Range("H1").Text}");
        sheet.Range("C1").Value2 = -2146826281.0;
        Print("C1 a number equal to the #DIV/0! HRESULT", sheet.Range("C1").Value2);
    }

    private static void Formulas(dynamic sheet)
    {
        sheet.Cells.ClearContents();
        sheet.Range("A1:C1").Value2 = new object?[,] { { 1, "text", null } };
        sheet.Range("A2").Formula = "=SUM(A1,1)";
        sheet.Range("B2").Formula = "=IF(A1>0,\"yes\",\"no\")";
        Print("A1:C2 Formula", sheet.Range("A1:C2").Formula);
        Print("A2 Formula (single cell)", sheet.Range("A2").Formula);
        Print("C1 Formula (empty)", sheet.Range("C1").Formula);

        sheet.Range("A4:C4").Formula = new object?[,] { { "=1+2", "plain", "" } };
        Print("A4:C4 Value2 after writing formulas incl. a constant and \"\"", sheet.Range("A4:C4").Value2);
        Print("A4:C4 Formula", sheet.Range("A4:C4").Formula);
        sheet.Range("A5").Formula = "=SUMME(1,2)";
        Print("A5 a localized function name through Formula", sheet.Range("A5").Value2);
        sheet.Range("A6").Formula = "=SEQUENCE(2)";
        Print("A6:A7 after =SEQUENCE(2) through Formula", sheet.Range("A6:A7").Value2);
        Print("A6 Formula after the spill", sheet.Range("A6").Formula);
    }

    private static void WriteFailures(dynamic sheet)
    {
        sheet.Cells.ClearContents();
        Attempt("Formula \"=SUM(\"", () => sheet.Range("A1").Formula = "=SUM(");
        Attempt("Formula array with \"=SUM(\" second", () =>
            sheet.Range("A2:C2").Formula = new object?[,] { { "=1", "=SUM(", "=3" } });
        Print("A2:C2 after the failed array write", sheet.Range("A2:C2").Value2);
        Attempt("Value2 array larger than the range", () => sheet.Range("A3:B3").Value2 = new object?[,] { { 1, 2, 3 } });
        Print("A3:B3", sheet.Range("A3:B3").Value2);
        Attempt("Value2 array smaller than the range", () => sheet.Range("A4:C4").Value2 = new object?[,] { { 1, 2 } });
        Print("A4:C4", sheet.Range("A4:C4").Value2);
        Attempt("Value2 a decimal", () => sheet.Range("A5").Value2 = 1.25m);
        Attempt("Value2 a long", () => sheet.Range("A6").Value2 = 9007199254740993L);
        Print("A5:A6", sheet.Range("A5:A6").Value2);

        sheet.Range("E1:F2").Merge();
        Attempt("Value2 over part of a merged cell", () => sheet.Range("E1:E2").Value2 = new object?[,] { { 1 }, { 2 } });
        Attempt("ClearContents over part of a merged cell", () => sheet.Range("E1").ClearContents());
        sheet.Range("E1:F2").UnMerge();

        sheet.Protect();
        Attempt("Value2 on a protected sheet", () => sheet.Range("A1").Value2 = 1);
        Attempt("Formula on a protected sheet", () => sheet.Range("A1").Formula = "=1");
        Attempt("ClearContents on a protected sheet", () => sheet.Range("A1").ClearContents());
        Attempt("Read on a protected sheet", () => Print("A1", sheet.Range("A1").Value2));
        sheet.Unprotect();
    }

    private static void Addresses(dynamic book, dynamic sheet)
    {
        foreach (var address in new[] { "a1:b2", "$A$1:$B$2", "B2:A1", "XFE1", "A0", "A1048577", "Sheet2!A1", "A1,B2", "A:A", "1:1", "NoSuchName" })
        {
            Attempt($"Range(\"{address}\")", () =>
            {
                dynamic range = sheet.Range(address);
                Console.WriteLine($"    Address {(string)range.Address} Areas {(int)range.Areas.Count} CountLarge {range.CountLarge}");
            });
        }

        Attempt("Worksheets.Item(\"NoSuchSheet\")", () => book.Worksheets.Item("NoSuchSheet"));
        Attempt("Worksheets.Item(\"sheet1\") (lower case)", () => Console.WriteLine($"    {(string)book.Worksheets.Item("sheet1").Name}"));
    }

    private static void Calculation(dynamic app, dynamic sheet)
    {
        sheet.Cells.ClearContents();
        Console.WriteLine($"  Calculation {app.Calculation} CalculationState {app.CalculationState}");
        app.Calculation = -4135; // xlCalculationManual
        sheet.Range("A1").Value2 = 1;
        sheet.Range("A2").Formula = "=A1*10";
        sheet.Range("A1").Value2 = 2;
        Print("A2 before Calculate (manual)", sheet.Range("A2").Value2);
        Console.WriteLine($"  CalculationState before Calculate {app.CalculationState}");
        app.Calculate();
        Console.WriteLine($"  CalculationState after Calculate {app.CalculationState}");
        Print("A2 after Calculate", sheet.Range("A2").Value2);
        app.Calculation = -4105; // xlCalculationAutomatic
    }

    private static void Size(dynamic sheet)
    {
        sheet.Cells.ClearContents();
        var data = new object?[100, 100];
        for (var r = 0; r < 100; r++)
        for (var c = 0; c < 100; c++)
        {
            data[r, c] = r * 100 + c;
        }

        var watch = Stopwatch.StartNew();
        sheet.Range("A1:CV100").Value2 = data;
        Console.WriteLine($"  write 10,000 cells: {watch.ElapsedMilliseconds} ms");
        watch.Restart();
        object[,] read = sheet.Range("A1:CV100").Value2;
        Console.WriteLine($"  read 10,000 cells: {watch.ElapsedMilliseconds} ms, lower bounds {read.GetLowerBound(0)},{read.GetLowerBound(1)}");
        watch.Restart();
        sheet.Range("A1:XFD1048576").ClearContents();
        Console.WriteLine($"  ClearContents on the whole sheet: {watch.ElapsedMilliseconds} ms");
    }

    private static void Attempt(string name, Action body)
    {
        try
        {
            body();
            Console.WriteLine($"  {name}: no exception");
        }
        catch (Exception e)
        {
            Console.WriteLine($"  {name}: {Describe(e)}");
        }
    }

    private static void Print(string label, object? value)
    {
        if (value is Array array && array.Rank == 2)
        {
            Console.WriteLine($"  {label}: {array.GetType().Name} [{array.GetLowerBound(0)}..{array.GetUpperBound(0)}, {array.GetLowerBound(1)}..{array.GetUpperBound(1)}]");
            for (var r = array.GetLowerBound(0); r <= array.GetUpperBound(0); r++)
            {
                var cells = new List<string>();
                for (var c = array.GetLowerBound(1); c <= array.GetUpperBound(1); c++)
                {
                    cells.Add(Show(array.GetValue(r, c)));
                }

                Console.WriteLine("    " + string.Join(" | ", cells));
            }

            return;
        }

        Console.WriteLine($"  {label}: {Show(value)}");
    }

    private static string Show(object? value) => value is null ? "null" : $"{value.GetType().Name} {value}";

    private static string Describe(Exception e)
    {
        var inner = e is System.Reflection.TargetInvocationException { InnerException: { } i } ? i : e;
        return $"{inner.GetType().Name} 0x{inner.HResult:X8} {inner.Message.ReplaceLineEndings(" ")}";
    }
}
