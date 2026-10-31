using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spike0;

internal sealed partial class ExcelScenario
{
    private readonly Dictionary<string, CellProbe.Cell> cellEvidence = [];
    private Excel.Worksheet? cellSheet;
    private Excel.Range CellRange(string address) => gate.Step("cell-range-" + address,
        () => Track(cellSheet!.Range[address]));
    private void WithCell(string address, Action<Excel.Range> action)
    {
        var range = CellRange(address);
        gate.Step("cell-action-" + address, () => action(range));
    }
    private void Value(string address, object? value) => WithCell(address, range => range.Value2 = value);
    private void Formula2(string address, string formula) => WithCell(address,
        range => { ((dynamic)range).Formula2 = formula; });
    private CellProbe.Cell ObserveCell(string address, bool formula2 = false)
    {
        var range = CellRange(address);
        object? value = gate.Step("cell-read-" + address, () => range.Value2);
        object hasFormula = gate.Step("cell-has-formula-" + address, () => range.HasFormula);
        object? formula = gate.Step("cell-formula-read-" + address,
            () => formula2 ? (object?)((dynamic)range).Formula2 : range.Formula);
        object prefix = gate.Step("cell-prefix-" + address, () => range.PrefixCharacter);
        object format = gate.Step("cell-format-" + address, () => range.NumberFormat);
        log.Write("cell.observed", gate, new { scenario = cellCase, address, value, value_type = value?.GetType().FullName,
            has_formula = hasFormula, formula, prefix, number_format = format, projected_value = CellProbe.Project(value) });
        CellProbe.Require(hasFormula is bool, "single-cell HasFormula type");
        var cell = new CellProbe.Cell(address, value, value?.GetType().FullName, (bool)hasFormula, formula as string);
        cellEvidence[address] = cell;
        return cell;
    }
    private void Matrix(string address, string[] addresses)
    {
        var range = CellRange(address);
        object raw = gate.Step("cell-matrix-" + address, () => range.Value2);
        CellProbe.Require(raw is Array { Rank: 2 }, "matrix type " + address);
        var array = (Array)raw;
        object?[] values = array.Cast<object?>().ToArray();
        log.Write("cell.matrix", gate, new { address, rows = array.GetLength(0), columns = array.GetLength(1),
            row_lower_bound = array.GetLowerBound(0), column_lower_bound = array.GetLowerBound(1),
            values, types = values.Select(item => item?.GetType().FullName).ToArray() });
        CellProbe.Require(values.Length == addresses.Length, "matrix shape");
        for (int i = 0; i < addresses.Length; i++)
        {
            var scalar = cellEvidence[addresses[i]];
            CellProbe.Require(Equals(values[i], scalar.Value) && values[i]?.GetType().FullName == scalar.ValueType,
                "scalar/matrix agreement " + addresses[i]);
        }
    }
    private void Expect(string address, object? expected, bool hasFormula, bool formula2 = false)
    {
        var cell = ObserveCell(address, formula2);
        CellProbe.Require(Equals(cell.Value, expected) && cell.HasFormula == hasFormula, address);
    }
    private void ProbeCells()
    {
        gate.Begin("cell-" + cellCase, TimeSpan.FromSeconds(30));
        WithEventsSuppressed(() =>
        {
            var sheets = gate.Step("cell-sheets", () => Track(workbook!.Worksheets));
            cellSheet = gate.Step("cell-sheet", () => Track((Excel.Worksheet)sheets["Lifecycle"]));
            ObserveSettings("cell-edit");
            log.Write("cell.environment", gate, new
            {
                scenario = cellCase,
                excel_version = gate.Step("cell-excel-version", () => application!.Version),
                excel_build = gate.Step("cell-excel-build", () => application!.Build)
            });
            switch (cellCase)
            {
                case "s1": ProbeStrings(); break;
                case "s2": ProbeClear(); break;
                case "s3": ProbeFormulas(); break;
                case "s4": ProbeErrors(); break;
                default: throw new InvalidDataException("Unknown cell case.");
            }
            gate.Step("cell-write-observations", () => File.WriteAllText(
                Path.Combine(Path.GetDirectoryName(workbookPath)!, "cell-observations.json"),
                JsonSerializer.Serialize(cellEvidence.Values.ToArray())));
        });
        ObserveSettings("cell-after-restoration");
        gate.Complete("ready");
    }
    private void ProbeStrings()
    {
        WithCell("D2:G14", range => range.NumberFormat = "General");
        var prefixed = new object[CellProbe.Strings.Length, 1];
        for (int i = 0; i < CellProbe.Strings.Length; i++)
        {
            string input = CellProbe.Strings[i];
            log.Write("cell.input", gate, new { row = i + 2, input, method = "raw D / prefixed E / prefixed matrix F" });
            Value("D" + (i + 2), input);
            ObserveCell("D" + (i + 2));
            Value("E" + (i + 2), "'" + input);
            Expect("E" + (i + 2), input, false);
            prefixed[i, 0] = "'" + input;
        }
        Value("F2:F14", prefixed);
        for (int i = 0; i < CellProbe.Strings.Length; i++)
            Expect("F" + (i + 2), CellProbe.Strings[i], false);
        Expect("G2", null, false);
        Matrix("E2:E14", Enumerable.Range(2, 13).Select(row => "E" + row).ToArray());
        Matrix("F2:F14", Enumerable.Range(2, 13).Select(row => "F" + row).ToArray());
        foreach (string column in new[] { "E", "F" })
            foreach (int row in Enumerable.Range(2, 13))
            {
                var range = CellRange(column + row);
                CellProbe.Require((string)gate.Step("verify-general", () => range.NumberFormat) == "General", "string format");
            }
    }
    private string Style(string address)
    {
        var range = CellRange(address);
        var font = gate.Step("style-font", () => Track(range.Font));
        var fill = gate.Step("style-fill", () => Track(range.Interior));
        var borders = gate.Step("style-borders", () => Track(range.Borders));
        var merge = gate.Step("style-merge", () => Track(range.MergeArea));
        return JsonSerializer.Serialize(new
        {
            format = gate.Step("style-format", () => range.NumberFormat),
            bold = gate.Step("style-bold", () => font.Bold),
            color = gate.Step("style-color", () => fill.Color),
            line = gate.Step("style-line", () => borders.LineStyle),
            weight = gate.Step("style-weight", () => borders.Weight),
            horizontal = gate.Step("style-horizontal", () => range.HorizontalAlignment),
            vertical = gate.Step("style-vertical", () => range.VerticalAlignment),
            merged = gate.Step("style-merged", () => range.MergeCells),
            merge_address = gate.Step("style-merge-address", () => merge.Address)
        });
    }
    private void SetupClear(string rectangle, string top, string bottom, string rightTop, string rightBottom)
    {
        var range = CellRange(rectangle);
        gate.Step("setup-format", () => range.NumberFormat = "0.00");
        var font = gate.Step("setup-font", () => Track(range.Font));
        gate.Step("setup-bold", () => font.Bold = true);
        var fill = gate.Step("setup-fill", () => Track(range.Interior));
        gate.Step("setup-color", () => fill.Color = 65535);
        var borders = gate.Step("setup-borders", () => Track(range.Borders));
        gate.Step("setup-line", () => borders.LineStyle = Excel.XlLineStyle.xlContinuous);
        gate.Step("setup-weight", () => borders.Weight = Excel.XlBorderWeight.xlThin);
        gate.Step("setup-horizontal", () => range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter);
        gate.Step("setup-vertical", () => range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter);
        Value(top, 5d); Formula2(bottom, "=1+1"); Value(rightTop, 17d); Value(rightBottom, true);
    }
    private void ProbeClear()
    {
        SetupClear("D2:E3", "D2", "D3", "E2", "E3");
        SetupClear("H2:I3", "H2", "H3", "I2", "I3");
        WithCell("D6:E6", range => range.Merge());
        WithCell("H6:I6", range => range.Merge());
        Value("D6", "reference"); Value("H6", "candidate");
        string[] addresses = ["D2", "E2", "D3", "E3", "H2", "I2", "H3", "I3", "D6", "E6", "H6", "I6"];
        var styles = addresses.ToDictionary(address => address, Style);
        foreach (string address in addresses) ObserveCell(address);
        log.Write("cell.styles.before", gate, styles);
        WithCell("D2", range => range.ClearContents());
        WithCell("D3", range => range.ClearContents());
        var anchor = CellRange("D6");
        var merge = gate.Step("cell-get-merge-reference", () => Track(anchor.MergeArea));
        gate.Step("cell-clear-merge-reference", () => merge.ClearContents());
        Expect("D2", null, false); Expect("D3", null, false); Expect("D6", null, false); Expect("E6", null, false);
        Value("H2:I3", new object?[,] { { null, 17d }, { null, true } });
        Value("H6", new object?[1, 1]);
        foreach (string address in new[] { "H2", "H3", "H6", "I6" }) Expect(address, null, false);
        foreach (string address in new[] { "E2", "I2" }) Expect(address, 17d, false);
        foreach (string address in new[] { "E3", "I3" }) Expect(address, true, false);
        Matrix("H2:I3", ["H2", "I2", "H3", "I3"]);
        foreach (string address in addresses)
        {
            string after = Style(address);
            log.Write("cell.style.after", gate, new { address, style = after });
            CellProbe.Require(styles[address] == after, "style/merge preserved " + address);
        }
    }
    private void ProbeFormulas()
    {
        Value("D2:D4", new object[,] { { 1d }, { 4d }, { 9d } });
        Formula2("F2", "=SQRT(D2:D4)"); Formula2("F7", "=SEQUENCE(2,2)"); Formula2("F12", "=1+1");
        WithCell("J2", range => range.Formula = "=SQRT(D2:D4)");
        Value("L2", 11d);
        gate.Step("cell-calculate", () => cellSheet!.Calculate());
        Expect("F2", 1d, true, true); Expect("F3", 2d, false, true); Expect("F4", 3d, false, true);
        Expect("F7", 1d, true, true); Expect("G7", 2d, false, true); Expect("F8", 3d, false, true); Expect("G8", 4d, false, true);
        Expect("F12", 2d, true, true); Expect("J2", 1d, true, true);
        Expect("L2", 11d, false, true); Expect("L3", null, false, true);
        CellProbe.Require(cellEvidence["F2"].Formula == "=SQRT(D2:D4)" &&
            cellEvidence["F7"].Formula == "=SEQUENCE(2,2)" && cellEvidence["F12"].Formula == "=1+1", "Formula2 round trip without @");
        CellProbe.Require(cellEvidence["J2"].Formula == "=SQRT(@D2:D4)", "legacy implicit intersection");
        Matrix("F2:F4", ["F2", "F3", "F4"]); Matrix("F7:G8", ["F7", "G7", "F8", "G8"]);
    }
    private void ProbeErrors()
    {
        Value("D2:D3", new object[,] { { 1d }, { 2d } });
        for (int i = 0; i < CellProbe.Errors.Length; i++) Formula2("H" + (i + 2), CellProbe.Errors[i].Formula);
        // H9 is the SEQUENCE anchor; H10 is its blocker and contains the independent FILTER error.
        gate.Step("cell-calculate", () => cellSheet!.Calculate());
        for (int i = 0; i < CellProbe.Errors.Length; i++)
        {
            var expected = CellProbe.Errors[i];
            log.Write("cell.error.expected", gate, new { address = "H" + (i + 2), expected.Formula, expected.Code, raw = CellProbe.ErrorRaw(expected.Number) });
            Expect("H" + (i + 2), CellProbe.ErrorRaw(expected.Number), true, true);
        }
        object[] controls = [2007d, 2045d, 2050d, (double)CellProbe.ErrorRaw(2007), "#DIV/0!"];
        for (int i = 0; i < controls.Length; i++)
        {
            Value("J" + (i + 2), controls[i] is string text ? "'" + text : controls[i]);
            Expect("J" + (i + 2), controls[i], false);
        }
        Matrix("H2:H10", Enumerable.Range(2, 9).Select(row => "H" + row).ToArray());
        Matrix("J2:J6", Enumerable.Range(2, 5).Select(row => "J" + row).ToArray());
    }
}
