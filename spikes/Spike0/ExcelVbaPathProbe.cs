using Excel = Microsoft.Office.Interop.Excel;

namespace Spike0;

internal sealed partial class ExcelScenario
{
    private void ProbeVbaPath()
    {
        gate.Begin("t15-probe", TimeSpan.FromSeconds(30));
        object Tracked(string stage, Func<object> read) => gate.Step(stage, () => Track(read()));
        log.Write("t15.environment", gate, new
        {
            version = gate.Step("t15-version", () => application!.Version),
            build = gate.Step("t15-build", () => application!.Build)
        });
        var sheets = gate.Step("t15-sheets", () => Track(workbook!.Worksheets));
        var sheet = gate.Step("t15-sheet", () => Track((Excel.Worksheet)sheets["Lifecycle"]));
        var cell = gate.Step("t15-cell", () => Track(sheet.Range["A2"]));
        if (gate.Step("t15-read-initial", () => cell.Value2) is not double initial || initial != 0)
            throw new InvalidDataException("The fixture cell does not contain the expected numeric zero.");

        object project = Tracked("t15-vbproject", () => VbaPathProbe.Get(workbook!, "VBProject"));
        int protection = gate.Step("t15-protection", () => Convert.ToInt32(VbaPathProbe.Get(project, "Protection")));
        object components = Tracked("t15-components", () => VbaPathProbe.Get(project, "VBComponents"));
        int before = gate.Step("t15-count-before", () => Convert.ToInt32(VbaPathProbe.Get(components, "Count")));
        log.Write("t15.vbproject", gate, new { protection, components_before = before });

        string modulePath = VbaPathProbe.ModulePath(Path.GetDirectoryName(workbookPath)!);
        object? imported = null;
        MeasureS11("vba-import", () => imported = Track(VbaPathProbe.Call(components, "Import", modulePath)));
        string name = gate.Step("t15-imported-name", () => (string)VbaPathProbe.Get(imported!, "Name"));
        int type = gate.Step("t15-imported-type", () => Convert.ToInt32(VbaPathProbe.Get(imported!, "Type")));
        int after = gate.Step("t15-count-after", () => Convert.ToInt32(VbaPathProbe.Get(components, "Count")));
        log.Write("t15.imported", gate, new { name, type, components_after = after });
        if (name != VbaPathProbe.ModuleName || type != VbaPathProbe.StandardModuleType || after != before + 1)
            throw new InvalidDataException("The imported component differs from the expected module; do not retry.");

        string macro = "'" + workbookPath.Replace("'", "''") + "'!" + VbaPathProbe.ModuleName + "." + VbaPathProbe.MacroName;
        MeasureS11("macro-call", () => application!.Run(macro));
        object marked = gate.Step("t15-read-marked", () => cell.Value2);
        log.Write("t15.macro", gate, new { result = marked, result_type = marked?.GetType().FullName });
        if (marked is not double value || value != VbaPathProbe.MarkedValue)
            throw new InvalidDataException("The macro result differs from the expected value; do not retry.");
        gate.Complete("ready");
    }
}
