using Excel = Microsoft.Office.Interop.Excel;

namespace Spike0;

internal sealed partial class ExcelScenario
{
    private string[] ReadTrace() => gate.Step("read-s10-trace", () =>
        File.ReadAllLines(Path.Combine(Path.GetDirectoryName(workbookPath)!, "s10-vba.log")));

    private void ObserveTrace(string phase) => log.Write("s10.trace", gate,
        new { phase, scenario = startupCase, trace = ReadTrace() });

    private void ObserveSettings(string phase)
    {
        var security = gate.Step("observe-security", () => application!.AutomationSecurity);
        bool events = gate.Step("observe-events", () => application!.EnableEvents);
        bool alerts = gate.Step("observe-alerts", () => application!.DisplayAlerts);
        int? calculation = StartupProbe.ReadCalculation(workbook is not null,
            () => (int)gate.Step("observe-calculation", () => application!.Calculation));
        log.Write("s10.settings", gate, new { phase, security, events, alerts, calculation,
            calculation_observed = calculation.HasValue,
            calculation_unobserved_reason = calculation.HasValue ? null : "managed_workbook_not_open" });
    }

    private void WithStartupEvents(Action action)
    {
        if (startupCase is null or "B") { WithEventsSuppressed(action); return; }
        bool original = gate.Step("read-events", () => application!.EnableEvents);
        gate.Step("enable-control-events", () => application!.EnableEvents = true);
        action();
        gate.Step("restore-events", () => application!.EnableEvents = original);
    }

    private void ProbeStartup()
    {
        gate.Begin("s10-probe", TimeSpan.FromSeconds(30));
        WithEventsSuppressed(() =>
        {
            var sheets = gate.Step("s10-sheets", () => Track(workbook!.Worksheets));
            var sheet = gate.Step("s10-sheet", () => Track((Excel.Worksheet)sheets["Probe"]));
            var input = gate.Step("s10-input", () => Track(sheet.Range["B2"]));
            var output = gate.Step("s10-output", () => Track(sheet.Range["C2"]));
            object initial = gate.Step("s10-initial-udf", () => output.Value2);
            double value = System.Security.Cryptography.RandomNumberGenerator.GetInt32(10000, 1000000);
            gate.Step("s10-write-input", () => input.Value2 = value);
            gate.Step("s10-write-formula", () => output.Formula = "=S10Probe(B2)");
            gate.Step("s10-calculate", () => output.Calculate());
            object result = gate.Step("s10-read-udf", () => output.Value2);
            log.Write("s10.udf", gate, new { initial, input = value, formula = "=S10Probe(B2)", result,
                result_type = result?.GetType().FullName });
            bool computed = result is double number && number == value + 1;
            if (startupCase == "C" ? computed || result is not int : !computed)
                throw new InvalidDataException("S10 UDF result differs from the expected case.");
            if (startupCase == "B")
            {
                gate.Step("s10-auto-open", () => workbook!.RunAutoMacros(Excel.XlRunAutoMacro.xlAutoOpen));
                ObserveTrace("explicit-auto-open");
            }
            StartupProbe.RequireTrace(startupCase!, "probe", ReadTrace());
        });
        ObserveSettings("after-probe-restoration");
        gate.Complete("ready");
    }

}
