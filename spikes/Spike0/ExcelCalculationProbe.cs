using System.Diagnostics;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spike0;

internal sealed partial class ExcelScenario
{
    private void ProbeCalculation()
    {
        gate.Begin("s14-probe", TimeSpan.FromSeconds(30));
        var timer = Stopwatch.StartNew();
        var observations = new List<object>();
        var sheets = gate.Step("s14-sheets", () => Track(workbook!.Worksheets));
        var probe = gate.Step("s14-probe-sheet", () => Track((Excel.Worksheet)sheets["Probe"]));
        var resultSheet = gate.Step("s14-result-sheet", () => Track((Excel.Worksheet)sheets["Result"]));
        var input = gate.Step("s14-input", () => Track(probe.Range["B2"]));
        var local = gate.Step("s14-local", () => Track(probe.Range["C2"]));
        var dependent = gate.Step("s14-dependent", () => Track(resultSheet.Range["B2"]));
        log.Write("s14.environment", gate, new
        {
            version = gate.Step("s14-version", () => application!.Version),
            build = gate.Step("s14-build", () => application!.Build)
        });
        (object? Local, object? Dependent, int Mode, int State) Observe(string phase)
        {
            double started = timer.Elapsed.TotalMilliseconds;
            int before = gate.Step("s14-state-before", () => (int)application!.CalculationState);
            object? source = gate.Step("s14-read-input", () => input.Value2);
            object? first = gate.Step("s14-read-local", () => local.Value2);
            object? second = gate.Step("s14-read-dependent", () => dependent.Value2);
            int after = gate.Step("s14-state-after", () => (int)application!.CalculationState);
            int mode = gate.Step("s14-mode", () => (int)application!.Calculation);
            var observation = new
            {
                phase, started_ms = started, finished_ms = timer.Elapsed.TotalMilliseconds,
                state_before = before, state_after = after, mode, input = source,
                local = first, local_type = first?.GetType().FullName,
                dependent = second, dependent_type = second?.GetType().FullName,
                local_formula = gate.Step("s14-local-formula", () => local.Formula),
                dependent_formula = gate.Step("s14-dependent-formula", () => dependent.Formula),
                calculate_before_save = gate.Step("s14-before-save", () => application!.CalculateBeforeSave),
                events = gate.Step("s14-events", () => application!.EnableEvents),
                alerts = gate.Step("s14-alerts", () => application!.DisplayAlerts)
            };
            observations.Add(observation);
            log.Write("s14.observed", gate, observation);
            CellProbe.Require(source is double && (string)observation.local_formula == "=B2*2" &&
                (string)observation.dependent_formula == "=Probe!C2+1", "S14 inputs/formulas");
            return (first, second, mode, after);
        }
        void Pause()
        {
            gate.Step("s14-pump", System.Windows.Forms.Application.DoEvents);
            gate.Step("s14-pause", () => Thread.Sleep(50));
        }
        void RequireValues((object? Local, object? Dependent, int Mode, int State) observed, double value, int mode)
        {
            CellProbe.Require(Equals(observed.Local, value * 2) && Equals(observed.Dependent, value * 2 + 1) && observed.Mode == mode,
                "S14 values/mode");
        }
        void RunMacro(string name) => gate.Step("s14-macro-" + name, () =>
        {
            application!.Run("'probe.xlsm'!S14Module." + name);
        });
        void Change(string phase, double value, int mode, double oldValue, Action change)
        {
            Observe(phase + "-before");
            gate.Step("s14-change", change);
            var immediate = Observe(phase + "-immediate");
            CellProbe.Require(immediate.Mode == mode && Equals(gate.Step("s14-confirm-input", () => input.Value2), value), "S14 changed input/mode");
            long passiveStart = Stopwatch.GetTimestamp();
            while (Stopwatch.GetElapsedTime(passiveStart).TotalMilliseconds < 250) Pause();
            var passive = Observe(phase + "-passive");
            if (mode == (int)Excel.XlCalculation.xlCalculationManual)
            {
                RequireValues(immediate, oldValue, mode);
                RequireValues(passive, oldValue, mode);
                gate.Step("s14-calculate", () => application!.Calculate());
                passive = Observe(phase + "-calculate-returned");
            }
            while (passive.State != (int)Excel.XlCalculationState.xlDone)
            {
                Pause();
                passive = Observe(phase + "-waiting");
            }
            var completed = Observe(phase + "-done");
            RequireValues(completed, value, mode);
            CellProbe.Require(completed.State == (int)Excel.XlCalculationState.xlDone, "S14 final calculation state");
        }
        int automatic = (int)Excel.XlCalculation.xlCalculationAutomatic;
        int manual = (int)Excel.XlCalculation.xlCalculationManual;
        RequireValues(Observe("opened"), 1, calculationCase == "manual" ? manual : automatic);
        if (calculationCase == "mode-change")
        {
            Change("manual31", 31, manual, 1, () => RunMacro("SetManual31"));
            Change("automatic41", 41, automatic, 31, () => RunMacro("SetAutomatic41"));
        }
        else
        {
            int mode = calculationCase == "manual" ? manual : automatic;
            Change("com11", 11, mode, 1, () => WithEventsSuppressed(() =>
                gate.Step("s14-write-input", () => input.Value2 = 11d)));
            Change("macro21", 21, mode, 11, () => RunMacro("SetInput21"));
        }
        gate.Step("s14-write-observations", () => File.WriteAllText(
            Path.Combine(Path.GetDirectoryName(workbookPath)!, "calculation-observations.json"), JsonSerializer.Serialize(observations)));
        gate.Complete("ready");
    }
}
