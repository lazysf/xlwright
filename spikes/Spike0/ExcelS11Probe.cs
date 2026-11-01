namespace Spike0;

internal sealed partial class ExcelScenario
{
    private void MeasureS11(string phase, Action action)
    {
        if (s11 is null) action();
        else s11.Evidence.Measure(phase, () => gate.Step(phase + "-permitted", action));
    }

    private void ProbeS11()
    {
        gate.Begin("s11-observe", TimeSpan.FromSeconds(30));
        bool visible = gate.Step("s11-visible", () => application!.Visible);
        bool userControl = gate.Step("s11-user-control", () => application!.UserControl);
        s11!.Evidence.Write("excel-settings", "observed", new { visible, user_control = userControl, process = identity });
        if (s11.Scenario == "protected-view")
        {
            var views = gate.Step("s11-protected-views", () => Track(application!.ProtectedViewWindows));
            gate.Step("s11-open-protected-view", () => Track(views.Open(s11.OtherPath, AddToMru: false)));
            int count = gate.Step("s11-protected-count", () => views.Count);
            s11.Evidence.Write("protected-view", "observed", new { count });
            if (count <= 0) throw new InvalidOperationException("Protected View was not observed; do not retry.");
        }
        else if (s11.Scenario is "wait" or "modal")
        {
            string macro = "'" + workbookPath.Replace("'", "''") + "'!S11Module." +
                (s11.Scenario == "wait" ? "S11FiniteWait" : "S11Modal");
            gate.Step("s11-run-macro", () => MeasureS11("macro-call", () => application!.Run(macro)));
            if (s11.Scenario == "modal")
                throw new InvalidOperationException("Modal returned before the observation deadline; preserve evidence.");
        }
        else if (s11.Scenario == "shell2")
        {
            s11.ObserveExcelProcesses("shell-processes-before");
            // The shell action belongs to the person. The program does not launch an unknown Excel instance.
            s11.Evidence.Write("shell-input", "required", new { path = s11.OtherPath });
            gate.Complete("ready");
            // A person must act here; diagnostics are long JSON lines and hid this moment in the first attempt.
            Console.WriteLine($"{Environment.NewLine}>>> ACTION REQUIRED: open this file from the shell, then type go and press Enter.{Environment.NewLine}>>> {s11.OtherPath}{Environment.NewLine}");
            s11.WaitForShell();
            gate.Begin("s11-shell-observe", TimeSpan.FromSeconds(30));
            s11.ObserveExcelProcesses("shell-processes-after");
            int count = gate.Step("s11-shell-count", () => workbooks!.Count);
            var paths = new List<string>();
            for (int index = 1; index <= count; index++)
            {
                var book = gate.Step("s11-shell-book", () => Track(workbooks![index]));
                paths.Add(gate.Step("s11-shell-path", () => book.FullName));
            }
            s11.Evidence.Write("shell-owned-books", "observed", new { paths });
            if (paths.Any(path => SamePath(path, s11.OtherPath)))
            {
                gate.Complete("ready");
                End(save: true);
                throw new InvalidOperationException("Shell workbook unexpectedly allowed termination; do not retry.");
            }
            throw new InvalidOperationException("S11 shell observation requires human review; no automatic cleanup.");
        }
        gate.Complete("ready");
    }
}
