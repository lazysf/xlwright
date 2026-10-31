using Excel = Microsoft.Office.Interop.Excel;

namespace Spike0;

internal sealed partial class ExcelScenario
{
    private void ProbeS9()
    {
        gate.Begin("s9-prepare", TimeSpan.FromSeconds(30));
        gate.Step("s9-visible", () => application!.Visible = true);
        gate.Step("s9-caption", () => application!.Caption = "xlwright S9 " + Path.GetFileName(Path.GetDirectoryName(workbookPath)));
        var sheets = gate.Step("s9-sheets", () => Track(workbook!.Worksheets));
        var sheet = gate.Step("s9-sheet", () => Track((Excel.Worksheet)sheets["Lifecycle"]));
        var target = gate.Step("s9-target", () => Track(sheet.Range["A2"]));
        object initial = gate.Step("s9-initial", () => target.Value2);
        if (initial is not double value || value != 0) throw new InvalidDataException("S9 initial value differs.");
        bool events = gate.Step("s9-read-events", () => application!.EnableEvents);
        gate.Step("s9-suppress-events", () => application!.EnableEvents = false);
        gate.Complete("ready");
        s9!.WaitForStart();
        gate.Begin("s9-write", s9.Timeout);
        var filter = s9.Install();
        try
        {
            filter.Register();
            gate.Step("s9-single-setter", () =>
            {
                filter.BeginCall();
                int hresult = 0;
                try { target.Value2 = 1d; }
                catch (Exception error) { hresult = error.HResult; throw; }
                finally { filter.EndCall(hresult); }
            });
            gate.Step("s9-accept-return", () =>
            {
                if (!filter.ObservedRetry) throw new InvalidDataException("No actual RETRYLATER retry was observed.");
            });
            // B/C must never continue even if Excel unexpectedly accepts the setter.
            if (!s9.IsRelease)
                throw new InvalidOperationException("S9 unexpected acceptance.");
        }
        catch (System.Runtime.InteropServices.COMException) when (s9.IsDeadline && filter.DeadlineRefusal)
        {
            // The last sub-100ms interval permits no retry. Observe the real deadline without COM.
            while (!gate.CheckTimeout()) Thread.Sleep(1);
            throw;
        }
        finally
        {
            try { filter.Dispose(); }
            finally { s9.SaveEvidence("returned"); }
        }
        object after = gate.Step("s9-readback", () => target.Value2);
        if (after is not double written || written != 1) throw new InvalidDataException("S9 readback differs.");
        gate.Step("s9-restore-events", () => application!.EnableEvents = events);
        gate.Complete("ready");
        log.Write("s9.verified", gate);
    }
}
