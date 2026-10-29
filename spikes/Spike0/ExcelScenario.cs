using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Spike0;

// This is one lifecycle experiment, not the Core or a reusable session service.
internal sealed class ExcelScenario(string workbookPath, OwnershipRecordFile record,
    ExecutionGate gate, DiagnosticLog log)
{
    private static readonly TimeSpan SessionTimeout = TimeSpan.FromSeconds(120);
    private readonly List<object> references = [];
    private readonly HashSet<object> seenReferences = new(ReferenceEqualityComparer.Instance);
    private Excel.Application? application;
    private Excel.Workbooks? workbooks;
    private Excel.Workbook? workbook;
    private Process? process;
    private ProcessIdentity? identity;
    private bool generationStarted;
    private bool closeStarted;
    private bool quitReturned;
    private bool recordReleased;

    internal bool Run()
    {
        try
        {
            Start();
            EditKnownCell();
            End(save: true);
            log.Write("smoke.closed", gate, new { workbook_path = workbookPath, process = identity });
            return true;
        }
        catch (WorkStoppedException) when (gate.Snapshot.StopReason == "shutdown" && identity is not null)
        {
            // The active call has returned before this catch; no parallel COM cleanup is used.
            gate.Fault(generationStarted);
            try
            {
                gate.Begin("shutdown", SessionTimeout, forShutdown: true);
                CloseAndQuit();
                log.Write("shutdown.closed_without_save", gate);
                return false;
            }
            catch (Exception error)
            {
                Fail(error);
                throw;
            }
        }
        catch (Exception error)
        {
            // Start failure never triggers automatic discard/quit.
            Fail(error);
            throw;
        }
        finally { process?.Dispose(); }
    }

    private void Start()
    {
        gate.Begin("start", SessionTimeout);
        var existingPids = gate.Step("enumerate-existing-excel", () =>
        {
            var existing = Process.GetProcessesByName("EXCEL");
            try { return existing.Select(item => item.Id).ToHashSet(); }
            finally { foreach (var item in existing) item.Dispose(); }
        });
        RecordTransition.Run(gate, record, new(1, "starting", workbookPath), () => gate.Step("generate-excel", () =>
        {
            generationStarted = true;
            application = Track(new Excel.Application());
        }));
        int hwnd = gate.Step("excel-hwnd", () => application!.Hwnd);
        gate.Step("identify-excel", () =>
        {
            if (GetWindowThreadProcessId(new nint(hwnd), out uint pid) == 0 || pid == 0)
                throw new InvalidOperationException("Excel window has no identifiable process.");
            process = Process.GetProcessById(checked((int)pid));
            _ = process.SafeHandle; // Retain the handle to this process, not just a reusable PID.
            identity = ProcessIdentity.Read(process);
            if (existingPids.Contains(identity.Pid) ||
                !StringComparer.OrdinalIgnoreCase.Equals(identity.Name, "EXCEL"))
                throw new InvalidOperationException("New Excel ownership could not be established.");
        });
        RecordTransition.Run(gate, record, new(1, "owned", workbookPath, identity), OpenCopy);
    }

    private void OpenCopy()
    {
        log.Write("ownership.confirmed", gate, new { workbook_path = workbookPath, process = identity });
        gate.Step("excel-visible", () => application!.Visible =
            Environment.GetEnvironmentVariable("XLWRIGHT_EXCEL_VISIBLE") == "1");
        gate.Step("excel-caption", () => application!.Caption = "xlwright Spike0");
        gate.Step("automation-security", () => application!.AutomationSecurity =
            Office.MsoAutomationSecurity.msoAutomationSecurityLow);
        workbooks = gate.Step("get-workbooks", () => Track(application!.Workbooks));
        WithEventsSuppressed(() =>
        {
            workbook = gate.Step("open-copy", () => Track(workbooks.Open(workbookPath,
                UpdateLinks: 0, ReadOnly: false, IgnoreReadOnlyRecommended: true,
                Notify: false, AddToMru: false)));
            if (gate.Step("check-read-only", () => workbook.ReadOnly))
                throw new InvalidOperationException("Workbook was opened read-only.");
            string openedPath = gate.Step("check-open-path", () => workbook.FullName);
            if (!SamePath(openedPath, workbookPath))
                throw new InvalidOperationException("Opened workbook path does not match the copy.");
        });
        gate.Complete("ready");
        log.Write("start.ready", gate);
    }

    private void EditKnownCell()
    {
        gate.Begin("edit", TimeSpan.FromSeconds(30));
        WithEventsSuppressed(() =>
        {
            var sheets = gate.Step("get-sheets", () => Track(workbook!.Worksheets));
            var sheet = gate.Step("get-lifecycle-sheet", () => Track((Excel.Worksheet)sheets["Lifecycle"]));
            var cell = gate.Step("get-cell", () => Track(sheet.Range["A2"]));
            object before = gate.Step("read-initial-value", () => cell.Value2);
            if (before is not double initial || initial != 0)
                throw new InvalidDataException("The fixture cell does not contain the expected numeric zero.");
            gate.Step("write-known-value", () => cell.Value2 = 1d);
            object after = gate.Step("read-written-value", () => cell.Value2);
            if (after is not double written || written != 1)
                throw new InvalidDataException("Written numeric value did not read back.");
        });
        gate.Complete("ready");
        log.Write("cell.verified", gate);
    }

    private void End(bool save)
    {
        gate.Begin("end", SessionTimeout);
        if (save)
        {
            WithEventsSuppressed(() =>
            {
                try { gate.Step("save", () => workbook!.Save()); }
                catch (COMException error) when (error.HResult == unchecked((int)0x800A03EC))
                {
                    // An explicit Excel failure is distinct from transport failure/timeout.
                    throw new SaveFailedException(error);
                }
                if (!gate.Step("check-saved", () => workbook!.Saved))
                    throw new SaveFailedException();
            });
            log.Write("save.confirmed", gate);
        }
        CloseAndQuit();
    }

    private void CloseAndQuit()
    {
        if (!quitReturned)
        {
            EndSafety.CloseAndQuit(VerifyEndSafety, () => WithEventsSuppressed(() =>
            {
                bool originalAlerts = gate.Step("read-alerts", () => application!.DisplayAlerts);
                gate.Step("suppress-close-alerts", () => application!.DisplayAlerts = false);
                if (workbook is not null)
                    gate.Step("close-without-save", () =>
                    {
                        closeStarted = true;
                        workbook.Close(SaveChanges: false);
                        workbook = null;
                    });
                gate.Step("restore-close-alerts", () => application!.DisplayAlerts = originalAlerts);
            }), () => gate.Step("quit", () =>
            {
                application!.Quit();
                quitReturned = true;
            }));
        }
        gate.Step("release-owned-references", () =>
        {
            for (int index = references.Count - 1; index >= 0; index--)
                Marshal.FinalReleaseComObject(references[index]);
            references.Clear();
            seenReferences.Clear();
            workbook = null;
            workbooks = null;
            application = null;
        });
        bool exited = gate.Step("wait-process-exit", () => process!.WaitForExit(
            Math.Max(1, (int)Math.Ceiling(gate.Remaining.TotalMilliseconds))));
        if (!exited) throw new WorkStoppedException();
        if (!recordReleased)
        {
            RecordTransition.Run(gate, record, new(1, "completed"), () => gate.Step("release-record-lock", () =>
            {
                record.Dispose();
                recordReleased = true;
            }));
        }
        gate.Complete("closed");
    }

    private void VerifyEndSafety(bool requireEmpty = false)
    {
        gate.Step("verify-process-identity", () =>
        {
            using var current = Process.GetProcessById(identity!.Pid);
            EndSafety.RequireIdentity(identity.Matches(ProcessIdentity.Read(current)));
        });
        int hwnd = gate.Step("verify-owned-hwnd", () => application!.Hwnd);
        gate.Step("verify-window-pid", () =>
        {
            if (GetWindowThreadProcessId(new nint(hwnd), out uint pid) == 0 || pid != identity!.Pid)
                throw new InvalidOperationException("EXCEL_NEEDS_REVIEW: Excel window PID changed.");
        });
        // EOF may arrive after ownership is known but before normal startup obtains Workbooks.
        workbooks ??= gate.Step("get-workbooks-for-safety", () => Track(application!.Workbooks));
        var protectedViews = gate.Step("get-protected-views", () => Track(application!.ProtectedViewWindows));
        int protectedCount = gate.Step("count-protected-views", () => protectedViews.Count);
        int count = gate.Step("count-workbooks", () => workbooks!.Count);
        IEnumerable<string> ReadPaths()
        {
            for (int index = 1; index <= count; index++)
            {
                var item = gate.Step("get-open-workbook", () => Track(workbooks![index]));
                yield return gate.Step("get-open-workbook-path", () => item.FullName);
            }
        }
        EndSafety.RequireBooks(workbookPath, ReadPaths(), protectedCount, count, requireEmpty);
    }

    private void WithEventsSuppressed(Action action)
    {
        bool original = gate.Step("read-events", () => application!.EnableEvents);
        gate.Step("suppress-events", () => application!.EnableEvents = false);
        action();
        // No finally: after a failed/expired COM call, even restoration is a new operation.
        gate.Step("restore-events", () => application!.EnableEvents = original);
    }

    private T Track<T>(T reference) where T : class
    {
        if (seenReferences.Add(reference)) references.Add(reference);
        return reference;
    }

    private void Fail(Exception error)
    {
        EndSafety.Fault(gate, generationStarted, closeStarted, error);
        log.Write("operation.failed", gate, new
        {
            error = error.GetType().Name,
            error.Message,
            hresult = error.HResult,
            process = identity,
            workbook_path = workbookPath,
            excel_may_remain = generationStarted,
            manual_recovery = generationStarted
        });
    }

    private static bool SamePath(string left, string right) =>
        StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(left), Path.GetFullPath(right));

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
}

internal sealed class EndRefusedException : Exception
{
    internal EndRefusedException() : base("EXCEL_NEEDS_REVIEW: unmanaged workbook or Protected View detected.") { }
}

internal sealed class SaveFailedException(Exception? inner = null)
    : Exception("Excel reported a save failure. Keep the workbook and ownership record.", inner);
