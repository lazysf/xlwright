using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Spike0;

// This is one lifecycle experiment, not the Core or a reusable session service.
internal sealed partial class ExcelScenario(string workbookPath, OwnershipRecordFile record,
    ExecutionGate gate, DiagnosticLog log, string? unmanagedPath = null, string? recordFault = null,
    string? startupCase = null, string? cellCase = null, string? calculationCase = null, S9Probe? s9 = null, S11Probe? s11 = null)
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
    private bool exitConfirmed;
    private HashSet<int>? existingExcelPids;
    private bool saveConfirmed;
    private bool refusalExpected;
    private int observedProtectedCount;
    private int observedBookCount;
    private readonly List<string> observedPaths = [];

    internal bool Run()
    {
        try
        {
            Start();
            if (s11 is not null) { EditKnownCell(); ProbeS11(); }
            else if (s9 is not null) ProbeS9();
            else if (calculationCase is not null) ProbeCalculation();
            else if (cellCase is not null) ProbeCells();
            else if (startupCase is null) EditKnownCell();
            else ProbeStartup();
            if (unmanagedPath is not null) OpenUnmanagedCopy();
            End(save: true);
            if (s11?.Scenario == "protected-view")
                throw new InvalidOperationException("Protected View unexpectedly allowed termination; do not retry.");
            if (unmanagedPath is not null)
                throw new InvalidOperationException("L3 unexpectedly allowed termination; do not retry.");
            if (recordFault is not null)
                throw new InvalidOperationException("L2 record fault was not injected; do not retry.");
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
            if (unmanagedPath is not null && error is EndRefusedException)
            {
                try { VerifyRefusalEvidence(); }
                catch (Exception evidenceError)
                {
                    Fail(evidenceError);
                    throw;
                }
            }
            if (recordFault is not null && error is RecordIoException fault)
            {
                try { VerifyRecordFault(fault); }
                catch (Exception evidenceError)
                {
                    Fail(evidenceError);
                    throw;
                }
            }
            throw;
        }
        finally { process?.Dispose(); }
    }

    private void OpenUnmanagedCopy()
    {
        gate.Begin("introduce-unmanaged-copy", TimeSpan.FromSeconds(120));
        WithEventsSuppressed(() =>
        {
            var other = gate.Step("open-unmanaged-copy", () => Track(workbooks!.Open(unmanagedPath!,
                UpdateLinks: 0, ReadOnly: false, IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false)));
            if (gate.Step("unmanaged-read-only", () => other.ReadOnly) ||
                !SamePath(gate.Step("unmanaged-full-name", () => other.FullName), unmanagedPath!))
                throw new InvalidOperationException("The unmanaged copy could not be verified.");
        });
        gate.Complete("ready");
    }

    private void VerifyRefusalEvidence()
    {
        // Only retained data and OS/file reads are allowed after the refusal. No COM calls.
        if (!refusalExpected || !saveConfirmed || closeStarted || quitReturned || recordReleased ||
            identity is null || process is null || process.HasExited ||
            gate.Snapshot is not { State: "ready", Busy: false })
            throw new InvalidOperationException("L3 refusal evidence is incomplete; preserve all records.");
        using var current = Process.GetProcessById(identity.Pid);
        if (!identity.Matches(ProcessIdentity.Read(current)))
            throw new InvalidOperationException("L3 process identity could not be reverified.");
        var owned = new OwnershipRecord(1, "owned", workbookPath, identity);
        record.VerifyRetained(owned);
        log.Write("end-refusal.verified", gate, new
        {
            process = identity, record = owned, observed_paths = observedPaths,
            protected_view_count = observedProtectedCount, workbook_count = observedBookCount,
            identity_matched = true, excel_alive = true, sharing_violation_observed = true,
            close_started = closeStarted, quit_returned = quitReturned, record_released = recordReleased,
            manual_recovery = true
        });
    }

    private void VerifyRecordFault(RecordIoException fault)
    {
        // Only retained data and OS/file reads are allowed after the injected failure. No COM calls.
        var snapshot = gate.Snapshot;
        if (fault.Boundary != recordFault || fault.Operation != "write" || snapshot.Busy || recordReleased)
            throw new InvalidOperationException("L2 fault evidence is incomplete; preserve all records.");
        OwnershipRecord retained;
        int[]? newExcelPids = null;
        if (recordFault == "starting")
        {
            if (generationStarted || identity is not null || snapshot.State != "closed" || existingExcelPids is null)
                throw new InvalidOperationException("L2 starting fault evidence is incomplete; preserve all records.");
            var current = Process.GetProcessesByName("EXCEL");
            try { newExcelPids = current.Select(item => item.Id).Where(pid => !existingExcelPids.Contains(pid)).ToArray(); }
            finally { foreach (var item in current) item.Dispose(); }
            // A new PID may belong to the user; stop for review without choosing it for termination.
            if (newExcelPids.Length != 0)
                throw new InvalidOperationException("A new Excel process appeared during the L2 starting fault: PID " +
                    string.Join(", ", newExcelPids));
            RuntimeAssets.VerifyCell(workbookPath, 0);
            retained = new(1, "completed");
        }
        else if (recordFault == "owned")
        {
            if (!generationStarted || identity is null || process is not { HasExited: false } || exitConfirmed ||
                closeStarted || workbooks is not null || workbook is not null || snapshot.State != "isolated")
                throw new InvalidOperationException("L2 owned fault evidence is incomplete; preserve all records.");
            using var current = Process.GetProcessById(identity.Pid);
            if (!identity.Matches(ProcessIdentity.Read(current)))
                throw new InvalidOperationException("L2 owned process identity could not be reverified.");
            RuntimeAssets.VerifyCell(workbookPath, 0);
            retained = new(1, "starting", workbookPath);
        }
        else
        {
            if (!saveConfirmed || !closeStarted || !quitReturned || !exitConfirmed || identity is null ||
                process is not { HasExited: true } || snapshot.State != "isolated")
                throw new InvalidOperationException("L2 completed fault evidence is incomplete; preserve all records.");
            RuntimeAssets.VerifyCell(workbookPath, 1);
            retained = new(1, "owned", workbookPath, identity);
        }
        // Before this process exits its own handle still holds the lock, including for starting (released on exit).
        record.VerifyRetained(retained);
        log.Write("record-fault.verified", gate, new
        {
            boundary = fault.Boundary, operation = fault.Operation, process = identity, record = retained,
            generation_started = generationStarted, excel_alive = process is { HasExited: false },
            excel_exit_confirmed = exitConfirmed, new_excel_pids = newExcelPids,
            close_started = closeStarted, quit_returned = quitReturned, record_released = recordReleased,
            sharing_violation_observed = true, manual_recovery = generationStarted
        });
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
        existingExcelPids = existingPids;
        RecordTransition.Run(gate, record, new(1, "starting", workbookPath), () => gate.Step("generate-excel", () =>
        {
            if (recordFault == "starting")
                throw new InvalidOperationException("L2 starting fault was not injected; Excel was not created.");
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
        s11?.Identify(identity!, new nint(hwnd));
        if (recordFault == "owned")
        {
            // Approved exception for this fixed case only: show Excel before the owned record so a person can
            // identify and close it normally when it remains without identity in the record.
            gate.Step("l2-owned-visible", () => application!.Visible = true);
            gate.Step("l2-owned-caption", () => application!.Caption =
                "xlwright L2 owned " + Path.GetFileName(Path.GetDirectoryName(workbookPath)));
        }
        RecordTransition.Run(gate, record, new(1, "owned", workbookPath, identity), () =>
        {
            if (recordFault == "owned")
                gate.Step("l2-owned-not-injected", () => throw new InvalidOperationException(
                    "L2 owned fault was not injected; the workbook was not opened."));
            OpenCopy();
        });
    }

    private void OpenCopy()
    {
        log.Write("ownership.confirmed", gate, new { workbook_path = workbookPath, process = identity });
        gate.Step("excel-visible", () => application!.Visible = unmanagedPath is not null ||
            Environment.GetEnvironmentVariable("XLWRIGHT_EXCEL_VISIBLE") == "1");
        gate.Step("excel-caption", () => application!.Caption = unmanagedPath is null ? "xlwright Spike0" :
            "xlwright L3 " + Path.GetFileName(Path.GetDirectoryName(workbookPath)));
        if (startupCase is not null) ObserveSettings("generated");
        var originalSecurity = startupCase is null && cellCase is null && calculationCase is null ? Office.MsoAutomationSecurity.msoAutomationSecurityLow :
            gate.Step("read-automation-security", () => application!.AutomationSecurity);
        gate.Step("automation-security", () => application!.AutomationSecurity =
            startupCase == "C" ? Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable :
            Office.MsoAutomationSecurity.msoAutomationSecurityLow);
        workbooks = gate.Step("get-workbooks", () => Track(application!.Workbooks));
        WithStartupEvents(() =>
        {
            if (startupCase is not null) ObserveSettings("before-open");
            workbook = gate.Step("open-copy", () => Track(workbooks.Open(workbookPath,
                UpdateLinks: 0, ReadOnly: false, IgnoreReadOnlyRecommended: true,
                Notify: false, AddToMru: false)));
            if (gate.Step("check-read-only", () => workbook.ReadOnly))
                throw new InvalidOperationException("Workbook was opened read-only.");
            string openedPath = gate.Step("check-open-path", () => workbook.FullName);
            if (!SamePath(openedPath, workbookPath))
                throw new InvalidOperationException("Opened workbook path does not match the copy.");
            if (startupCase is not null)
            {
                ObserveSettings("after-open");
                gate.Step("restore-automation-security", () => application!.AutomationSecurity = originalSecurity);
                ObserveSettings("after-security-restoration");
                ObserveTrace("after-open");
                StartupProbe.RequireTrace(startupCase, "open", ReadTrace());
            }
        });
        if (cellCase is not null || calculationCase is not null)
        {
            gate.Step("restore-cell-security", () => application!.AutomationSecurity = originalSecurity);
            ObserveSettings("cell-open");
        }
        if (startupCase is not null) ObserveSettings("after-open-events-restoration");
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
            saveConfirmed = true;
        }
        CloseAndQuit();
    }

    private void CloseAndQuit()
    {
        if (!quitReturned)
        {
            EndSafety.CloseAndQuit(VerifyEndSafety, () => WithStartupEvents(() =>
            {
                if (startupCase is not null) ObserveSettings("before-close");
                bool originalAlerts = gate.Step("read-alerts", () => application!.DisplayAlerts);
                gate.Step("suppress-close-alerts", () => application!.DisplayAlerts = false);
                if (workbook is not null)
                    gate.Step("close-without-save", () =>
                    {
                        closeStarted = true;
                        log.Write("close.started", gate);
                        MeasureS11("close-call", () => workbook.Close(SaveChanges: false));
                        workbook = null;
                    });
                gate.Step("restore-close-alerts", () => application!.DisplayAlerts = originalAlerts);
                if (startupCase is not null)
                {
                    ObserveSettings("after-close");
                    ObserveTrace("after-close");
                    StartupProbe.RequireTrace(startupCase, "closed", ReadTrace());
                }
            }), () => gate.Step("quit", () =>
            {
                log.Write("quit.started", gate);
                MeasureS11("quit-call", () => application!.Quit());
                quitReturned = true;
            }));
        }
        gate.Step("release-owned-references", () =>
        {
            for (int index = references.Count - 1; index >= 0; index--)
            {
                int current = index;
                MeasureS11("release-reference", () => Marshal.FinalReleaseComObject(references[current]));
            }
            references.Clear();
            seenReferences.Clear();
            workbook = null;
            workbooks = null;
            application = null;
        });
        bool exited = false;
        gate.Step("wait-process-exit", () => MeasureS11("wait-process-exit", () => exited = process!.WaitForExit(
            Math.Max(1, (int)Math.Ceiling(gate.Remaining.TotalMilliseconds)))));
        if (!exited) throw new WorkStoppedException();
        exitConfirmed = true;
        if (!recordReleased)
        {
            log.Write("record-completed.started", gate);
            MeasureS11("record-completed", () => RecordTransition.Run(gate, record, new(1, "completed"), () => gate.Step("release-record-lock", () =>
            {
                record.Dispose();
                recordReleased = true;
            })));
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
        observedPaths.Clear();
        observedProtectedCount = protectedCount;
        observedBookCount = count;
        refusalExpected = false;
        IEnumerable<string> ReadPaths()
        {
            for (int index = 1; index <= count; index++)
            {
                var item = gate.Step("get-open-workbook", () => Track(workbooks![index]));
                string path = gate.Step("get-open-workbook-path", () => item.FullName);
                observedPaths.Add(path);
                yield return path;
            }
        }
        IEnumerable<string> paths = ReadPaths();
        // The fixed experiment records both known copies before asking the shared policy to decide.
        if (unmanagedPath is not null) paths = paths.ToArray();
        try { EndSafety.RequireBooks(workbookPath, paths, protectedCount, count, requireEmpty); }
        catch (EndRefusedException)
        {
            refusalExpected = !requireEmpty && unmanagedPath is not null && protectedCount == 0 && count == 2 &&
                observedPaths.Any(path => SamePath(path, unmanagedPath)) &&
                observedPaths.Any(path => SamePath(path, workbookPath));
            throw;
        }
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
            excel_may_remain = generationStarted && !exitConfirmed,
            excel_exit_confirmed = exitConfirmed,
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
