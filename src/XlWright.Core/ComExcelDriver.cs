using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace XlWright.Core;

// Moved from Spike0's ExcelScenario and ExcelCellProbe. Every COM call goes through the gate, so a revoked
// permission stops the next call, and no restoration runs after a failed call: that would be a new operation.
internal sealed class ComExcelDriver(ExecutionGate gate) : IExcelDriver
{
    private const int ExcelApplicationFacility = unchecked((int)0x800A0000);
    private const int CallRejected = unchecked((int)0x80010001);
    private static readonly int[] Disconnected =
    [
        unchecked((int)0x80010007), // RPC_E_SERVER_DIED
        unchecked((int)0x80010012), // RPC_E_SERVER_DIED_DNE
        unchecked((int)0x80010108), // RPC_E_DISCONNECTED
        unchecked((int)0x800706BA), // RPC_S_SERVER_UNAVAILABLE
        unchecked((int)0x800706BE)  // RPC_S_CALL_FAILED
    ];

    // tradeoff: references live until the session ends, so a long session holds more RCWs.
    // Release per operation if the number of requests in a session grows large.
    private readonly List<object> references = [];
    private readonly HashSet<object> seenReferences = new(ReferenceEqualityComparer.Instance);
    private HashSet<int> existingPids = [];
    private Excel.Application? application;
    private Excel.Workbooks? workbooks;
    private Excel.Workbook? workbook;
    private Process? process;
    private ProcessIdentity? identity;

    public void Generate()
    {
        existingPids = Call("excel_start", () =>
        {
            var existing = Process.GetProcessesByName("EXCEL");
            try { return existing.Select(item => item.Id).ToHashSet(); }
            finally { foreach (var item in existing) item.Dispose(); }
        });
        application = Call("excel_start", () => Track(new Excel.Application()));
    }

    public ProcessIdentity Identify()
    {
        int hwnd = Call("excel_start", () => application!.Hwnd);
        return Call("excel_start", () =>
        {
            if (GetWindowThreadProcessId(new nint(hwnd), out uint pid) == 0 || pid == 0)
                throw new InvalidOperationException("Excel window has no identifiable process.");
            process = Process.GetProcessById(checked((int)pid));
            _ = process.SafeHandle; // Retain the handle to this process, not just a reusable PID.
            var read = ProcessIdentity.Read(process);
            if (existingPids.Contains(read.Pid) || !StringComparer.OrdinalIgnoreCase.Equals(read.Name, "EXCEL"))
                throw new InvalidOperationException("New Excel ownership could not be established.");
            identity = read;
            return read;
        });
    }

    public void Open(string workbookPath)
    {
        Call("open", () => application!.Visible = Environment.GetEnvironmentVariable("XLWRIGHT_EXCEL_VISIBLE") == "1");
        Call("open", () => application!.Caption = "xlwright");
        var originalSecurity = Call("open", () => application!.AutomationSecurity);
        Call("open", () => application!.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityLow);
        workbooks = Call("open", () => Track(application!.Workbooks));
        WithEventsSuppressed("open", () =>
        {
            workbook = Call("open", () => Track(workbooks.Open(workbookPath, UpdateLinks: 0, ReadOnly: false,
                IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false)));
            if (Call("open", () => workbook.ReadOnly))
                throw new InvalidOperationException("Workbook was opened read-only.");
            if (!SamePath(Call("open", () => workbook.FullName), workbookPath))
                throw new InvalidOperationException("Opened workbook path does not match.");
        });
        Call("open", () => application!.AutomationSecurity = originalSecurity);
    }

    public CellContent Read(string sheet, string address, bool values, bool formulas)
    {
        CellContent content = new(null, null);
        WithEventsSuppressed("read", () => Rejectable(() =>
        {
            var cell = FindCell(sheet, address, "read");
            if (values) Calculate();
            object? value = values ? Call("read", () => (object?)cell.Value2) : null;
            string? formula = null;
            if (formulas && Call("read", () => (bool)cell.HasFormula))
                formula = Call("read", () => (string)((dynamic)cell).Formula2);
            content = new(value, formula);
        }));
        return content;
    }

    public void Write(string sheet, string address, CellWrite write) =>
        WithEventsSuppressed("write", () => Rejectable(() =>
        {
            var cell = FindCell(sheet, address, "write");
            RequireWritable(sheet, cell);
            // A 1x1 array clears the value and formula and keeps the format and merge (S2).
            if (write.Value is null) Call("write", () => cell.Value2 = new object?[1, 1]);
            else if (write.IsFormula) Call("write", () => { ((dynamic)cell).Formula2 = (string)write.Value; });
            else Call("write", () => cell.Value2 = CellValues.ForValue2(write.Value));
            object? value = Call("verify", () => (object?)cell.Value2);
            bool hasFormula = Call("verify", () => (bool)cell.HasFormula);
            if (!CellValues.Verified(write, value, hasFormula))
                throw new CoreException(ErrorCode.ExcelError, ErrorDetail.WriteNotVerified,
                    "Reading the cell back did not match the write. Check the cell before retrying.");
        }));

    public void Save() => WithEventsSuppressed("save", () =>
    {
        try { Call("save", () => workbook!.Save()); }
        catch (COMException error) when (IsExcelRejection(error)) { throw new SaveFailedException(error); }
        if (!Call("save", () => workbook!.Saved)) throw new SaveFailedException();
    });

    public void VerifyEndSafety(string workbookPath, bool requireEmpty)
    {
        Call("safety_check", () =>
        {
            using var current = Process.GetProcessById(identity!.Pid);
            EndSafety.RequireIdentity(identity.Matches(ProcessIdentity.Read(current)));
        });
        int hwnd = Call("safety_check", () => application!.Hwnd);
        Call("safety_check", () => EndSafety.RequireIdentity(
            GetWindowThreadProcessId(new nint(hwnd), out uint pid) != 0 && pid == identity!.Pid));
        // EOF may arrive after ownership is known but before startup obtained Workbooks.
        workbooks ??= Call("safety_check", () => Track(application!.Workbooks));
        var protectedViews = Call("safety_check", () => Track(application!.ProtectedViewWindows));
        int protectedCount = Call("safety_check", () => protectedViews.Count);
        int count = Call("safety_check", () => workbooks.Count);
        var paths = new List<string>();
        for (int index = 1; index <= count; index++)
        {
            int current = index;
            var item = Call("safety_check", () => Track(workbooks[current]));
            paths.Add(Call("safety_check", () => item.FullName));
        }
        EndSafety.RequireBooks(workbookPath, paths, protectedCount, count, requireEmpty);
    }

    public void Close() => WithEventsSuppressed("close", () =>
    {
        bool originalAlerts = Call("close", () => application!.DisplayAlerts);
        Call("close", () => application!.DisplayAlerts = false);
        if (workbook is not null)
        {
            Call("close", () => workbook.Close(SaveChanges: false));
            workbook = null;
        }
        Call("close", () => application!.DisplayAlerts = originalAlerts);
    });

    public void Quit() => Call("close", () => application!.Quit());

    public bool ReleaseAndWaitForExit(TimeSpan timeout)
    {
        Call("process_exit", () =>
        {
            for (int index = references.Count - 1; index >= 0; index--) Marshal.FinalReleaseComObject(references[index]);
            references.Clear();
            seenReferences.Clear();
            workbook = null;
            workbooks = null;
            application = null;
        });
        using var owned = process!;
        return Call("process_exit", () => owned.WaitForExit(Math.Max(1, (int)Math.Ceiling(timeout.TotalMilliseconds))));
    }

    private Excel.Range FindCell(string sheet, string address, string stage)
    {
        var sheets = Call(stage, () => Track(workbook!.Worksheets));
        int count = Call(stage, () => sheets.Count);
        for (int index = 1; index <= count; index++)
        {
            int current = index;
            var item = Call(stage, () => Track((Excel.Worksheet)sheets[current]));
            // Indexing by name would ignore case; the contract matches sheet names exactly.
            if (Call(stage, () => item.Name) == sheet) return Call(stage, () => Track(item.Range[address]));
        }
        throw new CoreException(ErrorCode.NotFound, ErrorDetail.SheetNotFound, "The sheet does not exist. Check the name and retry.");
    }

    private void RequireWritable(string sheet, Excel.Range cell)
    {
        if (Call("write", () => (bool)cell.MergeCells))
        {
            var merge = Call("write", () => Track(cell.MergeArea));
            if (Call("write", () => merge.Row != cell.Row || merge.Column != cell.Column))
                throw new CoreException(ErrorCode.ExcelError, ErrorDetail.MergedCell,
                    "Only the top-left cell of a merged range can be written. Use that cell.");
        }
        var owner = Call("write", () => Track(cell.Worksheet));
        if (Call("write", () => owner.ProtectContents) && Call("write", () => (bool)cell.Locked))
            throw new CoreException(ErrorCode.ExcelError, ErrorDetail.CellProtected,
                "The cell is protected. xlwright does not remove protection.");
    }

    private void Calculate()
    {
        Call("calculate", () => application!.Calculate());
        // The gate ends this loop at the deadline; a stale value is never returned.
        while (Call("calculate", () => application!.CalculationState) != Excel.XlCalculationState.xlDone)
        {
            Call("calculate", System.Windows.Forms.Application.DoEvents);
            Call("calculate", () => Thread.Sleep(50));
        }
    }

    private void WithEventsSuppressed(string stage, Action action)
    {
        bool original = Call(stage, () => application!.EnableEvents);
        Call(stage, () => application!.EnableEvents = false);
        try { action(); }
        catch (CoreException error) when (error.Code is ErrorCode.ExcelError or ErrorCode.NotFound)
        {
            // A clear refusal keeps the session ready, so the settings must not stay changed for the next operation.
            Call(stage, () => application!.EnableEvents = original);
            throw;
        }
        // No finally: after an unknown or expired COM call, even restoration is a new operation.
        Call(stage, () => application!.EnableEvents = original);
    }

    private static void Rejectable(Action action)
    {
        try { action(); }
        catch (COMException error) when (IsExcelRejection(error))
        {
            throw new CoreException(ErrorCode.ExcelError, ErrorDetail.ExcelRejected,
                "Excel rejected the call. Check the arguments and the workbook, then retry.", error);
        }
    }

    // tradeoff: SERVERCALL_RETRYLATER is reported as a rejection without the bounded retry in docs/architecture.md.
    // Add the message filter from Spike0's S9 when T20 evaluates this path.
    private static bool IsExcelRejection(COMException error) =>
        (error.HResult & ~0xFFFF) == ExcelApplicationFacility || error.HResult == CallRejected;

    private T Call<T>(string stage, Func<T> action) => gate.Step(stage, () =>
    {
        try { return action(); }
        catch (COMException error) when (Disconnected.Contains(error.HResult))
        {
            throw new CoreException(ErrorCode.StateUnknown, ErrorDetail.ExcelDisconnected,
                "The connection to Excel was lost. Do not resend; follow manual recovery.", error);
        }
    });

    private void Call(string stage, Action action) => Call(stage, () =>
    {
        action();
        return true;
    });

    private T Track<T>(T reference) where T : class
    {
        if (seenReferences.Add(reference)) references.Add(reference);
        return reference;
    }

    private static bool SamePath(string left, string right) =>
        StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(left), Path.GetFullPath(right));

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
}
