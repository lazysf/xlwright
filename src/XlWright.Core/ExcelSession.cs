using System.Text.Json;
using System.Text.Json.Serialization;

namespace XlWright.Core;

public sealed record SessionSummary(string State, bool Busy);

public sealed record ProcessInfo(int Pid, DateTime StartedAtUtc, string Name);

public sealed record ErrorInfo(string Code, string Detail, string Message, string Operation, string? Stage,
    bool? ExcelMayRemain, ProcessInfo? Process, string? Log);

public abstract record OperationResult
{
    public SessionSummary Session { get; internal set; } = new("closed", false);
}

public sealed record WorkbookResult(string Workbook) : OperationResult;

public sealed record RangeResult(string Sheet, string Address,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] object?[][]? Values = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string?[][]? Formulas = null) : OperationResult;

public sealed class SessionOperationException(ErrorInfo error, SessionSummary session) : Exception(error.Message)
{
    public ErrorInfo Error { get; } = error;
    public SessionSummary Session { get; } = session;
}

// The only entry the MCP layer uses. It decides acceptance, order, deadlines and the one response of each operation;
// all work runs on the STA, and COM stays behind IExcelDriver. Order and failure rules are in docs/architecture.md.
public sealed class ExcelSession : IDisposable
{
    private static readonly string[] Stages = ["config", "lock", "record", "process_query", "excel_start", "open",
        "calculate", "read", "write", "verify", "save", "safety_check", "close", "process_exit"];

    private readonly Func<ExecutionGate, IExcelDriver> createDriver;
    private readonly bool excelAllowed;
    private readonly TimeSpan cellTimeout;
    private readonly TimeSpan lifecycleTimeout;
    private readonly ExecutionGate gate = new();
    private readonly StaThread sta = new();
    private IExcelDriver? driver;
    private OwnershipRecordFile? record;
    private volatile DiagnosticLog? log;
    private string? workbookPath;
    private volatile ProcessIdentity? identity;
    private volatile bool generationStarted;
    private volatile bool exitConfirmed;
    private bool closeStarted;

    public ExcelSession(bool excelAllowed)
        : this(activeGate => new ComExcelDriver(activeGate), excelAllowed, TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(120)) { }

    internal ExcelSession(Func<ExecutionGate, IExcelDriver> createDriver, bool excelAllowed, TimeSpan cellTimeout, TimeSpan lifecycleTimeout)
    {
        this.createDriver = createDriver;
        this.excelAllowed = excelAllowed;
        this.cellTimeout = cellTimeout;
        this.lifecycleTimeout = lifecycleTimeout;
    }

    public SessionSummary Summary
    {
        get
        {
            var snapshot = gate.Snapshot;
            return new(snapshot.State.ToString().ToLowerInvariant(), snapshot.Busy);
        }
    }

    public SessionOperationException InvalidArgument(string operation, string message) => Rejected(operation,
        new CoreException(ErrorCode.InvalidRequest, ErrorDetail.InvalidArgument, message + " Fix the arguments and retry."));

    public Task<WorkbookResult> StartAsync(string projectPath, bool force = false)
    {
        const string operation = "session_start";
        return Accepted(operation, () =>
        {
            if (force) throw NotImplemented("force");
            if (!Path.IsPathFullyQualified(projectPath) || !ProjectConfig.IsLocal(Path.GetFullPath(projectPath)))
                throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.InvalidArgument,
                    "project_path must be an absolute path on a local drive. Fix it and retry.");
            return Run(operation, lifecycleTimeout, SessionState.Closed, SessionState.Ready, () => Start(Path.GetFullPath(projectPath)));
        });
    }

    public Task<RangeResult> GetRangeAsync(string sheet, string address, IReadOnlyList<string>? fields = null)
    {
        const string operation = "range_get";
        return Accepted(operation, () =>
        {
            string normalized = CellAddress.Normalize(address);
            fields ??= ["values"];
            if (fields.Count == 0 || fields.Distinct().Count() != fields.Count || fields.Except(["values", "formulas", "merged_ranges"]).Any())
                throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.InvalidArgument,
                    "fields must be unique values among values, formulas and merged_ranges. Fix them and retry.");
            if (fields.Contains("merged_ranges")) throw NotImplemented("merged_ranges");
            bool values = fields.Contains("values");
            bool formulas = fields.Contains("formulas");
            return Run(operation, cellTimeout, SessionState.Ready, SessionState.Ready, () =>
            {
                var content = gate.Step("read", () => driver!.Read(sheet, normalized, values, formulas));
                return new RangeResult(sheet, normalized, values ? [[CellValues.Project(content.Value)]] : null,
                    formulas ? [[content.Formula]] : null);
            });
        });
    }

    public Task<RangeResult> SetRangeAsync(string sheet, string address, object?[][]? values, string?[][]? formulas)
    {
        const string operation = "range_set";
        return Accepted(operation, () =>
        {
            string normalized = CellAddress.Normalize(address);
            var write = CellValues.Validate(values, formulas);
            return Run(operation, cellTimeout, SessionState.Ready, SessionState.Ready, () =>
            {
                gate.Step("write", () => driver!.Write(sheet, normalized, write));
                return new RangeResult(sheet, normalized);
            });
        });
    }

    public Task<WorkbookResult> SaveAsync()
    {
        const string operation = "session_save";
        return Accepted(operation, () => Run(operation, lifecycleTimeout, SessionState.Ready, SessionState.Ready, () =>
        {
            gate.Step("save", () => driver!.Save());
            return new WorkbookResult(workbookPath!);
        }));
    }

    public Task<WorkbookResult> EndAsync(bool discard = false, bool force = false)
    {
        const string operation = "session_end";
        return Accepted(operation, () =>
        {
            if (discard) throw NotImplemented("discard");
            if (force) throw NotImplemented("force");
            return Run(operation, lifecycleTimeout, SessionState.Ready, SessionState.Closed, () =>
            {
                string path = workbookPath!;
                gate.Step("save", () => driver!.Save());
                log?.Write("save.confirmed", new { workbook_path = path });
                CloseAndQuit();
                return new WorkbookResult(path);
            });
        });
    }

    // EOF: closes without saving, only from a ready session with no COM call pending and when the end safety checks
    // pass. An isolated session belongs to manual recovery, so its Excel, record and evidence stay untouched.
    // The OS releasing the lock at exit is not evidence of a cleared record.
    public async Task ShutdownAsync()
    {
        gate.RequestShutdown();
        var snapshot = gate.Snapshot;
        if (snapshot.State == SessionState.Closed) return;
        if (snapshot is not { State: SessionState.Ready, Busy: false })
        {
            log?.Write("shutdown.skipped", new { state = snapshot.State.ToString(), snapshot.Busy, snapshot.Stage, process = identity });
            return;
        }
        try { gate.Begin("shutdown", lifecycleTimeout, forShutdown: true); }
        catch (WorkStoppedException)
        {
            log?.Write("shutdown.skipped", new { state = snapshot.State.ToString(), process = identity });
            return;
        }
        Task<bool> closing = sta.Invoke(() =>
        {
            try
            {
                CloseAndQuit();
                gate.Complete(SessionState.Closed);
                log?.Write("shutdown.closed_without_save", new { process = identity });
                return true;
            }
            catch (Exception error)
            {
                _ = Failed("shutdown", error);
                return false;
            }
        });
        await Task.WhenAny(closing, Task.Delay(lifecycleTimeout));
        if (!closing.IsCompleted) log?.Write("shutdown.timed_out", new { gate.Snapshot.Stage, process = identity });
    }

    public void Dispose() => sta.Dispose();

    private WorkbookResult Start(string root)
    {
        generationStarted = false;
        exitConfirmed = false;
        closeStarted = false;
        identity = null;
        workbookPath = null;
        log = null;
        var config = gate.Step("config", () => ProjectConfig.Load(root));
        if (!excelAllowed)
            throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.NotImplemented,
                "Development mode does not start Excel. The configuration is valid; use a prepared copy without --development.");
        log = new DiagnosticLog(Path.Combine(config.Workspace, "logs"));
        var file = gate.Step("lock", () => AcquireRecord(root));
        record = file;
        var replaced = gate.Step("record", () => file.CheckPrevious());
        if (replaced is not null) log.Write("record.replaced", new { record = replaced });
        // An empty record is refused on the next start, so a new file must be valid before anything can fail.
        if (file.CreatedNew) gate.Step("record", () => file.Write(new(1, "completed")));
        gate.Step("lock", () => RequireWritable(config.Workbook));
        var active = createDriver(gate);
        driver = active;
        RecordTransition.Run(gate, file, new(1, "starting", config.Workbook), () =>
        {
            generationStarted = true;
            StartStep("excel_start", ErrorDetail.ExcelStartFailed, active.Generate);
        });
        StartStep("excel_start", ErrorDetail.ExcelStartFailed, () => identity = active.Identify());
        RecordTransition.Run(gate, file, new(1, "owned", config.Workbook, identity), () =>
            StartStep("open", ErrorDetail.WorkbookOpenFailed, () => active.Open(config.Workbook)));
        workbookPath = config.Workbook;
        return new WorkbookResult(config.Workbook);
    }

    private void StartStep(string stage, string detail, Action action)
    {
        try { gate.Step(stage, action); }
        catch (Exception error) when (error is not (CoreException or WorkStoppedException))
        {
            throw new CoreException(ErrorCode.ExcelError, detail, "Excel could not start the session. Do not resend; follow manual recovery.", error);
        }
    }

    private static OwnershipRecordFile AcquireRecord(string root)
    {
        string path = Path.Combine(root, ".xlwright.lock");
        try { return OwnershipRecordFile.Acquire(path); }
        catch (IOException error) when (IsSharingViolation(error))
        {
            throw new CoreException(ErrorCode.Conflict, ErrorDetail.ProjectLocked,
                "Another server holds this project. Stop it or wait, then retry.", error);
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            throw new RecordIoException("existing", "open", error);
        }
    }

    // Probing before Excel starts gives closed-state refusals; Excel would otherwise open a read-only copy.
    private static void RequireWritable(string workbook)
    {
        if (File.GetAttributes(workbook).HasFlag(FileAttributes.ReadOnly)) throw ReadOnly(null);
        try { new FileStream(workbook, FileMode.Open, FileAccess.ReadWrite, FileShare.None).Dispose(); }
        catch (IOException error) when (IsSharingViolation(error))
        {
            throw new CoreException(ErrorCode.Conflict, ErrorDetail.WorkbookLocked,
                "Another process is using the workbook. Close it and retry.", error);
        }
        catch (UnauthorizedAccessException error) { throw ReadOnly(error); }
    }

    private static CoreException ReadOnly(Exception? inner) => new(ErrorCode.Conflict, ErrorDetail.WorkbookReadOnly,
        "The workbook can be opened only read-only. Make it writable and retry.", inner);

    private static bool IsSharingViolation(IOException error) => (error.HResult & 0xFFFF) is 32 or 33;

    private void CloseAndQuit()
    {
        var active = driver!;
        string path = workbookPath!;
        EndSafety.CloseAndQuit(
            requireEmpty =>
            {
                try { gate.Step("safety_check", () => active.VerifyEndSafety(path, requireEmpty)); }
                catch (Exception error) when (error is not (CoreException or WorkStoppedException))
                {
                    throw new CoreException(ErrorCode.StateUnknown, ErrorDetail.ExcelNeedsReview,
                        "The end safety check could not finish. Do not resend; follow manual recovery.", error);
                }
            },
            () =>
            {
                closeStarted = true;
                gate.Step("close", active.Close);
            },
            () => gate.Step("close", active.Quit));
        if (!gate.Step("process_exit", () => active.ReleaseAndWaitForExit(gate.Remaining)))
            throw new WorkStoppedException();
        exitConfirmed = true;
        var file = record!;
        RecordTransition.Run(gate, file, new(1, "completed"), () => gate.Step("close", () =>
        {
            file.Dispose();
            record = null;
        }));
        driver = null;
    }

    private async Task<T> Accepted<T>(string operation, Func<Task<T>> validateAndRun) where T : OperationResult
    {
        Task<T> running;
        try { running = validateAndRun(); }
        catch (CoreException error) { throw Rejected(operation, error); }
        return await running;
    }

    private async Task<T> Run<T>(string operation, TimeSpan timeout, SessionState required, SessionState next, Func<T> work)
        where T : OperationResult
    {
        try { gate.Accept(operation, timeout, required); }
        catch (CoreException error) { throw Rejected(operation, error); }
        Task<T> running = sta.Invoke(() =>
        {
            try
            {
                T result = work();
                gate.Complete(next);
                log?.Write(operation + ".completed", new { workbook_path = workbookPath, process = identity });
                return result;
            }
            catch (Exception error) { throw Failed(operation, error); }
        });
        await Task.WhenAny(running, Task.Delay(timeout));
        // A timer can fire just before the gate's clock expires. Only the gate decides, so a hung call must not be awaited.
        while (!running.IsCompleted && !gate.CheckTimeout()) await Task.WhenAny(running, Task.Delay(10));
        if (!running.IsCompleted)
        {
            // The active call is not stopped. Its late outcome is only logged, and busy stays until it unwinds.
            _ = running.ContinueWith(late => _ = late.Exception, TaskContinuationOptions.OnlyOnFaulted);
            string stage = gate.Snapshot.Stage;
            string? path = log?.Write("operation.timeout", new { operation, stage, process = identity });
            throw new SessionOperationException(Error(operation, new CoreException(ErrorCode.StateUnknown, ErrorDetail.OperationTimeout, TimeoutMessage),
                stage, path), Summary);
        }
        T completed = await running;
        completed.Session = Summary;
        return completed;
    }

    private const string TimeoutMessage = "The operation did not finish in time and its result is unknown. Do not resend; follow manual recovery.";

    private SessionOperationException Failed(string operation, Exception error)
    {
        bool timedOut = gate.CheckTimeout();
        string stage = error is CoreException { Detail: ErrorDetail.ProcessQueryFailed } ? "process_query" : gate.Snapshot.Stage;
        EndSafety.Fault(gate, generationStarted || timedOut, closeStarted, error);
        if (gate.Snapshot.State == SessionState.Closed)
        {
            // A failure before Excel generation releases this server's lock; the record stays on disk.
            record?.Dispose();
            record = null;
            driver = null;
        }
        CoreException reported = timedOut || error is WorkStoppedException ?
            new CoreException(ErrorCode.StateUnknown, ErrorDetail.OperationTimeout, TimeoutMessage, error) :
            error as CoreException ?? new CoreException(ErrorCode.InternalError, ErrorDetail.InternalFailure,
                "xlwright failed internally. Do not resend; check the log and follow manual recovery.", error);
        string? path = log?.Write("operation.failed", new
        {
            operation, stage, reported.Detail, error = error.GetType().Name, error.Message, hresult = error.HResult,
            stack = error.StackTrace, inner = error.InnerException?.GetType().Name, inner_message = error.InnerException?.Message,
            inner_hresult = error.InnerException?.HResult, process = identity, workbook_path = workbookPath,
            excel_may_remain = generationStarted && !exitConfirmed, excel_exit_confirmed = exitConfirmed
        });
        return new(Error(operation, reported, stage, path), Summary);
    }

    private SessionOperationException Rejected(string operation, CoreException error) =>
        new(new(error.CodeName, error.Detail, error.Message, operation, null, null, null, null), Summary);

    private ErrorInfo Error(string operation, CoreException error, string stage, string? path)
    {
        var owned = identity;
        return new(error.CodeName, error.Detail, error.Message, operation,
            Stages.FirstOrDefault(stage.StartsWith), generationStarted ? !exitConfirmed : null,
            owned is null ? null : new(owned.Pid, owned.StartedAtUtc, owned.Name), path);
    }

    private static CoreException NotImplemented(string feature) => new(ErrorCode.InvalidRequest, ErrorDetail.NotImplemented,
        $"{feature} is not implemented yet. Remove it and retry.");
}
