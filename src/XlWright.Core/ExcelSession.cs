using System.Diagnostics;
using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>The session state that every stage 1 result reports (docs/api.md).</summary>
public sealed record SessionSnapshot(
    string State, bool Busy, string? Operation = null, TimeSpan Elapsed = default, string? App = null, string? Workbook = null);

/// <summary>A tool's outcome: the result fields without "session", or a failure; and the session after it.</summary>
public sealed record SessionOutcome(JsonObject? Result, ToolFailure? Failure, SessionSnapshot Session);

/// <summary>
/// One session: start (docs/architecture.md "開始の順序" 1, 5, 6 and "所有の確定"), status, save, and end ("終了").
/// Timeouts end only the Excel whose ownership this session confirmed ("状態と結果不明").
/// </summary>
public sealed class ExcelSession(IExcelLauncher launcher, SessionTimeouts timeouts) : IDisposable
{
    private readonly object _gate = new();
    private string? _operation;
    private long _operationStarted;
    private bool _unusable;
    private Active? _active;

    private sealed record Active(AppFolder App, StaWorker Worker, IExcel Excel, IOwnedProcess Process);

    public SessionSnapshot Status()
    {
        lock (_gate)
        {
            return Snapshot();
        }
    }

    public Task<SessionOutcome> StartAsync(AppFolder app) => Run("session_start", () =>
    {
        if (_unusable)
        {
            return Fail(ToolFailure.Conflict, "SERVER_UNUSABLE",
                "A call to an earlier Excel has not returned, so this server cannot start a new session. Close this conversation and start the session in a new one.");
        }

        return _active is not null
            ? Fail(ToolFailure.Conflict, "SESSION_ACTIVE", "A session is already open in this conversation. End it with session_end first.")
            : null;
    }, () => StartCore(app));

    public Task<SessionOutcome> EndAsync(bool save) => Run("session_end", () =>
        _active is null ? Fail(ToolFailure.Conflict, "NO_SESSION", "There is no session to end. Start one with session_start.") : null,
        () => EndCore(_active!, save));

    public Task<SessionOutcome> SaveAsync() => Run("workbook_save", () =>
        _active is null ? Fail(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start.") : null,
        async () =>
        {
            var active = _active!;
            return await SaveCore(active, "Fix the cause and call workbook_save again, or call session_end with save false to discard the changes.")
                ?? (new JsonObject { ["workbook"] = active.App.WorkbookName }, null);
        });

    public void Dispose()
    {
        lock (_gate)
        {
            _active?.Worker.Dispose();
        }
    }

    // Checks the state and marks the operation under the lock, runs it, then reports the session after it.
    private async Task<SessionOutcome> Run(string operation, Func<(JsonObject?, ToolFailure?)?> check, Func<Task<(JsonObject?, ToolFailure?)>> body)
    {
        lock (_gate)
        {
            if (_operation is not null)
            {
                return new SessionOutcome(null, new ToolFailure(ToolFailure.Conflict, "BUSY",
                    $"{_operation} is running. Check it with session_status and call again when it has finished."), Snapshot());
            }

            if (check() is var (result, failure))
            {
                return new SessionOutcome(result, failure, Snapshot());
            }

            _operation = operation;
            _operationStarted = Stopwatch.GetTimestamp();
        }

        try
        {
            var (result, failure) = await body();
            lock (_gate)
            {
                _operation = null;
                return new SessionOutcome(result, failure, Snapshot());
            }
        }
        finally
        {
            lock (_gate)
            {
                _operation = null;
            }
        }
    }

    private async Task<(JsonObject?, ToolFailure?)> StartCore(AppFolder app)
    {
        var deadline = Stopwatch.StartNew();
        var worker = new StaWorker();
        var requestedUtc = DateTime.UtcNow;

        var launch = worker.Run(launcher.Launch);
        if (!await Completes(launch, timeouts.Start))
        {
            // Nothing is owned yet, so the Excel is not touched; the call is still pending on the STA.
            MarkUnusable();
            return Unknown("TIMEOUT", excelEnded: false, restartable: false, null,
                "Starting Excel did not finish within the time limit. The Excel it may have started is left untouched. Close this conversation and start again in a new one.");
        }

        if (launch.IsFaulted)
        {
            worker.Dispose();
            return Fail(ToolFailure.ExcelError, null, $"Excel did not start: {launch.Exception!.InnerException!.Message}");
        }

        var (excel, process) = launch.Result;
        var problem = process.Facts.OwnershipProblem(requestedUtc);
        if (problem is not null)
        {
            // Without a confirmed owner there is no ground to end that Excel (docs/architecture.md "所有の確定").
            return Unknown("OWNERSHIP_NOT_CONFIRMED", excelEnded: false, restartable: false, process.Facts,
                $"The started Excel could not be confirmed as this session's ({problem}), so it was left untouched. Ask the person to check the process in Task Manager.");
        }

        var active = new Active(app, worker, excel, process);
        var open = worker.Run(() =>
        {
            excel.TakeControl();
            return excel.Open(app.WorkbookPath);
        });
        if (!await Completes(open, Remaining(deadline, timeouts.Start)))
        {
            return await AfterTimeout(active, open, "Starting the session did not finish within the time limit.");
        }

        if (open.IsFaulted || open.Result)
        {
            var (code, detail, message) = open.IsFaulted
                ? (ToolFailure.ExcelError, "OPEN_FAILED", $"Excel could not open {app.WorkbookName}: {open.Exception!.InnerException!.Message}")
                : (ToolFailure.Conflict, "WORKBOOK_READ_ONLY", $"{app.WorkbookName} opened read-only. Close it in any other program, then start again.");
            var exit = await QuitAndWait(active, Remaining(deadline, timeouts.Start), new JsonObject());
            if (exit is { Item2: { Code: ToolFailure.StateUnknown } })
            {
                return exit.Value;
            }

            // The reason stays; when the exit was not confirmed, the handover goes in data.
            return exit is { Item2: { } notEnded }
                ? Fail(code, detail, message + " " + notEnded.Message, notEnded.Data)
                : Fail(code, detail, message);
        }

        lock (_gate)
        {
            _active = active;
        }

        return (new JsonObject
        {
            ["app"] = app.RelativePath,
            ["workbook"] = app.WorkbookName,
            ["excel"] = ExcelJson(process.Facts),
        }, null);
    }

    private async Task<(JsonObject?, ToolFailure?)> EndCore(Active active, bool save)
    {
        if (save && await SaveCore(active, "Call session_end with save false to discard the changes.") is { } saveFailed)
        {
            return saveFailed;
        }

        var listing = active.Worker.Run(active.Excel.OpenWorkbookPaths);
        if (!await Completes(listing, timeouts.Sync))
        {
            return await AfterTimeout(active, listing, "Checking the open workbooks did not finish within the time limit.");
        }

        var foreign = listing.IsFaulted ? null : listing.Result.Where(path => path is not null && !IsInside(path, active.App.FullPath)).ToList();
        if (foreign is null || foreign.Count > 0)
        {
            // Left running and untouched; only the references are released (UserControl is true, so it stays).
            if (!await Completes(active.Worker.Run(active.Excel.ReleaseReferences), timeouts.Sync))
            {
                MarkUnusable();
            }

            Close(active);
            var data = Handover(active.Process.Facts);
            data["saved"] = save;
            var why = foreign is null
                ? "The open workbooks could not be checked"
                : "A workbook outside the app folder is open in the dedicated Excel";
            return Fail(ToolFailure.Conflict, "FOREIGN_WORKBOOK",
                $"{why}, so it was left running. " +
                $"Ask the person to check Excel (PID {active.Process.Facts.Pid}), save what they need, and close it.", data);
        }

        var exit = await QuitAndWait(active, timeouts.QuitExit, new JsonObject { ["saved"] = save });
        return exit ?? (new JsonObject { ["saved"] = save }, null);
    }

    // Returns null when the main workbook was saved. A failed save keeps the session; `next` says what to do then.
    private async Task<(JsonObject?, ToolFailure?)?> SaveCore(Active active, string next)
    {
        var saving = active.Worker.Run(active.Excel.Save);
        if (!await Completes(saving, timeouts.Save))
        {
            return await AfterTimeout(active, saving,
                "Saving did not finish within the time limit. The main workbook file may hold the contents from before or after the save.");
        }

        return saving.IsFaulted
            ? Fail(ToolFailure.ExcelError, "SAVE_FAILED",
                $"Saving {active.App.WorkbookName} failed: {saving.Exception!.InnerException!.Message} The session is still open. {next}")
            : null;
    }

    // docs/architecture.md "終了" 3-5: close and quit, wait, then end the owned process through its handle.
    // A close or quit that does not return in time is a timeout ("状態と結果不明"). `data` goes into EXIT_NOT_CONFIRMED.
    private async Task<(JsonObject?, ToolFailure?)?> QuitAndWait(Active active, TimeSpan quitExit, JsonObject data)
    {
        var quit = active.Worker.Run(active.Excel.CloseAllAndQuit);
        if (!await Completes(quit, timeouts.Sync))
        {
            return await AfterTimeout(active, quit, "Closing the workbooks and quitting Excel did not finish within the time limit.");
        }

        var ended = active.Process.WaitForExit(quitExit);
        if (!ended)
        {
            active.Process.Terminate();
            ended = active.Process.WaitForExit(timeouts.TerminateExit);
        }

        Close(active);
        if (ended)
        {
            return null;
        }

        data["excel"] = ExcelJson(active.Process.Facts);
        return Fail(ToolFailure.Conflict, "EXIT_NOT_CONFIRMED",
            $"The dedicated Excel (PID {active.Process.Facts.Pid}) did not end. Ask the person to check it in Task Manager.",
            data);
    }

    // A time limit passed: end the owned Excel without checking workbooks (docs/architecture.md "状態と結果不明").
    private async Task<(JsonObject?, ToolFailure?)> AfterTimeout(Active active, Task pending, string what)
    {
        active.Process.Terminate();
        var ended = active.Process.WaitForExit(timeouts.TerminateExit);
        var returned = await Completes(pending, timeouts.PendingReturn);
        if (!returned)
        {
            MarkUnusable();
        }

        Close(active);
        var lost = " The changes since the last save are lost.";
        return ended && returned
            ? Unknown("TIMEOUT", excelEnded: true, restartable: true, null, what + " The dedicated Excel was ended." + lost + " You can start a new session.")
            : Unknown("TIMEOUT", ended, restartable: false, ended ? null : active.Process.Facts,
                what + (ended ? " The dedicated Excel was ended." + lost : " The dedicated Excel could not be ended.") +
                " Close this conversation and start the session in a new one.");
    }

    private void Close(Active active)
    {
        lock (_gate)
        {
            if (ReferenceEquals(_active, active))
            {
                _active = null;
            }
        }

        if (!_unusable)
        {
            active.Worker.Dispose();
        }

        active.Process.Dispose();
    }

    private void MarkUnusable()
    {
        lock (_gate)
        {
            _unusable = true;
        }
    }

    private SessionSnapshot Snapshot() => new(
        _active is null ? "closed" : "ready",
        _operation is not null,
        _operation,
        _operation is null ? TimeSpan.Zero : Stopwatch.GetElapsedTime(_operationStarted),
        _active?.App.RelativePath,
        _active?.App.WorkbookName);

    private static async Task<bool> Completes(Task task, TimeSpan timeout) =>
        await Task.WhenAny(task, Task.Delay(timeout)) == task;

    private static TimeSpan Remaining(Stopwatch elapsed, TimeSpan limit) =>
        limit > elapsed.Elapsed ? limit - elapsed.Elapsed : TimeSpan.Zero;

    // A path Excel reports that is not a full local path (a URL, a bare name) counts as outside the app folder.
    private static bool IsInside(string path, string folder) =>
        Path.IsPathFullyQualified(path) &&
        Path.GetFullPath(path).StartsWith(Path.TrimEndingDirectorySeparator(folder) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    private static JsonObject ExcelJson(ProcessFacts facts) => new()
    {
        ["pid"] = facts.Pid,
        ["started_at_utc"] = facts.CreatedUtc.ToString("O"),
    };

    private static JsonObject Handover(ProcessFacts facts) => new() { ["excel"] = ExcelJson(facts) };

    private static (JsonObject?, ToolFailure?) Fail(string code, string? detail, string message, JsonObject? data = null) =>
        (null, new ToolFailure(code, detail, message, data));

    private static (JsonObject?, ToolFailure?) Unknown(string detail, bool excelEnded, bool restartable, ProcessFacts? leftRunning, string message)
    {
        var data = leftRunning is null ? new JsonObject() : Handover(leftRunning);
        data["excel_ended"] = excelEnded;
        data["restartable"] = restartable;
        return Fail(ToolFailure.StateUnknown, detail, message, data);
    }
}
