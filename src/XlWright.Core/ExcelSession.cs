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
    // docs/notes/dialog-probe.md: every dialog was found at the first 250 ms check.
    private static readonly TimeSpan DialogInterval = TimeSpan.FromMilliseconds(250);

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

    public Task<SessionOutcome> SaveAsync() => Run("workbook_save", NoSession,
        async () =>
        {
            var active = _active!;
            return await SaveCore(active, "Fix the cause and call workbook_save again, or call session_end with save false to discard the changes.")
                ?? (new JsonObject { ["workbook"] = active.App.WorkbookName }, null);
        });

    public Task<SessionOutcome> ReadAsync(string sheet, CellRange range, bool formulas) => Run("range_read", NoSession, async () =>
    {
        var active = _active!;
        var reading = active.Worker.Run<Array?>(formulas
            ? () => active.Excel.ReadFormulas(sheet, range)
            : () => active.Excel.ReadValues(sheet, range));
        if (!await Completes(reading, timeouts.Sync))
        {
            return await AfterTimeout(active, reading, $"Reading {range.Address} did not finish within the time limit.");
        }

        if (reading.IsFaulted)
        {
            return Fail(ToolFailure.ExcelError, null, $"Excel could not read {range.Address} on {sheet}: {reading.Exception!.InnerException!.Message}");
        }

        return reading.Result is { } cells
            ? (new JsonObject { ["sheet"] = sheet, ["range"] = range.Address, [formulas ? "formulas" : "values"] = CellsJson(cells) }, null)
            : SheetNotFound(sheet);
    });

    public Task<SessionOutcome> WriteAsync(string sheet, CellRange range, RangeWrite write) => Run("range_write", NoSession, async () =>
    {
        var active = _active!;
        var writing = active.Worker.Run(() => active.Excel.Write(sheet, range, write));
        if (!await Completes(writing, timeouts.Sync))
        {
            return await AfterTimeout(active, writing, $"Writing {range.Address} did not finish within the time limit.");
        }

        if (writing.IsFaulted)
        {
            // A failed array write can stop part-way (docs/notes/range-values.md), so what was written is unknown.
            return Fail(ToolFailure.ExcelError, "WRITE_FAILED",
                $"Excel could not write {range.Address} on {sheet}: {writing.Exception!.InnerException!.Message} " +
                "Some cells may have been written; read the range again with range_read before you continue.");
        }

        return writing.Result
            ? (new JsonObject { ["sheet"] = sheet, ["range"] = range.Address, ["cells"] = range.Count }, null)
            : SheetNotFound(sheet);
    });

    public Task<SessionOutcome> SyncAsync() => Run("vba_sync", NoSession, () => SyncCore(_active!));

    /// <summary>
    /// Syncs, then runs the macro and waits up to <paramref name="limit"/>, checking for a dialog every 250 ms
    /// (docs/architecture.md "マクロの実行"). A dialog or the limit ends the owned Excel.
    /// </summary>
    public Task<SessionOutcome> RunAsync(MacroName macro, IReadOnlyList<object> args, TimeSpan limit) => Run("macro_run", NoSession, async () =>
    {
        var active = _active!;
        var (sync, syncFailure) = await SyncCore(active);
        if (syncFailure is not null)
        {
            return (null, syncFailure);
        }

        var elapsed = Stopwatch.StartNew();
        var running = active.Worker.Run(() => active.Excel.RunMacro(macro, args));
        while (!await Completes(running, DialogInterval))
        {
            if (active.Process.FindDialog() is { } dialog)
            {
                return await AfterTimeout(active, running, $"{macro} stopped at a dialog.", "DIALOG",
                    new JsonObject { ["dialog"] = new JsonObject { ["title"] = dialog.Title, ["text"] = dialog.Text } });
            }

            if (elapsed.Elapsed >= limit)
            {
                return await AfterTimeout(active, running, $"{macro} did not finish within {limit.TotalSeconds:0.###} seconds.");
            }
        }

        if (running.IsFaulted)
        {
            // Excel refused the call itself, such as after an End statement (docs/notes/macro-run.md).
            var error = running.Exception!.InnerException!;
            return MacroError(macro, error.HResult, error.Message, "");
        }

        var time = (long)elapsed.Elapsed.TotalMilliseconds;
        return running.Result switch
        {
            MacroResult.NotFound missing => Fail(ToolFailure.NotFound, "MACRO_NOT_FOUND", missing.Reason),
            MacroResult.Failed failed => MacroError(macro, failed.Number, failed.Description, failed.Source),
            MacroResult.OfType typed => (Returned(macro, null, typed.TypeName, time, sync!), null),
            MacroResult.Returned returned => (ReturnedValue(active.App, macro, returned.Value, time, sync!), null),
            _ => throw new InvalidOperationException("Unknown macro result."),
        };
    });

    private async Task<(JsonObject?, ToolFailure?)> SyncCore(Active active)
    {
        if (!Path.GetExtension(active.App.WorkbookName).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return Fail(ToolFailure.InvalidRequest, "NOT_MACRO_ENABLED",
                $"{active.App.WorkbookName} is not a macro-enabled workbook (.xlsm), so it cannot hold VBA.");
        }

        var (modules, invalid) = VbaSources.Prepare(active.App, VbaSources.SystemAnsiEncoding(), ServerIdentity.Current.Version);
        if (invalid is not null)
        {
            return (null, invalid);
        }

        try
        {
            var syncing = active.Worker.Run(() => active.Excel.SyncVba(modules!));
            if (!await Completes(syncing, timeouts.Sync))
            {
                return await AfterTimeout(active, syncing, "Syncing the VBA modules did not finish within the time limit.");
            }

            if (syncing.IsFaulted)
            {
                return syncing.Exception!.InnerException switch
                {
                    VbaSyncException { Detail: "SYNC_FAILED" } e => SyncFailed(e.Message, e.Done),
                    VbaSyncException e => Fail(ToolFailure.Conflict, e.Detail, e.Message),
                    var e => SyncFailed($"Syncing the VBA modules failed: {e!.Message}", []),
                };
            }

            var listed = new JsonArray();
            foreach (var module in modules!)
            {
                listed.Add(new JsonObject { ["name"] = module.Name, ["type"] = module.Kind == VbaModuleKind.Class ? "class" : "standard" });
            }

            return (new JsonObject { ["modules"] = listed, ["removed"] = new JsonArray(syncing.Result.Select(name => (JsonNode?)name).ToArray()) }, null);
        }
        finally
        {
            // Prepare always adds XlWright, so there is a first module. When an Import timed out, Excel may still
            // hold a file; the folder then stays under .xlwright/sync/, which holds only converted copies.
            try
            {
                Directory.Delete(Path.GetDirectoryName(modules![0].ImportPath)!, recursive: true);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
            }
        }
    }

    private static (JsonObject?, ToolFailure?) MacroError(MacroName macro, int number, string description, string source) =>
        Fail(ToolFailure.ExcelError, "MACRO_ERROR",
            $"{macro} raised run-time error {number}: {description} The session is still open; nothing was saved.",
            new JsonObject { ["number"] = number, ["description"] = description, ["source"] = source });

    // A return value too long for the response goes whole into <app>/.xlwright/results/<id>.json (docs/api.md "成功と失敗").
    private static JsonObject ReturnedValue(AppFolder app, MacroName macro, object? value, long time, JsonObject sync)
    {
        var (json, typeName) = MacroCall.ReturnValue(value);
        var result = Returned(macro, json, typeName, time, sync);
        if (value is not string text || JsonValue.Create(text).ToJsonString(JsonText.Options).Length <= MacroCall.MaxReturnCharacters)
        {
            return result;
        }

        var id = Guid.NewGuid().ToString("N");
        var folder = Path.Combine(app.FullPath, ".xlwright", "results");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, id + ".json"), new JsonObject { ["macro"] = macro.ToString(), ["return"] = text }.ToJsonString(JsonText.Options));
        var summary = Returned(macro, JsonValue.Create(text[..MacroCall.SummaryCharacters]), null, time, sync);
        summary["result_file"] = $"{app.RelativePath}/.xlwright/results/{id}.json";
        return summary;
    }

    private static JsonObject Returned(MacroName macro, JsonNode? value, string? typeName, long time, JsonObject sync)
    {
        var result = new JsonObject { ["macro"] = macro.ToString(), ["return"] = value };
        if (typeName is not null)
        {
            result["return_type"] = typeName;
        }

        result["elapsed_ms"] = time;
        result["sync"] = sync.DeepClone();
        return result;
    }

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
    // A dialog ends the owned Excel the same way, with its own detail and the dialog in data.
    private async Task<(JsonObject?, ToolFailure?)> AfterTimeout(Active active, Task pending, string what, string detail = "TIMEOUT", JsonObject? extra = null)
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
            ? Unknown(detail, excelEnded: true, restartable: true, null, what + " The dedicated Excel was ended." + lost + " You can start a new session.", extra)
            : Unknown(detail, ended, restartable: false, ended ? null : active.Process.Facts,
                what + (ended ? " The dedicated Excel was ended." + lost : " The dedicated Excel could not be ended.") +
                " Close this conversation and start the session in a new one.", extra);
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

    // Called under the lock by Run.
    private (JsonObject?, ToolFailure?)? NoSession() =>
        _active is null ? Fail(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start.") : null;

    // Nothing is saved, so the next sync replaces every module again (docs/architecture.md "VBAの反映").
    private static (JsonObject?, ToolFailure?) SyncFailed(string message, IReadOnlyList<string> done) =>
        Fail(ToolFailure.ExcelError, "SYNC_FAILED", message + " The workbook was not saved; fix the cause and call vba_sync again.",
            new JsonObject { ["done"] = new JsonArray(done.Select(name => (JsonNode?)name).ToArray()) });

    private static (JsonObject?, ToolFailure?) SheetNotFound(string sheet) =>
        Fail(ToolFailure.NotFound, "SHEET_NOT_FOUND", $"The main workbook has no sheet named \"{sheet}\". Check the sheet name.");

    // Values are double, string, bool, null, or CellError (IExcel.ReadValues); formulas are strings.
    private static JsonArray CellsJson(Array cells)
    {
        var rows = new JsonArray();
        for (var row = 0; row < cells.GetLength(0); row++)
        {
            var columns = new JsonArray();
            for (var column = 0; column < cells.GetLength(1); column++)
            {
                columns.Add(cells.GetValue(row, column) switch
                {
                    null => null,
                    CellError error => new JsonObject { ["error"] = error.Name },
                    double number => JsonValue.Create(number),
                    string text => JsonValue.Create(text),
                    bool flag => JsonValue.Create(flag),
                    var other => throw new InvalidOperationException($"Excel returned a cell of type {other.GetType().Name}."),
                });
            }

            rows.Add(columns);
        }

        return rows;
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

    private static (JsonObject?, ToolFailure?) Unknown(
        string detail, bool excelEnded, bool restartable, ProcessFacts? leftRunning, string message, JsonObject? extra = null)
    {
        var data = leftRunning is null ? new JsonObject() : Handover(leftRunning);
        data["excel_ended"] = excelEnded;
        data["restartable"] = restartable;
        foreach (var (name, value) in extra ?? [])
        {
            data[name] = value?.DeepClone();
        }

        return Fail(ToolFailure.StateUnknown, detail, message, data);
    }
}
