using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>The session state that every stage 1 result reports (docs/api.md).</summary>
public sealed record SessionSnapshot(
    string State, bool Busy, string? Operation = null, TimeSpan Elapsed = default, string? App = null, string? Workbook = null);

/// <summary>A tool's outcome: the result fields without "session", or a failure; and the session after it.</summary>
public sealed record SessionOutcome(JsonObject? Result, ToolFailure? Failure, SessionSnapshot Session);

/// <summary>
/// One session: start (docs/architecture.md "開始の順序" 1-6, "所有記録の形式", and "所有の確定"), status, save, and end ("終了").
/// Timeouts end only the Excel whose ownership this session confirmed ("状態と結果不明").
/// </summary>
public sealed class ExcelSession(IExcelLauncher launcher, IProcessQuery processes, IExcelConnector connector, SessionTimeouts timeouts) : IDisposable
{
    // docs/notes/dialog-probe.md: every dialog was found at the first 250 ms check.
    private static readonly TimeSpan DialogInterval = TimeSpan.FromMilliseconds(250);

    private readonly object _gate = new();
    private string? _operation;
    private long _operationStarted;
    private bool _unusable;
    private Active? _active;
    private DiagnosticLog? _log;

    // App locks kept while a call is pending; holding them here keeps the finalizer from closing the file.
    private readonly List<AppLock> _keptLocks = [];

    private sealed record Active(AppFolder App, StaWorker Worker, IExcel Excel, IOwnedProcess Process, AppLock Lock, DateTime Marked);

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
    /// Syncs, compiles the main workbook's project, and watches for the compile error dialog (docs/api.md
    /// "vba_check"). A compile error is a result: the dialog is read, closed with WM_CLOSE, and the VBE is hidden.
    /// </summary>
    public Task<SessionOutcome> CheckAsync() => Run("vba_check", NoSession, async () =>
    {
        var active = _active!;
        var (sync, syncFailure) = await SyncCore(active);
        if (syncFailure is not null)
        {
            return (null, syncFailure);
        }

        var starting = active.Worker.Run(active.Excel.StartCompile);
        if (!await Completes(starting, timeouts.Sync))
        {
            return await AfterTimeout(active, starting, "Compiling the VBA project did not finish within the time limit.");
        }

        if (starting.IsFaulted)
        {
            return Fail(ToolFailure.ExcelError, null, $"Excel could not compile the VBA project: {starting.Exception!.InnerException!.Message}");
        }

        if (starting.Result == CompileStart.AlreadyCompiled)
        {
            return (Checked(true, null, sync!), null);
        }

        // The control returns at once and a compile error shows its dialog later; the control stays enabled then.
        var watch = Stopwatch.StartNew();
        while (watch.Elapsed < timeouts.CompileWatch)
        {
            if (active.Process.FindDialog() is { } dialog)
            {
                return await CompileError(active, dialog, sync!);
            }

            var checking = active.Worker.Run(active.Excel.IsCompiled);
            if (!await Completes(checking, timeouts.Sync))
            {
                return await AfterTimeout(active, checking, "Checking the compile did not finish within the time limit.");
            }

            if (checking.IsFaulted)
            {
                return Fail(ToolFailure.ExcelError, null, $"Excel could not report the compile: {checking.Exception!.InnerException!.Message}");
            }

            if (checking.Result)
            {
                return (Checked(true, null, sync!), null);
            }

            await Task.Delay(DialogInterval);
        }

        return Fail(ToolFailure.ExcelError, "COMPILE_NOT_CONFIRMED",
            $"The compile showed no error and was not confirmed within {timeouts.CompileWatch.TotalSeconds:0.###} seconds. The session is still open; call vba_check again.");
    });

    private async Task<(JsonObject?, ToolFailure?)> CompileError(Active active, DialogText dialog, JsonObject sync)
    {
        var shown = new JsonObject { ["dialog"] = new JsonObject { ["title"] = dialog.Title, ["text"] = dialog.Text } };
        var reading = active.Worker.Run(active.Excel.ReadErrorPlace);
        if (!await Completes(reading, timeouts.Sync))
        {
            return await AfterTimeout(active, reading, "Reading the place of the compile error did not finish within the time limit.", "DIALOG", shown);
        }

        // The place is optional in the result; the VBE does not always have a code pane (docs/notes/test-copy.md).
        var place = reading.IsFaulted ? null : reading.Result;
        active.Process.CloseDialog();
        var closing = Stopwatch.StartNew();
        while (active.Process.FindDialog() is not null)
        {
            if (closing.Elapsed >= timeouts.DialogClose)
            {
                // Excel did not quit with the dialog left open (docs/notes/test-copy.md), so it is ended now.
                return await AfterTimeout(active, Task.CompletedTask, "The compile error dialog stayed open after it was closed.", "DIALOG", shown);
            }

            await Task.Delay(DialogInterval / 5);
        }

        var hiding = active.Worker.Run(active.Excel.HideEditor);
        if (!await Completes(hiding, timeouts.Sync))
        {
            return await AfterTimeout(active, hiding, "Hiding the VBA editor did not finish within the time limit.");
        }

        var error = new JsonObject { ["message"] = string.Join(" ", dialog.Text.Split('\n').Select(line => line.Trim()).Where(line => line.Length > 0)) };
        if (place is not null)
        {
            error["module"] = place.Module;
            error["line"] = place.Line;
        }

        return (Checked(false, error, sync), null);
    }

    private static JsonObject Checked(bool compiled, JsonObject? error, JsonObject sync)
    {
        var result = new JsonObject { ["compiled"] = compiled };
        if (error is not null)
        {
            result["error"] = error;
        }

        result["sync"] = sync;
        return result;
    }

    /// <summary>Lists the modules of the main workbook as they are, without a sync (docs/api.md "vba_get").</summary>
    public Task<SessionOutcome> ListModulesAsync() => Run("vba_get", NoSession, async () =>
    {
        var active = _active!;
        if (MacrosRejected(active) is { } rejected)
        {
            return rejected;
        }

        var listing = active.Worker.Run(active.Excel.ListVba);
        if (!await Completes(listing, timeouts.Sync))
        {
            return await AfterTimeout(active, listing, "Listing the VBA modules did not finish within the time limit.");
        }

        if (listing.IsFaulted)
        {
            return listing.Exception!.InnerException switch
            {
                VbaSyncException e => Fail(ToolFailure.Conflict, e.Detail, e.Message),
                var e => Fail(ToolFailure.ExcelError, null, $"Excel could not list the VBA modules: {e!.Message}"),
            };
        }

        var modules = new JsonArray();
        foreach (var component in listing.Result)
        {
            var module = new JsonObject { ["name"] = component.Name, ["type"] = component.Type, ["lines"] = component.Lines };
            if (component.Sheet is not null)
            {
                module["sheet"] = component.Sheet;
            }

            modules.Add(module);
        }

        return (new JsonObject { ["modules"] = modules }, null);
    });

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
    }, new JsonObject { ["macro"] = macro.ToString() });

    private static (JsonObject?, ToolFailure?)? MacrosRejected(Active active) =>
        Path.GetExtension(active.App.WorkbookName).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            ? null
            : Fail(ToolFailure.InvalidRequest, "NOT_MACRO_ENABLED",
                $"{active.App.WorkbookName} is not a macro-enabled workbook (.xlsm), so it cannot hold VBA.");

    private async Task<(JsonObject?, ToolFailure?)> SyncCore(Active active)
    {
        if (MacrosRejected(active) is { } rejected)
        {
            return rejected;
        }

        var (modules, documents, invalid) = VbaSources.Prepare(active.App, VbaSources.SystemAnsiEncoding(), ServerIdentity.Current.Version);
        if (invalid is not null)
        {
            return (null, invalid);
        }

        try
        {
            var syncing = active.Worker.Run(() => active.Excel.SyncVba(modules!, documents!));
            if (!await Completes(syncing, timeouts.Sync))
            {
                return await AfterTimeout(active, syncing, "Syncing the VBA modules did not finish within the time limit.");
            }

            if (syncing.IsFaulted)
            {
                return syncing.Exception!.InnerException switch
                {
                    VbaSyncException { Detail: "SYNC_FAILED" } e => SyncFailed(e.Message, e.Done),
                    VbaSyncException { Detail: "WORKBOOK_MODULE_NOT_FOUND" } e => Fail(ToolFailure.NotFound, e.Detail, e.Message, new JsonObject { ["file"] = e.File }),
                    VbaSyncException e => Fail(ToolFailure.Conflict, e.Detail, e.Message),
                    var e => SyncFailed($"Syncing the VBA modules failed: {e!.Message}", []),
                };
            }

            var listed = new JsonArray();
            foreach (var module in modules!)
            {
                listed.Add(new JsonObject { ["name"] = module.Name, ["type"] = module.Kind == VbaModuleKind.Class ? "class" : "standard" });
            }

            foreach (var document in documents!)
            {
                listed.Add(new JsonObject { ["name"] = document.Name, ["type"] = "document" });
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
            _active?.Lock.Dispose();
            _keptLocks.ForEach(kept => kept.Dispose());
        }
    }

    // Checks the state and marks the operation under the lock, runs it, then reports the session after it.
    // `call` adds names to the log's call line, never values (docs/architecture.md "診断の記録の形式").
    private async Task<SessionOutcome> Run(string operation, Func<(JsonObject?, ToolFailure?)?> check, Func<Task<(JsonObject?, ToolFailure?)>> body,
        JsonObject? call = null)
    {
        var started = Stopwatch.GetTimestamp();
        lock (_gate)
        {
            if (_operation is not null)
            {
                return new SessionOutcome(null, Logged(operation, started, call, new ToolFailure(ToolFailure.Conflict, "BUSY",
                    $"{_operation} is running. Check it with session_status and call again when it has finished.")), Snapshot());
            }

            if (check() is var (result, failure))
            {
                return new SessionOutcome(result, Logged(operation, started, call, failure), Snapshot());
            }

            _operation = operation;
            _operationStarted = started;
        }

        try
        {
            var (result, failure) = await body();
            lock (_gate)
            {
                _operation = null;
                return new SessionOutcome(result, Logged(operation, started, call, failure), Snapshot());
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

    // Called under _gate: writes the call line, then points a failure to the log while it belongs to a session.
    // The log stays with the session, so it is dropped once no session is open (a failed start, or an end).
    private ToolFailure? Logged(string operation, long started, JsonObject? call, ToolFailure? failure)
    {
        var log = _log;
        if (log is null)
        {
            return failure;
        }

        var line = call?.DeepClone().AsObject() ?? [];
        line["tool"] = operation;
        line["elapsed_ms"] = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds;
        if (failure is null)
        {
            line["ok"] = true;
        }
        else
        {
            line["code"] = failure.Code;
            if (failure.Detail is not null)
            {
                line["detail"] = failure.Detail;
            }
        }

        log.Write("call", line);
        if (_active is null)
        {
            _log = null;
        }

        if (failure is null || log.RelativePath is not { } path)
        {
            return failure;
        }

        var data = failure.Data ?? new JsonObject();
        data["log"] = path;
        return failure with { Data = data };
    }

    // The cleanup goes into the result, and into a failure after it, since the record is completed by then
    // (docs/api.md "session_start に段階2で加えるもの").
    private async Task<(JsonObject?, ToolFailure?)> StartCore(AppFolder app)
    {
        var log = DiagnosticLog.Create(app);
        log.Write("start", new JsonObject
        {
            ["app"] = app.RelativePath,
            ["build_id"] = ServerIdentity.Current.BuildId,
            ["server_pid"] = Environment.ProcessId,
        });
        lock (_gate)
        {
            _log = log;
        }

        var cleanedUp = new StrongBox<JsonObject?>();
        var (result, failure) = await StartSteps(app, cleanedUp);
        if (cleanedUp.Value is not { } cleanup)
        {
            // Only the failures about an earlier session's Excel carry both the record and that Excel.
            if (failure?.Data is { } stop && stop["record"] is not null && stop["excel"] is JsonObject leftover)
            {
                var stopped = new JsonObject
                {
                    ["pid"] = leftover["pid"]!.DeepClone(),
                    ["started_at_utc"] = leftover["started_at_utc"]!.DeepClone(),
                };
                if (stop["cleanup"]?["ended_by"] is { } endedBy)
                {
                    stopped["ended_by"] = endedBy.DeepClone();
                }

                stopped["detail"] = failure.Detail;
                log.Write("cleanup", stopped);
            }

            return (result, failure);
        }

        log.Write("cleanup", new JsonObject
        {
            ["pid"] = cleanup["excel"]!["pid"]!.DeepClone(),
            ["started_at_utc"] = cleanup["excel"]!["started_at_utc"]!.DeepClone(),
            ["ended_by"] = cleanup["ended_by"]!.DeepClone(),
        });

        if (failure is null)
        {
            result!["cleanup"] = cleanup;
            return (result, null);
        }

        var data = failure.Data ?? new JsonObject();
        data["cleanup"] ??= cleanup;
        return (null, failure with { Data = data });
    }

    // docs/architecture.md "開始の順序": 2 the app lock, 3 the record, 4 the sharing check, 5 the mark and the start.
    private async Task<(JsonObject?, ToolFailure?)> StartSteps(AppFolder app, StrongBox<JsonObject?> cleanedUp)
    {
        var deadline = Stopwatch.StartNew();
        AppLock? appLock;
        try
        {
            appLock = AppLock.Take(app);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            return RecordInvalid(app, $"The ownership record could not be opened: {e.Message}");
        }

        if (appLock is null)
        {
            return Fail(ToolFailure.Conflict, "SESSION_IN_USE",
                $"Another conversation has a session for {app.RelativePath}. End it there with session_end, or close that conversation, then start again.");
        }

        DateTime marked;
        try
        {
            (var ready, cleanedUp.Value) = await CheckRecord(app, appLock, deadline);
            ready ??= CheckWorkbookFile(app);
            if (ready is not null)
            {
                ReleaseLock(appLock);
                return ready.Value;
            }

            RemoveResults(app);
            marked = DateTime.UtcNow;
            appLock.Write(new OwnershipRecord(OwnershipRecord.Starting, marked, app.WorkbookName, null));
        }
        catch
        {
            // This session has started nothing yet, so the app is freed now instead of by the finalizer.
            ReleaseLock(appLock);
            throw;
        }

        var worker = new StaWorker();
        var requestedUtc = DateTime.UtcNow;
        var launch = worker.Run(launcher.Launch);
        if (!await Completes(launch, timeouts.Start))
        {
            // Nothing is owned yet, so the Excel is not touched; the call is still pending on the STA, so the app stays locked.
            MarkUnusable();
            KeepLock(appLock);
            return Unknown("TIMEOUT", excelEnded: false, restartable: false, null,
                "Starting Excel did not finish within the time limit. The Excel it may have started is left untouched. Close this conversation and start again in a new one.");
        }

        if (launch.IsFaulted)
        {
            worker.Dispose();
            appLock.Dispose();
            return Fail(ToolFailure.ExcelError, "EXCEL_NOT_STARTED",
                $"Excel did not start: {launch.Exception!.InnerException!.Message} Check that Excel is installed and starts, then call session_start again.");
        }

        var (excel, process) = launch.Result;
        var problem = process.Facts.OwnershipProblem(requestedUtc);
        if (problem is not null)
        {
            // Without a confirmed owner there is no ground to end that Excel (docs/architecture.md "所有の確定").
            // The starting mark stays, so the next start stops at an Excel started after it.
            appLock.Dispose();
            return Unknown("OWNERSHIP_NOT_CONFIRMED", excelEnded: false, restartable: false, process.Facts,
                $"The started Excel could not be confirmed as this session's ({problem}), so it was left untouched. Ask the person to check the process in Task Manager.");
        }

        appLock.Write(new OwnershipRecord(OwnershipRecord.Owned, marked, app.WorkbookName,
            new RecordedProcess(process.Facts.Pid, process.Facts.CreatedUtc, Path.GetFileNameWithoutExtension(process.Facts.ImageName))));
        _log?.Write("owned", ExcelJson(process.Facts));
        var active = new Active(app, worker, excel, process, appLock, marked);
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

        var result = new JsonObject
        {
            ["app"] = app.RelativePath,
            ["workbook"] = app.WorkbookName,
            ["excel"] = ExcelJson(process.Facts),
        };
        return (result, null);
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

            Close(active, ended: false, endedBy: null);
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
        var endedBy = ended ? "quit" : "terminate";
        if (!ended)
        {
            active.Process.Terminate();
            ended = active.Process.WaitForExit(timeouts.TerminateExit);
        }

        Close(active, ended, endedBy);
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
        _log?.Write("windows", new JsonObject
        {
            ["xlmain_enabled"] = new JsonArray(active.Process.MainWindowsEnabled().Select(enabled => (JsonNode?)enabled).ToArray()),
            ["dialog"] = detail == "DIALOG" || active.Process.FindDialog() is not null,
        });
        active.Process.Terminate();
        var ended = active.Process.WaitForExit(timeouts.TerminateExit);
        var returned = await Completes(pending, timeouts.PendingReturn);
        if (!returned)
        {
            MarkUnusable();
        }

        Close(active, ended, "terminate");
        var lost = " The changes since the last save are lost.";
        return ended && returned
            ? Unknown(detail, excelEnded: true, restartable: true, null, what + " The dedicated Excel was ended." + lost + " You can start a new session.", extra)
            : Unknown(detail, ended, restartable: false, ended ? null : active.Process.Facts,
                what + (ended ? " The dedicated Excel was ended." + lost : " The dedicated Excel could not be ended.") +
                " Close this conversation and start the session in a new one.", extra);
    }

    // The record is completed only when the OS confirmed the exit (docs/architecture.md "完了にする条件"); otherwise it
    // stays owned for the next start. The app stays locked while a call is pending ("状態と結果不明").
    private void Close(Active active, bool ended, string? endedBy)
    {
        var end = new JsonObject();
        if (endedBy is not null)
        {
            end["ended_by"] = endedBy;
        }

        end["exit_confirmed"] = ended;
        _log?.Write("end", end);
        lock (_gate)
        {
            if (ReferenceEquals(_active, active))
            {
                _active = null;
            }
        }

        // Kept before the write, so a failed write does not leave the lock to the finalizer.
        if (_unusable)
        {
            KeepLock(active.Lock);
        }

        try
        {
            if (ended)
            {
                active.Lock.Write(new OwnershipRecord(OwnershipRecord.Completed, active.Marked, active.App.WorkbookName, null));
            }
        }
        finally
        {
            if (!_unusable)
            {
                active.Worker.Dispose();
                active.Lock.Dispose();
            }

            active.Process.Dispose();
        }
    }

    private void KeepLock(AppLock appLock)
    {
        lock (_gate)
        {
            _keptLocks.Add(appLock);
        }
    }

    // docs/architecture.md "所有記録の形式" and "完了にする条件": a first start completes the new record; a starting
    // mark completes when no Excel started after it; an owned record completes when its Excel has ended or was
    // cleaned up. Returns the cleanup for the result when a leftover was ended.
    private async Task<((JsonObject?, ToolFailure?)? Failure, JsonObject? Cleanup)> CheckRecord(AppFolder app, AppLock appLock, Stopwatch deadline)
    {
        var completed = new OwnershipRecord(OwnershipRecord.Completed, DateTime.UtcNow, app.WorkbookName, null);
        if (appLock.Created)
        {
            appLock.Write(completed);
            return (null, null);
        }

        OwnershipRecord? record;
        try
        {
            record = OwnershipRecord.Parse(appLock.Read());
        }
        catch (Exception e) when (e is IOException or System.Text.DecoderFallbackException)
        {
            return (RecordInvalid(app, $"The ownership record could not be read: {e.Message}"), null);
        }

        JsonObject? cleanup = null;
        switch (record?.State)
        {
            case null:
                return (RecordInvalid(app, "The ownership record is empty or not in the expected form."), null);
            case OwnershipRecord.Starting:
                var started = processes.AutomationExcelsSince(record.MarkedAtUtc);
                if (started.Count > 0)
                {
                    var data = RecordData(app);
                    data["processes"] = new JsonArray(started.Select(facts => (JsonNode)ExcelJson(facts)).ToArray());
                    return (Fail(ToolFailure.Conflict, "UNKNOWN_EXCEL",
                        $"An earlier start of {app.RelativePath} stopped while starting Excel, and an Excel started from COM since then is running " +
                        $"(PID {string.Join(", ", started.Select(facts => facts.Pid))}). Ask the person to check it in Task Manager and close it, then start again.",
                        data), null);
                }

                break;
            case OwnershipRecord.Owned:
                var lookup = processes.Find(record.Process!);
                if (lookup.Presence == ProcessPresence.Unknown)
                {
                    return (Leftover(app, record.Process!, "could not be checked"), null);
                }

                if (lookup.Presence == ProcessPresence.Running)
                {
                    using var leftover = lookup.Process!;
                    // Started from COM, as at ownership, so a rewritten record cannot end a daily Excel (人の判断, 2026-11-04).
                    if (leftover.Facts.CommandLine is not { } commandLine)
                    {
                        return (Leftover(app, record.Process!, "has a command line that could not be read"), null);
                    }

                    if (!commandLine.Contains("/automation", StringComparison.OrdinalIgnoreCase))
                    {
                        return (Leftover(app, record.Process!, "has a command line without /automation"), null);
                    }

                    (var failure, cleanup) = await CleanUp(app, record.Process!, leftover, deadline);
                    if (_unusable && failure is not { Item2.Code: ToolFailure.StateUnknown })
                    {
                        // A call to the leftover is still pending, so this server starts no session ("状態と結果不明").
                        if (cleanup is not null)
                        {
                            appLock.Write(completed);
                        }

                        return (CleanupPending(app, record.Process!, cleanup), null);
                    }

                    if (failure is not null)
                    {
                        return (failure, null);
                    }
                }

                break;
        }

        appLock.Write(completed);
        return (null, cleanup);
    }

    // docs/architecture.md "残った専用のExcelの片付け" 3-5, after the lock and the identity check (1-2). Every call to
    // the leftover runs on a throwaway STA with the start's time limit; a call that does not return is pending.
    private async Task<((JsonObject?, ToolFailure?)? Failure, JsonObject? Cleanup)> CleanUp(
        AppFolder app, RecordedProcess recorded, IOwnedProcess leftover, Stopwatch deadline)
    {
        var worker = new StaWorker();
        var connect = worker.Run(() => connector.Connect(recorded.Pid));
        if (!await Completes(connect, Within(deadline, timeouts.Sync)))
        {
            return (CheckTimedOut(app, recorded), null);
        }

        if (connect.IsFaulted || connect.Result is LeftoverWindow.NotReachable)
        {
            worker.Dispose();
            var why = connect.IsFaulted ? connect.Exception!.InnerException!.Message : ((LeftoverWindow.NotReachable)connect.Result).Reason;
            return (Leftover(app, recorded, $"could not be reached ({why})"), null);
        }

        if (connect.Result is LeftoverWindow.NoWorkbookWindow)
        {
            // No workbook window, so nothing to lose; the lock and the identity are checked (人の判断, 2026-11-04).
            worker.Dispose();
            leftover.Terminate();
            return Ended(app, recorded, leftover.WaitForExit(timeouts.TerminateExit), "terminate");
        }

        var excel = ((LeftoverWindow.Connected)connect.Result).Excel;
        var listing = worker.Run(excel.OpenWorkbookPaths);
        if (!await Completes(listing, Within(deadline, timeouts.Sync)))
        {
            return (CheckTimedOut(app, recorded), null);
        }

        var outside = listing.IsFaulted ? null : listing.Result.Where(path => path is null || !IsInside(path, app.FullPath)).ToList();
        if (outside is null || outside.Count > 0)
        {
            if (!await Completes(worker.Run(excel.ReleaseReferences), timeouts.Sync))
            {
                MarkUnusable();
            }
            else
            {
                worker.Dispose();
            }

            var why = outside is null ? "its open workbooks could not be listed" : "a workbook outside the app folder or without a path is open";
            return (Leftover(app, recorded, $"was not closed because {why}"), null);
        }

        // Checked: only app workbooks are open, so a close or quit that does not end it in time ends it by handle.
        var quit = worker.Run(excel.CloseAllAndQuit);
        if (await Completes(quit, Within(deadline, timeouts.Sync)) && !quit.IsFaulted
            && leftover.WaitForExit(Within(deadline, timeouts.CleanupQuitExit)))
        {
            worker.Dispose();
            return Ended(app, recorded, ended: true, "quit");
        }

        leftover.Terminate();
        var ended = leftover.WaitForExit(timeouts.TerminateExit);
        if (await Completes(quit, timeouts.PendingReturn))
        {
            worker.Dispose();
        }
        else
        {
            MarkUnusable();
        }

        return Ended(app, recorded, ended, "terminate");
    }

    private ((JsonObject?, ToolFailure?)? Failure, JsonObject? Cleanup) Ended(AppFolder app, RecordedProcess recorded, bool ended, string endedBy)
    {
        if (ended)
        {
            return (null, new JsonObject { ["excel"] = RecordedJson(recorded), ["ended_by"] = endedBy });
        }

        var data = RecordData(app);
        data["excel"] = RecordedJson(recorded);
        return (Fail(ToolFailure.Conflict, "EXIT_NOT_CONFIRMED",
            $"The dedicated Excel of an earlier session (PID {recorded.Pid}) was asked to end but its end could not be confirmed. " +
            "Ask the person to check it in Task Manager, then start again.", data), null);
    }

    // The check did not finish within the start's limit: the leftover is left untouched ("残った専用のExcelの片付け").
    private (JsonObject?, ToolFailure?) CheckTimedOut(AppFolder app, RecordedProcess recorded)
    {
        MarkUnusable();
        var extra = RecordData(app);
        extra["excel"] = RecordedJson(recorded);
        return Unknown("TIMEOUT", excelEnded: false, restartable: false, null,
            $"Checking the dedicated Excel of an earlier session (PID {recorded.Pid}) did not finish within the time limit, so it was left untouched. " +
            "Close this conversation and start the session in a new one.", extra);
    }

    private (JsonObject?, ToolFailure?) CleanupPending(AppFolder app, RecordedProcess recorded, JsonObject? cleanup)
    {
        var extra = RecordData(app);
        extra["excel"] = RecordedJson(recorded);
        if (cleanup is not null)
        {
            extra["cleanup"] = cleanup;
        }

        var outcome = cleanup is null
            ? "It was left running."
            : "It was ended, and the changes in it since its last save are lost.";
        return Unknown("TIMEOUT", excelEnded: cleanup is not null, restartable: false, null,
            $"A call to the dedicated Excel of an earlier session (PID {recorded.Pid}) did not return within the time limit. {outcome} " +
            "Close this conversation and start the session in a new one.", extra);
    }

    private static (JsonObject?, ToolFailure?) Leftover(AppFolder app, RecordedProcess recorded, string state)
    {
        var data = RecordData(app);
        data["excel"] = RecordedJson(recorded);
        return Fail(ToolFailure.Conflict, "LEFTOVER_EXCEL",
            $"The dedicated Excel of an earlier session (PID {recorded.Pid}) {state} and was left untouched. " +
            "Ask the person to check it in Task Manager, save what they need, and close it, then start again.",
            data);
    }

    private static JsonObject RecordedJson(RecordedProcess recorded) => new()
    {
        ["pid"] = recorded.Pid,
        ["started_at_utc"] = recorded.StartedAtUtc.ToString("O"),
    };

    // A lock held while a call is pending stays until the process ends ("状態と結果不明").
    private void ReleaseLock(AppLock appLock)
    {
        if (_unusable)
        {
            KeepLock(appLock);
        }
        else
        {
            appLock.Dispose();
        }
    }

    // docs/architecture.md "開始の順序" 4: open the file for writing for a moment, letting others read and write.
    private static (JsonObject?, ToolFailure?)? CheckWorkbookFile(AppFolder app)
    {
        if (File.GetAttributes(app.WorkbookPath).HasFlag(FileAttributes.ReadOnly))
        {
            return Fail(ToolFailure.Conflict, "WORKBOOK_FILE_READ_ONLY",
                $"{app.WorkbookName} has the read-only file attribute. Ask the person to clear it, then start again.");
        }

        try
        {
            using var probe = new FileStream(app.WorkbookPath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            return null;
        }
        catch (IOException e) when (e.HResult == unchecked((int)0x80070020))
        {
            return Fail(ToolFailure.Conflict, "WORKBOOK_IN_USE",
                $"{app.WorkbookName} is open for editing in another program, such as Excel. Ask the person to close it there, then start again.");
        }
    }

    // The results of the last session go once no leftover Excel can still be writing them ("開始の順序" 3).
    private static void RemoveResults(AppFolder app)
    {
        var results = Path.Combine(app.FullPath, ".xlwright", "results");
        if (Directory.Exists(results))
        {
            Directory.Delete(results, recursive: true);
        }
    }

    private static JsonObject RecordData(AppFolder app) => new() { ["record"] = AppLock.RelativePath(app) };

    private static (JsonObject?, ToolFailure?) RecordInvalid(AppFolder app, string why) =>
        Fail(ToolFailure.Conflict, "RECORD_INVALID",
            $"{why} Ask the person to check {AppLock.RelativePath(app)}; when it is empty, delete it, then start again.",
            RecordData(app));

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

    // A step of the start's cleanup: its own limit from the table, cut short by what is left of the start.
    private TimeSpan Within(Stopwatch startElapsed, TimeSpan step)
    {
        var left = Remaining(startElapsed, timeouts.Start);
        return step < left ? step : left;
    }

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
