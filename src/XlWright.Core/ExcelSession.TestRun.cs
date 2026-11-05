using System.Diagnostics;
using System.Text.Json.Nodes;

namespace XlWright.Core;

// test_run (docs/api.md "test_run", docs/architecture.md "テスト"): sync and save the main workbook, close it, copy the
// app folder to .xlwright/test/<run-id>/, open the copy with events off, import tests/ and XlWrightAssert, run each
// test through the generated module, close the copy without saving, remove it, and open the main workbook again.
public sealed partial class ExcelSession
{
    // The response keeps room for "session" within the 20,000 characters of docs/api.md "成功と失敗".
    private const int MaxTestResultCharacters = 18_000;

    /// <summary>
    /// Runs the tests of tests/ in a copy of the app folder. <paramref name="only"/> keeps the named tests;
    /// <paramref name="limit"/> is measured over the whole call and checked while the tests run.
    /// </summary>
    public Task<SessionOutcome> TestRunAsync(IReadOnlyList<string>? only, TimeSpan limit) => Run("test_run", NoSession, async () =>
    {
        var active = _active!;
        var elapsed = Stopwatch.StartNew();
        if (MacrosRejected(active) is { } rejected)
        {
            return rejected;
        }

        var (plan, invalid) = VbaSources.PrepareTests(active.App, VbaSources.SystemAnsiEncoding(), only);
        if (invalid is not null)
        {
            return (null, invalid);
        }

        try
        {
            return await TestRunCore(active, plan!, limit, elapsed);
        }
        finally
        {
            lock (_gate)
            {
                _test = null;
            }

            RemoveFolder(Path.GetDirectoryName(plan!.Modules[0].ImportPath)!);
        }
    });

    private async Task<(JsonObject?, ToolFailure?)> TestRunCore(Active active, TestPlan plan, TimeSpan limit, Stopwatch elapsed)
    {
        var (sync, syncFailure) = await SyncCore(active);
        if (syncFailure is not null)
        {
            return (null, syncFailure);
        }

        if (await SaveCore(active, "No test ran. Fix the cause and call test_run again.") is { } saveFailed)
        {
            return saveFailed;
        }

        var closing = active.Worker.Run(active.Excel.CloseMain);
        if (!await Completes(closing, timeouts.Sync))
        {
            return await AfterTimeout(active, closing, "Closing the main workbook for the tests did not finish within the time limit.");
        }

        if (closing.IsFaulted)
        {
            return Fail(ToolFailure.ExcelError, null,
                $"Excel could not close {active.App.WorkbookName} for the tests: {closing.Exception!.InnerException!.Message} No test ran.");
        }

        // From here the main workbook is closed, so every way out opens it again.
        var results = new JsonArray();
        var copy = Path.Combine(active.App.FullPath, ".xlwright", "test", Guid.NewGuid().ToString("N"));
        try
        {
            CopyApp(active.App.FullPath, copy);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            return await StopAndReopen(active, copy, results, Fail(ToolFailure.InternalError, "COPY_FAILED",
                $"The app folder could not be copied for the tests: {e.Message} No test ran."));
        }

        var opening = active.Worker.Run(() => active.Excel.Open(Path.Combine(copy, active.App.WorkbookName!)));
        if (!await Completes(opening, timeouts.Save))
        {
            return await AfterTimeout(active, opening, "Opening the test copy did not finish within the time limit." + Saved, extra: TestData(results, null));
        }

        if (opening.IsFaulted)
        {
            return await StopAndReopen(active, copy, results, Fail(ToolFailure.ExcelError, "OPEN_FAILED",
                $"Excel could not open the test copy: {opening.Exception!.InnerException!.Message} No test ran."));
        }

        var importing = active.Worker.Run(() => active.Excel.ImportModules(plan.Modules));
        if (!await Completes(importing, timeouts.Sync))
        {
            return await AfterTimeout(active, importing, "Importing the tests did not finish within the time limit." + Saved, extra: TestData(results, null));
        }

        if (importing.IsFaulted)
        {
            var error = importing.Exception!.InnerException!;
            return await StopAndReopen(active, copy, results, SyncFailed(
                $"Importing the tests into the copy failed: {error.Message} No test ran.", (error as VbaSyncException)?.Done ?? []));
        }

        // A test that calls what is not written yet stops here, without ending the Excel (docs/api.md "test_runの前のコンパイル").
        // A compile that is not confirmed runs the tests, whose dialog watch still catches an error.
        switch (await CompileCore(active, Saved, TestData(results, null)))
        {
            case Compiled.Error compileError:
                return await StopAndReopen(active, copy, results, Fail(ToolFailure.ExcelError, "COMPILE_ERROR",
                    $"The tests did not compile: {compileError.Value["message"]} No test ran. Fix the code, then call test_run again.",
                    new JsonObject { ["error"] = compileError.Value }));
            case Compiled.Stopped { Failure: (_, { Code: ToolFailure.StateUnknown }) } ended:
                RemoveCopy(active.App, copy);
                return ended.Failure;
            case Compiled.Stopped failed:
                return await StopAndReopen(active, copy, results, failed.Failure);
        }

        for (var i = 0; i < plan.Tests.Count; i++)
        {
            var test = plan.Tests[i];
            lock (_gate)
            {
                _test = new TestProgress(test.ToString(), i + 1, plan.Tests.Count);
            }

            var started = Stopwatch.StartNew();
            var running = active.Worker.Run(() => active.Excel.RunMacro(test, [], []));
            while (!await Completes(running, DialogInterval))
            {
                if (active.Process.FindDialog() is { } dialog)
                {
                    var data = TestData(results, test);
                    data["dialog"] = new JsonObject { ["title"] = dialog.Title, ["text"] = dialog.Text };
                    return await AfterTimeoutAndRemove(active, running, copy, $"{test} stopped at a dialog." + Saved, "DIALOG", data);
                }

                if (elapsed.Elapsed >= limit)
                {
                    return await AfterTimeoutAndRemove(active, running, copy,
                        $"The tests did not finish within {limit.TotalSeconds:0.###} seconds; {test} was running." + Saved, "TIMEOUT", TestData(results, test));
                }
            }

            results.Add(TestResult(test, running, (long)started.Elapsed.TotalMilliseconds));
        }

        lock (_gate)
        {
            _test = null;
        }

        var (leftover, stop) = await CloseCopyAndReopen(active, copy, results);
        if (stop is { } stopped)
        {
            return stopped;
        }

        return (TestRunResult(active.App, results, (long)elapsed.Elapsed.TotalMilliseconds, sync!, leftover), null);
    }

    private const string Saved = " The main workbook was saved before the tests, so the tests changed only the copy.";

    // A failure before the tests ran: close the copy, remove it, and open the main workbook again.
    private async Task<(JsonObject?, ToolFailure?)> StopAndReopen(Active active, string copy, JsonArray results, (JsonObject?, ToolFailure?) failure)
    {
        var (_, stop) = await CloseCopyAndReopen(active, copy, results);
        return stop ?? failure;
    }

    // docs/architecture.md "テスト" 5. A main workbook that does not open again, or opens read-only, ends the Excel
    // ("終了" 2-5) and closes the session, with the results so far.
    private async Task<(string? Leftover, (JsonObject?, ToolFailure?)? Stop)> CloseCopyAndReopen(Active active, string copy, JsonArray results)
    {
        var closing = active.Worker.Run(() => active.Excel.CloseUnder(copy));
        if (!await Completes(closing, timeouts.Sync))
        {
            return (null, await AfterTimeoutAndRemove(active, closing, copy, "Closing the test copy did not finish within the time limit." + Saved, "TIMEOUT", TestData(results, null)));
        }

        var leftover = RemoveCopy(active.App, copy);
        var opening = active.Worker.Run(() => active.Excel.Open(active.App.WorkbookPath!));
        if (!await Completes(opening, timeouts.Save))
        {
            return (leftover, await AfterTimeout(active, opening, $"Opening {active.App.WorkbookName} again did not finish within the time limit." + Saved, extra: TestData(results, null)));
        }

        if (!opening.IsFaulted && !opening.Result)
        {
            return (leftover, null);
        }

        var why = opening.IsFaulted
            ? $"Excel could not open {active.App.WorkbookName} again after the tests: {opening.Exception!.InnerException!.Message}"
            : $"{active.App.WorkbookName} opened read-only after the tests, so another program may have it open.";
        var (_, endFailure) = await EndCore(active, save: false);
        if (endFailure is not null)
        {
            var data = endFailure.Data ?? [];
            data["tests"] = results.DeepClone();
            return (leftover, (null, endFailure with { Message = why + " " + endFailure.Message, Data = data }));
        }

        return (leftover, Fail(ToolFailure.Conflict, "WORKBOOK_NOT_REOPENED",
            why + " The dedicated Excel was ended and the session is closed." + Saved + " Start a new session when the workbook can be edited.",
            TestData(results, null)));
    }

    private async Task<(JsonObject?, ToolFailure?)> AfterTimeoutAndRemove(Active active, Task pending, string copy, string what, string detail, JsonObject data)
    {
        var outcome = await AfterTimeout(active, pending, what, detail, data);
        RemoveCopy(active.App, copy);
        return outcome;
    }

    // Returns the copy's path from the repository when it cannot be removed, such as while a program holds a file.
    private static string? RemoveCopy(AppFolder app, string copy)
    {
        try
        {
            if (Directory.Exists(copy))
            {
                Directory.Delete(copy, recursive: true);
            }

            RemoveIfEmpty(Path.GetDirectoryName(copy)!);
            return null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            return $"{app.RelativePath}/.xlwright/test/{Path.GetFileName(copy)}";
        }
    }

    // Everything but src/, tests/, and what starts with a dot (.xlwright/, .git/), so paths from ThisWorkbook.Path work.
    private static void CopyApp(string app, string copy)
    {
        Directory.CreateDirectory(copy);
        foreach (var entry in Directory.EnumerateFileSystemEntries(app))
        {
            var name = Path.GetFileName(entry);
            if (name.StartsWith('.') || name.Equals("src", StringComparison.OrdinalIgnoreCase) || name.Equals("tests", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            CopyEntry(entry, Path.Combine(copy, name));
        }
    }

    private static void CopyEntry(string source, string target)
    {
        if (!Directory.Exists(source))
        {
            File.Copy(source, target);
            return;
        }

        Directory.CreateDirectory(target);
        foreach (var entry in Directory.EnumerateFileSystemEntries(source))
        {
            CopyEntry(entry, Path.Combine(target, Path.GetFileName(entry)));
        }
    }

    private static JsonObject TestResult(MacroName test, Task<MacroResult> run, long time)
    {
        var result = new JsonObject { ["name"] = test.ToString() };
        if (run.IsFaulted)
        {
            // Excel refused the call itself, such as after an End statement (docs/notes/macro-run.md).
            var error = run.Exception!.InnerException!;
            return Errored(result, error.HResult, error.Message, "", time);
        }

        switch (run.Result)
        {
            case MacroResult.Failed { Source: VbaSources.AssertName } failed:
                result["result"] = "failed";
                result["message"] = failed.Description;
                break;
            case MacroResult.Failed failed:
                return Errored(result, failed.Number, failed.Description, failed.Source, time);
            case MacroResult.NotFound missing:
                return Errored(result, 0, missing.Reason, "", time);
            default:
                result["result"] = "passed";
                break;
        }

        result["elapsed_ms"] = time;
        return result;
    }

    private static JsonObject Errored(JsonObject result, int number, string description, string source, long time)
    {
        result["result"] = "error";
        result["number"] = number;
        result["description"] = description;
        result["source"] = source;
        result["elapsed_ms"] = time;
        return result;
    }

    private static JsonObject TestData(JsonArray results, MacroName? test)
    {
        var data = new JsonObject();
        if (test is not null)
        {
            data["test"] = test.ToString();
        }

        data["tests"] = results.DeepClone();
        return data;
    }

    // Many results go whole into <app>/.xlwright/results/<id>.json; the response keeps the summary and as many of the
    // tests that did not pass as fit (docs/api.md "成功と失敗").
    private static JsonObject TestRunResult(AppFolder app, JsonArray results, long time, JsonObject sync, string? leftover)
    {
        var summary = new JsonObject
        {
            ["passed"] = results.Count(r => r!["result"]!.GetValue<string>() == "passed"),
            ["failed"] = results.Count(r => r!["result"]!.GetValue<string>() == "failed"),
            ["errors"] = results.Count(r => r!["result"]!.GetValue<string>() == "error"),
        };
        JsonObject Build(JsonArray tests)
        {
            var result = new JsonObject { ["summary"] = summary.DeepClone(), ["tests"] = tests, ["elapsed_ms"] = time, ["sync"] = sync.DeepClone() };
            if (leftover is not null)
            {
                result["leftover_copy"] = leftover;
            }

            return result;
        }

        var whole = Build((JsonArray)results.DeepClone());
        if (whole.ToJsonString(JsonText.Options).Length <= MaxTestResultCharacters)
        {
            return whole;
        }

        var id = Guid.NewGuid().ToString("N");
        var folder = Path.Combine(app.FullPath, ".xlwright", "results");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, id + ".json"),
            new JsonObject { ["summary"] = summary.DeepClone(), ["tests"] = results.DeepClone() }.ToJsonString(JsonText.Options));
        var kept = new JsonArray();
        var shortened = Build(kept);
        shortened["result_file"] = $"{app.RelativePath}/.xlwright/results/{id}.json";
        foreach (var result in results.Where(r => r!["result"]!.GetValue<string>() != "passed"))
        {
            kept.Add(result!.DeepClone());
            if (shortened.ToJsonString(JsonText.Options).Length > MaxTestResultCharacters)
            {
                kept.RemoveAt(kept.Count - 1);
                break;
            }
        }

        return shortened;
    }
}
