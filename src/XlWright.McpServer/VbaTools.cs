using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>vba_sync and macro_run (docs/api.md, stage 1), vba_get, vba_check, and test_run (stage 3), and vba_edit and vba_extract (stage 4).</summary>
internal sealed class VbaTools(ExcelSession? session)
{
    private const int DefaultTimeoutSeconds = 300;
    private const int MaxTimeoutSeconds = 1200;

    public IEnumerable<McpServerTool> Create() =>
    [
        ToolDefinition.Create(Sync, new()
        {
            Name = "vba_sync",
            Description = "Replaces every standard and class module of the main workbook with src/modules/*.bas and src/classes/*.cls " +
                "(UTF-8). A module that is not in the source is removed from the workbook. It also replaces the code of each sheet " +
                "and workbook module that has a file src/workbook/<CodeName>.cls; the others are left as they are. Before it syncs, " +
                "xlwright writes src/modules/XlWright.bas; do not edit that file. It does not save. Takes no arguments.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, SyncOutput),
        ToolDefinition.Create(Get, new()
        {
            Name = "vba_get",
            Description = "Lists the VBA modules of the main workbook as they are now, without syncing: name, type " +
                "(standard, class, document, form), and number of lines. A sheet module also has the sheet name; its module " +
                "name is the CodeName to use for src/workbook/<CodeName>.cls. It also lists the references of the VBA project: " +
                "name, guid, major, minor, builtin, and broken. Takes no arguments.",
            ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, GetOutput),
        ToolDefinition.Create(Edit, new()
        {
            Name = "vba_edit",
            Description = "Adds or removes a reference of the main workbook's VBA project. action is add_reference or remove_reference. " +
                "Pass guid as {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}; add_reference also takes major and minor. The values of " +
                "references from vba_get can be used as they are. Adding a reference that is already there changes nothing " +
                "(changed is false). The built-in references (VBA, Excel) cannot be removed. It does not save.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, EditInput, EditOutput),
        ToolDefinition.Create(Extract, new()
        {
            Name = "vba_extract",
            Description = "Extracts the VBA source of an existing workbook into src/: standard modules to src/modules/*.bas, class " +
                "modules to src/classes/*.cls, and sheet and workbook modules to src/workbook/<CodeName>.cls, in UTF-8. Call it " +
                "before vba_sync on an existing workbook, because vba_sync removes the modules that are not in src/. workbook is a " +
                "path inside the app folder; leave it out for the main workbook. It reads the saved file, through a copy opened " +
                "with macros disabled, so no workbook code runs and unsaved changes are not extracted. UserForms are not " +
                "extracted (skipped). It stops if src/ already has source files, and never overwrites them.",
            ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = false,
        }, ExtractInput, ExtractOutput),
        ToolDefinition.Create(Check, new()
        {
            Name = "vba_check",
            Description = "Syncs src/ as vba_sync does, then compiles the VBA project of the main workbook. A compile error is a " +
                "result, not a failure: compiled is false and error has the message, and the module and line when the editor " +
                "gives them. On a compile error the dialog and the VBA editor show on the person's screen for a moment; xlwright " +
                "closes the dialog and hides the editor. It does not save. Takes no arguments.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, CheckOutput),
        ToolDefinition.Create(Run, new()
        {
            Name = "macro_run",
            Description = "Syncs src/ as vba_sync does, then runs a Public Sub or Function of a standard module in the main workbook " +
                "and returns its return value. Pass macro as <module>.<procedure>, such as Calc.Twice, without a workbook name. " +
                "It waits until the macro ends or timeout_seconds passes. A run-time error comes back as MACRO_ERROR. " +
                "If the macro stops at a dialog (a compile error, an uncaught error, a MsgBox) or runs past the limit, " +
                "xlwright ends the dedicated Excel and the changes since the last save are lost. It does not save.",
            ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true,
        }, RunInput, RunOutput),
        ToolDefinition.Create(TestRun, new()
        {
            Name = "test_run",
            Description = "Syncs src/ as vba_sync does and saves the main workbook, including the changes made before this call. " +
                "Then it copies the app folder (without src/, tests/, and folders that start with a dot) to .xlwright/test/, " +
                "imports tests/ and the assertion module XlWrightAssert into the copy, runs each test, and opens the main workbook " +
                "again; the tests' changes stay in the copy, which is removed. A test is a Sub without arguments, not Private, " +
                "whose name starts with Test, in tests/*.bas. Assert with XlWrightAssert.AreEqual expected, actual[, message], " +
                "XlWrightAssert.IsTrue condition[, message], and XlWrightAssert.Fail message. A failing test is a result, not a " +
                "failure. xlwright does not stop a test from changing files outside the app folder, or the main workbook's folder " +
                "through a full path. If a test stops at a dialog or runs past timeout_seconds (the whole call), xlwright ends " +
                "the dedicated Excel; the main workbook was saved before the tests.",
            ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = true,
        }, TestRunInput, TestRunOutput),
    ];

    private async Task<CallToolResult> Sync(RequestContext<CallToolRequestParams> context)
    {
        var invalid = ToolArguments.From("vba_sync", context).RejectUnknown();
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.SyncAsync());
    }

    private async Task<CallToolResult> Get(RequestContext<CallToolRequestParams> context)
    {
        var invalid = ToolArguments.From("vba_get", context).RejectUnknown();
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.ListModulesAsync());
    }

    private async Task<CallToolResult> Edit(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("vba_edit", context);
        var values = context.Params!.Arguments ?? new Dictionary<string, JsonElement>();
        var (action, guid) = ("", "");
        var invalid = arguments.RejectUnknown("action", "guid", "major", "minor")
            ?? arguments.RequireString("action", out action)
            ?? (action is "add_reference" or "remove_reference" ? null : "vba_edit: action must be add_reference or remove_reference.")
            ?? arguments.RequireString("guid", out guid)
            ?? (Guid.TryParseExact(guid, "B", out _) ? null : "vba_edit: guid must be {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}, as vba_get gives it.");
        var (major, minor) = (0, 0);
        if (invalid is null && action == "add_reference")
        {
            invalid = Version(values, "major", out major) ?? Version(values, "minor", out minor);
        }
        else if (invalid is null && (values.ContainsKey("major") || values.ContainsKey("minor")))
        {
            invalid = "vba_edit: major and minor are only for add_reference. Call it again without them.";
        }

        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        return session is null
            ? NoSession()
            : ToolResults.From(await session.EditReferenceAsync(new ReferenceEdit(action == "add_reference", guid, major, minor)));
    }

    private async Task<CallToolResult> Extract(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("vba_extract", context);
        var values = context.Params!.Arguments ?? new Dictionary<string, JsonElement>();
        string? workbook = null;
        var invalid = arguments.RejectUnknown("workbook");
        if (invalid is null && values.ContainsKey("workbook"))
        {
            invalid = arguments.RequireString("workbook", out var text)
                ?? (text.Length == 0 ? "vba_extract: workbook must not be empty; leave it out for the main workbook." : null);
            workbook = text;
        }

        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.ExtractAsync(workbook));
    }

    private static string? Version(IDictionary<string, JsonElement> arguments, string name, out int value)
    {
        value = 0;
        return arguments.TryGetValue(name, out var element) && element.ValueKind == JsonValueKind.Number &&
            element.TryGetInt32(out value) && value >= 0
            ? null
            : $"vba_edit: add_reference requires {name} (a whole number from 0), as vba_get gives it.";
    }

    private async Task<CallToolResult> Check(RequestContext<CallToolRequestParams> context)
    {
        var invalid = ToolArguments.From("vba_check", context).RejectUnknown();
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.CheckAsync());
    }

    private async Task<CallToolResult> TestRun(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("test_run", context);
        var (tests, seconds) = ((IReadOnlyList<string>?)null, DefaultTimeoutSeconds);
        var invalid = arguments.RejectUnknown("tests", "timeout_seconds")
            ?? Tests(context.Params!.Arguments ?? new Dictionary<string, JsonElement>(), out tests)
            ?? Timeout("test_run", context.Params!.Arguments ?? new Dictionary<string, JsonElement>(), out seconds);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        if (tests?.FirstOrDefault(name => MacroName.Parse(name).Failure is not null) is { } wrong)
        {
            return ToolResults.Failure(new ToolFailure(ToolFailure.InvalidRequest, "INVALID_TEST_NAME",
                $"tests has \"{wrong}\", which is not <module>.<procedure>, such as CalcTests.TestAdd. Leave out the workbook name."), Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.TestRunAsync(tests, TimeSpan.FromSeconds(seconds)));
    }

    private async Task<CallToolResult> Run(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("macro_run", context);
        var (text, args, seconds) = ("", (IReadOnlyList<object>)[], DefaultTimeoutSeconds);
        var invalid = arguments.RejectUnknown("macro", "args", "timeout_seconds")
            ?? arguments.RequireString("macro", out text)
            ?? Arguments(context.Params!.Arguments!, out args)
            ?? Timeout("macro_run", context.Params!.Arguments!, out seconds);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        var (macro, failure) = MacroName.Parse(text);
        if (failure is not null)
        {
            return ToolResults.Failure(failure, Current());
        }

        return session is null ? NoSession() : ToolResults.From(await session.RunAsync(macro!, args, TimeSpan.FromSeconds(seconds)));
    }

    private SessionState Current() => session is null ? SessionState.Closed : SessionState.From(session.Status());

    private static CallToolResult NoSession() => ToolResults.Failure(
        new ToolFailure(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start."), SessionState.Closed);

    private static string? Arguments(IDictionary<string, JsonElement> arguments, out IReadOnlyList<object> args)
    {
        args = [];
        if (!arguments.TryGetValue("args", out var element))
        {
            return null;
        }

        const string Message = "macro_run: args must be an array of at most 30 numbers, strings, or booleans.";
        if (element.ValueKind != JsonValueKind.Array || element.GetArrayLength() > MacroCall.MaxArguments)
        {
            return Message;
        }

        var list = new List<object>();
        foreach (var item in element.EnumerateArray())
        {
            object? value = item.ValueKind switch
            {
                JsonValueKind.Number => item.GetDouble(),
                JsonValueKind.String => item.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                _ => null,
            };
            if (value is null)
            {
                return Message;
            }

            list.Add(value);
        }

        args = list;
        return null;
    }

    private static string? Timeout(string tool, IDictionary<string, JsonElement> arguments, out int seconds)
    {
        seconds = DefaultTimeoutSeconds;
        if (!arguments.TryGetValue("timeout_seconds", out var element))
        {
            return null;
        }

        return element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out seconds) && seconds is >= 1 and <= MaxTimeoutSeconds
            ? null
            : $"{tool}: timeout_seconds must be a whole number from 1 to {MaxTimeoutSeconds}.";
    }

    private static string? Tests(IDictionary<string, JsonElement> arguments, out IReadOnlyList<string>? tests)
    {
        tests = null;
        if (!arguments.TryGetValue("tests", out var element))
        {
            return null;
        }

        if (element.ValueKind != JsonValueKind.Array || element.GetArrayLength() == 0 ||
            element.EnumerateArray().Any(item => item.ValueKind != JsonValueKind.String))
        {
            return "test_run: tests must be a non-empty array of <module>.<procedure> strings; leave it out to run every test.";
        }

        tests = element.EnumerateArray().Select(item => item.GetString()!).ToList();
        return null;
    }

    private static readonly string SyncResult = """
        "modules":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"type":{"enum":["standard","class","document"]}},"required":["name","type"]}},
         "removed":{"type":"array","items":{"type":"string"}}
        """;

    private static readonly string SyncOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},{{{SyncResult}}} },"required":["session","modules","removed"]}
        """;

    private static readonly string GetOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},
         "modules":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},
          "type":{"enum":["standard","class","document","form"]},"lines":{"type":"integer"},"sheet":{"type":"string"}},
          "required":["name","type","lines"]}},
         "references":{{{References}}} },
         "required":["session","modules","references"]}
        """;

    private const string References = """
        {"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"guid":{"type":"string"},
         "major":{"type":"integer"},"minor":{"type":"integer"},"builtin":{"type":"boolean"},"broken":{"type":"boolean"}},
         "required":["guid","major","minor","builtin","broken"]}}
        """;

    private const string EditInput = """
        {"type":"object","properties":{
         "action":{"enum":["add_reference","remove_reference"],"description":"What to do with the reference."},
         "guid":{"type":"string","description":"The GUID of the type library, such as {420B2830-E718-11CF-893D-00A0C9054228}."},
         "major":{"type":"integer","minimum":0,"description":"The major version, for add_reference only."},
         "minor":{"type":"integer","minimum":0,"description":"The minor version, for add_reference only."}},
         "required":["action","guid"],"additionalProperties":false}
        """;

    private const string ExtractInput = """
        {"type":"object","properties":{
         "workbook":{"type":"string","minLength":1,"description":"The workbook to extract from, as a path inside the app folder, such as data/Legacy.xlsm. Leave it out for the main workbook."}},
         "additionalProperties":false}
        """;

    private static readonly string ExtractOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"workbook":{"type":"string"},
         "files":{"type":"array","items":{"type":"object","properties":{"file":{"type":"string"},"type":{"enum":["standard","class","document"]}},"required":["file","type"]}},
         "skipped":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"type":{"type":"string"}},"required":["name","type"]}} },
         "required":["session","workbook","files","skipped"]}
        """;

    private static readonly string EditOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"changed":{"type":"boolean"},
         "references":{{{References}}} },
         "required":["session","changed","references"]}
        """;

    private static readonly string CheckOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"compiled":{"type":"boolean"},
         "error":{"type":"object","properties":{"message":{"type":"string"},"module":{"type":"string"},"line":{"type":"integer"}},"required":["message"]},
         "sync":{"type":"object","properties":{ {{{SyncResult}}} },"required":["modules","removed"]}},
         "required":["session","compiled","sync"]}
        """;

    private const string TestRunInput = """
        {"type":"object","properties":{
         "tests":{"type":"array","minItems":1,"items":{"type":"string"},"description":"The tests to run, as <module>.<procedure>, such as CalcTests.TestAdd. Leave it out to run every test."},
         "timeout_seconds":{"type":"integer","minimum":1,"maximum":1200,"description":"How long the whole call may take. The default is 300."}},
         "additionalProperties":false}
        """;

    private static readonly string TestRunOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},
         "summary":{"type":"object","properties":{"passed":{"type":"integer"},"failed":{"type":"integer"},"errors":{"type":"integer"}},"required":["passed","failed","errors"]},
         "tests":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"result":{"enum":["passed","failed","error"]},
          "message":{"type":"string"},"number":{"type":"integer"},"description":{"type":"string"},"source":{"type":"string"},"elapsed_ms":{"type":"integer"}},
          "required":["name","result"]}},
         "elapsed_ms":{"type":"integer"},"result_file":{"type":"string"},"leftover_copy":{"type":"string"},
         "sync":{"type":"object","properties":{ {{{SyncResult}}} },"required":["modules","removed"]}},
         "required":["session","summary","tests","elapsed_ms","sync"]}
        """;

    private const string RunInput = """
        {"type":"object","properties":{
         "macro":{"type":"string","description":"<module>.<procedure> of a Public Sub or Function in a standard module, such as Calc.Twice."},
         "args":{"type":"array","maxItems":30,"items":{"type":["number","string","boolean"]},"description":"The arguments, in order."},
         "timeout_seconds":{"type":"integer","minimum":1,"maximum":1200,"description":"How long the macro may run. The default is 300."}},
         "required":["macro"],"additionalProperties":false}
        """;

    private static readonly string RunOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"macro":{"type":"string"},
         "return":{"type":["number","string","boolean","null"]},"return_type":{"type":"string"},"result_file":{"type":"string"},
         "elapsed_ms":{"type":"integer"},
         "sync":{"type":"object","properties":{ {{{SyncResult}}} },"required":["modules","removed"]}},
         "required":["session","macro","return","elapsed_ms","sync"]}
        """;
}
