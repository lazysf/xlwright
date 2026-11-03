using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>vba_sync and macro_run (docs/api.md, stage 1).</summary>
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
                "(UTF-8). A module that is not in the source is removed from the workbook. Before it syncs, xlwright writes " +
                "src/modules/XlWright.bas; do not edit that file. It does not save. Takes no arguments.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, SyncOutput),
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

    private async Task<CallToolResult> Run(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("macro_run", context);
        var (text, args, seconds) = ("", (IReadOnlyList<object>)[], DefaultTimeoutSeconds);
        var invalid = arguments.RejectUnknown("macro", "args", "timeout_seconds")
            ?? arguments.RequireString("macro", out text)
            ?? Arguments(context.Params!.Arguments!, out args)
            ?? Timeout(context.Params!.Arguments!, out seconds);
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

    private static string? Timeout(IDictionary<string, JsonElement> arguments, out int seconds)
    {
        seconds = DefaultTimeoutSeconds;
        if (!arguments.TryGetValue("timeout_seconds", out var element))
        {
            return null;
        }

        return element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out seconds) && seconds is >= 1 and <= MaxTimeoutSeconds
            ? null
            : $"macro_run: timeout_seconds must be a whole number from 1 to {MaxTimeoutSeconds}.";
    }

    private static readonly string SyncResult = """
        "modules":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"type":{"enum":["standard","class"]}},"required":["name","type"]}},
         "removed":{"type":"array","items":{"type":"string"}}
        """;

    private static readonly string SyncOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},{{{SyncResult}}} },"required":["session","modules","removed"]}
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
