using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>
/// session_start, session_status, session_end, and workbook_save (docs/api.md, stage 1). With --no-excel there is no
/// ExcelSession: session_start stops after the app folder checks, and there is never a session.
/// </summary>
internal sealed class SessionTools(ServerOptions options, ExcelSession? session)
{
    public IEnumerable<McpServerTool> Create() =>
    [
        ToolDefinition.Create(Start, new()
        {
            Name = "session_start",
            Description = "Starts a session for one app folder: starts a dedicated hidden Excel and opens the app's main workbook " +
                "(the only .xlsm or .xlsx directly in the folder). Pass app as a path relative to the repository, or absolute. " +
                "A conversation has at most one session; end it with session_end before you report.",
            ReadOnly = false, Destructive = false, Idempotent = false, OpenWorld = false,
        }, StartInput, StartOutput),
        ToolDefinition.Create(Status, new()
        {
            Name = "session_status",
            Description = "Returns the session state (closed or ready), whether an operation is running, and how long it has run. " +
                "Answers even while a macro runs. During test_run it also gives the running test. Takes no arguments.",
            ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, StatusOutput),
        ToolDefinition.Create(End, new()
        {
            Name = "session_end",
            Description = "Ends the session: saves the main workbook if save is true (otherwise discards the changes since the last save), " +
                "closes it, and ends the dedicated Excel. Call it before you report.",
            ReadOnly = false, Destructive = true, Idempotent = false, OpenWorld = false,
        }, EndInput, EndOutput),
        ToolDefinition.Create(Save, new()
        {
            Name = "workbook_save",
            Description = "Saves the main workbook over its file. The session stays open. Takes no arguments.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, ToolDefinition.NoInput, SaveOutput),
    ];

    private async Task<CallToolResult> Start(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("session_start", context);
        var app = "";
        var invalid = arguments.RejectUnknown("app") ?? arguments.RequireString("app", out app);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        var (folder, failure) = AppFolder.Resolve(options.Project, app);
        if (failure is not null)
        {
            return ToolResults.Failure(failure, Current());
        }

        if (session is null)
        {
            return ToolResults.Failure(new ToolFailure(ToolFailure.InvalidRequest, "NOT_IMPLEMENTED",
                "This server runs with --no-excel, so it does not start Excel. The app folder passed the checks."), SessionState.Closed);
        }

        return ToolResults.From(await session.StartAsync(folder!));
    }

    private Task<CallToolResult> Status(RequestContext<CallToolRequestParams> context)
    {
        var invalid = ToolArguments.From("session_status", context).RejectUnknown();
        if (invalid is not null)
        {
            return Task.FromResult(ToolResults.InvalidRequest(invalid, Current()));
        }

        var snapshot = session?.Status() ?? new SessionSnapshot("closed", false);
        var result = new JsonObject { ["session"] = SessionState.From(snapshot).ToJson() };
        if (snapshot.App is not null)
        {
            result["app"] = snapshot.App;
            result["workbook"] = snapshot.Workbook;
        }

        if (snapshot.Operation is not null)
        {
            var operation = new JsonObject
            {
                ["tool"] = snapshot.Operation,
                ["elapsed_ms"] = (long)snapshot.Elapsed.TotalMilliseconds,
            };
            if (snapshot.Test is { } test)
            {
                operation["test"] = new JsonObject { ["name"] = test.Name, ["index"] = test.Index, ["count"] = test.Count };
            }

            result["operation"] = operation;
        }

        return Task.FromResult(ToolResults.Success(result));
    }

    private async Task<CallToolResult> End(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("session_end", context);
        var save = false;
        var invalid = arguments.RejectUnknown("save") ?? arguments.RequireBoolean("save", out save);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        if (session is null)
        {
            return ToolResults.Failure(
                new ToolFailure(ToolFailure.Conflict, "NO_SESSION", "There is no session to end. Start one with session_start."),
                SessionState.Closed);
        }

        return ToolResults.From(await session.EndAsync(save));
    }

    private async Task<CallToolResult> Save(RequestContext<CallToolRequestParams> context)
    {
        var invalid = ToolArguments.From("workbook_save", context).RejectUnknown();
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        if (session is null)
        {
            return ToolResults.Failure(
                new ToolFailure(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start."),
                SessionState.Closed);
        }

        return ToolResults.From(await session.SaveAsync());
    }

    private SessionState Current() => session is null ? SessionState.Closed : SessionState.From(session.Status());

    private const string StartInput = """
        {"type":"object","properties":{"app":{"type":"string","description":"The app folder, relative to the repository or absolute."}},"required":["app"],"additionalProperties":false}
        """;

    private const string EndInput = """
        {"type":"object","properties":{"save":{"type":"boolean","description":"true saves the main workbook before closing; false discards the changes since the last save."}},"required":["save"],"additionalProperties":false}
        """;

    private static readonly string StartOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"app":{"type":"string"},"workbook":{"type":"string"},
         "excel":{"type":"object","properties":{"pid":{"type":"integer"},"started_at_utc":{"type":"string"}},"required":["pid","started_at_utc"]}},
         "required":["session","app","workbook","excel"]}
        """;

    private static readonly string StatusOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"app":{"type":"string"},"workbook":{"type":"string"},
         "operation":{"type":"object","properties":{"tool":{"type":"string"},"elapsed_ms":{"type":"integer"},
          "test":{"type":"object","properties":{"name":{"type":"string"},"index":{"type":"integer"},"count":{"type":"integer"}},"required":["name","index","count"]}},
          "required":["tool","elapsed_ms"]}},
         "required":["session"]}
        """;

    private static readonly string EndOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"saved":{"type":"boolean"}},"required":["session","saved"]}
        """;

    private static readonly string SaveOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"workbook":{"type":"string"}},"required":["session","workbook"]}
        """;
}
