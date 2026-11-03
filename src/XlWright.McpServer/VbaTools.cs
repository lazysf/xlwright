using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>vba_sync (docs/api.md, stage 1).</summary>
internal sealed class VbaTools(ExcelSession? session)
{
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
    ];

    private async Task<CallToolResult> Sync(RequestContext<CallToolRequestParams> context)
    {
        var current = session is null ? SessionState.Closed : SessionState.From(session.Status());
        var invalid = ToolArguments.From("vba_sync", context).RejectUnknown();
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, current);
        }

        if (session is null)
        {
            return ToolResults.Failure(
                new ToolFailure(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start."),
                SessionState.Closed);
        }

        return ToolResults.From(await session.SyncAsync());
    }

    private static readonly string SyncOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},
         "modules":{"type":"array","items":{"type":"object","properties":{"name":{"type":"string"},"type":{"enum":["standard","class"]}},"required":["name","type"]}},
         "removed":{"type":"array","items":{"type":"string"}} },"required":["session","modules","removed"]}
        """;
}
