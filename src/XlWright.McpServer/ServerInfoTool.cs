using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

[McpServerToolType]
public sealed class ServerInfoTool
{
    [McpServerTool(
        Name = "server_info",
        ReadOnly = true,
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        UseStructuredContent = true,
        OutputSchemaType = typeof(ServerIdentity))]
    [Description("Returns the server name, version, and build id. Use it to check which build is running. Takes no arguments.")]
    public static CallToolResult ServerInfo(RequestContext<CallToolRequestParams> context)
    {
        // The SDK ignores unknown arguments, so the tool rejects them itself (docs/api.md).
        var arguments = context.Params?.Arguments;
        if (arguments is { Count: > 0 })
        {
            return ToolResults.InvalidRequest(
                $"server_info takes no arguments. Call it again without: {string.Join(", ", arguments.Keys)}.");
        }

        return ToolResults.Success(JsonSerializer.SerializeToElement(ServerIdentity.Current));
    }
}
