using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace XlWright.McpServer;

/// <summary>
/// Creates a tool with hand-written input and output schemas (docs/api.md "成功と失敗"), and the schema parts
/// that the stage 1 tools share.
/// </summary>
internal static class ToolDefinition
{
    public const string Session = """
        {"type":"object","properties":{"state":{"enum":["closed","ready"]},"busy":{"type":"boolean"}},"required":["state","busy"],"additionalProperties":false}
        """;

    public const string NoInput = """
        {"type":"object","properties":{},"additionalProperties":false}
        """;

    public static McpServerTool Create(
        Func<RequestContext<CallToolRequestParams>, Task<CallToolResult>> handler,
        McpServerToolCreateOptions toolOptions,
        string input,
        string output)
    {
        toolOptions.UseStructuredContent = true;
        toolOptions.OutputSchema = Parse(output);
        var tool = McpServerTool.Create(handler, toolOptions);
        tool.ProtocolTool.InputSchema = Parse(input);
        return tool;
    }

    private static JsonElement Parse(string schema) => JsonDocument.Parse(schema).RootElement.Clone();
}
