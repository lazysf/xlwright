using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;

namespace XlWright.McpServer;

/// <summary>
/// Builds tool results in the shape that docs/api.md "成功と失敗" defines.
/// </summary>
internal static class ToolResults
{
    public static CallToolResult Success(JsonElement structured) => new()
    {
        StructuredContent = structured,
        Content = [new TextContentBlock { Text = structured.GetRawText() }],
    };

    // There is no session before stage 1, so failures report a closed, idle session.
    public static CallToolResult InvalidRequest(string message)
    {
        var body = new JsonObject
        {
            ["error"] = new JsonObject
            {
                ["code"] = "INVALID_REQUEST",
                ["message"] = message,
            },
            ["session"] = new JsonObject
            {
                ["state"] = "closed",
                ["busy"] = false,
            },
        };
        return new CallToolResult
        {
            IsError = true,
            Content = [new TextContentBlock { Text = body.ToJsonString() }],
        };
    }
}
