using System.Text.Json;
using ModelContextProtocol.Protocol;
using XlWright.Core;

namespace XlWright.McpServer;

// The response shapes in docs/api.md. ADR 0005 explains why failures are text only.
internal static class ToolResponses
{
    private static readonly JsonSerializerOptions Options = new() { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower };

    internal static CallToolResult Success(OperationResult result)
    {
        var content = JsonSerializer.SerializeToElement(result, result.GetType(), Options);
        return new() { Content = [new TextContentBlock { Text = content.GetRawText() }], StructuredContent = content, IsError = false };
    }

    // Returned rather than thrown as McpException: the SDK prefixes that message, and the text must be only the JSON.
    internal static CallToolResult Failure(SessionOperationException failure) => new()
    {
        Content = [new TextContentBlock { Text = JsonSerializer.Serialize(new { error = failure.Error, session = failure.Session }, Options) }],
        IsError = true
    };
}
