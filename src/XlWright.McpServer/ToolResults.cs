using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using XlWright.Core;

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

    public static CallToolResult Success(JsonObject structured) =>
        Success(JsonSerializer.SerializeToElement(structured));

    // Success adds "session" to the result fields; failure reports the session after the operation.
    public static CallToolResult From(SessionOutcome outcome)
    {
        var state = SessionState.From(outcome.Session);
        if (outcome.Failure is not null)
        {
            return Failure(outcome.Failure, state);
        }

        var result = new JsonObject { ["session"] = state.ToJson() };
        foreach (var (name, value) in outcome.Result!)
        {
            result[name] = value?.DeepClone();
        }

        return Success(result);
    }

    public static CallToolResult InvalidRequest(string message) => InvalidRequest(message, SessionState.Closed);

    public static CallToolResult InvalidRequest(string message, SessionState session) =>
        Failure(new ToolFailure(ToolFailure.InvalidRequest, null, message), session);

    // One line of JSON; detail and data are left out when the tool's contract defines none.
    public static CallToolResult Failure(ToolFailure failure, SessionState session)
    {
        var error = new JsonObject { ["code"] = failure.Code };
        if (failure.Detail is not null)
        {
            error["detail"] = failure.Detail;
        }

        error["message"] = failure.Message;
        if (failure.Data is not null)
        {
            error["data"] = failure.Data.DeepClone();
        }

        var body = new JsonObject { ["error"] = error, ["session"] = session.ToJson() };
        return new CallToolResult
        {
            IsError = true,
            Content = [new TextContentBlock { Text = body.ToJsonString() }],
        };
    }
}

/// <summary>
/// The "session" object in every stage 1 result: state (closed or ready) and busy.
/// </summary>
internal sealed record SessionState(string State, bool Busy)
{
    public static SessionState Closed { get; } = new("closed", false);

    public static SessionState From(SessionSnapshot snapshot) => new(snapshot.State, snapshot.Busy);

    public JsonObject ToJson() => new() { ["state"] = State, ["busy"] = Busy };
}
