using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>
/// A failure in the shape that docs/api.md "成功と失敗" defines: code, optional detail, message, optional data.
/// </summary>
public sealed record ToolFailure(string Code, string? Detail, string Message, JsonObject? Data = null)
{
    public const string InvalidRequest = "INVALID_REQUEST";
    public const string NotFound = "NOT_FOUND";
    public const string Conflict = "CONFLICT";
    public const string ExcelError = "EXCEL_ERROR";
    public const string StateUnknown = "STATE_UNKNOWN";
}
