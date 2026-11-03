using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace XlWright.McpServer;

/// <summary>
/// Reads tool arguments. The SDK ignores unknown arguments, so the tools check them here (docs/api.md "引数").
/// Each method returns null on success, or the message for an INVALID_REQUEST.
/// </summary>
internal sealed class ToolArguments(string tool, IDictionary<string, JsonElement>? arguments)
{
    public static ToolArguments From(string tool, RequestContext<CallToolRequestParams> context) =>
        new(tool, context.Params?.Arguments);

    public string? RejectUnknown(params string[] allowed)
    {
        var unknown = (arguments?.Keys ?? []).Where(name => !allowed.Contains(name, StringComparer.Ordinal)).ToList();
        if (unknown.Count == 0)
        {
            return null;
        }

        var accepted = allowed.Length == 0 ? "takes no arguments" : $"takes only {string.Join(", ", allowed)}";
        return $"{tool} {accepted}. Call it again without: {string.Join(", ", unknown)}.";
    }

    public string? RequireString(string name, out string value)
    {
        value = "";
        if (arguments is null || !arguments.TryGetValue(name, out var element))
        {
            return $"{tool} requires {name} (a string).";
        }

        if (element.ValueKind != JsonValueKind.String)
        {
            return $"{tool}: {name} must be a string.";
        }

        value = element.GetString()!;
        return null;
    }

    public string? RequireBoolean(string name, out bool value)
    {
        value = false;
        if (arguments is null || !arguments.TryGetValue(name, out var element))
        {
            return $"{tool} requires {name} (true or false).";
        }

        if (element.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            return $"{tool}: {name} must be true or false.";
        }

        value = element.GetBoolean();
        return null;
    }
}
