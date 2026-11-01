using System.Text.Json;
using XlWright.Core;

namespace XlWright.McpServer;

// Checks JSON shapes here because the SDK's default failure text has no details. Core checks the meaning.
internal sealed class ToolArguments
{
    private readonly string operation;
    private readonly ExcelSession session;
    private readonly IDictionary<string, JsonElement> values;

    internal ToolArguments(string operation, ExcelSession session, IDictionary<string, JsonElement>? values, IReadOnlyCollection<string> known)
    {
        this.operation = operation;
        this.session = session;
        this.values = values ?? new Dictionary<string, JsonElement>();
        if (this.values.Keys.Any(name => !known.Contains(name))) throw Invalid("Unknown argument.");
    }

    internal string Text(string name)
    {
        if (!values.TryGetValue(name, out var value) || value.ValueKind != JsonValueKind.String)
            throw Invalid($"{name} is required and must be a string.");
        return value.GetString()!;
    }

    internal bool Flag(string name)
    {
        if (!values.TryGetValue(name, out var value)) return false;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => throw Invalid($"{name} must be a boolean.")
        };
    }

    internal string[]? TextList(string name)
    {
        if (!values.TryGetValue(name, out var value)) return null;
        if (value.ValueKind != JsonValueKind.Array || value.EnumerateArray().Any(item => item.ValueKind != JsonValueKind.String))
            throw Invalid($"{name} must be an array of strings.");
        return value.EnumerateArray().Select(item => item.GetString()!).ToArray();
    }

    internal object?[][]? CellValues(string name) => Grid<object?>(name, item => item.ValueKind switch
    {
        JsonValueKind.Number => item.GetDouble(),
        JsonValueKind.String => item.GetString(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        _ => throw Invalid($"{name} must hold numbers, strings, booleans or null.")
    });

    internal string?[][]? Formulas(string name) => Grid<string?>(name, item => item.ValueKind switch
    {
        JsonValueKind.String => item.GetString(),
        JsonValueKind.Null => null,
        _ => throw Invalid($"{name} must hold strings or null.")
    });

    private T[][]? Grid<T>(string name, Func<JsonElement, T> read)
    {
        if (!values.TryGetValue(name, out var value)) return null;
        if (value.ValueKind != JsonValueKind.Array || value.EnumerateArray().Any(row => row.ValueKind != JsonValueKind.Array))
            throw Invalid($"{name} must be a two-dimensional array.");
        return value.EnumerateArray().Select(row => row.EnumerateArray().Select(read).ToArray()).ToArray();
    }

    private SessionOperationException Invalid(string message) => session.InvalidArgument(operation, message);
}
