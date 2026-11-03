using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>
/// range_read and range_write (docs/api.md, stage 1). The arguments, the range, and the array are checked before
/// the session, so those failures do not depend on the session state.
/// </summary>
internal sealed class RangeTools(ExcelSession? session)
{
    // docs/api.md "成功と失敗": the body of a response is at most 20,000 characters.
    private const int MaxResponseCharacters = 20_000;

    public IEnumerable<McpServerTool> Create() =>
    [
        ToolDefinition.Create(Read, new()
        {
            Name = "range_read",
            Description = "Reads the values or the formulas of one range on one sheet of the main workbook, recalculating before it reads values. " +
                "Pass the sheet by name and the range as one cell or one rectangle in A1 notation (A1 or A1:C10). " +
                "Values are numbers, strings, booleans, null for an empty cell, and {\"error\":\"#N/A\"} for a cell error; dates are serial numbers. " +
                "Formulas use English function names. Reads at most 10,000 cells and 20,000 characters.",
            ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false,
        }, ReadInput, ReadOutput),
        ToolDefinition.Create(Write, new()
        {
            Name = "range_write",
            Description = "Writes values or formulas to one range on one sheet of the main workbook, or clears its contents and keeps the formats. " +
                "Pass exactly one of values, formulas, or clear. An array is a list of rows and must have the size of the range. " +
                "Excel reads a value string as if a person typed it, so \"1/2\" can become a date; pass formulas (English function names) in formulas. " +
                "Events do not run while it writes. It does not save.",
            ReadOnly = false, Destructive = true, Idempotent = true, OpenWorld = false,
        }, WriteInput, WriteOutput),
    ];

    private async Task<CallToolResult> Read(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("range_read", context);
        var (sheet, address, formulas) = ("", "", false);
        var invalid = arguments.RejectUnknown("sheet", "range", "content")
            ?? arguments.RequireString("sheet", out sheet)
            ?? arguments.RequireString("range", out address)
            ?? ReadContent(context, out formulas);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        var (range, failure) = CellRange.Parse(address);
        failure ??= RangeChecks.Read(range!);
        if (failure is not null || session is null)
        {
            return ToolResults.Failure(failure ?? NoSession, Current());
        }

        var result = ToolResults.From(await session.ReadAsync(sheet, range!, formulas));
        var length = ((TextContentBlock)result.Content[0]).Text.Length;
        return result.IsError != true && length > MaxResponseCharacters
            ? ToolResults.Failure(new ToolFailure(ToolFailure.InvalidRequest, "TOO_LARGE",
                $"The response for {range!.Address} has {length} characters; the limit is {MaxResponseCharacters}. Read a smaller range, such as fewer rows at a time."),
                Current())
            : result;
    }

    private async Task<CallToolResult> Write(RequestContext<CallToolRequestParams> context)
    {
        var arguments = ToolArguments.From("range_write", context);
        var (sheet, address) = ("", "");
        var invalid = arguments.RejectUnknown("sheet", "range", "values", "formulas", "clear")
            ?? arguments.RequireString("sheet", out sheet)
            ?? arguments.RequireString("range", out address);
        if (invalid is not null)
        {
            return ToolResults.InvalidRequest(invalid, Current());
        }

        var (range, failure) = CellRange.Parse(address);
        if (failure is not null)
        {
            return ToolResults.Failure(failure, Current());
        }

        var (write, writeInvalid, shape) = WriteContent(context.Params?.Arguments!, range!);
        if (writeInvalid is not null)
        {
            return ToolResults.InvalidRequest(writeInvalid, Current());
        }

        failure = shape ?? RangeChecks.Write(range!, write!);
        if (failure is not null || session is null)
        {
            return ToolResults.Failure(failure ?? NoSession, Current());
        }

        return ToolResults.From(await session.WriteAsync(sheet, range!, write!));
    }

    private SessionState Current() => session is null ? SessionState.Closed : SessionState.From(session.Status());

    private static ToolFailure NoSession { get; } =
        new(ToolFailure.Conflict, "NO_SESSION", "There is no session. Start one with session_start.");

    private static string? ReadContent(RequestContext<CallToolRequestParams> context, out bool formulas)
    {
        formulas = false;
        if (context.Params?.Arguments?.TryGetValue("content", out var content) != true)
        {
            return null;
        }

        formulas = content.ValueKind == JsonValueKind.String && content.GetString() == "formulas";
        return formulas || (content.ValueKind == JsonValueKind.String && content.GetString() == "values")
            ? null
            : "range_read: content must be \"values\" or \"formulas\".";
    }

    // Returns what to write, or the INVALID_REQUEST message, or a SHAPE_MISMATCH for rows of different lengths.
    private static (RangeWrite? Write, string? Invalid, ToolFailure? Shape) WriteContent(IDictionary<string, JsonElement> arguments, CellRange range)
    {
        var given = new[] { "values", "formulas", "clear" }.Where(arguments.ContainsKey).ToList();
        if (given.Count != 1)
        {
            return (null, "range_write takes exactly one of values, formulas, or clear.", null);
        }

        var element = arguments[given[0]];
        if (given[0] == "clear")
        {
            return element.ValueKind == JsonValueKind.True
                ? (RangeWrite.Clear, null, null)
                : (null, "range_write: clear must be true. To write, pass values or formulas instead.", null);
        }

        var isValues = given[0] == "values";
        var rows = Rows(element, isValues
            ? cell => cell.ValueKind is JsonValueKind.Number or JsonValueKind.String or JsonValueKind.True or JsonValueKind.False or JsonValueKind.Null
            : cell => cell.ValueKind == JsonValueKind.String);
        if (rows is null)
        {
            return (null, isValues
                ? "range_write: values must be an array of rows, each an array of numbers, strings, booleans, or null."
                : "range_write: formulas must be an array of rows, each an array of strings.", null);
        }

        var columns = rows.Count == 0 ? 0 : rows[0].Length;
        if (rows.Any(row => row.Length != columns))
        {
            return (null, null, new ToolFailure(ToolFailure.InvalidRequest, "SHAPE_MISMATCH",
                $"The rows of {given[0]} have different numbers of cells. Pass {range.Rows} row(s) of {range.Columns} cell(s) for {range.Address}."));
        }

        if (isValues)
        {
            var values = new object?[rows.Count, columns];
            Fill(rows, (row, column, cell) => values[row, column] = cell.ValueKind switch
            {
                JsonValueKind.Number => cell.GetDouble(),
                JsonValueKind.String => cell.GetString(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                _ => null,
            });
            return (RangeWrite.OfValues(values), null, null);
        }

        var formulas = new string[rows.Count, columns];
        Fill(rows, (row, column, cell) => formulas[row, column] = cell.GetString()!);
        return (RangeWrite.OfFormulas(formulas), null, null);
    }

    private static List<JsonElement[]>? Rows(JsonElement element, Func<JsonElement, bool> cellAllowed)
    {
        if (element.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var rows = new List<JsonElement[]>();
        foreach (var row in element.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array || !row.EnumerateArray().All(cellAllowed))
            {
                return null;
            }

            rows.Add(row.EnumerateArray().ToArray());
        }

        return rows;
    }

    private static void Fill(List<JsonElement[]> rows, Action<int, int, JsonElement> set)
    {
        for (var row = 0; row < rows.Count; row++)
        {
            for (var column = 0; column < rows[row].Length; column++)
            {
                set(row, column, rows[row][column]);
            }
        }
    }

    private const string ReadInput = """
        {"type":"object","properties":{
         "sheet":{"type":"string","description":"The sheet name."},
         "range":{"type":"string","description":"One cell or one rectangle in A1 notation, such as A1 or A1:C10, without a sheet name."},
         "content":{"enum":["values","formulas"],"description":"values (the default) reads calculated values; formulas reads the formulas."}},
         "required":["sheet","range"],"additionalProperties":false}
        """;

    private const string WriteInput = """
        {"type":"object","properties":{
         "sheet":{"type":"string","description":"The sheet name."},
         "range":{"type":"string","description":"One cell or one rectangle in A1 notation, such as A1 or A1:C10, without a sheet name."},
         "values":{"type":"array","items":{"type":"array","items":{"type":["number","string","boolean","null"]}},"description":"Rows of values; null empties a cell. A string that starts with = is rejected."},
         "formulas":{"type":"array","items":{"type":"array","items":{"type":"string"}},"description":"Rows of formulas in English, such as =SUM(A1:A3)."},
         "clear":{"const":true,"description":"true clears the values and formulas, and keeps the formats."}},
         "required":["sheet","range"],"additionalProperties":false}
        """;

    private const string Cells = """
        {"type":"array","items":{"type":"array"}}
        """;

    private static readonly string ReadOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"sheet":{"type":"string"},"range":{"type":"string"},
         "values":{{{Cells}}},"formulas":{{{Cells}}}},"required":["session","sheet","range"]}
        """;

    private static readonly string WriteOutput = $$$"""
        {"type":"object","properties":{"session":{{{ToolDefinition.Session}}},"sheet":{"type":"string"},"range":{"type":"string"},
         "cells":{"type":"integer"}},"required":["session","sheet","range","cells"]}
        """;
}
