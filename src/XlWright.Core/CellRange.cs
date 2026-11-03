using System.Text.RegularExpressions;

namespace XlWright.Core;

/// <summary>
/// One cell or one rectangle on a sheet, in A1 notation (docs/api.md "range_read"). Rows and columns are 1-based.
/// </summary>
public sealed partial record CellRange(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public const int MaxRows = 1_048_576;
    public const int MaxColumns = 16_384;

    public int Rows => LastRow - FirstRow + 1;

    public int Columns => LastColumn - FirstColumn + 1;

    public long Count => (long)Rows * Columns;

    /// <summary>The address without $, upper case, top-left first; one cell has no colon.</summary>
    public string Address => Rows == 1 && Columns == 1
        ? Cell(FirstRow, FirstColumn)
        : Cell(FirstRow, FirstColumn) + ":" + Cell(LastRow, LastColumn);

    /// <summary>The address of the cell at a 0-based offset from the top-left cell.</summary>
    public string CellAt(int row, int column) => Cell(FirstRow + row, FirstColumn + column);

    /// <summary>
    /// Accepts a cell (A1) or a rectangle (A1:C10), with or without $, in any case and either corner first.
    /// Excel's Range also takes whole columns and rows, several areas, names, and sheet prefixes; they are rejected
    /// here so that one call reads or writes exactly one rectangle on the named sheet (docs/notes/range-values.md).
    /// </summary>
    public static (CellRange? Range, ToolFailure? Failure) Parse(string text)
    {
        var match = A1Notation().Match(text);
        if (!match.Success)
        {
            return Invalid(text);
        }

        var first = ParseCell(match.Groups["c1"].Value, match.Groups["r1"].Value);
        var last = match.Groups["c2"].Success ? ParseCell(match.Groups["c2"].Value, match.Groups["r2"].Value) : first;
        if (first is null || last is null)
        {
            return Invalid(text);
        }

        var (row1, column1) = first.Value;
        var (row2, column2) = last.Value;
        return (new CellRange(Math.Min(row1, row2), Math.Min(column1, column2), Math.Max(row1, row2), Math.Max(column1, column2)), null);
    }

    private static (int Row, int Column)? ParseCell(string column, string row)
    {
        var number = 0;
        foreach (var letter in column.ToUpperInvariant())
        {
            number = number * 26 + (letter - 'A' + 1);
        }

        return int.TryParse(row, out var rowNumber) && rowNumber is >= 1 and <= MaxRows && number <= MaxColumns
            ? (rowNumber, number)
            : null;
    }

    private static string Cell(int row, int column)
    {
        var letters = "";
        for (var rest = column; rest > 0; rest = (rest - 1) / 26)
        {
            letters = (char)('A' + (rest - 1) % 26) + letters;
        }

        return letters + row;
    }

    private static (CellRange?, ToolFailure?) Invalid(string text) =>
        (null, new ToolFailure(ToolFailure.InvalidRequest, "INVALID_RANGE",
            $"range \"{text}\" is not one cell or one rectangle in A1 notation. Pass it like A1 or A1:C10, without a sheet name; " +
            "whole columns or rows, several areas, and names are not accepted."));

    [GeneratedRegex(@"^\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>[0-9]{1,7})(:\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>[0-9]{1,7}))?\z", RegexOptions.CultureInvariant)]
    private static partial Regex A1Notation();
}
