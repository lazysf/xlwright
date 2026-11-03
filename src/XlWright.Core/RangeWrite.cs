namespace XlWright.Core;

/// <summary>
/// What range_write writes: values (Value2), formulas (Formula), or neither, which clears the contents.
/// Values hold double, string, bool, or null.
/// </summary>
public sealed record RangeWrite(object?[,]? Values, string[,]? Formulas)
{
    public static RangeWrite Clear { get; } = new(null, null);

    public static RangeWrite OfValues(object?[,] values) => new(values, null);

    public static RangeWrite OfFormulas(string[,] formulas) => new(null, formulas);
}

/// <summary>The checks of range_read and range_write that need no Excel (docs/api.md).</summary>
public static class RangeChecks
{
    /// <summary>
    /// The most cells range_read reads. A cell takes at least two characters in the response (a digit and a
    /// comma), so more cells cannot fit in the 20,000 characters; this rejects them before Excel builds the array.
    /// </summary>
    public const int MaxReadCells = 10_000;

    public static ToolFailure? Read(CellRange range) => range.Count > MaxReadCells
        ? new ToolFailure(ToolFailure.InvalidRequest, "TOO_LARGE",
            $"{range.Address} has {range.Count} cells; range_read reads at most {MaxReadCells}. Read a smaller range, such as fewer rows at a time.")
        : null;

    public static ToolFailure? Write(CellRange range, RangeWrite write)
    {
        Array? array = (Array?)write.Values ?? write.Formulas;
        if (array is not null && (array.GetLength(0) != range.Rows || array.GetLength(1) != range.Columns))
        {
            return new ToolFailure(ToolFailure.InvalidRequest, "SHAPE_MISMATCH",
                $"{range.Address} is {range.Rows} row(s) by {range.Columns} column(s), but the array is {array.GetLength(0)} by {array.GetLength(1)}. " +
                "Pass an array of the same size as the range.");
        }

        if (write.Values is { } values)
        {
            // Excel turns a value that starts with = into a formula (docs/notes/range-values.md).
            for (var row = 0; row < range.Rows; row++)
            {
                for (var column = 0; column < range.Columns; column++)
                {
                    if (values[row, column] is string text && text.StartsWith('='))
                    {
                        return new ToolFailure(ToolFailure.InvalidRequest, "VALUE_LOOKS_LIKE_FORMULA",
                            $"The value for {range.CellAt(row, column)} starts with =, so Excel would write it as a formula. " +
                            "Pass formulas in formulas instead of values.");
                    }
                }
            }
        }

        return null;
    }
}
