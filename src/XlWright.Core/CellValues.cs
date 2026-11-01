using System.Text.Json.Serialization;

namespace XlWright.Core;

// An Excel error in a cell value. Raw is set only when the code is unknown.
public sealed record CellError(string? Code, [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] int? Raw = null)
{
    public string Type => "error";
}

internal sealed record CellContent(object? Value, string? Formula);

internal sealed record CellWrite(bool IsFormula, object? Value);

internal static class CellValues
{
    // The numbers are the S4 observations; Value2 returns them as Int32.
    private static readonly Dictionary<int, string> Errors = new()
    {
        [2000] = "#NULL!", [2007] = "#DIV/0!", [2015] = "#VALUE!", [2023] = "#REF!", [2029] = "#NAME?",
        [2036] = "#NUM!", [2042] = "#N/A", [2045] = "#SPILL!", [2050] = "#CALC!"
    };

    internal static object? Project(object? value)
    {
        if (value is not int raw) return value;
        return Errors.TryGetValue(raw & 0xFFFF, out string? code) && (raw & ~0xFFFF) == unchecked((int)0x800A0000) ?
            new CellError(code) : new CellError(null, raw);
    }

    // The contract requires a [[x]] array even for one cell. Exactly one of the two arrays must be given.
    internal static CellWrite Validate(object?[][]? values, string?[][]? formulas)
    {
        if (values is null == formulas is null)
            throw Invalid("Give exactly one of values and formulas.");
        int rows = values?.Length ?? formulas!.Length;
        int columns = values?.FirstOrDefault()?.Length ?? formulas?.FirstOrDefault()?.Length ?? 0;
        if (rows != 1 || columns != 1)
            throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.ShapeMismatch,
                "The array must be [[x]] for one cell. Fix it and retry.");
        if (values is not null)
        {
            if (values[0][0] is not (null or double or string or bool) || values[0][0] is double number && !double.IsFinite(number))
                throw Invalid("A cell value must be a finite number, string, boolean or null.");
            return new(false, values[0][0]);
        }
        if (formulas![0][0] is { } formula && !formula.StartsWith('='))
            throw Invalid("A formula must start with =.");
        return new(true, formulas[0][0]);
    }

    // Prefixing keeps a string as text instead of letting Excel parse it (S1).
    internal static object? ForValue2(object? value) => value is string text ? "'" + text : value;

    internal static bool Verified(CellWrite write, object? value, bool hasFormula)
    {
        if (write.Value is null) return value is null && !hasFormula;
        return write.IsFormula ? hasFormula : !hasFormula && Equals(value, write.Value);
    }

    private static CoreException Invalid(string message) =>
        new(ErrorCode.InvalidRequest, ErrorDetail.InvalidArgument, message + " Fix the arguments and retry.");
}
