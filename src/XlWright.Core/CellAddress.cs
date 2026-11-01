using System.Text.RegularExpressions;

namespace XlWright.Core;

internal static partial class CellAddress
{
    private const int LastColumn = 16384; // XFD
    private const int LastRow = 1048576;

    [GeneratedRegex(@"\A\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\z", RegexOptions.CultureInvariant)]
    private static partial Regex Cell();

    // Returns the cell without $ in upper case. PoC-1 serves single cells; a rectangle is valid but not implemented.
    internal static string Normalize(string address)
    {
        string[] parts = address.Split(':');
        if (parts.Length > 2 || parts.Any(part => Parse(part) is null))
            throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.InvalidAddress,
                "address must be one A1 cell such as B2. Fix it and retry.");
        if (parts.Length == 2)
            throw new CoreException(ErrorCode.InvalidRequest, ErrorDetail.NotImplemented,
                "Rectangular ranges are not implemented yet. Use one cell.");
        return Parse(parts[0])!;
    }

    private static string? Parse(string cell)
    {
        var match = Cell().Match(cell);
        if (!match.Success) return null;
        string column = match.Groups[1].Value.ToUpperInvariant();
        int number = column.Aggregate(0, (total, letter) => total * 26 + letter - 'A' + 1);
        return number <= LastColumn && int.Parse(match.Groups[2].ValueSpan) is >= 1 and <= LastRow ?
            column + int.Parse(match.Groups[2].ValueSpan) : null;
    }
}
