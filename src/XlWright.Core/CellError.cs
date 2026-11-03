namespace XlWright.Core;

/// <summary>A cell error value, such as #N/A. range_read returns it as {"error":"#N/A"}.</summary>
public sealed record CellError(string Name)
{
    // The XlCVError values Microsoft documents, and #CALC! (2050), which Excel returned in docs/notes/range-values.md.
    private static readonly Dictionary<int, string> Names = new()
    {
        [2000] = "#NULL!",
        [2007] = "#DIV/0!",
        [2015] = "#VALUE!",
        [2023] = "#REF!",
        [2029] = "#NAME?",
        [2036] = "#NUM!",
        [2042] = "#N/A",
        [2045] = "#SPILL!",
        [2050] = "#CALC!",
    };

    /// <summary>
    /// Names the error that Value2 returns as an Int32 HRESULT (0x800A0000 plus the error number). A number that
    /// is not in the list is kept as #ERROR(number) rather than guessed.
    /// </summary>
    public static CellError FromHResult(int hresult)
    {
        var number = hresult & 0xFFFF;
        return new CellError(Names.TryGetValue(number, out var name) ? name : $"#ERROR({number})");
    }
}
