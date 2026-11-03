using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The checks of range_read and range_write that need no Excel (docs/api.md, docs/notes/range-values.md).
public sealed class RangeTests
{
    [Theory]
    [InlineData("A1", 1, 1, 1, 1, "A1")]
    [InlineData("A1:C2", 1, 1, 2, 3, "A1:C2")]
    [InlineData("a1:c2", 1, 1, 2, 3, "A1:C2")]
    [InlineData("$A$1:$C$2", 1, 1, 2, 3, "A1:C2")]
    [InlineData("C2:A1", 1, 1, 2, 3, "A1:C2")]
    [InlineData("B5:B5", 5, 2, 5, 2, "B5")]
    [InlineData("XFD1048576", 1048576, 16384, 1048576, 16384, "XFD1048576")]
    [InlineData("AA10:AZ20", 10, 27, 20, 52, "AA10:AZ20")]
    public void ParsesACellOrARectangle(string text, int firstRow, int firstColumn, int lastRow, int lastColumn, string address)
    {
        var (range, failure) = CellRange.Parse(text);

        Assert.Null(failure);
        Assert.Equal(new CellRange(firstRow, firstColumn, lastRow, lastColumn), range);
        Assert.Equal(address, range!.Address);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" A1")]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("A:A")]
    [InlineData("1:1")]
    [InlineData("A1,B2")]
    [InlineData("Sheet1!A1")]
    [InlineData("SalesTotal")]
    [InlineData("A1:B2:C3")]
    [InlineData("R1C1")]
    [InlineData("A1\n")]
    public void RejectsAnythingElseAsInvalidRange(string text)
    {
        var (range, failure) = CellRange.Parse(text);

        Assert.Null(range);
        Assert.Equal(ToolFailure.InvalidRequest, failure!.Code);
        Assert.Equal("INVALID_RANGE", failure.Detail);
    }

    [Fact]
    public void CountsRowsColumnsAndCells()
    {
        var range = CellRange.Parse("B2:D11").Range!;

        Assert.Equal(10, range.Rows);
        Assert.Equal(3, range.Columns);
        Assert.Equal(30, range.Count);
        Assert.Equal(17_179_869_184L, CellRange.Parse("A1:XFD1048576").Range!.Count);
    }

    [Fact]
    public void ReadsUpToTheCellLimit()
    {
        Assert.Null(RangeChecks.Read(CellRange.Parse("A1:CV100").Range!));

        var failure = RangeChecks.Read(CellRange.Parse("A1:CV101").Range!);

        Assert.Equal(ToolFailure.InvalidRequest, failure!.Code);
        Assert.Equal("TOO_LARGE", failure.Detail);
    }

    [Theory]
    [InlineData("A1:B2", 2, 2, null)]
    [InlineData("A1:B2", 2, 3, "SHAPE_MISMATCH")]
    [InlineData("A1:B2", 1, 2, "SHAPE_MISMATCH")]
    [InlineData("A1", 1, 1, null)]
    public void WritesOnlyAnArrayOfTheRangeShape(string address, int rows, int columns, string? detail)
    {
        var range = CellRange.Parse(address).Range!;

        var values = RangeChecks.Write(range, RangeWrite.OfValues(new object?[rows, columns]));
        var formulas = RangeChecks.Write(range, RangeWrite.OfFormulas(new string[rows, columns]));

        Assert.Equal(detail, values?.Detail);
        Assert.Equal(detail, formulas?.Detail);
    }

    [Fact]
    public void ClearsAnyRange()
    {
        Assert.Null(RangeChecks.Write(CellRange.Parse("A1:XFD1048576").Range!, RangeWrite.Clear));
    }

    [Theory]
    [InlineData("=1+1")]
    [InlineData("=")]
    public void RejectsAValueThatLooksLikeAFormula(string text)
    {
        var values = new object?[,] { { 1.0, "plain" }, { true, text } };

        var failure = RangeChecks.Write(CellRange.Parse("A1:B2").Range!, RangeWrite.OfValues(values));

        Assert.Equal(ToolFailure.InvalidRequest, failure!.Code);
        Assert.Equal("VALUE_LOOKS_LIKE_FORMULA", failure.Detail);
        Assert.Contains("B2", failure.Message);
        Assert.Contains("formulas", failure.Message);
    }

    [Fact]
    public void AcceptsValuesThatOnlyContainAnEqualsSign()
    {
        var values = new object?[,] { { "a=b", " =1", null } };

        Assert.Null(RangeChecks.Write(CellRange.Parse("A1:C1").Range!, RangeWrite.OfValues(values)));
    }

    [Theory]
    [InlineData(-2146826281, "#DIV/0!")]
    [InlineData(-2146826246, "#N/A")]
    [InlineData(-2146826252, "#NUM!")]
    [InlineData(-2146826259, "#NAME?")]
    [InlineData(-2146826273, "#VALUE!")]
    [InlineData(-2146826288, "#NULL!")]
    [InlineData(-2146826265, "#REF!")]
    [InlineData(-2146826238, "#CALC!")]
    public void NamesTheCellErrorsExcelReturned(int hresult, string name)
    {
        Assert.Equal(name, CellError.FromHResult(hresult).Name);
    }

    [Fact]
    public void KeepsTheCodeOfAnUnknownCellError()
    {
        Assert.Equal("#ERROR(2099)", CellError.FromHResult(unchecked((int)0x800A0833)).Name);
    }
}
