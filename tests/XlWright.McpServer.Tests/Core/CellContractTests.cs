using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

public sealed class CellContractTests
{
    [Theory]
    [InlineData("C:\\project", true)]
    [InlineData("C:\\", true)]
    [InlineData("\\\\server\\share", false)]
    [InlineData("\\\\?\\C:\\project", false)]
    [InlineData("\\\\.\\C:\\project", false)]
    [InlineData("C:\\project:stream", false)]
    public void OnlyDrivePathsAreLocal(string path, bool expected) => Assert.Equal(expected, ProjectConfig.IsLocal(path));

    [Theory]
    [InlineData("A1", "A1")]
    [InlineData("$b$2", "B2")]
    [InlineData("xfd1048576", "XFD1048576")]
    public void AddressesAreNormalized(string address, string expected) =>
        Assert.Equal(expected, CellAddress.Normalize(address));

    [Theory]
    [InlineData("")]
    [InlineData("A0")]
    [InlineData("A1048577")]
    [InlineData("XFE1")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A:A")]
    [InlineData("1:1")]
    [InlineData("A1,B2")]
    [InlineData("A1:B2:C3")]
    [InlineData("A1:XFE1")]
    [InlineData("Sheet!A1")]
    [InlineData("A1\n")]
    public void InvalidAddressesAreRefused(string address) =>
        Assert.Equal(ErrorDetail.InvalidAddress, Assert.Throws<CoreException>(() => CellAddress.Normalize(address)).Detail);

    [Fact]
    public void RectanglesAreNotImplemented() =>
        Assert.Equal(ErrorDetail.NotImplemented, Assert.Throws<CoreException>(() => CellAddress.Normalize("A1:B2")).Detail);

    [Theory]
    [InlineData(2000, "#NULL!")]
    [InlineData(2007, "#DIV/0!")]
    [InlineData(2015, "#VALUE!")]
    [InlineData(2023, "#REF!")]
    [InlineData(2029, "#NAME?")]
    [InlineData(2036, "#NUM!")]
    [InlineData(2042, "#N/A")]
    [InlineData(2045, "#SPILL!")]
    [InlineData(2050, "#CALC!")]
    public void KnownErrorsAreProjected(int number, string code) =>
        Assert.Equal(new CellError(code), CellValues.Project(unchecked((int)0x800A0000) | number));

    [Fact]
    public void UnknownErrorsKeepTheRawValue()
    {
        int raw = unchecked((int)0x800A0000) | 9999;

        Assert.Equal(new CellError(null, raw), CellValues.Project(raw));
    }

    [Theory]
    [InlineData(1d)]
    [InlineData("text")]
    [InlineData(true)]
    [InlineData(null)]
    public void OrdinaryValuesPassThrough(object? value) => Assert.Equal(value, CellValues.Project(value));

    [Fact]
    public void StringsAreWrittenAsText()
    {
        Assert.Equal("'=1+1", CellValues.ForValue2("=1+1"));
        Assert.Equal("'", CellValues.ForValue2(""));
        Assert.Equal(1d, CellValues.ForValue2(1d));
    }

    [Fact]
    public void ExactlyOneArrayIsRequired()
    {
        Assert.Equal(ErrorDetail.InvalidArgument, Assert.Throws<CoreException>(() => CellValues.Validate(null, null)).Detail);
        Assert.Equal(ErrorDetail.InvalidArgument, Assert.Throws<CoreException>(() => CellValues.Validate([[1d]], [["=1"]])).Detail);
    }

    [Fact]
    public void ArraysMustBeOneCell()
    {
        Assert.Equal(ErrorDetail.ShapeMismatch, Assert.Throws<CoreException>(() => CellValues.Validate([], null)).Detail);
        Assert.Equal(ErrorDetail.ShapeMismatch, Assert.Throws<CoreException>(() => CellValues.Validate([[]], null)).Detail);
        Assert.Equal(ErrorDetail.ShapeMismatch, Assert.Throws<CoreException>(() => CellValues.Validate([[1d, 2d]], null)).Detail);
        Assert.Equal(ErrorDetail.ShapeMismatch, Assert.Throws<CoreException>(() => CellValues.Validate([[1d], [2d]], null)).Detail);
    }

    [Fact]
    public void NumbersMustBeFinite() =>
        Assert.Equal(ErrorDetail.InvalidArgument, Assert.Throws<CoreException>(() => CellValues.Validate([[double.PositiveInfinity]], null)).Detail);

    [Fact]
    public void FormulasMustStartWithEquals()
    {
        Assert.Equal(ErrorDetail.InvalidArgument, Assert.Throws<CoreException>(() => CellValues.Validate(null, [["1+1"]])).Detail);
        Assert.Equal(new CellWrite(true, "=1+1"), CellValues.Validate(null, [["=1+1"]]));
        Assert.Equal(new CellWrite(true, null), CellValues.Validate(null, [[null]]));
    }

    [Fact]
    public void ReadBackComparesTypeAndValue()
    {
        Assert.True(CellValues.Verified(new(false, 1d), 1d, false));
        Assert.False(CellValues.Verified(new(false, 1d), "1", false));
        Assert.False(CellValues.Verified(new(false, "=1"), 1d, true));
        Assert.True(CellValues.Verified(new(false, null), null, false));
        Assert.False(CellValues.Verified(new(false, null), null, true));
        Assert.True(CellValues.Verified(new(true, "=1+1"), 2d, true));
        Assert.False(CellValues.Verified(new(true, "=1+1"), 2d, false));
    }
}
