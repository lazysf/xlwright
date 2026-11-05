using System.Text.Json.Nodes;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "マクロの実行" and "実行ごとに作るモジュール": the parts that need no Excel.
public sealed class MacroCallTests
{
    [Theory]
    [InlineData("Calc.Twice", "Calc", "Twice")]
    [InlineData("日本語名.実行_1", "日本語名", "実行_1")]
    [InlineData("calc.twice", "calc", "twice")]
    public void ParsesAModuleAndAProcedure(string text, string module, string procedure)
    {
        var (name, failure) = MacroName.Parse(text);

        Assert.Null(failure);
        Assert.Equal(new MacroName(module, procedure), name);
        Assert.Equal($"{module}.{procedure}", name!.ToString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("Twice")]
    [InlineData("Calc.Twice.Again")]
    [InlineData("'Book.xlsm'!Calc.Twice")]
    [InlineData("Book.xlsm!Calc.Twice")]
    [InlineData("[Book.xlsm]Calc.Twice")]
    [InlineData(@"C:\Book.xlsm!Calc.Twice")]
    [InlineData("apps/Calc.Twice")]
    [InlineData("1Calc.Twice")]
    [InlineData("Calc. Twice")]
    [InlineData("Calc.Twice()")]
    [InlineData("_Calc.Twice")]
    public void RejectsAnythingElseAsInvalidMacroName(string text)
    {
        var (name, failure) = MacroName.Parse(text);

        Assert.Null(name);
        Assert.Equal(ToolFailure.InvalidRequest, failure!.Code);
        Assert.Equal("INVALID_MACRO_NAME", failure.Detail);
    }

    [Fact]
    public void GeneratesACallToAFunction()
    {
        var code = MacroCall.Generate(new MacroName("Calc", "Twice"), argumentCount: 2, isFunction: true);

        Assert.Equal(
            "Option Explicit\r\n\r\n" +
            "Public Function Run(ByVal args As Variant, ByVal answers As Variant) As Variant\r\n" +
            "    Dim result As Variant\r\n" +
            "    On Error GoTo Failed\r\n" +
            "    XlWright.BeginRun answers\r\n" +
            "    result = XlWright.Result(Calc.Twice((args(0)), (args(1))))\r\n" +
            "    Run = Array(result, XlWright.EndRun())\r\n" +
            "    Exit Function\r\n" +
            "Failed:\r\n" +
            "    result = Array(1, Err.Number, Err.Description, Err.Source)\r\n" +
            "    Run = Array(result, XlWright.EndRun())\r\n" +
            "End Function\r\n",
            code);
    }

    [Theory]
    [InlineData(0, "    Call Calc.Write\r\n")]
    [InlineData(1, "    Call Calc.Write((args(0)))\r\n")]
    public void GeneratesACallToASub(int argumentCount, string call)
    {
        var code = MacroCall.Generate(new MacroName("Calc", "Write"), argumentCount, isFunction: false);

        Assert.Contains(call + "    result = Array(0, Empty)\r\n", code);
    }

    [Fact]
    public void GeneratesAFunctionCallWithoutArguments()
    {
        var code = MacroCall.Generate(new MacroName("Calc", "Now"), argumentCount: 0, isFunction: true);

        Assert.Contains("    result = XlWright.Result(Calc.Now())\r\n", code);
    }

    public static TheoryData<object?, string?> Returns => new()
    {
        { 42, "42" },
        { (short)7, "7" },
        { 2.5, "2.5" },
        { 1.25m, "1.25" },
        { (byte)3, "3" },
        { 1.5f, "1.5" },
        { "日本語", "\"日本語\"" },
        { true, "true" },
        { null, null },
        { DBNull.Value, null },
    };

    [Theory]
    [MemberData(nameof(Returns))]
    public void ReturnsNumbersStringsBooleansAndNull(object? value, string? json)
    {
        var (node, typeName) = MacroCall.ReturnValue(value);

        Assert.Null(typeName);
        Assert.Equal(json, node?.ToJsonString(new System.Text.Json.JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
    }

    [Fact]
    public void NamesAnyOtherType()
    {
        var (node, typeName) = MacroCall.ReturnValue(new DateTime(2026, 11, 3));

        Assert.Null(node);
        Assert.Equal("DateTime", typeName);
    }
}
