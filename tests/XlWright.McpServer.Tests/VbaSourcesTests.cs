using System.Text;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "VBAの反映" and "xlwrightのVBAのモジュール": the checks and conversions before Excel.
public sealed class VbaSourcesTests : IDisposable
{
    private static readonly Encoding ShiftJis = VbaSources.AnsiEncoding(932);

    private readonly TempRepository _repo = new();
    private readonly AppFolder _app;

    public VbaSourcesTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
    }

    public void Dispose() => _repo.Dispose();

    [Fact]
    public void ConvertsStandardAndClassModulesToAnsiWithCrlfAndNames()
    {
        Source("src/modules/Calc.bas", "Option Explicit\n' 日本語\nPublic Function One() As Long\n    One = 1\nEnd Function");
        Source("src/classes/Counter.cls", "Option Explicit\r\nPublic Value As Long\r\n");

        var modules = Prepare();

        Assert.Equal(["Calc", "XlWright", "Counter"], modules.Select(m => m.Name));
        Assert.Equal([VbaModuleKind.Standard, VbaModuleKind.Standard, VbaModuleKind.Class], modules.Select(m => m.Kind));
        Assert.Equal(
            "Attribute VB_Name = \"Calc\"\r\nOption Explicit\r\n' 日本語\r\nPublic Function One() As Long\r\n    One = 1\r\nEnd Function\r\n",
            Imported(modules[0]));
        Assert.Equal(
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Counter\"\r\nOption Explicit\r\nPublic Value As Long\r\n",
            Imported(modules[2]));
        Assert.All(modules, m => Assert.StartsWith(Path.Combine(_app.FullPath, ".xlwright", "sync"), m.ImportPath));
    }

    [Fact]
    public void KeepsAClassHeaderAndAddsTheNameAfterIt()
    {
        Source("src/classes/Counter.cls",
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_PredeclaredId = True\r\nOption Explicit\r\n");

        var counter = Prepare().Single(m => m.Name == "Counter");

        Assert.Equal(
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Counter\"\r\nAttribute VB_PredeclaredId = True\r\nOption Explicit\r\n",
            Imported(counter));
    }

    [Fact]
    public void KeepsANameThatMatchesTheFileAndReadsAByteOrderMark()
    {
        Source("src/modules/Calc.bas", "﻿Attribute VB_Name = \"calc\"\r\nOption Explicit\r\n");

        var calc = Prepare().Single(m => m.Name == "calc");

        Assert.Equal("Attribute VB_Name = \"calc\"\r\nOption Explicit\r\n", Imported(calc));
    }

    [Fact]
    public void SkipsFilesThatStartWithADot()
    {
        Source("src/modules/.gitkeep", "");
        Source("src/workbook/.gitkeep", "");

        Assert.Equal(["XlWright"], Prepare().Select(m => m.Name));
        Assert.Empty(PrepareDocuments());
    }

    [Fact]
    public void ConvertsWorkbookModulesToCodeWithoutTheHeaderAndAttributes()
    {
        Source("src/workbook/Sheet1.cls",
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Sheet1\"\r\nAttribute VB_PredeclaredId = True\r\n" +
            "Option Explicit\n' 日本語\nPublic Function Twice(x)\nAttribute Twice.VB_Description = \"x\"\n    Twice = x * 2\nEnd Function\n");
        Source("src/workbook/ThisWorkbook.cls", "");

        var documents = PrepareDocuments();

        Assert.Equal(["Sheet1", "ThisWorkbook"], documents.Select(d => d.Name));
        Assert.Equal("Option Explicit\r\n' 日本語\r\nPublic Function Twice(x)\r\n    Twice = x * 2\r\nEnd Function", documents[0].Code);
        Assert.Equal("src/workbook/Sheet1.cls", documents[0].File);
        Assert.Equal("", documents[1].Code);
    }

    [Fact]
    public void RejectsAWorkbookModuleNameThatDiffersFromTheFile()
    {
        Source("src/workbook/Sheet1.cls", "Attribute VB_Name = \"Sheet2\"\r\nOption Explicit\r\n");

        var failure = PrepareFailure();

        Assert.Equal("SOURCE_INVALID", failure.Detail);
        Assert.Equal("src/workbook/Sheet1.cls", failure.Data!["file"]!.GetValue<string>());
        Assert.Equal(1, failure.Data["line"]!.GetValue<int>());
    }

    [Fact]
    public void RejectsAWorkbookModuleCharacterTheAnsiCodePageCannotHold()
    {
        Source("src/workbook/ThisWorkbook.cls", "Option Explicit\r\n' 😀\r\n");

        var failure = PrepareFailure();

        Assert.Equal("src/workbook/ThisWorkbook.cls", failure.Data!["file"]!.GetValue<string>());
        Assert.Equal(2, failure.Data["line"]!.GetValue<int>());
        Assert.False(File.Exists(Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas")));
    }

    [Fact]
    public void WritesTheXlWrightModuleWhenItIsMissingOrOlder()
    {
        var path = Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas");

        Prepare();
        var written = File.ReadAllText(path);
        File.WriteAllText(path, "' XlWright 0.0.9: old\r\n");
        Prepare();

        Assert.StartsWith("' XlWright 0.1.0:", written);
        Assert.Equal(written, File.ReadAllText(path));
    }

    [Fact]
    public void RewritesAnXlWrightModuleOfTheSameVersionWithOtherContent()
    {
        Source("src/modules/XlWright.bas", "' XlWright 0.1.0: edited\r\nOption Explicit\r\n");

        Prepare();

        Assert.Equal(VbaSources.XlWrightText("0.1.0"), File.ReadAllText(Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas")));
    }

    [Fact]
    public void TheXlWrightModuleTurnsObjectsArraysErrorsAndDatesIntoTypeNames()
    {
        var text = VbaSources.XlWrightText("0.1.0");

        Assert.Contains("Public Function Result(ByVal value As Variant) As Variant", text);
        Assert.Contains("IsObject(value)", text);
        Assert.Contains("IsArray(value)", text);
        Assert.Contains("vbError", text);
        Assert.Contains("vbDate", text);
    }

    [Theory]
    [InlineData("src/modules/XlWrightCall.bas")]
    [InlineData("src/classes/xlwrightcall.cls")]
    public void RejectsAModuleNamedLikeTheGeneratedOne(string file)
    {
        Source(file, "Option Explicit\r\n");

        var failure = PrepareFailure();

        Assert.Equal(file, failure.Data!["file"]!.GetValue<string>());
        Assert.Contains("XlWrightCall", failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void LeavesAnXlWrightModuleFromANewerVersion()
    {
        Source("src/modules/XlWright.bas", "' XlWright 0.2.0: newer\r\nOption Explicit\r\n");

        Prepare();

        Assert.Equal("' XlWright 0.2.0: newer\r\nOption Explicit\r\n", File.ReadAllText(Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas")));
    }

    [Theory]
    [InlineData("src/modules/Calc.cls", null, "Calc.cls")]
    [InlineData("src/classes/Calc.bas", null, "Calc.bas")]
    [InlineData("src/modules/notes.txt", null, "notes.txt")]
    [InlineData("src/modules/sub/Calc.bas", null, "sub")]
    [InlineData("src/classes/XlWright.cls", null, "XlWright")]
    [InlineData("src/workbook/Sheet1.bas", null, "Sheet1.bas")]
    [InlineData("src/workbook/sub/Sheet1.cls", null, "sub")]
    public void RejectsAFileInTheWrongPlace(string file, int? line, string reason)
    {
        Source(file, "Option Explicit\r\n");

        var failure = PrepareFailure();

        Assert.Equal("SOURCE_INVALID", failure.Detail);
        Assert.StartsWith(file.Contains("/sub/") ? file[..(file.IndexOf("/sub/") + 4)] : file, failure.Data!["file"]!.GetValue<string>());
        Assert.Equal(line, failure.Data["line"]?.GetValue<int>());
        Assert.Contains(reason, failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void RejectsANameThatDiffersFromTheFile()
    {
        Source("src/modules/Calc.bas", "' comment\r\nAttribute VB_Name = \"Other\"\r\n");

        var failure = PrepareFailure();

        Assert.Equal("src/modules/Calc.bas", failure.Data!["file"]!.GetValue<string>());
        Assert.Equal(2, failure.Data["line"]!.GetValue<int>());
        Assert.Contains("Other", failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void RejectsAnotherModuleNamedXlWright()
    {
        Source("src/modules/Helper.bas", "Attribute VB_Name = \"Helper\"\r\n");
        Source("src/classes/Thing.cls", "Attribute VB_Name = \"xlwright\"\r\n");

        var failure = PrepareFailure();

        Assert.Equal("src/classes/Thing.cls", failure.Data!["file"]!.GetValue<string>());
    }

    [Fact]
    public void RejectsACharacterTheAnsiCodePageCannotHold()
    {
        Source("src/modules/Calc.bas", "Option Explicit\r\n' ok\r\n' 😀\r\n");

        var failure = PrepareFailure();

        Assert.Equal("src/modules/Calc.bas", failure.Data!["file"]!.GetValue<string>());
        Assert.Equal(3, failure.Data["line"]!.GetValue<int>());
        Assert.Contains("932", failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void RejectsAFileThatIsNotUtf8()
    {
        File.WriteAllBytes(Write("src/modules/Calc.bas"), [0x27, 0x20, 0x82, 0xA0, 0x0D, 0x0A]);

        var failure = PrepareFailure();

        Assert.Equal("src/modules/Calc.bas", failure.Data!["file"]!.GetValue<string>());
        Assert.Contains("UTF-8", failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void DoesNotWriteTheXlWrightModuleWhenASourceIsRejected()
    {
        Source("src/modules/Calc.cls", "Option Explicit\r\n");

        PrepareFailure();

        Assert.False(File.Exists(Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas")));
    }

    private IReadOnlyList<VbaModule> Prepare()
    {
        var (modules, _, failure) = VbaSources.Prepare(_app, ShiftJis, "0.1.0");
        Assert.Null(failure);
        return modules!;
    }

    private IReadOnlyList<VbaDocumentModule> PrepareDocuments()
    {
        var (_, documents, failure) = VbaSources.Prepare(_app, ShiftJis, "0.1.0");
        Assert.Null(failure);
        return documents!;
    }

    private ToolFailure PrepareFailure()
    {
        var (modules, _, failure) = VbaSources.Prepare(_app, ShiftJis, "0.1.0");
        Assert.Null(modules);
        Assert.NotNull(failure);
        return failure;
    }

    private static string Imported(VbaModule module) => ShiftJis.GetString(File.ReadAllBytes(module.ImportPath));

    private void Source(string file, string text) => File.WriteAllText(Write(file), text, new UTF8Encoding(false));

    private string Write(string file)
    {
        var path = Path.Combine(_app.FullPath, file);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        return path;
    }
}
