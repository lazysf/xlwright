using System.IO.Compression;
using System.Text;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// The checks and the writing of vba_extract (docs/api.md "vba_extract"), without Excel.
public sealed class VbaExtractionTests : IDisposable
{
    private static readonly Encoding Ansi = VbaSources.AnsiEncoding(932);

    private readonly TempRepository _repo = new();
    private readonly AppFolder _app;

    public VbaExtractionTests()
    {
        _repo.AddApp("apps/old");
        Package(Path.Combine(_repo.Root, "apps", "old", "Old.xlsm"), "xl/workbook.xml", "xl/vbaProject.bin");
        _app = AppFolder.Resolve(_repo.Root, "apps/old").Folder!;
    }

    public void Dispose() => _repo.Dispose();

    [Fact]
    public void ChecksTheMainWorkbookWhenNoneIsGiven()
    {
        var (source, failure) = VbaExtraction.Check(_app, null);

        Assert.Null(failure);
        Assert.Equal(_app.WorkbookPath, source);
    }

    // Excel holds the main workbook it opened for editing with write access and read sharing.
    [Fact]
    public void ChecksAndCopiesTheMainWorkbookWhileItIsHeldForEditing()
    {
        using var held = new FileStream(_app.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        var target = Path.Combine(_repo.Root, "copy.xlsm");

        Assert.Null(VbaExtraction.Check(_app, null).Failure);
        VbaExtraction.CopyShared(_app.WorkbookPath, target);

        Assert.Equal(held.Length, new FileInfo(target).Length);
    }

    [Fact]
    public void ChecksAnotherWorkbookInTheAppFolder()
    {
        Package(Path.Combine(_app.FullPath, "data", "Legacy.xlsm"), "xl/workbook.xml");

        var (source, failure) = VbaExtraction.Check(_app, "data/Legacy.xlsm");

        Assert.Null(failure);
        Assert.Equal(Path.Combine(_app.FullPath, "data", "Legacy.xlsm"), source);
    }

    [Theory]
    [InlineData("../Other.xlsm")]
    [InlineData("C:/Windows/Other.xlsm")]
    [InlineData("//server/share/Other.xlsm")]
    public void RejectsAWorkbookOutsideTheAppFolder(string workbook)
    {
        AssertFailure(VbaExtraction.Check(_app, workbook).Failure, "INVALID_REQUEST", "PATH_NOT_ALLOWED");
    }

    [Fact]
    public void ReportsAMissingWorkbook()
    {
        AssertFailure(VbaExtraction.Check(_app, "data/None.xlsm").Failure, "NOT_FOUND", "WORKBOOK_NOT_FOUND");
    }

    [Fact]
    public void RejectsAWorkbookThatIsNotMacroEnabled()
    {
        Package(Path.Combine(_app.FullPath, "data", "Plain.xlsx"), "xl/workbook.xml");

        AssertFailure(VbaExtraction.Check(_app, "data/Plain.xlsx").Failure, "INVALID_REQUEST", "NOT_MACRO_ENABLED");
    }

    [Fact]
    public void RejectsAWorkbookWithExcel4MacroSheets()
    {
        Package(Path.Combine(_app.FullPath, "data", "Xlm.xlsm"), "xl/workbook.xml", "xl/macrosheets/sheet1.xml");

        AssertFailure(VbaExtraction.Check(_app, "data/Xlm.xlsm").Failure, "INVALID_REQUEST", "EXCEL4_MACROS");
    }

    [Fact]
    public void RejectsAFileThatIsNotAPackageAsOpenFailed()
    {
        File.WriteAllText(Path.Combine(_app.FullPath, "Old.xlsm"), "not a zip");

        AssertFailure(VbaExtraction.Check(_app, null).Failure, "EXCEL_ERROR", "OPEN_FAILED");
    }

    [Fact]
    public void RejectsSourcesThatAreAlreadyThereAndListsThem()
    {
        Source("src/modules/Calc.bas", "x");
        Source("src/workbook/Sheet1.cls", "x");

        var failure = VbaExtraction.Check(_app, null).Failure;

        AssertFailure(failure, "CONFLICT", "SOURCE_EXISTS");
        Assert.Equal("""["src/modules/Calc.bas","src/workbook/Sheet1.cls"]""", failure!.Data!["files"]!.ToJsonString());
    }

    [Fact]
    public void AcceptsDotFilesAndTheXlWrightModuleInSrc()
    {
        Source("src/modules/XlWright.bas", "x");
        Source("src/modules/.gitkeep", "");
        Source("src/classes/.gitkeep", "");

        Assert.Null(VbaExtraction.Check(_app, null).Failure);
    }

    [Fact]
    public void WritesEachModuleAsUtf8InTheFolderOfItsType()
    {
        var export = Directory.CreateDirectory(Path.Combine(_repo.Root, "export")).FullName;
        var modules = new[]
        {
            Exported(export, "Calc", "standard", "Attribute VB_Name = \"Calc\"\r\n' 日本語\r\n"),
            Exported(export, "Counter", "class", "VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Counter\"\r\n"),
            Exported(export, "Sheet1", "document", "VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Sheet1\"\r\n"),
            new ExportedModule("Form1", "form", null),
            Exported(export, "XlWright", "standard", "Attribute VB_Name = \"XlWright\"\r\n"),
        };

        var (files, skipped, failure) = VbaExtraction.Write(_app, modules, Ansi);

        Assert.Null(failure);
        Assert.Equal(
            """[{"file":"src/modules/Calc.bas","type":"standard"},{"file":"src/classes/Counter.cls","type":"class"},{"file":"src/workbook/Sheet1.cls","type":"document"}]""",
            files!.ToJsonString());
        Assert.Equal("""[{"name":"Form1","type":"form"}]""", skipped!.ToJsonString());
        var calc = File.ReadAllBytes(Path.Combine(_app.FullPath, "src", "modules", "Calc.bas"));
        Assert.Equal("Attribute VB_Name = \"Calc\"\r\n' 日本語\r\n", new UTF8Encoding(false, true).GetString(calc));
        Assert.NotEqual(0xEF, calc[0]);
        Assert.False(File.Exists(Path.Combine(_app.FullPath, "src", "modules", "XlWright.bas")));
        Assert.True(File.Exists(Path.Combine(_app.FullPath, "src", "classes", "Counter.cls")));
        Assert.True(File.Exists(Path.Combine(_app.FullPath, "src", "workbook", "Sheet1.cls")));
    }

    [Fact]
    public void ReportsTheFilesWrittenBeforeAFailure()
    {
        var export = Directory.CreateDirectory(Path.Combine(_repo.Root, "export")).FullName;
        var modules = new[]
        {
            Exported(export, "Calc", "standard", "Attribute VB_Name = \"Calc\"\r\n"),
            new ExportedModule("Gone", "standard", Path.Combine(export, "Gone.bas")),
        };

        var (files, _, failure) = VbaExtraction.Write(_app, modules, Ansi);

        Assert.Null(files);
        AssertFailure(failure, "EXCEL_ERROR", "EXTRACT_FAILED");
        Assert.Equal("""["src/modules/Calc.bas"]""", failure!.Data!["done"]!.ToJsonString());
        Assert.True(File.Exists(Path.Combine(_app.FullPath, "src", "modules", "Calc.bas")));
    }

    private static ExportedModule Exported(string folder, string name, string type, string text)
    {
        var path = Path.Combine(folder, name + (type == "standard" ? ".bas" : ".cls"));
        File.WriteAllBytes(path, Ansi.GetBytes(text));
        return new ExportedModule(name, type, path);
    }

    private void Source(string relative, string text)
    {
        var path = Path.Combine(_app.FullPath, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, text);
    }

    private static void Package(string path, params string[] entries)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.Delete(path);
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var entry in entries)
        {
            zip.CreateEntry(entry);
        }
    }

    private static void AssertFailure(ToolFailure? failure, string code, string detail)
    {
        Assert.NotNull(failure);
        Assert.Equal(code, failure.Code);
        Assert.Equal(detail, failure.Detail);
    }
}
