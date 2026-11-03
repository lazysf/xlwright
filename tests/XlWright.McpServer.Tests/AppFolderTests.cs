using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "アプリのフォルダ" and "パスの制限".
public sealed class AppFolderTests : IDisposable
{
    private readonly TempRepository _repo = new();

    public void Dispose() => _repo.Dispose();

    [Fact]
    public void ResolvesARelativeFolderAndItsMainWorkbook()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");

        var app = Resolve("apps/sales");

        Assert.Equal(Path.Combine(_repo.Root, "apps", "sales"), app.FullPath);
        Assert.Equal("apps/sales", app.RelativePath);
        Assert.Equal("Sales.xlsm", app.WorkbookName);
        Assert.Equal(Path.Combine(_repo.Root, "apps", "sales", "Sales.xlsm"), app.WorkbookPath);
    }

    [Fact]
    public void ResolvesAnAbsoluteFolderWithDifferentCase()
    {
        _repo.AddApp("apps/sales", "Sales.xlsx");

        var app = Resolve(Path.Combine(_repo.Root, "APPS", "Sales"));

        Assert.Equal("apps/sales", app.RelativePath, ignoreCase: true);
        Assert.Equal("Sales.xlsx", app.WorkbookName);
    }

    [Fact]
    public void AcceptsTheRepositoryRootAsTheAppFolder()
    {
        _repo.AddApp(".", "Book.xlsm");

        var app = Resolve(".");

        Assert.Equal(".", app.RelativePath);
        Assert.Equal("Book.xlsm", app.WorkbookName);
    }

    [Theory]
    [InlineData("..")]
    [InlineData("../outside")]
    [InlineData(@"\\server\share\apps\sales")]
    [InlineData(@"\\?\C:\apps\sales")]
    [InlineData(@"\\.\C:\apps\sales")]
    [InlineData("//server/share/apps")]
    public void RejectsPathsOutsideTheRepositoryOrNotAllowed(string app)
    {
        var failure = Fail(app);

        Assert.Equal("INVALID_REQUEST", failure.Code);
        Assert.Equal("PATH_NOT_ALLOWED", failure.Detail);
    }

    [Fact]
    public void RejectsASiblingFolderThatSharesTheRepositoryPrefix()
    {
        var sibling = _repo.Root + "2";
        Directory.CreateDirectory(sibling);
        try
        {
            var failure = Fail(sibling);

            Assert.Equal("PATH_NOT_ALLOWED", failure.Detail);
        }
        finally
        {
            Directory.Delete(sibling);
        }
    }

    [Fact]
    public void ReportsAMissingFolder()
    {
        var failure = Fail("apps/missing");

        Assert.Equal("NOT_FOUND", failure.Code);
        Assert.Equal("APP_NOT_FOUND", failure.Detail);
    }

    [Fact]
    public void RejectsAFolderWithoutAWorkbook()
    {
        _repo.AddApp("apps/empty", "notes.txt", "data.csv", "Old.xls");

        var failure = Fail("apps/empty");

        Assert.Equal("INVALID_REQUEST", failure.Code);
        Assert.Equal("NO_WORKBOOK", failure.Detail);
    }

    [Fact]
    public void IgnoresOwnerFilesAndHiddenFiles()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm", "~$Sales.xlsm", "Hidden.xlsx");
        File.SetAttributes(Path.Combine(_repo.Root, "apps", "sales", "Hidden.xlsx"), FileAttributes.Hidden);

        var app = Resolve("apps/sales");

        Assert.Equal("Sales.xlsm", app.WorkbookName);
    }

    [Fact]
    public void IgnoresWorkbooksInSubfolders()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm", "data/Input.xlsx");

        Assert.Equal("Sales.xlsm", Resolve("apps/sales").WorkbookName);
    }

    [Fact]
    public void RejectsTwoWorkbooksAndNamesThem()
    {
        _repo.AddApp("apps/sales", "Sales.xlsx", "Sales.xlsm");

        var failure = Fail("apps/sales");

        Assert.Equal("INVALID_REQUEST", failure.Code);
        Assert.Equal("MULTIPLE_WORKBOOKS", failure.Detail);
        Assert.NotNull(failure.Data);
        var names = failure.Data["workbooks"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray();
        Assert.Equal(["Sales.xlsm", "Sales.xlsx"], names);
    }

    [Fact]
    public void RejectsAnEmptyPath()
    {
        var failure = Fail("");

        Assert.Equal("INVALID_REQUEST", failure.Code);
        Assert.Null(failure.Detail);
    }

    private AppFolder Resolve(string app)
    {
        var result = AppFolder.Resolve(_repo.Root, app);
        Assert.Null(result.Failure);
        return Assert.IsType<AppFolder>(result.Folder);
    }

    private ToolFailure Fail(string app)
    {
        var result = AppFolder.Resolve(_repo.Root, app);
        Assert.Null(result.Folder);
        return Assert.IsType<ToolFailure>(result.Failure);
    }
}
