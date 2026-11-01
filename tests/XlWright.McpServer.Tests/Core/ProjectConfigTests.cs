using System.Diagnostics;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

public sealed class ProjectConfigTests : IDisposable
{
    private readonly string root = Directory.CreateTempSubdirectory("xlwright-config-").FullName;

    public ProjectConfigTests()
    {
        File.WriteAllText(Path.Combine(root, "book.xlsm"), "");
        Directory.CreateDirectory(Path.Combine(root, "src"));
        Directory.CreateDirectory(Path.Combine(root, "tests"));
    }

    public void Dispose() => Directory.Delete(root, recursive: true);

    private ProjectConfig Load(string text)
    {
        File.WriteAllText(Path.Combine(root, "project.json"), text);
        return ProjectConfig.Load(root);
    }

    [Fact]
    public void DefaultsResolveFromTheProjectRoot()
    {
        var config = Load("""{"paths":{"workbook":"book.xlsm"}}""");

        Assert.Equal(Path.Combine(root, "book.xlsm"), config.Workbook);
        Assert.Equal(Path.Combine(root, "src"), config.VbaSource);
        Assert.Equal(Path.Combine(root, "tests"), config.Tests);
        Assert.Equal(Path.Combine(root, ".xlwright"), config.Workspace);
        Assert.Equal(Path.Combine(root, "dist"), config.Dist);
    }

    [Fact]
    public void MissingFileIsNotFound()
    {
        var error = Assert.Throws<CoreException>(() => ProjectConfig.Load(root));

        Assert.Equal(ErrorCode.NotFound, error.Code);
        Assert.Equal(ErrorDetail.ConfigNotFound, error.Detail);
    }

    [Theory]
    [InlineData("{")]
    [InlineData("[]")]
    [InlineData("{}")]
    [InlineData("""{"paths":{}}""")]
    [InlineData("""{"paths":{"workbook":1}}""")]
    [InlineData("""{"paths":{"workbook":""}}""")]
    [InlineData("""{"paths":{"workbook":"book.xls"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","unknown":"x"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm"},"unknown":1}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm",}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","workbook":"book.xlsm"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm"},"paths":{"workbook":"book.xlsm"}}""")]
    [InlineData("""{/* c */"paths":{"workbook":"book.xlsm"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","workspace":"src"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","workspace":"src/out"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","dist":"."}}""")]
    [InlineData("""{"paths":{"workbook":"\\\\server.invalid\\share\\book.xlsm"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","vba_source":"//server.invalid/share"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","tests":"\\\\server.invalid\\share"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","workspace":"\\\\?\\C:\\out"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","dist":"\\\\.\\C:\\out"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","workspace":"src::$INDEX_ALLOCATION"}}""")]
    public void InvalidConfigurationsAreRefused(string text)
    {
        var error = Assert.Throws<CoreException>(() => Load(text));

        Assert.Equal(ErrorCode.InvalidRequest, error.Code);
        Assert.Equal(ErrorDetail.InvalidConfig, error.Detail);
    }

    [Theory]
    [InlineData("""{"paths":{"workbook":"none.xlsx"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","vba_source":"none"}}""")]
    [InlineData("""{"paths":{"workbook":"book.xlsm","tests":"none"}}""")]
    public void MissingPathsAreNotFound(string text)
    {
        var error = Assert.Throws<CoreException>(() => Load(text));

        Assert.Equal(ErrorDetail.PathNotFound, error.Detail);
    }

    private void CreateJunction(string link, string target)
    {
        var info = new ProcessStartInfo("cmd.exe") { UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in new[] { "/c", "mklink", "/J", link, target }) info.ArgumentList.Add(argument);
        info.RedirectStandardOutput = true;
        using var process = Process.Start(info)!;
        process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        Assert.Equal(0, process.ExitCode);
    }

    [Theory]
    [InlineData("out", """{"paths":{"workbook":"book.xlsm","workspace":"out"}}""")]
    [InlineData(".xlwright/logs", """{"paths":{"workbook":"book.xlsm"}}""")]
    public void LinksCannotHideAnOverlap(string link, string text)
    {
        string path = Path.GetFullPath(Path.Combine(root, link));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        CreateJunction(path, Path.Combine(root, "src"));

        var error = Assert.Throws<CoreException>(() => Load(text));

        Assert.Equal(ErrorDetail.InvalidConfig, error.Detail);
        Directory.Delete(path);
    }

    [Fact]
    public void ADriveRootOverlapsEverything()
    {
        string drive = Path.GetPathRoot(root)!.Replace("\\", "\\\\");

        var error = Assert.Throws<CoreException>(() => Load("{\"paths\":{\"workbook\":\"book.xlsm\",\"dist\":\"" + drive + "\"}}"));

        Assert.Equal(ErrorDetail.InvalidConfig, error.Detail);
    }

    [Fact]
    public void WorkspaceBesideTheProjectIsAllowed()
    {
        var config = Load("""{"paths":{"workbook":"book.xlsm","workspace":"../outside"}}""");

        Assert.Equal(Path.GetFullPath(Path.Combine(root, "..", "outside")), config.Workspace);
    }
}
