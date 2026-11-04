using Xunit;

namespace XlWright.McpServer.Tests;

public sealed class ServerArgumentsTests
{
    public static TheoryData<string[], string> InvalidArguments => new()
    {
        { ["--unknown"], "--unknown" },
        { ["--project"], "--project" },
        { ["--project", "--no-excel"], "--project" },
        { ["--project", Directory.GetCurrentDirectory(), "--project", Directory.GetCurrentDirectory()], "--project" },
        { ["--no-excel"], "--project" },
        { ["--project", Path.Combine(Path.GetTempPath(), "xlwright-tests", "missing-" + Guid.NewGuid().ToString("N"))], "--project" },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task ExitsWithUsageErrorOnInvalidArguments(string[] arguments, string named)
    {
        var (exitCode, standardError) = await ServerProcess.RunAsync(
            arguments,
            TestContext.Current.CancellationToken);

        Assert.Equal(2, exitCode);
        Assert.Contains(named, standardError);
    }

    // docs/api.md "表示要求のコマンド".
    public static TheoryData<string[], string> InvalidShowArguments => new()
    {
        { ["show"], "--project" },
        { ["show", "--project", Directory.GetCurrentDirectory()], "--app" },
        { ["show", "--project", Directory.GetCurrentDirectory(), "--app"], "--app" },
        { ["show", "--project", Directory.GetCurrentDirectory(), "--app", "apps/sales", "--no-excel"], "--no-excel" },
        { ["show", "--project", Directory.GetCurrentDirectory(), "--app", Path.GetTempPath()], "PATH_NOT_ALLOWED" },
    };

    [Theory]
    [MemberData(nameof(InvalidShowArguments))]
    public async Task ShowExitsWithUsageErrorOnInvalidArguments(string[] arguments, string named)
    {
        var (exitCode, standardError) = await ServerProcess.RunAsync(arguments, TestContext.Current.CancellationToken);

        Assert.Equal(2, exitCode);
        Assert.Contains(named, standardError);
    }

    [Fact]
    public async Task ShowEndsWithoutCreatingARecordWhenThereIsNone()
    {
        using var repo = new TempRepository();
        repo.AddApp("apps/sales", "Sales.xlsm");

        var (exitCode, _) = await ServerProcess.RunAsync(["show", "--project", repo.Root, "--app", "apps/sales"],
            TestContext.Current.CancellationToken);

        Assert.Equal(0, exitCode);
        Assert.False(File.Exists(Path.Combine(repo.Root, "apps", "sales", ".xlwright", "session.json")));
    }
}
