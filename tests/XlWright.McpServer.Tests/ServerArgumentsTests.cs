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
}
