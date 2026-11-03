using Xunit;

namespace XlWright.McpServer.Tests;

public sealed class ServerArgumentsTests
{
    [Theory]
    [InlineData("--unknown")]
    [InlineData("--project")]
    public async Task ExitsWithUsageErrorOnInvalidArguments(string argument)
    {
        var (exitCode, standardError) = await ServerProcess.RunAsync(
            [argument],
            TestContext.Current.CancellationToken);

        Assert.Equal(2, exitCode);
        Assert.Contains(argument, standardError);
    }
}
