using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// These tests start the built server as a real process and talk to it over stdio.
public sealed class ServerProcessTests
{
    private static readonly string BuildDirectory = new DirectoryInfo(AppContext.BaseDirectory).Name;
    private static readonly string ServerPath = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "XlWright.McpServer", BuildDirectory, "XlWright.McpServer.dll"));

    private static ProcessStartInfo Launch(params string[] arguments)
    {
        var info = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetTempPath(),
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        info.ArgumentList.Add(ServerPath);
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        return info;
    }

    private static Task<McpClient> Connect(CancellationToken cancellation, params string[] arguments) => McpClient.CreateAsync(
        new StdioClientTransport(new()
        {
            Name = "xlwright-test", Command = Environment.ProcessPath!,
            Arguments = [ServerPath, .. arguments], WorkingDirectory = Path.GetTempPath()
        }),
        cancellationToken: cancellation);

    // A hung server must fail the test, not hold the integrated check open.
    private static async Task WaitForExit(Process process)
    {
        try
        {
            await process.WaitForExitAsync(TestContext.Current.CancellationToken)
                .WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken);
        }
        finally
        {
            if (!process.HasExited) process.Kill();
        }
    }

    [Theory(Timeout = 60000)]
    [InlineData]
    [InlineData("--development")]
    public async Task ExitsOnEofWithoutProtocolOutput(params string[] arguments)
    {
        using var process = Process.Start(Launch(arguments))!;
        Task<string> stdout = process.StandardOutput.ReadToEndAsync(TestContext.Current.CancellationToken);
        Task<string> stderr = process.StandardError.ReadToEndAsync(TestContext.Current.CancellationToken);

        process.StandardInput.Close();
        await WaitForExit(process);

        Assert.Equal(0, process.ExitCode);
        Assert.Equal("", await stdout);
        await stderr;
    }

    [Fact(Timeout = 60000)]
    public async Task RejectsUnknownArguments()
    {
        using var process = Process.Start(Launch("--unknown"))!;
        Task<string> stderr = process.StandardError.ReadToEndAsync(TestContext.Current.CancellationToken);

        await WaitForExit(process);

        Assert.Equal(1, process.ExitCode);
        Assert.Contains("Usage", await stderr);
    }

    [Fact(Timeout = 60000)]
    public async Task NormalModeAdvertisesNoTools()
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken);

        Assert.Null(client.ServerCapabilities.Tools);
    }

    [Fact(Timeout = 60000)]
    public async Task DevelopmentModePublishesOnlyStatus()
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");

        var tools = await client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal("dev_status", Assert.Single(tools).Name);
    }

    [Fact(Timeout = 60000)]
    public async Task StatusIdentifiesTheRunningBuild()
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");

        CallToolResult result = await client.CallToolAsync("dev_status",
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.NotEqual(true, result.IsError);
        using var status = JsonDocument.Parse(Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(ServerPath,
            TestContext.Current.CancellationToken))), status.RootElement.GetProperty("build").GetString());
        Assert.False(string.IsNullOrEmpty(status.RootElement.GetProperty("invocation").GetString()));
    }
}
