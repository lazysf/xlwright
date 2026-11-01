using System.Security.Cryptography;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.RealExcel.Tests;

// Keeps one connection across requests, which the Inspector CLI cannot do. This scenario does not reach Excel.
public sealed class CopiedServerTests(PreparedCopy copy)
{
    [Fact(Timeout = 60000)]
    public async Task ConnectionToTheCopiedServerSpansRequests()
    {
        // Development mode is used only because it has a read-only tool; product tools arrive in T19.
        await using McpClient client = await McpClient.CreateAsync(new StdioClientTransport(new()
        {
            Name = "xlwright-real-excel", Command = Environment.ProcessPath!,
            Arguments = [copy.ServerPath, "--development"], WorkingDirectory = copy.Repository
        }), cancellationToken: TestContext.Current.CancellationToken);

        var first = await Status(client);
        var second = await Status(client);

        Assert.Equal(first, second);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(copy.ServerPath,
            TestContext.Current.CancellationToken))), first.Build);
    }

    private static async Task<(string? Invocation, string? Build)> Status(McpClient client)
    {
        CallToolResult result = await client.CallToolAsync("dev_status", cancellationToken: TestContext.Current.CancellationToken);
        Assert.NotEqual(true, result.IsError);
        using var status = JsonDocument.Parse(Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        return (status.RootElement.GetProperty("invocation").GetString(), status.RootElement.GetProperty("build").GetString());
    }
}
