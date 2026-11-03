using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// Starts the built server as a real process over stdio, the same way Claude Code does.
public sealed class ServerInfoTests : IAsyncLifetime
{
    private McpClient? _client;

    private McpClient Client => _client ?? throw new InvalidOperationException("The client is not connected.");

    public async ValueTask InitializeAsync()
    {
        _client = await ServerProcess.ConnectAsync(TestContext.Current.CancellationToken);
    }

    public async ValueTask DisposeAsync()
    {
        if (_client is not null)
        {
            await _client.DisposeAsync();
        }
    }

    [Fact]
    public async Task ListsOnlyServerInfoWithItsAnnotations()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var tool = Assert.Single(tools);
        Assert.Equal("server_info", tool.Name);
        var annotations = tool.ProtocolTool.Annotations;
        Assert.NotNull(annotations);
        Assert.True(annotations.ReadOnlyHint);
        Assert.False(annotations.DestructiveHint);
        Assert.True(annotations.IdempotentHint);
        Assert.False(annotations.OpenWorldHint);
        Assert.NotNull(tool.ProtocolTool.OutputSchema);
    }

    [Fact]
    public async Task ReturnsNameVersionAndBuildId()
    {
        var result = await Client.CallToolAsync(
            "server_info",
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.NotEqual(true, result.IsError);
        Assert.NotNull(result.StructuredContent);
        var structured = result.StructuredContent.Value;
        Assert.Equal("xlwright", structured.GetProperty("name").GetString());
        Assert.Equal("0.1.0", structured.GetProperty("version").GetString());
        Assert.Equal(ExpectedBuildId(), structured.GetProperty("build_id").GetString());
        Assert.Equal(3, structured.EnumerateObject().Count());

        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.True(JsonElement.DeepEquals(structured, body.RootElement));
    }

    [Fact]
    public async Task RejectsArgumentsAsInvalidRequest()
    {
        var result = await Client.CallToolAsync(
            "server_info",
            new Dictionary<string, object?> { ["verbose"] = true },
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        Assert.DoesNotContain('\n', text.Text);
        using var body = JsonDocument.Parse(text.Text);
        var error = body.RootElement.GetProperty("error");
        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.Contains("verbose", error.GetProperty("message").GetString());
        var session = body.RootElement.GetProperty("session");
        Assert.Equal("closed", session.GetProperty("state").GetString());
        Assert.False(session.GetProperty("busy").GetBoolean());
    }

    // The build embeds the id that the check entry point computed; a plain build embeds "unknown".
    private static string ExpectedBuildId()
    {
        var core = Assembly.Load("XlWright.Core");
        var buildId = core.GetCustomAttributes<AssemblyMetadataAttribute>()
            .Single(attribute => attribute.Key == "XlWrightBuildId")
            .Value;
        Assert.NotNull(buildId);
        Assert.Matches(new Regex("^(unknown|[0-9a-f]{7,}(-dirty)?-[0-9]{8}T[0-9]{6}Z)$"), buildId);
        return buildId;
    }
}
