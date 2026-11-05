using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// test_run on a --no-excel server (docs/api.md, stage 3). The argument checks run before the session check.
[Collection(ServerProcess.Collection)]
public sealed class TestToolsTests : IAsyncLifetime
{
    private readonly TempRepository _repo = new();
    private McpClient? _client;

    private McpClient Client => _client ?? throw new InvalidOperationException("The client is not connected.");

    public async ValueTask InitializeAsync() =>
        _client = await ServerProcess.ConnectAsync(_repo.Root, TestContext.Current.CancellationToken);

    public async ValueTask DisposeAsync()
    {
        if (_client is not null)
        {
            await _client.DisposeAsync();
        }

        _repo.Dispose();
    }

    [Fact]
    public async Task ListsTheToolWithItsAnnotationsAndSchemas()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var tool = Assert.Single(tools, t => t.Name == "test_run").ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.False(tool.Annotations.ReadOnlyHint);
        Assert.True(tool.Annotations.DestructiveHint);
        Assert.False(tool.Annotations.IdempotentHint);
        Assert.True(tool.Annotations.OpenWorldHint);
        Assert.Contains("COMPILE_ERROR", tool.Description);
        Assert.Contains("person's screen", tool.Description);
        Assert.NotNull(tool.OutputSchema);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["tests", "timeout_seconds"], tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.False(tool.InputSchema.TryGetProperty("required", out _));
    }

    public static TheoryData<Dictionary<string, object?>> InvalidArguments => new()
    {
        new Dictionary<string, object?> { ["tests"] = "Tests.TestAdd" },
        new Dictionary<string, object?> { ["tests"] = new object[] { 1 } },
        new Dictionary<string, object?> { ["tests"] = Array.Empty<string>() },
        new Dictionary<string, object?> { ["timeout_seconds"] = 0 },
        new Dictionary<string, object?> { ["timeout_seconds"] = 1201 },
        new Dictionary<string, object?> { ["timeout_seconds"] = "300" },
        new Dictionary<string, object?> { ["save"] = true },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    [Theory]
    [InlineData("TestAdd")]
    [InlineData("'Sales.xlsm'!Tests.TestAdd")]
    public async Task RejectsATestNameThatIsNotModuleDotProcedure(string name)
    {
        var error = await Failure(new() { ["tests"] = new[] { "Tests.TestOk", name } });

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.Equal("INVALID_TEST_NAME", error.GetProperty("detail").GetString());
    }

    [Fact]
    public async Task ReportsThatThereIsNoSession()
    {
        var error = await Failure(new() { ["tests"] = new[] { "Tests.TestAdd" }, ["timeout_seconds"] = 1200 });

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private async Task<JsonElement> Failure(Dictionary<string, object?> arguments)
    {
        var result = await Client.CallToolAsync("test_run", arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        return body.RootElement.GetProperty("error").Clone();
    }
}
