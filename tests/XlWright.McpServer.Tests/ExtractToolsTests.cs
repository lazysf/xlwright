using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// vba_extract on a --no-excel server (docs/api.md, stage 4). The argument checks run before the session check.
public sealed class ExtractToolsTests : IAsyncLifetime
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

        var tool = Assert.Single(tools, t => t.Name == "vba_extract").ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.False(tool.Annotations.ReadOnlyHint);
        Assert.False(tool.Annotations.DestructiveHint);
        Assert.False(tool.Annotations.IdempotentHint);
        Assert.False(tool.Annotations.OpenWorldHint);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["workbook"], tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.False(tool.InputSchema.TryGetProperty("required", out _));
        Assert.Equal(["session", "workbook", "files", "skipped"],
            tool.OutputSchema!.Value.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
    }

    public static TheoryData<Dictionary<string, object?>> InvalidArguments => new()
    {
        new Dictionary<string, object?> { ["workbook"] = 1 },
        new Dictionary<string, object?> { ["workbook"] = "" },
        new Dictionary<string, object?> { ["overwrite"] = true },
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
    [InlineData(null)]
    [InlineData("data/Legacy.xlsm")]
    public async Task ReportsThatThereIsNoSession(string? workbook)
    {
        var error = await Failure(workbook is null ? new() : new() { ["workbook"] = workbook });

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private async Task<JsonElement> Failure(Dictionary<string, object?> arguments)
    {
        var result = await Client.CallToolAsync("vba_extract", arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        return body.RootElement.GetProperty("error").Clone();
    }
}
