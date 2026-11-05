using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// workbook_create and the session without a main workbook on a --no-excel server (docs/api.md, stage 4).
[Collection(ServerProcess.Collection)]
public sealed class WorkbookToolsTests : IAsyncLifetime
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
    public async Task ListsWorkbookCreateWithItsAnnotationsAndSchemas()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var tool = Assert.Single(tools, t => t.Name == "workbook_create").ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.False(tool.Annotations.ReadOnlyHint);
        Assert.False(tool.Annotations.DestructiveHint);
        Assert.False(tool.Annotations.IdempotentHint);
        Assert.False(tool.Annotations.OpenWorldHint);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["name"], tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.Equal(["name"], tool.InputSchema.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
        Assert.Equal(["session", "workbook"], tool.OutputSchema!.Value.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
    }

    [Theory]
    [InlineData("session_start")]
    [InlineData("session_status")]
    public async Task TheWorkbookOfStartAndStatusCanBeNull(string name)
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var workbook = Assert.Single(tools, t => t.Name == name).ProtocolTool.OutputSchema!.Value.GetProperty("properties").GetProperty("workbook");
        Assert.Equal(["string", "null"], workbook.GetProperty("type").EnumerateArray().Select(e => e.GetString()));
    }

    public static TheoryData<Dictionary<string, object?>> InvalidArguments => new()
    {
        new Dictionary<string, object?>(),
        new Dictionary<string, object?> { ["name"] = 1 },
        new Dictionary<string, object?> { ["name"] = "New.xlsm", ["format"] = "xlsm" },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    [Fact]
    public async Task ReportsThatThereIsNoSession()
    {
        var error = await Failure(new() { ["name"] = "New.xlsm" });

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    [Fact]
    public async Task ListsWorkbookGetAndEditWithTheirAnnotationsAndSchemas()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var get = Assert.Single(tools, t => t.Name == "workbook_get").ProtocolTool;
        Assert.NotNull(get.Annotations);
        Assert.True(get.Annotations.ReadOnlyHint);
        Assert.False(get.Annotations.DestructiveHint);
        Assert.True(get.Annotations.IdempotentHint);
        Assert.False(get.Annotations.OpenWorldHint);
        Assert.Empty(get.InputSchema.GetProperty("properties").EnumerateObject());
        Assert.Equal(["session", "workbook", "format", "saved"], get.OutputSchema!.Value.GetProperty("required").EnumerateArray().Select(e => e.GetString()));

        var edit = Assert.Single(tools, t => t.Name == "workbook_edit").ProtocolTool;
        Assert.NotNull(edit.Annotations);
        Assert.False(edit.Annotations.ReadOnlyHint);
        Assert.False(edit.Annotations.DestructiveHint);
        Assert.False(edit.Annotations.IdempotentHint);
        Assert.False(edit.Annotations.OpenWorldHint);
        Assert.Contains("MULTIPLE_WORKBOOKS", edit.Description);
        Assert.False(edit.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["convert_to_xlsm"], edit.InputSchema.GetProperty("properties").GetProperty("action").GetProperty("enum").EnumerateArray().Select(e => e.GetString()));
        Assert.Equal(["action"], edit.InputSchema.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
        Assert.Equal(["session", "workbook", "previous"], edit.OutputSchema!.Value.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
    }

    public static TheoryData<string, Dictionary<string, object?>> InvalidGetAndEditArguments => new()
    {
        { "workbook_get", new Dictionary<string, object?> { ["name"] = "Sales.xlsx" } },
        { "workbook_edit", new Dictionary<string, object?>() },
        { "workbook_edit", new Dictionary<string, object?> { ["action"] = 1 } },
        { "workbook_edit", new Dictionary<string, object?> { ["action"] = "protect" } },
        { "workbook_edit", new Dictionary<string, object?> { ["action"] = "convert_to_xlsm", ["name"] = "Sales.xlsm" } },
    };

    [Theory]
    [MemberData(nameof(InvalidGetAndEditArguments))]
    public async Task GetAndEditRejectInvalidArgumentsWithoutDetail(string tool, Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments, tool);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    [Theory]
    [InlineData("workbook_get", null)]
    [InlineData("workbook_edit", "convert_to_xlsm")]
    public async Task GetAndEditReportThatThereIsNoSession(string tool, string? action)
    {
        var error = await Failure(action is null ? new() : new() { ["action"] = action }, tool);

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private async Task<JsonElement> Failure(Dictionary<string, object?> arguments, string tool = "workbook_create")
    {
        var result = await Client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        return body.RootElement.GetProperty("error").Clone();
    }
}
