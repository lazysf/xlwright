using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// range_read and range_write on a --no-excel server (docs/api.md, stage 1). The argument checks run before the
// session check, so every failure but NO_SESSION is reachable without Excel.
[Collection(ServerProcess.Collection)]
public sealed class RangeToolsTests : IAsyncLifetime
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

    [Theory]
    [InlineData("range_read", true, false, true, false, new[] { "sheet", "range", "content" }, new[] { "sheet", "range" })]
    [InlineData("range_write", false, true, true, false, new[] { "sheet", "range", "values", "formulas", "clear" }, new[] { "sheet", "range" })]
    public async Task ListsTheToolWithItsAnnotationsAndSchemas(
        string name, bool readOnly, bool destructive, bool idempotent, bool openWorld, string[] properties, string[] required)
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var tool = Assert.Single(tools, t => t.Name == name).ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.Equal(readOnly, tool.Annotations.ReadOnlyHint);
        Assert.Equal(destructive, tool.Annotations.DestructiveHint);
        Assert.Equal(idempotent, tool.Annotations.IdempotentHint);
        Assert.Equal(openWorld, tool.Annotations.OpenWorldHint);
        Assert.NotNull(tool.OutputSchema);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(properties, tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.Equal(required, tool.InputSchema.GetProperty("required").EnumerateArray().Select(r => r.GetString()));
    }

    public static TheoryData<string, Dictionary<string, object?>> InvalidArguments => new()
    {
        { "range_read", new() { ["range"] = "A1" } },
        { "range_read", new() { ["sheet"] = "Data" } },
        { "range_read", new() { ["sheet"] = 1, ["range"] = "A1" } },
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A1", ["content"] = "text" } },
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A1", ["workbook"] = "Other.xlsm" } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1" } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["values"] = Grid(1), ["clear"] = true } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["clear"] = false } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["values"] = 1 } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["values"] = new object[] { 1 } } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["values"] = new[] { new object[] { new[] { 1 } } } } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["formulas"] = new[] { new object[] { 1 } } } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["clear"] = true, ["save"] = true } },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(string tool, Dictionary<string, object?> arguments)
    {
        var error = await Failure(tool, arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    public static TheoryData<string, Dictionary<string, object?>, string> RejectedBeforeTheSession => new()
    {
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A:A" }, "INVALID_RANGE" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "Data!A1", ["clear"] = true }, "INVALID_RANGE" },
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A1:CV101" }, "TOO_LARGE" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:B1", ["values"] = new[] { new object?[] { 1 } } }, "SHAPE_MISMATCH" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:B2", ["values"] = new[] { new object?[] { 1, 2 }, new object?[] { 3 } } }, "SHAPE_MISMATCH" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["values"] = Array.Empty<object[]>() }, "SHAPE_MISMATCH" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:B1", ["formulas"] = new[] { new[] { "=1" } } }, "SHAPE_MISMATCH" },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:B1", ["values"] = new[] { new object?[] { 1, "=A1" } } }, "VALUE_LOOKS_LIKE_FORMULA" },
    };

    [Theory]
    [MemberData(nameof(RejectedBeforeTheSession))]
    public async Task RejectsBeforeTheSessionCheck(string tool, Dictionary<string, object?> arguments, string detail)
    {
        var error = await Failure(tool, arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.Equal(detail, error.GetProperty("detail").GetString());
    }

    public static TheoryData<string, Dictionary<string, object?>> Valid => new()
    {
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A1:C2" } },
        { "range_read", new() { ["sheet"] = "Data", ["range"] = "A1", ["content"] = "formulas" } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:D1", ["values"] = new[] { new object?[] { 1.5, "a", true, null } } } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1", ["formulas"] = new[] { new[] { "=1+1" } } } },
        { "range_write", new() { ["sheet"] = "Data", ["range"] = "A1:XFD1048576", ["clear"] = true } },
    };

    [Theory]
    [MemberData(nameof(Valid))]
    public async Task ReportsThatThereIsNoSession(string tool, Dictionary<string, object?> arguments)
    {
        var error = await Failure(tool, arguments);

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private static object[][] Grid(int value) => [[value]];

    // Returns the error object after checking the shared failure shape: isError, no structured content,
    // one line of JSON, and a closed, idle session.
    private async Task<JsonElement> Failure(string tool, Dictionary<string, object?> arguments)
    {
        var result = await Client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        Assert.DoesNotContain('\n', text.Text);
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        var error = body.RootElement.GetProperty("error");
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("message").GetString()));
        return error.Clone();
    }
}
