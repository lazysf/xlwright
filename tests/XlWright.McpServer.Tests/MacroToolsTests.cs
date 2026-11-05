using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// macro_run on a --no-excel server (docs/api.md, stage 1). The argument checks run before the session check.
[Collection(ServerProcess.Collection)]
public sealed class MacroToolsTests : IAsyncLifetime
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

        var tool = Assert.Single(tools, t => t.Name == "macro_run").ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.False(tool.Annotations.ReadOnlyHint);
        Assert.True(tool.Annotations.DestructiveHint);
        Assert.False(tool.Annotations.IdempotentHint);
        Assert.True(tool.Annotations.OpenWorldHint);
        Assert.NotNull(tool.OutputSchema);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["macro", "args", "dialogs", "timeout_seconds"], tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.Equal(["macro"], tool.InputSchema.GetProperty("required").EnumerateArray().Select(r => r.GetString()));
    }

    public static TheoryData<Dictionary<string, object?>> InvalidArguments => new()
    {
        new Dictionary<string, object?>(),
        new Dictionary<string, object?> { ["macro"] = 1 },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["args"] = 1 },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["args"] = new object?[] { null } },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["args"] = new object[] { new[] { 1 } } },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["args"] = Enumerable.Range(1, 31).Cast<object>().ToArray() },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["timeout_seconds"] = 0 },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["timeout_seconds"] = 1201 },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["timeout_seconds"] = 1.5 },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["timeout_seconds"] = "300" },
        new Dictionary<string, object?> { ["macro"] = "Calc.Twice", ["save"] = true },
        Dialogs("message"),
        Dialogs(new object[] { "yes" }),
        Dialogs(Enumerable.Range(1, 31).Select(_ => (object)new Dictionary<string, object> { ["kind"] = "message", ["answer"] = "ok" }).ToArray()),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "color", ["value"] = "red" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["answer"] = "yes" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "message" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "message", ["answer"] = "maybe" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "message", ["value"] = "yes" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "input", ["value"] = 1 } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "input", ["value"] = "a", ["cancel"] = true } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "folder", ["cancel"] = false } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "open_file", ["path"] = "" } }),
        Dialogs(new object[] { new Dictionary<string, object> { ["kind"] = "save_file", ["path"] = "out.csv", ["note"] = "x" } }),
    };

    private static Dictionary<string, object?> Dialogs(object dialogs) => new() { ["macro"] = "Calc.Twice", ["dialogs"] = dialogs };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    [Theory]
    [InlineData("'Sales.xlsm'!Calc.Twice")]
    [InlineData("Twice")]
    public async Task RejectsAQualifiedOrBareName(string macro)
    {
        var error = await Failure(new() { ["macro"] = macro });

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.Equal("INVALID_MACRO_NAME", error.GetProperty("detail").GetString());
    }

    [Fact]
    public async Task ReportsThatThereIsNoSession()
    {
        var error = await Failure(new()
        {
            ["macro"] = "Calc.Twice",
            ["args"] = Enumerable.Range(1, 30).Select(i => (object)(i % 3 == 0 ? "a" : i % 3 == 1 ? i : true)).ToArray(),
            ["timeout_seconds"] = 1200,
            ["dialogs"] = new object[]
            {
                new Dictionary<string, object> { ["kind"] = "message", ["answer"] = "retry" },
                new Dictionary<string, object> { ["kind"] = "input", ["value"] = "東京" },
                new Dictionary<string, object> { ["kind"] = "open_file", ["path"] = "data/in.csv" },
                new Dictionary<string, object> { ["kind"] = "save_file", ["cancel"] = true },
                new Dictionary<string, object> { ["kind"] = "folder", ["path"] = "data" },
            },
        });

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private async Task<JsonElement> Failure(Dictionary<string, object?> arguments)
    {
        var result = await Client.CallToolAsync("macro_run", arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        return body.RootElement.GetProperty("error").Clone();
    }
}
