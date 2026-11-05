using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// vba_edit and the references of vba_get on a --no-excel server (docs/api.md, stage 4). The argument checks run before
// the session check.
[Collection(ServerProcess.Collection)]
public sealed class ReferenceToolsTests : IAsyncLifetime
{
    private const string ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}";

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
    public async Task ListsVbaEditWithItsAnnotationsAndSchemas()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var tool = Assert.Single(tools, t => t.Name == "vba_edit").ProtocolTool;
        Assert.NotNull(tool.Annotations);
        Assert.False(tool.Annotations.ReadOnlyHint);
        Assert.True(tool.Annotations.DestructiveHint);
        Assert.True(tool.Annotations.IdempotentHint);
        Assert.False(tool.Annotations.OpenWorldHint);
        Assert.False(tool.InputSchema.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(["action", "guid", "major", "minor"], tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name));
        Assert.Equal(["action", "guid"], tool.InputSchema.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
        Assert.NotNull(tool.OutputSchema);
        Assert.Contains("references", tool.OutputSchema.Value.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
    }

    [Fact]
    public async Task VbaGetReturnsTheReferences()
    {
        var tools = await Client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        var schema = Assert.Single(tools, t => t.Name == "vba_get").ProtocolTool.OutputSchema!.Value;
        Assert.Contains("references", schema.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
        var reference = schema.GetProperty("properties").GetProperty("references").GetProperty("items");
        Assert.Equal(["guid", "major", "minor", "builtin", "broken"], reference.GetProperty("required").EnumerateArray().Select(e => e.GetString()));
    }

    public static TheoryData<Dictionary<string, object?>> InvalidArguments => new()
    {
        new Dictionary<string, object?> { ["guid"] = ScriptingGuid },
        new Dictionary<string, object?> { ["action"] = "rename_reference", ["guid"] = ScriptingGuid },
        new Dictionary<string, object?> { ["action"] = "remove_reference" },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = "420B2830-E718-11CF-893D-00A0C9054228" },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = "{420B2830}" },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = 1 },
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["minor"] = 0 },
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["major"] = 1 },
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["major"] = -1, ["minor"] = 0 },
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["major"] = "1", ["minor"] = 0 },
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["major"] = 1.5, ["minor"] = 0 },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = ScriptingGuid, ["major"] = 1 },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = ScriptingGuid, ["path"] = "C:/x.dll" },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    public static TheoryData<Dictionary<string, object?>> ValidArguments => new()
    {
        new Dictionary<string, object?> { ["action"] = "add_reference", ["guid"] = ScriptingGuid, ["major"] = 1, ["minor"] = 0 },
        new Dictionary<string, object?> { ["action"] = "remove_reference", ["guid"] = ScriptingGuid.ToLowerInvariant() },
    };

    [Theory]
    [MemberData(nameof(ValidArguments))]
    public async Task ReportsThatThereIsNoSession(Dictionary<string, object?> arguments)
    {
        var error = await Failure(arguments);

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private async Task<JsonElement> Failure(Dictionary<string, object?> arguments)
    {
        var result = await Client.CallToolAsync("vba_edit", arguments, cancellationToken: TestContext.Current.CancellationToken);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.Equal("closed", body.RootElement.GetProperty("session").GetProperty("state").GetString());
        return body.RootElement.GetProperty("error").Clone();
    }
}
