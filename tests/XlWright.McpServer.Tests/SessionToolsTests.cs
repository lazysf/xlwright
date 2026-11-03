using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// session_start, session_status, session_end, and workbook_save on a --no-excel server (docs/api.md, stage 1).
public sealed class SessionToolsTests : IAsyncLifetime
{
    private readonly TempRepository _repo = new();
    private McpClient? _client;

    private McpClient Client => _client ?? throw new InvalidOperationException("The client is not connected.");

    public async ValueTask InitializeAsync()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _repo.AddApp("apps/two", "A.xlsm", "B.xlsx");
        _client = await ServerProcess.ConnectAsync(_repo.Root, TestContext.Current.CancellationToken);
    }

    public async ValueTask DisposeAsync()
    {
        if (_client is not null)
        {
            await _client.DisposeAsync();
        }

        _repo.Dispose();
    }

    [Theory]
    [InlineData("session_start", false, false, false, false, "app")]
    [InlineData("session_status", true, false, true, false, null)]
    [InlineData("session_end", false, true, false, false, "save")]
    [InlineData("workbook_save", false, true, true, false, null)]
    public async Task ListsTheToolWithItsAnnotationsAndSchemas(
        string name, bool readOnly, bool destructive, bool idempotent, bool openWorld, string? required)
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
        var properties = tool.InputSchema.GetProperty("properties").EnumerateObject().Select(p => p.Name).ToArray();
        Assert.Equal(required is null ? [] : [required], properties);
        if (required is not null)
        {
            Assert.Equal(required, Assert.Single(tool.InputSchema.GetProperty("required").EnumerateArray()).GetString());
        }
    }

    [Fact]
    public async Task StatusReportsAClosedSession()
    {
        var result = await Call("session_status", new());

        Assert.NotEqual(true, result.IsError);
        Assert.NotNull(result.StructuredContent);
        var structured = result.StructuredContent.Value;
        AssertSession(structured, "closed");
        Assert.Single(structured.EnumerateObject());
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        Assert.True(JsonElement.DeepEquals(structured, body.RootElement));
    }

    [Fact]
    public async Task StartChecksTheAppFolderThenReportsNotImplementedWithoutExcel()
    {
        var error = await Failure("session_start", new() { ["app"] = "apps/sales" });

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.Equal("NOT_IMPLEMENTED", error.GetProperty("detail").GetString());
    }

    [Theory]
    [InlineData("../outside", "INVALID_REQUEST", "PATH_NOT_ALLOWED")]
    [InlineData(@"\\server\share\app", "INVALID_REQUEST", "PATH_NOT_ALLOWED")]
    [InlineData("apps/missing", "NOT_FOUND", "APP_NOT_FOUND")]
    [InlineData("apps", "INVALID_REQUEST", "NO_WORKBOOK")]
    [InlineData("apps/two", "INVALID_REQUEST", "MULTIPLE_WORKBOOKS")]
    public async Task StartRejectsABadAppFolder(string app, string code, string detail)
    {
        var error = await Failure("session_start", new() { ["app"] = app });

        Assert.Equal(code, error.GetProperty("code").GetString());
        Assert.Equal(detail, error.GetProperty("detail").GetString());
    }

    [Fact]
    public async Task StartNamesTheWorkbooksWhenThereAreTwo()
    {
        var error = await Failure("session_start", new() { ["app"] = "apps/two" });

        var names = error.GetProperty("data").GetProperty("workbooks").EnumerateArray().Select(n => n.GetString()).ToArray();
        Assert.Equal(["A.xlsm", "B.xlsx"], names);
    }

    public static TheoryData<string, Dictionary<string, object?>> InvalidArguments => new()
    {
        { "session_start", new() },
        { "session_start", new() { ["app"] = 3 } },
        { "session_start", new() { ["app"] = "apps/sales", ["workbook"] = "Sales.xlsm" } },
        { "session_status", new() { ["verbose"] = true } },
        { "session_end", new() },
        { "session_end", new() { ["save"] = "yes" } },
        { "session_end", new() { ["save"] = true, ["force"] = true } },
        { "workbook_save", new() { ["path"] = "Other.xlsm" } },
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public async Task RejectsInvalidArgumentsWithoutDetail(string tool, Dictionary<string, object?> arguments)
    {
        var error = await Failure(tool, arguments);

        Assert.Equal("INVALID_REQUEST", error.GetProperty("code").GetString());
        Assert.False(error.TryGetProperty("detail", out _));
    }

    [Fact]
    public async Task EndReportsThatThereIsNoSession()
    {
        var error = await Failure("session_end", new() { ["save"] = false });

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    [Fact]
    public async Task SaveReportsThatThereIsNoSession()
    {
        var error = await Failure("workbook_save", new());

        Assert.Equal("CONFLICT", error.GetProperty("code").GetString());
        Assert.Equal("NO_SESSION", error.GetProperty("detail").GetString());
    }

    private Task<CallToolResult> Call(string tool, Dictionary<string, object?> arguments) =>
        Client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken).AsTask();

    // Returns the error object after checking the shared failure shape: isError, no structured content,
    // one line of JSON, and a closed, idle session.
    private async Task<JsonElement> Failure(string tool, Dictionary<string, object?> arguments)
    {
        var result = await Call(tool, arguments);

        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        Assert.DoesNotContain('\n', text.Text);
        using var body = JsonDocument.Parse(text.Text);
        AssertSession(body.RootElement, "closed");
        var error = body.RootElement.GetProperty("error");
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("message").GetString()));
        return error.Clone();
    }

    private static void AssertSession(JsonElement body, string state)
    {
        var session = body.GetProperty("session");
        Assert.Equal(state, session.GetProperty("state").GetString());
        Assert.False(session.GetProperty("busy").GetBoolean());
    }
}
