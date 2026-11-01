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
    private static readonly string[] ProductToolNames = ["range_get", "range_set", "session_end", "session_save", "session_start"];
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
    public async Task NormalModePublishesOnlyProductTools()
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken);

        var tools = await client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(ProductToolNames.Order(), tools.Select(tool => tool.Name).Order());
    }

    [Fact(Timeout = 60000)]
    public async Task DevelopmentModeAddsStatus()
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");

        var tools = await client.ListToolsAsync(cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(ProductToolNames.Append("dev_status").Order(), tools.Select(tool => tool.Name).Order());
    }

    // Every call uses development mode, which cannot start Excel, or fails before reaching it.
    private static async Task<JsonElement> Refusal(McpClient client, string tool, Dictionary<string, object?>? arguments = null)
    {
        CallToolResult result = await client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken);
        Assert.True(result.IsError);
        Assert.Null(result.StructuredContent);
        string text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text;
        Assert.DoesNotContain("\n", text);
        return JsonDocument.Parse(text).RootElement;
    }

    private static void AssertRefused(JsonElement response, string code, string detail, string operation)
    {
        var error = response.GetProperty("error");
        Assert.Equal(code, error.GetProperty("code").GetString());
        Assert.Equal(detail, error.GetProperty("detail").GetString());
        Assert.Equal(operation, error.GetProperty("operation").GetString());
        foreach (string field in new[] { "stage", "excel_may_remain", "process", "log" }) Assert.True(error.TryGetProperty(field, out _));
        Assert.Equal("closed", response.GetProperty("session").GetProperty("state").GetString());
        Assert.False(response.GetProperty("session").GetProperty("busy").GetBoolean());
    }

    public static TheoryData<string, string, string> RefusedArguments => new()
    {
        { "session_start", "{}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":1}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"relative\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"\\\\\\\\server.invalid\\\\share\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"//server.invalid/share\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"\\\\\\\\?\\\\C:\\\\project\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"\\\\\\\\.\\\\C:\\\\project\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"C:\\\\x\",\"unknown\":1}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"C:\\\\x\",\"force\":\"yes\"}", "INVALID_ARGUMENT" },
        { "session_start", "{\"project_path\":\"C:\\\\x\",\"force\":true}", "NOT_IMPLEMENTED" },
        { "session_end", "{\"discard\":true}", "NOT_IMPLEMENTED" },
        { "session_save", "{\"unknown\":1}", "INVALID_ARGUMENT" },
        { "range_get", "{\"sheet\":\"S\"}", "INVALID_ARGUMENT" },
        { "range_get", "{\"sheet\":\"S\",\"address\":\"A:A\"}", "INVALID_ADDRESS" },
        { "range_get", "{\"sheet\":\"S\",\"address\":\"A1:B2\"}", "NOT_IMPLEMENTED" },
        { "range_get", "{\"sheet\":\"S\",\"address\":\"A1\",\"fields\":[\"merged_ranges\"]}", "NOT_IMPLEMENTED" },
        { "range_get", "{\"sheet\":\"S\",\"address\":\"A1\",\"fields\":[]}", "INVALID_ARGUMENT" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\"}", "INVALID_ARGUMENT" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\",\"values\":[[1]],\"formulas\":[[\"=1\"]]}", "INVALID_ARGUMENT" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\",\"values\":[[1,2]]}", "SHAPE_MISMATCH" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\",\"values\":[1]}", "INVALID_ARGUMENT" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\",\"values\":[[{\"type\":\"error\"}]]}", "INVALID_ARGUMENT" },
        { "range_set", "{\"sheet\":\"S\",\"address\":\"A1\",\"formulas\":[[\"1+1\"]]}", "INVALID_ARGUMENT" }
    };

    [Theory(Timeout = 60000)]
    [MemberData(nameof(RefusedArguments))]
    public async Task InvalidRequestsUseTheContractShape(string tool, string arguments, string detail)
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");

        var response = await Refusal(client, tool, JsonSerializer.Deserialize<Dictionary<string, object?>>(arguments));

        AssertRefused(response, "INVALID_REQUEST", detail, tool);
    }

    [Theory(Timeout = 60000)]
    [InlineData("session_save")]
    [InlineData("session_end")]
    public async Task OperationsNeedASession(string tool)
    {
        await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");

        AssertRefused(await Refusal(client, tool), "CONFLICT", "SESSION_NOT_STARTED", tool);
    }

    [Fact(Timeout = 60000)]
    public async Task DevelopmentModeValidatesTheProjectAndDoesNotStartExcel()
    {
        string project = Directory.CreateTempSubdirectory("xlwright-server-").FullName;
        try
        {
            await using McpClient client = await Connect(TestContext.Current.CancellationToken, "--development");
            var arguments = new Dictionary<string, object?> { ["project_path"] = project };

            AssertRefused(await Refusal(client, "session_start", arguments), "NOT_FOUND", "CONFIG_NOT_FOUND", "session_start");

            File.WriteAllText(Path.Combine(project, "project.json"), """{"paths":{"workbook":"book.xlsx"}}""");
            File.WriteAllText(Path.Combine(project, "book.xlsx"), "");
            Directory.CreateDirectory(Path.Combine(project, "src"));
            Directory.CreateDirectory(Path.Combine(project, "tests"));

            AssertRefused(await Refusal(client, "session_start", arguments), "INVALID_REQUEST", "NOT_IMPLEMENTED", "session_start");
            Assert.False(File.Exists(Path.Combine(project, ".xlwright.lock")));
        }
        finally { Directory.Delete(project, recursive: true); }
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
