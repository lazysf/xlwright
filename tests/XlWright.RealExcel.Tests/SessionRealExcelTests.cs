using System.Diagnostics;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.McpServer.Tests;

// Starts a real dedicated Excel through the server, the same way Claude Code does. The tests never end a process
// themselves: the server ends only the Excel whose ownership it confirmed, and the tests check what is left.
public sealed class SessionRealExcelTests : IAsyncLifetime
{
    private readonly TempRepository _repo = new();
    private readonly Dictionary<int, DateTime> _before = ExcelProcesses();
    private McpClient? _client;

    private McpClient Client => _client ?? throw new InvalidOperationException("The client is not connected.");

    private string Workbook => Path.Combine(_repo.Root, "apps", "sales", "Sales.xlsx");

    public async ValueTask InitializeAsync()
    {
        _repo.AddApp("apps/sales");
        MinimalWorkbook.Write(Workbook);
        _client = await ServerProcess.ConnectAsync(_repo.Root, TestContext.Current.CancellationToken, noExcel: false);
    }

    public async ValueTask DisposeAsync()
    {
        if (_client is not null)
        {
            // A test that failed half-way must not leave its Excel to the next start.
            var status = await Call("session_status", new());
            if (Body(status).GetProperty("session").GetProperty("state").GetString() == "ready")
            {
                await Call("session_end", new() { ["save"] = false });
            }

            await _client.DisposeAsync();
        }

        File.SetAttributes(Workbook, FileAttributes.Normal);
        _repo.Dispose();
    }

    [Fact]
    public async Task StartsReportsAndEndsWithSave()
    {
        var savedBefore = File.GetLastWriteTimeUtc(Workbook);

        var start = await Call("session_start", new() { ["app"] = "apps/sales" });

        var started = Success(start);
        Assert.Equal("ready", started.GetProperty("session").GetProperty("state").GetString());
        Assert.Equal("apps/sales", started.GetProperty("app").GetString());
        Assert.Equal("Sales.xlsx", started.GetProperty("workbook").GetString());
        var pid = started.GetProperty("excel").GetProperty("pid").GetInt32();
        Assert.DoesNotContain(pid, _before.Keys);
        using (var excel = Process.GetProcessById(pid))
        {
            Assert.Equal("EXCEL", excel.ProcessName, ignoreCase: true);
        }

        var status = Success(await Call("session_status", new()));
        Assert.Equal("ready", status.GetProperty("session").GetProperty("state").GetString());
        Assert.False(status.GetProperty("session").GetProperty("busy").GetBoolean());
        Assert.Equal("Sales.xlsx", status.GetProperty("workbook").GetString());

        // A rejected argument leaves the session as it is, and the failure says so.
        var invalid = await Call("session_end", new() { ["save"] = "yes" });
        Assert.True(invalid.IsError);
        Assert.Equal("ready", Body(invalid).GetProperty("session").GetProperty("state").GetString());

        var end = Success(await Call("session_end", new() { ["save"] = true }));

        Assert.Equal("closed", end.GetProperty("session").GetProperty("state").GetString());
        Assert.True(end.GetProperty("saved").GetBoolean());
        Assert.True(File.GetLastWriteTimeUtc(Workbook) > savedBefore);
        AssertNothingLeft(pid);
    }

    [Fact]
    public async Task SavesAndKeepsTheSession()
    {
        var savedBefore = File.GetLastWriteTimeUtc(Workbook);
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/sales" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var save = Success(await Call("workbook_save", new()));

        Assert.Equal("ready", save.GetProperty("session").GetProperty("state").GetString());
        Assert.Equal("Sales.xlsx", save.GetProperty("workbook").GetString());
        Assert.True(File.GetLastWriteTimeUtc(Workbook) > savedBefore);
        var status = Success(await Call("session_status", new()));
        Assert.Equal("ready", status.GetProperty("session").GetProperty("state").GetString());

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    [Fact]
    public async Task EndsTheExcelWhenTheWorkbookOpensReadOnly()
    {
        File.SetAttributes(Workbook, FileAttributes.ReadOnly);

        var start = await Call("session_start", new() { ["app"] = "apps/sales" });

        Assert.True(start.IsError);
        var body = Body(start);
        Assert.Equal("CONFLICT", body.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("WORKBOOK_READ_ONLY", body.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("closed", body.GetProperty("session").GetProperty("state").GetString());
        AssertNothingLeft(null);
    }

    // The Excel processes from before the test are unchanged, and no Excel started by the test is still running.
    private void AssertNothingLeft(int? pid)
    {
        var after = ExcelProcesses();
        Assert.All(_before, p => Assert.Equal(p.Value, Assert.Contains(p.Key, after)));
        Assert.Empty(after.Keys.Except(_before.Keys));
        if (pid is not null)
        {
            Assert.DoesNotContain(pid.Value, after.Keys);
        }
    }

    private static Dictionary<int, DateTime> ExcelProcesses()
    {
        var result = new Dictionary<int, DateTime>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                result[process.Id] = process.StartTime.ToUniversalTime();
            }
        }

        return result;
    }

    private Task<CallToolResult> Call(string tool, Dictionary<string, object?> arguments) =>
        Client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken).AsTask();

    private static JsonElement Success(CallToolResult result)
    {
        Assert.True(result.IsError != true, Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.NotNull(result.StructuredContent);
        return result.StructuredContent.Value.Clone();
    }

    private static JsonElement Body(CallToolResult result)
    {
        var text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        using var body = JsonDocument.Parse(text.Text);
        return body.RootElement.Clone();
    }
}
