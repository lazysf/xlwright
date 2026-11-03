using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
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
        MinimalWorkbook.Write(Workbook, lockedSheet: true);
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
    public async Task WritesAndReadsValuesFormulasAndErrors()
    {
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/sales" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var written = Success(await Call("range_write", new()
        {
            ["sheet"] = "sheet1",
            ["range"] = "a1:d1",
            ["values"] = new[] { new object?[] { 1.5, "日本語", true, null } },
        }));
        Assert.Equal("A1:D1", written.GetProperty("range").GetString());
        Assert.Equal(4, written.GetProperty("cells").GetInt32());
        Success(await Call("range_write", new()
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A2:C2",
            ["formulas"] = new[] { new[] { "=A1*2", "=1/0", "=NA()" } },
        }));

        var values = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A1:D2" }));
        Assert.True(JsonNode.DeepEquals(
            JsonNode.Parse("""[[1.5,"日本語",true,null],[3,{"error":"#DIV/0!"},{"error":"#N/A"},null]]"""),
            JsonNode.Parse(values.GetProperty("values").GetRawText())), values.GetProperty("values").GetRawText());
        var formulas = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A2:B2", ["content"] = "formulas" }));
        Assert.Equal("""[["=A1*2","=1/0"]]""", formulas.GetProperty("formulas").GetRawText());
        var one = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "B1" }));
        Assert.Equal("B1", one.GetProperty("range").GetString());

        Success(await Call("range_write", new() { ["sheet"] = "Sheet1", ["range"] = "A1:A2", ["clear"] = true }));
        var cleared = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A1:A2" }));
        Assert.Equal("[[null],[null]]", cleared.GetProperty("values").GetRawText());

        var missing = Body(await Call("range_read", new() { ["sheet"] = "Missing", ["range"] = "A1" }));
        Assert.Equal("SHEET_NOT_FOUND", missing.GetProperty("error").GetProperty("detail").GetString());
        var locked = Body(await Call("range_write", new() { ["sheet"] = "Locked", ["range"] = "A1", ["values"] = new[] { new object?[] { 1 } } }));
        Assert.Equal("WRITE_FAILED", locked.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("ready", locked.GetProperty("session").GetProperty("state").GetString());

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    // Needs "Trust access to the VBA project object model" (AccessVBOM 1).
    [Fact]
    public async Task SyncsReplacesAndRefusesARenamedModule()
    {
        _repo.AddApp("apps/macro");
        var app = Path.Combine(_repo.Root, "apps", "macro");
        MinimalWorkbook.Write(Path.Combine(app, "Macro.xlsm"));
        Source(app, "src/modules/Calc.bas", "Option Explicit\n' 日本語のコメント\nPublic Function Name() As String\n    Name = \"日本語\"\nEnd Function\n");
        Source(app, "src/classes/Counter.cls", "Option Explicit\nPublic Value As Long\n");
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/macro" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var first = Success(await Call("vba_sync", new()));

        Assert.Equal("""[{"name":"Calc","type":"standard"},{"name":"XlWright","type":"standard"},{"name":"Counter","type":"class"}]""",
            first.GetProperty("modules").GetRawText());
        Assert.Equal("[]", first.GetProperty("removed").GetRawText());
        Assert.True(File.Exists(Path.Combine(app, "src", "modules", "XlWright.bas")));
        Assert.Empty(Directory.EnumerateFileSystemEntries(Path.Combine(app, ".xlwright", "sync")));

        File.Delete(Path.Combine(app, "src", "modules", "Calc.bas"));
        var second = Success(await Call("vba_sync", new()));
        Assert.Equal("""["Calc"]""", second.GetProperty("removed").GetRawText());

        // Sheet1 is the sheet module's name, so Import renames the module to Sheet11 (docs/notes/vba-sync.md).
        Source(app, "src/modules/Sheet1.bas", "Option Explicit\n");
        var renamed = Body(await Call("vba_sync", new()));
        Assert.Equal("SYNC_FAILED", renamed.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("[]", renamed.GetProperty("error").GetProperty("data").GetProperty("done").GetRawText());
        Assert.Equal("ready", renamed.GetProperty("session").GetProperty("state").GetString());

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

    private static void Source(string app, string file, string text)
    {
        var path = Path.Combine(app, file);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, text, new System.Text.UTF8Encoding(false));
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
