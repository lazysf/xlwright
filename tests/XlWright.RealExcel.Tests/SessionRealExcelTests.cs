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

        var read = await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A1:D2" });
        Assert.Contains("日本語", Assert.IsType<TextContentBlock>(Assert.Single(read.Content)).Text);
        var values = Success(read);
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

    // Needs AccessVBOM 1. The last call shows a MsgBox on the person's screen until the watch ends the Excel.
    [Fact]
    public async Task RunsMacrosAndReportsErrorsStatusAndADialog()
    {
        _repo.AddApp("apps/run");
        var app = Path.Combine(_repo.Root, "apps", "run");
        MinimalWorkbook.Write(Path.Combine(app, "Run.xlsm"));
        Source(app, "src/modules/Calc.bas", """
            Option Explicit

            Public Function Twice(ByVal x As Long) As Long
                Twice = x * 2
            End Function

            Public Function Greet(ByVal name As String) As String
                Greet = "こんにちは、" & name
            End Function

            Public Sub DivZero()
                Dim a As Double
                a = 1 / 0
            End Sub

            Public Function Slow() As Long
                Application.Wait Now + TimeSerial(0, 0, 3)
                Slow = 1
            End Function

            Public Function Items() As Collection
                Set Items = New Collection
            End Function

            Public Sub Ask()
                MsgBox "本文です", vbOKOnly, "題名です"
            End Sub
            """);
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/run" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var twice = Success(await Call("macro_run", new() { ["macro"] = "Calc.Twice", ["args"] = new object[] { 21 } }));
        Assert.Equal(42, twice.GetProperty("return").GetInt32());
        Assert.Equal("Calc", twice.GetProperty("sync").GetProperty("modules")[0].GetProperty("name").GetString());
        var greet = Success(await Call("macro_run", new() { ["macro"] = "Calc.Greet", ["args"] = new object[] { "世界" } }));
        Assert.Equal("こんにちは、世界", greet.GetProperty("return").GetString());
        var items = Success(await Call("macro_run", new() { ["macro"] = "Calc.Items" }));
        Assert.Equal(JsonValueKind.Null, items.GetProperty("return").ValueKind);
        Assert.Equal("Collection", items.GetProperty("return_type").GetString());

        var error = Body(await Call("macro_run", new() { ["macro"] = "Calc.DivZero" }));
        Assert.Equal("MACRO_ERROR", error.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal(11, error.GetProperty("error").GetProperty("data").GetProperty("number").GetInt32());
        Assert.Equal("ready", error.GetProperty("session").GetProperty("state").GetString());
        var missing = Body(await Call("macro_run", new() { ["macro"] = "Calc.Missing" }));
        Assert.Equal("MACRO_NOT_FOUND", missing.GetProperty("error").GetProperty("detail").GetString());

        // session_status answers while a macro runs (docs/architecture.md "マクロの実行").
        var slow = Call("macro_run", new() { ["macro"] = "Calc.Slow" });
        await Task.Delay(1_000, TestContext.Current.CancellationToken);
        var status = Success(await Call("session_status", new()));
        Assert.True(status.GetProperty("session").GetProperty("busy").GetBoolean());
        Assert.Equal("macro_run", status.GetProperty("operation").GetProperty("tool").GetString());
        Assert.Equal(1, Success(await slow).GetProperty("return").GetInt32());

        var dialog = Body(await Call("macro_run", new() { ["macro"] = "Calc.Ask" }));
        Assert.Equal("DIALOG", dialog.GetProperty("error").GetProperty("detail").GetString());
        var data = dialog.GetProperty("error").GetProperty("data");
        Assert.Equal("題名です", data.GetProperty("dialog").GetProperty("title").GetString());
        Assert.Contains("本文です", data.GetProperty("dialog").GetProperty("text").GetString());
        Assert.True(data.GetProperty("excel_ended").GetBoolean());
        Assert.Equal("closed", dialog.GetProperty("session").GetProperty("state").GetString());
        AssertNothingLeft(pid);
    }

    [Theory]
    [InlineData("WORKBOOK_FILE_READ_ONLY")]
    [InlineData("WORKBOOK_IN_USE")]
    public async Task StopsBeforeStartingExcelWhenTheWorkbookCannotBeEdited(string detail)
    {
        if (detail == "WORKBOOK_FILE_READ_ONLY")
        {
            File.SetAttributes(Workbook, FileAttributes.ReadOnly);
        }

        using var holder = detail == "WORKBOOK_IN_USE"
            ? new FileStream(Workbook, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)
            : null;

        var start = await Call("session_start", new() { ["app"] = "apps/sales" });

        Assert.True(start.IsError);
        var body = Body(start);
        Assert.Equal("CONFLICT", body.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal(detail, body.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("closed", body.GetProperty("session").GetProperty("state").GetString());
        AssertNothingLeft(null);
    }

    [Fact]
    public async Task TheNextStartCleansUpTheExcelOfAClosedConversation()
    {
        int leftover;
        await using (var closed = await ServerProcess.ConnectAsync(_repo.Root, TestContext.Current.CancellationToken, noExcel: false))
        {
            var start = await closed.CallToolAsync("session_start", new Dictionary<string, object?> { ["app"] = "apps/sales" },
                cancellationToken: TestContext.Current.CancellationToken);
            leftover = Success(start).GetProperty("excel").GetProperty("pid").GetInt32();
        }

        // Closing the conversation ends its server and leaves its Excel (docs/architecture.md "実行時の状態と所有記録").
        Assert.Contains(leftover, ExcelProcesses().Keys);

        var started = Success(await Call("session_start", new() { ["app"] = "apps/sales" }));

        var cleanup = started.GetProperty("cleanup");
        Assert.Equal(leftover, cleanup.GetProperty("excel").GetProperty("pid").GetInt32());
        Assert.Contains(cleanup.GetProperty("ended_by").GetString(), new[] { "quit", "terminate" });
        Assert.DoesNotContain(leftover, ExcelProcesses().Keys);
        var pid = started.GetProperty("excel").GetProperty("pid").GetInt32();
        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    [Fact]
    public async Task LocksTheAppAndCompletesTheRecord()
    {
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/sales" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();
        await using (var second = await ServerProcess.ConnectAsync(_repo.Root, TestContext.Current.CancellationToken, noExcel: false))
        {
            var other = await second.CallToolAsync("session_start", new Dictionary<string, object?> { ["app"] = "apps/sales" },
                cancellationToken: TestContext.Current.CancellationToken);

            Assert.True(other.IsError);
            Assert.Equal("SESSION_IN_USE", Body(other).GetProperty("error").GetProperty("detail").GetString());
        }

        Success(await Call("session_end", new() { ["save"] = false }));

        var record = JsonNode.Parse(File.ReadAllText(Path.Combine(_repo.Root, "apps", "sales", ".xlwright", "session.json")))!;
        Assert.Equal("completed", record["state"]!.GetValue<string>());
        AssertNothingLeft(pid);
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
