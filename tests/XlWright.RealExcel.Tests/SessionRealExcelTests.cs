using System.Diagnostics;
using System.Text;
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

        // ConvertsTheXlsxMainWorkbookToXlsm removes it.
        if (File.Exists(Workbook))
        {
            File.SetAttributes(Workbook, FileAttributes.Normal);
        }

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

        // The diagnostic log stays after the session (docs/architecture.md "診断の記録の形式").
        var log = Assert.Single(Directory.GetFiles(Path.Combine(_repo.Root, "apps", "sales", ".xlwright", "logs")));
        var events = File.ReadAllLines(log).Select(line => JsonNode.Parse(line)!["event"]!.GetValue<string>()).ToList();
        Assert.Equal(["start", "owned", "call", "end", "call"], events);
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

    // Needs AccessVBOM 1. The sheet is renamed, so its CodeName comes from vba_get, not from the sheet name.
    [Fact]
    public async Task ListsAndSyncsTheSheetAndWorkbookModulesByCodeName()
    {
        _repo.AddApp("apps/sheets");
        var app = Path.Combine(_repo.Root, "apps", "sheets");
        MinimalWorkbook.Write(Path.Combine(app, "Sheets.xlsm"), firstSheet: "データ");
        Source(app, "src/modules/Poke.bas", "Option Explicit\nPublic Sub WriteA1()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value = 5\nEnd Sub\n");
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/sheets" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var before = Success(await Call("vba_get", new())).GetProperty("modules").EnumerateArray().ToList();
        var sheet = Assert.Single(before, m => m.TryGetProperty("sheet", out var name) && name.GetString() == "データ");
        var codeName = sheet.GetProperty("name").GetString()!;
        Assert.Equal("document", sheet.GetProperty("type").GetString());
        Assert.Contains(before, m => m.GetProperty("name").GetString() == "ThisWorkbook" && !m.TryGetProperty("sheet", out _));

        Source(app, $"src/workbook/{codeName}.cls",
            "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"" + codeName + "\"\n" +
            "Option Explicit\n\nPrivate Sub Worksheet_Change(ByVal Target As Range)\n    If Target.Address = \"$A$1\" Then Me.Range(\"B1\").Value = \"変更を見た\"\nEnd Sub\n");
        Source(app, "src/workbook/ThisWorkbook.cls", "Option Explicit\n\nPublic Function Hello() As String\n    Hello = \"ブック\"\nEnd Function\n");
        var synced = Success(await Call("vba_sync", new())).GetProperty("modules").EnumerateArray()
            .Where(m => m.GetProperty("type").GetString() == "document").Select(m => m.GetProperty("name").GetString());
        Assert.Equal([codeName, "ThisWorkbook"], synced);

        var after = Success(await Call("vba_get", new())).GetProperty("modules").EnumerateArray().ToList();
        Assert.Equal(5, after.Single(m => m.GetProperty("name").GetString() == codeName).GetProperty("lines").GetInt32());
        Assert.Equal(5, after.Single(m => m.GetProperty("name").GetString() == "ThisWorkbook").GetProperty("lines").GetInt32());

        // macro_run leaves events on, so the replaced Worksheet_Change runs.
        Success(await Call("macro_run", new() { ["macro"] = "Poke.WriteA1" }));
        var b1 = Success(await Call("range_read", new() { ["sheet"] = "データ", ["range"] = "B1" }));
        Assert.Equal("変更を見た", b1.GetProperty("values")[0][0].GetString());

        Source(app, "src/workbook/NoSuchSheet.cls", "Option Explicit\n");
        var missing = Body(await Call("vba_sync", new()));
        Assert.Equal("WORKBOOK_MODULE_NOT_FOUND", missing.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("src/workbook/NoSuchSheet.cls", missing.GetProperty("error").GetProperty("data").GetProperty("file").GetString());
        Assert.Equal("ready", missing.GetProperty("session").GetProperty("state").GetString());

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    // Needs AccessVBOM 1. The name is Japanese, so the path is not ASCII.
    [Fact]
    public async Task CreatesTheMainWorkbookInAnEmptyAppFolder()
    {
        _repo.AddApp("apps/fresh");
        var app = Path.Combine(_repo.Root, "apps", "fresh");
        var started = Success(await Call("session_start", new() { ["app"] = "apps/fresh" }));
        var pid = started.GetProperty("excel").GetProperty("pid").GetInt32();
        Assert.Equal(JsonValueKind.Null, started.GetProperty("workbook").ValueKind);
        var none = Body(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A1" }));
        Assert.Equal("NO_WORKBOOK", none.GetProperty("error").GetProperty("detail").GetString());

        var created = Success(await Call("workbook_create", new() { ["name"] = "新規.xlsm" }));
        Assert.Equal("新規.xlsm", created.GetProperty("workbook").GetString());
        Assert.True(File.Exists(Path.Combine(app, "新規.xlsm")));
        Source(app, "src/modules/Calc.bas", "Option Explicit\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
        var twice = Success(await Call("macro_run", new() { ["macro"] = "Calc.Twice", ["args"] = new object[] { 21 } }));
        Assert.Equal(42, twice.GetProperty("return").GetInt32());
        var sheet = Success(await Call("vba_get", new())).GetProperty("modules").EnumerateArray()
            .Select(m => m.TryGetProperty("sheet", out var name) ? name.GetString() : null).First(name => name is not null)!;
        Success(await Call("range_write", new() { ["sheet"] = sheet, ["range"] = "A1", ["values"] = new[] { new object[] { "作った" } } }));
        var again = Body(await Call("workbook_create", new() { ["name"] = "Other.xlsm" }));
        Assert.Equal("WORKBOOK_EXISTS", again.GetProperty("error").GetProperty("detail").GetString());
        Success(await Call("session_end", new() { ["save"] = true }));
        AssertNothingLeft(pid);

        var next = Success(await Call("session_start", new() { ["app"] = "apps/fresh" }));
        Assert.Equal("新規.xlsm", next.GetProperty("workbook").GetString());
        var a1 = Success(await Call("range_read", new() { ["sheet"] = sheet, ["range"] = "A1" }));
        Assert.Equal("作った", a1.GetProperty("values")[0][0].GetString());
        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(next.GetProperty("excel").GetProperty("pid").GetInt32());
    }

    // Needs AccessVBOM 1. The .xlsx stays next to the .xlsm, so the next start stops with MULTIPLE_WORKBOOKS.
    [Fact]
    public async Task ConvertsTheXlsxMainWorkbookToXlsm()
    {
        var app = Path.Combine(_repo.Root, "apps", "sales");
        var started = Success(await Call("session_start", new() { ["app"] = "apps/sales" }));
        var pid = started.GetProperty("excel").GetProperty("pid").GetInt32();
        var before = Success(await Call("workbook_get", new()));
        Assert.Equal("xlsx", before.GetProperty("format").GetString());
        Assert.True(before.GetProperty("saved").GetBoolean());
        Success(await Call("range_write", new() { ["sheet"] = "Sheet1", ["range"] = "B1", ["values"] = new[] { new object[] { "変える前" } } }));
        Assert.False(Success(await Call("workbook_get", new())).GetProperty("saved").GetBoolean());

        var converted = Success(await Call("workbook_edit", new() { ["action"] = "convert_to_xlsm" }));
        Assert.Equal("Sales.xlsm", converted.GetProperty("workbook").GetString());
        Assert.Equal("Sales.xlsx", converted.GetProperty("previous").GetString());
        Assert.True(File.Exists(Path.Combine(app, "Sales.xlsm")));
        var after = Success(await Call("workbook_get", new()));
        Assert.Equal("Sales.xlsm", after.GetProperty("workbook").GetString());
        Assert.Equal("xlsm", after.GetProperty("format").GetString());
        var again = Body(await Call("workbook_edit", new() { ["action"] = "convert_to_xlsm" }));
        Assert.Equal("ALREADY_MACRO_ENABLED", again.GetProperty("error").GetProperty("detail").GetString());

        Source(app, "src/modules/Calc.bas", "Option Explicit\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
        var twice = Success(await Call("macro_run", new() { ["macro"] = "Calc.Twice", ["args"] = new object[] { 21 } }));
        Assert.Equal(42, twice.GetProperty("return").GetInt32());
        Success(await Call("session_end", new() { ["save"] = true }));
        AssertNothingLeft(pid);

        var next = Body(await Call("session_start", new() { ["app"] = "apps/sales" }));
        Assert.Equal("MULTIPLE_WORKBOOKS", next.GetProperty("error").GetProperty("detail").GetString());
        File.Delete(Workbook);
        var reopened = Success(await Call("session_start", new() { ["app"] = "apps/sales" }));
        Assert.Equal("Sales.xlsm", reopened.GetProperty("workbook").GetString());
        var b1 = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "B1" }));
        Assert.Equal("変える前", b1.GetProperty("values")[0][0].GetString());
        Assert.Contains(Success(await Call("vba_get", new())).GetProperty("modules").EnumerateArray(), m => m.GetProperty("name").GetString() == "Calc");
        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(reopened.GetProperty("excel").GetProperty("pid").GetInt32());
    }

    // Needs AccessVBOM 1. The existing workbook is made with vba_sync, then src/ is removed, so its VBA is only in the
    // workbook. Workbook_Open writes a marker into the app folder by full path, so the removed copy cannot hide it.
    [Fact]
    public async Task ExtractsAnExistingWorkbookWithoutRunningItsCode()
    {
        _repo.AddApp("apps/legacy");
        var app = Path.Combine(_repo.Root, "apps", "legacy");
        MinimalWorkbook.Write(Path.Combine(app, "Legacy.xlsm"));
        var marker = Path.Combine(app, "opened.txt");
        Source(app, "src/modules/Calc.bas", "Option Explicit\n' 日本語のコメント\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
        Source(app, "src/classes/Counter.cls", "Option Explicit\n' 数える\nPublic Value As Long\n");
        Source(app, "src/workbook/ThisWorkbook.cls",
            $"Option Explicit\n\nPrivate Sub Workbook_Open()\n    Dim f As Integer\n    f = FreeFile\n    Open \"{marker}\" For Output As #f\n    Close #f\nEnd Sub\n");
        Success(await Call("session_start", new() { ["app"] = "apps/legacy" }));
        Success(await Call("vba_sync", new()));
        Success(await Call("session_end", new() { ["save"] = true }));
        Directory.Delete(Path.Combine(app, "src"), recursive: true);

        var pid = Success(await Call("session_start", new() { ["app"] = "apps/legacy" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();
        var extracted = Success(await Call("vba_extract", new()));

        Assert.False(File.Exists(marker));
        Assert.Equal("Legacy.xlsm", extracted.GetProperty("workbook").GetString());
        var files = extracted.GetProperty("files").EnumerateArray().Select(f => f.GetProperty("file").GetString()).ToList();
        Assert.Contains("src/modules/Calc.bas", files);
        Assert.Contains("src/classes/Counter.cls", files);
        Assert.Contains("src/workbook/ThisWorkbook.cls", files);
        Assert.DoesNotContain("src/modules/XlWright.bas", files);
        Assert.Contains("日本語のコメント", File.ReadAllText(Path.Combine(app, "src", "modules", "Calc.bas"), Encoding.UTF8));
        Assert.Contains("Workbook_Open", File.ReadAllText(Path.Combine(app, "src", "workbook", "ThisWorkbook.cls"), Encoding.UTF8));
        Assert.Empty(Directory.EnumerateFileSystemEntries(Path.Combine(app, ".xlwright", "extract")));

        // The extracted files sync back as they are.
        var synced = Success(await Call("vba_sync", new())).GetProperty("modules").EnumerateArray().Select(m => m.GetProperty("name").GetString());
        Assert.Contains("Calc", synced);
        Assert.Contains("Counter", synced);
        var twice = Success(await Call("macro_run", new() { ["macro"] = "Calc.Twice", ["args"] = new object[] { 21 } }));
        Assert.Equal(42, twice.GetProperty("return").GetInt32());
        var exists = Body(await Call("vba_extract", new()));
        Assert.Equal("SOURCE_EXISTS", exists.GetProperty("error").GetProperty("detail").GetString());

        Success(await Call("session_end", new() { ["save"] = false }));
        Assert.False(File.Exists(marker));
        AssertNothingLeft(pid);
    }

    // Needs AccessVBOM 1 and the Scripting Runtime, which Windows registers.
    [Fact]
    public async Task AddsListsAndRemovesAReference()
    {
        const string Scripting = "{420B2830-E718-11CF-893D-00A0C9054228}";
        _repo.AddApp("apps/refs");
        MinimalWorkbook.Write(Path.Combine(_repo.Root, "apps", "refs", "Refs.xlsm"));
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/refs" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var before = Success(await Call("vba_get", new())).GetProperty("references").EnumerateArray().ToList();
        var vba = Assert.Single(before, r => r.GetProperty("name").GetString() == "VBA");
        Assert.True(vba.GetProperty("builtin").GetBoolean());
        Assert.DoesNotContain(before, r => r.GetProperty("guid").GetString() == Scripting);

        var added = Success(await Call("vba_edit", new() { ["action"] = "add_reference", ["guid"] = Scripting, ["major"] = 1, ["minor"] = 0 }));
        Assert.True(added.GetProperty("changed").GetBoolean());
        var listed = Success(await Call("vba_get", new())).GetProperty("references").EnumerateArray();
        var scripting = Assert.Single(listed, r => r.GetProperty("guid").GetString() == Scripting);
        Assert.Equal("Scripting", scripting.GetProperty("name").GetString());
        Assert.False(scripting.GetProperty("broken").GetBoolean());

        var again = Success(await Call("vba_edit", new() { ["action"] = "add_reference", ["guid"] = Scripting, ["major"] = 1, ["minor"] = 0 }));
        Assert.False(again.GetProperty("changed").GetBoolean());

        var removed = Success(await Call("vba_edit", new() { ["action"] = "remove_reference", ["guid"] = Scripting }));
        Assert.True(removed.GetProperty("changed").GetBoolean());
        Assert.DoesNotContain(removed.GetProperty("references").EnumerateArray(), r => r.GetProperty("guid").GetString() == Scripting);

        var builtIn = Body(await Call("vba_edit", new() { ["action"] = "remove_reference", ["guid"] = vba.GetProperty("guid").GetString() }));
        Assert.Equal("BUILTIN_REFERENCE", builtIn.GetProperty("error").GetProperty("detail").GetString());
        var missing = Body(await Call("vba_edit", new() { ["action"] = "remove_reference", ["guid"] = Scripting }));
        Assert.Equal("REFERENCE_NOT_FOUND", missing.GetProperty("error").GetProperty("detail").GetString());
        var unregistered = Body(await Call("vba_edit", new() { ["action"] = "add_reference", ["guid"] = "{00000000-1111-2222-3333-444444444444}", ["major"] = 1, ["minor"] = 0 }));
        Assert.Equal("REFERENCE_FAILED", unregistered.GetProperty("error").GetProperty("detail").GetString());
        Assert.Equal("ready", unregistered.GetProperty("session").GetProperty("state").GetString());

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    // Needs AccessVBOM 1. The tests change a cell and a data file; both changes stay in the copy.
    [Fact]
    public async Task RunsTheTestsInACopyAndLeavesTheAppAsItWas()
    {
        _repo.AddApp("apps/tested");
        var app = Path.Combine(_repo.Root, "apps", "tested");
        MinimalWorkbook.Write(Path.Combine(app, "Tested.xlsm"));
        Source(app, "data/note.txt", "original");
        Source(app, "src/modules/Calc.bas", "Option Explicit\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
        Source(app, "tests/CalcTests.bas",
            "Option Explicit\n\nPublic Sub TestTwice()\n    XlWrightAssert.AreEqual 6, Calc.Twice(3)\nEnd Sub\n\n" +
            "Public Sub TestWrong()\n    XlWrightAssert.AreEqual 7, Calc.Twice(3), \"三の二倍\"\nEnd Sub\n\n" +
            "Public Sub TestError()\n    Dim x As Long\n    x = 1 / 0\nEnd Sub\n\n" +
            "Public Sub TestWrites()\n    ThisWorkbook.Worksheets(1).Range(\"A1\").Value = \"changed by test\"\n    Dim f As Integer\n    f = FreeFile\n" +
            "    Open ThisWorkbook.Path & \"\\data\\note.txt\" For Output As #f\n    Print #f, \"changed by test\"\n    Close #f\n" +
            "    XlWrightAssert.IsTrue InStr(ThisWorkbook.Path, \"\\.xlwright\\test\\\") > 0, \"runs in the copy\"\nEnd Sub\n");
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/tested" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();
        Success(await Call("range_write", new() { ["sheet"] = "Sheet1", ["range"] = "A1", ["values"] = new[] { new[] { "original" } } }));

        var run = Success(await Call("test_run", new()));

        Assert.Equal("""{"passed":2,"failed":1,"errors":1}""", run.GetProperty("summary").GetRawText());
        var tests = run.GetProperty("tests").EnumerateArray().ToDictionary(t => t.GetProperty("name").GetString()!);
        Assert.Equal("AreEqual: expected 7, actual 6. 三の二倍", tests["CalcTests.TestWrong"].GetProperty("message").GetString());
        Assert.Equal(11, tests["CalcTests.TestError"].GetProperty("number").GetInt32());
        Assert.Equal("passed", tests["CalcTests.TestWrites"].GetProperty("result").GetString());
        Assert.Equal("ready", run.GetProperty("session").GetProperty("state").GetString());
        Assert.False(run.TryGetProperty("leftover_copy", out _));

        var a1 = Success(await Call("range_read", new() { ["sheet"] = "Sheet1", ["range"] = "A1" }));
        Assert.Equal("original", a1.GetProperty("values")[0][0].GetString());
        Assert.Equal("original", File.ReadAllText(Path.Combine(app, "data", "note.txt")));
        Assert.Empty(Directory.EnumerateFileSystemEntries(Path.Combine(app, ".xlwright", "test")));
        var modules = Success(await Call("vba_get", new())).GetProperty("modules").EnumerateArray().Select(m => m.GetProperty("name").GetString());
        Assert.DoesNotContain("CalcTests", modules);
        Assert.DoesNotContain("XlWrightAssert", modules);

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    // Needs AccessVBOM 1. A test written before the function it calls stops at the compile, not at a dialog that ends
    // the Excel (docs/notes/stage3-scenarios.md). The dialog and the VBA editor show on the person's screen for a moment.
    [Fact]
    public async Task ATestThatDoesNotCompileRunsNoTestAndKeepsTheSession()
    {
        _repo.AddApp("apps/first");
        var app = Path.Combine(_repo.Root, "apps", "first");
        MinimalWorkbook.Write(Path.Combine(app, "First.xlsm"));
        Source(app, "src/modules/Text.bas", "Option Explicit\n");
        Source(app, "tests/TextTests.bas",
            "Option Explicit\n\nPublic Sub TestJoin()\n    XlWrightAssert.AreEqual \"a,b\", Text.JoinFields(\"a\", \"b\")\nEnd Sub\n");
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/first" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var failed = Body(await Call("test_run", new()));

        var error = failed.GetProperty("error");
        Assert.Equal("EXCEL_ERROR", error.GetProperty("code").GetString());
        Assert.Equal("COMPILE_ERROR", error.GetProperty("detail").GetString());
        Assert.False(string.IsNullOrEmpty(error.GetProperty("data").GetProperty("error").GetProperty("message").GetString()));
        Assert.Equal("ready", failed.GetProperty("session").GetProperty("state").GetString());
        Assert.Empty(Directory.EnumerateFileSystemEntries(Path.Combine(app, ".xlwright", "test")));

        Source(app, "src/modules/Text.bas",
            "Option Explicit\nPublic Function JoinFields(ByVal a As String, ByVal b As String) As String\n    JoinFields = a & \",\" & b\nEnd Function\n");
        var run = Success(await Call("test_run", new()));
        Assert.Equal("""{"passed":1,"failed":0,"errors":0}""", run.GetProperty("summary").GetRawText());

        Success(await Call("session_end", new() { ["save"] = false }));
        AssertNothingLeft(pid);
    }

    // Needs AccessVBOM 1. The compile error shows its dialog and the VBA editor on the person's screen for a moment.
    [Fact]
    public async Task ChecksTheCompileAndReportsACompileErrorWithItsLine()
    {
        _repo.AddApp("apps/check");
        var app = Path.Combine(_repo.Root, "apps", "check");
        MinimalWorkbook.Write(Path.Combine(app, "Check.xlsm"));
        Source(app, "src/modules/Calc.bas", "Option Explicit\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
        var pid = Success(await Call("session_start", new() { ["app"] = "apps/check" }))
            .GetProperty("excel").GetProperty("pid").GetInt32();

        var good = Success(await Call("vba_check", new()));
        Assert.True(good.GetProperty("compiled").GetBoolean());
        Assert.False(good.TryGetProperty("error", out _));

        Source(app, "src/modules/Broken.bas", "Option Explicit\n\nPublic Sub Broken()\n    notDeclared = 1\nEnd Sub\n");
        var bad = Success(await Call("vba_check", new()));
        Assert.False(bad.GetProperty("compiled").GetBoolean());
        var error = bad.GetProperty("error");
        Assert.Contains("コンパイル エラー", error.GetProperty("message").GetString());
        Assert.Equal("Broken", error.GetProperty("module").GetString());
        Assert.Equal(4, error.GetProperty("line").GetInt32());
        Assert.Equal("ready", bad.GetProperty("session").GetProperty("state").GetString());

        File.Delete(Path.Combine(app, "src", "modules", "Broken.bas"));
        Assert.True(Success(await Call("vba_check", new())).GetProperty("compiled").GetBoolean());

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
