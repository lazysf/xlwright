using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Xunit;

namespace XlWright.RealExcel.Tests;

// T19. This scenario starts a real Excel, so it runs only when a person runs a prepared copy after approval.
// npm run check selects the connection scenario by name and never reaches this class.
public sealed class VerticalSliceTests(PreparedCopy copy)
{
    private const string Sheet = "Lifecycle";
    private const string Cell = "A2";
    private const string BookName = "lifecycle.xlsx";

    [Fact(Timeout = 300000)]
    public async Task OneCellPassesThroughTheWholeSession()
    {
        string fixture = Path.Combine(copy.Repository, "tests", "fixtures", BookName);
        string fixtureHash = Hash(fixture);
        var guard = ScenarioGuard.Create(copy.GuardPath, nameof(OneCellPassesThroughTheWholeSession));
        string project = CreateProject(fixture);

        await using (McpClient client = await McpClient.CreateAsync(new StdioClientTransport(new()
        {
            Name = "xlwright-real-excel", Command = Environment.ProcessPath!,
            Arguments = [copy.ServerPath], WorkingDirectory = copy.Repository
        }), cancellationToken: TestContext.Current.CancellationToken))
        {
            var started = await Call(client, "session_start", new() { ["project_path"] = project });
            Assert.Equal("ready", State(started));
            var cell = new Dictionary<string, object?> { ["sheet"] = Sheet, ["address"] = Cell };
            Assert.Equal(0d, FirstValue(await Call(client, "range_get", cell)));
            await Call(client, "range_set", new(cell) { ["values"] = new[] { new[] { 1d } } });
            Assert.Equal(1d, FirstValue(await Call(client, "range_get", cell)));
            await Call(client, "session_save", null);
            var ended = await Call(client, "session_end", null);
            Assert.Equal("closed", State(ended));
        }

        Assert.Equal(1d, SavedCell(Path.Combine(project, BookName)));
        using (var record = JsonDocument.Parse(File.ReadAllText(Path.Combine(project, ".xlwright.lock"))))
            Assert.Equal("completed", record.RootElement.GetProperty("kind").GetString());
        Assert.Equal(fixtureHash, Hash(fixture));
        guard.Release();
    }

    private string CreateProject(string fixture)
    {
        string project = Path.Combine(copy.Repository, ".xlwright", "development", "real-excel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(project, "src"));
        Directory.CreateDirectory(Path.Combine(project, "tests"));
        File.Copy(fixture, Path.Combine(project, BookName));
        File.WriteAllText(Path.Combine(project, "project.json"), JsonSerializer.Serialize(new { paths = new { workbook = BookName } }));
        return project;
    }

    private static async Task<JsonElement> Call(McpClient client, string tool, Dictionary<string, object?>? arguments)
    {
        CallToolResult result = await client.CallToolAsync(tool, arguments, cancellationToken: TestContext.Current.CancellationToken);
        Assert.True(result.IsError != true, Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        return Assert.IsType<JsonElement>(result.StructuredContent);
    }

    private static string? State(JsonElement response) => response.GetProperty("session").GetProperty("state").GetString();

    private static double FirstValue(JsonElement response) => response.GetProperty("values")[0][0].GetDouble();

    private static double SavedCell(string workbook)
    {
        using var archive = ZipFile.OpenRead(workbook);
        using var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var saved = XDocument.Load(reader).Descendants(main + "c").Single(item => (string?)item.Attribute("r") == Cell);
        return double.Parse(saved.Element(main + "v")!.Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
}
