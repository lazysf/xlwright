using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;

namespace Spike0;

internal static class CalculationProbe
{
    internal static string? Parse(string[] args) => args.Length == 1 ? args[0] switch
    {
        "--excel-s14-auto" => "auto", "--excel-s14-manual" => "manual",
        "--excel-s14-mode-change" => "mode-change", _ => null
    } : null;

    internal static string? Previous(string scenario) => scenario switch
    {
        "auto" => null, "manual" => "auto", "mode-change" => "manual",
        _ => throw new InvalidDataException("Unknown S14 case.")
    };

    internal static void RequirePrevious(string repository, string workspace, string scenario)
    {
        string? previous = Previous(scenario);
        if (previous is null) return;
        string attempt = Path.Combine(workspace, "s14-" + previous + ".attempt");
        string success = Path.Combine(workspace, "s14-" + previous + ".success");
        RuntimeAssets.RejectReparsePoints(repository, attempt);
        RuntimeAssets.RejectReparsePoints(repository, success);
        if (File.ReadAllText(attempt).Length == 0 || File.ReadAllText(attempt) != File.ReadAllText(success))
            throw new InvalidDataException("Previous S14 evidence is incomplete.");
    }

    internal static StartupProbe.Fixture VerifyFixture(string repository, string scenario)
    {
        _ = Previous(scenario);
        string path = Path.Combine(repository, "tests", "fixtures", "s14",
            scenario == "manual" ? "manual.xlsm" : "auto.xlsm");
        string pin = Path.ChangeExtension(path, ".sha256");
        RuntimeAssets.RejectReparsePoints(repository, path);
        RuntimeAssets.RejectReparsePoints(repository, pin);
        string hash = File.ReadAllText(pin).Trim();
        if (hash.Length != 64 || !hash.All(Uri.IsHexDigit)) throw new InvalidDataException("Invalid S14 pin.");
        StartupProbe.VerifyHash(path, hash);
        VerifyPackage(path, 1, scenario == "manual" ? "manual" : "auto");
        return new(path, hash);
    }

    internal static void VerifyPackage(string path, double input, string mode)
    {
        using var archive = ZipFile.OpenRead(path);
        if (archive.GetEntry("xl/vbaProject.bin") is null) throw new InvalidDataException("Missing S14 VBA.");
        foreach (var entry in archive.Entries)
        {
            string name = entry.FullName;
            if (name.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("xl/connections.xml", StringComparison.OrdinalIgnoreCase) ||
                name.Contains("macrosheet", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith("xl/activeX/", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Excluded S14 package part.");
            if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) &&
                CellProbe.ReadXml(archive, name).Descendants().Any(item => (string?)item.Attribute("TargetMode") == "External"))
                throw new InvalidDataException("External S14 relationship.");
        }
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var book = CellProbe.ReadXml(archive, "xl/workbook.xml");
        XNamespace compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        foreach (var element in book.Descendants())
            foreach (var attribute in element.Attributes().Where(attribute =>
                attribute.Name == compatibility + "Ignorable" || attribute.Name.LocalName == "Requires"))
                foreach (string prefix in attribute.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                    if (element.GetNamespaceOfPrefix(prefix) is null)
                        throw new InvalidDataException("Unbound S14 markup compatibility prefix.");
        if (!book.Descendants(ns + "sheet").Select(item => (string?)item.Attribute("name")).SequenceEqual(new[] { "Probe", "Result" }) ||
            ((string?)book.Root!.Element(ns + "calcPr")?.Attribute("calcMode") ?? "auto") != mode)
            throw new InvalidDataException("Unexpected S14 sheets or saved mode.");
        // Resolve sheet relationships rather than assuming Excel's part numbering.
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var relationships = CellProbe.ReadXml(archive, "xl/_rels/workbook.xml.rels");
        foreach (var sheet in book.Descendants(ns + "sheet"))
        {
            string target = (string)relationships.Root!.Elements().Single(item =>
                (string?)item.Attribute("Id") == (string?)sheet.Attribute(rel + "id")).Attribute("Target")!;
            string part = new Uri(new Uri("https://example.com/xl/workbook.xml"), target).AbsolutePath.TrimStart('/');
            var xml = CellProbe.ReadXml(archive, part);
            void RequireCell(string address, double expected, string? formula)
            {
                var cell = xml.Descendants(ns + "c").Single(item => (string?)item.Attribute("r") == address);
                if ((string?)cell.Attribute("t") is not (null or "n") ||
                    (string?)cell.Element(ns + "f") != formula ||
                    !double.TryParse((string?)cell.Element(ns + "v"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) || value != expected)
                    throw new InvalidDataException("S14 saved cell mismatch: " + address);
            }
            if ((string?)sheet.Attribute("name") == "Probe")
            {
                RequireCell("B2", input, null);
                RequireCell("C2", input * 2, "B2*2");
            }
            else RequireCell("B2", input * 2 + 1, "Probe!C2+1");
        }
    }

    internal static void VerifySaved(string path, string scenario)
    {
        VerifyPackage(path, scenario == "mode-change" ? 41 : 21, scenario == "manual" ? "manual" : "auto");
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(path)!, "calculation-saved.json"),
            JsonSerializer.Serialize(new { scenario, verified = true, workbook_hash = Hash(path) }));
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    internal static void WriteIdentity(string repository, string run, StartupProbe.Fixture fixture)
    {
        string[] sources = ["Program.cs", "ExcelScenario.cs", "CalculationProbe.cs", "ExcelCalculationProbe.cs", "ExecutionGate.cs", "StaRunner.cs", "RuntimeAssets.cs", "StartupProbe.cs", "CellProbe.cs", "OwnershipRecord.cs", "EndSafety.cs", "RecordTransition.cs", "DiagnosticLog.cs"];
        File.WriteAllText(Path.Combine(run, "calculation-build.json"), JsonSerializer.Serialize(new
        {
            runtime = Environment.Version.ToString(), architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
            assembly_hash = Hash(typeof(CalculationProbe).Assembly.Location), fixture_hash = fixture.Hash,
            source_hashes = sources.ToDictionary(name => name, name => Hash(Path.Combine(repository, "spikes", "Spike0", name)))
        }));
    }
}
