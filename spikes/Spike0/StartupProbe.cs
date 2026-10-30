using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace Spike0;

internal static class StartupProbe
{
    internal sealed record Fixture(string Path, string Hash);
    internal static string? Parse(string[] args) => args.Length == 1 ? args[0] switch
    {
        "--excel-s10-a2" => "A2", "--excel-s10-b" => "B",
        "--excel-s10-c" => "C", _ => null
    } : null;

    internal static string? PreviousCase(string scenario) => scenario switch
    {
        "A2" => null,
        "B" => "A2",
        "C" => "B",
        _ => throw new InvalidDataException("Unknown S10 case.")
    };

    // An absent managed workbook is explicitly unobserved, not a default calculation mode.
    // Getter failures with a workbook open must still stop the experiment.
    internal static int? ReadCalculation(bool managedWorkbookOpen, Func<int> read) =>
        managedWorkbookOpen ? read() : null;

    internal static void RequireNewAttempt(string path)
    {
        try { _ = File.GetAttributes(path); }
        catch (FileNotFoundException) { return; }
        catch (DirectoryNotFoundException) { return; }
        // Unknown I/O failures propagate; only a confirmed missing path permits preparation.
        throw new IOException("This S10 case was already attempted. Do not retry.");
    }

    internal static Fixture VerifyFixture(string repository)
    {
        string path = Path.Combine(repository, "tests", "fixtures", "s10", "core.xlsm");
        string pin = Path.ChangeExtension(path, ".sha256");
        RuntimeAssets.RejectReparsePoints(repository, path);
        RuntimeAssets.RejectReparsePoints(repository, pin);
        string expected = File.ReadAllText(pin).Trim();
        using (var stream = File.OpenRead(path))
            if (expected.Length != 64 || !StringComparer.OrdinalIgnoreCase.Equals(expected,
                Convert.ToHexString(SHA256.HashData(stream))))
                throw new InvalidDataException("S10 fixture hash does not match its reviewed pin.");
        using var archive = ZipFile.OpenRead(path);
        if (archive.GetEntry("xl/vbaProject.bin") is null)
            throw new InvalidDataException("S10 requires the reviewed VBA fixture.");
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.Equals("xl/connections.xml", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.Contains("macrosheet", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("External links, connections and XLM are excluded.");
            if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            using var content = entry.Open();
            using var reader = XmlReader.Create(content, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = 1024 * 1024
            });
            if (XDocument.Load(reader).Descendants().Any(item =>
                (string?)item.Attribute("TargetMode") == "External"))
                throw new InvalidDataException("External relationships are excluded.");
        }
        return new(path, expected);
    }

    internal static void VerifyHash(string path, string expected)
    {
        using var stream = File.OpenRead(path);
        if (!StringComparer.OrdinalIgnoreCase.Equals(expected, Convert.ToHexString(SHA256.HashData(stream))))
            throw new InvalidDataException("S10 asset differs from the original reviewed hash.");
    }

    internal static void RequireTrace(string scenario, string phase, string[] trace)
    {
        string[] expected = scenario switch
        {
            "A2" => phase == "closed" ? ["open", "before-close"] : ["open"],
            "B" => phase == "open" ? [] : ["auto-open"],
            "C" => [],
            _ => throw new InvalidDataException("Unknown S10 case.")
        };
        if (!trace.SequenceEqual(expected))
            throw new InvalidDataException("S10 trace differs from the expected phase; do not retry.");
    }
}
