using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace Spike0;

// Fixed experiment data, not a product range API.
internal static class CellProbe
{
    internal static readonly string[] Strings = ["日本語", "00123", "12345678901234567890", "2026-01-02",
        "=1+1", "+1", "-1", "@SUM(1,2)", "'abc", "''abc", "TRUE", "", " "];
    internal static readonly (string Formula, string Code, int Number)[] Errors =
    [
        ("=1/0", "#DIV/0!", 2007), ("=NA()", "#N/A", 2042),
        ("=T8UnknownName", "#NAME?", 2029), ("=SUM(D2:D3 F2:F3)", "#NULL!", 2000),
        ("=SQRT(-1)", "#NUM!", 2036), ("=INDIRECT(\"A0\")", "#REF!", 2023),
        ("=\"x\"+1", "#VALUE!", 2015), ("=SEQUENCE(2)", "#SPILL!", 2045),
        ("=FILTER(D2:D3,D2:D3=0)", "#CALC!", 2050)
    ];
    internal sealed record Cell(string Address, object? Value, string? ValueType, bool HasFormula, string? Formula);
    internal static string? Parse(string[] args) => args.Length == 1 ? args[0] switch
    {
        "--excel-s1" => "s1", "--excel-s2" => "s2", "--excel-s3" => "s3", "--excel-s4" => "s4", _ => null
    } : null;
    internal static string? Previous(string scenario) => scenario switch
    {
        "s1" => null, "s2" => "s1", "s3" => "s2", "s4" => "s3", _ => throw new InvalidDataException("Unknown cell case.")
    };
    internal static int ErrorRaw(int number) => unchecked((int)0x800A0000) | number;
    internal static void RequirePrevious(string repository, string workspace, string scenario)
    {
        string? previous = Previous(scenario);
        if (previous is null) return;
        string attempt = Path.Combine(workspace, previous + ".attempt");
        string success = Path.Combine(workspace, previous + ".success");
        RuntimeAssets.RejectReparsePoints(repository, attempt);
        RuntimeAssets.RejectReparsePoints(repository, success);
        Require(File.Exists(attempt) && File.Exists(success) && new FileInfo(success).Length > 0 &&
            File.ReadAllText(attempt) == File.ReadAllText(success), "previous case evidence");
    }
    internal static object? Project(object? value) => value is int raw ?
        new { type = "error", code = Errors.Where(item => ErrorRaw(item.Number) == raw).Select(item => item.Code).FirstOrDefault(), raw } : value;
    internal static void Require(bool condition, string reason)
    {
        if (!condition) throw new InvalidDataException("Cell hypothesis mismatch: " + reason + ". Preserve evidence; do not retry.");
    }
    internal static void WriteIdentity(string repository, string run, string scenario)
    {
        string[] sources = ["Program.cs", "ExcelScenario.cs", "ExcelCellProbe.cs", "CellProbe.cs", "SavedCellError.cs", "ExecutionGate.cs", "RuntimeAssets.cs"];
        File.WriteAllText(Path.Combine(run, "cell-build.json"), JsonSerializer.Serialize(new
        {
            scenario, runtime = Environment.Version.ToString(), framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            process_architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
            assembly_hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(CellProbe).Assembly.Location))),
            source_hashes = sources.ToDictionary(name => name, name => Convert.ToHexString(SHA256.HashData(
                File.ReadAllBytes(Path.Combine(repository, "spikes", "Spike0", name)))))
        }));
    }
    internal static XDocument ReadXml(ZipArchive archive, string part)
    {
        using var stream = (archive.GetEntry(part) ?? throw new InvalidDataException("Missing OOXML part: " + part)).Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024 });
        return XDocument.Load(reader);
    }
    internal static void RecheckS4(string repository)
    {
        const string runId = "7782619e5a6749e0bd1b39f3c085836e";
        string directory = Path.Combine(repository, ".xlwright", "spike0", "runs", runId);
        string workbook = Path.Combine(directory, "lifecycle.xlsx");
        string observations = Path.Combine(directory, "cell-observations.json");
        string resultDirectory = RuntimeAssets.CreateLocalDirectory(repository, ".xlwright/spike0/recovery/s4-" + runId);
        string result = Path.Combine(resultDirectory, "saved-recheck.json");
        RuntimeAssets.RejectReparsePoints(repository, result);
        StartupProbe.RequireNewAttempt(result);
        RuntimeAssets.RejectReparsePoints(repository, workbook);
        RuntimeAssets.RejectReparsePoints(repository, observations);
        // Keep original evidence read-only and deny writers/deletion throughout the recheck.
        using var bookHandle = new FileStream(workbook, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var observationHandle = new FileStream(observations, FileMode.Open, FileAccess.Read, FileShare.Read);
        const string bookHash = "29169F668B7AE98A76DFFC48C1D24F758F3C47A392A7E9CE58DFBB192E7C790C";
        const string observationHash = "1C3127D8FF7FF03B55FA099BD59BF6B12500267A174E5451A276751939ABB302";
        StartupProbe.VerifyHash(workbook, bookHash);
        StartupProbe.VerifyHash(observations, observationHash);
        _ = VerifySaved(workbook, writeSheet: false);
        StartupProbe.VerifyHash(workbook, bookHash);
        StartupProbe.VerifyHash(observations, observationHash);
        using var stream = new FileStream(result, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        JsonSerializer.Serialize(stream, new
        {
            at_utc = DateTime.UtcNow, run_id = runId, saved_verification_passed = true,
            original_exit_code = 1, workbook_hash = bookHash, observations_hash = observationHash,
            verifier_hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(CellProbe).Assembly.Location))),
            original_evidence_unchanged = true, excel_started = false, guard_changed = false, success_marker_created = false
        });
        stream.Flush(true);
        Console.WriteLine("S4 saved evidence verified without Excel. Original failure and guard are unchanged.");
    }
    internal static string VerifySaved(string workbook, bool writeSheet = true)
    {
        RuntimeAssets.VerifyCell(workbook, 0);
        string directory = Path.GetDirectoryName(workbook)!;
        var cells = JsonSerializer.Deserialize<Cell[]>(File.ReadAllText(Path.Combine(directory, "cell-observations.json")))!;
        using var archive = ZipFile.OpenRead(workbook);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheet = ReadXml(archive, "xl/worksheets/sheet1.xml");
        string[] strings = archive.GetEntry("xl/sharedStrings.xml") is null ? [] :
            ReadXml(archive, "xl/sharedStrings.xml").Descendants(ns + "si")
                .Select(item => string.Concat(item.Descendants(ns + "t").Select(text => text.Value))).ToArray();
        foreach (var observed in cells)
        {
            var saved = sheet.Descendants(ns + "c").SingleOrDefault(item => (string?)item.Attribute("r") == observed.Address);
            string? kind = (string?)saved?.Attribute("t");
            string? text = saved?.Element(ns + "v")?.Value;
            if (kind == "s") text = strings[int.Parse(text!, CultureInfo.InvariantCulture)];
            if (kind == "inlineStr") text = string.Concat(saved!.Descendants(ns + "t").Select(item => item.Value));
            if (kind == "e") text = SavedCellError.Read(saved!, part => ReadXml(archive, part));
            var value = (JsonElement?)observed.Value;
            bool matches = observed.ValueType switch
            {
                null => text is null && saved?.Element(ns + "f") is null,
                "System.String" => kind is "s" or "inlineStr" or "str" && (text ?? "") == value!.Value.GetString(),
                "System.Boolean" => kind == "b" && text == (value!.Value.GetBoolean() ? "1" : "0"),
                "System.Double" => kind is null or "n" && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && number == value!.Value.GetDouble(),
                "System.Int32" => kind == "e" && Errors.Any(item => ErrorRaw(item.Number) == value!.Value.GetInt32() && item.Code == text),
                _ => false
            };
            Require(matches, "saved value " + observed.Address);
            if (observed.HasFormula) Require(saved?.Element(ns + "f") is not null, "saved formula " + observed.Address);
        }
        string sheetXml = sheet.ToString();
        if (writeSheet) File.WriteAllText(Path.Combine(directory, "saved-sheet.xml"), sheetXml);
        return sheetXml;
    }
}
