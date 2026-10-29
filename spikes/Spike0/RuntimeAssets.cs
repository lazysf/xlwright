using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;
using System.IO.Compression;

namespace Spike0;

internal static class RuntimeAssets
{
    internal const string FixtureHash = "E87C630F4B14250700A40359FD1788875D145AE2019974D8C5264D591BEBBDA2";

    internal static string FindRepository()
    {
        var directory = new DirectoryInfo(Environment.CurrentDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "global.json")) &&
                File.Exists(Path.Combine(directory.FullName, "spikes", "Spike0", "Spike0.csproj")))
                return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("Run Spike0 from this repository.");
    }

    internal static void VerifyPias()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "assembly", "GAC_MSIL");
        VerifyHash(Path.Combine(root, "Microsoft.Office.Interop.Excel",
            "15.0.0.0__71e9bce111e9429c", "Microsoft.Office.Interop.Excel.dll"),
            "6F262BBA82EED8A8D69FAC44E491B99CCA2D4CD448166291CE2186833E730A85");
        VerifyHash(Path.Combine(root, "office", "15.0.0.0__71e9bce111e9429c", "OFFICE.DLL"),
            "5512CABD57B6E1FBD2B96C298D804A3795CD317F61E154AEDB335F6C119EAF62");
    }

    internal static void VerifyFixture(string repository)
    {
        string path = Path.Combine(repository, "tests", "fixtures", "lifecycle.xlsx");
        RejectReparsePoints(repository, path);
        VerifyHash(path, FixtureHash);
        VerifyCell(path, 0);
    }

    internal static void VerifyCell(string workbook, double expected)
    {
        using var archive = ZipFile.OpenRead(workbook);
        var entry = archive.GetEntry("xl/worksheets/sheet1.xml") ??
            throw new InvalidDataException("Known fixture sheet is missing.");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 1024 * 1024
        });
        var document = XDocument.Load(reader);
        XNamespace sheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cell = document.Descendants(sheetNamespace + "c").Single(item => (string?)item.Attribute("r") == "A2");
        if ((string?)cell.Attribute("t") is not (null or "n") || cell.Element(sheetNamespace + "f") is not null ||
            (double?)cell.Element(sheetNamespace + "v") != expected)
            throw new InvalidDataException("Saved fixture cell is not the expected numeric constant.");
    }

    internal static string CreateLocalDirectory(string repository, string relativePath)
    {
        string path = Path.GetFullPath(relativePath, repository);
        RejectReparsePoints(repository, path);
        Directory.CreateDirectory(path);
        return path;
    }

    internal static void RejectReparsePoints(string repository, string path)
    {
        string relative = Path.GetRelativePath(repository, path);
        if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar))
            throw new IOException("Experiment paths must stay within the repository.");
        string current = repository;
        foreach (string component in relative.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, component);
            if ((File.Exists(current) || Directory.Exists(current)) &&
                File.GetAttributes(current).HasFlag(FileAttributes.ReparsePoint))
                throw new IOException("Experiment paths may not use reparse points.");
        }
    }

    private static void VerifyHash(string path, string expected)
    {
        using var stream = File.OpenRead(path);
        string actual = Convert.ToHexString(SHA256.HashData(stream));
        if (!StringComparer.Ordinal.Equals(actual, expected))
            throw new InvalidDataException("Pinned file hash changed: " + Path.GetFileName(path));
    }
}
