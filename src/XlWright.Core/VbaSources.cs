using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace XlWright.Core;

public enum VbaModuleKind
{
    Standard,
    Class,
}

/// <summary>A module ready for VBComponents.Import: the converted file at ImportPath.</summary>
public sealed record VbaModule(string Name, VbaModuleKind Kind, string ImportPath);

/// <summary>
/// Reads src/modules and src/classes, writes the xlwright VBA module, and converts each source for Import
/// (docs/architecture.md "VBAの反映", "xlwrightのVBAのモジュール"; docs/notes/vba-sync.md).
/// </summary>
public static partial class VbaSources
{
    public const string XlWrightName = "XlWright";

    // The class header Import needs to make a class module; without it a .cls becomes a standard module.
    private const string ClassHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    static VbaSources() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>The system ANSI code page, which Import reads files in (docs/notes/vba-run.md).</summary>
    public static Encoding SystemAnsiEncoding() => AnsiEncoding((int)Com.NativeMethods.GetACP());

    /// <summary>The code page's encoding, failing on characters it cannot hold.</summary>
    public static Encoding AnsiEncoding(int codePage) =>
        Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

    /// <summary>The text of src/modules/XlWright.bas. The first line carries the version that wrote it.</summary>
    public static string XlWrightText(string version) =>
        $"' XlWright {version}: xlwright writes this module before each sync. Do not edit it.\r\nOption Explicit\r\n";

    /// <summary>
    /// Checks every source before Excel is touched, writes XlWright.bas when it is missing or older than
    /// <paramref name="version"/>, and writes the converted files under &lt;app&gt;/.xlwright/sync/&lt;id&gt;/.
    /// The caller deletes that folder after the sync.
    /// </summary>
    public static (IReadOnlyList<VbaModule>? Modules, ToolFailure? Failure) Prepare(AppFolder app, Encoding ansi, string version)
    {
        var src = Path.Combine(app.FullPath, "src");
        var workbook = Path.Combine(src, "workbook");
        if (Directory.Exists(workbook) && Directory.EnumerateFileSystemEntries(workbook).Any(entry => !Path.GetFileName(entry).StartsWith('.')))
        {
            return (null, new ToolFailure(ToolFailure.InvalidRequest, "WORKBOOK_MODULES_NOT_SUPPORTED",
                "src/workbook/ has files. Sheet and workbook modules are not synced yet; move or remove those files, then call again."));
        }

        var files = new List<(string Path, VbaModuleKind Kind)>();
        foreach (var (folder, extension, kind) in new[] { ("modules", ".bas", VbaModuleKind.Standard), ("classes", ".cls", VbaModuleKind.Class) })
        {
            var full = Path.Combine(src, folder);
            if (!Directory.Exists(full))
            {
                continue;
            }

            foreach (var entry in Directory.EnumerateFileSystemEntries(full).Order(StringComparer.Ordinal))
            {
                var name = Path.GetFileName(entry);
                if (name.StartsWith('.'))
                {
                    continue;
                }

                if (Directory.Exists(entry))
                {
                    return Invalid(app, entry, null, $"{Relative(app, entry)} is a folder inside src/{folder}/. Put module files directly in src/{folder}/.");
                }

                if (!Path.GetExtension(name).Equals(extension, StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid(app, entry, null,
                        $"{name} is in src/{folder}/, which holds only {extension} files. Standard modules go in src/modules/*.bas and class modules in src/classes/*.cls.");
                }

                if (kind == VbaModuleKind.Class && Path.GetFileNameWithoutExtension(name).Equals(XlWrightName, StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid(app, entry, null, $"{name}: the name {XlWrightName} is kept for the module xlwright writes. Rename the class.");
                }

                files.Add((entry, kind));
            }
        }

        // Read and convert everything before writing anything, so a rejected source leaves the app folder as it was.
        var converted = new List<(string Name, VbaModuleKind Kind, byte[] Bytes)>();
        foreach (var (path, kind) in files)
        {
            var (module, failure) = Convert(app, path, kind, File.ReadAllBytes(path), ansi);
            if (failure is not null)
            {
                return (null, failure);
            }

            converted.Add(module!.Value);
        }

        var xlWrightPath = Path.Combine(src, "modules", XlWrightName + ".bas");
        if (NeedsXlWright(xlWrightPath, version))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(xlWrightPath)!);
            File.WriteAllText(xlWrightPath, XlWrightText(version), StrictUtf8);
            converted.RemoveAll(module => module.Name.Equals(XlWrightName, StringComparison.OrdinalIgnoreCase));
            var (module, failure) = Convert(app, xlWrightPath, VbaModuleKind.Standard, File.ReadAllBytes(xlWrightPath), ansi);
            if (failure is not null)
            {
                return (null, failure);
            }

            converted.Add(module!.Value);
        }

        var syncFolder = Path.Combine(app.FullPath, ".xlwright", "sync", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(syncFolder);
        var modules = new List<VbaModule>();
        foreach (var (kind, group) in new[] { VbaModuleKind.Standard, VbaModuleKind.Class }.Select(k => (k, converted.Where(m => m.Kind == k))))
        {
            foreach (var (name, _, bytes) in group.OrderBy(m => m.Name, StringComparer.Ordinal))
            {
                var importPath = Path.Combine(syncFolder, name + (kind == VbaModuleKind.Class ? ".cls" : ".bas"));
                File.WriteAllBytes(importPath, bytes);
                modules.Add(new VbaModule(name, kind, importPath));
            }
        }

        return (modules, null);
    }

    // UTF-8 to the ANSI code page with CRLF, the class header when missing, and Attribute VB_Name from the file name.
    private static ((string, VbaModuleKind, byte[])?, ToolFailure?) Convert(AppFolder app, string path, VbaModuleKind kind, byte[] bytes, Encoding ansi)
    {
        string text;
        try
        {
            text = StrictUtf8.GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return (null, Invalid(app, path, null, "The file is not valid UTF-8. Save it as UTF-8.").Failure);
        }

        var lines = text.TrimStart('﻿').ReplaceLineEndings("\n").TrimEnd('\n').Split('\n').ToList();
        var fileName = Path.GetFileNameWithoutExtension(path);
        var name = fileName;
        var nameLine = lines.FindIndex(line => VbName().IsMatch(line));
        if (nameLine >= 0)
        {
            name = VbName().Match(lines[nameLine]).Groups["name"].Value;
            if (!name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
            {
                return (null, Invalid(app, path, nameLine + 1,
                    $"Attribute VB_Name is \"{name}\", but the file is {Path.GetFileName(path)}. Make them the same.").Failure);
            }
        }

        if (name.Equals(XlWrightName, StringComparison.OrdinalIgnoreCase) && !Path.GetFileName(path).Equals(XlWrightName + ".bas", StringComparison.OrdinalIgnoreCase))
        {
            return (null, Invalid(app, path, nameLine + 1, $"The name {XlWrightName} is kept for the module xlwright writes. Rename the module.").Failure);
        }

        for (var i = 0; i < lines.Count; i++)
        {
            try
            {
                ansi.GetBytes(lines[i]);
            }
            catch (EncoderFallbackException)
            {
                return (null, Invalid(app, path, i + 1,
                    $"The line has a character that the system code page ({ansi.CodePage}) cannot hold, so the VBA editor would change it. Replace it, for example with ChrW.").Failure);
            }
        }

        var addHeader = kind == VbaModuleKind.Class && !lines.FirstOrDefault(line => line.Trim().Length > 0)?.Trim().StartsWith("VERSION 1.0 CLASS", StringComparison.OrdinalIgnoreCase) == true;
        if (nameLine < 0)
        {
            var at = kind == VbaModuleKind.Class && !addHeader ? lines.FindIndex(line => line.Trim().Equals("END", StringComparison.OrdinalIgnoreCase)) + 1 : 0;
            lines.Insert(at, $"Attribute VB_Name = \"{name}\"");
        }

        var result = (addHeader ? ClassHeader : "") + string.Join("\r\n", lines) + "\r\n";
        return ((name, kind, ansi.GetBytes(result)), null);
    }

    private static bool NeedsXlWright(string path, string version)
    {
        if (!File.Exists(path))
        {
            return true;
        }

        using var reader = new StreamReader(path);
        var match = VersionLine().Match(reader.ReadLine() ?? "");
        return !match.Success || Version.Parse(match.Groups["version"].Value) < Version.Parse(version.Split('+')[0]);
    }

    private static (IReadOnlyList<VbaModule>?, ToolFailure? Failure) Invalid(AppFolder app, string path, int? line, string reason) =>
        (null, new ToolFailure(ToolFailure.InvalidRequest, "SOURCE_INVALID",
            $"{Relative(app, path)}{(line is null ? "" : $" line {line}")}: {reason} Nothing was changed in the workbook.",
            new JsonObject { ["file"] = Relative(app, path), ["line"] = line, ["reason"] = reason }));

    private static string Relative(AppFolder app, string path) => Path.GetRelativePath(app.FullPath, path).Replace('\\', '/');

    [GeneratedRegex("""^\s*Attribute\s+VB_Name\s*=\s*"(?<name>[^"]*)"\s*\z""", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex VbName();

    [GeneratedRegex(@"^' XlWright (?<version>[0-9]+\.[0-9]+\.[0-9]+)[: ]", RegexOptions.CultureInvariant)]
    private static partial Regex VersionLine();
}
