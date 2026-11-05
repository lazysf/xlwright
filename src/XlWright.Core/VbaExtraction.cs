using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>
/// One module that Excel exported from the copy for vba_extract: Type is standard, class, document, or form; File is
/// the exported file, and null for a form, which is not exported.
/// </summary>
public sealed record ExportedModule(string Name, string Type, string? File);

/// <summary>The checks before vba_extract opens a workbook, and the writing of what it exported (docs/api.md "vba_extract").</summary>
public static class VbaExtraction
{
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Returns the full path of the workbook to extract from, or the failure that stops before Excel.</summary>
    public static (string? Source, ToolFailure? Failure) Check(AppFolder app, string? workbook)
    {
        var source = app.WorkbookPath!;
        if (workbook is not null)
        {
            source = Path.GetFullPath(Path.Combine(app.FullPath, workbook));
            if (!source.StartsWith(app.FullPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                return Fail(ToolFailure.InvalidRequest, "PATH_NOT_ALLOWED",
                    $"workbook must be a path inside the app folder, relative to it. Rejected: {workbook}");
            }
        }

        var name = Path.GetRelativePath(app.FullPath, source).Replace(Path.DirectorySeparatorChar, '/');
        if (!File.Exists(source))
        {
            return Fail(ToolFailure.NotFound, "WORKBOOK_NOT_FOUND", $"The app folder has no workbook {name}.");
        }

        if (!Path.GetExtension(source).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return Fail(ToolFailure.InvalidRequest, "NOT_MACRO_ENABLED", $"{name} is not a macro-enabled workbook (.xlsm), so it holds no VBA to extract.");
        }

        try
        {
            using var stream = OpenShared(source);
            using var package = new ZipArchive(stream, ZipArchiveMode.Read);
            if (package.Entries.Any(entry => entry.FullName.StartsWith("xl/macrosheets/", StringComparison.OrdinalIgnoreCase)))
            {
                // Whether ForceDisable stops Excel 4.0 macros is not known (docs/architecture.md "VBAの反映").
                return Fail(ToolFailure.InvalidRequest, "EXCEL4_MACROS",
                    $"{name} has Excel 4.0 macro sheets, which xlwright does not open. Ask the person how to handle this workbook.");
            }
        }
        catch (Exception e) when (e is InvalidDataException or IOException or UnauthorizedAccessException)
        {
            return Fail(ToolFailure.ExcelError, "OPEN_FAILED", $"{name} could not be read as a workbook package: {e.Message}");
        }

        var existing = ExistingSources(app);
        if (existing.Count > 0)
        {
            var data = new JsonObject { ["files"] = new JsonArray(existing.Select(file => JsonValue.Create(file)).ToArray<JsonNode?>()) };
            return Fail(ToolFailure.Conflict, "SOURCE_EXISTS",
                "src/ already has source files, and vba_extract does not overwrite them. Move or remove them with Git first, or keep them and skip the extraction.", data);
        }

        return (source, null);
    }

    /// <summary>
    /// Converts each exported file from <paramref name="ansi"/> to UTF-8 and writes it under src/. Returns the files
    /// and the skipped modules for the result, or EXTRACT_FAILED with data.done.
    /// </summary>
    public static (JsonArray? Files, JsonArray? Skipped, ToolFailure? Failure) Write(AppFolder app, IReadOnlyList<ExportedModule> modules, Encoding ansi)
    {
        var files = new JsonArray();
        var skipped = new JsonArray();
        var done = new List<string>();
        foreach (var module in modules)
        {
            var folder = module.Type switch
            {
                "standard" => "modules",
                "class" => "classes",
                "document" => "workbook",
                _ => null,
            };

            // xlwright writes its own module before every sync (docs/architecture.md "xlwrightのVBAのモジュール").
            if (module.Name.Equals(VbaSources.XlWrightName, StringComparison.OrdinalIgnoreCase) && module.Type == "standard")
            {
                continue;
            }

            if (folder is null || module.File is null)
            {
                skipped.Add(new JsonObject { ["name"] = module.Name, ["type"] = module.Type });
                continue;
            }

            var relative = $"src/{folder}/{module.Name}{(module.Type == "standard" ? ".bas" : ".cls")}";
            try
            {
                var text = ansi.GetString(File.ReadAllBytes(module.File));
                var target = Path.Combine(app.FullPath, relative.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.WriteAllText(target, text, Utf8);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException or DecoderFallbackException)
            {
                var data = new JsonObject { ["done"] = new JsonArray(done.Select(file => JsonValue.Create(file)).ToArray<JsonNode?>()) };
                return (null, null, new ToolFailure(ToolFailure.ExcelError, "EXTRACT_FAILED",
                    $"Writing {relative} failed: {e.Message} The files in data.done were written and are left as they are.", data));
            }

            done.Add(relative);
            files.Add(new JsonObject { ["file"] = relative, ["type"] = module.Type });
        }

        return (files, skipped, null);
    }

    /// <summary>Copies the workbook while Excel may hold it open for editing, as the main workbook is.</summary>
    public static void CopyShared(string source, string target)
    {
        using var from = OpenShared(source);
        using var to = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        from.CopyTo(to);
    }

    // Excel holds a workbook it opened for editing with write access, so a reader must share writing too.
    private static FileStream OpenShared(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

    // The files under src/ that the extraction would overwrite or mix with: all but dot files and XlWright.bas.
    private static List<string> ExistingSources(AppFolder app)
    {
        var src = Path.Combine(app.FullPath, "src");
        if (!Directory.Exists(src))
        {
            return [];
        }

        var xlWright = Path.Combine(src, "modules", VbaSources.XlWrightName + ".bas");
        return Directory.EnumerateFiles(src, "*", SearchOption.AllDirectories)
            .Where(path => !Path.GetFileName(path).StartsWith('.') && !path.Equals(xlWright, StringComparison.OrdinalIgnoreCase))
            .Select(path => Path.GetRelativePath(app.FullPath, path).Replace(Path.DirectorySeparatorChar, '/'))
            .Order(StringComparer.Ordinal)
            .ToList();
    }

    private static (string?, ToolFailure?) Fail(string code, string detail, string message, JsonObject? data = null) =>
        (null, new ToolFailure(code, detail, message, data));
}
