using System.Text.Json;

namespace XlWright.Core;

// project.json in the project root. The rules are in docs/architecture.md; ADR 0006 explains the format.
internal sealed record ProjectConfig(string Root, string Workbook, string VbaSource, string Tests, string Workspace, string Dist)
{
    private static readonly string[] Keys = ["workbook", "vba_source", "tests", "workspace", "dist"];
    private const string RetryHint = " Fix project.json and retry.";

    internal static ProjectConfig Load(string root)
    {
        string path = Path.Combine(root, "project.json");
        if (!File.Exists(path))
            throw new CoreException(ErrorCode.NotFound, ErrorDetail.ConfigNotFound,
                "project.json was not found in project_path. Check project_path and retry.");
        var paths = ReadPaths(File.ReadAllText(path));
        string Resolve(string key, string fallback) =>
            Path.GetFullPath(Path.Combine(root, paths.GetValueOrDefault(key, fallback)));
        if (!paths.TryGetValue("workbook", out _)) throw Invalid("paths.workbook is required.");
        var config = new ProjectConfig(root, Resolve("workbook", ""), Resolve("vba_source", "src"), Resolve("tests", "tests"),
            Resolve("workspace", ".xlwright"), Resolve("dist", "dist"));
        config.Validate();
        return config;
    }

    private static Dictionary<string, string> ReadPaths(string text)
    {
        try
        {
            using var document = JsonDocument.Parse(text, new JsonDocumentOptions { AllowDuplicateProperties = false });
            var top = document.RootElement;
            if (top.ValueKind != JsonValueKind.Object || top.EnumerateObject().Any(item => item.Name != "paths") ||
                !top.TryGetProperty("paths", out var section) || section.ValueKind != JsonValueKind.Object)
                throw Invalid("The root must be an object whose only key is paths.");
            var paths = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var item in section.EnumerateObject())
            {
                if (!Keys.Contains(item.Name)) throw Invalid("paths has an unknown key.");
                if (item.Value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(item.Value.GetString()))
                    throw Invalid("Every value in paths must be a non-empty string.");
                paths[item.Name] = item.Value.GetString()!;
            }
            return paths;
        }
        catch (JsonException error) { throw Invalid("project.json is not strict JSON.", error); }
    }

    // Only a drive path. UNC and device paths could reach the network, which sends the user's credentials
    // during validation; alternate data stream syntax would also hide what a path names.
    internal static bool IsLocal(string fullPath) =>
        fullPath.Length >= 3 && char.IsAsciiLetter(fullPath[0]) && fullPath[1] == ':' &&
        fullPath[2] == Path.DirectorySeparatorChar && !fullPath.AsSpan(2).Contains(':');

    private void Validate()
    {
        // This comes first because every check below touches the file system.
        if (!new[] { Workbook, VbaSource, Tests, Workspace, Dist }.All(IsLocal))
            throw Invalid("Every value in paths must be a local drive path.");
        if (Path.GetExtension(Workbook).ToLowerInvariant() is not (".xlsx" or ".xlsm"))
            throw Invalid("paths.workbook must end with .xlsx or .xlsm.");
        // logs is included because it is what the session writes, and it can be a link even when workspace is not.
        foreach (string output in new[] { Workspace, Path.Combine(Workspace, "logs"), Dist })
            foreach (string input in new[] { Workbook, VbaSource, Tests })
                if (Overlaps(Resolved(output), Resolved(input)))
                    throw Invalid("paths.workspace and paths.dist must not overlap the workbook, sources or tests.");
        if (!File.Exists(Workbook)) throw Missing("paths.workbook");
        if (!Directory.Exists(VbaSource)) throw Missing("paths.vba_source");
        if (!Directory.Exists(Tests)) throw Missing("paths.tests");
    }

    private static bool Overlaps(string left, string right) => Contains(left, right) || Contains(right, left);

    private static bool Contains(string parent, string child) =>
        StringComparer.OrdinalIgnoreCase.Equals(parent, child) ||
        child.StartsWith(Path.TrimEndingDirectorySeparator(parent).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    // Follows links in every existing component, so a link cannot hide an overlap.
    private static string Resolved(string path)
    {
        string current = Path.GetPathRoot(path)!;
        foreach (string component in path[current.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries))
        {
            current = Path.Combine(current, component);
            FileSystemInfo item = Directory.Exists(current) ? new DirectoryInfo(current) : new FileInfo(current);
            if (item.Exists && item.LinkTarget is not null)
                current = item.ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? current;
        }
        return current;
    }

    private static CoreException Invalid(string message, Exception? inner = null) =>
        new(ErrorCode.InvalidRequest, ErrorDetail.InvalidConfig, message + RetryHint, inner);

    private static CoreException Missing(string key) =>
        new(ErrorCode.NotFound, ErrorDetail.PathNotFound, $"The path in {key} does not exist. Create it or fix project.json, then retry.");
}
