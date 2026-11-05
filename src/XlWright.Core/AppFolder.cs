using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>
/// An app folder inside the repository and its main workbook (docs/architecture.md "アプリのフォルダ"). The workbook is null
/// for a session without one, until workbook_create makes it (docs/api.md "主ブックの無いセッション").
/// </summary>
public sealed record AppFolder(string FullPath, string RelativePath, string? WorkbookPath, string? WorkbookName)
{
    public bool HasWorkbook => WorkbookName is not null;

    /// <summary>The folder with the workbook that workbook_create saved directly under it as the main workbook.</summary>
    public AppFolder WithWorkbook(string name) => this with { WorkbookPath = Path.Combine(FullPath, name), WorkbookName = name };

    /// <summary>
    /// Checks the file name workbook_create is given (docs/api.md "workbook_create"): a name without folders, ending in
    /// .xlsm or .xlsx, not an owner file (~$), and one Windows accepts. Returns null when it is good.
    /// </summary>
    public static ToolFailure? CheckNewWorkbookName(string name)
    {
        var stem = Path.GetFileNameWithoutExtension(name);
        var good = WorkbookExtensions.Contains(Path.GetExtension(name), StringComparer.OrdinalIgnoreCase)
            && stem.Length > 0 && stem == stem.Trim()
            && !name.StartsWith("~$", StringComparison.Ordinal)
            && name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;
        return good
            ? null
            : new ToolFailure(ToolFailure.InvalidRequest, "INVALID_WORKBOOK_NAME",
                $"name must be a file name without folders that ends in .xlsm or .xlsx, such as Sales.xlsm. Rejected: {name}");
    }

    private static readonly string[] WorkbookExtensions = [".xlsm", ".xlsx"];

    /// <summary>
    /// Resolves <paramref name="app"/> (relative to <paramref name="repository"/>, or absolute) and finds the main workbook.
    /// Checks the path before touching it (docs/architecture.md "パスの制限").
    /// </summary>
    public static (AppFolder? Folder, ToolFailure? Failure) Resolve(string repository, string app)
    {
        if (string.IsNullOrWhiteSpace(app))
        {
            return Fail(ToolFailure.InvalidRequest, null, "app is required: pass the app folder, relative to the repository or absolute.");
        }

        // UNC and device paths (\\server\share, \\?\, \\.\) are rejected as text, before any file system call.
        if (app.StartsWith(@"\\", StringComparison.Ordinal) || app.StartsWith("//", StringComparison.Ordinal))
        {
            return NotAllowed(app, "UNC and device paths are not supported.");
        }

        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(repository));
        var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.Combine(root, app)));
        if (!IsSameOrInside(full, root))
        {
            return NotAllowed(app, "The app folder must be inside the repository.");
        }

        var drive = Path.GetPathRoot(full);
        if (string.IsNullOrEmpty(drive) || drive.StartsWith(@"\\", StringComparison.Ordinal) || new DriveInfo(drive).DriveType == DriveType.Network)
        {
            return NotAllowed(app, "Network drives are not supported.");
        }

        if (!Directory.Exists(full))
        {
            return Fail(ToolFailure.NotFound, "APP_NOT_FOUND", $"The app folder does not exist: {app}. Check the path.");
        }

        var workbooks = new DirectoryInfo(full).EnumerateFiles()
            .Where(file => WorkbookExtensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase))
            .Where(file => !file.Name.StartsWith("~$", StringComparison.Ordinal))
            .Where(file => !file.Attributes.HasFlag(FileAttributes.Hidden))
            .Select(file => file.Name)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var relative = Path.GetRelativePath(root, full).Replace(Path.DirectorySeparatorChar, '/');
        if (workbooks.Count == 0)
        {
            return (new AppFolder(full, relative, null, null), null);
        }

        if (workbooks.Count > 1)
        {
            var data = new JsonObject { ["workbooks"] = new JsonArray(workbooks.Select(name => JsonValue.Create(name)).ToArray<JsonNode?>()) };
            return Fail(ToolFailure.InvalidRequest, "MULTIPLE_WORKBOOKS", "The app folder has more than one workbook directly under it. Keep only the main workbook there; move others into a subfolder.", data);
        }

        return (new AppFolder(full, relative, Path.Combine(full, workbooks[0]), workbooks[0]), null);
    }

    // Compares full paths with the trailing separator, ignoring case, so apps\sales2 is not inside apps\sales.
    private static bool IsSameOrInside(string path, string root) =>
        string.Equals(path, root, StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    private static (AppFolder?, ToolFailure?) NotAllowed(string app, string reason) =>
        Fail(ToolFailure.InvalidRequest, "PATH_NOT_ALLOWED", $"{reason} Rejected: {app}");

    private static (AppFolder?, ToolFailure?) Fail(string code, string? detail, string message, JsonObject? data = null) =>
        (null, new ToolFailure(code, detail, message, data));
}
