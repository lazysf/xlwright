using XlWright.Core;

namespace XlWright.McpServer;

/// <summary>
/// Arguments of the show command (docs/api.md "表示要求のコマンド"): show --project &lt;repository&gt; --app &lt;app&gt;.
/// The app folder passes the same checks as session_start's app.
/// </summary>
internal static class ShowOptions
{
    public const string Usage = "Usage: XlWright.McpServer show --project <repository> --app <app folder>";

    /// <summary>Returns the app folder, or null with an error message that names the offending argument.</summary>
    public static AppFolder? Parse(IReadOnlyList<string> args, out string? error)
    {
        string? project = null;
        string? app = null;
        for (var i = 0; i < args.Count; i++)
        {
            var hasValue = i + 1 < args.Count && !args[i + 1].StartsWith("--", StringComparison.Ordinal);
            switch (args[i])
            {
                case "--project" when project is null && hasValue:
                    project = args[++i];
                    break;
                case "--app" when app is null && hasValue:
                    app = args[++i];
                    break;
                case "--project" or "--app":
                    error = $"{args[i]} is given more than once or without a value.";
                    return null;
                default:
                    error = $"Unknown argument: {args[i]}";
                    return null;
            }
        }

        if (project is null || !Directory.Exists(project))
        {
            error = "--project <repository> is required and must be an existing folder.";
            return null;
        }

        if (app is null)
        {
            error = "--app <app folder> is required.";
            return null;
        }

        var (folder, failure) = AppFolder.Resolve(Path.GetFullPath(project), app);
        error = failure is null ? null : $"--app: {failure.Detail ?? failure.Code}: {failure.Message}";
        return folder;
    }
}
