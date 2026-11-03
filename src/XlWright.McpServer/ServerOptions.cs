namespace XlWright.McpServer;

/// <summary>
/// Command-line options: --project &lt;repository&gt; and --no-excel.
/// </summary>
internal sealed record ServerOptions(string? Project, bool NoExcel)
{
    /// <summary>
    /// Returns the options, or null with an error message that names the offending argument.
    /// </summary>
    public static ServerOptions? Parse(IReadOnlyList<string> args, out string? error)
    {
        string? project = null;
        var noExcel = false;
        for (var i = 0; i < args.Count; i++)
        {
            switch (args[i])
            {
                case "--project" when i + 1 < args.Count:
                    project = args[++i];
                    break;
                case "--project":
                    error = "--project requires a repository path.";
                    return null;
                case "--no-excel":
                    noExcel = true;
                    break;
                default:
                    error = $"Unknown argument: {args[i]}";
                    return null;
            }
        }

        error = null;
        return new ServerOptions(project, noExcel);
    }
}
