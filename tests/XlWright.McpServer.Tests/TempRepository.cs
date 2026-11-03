namespace XlWright.McpServer.Tests;

// A throwaway folder that stands in for the repository passed with --project.
internal sealed class TempRepository : IDisposable
{
    public string Root { get; } = Path.Combine(Path.GetTempPath(), "xlwright-tests", Guid.NewGuid().ToString("N"));

    public TempRepository() => Directory.CreateDirectory(Root);

    // Creates the app folder and empty files at the given paths under it. The files only need to exist.
    public void AddApp(string app, params string[] files)
    {
        var folder = Path.Combine(Root, app);
        Directory.CreateDirectory(folder);
        foreach (var file in files)
        {
            var path = Path.Combine(folder, file);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, []);
        }
    }

    public void Dispose()
    {
        foreach (var file in Directory.EnumerateFiles(Root, "*", SearchOption.AllDirectories))
        {
            File.SetAttributes(file, FileAttributes.Normal);
        }

        Directory.Delete(Root, recursive: true);
    }
}
