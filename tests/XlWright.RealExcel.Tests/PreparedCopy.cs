using System.Security.Cryptography;
using System.Text.Json;
using XlWright.RealExcel.Tests;
using Xunit;

[assembly: AssemblyFixture(typeof(PreparedCopy))]

namespace XlWright.RealExcel.Tests;

// Real-Excel scenarios run only from a copy made by npm run prepare:product, so every run names a checked build.
// The fixture runs before any test; if it throws, every test fails and no scenario reaches the server.
// This stops accidents such as running a stale build. It does not protect against tampering by the same user.
public sealed class PreparedCopy
{
    internal string Repository { get; }
    internal string ServerPath { get; }
    internal string CheckRun { get; }
    internal string GuardPath { get; }

    public PreparedCopy()
    {
        // The copy keeps the artifacts layout: product-<id>/bin/XlWright.RealExcel.Tests/release/.
        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        string prepared = Path.GetDirectoryName(root)!;
        Repository = Path.GetFullPath(Path.Combine(prepared, "..", "..", ".."));
        if (!Path.GetFileName(root).StartsWith("product-", StringComparison.Ordinal) ||
            !string.Equals(prepared, Path.Combine(Repository, ".xlwright", "development", "prepared"), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Run real-Excel scenarios from a copy made by npm run prepare:product.");
        RejectReparsePoints(Repository, root);
        CheckRun = Verify(root, Repository);
        GuardPath = Path.Combine(Repository, ".xlwright", "mcp-development", "active-smoke.json");
        RejectReparsePoints(Repository, GuardPath);
        RequireNoGuard(GuardPath);
        ServerPath = Path.Combine(root, "bin", "XlWright.McpServer", "release", "XlWright.McpServer.dll");
    }

    private static string Verify(string root, string repository)
    {
        string manifestPath = Path.Combine(root, "manifest.json");
        if (!File.Exists(manifestPath))
            throw new InvalidDataException("The prepared copy has no manifest; run npm run prepare:product.");
        using var manifest = JsonDocument.Parse(File.ReadAllBytes(manifestPath));
        // The layout alone cannot tell a moved copy: the repository is derived from the copy's own location.
        if (!string.Equals(manifest.RootElement.GetProperty("repository").GetString(), repository, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The prepared copy was made for another location; run npm run prepare:product.");
        var expected = manifest.RootElement.GetProperty("files").EnumerateArray().ToDictionary(
            file => file.GetProperty("path").GetString()!, file => file.GetProperty("sha256").GetString()!);
        var actual = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(root, path).Replace('\\', '/'))
            .Where(path => path != "manifest.json").ToList();
        if (actual.Count != expected.Count || !actual.All(expected.ContainsKey))
            throw new InvalidDataException("The prepared copy differs from its manifest; run npm run prepare:product.");
        foreach (var (path, hash) in expected)
            if (Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(root, path)))) != hash)
                throw new InvalidDataException("The prepared copy differs from its manifest; run npm run prepare:product.");
        return manifest.RootElement.GetProperty("checkRun").GetString() ??
            throw new InvalidDataException("The prepared copy has no check run; run npm run prepare:product.");
    }

    private static void RequireNoGuard(string guardPath)
    {
        try { _ = File.GetAttributes(guardPath); }
        catch (FileNotFoundException) { return; }
        catch (DirectoryNotFoundException) { return; }
        // Unknown I/O failures propagate; only a confirmed missing guard permits a run.
        throw new InvalidOperationException("An experiment guard remains; manual review is required. No scenario was run.");
    }

    private static void RejectReparsePoints(string repository, string path)
    {
        string relative = Path.GetRelativePath(repository, path);
        if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar))
            throw new IOException("Scenario paths must stay within the repository.");
        string current = repository;
        foreach (string component in relative.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, component);
            if ((File.Exists(current) || Directory.Exists(current)) &&
                File.GetAttributes(current).HasFlag(FileAttributes.ReparsePoint))
                throw new IOException("Scenario paths may not use reparse points.");
        }
    }
}
