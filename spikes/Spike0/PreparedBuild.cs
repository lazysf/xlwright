using System.Security.Cryptography;
using System.Text.Json;

namespace Spike0;

// Excel entries run only from a copy made by npm run prepare:experiment, so every run names a checked build.
// This stops accidents such as running a stale DLL. It does not protect against tampering by the same user.
internal static class PreparedBuild
{
    internal sealed record Identity(string CheckRun, string Assembly);

    internal static Identity Require(string repository, string directory)
    {
        string prepared = Path.Combine(repository, ".xlwright", "development", "prepared");
        string location = Path.GetFullPath(Path.TrimEndingDirectorySeparator(directory));
        if (!string.Equals(Path.GetDirectoryName(location), prepared, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Run Excel experiments from a copy made by npm run prepare:experiment.");
        RuntimeAssets.RejectReparsePoints(repository, location);
        return Verify(location);
    }

    internal static Identity Verify(string directory)
    {
        string manifestPath = Path.Combine(directory, "manifest.json");
        if (!File.Exists(manifestPath))
            throw new InvalidDataException("The prepared copy has no manifest; run npm run prepare:experiment.");
        using var manifest = JsonDocument.Parse(File.ReadAllBytes(manifestPath));
        var expected = manifest.RootElement.GetProperty("files").EnumerateArray().ToDictionary(
            file => file.GetProperty("path").GetString()!, file => file.GetProperty("sha256").GetString()!);
        var actual = Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(directory, path).Replace('\\', '/'))
            .Where(path => path != "manifest.json").ToList();
        if (actual.Count != expected.Count || !actual.All(expected.ContainsKey))
            throw new InvalidDataException("The prepared copy differs from its manifest; run npm run prepare:experiment.");
        foreach (var (path, hash) in expected)
            if (Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(directory, path)))) != hash)
                throw new InvalidDataException("The prepared copy differs from its manifest; run npm run prepare:experiment.");
        string checkRun = manifest.RootElement.GetProperty("checkRun").GetString() ??
            throw new InvalidDataException("The prepared copy has no check run; run npm run prepare:experiment.");
        return new(checkRun, expected["Spike0.dll"]);
    }
}
