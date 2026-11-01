using System.Text.Json;

namespace XlWright.Core;

// COM details and stack traces stay here; responses carry only this file's path. One file per started session.
internal sealed class DiagnosticLog(string directory)
{
    private readonly string path = Path.Combine(directory, $"session-{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}-{Guid.NewGuid():N}.jsonl");

    // Returns null when the log could not be written: a diagnostic failure must not replace the reported failure.
    internal string? Write(string name, object details)
    {
        try
        {
            Directory.CreateDirectory(directory);
            File.AppendAllText(path, JsonSerializer.Serialize(new { at_utc = DateTime.UtcNow, name, details }) + "\n");
            return path;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return null; }
    }
}
