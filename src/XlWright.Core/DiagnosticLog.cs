using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>
/// The diagnostic log of one session_start: &lt;app&gt;/.xlwright/logs/&lt;session-id&gt;.jsonl, one JSON object per line
/// (docs/architecture.md "診断の記録の形式"). It never holds values, sources, dialog text, or messages. A failure to
/// write never fails the tool; the log is then left out of the responses.
/// </summary>
public sealed class DiagnosticLog
{
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    private readonly object _gate = new();
    private readonly string _path;
    private readonly string _relativePath;
    private bool _broken;

    private DiagnosticLog(string path, string relativePath)
    {
        _path = path;
        _relativePath = relativePath;
    }

    /// <summary>The path from the repository with '/', or null once the log could not be written.</summary>
    public string? RelativePath
    {
        get { lock (_gate) return _broken ? null : _relativePath; }
    }

    public static DiagnosticLog Create(AppFolder app)
    {
        var id = $"{DateTime.UtcNow:yyyyMMdd'T'HHmmss'Z'}-{RandomNumberGenerator.GetHexString(8, lowercase: true)}";
        var folder = Path.Combine(app.FullPath, ".xlwright", "logs");
        var log = new DiagnosticLog(Path.Combine(folder, id + ".jsonl"), $"{app.RelativePath}/.xlwright/logs/{id}.jsonl");
        try
        {
            Directory.CreateDirectory(folder);
            new FileStream(log._path, FileMode.CreateNew, FileAccess.Write, FileShare.Read).Dispose();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            log._broken = true;
        }

        return log;
    }

    /// <summary>Appends one line with time_utc, event, and the fields; each line is flushed by closing the file.</summary>
    public void Write(string name, JsonObject? fields = null)
    {
        var line = new JsonObject { ["time_utc"] = DateTime.UtcNow.ToString("O"), ["event"] = name };
        foreach (var (key, value) in fields ?? [])
        {
            line[key] = value?.DeepClone();
        }

        lock (_gate)
        {
            if (_broken)
            {
                return;
            }

            try
            {
                File.AppendAllText(_path, line.ToJsonString(JsonText.Options) + "\n", Utf8);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                _broken = true;
            }
        }
    }
}
