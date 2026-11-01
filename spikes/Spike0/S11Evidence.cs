using System.Diagnostics;
using System.Text.Json;

namespace Spike0;

// S11 has its own file and lock, so a blocked console writer cannot hold up window evidence.
internal sealed class S11Evidence(string path) : IDisposable
{
    private readonly object sync = new();
    private readonly StreamWriter writer = new(new FileStream(path, FileMode.CreateNew,
        FileAccess.Write, FileShare.Read)) { AutoFlush = true };

    internal void Write(string phase, string outcome, object? detail = null)
    {
        var line = JsonSerializer.Serialize(new { at_utc = DateTime.UtcNow,
            timestamp = Stopwatch.GetTimestamp(), frequency = Stopwatch.Frequency, phase, outcome, detail });
        lock (sync) writer.WriteLine(line);
    }

    internal void Measure(string phase, Action action)
    {
        Write(phase, "started");
        long started = Stopwatch.GetTimestamp();
        try
        {
            action();
            Write(phase, "returned", new { elapsed_ms = Stopwatch.GetElapsedTime(started).TotalMilliseconds });
        }
        catch (Exception error)
        {
            Write(phase, "failed", new { elapsed_ms = Stopwatch.GetElapsedTime(started).TotalMilliseconds,
                error = error.GetType().Name, hresult = error.HResult });
            throw;
        }
    }

    public void Dispose() { lock (sync) writer.Dispose(); }
}
