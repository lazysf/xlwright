using System.Text.Json;

namespace Spike0;

internal sealed class DiagnosticLog(string path) : IDisposable
{
    private readonly object sync = new();
    private readonly StreamWriter writer = new(new FileStream(path, FileMode.CreateNew,
        FileAccess.Write, FileShare.Read)) { AutoFlush = true };

    internal void Write(string observation, ExecutionGate gate, object? detail = null)
    {
        lock (sync)
        {
            var line = JsonSerializer.Serialize(new
            {
                at_utc = DateTime.UtcNow,
                observation,
                snapshot = gate.Snapshot,
                detail
            });
            writer.WriteLine(line);
            Console.WriteLine(line);
        }
    }

    public void Dispose() => writer.Dispose();
}
