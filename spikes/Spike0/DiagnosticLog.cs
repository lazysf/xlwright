using System.Text.Json;

namespace Spike0;

internal sealed class DiagnosticLog(string path, S11Evidence? timing = null, TextWriter? output = null) : IDisposable
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
            if (timing is null)
            {
                writer.WriteLine(line);
                (output ?? Console.Out).WriteLine(line);
            }
            else
            {
                timing.Measure("diagnostic-file", () => writer.WriteLine(line));
                timing.Measure("diagnostic-stdout", () => (output ?? Console.Out).WriteLine(line));
            }
        }
    }

    public void Dispose() => writer.Dispose();
}
