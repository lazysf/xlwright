using System.Text.Json;

namespace XlWright.RealExcel.Tests;

// The guard Spike0 uses too. CreateNew makes concurrent real-Excel runs fail, and a run that does not confirm its
// normal end leaves the guard, which stops the next run until a person has reviewed the evidence.
internal sealed class ScenarioGuard
{
    private readonly string path;
    private readonly string scenario;
    private readonly string run;

    private ScenarioGuard(string path, string scenario, string run)
    {
        this.path = path;
        this.scenario = scenario;
        this.run = run;
    }

    internal static ScenarioGuard Create(string path, string scenario)
    {
        string run = Guid.NewGuid().ToString("N");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var guard = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        JsonSerializer.Serialize(guard, new { scenario, run });
        guard.Flush(true);
        return new(path, scenario, run);
    }

    // Call only after the scenario has confirmed its normal end.
    internal void Release()
    {
        using (var guard = JsonDocument.Parse(File.ReadAllText(path)))
            if (guard.RootElement.GetProperty("run").GetString() != run ||
                guard.RootElement.GetProperty("scenario").GetString() != scenario)
                throw new InvalidDataException("Experiment guard changed; preserve it for review.");
        File.Delete(path);
    }
}
