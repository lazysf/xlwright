using XlWright.RealExcel.Tests;
using Xunit;

namespace XlWright.McpServer.Tests;

public sealed class ScenarioGuardTests : IDisposable
{
    private readonly string directory = Directory.CreateTempSubdirectory("xlwright-guard-").FullName;
    private string GuardPath => Path.Combine(directory, "guard", "active-smoke.json");

    public void Dispose() => Directory.Delete(directory, recursive: true);

    [Fact]
    public void ReleaseRemovesTheGuardOfTheSameRun()
    {
        var guard = ScenarioGuard.Create(GuardPath, "scenario");
        Assert.True(File.Exists(GuardPath));

        guard.Release();

        Assert.False(File.Exists(GuardPath));
    }

    [Fact]
    public void GuardRemainsUntilReleased()
    {
        _ = ScenarioGuard.Create(GuardPath, "scenario");

        Assert.True(File.Exists(GuardPath));
    }

    [Fact]
    public void SecondRunIsRefused()
    {
        _ = ScenarioGuard.Create(GuardPath, "first");

        Assert.Throws<IOException>(() => ScenarioGuard.Create(GuardPath, "second"));
    }

    [Fact]
    public void ChangedGuardIsPreserved()
    {
        var guard = ScenarioGuard.Create(GuardPath, "scenario");
        string changed = "{\"scenario\":\"other\",\"run\":\"other\"}";
        File.WriteAllText(GuardPath, changed);

        Assert.Throws<InvalidDataException>(guard.Release);

        Assert.Equal(changed, File.ReadAllText(GuardPath));
    }
}
