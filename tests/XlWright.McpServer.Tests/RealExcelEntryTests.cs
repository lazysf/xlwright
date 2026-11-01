using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace XlWright.McpServer.Tests;

// Checks the fixture that decides whether real-Excel scenarios may run. Every run selects only the connection
// scenario, which does not reach Excel, so scenarios added later cannot start Excel from npm run check.
public sealed class RealExcelEntryTests : IDisposable
{
    private const string ConnectionScenario = "XlWright.RealExcel.Tests.CopiedServerTests.ConnectionToTheCopiedServerSpansRequests";
    private static readonly string BuildDirectory = new DirectoryInfo(AppContext.BaseDirectory).Name;
    private static readonly string Artifacts = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", ".."));
    private static readonly string ScenarioOutput = Path.Combine(Artifacts, "XlWright.RealExcel.Tests", BuildDirectory);
    private readonly string repository = Directory.CreateTempSubdirectory("xlwright-real-excel-").FullName;
    private readonly string copy;

    public RealExcelEntryTests()
    {
        copy = Path.Combine(repository, ".xlwright", "development", "prepared", "product-test");
        foreach (string project in new[] { "XlWright.RealExcel.Tests", "XlWright.McpServer" })
            CopyDirectory(Path.Combine(Artifacts, project, BuildDirectory), Path.Combine(copy, "bin", project, "release"));
        var files = Directory.EnumerateFiles(copy, "*", SearchOption.AllDirectories).Select(path => new
        {
            path = Path.GetRelativePath(copy, path).Replace('\\', '/'),
            sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)))
        }).ToList();
        File.WriteAllText(Path.Combine(copy, "manifest.json"), JsonSerializer.Serialize(new { checkRun = "test", files }));
    }

    private string GuardPath => Path.Combine(repository, ".xlwright", "mcp-development", "active-smoke.json");

    public void Dispose()
    {
        // Antivirus can hold a just-run DLL; a leftover temporary directory must not fail the check.
        try { Directory.Delete(repository, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static void CopyDirectory(string source, string target)
    {
        foreach (string path in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            string destination = Path.Combine(target, Path.GetRelativePath(source, path));
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(path, destination);
        }
    }

    private async Task<(int ExitCode, string Output)> Run(string dll, CancellationToken cancellation)
    {
        var info = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = repository,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        info.ArgumentList.Add(dll);
        info.ArgumentList.Add("-method");
        info.ArgumentList.Add(ConnectionScenario);
        using var process = Process.Start(info)!;
        try
        {
            process.StandardInput.Close();
            Task<string> stdout = process.StandardOutput.ReadToEndAsync(cancellation);
            Task<string> stderr = process.StandardError.ReadToEndAsync(cancellation);
            await process.WaitForExitAsync(cancellation);
            return (process.ExitCode, await stdout + await stderr);
        }
        finally
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        }
    }

    private string CopiedScenarios => Path.Combine(copy, "bin", "XlWright.RealExcel.Tests", "release", "XlWright.RealExcel.Tests.dll");

    [Fact(Timeout = 120000)]
    public async Task PreparedCopyRunsTheConnectionScenario()
    {
        var (exitCode, output) = await Run(CopiedScenarios, TestContext.Current.CancellationToken);

        Assert.True(exitCode == 0, output);
        // xUnit exits 0 when the filter matches nothing, so a renamed scenario must not pass silently.
        Assert.Contains("Total: 1,", output);
    }

    [Fact(Timeout = 120000)]
    public async Task UnpreparedBuildIsRefused()
    {
        var (exitCode, output) = await Run(Path.Combine(ScenarioOutput, "XlWright.RealExcel.Tests.dll"),
            TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("npm run prepare:product", output);
    }

    [Fact(Timeout = 120000)]
    public async Task ChangedCopyIsRefused()
    {
        File.AppendAllText(Path.Combine(copy, "bin", "XlWright.McpServer", "release", "XlWright.McpServer.runtimeconfig.json"), " ");

        var (exitCode, output) = await Run(CopiedScenarios, TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("differs from its manifest", output);
    }

    [Fact(Timeout = 120000)]
    public async Task ExtraFileInCopyIsRefused()
    {
        File.WriteAllText(Path.Combine(copy, "bin", "extra.dll"), "");

        var (exitCode, output) = await Run(CopiedScenarios, TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("differs from its manifest", output);
    }

    [Fact(Timeout = 120000)]
    public async Task RemainingGuardIsRefused()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(GuardPath)!);
        File.WriteAllText(GuardPath, "{}");

        var (exitCode, output) = await Run(CopiedScenarios, TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("experiment guard remains", output);
    }
}
