using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace XlWright.McpServer.Tests;

// Checks that decide whether an Excel entry may start. Every run uses --preflight, which returns before any
// Excel code, so these tests cannot start Excel even if the checks break.
public sealed class SpikeEntryTests : IDisposable
{
    private static readonly string SpikeOutput = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "Spike0", new DirectoryInfo(AppContext.BaseDirectory).Name));
    private readonly string repository = Directory.CreateTempSubdirectory("xlwright-entry-").FullName;
    private readonly string prepared;

    public SpikeEntryTests()
    {
        File.WriteAllText(Path.Combine(repository, "global.json"), "{}");
        Directory.CreateDirectory(Path.Combine(repository, "spikes", "Spike0"));
        File.WriteAllText(Path.Combine(repository, "spikes", "Spike0", "Spike0.csproj"), "<Project />");
        Directory.CreateDirectory(Path.GetDirectoryName(GuardPath)!);

        prepared = Path.Combine(repository, ".xlwright", "development", "prepared", "experiment-test");
        Directory.CreateDirectory(prepared);
        var files = new List<object>();
        foreach (string source in Directory.EnumerateFiles(SpikeOutput))
        {
            string name = Path.GetFileName(source);
            File.Copy(source, Path.Combine(prepared, name));
            files.Add(new { path = name, sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))) });
        }
        File.WriteAllText(Path.Combine(prepared, "manifest.json"), JsonSerializer.Serialize(new { checkRun = "test", files }));
    }

    private string GuardPath => Path.Combine(repository, ".xlwright", "mcp-development", "active-smoke.json");

    public void Dispose()
    {
        // Antivirus can hold a just-run DLL; a leftover temporary directory must not fail the check.
        try { Directory.Delete(repository, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private async Task<(int ExitCode, string Output, string Error)> Preflight(string dll, string argument, CancellationToken cancellation)
    {
        var info = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = repository,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        info.ArgumentList.Add(dll);
        info.ArgumentList.Add("--preflight");
        info.ArgumentList.Add(argument);
        using var process = Process.Start(info)!;
        try
        {
            process.StandardInput.Close();
            Task<string> stdout = process.StandardOutput.ReadToEndAsync(cancellation);
            Task<string> stderr = process.StandardError.ReadToEndAsync(cancellation);
            await process.WaitForExitAsync(cancellation);
            return (process.ExitCode, await stdout, await stderr);
        }
        finally
        {
            if (!process.HasExited) process.Kill();
        }
    }

    public static TheoryData<string> ExcelEntries => new(
        "--excel", "--excel-end-refusal",
        "--excel-record-fault-starting", "--excel-record-fault-owned", "--excel-record-fault-completed",
        "--excel-s1", "--excel-s2", "--excel-s3", "--excel-s4",
        "--excel-s9-release", "--excel-s9-limit", "--excel-s9-deadline",
        "--excel-s10-a2", "--excel-s10-b", "--excel-s10-c",
        "--excel-s11r2-baseline", "--excel-s11r2-protected-view", "--excel-s11r2-wait", "--excel-s11r2-modal", "--excel-s11r2-shell2",
        "--excel-s14-auto", "--excel-s14-manual", "--excel-s14-mode-change",
        "--excel-t15-vba");

    [Theory(Timeout = 60000)]
    [MemberData(nameof(ExcelEntries))]
    public async Task EveryExcelEntryStopsAtARemainingGuard(string argument)
    {
        File.WriteAllText(GuardPath, "{}");

        var (exitCode, _, error) = await Preflight(Path.Combine(prepared, "Spike0.dll"), argument,
            TestContext.Current.CancellationToken);

        Assert.Equal(1, exitCode);
        Assert.Contains("experiment guard remains", error);
    }

    [Fact(Timeout = 60000)]
    public async Task GuardThatCannotBeConfirmedAbsentStopsTheEntry()
    {
        Directory.CreateDirectory(GuardPath);

        var (exitCode, _, error) = await Preflight(Path.Combine(prepared, "Spike0.dll"), "--excel",
            TestContext.Current.CancellationToken);

        Assert.Equal(1, exitCode);
        Assert.Contains("experiment guard remains", error);
    }

    [Theory(Timeout = 60000)]
    [MemberData(nameof(ExcelEntries))]
    public async Task EveryExcelEntryPassesWithoutGuard(string argument)
    {
        var (exitCode, output, _) = await Preflight(Path.Combine(prepared, "Spike0.dll"), argument,
            TestContext.Current.CancellationToken);

        Assert.Equal(0, exitCode);
        Assert.Contains("Excel was not started", output);
    }

    [Fact(Timeout = 60000)]
    public async Task UnpreparedBuildIsRefused()
    {
        var (exitCode, _, error) = await Preflight(Path.Combine(SpikeOutput, "Spike0.dll"), "--excel",
            TestContext.Current.CancellationToken);

        Assert.Equal(1, exitCode);
        Assert.Contains("npm run prepare:experiment", error);
    }

    [Fact(Timeout = 60000)]
    public async Task ChangedCopyIsRefused()
    {
        File.AppendAllText(Path.Combine(prepared, "Spike0.runtimeconfig.json"), " ");

        var (exitCode, _, error) = await Preflight(Path.Combine(prepared, "Spike0.dll"), "--excel",
            TestContext.Current.CancellationToken);

        Assert.Equal(1, exitCode);
        Assert.Contains("differs from its manifest", error);
    }
}
