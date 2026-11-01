using System.Diagnostics;
using Xunit;

namespace XlWright.McpServer.Tests;

// Runs the manual recovery script with -CheckOnly against synthetic records. Every case is refused before the
// script looks for windows, so no display request is sent and no Excel window is touched.
public sealed class ShowOwnedExcelTests : IDisposable
{
    private static readonly string Script = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "..", "..", "..", "scripts", "Show-OwnedExcel.ps1"));
    private readonly string directory = Directory.CreateTempSubdirectory("xlwright-show-").FullName;

    public void Dispose() => Directory.Delete(directory, recursive: true);

    private async Task<(int ExitCode, string Output)> CheckOnly(string recordText, CancellationToken cancellation,
        Func<IDisposable>? hold = null)
    {
        string record = Path.Combine(directory, ".xlwright.lock");
        File.WriteAllText(record, recordText);
        using var held = hold?.Invoke();
        var info = new ProcessStartInfo("powershell.exe")
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = directory,
            RedirectStandardOutput = true, RedirectStandardError = true
        };
        foreach (string argument in new[] { "-NoLogo", "-NoProfile", "-NonInteractive", "-File", Script, "-RecordPath", record, "-CheckOnly" })
            info.ArgumentList.Add(argument);
        using var process = Process.Start(info)!;
        try
        {
            Task<string> stdout = process.StandardOutput.ReadToEndAsync(cancellation);
            Task<string> stderr = process.StandardError.ReadToEndAsync(cancellation);
            await process.WaitForExitAsync(cancellation);
            return (process.ExitCode, await stdout + await stderr);
        }
        finally
        {
            if (!process.HasExited) process.Kill();
        }
    }

    public static TheoryData<string, string> RefusedRecords => new()
    {
        { "{", "not valid JSON" },
        { "{\"version\":1,\"kind\":\"completed\"}", "not an owned record" },
        { "{\"version\":\"1\",\"kind\":\"OWNED\",\"process\":{\"pid\":1,\"started_at_utc\":\"2026-10-30T00:00:00.0000000Z\",\"name\":\"EXCEL\"}}", "not an owned record" },
        { "{\"version\":1,\"kind\":\"starting\",\"workbook_path\":\"C:\\\\book.xlsx\"}", "not an owned record" },
        {
            "{\"version\":1,\"kind\":\"owned\",\"workbook_path\":\"C:\\\\book.xlsx\",\"process\":" +
            "{\"pid\":2147483647,\"started_at_utc\":\"2026-10-30T00:00:00.0000000Z\",\"name\":\"EXCEL\"}}",
            "not running"
        }
    };

    [Theory(Timeout = 60000)]
    [MemberData(nameof(RefusedRecords))]
    public async Task RecordsWithoutAShowableExcelAreRefused(string recordText, string reason)
    {
        var (exitCode, output) = await CheckOnly(recordText, TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains(reason, output);
        Assert.False(Directory.Exists(Path.Combine(directory, ".xlwright", "recovery")));
    }

    // A server holds the record with FileShare.None while it runs.
    [Fact(Timeout = 60000)]
    public async Task RecordHeldByAServerIsRefused()
    {
        string record = Path.Combine(directory, ".xlwright.lock");
        var (exitCode, output) = await CheckOnly("{\"version\":1,\"kind\":\"completed\"}", TestContext.Current.CancellationToken,
            hold: () => new FileStream(record, FileMode.Open, FileAccess.ReadWrite, FileShare.None));

        Assert.NotEqual(0, exitCode);
        Assert.Contains("held by another process", output);
    }

    [Fact(Timeout = 60000)]
    public async Task ChangedIdentityIsRefused()
    {
        using var current = Process.GetCurrentProcess();
        string started = current.StartTime.ToUniversalTime().AddSeconds(-1).ToString("O");
        var (exitCode, output) = await CheckOnly(
            "{\"version\":1,\"kind\":\"owned\",\"workbook_path\":\"C:\\\\book.xlsx\",\"process\":" +
            $"{{\"pid\":{current.Id},\"started_at_utc\":\"{started}\",\"name\":\"{current.ProcessName}\"}}}}", TestContext.Current.CancellationToken);

        Assert.NotEqual(0, exitCode);
        Assert.Contains("identity does not match", output);
    }
}
