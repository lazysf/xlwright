using System.Diagnostics;
using ModelContextProtocol.Client;

namespace XlWright.McpServer.Tests;

// Launches XlWright.McpServer.dll from the test output folder with `dotnet`.
internal static class ServerProcess
{
    public static string DllPath { get; } = Path.Combine(AppContext.BaseDirectory, "XlWright.McpServer.dll");

    public static Task<McpClient> ConnectAsync(CancellationToken cancellationToken) =>
        ConnectAsync(Directory.GetCurrentDirectory(), cancellationToken);

    public static Task<McpClient> ConnectAsync(string project, CancellationToken cancellationToken, bool noExcel = true)
    {
        var transport = new StdioClientTransport(new StdioClientTransportOptions
        {
            Name = "xlwright",
            Command = "dotnet",
            Arguments = noExcel ? [DllPath, "--project", project, "--no-excel"] : [DllPath, "--project", project],
        });
        return McpClient.CreateAsync(transport, cancellationToken: cancellationToken);
    }

    public static async Task<(int ExitCode, string StandardError)> RunAsync(
        IEnumerable<string> arguments,
        CancellationToken cancellationToken)
    {
        var startInfo = new ProcessStartInfo("dotnet")
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        startInfo.ArgumentList.Add(DllPath);
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("dotnet did not start.");
        process.StandardInput.Close();
        var standardError = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.StandardOutput.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        return (process.ExitCode, await standardError);
    }
}
