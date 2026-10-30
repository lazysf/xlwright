using System.ComponentModel;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;

namespace XlWright.McpServer;

// The development adapter owns only a fixed Spike0 process, never an Excel COM reference.
internal sealed class DevelopmentTools : IAsyncDisposable
{
    private readonly object sync = new();
    private readonly string repository = Path.GetFullPath(Environment.CurrentDirectory);
    private readonly string invocation = Guid.NewGuid().ToString("N");
    private readonly string build = Program.BuildHash;
    private readonly string spike;
    private readonly string spikeHash;
    private readonly string directory;
    private readonly string guardPath;
    private Task? running;
    private Process? child;
    private string? receipt;
    private string? scenario;
    private string state = "idle";
    private int? exitCode;
    private string? latestObservation;
    private string? errorType;
    private bool stopping;

    internal DevelopmentTools()
    {
        if (!File.Exists(Path.Combine(repository, "spikes", "Spike0", "Spike0.csproj")))
            throw new InvalidOperationException("Development mode requires the repository working directory.");
        spike = Path.Combine(repository, "spikes", "Spike0", "bin", "Release", "net10.0-windows", "Spike0.dll");
        RejectLinks(spike);
        spikeHash = Hash(spike);
        directory = Path.Combine(repository, ".xlwright", "mcp-development", invocation);
        guardPath = Path.Combine(repository, ".xlwright", "mcp-development", "active-smoke.json");
        RejectLinks(directory);
        RejectLinks(guardPath);
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "server.json"), JsonSerializer.Serialize(new
        {
            invocation, build, spikeHash, pid = Environment.ProcessId,
            startedAtUtc = Process.GetCurrentProcess().StartTime.ToUniversalTime(),
            atUtc = DateTime.UtcNow
        }));
        Console.Error.WriteLine($"xlwright development {invocation} started");
    }

    [Description("Read development server identity and current smoke status without Excel access.")]
    internal string Status()
    {
        lock (sync) return Snapshot();
    }

    [Description("Start the fixed Spike0 Excel smoke once. Requires explicit approval of this run. Do not retry after a lost response; use dev_status. Changes only a generated fixture copy and normally closes its owned Excel.")]
    internal string Start()
        => StartFixed("smoke", "--excel");

    [Description("Start the fixed L3 end-refusal experiment once. Requires explicit approval. Leaves ownership and guard for manual recovery even when refusal is expected. Do not retry.")]
    internal string StartEndRefusal()
        => StartFixed("end-refusal", "--excel-end-refusal");

    private string StartFixed(string experiment, string argument)
    {
        lock (sync)
        {
            if (stopping || running is { IsCompleted: false } || state is "failed" or "unknown")
                throw new InvalidOperationException("Smoke is busy or requires manual review; do not retry.");
            RejectLinks(spike);
            RejectLinks(guardPath);
            if (Hash(spike) != spikeHash)
                throw new InvalidOperationException("Spike build changed. Stop and reconnect before running.");

            // This preflight does not replace Spike0's exclusive record validation.
            string record = Path.Combine(repository, ".xlwright", "spike0", ".xlwright.lock");
            RejectLinks(record);
            using (var document = JsonDocument.Parse(File.ReadAllText(record)))
                if (document.RootElement.GetProperty("kind").GetString() != "completed")
                    throw new InvalidOperationException("Existing ownership record requires manual review.");

            string dotnet = Environment.ProcessPath ?? throw new InvalidOperationException("No host path.");
            if (!Path.GetFileName(dotnet).Equals("dotnet.exe", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Development mode must be launched with dotnet.exe.");

            receipt = Guid.NewGuid().ToString("D");
            scenario = experiment;
            // CreateNew prevents simultaneous or forgotten experiments across server restarts.
            using (var guard = new FileStream(guardPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                JsonSerializer.Serialize(guard, new { invocation, receipt, scenario, atUtc = DateTime.UtcNow });
                guard.Flush(true);
            }
            state = "running";
            exitCode = null;
            latestObservation = null;
            errorType = null;
            var info = new ProcessStartInfo(dotnet)
            {
                UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = repository,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
            };
            info.ArgumentList.Add(spike);
            info.ArgumentList.Add(argument);
            try
            {
                child?.Dispose();
                child = Process.Start(info) ?? throw new InvalidOperationException("Child did not start.");
                File.WriteAllText(Path.Combine(directory, receipt + ".process.json"), JsonSerializer.Serialize(new
                {
                    receipt, pid = child.Id, startedAtUtc = child.StartTime.ToUniversalTime(), spikeHash
                }));
                running = ObserveAsync(child, receipt);
            }
            catch (Exception error)
            {
                state = "unknown";
                errorType = error.GetType().Name;
                child?.StandardInput.Close();
                throw new InvalidOperationException("Start failed; preserve records and inspect dev_status.");
            }
            return Snapshot();
        }
    }

    [Description("Read a previously accepted fixed experiment by receipt. This never resends the Excel operation.")]
    internal string SmokeStatus(Guid receiptId)
    {
        lock (sync)
        {
            if (receiptId.ToString("D") != receipt) throw new ArgumentException("Unknown receipt.");
            return Snapshot();
        }
    }

    private string Snapshot() => JsonSerializer.Serialize(new
    {
        invocation, build, spikeHash, receipt, scenario, state, exitCode, latestObservation, errorType,
        logDirectory = Path.GetRelativePath(repository, directory).Replace('\\', '/'),
        manualReviewRequired = state is "failed" or "unknown" || File.Exists(guardPath) && state == "idle"
    });

    private async Task ObserveAsync(Process process, string acceptedReceipt)
    {
        try
        {
            Task stdout = DrainAsync(process.StandardOutput, acceptedReceipt + ".stdout.jsonl", true);
            Task stderr = DrainAsync(process.StandardError, acceptedReceipt + ".stderr.log", false);
            await Task.WhenAll(process.WaitForExitAsync(), stdout, stderr).WaitAsync(TimeSpan.FromSeconds(360));
            lock (sync)
            {
                exitCode = process.ExitCode;
                state = scenario == "smoke" && exitCode == 0 && latestObservation == "saved-file.verified" ? "completed" :
                    exitCode == 2 ? "unknown" : "failed";
                WriteResult();
                if (state == "completed") File.Delete(guardPath);
            }
        }
        catch (Exception error)
        {
            try
            {
                lock (sync)
                {
                    state = "unknown";
                    errorType = error.GetType().Name;
                    WriteUnknownResult();
                }
            }
            finally
            {
                try { process.StandardInput.Close(); } catch (IOException) { }
            }
        }
    }

    private async Task DrainAsync(StreamReader reader, string filename, bool observations)
    {
        using var writer = new StreamWriter(new FileStream(Path.Combine(directory, filename),
            FileMode.CreateNew, FileAccess.Write, FileShare.Read)) { AutoFlush = true };
        while (await reader.ReadLineAsync() is { } line)
        {
            await writer.WriteLineAsync(line);
            if (!observations) continue;
            using var document = JsonDocument.Parse(line);
            if (document.RootElement.TryGetProperty("observation", out var observation))
                lock (sync) latestObservation = observation.GetString();
        }
    }

    private void WriteResult() => File.WriteAllText(Path.Combine(directory, receipt + ".result.json"), Snapshot());

    private void WriteUnknownResult()
    {
        try { WriteResult(); }
        catch (IOException) { Console.Error.WriteLine("xlwright smoke result could not be saved; preserve guard."); }
        catch (UnauthorizedAccessException) { Console.Error.WriteLine("xlwright smoke result could not be saved; preserve guard."); }
    }

    public async ValueTask DisposeAsync()
    {
        Task? pending;
        lock (sync)
        {
            stopping = true;
            pending = running;
            if (child is not null)
                try { child.StandardInput.Close(); } catch (IOException) { }
        }
        if (pending is not null)
        {
            try { await pending.WaitAsync(TimeSpan.FromSeconds(125)); }
            catch (TimeoutException)
            {
                lock (sync) { state = "unknown"; errorType = "ShutdownUnconfirmed"; WriteUnknownResult(); }
            }
        }
        Console.Error.WriteLine($"xlwright development {invocation} stopped; {state}");
        // Dispose releases the local handle; it does not terminate an unconfirmed process.
        child?.Dispose();
    }

    private void RejectLinks(string path)
    {
        for (var current = new DirectoryInfo(repository); current is not null; current = current.Parent)
            if (current.Attributes.HasFlag(FileAttributes.ReparsePoint))
                throw new IOException("Repository ancestors may not be reparse points.");
        string relative = Path.GetRelativePath(repository, path);
        if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar))
            throw new IOException("Development paths must remain within the repository.");
        string currentPath = repository;
        foreach (string part in relative.Split(Path.DirectorySeparatorChar))
        {
            currentPath = Path.Combine(currentPath, part);
            if (Path.Exists(currentPath) && File.GetAttributes(currentPath).HasFlag(FileAttributes.ReparsePoint))
                throw new IOException("Development paths may not be reparse points.");
        }
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
}
