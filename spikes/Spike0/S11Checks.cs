using System.Diagnostics;
using System.Windows.Forms;

namespace Spike0;

internal static class S11Checks
{
    internal static void Run(string directory, Action<string, Action> check)
    {
        check("S11 fixed arguments reject arbitrary paths", () =>
        {
            Require(S11Probe.Parse(["--excel-s11r2-baseline"]) == "baseline");
            Require(S11Probe.Parse(["--excel-s11r2-modal", "other.xlsm"]) is null);
            Require(S11Probe.Parse(["--excel-s11r2-unknown"]) is null);
            // The first attempt's stamps stay in place, so its arguments must not reach them again.
            Require(S11Probe.Parse(["--excel-s11-baseline"]) is null);
            Require(S11Probe.Parse(["--excel-s11r2-shell"]) is null);
            Require(S11Probe.Parse(["--excel-s11r2-shell2"]) == "shell2");
        });
        check("S11 evidence failure revokes work before reporting the failure", () =>
        {
            using var process = Process.GetCurrentProcess();
            var evidence = new S11Evidence(Path.Combine(directory, "s11-disposed.jsonl"));
            using var probe = new S11Probe("baseline", directory, evidence);
            probe.Identify(ProcessIdentity.Read(process), 0);
            evidence.Dispose();
            var gate = new ExecutionGate();
            gate.Begin("active-call", TimeSpan.FromSeconds(30));
            try { probe.Observe(gate); }
            catch (ObjectDisposedException) { }
            Require(gate.Snapshot is { State: "isolated", Busy: true, StopReason: "diagnostic-failure" });
            bool next = false;
            try { gate.Step("next", () => next = true); }
            catch (WorkStoppedException) { }
            Require(!next);
        });
        check("S11 independent evidence survives a blocked output operation", () =>
        {
            string path = Path.Combine(directory, "s11-independent.jsonl");
            using var evidence = new S11Evidence(path);
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            using var writer = new BlockedWriter(entered, release);
            using var log = new DiagnosticLog(Path.Combine(directory, "s11-blocked-log.jsonl"), evidence, writer);
            Task output = Task.Run(() => log.Write("blocked-output", new ExecutionGate()));
            try
            {
                Require(entered.Wait(TimeSpan.FromSeconds(2)));
                evidence.Write("monitor", "observed");
                Require(ReadEvidence(path).Contains("monitor"));
                Require(!output.IsCompleted);
            }
            finally { release.Set(); output.GetAwaiter().GetResult(); }
        });
        check("S11 records failed phase without turning it into success", () =>
        {
            string path = Path.Combine(directory, "s11-failure.jsonl");
            using var evidence = new S11Evidence(path);
            try { evidence.Measure("failure", () => throw new IOException()); }
            catch (IOException) { }
            string text = ReadEvidence(path);
            Require(text.Contains("failed") && !text.Contains("returned"));
        });
        check("S11 native metadata uses only the identified process", () =>
        {
            Require(StaRunner.Start(() =>
            {
                using var process = Process.GetCurrentProcess();
                var identity = ProcessIdentity.Read(process);
                using var owner = new Form { Text = "S11 text must not be collected" };
                using var child = new Form { Text = "S11 private text" };
                owner.Show();
                child.Show(owner);
                var snapshot = S11Windows.Capture(identity, owner.Handle);
                Require(snapshot.Response == "responded");
                Require(snapshot.Windows.Any(window => window.Hwnd == child.Handle && window.Owner == owner.Handle));
                string json = System.Text.Json.JsonSerializer.Serialize(snapshot);
                Require(!json.Contains("private text") && !json.Contains("must not be collected"));
                try
                {
                    S11Windows.Capture(identity with { StartedAtUtc = identity.StartedAtUtc.AddSeconds(-1) }, owner.Handle);
                    throw new InvalidOperationException("Changed identity was accepted.");
                }
                catch (InvalidOperationException error) when (error.Message == "S11 process identity changed.") { }
                return true;
            }).WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult());
        });
        check("S11 cross-thread unresponsive window returns within its observation bound", () =>
        {
            using var process = Process.GetCurrentProcess();
            var identity = ProcessIdentity.Read(process);
            using var entered = new ManualResetEventSlim();
            using var release = new ManualResetEventSlim();
            nint window = 0;
            Task<bool> owner = StaRunner.Start(() =>
            {
                using var form = new Form();
                window = form.Handle;
                entered.Set();
                // WaitHandle waits on STA can dispatch sent messages. Sleep deliberately does not pump them.
                Thread.Sleep(1500);
                return true;
            });
            try
            {
                Require(entered.Wait(TimeSpan.FromSeconds(2)));
                long started = Stopwatch.GetTimestamp();
                var snapshot = S11Windows.Capture(identity, window);
                if (snapshot.Response != "failed-or-timed-out")
                    throw new InvalidOperationException("S11 cross-thread response: " + snapshot.Response);
                Require(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(2));
            }
            finally
            {
                release.Set();
                Require(owner.WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult());
            }
            Require(S11Windows.Capture(identity, window).Response == "window-unavailable");
        });
    }

    private static void Require(bool condition)
    {
        if (!condition) throw new InvalidOperationException("S11 check failed.");
    }

    private static string ReadEvidence(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private sealed class BlockedWriter(ManualResetEventSlim entered, ManualResetEventSlim release) : TextWriter
    {
        public override System.Text.Encoding Encoding => System.Text.Encoding.UTF8;
        public override void WriteLine(string? value)
        {
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(5))) throw new TimeoutException();
        }
    }
}
