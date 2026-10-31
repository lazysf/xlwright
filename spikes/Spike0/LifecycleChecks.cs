using System.Text;
using System.Text.Json;

namespace Spike0;

internal static class LifecycleChecks
{
    internal static void Run(string directory, Action<string, Action> check)
    {
        string workbook = Path.Combine(directory, "managed.xlsx");
        var identity = new ProcessIdentity(12345, new DateTime(2026, 10, 30, 0, 0, 0, DateTimeKind.Utc), "EXCEL");
        foreach (string kind in new[] { "starting", "owned", "completed" })
        foreach (string failure in new[] { "write", "partial-write", "set-length", "flush", "read-back", "mismatch", "corrupt" })
        {
            check($"L2 {kind}/{failure}", () =>
            {
                string path = Path.Combine(directory, $"{kind}-{failure}.lock");
                var record = kind == "completed" ? new OwnershipRecord(1, kind) :
                    new OwnershipRecord(1, kind, workbook, kind == "owned" ? identity : null);
                var previous = kind == "starting" ? new OwnershipRecord(1, "completed") :
                    new OwnershipRecord(1, kind == "owned" ? "starting" : "owned", workbook,
                        kind == "completed" ? identity : null);
                byte[] original = JsonSerializer.SerializeToUtf8Bytes(previous, OwnershipRecord.JsonOptions);
                byte[] partial = JsonSerializer.SerializeToUtf8Bytes(record, OwnershipRecord.JsonOptions)[..^2];
                using (var file = OwnershipRecordFile.Acquire(path))
                {
                    file.Write(previous);
                    var gate = new ExecutionGate();
                    gate.Begin("record-" + kind, TimeSpan.FromSeconds(120));
                    var calls = new List<string>();
                    file.BeforeIo = (_, operation, stream) =>
                    {
                        calls.Add(operation);
                        if (failure is "mismatch" or "corrupt" && operation == "read-back")
                        {
                            byte[] other = failure == "corrupt" ? Encoding.UTF8.GetBytes("{") : JsonSerializer.SerializeToUtf8Bytes(
                                new OwnershipRecord(1, "starting", workbook + ".other"), OwnershipRecord.JsonOptions);
                            stream.Position = 0;
                            stream.Write(other);
                            stream.SetLength(other.Length);
                            stream.Flush(true);
                            stream.Position = 0;
                            return;
                        }
                        if (failure == "partial-write" && operation == "write")
                        {
                            stream.Write(partial);
                            throw new IOException("Injected partial write.");
                        }
                        if (operation == failure) throw new IOException("Injected " + operation);
                    };
                    bool nextRan = false;
                    try
                    {
                        RecordTransition.Run(gate, file, record,
                            () => gate.Step("next-excel-operation", () => nextRan = true));
                        throw new InvalidOperationException("Expected record failure.");
                    }
                    catch (RecordIoException error)
                    {
                        EndSafety.Fault(gate, kind != "starting", false, error);
                        string expected = failure is "partial-write" ? "write" : failure is "mismatch" or "corrupt" ? "read-back" : failure;
                        Require(error.Boundary == kind && error.Operation == expected);
                        File.WriteAllText(path + ".failure.json", JsonSerializer.Serialize(new
                        {
                            boundary = error.Boundary, operation = error.Operation, snapshot = gate.Snapshot, calls
                        }));
                    }
                    Require(!nextRan && !gate.Snapshot.Busy && gate.Snapshot.State ==
                        (kind == "starting" ? "closed" : "isolated"));
                    Expect<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
                }
                // Reopening proves lock release, not lifecycle completion. Keep the failed bytes.
                using var reopened = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                byte[] retained = new byte[checked((int)reopened.Length)];
                reopened.ReadExactly(retained);
                File.WriteAllBytes(path + ".retained", retained);
                if (failure == "partial-write")
                {
                    byte[] expected = new byte[Math.Max(original.Length, partial.Length)];
                    original.CopyTo(expected, 0);
                    partial.CopyTo(expected, 0);
                    Require(retained.SequenceEqual(expected) && !retained.SequenceEqual(original));
                }
            });
        }

        check("L2 previous read failure is preserved", () =>
        {
            string path = Path.Combine(directory, "previous-read.lock");
            using (var created = OwnershipRecordFile.Acquire(path)) created.Write(new(1, "completed"));
            byte[] original = File.ReadAllBytes(path);
            using (var file = OwnershipRecordFile.Acquire(path))
            {
                file.BeforeIo = (_, _, _) => throw new IOException("Injected previous read.");
                try { file.CheckPrevious(); throw new InvalidOperationException("Expected read failure."); }
                catch (RecordIoException error) { Require(error.Boundary == "existing" && error.Operation == "read"); }
                Expect<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
            }
            Require(File.ReadAllBytes(path).SequenceEqual(original));
        });

        check("next start replaces an owned record with an absent PID and returns it", () =>
        {
            string path = Path.Combine(directory, "absent-pid.lock");
            var owned = new OwnershipRecord(1, "owned", workbook, identity with { Pid = int.MaxValue });
            using (var created = OwnershipRecordFile.Acquire(path)) created.Write(owned);
            using var file = OwnershipRecordFile.Acquire(path);
            Require(file.CheckPrevious() == owned);
            Require(file.Read() == new OwnershipRecord(1, "completed"));
            Require(file.CheckPrevious() is null);
        });

        foreach (string target in new[] { "starting", "owned", "completed" })
        {
            check("L2 fixed injection targets only " + target, () =>
            {
                using var file = OwnershipRecordFile.Acquire(Path.Combine(directory, "fixed-" + target + ".lock"));
                var previous = new OwnershipRecord(1, "completed");
                file.Write(previous);
                file.InjectWriteFailure(target);
                Require(file.Read() == previous);
                int failures = 0;
                foreach (var record in new OwnershipRecord[]
                    { new(1, "starting", workbook), new(1, "owned", workbook, identity), new(1, "completed") })
                {
                    try { file.Write(record); }
                    catch (RecordIoException error)
                    {
                        Require(record.Kind == target && error.Boundary == target && error.Operation == "write");
                        Require(file.Read() == previous);
                        failures++;
                        continue;
                    }
                    Require(record.Kind != target && file.Read() == record);
                    previous = record;
                }
                Require(failures == 1);
            });
        }

        foreach (string condition in new[] { "safe", "unmanaged", "protected", "identity", "books-unknown",
            "protected-unknown", "late-book", "late-protected", "isolated" })
        {
            check("L3 " + condition, () =>
            {
                string path = Path.Combine(directory, condition + ".lock");
                using (var file = OwnershipRecordFile.Acquire(path))
                {
                    var owned = new OwnershipRecord(1, "owned", workbook, identity);
                    file.Write(owned);
                    var gate = new ExecutionGate();
                    gate.Begin("prepare", TimeSpan.FromSeconds(120));
                    gate.Complete("ready");
                    if (condition == "isolated") gate.Fault(true);
                    var calls = new List<string>();
                    bool closeStarted = false;
                    try
                    {
                        gate.Begin("end", TimeSpan.FromSeconds(120));
                        EndSafety.CloseAndQuit(afterClose =>
                        {
                            gate.Step("observe", () =>
                            {
                                calls.Add(afterClose ? "verify-after-close" : "verify");
                                EndSafety.RequireIdentity(condition != "identity");
                                if (condition is "books-unknown" or "protected-unknown")
                                    throw new IOException("Injected observation failure.");
                                string[] paths = condition == "unmanaged" || afterClose && condition == "late-book" ?
                                    [workbook + ".other"] : afterClose ? [] : [workbook];
                                int protectedCount = condition == "protected" || afterClose && condition == "late-protected" ? 1 : 0;
                                EndSafety.RequireBooks(workbook, paths, protectedCount, paths.Length, afterClose);
                            });
                        }, () => gate.Step("close", () => { closeStarted = true; calls.Add("close"); }),
                        () => gate.Step("quit", () => calls.Add("quit")));
                        gate.Step("completed", () => { file.Write(new(1, "completed")); calls.Add("completed"); });
                        gate.Complete("closed");
                    }
                    catch (Exception error) when (error is EndRefusedException or IOException or InvalidOperationException or WorkStoppedException)
                    {
                        EndSafety.Fault(gate, true, closeStarted, error);
                    }
                    string expectedState = condition == "safe" ? "closed" :
                        condition is "unmanaged" or "protected" ? "ready" : "isolated";
                    Require(gate.Snapshot.State == expectedState && !gate.Snapshot.Busy);
                    Require(calls.Contains("quit") == (condition == "safe"));
                    Require(calls.Contains("completed") == (condition == "safe"));
                    Require(closeStarted == (condition is "safe" or "late-book" or "late-protected"));
                    if (condition == "isolated") Require(calls.Count == 0);
                    Require(file.Read() == (condition == "safe" ? new OwnershipRecord(1, "completed") : owned));
                    Expect<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
                    File.WriteAllText(path + ".observation.json", JsonSerializer.Serialize(new { condition, calls, snapshot = gate.Snapshot }));
                }
                using var reopened = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            });
        }
    }

    private static void Require(bool condition)
    {
        if (!condition) throw new InvalidOperationException("Lifecycle check failed.");
    }

    private static void Expect<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException("Expected " + typeof(T).Name);
    }
}
