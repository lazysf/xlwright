using System.Text;
using System.Text.Json;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

// L2, L3 and the record part of L5, ported from the Spike0 lifecycle checks. Excel calls are synthetic steps.
public sealed class LifecycleTests : IDisposable
{
    private static readonly TimeSpan Limit = TimeSpan.FromSeconds(120);
    private readonly string directory = Directory.CreateTempSubdirectory("xlwright-lifecycle-").FullName;
    private readonly string workbook;
    private readonly ProcessIdentity identity = new(12345, new DateTime(2026, 10, 30, 0, 0, 0, DateTimeKind.Utc), "EXCEL");

    public LifecycleTests() => workbook = Path.Combine(directory, "managed.xlsx");

    public void Dispose() => Directory.Delete(directory, recursive: true);

    public static TheoryData<string, string> RecordFailures()
    {
        var data = new TheoryData<string, string>();
        foreach (string kind in new[] { "starting", "owned", "completed" })
        foreach (string failure in new[] { "write", "partial-write", "set-length", "flush", "read-back", "mismatch", "corrupt" })
            data.Add(kind, failure);
        return data;
    }

    // A failed record update never permits the next Excel step. Before generation the session closes;
    // after generation (including the completed update of L5) it isolates and keeps the record and lock.
    [Theory]
    [MemberData(nameof(RecordFailures))]
    public void RecordFailureStopsTheNextStep(string kind, string failure)
    {
        string path = Path.Combine(directory, $"{kind}-{failure}.lock");
        var record = kind == "completed" ? new OwnershipRecord(1, kind) :
            new OwnershipRecord(1, kind, workbook, kind == "owned" ? identity : null);
        var previous = kind == "starting" ? new OwnershipRecord(1, "completed") :
            new OwnershipRecord(1, kind == "owned" ? "starting" : "owned", workbook, kind == "completed" ? identity : null);
        byte[] original = JsonSerializer.SerializeToUtf8Bytes(previous, OwnershipRecord.JsonOptions);
        byte[] partial = JsonSerializer.SerializeToUtf8Bytes(record, OwnershipRecord.JsonOptions)[..^2];
        using (var file = OwnershipRecordFile.Acquire(path))
        {
            file.Write(previous);
            var gate = new ExecutionGate();
            gate.Begin("record-" + kind, Limit);
            file.BeforeIo = (_, operation, stream) =>
            {
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
            var error = Assert.Throws<RecordIoException>(() => RecordTransition.Run(gate, file, record,
                () => gate.Step("next-excel-operation", () => nextRan = true)));
            EndSafety.Fault(gate, kind != "starting", false, error);

            string expected = failure is "partial-write" ? "write" : failure is "mismatch" or "corrupt" ? "read-back" : failure;
            Assert.Equal((kind, expected), (error.Boundary, error.Operation));
            Assert.False(nextRan);
            Assert.False(gate.Snapshot.Busy);
            Assert.Equal(kind == "starting" ? SessionState.Closed : SessionState.Isolated, gate.Snapshot.State);
            Assert.ThrowsAny<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
        }
        // Reopening proves lock release, not lifecycle completion.
        if (failure == "partial-write")
        {
            byte[] retained = File.ReadAllBytes(path);
            byte[] expected = new byte[Math.Max(original.Length, partial.Length)];
            original.CopyTo(expected, 0);
            partial.CopyTo(expected, 0);
            Assert.Equal(expected, retained);
            Assert.NotEqual(original, retained);
        }
    }

    public static TheoryData<string> EndConditions() => new(
        "safe", "unmanaged", "protected", "identity", "books-unknown", "protected-unknown", "late-book", "late-protected", "isolated");

    // L3: only a confirmed unmanaged workbook or Protected View keeps a ready session ready.
    // Unknown observations, identity changes and anything after Close isolate it.
    [Theory]
    [MemberData(nameof(EndConditions))]
    public void EndSafetyDecidesTheState(string condition)
    {
        string path = Path.Combine(directory, condition + ".lock");
        using (var file = OwnershipRecordFile.Acquire(path))
        {
            var owned = new OwnershipRecord(1, "owned", workbook, identity);
            file.Write(owned);
            var gate = new ExecutionGate();
            gate.Begin("prepare", Limit);
            gate.Complete(SessionState.Ready);
            if (condition == "isolated") gate.Fault(true);
            var calls = new List<string>();
            bool closeStarted = false;
            Exception? failure = null;
            try
            {
                gate.Begin("end", Limit);
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
                gate.Complete(SessionState.Closed);
            }
            catch (Exception error) when (error is CoreException or IOException or WorkStoppedException)
            {
                failure = error;
                EndSafety.Fault(gate, true, closeStarted, error);
            }

            var expectedState = condition == "safe" ? SessionState.Closed :
                condition is "unmanaged" or "protected" ? SessionState.Ready : SessionState.Isolated;
            Assert.Equal(expectedState, gate.Snapshot.State);
            Assert.False(gate.Snapshot.Busy);
            Assert.Equal(condition == "safe", calls.Contains("quit"));
            Assert.Equal(condition == "safe", calls.Contains("completed"));
            Assert.Equal(condition is "safe" or "late-book" or "late-protected", closeStarted);
            if (condition == "isolated") Assert.Empty(calls);
            if (condition is "unmanaged" or "protected" or "identity" or "late-book" or "late-protected")
                Assert.Equal(ErrorDetail.ExcelNeedsReview, Assert.IsAssignableFrom<CoreException>(failure).Detail);
            Assert.Equal(condition == "safe" ? new OwnershipRecord(1, "completed") : owned, file.Read());
            Assert.ThrowsAny<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
        }
    }

    // L5: only a clear save failure keeps ready; it never starts Close.
    [Fact]
    public void ClearSaveFailureKeepsReady()
    {
        var gate = new ExecutionGate();
        gate.Begin("prepare", Limit);
        gate.Complete(SessionState.Ready);
        gate.Begin("end", Limit);
        var error = Assert.Throws<SaveFailedException>(() =>
            gate.Step("save", () => throw new SaveFailedException(new IOException("Injected save failure."))));
        EndSafety.Fault(gate, true, false, error);

        Assert.Equal((ErrorCode.ExcelError, ErrorDetail.SaveFailed), (error.Code, error.Detail));
        Assert.Equal(SessionState.Ready, gate.Snapshot.State);
    }

    [Fact]
    public void UnknownSaveResultIsolates()
    {
        var gate = new ExecutionGate();
        gate.Begin("prepare", Limit);
        gate.Complete(SessionState.Ready);
        gate.Begin("end", Limit);
        var error = Assert.Throws<IOException>(() => gate.Step("save", () => throw new IOException("Disconnected.")));
        EndSafety.Fault(gate, true, false, error);

        Assert.Equal(SessionState.Isolated, gate.Snapshot.State);
    }
}
