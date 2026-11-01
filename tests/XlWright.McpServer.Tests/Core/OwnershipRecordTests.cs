using System.Text;
using System.Text.Json;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

// Record format, exclusive handle and the next-start rule (L4 and L8 without Excel).
public sealed class OwnershipRecordTests : IDisposable
{
    private readonly string directory = Directory.CreateTempSubdirectory("xlwright-record-").FullName;
    private readonly string workbook;
    private readonly ProcessIdentity owner = new(12345, new DateTime(2026, 10, 29, 0, 0, 0, DateTimeKind.Utc), "EXCEL");
    private readonly OwnershipRecord owned;
    private readonly string ownedJson;

    public OwnershipRecordTests()
    {
        workbook = Path.Combine(directory, "managed.xlsx");
        owned = new(1, "owned", workbook, owner);
        ownedJson = JsonSerializer.Serialize(owned, OwnershipRecord.JsonOptions);
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);

    private string LockPath(string name) => Path.Combine(directory, name + ".lock");

    [Fact]
    public void RecordRoundTripsWithUtcIdentity() => Assert.Equal(owned, OwnershipRecord.Parse(ownedJson));

    [Fact]
    public void InvalidRecordsFailClosed()
    {
        foreach (string text in new[]
        {
            "", "{", "[]", "{}", "{\"version\":2,\"kind\":\"completed\"}",
            "{\"version\":1,\"kind\":\"unknown\"}", "{\"version\":1,\"kind\":\"owned\"}",
            "{\"version\":1,\"kind\":\"starting\",\"workbook_path\":\"relative.xlsx\"}",
            "{\"version\":1,\"version\":1,\"kind\":\"completed\"}",
            "{\"version\":1,\"kind\":\"completed\",\"extra\":1}",
            ownedJson.Replace("12345", "0"), ownedJson.Replace(".0000000Z", ".0000000+09:00")
        })
        {
            var error = Assert.Throws<SessionRecordInvalidException>(() => OwnershipRecord.Parse(text));
            Assert.Equal((ErrorCode.Conflict, ErrorDetail.SessionRecordInvalid), (error.Code, error.Detail));
        }
    }

    [Fact]
    public void ExclusiveHandleTruncatesWithoutBom()
    {
        string path = LockPath("exclusive");
        using (var file = OwnershipRecordFile.Acquire(path))
        {
            Assert.True(file.CreatedNew);
            file.Write(owned);
            Assert.ThrowsAny<IOException>(() => { using var other = OwnershipRecordFile.Acquire(path); });
            file.Write(new(1, "completed"));
            Assert.Equal(new OwnershipRecord(1, "completed"), file.Read());
        }
        Assert.Equal("{\"version\":1,\"kind\":\"completed\"}", Encoding.UTF8.GetString(File.ReadAllBytes(path)));
        using var reopened = OwnershipRecordFile.Acquire(path);
        Assert.False(reopened.CreatedNew);
        Assert.Null(reopened.CheckPrevious());
    }

    [Fact]
    public void NewRecordFileAllowsStart()
    {
        using var file = OwnershipRecordFile.Acquire(LockPath("new"));
        Assert.Null(file.CheckPrevious(_ => throw new InvalidOperationException("The process must not be queried.")));
    }

    // L8: empty, corrupt, unknown version or kind is refused and the bytes are kept.
    [Theory]
    [InlineData("")]
    [InlineData("broken")]
    [InlineData("{\"version\":2,\"kind\":\"completed\"}")]
    [InlineData("{\"version\":1,\"kind\":\"other\"}")]
    public void InvalidPreviousRecordIsPreserved(string text)
    {
        string path = LockPath(Guid.NewGuid().ToString("N"));
        File.WriteAllText(path, text);
        using (var file = OwnershipRecordFile.Acquire(path))
            Assert.Throws<SessionRecordInvalidException>(() => file.CheckPrevious());
        Assert.Equal(text, File.ReadAllText(path));
    }

    // L4: a start that stopped after the starting mark may have created Excel.
    [Fact]
    public void StartingRecordNeedsReviewAndIsPreserved()
    {
        string path = LockPath("starting");
        string text = JsonSerializer.Serialize(new OwnershipRecord(1, "starting", workbook), OwnershipRecord.JsonOptions);
        File.WriteAllText(path, text);
        using (var file = OwnershipRecordFile.Acquire(path))
            AssertNeedsReview(() => file.CheckPrevious(_ => throw new InvalidOperationException("Not queried.")));
        Assert.Equal(text, File.ReadAllText(path));
    }

    [Fact]
    public void AbsentPidIsReplacedAndReturned()
    {
        string path = LockPath("absent");
        using (var created = OwnershipRecordFile.Acquire(path)) created.Write(owned);
        using var file = OwnershipRecordFile.Acquire(path);

        Assert.Equal(owned, file.CheckPrevious(_ => null));
        Assert.Equal(new OwnershipRecord(1, "completed"), file.Read());
        Assert.Null(file.CheckPrevious());
    }

    [Fact]
    public void RealProcessQueryTreatsAMissingPidAsAbsent()
    {
        string path = LockPath("missing-pid");
        var missing = owned with { Process = owner with { Pid = int.MaxValue } };
        using (var created = OwnershipRecordFile.Acquire(path)) created.Write(missing);
        using var file = OwnershipRecordFile.Acquire(path);

        Assert.Equal(missing, file.CheckPrevious());
    }

    // L8: a live process, a reused PID and a failed query all keep the record.
    [Fact]
    public void LiveProcessNeedsReviewWithoutAlteringRecord()
    {
        using var current = System.Diagnostics.Process.GetCurrentProcess();
        var live = owned with { Process = ProcessIdentity.Read(current) };
        AssertRefusedAndPreserved(live, null, ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview);
    }

    [Fact]
    public void ReusedPidNeedsReviewWithoutAlteringRecord() =>
        AssertRefusedAndPreserved(owned, _ => owner with { StartedAtUtc = owner.StartedAtUtc.AddSeconds(1) },
            ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview);

    [Fact]
    public void DifferentProcessNameNeedsReview() =>
        AssertRefusedAndPreserved(owned, _ => owner with { Name = "different" }, ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview);

    [Theory]
    [InlineData("exited")]
    [InlineData("denied")]
    public void ProcessQueryFailureIsNotAbsence(string failure) =>
        AssertRefusedAndPreserved(owned, _ => throw (failure == "exited" ?
                new InvalidOperationException("Process exited.") : new System.ComponentModel.Win32Exception(5)),
            ErrorCode.InternalError, ErrorDetail.ProcessQueryFailed);

    private void AssertRefusedAndPreserved(OwnershipRecord record, Func<int, ProcessIdentity?>? findProcess,
        ErrorCode code, string detail)
    {
        string path = LockPath(Guid.NewGuid().ToString("N"));
        using (var created = OwnershipRecordFile.Acquire(path)) created.Write(record);
        byte[] original = File.ReadAllBytes(path);
        using (var file = OwnershipRecordFile.Acquire(path))
        {
            var error = Assert.ThrowsAny<CoreException>(() => file.CheckPrevious(findProcess));
            Assert.Equal((code, detail), (error.Code, error.Detail));
        }
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    private static void AssertNeedsReview(Action action)
    {
        var error = Assert.ThrowsAny<CoreException>(action);
        Assert.Equal((ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview), (error.Code, error.Detail));
    }

    [Fact]
    public void PreviousReadFailureIsPreserved()
    {
        string path = LockPath("previous-read");
        using (var created = OwnershipRecordFile.Acquire(path)) created.Write(new(1, "completed"));
        byte[] original = File.ReadAllBytes(path);
        using (var file = OwnershipRecordFile.Acquire(path))
        {
            file.BeforeIo = (_, _, _) => throw new IOException("Injected previous read.");
            var error = Assert.Throws<RecordIoException>(() => file.CheckPrevious());
            Assert.Equal(("existing", "read"), (error.Boundary, error.Operation));
            Assert.Equal((ErrorCode.InternalError, ErrorDetail.SessionRecordIoError), (error.Code, error.Detail));
            Assert.ThrowsAny<IOException>(() => { using var competing = OwnershipRecordFile.Acquire(path); });
        }
        Assert.Equal(original, File.ReadAllBytes(path));
    }
}
