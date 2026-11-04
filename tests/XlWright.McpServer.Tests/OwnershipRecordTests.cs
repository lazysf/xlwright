using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/architecture.md "所有記録の形式".
public sealed class OwnershipRecordTests
{
    private static readonly DateTime Marked = new(2026, 11, 2, 1, 23, 45, 678, DateTimeKind.Utc);
    private static readonly DateTime Started = new(2026, 11, 2, 1, 23, 46, 123, DateTimeKind.Utc);

    [Fact]
    public void WritesAnOwnedRecordInTheDesignedShape()
    {
        var record = new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", new RecordedProcess(12345, Started, "EXCEL"));

        Assert.Equal(
            """{"version":1,"state":"owned","marked_at_utc":"2026-11-02T01:23:45.6780000Z","workbook":"Sales.xlsm","process":{"pid":12345,"started_at_utc":"2026-11-02T01:23:46.1230000Z","name":"EXCEL"}}""",
            record.ToJson());
    }

    [Theory]
    [InlineData("starting")]
    [InlineData("completed")]
    public void ReadsBackARecordWithoutAProcess(string state)
    {
        var record = new OwnershipRecord(state, Marked, "販売.xlsm", null);

        Assert.Equal(record, OwnershipRecord.Parse(record.ToJson()));
    }

    [Fact]
    public void ReadsBackAnOwnedRecord()
    {
        var record = new OwnershipRecord(OwnershipRecord.Owned, Marked, "Sales.xlsm", new RecordedProcess(12345, Started, "EXCEL"));

        var read = OwnershipRecord.Parse(record.ToJson())!;

        Assert.Equal(record.Process, read.Process);
        Assert.Equal(DateTimeKind.Utc, read.MarkedAtUtc.Kind);
        Assert.Equal(record with { Process = null }, read with { Process = null });
    }

    private const string Head = "\"version\":1,\"marked_at_utc\":\"2026-11-02T01:23:45.6789012Z\",\"workbook\":\"Sales.xlsm\"";
    private const string Process = "\"process\":{\"pid\":12345,\"started_at_utc\":\"2026-11-02T01:23:46.1234567Z\",\"name\":\"EXCEL\"}";

    public static TheoryData<string> Invalid => new()
    {
        "",
        "{}",
        "[]",
        "{" + Head + ",\"state\":\"completed\"} trailing",
        "{" + Head + ",\"state\":\"completed\",}",
        "{" + Head + ",\"state\":\"completed\" /* note */}",
        "{" + Head + ",\"state\":\"completed\",\"state\":\"completed\"}",
        "{" + Head + ",\"state\":\"completed\",\"note\":1}",
        "{" + Head.Replace("\"version\":1", "\"version\":2") + ",\"state\":\"completed\"}",
        "{" + Head + ",\"state\":\"done\"}",
        "{" + Head + ",\"state\":1}",
        "{" + Head + ",\"state\":null}",
        "{" + Head + ",\"state\":\"owned\"}",
        "{" + Head + ",\"state\":\"completed\"," + Process + "}",
        "{" + Head.Replace("45.6789012Z", "45.6789012+09:00") + ",\"state\":\"completed\"}",
        "{" + Head.Replace("2026-11-02T01:23:45.6789012Z", "2026/11/02") + ",\"state\":\"completed\"}",
        "{" + Head + ",\"state\":\"owned\"," + Process.Replace("12345", "\"12345\"") + "}",
        "{" + Head + ",\"state\":\"owned\"," + Process.Replace(",\"name\":\"EXCEL\"", "") + "}",
        "{" + Head + ",\"state\":\"owned\"," + Process.Replace("\"name\":\"EXCEL\"", "\"name\":\"EXCEL\",\"pid\":1") + "}",
        "{" + Head.Replace(",\"workbook\":\"Sales.xlsm\"", "") + ",\"state\":\"completed\"}",
    };

    [Theory]
    [MemberData(nameof(Invalid))]
    public void RejectsAnythingButTheStrictShape(string text)
    {
        Assert.Null(OwnershipRecord.Parse(text));
    }
}
