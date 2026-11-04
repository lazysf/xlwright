using System.Diagnostics;
using XlWright.Core;
using XlWright.Core.Com;
using Xunit;

namespace XlWright.McpServer.Tests;

// The identity check of a recorded process (docs/architecture.md "完了にする条件"), against this test process.
public sealed class WindowsProcessQueryTests
{
    private static readonly RecordedProcess Self = SelfRecord();

    private readonly WindowsProcessQuery _query = new();

    [Fact]
    public void FindsTheRecordedProcessThroughAHandle()
    {
        var lookup = _query.Find(Self);

        Assert.Equal(ProcessPresence.Running, lookup.Presence);
        using var process = lookup.Process!;
        Assert.Equal(Environment.ProcessId, process.Facts.Pid);
        Assert.False(process.WaitForExit(TimeSpan.Zero));
    }

    [Fact]
    public void TakesAnotherCreationTimeForAReusedPid()
    {
        var lookup = _query.Find(Self with { StartedAtUtc = Self.StartedAtUtc.AddTicks(1) });

        Assert.Equal(ProcessPresence.Gone, lookup.Presence);
        Assert.Null(lookup.Process);
    }

    [Fact]
    public void DoesNotTakeAnotherNameForTheRecordedProcess()
    {
        var lookup = _query.Find(Self with { Name = "EXCEL" });

        Assert.Equal(ProcessPresence.Unknown, lookup.Presence);
        Assert.Null(lookup.Process);
    }

    [Fact]
    public void ReportsAPidWithoutAProcessAsGone()
    {
        // PIDs are multiples of 4, so this one never names a process.
        var lookup = _query.Find(Self with { Pid = 0x7FFFFFFF });

        Assert.Equal(ProcessPresence.Gone, lookup.Presence);
    }

    private static RecordedProcess SelfRecord()
    {
        using var self = Process.GetCurrentProcess();
        return new RecordedProcess(self.Id, self.StartTime.ToUniversalTime(), Path.GetFileNameWithoutExtension(Environment.ProcessPath!));
    }
}
