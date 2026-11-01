using System.Text.Json;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

public sealed class ExcelSessionTests : IDisposable
{
    private static readonly TimeSpan Generous = TimeSpan.FromSeconds(30);
    private readonly string project = Directory.CreateTempSubdirectory("xlwright-session-").FullName;
    private readonly FakeExcelDriver driver = new();
    private readonly ExcelSession session;

    public ExcelSessionTests()
    {
        File.WriteAllText(Path.Combine(project, "project.json"), """{"paths":{"workbook":"book.xlsx"}}""");
        File.WriteAllText(Path.Combine(project, "book.xlsx"), "");
        Directory.CreateDirectory(Path.Combine(project, "src"));
        Directory.CreateDirectory(Path.Combine(project, "tests"));
        session = Create();
    }

    private string RecordPath => Path.Combine(project, ".xlwright.lock");

    private ExcelSession Create(bool excelAllowed = true, TimeSpan? cellTimeout = null, TimeSpan? lifecycleTimeout = null) =>
        new(_ => driver, excelAllowed, cellTimeout ?? Generous, lifecycleTimeout ?? Generous);

    public void Dispose()
    {
        session.Dispose();
        GC.SuppressFinalize(this);
        // The record of an isolated session stays held by this test process until its handle is finalized.
        try { Directory.Delete(project, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static async Task<ErrorInfo> Refused(Func<Task> operation) =>
        (await Assert.ThrowsAsync<SessionOperationException>(operation)).Error;

    private static async Task WaitUntil(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 500 && !condition(); attempt++) await Task.Delay(10);
        Assert.True(condition());
    }

    [Fact]
    public async Task VerticalSliceFollowsTheLifecycleOrder()
    {
        var started = await session.StartAsync(project);
        Assert.Equal(new SessionSummary("ready", false), started.Session);
        Assert.Equal(Path.Combine(project, "book.xlsx"), started.Workbook);

        var written = await session.SetRangeAsync("Lifecycle", "$a$2", [[1d]], null);
        Assert.Equal("A2", written.Address);
        var read = await session.GetRangeAsync("Lifecycle", "A2");
        Assert.Equal(1d, read.Values![0][0]);
        Assert.Null(read.Formulas);
        await session.SaveAsync();
        var ended = await session.EndAsync();

        Assert.Equal(new SessionSummary("closed", false), ended.Session);
        Assert.Equal(["Generate", "Identify", "Open", "Write", "Read", "Save", "Save", "Verify", "Close", "VerifyEmpty", "Quit", "WaitForExit"], driver.Calls);
        Assert.Equal("completed", JsonDocument.Parse(File.ReadAllText(RecordPath)).RootElement.GetProperty("kind").GetString());
    }

    [Fact]
    public async Task AnEndedSessionCanStartAgain()
    {
        await session.StartAsync(project);
        await session.EndAsync();

        var again = await session.StartAsync(project);

        Assert.Equal("ready", again.Session.State);
    }

    [Fact]
    public async Task OperationsNeedTheRightState()
    {
        Assert.Equal("SESSION_NOT_STARTED", (await Refused(() => session.GetRangeAsync("S", "A1"))).Detail);
        Assert.Equal("SESSION_NOT_STARTED", (await Refused(session.SaveAsync)).Detail);

        await session.StartAsync(project);

        var refused = await Refused(() => session.StartAsync(project));
        Assert.Equal("CONFLICT", refused.Code);
        Assert.Equal("SESSION_ACTIVE", refused.Detail);
        Assert.Null(refused.Stage);
    }

    [Theory]
    [InlineData("A1:B2", "NOT_IMPLEMENTED")]
    [InlineData("A1,B2", "INVALID_ADDRESS")]
    [InlineData("A:A", "INVALID_ADDRESS")]
    public async Task AddressesAreCheckedBeforeTheState(string address, string detail)
    {
        var refused = await Refused(() => session.GetRangeAsync("S", address));

        Assert.Equal("INVALID_REQUEST", refused.Code);
        Assert.Equal(detail, refused.Detail);
    }

    [Fact]
    public async Task UnsupportedRequestsAreRefusedWithoutSideEffects()
    {
        Assert.Equal("NOT_IMPLEMENTED", (await Refused(() => session.StartAsync(project, force: true))).Detail);
        Assert.Equal("NOT_IMPLEMENTED", (await Refused(() => session.EndAsync(discard: true))).Detail);
        Assert.Equal("NOT_IMPLEMENTED", (await Refused(() => session.GetRangeAsync("S", "A1", ["merged_ranges"]))).Detail);
        Assert.Equal("INVALID_ARGUMENT", (await Refused(() => session.StartAsync("relative"))).Detail);

        Assert.Empty(driver.Calls);
        Assert.False(File.Exists(RecordPath));
    }

    [Fact]
    public async Task DevelopmentModeStopsAfterTheConfiguration()
    {
        using var development = Create(excelAllowed: false);

        var refused = await Refused(() => development.StartAsync(project));

        Assert.Equal("NOT_IMPLEMENTED", refused.Detail);
        Assert.Equal("closed", development.Summary.State);
        Assert.Empty(driver.Calls);
        Assert.False(File.Exists(RecordPath));
        Assert.False(Directory.Exists(Path.Combine(project, ".xlwright")));
    }

    [Fact]
    public async Task DevelopmentModeStillReportsConfigurationFailures()
    {
        using var development = Create(excelAllowed: false);
        File.Delete(Path.Combine(project, "project.json"));

        Assert.Equal("CONFIG_NOT_FOUND", (await Refused(() => development.StartAsync(project))).Detail);
    }

    [Fact]
    public async Task BusyIsCheckedAfterArgumentsAndBeforeState()
    {
        await session.StartAsync(project);
        using var release = new ManualResetEventSlim();
        driver.Before = name => { if (name == "Read") release.Wait(); };
        Task<RangeResult> running = session.GetRangeAsync("S", "A1");
        await WaitUntil(() => driver.Calls.Contains("Read"));

        Assert.True(session.Summary.Busy);
        Assert.Equal("OPERATION_BUSY", (await Refused(() => session.SetRangeAsync("S", "A1", [[1d]], null))).Detail);
        Assert.Equal("INVALID_ADDRESS", (await Refused(() => session.SetRangeAsync("S", "?", [[1d]], null))).Detail);

        release.Set();
        await running;
        Assert.False(session.Summary.Busy);
    }

    [Fact]
    public async Task ADeadlineIsolatesAndTheLateResultChangesNothing()
    {
        using var timed = Create(cellTimeout: TimeSpan.FromMilliseconds(200));
        await timed.StartAsync(project);
        using var release = new ManualResetEventSlim();
        driver.Before = name => { if (name == "Read") release.Wait(); };

        // A bound here turns a missing response into a failure instead of a hung check.
        var failure = await Assert.ThrowsAsync<SessionOperationException>(() =>
            timed.GetRangeAsync("S", "A1").WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken));

        Assert.Equal("STATE_UNKNOWN", failure.Error.Code);
        Assert.Equal("OPERATION_TIMEOUT", failure.Error.Detail);
        Assert.Equal("read", failure.Error.Stage);
        Assert.Equal(true, failure.Error.ExcelMayRemain);
        Assert.NotNull(failure.Error.Process);
        Assert.NotNull(failure.Error.Log);
        Assert.Equal(new SessionSummary("isolated", true), failure.Session);

        release.Set();
        await WaitUntil(() => !timed.Summary.Busy);

        Assert.Equal("isolated", timed.Summary.State);
        Assert.Equal("SESSION_ISOLATED", (await Refused(timed.SaveAsync)).Detail);
    }

    [Fact]
    public async Task AClearSaveFailureKeepsTheSessionReady()
    {
        await session.StartAsync(project);
        driver.Before = name => { if (name == "Save") throw new SaveFailedException(); };

        var refused = await Refused(() => session.EndAsync());

        Assert.Equal("SAVE_FAILED", refused.Detail);
        Assert.Equal("save", refused.Stage);
        Assert.Equal("ready", session.Summary.State);
        Assert.DoesNotContain("Close", driver.Calls);
    }

    [Fact]
    public async Task AnUnmanagedWorkbookRefusesTheEndAndKeepsReady()
    {
        await session.StartAsync(project);
        driver.Before = name => { if (name == "Verify") throw new EndRefusedException(); };

        var refused = await Refused(() => session.EndAsync());

        Assert.Equal("EXCEL_NEEDS_REVIEW", refused.Detail);
        Assert.Equal("safety_check", refused.Stage);
        Assert.Equal("ready", session.Summary.State);
    }

    [Fact]
    public async Task AFailureAfterCloseStartedIsolates()
    {
        await session.StartAsync(project);
        driver.Before = name => { if (name == "Quit") throw new InvalidOperationException("quit failed"); };

        var refused = await Refused(() => session.EndAsync());

        Assert.Equal("INTERNAL_FAILURE", refused.Detail);
        Assert.Equal("isolated", session.Summary.State);
        Assert.NotNull(refused.Log);
        Assert.Contains("quit failed", File.ReadAllText(refused.Log));
    }

    [Fact]
    public async Task AMissingSheetKeepsReady()
    {
        await session.StartAsync(project);
        driver.Before = name =>
        {
            if (name == "Read") throw new CoreException(ErrorCode.NotFound, ErrorDetail.SheetNotFound, "missing");
        };

        Assert.Equal("SHEET_NOT_FOUND", (await Refused(() => session.GetRangeAsync("S", "A1"))).Detail);
        Assert.Equal("ready", session.Summary.State);
    }

    [Theory]
    [InlineData("Generate", "EXCEL_START_FAILED")]
    [InlineData("Identify", "EXCEL_START_FAILED")]
    [InlineData("Open", "WORKBOOK_OPEN_FAILED")]
    public async Task StartFailuresAfterGenerationIsolateAndKeepTheLock(string call, string detail)
    {
        driver.Before = name => { if (name == call) throw new InvalidOperationException("failed"); };

        var refused = await Refused(() => session.StartAsync(project));

        Assert.Equal("EXCEL_ERROR", refused.Code);
        Assert.Equal(detail, refused.Detail);
        Assert.Equal(true, refused.ExcelMayRemain);
        Assert.Equal("isolated", session.Summary.State);
        using var second = Create();
        Assert.Equal("PROJECT_LOCKED", (await Refused(() => second.StartAsync(project))).Detail);
    }

    [Fact]
    public async Task ASecondServerIsRefusedByTheLock()
    {
        await session.StartAsync(project);
        using var second = Create();

        var refused = await Refused(() => second.StartAsync(project));

        Assert.Equal("CONFLICT", refused.Code);
        Assert.Equal("PROJECT_LOCKED", refused.Detail);
        Assert.Equal("closed", second.Summary.State);
    }

    public static TheoryData<string, string> RefusedRecords => new()
    {
        { "{", "SESSION_RECORD_INVALID" },
        { "{\"version\":1,\"kind\":\"starting\",\"workbook_path\":\"C:\\\\book.xlsx\"}", "EXCEL_NEEDS_REVIEW" }
    };

    [Theory]
    [MemberData(nameof(RefusedRecords))]
    public async Task PreviousRecordsThatNeedReviewRefuseAndStayUnchanged(string text, string detail)
    {
        File.WriteAllText(RecordPath, text);

        var refused = await Refused(() => session.StartAsync(project));

        Assert.Equal(detail, refused.Detail);
        Assert.Equal("closed", session.Summary.State);
        Assert.Empty(driver.Calls);
        Assert.Equal(text, File.ReadAllText(RecordPath));
    }

    [Fact]
    public async Task ALockedWorkbookIsRefusedAndReleasesTheProject()
    {
        using (new FileStream(Path.Combine(project, "book.xlsx"), FileMode.Open, FileAccess.Read, FileShare.None))
            Assert.Equal("WORKBOOK_LOCKED", (await Refused(() => session.StartAsync(project))).Detail);

        Assert.Equal("ready", (await session.StartAsync(project)).Session.State);
    }

    [Fact]
    public async Task AReadOnlyWorkbookIsRefused()
    {
        string book = Path.Combine(project, "book.xlsx");
        File.SetAttributes(book, FileAttributes.ReadOnly);
        try { Assert.Equal("WORKBOOK_READ_ONLY", (await Refused(() => session.StartAsync(project))).Detail); }
        finally { File.SetAttributes(book, FileAttributes.Normal); }
    }

    [Fact]
    public async Task ShutdownClosesWithoutSaving()
    {
        await session.StartAsync(project);

        await session.ShutdownAsync();

        Assert.Equal("closed", session.Summary.State);
        Assert.DoesNotContain("Save", driver.Calls);
        Assert.Contains("Quit", driver.Calls);
        Assert.Equal("completed", JsonDocument.Parse(File.ReadAllText(RecordPath)).RootElement.GetProperty("kind").GetString());
    }

    [Fact]
    public async Task ShutdownAddsNoWorkWhileACallIsPending()
    {
        await session.StartAsync(project);
        using var release = new ManualResetEventSlim();
        driver.Before = name => { if (name == "Read") release.Wait(); };
        Task<RangeResult> running = session.GetRangeAsync("S", "A1");
        await WaitUntil(() => driver.Calls.Contains("Read"));

        await session.ShutdownAsync();
        Assert.DoesNotContain("Verify", driver.Calls);

        release.Set();
        await Assert.ThrowsAsync<SessionOperationException>(() => running);
        Assert.DoesNotContain("Verify", driver.Calls);
    }

    // A timer can fire just before the gate's clock, so one call rarely shows this; many calls do.
    [Fact]
    public async Task EveryHungCallGetsItsTimeoutResponse()
    {
        using var release = new ManualResetEventSlim();
        var projects = Enumerable.Range(0, 30).Select(index =>
        {
            string root = Path.Combine(project, "p" + index);
            Directory.CreateDirectory(Path.Combine(root, "src"));
            Directory.CreateDirectory(Path.Combine(root, "tests"));
            File.WriteAllText(Path.Combine(root, "project.json"), """{"paths":{"workbook":"book.xlsx"}}""");
            File.WriteAllText(Path.Combine(root, "book.xlsx"), "");
            return root;
        }).ToArray();
        var hung = new FakeExcelDriver { Before = name => { if (name == "Read") release.Wait(); } };
        var sessions = projects.Select(_ => new ExcelSession(_ => hung, true, TimeSpan.FromMilliseconds(100), Generous)).ToArray();
        try
        {
            await Task.WhenAll(sessions.Zip(projects, (item, root) => item.StartAsync(root)));

            var calls = sessions.Select(async item =>
                (await Assert.ThrowsAsync<SessionOperationException>(() => item.GetRangeAsync("S", "A1"))).Error.Detail);
            string[] details = await Task.WhenAll(calls).WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken);

            Assert.All(details, detail => Assert.Equal("OPERATION_TIMEOUT", detail));
        }
        finally
        {
            release.Set();
            foreach (var item in sessions) item.Dispose();
        }
    }

    [Theory]
    [InlineData("Generate")]
    [InlineData("Identify")]
    [InlineData("Open")]
    public async Task ShutdownLeavesAFailedStartForManualRecovery(string call)
    {
        driver.Before = name => { if (name == call) throw new InvalidOperationException("failed"); };
        await Refused(() => session.StartAsync(project));
        driver.Calls.Clear();

        await session.ShutdownAsync();

        Assert.Empty(driver.Calls);
        Assert.Equal("isolated", session.Summary.State);
    }

    [Fact]
    public async Task ShutdownLeavesAnIsolatedSessionForManualRecovery()
    {
        await session.StartAsync(project);
        driver.Before = name => { if (name == "Write") throw new InvalidOperationException("unknown"); };
        await Refused(() => session.SetRangeAsync("S", "A1", [[1d]], null));
        driver.Calls.Clear();

        await session.ShutdownAsync();

        Assert.Empty(driver.Calls);
        Assert.Equal("isolated", session.Summary.State);
    }

    [Fact]
    public async Task AnUnconfirmedExitIsolatesAndKeepsTheLock()
    {
        await session.StartAsync(project);
        driver.ExitConfirmed = false;

        var refused = await Refused(() => session.EndAsync());

        Assert.Equal("STATE_UNKNOWN", refused.Code);
        Assert.Equal("isolated", session.Summary.State);
        using var second = Create();
        Assert.Equal("PROJECT_LOCKED", (await Refused(() => second.StartAsync(project))).Detail);
    }

    [Fact]
    public async Task ALateOpenDoesNotBecomeReady()
    {
        using var timed = Create(lifecycleTimeout: TimeSpan.FromMilliseconds(200));
        using var release = new ManualResetEventSlim();
        driver.Before = name => { if (name == "Open") release.Wait(); };

        var failure = await Assert.ThrowsAsync<SessionOperationException>(() =>
            timed.StartAsync(project).WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken));
        release.Set();
        await WaitUntil(() => !timed.Summary.Busy);

        Assert.Equal("OPERATION_TIMEOUT", failure.Error.Detail);
        Assert.Equal("isolated", timed.Summary.State);
        Assert.DoesNotContain("session_start.completed", File.ReadAllText(failure.Error.Log!));
    }

    [Fact]
    public async Task AnUnfinishedSafetyCheckIsUnknownAndIsolates()
    {
        await session.StartAsync(project);
        driver.Before = name => { if (name == "Verify") throw new InvalidOperationException("no answer"); };

        var refused = await Refused(() => session.EndAsync());

        Assert.Equal("STATE_UNKNOWN", refused.Code);
        Assert.Equal("EXCEL_NEEDS_REVIEW", refused.Detail);
        Assert.Equal("isolated", session.Summary.State);
    }

    [Fact]
    public async Task TheLogKeepsTheWrappedCause()
    {
        driver.Before = name => { if (name == "Open") throw new InvalidOperationException("cause detail"); };

        var refused = await Refused(() => session.StartAsync(project));

        Assert.Contains("cause detail", File.ReadAllText(refused.Log!));
    }

    [Fact]
    public async Task ShutdownWithoutASessionDoesNothing()
    {
        await session.ShutdownAsync();

        Assert.Empty(driver.Calls);
    }
}
