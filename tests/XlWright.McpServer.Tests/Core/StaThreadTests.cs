using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests.Core;

public sealed class StaThreadTests
{
    private static readonly TimeSpan Wait = TimeSpan.FromSeconds(5);

    [Fact(Timeout = 10000)]
    public async Task RunsWorkOnOneStaThreadInOrder()
    {
        using var sta = new StaThread();
        var order = new List<int>();
        Task<(ApartmentState, int)>[] work = [.. Enumerable.Range(0, 5).Select(index => sta.Invoke(() =>
        {
            order.Add(index);
            return (Thread.CurrentThread.GetApartmentState(), Environment.CurrentManagedThreadId);
        }))];
        var results = await Task.WhenAll(work).WaitAsync(Wait, TestContext.Current.CancellationToken);

        Assert.All(results, result => Assert.Equal(ApartmentState.STA, result.Item1));
        Assert.Single(results.Select(result => result.Item2).Distinct());
        Assert.NotEqual(Environment.CurrentManagedThreadId, results[0].Item2);
        Assert.Equal([0, 1, 2, 3, 4], order);
    }

    [Fact(Timeout = 10000)]
    public async Task WorkFailureReachesTheCallerAndTheLoopContinues()
    {
        using var sta = new StaThread();
        await Assert.ThrowsAsync<IOException>(() =>
            sta.Invoke<bool>(() => throw new IOException("Synthetic failure.")).WaitAsync(Wait, TestContext.Current.CancellationToken));

        Assert.True(await sta.Invoke(() => true).WaitAsync(Wait, TestContext.Current.CancellationToken));
    }

    [Fact(Timeout = 10000)]
    public async Task DisposeRunsPostedWorkAndRefusesNewWork()
    {
        var sta = new StaThread();
        Task<bool> posted = sta.Invoke(() => true);
        sta.Dispose();

        Assert.True(await posted.WaitAsync(Wait, TestContext.Current.CancellationToken));
        // Invoke refuses synchronously; it does not return a faulted task.
        Assert.Throws<ObjectDisposedException>(() => { _ = sta.Invoke(() => true); });
        sta.Dispose();
    }
}
