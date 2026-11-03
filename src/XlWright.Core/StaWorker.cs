using System.Collections.Concurrent;

namespace XlWright.Core;

/// <summary>
/// The dedicated STA thread of one session. Excel objects are created, used, and released only on it
/// (docs/architecture.md "構成"). A call that never returns leaves the thread blocked; the session then
/// treats it as a pending COM call.
/// </summary>
public sealed class StaWorker : IDisposable
{
    private readonly BlockingCollection<Action> _work = new();
    private readonly Thread _thread;

    public StaWorker()
    {
        _thread = new Thread(Loop) { IsBackground = true, Name = "xlwright STA" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> Run<T>(Func<T> work)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _work.Add(() =>
        {
            try
            {
                completion.SetResult(work());
            }
            catch (Exception e)
            {
                completion.SetException(e);
            }
        });
        return completion.Task;
    }

    public Task Run(Action work) => Run(() =>
    {
        work();
        return true;
    });

    public void Dispose() => _work.CompleteAdding();

    private void Loop()
    {
        foreach (var work in _work.GetConsumingEnumerable())
        {
            work();
        }
    }
}
