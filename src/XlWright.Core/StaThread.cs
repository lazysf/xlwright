using System.Windows.Forms;

namespace XlWright.Core;

// One long-lived STA with a message loop owns every COM object. Callers receive results, never COM references.
// ExecutionGate decides whether work may start; this thread only runs posted work in order.
internal sealed class StaThread : IDisposable
{
    private readonly object sync = new();
    private readonly SynchronizationContext loop;
    private bool stopped;

    internal StaThread()
    {
        var ready = new TaskCompletionSource<SynchronizationContext>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
                // Posts become window messages, so they run inside Application.Run like the Spike0 dispatch check.
                var context = new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(context);
                ready.SetResult(context);
                Application.Run();
            }
            catch (Exception error) { ready.TrySetException(error); }
            finally
            {
                // tradeoff: work posted while the loop fails is never completed. The caller's deadline and
                // ExecutionGate isolate the session; rebuild with tracked completions if a caller cannot bound the wait.
                lock (sync) stopped = true;
            }
        }) { IsBackground = true, Name = "xlwright Excel STA" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        loop = ready.Task.GetAwaiter().GetResult();
    }

    internal Task<T> Invoke<T>(Func<T> work)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        lock (sync)
        {
            ObjectDisposedException.ThrowIf(stopped, this);
            loop.Post(_ =>
            {
                try { completion.SetResult(work()); }
                catch (Exception error) { completion.SetException(error); }
            }, null);
        }
        return completion.Task;
    }

    // Ends the loop after work already posted. A running call is not stopped; only new work is refused.
    public void Dispose()
    {
        lock (sync)
        {
            if (stopped) return;
            stopped = true;
            loop.Post(_ => Application.ExitThread(), null);
        }
    }
}
