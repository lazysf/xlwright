using System.Windows.Forms;

namespace Spike0;

internal static class StaRunner
{
    internal static Task<bool> Start(Func<bool> scenario)
    {
        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            EventHandler? runOnce = null;
            runOnce = (_, _) =>
            {
                Application.Idle -= runOnce;
                try { completion.TrySetResult(scenario()); }
                catch (Exception error) { completion.TrySetException(error); }
                finally { Application.ExitThread(); }
            };
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
                Application.Idle += runOnce;
                Application.Run();
            }
            catch (Exception error) { completion.TrySetException(error); }
            finally { Application.Idle -= runOnce; }
        }) { IsBackground = true, Name = "Spike0 Excel STA" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }
}
