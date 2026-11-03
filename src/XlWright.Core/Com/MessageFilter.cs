using System.Runtime.InteropServices;

namespace XlWright.Core.Com;

/// <summary>
/// Retries SERVERCALL_RETRYLATER for up to 10 seconds in total; does not retry SERVERCALL_REJECTED
/// (docs/architecture.md "期限と大きさの上限"). Registered once on the session's STA thread.
/// </summary>
internal sealed class MessageFilter : IOleMessageFilter
{
    private const int ServerCallRetryLater = 2;
    private const int RetryLimitMs = 10_000;
    private const int RetryDelayMs = 100;
    private const int PendingMsgWaitDefProcess = 2;

    public static void Register() =>
        Marshal.ThrowExceptionForHR(NativeMethods.CoRegisterMessageFilter(new MessageFilter(), out _));

    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) =>
        dwRejectType == ServerCallRetryLater && dwTickCount < RetryLimitMs ? RetryDelayMs : -1;

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => PendingMsgWaitDefProcess;
}
