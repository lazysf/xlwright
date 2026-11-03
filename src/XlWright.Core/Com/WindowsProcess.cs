using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace XlWright.Core.Com;

/// <summary>
/// A process held through one handle. The facts are read through the same handle that waits and terminates,
/// so a reused PID cannot be mistaken for it (docs/architecture.md "完了にする条件").
/// </summary>
internal sealed class WindowsProcess : IOwnedProcess
{
    private readonly SafeProcessHandle _handle;

    private WindowsProcess(SafeProcessHandle handle, ProcessFacts facts)
    {
        _handle = handle;
        Facts = facts;
    }

    public ProcessFacts Facts { get; }

    /// <summary>
    /// Opens the process. If it cannot be opened or read, the facts carry no image name or command line,
    /// so ownership is not confirmed and the process is never terminated.
    /// </summary>
    public static WindowsProcess Open(int pid)
    {
        var handle = NativeMethods.OpenProcess(NativeMethods.ProcessAccess, false, pid);
        if (handle.IsInvalid)
        {
            return new WindowsProcess(handle, new ProcessFacts(pid, "", UnknownTime, null));
        }

        var created = NativeMethods.GetProcessTimes(handle, out var creation, out _, out _, out _)
            ? DateTime.FromFileTimeUtc(creation)
            : UnknownTime;
        return new WindowsProcess(handle, new ProcessFacts(pid, ImagePath(handle), created, CommandLine(handle)));
    }

    /// <summary>A process that could not be found at all; it is never waited on or terminated.</summary>
    public static WindowsProcess Unidentified() =>
        new(new SafeProcessHandle(IntPtr.Zero, ownsHandle: false), new ProcessFacts(0, "", UnknownTime, null));

    private static DateTime UnknownTime => DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);

    public bool WaitForExit(TimeSpan timeout) =>
        !_handle.IsInvalid && NativeMethods.WaitForSingleObject(_handle, (int)timeout.TotalMilliseconds) == 0;

    public void Terminate()
    {
        if (!_handle.IsInvalid)
        {
            NativeMethods.TerminateProcess(_handle, 1);
        }
    }

    public void Dispose() => _handle.Dispose();

    private static string ImagePath(SafeProcessHandle handle)
    {
        var size = 1024;
        var name = new StringBuilder(size);
        return NativeMethods.QueryFullProcessImageName(handle, 0, name, ref size) ? name.ToString() : "";
    }

    // ProcessCommandLineInformation (60) returns a UNICODE_STRING followed by its characters.
    private static string? CommandLine(SafeProcessHandle handle)
    {
        const int ProcessCommandLineInformation = 60;
        _ = NativeMethods.NtQueryInformationProcess(handle, ProcessCommandLineInformation, IntPtr.Zero, 0, out var length);
        if (length <= 0)
        {
            return null;
        }

        var buffer = Marshal.AllocHGlobal(length);
        try
        {
            if (NativeMethods.NtQueryInformationProcess(handle, ProcessCommandLineInformation, buffer, length, out _) != 0)
            {
                return null;
            }

            var bytes = (ushort)Marshal.ReadInt16(buffer, 0);
            return Marshal.PtrToStringUni(Marshal.ReadIntPtr(buffer, IntPtr.Size), bytes / 2);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }
}
