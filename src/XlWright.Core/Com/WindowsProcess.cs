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

    private WindowsProcess(SafeProcessHandle handle, ProcessFacts facts, int openError = 0)
    {
        _handle = handle;
        Facts = facts;
        OpenError = openError;
    }

    public ProcessFacts Facts { get; }

    /// <summary>The Win32 error of OpenProcess when the handle is invalid; 0 when it opened.</summary>
    public int OpenError { get; }

    /// <summary>The Win32 error OpenProcess returns when no process has the PID.</summary>
    public const int NoSuchProcess = 87;

    /// <summary>
    /// Opens the process. If it cannot be opened or read, the facts carry no image name or command line,
    /// so ownership is not confirmed and the process is never terminated.
    /// </summary>
    public static WindowsProcess Open(int pid)
    {
        var handle = NativeMethods.OpenProcess(NativeMethods.ProcessAccess, false, pid);
        if (handle.IsInvalid)
        {
            return new WindowsProcess(handle, new ProcessFacts(pid, "", UnknownTime, null), Marshal.GetLastPInvokeError());
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

    // docs/architecture.md "失敗の返し方と記録": an XLMAIN of the process is disabled, and the process has a visible,
    // enabled window that is not XLMAIN. A dialog a macro shows has no owner, or the VBE as owner (docs/notes/dialog-probe.md).
    public DialogText? FindDialog()
    {
        if (_handle.IsInvalid)
        {
            return null;
        }

        var windows = new List<IntPtr>();
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(hwnd, out var pid);
            if (pid == (uint)Facts.Pid)
            {
                windows.Add(hwnd);
            }

            return true;
        }, IntPtr.Zero);

        if (!windows.Any(hwnd => ClassOf(hwnd) == "XLMAIN" && !NativeMethods.IsWindowEnabled(hwnd)))
        {
            return null;
        }

        var dialog = windows.FirstOrDefault(hwnd =>
            ClassOf(hwnd) != "XLMAIN" && NativeMethods.IsWindowVisible(hwnd) && NativeMethods.IsWindowEnabled(hwnd));
        return dialog == IntPtr.Zero ? null : ReadDialog(dialog);
    }

    public IReadOnlyList<bool> MainWindowsEnabled()
    {
        var enabled = new List<bool>();
        if (_handle.IsInvalid)
        {
            return enabled;
        }

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(hwnd, out var pid);
            if (pid == (uint)Facts.Pid && ClassOf(hwnd) == "XLMAIN")
            {
                enabled.Add(NativeMethods.IsWindowEnabled(hwnd));
            }

            return true;
        }, IntPtr.Zero);
        return enabled;
    }

    public void Dispose() => _handle.Dispose();

    // The title and the visible Static children, read with WM_GETTEXT within 500 ms in all, as ExcelMcp does,
    // so a hung window cannot hold the watch; the caller ends the Excel whether or not the text was read.
    private static DialogText ReadDialog(IntPtr dialog)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var title = Text(dialog, watch);
        var lines = new List<string>();
        NativeMethods.EnumChildWindows(dialog, (hwnd, _) =>
        {
            if (ClassOf(hwnd) == "Static" && NativeMethods.IsWindowVisible(hwnd) && Text(hwnd, watch) is { Length: > 0 } text)
            {
                lines.Add(text);
            }

            return watch.ElapsedMilliseconds < ReadBudgetMs;
        }, IntPtr.Zero);
        return new DialogText(title, string.Join("\n", lines));
    }

    private const int ReadBudgetMs = 500;

    private static string Text(IntPtr hwnd, System.Diagnostics.Stopwatch watch)
    {
        const uint WmGetText = 0x000D;
        const uint SmtoAbortIfHung = 0x0002;
        var remaining = ReadBudgetMs - (int)watch.ElapsedMilliseconds;
        if (remaining <= 0)
        {
            return "";
        }

        var buffer = new StringBuilder(4096);
        return NativeMethods.SendMessageTimeout(hwnd, WmGetText, buffer.Capacity, buffer, SmtoAbortIfHung, (uint)Math.Min(50, remaining), out _) == IntPtr.Zero
            ? ""
            : buffer.ToString().ReplaceLineEndings("\n").Trim();
    }

    private static string ClassOf(IntPtr hwnd)
    {
        var name = new StringBuilder(256);
        _ = NativeMethods.GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }

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
