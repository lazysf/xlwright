using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Spike0;

internal sealed record S11Window(long Hwnd, string ClassName, bool Visible, bool Enabled, long Owner);
internal sealed record S11WindowSnapshot(string Status, S11Window[] Windows, string Response, int Error);

// This probe reads window metadata only. It never obtains an Excel COM object or window text.
internal static class S11Windows
{
    internal static S11WindowSnapshot Capture(ProcessIdentity identity, nint mainWindow)
    {
        Process process;
        try { process = Process.GetProcessById(identity.Pid); }
        catch (ArgumentException) { return new("process-absent", [], "unavailable", 0); }
        using var lifetime = process;
        if (!identity.Matches(ProcessIdentity.Read(process)))
            throw new InvalidOperationException("S11 process identity changed.");
        _ = process.SafeHandle;
        var windows = new List<S11Window>();
        Exception? failure = null;
        bool complete = EnumWindows((window, _) =>
        {
            try
            {
                if (GetWindowThreadProcessId(window, out uint pid) == 0 || pid != identity.Pid) return true;
                var name = new StringBuilder(256);
                if (GetClassName(window, name, name.Capacity) == 0)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                nint owner = GetWindow(window, 4);
                if (owner != 0 && (GetWindowThreadProcessId(owner, out uint ownerPid) == 0 || ownerPid != pid))
                    owner = 0;
                windows.Add(new(window, name.ToString(), IsWindowVisible(window), IsWindowEnabled(window), owner));
                return true;
            }
            catch (Exception error) { failure = error; return false; }
        }, 0);
        if (failure is not null) throw failure;
        if (!complete) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        if (process.HasExited) return new("process-exited", windows.ToArray(), "unavailable", 0);
        if (!identity.Matches(ProcessIdentity.Read(process)))
            throw new InvalidOperationException("S11 process identity changed during observation.");
        if (mainWindow == 0 || GetWindowThreadProcessId(mainWindow, out uint mainPid) == 0 || mainPid != identity.Pid)
            return new("observed", windows.ToArray(), "window-unavailable", 0);
        Marshal.SetLastPInvokeError(0);
        // WM_NULL has no dialog response. Do not use BROADCAST or NOTIMEOUTIFNOTHUNG.
        nint response = SendMessageTimeout(mainWindow, 0, 0, 0, 0x23, 200, out _);
        int nativeError = Marshal.GetLastWin32Error();
        return new("observed", windows.ToArray(), response != 0 ? "responded" : "failed-or-timed-out", nativeError);
    }

    private delegate bool WindowCallback(nint window, nint parameter);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowCallback callback, nint parameter);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(nint window, out uint pid);
    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetClassName(nint window, StringBuilder name, int capacity);
    [DllImport("user32.dll")]
    private static extern nint GetWindow(nint window, uint command);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint window);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowEnabled(nint window);
    [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
    private static extern nint SendMessageTimeout(nint window, uint message, nuint wparam, nint lparam,
        uint flags, uint timeout, out nuint result);
}
