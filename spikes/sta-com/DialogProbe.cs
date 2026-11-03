// Finds a dialog that blocks the owned Excel and reads its text, the way docs/architecture.md describes:
// XLMAIN is disabled, and a visible, enabled window of the process has an owner chain that reaches XLMAIN.

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

internal static class DialogProbe
{
    private const uint GwOwner = 4;
    private const uint WmGetText = 0x000D;
    private const uint SmtoAbortIfHung = 0x0002;

    private sealed record Window(IntPtr Handle, IntPtr Owner, bool Visible, bool Enabled, string ClassName, string Title);

    // Returns the dialog handle, or IntPtr.Zero. `rule` names the rule that matched:
    // "owner chain" (the design), or "disabled XLMAIN" (any XLMAIN of the process is disabled and the process has
    // a visible, enabled window that is not XLMAIN). `windows` describes the visible windows and every XLMAIN.
    public static IntPtr Find(int pid, out string windows, out string rule)
    {
        var list = new List<Window>();
        EnumWindows((hwnd, lParam) =>
        {
            _ = Native.GetWindowThreadProcessId(hwnd, out var owner);
            if (owner != (uint)pid) return true;
            list.Add(new Window(hwnd, GetWindow(hwnd, GwOwner), IsWindowVisible(hwnd), IsWindowEnabled(hwnd), ClassOf(hwnd), TitleOf(hwnd)));
            return true;
        }, IntPtr.Zero);

        var byHandle = list.ToDictionary(w => w.Handle);
        var shown = list.Where(w => w.Visible || w.ClassName == "XLMAIN").Select(w =>
        {
            var chain = new List<string>();
            for (var o = w.Owner; o != IntPtr.Zero && chain.Count < 10; o = byHandle.TryGetValue(o, out var ow) ? ow.Owner : IntPtr.Zero)
                chain.Add(byHandle.TryGetValue(o, out var x) ? x.ClassName : "<other process>");
            return $"{w.ClassName} visible={w.Visible} enabled={w.Enabled} \"{w.Title}\" owners=[{string.Join(" > ", chain)}]";
        });
        windows = string.Join(" | ", shown);

        rule = "owner chain";
        var disabledRoots = list.Where(w => w.ClassName == "XLMAIN" && !w.Enabled).ToList();
        foreach (var root in disabledRoots)
        {
            foreach (var w in list.Where(w => w.Visible && w.Enabled && w.Handle != root.Handle))
            {
                for (var o = w.Owner; o != IntPtr.Zero; o = byHandle.TryGetValue(o, out var ow) ? ow.Owner : IntPtr.Zero)
                {
                    if (o == root.Handle) return w.Handle;
                    if (!byHandle.ContainsKey(o)) break;
                }
            }
        }

        rule = "disabled XLMAIN";
        if (disabledRoots.Count > 0)
        {
            var candidate = list.FirstOrDefault(w => w.Visible && w.Enabled && w.ClassName != "XLMAIN");
            if (candidate != null) return candidate.Handle;
        }

        rule = "";
        return IntPtr.Zero;
    }

    // The title and the visible Static and Button children, each read with WM_GETTEXT, within the time budget.
    public static string ReadText(IntPtr dialog, int budgetMs)
    {
        var watch = Stopwatch.StartNew();
        var parts = new List<string>();
        void Read(IntPtr hwnd)
        {
            var remaining = budgetMs - (int)watch.ElapsedMilliseconds;
            if (remaining <= 0) return;
            var buffer = new StringBuilder(2048);
            if (SendMessageTimeout(hwnd, WmGetText, (IntPtr)buffer.Capacity, buffer, SmtoAbortIfHung, (uint)Math.Min(50, remaining), out _) == IntPtr.Zero)
            {
                parts.Add("<timeout>");
                return;
            }

            var text = buffer.ToString().Trim();
            if (text.Length > 0) parts.Add(text.Replace("\r", "").Replace("\n", "\\n"));
        }

        Read(dialog);
        EnumChildWindows(dialog, (hwnd, lParam) =>
        {
            if (watch.ElapsedMilliseconds >= budgetMs) return false;
            var cls = ClassOf(hwnd);
            if (IsWindowVisible(hwnd) && (cls == "Static" || cls == "Button")) Read(hwnd);
            return true;
        }, IntPtr.Zero);
        return parts.Count == 0 ? "<no text>" : string.Join(" / ", parts.Select(p => $"\"{p}\""));
    }

    private static string ClassOf(IntPtr hwnd)
    {
        var name = new StringBuilder(256);
        _ = GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }

    private static string TitleOf(IntPtr hwnd)
    {
        var title = new StringBuilder(256);
        _ = GetWindowText(hwnd, title, title.Capacity);
        return title.ToString();
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr hwnd, uint command);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr hwnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam, StringBuilder lParam, uint flags, uint timeoutMs, out IntPtr result);
}
