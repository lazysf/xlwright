using System.Text;

namespace XlWright.Core.Com;

/// <summary>
/// The XLMAIN windows of a process and the workbook windows (EXCEL7) under each (docs/notes/leftover-excel.md).
/// It reads windows only and never touches COM.
/// </summary>
internal static class ExcelWindows
{
    public static List<(IntPtr Main, List<IntPtr> Books)> Of(int pid)
    {
        var mains = new List<(IntPtr, List<IntPtr>)>();
        NativeMethods.EnumWindows((hwnd, parameter) =>
        {
            _ = NativeMethods.GetWindowThreadProcessId(hwnd, out var owner);
            if (owner == pid && ClassOf(hwnd) == "XLMAIN")
            {
                var books = new List<IntPtr>();
                NativeMethods.EnumChildWindows(hwnd, (child, childParameter) =>
                {
                    if (ClassOf(child) == "EXCEL7")
                    {
                        books.Add(child);
                    }

                    return true;
                }, IntPtr.Zero);
                mains.Add((hwnd, books));
            }

            return true;
        }, IntPtr.Zero);
        return mains;
    }

    private static string ClassOf(IntPtr hwnd)
    {
        var name = new StringBuilder(256);
        _ = NativeMethods.GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }
}

/// <summary>Shows the workbook windows of a leftover Excel without COM (docs/notes/human-open-and-show.md).</summary>
public sealed class WorkbookWindows : IWorkbookWindows
{
    private const int ShowCommand = 5; // SW_SHOW

    public int Show(int pid)
    {
        var shown = 0;
        foreach (var (main, _) in ExcelWindows.Of(pid).Where(window => window.Books.Count > 0))
        {
            _ = NativeMethods.ShowWindow(main, ShowCommand);
            _ = NativeMethods.SetForegroundWindow(main);
            shown++;
        }

        return shown;
    }

    public bool HasWorkbookWindow(int pid) => ExcelWindows.Of(pid).Any(window => window.Books.Count > 0);
}
