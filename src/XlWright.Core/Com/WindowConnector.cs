using System.Runtime.InteropServices;
using System.Text;

namespace XlWright.Core.Com;

/// <summary>
/// Connects to a leftover Excel through a workbook window (EXCEL7) under its XLMAIN, the way
/// docs/notes/leftover-excel.md found: XLMAIN itself gives no object, and an Excel without a workbook has no EXCEL7.
/// </summary>
public sealed class WindowConnector : IExcelConnector
{
    private const uint NativeObjectModel = 0xFFFFFFF0;

    public LeftoverWindow Connect(int pid)
    {
        var books = WorkbookWindows(pid);
        if (books.Count == 0)
        {
            return new LeftoverWindow.NoWorkbookWindow();
        }

        MessageFilter.Register();
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        foreach (var book in books)
        {
            if (NativeMethods.AccessibleObjectFromWindow(book, NativeObjectModel, ref iid, out var window) != 0 || window is null)
            {
                continue;
            }

            dynamic? app = null;
            try
            {
                // A Protected View window gives an object without Application (docs/notes/leftover-excel.md).
                app = ((dynamic)window).Application;
                _ = NativeMethods.GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out var owner);
                if (owner == pid)
                {
                    var connected = new LeftoverWindow.Connected(new ComExcel(app));
                    app = null;
                    return connected;
                }
            }
            catch (Exception e) when (e is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
            finally
            {
                if (app is not null)
                {
                    Marshal.FinalReleaseComObject(app);
                }

                Marshal.FinalReleaseComObject(window);
            }
        }

        return new LeftoverWindow.NotReachable("no workbook window gave an Excel of the process");
    }

    private static List<IntPtr> WorkbookWindows(int pid)
    {
        var books = new List<IntPtr>();
        NativeMethods.EnumWindows((hwnd, parameter) =>
        {
            _ = NativeMethods.GetWindowThreadProcessId(hwnd, out var owner);
            if (owner == pid && ClassOf(hwnd) == "XLMAIN")
            {
                NativeMethods.EnumChildWindows(hwnd, (child, childParameter) =>
                {
                    if (ClassOf(child) == "EXCEL7")
                    {
                        books.Add(child);
                    }

                    return true;
                }, IntPtr.Zero);
            }

            return true;
        }, IntPtr.Zero);
        return books;
    }

    private static string ClassOf(IntPtr hwnd)
    {
        var name = new StringBuilder(256);
        _ = NativeMethods.GetClassName(hwnd, name, name.Capacity);
        return name.ToString();
    }
}
