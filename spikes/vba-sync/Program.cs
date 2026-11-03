// Throwaway experiment for vba_sync (docs/api.md, stage 1). It starts one dedicated Excel on an STA thread,
// works on a new workbook that has no file, imports module files written to a temporary folder, and prints the
// name and type of each imported component and the exceptions Import raises.
// It only ends the Excel whose ownership it confirmed: a new PID, named EXCEL.EXE, created after the start request.
// Every COM object it obtains is released, so Excel can end after Quit.
//
// Usage: dotnet run --project spikes\vba-sync

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    private const int StdModule = 1;
    private const int ClassModule = 2;

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

    [DllImport("kernel32.dll")]
    private static extern uint GetACP();

    private static readonly string Folder = Path.Combine(Path.GetTempPath(), "xlwright-spike-vba-sync-" + Guid.NewGuid().ToString("N"));
    private static Encoding s_ansi = Encoding.ASCII;

    private static int Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        s_ansi = Encoding.GetEncoding((int)GetACP());
        Console.WriteLine($"ANSI code page {GetACP()} ({s_ansi.WebName})");
        var exit = 1;
        var thread = new Thread(() => exit = Run());
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        Directory.Delete(Folder, recursive: true);
        Console.WriteLine($"temporary folder removed: {!Directory.Exists(Folder)}");
        return exit;
    }

    private static int Run()
    {
        Directory.CreateDirectory(Folder);
        var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Console.WriteLine($"before: {before.Count} EXCEL process(es)");
        var requested = DateTime.Now;
        dynamic? app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!)!;
        GetWindowThreadProcessId(new IntPtr((int)app!.Hwnd), out var rawPid);
        var process = Process.GetProcessById((int)rawPid);
        var owned = !before.Contains(process.Id) &&
            process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) &&
            process.StartTime > requested.AddSeconds(-1);
        Console.WriteLine($"started PID {process.Id}, owned {owned}");
        if (!owned)
        {
            Console.WriteLine("ownership not confirmed; leaving it untouched");
            return 1;
        }

        app.UserControl = true;
        dynamic workbooks = app.Workbooks;
        dynamic book = workbooks.Add();
        dynamic project = book.VBProject;
        dynamic components = project.VBComponents;
        try
        {
            Console.WriteLine($"Protection {project.Protection}");
            Trials(components);
        }
        finally
        {
            Release(components);
            Release(project);
            app.DisplayAlerts = false;
            book.Close(false);
            Release(book);
            Release(workbooks);
            app.Quit();
            Release(app);
            app = null;
        }

        var ended = process.WaitForExit(TimeSpan.FromSeconds(75));
        if (!ended)
        {
            process.Kill();
            ended = process.WaitForExit(TimeSpan.FromSeconds(5));
            Console.WriteLine("terminated after Quit");
        }

        var after = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        Console.WriteLine($"ended {ended}; after: {after.Count} EXCEL process(es), unchanged from before {after.SetEquals(before)}");
        return 0;
    }

    private static void Trials(dynamic components)
    {
        const string Header = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";
        Import(components, "NoHeader.cls", "Attribute VB_Name = \"NoHeader\"\r\nOption Explicit\r\nPublic Value As Long\r\n");
        Import(components, "OnlyVersion.cls", Header + "Attribute VB_Name = \"OnlyVersion\"\r\nOption Explicit\r\nPublic Value As Long\r\n");
        Import(components, "NoName.bas", "Option Explicit\r\nPublic Function One() As Long\r\n    One = 1\r\nEnd Function\r\n");
        Import(components, "NoNameClass.cls", Header + "Option Explicit\r\n");
        Import(components, "CommentFirst.bas", "' a comment before the name\r\nAttribute VB_Name = \"CommentFirst\"\r\nOption Explicit\r\n");
        Import(components, "Sheet1.bas", "Attribute VB_Name = \"Sheet1\"\r\nOption Explicit\r\n");
        Import(components, "ThisWorkbook.bas", "Attribute VB_Name = \"ThisWorkbook\"\r\nOption Explicit\r\n");
        Import(components, "Twice.bas", "Attribute VB_Name = \"Twice\"\r\nOption Explicit\r\n");
        Import(components, "Twice2.bas", "Attribute VB_Name = \"Twice\"\r\nOption Explicit\r\n");
        Import(components, "Digit.bas", "Attribute VB_Name = \"1Digit\"\r\nOption Explicit\r\n");
        Import(components, "Long.bas", $"Attribute VB_Name = \"{new string('L', 32)}\"\r\nOption Explicit\r\n");
        Import(components, "Thirty1.bas", $"Attribute VB_Name = \"{new string('T', 31)}\"\r\nOption Explicit\r\n");
        Import(components, "Japanese.bas", "Attribute VB_Name = \"日本語名\"\r\nOption Explicit\r\n");
        Import(components, "Reserved.bas", "Attribute VB_Name = \"Option\"\r\nOption Explicit\r\n");
        Import(components, "LfOnly.bas", "Attribute VB_Name = \"LfOnly\"\nOption Explicit\nPublic Function Two() As Long\n    Two = 2\nEnd Function\n");
        Import(components, "Empty.bas", "");

        Console.WriteLine("== components after the imports");
        for (var i = 1; i <= (int)components.Count; i++)
        {
            dynamic component = components.Item(i);
            dynamic code = component.CodeModule;
            Console.WriteLine($"  {(string)component.Name} type {(int)component.Type} lines {(int)code.CountOfLines}");
            Release(code);
            Release(component);
        }

        Console.WriteLine("== remove every standard and class module");
        for (var i = (int)components.Count; i >= 1; i--)
        {
            dynamic component = components.Item(i);
            int type = component.Type;
            if (type is StdModule or ClassModule)
            {
                components.Remove(component);
            }

            Release(component);
        }

        Console.WriteLine($"  left: {(int)components.Count}");
        Import(components, "Twice3.bas", "Attribute VB_Name = \"Twice\"\r\nOption Explicit\r\n");
    }

    private static void Import(dynamic components, string file, string text)
    {
        var path = Path.Combine(Folder, file);
        File.WriteAllText(path, text, s_ansi);
        try
        {
            dynamic imported = components.Import(path);
            Console.WriteLine($"  {file}: imported as {(string)imported.Name} type {(int)imported.Type}");
            Release(imported);
        }
        catch (Exception e)
        {
            Console.WriteLine($"  {file}: {e.GetType().Name} 0x{e.HResult:X8} {e.Message.ReplaceLineEndings(" ")}");
        }
    }

    private static void Release(object comObject) => Marshal.FinalReleaseComObject(comObject);
}
