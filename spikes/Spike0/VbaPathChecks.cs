using System.Runtime.InteropServices;
using System.Text;

namespace Spike0;

internal static class VbaPathChecks
{
    internal static void Run(string directory, Action<string, Action> check)
    {
        check("T15 fixed argument rejects arbitrary paths", () =>
        {
            Require(VbaPathProbe.Parse(["--excel-t15-vba"]));
            Require(!VbaPathProbe.Parse(["--excel-t15-vba", "other.xlsm"]));
            Require(!VbaPathProbe.Parse(["--excel-t15"]));
        });
        check("T15 requires user access and no denying policy", () =>
        {
            VbaPathProbe.RequireAccess(new(1, 2, null, null, null));
            VbaPathProbe.RequireAccess(new(1, null, 1, 1, 1));
            foreach (VbaPathProbe.Settings refused in new VbaPathProbe.Settings[]
            {
                new(null, 1, null, null, null), new(0, 1, null, null, null), new("1", 1, null, null, null),
                new(1, 1, 0, null, null), new(1, 1, null, 0, null), new(1, 1, null, null, 0)
            })
                Refused(refused);
        });
        check("T15 module is one ASCII standard module and is never overwritten", () =>
        {
            string run = Directory.CreateDirectory(Path.Combine(directory, "t15-module")).FullName;
            VbaPathProbe.WriteModule(run);
            byte[] bytes = File.ReadAllBytes(VbaPathProbe.ModulePath(run));
            Require(bytes.All(value => value < 0x80));
            string text = Encoding.ASCII.GetString(bytes);
            Require(text == VbaPathProbe.ModuleSource && !text.Replace("\r\n", "").Contains('\n'));
            Require(text.StartsWith($"Attribute VB_Name = \"{VbaPathProbe.ModuleName}\"\r\n", StringComparison.Ordinal));
            Require(text.Contains($"Public Sub {VbaPathProbe.MacroName}()", StringComparison.Ordinal));
            Require(text.Contains($"= {VbaPathProbe.MarkedValue}\r\n", StringComparison.Ordinal));
            try { VbaPathProbe.WriteModule(run); }
            catch (IOException) { return; }
            throw new InvalidOperationException("The module was overwritten.");
        });
        check("T15 late binding reaches IDispatch without a PIA and keeps the COM failure", () =>
        {
            // An in-process system object stands in for VBIDE, so no Excel is activated.
            object dictionary = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", throwOnError: true)!)!;
            try
            {
                Require(Marshal.IsComObject(dictionary));
                Require(Convert.ToInt32(VbaPathProbe.Get(dictionary, "Count")) == 0);
                try { VbaPathProbe.Call(dictionary, "Add", "key", 1); }
                catch (InvalidDataException) { } // Add returns no value; the helper reports that on purpose.
                Require(Convert.ToInt32(VbaPathProbe.Get(dictionary, "Count")) == 1);
                try { VbaPathProbe.Call(dictionary, "Add", "key", 2); }
                catch (COMException) { return; }
                throw new InvalidOperationException("The COM failure was not preserved.");
            }
            finally { Marshal.FinalReleaseComObject(dictionary); }
        });
    }

    private static void Refused(VbaPathProbe.Settings settings)
    {
        try { VbaPathProbe.RequireAccess(settings); }
        catch (InvalidOperationException) { return; }
        throw new InvalidOperationException("A disabled VBA setting was accepted.");
    }

    private static void Require(bool condition)
    {
        if (!condition) throw new InvalidOperationException("T15 check failed.");
    }
}
