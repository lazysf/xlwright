using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using Microsoft.Win32;

namespace Spike0;

// T15 confirms only that the VBA path exists. S5, S6, S7 and S15 stay in R4.
internal static class VbaPathProbe
{
    internal const string Scenario = "t15-vba";
    internal const string ModuleName = "T15Module";
    internal const string MacroName = "T15Mark";
    internal const double MarkedValue = 15;
    internal const int StandardModuleType = 1; // vbext_ct_StdModule
    // Kept in source so the prepared manifest's source hash covers the imported code.
    internal const string ModuleSource =
        "Attribute VB_Name = \"T15Module\"\r\n" +
        "Option Explicit\r\n" +
        "\r\n" +
        "Public Sub T15Mark()\r\n" +
        "    ThisWorkbook.Worksheets(\"Lifecycle\").Range(\"A2\").Value2 = 15\r\n" +
        "End Sub\r\n";

    private const string SecurityPath = @"Software\Microsoft\Office\16.0\Excel\Security";
    private const string PolicyPath = @"Software\Policies\Microsoft\Office\16.0\Excel\Security";
    private const string CloudPolicyPath = @"Software\Policies\Microsoft\Cloud\Office\16.0\Excel\Security";

    internal sealed record Settings(object? AccessVbom, object? VbaWarnings, object? UserPolicy,
        object? MachinePolicy, object? CloudPolicy);

    internal static bool Parse(string[] args) => args is ["--excel-t15-vba"];

    internal static Settings ReadSettings()
    {
        static object? Read(RegistryKey hive, string path, string name)
        {
            using var key = hive.OpenSubKey(path);
            return key?.GetValue(name);
        }
        return new(Read(Registry.CurrentUser, SecurityPath, "AccessVBOM"),
            Read(Registry.CurrentUser, SecurityPath, "VBAWarnings"),
            Read(Registry.CurrentUser, PolicyPath, "AccessVBOM"),
            Read(Registry.LocalMachine, PolicyPath, "AccessVBOM"),
            Read(Registry.CurrentUser, CloudPolicyPath, "AccessVBOM"));
    }

    // Runs before the attempt stamp, so a setting that is off stops without consuming the single attempt.
    internal static void RequireAccess(Settings settings)
    {
        if (settings.AccessVbom is not 1 || settings.UserPolicy is not (null or 1) ||
            settings.MachinePolicy is not (null or 1) || settings.CloudPolicy is not (null or 1))
            throw new InvalidOperationException(
                "VBA project access is not enabled for this user. No Excel was started and no attempt was recorded.");
    }

    internal static string ModulePath(string directory) => Path.Combine(directory, ModuleName + ".bas");

    internal static void WriteModule(string directory)
    {
        using var stream = new FileStream(ModulePath(directory), FileMode.CreateNew, FileAccess.Write, FileShare.Read);
        stream.Write(Encoding.ASCII.GetBytes(ModuleSource));
        stream.Flush(true);
    }

    // IDispatch through InvokeMember, not dynamic: the DLR caches type information, and a retained
    // reference into Excel could keep the owned process alive after Quit.
    internal static object Get(object target, string property) =>
        Invoke(target, property, BindingFlags.GetProperty, []);

    internal static object Call(object target, string method, params object[] arguments) =>
        Invoke(target, method, BindingFlags.InvokeMethod, arguments);

    private static object Invoke(object target, string member, BindingFlags kind, object[] arguments)
    {
        try
        {
            return target.GetType().InvokeMember(member, kind, null, target, arguments) ??
                throw new InvalidDataException("VBA object model returned no value: " + member);
        }
        catch (TargetInvocationException error) when (error.InnerException is not null)
        {
            // Keep the COM HRESULT and message in the diagnostics instead of the reflection wrapper.
            ExceptionDispatchInfo.Capture(error.InnerException).Throw();
            throw;
        }
    }

    // Diagnostics are long JSON lines; the person running the console reads these instead.
    internal static void AnnounceStart() => Announce(
        "T15 started. No input is needed. Do not type in this console.", "Wait for the final >>> line.");

    internal static void AnnounceEnd(int exitCode)
    {
        if (exitCode == 0) Announce("T15 finished with exit code 0.", "Report the exit code. Do not run it again.");
        else Announce($"STOP: T15 ended with exit code {exitCode}.",
            "Do not run it again. Do not close Excel or delete files. Report the exit code and wait for recovery approval.");
    }

    private static void Announce(string first, string second) =>
        Console.WriteLine($"{Environment.NewLine}>>> {first}{Environment.NewLine}>>> {second}{Environment.NewLine}");
}
