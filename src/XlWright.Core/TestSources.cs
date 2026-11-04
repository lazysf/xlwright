using System.Text;
using System.Text.RegularExpressions;

namespace XlWright.Core;

/// <summary>What test_run imports into the test copy, and the tests it runs in order (docs/api.md "test_run").</summary>
public sealed record TestPlan(IReadOnlyList<VbaModule> Modules, IReadOnlyList<MacroName> Tests);

public static partial class VbaSources
{
    public const string AssertName = "XlWrightAssert";

    /// <summary>
    /// The assertion module, imported into the test copy only (docs/architecture.md "テスト"). A failure raises an
    /// error whose Source is XlWrightAssert, which tells it apart from other run-time errors.
    /// </summary>
    public static string AssertText() =>
        $"Attribute VB_Name = \"{AssertName}\"\r\n" +
        "' test_run imports this module into the test copy only.\r\n" +
        "Option Explicit\r\n" +
        "\r\n" +
        "Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = \"\")\r\n" +
        "    If IsObject(expected) Or IsObject(actual) Or IsArray(expected) Or IsArray(actual) Then\r\n" +
        "        Report \"AreEqual: objects and arrays are not compared\", message\r\n" +
        "    End If\r\n" +
        "    Dim same As Boolean\r\n" +
        "    On Error GoTo Different\r\n" +
        "    same = (expected = actual)\r\n" +
        "    On Error GoTo 0\r\n" +
        "    If same Then Exit Sub\r\n" +
        "Different:\r\n" +
        "    Report \"AreEqual: expected \" & Show(expected) & \", actual \" & Show(actual), message\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Public Sub IsTrue(ByVal condition As Boolean, Optional ByVal message As String = \"\")\r\n" +
        "    If Not condition Then Report \"IsTrue: the condition is false\", message\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Public Sub Fail(ByVal message As String)\r\n" +
        $"    Err.Raise vbObjectError + 1, \"{AssertName}\", message\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Private Sub Report(ByVal text As String, ByVal message As String)\r\n" +
        "    If Len(message) > 0 Then text = text & \". \" & message\r\n" +
        $"    Err.Raise vbObjectError + 1, \"{AssertName}\", text\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Private Function Show(ByVal value As Variant) As String\r\n" +
        "    If IsNull(value) Then\r\n" +
        "        Show = \"Null\"\r\n" +
        "    ElseIf IsEmpty(value) Then\r\n" +
        "        Show = \"Empty\"\r\n" +
        "    ElseIf VarType(value) = vbString Then\r\n" +
        "        Show = \"\"\"\" & value & \"\"\"\"\r\n" +
        "    Else\r\n" +
        "        Show = CStr(value)\r\n" +
        "    End If\r\n" +
        "End Function\r\n";

    /// <summary>
    /// Checks tests/ before Excel is touched, finds the tests (a Sub without arguments, not Private, whose name starts
    /// with Test) in file name order, keeps those in <paramref name="only"/> when it is given, and writes the converted
    /// test modules, helper classes, and the assertion module under &lt;app&gt;/.xlwright/sync/&lt;id&gt;/.
    /// The caller deletes that folder after the import.
    /// </summary>
    public static (TestPlan? Plan, ToolFailure? Failure) PrepareTests(AppFolder app, Encoding ansi, IReadOnlyList<string>? only)
    {
        var src = Path.Combine(app.FullPath, "src");
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { XlWrightName };
        foreach (var (folder, extension) in new[] { ("modules", ".bas"), ("classes", ".cls") })
        {
            var full = Path.Combine(src, folder);
            if (Directory.Exists(full))
            {
                taken.UnionWith(Directory.EnumerateFiles(full, "*" + extension).Select(Path.GetFileNameWithoutExtension)!);
            }
        }

        var files = new List<(string Path, VbaModuleKind Kind)>();
        var tests = Path.Combine(app.FullPath, "tests");
        var entries = Directory.Exists(tests) ? Directory.EnumerateFileSystemEntries(tests).Order(StringComparer.Ordinal).ToList() : [];
        foreach (var entry in entries)
        {
            var name = Path.GetFileName(entry);
            var module = Path.GetFileNameWithoutExtension(name);
            if (name.StartsWith('.'))
            {
                continue;
            }

            if (Directory.Exists(entry))
            {
                return (null, Invalid(app, entry, null, $"{Relative(app, entry)} is a folder inside tests/. Put test files directly in tests/.").Failure);
            }

            var extension = Path.GetExtension(name);
            if (!extension.Equals(".bas", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".cls", StringComparison.OrdinalIgnoreCase))
            {
                return (null, Invalid(app, entry, null, $"{name} is in tests/, which holds only test modules (.bas) and helper classes (.cls).").Failure);
            }

            var kind = extension.Equals(".cls", StringComparison.OrdinalIgnoreCase) ? VbaModuleKind.Class : VbaModuleKind.Standard;

            if (module.Equals(AssertName, StringComparison.OrdinalIgnoreCase) || module.Equals(MacroCall.ModuleName, StringComparison.OrdinalIgnoreCase) ||
                module.Equals(XlWrightName, StringComparison.OrdinalIgnoreCase))
            {
                return (null, Invalid(app, entry, null, $"{name}: the name {module} is kept for a module xlwright adds. Rename the test module.").Failure);
            }

            if (taken.Contains(module))
            {
                return (null, Invalid(app, entry, null, $"{name}: src/ has a module named {module}, and the test copy needs both. Rename the test module.").Failure);
            }

            files.Add((entry, kind));
        }

        var found = new List<MacroName>();
        var converted = new List<(string Name, VbaModuleKind Kind, byte[] Bytes)>();
        foreach (var (path, kind) in files)
        {
            var bytes = File.ReadAllBytes(path);
            var (lines, _, name, failure) = Read(app, path, bytes, ansi);
            if (failure is not null)
            {
                return (null, failure);
            }

            if (kind == VbaModuleKind.Standard)
            {
                found.AddRange(lines.Select(line => TestSub().Match(line)).Where(match => match.Success)
                    .Select(match => new MacroName(name, match.Groups["name"].Value)));
            }

            var (module, invalid) = Convert(app, path, kind, bytes, ansi);
            if (invalid is not null)
            {
                return (null, invalid);
            }

            converted.Add(module!.Value);
        }

        if (only is not null)
        {
            var missing = only.FirstOrDefault(name => !found.Any(test => test.ToString().Equals(name, StringComparison.OrdinalIgnoreCase)));
            if (missing is not null)
            {
                return (null, new ToolFailure(ToolFailure.NotFound, "TEST_NOT_FOUND",
                    $"tests/ has no test {missing}. A test is a Sub without arguments, not Private, whose name starts with Test, in tests/*.bas."));
            }

            found = found.Where(test => only.Any(name => test.ToString().Equals(name, StringComparison.OrdinalIgnoreCase))).ToList();
        }

        if (found.Count == 0)
        {
            return (null, new ToolFailure(ToolFailure.NotFound, "TEST_NOT_FOUND",
                "tests/ has no tests. A test is a Sub without arguments, not Private, whose name starts with Test, in tests/*.bas."));
        }

        converted.Add((AssertName, VbaModuleKind.Standard, ansi.GetBytes(AssertText())));
        var syncFolder = Path.Combine(app.FullPath, ".xlwright", "sync", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(syncFolder);
        var modules = new List<VbaModule>();
        foreach (var kind in new[] { VbaModuleKind.Standard, VbaModuleKind.Class })
        {
            foreach (var (name, _, bytes) in converted.Where(m => m.Kind == kind).OrderBy(m => m.Name, StringComparer.Ordinal))
            {
                var importPath = Path.Combine(syncFolder, name + (kind == VbaModuleKind.Class ? ".cls" : ".bas"));
                File.WriteAllBytes(importPath, bytes);
                modules.Add(new VbaModule(name, kind, importPath));
            }
        }

        return (new TestPlan(modules, found), null);
    }

    // VBA keywords are not case sensitive; the Test prefix is, as docs/api.md writes it.
    [GeneratedRegex(@"^\s*(?i:public\s+)?(?i:sub)\s+(?<name>Test[\p{L}\p{Nd}_]*)\s*\(\s*\)", RegexOptions.CultureInvariant)]
    private static partial Regex TestSub();
}
