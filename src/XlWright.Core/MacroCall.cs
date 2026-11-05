using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace XlWright.Core;

/// <summary>
/// A macro as macro_run takes it: &lt;module&gt;.&lt;procedure&gt;, without a workbook qualifier, which xlwright adds
/// (docs/architecture.md "マクロの実行").
/// </summary>
public sealed partial record MacroName(string Module, string Procedure)
{
    public override string ToString() => $"{Module}.{Procedure}";

    public static (MacroName? Name, ToolFailure? Failure) Parse(string text)
    {
        var match = Pattern().Match(text);
        return match.Success
            ? (new MacroName(match.Groups["module"].Value, match.Groups["procedure"].Value), null)
            : (null, new ToolFailure(ToolFailure.InvalidRequest, "INVALID_MACRO_NAME",
                $"macro \"{text}\" is not <module>.<procedure>, such as Calc.Twice. Leave out the workbook name; xlwright runs the macro in the main workbook."));
    }

    // A VBA name starts with a letter and holds letters, digits, and underscores (docs/notes/vba-sync.md).
    [GeneratedRegex(@"^(?<module>\p{L}[\p{L}\p{Nd}_]*)\.(?<procedure>\p{L}[\p{L}\p{Nd}_]*)\z", RegexOptions.CultureInvariant)]
    private static partial Regex Pattern();
}

/// <summary>
/// One answer of macro_run's dialogs (docs/api.md "ダイアログへの事前の応答"): Kind is message, input, open_file,
/// save_file, or folder; Text is the button name, the text, or the path; null cancels.
/// </summary>
public sealed record DialogAnswer(string Kind, string? Text)
{
    public static readonly IReadOnlyList<string> Kinds = ["message", "input", "open_file", "save_file", "folder"];

    public static readonly IReadOnlyList<string> Buttons = ["ok", "cancel", "yes", "no", "abort", "retry", "ignore"];

    public bool IsPath => Kind is "open_file" or "save_file" or "folder";

    /// <summary>
    /// The answers with each path made full from the app folder, or PATH_NOT_ALLOWED when one points outside it.
    /// Checked before Excel is touched.
    /// </summary>
    public static (IReadOnlyList<DialogAnswer>? Answers, ToolFailure? Failure) Resolve(AppFolder app, IReadOnlyList<DialogAnswer> answers)
    {
        var resolved = new List<DialogAnswer>();
        foreach (var answer in answers)
        {
            if (!answer.IsPath || answer.Text is null)
            {
                resolved.Add(answer);
                continue;
            }

            // The folder itself is a good answer for folder.
            var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.Combine(app.FullPath, answer.Text)));
            if (!full.Equals(app.FullPath, StringComparison.OrdinalIgnoreCase)
                && !full.StartsWith(app.FullPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                return (null, new ToolFailure(ToolFailure.InvalidRequest, "PATH_NOT_ALLOWED",
                    $"dialogs: path must be inside the app folder, relative to it. Rejected: {answer.Text}"));
            }

            resolved.Add(answer with { Text = full });
        }

        return (resolved, null);
    }
}

/// <summary>The module macro_run adds for one call, and the conversion of what it returns (docs/architecture.md "実行ごとに作るモジュール").</summary>
public static class MacroCall
{
    public const string ModuleName = "XlWrightCall";
    public const int MaxArguments = 30;

    /// <summary>
    /// A string return value longer than this in JSON goes into a result file, and the response carries its first
    /// <see cref="SummaryCharacters"/> characters. It leaves room in the 20,000 characters for the sync lists.
    /// </summary>
    public const int MaxReturnCharacters = 10_000;

    public const int SummaryCharacters = 1_000;

    /// <summary>
    /// A Run function that calls the macro by name, so its run-time error reaches this On Error; Application.Run
    /// does not pass it on (docs/notes/macro-run.md). Each argument is in parentheses, so a typed ByRef parameter
    /// takes a copy. Run returns Array(xlwrightResult, XlWright.EndRun()): xlwrightResult is Array(0, value), Array(1, number,
    /// description, source), or XlWright.Result's array, and EndRun gives Array(answered, unused). XlWright.BeginRun
    /// takes the dialog answers and marks the call as inside xlwright (docs/architecture.md "xlwrightのVBAのモジュール").
    /// Err is read into xlwrightResult before EndRun. The name is not "result", which differs from XlWright.Result
    /// only in case and which the VBE may recase across the project (not checked on a real Excel).
    /// </summary>
    public static string Generate(MacroName macro, int argumentCount, bool isFunction)
    {
        var arguments = string.Join(", ", Enumerable.Range(0, argumentCount).Select(i => $"(args({i}))"));
        var call = isFunction
            ? $"    xlwrightResult = XlWright.Result({macro}({arguments}))\r\n"
            : $"    Call {macro}{(argumentCount == 0 ? "" : $"({arguments})")}\r\n    xlwrightResult = Array(0, Empty)\r\n";
        return new StringBuilder()
            .Append("Option Explicit\r\n\r\n")
            .Append("Public Function Run(ByVal args As Variant, ByVal answers As Variant) As Variant\r\n")
            .Append("    Dim xlwrightResult As Variant\r\n")
            .Append("    On Error GoTo Failed\r\n")
            .Append("    XlWright.BeginRun answers\r\n")
            .Append(call)
            .Append("    Run = Array(xlwrightResult, XlWright.EndRun())\r\n")
            .Append("    Exit Function\r\n")
            .Append("Failed:\r\n")
            .Append("    xlwrightResult = Array(1, Err.Number, Err.Description, Err.Source)\r\n")
            .Append("    Run = Array(xlwrightResult, XlWright.EndRun())\r\n")
            .Append("End Function\r\n")
            .ToString();
    }

    /// <summary>The return value as JSON (a number, string, boolean, or null), or else the name of its type.</summary>
    public static (JsonNode? Value, string? TypeName) ReturnValue(object? value) => value switch
    {
        null or DBNull => (null, null),
        string text => (JsonValue.Create(text), null),
        bool flag => (JsonValue.Create(flag), null),
        byte or short or int or long or float or double or decimal => (JsonValue.Create(Convert.ToDouble(value)), null),
        _ => (null, value.GetType().Name),
    };
}
