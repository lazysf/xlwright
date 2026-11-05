using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace XlWright.Core;

public enum VbaModuleKind
{
    Standard,
    Class,
}

/// <summary>A module ready for VBComponents.Import: the converted file at ImportPath.</summary>
public sealed record VbaModule(string Name, VbaModuleKind Kind, string ImportPath);

/// <summary>
/// The code for the sheet or workbook module whose CodeName is Name, without the header and Attribute lines, for
/// CodeModule.AddFromString. File is the source relative to the app folder.
/// </summary>
public sealed record VbaDocumentModule(string Name, string File, string Code);

/// <summary>
/// Reads src/modules, src/classes, and src/workbook, writes the xlwright VBA module, and converts each source for
/// Import or AddFromString (docs/architecture.md "VBAの反映", "xlwrightのVBAのモジュール"; docs/notes/vba-sync.md,
/// docs/notes/test-copy.md).
/// </summary>
public static partial class VbaSources
{
    public const string XlWrightName = "XlWright";

    // The class header Import needs to make a class module; without it a .cls becomes a standard module.
    private const string ClassHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    static VbaSources() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>The system ANSI code page, which Import reads files in (docs/notes/vba-run.md).</summary>
    public static Encoding SystemAnsiEncoding() => AnsiEncoding((int)Com.NativeMethods.GetACP());

    /// <summary>The code page's encoding, failing on characters it cannot hold.</summary>
    public static Encoding AnsiEncoding(int codePage) =>
        Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

    /// <summary>
    /// The text of src/modules/XlWright.bas. The first line carries the version that wrote it. Result receives a
    /// macro's return value as a ByVal Variant, so an object stays an object and its default member is not read;
    /// a CVErr value and a Long both reach .NET as Int32 (docs/notes/macro-run.md).
    /// </summary>
    public static string XlWrightText(string version) =>
        $"' XlWright {version}: xlwright writes this module before each sync. Do not edit it.\r\n" +
        "Option Explicit\r\n" +
        "\r\n" +
        "' Whether xlwright runs the code, and the answers for the dialog functions.\r\n" +
        "Private mRunning As Boolean\r\n" +
        "Private mAnswers As Collection\r\n" +
        "Private mUsed As Long\r\n" +
        "\r\n" +
        "' macro_run passes a macro's return value here: Array(0, value), or Array(2, type name) for a value it does not return.\r\n" +
        "Public Function Result(ByVal value As Variant) As Variant\r\n" +
        "    If IsObject(value) Or IsArray(value) Then\r\n" +
        "        Result = Array(2, TypeName(value))\r\n" +
        "    ElseIf VarType(value) = vbError Or VarType(value) = vbDate Then\r\n" +
        "        Result = Array(2, TypeName(value))\r\n" +
        "    Else\r\n" +
        "        Result = Array(0, value)\r\n" +
        "    End If\r\n" +
        "End Function\r\n" +
        DialogFunctions;

    // The dialog functions. Outside xlwright they show the standard dialog; inside, they use the answers in order.
    private static readonly string DialogFunctions = """

        ' xlwright calls BeginRun before each macro_run call and each test, and EndRun after it. The app's code does not call them.
        ' answers is macro_run's dialogs as an array of Array(kind, value); EndRun returns Array(answered, unused).
        Public Sub BeginRun(Optional ByVal answers As Variant)
            mRunning = True
            mUsed = 0
            Set mAnswers = New Collection
            If IsMissing(answers) Then Exit Sub
            If Not IsArray(answers) Then Exit Sub
            Dim i As Long
            For i = LBound(answers) To UBound(answers)
                mAnswers.Add answers(i)
            Next
        End Sub

        Public Function EndRun() As Variant
            If mAnswers Is Nothing Then
                EndRun = Array(mUsed, 0)
            Else
                EndRun = Array(mUsed, mAnswers.Count)
            End If
            mRunning = False
            Set mAnswers = Nothing
        End Function

        ' A test queues the answer for the next dialog function. kind is message, input, open_file, save_file, or folder.
        ' value is a button name such as "yes", the text, or a path; False cancels. A relative path is from ThisWorkbook.Path.
        Public Sub Answer(ByVal kind As String, ByVal value As Variant)
            If Not mRunning Then Fail "Answer", "it works only while xlwright runs the code."
            Select Case kind
            Case "message", "input", "open_file", "save_file", "folder"
            Case Else
                Fail "Answer", "kind must be message, input, open_file, save_file, or folder, not " & kind & "."
            End Select
            mAnswers.Add Array(kind, value)
        End Sub

        ' Use in place of MsgBox. Returns the button pressed, such as vbYes.
        Public Function Message(ByVal prompt As Variant, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As Variant) As VbMsgBoxResult
            If Not mRunning Then
                If IsMissing(title) Then
                    Message = VBA.MsgBox(prompt, buttons)
                Else
                    Message = VBA.MsgBox(prompt, buttons, title)
                End If
                Exit Function
            End If

            Dim value As Variant
            value = NextAnswer("Message", "message")
            If Cancels("Message", value) Then value = "cancel"
            Dim names As String
            names = Choose((buttons And 7) + 1, "ok", "ok cancel", "abort retry ignore", "yes no cancel", "yes no", "retry cancel", "", "")
            If InStr(" " & names & " ", " " & value & " ") = 0 Or InStr(value, " ") > 0 Then
                Fail "Message", "the answer " & value & " is not a button of the dialog (" & names & ")."
            End If

            Select Case value
            Case "ok": Message = vbOK
            Case "cancel": Message = vbCancel
            Case "abort": Message = vbAbort
            Case "retry": Message = vbRetry
            Case "ignore": Message = vbIgnore
            Case "yes": Message = vbYes
            Case "no": Message = vbNo
            End Select
        End Function

        ' Asks for text. Returns the text, or False when cancelled.
        Public Function InputText(ByVal prompt As Variant, Optional ByVal title As Variant, Optional ByVal default As Variant) As Variant
            If Not mRunning Then
                InputText = Application.InputBox(prompt, title, default, Type:=2)
                Exit Function
            End If

            Dim value As Variant
            value = NextAnswer("InputText", "input")
            If Cancels("InputText", value) Then
                InputText = False
            Else
                InputText = CStr(value)
            End If
        End Function

        ' Asks for a file to open. Returns its full path, or False when cancelled.
        Public Function OpenFile(Optional ByVal filter As Variant, Optional ByVal title As Variant) As Variant
            If mRunning Then
                OpenFile = AnsweredPath("OpenFile", "open_file")
            Else
                OpenFile = Application.GetOpenFilename(filter, , title)
            End If
        End Function

        ' Asks where to save a file. Returns its full path, or False when cancelled.
        Public Function SaveFile(Optional ByVal initial As Variant, Optional ByVal filter As Variant, Optional ByVal title As Variant) As Variant
            If mRunning Then
                SaveFile = AnsweredPath("SaveFile", "save_file")
            Else
                SaveFile = Application.GetSaveAsFilename(initial, filter, , title)
            End If
        End Function

        ' Asks for a folder. Returns its full path, or False when cancelled.
        Public Function Folder(Optional ByVal title As Variant) As Variant
            If mRunning Then
                Folder = AnsweredPath("Folder", "folder")
                Exit Function
            End If

            Dim picker As Object
            Set picker = Application.FileDialog(4) ' msoFileDialogFolderPicker
            If Not IsMissing(title) Then picker.Title = title
            If picker.Show = -1 Then
                Folder = picker.SelectedItems(1)
            Else
                Folder = False
            End If
        End Function

        Private Function NextAnswer(ByVal name As String, ByVal kind As String) As Variant
            If mAnswers.Count = 0 Then Fail name, "no answer is left for the dialog."
            Dim item As Variant
            item = mAnswers(1)
            mAnswers.Remove 1
            mUsed = mUsed + 1
            If item(0) <> kind Then Fail name, "the next answer is for " & item(0) & ", not " & kind & "."
            NextAnswer = item(1)
        End Function

        Private Function AnsweredPath(ByVal name As String, ByVal kind As String) As Variant
            Dim value As Variant
            value = NextAnswer(name, kind)
            If Cancels(name, value) Then
                AnsweredPath = False
            ElseIf CStr(value) Like "?:\*" Or Left$(CStr(value), 2) = "\\" Then
                AnsweredPath = CStr(value)
            Else
                AnsweredPath = ThisWorkbook.Path & "\" & Replace(CStr(value), "/", "\")
            End If
        End Function

        ' False cancels; True is not an answer.
        Private Function Cancels(ByVal name As String, ByVal value As Variant) As Boolean
            If VarType(value) <> vbBoolean Then Exit Function
            If value Then Fail name, "True is not an answer; pass False to cancel."
            Cancels = True
        End Function

        Private Sub Fail(ByVal name As String, ByVal reason As String)
            Err.Raise vbObjectError + 1, "XlWright", "XlWright." & name & ": " & reason
        End Sub

        """.ReplaceLineEndings("\r\n");

    /// <summary>
    /// Checks every source before Excel is touched, writes XlWright.bas when it is missing, older than
    /// <paramref name="version"/>, or of the same version with other text, and writes the converted files under
    /// &lt;app&gt;/.xlwright/sync/&lt;id&gt;/. The sheet and workbook modules come back as code, in file name order.
    /// The caller deletes that folder after the sync.
    /// </summary>
    public static (IReadOnlyList<VbaModule>? Modules, IReadOnlyList<VbaDocumentModule>? Documents, ToolFailure? Failure) Prepare(
        AppFolder app, Encoding ansi, string version)
    {
        var src = Path.Combine(app.FullPath, "src");
        var files = new List<(string Path, VbaModuleKind Kind)>();
        var documentFiles = new List<string>();
        foreach (var (folder, extension, kind) in new[] { ("modules", ".bas", VbaModuleKind.Standard), ("classes", ".cls", VbaModuleKind.Class), ("workbook", ".cls", (VbaModuleKind?)null) })
        {
            var full = Path.Combine(src, folder);
            if (!Directory.Exists(full))
            {
                continue;
            }

            foreach (var entry in Directory.EnumerateFileSystemEntries(full).Order(StringComparer.Ordinal))
            {
                var name = Path.GetFileName(entry);
                if (name.StartsWith('.'))
                {
                    continue;
                }

                if (Directory.Exists(entry))
                {
                    return Invalid(app, entry, null, $"{Relative(app, entry)} is a folder inside src/{folder}/. Put module files directly in src/{folder}/.");
                }

                if (!Path.GetExtension(name).Equals(extension, StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid(app, entry, null,
                        $"{name} is in src/{folder}/, which holds only {extension} files. Standard modules go in src/modules/*.bas, " +
                        "class modules in src/classes/*.cls, and sheet and workbook modules in src/workbook/<CodeName>.cls.");
                }

                if (kind is null)
                {
                    documentFiles.Add(entry);
                    continue;
                }

                if (Path.GetFileNameWithoutExtension(name).Equals(MacroCall.ModuleName, StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid(app, entry, null, $"{name}: the name {MacroCall.ModuleName} is kept for the module macro_run adds while it runs. Rename the module.");
                }

                if (kind == VbaModuleKind.Class && Path.GetFileNameWithoutExtension(name).Equals(XlWrightName, StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid(app, entry, null, $"{name}: the name {XlWrightName} is kept for the module xlwright writes. Rename the class.");
                }

                files.Add((entry, kind.Value));
            }
        }

        // Read and convert everything before writing anything, so a rejected source leaves the app folder as it was.
        var converted = new List<(string Name, VbaModuleKind Kind, byte[] Bytes)>();
        foreach (var (path, kind) in files)
        {
            var (module, failure) = Convert(app, path, kind, File.ReadAllBytes(path), ansi);
            if (failure is not null)
            {
                return (null, null, failure);
            }

            converted.Add(module!.Value);
        }

        var documents = new List<VbaDocumentModule>();
        foreach (var path in documentFiles)
        {
            var (document, failure) = ConvertDocument(app, path, File.ReadAllBytes(path), ansi);
            if (failure is not null)
            {
                return (null, null, failure);
            }

            documents.Add(document!);
        }

        var xlWrightPath = Path.Combine(src, "modules", XlWrightName + ".bas");
        if (NeedsXlWright(xlWrightPath, version))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(xlWrightPath)!);
            File.WriteAllText(xlWrightPath, XlWrightText(version), StrictUtf8);
            converted.RemoveAll(module => module.Name.Equals(XlWrightName, StringComparison.OrdinalIgnoreCase));
            var (module, failure) = Convert(app, xlWrightPath, VbaModuleKind.Standard, File.ReadAllBytes(xlWrightPath), ansi);
            if (failure is not null)
            {
                return (null, null, failure);
            }

            converted.Add(module!.Value);
        }

        var syncFolder = Path.Combine(app.FullPath, ".xlwright", "sync", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(syncFolder);
        var modules = new List<VbaModule>();
        foreach (var (kind, group) in new[] { VbaModuleKind.Standard, VbaModuleKind.Class }.Select(k => (k, converted.Where(m => m.Kind == k))))
        {
            foreach (var (name, _, bytes) in group.OrderBy(m => m.Name, StringComparer.Ordinal))
            {
                var importPath = Path.Combine(syncFolder, name + (kind == VbaModuleKind.Class ? ".cls" : ".bas"));
                File.WriteAllBytes(importPath, bytes);
                modules.Add(new VbaModule(name, kind, importPath));
            }
        }

        return (modules, documents, null);
    }

    // UTF-8 to the ANSI code page with CRLF, the class header when missing, and Attribute VB_Name from the file name.
    private static ((string, VbaModuleKind, byte[])?, ToolFailure?) Convert(AppFolder app, string path, VbaModuleKind kind, byte[] bytes, Encoding ansi)
    {
        var (lines, nameLine, name, failure) = Read(app, path, bytes, ansi);
        if (failure is not null)
        {
            return (null, failure);
        }

        var addHeader = kind == VbaModuleKind.Class && !lines.FirstOrDefault(line => line.Trim().Length > 0)?.Trim().StartsWith("VERSION 1.0 CLASS", StringComparison.OrdinalIgnoreCase) == true;
        if (nameLine < 0)
        {
            var at = kind == VbaModuleKind.Class && !addHeader ? lines.FindIndex(line => line.Trim().Equals("END", StringComparison.OrdinalIgnoreCase)) + 1 : 0;
            lines.Insert(at, $"Attribute VB_Name = \"{name}\"");
        }

        var result = (addHeader ? ClassHeader : "") + string.Join("\r\n", lines) + "\r\n";
        return ((name, kind, ansi.GetBytes(result)), null);
    }

    // The code only: the class header (VERSION 1.0 CLASS to END) and every Attribute line are removed, because
    // AddFromString writes them as code. The name is the CodeName from the file name.
    private static (VbaDocumentModule?, ToolFailure?) ConvertDocument(AppFolder app, string path, byte[] bytes, Encoding ansi)
    {
        var (lines, _, name, failure) = Read(app, path, bytes, ansi);
        if (failure is not null)
        {
            return (null, failure);
        }

        var first = lines.FindIndex(line => line.Trim().Length > 0);
        if (first >= 0 && lines[first].Trim().StartsWith("VERSION 1.0 CLASS", StringComparison.OrdinalIgnoreCase))
        {
            var end = lines.FindIndex(first, line => line.Trim().Equals("END", StringComparison.OrdinalIgnoreCase));
            lines.RemoveRange(0, end + 1);
        }

        lines.RemoveAll(line => AttributeLine().IsMatch(line));
        var code = string.Join("\r\n", lines).Trim('\r', '\n');
        return (new VbaDocumentModule(name, Relative(app, path), code), null);
    }

    // UTF-8 to lines, with the checks every source passes: Attribute VB_Name matches the file name, the name
    // XlWright is kept, and each line fits the ANSI code page. The name is VB_Name, or the file name without it.
    private static (List<string> Lines, int NameLine, string Name, ToolFailure? Failure) Read(AppFolder app, string path, byte[] bytes, Encoding ansi)
    {
        string text;
        try
        {
            text = StrictUtf8.GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return ([], -1, "", Invalid(app, path, null, "The file is not valid UTF-8. Save it as UTF-8.").Failure);
        }

        var lines = text.TrimStart('﻿').ReplaceLineEndings("\n").TrimEnd('\n').Split('\n').ToList();
        var fileName = Path.GetFileNameWithoutExtension(path);
        var name = fileName;
        var nameLine = lines.FindIndex(line => VbName().IsMatch(line));
        if (nameLine >= 0)
        {
            name = VbName().Match(lines[nameLine]).Groups["name"].Value;
            if (!name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
            {
                return ([], -1, "", Invalid(app, path, nameLine + 1,
                    $"Attribute VB_Name is \"{name}\", but the file is {Path.GetFileName(path)}. Make them the same.").Failure);
            }
        }

        if (name.Equals(XlWrightName, StringComparison.OrdinalIgnoreCase) && !Path.GetFileName(path).Equals(XlWrightName + ".bas", StringComparison.OrdinalIgnoreCase))
        {
            return ([], -1, "", Invalid(app, path, nameLine + 1, $"The name {XlWrightName} is kept for the module xlwright writes. Rename the module.").Failure);
        }

        for (var i = 0; i < lines.Count; i++)
        {
            try
            {
                ansi.GetBytes(lines[i]);
            }
            catch (EncoderFallbackException)
            {
                return ([], -1, "", Invalid(app, path, i + 1,
                    $"The line has a character that the system code page ({ansi.CodePage}) cannot hold, so the VBA editor would change it. Replace it, for example with ChrW.").Failure);
            }
        }

        return (lines, nameLine, name, null);
    }

    private static bool NeedsXlWright(string path, string version)
    {
        if (!File.Exists(path))
        {
            return true;
        }

        // A development build keeps the version while the text changes, so the same version is compared by text.
        var text = File.ReadAllText(path);
        var match = VersionLine().Match(text);
        if (!match.Success)
        {
            return true;
        }

        var written = Version.Parse(match.Groups["version"].Value);
        var current = Version.Parse(version.Split('+')[0]);
        return written < current || (written == current && text != XlWrightText(version));
    }

    private static (IReadOnlyList<VbaModule>?, IReadOnlyList<VbaDocumentModule>?, ToolFailure? Failure) Invalid(AppFolder app, string path, int? line, string reason) =>
        (null, null, new ToolFailure(ToolFailure.InvalidRequest, "SOURCE_INVALID",
            $"{Relative(app, path)}{(line is null ? "" : $" line {line}")}: {reason} Nothing was changed in the workbook.",
            new JsonObject { ["file"] = Relative(app, path), ["line"] = line, ["reason"] = reason }));

    private static string Relative(AppFolder app, string path) => Path.GetRelativePath(app.FullPath, path).Replace('\\', '/');

    [GeneratedRegex("""^\s*Attribute\s+VB_Name\s*=\s*"(?<name>[^"]*)"\s*\z""", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex VbName();

    [GeneratedRegex(@"^\s*Attribute\s", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex AttributeLine();

    [GeneratedRegex(@"^' XlWright (?<version>[0-9]+\.[0-9]+\.[0-9]+)[: ]", RegexOptions.CultureInvariant)]
    private static partial Regex VersionLine();
}
