<#
.SYNOPSIS
Compiles the VBA project of a workbook and reports the compile error, if any.

.DESCRIPTION
Opens the workbook in a hidden Excel with macros disabled, makes its project the active one in the VBE and
executes the VBE command "Compile VBAProject" (command bar control 578). There is no COM method for
compiling. The workbook is closed without saving.

The outcome is read from what the VBE does next:
- The control becomes disabled: the project compiled.
- A dialog (#32770) of this Excel instance appears: a compile error. Its text and the selection of the
  active code pane (the place of the error) are recorded, the dialog is closed with WM_CLOSE and the VBE
  window, which the error makes visible on the screen, is hidden again.
- Neither within -DialogTimeoutSeconds: reported as an error, not as a pass.

Only one compile error is reported: the one the VBE shows. The place is left null when the VBE has no
active code pane. A locked project is not compiled and is reported as an error.

If the dialog cannot be closed, the Excel instance started by this script is stopped by its PID. Other
Excel processes are never touched.

The result is written as UTF-8 JSON to -OutFile (stdout of powershell.exe garbles Japanese in Git Bash):
Compiled (true / false, null when the check itself failed), CompileError (Message, Module, Line, Column)
and Error. Exit code is 0 only when the project compiled.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300,
    # How long to wait for the compile to finish or the error dialog to appear.
    [ValidateRange(1, 3600)][int]$DialogTimeoutSeconds = 5
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

if (-not ('XlWright.CompileDialog' -as [type])) {
    Add-Type -Namespace XlWright -Name CompileDialog -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr hWnd, System.IntPtr lParam);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, System.IntPtr lParam);
[DllImport("user32.dll")] static extern bool EnumChildWindows(System.IntPtr parent, EnumProc callback, System.IntPtr lParam);
[DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
[DllImport("user32.dll")] static extern bool IsWindowVisible(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsWindow(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool PostMessage(System.IntPtr hWnd, uint message, System.IntPtr wParam, System.IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(System.IntPtr hWnd, System.Text.StringBuilder text, int count);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(System.IntPtr hWnd, System.Text.StringBuilder text, int count);

static string ClassOf(System.IntPtr hWnd) {
    var text = new System.Text.StringBuilder(256);
    GetClassName(hWnd, text, text.Capacity);
    return text.ToString();
}

// The compile error dialog is owned by the VBE window, not by the Excel window, so it is looked up among
// the top-level windows of the process.
public static System.IntPtr FindDialog(uint processId) {
    System.IntPtr found = System.IntPtr.Zero;
    EnumWindows((hWnd, lParam) => {
        uint owner;
        GetWindowThreadProcessId(hWnd, out owner);
        if (owner != processId || !IsWindowVisible(hWnd) || ClassOf(hWnd) != "#32770") return true;
        found = hWnd;
        return false;
    }, System.IntPtr.Zero);
    return found;
}

// The message is in the Static children; the buttons (OK, Help) are left out.
public static string[] StaticTexts(System.IntPtr dialog) {
    var texts = new System.Collections.Generic.List<string>();
    EnumChildWindows(dialog, (hWnd, lParam) => {
        if (ClassOf(hWnd) != "Static") return true;
        var text = new System.Text.StringBuilder(2048);
        GetWindowText(hWnd, text, text.Capacity);
        if (text.Length > 0) texts.Add(text.ToString());
        return true;
    }, System.IntPtr.Zero);
    return texts.ToArray();
}
'@
}

$result = [ordered]@{ Compiled = $null; CompileError = $null; Error = $null }
try {
    Use-ExcelWorkbook -Path $Path -TimeoutSeconds $TimeoutSeconds -ScriptBlock {
        param($xl, $wb)
        $excelPid = Get-ExcelProcessId $xl
        $project = $wb.VBProject
        if ($project.Protection -ne 0) { throw 'The VBA project is locked and cannot be compiled from outside.' }
        $vbe = $xl.VBE
        # Only the active project is compiled.
        $vbe.ActiveVBProject = $project
        $control = $vbe.CommandBars.FindControl(1, 578)  # msoControlButton, "Compile VBAProject"
        if ($null -eq $control) { throw 'The VBE command "Compile VBAProject" (578) was not found.' }
        # A disabled control means the project is already in the compiled state; Execute would fail.
        if (-not $control.Enabled) { $result.Compiled = $true; return }

        $vbeWasVisible = [bool]$vbe.MainWindow.Visible
        $control.Execute()
        # Execute returns at once; the error dialog appears a moment later.
        $dialog = [IntPtr]::Zero
        $deadline = (Get-Date).AddSeconds($DialogTimeoutSeconds)
        do {
            $dialog = [XlWright.CompileDialog]::FindDialog([uint32]$excelPid)
            if ($dialog -ne [IntPtr]::Zero) { break }
            if (-not $control.Enabled) { $result.Compiled = $true; return }
            Start-Sleep -Milliseconds 100
        } while ((Get-Date) -lt $deadline)
        if ($dialog -eq [IntPtr]::Zero) {
            throw "The compile could not be confirmed: no error dialog and the project is not in the compiled state after $DialogTimeoutSeconds s."
        }

        $message = ([XlWright.CompileDialog]::StaticTexts($dialog) -join ' ') -replace '\s+', ' '
        $place = [ordered]@{ Message = $message.Trim(); Module = $null; Line = $null; Column = $null }
        $pane = $vbe.ActiveCodePane
        if ($null -ne $pane) {
            $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
            $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
            if ($startLine -gt 0) {
                $place.Module = $pane.CodeModule.Parent.Name
                $place.Line = $startLine
                $place.Column = $startColumn
            }
        }
        $result.Compiled = $false
        $result.CompileError = $place

        [void][XlWright.CompileDialog]::PostMessage($dialog, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
        $deadline = (Get-Date).AddSeconds(3)
        while ([XlWright.CompileDialog]::IsWindow($dialog) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 50 }
        if ([XlWright.CompileDialog]::IsWindow($dialog)) {
            # The instance is blocked in a modal state; closing the workbook would hang. Stop only this instance.
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            throw "The compile error dialog could not be closed. Stopped the Excel instance of this script (PID $excelPid)."
        }
        if (-not $vbeWasVisible) { $vbe.MainWindow.Visible = $false }
    }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
if (-not $result.Compiled) { exit 1 }
