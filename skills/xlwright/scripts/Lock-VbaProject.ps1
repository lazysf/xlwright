# Locks a closed, disposable publish copy and verifies the saved project.
param([Parameter(Mandatory)][string]$Path)
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1')
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
if (-not ('VbaProjectLock.Native' -as [type])) {
    Add-Type -Namespace VbaProjectLock -Name Native -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
private static extern IntPtr SendMessageTimeout(
    IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);

[DllImport("user32.dll", SetLastError = true)]
public static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

public static long Send(IntPtr window, uint message, int wParam, IntPtr lParam)
{
    IntPtr result;
    // SMTO_ABORTIFHUNG | SMTO_ERRORONEXIT
    if (SendMessageTimeout(window, message, new IntPtr(wParam), lParam, 0x22, 2000, out result) == IntPtr.Zero)
        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
    return result.ToInt64();
}
'@
}

$setProtection = {
    param($vbeHandle, $projectName, $excelPid, $commandReturned)
    $ErrorActionPreference = 'Stop'
    $element = [System.Windows.Automation.AutomationElement]
    $scope = [System.Windows.Automation.TreeScope]
    $dialogCondition = [System.Windows.Automation.PropertyCondition]::new($element::ClassNameProperty, '#32770')
    $dialogHandle = [IntPtr]::Zero
    $accepted = $false

    function Get-ControlHandle([string]$Id, [string]$ClassName) {
        $idCondition = [System.Windows.Automation.PropertyCondition]::new($element::AutomationIdProperty, $Id)
        $controls = @($dialog.FindAll($scope::Descendants, $idCondition) |
            Where-Object { $_.Current.ClassName -eq $ClassName })
        if ($controls.Count -ne 1 -or !$controls[0].Current.NativeWindowHandle) {
            throw "Expected one VBE protection control with ID $Id ($ClassName)."
        }
        [IntPtr]$controls[0].Current.NativeWindowHandle
    }

    try {
        $vbeWindow = $element::FromHandle($vbeHandle)
        $deadline = [DateTime]::UtcNow.AddSeconds(10)
        do {
            $matches = @($vbeWindow.FindAll($scope::Descendants, $dialogCondition) |
                Where-Object { $_.Current.ProcessId -eq $excelPid -and
                    $_.Current.Name.StartsWith($projectName + ' - ', [StringComparison]::Ordinal) })
            if ($matches.Count -gt 1) { throw 'Multiple VBE project properties dialogs found.' }
            if ($matches.Count -eq 1) { break }
        } while ([DateTime]::UtcNow -lt $deadline -and !$commandReturned.Wait(150))
        if ($matches.Count -ne 1) { throw 'VBE project properties dialog not found.' }
        $dialog = $matches[0]
        $dialogHandle = [IntPtr]$dialog.Current.NativeWindowHandle

        # The VBE exposes these legacy controls as panes. Use their Win32 IDs and messages.
        $tabs = Get-ControlHandle '12320' 'SysTabControl32'
        if ([VbaProjectLock.Native]::Send($tabs, 0x1304, 0, 0) -ne 2) {  # TCM_GETITEMCOUNT
            throw 'Unexpected VBE project properties tab layout.'
        }
        [void][VbaProjectLock.Native]::Send($tabs, 0x1330, 1, 0)  # TCM_SETCURFOCUS notifies the dialog
        $lock = Get-ControlHandle '5463' 'Button'
        $password = Get-ControlHandle '5461' 'Edit'
        $confirmation = Get-ControlHandle '5462' 'Edit'
        $ok = Get-ControlHandle '1' 'Button'

        # Use the Edit controls' input path so VBE receives the change notifications.
        $random = [Security.Cryptography.RandomNumberGenerator]::Create()
        $bytes = [byte[]]::new(8)
        try { $random.GetBytes($bytes) }
        finally { $random.Dispose() }
        $text = [BitConverter]::ToString($bytes).Replace('-', '')
        foreach ($field in $password, $confirmation) {
            [void][VbaProjectLock.Native]::Send($field, 0x00B1, 0, -1)  # EM_SETSEL
            foreach ($character in $text.ToCharArray()) {
                [void][VbaProjectLock.Native]::Send($field, 0x0102, [int]$character, 1)  # WM_CHAR
            }
            if ([VbaProjectLock.Native]::Send($field, 0x000E, 0, 0) -ne $text.Length) {  # WM_GETTEXTLENGTH
                throw 'VBE did not accept the generated password.'
            }
        }
        if ([VbaProjectLock.Native]::Send($lock, 0x00F0, 0, 0) -ne 1) {  # BM_GETCHECK
            [void][VbaProjectLock.Native]::Send($lock, 0x00F5, 0, 0)  # BM_CLICK
        }
        if ([VbaProjectLock.Native]::Send($lock, 0x00F0, 0, 0) -ne 1) {
            throw 'VBE did not select Lock project for viewing.'
        }
        [void][VbaProjectLock.Native]::Send($ok, 0x00F5, 0, 0)  # BM_CLICK
        if (!$commandReturned.Wait(5000)) {
            throw 'VBE did not accept the project protection settings.'
        }
        $accepted = $true
    }
    finally {
        if (!$accepted -and !$commandReturned.IsSet) {
            if ($dialogHandle -ne [IntPtr]::Zero) {
                [void][VbaProjectLock.Native]::PostMessage($dialogHandle, 0x0010, 0, 0)  # WM_CLOSE
            }
            if (!$commandReturned.Wait(3000)) {
                # Only the dedicated Excel process can be blocked by our modal command.
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

Use-ExcelWorkbook -Path $Path -Visible -Save -ScriptBlock {
    param($xl, $wb)
    $project = $vbe = $vbeWindow = $commandBars = $properties = $null
    try {
        $project = $wb.VBProject
        if ($project.Protection -ne 0) { throw 'The publish copy must have an unlocked VBA project.' }
        $vbe = $xl.VBE
        $vbeWindow = $vbe.MainWindow
        $vbeWindow.Visible = $true
        $vbe.ActiveVBProject = $project
        $commandBars = $vbe.CommandBars
        $properties = $commandBars.FindControl(1, 2578)  # msoControlButton, Project Properties
        if (!$properties -or !$properties.Enabled) { throw 'VBE Project Properties command is unavailable.' }

        # Execute blocks until the modal dialog closes, so operate the controls from a second runspace.
        $commandReturned = [Threading.ManualResetEventSlim]::new()
        $worker = [PowerShell]::Create()
        try {
            [void]$worker.AddScript($setProtection.ToString()).AddArgument([IntPtr]$vbeWindow.HWnd).
                AddArgument($project.Name).AddArgument((Get-ExcelProcessId $xl)).AddArgument($commandReturned)
            $operation = $worker.BeginInvoke()
            try { $properties.Execute() }
            finally {
                $commandReturned.Set()
                [void]$worker.EndInvoke($operation)
            }
        }
        finally {
            $worker.Dispose()
            $commandReturned.Dispose()
        }
    }
    finally {
        # Release the VBE objects before Use-ExcelWorkbook closes the workbook and quits Excel.
        foreach ($reference in $properties, $commandBars, $vbeWindow, $vbe, $project) {
            if ($null -ne $reference) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference)
            }
        }
    }
}

Use-ExcelWorkbook -Path $Path -ScriptBlock {
    param($xl, $wb)
    if ($wb.VBProject.Protection -ne 1) {  # vbext_pp_locked
        throw 'The saved VBA project is not locked. The publish copy was rejected.'
    }
}
