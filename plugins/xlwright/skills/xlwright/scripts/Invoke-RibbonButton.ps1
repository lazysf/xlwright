<#
.SYNOPSIS
Clicks a custom ribbon button through UI Automation and reports the dialogs it shows.

.DESCRIPTION
Opens the workbook in a visible, maximized Excel, selects the tab, invokes the button, and handles the dialogs
(#32770, e.g. MsgBox) that appear afterwards: the names of each dialog's elements are recorded and the
dialog is closed with WindowPattern.Close (the same as the x button). Only this check has been done with
a visible window; whether the ribbon is built for a hidden one has not been checked. The workbook is
closed without saving. Workbook_Open does not run (see XlWright.psm1), the button's own macro does.
The tab and the button must be uniquely named in their respective search scopes.

In a MsgBox both the message and the buttons are exposed as Pane elements, so they cannot be told
apart by control type. Each entry of Dialogs lists them all, e.g. ["OK", "<message>"].

If a dialog cannot be closed (a Yes/No MsgBox has no Cancel and ignores the x button), or the click
call does not return in time, the Excel instance started by this script is stopped by its PID and the
result records the error. Other Excel processes are never touched.

The click call returns without waiting for the macro to finish. A macro that keeps running without a
dialog (e.g. Application.Wait) is waited for when the workbook is closed; a macro that never ends
makes this script hang there.

The result is written as UTF-8 JSON to -OutFile (stdout of powershell.exe garbles Japanese in Git Bash).
Exit code is non-zero when the click failed.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$TabLabel,
    [Parameter(Mandatory)][string]$ButtonLabel,
    [Parameter(Mandatory)][string]$OutFile,
    # How long to wait for the tab / button to appear.
    [int]$FindTimeoutSeconds = 15,
    # How long to wait for each dialog. No dialog within this time is not an error.
    [int]$DialogTimeoutSeconds = 5,
    # How long to wait for the click call to return after the last dialog.
    [int]$ActionTimeoutSeconds = 10,
    # Safety limit on the number of dialogs handled for one click.
    [int]$MaxDialogs = 5
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Ribbon.psm1') -Force

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$ControlType = [System.Windows.Automation.ControlType]

function Find-Element($Root, $Condition, [double]$TimeoutSeconds) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $el = $Root.FindFirst($Scope::Descendants, $Condition)
        if ($el) { return $el }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    $null
}

function Get-DialogNames($Dialog) {
    $names = @()
    foreach ($e in $Dialog.FindAll($Scope::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        if ($e.Current.Name) { $names += $e.Current.Name }
    }
    , $names
}

$result = [ordered]@{ TabFound = $false; ButtonFound = $false; Clicked = $false; Dialogs = @(); Error = $null }
try {
    Use-ExcelWorkbook -Path $Path -Visible -EnableMacros -ScriptBlock {
        param($xl, $wb)
        $excelPid = Get-ExcelProcessId $xl
        $xl.WindowState = -4137  # xlMaximized: in a narrow window Excel collapses groups and hides their buttons
        $win = $AE::FromHandle([IntPtr]$xl.Hwnd)

        $tab = Find-RibbonTab $win $TabLabel $FindTimeoutSeconds
        $result.TabFound = $true
        $tabRoot = [System.Windows.Automation.TreeWalker]::RawViewWalker.GetParent($tab)
        if (!$tabRoot -or $tabRoot.Current.ControlType -ne $ControlType::Tab) {
            throw "Ribbon tab container not found: $TabLabel"
        }
        $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()

        # Excel exposes the selected tab's controls in a group named after the tab, not under the tab item.
        # Depending on the Excel build the group is under the tab container or elsewhere in the window,
        # so search the whole window.
        $tabContent = Find-TabContent $win (New-NamedControlCondition $TabLabel $ControlType::Group) $FindTimeoutSeconds "tab content '$TabLabel'"
        $btn = Find-UniqueElement $tabContent (New-NamedControlCondition $ButtonLabel $ControlType::Button) $FindTimeoutSeconds "button '$ButtonLabel' in '$TabLabel'"
        $result.ButtonFound = $true

        # Invoke does not return while the macro shows a modal dialog, so click from another runspace.
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript({ param($b)
            $ErrorActionPreference = 'Stop'
            $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }).AddArgument($btn)
        $async = $ps.BeginInvoke()

        $stuck = $null
        $dlgCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, '#32770')
        while ($result.Dialogs.Count -lt $MaxDialogs) {
            # A MsgBox shows up as a descendant of the Excel window in the UIA tree.
            $dlg = Find-Element $win $dlgCond $DialogTimeoutSeconds
            if (-not $dlg) { break }
            $names = Get-DialogNames $dlg
            $result.Dialogs += , ([ordered]@{ Title = $dlg.Current.Name; Names = $names })
            $handle = $dlg.Current.NativeWindowHandle
            try { $dlg.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { }
            # The next MsgBox may reuse the same window, so "closed" means the same window with the
            # same contents is no longer there.
            $closed = $false
            for ($i = 0; $i -lt 10 -and -not $closed; $i++) {
                Start-Sleep -Milliseconds 300
                $now = $win.FindFirst($Scope::Descendants, $dlgCond)
                $closed = (-not $now) -or ($now.Current.NativeWindowHandle -ne $handle) -or
                    ((Get-DialogNames $now) -join "`n") -ne ($names -join "`n")
            }
            if (-not $closed) { $stuck = 'A dialog could not be closed'; break }
        }
        if (-not $stuck -and -not $async.AsyncWaitHandle.WaitOne($ActionTimeoutSeconds * 1000)) {
            $stuck = 'The click call did not return in time'
        }
        if ($stuck) {
            # The instance is blocked in a modal state; COM calls to it would hang. Stop only this instance.
            # The runspace is left undisposed on purpose: Dispose waits for the blocked Invoke.
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            throw "$stuck. Stopped the Excel instance of this script (PID $excelPid)."
        }
        try {
            [void]$ps.EndInvoke($async)
            if ($ps.HadErrors) { throw ($ps.Streams.Error | Select-Object -First 1).Exception }
        }
        finally { $ps.Dispose() }
        $result.Clicked = $true
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
