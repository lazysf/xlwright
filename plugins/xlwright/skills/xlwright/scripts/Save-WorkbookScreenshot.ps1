<#
.SYNOPSIS
Saves an image of the Excel window showing a workbook, as a person would see it.

.DESCRIPTION
Opens the workbook in a new Excel instance, by default with macros and events disabled and without updating
links, and closes it without saving. The window is placed outside every monitor before it is shown, so it does
not appear on the screen and does not take the foreground; the clipboard is not used. After the window has
been shown for a second, the sheet is activated, -Range is scrolled to the top left, and the window is drawn
with PrintWindow. The image holds the ribbon, the formula bar, the cells with their formats, conditional
formats, charts, shapes and buttons, the sheet tabs and the status bar, in the Office theme of this user.
A hidden Excel window draws as black, so it is shown (docs/notes/screenshot.md).

With -UserForm, the image is that UserForm instead of the Excel window. It needs -EnableMacros, because the
form's code (UserForm_Initialize) runs. A temporary standard module, which is not saved, loads the form, places
it outside every monitor and shows it modeless (Show without vbModeless would stop the call until the form is
closed); its window (class ThunderDFrame) is drawn with PrintWindow and the form is unloaded
(docs/notes/userform-screenshot.md). A form whose code shows a MsgBox is expected to stop the call until -TimeoutSeconds (not checked).

The result is written as UTF-8 JSON to -OutFile:
- Path, Sheet, VisibleRange (the cells in the window, e.g. "A1:P20"), UserForm, ImageFile, Width and Height
  (pixels).
- Error and Warnings.

Exit code is 0 when the image was saved and 1 otherwise.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # The PNG to write. An existing file is replaced.
    [Parameter(Mandatory)][string]$ImageFile,
    # The sheet to show. Default: the sheet that is active when the workbook opens.
    [string]$Sheet,
    # A cell or range of -Sheet (or of the active sheet) to scroll to the top left, e.g. A100.
    [string]$Range,
    # The window size in pixels. Excel makes a window larger than the screen smaller; the image then has
    # the smaller size and Warnings says so.
    [ValidateRange(400, 8000)][int]$Width = 1280,
    [ValidateRange(300, 8000)][int]$Height = 800,
    # Let macros and events run, e.g. Workbook_Open or a ribbon onLoad that changes what is shown.
    [switch]$EnableMacros,
    # A UserForm of the workbook to show and draw instead of the window. Needs -EnableMacros.
    [string]$UserForm,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 120
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force
Add-Type -AssemblyName System.Drawing
# Add-Type fails if the type already exists, e.g. when the script runs twice in one session.
if (-not ('XlWright.Screenshot' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
namespace XlWright {
    public static class Screenshot {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
        [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hWnd, StringBuilder name, int size);
        // The top-level windows of a process with a window class, e.g. ThunderDFrame for a UserForm.
        public static IntPtr[] FindWindows(uint processId, string className) {
            var found = new List<IntPtr>();
            EnumWindows((hWnd, lParam) => {
                uint owner;
                GetWindowThreadProcessId(hWnd, out owner);
                if (owner == processId) {
                    var name = new StringBuilder(256);
                    GetClassName(hWnd, name, name.Capacity);
                    if (name.ToString() == className) { found.Add(hWnd); }
                }
                return true;
            }, IntPtr.Zero);
            return found.ToArray();
        }
    }
}
'@
}

# Shows a form modeless outside every monitor. With StartUpPosition left at CenterOwner, the form came out on
# the screen although its owner was outside (docs/notes/userform-screenshot.md).
$formHelper = @'
Option Explicit
Private shownForm As Object

Public Sub ShowForm(ByVal formName As String, ByVal leftPoints As Double)
    Set shownForm = VBA.UserForms.Add(formName)
    shownForm.StartUpPosition = 0
    shownForm.Left = leftPoints
    shownForm.Top = 0
    shownForm.Show vbModeless
End Sub

Public Sub CloseForm()
    If Not shownForm Is Nothing Then Unload shownForm
    Set shownForm = Nothing
End Sub
'@ -replace "`r?`n", "`r`n"

$result = [ordered]@{
    Path = $null; Sheet = $null; VisibleRange = $null; UserForm = $null; ImageFile = $null; Width = 0; Height = 0
    Error = $null; Warnings = @()
}
$warnings = @()
try {
    if ($UserForm -and -not $EnableMacros) {
        throw '-UserForm needs -EnableMacros: the code of the form (UserForm_Initialize) runs.'
    }
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $result.Path = $full
    $image = [IO.Path]::GetFullPath($ImageFile)
    if (-not (Test-Path -LiteralPath (Split-Path -Parent $image) -PathType Container)) {
        throw "The folder of -ImageFile does not exist: $image"
    }

    $graphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    try { $dpi = $graphics.DpiX } finally { $graphics.Dispose() }
    if ($dpi -ne 96) {
        $warnings += "Display scaling is not 100 % ($dpi dpi). The image size and look were checked only at 100 %."
    }
    $pointsPerPixel = 72.0 / $dpi
    # Left of the leftmost monitor (SM_XVIRTUALSCREEN), so the window is on no monitor.
    $left = [XlWright.Screenshot]::GetSystemMetrics(76) - $Width - 100

    $shot = Use-ExcelWorkbook -Path $full -EnableMacros:$EnableMacros -EnableEvents:$EnableMacros -TimeoutSeconds $TimeoutSeconds `
        -WarningVariable +warnings -ScriptBlock {
        param($xl, $wb)
        $target = $wb.ActiveSheet
        if ($Sheet) {
            $target = $null
            foreach ($item in $wb.Sheets) { if ($item.Name -eq $Sheet) { $target = $item } }
            if ($null -eq $target) {
                $names = @(foreach ($item in $wb.Sheets) { $item.Name })
                throw ('Not a sheet of the workbook: ' + $Sheet + '. Sheets: ' + ($names -join ', ') + '.')
            }
            if ([int]$target.Visible -ne -1) { throw "The sheet is hidden: $Sheet" }  # -1 is xlSheetVisible
        }
        $cell = $null
        if ($Range) {
            try { $cell = $target.Range($Range) }
            catch { throw ('Not a range of the sheet ' + $target.Name + ': ' + $Range) }
        }

        # Placed and sized while hidden, then shown: the window never appears on a monitor.
        $xl.WindowState = -4143  # xlNormal
        $xl.Width = $Width * $pointsPerPixel
        $xl.Height = $Height * $pointsPerPixel
        $xl.Left = $left * $pointsPerPixel
        $xl.Top = 0
        $xl.Visible = $true
        # 200 ms after showing, the sheet tabs were not drawn yet; after 1 s they were (docs/notes/screenshot.md).
        Start-Sleep -Milliseconds 1000
        $target.Activate()
        if ($cell) { $xl.Goto($cell, $true) }

        $hwnd = [IntPtr]$xl.Hwnd
        $helper = $null
        if ($UserForm) {
            $form = $null
            foreach ($component in $wb.VBProject.VBComponents) { if ($component.Name -eq $UserForm) { $form = $component } }
            if ($null -eq $form -or [int]$form.Type -ne 3) { throw "Not a UserForm of the workbook: $UserForm" }
            # The workbook is closed without saving, so the module does not stay in it.
            $helper = $wb.VBProject.VBComponents.Add(1)
            $helper.CodeModule.AddFromString($formHelper)
            $before = @([XlWright.Screenshot]::FindWindows((Get-ExcelProcessId $xl), 'ThunderDFrame')).Count
            # tradeoff: 9000 px left of every monitor; a form wider than that would reach a monitor.
            $formLeft = ([XlWright.Screenshot]::GetSystemMetrics(76) - 9000) * $pointsPerPixel
            [void]$xl.Run(($helper.Name + '.ShowForm'), $UserForm, $formLeft)
            Start-Sleep -Milliseconds 500
            $forms = @([XlWright.Screenshot]::FindWindows((Get-ExcelProcessId $xl), 'ThunderDFrame'))
            if ($before -ne 0 -or $forms.Count -ne 1) {
                throw ('Expected one UserForm window after showing ' + $UserForm + ', found ' + $forms.Count + '.')
            }
            $hwnd = $forms[0]
        }
        $rect = New-Object XlWright.Screenshot+RECT
        if (-not [XlWright.Screenshot]::GetWindowRect($hwnd, [ref]$rect)) { throw 'GetWindowRect failed.' }
        $bitmap = New-Object System.Drawing.Bitmap(($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top))
        try {
            $g = [System.Drawing.Graphics]::FromImage($bitmap)
            $hdc = $g.GetHdc()
            # PW_RENDERFULLCONTENT (2) draws the window itself, not what is on the screen in front of it.
            try { $drawn = [XlWright.Screenshot]::PrintWindow($hwnd, $hdc, 2) }
            finally { $g.ReleaseHdc($hdc); $g.Dispose() }
            if (-not $drawn) { throw 'PrintWindow failed.' }
            $bitmap.Save($image, [System.Drawing.Imaging.ImageFormat]::Png)
            $size = @($bitmap.Width, $bitmap.Height)
        }
        finally { $bitmap.Dispose() }
        if ($helper) { [void]$xl.Run(($helper.Name + '.CloseForm')) }
        $visible = $null
        if ($null -ne $xl.ActiveWindow.VisibleRange) { $visible = $xl.ActiveWindow.VisibleRange.Address($false, $false) }
        $xl.Visible = $false
        $shownForm = if ($UserForm) { $UserForm } else { $null }
        [ordered]@{ Sheet = [string]$target.Name; VisibleRange = $visible; UserForm = $shownForm; ImageFile = $image; Width = $size[0]; Height = $size[1] }
    }
    foreach ($key in @($shot.Keys)) { $result[$key] = $shot[$key] }
    # Excel makes a window larger than the screen smaller when Width and Height are set, while still hidden
    # (docs/notes/screenshot.md). The image is still usable, so this is a warning.
    if (-not $UserForm -and ($result.Width -ne $Width -or $result.Height -ne $Height)) {
        $message = "The image is $($result.Width) x $($result.Height) px, not the requested $Width x $Height."
        $screenWidth = [XlWright.Screenshot]::GetSystemMetrics(0)   # SM_CXSCREEN
        $screenHeight = [XlWright.Screenshot]::GetSystemMetrics(1)  # SM_CYSCREEN
        # Only at 100 %: with other scaling these metrics may be scaled for this process while the image is in
        # pixels of the screen, so the comparison would not hold; the warning about the scaling stands instead.
        if ($dpi -eq 96 -and ($Width -gt $screenWidth -or $Height -gt $screenHeight)) {
            $message += " Excel makes a window larger than the screen ($screenWidth x $screenHeight px) smaller."
        }
        $warnings += $message
    }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $result.Warnings = @($warnings | ForEach-Object { [string]$_ })
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 4
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
