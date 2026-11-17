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

The result is written as UTF-8 JSON to -OutFile:
- Path, Sheet, VisibleRange (the cells in the window, e.g. "A1:P20"), ImageFile, Width and Height (pixels).
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
    # The window size in pixels.
    [ValidateRange(400, 8000)][int]$Width = 1280,
    [ValidateRange(300, 8000)][int]$Height = 800,
    # Let macros and events run, e.g. Workbook_Open or a ribbon onLoad that changes what is shown.
    [switch]$EnableMacros,
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
using System.Runtime.InteropServices;
namespace XlWright {
    public static class Screenshot {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
        [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
    }
}
'@
}

$result = [ordered]@{
    Path = $null; Sheet = $null; VisibleRange = $null; ImageFile = $null; Width = 0; Height = 0; Error = $null; Warnings = @()
}
$warnings = @()
try {
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
        $visible = $null
        if ($null -ne $xl.ActiveWindow.VisibleRange) { $visible = $xl.ActiveWindow.VisibleRange.Address($false, $false) }
        $xl.Visible = $false
        [ordered]@{ Sheet = [string]$target.Name; VisibleRange = $visible; ImageFile = $image; Width = $size[0]; Height = $size[1] }
    }
    foreach ($key in @($shot.Keys)) { $result[$key] = $shot[$key] }
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
