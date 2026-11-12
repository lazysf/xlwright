<#
.SYNOPSIS
Checks that the custom ribbon shown by Excel matches the customUI XML embedded in the workbook.

.DESCRIPTION
Reads customUI/customUI14.xml from the workbook, opens the workbook in a visible, maximized Excel
(macros enabled, workbook events disabled), selects each tab of the XML, and compares what UI Automation
exposes with the XML:

- tabs: shown or not (visible), position (insertAfterMso / insertBeforeMso), and with
  startFromScratch="true" that no other built-in tab is shown
- groups: label and order. A built-in group (idMso) in a custom tab is checked by position only,
  because its label depends on the Office language.
- button and toggleButton: label, order, size (large / normal), enabled, and showLabel for normal
  buttons with an image

Values set by callbacks (getLabel, getVisible, getEnabled, getSize, getShowLabel) are not compared;
the values Excel shows are listed in Dynamic. Callbacks run with workbook events disabled, so a value
that depends on state made in Workbook_Open can differ from normal use.

Not compared and listed in Unsupported: groups that hold other controls (menu, splitButton, gallery,
box, ...), tabs whose label comes from a callback, hidden built-in groups, insertAfterMso /
insertBeforeMso of groups and controls, and contextual tabs. Never compared: tooltips (UI Automation
does not expose screentip / supertip) and icons. With -ScreenshotDir a PNG of each tab is written, so
that the icons can be checked by eye; a tab whose image fails has File = null and the reason in Error.
Separators are not exposed by UI Automation and are ignored.

A group that Excel collapsed into one button for lack of width is listed in Collapsed and its controls
are not compared. The window is maximized to avoid this.

The result is written as UTF-8 JSON to -OutFile.
Exit code: 0 everything matched, 1 a mismatch was found, 2 the check could not be done (Error).
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # Write tab-1.png, tab-2.png, ... (in the order of the XML) to this existing directory.
    [string]$ScreenshotDir,
    # How long to wait for the ribbon to appear.
    [int]$FindTimeoutSeconds = 15,
    # For the tests: a normal window this many points wide instead of a maximized one.
    [int]$WindowWidth = 0
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem, System.Drawing
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'Ribbon.psm1') -Force
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace XlWright {
    public static class WindowCapture {
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
    }
}
'@

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$ControlType = [System.Windows.Automation.ControlType]
$Walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
$ns = 'http://schemas.microsoft.com/office/2009/07/customui'
$supportedControls = @('button', 'toggleButton')

$result = [ordered]@{
    Expected = @(); Actual = @(); Mismatches = @(); Collapsed = @(); Dynamic = @(); Unsupported = @()
    Screenshots = @(); Error = $null
}

function Add-Mismatch([string]$At, [string]$Item, $Expected, $Actual) {
    $result.Mismatches += , ([ordered]@{ Path = $At; Item = $Item; Expected = $Expected; Actual = $Actual })
}
function Add-Dynamic([string]$At, [string]$Item, $Actual) {
    $result.Dynamic += , ([ordered]@{ Path = $At; Item = $Item; Actual = $Actual })
}
function Add-Unsupported([string]$At, [string]$Reason) {
    $result.Unsupported += , ([ordered]@{ Path = $At; Reason = $Reason })
}

# --- Expected values from the XML ------------------------------------------------------------

function Get-Flag($Node, [string]$Name, [string]$Callback, [string]$Default) {
    if ($Node.HasAttribute($Callback)) { return 'dynamic' }
    if (!$Node.HasAttribute($Name)) { return $Default }
    $value = $Node.GetAttribute($Name)
    if ($value -eq '1') { return 'true' }
    if ($value -eq '0') { return 'false' }
    $value
}

function Get-Label($Node) {
    if ($Node.HasAttribute('getLabel')) { return $null }
    $Node.GetAttribute('label')
}

# insertAfterMso / insertBeforeMso of a group or control, e.g. 'insertAfterMso=GroupClipboard'.
function Get-InsertMso($Node) {
    foreach ($a in 'insertAfterMso', 'insertBeforeMso') {
        if ($Node.HasAttribute($a)) { return $a + '=' + $Node.GetAttribute($a) }
    }
    $null
}

function Read-CustomUI([string]$Workbook) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Workbook)
    try {
        $entry = $zip.GetEntry('customUI/customUI14.xml')
        if (!$entry) { throw 'The workbook has no customUI/customUI14.xml' }
        $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    finally { $zip.Dispose() }
    $doc = New-Object Xml.XmlDocument
    $doc.LoadXml($text)
    if ($doc.DocumentElement.NamespaceURI -ne $ns) { throw "customUI14.xml must use the namespace $ns" }
    $doc
}

function ConvertTo-ExpectedControl($Node) {
    [ordered]@{
        Type = $Node.LocalName; Id = $Node.GetAttribute('id'); Label = (Get-Label $Node)
        Visible = (Get-Flag $Node 'visible' 'getVisible' 'true')
        Size = (Get-Flag $Node 'size' 'getSize' 'normal')
        Enabled = (Get-Flag $Node 'enabled' 'getEnabled' 'true')
        ShowLabel = (Get-Flag $Node 'showLabel' 'getShowLabel' 'true')
        HasImage = ($Node.HasAttribute('imageMso') -or $Node.HasAttribute('image') -or $Node.HasAttribute('getImage'))
        Insert = (Get-InsertMso $Node)
    }
}

function ConvertTo-ExpectedGroup($Node) {
    $controls = @()
    $others = @()
    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -ne 'Element') { continue }
        if ($supportedControls -contains $child.LocalName) { $controls += , (ConvertTo-ExpectedControl $child) }
        elseif ($child.LocalName -ne 'separator') { $others += $child.LocalName }
    }
    [ordered]@{
        BuiltIn = $Node.HasAttribute('idMso'); Id = $(if ($Node.HasAttribute('idMso')) { $Node.GetAttribute('idMso') } else { $Node.GetAttribute('id') })
        Label = $(if ($Node.HasAttribute('idMso')) { $null } else { Get-Label $Node })
        Visible = (Get-Flag $Node 'visible' 'getVisible' 'true')
        Insert = (Get-InsertMso $Node)
        Controls = $controls; OtherControls = $others
    }
}

function ConvertTo-ExpectedTab($Node) {
    $groups = @()
    foreach ($g in $Node.ChildNodes) {
        if ($g.NodeType -eq 'Element' -and $g.LocalName -eq 'group') { $groups += , (ConvertTo-ExpectedGroup $g) }
    }
    [ordered]@{
        BuiltIn = $Node.HasAttribute('idMso'); Id = $(if ($Node.HasAttribute('idMso')) { $Node.GetAttribute('idMso') } else { $Node.GetAttribute('id') })
        Label = $(if ($Node.HasAttribute('idMso')) { $null } else { Get-Label $Node })
        DynamicLabel = $Node.HasAttribute('getLabel')
        Visible = (Get-Flag $Node 'visible' 'getVisible' 'true')
        InsertAfterMso = $Node.GetAttribute('insertAfterMso'); InsertBeforeMso = $Node.GetAttribute('insertBeforeMso')
        Groups = $groups
    }
}

function Get-Name($Item) {
    if ($Item.Label) { return $Item.Label }
    $Item.Id
}

# --- Actual values from UI Automation ---------------------------------------------------------

function Get-Children($Element) {
    $list = @()
    $child = $Walker.GetFirstChild($Element)
    while ($child) { $list += , $child; $child = $Walker.GetNextSibling($child) }
    , $list
}

function Read-ActualGroup($Element) {
    $c = $Element.Current
    $height = $c.BoundingRectangle.Height
    $children = Get-Children $Element
    $collapsed = Test-CollapsedGroup $Element
    $controls = @(foreach ($ch in $children) {
        $cc = $ch.Current
        $r = $cc.BoundingRectangle
        [ordered]@{
            Type = $cc.ControlType.ProgrammaticName -replace '^ControlType\.', ''; Name = $cc.Name
            # A large button spans most of the group height (76 of 97 px); a normal one is a third of it.
            Size = $(if ($r.Height -gt $height / 2) { 'large' } else { 'normal' })
            Enabled = $cc.IsEnabled
            # A normal button without its label is a square icon; with a label it is wider.
            IconOnly = ($r.Width -le $r.Height + 4)
            Width = [int]$r.Width; Height = [int]$r.Height
        }
    })
    [ordered]@{ Name = $c.Name; Collapsed = $collapsed; Controls = $controls }
}

function Save-TabImage($Xl, $Content, [string]$File) {
    $hwnd = [IntPtr]$Xl.Hwnd
    $wr = New-Object XlWright.WindowCapture+RECT
    if (![XlWright.WindowCapture]::GetWindowRect($hwnd, [ref]$wr)) { throw 'GetWindowRect failed' }
    $bmp = New-Object System.Drawing.Bitmap(($wr.Right - $wr.Left), ($wr.Bottom - $wr.Top))
    try {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        # PW_RENDERFULLCONTENT (2) draws the window itself, not what is on the screen in front of it.
        try { $ok = [XlWright.WindowCapture]::PrintWindow($hwnd, $hdc, 2) }
        finally { $g.ReleaseHdc($hdc); $g.Dispose() }
        if (!$ok) { throw 'PrintWindow failed' }
        $r = $Content.Current.BoundingRectangle
        $area = New-Object System.Drawing.Rectangle(([int]$r.X - $wr.Left), ([int]$r.Y - $wr.Top), [int]$r.Width, [int]$r.Height)
        # Checked only at 100 % display scaling. With other scaling the window rectangle of this (not DPI
        # aware) process and the UI Automation rectangle may disagree.
        if ($area.X -lt 0 -or $area.Y -lt 0 -or $area.Right -gt $bmp.Width -or $area.Bottom -gt $bmp.Height) {
            throw "The tab area $area is outside the window image ($($bmp.Width) x $($bmp.Height)). Is display scaling not 100 %?"
        }
        $crop = $bmp.Clone($area, $bmp.PixelFormat)
        try { $crop.Save($File, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $crop.Dispose() }
    }
    finally { $bmp.Dispose() }
}

# --- Comparison ------------------------------------------------------------------------------

# UI Automation names a group without a label " " and a button without a label "".
function Test-SameName([string]$Actual, [string]$Expected) {
    $Actual.Trim() -ceq $Expected.Trim()
}

# Reports an actual item that no expected item took. One that the XML hides is reported as such.
function Add-Extra($Item, [string]$At, [string]$Kind, $Hidden, [bool]$AllowExtra) {
    if (@($Hidden | Where-Object { Test-SameName $Item.Name $_ }).Count -gt 0) {
        Add-Mismatch "$At/$($Item.Name)" $Kind 'hidden' 'shown'
    }
    elseif (-not $AllowExtra) { Add-Mismatch "$At/$($Item.Name)" $Kind 'not in the XML' 'shown' }
}

# Pairs expected items (in XML order) with actual items (in screen order). An expected item without a
# label (built-in group, getLabel) takes the next actual item, unless that item carries the label of a
# later expected item: then the item without a label is absent. Items whose visibility comes from a
# callback may be absent. $Hidden holds the labels of items the XML hides. With -AllowExtra, actual items
# that are not in the XML (the built-in groups of a built-in tab) are not reported.
function Get-Pairs($Expected, $Actual, [string]$At, [string]$Kind, $Hidden = @(), [switch]$AllowExtra) {
    $pairs = @()
    $p = 0
    for ($k = 0; $k -lt $Expected.Count; $k++) {
        $e = $Expected[$k]
        $name = Get-Name $e
        if ($null -eq $e.Label) {
            $later = @($Expected | Select-Object -Skip ($k + 1) | Where-Object { $null -ne $_.Label } | ForEach-Object { $_.Label })
            $takenByLater = $p -lt $Actual.Count -and @($later | Where-Object { Test-SameName $Actual[$p].Name $_ }).Count -gt 0
            if ($p -lt $Actual.Count -and -not $takenByLater) { $pairs += , @($e, $Actual[$p]); $p++ }
            elseif ($e.Visible -eq 'dynamic') { Add-Dynamic "$At/$name" 'Visible' $false }
            else { Add-Mismatch "$At/$name" $Kind 'shown' 'missing' }
            continue
        }
        $q = $p
        while ($q -lt $Actual.Count -and -not (Test-SameName $Actual[$q].Name $e.Label)) { $q++ }
        if ($q -lt $Actual.Count) {
            for ($i = $p; $i -lt $q; $i++) { Add-Extra $Actual[$i] $At $Kind $Hidden $AllowExtra.IsPresent }
            $pairs += , @($e, $Actual[$q])
            $p = $q + 1
        }
        elseif ($e.Visible -eq 'dynamic') { Add-Dynamic "$At/$name" 'Visible' $false }
        else {
            $elsewhere = @($Actual | Where-Object { Test-SameName $_.Name $e.Label }).Count -gt 0
            Add-Mismatch "$At/$name" $Kind 'shown' $(if ($elsewhere) { 'out of order' } else { 'missing' })
        }
    }
    for ($i = $p; $i -lt $Actual.Count; $i++) { Add-Extra $Actual[$i] $At $Kind $Hidden $AllowExtra.IsPresent }
    , $pairs
}

# Labels used only by hidden items. A label that a shown item also uses does not tell which one appeared.
function Get-HiddenLabels($Items) {
    $shown = @($Items | Where-Object { $_.Visible -ne 'false' -and $null -ne $_.Label } | ForEach-Object { $_.Label })
    , @($Items | Where-Object { $_.Visible -eq 'false' -and $null -ne $_.Label -and $shown -cnotcontains $_.Label } |
        ForEach-Object { $_.Label })
}

# Positions relative to built-in groups or controls cannot be checked: their labels depend on the
# Office language and UI Automation gives groups no id.
function Add-InsertUnsupported($Items, [string]$At) {
    foreach ($i in @($Items | Where-Object { $_.Insert })) {
        Add-Unsupported "$At/$(Get-Name $i)" "position $($i.Insert) not compared"
    }
}

function Compare-Control($E, $A, [string]$At) {
    if ($A.Type -ne 'Button') { Add-Mismatch $At 'Type' $E.Type $A.Type; return }
    if ($null -eq $E.Label) { Add-Dynamic $At 'Label' $A.Name }
    if ($E.Visible -eq 'dynamic') { Add-Dynamic $At 'Visible' $true }
    if ($E.Size -eq 'dynamic') { Add-Dynamic $At 'Size' $A.Size }
    elseif ($E.Size -ne $A.Size) { Add-Mismatch $At 'Size' $E.Size $A.Size }
    if ($E.Enabled -eq 'dynamic') { Add-Dynamic $At 'Enabled' $A.Enabled }
    elseif (($E.Enabled -eq 'true') -ne $A.Enabled) { Add-Mismatch $At 'Enabled' ($E.Enabled -eq 'true') $A.Enabled }
    # Without a label there is nothing to show beside the icon, whatever showLabel says.
    if ($E.HasImage -and $A.Size -eq 'normal' -and $E.Label -ne '') {
        if ($E.ShowLabel -eq 'dynamic') { Add-Dynamic $At 'ShowLabel' (-not $A.IconOnly) }
        elseif (($E.ShowLabel -eq 'true') -eq $A.IconOnly) { Add-Mismatch $At 'ShowLabel' ($E.ShowLabel -eq 'true') (-not $A.IconOnly) }
    }
}

function Compare-Group($E, $A, [string]$At) {
    Add-InsertUnsupported $E.Controls $At
    if ($E.Visible -eq 'dynamic') { Add-Dynamic $At 'Visible' $true }
    if ($null -eq $E.Label -and -not $E.BuiltIn) { Add-Dynamic $At 'Label' $A.Name }
    if ($E.BuiltIn) { return }
    if ($A.Collapsed) { $result.Collapsed += $At; return }
    if ($E.OtherControls.Count -gt 0) {
        Add-Unsupported $At ('controls not compared because the group holds ' + (($E.OtherControls | Sort-Object -Unique) -join ', '))
        return
    }
    $expected = @($E.Controls | Where-Object { $_.Visible -ne 'false' })
    foreach ($pair in (Get-Pairs $expected $A.Controls $At 'Control' (Get-HiddenLabels $E.Controls))) {
        Compare-Control $pair[0] $pair[1] "$At/$(Get-Name $pair[0])"
    }
}

function Compare-Tab($E, $Content, [string]$At) {
    $actualGroups = @(foreach ($g in (Get-Children $Content)) {
        if ($g.Current.ControlType -eq $ControlType::Group) { Read-ActualGroup $g }
    })
    Add-InsertUnsupported $E.Groups $At
    # A built-in group has no label in the XML and no id in UI Automation, so a hidden one cannot be found.
    foreach ($g in @($E.Groups | Where-Object { $_.Visible -eq 'false' -and $null -eq $_.Label })) {
        Add-Unsupported "$At/$(Get-Name $g)" 'hidden group without a fixed label: not compared'
    }
    $expected = @($E.Groups | Where-Object { $_.Visible -ne 'false' })
    $hidden = Get-HiddenLabels $E.Groups
    if ($E.BuiltIn) {
        # The built-in groups of the tab are not in the XML: only the custom groups and their order are checked.
        foreach ($g in @($expected | Where-Object { $_.BuiltIn -or $null -eq $_.Label })) {
            Add-Unsupported "$At/$(Get-Name $g)" 'group in a built-in tab without a fixed label: not compared'
        }
        $custom = @($expected | Where-Object { -not $_.BuiltIn -and $null -ne $_.Label })
        $pairs = Get-Pairs $custom $actualGroups $At 'Group' $hidden -AllowExtra
    }
    else { $pairs = Get-Pairs $expected $actualGroups $At 'Group' $hidden }
    foreach ($pair in $pairs) { Compare-Group $pair[0] $pair[1] "$At/$(Get-Name $pair[0])" }
    , $actualGroups
}

# --- Main ------------------------------------------------------------------------------------

try {
    if ($ScreenshotDir -and !(Test-Path -LiteralPath $ScreenshotDir -PathType Container)) {
        throw "ScreenshotDir does not exist: $ScreenshotDir"
    }
    $doc = Read-CustomUI $Path
    $nsm = New-Object Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('c', $ns)
    $ribbon = $doc.SelectSingleNode('/c:customUI/c:ribbon', $nsm)
    if (!$ribbon) { throw 'customUI14.xml has no ribbon element' }
    $fromScratch = @('true', '1') -contains $ribbon.GetAttribute('startFromScratch')
    if ($doc.SelectSingleNode('/c:customUI/c:ribbon/c:contextualTabs', $nsm)) { Add-Unsupported 'contextualTabs' 'not compared' }
    $expectedTabs = @(foreach ($t in $doc.SelectNodes('/c:customUI/c:ribbon/c:tabs/c:tab', $nsm)) { ConvertTo-ExpectedTab $t })
    $result.Expected = [ordered]@{ StartFromScratch = $fromScratch; Tabs = $expectedTabs }
    if ($expectedTabs.Count -eq 0) { throw 'customUI14.xml has no tab' }

    Use-ExcelWorkbook -Path $Path -Visible -EnableMacros -ScriptBlock {
        param($xl, $wb)
        Set-RibbonWindow $xl $WindowWidth
        $win = $AE::FromHandle([IntPtr]$xl.Hwnd)

        # Wait until the ribbon has the tabs of the XML (or the timeout, when the XML was rejected).
        $deadline = (Get-Date).AddSeconds($FindTimeoutSeconds)
        do {
            $tabs = Get-RibbonTabs $win
            $missing = @($expectedTabs | Where-Object {
                $e = $_
                $e.Visible -eq 'true' -and -not $e.DynamicLabel -and @($tabs | Where-Object {
                    ($e.BuiltIn -and $_.Current.AutomationId -eq $e.Id) -or (-not $e.BuiltIn -and $_.Current.Name -ceq $e.Label)
                }).Count -eq 0
            })
            if ($missing.Count -eq 0) { break }
            Start-Sleep -Milliseconds 300
        } while ((Get-Date) -lt $deadline)
        if ($tabs.Count -eq 0) { throw 'No ribbon tab found. Is the ribbon collapsed or hidden?' }

        if ($fromScratch) {
            $allowed = @($expectedTabs | Where-Object { $_.BuiltIn } | ForEach-Object { $_.Id })
            foreach ($t in $tabs) {
                $aid = $t.Current.AutomationId
                if ($aid -and $allowed -notcontains $aid) { Add-Mismatch $t.Current.Name 'Tab' 'hidden by startFromScratch' 'shown' }
            }
        }

        $index = 0
        foreach ($e in $expectedTabs) {
            $index++
            $at = Get-Name $e
            if ($e.DynamicLabel) { Add-Unsupported $at 'tab label from getLabel: the tab cannot be identified'; continue }
            $found = @($tabs | Where-Object {
                ($e.BuiltIn -and $_.Current.AutomationId -eq $e.Id) -or (-not $e.BuiltIn -and $_.Current.Name -ceq $e.Label)
            })
            if ($e.Visible -eq 'false') {
                if ($found.Count -gt 0) { Add-Mismatch $at 'Tab' 'hidden' 'shown' }
                continue
            }
            if ($found.Count -eq 0) {
                if ($e.Visible -eq 'dynamic') { Add-Dynamic $at 'Visible' $false } else { Add-Mismatch $at 'Tab' 'shown' 'missing' }
                continue
            }
            if ($found.Count -gt 1) { Add-Mismatch $at 'Tab' 'one tab' "$($found.Count) tabs with this label"; continue }
            $tab = $found[0]
            if ($e.Visible -eq 'dynamic') { Add-Dynamic $at 'Visible' $true }
            $pos = [array]::IndexOf($tabs, $tab)
            # Custom tabs have no AutomationId. Several of them can follow the same built-in tab, so the
            # position is checked against the nearest built-in tab on each side.
            $before = $null
            for ($i = $pos - 1; $i -ge 0 -and -not $before; $i--) { $before = $tabs[$i].Current.AutomationId }
            $after = $null
            for ($i = $pos + 1; $i -lt $tabs.Count -and -not $after; $i++) { $after = $tabs[$i].Current.AutomationId }
            if (-not $before) { $before = $null }
            if (-not $after) { $after = $null }
            if ($e.InsertAfterMso -and $before -ne $e.InsertAfterMso) { Add-Mismatch $at 'InsertAfterMso' $e.InsertAfterMso $before }
            if ($e.InsertBeforeMso -and $after -ne $e.InsertBeforeMso) { Add-Mismatch $at 'InsertBeforeMso' $e.InsertBeforeMso $after }

            $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
            $name = $tab.Current.Name
            $content = Find-TabContent $win (New-NamedControlCondition $name $ControlType::Group) $FindTimeoutSeconds "tab content '$name'"
            # The classic ribbon is about 100 px high; the single-line (simplified) ribbon is far lower.
            if ($content.Current.BoundingRectangle.Height -lt 60) {
                throw 'The ribbon looks simplified (single line). Ask the user to switch to the classic ribbon.'
            }
            if ($ScreenshotDir) {
                $file = Join-Path $ScreenshotDir "tab-$index.png"
                # A failed image does not stop the comparison; the entry carries the reason instead.
                try { Save-TabImage $xl $content $file; $shotError = $null }
                catch { $file = $null; $shotError = $_.Exception.Message }
                $result.Screenshots += , ([ordered]@{ Tab = $at; File = $file; Error = $shotError })
            }
            $groups = Compare-Tab $e $content $at
            $result.Actual += , ([ordered]@{
                Tab = $at; Name = $name; Position = $pos + 1; Before = $before; After = $after; Groups = $groups
            })
        }
    }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error -ErrorAction Continue; exit 2 }
if ($result.Mismatches.Count -gt 0) { exit 1 }
exit 0
