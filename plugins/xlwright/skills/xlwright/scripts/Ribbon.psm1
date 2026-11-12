<#
.SYNOPSIS
UI Automation helpers shared by the ribbon scripts (Invoke-RibbonButton.ps1, Test-RibbonLayout.ps1).
#>
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$ControlType = [System.Windows.Automation.ControlType]
$Walker = [System.Windows.Automation.TreeWalker]::RawViewWalker

# Maximizes the window: a narrow window makes Excel collapse groups and hide their buttons.
# A width (in points) gives a normal window of that width instead; the tests use it to make groups collapse.
function Set-RibbonWindow($Xl, [int]$Width) {
    if ($Width -gt 0) {
        $Xl.WindowState = -4143  # xlNormal
        $Xl.Width = $Width
    }
    else { $Xl.WindowState = -4137 }  # xlMaximized
}

# A group that Excel collapsed for lack of width is one drop-down (MenuItem) carrying the group label
# and filling the whole group (97 of 97 px). A menu with the group label is at most a large button (76 px).
function Test-CollapsedGroup($Group) {
    $first = $Walker.GetFirstChild($Group)
    if (!$first -or $Walker.GetNextSibling($first)) { return $false }
    $c = $first.Current
    $c.ControlType -eq $ControlType::MenuItem -and $c.Name -eq $Group.Current.Name -and
        $c.BoundingRectangle.Height -ge $Group.Current.BoundingRectangle.Height - 4
}

function Get-CollapsedGroups($TabContent) {
    $names = @()
    $g = $Walker.GetFirstChild($TabContent)
    while ($g) {
        if ($g.Current.ControlType -eq $ControlType::Group -and (Test-CollapsedGroup $g)) { $names += $g.Current.Name }
        $g = $Walker.GetNextSibling($g)
    }
    , $names
}

# The ribbon tabs in screen order. The sheet tabs are TabItems too; the ribbon tabs have this class.
function Get-RibbonTabs($Window) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ControlType::TabItem)
    $tabs = @()
    foreach ($t in $Window.FindAll($Scope::Descendants, $cond)) {
        if ($t.Current.ClassName -eq 'NetUIRibbonTab') { $tabs += , $t }
    }
    , $tabs
}

function Find-RibbonTab($Window, [string]$Label, [double]$TimeoutSeconds) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $tabs = Get-RibbonTabs $Window
        $found = @($tabs | Where-Object { $_.Current.Name -ceq $Label })
        if ($found.Count -gt 1) { throw "Ambiguous tab '${Label}': $($found.Count) matching elements" }
        if ($found.Count -eq 1) { return $found[0] }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    throw "Not found: tab '$Label'"
}

function Find-UniqueElement($Root, $Condition, [double]$TimeoutSeconds, [string]$Description) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $found = $Root.FindAll($Scope::Descendants, $Condition)
        if ($found.Count -gt 1) { throw "Ambiguous ${Description}: $($found.Count) matching elements" }
        if ($found.Count -eq 1) { return $found.Item(0) }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    throw "Not found: $Description"
}

function Test-IsDescendant($Element, $Ancestor) {
    $parent = [System.Windows.Automation.TreeWalker]::RawViewWalker.GetParent($Element)
    while ($parent) {
        if ([System.Windows.Automation.Automation]::Compare($parent, $Ancestor)) { return $true }
        $parent = [System.Windows.Automation.TreeWalker]::RawViewWalker.GetParent($parent)
    }
    $false
}

# The tab content is a group named after the tab. A custom group with the same label is inside it,
# so when several groups match, take the one that contains all the others.
function Find-TabContent($Root, $Condition, [double]$TimeoutSeconds, [string]$Description) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $found = @($Root.FindAll($Scope::Descendants, $Condition))
        if ($found.Count -eq 1) { return $found[0] }
        if ($found.Count -gt 1) {
            $outer = @($found | Where-Object {
                $candidate = $_
                @($found | Where-Object {
                    -not [System.Windows.Automation.Automation]::Compare($_, $candidate) -and
                    -not (Test-IsDescendant $_ $candidate)
                }).Count -eq 0
            })
            if ($outer.Count -eq 1) { return $outer[0] }
            throw "Ambiguous ${Description}: $($found.Count) matching elements"
        }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    throw "Not found: $Description"
}

function New-NamedControlCondition([string]$Name, $Type) {
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $Type)
    [System.Windows.Automation.AndCondition]::new([System.Windows.Automation.Condition[]]@($nameCondition, $typeCondition))
}

Export-ModuleMember -Function Set-RibbonWindow, Test-CollapsedGroup, Get-CollapsedGroups, Get-RibbonTabs, Find-RibbonTab, Find-UniqueElement,
    Test-IsDescendant, Find-TabContent, New-NamedControlCondition
