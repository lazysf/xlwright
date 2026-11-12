<#
.SYNOPSIS
UI Automation helpers shared by the ribbon scripts (Invoke-RibbonButton.ps1, Test-RibbonLayout.ps1).
#>
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]

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

Export-ModuleMember -Function Find-UniqueElement, Test-IsDescendant, Find-TabContent, New-NamedControlCondition
