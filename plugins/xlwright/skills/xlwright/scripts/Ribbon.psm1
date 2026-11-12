<#
.SYNOPSIS
UI Automation helpers shared by the ribbon scripts (Invoke-RibbonButton.ps1, Test-RibbonLayout.ps1).
Tabs are switched and buttons pressed through MSAA, not UI Automation: see Select-RibbonTab.
#>
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$ControlType = [System.Windows.Automation.ControlType]
$Walker = [System.Windows.Automation.TreeWalker]::RawViewWalker

# Add-Type fails if the type already exists, e.g. when the module is imported again with -Force.
if (-not ('XlWright.RibbonMsaa' -as [type])) {
    Add-Type -ReferencedAssemblies Accessibility -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;

namespace XlWright
{
    public static class RibbonMsaa
    {
        [DllImport("oleacc.dll")]
        static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out object accessible);

        [DllImport("oleacc.dll")]
        static extern int AccessibleChildren(IAccessible container, int start, int count,
            [Out] object[] children, out int obtained);

        const uint ObjIdClient = 0xFFFFFFFC;  // OBJID_CLIENT
        const int MaxDepth = 20;
        const int MaxVisited = 20000;

        // Runs the default action of the element with this name (and this MSAA role, unless role is 0)
        // in the window. Elements with the same name are told apart by their screen rectangle.
        // Returns how many elements matched; the action runs only when exactly one matched.
        public static int DoDefaultAction(IntPtr hwnd, string name, int role, int left, int top, int width, int height)
        {
            Guid iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");  // IID_IAccessible
            object root;
            int hr = AccessibleObjectFromWindow(hwnd, ObjIdClient, ref iid, out root);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            List<IAccessible> found = new List<IAccessible>();
            int visited = 0;
            Collect((IAccessible)root, name, role, 0, found, ref visited);
            if (found.Count > 1)
            {
                found = found.FindAll(delegate(IAccessible element)
                {
                    int l, t, w, h;
                    element.accLocation(out l, out t, out w, out h, 0);
                    return Math.Abs(l - left) <= 2 && Math.Abs(t - top) <= 2 &&
                        Math.Abs(w - width) <= 2 && Math.Abs(h - height) <= 2;
                });
            }
            if (found.Count == 1) found[0].accDoDefaultAction(0);
            return found.Count;
        }

        static bool Matches(IAccessible element, string name, int role)
        {
            string actualName;
            object actualRole;
            try
            {
                actualName = element.get_accName(0);
                actualRole = element.get_accRole(0);
            }
            catch (COMException) { return false; }  // the element went away or exposes no name or role
            return actualName == name && (role == 0 || (actualRole is int && (int)actualRole == role));
        }

        static void Collect(IAccessible parent, string name, int role, int depth, List<IAccessible> found, ref int visited)
        {
            if (depth > MaxDepth || visited > MaxVisited) return;
            int count;
            try { count = parent.accChildCount; }
            catch (COMException) { return; }  // the element went away while walking
            if (count <= 0) return;
            object[] children = new object[count];
            int obtained;
            int hr = AccessibleChildren(parent, 0, count, children, out obtained);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            for (int i = 0; i < obtained; i++)
            {
                visited++;
                // Children that are not objects (simple elements) are never ribbon tabs or buttons.
                IAccessible child = children[i] as IAccessible;
                if (child == null) continue;
                if (Matches(child, name, role)) found.Add(child);
                else Collect(child, name, role, depth + 1, found, ref visited);
            }
        }
    }
}
'@
}

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

# The arguments of [XlWright.RibbonMsaa]::DoDefaultAction for a UI Automation element: the window that
# hosts it (the ribbon is a NetUIHWND child window), its name, and its screen rectangle.
function Get-RibbonAccessibleTarget($Element) {
    $window = $Element
    while ($window -and $window.Current.NativeWindowHandle -eq 0) { $window = $Walker.GetParent($window) }
    if (!$window) { throw "The window hosting '$($Element.Current.Name)' was not found" }
    $r = $Element.Current.BoundingRectangle
    [pscustomobject]@{
        Hwnd = [IntPtr]$window.Current.NativeWindowHandle; Name = $Element.Current.Name
        Left = [int]$r.X; Top = [int]$r.Y; Width = [int]$r.Width; Height = [int]$r.Height
    }
}

# UI Automation's SelectionItemPattern.Select() and InvokePattern.Invoke() bring Excel to the foreground and
# take the keyboard from the application the user is working in. The MSAA default action of the same
# element ("switch" for a tab, "click" for a button) does not.
function Select-RibbonTab($Tab, [double]$TimeoutSeconds = 5) {
    $t = Get-RibbonAccessibleTarget $Tab
    $count = [XlWright.RibbonMsaa]::DoDefaultAction($t.Hwnd, $t.Name, 37, $t.Left, $t.Top, $t.Width, $t.Height)  # ROLE_SYSTEM_PAGETAB
    if ($count -ne 1) { throw "Expected one accessible tab '$($t.Name)', found $count" }
    $selection = $Tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while (-not $selection.Current.IsSelected) {
        if ((Get-Date) -ge $deadline) { throw "Tab '$($t.Name)' was not selected" }
        Start-Sleep -Milliseconds 100
    }
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
    Get-RibbonAccessibleTarget, Select-RibbonTab,
    Test-IsDescendant, Find-TabContent, New-NamedControlCondition
