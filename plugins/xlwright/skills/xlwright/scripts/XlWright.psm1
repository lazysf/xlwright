<#
.SYNOPSIS
Runs Excel through COM in a dedicated instance and always cleans up.

.DESCRIPTION
Use-ExcelWorkbook opens an existing workbook and invokes a script block with the Excel.Application
and Workbook objects. New-ExcelWorkbook creates a new empty workbook file. Invoke-ExcelMacro runs a macro
inside that script block and handles the dialogs the macro shows. Get-ExcelPrerequisite checks, without
starting Excel, what stands in the way of opening a workbook.
Read-ExcelRange and Write-ExcelRange read and write the cells of a range in one call.
Import-ExcelVba, Test-ExcelVbaCompile and Find-ExcelCellError do the work of Import-Vba.ps1, Test-VbaCompile.ps1
and Find-CellError.ps1 inside a script block, so that several steps share one Excel instance.

Use-ExcelWorkbook disables VBA macros and workbook events by default and opens without updating
external workbook links. Use -EnableMacros for an intentional macro test. Event handlers such as
Workbook_Open also require -EnableEvents.

On exit the workbook is closed without saving (unless -Save was given and the block completed),
Excel is asked to quit, and the COM objects are released. Each step is attempted even if an earlier
one fails. If this instance's process is still alive afterwards, its PID is reported as a warning.

Each started instance gets a watchdog: a separate hidden PowerShell process that stops this instance
(and only this instance, checked by PID and start time) when -TimeoutSeconds have passed. The
watchdog is created through WMI so that it keeps running even if the calling shell is killed, e.g.
by a command timeout. It ends by itself as soon as the instance exits. If the watchdog cannot be
started, the Excel instance is closed and an error is thrown; the skill is not usable without it.
#>

# Add-Type fails if the type already exists, e.g. when the module is imported again with -Force.
if (-not ('XlWright.Native' -as [type])) {
    Add-Type -Namespace XlWright -Name Native -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
}

# Dialogs (#32770) shown by one Excel process: the VBA runtime and compile error dialogs and MsgBox.
if (-not ('XlWright.Dialog' -as [type])) {
    Add-Type -Namespace XlWright -Name Dialog -MemberDefinition @'
public delegate bool EnumProc(System.IntPtr hWnd, System.IntPtr lParam);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, System.IntPtr lParam);
[DllImport("user32.dll")] static extern bool EnumChildWindows(System.IntPtr parent, EnumProc callback, System.IntPtr lParam);
[DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
[DllImport("user32.dll")] static extern bool IsWindowVisible(System.IntPtr hWnd);
[DllImport("user32.dll")] static extern int GetDlgCtrlID(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsWindow(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern System.IntPtr GetDlgItem(System.IntPtr dialog, int id);
[DllImport("user32.dll")] public static extern bool PostMessage(System.IntPtr hWnd, uint message, System.IntPtr wParam, System.IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(System.IntPtr hWnd, System.Text.StringBuilder text, int count);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(System.IntPtr hWnd, System.Text.StringBuilder text, int count);

// Raised with every change of the members; see Assert-TypeRevision.
public const int Revision = 2;

public const uint WM_CLOSE = 0x10;
public const uint BM_CLICK = 0xF5;

static string ClassOf(System.IntPtr hWnd) {
    var text = new System.Text.StringBuilder(256);
    GetClassName(hWnd, text, text.Capacity);
    return text.ToString();
}

public static string Title(System.IntPtr hWnd) {
    var text = new System.Text.StringBuilder(2048);
    GetWindowText(hWnd, text, text.Capacity);
    return text.ToString();
}

// The first visible top-level window of this class in the process, or zero.
static System.IntPtr FindWindowOf(uint processId, string className) {
    System.IntPtr found = System.IntPtr.Zero;
    EnumWindows((hWnd, lParam) => {
        uint owner;
        GetWindowThreadProcessId(hWnd, out owner);
        if (owner != processId || !IsWindowVisible(hWnd) || ClassOf(hWnd) != className) return true;
        found = hWnd;
        return false;
    }, System.IntPtr.Zero);
    return found;
}

// These dialogs are not owned by the Excel window (a compile error dialog is owned by the VBE window),
// so they are looked up among the top-level windows of the process.
public static System.IntPtr Find(uint processId) { return FindWindowOf(processId, "#32770"); }

public static bool IsEditorVisible(uint processId) {
    return FindWindowOf(processId, "wndclass_desked_gsk") != System.IntPtr.Zero;
}

// The message is in the Static children; the buttons are left out. Line breaks become single spaces.
public static string Message(System.IntPtr dialog) {
    var texts = new System.Collections.Generic.List<string>();
    EnumChildWindows(dialog, (hWnd, lParam) => {
        if (ClassOf(hWnd) != "Static") return true;
        var text = Title(hWnd);
        if (text.Length > 0) texts.Add(text);
        return true;
    }, System.IntPtr.Zero);
    return System.Text.RegularExpressions.Regex.Replace(string.Join(" ", texts.ToArray()), @"\s+", " ").Trim();
}

// The control ids of the Button children, e.g. 1 (OK) and 2 (Cancel) of a MsgBox with OK and Cancel.
static int[] ButtonIds(System.IntPtr dialog) {
    var ids = new System.Collections.Generic.List<int>();
    EnumChildWindows(dialog, (hWnd, lParam) => {
        if (ClassOf(hWnd) == "Button") ids.Add(GetDlgCtrlID(hWnd));
        return true;
    }, System.IntPtr.Zero);
    return ids.ToArray();
}

// True when closing the dialog would choose an answer for the macro. WM_CLOSE and the x button close a MsgBox
// with OK and Cancel too, with an answer the macro then acts on (which one was not read). Only a dialog whose
// buttons are 2 and Help (9) has a single answer: the OK button of a MsgBox with only OK has id 2
// (docs/notes/macro-dialog.md).
public static bool AsksForAnswer(System.IntPtr dialog) {
    foreach (var id in ButtonIds(dialog)) { if (id != 2 && id != 9) return true; }
    return false;
}

// Waits until the dialog window is gone. False when it is still there after the time.
public static bool WaitClosed(System.IntPtr dialog, int milliseconds) {
    var watch = System.Diagnostics.Stopwatch.StartNew();
    while (IsWindow(dialog)) {
        if (watch.ElapsedMilliseconds >= milliseconds) return false;
        System.Threading.Thread.Sleep(50);
    }
    return true;
}
'@
}

# Stops an Excel process only when it is still the instance that started at the given time (AGENTS.md: never stop
# an Excel that cannot be confirmed as the own one). A type, so that the other runspaces of the scripts can use it.
if (-not ('XlWright.ExcelProcess' -as [type])) {
    Add-Type -Namespace XlWright -Name ExcelProcess -MemberDefinition @'
// Raised with every change of the members; see Assert-TypeRevision.
public const int Revision = 1;

// Gone: no such process (it has exited). NotSame: another process, or its start time could not be read.
// Failed: it is the same process, but it could not be stopped; the reason is in the out parameter.
public enum StopResult { Stopped, Gone, NotSame, Failed }

public static long StartTicks(int processId) {
    using (var process = System.Diagnostics.Process.GetProcessById(processId)) {
        return process.StartTime.ToUniversalTime().Ticks;
    }
}

public static StopResult StopIfSame(int processId, long startTicks) {
    string reason;
    return StopIfSame(processId, startTicks, out reason);
}

public static StopResult StopIfSame(int processId, long startTicks, out string reason) {
    reason = null;
    System.Diagnostics.Process process;
    try { process = System.Diagnostics.Process.GetProcessById(processId); }
    catch (System.ArgumentException) { return StopResult.Gone; }
    using (process) {
        try {
            if (process.StartTime.ToUniversalTime().Ticks != startTicks) return StopResult.NotSame;
        }
        catch (System.InvalidOperationException) { return StopResult.Gone; }
        catch (System.ComponentModel.Win32Exception) { return StopResult.NotSame; }
        try {
            process.Kill();
            return StopResult.Stopped;
        }
        catch (System.InvalidOperationException) { return StopResult.Gone; }
        catch (System.ComponentModel.Win32Exception e) { reason = e.Message; return StopResult.Failed; }
    }
}
'@
}

# A PowerShell session keeps the types it loaded first: Import-Module -Force does not replace them, so a session
# that loaded an older module would run this one with old types and fail inside another runspace.
function Assert-TypeRevision([type]$Type, [int]$Revision) {
    $field = $Type.GetField('Revision')
    $loaded = if ($null -ne $field) { [int]$field.GetValue($null) } else { 0 }
    if ($loaded -ne $Revision) {
        throw ($Type.FullName + " of another version of XlWright.psm1 (revision " + $loaded + ", this one needs " +
            $Revision + ") is loaded in this PowerShell session. Run the script in a new powershell.exe.")
    }
}
Assert-TypeRevision ([XlWright.Dialog]) 2
Assert-TypeRevision ([XlWright.ExcelProcess]) 1

function Get-ExcelProcessId($Xl) {
    # Hwnd is available even for a hidden instance.
    $processId = [uint32]0
    [void][XlWright.Native]::GetWindowThreadProcessId([IntPtr]$Xl.Hwnd, [ref]$processId)
    [int]$processId
}

function Start-ExcelWatchdog([int]$ProcessId, [int]$TimeoutSeconds) {
    $process = Get-Process -Id $ProcessId -ErrorAction Stop
    $ticks = $process.StartTime.ToUniversalTime().Ticks
    # Only integers are embedded, so the script cannot be altered by the values.
    $script = @"
`$ErrorActionPreference = 'SilentlyContinue'
`$p = Get-Process -Id $ProcessId
if (-not `$p -or `$p.StartTime.ToUniversalTime().Ticks -ne $ticks) { exit }
if (-not `$p.WaitForExit($($TimeoutSeconds * 1000))) {
    `$q = Get-Process -Id $ProcessId
    if (`$q -and `$q.StartTime.ToUniversalTime().Ticks -eq $ticks) { Stop-Process -Id $ProcessId -Force }
}
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
    $commandLine = 'powershell.exe -NoProfile -NonInteractive -EncodedCommand ' + $encoded
    try { Start-WmiProcess $commandLine }
    catch {
        throw ("Could not start the watchdog for Excel PID " + $ProcessId + ": " + $_.Exception.Message + " " +
            "This skill requires starting a process through WMI (Win32_Process.Create); without the " +
            "watchdog it cannot be used. Stop and report this to the user; do not work around it.")
    }
}

# Starts a hidden process through WMI, so that it outlives the calling shell. Throws when it cannot.
function Start-WmiProcess([string]$CommandLine) {
    $startup = New-CimInstance -ClassName Win32_ProcessStartup -ClientOnly -Property @{ ShowWindow = [uint16]0 }
    $created = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
        CommandLine = $CommandLine; ProcessStartupInformation = $startup
    }
    if ($created.ReturnValue -ne 0) { throw ("Win32_Process.Create returned " + $created.ReturnValue + ".") }
}

function Start-ExcelSession([bool]$Visible, [bool]$EnableEvents, [int]$TimeoutSeconds) {
    $xl = New-Object -ComObject Excel.Application
    $session = @{
        Xl = $xl
        ProcessId = (Get-ExcelProcessId $xl)
        StartTicks = $null
        TimeoutSeconds = $TimeoutSeconds
        Deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        Watched = $false
    }
    try {
        $session.StartTicks = (Get-Process -Id $session.ProcessId -ErrorAction Stop).StartTime.ToUniversalTime().Ticks
        Start-ExcelWatchdog $session.ProcessId $TimeoutSeconds
        $session.Watched = $true
    }
    catch {
        Stop-ExcelSession $session $null
        throw
    }
    $xl.Visible = $Visible
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $EnableEvents
    $session
}

function Stop-ExcelSession($Session, $Workbook) {
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch { Write-Warning ("Workbook.Close failed: " + $_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch { }
    }
    # Release unreachable wrappers while Excel still answers. After Quit, a finalizer that calls into the
    # exiting Excel can block for about a minute, so no collection is forced after Quit.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    try { $Session.Xl.Quit() } catch { Write-Warning ("Excel.Quit failed: " + $_.Exception.Message) }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Session.Xl) } catch { }
    $Session.Xl = $null

    # Quit returns before the process exits, and some instances do not exit for minutes. The workbook is closed
    # by now, so after a short wait this call stops its own instance (same PID and start time) instead of
    # leaving it behind. Most instances exited 2.2 to 2.9 s after Quit; the ones that did not exit in 60 s had
    # started an Office license renewal (docs/notes/excel-quit.md). Those would still be stopped after a longer
    # wait, only later.
    for ($i = 0; $i -lt 10; $i++) {
        if (-not (Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 500
    }
    $stopped = 'NotSame'
    $reason = $null
    if ($null -ne $Session.StartTicks) {
        $stopped = [string][XlWright.ExcelProcess]::StopIfSame($Session.ProcessId, $Session.StartTicks, [ref]$reason)
    }
    elseif (-not (Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue)) { return }
    switch ($stopped) {
        'Stopped' { Write-Verbose ("Stopped the Excel instance of this call (PID " + $Session.ProcessId + ") after Quit.") }
        'Gone' { }
        'NotSame' {
            Write-Warning ("PID " + $Session.ProcessId + " could not be confirmed as the Excel instance of this call " +
                "and was left running. Report it to the user.")
        }
        default {
            Write-Warning ("Could not stop the Excel instance of this call (PID " + $Session.ProcessId + "): " +
                $reason + " Report it to the user.")
        }
    }
}

# The formats this skill handles (the table of workbook formats in SKILL.md). Others, e.g. .xltm and .xls, were not
# checked with every script, so they are rejected before Excel or the zip is touched.
$script:WorkbookExtensions = '.xlsx', '.xlsm', '.xlsb', '.xlam'

function Test-WorkbookExtension([string]$Path) {
    $script:WorkbookExtensions -contains [IO.Path]::GetExtension($Path).ToLowerInvariant()
}

function Assert-WorkbookExtension {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-WorkbookExtension $Path)) {
        throw (New-Object System.ArgumentException(("Unsupported workbook format: " + $Path + ". This skill handles " +
            ($script:WorkbookExtensions -join ', ') + "."), 'Path'))
    }
}

function Use-ExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        # Show the window. Needed when the ribbon (customUI) must be built, e.g. to click a button.
        [switch]$Visible,
        # Save the workbook after the script block completes without throwing.
        [switch]$Save,
        # Let Workbook_Open and other event handlers run.
        [switch]$EnableEvents,
        [switch]$EnableMacros,
        # The watchdog stops this Excel instance after this many seconds.
        [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300
    )
    $Path = (Resolve-Path -LiteralPath $Path).ProviderPath
    Assert-WorkbookExtension $Path
    $session = Start-ExcelSession ([bool]$Visible) ([bool]$EnableEvents) $TimeoutSeconds
    $wb = $null
    try {
        if (!$EnableMacros) {
            $session.Xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        }
        $wb = $session.Xl.Workbooks.Open($Path, 0)  # UpdateLinks = 0
        & $ScriptBlock $session.Xl $wb
        if ($Save) { $wb.Save() }
    }
    catch {
        if ([DateTime]::UtcNow -ge $session.Deadline) {
            throw ("Timed out after " + $TimeoutSeconds + " s. The watchdog stopped the Excel instance of this " +
                "call (PID " + $session.ProcessId + "). Unsaved changes are lost. Cause: " + $_.Exception.Message)
        }
        throw
    }
    finally {
        Stop-ExcelSession $session $wb
    }
}

function New-ExcelWorkbook {
    [CmdletBinding()]
    param(
        # Target file. The format follows the extension: .xlsm (macro-enabled), .xlsb (binary) or .xlsx.
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(1, 86400)][int]$TimeoutSeconds = 120
    )
    $Path = [IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $Path) { throw "Already exists: $Path" }
    $format = switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.xlsm' { 52 }   # xlOpenXMLWorkbookMacroEnabled
        '.xlsx' { 51 }   # xlOpenXMLWorkbook
        '.xlsb' { 50 }   # xlExcel12
        default { throw "Unsupported extension: $Path. A new workbook can be .xlsm, .xlsb or .xlsx." }
    }
    $session = Start-ExcelSession $false $false $TimeoutSeconds
    $wb = $null
    try {
        $wb = $session.Xl.Workbooks.Add()
        $wb.SaveAs($Path, $format)
    }
    finally {
        Stop-ExcelSession $session $wb
    }
    $Path
}

# Runs in another runspace while Application.Run is blocked by a dialog. It uses XlWright.Dialog and the
# values in $State. The only COM object it touches is $State.Excel, to reset the VBE after a compile error.
$script:MacroDialogWatcher = {
    param($State)
    $excelPid = [uint32]$State.ProcessId
    function Stop-OwnExcel([string]$Reason) {
        $State.Fatal = $Reason
        # Only the instance of this call: same PID and start time.
        [void][XlWright.ExcelProcess]::StopIfSame($State.ProcessId, $State.StartTicks)
    }
    while (-not $State.Stop) {
        $dialog = [XlWright.Dialog]::Find($excelPid)
        if ($dialog -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 100; continue }
        $text = [XlWright.Dialog]::Title($dialog) + ': ' + [XlWright.Dialog]::Message($dialog)

        # 4800 is the End button of the VBA runtime error dialog. The button texts depend on the display
        # language, the control id does not. WM_CLOSE does not close this dialog.
        $endButton = [XlWright.Dialog]::GetDlgItem($dialog, 4800)
        if ($endButton -ne [IntPtr]::Zero) {
            $State.RuntimeError = $text
            [void][XlWright.Dialog]::PostMessage($endButton, [XlWright.Dialog]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
        }
        elseif ([XlWright.Dialog]::IsEditorVisible($excelPid)) {
            # A compile error shows the VBE. Closing its dialog leaves the VBE in break mode and Run does not
            # return until the VBE is reset (docs/notes/macro-dialog.md).
            $cannotContinue = 'The macro stopped with a dialog of the VBE and cannot continue: ' + $text
            [void][XlWright.Dialog]::PostMessage($dialog, [XlWright.Dialog]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
            if (-not [XlWright.Dialog]::WaitClosed($dialog, 3000)) { Stop-OwnExcel $cannotContinue; return }
            try {
                # COM calls from this runspace go through while Run is blocked in the calling one.
                $vbe = $State.Excel.VBE
                $reset = $vbe.CommandBars.FindControl(1, 228)  # msoControlButton, "Reset"
                $watch = [Diagnostics.Stopwatch]::StartNew()
                # Reset was enabled in break mode. Whether it is disabled before that was not checked.
                while (-not $reset.Enabled -and $watch.ElapsedMilliseconds -lt 3000) { Start-Sleep -Milliseconds 50 }
                $reset.Execute()
                if (-not $State.EditorWasVisible) { $vbe.MainWindow.Visible = $false }
            }
            catch { Stop-OwnExcel ($cannotContinue + ' Reset failed: ' + $_.Exception.Message); return }
            # Run fails a few milliseconds after the reset; the caller then sets Stop.
            $watch = [Diagnostics.Stopwatch]::StartNew()
            while (-not $State.Stop -and $watch.ElapsedMilliseconds -lt 3000) { Start-Sleep -Milliseconds 50 }
            if (-not $State.Stop) { Stop-OwnExcel $cannotContinue }
            else { $State.CompileError = $text }
            return
        }
        else {
            # WM_CLOSE would close a MsgBox with OK and Cancel too, choosing an answer for the macro.
            if ([XlWright.Dialog]::AsksForAnswer($dialog)) {
                Stop-OwnExcel ('The macro shows a dialog that cannot be closed without choosing an answer: ' + $text)
                return
            }
            [void][XlWright.Dialog]::PostMessage($dialog, [XlWright.Dialog]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
        }
        if (-not [XlWright.Dialog]::WaitClosed($dialog, 3000)) {
            # E.g. a Yes/No MsgBox: answering it would decide what the macro does.
            Stop-OwnExcel ('The macro shows a dialog that cannot be closed without choosing an answer: ' + $text)
            return
        }
        if ($endButton -eq [IntPtr]::Zero) { [void]$State.Closed.Add($text) }
    }
}

function Invoke-ExcelMacro {
    <#
    .SYNOPSIS
    Runs a macro like Application.Run and handles the dialogs that would block it.

    .DESCRIPTION
    Application.Run does not return while the macro shows a dialog, and a hidden Excel instance gives no
    hint of it. While the macro runs, another runspace watches the dialogs (#32770) of this instance:

    - VBA runtime error: the dialog is ended with its End button and its text is thrown. Excel stays usable.
    - Compile error (the dialog comes with the VBE window): the dialog is closed, the VBE is reset with its
      Reset command and hidden again, and the text is thrown. Excel stays usable.
    - A dialog whose only button is the one WM_CLOSE presses (a MsgBox with only OK): it is closed, the macro
      goes on, and the title and text are written as a warning.
    - Anything else (a MsgBox with Yes/No or OK/Cancel, whose answer would decide what the macro does, a
      compile error that the reset does not end): the text is read, this
      Excel instance is stopped by its PID and start time, and the text is thrown. Unsaved changes are lost.

    Returns what the macro returns, and nothing for a Sub. Use it inside the script block of
    Use-ExcelWorkbook with -EnableMacros.
    #>
    [CmdletBinding()]
    param(
        # The Excel.Application given to the script block.
        [Parameter(Mandatory)]$Excel,
        # As for Application.Run, e.g. 'Module1.Main'.
        [Parameter(Mandatory)][string]$Name,
        # Application.Run takes at most 30 arguments.
        [ValidateCount(0, 30)][object[]]$ArgumentList = @()
    )
    $processId = Get-ExcelProcessId $Excel
    $state = [hashtable]::Synchronized(@{
        Stop = $false; ProcessId = $processId; RuntimeError = $null; CompileError = $null; Fatal = $null
        Excel = $Excel
        EditorWasVisible = [XlWright.Dialog]::IsEditorVisible([uint32]$processId)
        StartTicks = [XlWright.ExcelProcess]::StartTicks($processId)
        Closed = [Collections.ArrayList]::Synchronized((New-Object Collections.ArrayList))
    })
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript($script:MacroDialogWatcher.ToString()).AddArgument($state)
    $async = $ps.BeginInvoke()
    # Through InvokeMember, a value wrapped in a PSObject (e.g. a string from Join-Path or Get-Content) fails a String
    # parameter with DISP_E_TYPEMISMATCH; through $Excel.Run(...) it does not (docs/notes/macro-arguments.md).
    # A loop rather than the pipeline, which would unroll an argument that is an array.
    $arguments = New-Object object[] $ArgumentList.Count
    for ($i = 0; $i -lt $ArgumentList.Count; $i++) {
        $argument = $ArgumentList[$i]
        if ($argument -is [psobject]) { $argument = $argument.psobject.BaseObject }
        $arguments[$i] = $argument
    }
    $value = $null
    $failure = $null
    try {
        # InvokeMember passes a variable number of arguments, which $Excel.Run(...) cannot.
        $value = [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $Excel,
            [object[]](@($Name) + $arguments))
    }
    catch { $failure = $_ }
    finally {
        $state.Stop = $true
        try { [void]$ps.EndInvoke($async) } finally { $ps.Dispose() }
    }
    foreach ($text in $state.Closed) { Write-Warning ("Closed a dialog shown by macro '" + $Name + "': " + $text) }
    $message = $null
    if ($state.Fatal) {
        $message = $state.Fatal + ' Stopped the Excel instance of this call (PID ' + $processId + '). Unsaved changes are lost.'
    }
    elseif ($state.CompileError) {
        $message = "Macro '" + $Name + "' stopped with a compile error: " + $state.CompileError
    }
    elseif ($state.RuntimeError) {
        $message = "Macro '" + $Name + "' stopped with a runtime error: " + $state.RuntimeError
    }
    if ($message) { throw (New-Object System.InvalidOperationException($message)) }
    if ($failure) { throw $failure }
    # A Sub returns nothing; do not put a null into the output of the script block.
    if ($null -ne $value) { $value }
}

function Get-ExcelPrerequisite {
    <#
    .SYNOPSIS
    Checks what this skill needs before a workbook is opened, without starting Excel or changing anything.

    .DESCRIPTION
    Returns Ready, Problems, Excel, AccessVBOM, Wmi and Workbook. Ready is true when Problems is empty.
    A problem is: a running Excel, a process that cannot be started through WMI (one hidden "cmd.exe /c exit"
    is started to find out, the way the watchdog is), with -Vba an AccessVBOM that is not 1, and with -Path a
    workbook that does not exist, is not in a format this skill handles, or has an owner file (~$<name>) next to it,
    i.e. is open somewhere. An AccessVBOM key that cannot be read (not one that is missing) is a problem of its own
    with -Vba, which asks for the rights to read it rather than for the setting.
    The Git state of the workbook is reported but is never a problem.
    #>
    [CmdletBinding()]
    param(
        # The workbook that is going to be opened. Without it, Workbook is null.
        [string]$Path,
        # The work reads or changes VBA, which needs AccessVBOM.
        [switch]$Vba
    )
    $problems = New-Object System.Collections.Generic.List[string]

    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Sort-Object Id | ForEach-Object {
        # StartTime of a process of another user cannot be read; the process is still reported.
        $startTime = $null
        try { $startTime = $_.StartTime.ToString('s') } catch { }
        [ordered]@{ Id = $_.Id; StartTime = $startTime; WindowTitle = $_.MainWindowTitle }
    })
    if ($processes.Count -gt 0) {
        $problems.Add('Excel is running (PID ' + (($processes | ForEach-Object { $_.Id }) -join ', ') +
            '). Ask the user before going on; do not stop it.')
    }

    # A value set by policy is taken before the one the user set in the Trust Center. The order of the two policy keys,
    # and that a policy value wins at all, is not in the Microsoft documentation and was not checked on a machine with
    # a policy (docs/notes/prerequisite.md).
    # A missing key or value is "not set". Any other failure is a key that could not be read, e.g. in a sandbox:
    # a policy above the value found is then unknown, so the value is not taken as enabled.
    $value = $null
    $readError = $null
    foreach ($key in 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security',
        'HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Security', 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security') {
        try { $value = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction Stop).AccessVBOM }
        catch [System.Management.Automation.ItemNotFoundException], [System.Management.Automation.PSArgumentException] { }
        catch { if (-not $readError) { $readError = $key + ': ' + $_.Exception.Message } }
        if ($null -ne $value) { break }
    }
    $enabled = ($value -eq 1) -and -not $readError
    if ($Vba -and $readError) {
        $problems.Add('AccessVBOM could not be read (' + $readError + '). This is a permission or sandbox problem, ' +
            'not the setting: ask to run outside the sandbox or with the needed rights first; do not ask the user to ' +
            'change the setting for this.')
    }
    elseif ($Vba -and -not $enabled) {
        $problems.Add('AccessVBOM is not 1, so VBA cannot be read or changed. Ask the user to trust access to the ' +
            'VBA project object model; do not change the setting yourself.')
    }

    $wmiError = $null
    try { Start-WmiProcess 'cmd.exe /c exit' } catch { $wmiError = $_.Exception.Message }
    if ($wmiError) {
        $problems.Add('A process cannot be started through WMI (Win32_Process.Create), so this skill cannot be ' +
            'used: ' + $wmiError + ' Stop and report this to the user; do not work around it.')
    }

    $workbook = $null
    if ($Path) {
        $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $directory = Split-Path -Parent $full
        $name = Split-Path -Leaf $full
        $exists = Test-Path -LiteralPath $full -PathType Leaf
        # Excel keeps this file next to a workbook it has open, also when its window is hidden.
        $open = Test-Path -LiteralPath (Join-Path $directory ('~$' + $name)) -PathType Leaf
        if (-not $exists) { $problems.Add('Workbook not found: ' + $full) }
        if (-not (Test-WorkbookExtension $full)) {
            $problems.Add('Unsupported workbook format: ' + $name + '. This skill handles ' +
                ($script:WorkbookExtensions -join ', ') + '.')
        }
        if ($open) {
            $problems.Add('The workbook seems to be open: ~$' + $name + ' is next to it. Ask the user to close it.')
        }
        $git = [ordered]@{ Available = $false; InRepository = $null; Tracked = $null; Uncommitted = $null }
        if (Get-Command git -CommandType Application -ErrorAction SilentlyContinue) {
            $git.Available = $true
            # The arguments go in as one array: PowerShell would bind "--" itself when calling a script block.
            $runGit = {
                param([string[]]$Arguments)
                # git writes to stderr outside a repository; that must not stop the caller.
                $ErrorActionPreference = 'Continue'
                $output = & git -C $directory @Arguments 2>$null
                [pscustomobject]@{ Exit = $LASTEXITCODE; Output = @($output) }
            }
            $inside = if (Test-Path -LiteralPath $directory -PathType Container) {
                & $runGit @('rev-parse', '--is-inside-work-tree')
            }
            $git.InRepository = ($null -ne $inside -and $inside.Exit -eq 0 -and $inside.Output[0] -eq 'true')
            if ($git.InRepository) {
                $git.Tracked = (& $runGit @('ls-files', '--error-unmatch', '--', $name)).Exit -eq 0
                if ($git.Tracked) {
                    $git.Uncommitted = (& $runGit @('status', '--porcelain=v1', '--', $name)).Output.Count -gt 0
                }
            }
        }
        $workbook = [ordered]@{ Path = $full; Exists = $exists; OpenElsewhere = $open; Git = $git }
    }

    [ordered]@{
        Ready = ($problems.Count -eq 0); Problems = $problems.ToArray()
        Excel = [ordered]@{ Running = ($processes.Count -gt 0); Processes = $processes }
        AccessVBOM = [ordered]@{ Value = $value; Enabled = $enabled }
        Wmi = [ordered]@{ Ok = (-not $wmiError); Error = $wmiError }
        Workbook = $workbook
    }
}

# Range.Value2 and the like are set and read through InvokeMember. Setting them with PowerShell's property syntax
# fails with InvalidCastException once the same line has seen a value of another type, or after a value from
# ConvertFrom-Json (docs/notes/range-value.md).
function Set-ComProperty($Object, [string]$Name, $Value) {
    [void][System.__ComObject].InvokeMember($Name, [Reflection.BindingFlags]::SetProperty, $null, $Object, [object[]]@(, $Value))
}

function Get-ComProperty($Object, [string]$Name) {
    $value = [System.__ComObject].InvokeMember($Name, [Reflection.BindingFlags]::GetProperty, $null, $Object, $null)
    # Without the comma PowerShell would unroll the array into its elements.
    if ($value -is [Array]) { return , $value }
    $value
}

# Value2 of one cell is the value itself, not an array. This returns a 1-based [row, column] array in every case.
function Get-RangeArray($Range, [string]$Name) {
    $raw = Get-ComProperty $Range $Name
    if ($raw -is [Array] -and $raw.Rank -eq 2) { return , $raw }
    $one = [Array]::CreateInstance([object], [int[]]@(1, 1), [int[]]@(1, 1))
    $one[1, 1] = $raw
    , $one
}

function Assert-SingleArea($Range) {
    if ($Range.Areas.Count -ne 1) {
        throw (New-Object System.ArgumentException(("The range " + $Range.Address($false, $false) + " has " +
            $Range.Areas.Count + " areas. Pass one rectangular range."), 'Range'))
    }
}

# XlCVError in the Microsoft documentation, and 2043 and 2050, which Excel returned but the documentation does not list.
$script:CellErrorName = @{
    2000 = '#NULL!'; 2007 = '#DIV/0!'; 2015 = '#VALUE!'; 2023 = '#REF!'; 2029 = '#NAME?'; 2036 = '#NUM!'; 2042 = '#N/A'
    2043 = '#GETTING_DATA'; 2045 = '#SPILL!'; 2050 = '#CALC!'
}

function Read-ExcelRange {
    <#
    .SYNOPSIS
    Reads a range in one COM call and returns it as an array of rows.

    .DESCRIPTION
    $rows[0][0] is the top-left cell. A single cell and a single row come back in the same shape. The values are
    those of Range.Value2: null for an empty cell, strings, doubles (dates are serial numbers) and booleans.
    An error value becomes @{ Error = '#DIV/0!' }; a code that is not known becomes '#ERROR<code>'.
    With -Formula the strings of Range.Formula2 are returned instead: the formula, the constant as text, or ''.
    #>
    [CmdletBinding()]
    param(
        # One rectangular Range, e.g. $ws.Range('A1:C10') or $ws.UsedRange.
        [Parameter(Mandatory)]$Range,
        [switch]$Formula
    )
    Assert-SingleArea $Range
    $raw = Get-RangeArray $Range $(if ($Formula) { 'Formula2' } else { 'Value2' })
    $firstRow = $raw.GetLowerBound(0)
    $firstColumn = $raw.GetLowerBound(1)
    $rowCount = $raw.GetLength(0)
    $columnCount = $raw.GetLength(1)
    $rows = New-Object 'object[]' $rowCount
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = New-Object 'object[]' $columnCount
        for ($c = 0; $c -lt $columnCount; $c++) {
            $value = $raw[($firstRow + $r), ($firstColumn + $c)]
            # Value2 returns numbers as Double; an Int32 is an error value with the XlCVError code in its low 16 bits.
            if ($value -is [int]) {
                $code = $value -band 0xFFFF
                $name = $script:CellErrorName[$code]
                if (-not $name) { $name = '#ERROR' + $code }
                $value = [ordered]@{ Error = $name }
            }
            $row[$c] = $value
        }
        $rows[$r] = $row
    }
    , $rows
}

function ConvertTo-CellValue($Value, [string]$Position) {
    if ($Value -is [psobject]) { $Value = $Value.psobject.BaseObject }
    if ($null -eq $Value -or $Value -is [DBNull]) { return $null }
    if ($Value -is [string]) {
        # Excel cuts longer strings without an error.
        if ($Value.Length -gt 32767) {
            throw (New-Object System.ArgumentException(("Value" + $Position + " is a string of " + $Value.Length +
                " characters; a cell holds at most 32767."), 'Value'))
        }
        if ($Value.Length -eq 0) { return $null }
        return $Value
    }
    if ($Value -is [double] -or $Value -is [single]) {
        # NaN becomes #NUM! in the cell.
        if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
            throw (New-Object System.ArgumentException(("Value" + $Position + " is " + $Value +
                ", which a cell cannot hold."), 'Value'))
        }
        return $Value
    }
    if ($Value -is [bool] -or $Value -is [datetime] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long] -or
        $Value -is [int16] -or $Value -is [byte] -or $Value -is [sbyte] -or $Value -is [uint16] -or $Value -is [uint32] -or
        $Value -is [uint64]) {
        return $Value
    }
    # A char is written as its character code, an array as an empty cell; others fail inside Excel.
    throw (New-Object System.ArgumentException(("Value" + $Position + " is a " + $Value.GetType().FullName +
        ". Use null, a string, a number, a boolean or a DateTime."), 'Value'))
}

function Write-ExcelRange {
    <#
    .SYNOPSIS
    Writes an array of rows to a range in one COM call.

    .DESCRIPTION
    -Value is an array of rows of the same length, the shape Read-ExcelRange returns; write one row as ,@(1, 2, 3).
    -Range is the top-left cell, or a range of exactly the size of -Value.

    A cell gets null (empty), a string, a number, a boolean or a DateTime (written as its serial number; the number
    format is left as it is). Other values, and strings longer than 32767 characters, are rejected before anything
    is written. An empty string leaves the cell empty.

    Strings are written as text: '001', '1/2', 'TRUE' or '=A1' stay as they are. Excel would read such a string as
    typed input, so the strings it changed are written again with a leading apostrophe (PrefixCharacter). A string that
    starts with "=" (or the full-width U+FF1D) gets the apostrophe in the first write already, so that it is never calculated as a formula.
    With -Formula, the values go to Range.Formula2 and Excel reads them as typed input: '=A1*2' is a formula,
    '001' becomes 1.

    Nothing is returned. If Excel fails after the writing has started (e.g. on a protected sheet), the exception is
    passed on and the range may be partly written.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Range,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Value,
        [switch]$Formula
    )
    Assert-SingleArea $Range
    $rowCount = $Value.Count
    if ($rowCount -eq 0) { throw (New-Object System.ArgumentException('Value has no rows.', 'Value')) }
    $rows = New-Object 'object[]' $rowCount
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $Value[$r]
        if ($row -is [psobject]) { $row = $row.psobject.BaseObject }
        if ($row -is [string] -or $row -isnot [Collections.IList]) {
            throw (New-Object System.ArgumentException(("Value[" + $r + "] is not an array. Each row must be an " +
                "array; write one row as ,@(1, 2, 3)."), 'Value'))
        }
        $rows[$r] = $row
    }
    $columnCount = $rows[0].Count
    for ($r = 0; $r -lt $rowCount; $r++) {
        if ($rows[$r].Count -ne $columnCount -or $columnCount -eq 0) {
            throw (New-Object System.ArgumentException(("Value[" + $r + "] has " + $rows[$r].Count + " values and " +
                "Value[0] has " + $columnCount + ". All rows must have the same number of values, at least one."), 'Value'))
        }
    }

    $given = $Range.Rows.Count.ToString() + 'x' + $Range.Columns.Count
    if ($given -ne '1x1' -and $given -ne ($rowCount.ToString() + 'x' + $columnCount)) {
        throw (New-Object System.ArgumentException(("The range " + $Range.Address($false, $false) + " is " + $given +
            " but Value is " + $rowCount + "x" + $columnCount + ". Pass the top-left cell or a range of that size."), 'Range'))
    }
    $anchor = $Range.Cells.Item(1, 1)
    $sheet = $Range.Worksheet
    if ($anchor.Row + $rowCount - 1 -gt $sheet.Rows.Count -or $anchor.Column + $columnCount - 1 -gt $sheet.Columns.Count) {
        throw (New-Object System.ArgumentException(("Value is " + $rowCount + "x" + $columnCount + " and does not fit " +
            "on the sheet from " + $anchor.Address($false, $false) + "."), 'Value'))
    }

    $array = [Array]::CreateInstance([object], [int[]]@($rowCount, $columnCount), [int[]]@(1, 1))
    $hasString = $false
    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $columnCount; $c++) {
            $cell = ConvertTo-CellValue $row[$c] ('[' + $r + '][' + $c + ']')
            if ($cell -is [string]) { $hasString = $true }
            $array[($r + 1), ($c + 1)] = $cell
        }
    }
    $target = $anchor.Resize($rowCount, $columnCount)
    if ($Formula) {
        Set-ComProperty $target 'Formula2' $array
        return
    }
    if (-not $hasString) {
        Set-ComProperty $target 'Value2' $array
        return
    }

    $intended = $array.Clone()
    # A string that starts with "=" or the full-width U+FF1D becomes a formula and is calculated as soon as Excel gets
    # it, so the apostrophe has to be there in the first write; writing it again afterwards would be too late.
    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $columnCount; $c++) {
            $text = $intended[$r, $c]
            if ($text -is [string] -and ($text[0] -eq [char]'=' -or $text[0] -eq [char]0xFF1D)) { $array[$r, $c] = "'" + $text }
        }
    }
    Set-ComProperty $target 'Value2' $array

    # Excel parses the other strings in an array as typed input too. The strings it changed are written again as text.
    $back = Get-RangeArray $target 'Value2'
    $changed = $false
    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $columnCount; $c++) {
            $text = $intended[$r, $c]
            if ($text -is [string] -and -not ($back[$r, $c] -is [string] -and $back[$r, $c] -ceq $text)) {
                $array[$r, $c] = "'" + $text
                $changed = $true
            }
        }
    }
    if (-not $changed) { return }
    Set-ComProperty $target 'Value2' $array
    $back = Get-RangeArray $target 'Value2'
    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $columnCount; $c++) {
            $text = $intended[$r, $c]
            if ($text -is [string] -and -not ($back[$r, $c] -is [string] -and $back[$r, $c] -ceq $text)) {
                throw ("Value[" + ($r - 1) + "][" + ($c - 1) + "] was not kept as text in " +
                    $target.Cells.Item($r, $c).Address($false, $false) + ". The range has been written.")
            }
        }
    }
}

# The bodies of Import-Vba.ps1, Test-VbaCompile.ps1 and Find-CellError.ps1. Called in one script block, they share
# one Excel instance: every start and exit of Excel takes about 4 s (docs/notes/dev-cycle-timing.md).

# The folders of a VBA source directory (the layout in SKILL.md) and the component type each one holds.
$script:VbaFolderTypes = @{ standard = 1; classes = 2; forms = 3; documents = 100 }
$script:VbaExtensions = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }

# Reads and checks the sources without Excel, so that a bad source fails before any code of the workbook changes.
# Returns Sources (File, Name, Text, ImportBytes, Sidecar, Type) and References (the entries of references.json).
function Read-VbaSource([string]$SourceDir) {
    $cp932 = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback,
        [Text.DecoderFallback]::ExceptionFallback)
    if (Get-ChildItem -LiteralPath $SourceDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' }) {
        throw "Place VBA files in standard/, classes/, documents/, or forms/, not directly in $SourceDir"
    }
    $sourceFolders = @(Get-ChildItem -LiteralPath $SourceDir -Directory |
        Where-Object { $script:VbaFolderTypes.ContainsKey($_.Name) })
    $files = @(foreach ($folder in $sourceFolders) {
        Get-ChildItem -LiteralPath $folder.FullName -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' }
    })
    if ($files.Count -eq 0) { throw "No .bas / .cls / .frm files in VBA component directories under $SourceDir" }

    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $sources = @(foreach ($f in $files) {
        $bytes = [IO.File]::ReadAllBytes($f.FullName)
        try { $text = $strictUtf8.GetString($bytes) }
        catch { throw ("Invalid UTF-8 in " + $f.FullName) }
        if ($text.StartsWith([string][char]0xFEFF, [StringComparison]::Ordinal) -or $text.Contains("`r")) {
            throw ("Expected UTF-8 without BOM and LF line endings in " + $f.FullName)
        }
        $text = $text.Replace("`n", "`r`n")
        try { $importBytes = $cp932.GetBytes($text) }
        catch { throw ("Cannot encode as CP932: " + $f.FullName) }
        if (![string]::Equals($cp932.GetString($importBytes), $text, [StringComparison]::Ordinal)) {
            throw ("Cannot round-trip through CP932: " + $f.FullName)
        }
        $m = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\r$')
        if (-not $m.Success) { throw ("No Attribute VB_Name in " + $f.Name) }
        $name = $m.Groups[1].Value
        $type = $script:VbaFolderTypes[$f.Directory.Name]
        if ($f.Extension -ne $script:VbaExtensions[$type] -or $f.BaseName -ne $name) {
            throw ("Unexpected extension or VB_Name in " + $f.FullName)
        }
        $sidecar = $null
        if ($type -eq 3) {
            $designer = $text.Substring(0, $m.Index)
            if ($designer -match '(?m)^[ \t]*OleObjectBlob[ \t]*=') {
                $blob = [regex]::Match($designer, '(?m)^[ \t]*OleObjectBlob[ \t]*=[ \t]*"([^"]+)":[0-9A-Fa-f]+[ \t]*\r$')
                $sidecarName = $f.BaseName + '.frx'
                if (!$blob.Success -or $blob.Groups[1].Value -ne $sidecarName) {
                    throw ("Expected a reference to " + $sidecarName + " in " + $f.FullName)
                }
                $sidecar = Join-Path $f.DirectoryName $sidecarName
                if (!(Test-Path -LiteralPath $sidecar -PathType Leaf)) {
                    throw ("Missing UserForm resource: " + $sidecar)
                }
            }
        }
        [pscustomobject]@{
            File = $f.Name
            Name = $name
            Text = $text
            ImportBytes = $importBytes
            Sidecar = $sidecar
            Type = $type
        }
    })

    $duplicate = $sources | Group-Object Name | Where-Object Count -gt 1
    if ($duplicate) { throw ("Duplicate VB_Name: " + ($duplicate.Name -join ', ')) }

    $expectedReferences = @()
    $referencesFile = Join-Path $SourceDir 'references.json'
    if (Test-Path -LiteralPath $referencesFile -PathType Leaf) {
        try { $parsed = $strictUtf8.GetString([IO.File]::ReadAllBytes($referencesFile)) | ConvertFrom-Json }
        catch { throw ("Invalid references.json: " + $referencesFile + ": " + $_.Exception.Message) }
        $expectedReferences = @(foreach ($entry in @($parsed.References)) {
            if ($null -eq $entry -or $entry.Name -isnot [string] -or $entry.Guid -isnot [string] -or
                $entry.Guid -notmatch '^\{[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$') {
                throw ("Expected References with a Name and a Guid for each entry in " + $referencesFile)
            }
            $entry
        })
    }
    [pscustomobject]@{ Sources = $sources; References = $expectedReferences }
}

function Import-ExcelVba {
    <#
    .SYNOPSIS
    Imports the UTF-8 (no BOM), LF .bas / .cls / .frm files of a source directory into an open workbook.

    .DESCRIPTION
    The rules are those of Import-Vba.ps1, which calls this function and then saves the workbook. This function
    does not save: run it in a script block of Use-ExcelWorkbook with -Save, or in Invoke-ExcelScript.ps1 with -Save.
    All sources are checked before any code of the workbook changes.

    Outputs one line per component: "imported: <name>" or "replaced document module: <name>".
    A reference in references.json that the workbook does not have is written as a warning.
    #>
    [CmdletBinding()]
    param(
        # The Workbook given to the script block.
        [Parameter(Mandatory)]$Workbook,
        # The folder with standard/, classes/, documents/ and forms/ (and references.json).
        [Parameter(Mandatory)][string]$SourceDir
    )
    $read = Read-VbaSource $SourceDir
    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-import-' + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    try {
        $components = $Workbook.VBProject.VBComponents
        $existingByName = @{}
        foreach ($component in $components) { $existingByName[$component.Name] = $component }
        foreach ($s in $read.Sources) {
            $existing = $existingByName[$s.Name]
            if ($null -ne $existing -and $existing.Type -ne $s.Type) {
                throw ("Component type mismatch for " + $s.Name + ": workbook " + $existing.Type + ", source " + $s.Type)
            }
            if ($s.Type -eq 100 -and $null -eq $existing) {
                throw ("No document module named " + $s.Name + " in the workbook")
            }
        }
        foreach ($s in $read.Sources) {
            $existing = $existingByName[$s.Name]
            if ($s.Type -eq 100) {
                $lines = $s.Text -split "`r`n"
                $i = 0
                if ($lines[0] -eq 'VERSION 1.0 CLASS') {
                    while ($i -lt $lines.Count -and $lines[$i] -ne 'END') { $i++ }
                    $i++
                }
                while ($i -lt $lines.Count -and $lines[$i] -like 'Attribute *') { $i++ }
                $code = ''
                if ($i -lt $lines.Count) { $code = ($lines[$i..($lines.Count - 1)] -join "`r`n").TrimEnd("`r", "`n") }
                $cm = $existing.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                if ($code.Length -gt 0) { $cm.AddFromString($code) }
                "replaced document module: " + $s.Name
                continue
            }
            $tmp = Join-Path $tmpDir $s.File
            [IO.File]::WriteAllBytes($tmp, $s.ImportBytes)
            if ($s.Sidecar) {
                [IO.File]::Copy($s.Sidecar, [IO.Path]::ChangeExtension($tmp, '.frx'))
            }
            if ($null -ne $existing) { $components.Remove($existing) }
            $c = $components.Import($tmp)
            if ($c.Name -ne $s.Name -or $c.Type -ne $s.Type) {
                throw ("Imported " + $s.File + " as " + $c.Name + " (type " + $c.Type + "), expected " +
                    $s.Name + " (type " + $s.Type + ")")
            }
            "imported: " + $s.Name
        }
        # The keys of this hashtable are not case-sensitive, and neither are GUIDs.
        $present = @{}
        foreach ($r in $Workbook.VBProject.References) { $present[[string]$r.Guid] = $true }
        foreach ($e in $read.References) {
            if (-not $present.ContainsKey($e.Guid)) {
                Write-Warning ("Reference not in the workbook: " + $e.Name + " " + $e.Guid +
                    "; code that uses it does not compile until it is added")
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Test-ExcelVbaCompile {
    <#
    .SYNOPSIS
    Compiles the VBA project of an open workbook and returns the compile error, if any.

    .DESCRIPTION
    Makes the project the active one in the VBE and executes the VBE command "Compile VBAProject" (command bar
    control 578); there is no COM method for compiling. The outcome is read from what the VBE does next:
    - The control becomes disabled: the project compiled.
    - A dialog (#32770) of this Excel instance appears: a compile error. Its text and the selection of the active
      code pane (the place of the error) are recorded, the dialog is closed with WM_CLOSE and the VBE window,
      which the error makes visible on the screen, is hidden again.
    - Neither within -DialogTimeoutSeconds: an exception, not a pass.

    Returns Compiled (true / false) and CompileError (Message, Module, Line, Column; null when it compiled). Only
    the error the VBE shows is reported; the place is null when the VBE has no active code pane. A locked project
    is not compiled and throws. If the dialog cannot be closed, this Excel instance is stopped by its PID and start
    time and an exception is thrown; its Data['CompileResult'] holds what was read. Excel stays usable after a
    compile error, also for macros of the modules that compile (docs/notes/vba-one-session.md).
    #>
    [CmdletBinding()]
    param(
        # The Excel.Application given to the script block.
        [Parameter(Mandatory)]$Excel,
        # The Workbook given to the script block.
        [Parameter(Mandatory)]$Workbook,
        # How long to wait for the compile to finish or the error dialog to appear.
        [ValidateRange(1, 3600)][int]$DialogTimeoutSeconds = 5
    )
    $result = [ordered]@{ Compiled = $null; CompileError = $null }
    $excelPid = Get-ExcelProcessId $Excel
    $excelStartTicks = [XlWright.ExcelProcess]::StartTicks($excelPid)
    $project = $Workbook.VBProject
    if ($project.Protection -ne 0) { throw 'The VBA project is locked and cannot be compiled from outside.' }
    $vbe = $Excel.VBE
    # Only the active project is compiled.
    $vbe.ActiveVBProject = $project
    $control = $vbe.CommandBars.FindControl(1, 578)  # msoControlButton, "Compile VBAProject"
    if ($null -eq $control) { throw 'The VBE command "Compile VBAProject" (578) was not found.' }
    # A disabled control means the project is already in the compiled state; Execute would fail.
    if (-not $control.Enabled) { $result.Compiled = $true; return $result }

    $vbeWasVisible = [bool]$vbe.MainWindow.Visible
    $control.Execute()
    # Execute returns at once; the error dialog appears a moment later.
    $dialog = [IntPtr]::Zero
    $deadline = (Get-Date).AddSeconds($DialogTimeoutSeconds)
    do {
        $dialog = [XlWright.Dialog]::Find([uint32]$excelPid)
        if ($dialog -ne [IntPtr]::Zero) { break }
        if (-not $control.Enabled) { $result.Compiled = $true; return $result }
        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $deadline)
    if ($dialog -eq [IntPtr]::Zero) {
        throw "The compile could not be confirmed: no error dialog and the project is not in the compiled state after $DialogTimeoutSeconds s."
    }

    $place = [ordered]@{ Message = [XlWright.Dialog]::Message($dialog); Module = $null; Line = $null; Column = $null }
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

    [void][XlWright.Dialog]::PostMessage($dialog, [XlWright.Dialog]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
    if (-not [XlWright.Dialog]::WaitClosed($dialog, 3000)) {
        # The instance is blocked in a modal state; closing the workbook would hang. Stop only this instance:
        # same PID and start time.
        $stuck = 'The compile error dialog could not be closed.'
        $message = switch ([string][XlWright.ExcelProcess]::StopIfSame($excelPid, $excelStartTicks)) {
            'Stopped' { "$stuck Stopped the Excel instance of this script (PID $excelPid)." }
            'Gone' { "$stuck The Excel instance of this script (PID $excelPid) had already exited." }
            default { "$stuck Excel PID $excelPid could not be stopped as the instance of this script. Report it to the user." }
        }
        $failure = New-Object System.InvalidOperationException($message)
        $failure.Data['CompileResult'] = $result
        throw $failure
    }
    if (-not $vbeWasVisible) { $vbe.MainWindow.Visible = $false }
    $result
}

function Find-ExcelCellError {
    <#
    .SYNOPSIS
    Lists the cells of an open workbook whose value is an error (#REF!, #NAME?, #DIV/0! and the others).

    .DESCRIPTION
    Nothing is changed or recalculated: the values are those Excel holds now. The used range of every worksheet is
    read with Range.Value2 in blocks of rows (Range.SpecialCells is not used; docs/notes/cell-error.md). Hidden and
    protected sheets are read too; chart sheets are not.

    Returns Count (the error cells found), Truncated (true when Errors lists fewer than Count), Calculation
    ("automatic", "manual", "semiautomatic" or "other") and Sheets in the order of the workbook: Name, Count, Errors.
    Each error: Address (e.g. "B3"), Error (e.g. "#REF!"), Kind ("formula", "constant", or "spill" for a cell that
    a formula elsewhere spills into), Formula (Range.Formula2 of a formula cell, otherwise null) and SpillParent
    (the address of the spilling formula, only for Kind "spill").
    #>
    [CmdletBinding()]
    param(
        # The Excel.Application given to the script block.
        [Parameter(Mandatory)]$Excel,
        # The Workbook given to the script block.
        [Parameter(Mandatory)]$Workbook,
        # Only these worksheets. A name that is not a worksheet of the workbook fails.
        [string[]]$Sheet = @(),
        # Errors lists at most this many cells over all sheets; Count still counts all of them.
        [ValidateRange(0, 1000000)][int]$MaxErrors = 1000
    )
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Sheet = @($Sheet)
    $worksheets = @(foreach ($item in $Workbook.Sheets) {
        # -4167 is xlWorksheet. The type of a chart sheet is not a sheet type, so the class is checked too.
        if ([Microsoft.VisualBasic.Information]::TypeName($item) -eq 'Worksheet' -and [int]$item.Type -eq -4167) { $item }
    })
    if ($Sheet.Count -gt 0) {
        $known = @($worksheets | ForEach-Object { $_.Name })
        $missing = @($Sheet | Where-Object { $known -notcontains $_ })
        if ($missing.Count -gt 0) {
            throw ('Not a worksheet of the workbook: ' + ($missing -join ', ') + '. Worksheets: ' + ($known -join ', ') + '.')
        }
        $worksheets = @($worksheets | Where-Object { $Sheet -contains $_.Name })
    }

    $calculation = @{ -4105 = 'automatic'; -4135 = 'manual'; 2 = 'semiautomatic' }[[int]$Excel.Calculation]
    if (-not $calculation) { $calculation = 'other' }
    $total = 0
    $listed = 0
    $sheets = @(foreach ($worksheet in $worksheets) {
        $used = $worksheet.UsedRange
        $rowCount = $used.Rows.Count
        $columnCount = $used.Columns.Count
        # Bounds the size of one Value2 array also for a wide sheet.
        $blockRows = [Math]::Max(1, [Math]::Min(10000, [Math]::Floor(200000 / $columnCount)))
        $count = 0
        $errors = New-Object System.Collections.Generic.List[object]
        for ($start = 1; $start -le $rowCount; $start += $blockRows) {
            $block = $used.Rows.Item($start).Resize([Math]::Min($blockRows, $rowCount - $start + 1))
            $rows = Read-ExcelRange -Range $block
            for ($r = 0; $r -lt $rows.Count; $r++) {
                $row = $rows[$r]
                for ($c = 0; $c -lt $row.Count; $c++) {
                    if ($row[$c] -isnot [Collections.IDictionary]) { continue }
                    $count++
                    if ($listed -ge $MaxErrors) { continue }
                    $listed++
                    $cell = $block.Cells.Item($r + 1, $c + 1)
                    $entry = [ordered]@{
                        Address = $cell.Address($false, $false); Error = $row[$c].Error; Kind = 'constant'
                        Formula = $null; SpillParent = $null
                    }
                    if ($cell.HasFormula) {
                        $entry.Kind = 'formula'
                        $entry.Formula = [string]$cell.Formula2
                    }
                    # SpillParent fails for a cell outside a spill, so HasSpill is checked first.
                    elseif ($cell.HasSpill -eq $true) {
                        $entry.Kind = 'spill'
                        $entry.SpillParent = $cell.SpillParent.Address($false, $false)
                    }
                    $errors.Add($entry)
                }
            }
        }
        $total += $count
        [ordered]@{ Name = $worksheet.Name; Count = $count; Errors = $errors.ToArray() }
    })
    [ordered]@{ Count = $total; Truncated = ($listed -lt $total); Calculation = $calculation; Sheets = $sheets }
}

Export-ModuleMember -Function Use-ExcelWorkbook, New-ExcelWorkbook, Get-ExcelProcessId, Invoke-ExcelMacro,
    Get-ExcelPrerequisite, Read-ExcelRange, Write-ExcelRange, Assert-WorkbookExtension, Import-ExcelVba,
    Test-ExcelVbaCompile, Find-ExcelCellError
