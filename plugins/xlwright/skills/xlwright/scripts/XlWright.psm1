<#
.SYNOPSIS
Runs Excel through COM in a dedicated instance and always cleans up.

.DESCRIPTION
Use-ExcelWorkbook opens an existing workbook and invokes a script block with the Excel.Application
and Workbook objects. New-ExcelWorkbook creates a new empty workbook file. Invoke-ExcelMacro runs a macro
inside that script block and handles the dialogs the macro shows. Get-VbaTestName lists the VBA tests of
a project. Get-ExcelPrerequisite checks, without starting Excel, what stands in the way of opening a workbook.

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
[DllImport("user32.dll")] public static extern bool IsWindow(System.IntPtr hWnd);
[DllImport("user32.dll")] public static extern System.IntPtr GetDlgItem(System.IntPtr dialog, int id);
[DllImport("user32.dll")] public static extern bool PostMessage(System.IntPtr hWnd, uint message, System.IntPtr wParam, System.IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(System.IntPtr hWnd, System.Text.StringBuilder text, int count);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(System.IntPtr hWnd, System.Text.StringBuilder text, int count);

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

    # Quit returns before the process exits. PowerShell often keeps COM wrappers alive that GC does not
    # reclaim, and Excel then stays for minutes. The workbook is closed by now, so after a short wait
    # this call stops its own instance (same PID and start time) instead of leaving it behind.
    for ($i = 0; $i -lt 20; $i++) {
        if (-not (Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 500
    }
    $process = Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue
    if (-not $process) { return }
    if ($null -eq $Session.StartTicks -or $process.StartTime.ToUniversalTime().Ticks -ne $Session.StartTicks) {
        Write-Warning ("PID " + $Session.ProcessId + " could not be confirmed as the Excel instance of this call " +
            "and was left running. Report it to the user.")
        return
    }
    try {
        Stop-Process -Id $Session.ProcessId -Force -ErrorAction Stop
        Write-Verbose ("Stopped the Excel instance of this call (PID " + $Session.ProcessId + ") after Quit.")
    }
    catch {
        Write-Warning ("Could not stop the Excel instance of this call (PID " + $Session.ProcessId + "): " +
            $_.Exception.Message + " Report it to the user.")
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
        # Target file. The format follows the extension: .xlsm (macro-enabled) or .xlsx.
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(1, 86400)][int]$TimeoutSeconds = 120
    )
    $Path = [IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $Path) { throw "Already exists: $Path" }
    $format = switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.xlsm' { 52 }   # xlOpenXMLWorkbookMacroEnabled
        '.xlsx' { 51 }   # xlOpenXMLWorkbook
        default { throw "Unsupported extension: $Path" }
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
        $process = Get-Process -Id $State.ProcessId -ErrorAction SilentlyContinue
        # Only the instance of this call: same PID and start time.
        if ($process -and $process.StartTime.ToUniversalTime().Ticks -eq $State.StartTicks) {
            Stop-Process -Id $State.ProcessId -Force -ErrorAction SilentlyContinue
        }
    }
    while (-not $State.Stop) {
        $dialog = [XlWright.Dialog]::Find($excelPid)
        if ($dialog -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 100; continue }
        $State.DialogMessage = [XlWright.Dialog]::Message($dialog)
        $text = [XlWright.Dialog]::Title($dialog) + ': ' + $State.DialogMessage

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
    - A dialog that WM_CLOSE closes (a MsgBox with only OK): it is closed, the macro goes on, and the
      title and text are written as a warning.
    - Anything else (a Yes/No MsgBox, a compile error that the reset does not end): the text is read, this
      Excel instance is stopped by its PID and start time, and the text is thrown. Unsaved changes are lost.

    Returns what the macro returns, and nothing for a Sub. Use it inside the script block of
    Use-ExcelWorkbook with -EnableMacros.

    The exception of every case but the MsgBox with only OK carries Data['XlWrightFailure'] ('RuntimeError',
    'CompileError' or 'ExcelStopped') and Data['XlWrightDialogMessage'] (the text of the dialog without its
    title), so that a caller such as Invoke-VbaTest.ps1 does not have to parse the message.
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
        DialogMessage = $null; Excel = $Excel
        EditorWasVisible = [XlWright.Dialog]::IsEditorVisible([uint32]$processId)
        StartTicks = (Get-Process -Id $processId -ErrorAction Stop).StartTime.ToUniversalTime().Ticks
        Closed = [Collections.ArrayList]::Synchronized((New-Object Collections.ArrayList))
    })
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript($script:MacroDialogWatcher.ToString()).AddArgument($state)
    $async = $ps.BeginInvoke()
    $value = $null
    $failure = $null
    try {
        # InvokeMember passes a variable number of arguments, which $Excel.Run(...) cannot.
        $value = [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $Excel,
            [object[]](@($Name) + $ArgumentList))
    }
    catch { $failure = $_ }
    finally {
        $state.Stop = $true
        try { [void]$ps.EndInvoke($async) } finally { $ps.Dispose() }
    }
    foreach ($text in $state.Closed) { Write-Warning ("Closed a dialog shown by macro '" + $Name + "': " + $text) }
    $kind = $null
    $message = $null
    if ($state.Fatal) {
        $kind = 'ExcelStopped'
        $message = $state.Fatal + ' Stopped the Excel instance of this call (PID ' + $processId + '). Unsaved changes are lost.'
    }
    elseif ($state.CompileError) {
        $kind = 'CompileError'
        $message = "Macro '" + $Name + "' stopped with a compile error: " + $state.CompileError
    }
    elseif ($state.RuntimeError) {
        $kind = 'RuntimeError'
        $message = "Macro '" + $Name + "' stopped with a runtime error: " + $state.RuntimeError
    }
    if ($kind) {
        $exception = New-Object System.InvalidOperationException($message)
        $exception.Data['XlWrightFailure'] = $kind
        $exception.Data['XlWrightDialogMessage'] = $state.DialogMessage
        throw $exception
    }
    if ($failure) { throw $failure }
    # A Sub returns nothing; do not put a null into the output of the script block.
    if ($null -ne $value) { $value }
}

function Get-VbaTestName {
    <#
    .SYNOPSIS
    Lists the VBA tests of a project as "<module>.<Sub>", in the order of the components and of the code.

    .DESCRIPTION
    A test is a Sub in a standard module that takes no arguments, is not Private and whose name starts with
    "Test_". Invoke-VbaTest.ps1 runs these; Publish-Workbook.ps1 -ExcludeTests reports the ones it leaves.
    #>
    param(
        # The VBProject of an open workbook. It must not be locked.
        [Parameter(Mandatory)]$Project
    )
    foreach ($component in $Project.VBComponents) {
        if ($component.Type -ne 1) { continue }  # vbext_ct_StdModule
        $module = $component.CodeModule
        $count = $module.CountOfLines
        if ($count -eq 0) { continue }
        # The VBE normalizes the code, so a declaration always starts its (continued) line.
        $code = [regex]::Replace($module.Lines(1, $count), ' _\r?\n', ' ')
        foreach ($line in ($code -split '\r?\n')) {
            # Private is left out on purpose: Application.Run would run a Private Sub by its qualified name.
            if ($line -match '^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+(Test_\w*)\s*\(\s*\)') {
                $component.Name + '.' + $Matches[1]
            }
        }
    }
}

function Get-ExcelPrerequisite {
    <#
    .SYNOPSIS
    Checks what this skill needs before a workbook is opened, without starting Excel or changing anything.

    .DESCRIPTION
    Returns Ready, Problems, Excel, AccessVBOM, Wmi and Workbook. Ready is true when Problems is empty.
    A problem is: a running Excel, a process that cannot be started through WMI (one hidden "cmd.exe /c exit"
    is started to find out, the way the watchdog is), with -Vba an AccessVBOM that is not 1, and with -Path a
    workbook that does not exist or has an owner file (~$<name>) next to it, i.e. is open somewhere.
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

    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $value = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    $enabled = $value -eq 1
    if ($Vba -and -not $enabled) {
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

Export-ModuleMember -Function Use-ExcelWorkbook, New-ExcelWorkbook, Get-ExcelProcessId, Invoke-ExcelMacro, Get-VbaTestName,
    Get-ExcelPrerequisite
