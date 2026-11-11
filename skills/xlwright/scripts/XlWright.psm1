<#
.SYNOPSIS
Runs Excel through COM in a dedicated instance and always cleans up.

.DESCRIPTION
Use-ExcelWorkbook opens an existing workbook and invokes a script block with the Excel.Application
and Workbook objects. New-ExcelWorkbook creates a new empty workbook file.

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
    $requirement = "This skill requires starting a process through WMI (Win32_Process.Create); without the " +
        "watchdog it cannot be used. Stop and report this to the user; do not work around it."
    try {
        $startup = New-CimInstance -ClassName Win32_ProcessStartup -ClientOnly -Property @{ ShowWindow = [uint16]0 }
        $created = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
            CommandLine = $commandLine; ProcessStartupInformation = $startup
        }
    }
    catch {
        throw ("Could not start the watchdog for Excel PID " + $ProcessId + ": " + $_.Exception.Message + " " +
            $requirement)
    }
    if ($created.ReturnValue -ne 0) {
        throw ("Could not start the watchdog for Excel PID " + $ProcessId + " (Win32_Process.Create returned " +
            $created.ReturnValue + "). " + $requirement)
    }
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

Export-ModuleMember -Function Use-ExcelWorkbook, New-ExcelWorkbook, Get-ExcelProcessId
