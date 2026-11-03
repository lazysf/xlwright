# Spike: can the agent's shell start a dedicated Excel and end it?
# docs/architecture.md "本体の部品の境界", stage 0.
#
# Safety rules (docs/requirements.md goals 3 and 4, docs/architecture.md "所有の確定" and "終了"):
# - Starts Excel only with New-Object -ComObject (CoCreateInstance). Never uses GetActiveObject or the ROT.
# - Treats the started Excel as owned only when its Hwnd maps to a PID that did not exist before,
#   whose process is EXCEL.EXE and was created after the mark. If not, it releases the COM
#   reference, does not ask Excel to quit, does not terminate anything, and stops.
# - Asks Excel to quit only when it has no workbooks and no Protected View windows.
# - Terminates a process only through the handle opened at the ownership check, after checking
#   that it is still the same process (PID and start time).
# - Never touches an Excel process that existed before the run.
#
# Usage: powershell -NoProfile -File spikes\excel-start-stop\Test-ExcelStartStop.ps1

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Add-Type -Namespace XlWrightSpike -Name Native -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@

function Get-ExcelProcesses {
    @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | ForEach-Object {
        [pscustomobject]@{
            ProcessId = [int] $_.ProcessId
            CreatedUtc = $_.CreationDate.ToUniversalTime()
            CommandLine = $_.CommandLine
        }
    })
}

function Write-Result([string] $name, $value) {
    Write-Output ('{0}: {1}' -f $name, $value)
}

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'This spike must run on an STA thread (Windows PowerShell 5.1 console default).'
}

$before = Get-ExcelProcesses
Write-Result 'excel before' (($before | ForEach-Object { $_.ProcessId }) -join ',')

$markedAtUtc = (Get-Date).ToUniversalTime()
$excel = $null
$owned = $null
$ownedProcess = $null
$mayQuit = $false

try {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    $excel = New-Object -ComObject Excel.Application
    Write-Result 'created ms' $clock.ElapsedMilliseconds

    [uint32] $excelPid = 0
    [void] [XlWrightSpike.Native]::GetWindowThreadProcessId([System.IntPtr] [int64] $excel.Hwnd, [ref] $excelPid)
    $candidate = Get-CimInstance Win32_Process -Filter "ProcessId=$excelPid"
    $preexisting = @($before | Where-Object { $_.ProcessId -eq $excelPid }).Count -gt 0
    $checks = [ordered]@{
        pid_found = $excelPid -ne 0 -and $null -ne $candidate
        name_is_excel = $null -ne $candidate -and $candidate.Name -eq 'EXCEL.EXE'
        created_after_mark = $null -ne $candidate -and $candidate.CreationDate.ToUniversalTime() -gt $markedAtUtc
        not_preexisting = -not $preexisting
    }
    Write-Result 'pid' $excelPid
    Write-Result 'checks' (($checks.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ')
    if ($null -ne $candidate) {
        Write-Result 'command line has /automation' ($candidate.CommandLine -match '(?i)/automation')
        Write-Result 'command line' $candidate.CommandLine
    }

    if (@($checks.Values | Where-Object { -not $_ }).Count -gt 0) {
        Write-Result 'ownership' 'NOT CONFIRMED - leaving the process untouched and stopping'
    }
    else {
        # Open the handle now, so later waits and termination use this exact process.
        $ownedProcess = [System.Diagnostics.Process]::GetProcessById([int] $excelPid)
        [void] $ownedProcess.Handle
        $owned = [pscustomobject]@{
            ProcessId = [int] $excelPid
            StartTimeUtc = $ownedProcess.StartTime.ToUniversalTime()
        }
        Write-Result 'ownership' "confirmed (start $($owned.StartTimeUtc.ToString('o')))"

        Write-Result 'visible' $excel.Visible
        $workbookCount = $excel.Workbooks.Count
        $protectedCount = $excel.ProtectedViewWindows.Count
        Write-Result 'workbooks' $workbookCount
        Write-Result 'protected view windows' $protectedCount
        if ($workbookCount -eq 0 -and $protectedCount -eq 0) {
            $mayQuit = $true
        }
        else {
            Write-Result 'quit' 'SKIPPED - Excel has workbooks this spike did not open'
        }
    }

    if ($mayQuit) {
        $clock.Restart()
        $excel.Quit()
        Write-Result 'quit returned ms' $clock.ElapsedMilliseconds
    }
}
finally {
    if ($null -ne $excel) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}

if ($null -ne $ownedProcess -and $mayQuit) {
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    if ($ownedProcess.WaitForExit(30000)) {
        Write-Result 'exited after quit ms' $clock.ElapsedMilliseconds
        Write-Result 'exit code' $ownedProcess.ExitCode
    }
    else {
        $ownedProcess.Refresh()
        $same = -not $ownedProcess.HasExited -and $ownedProcess.Id -eq $owned.ProcessId -and
            $ownedProcess.StartTime.ToUniversalTime() -eq $owned.StartTimeUtc
        if ($same) {
            Write-Result 'exited after quit' 'NO within 30000 ms - terminating the owned process'
            $ownedProcess.Kill()
            Write-Result 'exited after kill' $ownedProcess.WaitForExit(10000)
        }
        else {
            Write-Result 'exited after quit' 'identity no longer matches - not terminating'
        }
    }
}

$after = Get-ExcelProcesses
Write-Result 'excel after' (($after | ForEach-Object { $_.ProcessId }) -join ',')
$changed = @($before | Where-Object {
    $old = $_
    @($after | Where-Object { $_.ProcessId -eq $old.ProcessId -and $_.CreatedUtc -eq $old.CreatedUtc }).Count -eq 0
})
$added = @($after | Where-Object {
    $new = $_
    @($before | Where-Object { $_.ProcessId -eq $new.ProcessId }).Count -eq 0
})
Write-Result 'preexisting excel unchanged' ($changed.Count -eq 0)
Write-Result 'excel left by this run' (($added | ForEach-Object { $_.ProcessId }) -join ',')
