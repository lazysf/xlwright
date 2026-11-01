#requires -Version 5.1
#requires -PSEdition Desktop
# Manual recovery step 3 in docs/architecture.md: shows a hidden owned Excel to a person.
# Run only after individual approval. Run -CheckOnly first; it sends nothing.
# Sends one display request to each hidden XLMAIN window of the recorded process. No COM, input or termination.
# A second run for the same process is refused, so a request is never repeated.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $RecordPath,
    [switch] $CheckOnly
)
$ErrorActionPreference = 'Stop'

if (-not [IO.Path]::IsPathRooted($RecordPath)) { throw 'RecordPath must be an absolute path.' }
$RecordPath = [IO.Path]::GetFullPath($RecordPath)
if (-not [IO.File]::Exists($RecordPath)) { throw 'The ownership record does not exist.' }
if ([IO.File]::GetAttributes($RecordPath).HasFlag([IO.FileAttributes]::ReparsePoint)) { throw 'The ownership record may not be a reparse point.' }

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class XlwrightOwnedWindows {
    public delegate bool Callback(IntPtr hwnd, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumWindows(Callback callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", EntryPoint="GetClassNameW", CharSet=CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hwnd, int command);
    public static uint OwnerOf(IntPtr hwnd) {
        uint owner;
        GetWindowThreadProcessId(hwnd, out owner);
        return owner;
    }
    public static List<long> FindXlmain(uint pid) {
        var found = new List<long>();
        EnumWindows((hwnd, parameter) => {
            var name = new StringBuilder(256);
            if (OwnerOf(hwnd) == pid && GetClassName(hwnd, name, name.Capacity) != 0 && name.ToString() == "XLMAIN") found.Add(hwnd.ToInt64());
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@

# A running server keeps the record open with FileShare.None. Holding it exclusively until the end proves that no
# server owns it while the request is sent, and keeps a new server from starting meanwhile.
try { $stream = [IO.FileStream]::new($RecordPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None) }
catch [IO.IOException] { throw 'The ownership record is held by another process; confirm the server has stopped first.' }
try {
    $text = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false)).ReadToEnd()
    try { $record = $text | ConvertFrom-Json }
    catch { throw 'The ownership record is not valid JSON; review it manually.' }
    if (-not ($record.version -is [int] -and $record.version -eq 1) -or $record.kind -cne 'owned' -or $null -eq $record.process -or
        -not ($record.process.pid -is [int] -and $record.process.pid -gt 0) -or $record.process.name -isnot [string]) {
        throw 'The ownership record is not an owned record; there is no recorded Excel to show.'
    }
    $expectedStart = [datetime]::ParseExact([string]$record.process.started_at_utc, 'o',
        [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)

    try { $process = [Diagnostics.Process]::GetProcessById($record.process.pid) }
    catch [ArgumentException] { throw 'The recorded Excel process is not running.' }
    try {
        # .NET Framework keeps this handle and reuses it for HasExited and StartTime. An open handle keeps the PID
        # from being reused, so the identity checked below stays the process the requests are sent to.
        $null = $process.Handle
        if ($process.HasExited -or -not [string]::Equals($process.ProcessName, $record.process.name, [StringComparison]::OrdinalIgnoreCase) -or
            $process.StartTime.ToUniversalTime() -ne $expectedStart) {
            throw 'Process identity does not match the record; do not show or end it.'
        }
        # The old server must be gone before a person works with its Excel (manual recovery step 2).
        # A dotnet process whose command line cannot be read is treated as a possible server.
        $servers = @(Get-CimInstance Win32_Process | Where-Object {
            $_.Name -eq 'XlWright.McpServer.exe' -or $_.CommandLine -like '*XlWright.McpServer.dll*' -or
            ($_.Name -eq 'dotnet.exe' -and [string]::IsNullOrEmpty($_.CommandLine))
        })
        if ($servers.Count -ne 0) { throw 'An xlwright server process may be running; confirm it has stopped first.' }

        $resultDirectory = Join-Path ([IO.Path]::GetDirectoryName($RecordPath)) '.xlwright/recovery'
        $resultPath = Join-Path $resultDirectory ('show-{0}-{1}.jsonl' -f $process.Id, $expectedStart.ToString('yyyyMMddTHHmmssfffffffZ'))
        if ([IO.File]::Exists($resultPath)) { throw 'Display was already requested for this process; do not retry.' }
        $windows = @([XlwrightOwnedWindows]::FindXlmain([uint32]$process.Id))
        if ($windows.Count -eq 0) { throw 'The recorded process has no XLMAIN window.' }
        foreach ($hwnd in $windows) {
            if ([XlwrightOwnedWindows]::IsWindowVisible([IntPtr]$hwnd)) { throw 'An XLMAIN window is already visible; ask the person first.' }
        }
        if ($CheckOnly) { "CHECK PASSED: $($windows.Count) hidden XLMAIN windows. No display requested."; return }

        [IO.Directory]::CreateDirectory($resultDirectory) | Out-Null
        $writer = [IO.StreamWriter]::new([IO.FileStream]::new($resultPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read))
        try {
            $writer.AutoFlush = $true
            foreach ($hwnd in $windows) {
                # A window handle can be reused by another process; check its owner again just before sending.
                if ($process.HasExited -or [XlwrightOwnedWindows]::OwnerOf([IntPtr]$hwnd) -ne [uint32]$process.Id) {
                    throw 'The recorded process or its window changed; do not retry.'
                }
                # 4 is SW_SHOWNOACTIVATE, the command used in the T11 recoveries.
                $started = [XlwrightOwnedWindows]::ShowWindowAsync([IntPtr]$hwnd, 4)
                $writer.WriteLine((@{ hwnd = $hwnd; requestStarted = $started; atUtc = [datetime]::UtcNow.ToString('o') } | ConvertTo-Json -Compress))
                if (-not $started) { throw 'Display request not accepted; do not retry.' }
            }
        }
        finally { $writer.Dispose() }
        "$($windows.Count) display requests started. No input was sent."
    }
    finally { $process.Dispose() }
}
finally { $stream.Dispose() }
