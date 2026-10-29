#requires -Version 7.0
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid] $ExperimentId,
    [Parameter(Mandatory)][ValidateSet('user-shell', 'agent-shell', 'mcp-launcher')][string] $Route
)

$ErrorActionPreference = 'Stop'
$invocationId = [guid]::NewGuid().ToString('N')
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$relativeDirectory = ".xlwright/mcp-launch-probe/$($ExperimentId.ToString('N'))"
$observation = [ordered]@{
    at_utc = [DateTime]::UtcNow.ToString('O')
    route = $Route
    invocation_id = $invocationId
    process_id = $PID
    process_started_at_utc = $null
    session_id = $null
    process_bits = [IntPtr]::Size * 8
    powershell_version = $PSVersionTable.PSVersion.ToString()
    script_sha256 = $null
    working_directory_matches = $false
    integrity_sid = $null
    measurement = 'unknown'
}
$exitCode = 1
$whoami = $null
try {
    # Inspect every existing ancestor before creating the fixed experiment directory.
    $current = [IO.Path]::GetPathRoot($repository)
    $fullDirectory = Join-Path $repository $relativeDirectory
    foreach ($component in $fullDirectory.Substring($current.Length).Split([IO.Path]::DirectorySeparatorChar, [StringSplitOptions]::RemoveEmptyEntries)) {
        $current = Join-Path $current $component
        if ((Test-Path -LiteralPath $current) -and
            (Get-Item -LiteralPath $current -Force).Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Reparse point in experiment directory.'
        }
    }
    [IO.Directory]::CreateDirectory($fullDirectory) | Out-Null
    $recordPath = Join-Path $fullDirectory "$invocationId.json"
    $self = [Diagnostics.Process]::GetCurrentProcess()
    try {
        $observation.process_started_at_utc = $self.StartTime.ToUniversalTime().ToString('O')
        $observation.session_id = $self.SessionId
    } finally { $self.Dispose() }
    $observation.script_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    $observation.working_directory_matches = [StringComparer]::OrdinalIgnoreCase.Equals(
        [IO.Path]::GetFullPath((Get-Location).Path), $repository)

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = Join-Path ([Environment]::GetFolderPath('System')) 'whoami.exe'
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in @('/groups', '/fo', 'csv', '/nh')) { $start.ArgumentList.Add($argument) }
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $whoami = [Diagnostics.Process]::Start($start)
    $output = $whoami.StandardOutput.ReadToEndAsync()
    $errors = $whoami.StandardError.ReadToEndAsync()
    if (-not $whoami.WaitForExit([Math]::Max(0, 10000 - [int]$timer.ElapsedMilliseconds))) {
        throw 'whoami result unknown.'
    }
    # Process exit and asynchronous pipe completion are separate events.
    $reads = [Threading.Tasks.Task]::WhenAll([Threading.Tasks.Task[]]@($output, $errors))
    if (-not $reads.Wait([Math]::Max(0, 10000 - [int]$timer.ElapsedMilliseconds)) -or $whoami.ExitCode -ne 0) {
        throw 'whoami result unavailable.'
    }
    # Raw group names and personal SIDs stay in memory and never enter diagnostics.
    $groups = $output.GetAwaiter().GetResult() | ConvertFrom-Csv -Header Group, Type, Sid, Attributes
    $integrity = @($groups | Where-Object { $_.Sid -match '^S-1-16-[0-9]+$' })
    if ($integrity.Count -ne 1) { throw 'Integrity SID is ambiguous.' }
    $observation.integrity_sid = $integrity[0].Sid
    $observation.measurement = 'observed'
    $exitCode = 0
} catch {
    # Exception text can include personal paths. Only the outcome is persisted.
    $observation.measurement = 'unknown'
} finally {
    if ($null -ne $whoami) { $whoami.Dispose() }
}

try {
    if (-not $recordPath) { throw 'No safe record path.' }
    $stream = [IO.FileStream]::new($recordPath, [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes(($observation | ConvertTo-Json -Depth 3))
        $stream.Write($bytes)
        $stream.Flush($true)
    } finally { $stream.Dispose() }
} catch {
    $observation.measurement = 'unknown'
    $exitCode = 1
}
[Console]::Error.WriteLine("xlwright-launch-probe $invocationId $($observation.measurement)")
exit $exitCode
