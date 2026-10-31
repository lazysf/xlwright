#requires -Version 5.1
[CmdletBinding()]
param([switch] $Restore, [string] $TestFilter)
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$development = Join-Path $repository '.xlwright/development'

function Assert-LocalPath([string] $Path) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($repository + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Development paths must remain inside the repository.'
    }
    if ([IO.File]::Exists($fullPath) -and [IO.File]::GetAttributes($fullPath).HasFlag([IO.FileAttributes]::ReparsePoint)) {
        throw 'Development files may not use reparse points.'
    }
    for ($directory = [IO.DirectoryInfo]::new($fullPath); $null -ne $directory; $directory = $directory.Parent) {
        if ($directory.Exists -and $directory.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Development paths may not use reparse points.'
        }
    }
}

Assert-LocalPath $development
[IO.Directory]::CreateDirectory($development) | Out-Null
$lockPath = Join-Path $development 'check.lock'
if (Test-Path -LiteralPath $lockPath) {
    if ((Get-Item -LiteralPath $lockPath).Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) { throw 'Invalid check lock.' }
}
$checkLock = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
$runId = [Guid]::NewGuid().ToString('N')
$runDirectory = Join-Path $development "runs/$runId"
Assert-LocalPath $runDirectory
[IO.Directory]::CreateDirectory($runDirectory) | Out-Null
$result = [ordered]@{ runId = $runId; state = 'running'; testFilter = $TestFilter; stages = @() }

function Write-StageLogTail([string] $Name) {
    foreach ($stream in @('stdout', 'stderr')) {
        $path = Join-Path $runDirectory "$Name.$stream.log"
        if (-not (Test-Path -LiteralPath $path)) { continue }
        Write-Output "--- $Name.$stream.log (last 80 lines)"
        # Console output often lands in reports, so keep home paths out of it.
        Get-Content -LiteralPath $path -Tail 80 -Encoding UTF8 | ForEach-Object {
            $line = $_
            foreach ($form in @($repository, $repository.Replace('\', '/'), $repository.Replace('\', '\\'))) { $line = $line.Replace($form, '.') }
            $line.Replace($env:USERPROFILE, '~').Replace($env:USERPROFILE.Replace('\', '/'), '~')
        }
    }
}

function Invoke-Stage([string] $Name, [string] $Command, [string[]] $Arguments) {
    Write-Output "Running $Name"
    $info = [Diagnostics.ProcessStartInfo]::new($Command)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.WorkingDirectory = $repository
    # Callers provide only fixed repository-relative tokens without shell syntax.
    foreach ($token in $Arguments) {
        if ($token -notmatch '^[A-Za-z0-9_./:*?=\-]+$') { throw 'Unexpected stage argument.' }
    }
    $info.Arguments = $Arguments -join ' '
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $stage = [ordered]@{ name = $Name; state = 'running'; exitCode = $null; durationMs = $null }
    $result.stages += $stage
    $process = $null
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    try {
        $process = [Diagnostics.Process]::Start($info)
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        [IO.File]::WriteAllText((Join-Path $runDirectory "$Name.stdout.log"), $stdout.GetAwaiter().GetResult())
        [IO.File]::WriteAllText((Join-Path $runDirectory "$Name.stderr.log"), $stderr.GetAwaiter().GetResult())
        $stage.exitCode = $process.ExitCode
        if ($process.ExitCode -ne 0) { $stage.state = 'failed'; throw "$Name failed; inspect $Name.stdout.log and $Name.stderr.log in .xlwright/development/runs/$runId" }
        $stage.state = 'passed'
    }
    catch {
        $stage.state = 'failed'
        [IO.File]::AppendAllText((Join-Path $runDirectory "$Name.stderr.log"), [Environment]::NewLine + $_.Exception.Message)
        Write-StageLogTail $Name
        throw
    }
    finally {
        $stage.durationMs = $stopwatch.ElapsedMilliseconds
        if ($null -ne $process) { $process.Dispose() }
    }
}

try {
    Assert-LocalPath (Join-Path $development 'artifacts')
    if (Test-Path -LiteralPath (Join-Path $development 'artifacts')) {
        foreach ($entry in Get-ChildItem -LiteralPath (Join-Path $development 'artifacts') -Recurse -Force) {
            if ($entry.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) { throw 'Artifacts may not use reparse points.' }
        }
    }
    $mcpArguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', 'src/XlWright.McpServer/check.ps1')
    if ($Restore) { $mcpArguments += '-Restore' }
    $spikeArguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', 'spikes/Spike0/check.ps1')
    if ($TestFilter) {
        # The tests run the built Spike0, so a filtered run still has to rebuild it.
        Invoke-Stage 'spike' 'powershell.exe' $spikeArguments
        Invoke-Stage 'mcp' 'powershell.exe' ($mcpArguments + @('-TestFilter', $TestFilter))
    }
    else {
        # A failed run must not leave an earlier pass valid for preparing experiments.
        [IO.File]::Delete((Join-Path $development 'last-passed.json'))
        Invoke-Stage 'sources' 'node.exe' @('scripts/record-build.mjs', '--begin', $runId)
        Invoke-Stage 'spike' 'powershell.exe' $spikeArguments
        Invoke-Stage 'mcp' 'powershell.exe' $mcpArguments
        Invoke-Stage 'inspector' 'node.exe' @('scripts/inspector.mjs', '--check')
        Invoke-Stage 'markdown' 'node.exe' @('node_modules/markdownlint-cli2/markdownlint-cli2-bin.mjs', '**/*.md')
        Invoke-Stage 'diff' 'git.exe' @('diff', '--check')
        Invoke-Stage 'record' 'node.exe' @('scripts/record-build.mjs', '--finish', $runId)
    }

    $result.state = 'passed'
}
catch { $result.state = 'failed'; throw }
finally {
    [IO.File]::WriteAllText((Join-Path $runDirectory 'result.json'), ($result | ConvertTo-Json -Depth 8))
    Write-Output "Result: .xlwright/development/runs/$runId/result.json"
    $checkLock.Dispose()
}
