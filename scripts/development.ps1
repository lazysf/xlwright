#requires -Version 5.1
[CmdletBinding()]
param([switch] $Restore)
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
$result = [ordered]@{ runId = $runId; state = 'running'; stages = @() }

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
    $stage = [ordered]@{ name = $Name; state = 'running'; exitCode = $null }
    $result.stages += $stage
    $process = $null
    try {
        $process = [Diagnostics.Process]::Start($info)
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        [IO.File]::WriteAllText((Join-Path $runDirectory "$Name.stdout.log"), $stdout.GetAwaiter().GetResult())
        [IO.File]::WriteAllText((Join-Path $runDirectory "$Name.stderr.log"), $stderr.GetAwaiter().GetResult())
        $stage.exitCode = $process.ExitCode
        if ($process.ExitCode -ne 0) { $stage.state = 'failed'; throw "$Name failed; inspect .xlwright/development/runs/$runId/$Name.stderr.log" }
        $stage.state = 'passed'
    }
    catch {
        $stage.state = 'failed'
        [IO.File]::AppendAllText((Join-Path $runDirectory "$Name.stderr.log"), [Environment]::NewLine + $_.Exception.Message)
        throw
    }
    finally { if ($null -ne $process) { $process.Dispose() } }
}

try {
    Assert-LocalPath (Join-Path $development 'artifacts')
    if (Test-Path -LiteralPath (Join-Path $development 'artifacts')) {
        foreach ($entry in Get-ChildItem -LiteralPath (Join-Path $development 'artifacts') -Recurse -Force) {
            if ($entry.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) { throw 'Artifacts may not use reparse points.' }
        }
    }
    Invoke-Stage 'spike' 'powershell.exe' @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', 'spikes/Spike0/check.ps1', '-Isolated')
    $mcpArguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', 'src/XlWright.McpServer/check.ps1', '-Isolated')
    if ($Restore) { $mcpArguments += '-Restore' }
    Invoke-Stage 'mcp' 'powershell.exe' $mcpArguments
    Invoke-Stage 'inspector' 'node.exe' @('scripts/inspector.mjs', '--check')
    Invoke-Stage 'markdown' 'node.exe' @('node_modules/markdownlint-cli2/markdownlint-cli2-bin.mjs', '**/*.md')
    Invoke-Stage 'diff' 'git.exe' @('diff', '--check')

    $result.state = 'passed'
}
catch { $result.state = 'failed'; throw }
finally {
    [IO.File]::WriteAllText((Join-Path $runDirectory 'result.json'), ($result | ConvertTo-Json -Depth 8))
    Write-Output "Result: .xlwright/development/runs/$runId/result.json"
    $checkLock.Dispose()
}
