# Spike: does renaming a build folder fail while a server runs from it?
# docs/architecture.md "サーバーとビルドとプラグイン", stage 0.
#
# Copies the latest passed build to a scratch folder under .xlwright/spikes/, starts the
# server from the copy (no Excel), tries to rename the folder, then stops the server and
# tries again as a control. Touches only the copy.
#
# Usage: powershell -NoProfile -File spikes\build-folder-rename\Test-BuildFolderRename.ps1

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repo = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$latest = Get-Content -Raw (Join-Path $repo '.xlwright\builds\latest.json') | ConvertFrom-Json
$source = Split-Path -Parent (Join-Path $repo $latest.server)
$stamp = (Get-Date).ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'")
$work = Join-Path $repo ".xlwright\spikes\build-folder-rename\$stamp"
$folder = Join-Path $work 'build'
New-Item -ItemType Directory -Force $work | Out-Null
Copy-Item -Recurse $source $folder

function Test-Rename([string] $path, [string] $newName) {
    try {
        Rename-Item -LiteralPath $path -NewName $newName
        return 'renamed'
    }
    catch {
        return "failed: $($_.Exception.GetType().Name): $($_.Exception.Message)"
    }
}

$startInfo = New-Object System.Diagnostics.ProcessStartInfo 'dotnet'
$startInfo.Arguments = "`"$(Join-Path $folder 'XlWright.McpServer.dll')`" --project `"$repo`" --no-excel"
$startInfo.UseShellExecute = $false
$startInfo.RedirectStandardInput = $true
$startInfo.RedirectStandardOutput = $true
$server = [System.Diagnostics.Process]::Start($startInfo)
try {
    Start-Sleep -Seconds 3
    "build id: $($latest.build_id)"
    "server alive: $(-not $server.HasExited) (pid $($server.Id))"
    "while running: $(Test-Rename $folder 'build-renamed-running')"
}
finally {
    # Closing stdin ends the stdio server; kill only the process this script started.
    $server.StandardInput.Close()
    if (-not $server.WaitForExit(10000)) { $server.Kill() }
}

$current = if (Test-Path (Join-Path $work 'build-renamed-running')) { Join-Path $work 'build-renamed-running' } else { $folder }
"server exited: $($server.HasExited) (exit code $($server.ExitCode))"
"after exit: $(Test-Rename $current 'build-renamed-after-exit')"
"work folder: $work"
