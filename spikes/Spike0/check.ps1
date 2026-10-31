#requires -Version 5.1
[CmdletBinding()]
param([switch] $Isolated)
$ErrorActionPreference = 'Stop'

$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$runtimeRoot = Join-Path $repositoryRoot '.xlwright'
if ((Test-Path -LiteralPath $runtimeRoot) -and
    (Get-Item -LiteralPath $runtimeRoot).Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
    throw 'The experiment workspace may not use a reparse point.'
}
$buildParent = Join-Path $runtimeRoot 'spike0-build'
if ((Test-Path -LiteralPath $buildParent) -and
    (Get-Item -LiteralPath $buildParent).Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
    throw 'The build workspace may not use a reparse point.'
}
# A new directory avoids consuming existing config/cache child paths or their links.
$buildId = [Guid]::NewGuid().ToString('N')
$buildRoot = Join-Path $buildParent $buildId
if (Test-Path -LiteralPath $buildRoot) { throw 'The build directory must be new.' }
[IO.Directory]::CreateDirectory($buildRoot) | Out-Null
$configPath = Join-Path $buildRoot 'NuGet.Config'
[IO.File]::WriteAllText($configPath,
    '<configuration><packageSources><clear /></packageSources></configuration>')
$dotnetPath = (Get-Command dotnet -CommandType Application).Source
$artifactArguments = @()
if ($Isolated) {
    $artifactRoot = Join-Path $PSScriptRoot '../../.xlwright/development/artifacts'
    for ($ancestor = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($artifactRoot)); $null -ne $ancestor; $ancestor = $ancestor.Parent) {
        if ($ancestor.Exists -and $ancestor.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Artifact path ancestors may not use reparse points.'
        }
    }
    if (Test-Path -LiteralPath $artifactRoot) {
        foreach ($entry in Get-ChildItem -LiteralPath $artifactRoot -Recurse -Force) {
            if ($entry.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) { throw 'Artifacts may not use reparse points.' }
        }
    }
    $artifactArguments = @('--artifacts-path', '.xlwright/development/artifacts')
}

function Invoke-SpikeDotnet([ValidatePattern('^[A-Za-z0-9_./:=\-]+$')][string[]] $DotnetArguments) {
    $processInfo = [Diagnostics.ProcessStartInfo]::new($dotnetPath)
    $processInfo.UseShellExecute = $false
    $processInfo.WorkingDirectory = $repositoryRoot
    # All callers use fixed tokens; reject values needing quoting on .NET Framework.
    $processInfo.Arguments = @($DotnetArguments + $artifactArguments) -join ' '
    # Child-process settings keep SDK initialization and NuGet away from personal config.
    $processInfo.EnvironmentVariables['DOTNET_CLI_HOME'] = Join-Path $buildRoot 'dotnet-home'
    $processInfo.EnvironmentVariables['APPDATA'] = Join-Path $buildRoot 'appdata'
    $processInfo.EnvironmentVariables['NUGET_PACKAGES'] = Join-Path $buildRoot 'packages'
    $processInfo.EnvironmentVariables['DOTNET_GENERATE_ASPNET_CERTIFICATE'] = 'false'
    $processInfo.EnvironmentVariables['DOTNET_ADD_GLOBAL_TOOLS_TO_PATH'] = 'false'
    $processInfo.EnvironmentVariables['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    $processInfo.EnvironmentVariables['DOTNET_NOLOGO'] = '1'
    $childProcess = [Diagnostics.Process]::Start($processInfo)
    try {
        $childProcess.WaitForExit()
        if ($childProcess.ExitCode -ne 0) {
            throw "dotnet exited with code $($childProcess.ExitCode)."
        }
    }
    finally { $childProcess.Dispose() }
}

# Use the repository-relative path so checkout names never enter the argument string.
Invoke-SpikeDotnet @('restore', 'spikes/Spike0/Spike0.csproj', '--configfile',
    ".xlwright/spike0-build/$buildId/NuGet.Config", '-p:NuGetAudit=false')
foreach ($configuration in @('Debug', 'Release')) {
    Invoke-SpikeDotnet @('build', 'spikes/Spike0/Spike0.csproj', '--configuration', $configuration, '--no-restore')
}
$selfCheck = if ($Isolated) { '.xlwright/development/artifacts/bin/Spike0/release/Spike0.dll' } else { 'spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll' }
$artifactArguments = @()
Invoke-SpikeDotnet @($selfCheck, '--self-check')
