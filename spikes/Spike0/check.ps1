#requires -Version 7.0
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
$buildRoot = Join-Path $buildParent ([Guid]::NewGuid().ToString('N'))
if (Test-Path -LiteralPath $buildRoot) { throw 'The build directory must be new.' }
[IO.Directory]::CreateDirectory($buildRoot) | Out-Null
$configPath = Join-Path $buildRoot 'NuGet.Config'
[IO.File]::WriteAllText($configPath,
    '<configuration><packageSources><clear /></packageSources></configuration>')
$dotnetPath = (Get-Command dotnet -CommandType Application).Source

function Invoke-SpikeDotnet([string[]] $DotnetArguments) {
    $processInfo = [Diagnostics.ProcessStartInfo]::new($dotnetPath)
    $processInfo.UseShellExecute = $false
    $processInfo.WorkingDirectory = $repositoryRoot
    foreach ($argument in $DotnetArguments) { $processInfo.ArgumentList.Add($argument) }
    # Child-process settings keep SDK initialization and NuGet away from personal config.
    $processInfo.Environment['DOTNET_CLI_HOME'] = Join-Path $buildRoot 'dotnet-home'
    $processInfo.Environment['APPDATA'] = Join-Path $buildRoot 'appdata'
    $processInfo.Environment['NUGET_PACKAGES'] = Join-Path $buildRoot 'packages'
    $processInfo.Environment['DOTNET_GENERATE_ASPNET_CERTIFICATE'] = 'false'
    $processInfo.Environment['DOTNET_ADD_GLOBAL_TOOLS_TO_PATH'] = 'false'
    $processInfo.Environment['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    $processInfo.Environment['DOTNET_NOLOGO'] = '1'
    $childProcess = [Diagnostics.Process]::Start($processInfo)
    try {
        $childProcess.WaitForExit()
        if ($childProcess.ExitCode -ne 0) {
            throw "dotnet exited with code $($childProcess.ExitCode)."
        }
    }
    finally { $childProcess.Dispose() }
}

Invoke-SpikeDotnet @('restore', 'spikes/Spike0/Spike0.csproj', '--configfile', $configPath, '-p:NuGetAudit=false')
foreach ($configuration in @('Debug', 'Release')) {
    Invoke-SpikeDotnet @('build', 'spikes/Spike0/Spike0.csproj', '--configuration', $configuration, '--no-restore')
}
Invoke-SpikeDotnet @('spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll', '--self-check')
