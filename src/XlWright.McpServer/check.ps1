#requires -Version 5.1
[CmdletBinding()]
param([switch] $Restore, [switch] $Isolated)
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$buildRoot = Join-Path $repository '.xlwright/mcp-build'
for ($ancestor = [IO.DirectoryInfo]::new($buildRoot); $null -ne $ancestor; $ancestor = $ancestor.Parent) {
    if ($ancestor.Exists -and $ancestor.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
        throw 'Build path ancestors may not be reparse points.'
    }
}
[IO.Directory]::CreateDirectory($buildRoot) | Out-Null
foreach ($entry in Get-ChildItem -LiteralPath $buildRoot -Recurse -Force) {
    if ($entry.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
        throw 'The build workspace may not contain reparse points.'
    }
}
$dotnet = (Get-Command dotnet -CommandType Application).Source
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

function Invoke-DevelopmentDotnet([ValidatePattern('^[A-Za-z0-9_./:=\-]+$')][string[]] $Arguments) {
    $info = [Diagnostics.ProcessStartInfo]::new($dotnet)
    $info.UseShellExecute = $false
    $info.WorkingDirectory = $repository
    # All callers use fixed tokens; reject values needing quoting on .NET Framework.
    $info.Arguments = @($Arguments + $artifactArguments) -join ' '
    $info.EnvironmentVariables['DOTNET_CLI_HOME'] = Join-Path $buildRoot 'dotnet-home'
    $info.EnvironmentVariables['APPDATA'] = Join-Path $buildRoot 'appdata'
    $info.EnvironmentVariables['NUGET_PACKAGES'] = Join-Path $buildRoot 'packages'
    $info.EnvironmentVariables['NUGET_HTTP_CACHE_PATH'] = Join-Path $buildRoot 'http-cache'
    $info.EnvironmentVariables['NUGET_PLUGINS_CACHE_PATH'] = Join-Path $buildRoot 'plugins-cache'
    $info.EnvironmentVariables['DOTNET_GENERATE_ASPNET_CERTIFICATE'] = 'false'
    $info.EnvironmentVariables['DOTNET_ADD_GLOBAL_TOOLS_TO_PATH'] = 'false'
    $info.EnvironmentVariables['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    $process = [Diagnostics.Process]::Start($info)
    try {
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "dotnet exited with code $($process.ExitCode)." }
    }
    finally { $process.Dispose() }
}

$checksProject = 'tests/XlWright.McpServer.Checks/XlWright.McpServer.Checks.csproj'
$offlineConfig = Join-Path $buildRoot 'offline.NuGet.Config'
if ($Isolated -and -not $Restore) {
    [IO.File]::WriteAllText($offlineConfig, '<configuration><packageSources><clear /></packageSources></configuration>')
    Invoke-DevelopmentDotnet @('restore', $checksProject, '--locked-mode', '--configfile',
        '.xlwright/mcp-build/offline.NuGet.Config', '-p:NuGetAudit=false')
}
elseif ($Restore) {
    Invoke-DevelopmentDotnet @('restore', $checksProject, '--locked-mode', '--configfile',
        'src/XlWright.McpServer/NuGet.Config')
}
foreach ($configuration in @('Debug', 'Release')) {
    Invoke-DevelopmentDotnet @('build', $checksProject, '-c', $configuration, '--no-restore')
}
$artifactArguments = @()
if ($Isolated) {
    Invoke-DevelopmentDotnet @('.xlwright/development/artifacts/bin/XlWright.McpServer.Checks/release/XlWright.McpServer.Checks.dll',
        '--server', '.xlwright/development/artifacts/bin/XlWright.McpServer/release/XlWright.McpServer.dll')
}
else {
    Invoke-DevelopmentDotnet @('tests/XlWright.McpServer.Checks/bin/Release/net10.0-windows/XlWright.McpServer.Checks.dll')
}

Write-Output 'Open the trusted repository root in Codex and verify dev_status using .codex/config.toml.'
