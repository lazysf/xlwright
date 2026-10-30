#requires -Version 7.0
[CmdletBinding()]
param([switch] $Restore)
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

function Invoke-DevelopmentDotnet([string[]] $Arguments) {
    $info = [Diagnostics.ProcessStartInfo]::new($dotnet)
    $info.UseShellExecute = $false
    $info.WorkingDirectory = $repository
    foreach ($argument in $Arguments) { $info.ArgumentList.Add($argument) }
    $info.Environment['DOTNET_CLI_HOME'] = Join-Path $buildRoot 'dotnet-home'
    $info.Environment['APPDATA'] = Join-Path $buildRoot 'appdata'
    $info.Environment['NUGET_PACKAGES'] = Join-Path $buildRoot 'packages'
    $info.Environment['NUGET_HTTP_CACHE_PATH'] = Join-Path $buildRoot 'http-cache'
    $info.Environment['NUGET_PLUGINS_CACHE_PATH'] = Join-Path $buildRoot 'plugins-cache'
    $info.Environment['DOTNET_GENERATE_ASPNET_CERTIFICATE'] = 'false'
    $info.Environment['DOTNET_ADD_GLOBAL_TOOLS_TO_PATH'] = 'false'
    $info.Environment['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    $process = [Diagnostics.Process]::Start($info)
    try {
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "dotnet exited with code $($process.ExitCode)." }
    }
    finally { $process.Dispose() }
}

$checksProject = 'tests/XlWright.McpServer.Checks/XlWright.McpServer.Checks.csproj'
if ($Restore) {
    Invoke-DevelopmentDotnet @('restore', $checksProject, '--locked-mode', '--configfile',
        'src/XlWright.McpServer/NuGet.Config')
}
foreach ($configuration in @('Debug', 'Release')) {
    Invoke-DevelopmentDotnet @('build', $checksProject, '-c', $configuration, '--no-restore')
}
Invoke-DevelopmentDotnet @('tests/XlWright.McpServer.Checks/bin/Release/net10.0-windows/XlWright.McpServer.Checks.dll')

Write-Output 'Open the trusted repository root in Codex and verify dev_status using .codex/config.toml.'
