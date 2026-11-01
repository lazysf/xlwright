#requires -Version 5.1
[CmdletBinding()]
param([switch] $Restore, [string] $TestFilter)
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$buildRoot = Join-Path $repository '.xlwright/mcp-build'
$artifactRoot = Join-Path $repository '.xlwright/development/artifacts'
foreach ($root in @($buildRoot, $artifactRoot)) {
    for ($ancestor = [IO.DirectoryInfo]::new($root); $null -ne $ancestor; $ancestor = $ancestor.Parent) {
        if ($ancestor.Exists -and $ancestor.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Build path ancestors may not be reparse points.'
        }
    }
    [IO.Directory]::CreateDirectory($root) | Out-Null
    foreach ($entry in Get-ChildItem -LiteralPath $root -Recurse -Force) {
        if ($entry.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Build workspaces may not contain reparse points.'
        }
    }
}
$dotnet = (Get-Command dotnet -CommandType Application).Source

function Invoke-DevelopmentDotnet([ValidatePattern('^[A-Za-z0-9_./:=*\-]+$')][string[]] $Arguments) {
    $info = [Diagnostics.ProcessStartInfo]::new($dotnet)
    $info.UseShellExecute = $false
    $info.WorkingDirectory = $repository
    # All callers use fixed tokens; reject values needing quoting on .NET Framework.
    $info.Arguments = $Arguments -join ' '
    $info.EnvironmentVariables['DOTNET_CLI_HOME'] = Join-Path $buildRoot 'dotnet-home'
    $info.EnvironmentVariables['APPDATA'] = Join-Path $buildRoot 'appdata'
    $info.EnvironmentVariables['NUGET_PACKAGES'] = Join-Path $buildRoot 'packages'
    $info.EnvironmentVariables['NUGET_HTTP_CACHE_PATH'] = Join-Path $buildRoot 'http-cache'
    $info.EnvironmentVariables['NUGET_PLUGINS_CACHE_PATH'] = Join-Path $buildRoot 'plugins-cache'
    $info.EnvironmentVariables['DOTNET_GENERATE_ASPNET_CERTIFICATE'] = 'false'
    $info.EnvironmentVariables['DOTNET_ADD_GLOBAL_TOOLS_TO_PATH'] = 'false'
    $info.EnvironmentVariables['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    $info.EnvironmentVariables['DOTNET_CLI_UI_LANGUAGE'] = 'en-US'
    $process = [Diagnostics.Process]::Start($info)
    try {
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "dotnet exited with code $($process.ExitCode)." }
    }
    finally { $process.Dispose() }
}

$testProject = 'tests/XlWright.McpServer.Tests/XlWright.McpServer.Tests.csproj'
# Built so its entry tests and prepare:product have an output, but never run here: its scenarios start Excel.
$realExcelProject = 'tests/XlWright.RealExcel.Tests/XlWright.RealExcel.Tests.csproj'
$artifactArguments = @('--artifacts-path', '.xlwright/development/artifacts')
if (-not $Restore) {
    $offlineConfig = Join-Path $buildRoot 'offline.NuGet.Config'
    [IO.File]::WriteAllText($offlineConfig, '<configuration><packageSources><clear /></packageSources></configuration>')
}
foreach ($project in @($testProject, $realExcelProject)) {
    if ($Restore) {
        Invoke-DevelopmentDotnet (@('restore', $project, '--locked-mode', '--configfile',
            'src/XlWright.McpServer/NuGet.Config') + $artifactArguments)
    }
    else {
        Invoke-DevelopmentDotnet (@('restore', $project, '--locked-mode', '--configfile',
            '.xlwright/mcp-build/offline.NuGet.Config', '-p:NuGetAudit=false') + $artifactArguments)
    }
    foreach ($configuration in @('Debug', 'Release')) {
        Invoke-DevelopmentDotnet (@('build', $project, '-c', $configuration, '--no-restore') + $artifactArguments)
    }
}
$testResults = '.xlwright/development/test-results.xml'
$testResultsPath = Join-Path $repository $testResults
if (Test-Path -LiteralPath $testResultsPath) { Remove-Item -LiteralPath $testResultsPath }
$testArguments = @('.xlwright/development/artifacts/bin/XlWright.McpServer.Tests/release/XlWright.McpServer.Tests.dll', '-result-xml', $testResults)
if ($TestFilter) { $testArguments += @('-method', $TestFilter) }
Invoke-DevelopmentDotnet $testArguments
# xUnit exits 0 when nothing ran, which would report a mistyped filter as a pass.
$executed = ([xml](Get-Content -LiteralPath $testResultsPath -Raw -Encoding UTF8)).assemblies.assembly.total
if ([int] $executed -eq 0) { throw 'No tests ran. Check the test filter.' }
Write-Output "Executed $executed tests."
