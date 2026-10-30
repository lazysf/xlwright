#requires -Version 7.0
[CmdletBinding()]
param([ValidateSet('Codex', 'Claude')] [string] $Client = 'Codex')
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$configurationPath = if ($Client -eq 'Codex') { Join-Path $repository '.codex/config.toml' } else { Join-Path $repository '.mcp.json' }
$configurationDirectory = [IO.Path]::GetDirectoryName($configurationPath)
$configurationName = [IO.Path]::GetRelativePath($repository, $configurationPath).Replace('\', '/')
$server = Join-Path $PSScriptRoot 'bin/Release/net10.0-windows/XlWright.McpServer.dll'
$spike = Join-Path $repository 'spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll'

function Assert-DevelopmentPath([string] $Path) {
    # Inspect existing ancestors before creating or opening any destination.
    for ($current = $Path; $null -ne $current -and $current.Length -gt 0; $current = [IO.Path]::GetDirectoryName($current)) {
        if ((Test-Path -LiteralPath $current) -and
            (Get-Item -LiteralPath $current -Force).Attributes.HasFlag([IO.FileAttributes]::ReparsePoint)) {
            throw 'Development setup paths may not use reparse points.'
        }
    }
}

Assert-DevelopmentPath $configurationPath
Assert-DevelopmentPath $server
Assert-DevelopmentPath $spike
if (Test-Path -LiteralPath $configurationPath) {
    throw "Project $configurationName already exists. Preserve it and review integration manually."
}
if (-not (Test-Path -LiteralPath $server -PathType Leaf) -or
    -not (Test-Path -LiteralPath $spike -PathType Leaf)) {
    throw 'Build Spike0 and the MCP server before running development setup.'
}
$dotnet = (Get-Command dotnet -CommandType Application).Source
if (-not [IO.Path]::GetFileName($dotnet).Equals('dotnet.exe', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Development setup requires the dotnet.exe application host.'
}
Assert-DevelopmentPath $dotnet

if ($Client -eq 'Codex') {
    # JSON string encoding is a subset of TOML basic strings for these slash-normalized
    # Windows paths. Formatting inserts each encoded value once, without evaluation.
    $values = @($dotnet, $server, $repository) | ForEach-Object {
        ConvertTo-Json -InputObject $_.Replace('\', '/') -Compress
    }
    $template = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'codex-config.toml.template'))
    $configuration = [string]::Format([Globalization.CultureInfo]::InvariantCulture, $template, [object[]]$values)
}
else {
    # Claude Code expands ${VAR} in command and args, so the launched path could differ from the checked one.
    if ($dotnet.Contains('$') -or $server.Contains('$')) {
        throw 'Claude Code setup does not support "$" in the dotnet or server path.'
    }
    # Claude Code documents no working-directory field; the server refuses to start without Spike0 in it.
    $configuration = ConvertTo-Json -Depth 4 -InputObject ([ordered]@{
        mcpServers = [ordered]@{
            'xlwright-dev' = [ordered]@{
                type = 'stdio'
                command = $dotnet.Replace('\', '/')
                args = @($server.Replace('\', '/'), '--development')
            }
        }
    })
    $configuration = $configuration.Replace("`r`n", "`n") + "`n"
}
[IO.Directory]::CreateDirectory($configurationDirectory) | Out-Null
Assert-DevelopmentPath $configurationPath
# CreateNew preserves an existing configuration even if it appeared after the check.
$stream = [IO.FileStream]::new($configurationPath, [IO.FileMode]::CreateNew,
    [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($configuration)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush($true)
}
finally { $stream.Dispose() }
Write-Output "Created project $configurationName. Reopen the trusted project and verify dev_status."
