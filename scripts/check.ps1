# The check entry point. Runs every check an agent needs, in the order that
# docs/architecture.md "サーバーとビルドとプラグイン" defines, and stops at the first failure.
# Requires Windows PowerShell 5.1, the .NET SDK in global.json, and `npm ci`.
#
# Usage (from any folder):
#   powershell -NoProfile -File scripts\check.ps1

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repo = Split-Path -Parent $PSScriptRoot
Push-Location $repo
try {
    function Invoke-Step([string] $name, [scriptblock] $action) {
        Write-Host "==> $name"
        & $action
        if ($LASTEXITCODE -ne 0) {
            throw "Step failed: $name (exit code $LASTEXITCODE)"
        }
    }

    # The build id is <short commit>[-dirty]-<UTC time>, as docs/api.md server_info defines.
    $commit = (git rev-parse --short HEAD).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'git rev-parse failed.' }
    $status = git status --porcelain
    if ($LASTEXITCODE -ne 0) { throw 'git status failed.' }
    $dirty = if ($status) { '-dirty' } else { '' }
    $time = (Get-Date).ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'")
    $buildId = "$commit$dirty-$time"
    $buildDir = Join-Path $repo ".xlwright\builds\$buildId"
    $serverDll = Join-Path $buildDir 'XlWright.McpServer.dll'
    Write-Host "Build id: $buildId"

    Invoke-Step 'Restore' { dotnet restore xlwright.slnx }

    Invoke-Step 'Build to the build folder' {
        dotnet build src\XlWright.McpServer\XlWright.McpServer.csproj --no-restore -c Release "-p:XlWrightBuildId=$buildId" -o $buildDir
    }

    Invoke-Step 'Tests without Excel' {
        dotnet test --solution xlwright.slnx --no-restore -c Release "-p:XlWrightBuildId=$buildId"
    }

    Invoke-Step 'Inspector: tools/list' {
        $output = npx --no-install mcp-inspector --cli --method tools/list -- dotnet $serverDll --project $repo --no-excel | Out-String
        if ($LASTEXITCODE -ne 0) { return }
        $names = @((ConvertFrom-Json $output).tools | ForEach-Object { $_.name })
        if ($names -notcontains 'server_info') {
            throw "tools/list does not contain server_info: $output"
        }
    }

    Invoke-Step 'Inspector: server_info' {
        $output = npx --no-install mcp-inspector --cli --method tools/call --tool-name server_info -- dotnet $serverDll --project $repo --no-excel | Out-String
        if ($LASTEXITCODE -ne 0) { return }
        $info = (ConvertFrom-Json $output).structuredContent
        if ($info.name -ne 'xlwright' -or $info.version -ne '0.1.0' -or $info.build_id -ne $buildId) {
            throw "server_info does not match this build ($buildId): $output"
        }
    }

    Invoke-Step 'Markdown lint' { npm run lint:md }

    Invoke-Step 'git diff --check' { git diff --check HEAD }

    # Record the build that passed. UTF-8 without BOM, which Windows PowerShell 5.1 Set-Content cannot write.
    $latest = [ordered]@{
        build_id = $buildId
        server = ".xlwright/builds/$buildId/XlWright.McpServer.dll"
    } | ConvertTo-Json
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText((Join-Path $repo '.xlwright\builds\latest.json'), $latest, $utf8)

    # Point the development plugin at this build. Claude Code expands the ${...} placeholders.
    $mcp = [ordered]@{
        mcpServers = [ordered]@{
            xlwright = [ordered]@{
                command = 'dotnet'
                args = @(
                    ('${CLAUDE_PLUGIN_ROOT}/../../.xlwright/builds/' + $buildId + '/XlWright.McpServer.dll'),
                    '--project',
                    '${CLAUDE_PROJECT_DIR}'
                )
            }
        }
    } | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText((Join-Path $repo 'plugins\xlwright-dev\.mcp.json'), $mcp, $utf8)
    Write-Host "All checks passed. Build $buildId is recorded in .xlwright\builds\latest.json and plugins\xlwright-dev\.mcp.json."
}
finally {
    Pop-Location
}
