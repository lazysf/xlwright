# Development entry of the show command (docs/api.md, the show command section).
# It runs the server of the last checked build with: show --project <repository> --app <App>.
# Keep this file ASCII only, for Windows PowerShell 5.1.
param(
    [Parameter(Mandatory = $true)]
    [string]$App
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$latest = Get-Content -Raw -Path (Join-Path $repository '.xlwright\builds\latest.json') | ConvertFrom-Json
$server = Join-Path $repository $latest.server
& dotnet $server show --project $repository --app $App
exit $LASTEXITCODE
