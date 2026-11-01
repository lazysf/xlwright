#requires -Version 5.1
# S13 one-off probe: a minimal stdio MCP server that never touches Excel. It records how an MCP client stops it:
# whether standard input reaches end of file, and how long the process survives afterwards. A hard kill shows as
# alive marks that stop before the final exit mark. Not used by npm run check or by the product server.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][guid] $ExperimentId,
    # Seconds to stay alive after end of file, writing a mark every 100 ms.
    [ValidateRange(1, 30)][int] $LingerSeconds = 5
)

$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$directory = Join-Path $repository ".xlwright/mcp-shutdown-probe/$($ExperimentId.ToString('N'))"
[IO.Directory]::CreateDirectory($directory) | Out-Null
$recordPath = Join-Path $directory "$([guid]::NewGuid().ToString('N')).jsonl"
$utf8 = New-Object Text.UTF8Encoding $false

function Write-Mark([string] $name, [hashtable] $details = @{}) {
    $details['at_utc'] = [DateTime]::UtcNow.ToString('O')
    $details['name'] = $name
    [IO.File]::AppendAllText($recordPath, ((New-Object PSObject -Property $details | ConvertTo-Json -Compress) + "`n"), $utf8)
}

$self = [Diagnostics.Process]::GetCurrentProcess()
$parent = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID").ParentProcessId
Write-Mark 'started' @{
    pid = $PID; parent_pid = $parent; started_at_utc = $self.StartTime.ToUniversalTime().ToString('O')
    powershell_version = $PSVersionTable.PSVersion.ToString()
    script_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
}

$stdin = New-Object IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = [Console]::OpenStandardOutput()
function Send([string] $json) {
    $bytes = $utf8.GetBytes($json + "`n")
    $stdout.Write($bytes, 0, $bytes.Length)
    $stdout.Flush()
}

while ($null -ne ($line = $stdin.ReadLine())) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    try { $message = $line | ConvertFrom-Json } catch { Write-Mark 'unparsed'; continue }
    Write-Mark 'received' @{ method = $message.method; has_id = $null -ne $message.id }
    if ($null -eq $message.id) { continue }
    $id = $message.id | ConvertTo-Json -Compress
    switch ($message.method) {
        'initialize' {
            $version = $message.params.protocolVersion | ConvertTo-Json -Compress
            Send "{`"jsonrpc`":`"2.0`",`"id`":$id,`"result`":{`"protocolVersion`":$version,`"capabilities`":{`"tools`":{}},`"serverInfo`":{`"name`":`"xlwright-shutdown-probe`",`"version`":`"0.0.1`"}}}"
        }
        'tools/list' { Send "{`"jsonrpc`":`"2.0`",`"id`":$id,`"result`":{`"tools`":[]}}" }
        'ping' { Send "{`"jsonrpc`":`"2.0`",`"id`":$id,`"result`":{}}" }
        default { Send "{`"jsonrpc`":`"2.0`",`"id`":$id,`"error`":{`"code`":-32601,`"message`":`"Method not found`"}}" }
    }
}

Write-Mark 'eof'
$timer = [Diagnostics.Stopwatch]::StartNew()
for ($mark = 1; $timer.Elapsed.TotalSeconds -lt $LingerSeconds; $mark++) {
    Start-Sleep -Milliseconds 100
    Write-Mark 'alive' @{ mark = $mark; elapsed_ms = [int]$timer.Elapsed.TotalMilliseconds }
}
Write-Mark 'exit'
exit 0
