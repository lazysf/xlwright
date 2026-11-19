<#
.SYNOPSIS
Runs eval cases through codex exec and checks what the agent read.

.DESCRIPTION
For each case under plugins/xlwright/evals/, sends the body of prompt.md followed by the sentence in
tools/data/codex-check-suffix.txt to codex exec. Each run uses a new empty folder, the read-only
sandbox, and no saved session. The prompt, the events (JSONL), the last message, and stderr are
written to plugins/xlwright/evals/results/codex/<time>/<case>/, which Git ignores, and a summary to
summary.md next to the case folders.

Graders of type tool_used are checked against the commands in the events, because Codex reads files
with shell commands instead of the Read and Skill tools of Claude Code:
- tool Read: the commands that match input_match.
- tool Skill: the commands that open skills/xlwright/SKILL.md.
min and max apply as in claude plugin eval (min defaults to 1). Graders of type llm are not judged
here; summary.md puts the last message next to the criterion for a person or Claude to judge.

Codex must already have xlwright installed (DEVELOPMENT.md). The script stops when codex plugin list
does not show it installed and enabled.

-Case all runs every case that has a prompt.md. -FromResults <run folder> checks a previous run
again without calling Codex.

Exit code 0 when every run finished and every tool_used grader passed, 1 otherwise.
#>
[CmdletBinding()]
param(
    [string[]]$Case = @('userform', 'not-tool', 'refused-script', 'safety'),
    [string]$Codex,
    [int]$TimeoutSec = 600,
    [string]$FromResults
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$evalDir = Join-Path $root 'plugins\xlwright\evals'
$suffixPath = Join-Path $PSScriptRoot 'data\codex-check-suffix.txt'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$skillPattern = 'skills[\\/]+xlwright[\\/]+SKILL\.md'

function Read-Utf8([string]$Path) {
    return [IO.File]::ReadAllText($Path, $utf8)
}

function Split-FrontMatter([string]$Text) {
    # Reads only the flat "key: value" lines that the eval files use.
    $lines = ($Text -replace "`r", '') -split "`n"
    $meta = @{}
    if ($lines.Count -lt 2 -or $lines[0] -ne '---') { return @{ Meta = $meta; Body = $Text.Trim() } }
    $end = -1
    for ($i = 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -eq '---') { $end = $i; break }
    }
    if ($end -lt 0) { throw 'The front matter is not closed with ---.' }
    for ($i = 1; $i -lt $end; $i++) {
        if ($lines[$i] -match '^([A-Za-z_]+):\s*(.*)$') {
            $key = $Matches[1]
            $value = $Matches[2].Trim()
            if ($value -match "^'(.*)'$") { $value = $Matches[1] -replace "''", "'" }
            $meta[$key] = $value
        }
    }
    $body = ''
    if ($end -lt $lines.Count - 1) { $body = ($lines[($end + 1)..($lines.Count - 1)] -join "`n").Trim() }
    return @{ Meta = $meta; Body = $body }
}

function Find-Codex {
    $package = Get-AppxPackage -Name 'OpenAI.Codex' | Sort-Object Version -Descending | Select-Object -First 1
    if (!$package) { throw 'The ChatGPT app (OpenAI.Codex) is not installed. Pass -Codex <path to codex.exe>.' }
    $path = Join-Path $package.InstallLocation 'app\resources\codex.exe'
    if (!(Test-Path -LiteralPath $path)) { throw "codex.exe is not at $path. Pass -Codex <path to codex.exe>." }
    return $path
}

function Get-InstalledRow([string]$Exe) {
    $out = & $Exe plugin list 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "codex plugin list failed (exit $LASTEXITCODE)." }
    return ($out -split "`r?`n") | Where-Object { $_ -match '^\s*xlwright@\S+\s+installed, enabled' } |
        Select-Object -First 1
}

function Invoke-CodexRun([string]$Exe, [string]$PromptFile, [string]$WorkDir, [string]$OutDir) {
    $events = Join-Path $OutDir 'events.jsonl'
    $last = Join-Path $OutDir 'last.md'
    $stderr = Join-Path $OutDir 'stderr.txt'
    # cmd hands the prompt file to stdin as bytes, so it stays UTF-8. A pipe from Windows
    # PowerShell would re-encode it with $OutputEncoding.
    $line = '"' + $Exe + '" exec --json -s read-only --ephemeral --skip-git-repo-check -C "' + $WorkDir +
        '" -o "' + $last + '" - < "' + $PromptFile + '" > "' + $events + '" 2> "' + $stderr + '"'
    $proc = Start-Process -FilePath $env:ComSpec -ArgumentList @('/d', '/s', '/c', ('"' + $line + '"')) `
        -WorkingDirectory $WorkDir -NoNewWindow -PassThru
    # Reading Handle keeps ExitCode available after the process ends.
    $null = $proc.Handle
    if (!$proc.WaitForExit($TimeoutSec * 1000)) {
        # Stops only the process tree started here.
        & taskkill.exe /T /F /PID $proc.Id | Out-Null
        throw "codex exec did not finish in $TimeoutSec seconds."
    }
    return $proc.ExitCode
}

function Read-Events([string]$Path) {
    $commands = New-Object System.Collections.Generic.List[string]
    $usage = $null
    foreach ($line in [IO.File]::ReadAllLines($Path, $utf8)) {
        if (!$line.Trim()) { continue }
        $event = $line | ConvertFrom-Json
        if ($event.type -eq 'item.completed' -and $event.item -and $event.item.type -eq 'command_execution') {
            $commands.Add([string]$event.item.command)
        }
        elseif ($event.type -eq 'turn.completed') {
            $usage = $event.usage
        }
    }
    return @{ Commands = $commands; Usage = $usage }
}

function Test-ToolGrader($Meta, $Commands) {
    $tool = $Meta['tool']
    if ($tool -eq 'Skill') {
        $pattern = $skillPattern
    }
    elseif ($tool -eq 'Read' -and $Meta['input_match']) {
        $pattern = $Meta['input_match']
    }
    else {
        return @{ Passed = $false; Detail = "no Codex mapping for tool '$tool'" }
    }
    $count = @($Commands | Where-Object { $_ -cmatch $pattern }).Count
    $min = 1
    if ($Meta.ContainsKey('min')) { $min = [int]$Meta['min'] }
    $max = [int]::MaxValue
    if ($Meta.ContainsKey('max')) { $max = [int]$Meta['max'] }
    $range = "min $min"
    if ($max -ne [int]::MaxValue) { $range += ", max $max" }
    return @{ Passed = ($count -ge $min -and $count -le $max); Detail = "$count command(s) match '$pattern' ($range)" }
}

$savedEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = $utf8
try {
    if ($FromResults) {
        $runDir = (Resolve-Path -LiteralPath $FromResults).Path
        $cases = @(Get-ChildItem -LiteralPath $runDir -Directory |
            Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'events.jsonl') } |
            ForEach-Object { $_.Name })
        if ($cases.Count -eq 0) { throw "No <case>\events.jsonl under $runDir." }
    }
    else {
        $cases = $Case
        if ($cases -contains 'all') {
            $cases = @(Get-ChildItem -LiteralPath $evalDir -Directory |
                Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'prompt.md') } |
                ForEach-Object { $_.Name })
        }
        foreach ($name in $cases) {
            if (!(Test-Path -LiteralPath (Join-Path $evalDir "$name\prompt.md"))) { throw "No eval case '$name'." }
        }
        if (!$Codex) { $Codex = Find-Codex }
        $row = Get-InstalledRow $Codex
        if (!$row) {
            Write-Output 'xlwright is not installed and enabled in Codex. Install the candidate first (DEVELOPMENT.md).'
            exit 1
        }
        Write-Output "Codex plugin: $($row.Trim())"
        $runDir = Join-Path $evalDir ('results\codex\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType Directory -Path $runDir | Out-Null
    }

    $suffix = Read-Utf8 $suffixPath
    $allPassed = $true
    $report = New-Object System.Collections.Generic.List[string]
    $report.Add('# Codex check')
    $report.Add('')
    $report.Add("Run folder: $runDir")

    foreach ($name in $cases) {
        $caseOut = Join-Path $runDir $name
        $caseDir = Join-Path $evalDir $name
        $prompt = Split-FrontMatter (Read-Utf8 (Join-Path $caseDir 'prompt.md'))
        $exitFile = Join-Path $caseOut 'exit-code.txt'

        if (!$FromResults) {
            New-Item -ItemType Directory -Path $caseOut | Out-Null
            $promptFile = Join-Path $caseOut 'prompt.txt'
            [IO.File]::WriteAllText($promptFile, $prompt.Body + "`n" + $suffix, $utf8)
            $work = Join-Path ([IO.Path]::GetTempPath()) ('xlwright-codex-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $work | Out-Null
            try {
                Write-Output "Running $name ..."
                $exit = Invoke-CodexRun $Codex $promptFile $work $caseOut
            }
            finally {
                Remove-Item -LiteralPath $work -Recurse -Force
            }
            [IO.File]::WriteAllText($exitFile, [string]$exit, $utf8)
        }

        $report.Add('')
        $report.Add("## $name")
        $report.Add('')
        $report.Add("Expected: $($prompt.Meta['expected_outcome'])")
        $report.Add('')

        $exitText = 'unknown'
        if (Test-Path -LiteralPath $exitFile) { $exitText = (Read-Utf8 $exitFile).Trim() }
        $eventsPath = Join-Path $caseOut 'events.jsonl'
        if ($exitText -ne '0' -or !(Test-Path -LiteralPath $eventsPath)) {
            $allPassed = $false
            $report.Add("- FAIL: codex exec exit code $exitText (see stderr.txt)")
            Write-Output "[$name] FAIL: codex exec exit code $exitText"
            continue
        }

        $events = Read-Events $eventsPath
        if ($events.Usage) {
            $report.Add("- Tokens: input $($events.Usage.input_tokens) (cached $($events.Usage.cached_input_tokens)), output $($events.Usage.output_tokens)")
        }
        $report.Add("- Commands: $($events.Commands.Count)")
        foreach ($command in $events.Commands) {
            $shown = $command
            if ($shown.Length -gt 240) { $shown = $shown.Substring(0, 240) + ' ...' }
            $report.Add("  - ``$shown``")
        }

        $judge = New-Object System.Collections.Generic.List[string]
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $caseDir 'graders') -Filter '*.md') {
            $grader = Split-FrontMatter (Read-Utf8 $file.FullName)
            $type = $grader.Meta['type']
            if ($type -eq 'tool_used') {
                $result = Test-ToolGrader $grader.Meta $events.Commands
                $mark = 'PASS'
                if (!$result.Passed) { $mark = 'FAIL'; $allPassed = $false }
                $report.Add("- ${mark}: $($file.BaseName) - $($result.Detail)")
                Write-Output "[$name] ${mark}: $($file.BaseName) - $($result.Detail)"
            }
            elseif ($type -eq 'llm') {
                $report.Add("- JUDGE: $($file.BaseName) (criterion below)")
                Write-Output "[$name] JUDGE: $($file.BaseName)"
                $judge.Add("### $($file.BaseName)")
                $judge.Add('')
                $judge.Add($grader.Body)
                $judge.Add('')
            }
            else {
                $allPassed = $false
                $report.Add("- FAIL: $($file.BaseName) - unknown grader type '$type'")
                Write-Output "[$name] FAIL: $($file.BaseName) - unknown grader type '$type'"
            }
        }

        if ($judge.Count -gt 0) {
            $report.Add('')
            $report.AddRange($judge)
        }
        $lastPath = Join-Path $caseOut 'last.md'
        $report.Add('')
        $report.Add('### Last message')
        $report.Add('')
        if (Test-Path -LiteralPath $lastPath) { $report.Add((Read-Utf8 $lastPath).Trim()) } else { $report.Add('(none)') }
    }

    $summary = Join-Path $runDir 'summary.md'
    [IO.File]::WriteAllText($summary, (($report -join "`n") + "`n"), $utf8)
    Write-Output "Summary: $summary"
    if ($allPassed) { exit 0 }
    exit 1
}
finally {
    [Console]::OutputEncoding = $savedEncoding
}
