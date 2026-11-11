<#
.SYNOPSIS
Checks the encoding and line endings of the text files in the repository.

.DESCRIPTION
Checks every file that Git tracks or would add (tracked, plus untracked files that are not ignored).
Files marked binary in .gitattributes, and symbolic links, are skipped.

Every text file must be valid UTF-8 without BOM, use LF line endings, and contain no control
characters other than tab and LF. Import-Vba.ps1 rejects BOM and CR, and stray control characters
are hard to see in review.

.ps1 and .psm1 files must be ASCII only, because Windows PowerShell 5.1 reads a script without BOM
in the ANSI code page. Test data under plugins/xlwright/skills/xlwright/tests/data/ is read as UTF-8 by
Invoke-ExcelScript.ps1, so it is exempt from the ASCII rule.

Exit code 0 when every file passes, 1 when a problem is found.
#>
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$asciiExempt = 'plugins/xlwright/skills/xlwright/tests/data/'
$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)

# git writes paths as UTF-8; the console default (CP932) would garble non-ASCII paths.
$savedEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = $strictUtf8
try {
    Push-Location -LiteralPath $root
    try {
        $listed = git ls-files -z --cached --others --exclude-standard
        if ($LASTEXITCODE -ne 0) { throw "git ls-files failed (exit $LASTEXITCODE)" }
        $paths = @(($listed -join '') -split "`0" | Where-Object { $_ } | Sort-Object -Unique)

        $attrs = $paths -join "`0" | git check-attr -z --stdin binary
        if ($LASTEXITCODE -ne 0) { throw "git check-attr failed (exit $LASTEXITCODE)" }
        # -z output is path, attribute, value repeated, each ending with NUL.
        $fields = ($attrs -join '') -split "`0"
        $binary = @{}
        for ($i = 0; $i + 2 -lt $fields.Count; $i += 3) {
            if ($fields[$i + 2] -eq 'set') { $binary[$fields[$i]] = $true }
        }
    }
    finally { Pop-Location }
}
finally { [Console]::OutputEncoding = $savedEncoding }

function Get-LineNumber([string]$Text, [int]$Index) {
    return ($Text.Substring(0, $Index).Split("`n").Count)
}

$problems = New-Object System.Collections.Generic.List[string]
$checked = 0
foreach ($path in $paths) {
    if ($binary.ContainsKey($path)) { continue }
    $full = Join-Path $root $path
    # A deleted but still tracked file has nothing to check.
    if (!(Test-Path -LiteralPath $full)) { continue }
    $item = Get-Item -LiteralPath $full -Force
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { continue }
    $checked++

    $bytes = [IO.File]::ReadAllBytes($full)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $problems.Add("${path}: has a UTF-8 BOM")
        continue
    }
    try { $text = $strictUtf8.GetString($bytes) }
    catch {
        $problems.Add("${path}: is not valid UTF-8")
        continue
    }

    $cr = $text.IndexOf("`r")
    if ($cr -ge 0) { $problems.Add("${path}:$(Get-LineNumber $text $cr): has CR (use LF line endings)") }

    $control = [regex]::Match($text, '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')
    if ($control.Success) {
        $code = '0x{0:X2}' -f [int][char]$control.Value
        $problems.Add("${path}:$(Get-LineNumber $text $control.Index): has control character $code")
    }

    $isScript = $path -match '\.(ps1|psm1)$'
    if ($isScript -and !$path.StartsWith($asciiExempt, [StringComparison]::Ordinal)) {
        $nonAscii = [regex]::Match($text, '[^\x00-\x7F]')
        if ($nonAscii.Success) {
            $problems.Add("${path}:$(Get-LineNumber $text $nonAscii.Index): has non-ASCII character (scripts must be ASCII)")
        }
    }
}

foreach ($p in $problems) { Write-Output $p }
Write-Output ("Checked $checked text files, found $($problems.Count) problems.")
if ($problems.Count -gt 0) { exit 1 }
exit 0
