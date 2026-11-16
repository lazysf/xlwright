<#
.SYNOPSIS
Runs the VBA tests of a workbook and reports which passed and which failed.

.DESCRIPTION
Opens the workbook in a hidden Excel with macros enabled and events disabled, runs every test with
Invoke-ExcelMacro and closes the workbook without saving.

A test is a Sub in a standard module that takes no arguments, is not Private and whose name starts with
"Test_". The tests are read from the code of the workbook, sorted by module name and kept in their order
inside a module. -Name limits the run to the given tests ("<module>.<Sub>").

A test passes when it returns. It fails when it stops with a VBA runtime error, e.g. one it raises itself
with Err.Raise; Message is then the text of the error dialog, which shows on the screen for a moment.
When a test shows a dialog that cannot be closed (a Yes/No MsgBox, a compile error), the Excel instance of
this script is stopped: that test fails, the tests after it are reported as "not_run" and Error says why.
A MsgBox with only OK is closed and reported in Warnings.

After a runtime error the VBA project is reset, so module-level variables do not survive a failed test.

The result is written as UTF-8 JSON to -OutFile: Ok, Summary (Total, Passed, Failed, NotRun), Tests (Name,
Result "passed" / "failed" / "not_run", ElapsedMs, Message), Error and Warnings. Exit code is 0 only when
every test passed; no test at all is an error.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # Run only these tests, each as "<module>.<Sub>".
    [string[]]$Name,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

# powershell.exe -File passes "a,b" as one string, so split it here.
$Name = @($Name | ForEach-Object { $_ -split ',' } | Where-Object { $_ })

$tests = New-Object System.Collections.Generic.List[object]
$warnings = @()
$errorText = $null
try {
    foreach ($item in $Name) {
        if ($item -notmatch '^\w+\.\w+$') { throw "Not a test name (<module>.<Sub>): $item" }
    }
    Use-ExcelWorkbook -Path $Path -EnableMacros -TimeoutSeconds $TimeoutSeconds -WarningVariable +warnings -ScriptBlock {
        param($xl, $wb)
        $excel = Get-Process -Id (Get-ExcelProcessId $xl)
        $project = $wb.VBProject
        if ($project.Protection -ne 0) { throw 'The VBA project is locked; its tests cannot be listed.' }

        $found = @(foreach ($component in $project.VBComponents) {
            if ($component.Type -ne 1) { continue }  # vbext_ct_StdModule
            $module = $component.CodeModule
            $count = $module.CountOfLines
            if ($count -eq 0) { continue }
            # The VBE normalizes the code, so a declaration always starts its (continued) line.
            $code = [regex]::Replace($module.Lines(1, $count), ' _\r?\n', ' ')
            foreach ($line in ($code -split '\r?\n')) {
                # Private is left out on purpose: Application.Run would run a Private Sub by its qualified name.
                if ($line -match '^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+(Test_\w*)\s*\(\s*\)') {
                    $component.Name + '.' + $Matches[1]
                }
            }
        })
        # Stable sort by module name; the order inside a module is the order in the code.
        $index = 0
        $found = @($found | ForEach-Object { [pscustomobject]@{ Name = $_; Module = ($_ -split '\.')[0]; Index = $index++ } } |
            Sort-Object Module, Index | ForEach-Object { $_.Name })
        if ($Name.Count -gt 0) {
            $missing = @($Name | Where-Object { $found -notcontains $_ })
            if ($missing.Count -gt 0) { throw ('Test not found: ' + ($missing -join ', ')) }
            $found = @($found | Where-Object { $Name -contains $_ })
        }
        if ($found.Count -eq 0) { throw 'No test found: no Sub named Test_* without arguments in a standard module.' }
        foreach ($testName in $found) {
            $tests.Add([ordered]@{ Name = $testName; Result = 'not_run'; ElapsedMs = $null; Message = $null })
        }

        foreach ($test in $tests) {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            try {
                [void](Invoke-ExcelMacro -Excel $xl -Name $test.Name)
                $test.Result = 'passed'
            }
            catch {
                $test.Result = 'failed'
                $failure = $_.Exception.Data['XlWrightFailure']
                $test.Message = if ($failure) { [string]$_.Exception.Data['XlWrightDialogMessage'] } else { $_.Exception.Message }
                # Without this instance (stopped for a dialog, or by the watchdog) no further test can run.
                # The call fails a moment before a stopped process is gone, hence the wait.
                if ($failure -eq 'ExcelStopped' -or (-not $failure -and $excel.WaitForExit(2000))) { throw }
            }
            finally {
                $test.ElapsedMs = $watch.ElapsedMilliseconds
            }
        }
    }
}
catch {
    $errorText = $_.Exception.Message
}
finally {
    $count = { param($value) @($tests | Where-Object { $_.Result -eq $value }).Count }
    $summary = [ordered]@{
        Total = $tests.Count; Passed = (& $count 'passed'); Failed = (& $count 'failed'); NotRun = (& $count 'not_run')
    }
    $ok = (-not $errorText) -and $summary.Total -gt 0 -and $summary.Passed -eq $summary.Total
    $result = [ordered]@{
        Ok = $ok; Summary = $summary; Tests = $tests.ToArray(); Error = $errorText
        Warnings = @($warnings | ForEach-Object { [string]$_ })
    }
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($errorText) { Write-Error $errorText; exit 1 }
if (-not $ok) { exit 1 }
