<#
.SYNOPSIS
Checks the xlwright scripts against the real Excel on this machine.

.DESCRIPTION
Every case creates its own workbooks in a new temporary folder and checks explicit expected values.
Nothing outside that folder is changed. The default cases keep Excel hidden. -Full adds the cases
that show Excel and the VBE on the screen (VBA compile error, macro dialogs, failing VBA tests, ribbon click, ribbon layout, publish with VBA lock, publish without the test modules); do not touch
those windows while they run. These cases also check that Excel never takes the foreground.

Preconditions: no Excel process is running, and AccessVBOM is 1. Otherwise nothing runs (exit 2).
After each case no Excel process may remain. If one does, the run stops without touching it.

-Case runs only the named cases, e.g. -Case timeout,userform.

The result is written as UTF-8 JSON to -OutFile and summarized on the console.
Exit code: 0 all passed, 1 a check failed, 2 preconditions not met or the run was stopped.
The temporary folder is removed at the end, whether the cases passed or not, unless -KeepTemp is given.
#>
param(
    [switch]$Full,
    # Run only these cases (names as printed). Cases that need -Full still need it.
    [string[]]$Case,
    [string]$OutFile,
    [switch]$KeepTemp
)
$ErrorActionPreference = 'Stop'
# powershell.exe -File passes "a,b" as one string, so split it here.
$Case = @($Case | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
$skill = Split-Path -Parent $PSScriptRoot
$scripts = Join-Path $skill 'scripts'
$data = Join-Path $PSScriptRoot 'data'
$module = Join-Path $scripts 'XlWright.psm1'
$runner = Join-Path $scripts 'Invoke-ExcelScript.ps1'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$S = [IO.File]::ReadAllText((Join-Path $data 'strings.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json

if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output 'Excel is running. Close it (or ask the user) before running the tests.'
    exit 2
}
$vbom = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name AccessVBOM `
    -ErrorAction SilentlyContinue).AccessVBOM
if ($vbom -ne 1) {
    Write-Output 'AccessVBOM is not 1. Ask the user to trust access to the VBA project object model.'
    exit 2
}

Import-Module $module -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
$temp = Join-Path ([IO.Path]::GetTempPath()) ('xlwright-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
$results = New-Object System.Collections.Generic.List[object]
$script:checks = $null
$stopped = $false

function Assert([string]$Label, [bool]$Condition, $Detail) {
    $script:checks.Add([ordered]@{ Check = $Label; Ok = $Condition; Detail = $(if ($Condition) { $null } else { [string]$Detail }) })
}

function New-CaseDir([string]$Name) {
    $dir = Join-Path $temp $Name
    New-Item -ItemType Directory -Path $dir | Out-Null
    $dir
}

function Invoke-Script([string]$File, [string[]]$Arguments) {
    $ErrorActionPreference = 'Continue'
    $output = & powershell.exe -NoProfile -File $File @Arguments 2>&1 | Out-String
    [pscustomobject]@{ Exit = $LASTEXITCODE; Output = $output }
}

function Invoke-Runner([string]$Book, [string]$Name, [string[]]$Extra = @()) {
    $json = Join-Path $temp ($Name + '-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script $runner (@('-Path', $Book, '-Script', (Join-Path $data "scripts\$Name.ps1"), '-OutFile', $json) + $Extra)
    $parsed = $null
    if (Test-Path -LiteralPath $json) { $parsed = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Output = $run.Output }
}

function Get-FileHashText([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }

function Wait-NoExcel([int]$Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    do {
        $left = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        if ($left.Count -eq 0) { return @() }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    $left
}

if (-not ('XlWrightTest.Foreground' -as [type])) {
    Add-Type -Namespace XlWrightTest -Name Foreground -MemberDefinition @'
[DllImport("user32.dll")]
public static extern System.IntPtr GetForegroundWindow();
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
}

# Samples the foreground window every 50 ms in another runspace. The tests start only when no Excel is
# running, so an Excel in the foreground is one this run started.
function Start-ForegroundWatch {
    $state = [hashtable]::Synchronized(@{ Stop = $false; Total = 0; Excel = 0 })
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({ param($s)
        while (-not $s.Stop) {
            $processId = [uint32]0
            [void][XlWrightTest.Foreground]::GetWindowThreadProcessId([XlWrightTest.Foreground]::GetForegroundWindow(), [ref]$processId)
            $s.Total++
            if ((Get-Process -Id $processId -ErrorAction SilentlyContinue).ProcessName -eq 'EXCEL') { $s.Excel++ }
            Start-Sleep -Milliseconds 50
        }
    }).AddArgument($state)
    [pscustomobject]@{ State = $state; PowerShell = $ps; Async = $ps.BeginInvoke() }
}

function Stop-ForegroundWatch($Watch) {
    $Watch.State.Stop = $true
    try { [void]$Watch.PowerShell.EndInvoke($Watch.Async) } finally { $Watch.PowerShell.Dispose() }
    $Watch.State
}

function Invoke-Case([string]$Name, [switch]$WatchForeground, [scriptblock]$Body) {
    if ($script:stopped) { return }
    if ($Case -and $Case -notcontains $Name) { return }
    $script:checks = New-Object System.Collections.Generic.List[object]
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $errorText = $null
    $foreground = $null
    if ($WatchForeground) { $foreground = Start-ForegroundWatch }
    try { & $Body } catch { $errorText = $_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage }
    finally {
        if ($foreground) {
            $samples = Stop-ForegroundWatch $foreground
            # The ribbon is operated through MSAA so that the user can keep working in another application.
            Assert 'Excel never took the foreground' ($samples.Excel -eq 0) ("$($samples.Excel) of $($samples.Total) samples")
        }
    }
    $left = Wait-NoExcel 20
    if ($left.Count -gt 0) {
        Assert 'no Excel left after the case' $false ('PIDs ' + ($left -join ','))
        $script:stopped = $true
    }
    if ($errorText) { Assert 'case ran without an unexpected error' $false $errorText }
    $failed = @($script:checks | Where-Object { -not $_.Ok })
    $results.Add([ordered]@{
        # ToArray, not @(): in PS 5.1, @() of this list inside the hashtable makes List.Add fail.
        Case = $Name; Ok = ($failed.Count -eq 0); Seconds = [int]$watch.Elapsed.TotalSeconds; Checks = $script:checks.ToArray()
    })
    $mark = if ($failed.Count -eq 0) { 'PASS' } else { 'FAIL' }
    Write-Output ('{0} {1} ({2} s, {3} checks)' -f $mark, $Name, [int]$watch.Elapsed.TotalSeconds, $script:checks.Count)
    foreach ($f in $failed) { Write-Output ('     - ' + $f.Check + ': ' + $f.Detail) }
}

function New-BookWithVba([string]$Dir) {
    $book = Join-Path $Dir 'book.xlsm'
    [void](New-ExcelWorkbook -Path $book)
    Copy-Item -LiteralPath (Join-Path $data 'vba') -Destination (Join-Path $Dir 'vba') -Recurse
    $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', (Join-Path $Dir 'vba'))
    Assert 'Import-Vba exit 0' ($import.Exit -eq 0) $import.Output
    $book
}

function Invoke-Compile([string]$Book) {
    $json = Join-Path $temp ('compile-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Test-VbaCompile.ps1') @('-Path', $Book, '-OutFile', $json)
    $parsed = $null
    if (Test-Path -LiteralPath $json) { $parsed = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Output = $run.Output }
}

# ---------------------------------------------------------------------------------------------
Invoke-Case 'cells-and-vba' {
    $dir = New-CaseDir 'cells-and-vba'
    $book = New-BookWithVba $dir
    Use-ExcelWorkbook -Path $book -Save -ScriptBlock {
        param($xl, $wb)
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = $S.Sheet
        $ws.Range('A1').Value2 = [double]2
        $ws.Range('B1').Formula = '=A1*3'
    }
    $r = Use-ExcelWorkbook -Path $book -EnableMacros -ScriptBlock {
        param($xl, $wb)
        $ws = $wb.Worksheets.Item(1)
        [ordered]@{
            Sheet = $ws.Name; A1 = $ws.Range('A1').Value2; B1 = $ws.Range('B1').Formula; B1Value = $ws.Range('B1').Value2
            Add = $xl.Run('AddNums', 2, 3); Greeting = $xl.Run('Greeting')
            Version = [System.__ComObject].InvokeMember('BookVersion', [Reflection.BindingFlags]::InvokeMethod, $null, $wb, $null)
            Components = (@($wb.VBProject.VBComponents | ForEach-Object { $_.Name + ':' + $_.Type }) -join ',')
        }
    }
    # In-process use: this process keeps running, so Use-ExcelWorkbook itself must end its Excel.
    $left = Wait-NoExcel 3
    Assert 'Excel ended when Use-ExcelWorkbook returned' ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
    Assert 'sheet name (Japanese)' ($r.Sheet -eq $S.Sheet) $r.Sheet
    Assert 'A1 = 2' ($r.A1 -eq 2) $r.A1
    Assert 'B1 formula' ($r.B1 -eq '=A1*3') $r.B1
    Assert 'B1 recalculated = 6' ($r.B1Value -eq 6) $r.B1Value
    Assert 'AddNums(2,3) = 5' ($r.Add -eq 5) $r.Add
    Assert 'Japanese string from VBA' ($r.Greeting -eq $S.Greeting) $r.Greeting
    Assert 'ThisWorkbook code replaced' ($r.Version -eq '1.1') $r.Version
    Assert 'components' ($r.Components -eq 'ThisWorkbook:100,Sheet1:100,CalcOptions:2,Util:1') $r.Components
    $export = Invoke-Script (Join-Path $scripts 'Export-Vba.ps1') @('-Path', $book, '-OutDir', (Join-Path $dir 'export'))
    Assert 'Export-Vba exit 0' ($export.Exit -eq 0) $export.Output
    foreach ($rel in 'standard\Util.bas', 'classes\CalcOptions.cls', 'documents\ThisWorkbook.cls') {
        $a = Get-FileHashText (Join-Path $dir ('vba\' + $rel))
        $exported = Join-Path $dir ('export\' + $rel)
        $b = if (Test-Path -LiteralPath $exported) { Get-FileHashText $exported } else { 'missing' }
        Assert ('export round trip identical: ' + $rel) ($a -eq $b) $b
    }
}

Invoke-Case 'invoke-excelscript' {
    $book = Join-Path (New-CaseDir 'invoke-excelscript') 'book.xlsm'
    $w = Invoke-Runner $book 'write' @('-Create', '-Save')
    Assert 'write exit 0' ($w.Exit -eq 0) $w.Output
    Assert 'write result (Japanese)' ($w.Json.Ok -and $w.Json.Result[0] -eq $S.WriteDone) ($w.Json | ConvertTo-Json -Compress)
    $before = Get-FileHashText $book
    $r = Invoke-Runner $book 'read'
    $v = $r.Json.Result
    Assert 'read exit 0' ($r.Exit -eq 0) $r.Output
    Assert 'read values' ($v[0].A1 -eq 2 -and $v[0].B1 -eq '=A1*3' -and $v[0].B1Value -eq 6) ($v[0] | ConvertTo-Json -Compress)
    Assert '2D array as rows' ($v[0].Rows[0][1] -eq 6 -and $v[0].Rows[1][0] -eq $S.Text) ($v[0].Rows | ConvertTo-Json -Compress)
    Assert 'COM object replaced' ($v[1] -eq '[COM object]') $v[1]
    Assert 'sheet list' ((@($v | Select-Object -Skip 2) -join ',') -eq $S.Sheet) (@($v | Select-Object -Skip 2) -join ',')
    Assert 'read leaves no warnings' (@($r.Json.Warnings).Count -eq 0) (@($r.Json.Warnings) -join ' | ')
    # Reading a sheet keeps COM wrappers alive; the call must still end its Excel before it returns.
    $left = Wait-NoExcel 3
    Assert 'Excel ended when the call returned' ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
    $f = Invoke-Runner $book 'fail'
    Assert 'failure exit 1' ($f.Exit -eq 1) $f.Exit
    Assert 'failure reported' ((-not $f.Json.Ok) -and $f.Json.Error) ($f.Json | ConvertTo-Json -Compress)
    Assert 'workbook unchanged by read and failure' ((Get-FileHashText $book) -eq $before) 'hash changed'
}

Invoke-Case 'vba-compile' {
    $dir = New-CaseDir 'vba-compile'
    $book = New-BookWithVba $dir
    $hash = Get-FileHashText $book
    $c = Invoke-Compile $book
    Assert 'valid code: exit 0' ($c.Exit -eq 0) $c.Output
    Assert 'valid code: compiled, no error' ($c.Json.Compiled -eq $true -and $null -eq $c.Json.CompileError -and $null -eq $c.Json.Error) ($c.Json | ConvertTo-Json -Compress)
    Assert 'workbook unchanged by the check' ((Get-FileHashText $book) -eq $hash) 'hash changed'
    $plain = Join-Path $dir 'plain.xlsx'
    [void](New-ExcelWorkbook -Path $plain)
    $p = Invoke-Compile $plain
    Assert 'workbook without VBA: compiled' ($p.Exit -eq 0 -and $p.Json.Compiled -eq $true) ($p.Json | ConvertTo-Json -Compress)
}

function New-BookWithMacros([string]$Dir, [string]$Source) {
    $book = Join-Path $Dir ((Split-Path -Leaf $Source) + '.xlsm')
    [void](New-ExcelWorkbook -Path $book)
    $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', (Join-Path $data $Source))
    Assert ($Source + ': import exit 0') ($import.Exit -eq 0) $import.Output
    $book
}

Invoke-Case 'macro' {
    $book = New-BookWithMacros (New-CaseDir 'macro') 'vba-macro'
    $r = Use-ExcelWorkbook -Path $book -EnableMacros -ScriptBlock {
        param($xl, $wb)
        Invoke-ExcelMacro -Excel $xl -Name 'Macros.CountCall'
        Invoke-ExcelMacro -Excel $xl -Name 'Macros.CountCall'
        [ordered]@{
            Sum = Invoke-ExcelMacro -Excel $xl -Name 'Macros.AddThree' -ArgumentList 1, 2, 3
            Calls = Invoke-ExcelMacro -Excel $xl -Name 'Macros.CallCount'
            Slow = Invoke-ExcelMacro -Excel $xl -Name 'Macros.Slow'
        }
    }
    Assert 'only the values of the functions are returned' (@($r).Count -eq 1) (@($r).Count)
    Assert 'arguments passed, value returned' ($r.Sum -eq 6) $r.Sum
    Assert 'Sub ran twice' ($r.Calls -eq 2) $r.Calls
    Assert 'a macro that takes time is waited for' ($r.Slow -eq 7) $r.Slow
}

function Invoke-VbaTest([string]$Book, [string[]]$Extra = @()) {
    $json = Join-Path $temp ('vba-test-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Invoke-VbaTest.ps1') (@('-Path', $Book, '-OutFile', $json) + $Extra)
    $parsed = $null
    if (Test-Path -LiteralPath $json) { $parsed = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Output = $run.Output }
}

Invoke-Case 'vba-test' {
    $dir = New-CaseDir 'vba-test'
    $book = New-BookWithMacros $dir 'vba-test\pass'
    $hash = Get-FileHashText $book
    $t = Invoke-VbaTest $book
    $detail = $t.Json | ConvertTo-Json -Compress -Depth 5
    Assert 'all passed: exit 0 and Ok' ($t.Exit -eq 0 -and $t.Json.Ok -eq $true -and $null -eq $t.Json.Error) $detail
    # Sorted by module, in the order of the code; no Private Sub, Function, Sub with arguments or other name.
    $names = @($t.Json.Tests | ForEach-Object { $_.Name }) -join ','
    Assert 'only the tests are listed, in order' ($names -eq 'Alpha.Test_Second,Alpha.Test_First,Test_Calc.Test_Add,Test_Calc.Test_ImplicitPublic') $names
    Assert 'every test ran and passed' (@($t.Json.Tests | Where-Object { $_.Result -ne 'passed' -or $null -eq $_.ElapsedMs }).Count -eq 0) $detail
    Assert 'summary counts' ($t.Json.Summary.Total -eq 4 -and $t.Json.Summary.Passed -eq 4 -and $t.Json.Summary.Failed -eq 0 -and $t.Json.Summary.NotRun -eq 0) $detail
    Assert 'workbook unchanged by the run' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    $one = Invoke-VbaTest $book @('-Name', 'Test_Calc.Test_Add,Alpha.Test_First')
    $names = @($one.Json.Tests | ForEach-Object { $_.Name }) -join ','
    Assert '-Name: only the named tests ran' ($one.Exit -eq 0 -and $names -eq 'Alpha.Test_First,Test_Calc.Test_Add') ($one.Json | ConvertTo-Json -Compress -Depth 5)

    # A Private Sub must not become runnable by naming it.
    $private = Invoke-VbaTest $book @('-Name', 'Test_Calc.Test_Private')
    Assert '-Name: a name that is not a test is an error' ($private.Exit -eq 1 -and $private.Json.Ok -eq $false -and [string]$private.Json.Error -like 'Test not found: Test_Calc.Test_Private*' -and @($private.Json.Tests).Count -eq 0) ($private.Json | ConvertTo-Json -Compress -Depth 5)

    $empty = Join-Path $dir 'empty.xlsm'
    [void](New-ExcelWorkbook -Path $empty)
    $none = Invoke-VbaTest $empty
    Assert 'no test: an error, not a pass' ($none.Exit -eq 1 -and $none.Json.Ok -eq $false -and [string]$none.Json.Error -like 'No test found*') ($none.Json | ConvertTo-Json -Compress -Depth 5)

    $hang = Invoke-VbaTest (New-BookWithMacros $dir 'vba-test\hang') @('-TimeoutSeconds', '20')
    $detail = $hang.Json | ConvertTo-Json -Compress -Depth 5
    $results = @($hang.Json.Tests | ForEach-Object { $_.Result }) -join ','
    Assert 'timeout: the test that hangs failed, the rest was not run' ($hang.Exit -eq 1 -and $results -eq 'passed,failed,not_run') $detail
    Assert 'timeout: Error says so' ([string]$hang.Json.Error -like 'Timed out after 20 s*') $detail
}

function Invoke-Compare([string]$Book, [string]$SourceDir) {
    $json = Join-Path $temp ('compare-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Compare-Vba.ps1') @('-Path', $Book, '-SourceDir', $SourceDir, '-OutFile', $json)
    $parsed = $null
    if (Test-Path -LiteralPath $json) { $parsed = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Output = $run.Output }
}

Invoke-Case 'vba-compare' {
    $dir = New-CaseDir 'vba-compare'
    $book = New-BookWithVba $dir
    $source = Join-Path $dir 'vba'
    $hash = Get-FileHashText $book
    $c = Invoke-Compare $book $source
    $detail = $c.Json | ConvertTo-Json -Compress -Depth 5
    $rows = @($c.Json.Components | ForEach-Object { $_.Kind + '/' + $_.Name + '=' + $_.Status }) -join ','
    Assert 'imported sources: exit 0 and in sync' ($c.Exit -eq 0 -and $c.Json.InSync -eq $true -and $null -eq $c.Json.Error) $detail
    Assert 'every component is listed, by kind and name' ($rows -eq 'standard/Util=same,classes/CalcOptions=same,documents/ThisWorkbook=same') $rows
    Assert 'workbook unchanged by the comparison' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    # Edited sources: one kind of difference per component.
    $edited = Join-Path $dir 'edited'
    Copy-Item -LiteralPath $source -Destination $edited -Recurse
    $util = Join-Path $edited 'standard\Util.bas'
    $text = [IO.File]::ReadAllText($util, $utf8)
    [IO.File]::WriteAllText($util, $text.Replace('AddNums = a + b', 'AddNums = a - b'), $utf8)
    $book2 = Join-Path $edited 'documents\ThisWorkbook.cls'
    $text = [IO.File]::ReadAllText($book2, $utf8)
    [IO.File]::WriteAllText($book2, $text.Replace('Public Function BookVersion', 'Public Function BOOKVERSION'), $utf8)
    Remove-Item -LiteralPath (Join-Path $edited 'classes\CalcOptions.cls')
    [IO.File]::WriteAllText((Join-Path $edited 'standard\Extra.bas'), "Attribute VB_Name = ""Extra""`nOption Explicit`n", $utf8)
    $editedHash = Get-FileHashText $util
    $e = Invoke-Compare $book $edited
    $detail = $e.Json | ConvertTo-Json -Compress -Depth 5
    $rows = @($e.Json.Components | ForEach-Object { $_.Kind + '/' + $_.Name + '=' + $_.Status }) -join ','
    Assert 'edited sources: exit 1 and not in sync' ($e.Exit -eq 1 -and $e.Json.InSync -eq $false -and $null -eq $e.Json.Error) $detail
    Assert 'status of every component' ($rows -eq 'standard/Extra=source_only,standard/Util=different,classes/CalcOptions=workbook_only,documents/ThisWorkbook=case_only') $rows
    $first = @($e.Json.Components | Where-Object { $_.Name -eq 'Util' })[0].FirstDifference
    Assert 'different: the first differing line of the source file' ($first.Line -eq 6 -and $first.Source -eq '    AddNums = a - b' -and $first.Workbook -eq '    AddNums = a + b') ($first | ConvertTo-Json -Compress)
    Assert 'sources unchanged by the comparison' ((Get-FileHashText $util) -eq $editedHash) 'hash changed'

    # What the VBE or an editor changes without changing the code is not a difference.
    $spaced = Join-Path $dir 'spaced'
    Copy-Item -LiteralPath $source -Destination $spaced -Recurse
    $util = Join-Path $spaced 'standard\Util.bas'
    $text = [IO.File]::ReadAllText($util, $utf8)
    [IO.File]::WriteAllText($util, ($text.Replace("""Util""`n", """Util""`n`n`n").TrimEnd("`n") + "`n`n`n"), $utf8)
    $p = Invoke-Compare $book $spaced
    Assert 'blank lines after the attributes and at the end are not a difference' ($p.Exit -eq 0 -and $p.Json.InSync -eq $true) ($p.Json | ConvertTo-Json -Compress -Depth 5)
    $text = [IO.File]::ReadAllText($util, $utf8)
    [IO.File]::WriteAllText($util, $text.Replace("End Function`n`nPublic Function Greeting", "End Function`n`n`nPublic Function Greeting"), $utf8)
    $q = Invoke-Compare $book $spaced
    $first = @($q.Json.Components | Where-Object { $_.Name -eq 'Util' })[0].FirstDifference
    Assert 'a blank line inside the code is a difference, at its line in the file' ($q.Exit -eq 1 -and $first.Line -eq 11 -and $first.Source -eq '') ($q.Json | ConvertTo-Json -Compress -Depth 5)

    $missing = Invoke-Compare $book (Join-Path $dir 'no-such-dir')
    Assert 'missing source directory: an error' ($missing.Exit -eq 1 -and $missing.Json.InSync -eq $false -and [string]$missing.Json.Error -like 'Source directory not found*') ($missing.Json | ConvertTo-Json -Compress -Depth 5)
}

Invoke-Case 'vba-prune' {
    $dir = New-CaseDir 'vba-prune'
    $book = New-BookWithVba $dir
    $source = Join-Path $dir 'vba'
    $importer = Join-Path $scripts 'Import-Vba.ps1'
    # Components that are not in the sources: a UserForm and a standard module (uf-create), a standard
    # module and the code of a sheet (imported from another folder).
    $created = Invoke-Runner $book 'uf-create' @('-Save')
    Assert 'UserForm added to the workbook' ($created.Exit -eq 0) $created.Output
    $more = Join-Path $dir 'more'
    New-Item -ItemType Directory -Path (Join-Path $more 'standard'), (Join-Path $more 'documents') | Out-Null
    [IO.File]::WriteAllText((Join-Path $more 'standard\Extra.bas'), "Attribute VB_Name = ""Extra""`nOption Explicit`n", $utf8)
    [IO.File]::WriteAllText((Join-Path $more 'documents\Sheet1.cls'), "Attribute VB_Name = ""Sheet1""`nOption Explicit`n", $utf8)
    $added = Invoke-Script $importer @('-Path', $book, '-SourceDir', $more)
    Assert 'more components imported' ($added.Exit -eq 0) $added.Output
    $extraRows = 'standard/Extra=workbook_only,standard/FormProbe=workbook_only,documents/Sheet1=workbook_only,forms/frmInfo=workbook_only'
    $rowsOf = { param($c) @($c.Json.Components | Where-Object { $_.Status -ne 'same' } | ForEach-Object { $_.Kind + '/' + $_.Name + '=' + $_.Status }) -join ',' }

    $plain = Invoke-Script $importer @('-Path', $book, '-SourceDir', $source)
    $rows = & $rowsOf (Invoke-Compare $book $source)
    Assert 'without -Prune: exit 0 and nothing reported as removed' ($plain.Exit -eq 0 -and $plain.Output -notmatch 'removed|cleared') $plain.Output
    Assert 'without -Prune: the components that are not in the sources stay' ($rows -eq $extraRows) $rows

    # A failed import must not leave a pruned workbook behind.
    $bad = Join-Path $dir 'bad'
    Copy-Item -LiteralPath $source -Destination $bad -Recurse
    [IO.File]::WriteAllText((Join-Path $bad 'documents\NoSuchSheet.cls'), "Attribute VB_Name = ""NoSuchSheet""`nOption Explicit`n", $utf8)
    $hash = Get-FileHashText $book
    $failed = Invoke-Script $importer @('-Path', $book, '-SourceDir', $bad, '-Prune')
    Assert '-Prune with a source that cannot be imported: exit 1' ($failed.Exit -eq 1 -and $failed.Output -match 'No document module named NoSuchSheet') $failed.Output
    Assert '-Prune with a source that cannot be imported: workbook unchanged' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    $pruned = Invoke-Script $importer @('-Path', $book, '-SourceDir', $source, '-Prune')
    $lines = @($pruned.Output -split "\r?\n" | Where-Object { $_ -match '^(removed|cleared document module): ' } | Sort-Object) -join ','
    Assert '-Prune: exit 0' ($pruned.Exit -eq 0) $pruned.Output
    Assert '-Prune: reports what it removed and cleared, and nothing else' ($lines -eq 'cleared document module: Sheet1,removed: Extra,removed: FormProbe,removed: frmInfo') $pruned.Output
    $after = Invoke-Compare $book $source
    Assert '-Prune: the saved workbook is in sync with the sources' ($after.Exit -eq 0 -and $after.Json.InSync -eq $true -and @($after.Json.Components).Count -eq 3) ($after.Json | ConvertTo-Json -Compress -Depth 5)

    $again = Invoke-Script $importer @('-Path', $book, '-SourceDir', $source, '-Prune')
    Assert '-Prune with nothing to remove: exit 0 and nothing reported' ($again.Exit -eq 0 -and $again.Output -notmatch 'removed|cleared') $again.Output
}

Invoke-Case 'timeout' {
    $book = Join-Path (New-CaseDir 'timeout') 'book.xlsm'
    [void](New-ExcelWorkbook -Path $book)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $h = Invoke-Runner $book 'hang' @('-EnableMacros', '-TimeoutSeconds', '20')
    $seconds = [int]$watch.Elapsed.TotalSeconds
    Assert 'timeout exit 1' ($h.Exit -eq 1) $h.Exit
    Assert 'timeout message' ([string]$h.Json.Error -like 'Timed out after 20 s*') $h.Json.Error
    Assert 'stopped within 60 s' ($seconds -lt 60) $seconds
}

Invoke-Case 'caller-killed' {
    $dir = New-CaseDir 'caller-killed'
    $book = Join-Path $dir 'book.xlsm'
    [void](New-ExcelWorkbook -Path $book)
    $pidFile = Join-Path $dir 'excel.pid'
    $child = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-File', ('"' + (Join-Path $data 'scripts\orphan-child.ps1') + '"'),
        '-Module', ('"' + $module + '"'), '-Book', ('"' + $book + '"'), '-PidFile', ('"' + $pidFile + '"'))
    $deadline = (Get-Date).AddSeconds(60)
    while (-not (Test-Path -LiteralPath $pidFile) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    Assert 'child started Excel' (Test-Path -LiteralPath $pidFile) 'no PID file'
    if (-not (Test-Path -LiteralPath $pidFile)) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue; return }
    $excelPid = [int][IO.File]::ReadAllText($pidFile)
    Stop-Process -Id $child.Id -Force
    Start-Sleep -Seconds 2
    Assert 'Excel survives the killed caller' ($null -ne (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) 'already gone'
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Process -Id $excelPid -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }
    Assert 'watchdog stopped Excel at the limit' ($null -eq (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) ('PID ' + $excelPid)
}

Invoke-Case 'input-rejected' {
    $dir = New-CaseDir 'input-rejected'
    $book = Join-Path $dir 'book.xlsm'
    [void](New-ExcelWorkbook -Path $book)
    $hash = Get-FileHashText $book
    $ascii = New-Object System.Text.UTF8Encoding($false)
    $sources = @(
        @{ Name = 'bom'; Reason = '*without BOM and LF*'; Bytes = ([byte[]](0xEF, 0xBB, 0xBF) + $ascii.GetBytes("Attribute VB_Name = ""Bad""`nOption Explicit`n")) },
        @{ Name = 'crlf'; Reason = '*without BOM and LF*'; Bytes = $ascii.GetBytes("Attribute VB_Name = ""Bad""`r`nOption Explicit`r`n") },
        @{ Name = 'not-cp932'; Reason = '*CP932*'; Bytes = $ascii.GetBytes("Attribute VB_Name = ""Bad""`n' " + [char]::ConvertFromUtf32(0x1F600) + "`n") },
        @{ Name = 'name-mismatch'; Reason = '*Unexpected extension or VB_Name*'; Bytes = $ascii.GetBytes("Attribute VB_Name = ""Other""`nOption Explicit`n") })
    foreach ($s in $sources) {
        $src = Join-Path $dir ('src-' + $s.Name)
        New-Item -ItemType Directory -Path (Join-Path $src 'standard') -Force | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $src 'standard\Bad.bas'), $s.Bytes)
        $r = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', $src)
        Assert ('Import-Vba rejects ' + $s.Name) ($r.Exit -ne 0 -and $r.Output -like $s.Reason) $r.Output
    }
    Assert 'workbook unchanged by rejected imports' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    $badXml = Join-Path $dir 'bad.xml'
    [IO.File]::WriteAllText($badXml, '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon>', $ascii)
    $r = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', $badXml)
    Assert 'Add-CustomUI rejects malformed XML' ($r.Exit -ne 0) $r.Output
    Assert 'workbook unchanged by rejected ribbon' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    $pubDir = Join-Path $dir 'publish'
    New-Item -ItemType Directory -Path (Join-Path $pubDir 'vba\standard') -Force | Out-Null
    $pubBook = Join-Path $pubDir 'book.xlsm'
    Copy-Item -LiteralPath $book -Destination $pubBook
    [IO.File]::WriteAllBytes((Join-Path $pubDir 'vba\standard\Bad.bas'), $sources[0].Bytes)
    $out = Join-Path $dir 'dist'
    $r = Invoke-Script (Join-Path $scripts 'Publish-Workbook.ps1') @('-Path', $pubBook, '-OutDir', $out)
    Assert 'publish fails on a bad source' ($r.Exit -ne 0 -and $r.Output -like '*without BOM and LF*') $r.Output
    Assert 'work book unchanged by failed publish' ((Get-FileHashText $pubBook) -eq $hash) 'hash changed'
    $leftovers = @(Get-ChildItem -LiteralPath $pubDir -File -Force | Where-Object { $_.Name -like '.excel-publish-*' -or $_.Name -like '.customui-*' })
    if (Test-Path -LiteralPath $out) { $leftovers += @(Get-ChildItem -LiteralPath $out -File -Force) }
    Assert 'no output or temporary file after failed publish' ($leftovers.Count -eq 0) (($leftovers | ForEach-Object { $_.Name }) -join ',')
}

Invoke-Case 'watchdog-unavailable' {
    $book = Join-Path (New-CaseDir 'watchdog-unavailable') 'book.xlsx'
    [void](New-ExcelWorkbook -Path $book)
    $excelModule = Get-Module XlWright
    $mocks = @(
        @{ Name = 'WMI call throws'; Mock = { throw 'simulated: access denied' } },
        @{ Name = 'WMI returns an error code'; Mock = { [pscustomobject]@{ ReturnValue = 2 } } })
    foreach ($m in $mocks) {
        # Shadow the cmdlet inside the module only; nothing on the system is changed.
        & $excelModule { param($body) Set-Item -Path function:script:Invoke-CimMethod -Value $body } $m.Mock
        $message = $null
        try {
            Use-ExcelWorkbook -Path $book -ScriptBlock { param($xl, $wb) 'should not run' } | Out-Null
        }
        catch { $message = $_.Exception.Message }
        # Remove-Item ignores the script: qualifier, so remove by the plain path and check it is gone.
        finally { & $excelModule { Remove-Item -Path function:Invoke-CimMethod -ErrorAction SilentlyContinue } }
        $restored = & $excelModule { (Get-Command Invoke-CimMethod).CommandType }
        Assert ($m.Name + ': mock removed') ([string]$restored -eq 'Cmdlet') $restored
        Assert ($m.Name + ': stops with the watchdog error') ([string]$message -like 'Could not start the watchdog*') $message
        Assert ($m.Name + ': says the skill cannot be used') ([string]$message -like '*cannot be used*') $message
        $left = Wait-NoExcel 20
        Assert ($m.Name + ': Excel closed') ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
    }
}

function Invoke-Prerequisite([string[]]$Arguments) {
    $json = Join-Path $temp ('prerequisite-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Test-Prerequisite.ps1') (@('-OutFile', $json) + $Arguments)
    $parsed = $null
    if (Test-Path -LiteralPath $json) { $parsed = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Output = $run.Output }
}

Invoke-Case 'prerequisite' {
    $dir = New-CaseDir 'prerequisite'
    $book = Join-Path $dir 'book.xlsm'
    [void](New-ExcelWorkbook -Path $book)
    $hash = Get-FileHashText $book

    # This run starts only when no Excel is running and AccessVBOM is 1.
    $ready = Invoke-Prerequisite @('-Vba')
    $detail = $ready.Json | ConvertTo-Json -Compress -Depth 6
    Assert 'nothing in the way: exit 0, Ready, no problem' ($ready.Exit -eq 0 -and $ready.Json.Ready -eq $true -and @($ready.Json.Problems).Count -eq 0 -and $null -eq $ready.Json.Error) $detail
    Assert 'no Excel, AccessVBOM enabled, WMI usable, no workbook' ($ready.Json.Excel.Running -eq $false -and @($ready.Json.Excel.Processes).Count -eq 0 -and $ready.Json.AccessVBOM.Value -eq 1 -and $ready.Json.AccessVBOM.Enabled -eq $true -and $ready.Json.Wmi.Ok -eq $true -and $null -eq $ready.Json.Workbook) $detail

    $closed = Invoke-Prerequisite @('-Path', $book)
    $detail = $closed.Json | ConvertTo-Json -Compress -Depth 6
    Assert 'closed workbook: exit 0, exists, not open' ($closed.Exit -eq 0 -and $closed.Json.Workbook.Path -eq $book -and $closed.Json.Workbook.Exists -eq $true -and $closed.Json.Workbook.OpenElsewhere -eq $false) $detail
    $missing = Invoke-Prerequisite @('-Path', (Join-Path $dir 'no-such.xlsm'))
    Assert 'missing workbook: exit 1 and a problem' ($missing.Exit -eq 1 -and $missing.Json.Ready -eq $false -and $missing.Json.Workbook.Exists -eq $false -and @($missing.Json.Problems).Count -eq 1 -and [string]$missing.Json.Problems[0] -like 'Workbook not found:*') ($missing.Json | ConvertTo-Json -Compress -Depth 6)

    # A hidden Excel that has the workbook open is a running Excel and leaves the owner file next to it.
    $open = Use-ExcelWorkbook -Path $book -ScriptBlock {
        param($xl, $wb)
        $r = Invoke-Prerequisite @('-Path', $book)
        [pscustomobject]@{ Exit = $r.Exit; Json = $r.Json; ExcelId = (Get-ExcelProcessId $xl) }
    }
    $detail = $open.Json | ConvertTo-Json -Compress -Depth 6
    Assert 'Excel running: exit 1, not Ready' ($open.Exit -eq 1 -and $open.Json.Ready -eq $false) $detail
    Assert 'Excel running: the process is listed' ($open.Json.Excel.Running -eq $true -and @($open.Json.Excel.Processes | Where-Object { $_.Id -eq $open.ExcelId }).Count -eq 1) $detail
    Assert 'open workbook: reported as open' ($open.Json.Workbook.OpenElsewhere -eq $true) $detail
    $problems = @($open.Json.Problems)
    Assert 'Excel running and open workbook: one problem each' ($problems.Count -eq 2 -and [string]$problems[0] -like 'Excel is running (PID*do not stop it.' -and [string]$problems[1] -like 'The workbook seems to be open:*') $detail
    Assert 'workbook unchanged by the checks' ((Get-FileHashText $book) -eq $hash) 'hash changed'
    # The process is gone a moment after Use-ExcelWorkbook returns; the checks below expect no Excel.
    $left = Wait-NoExcel 20
    Assert 'Excel closed before the next checks' ($left.Count -eq 0) ('PIDs ' + ($left -join ','))

    # What cannot be produced on this machine is shadowed inside the module only; nothing on the system is changed.
    $excelModule = Get-Module XlWright
    $mocked = {
        param([string]$Command, [scriptblock]$Mock, [scriptblock]$Body)
        & $excelModule { param($name, $body) Set-Item -Path ('function:script:' + $name) -Value $body } $Command $Mock
        try { & $Body }
        finally { & $excelModule { param($name) Remove-Item -Path ('function:' + $name) -ErrorAction SilentlyContinue } $Command }
        $restored = & $excelModule { param($name) (Get-Command $name).CommandType } $Command
        Assert ($Command + ': mock removed') ([string]$restored -eq 'Cmdlet') $restored
    }
    & $mocked 'Get-ItemProperty' { [pscustomobject]@{ AccessVBOM = 0 } } {
        $sheet = Get-ExcelPrerequisite
        $vba = Get-ExcelPrerequisite -Vba
        Assert 'AccessVBOM 0 without -Vba: reported, not a problem' ($sheet.Ready -eq $true -and $sheet.AccessVBOM.Value -eq 0 -and $sheet.AccessVBOM.Enabled -eq $false) ($sheet | ConvertTo-Json -Compress -Depth 6)
        Assert 'AccessVBOM 0 with -Vba: a problem' ($vba.Ready -eq $false -and @($vba.Problems).Count -eq 1 -and [string]$vba.Problems[0] -like 'AccessVBOM is not 1*do not change the setting yourself.') ($vba | ConvertTo-Json -Compress -Depth 6)
    }
    # A policy value is taken before the one of the Trust Center; the user policy before the machine policy.
    & $mocked 'Get-ItemProperty' {
        param($Path)
        if ($Path -like 'HKLM:\Software\Policies\*') { [pscustomobject]@{ AccessVBOM = 0 } }
        elseif ($Path -like 'HKCU:\Software\Microsoft\*') { [pscustomobject]@{ AccessVBOM = 1 } }
    } {
        $r = Get-ExcelPrerequisite -Vba
        Assert 'AccessVBOM 0 by machine policy, 1 by the user: a problem' ($r.Ready -eq $false -and $r.AccessVBOM.Value -eq 0 -and $r.AccessVBOM.Enabled -eq $false) ($r | ConvertTo-Json -Compress -Depth 6)
    }
    & $mocked 'Get-ItemProperty' {
        param($Path)
        if ($Path -like 'HKCU:\Software\Policies\*') { [pscustomobject]@{ AccessVBOM = 1 } }
        elseif ($Path -like 'HKLM:\Software\Policies\*') { [pscustomobject]@{ AccessVBOM = 0 } }
    } {
        $r = Get-ExcelPrerequisite -Vba
        Assert 'AccessVBOM 1 by user policy, 0 by machine policy: no problem' ($r.Ready -eq $true -and $r.AccessVBOM.Value -eq 1 -and $r.AccessVBOM.Enabled -eq $true) ($r | ConvertTo-Json -Compress -Depth 6)
    }
    & $mocked 'Invoke-CimMethod' { throw 'simulated: access denied' } {
        $r = Get-ExcelPrerequisite
        Assert 'WMI call throws: a problem with the reason' ($r.Ready -eq $false -and $r.Wmi.Ok -eq $false -and $r.Wmi.Error -eq 'simulated: access denied' -and [string]$r.Problems[0] -like '*cannot be used: simulated: access denied*') ($r | ConvertTo-Json -Compress -Depth 6)
    }
    & $mocked 'Invoke-CimMethod' { [pscustomobject]@{ ReturnValue = 2 } } {
        $r = Get-ExcelPrerequisite
        Assert 'WMI returns an error code: a problem with the code' ($r.Ready -eq $false -and $r.Wmi.Ok -eq $false -and $r.Wmi.Error -eq 'Win32_Process.Create returned 2.') ($r | ConvertTo-Json -Compress -Depth 6)
    }

    # Git: only where git is installed. The repository is made inside the temporary folder.
    if (Get-Command git -CommandType Application -ErrorAction SilentlyContinue) {
        $git = { param([string[]]$Arguments) $ErrorActionPreference = 'Continue'; & git -C $dir @Arguments 2>$null | Out-Null; $LASTEXITCODE }
        $state = { $g = (Get-ExcelPrerequisite -Path $book).Workbook.Git; '' + $g.Available + ',' + $g.InRepository + ',' + $g.Tracked + ',' + $g.Uncommitted }
        [void](& $git @('init', '-q'))
        $s = & $state
        Assert 'git: untracked workbook' ($s -eq 'True,True,False,') $s
        [void](& $git @('add', '--', 'book.xlsm'))
        $s = & $state
        Assert 'git: added workbook has an uncommitted change' ($s -eq 'True,True,True,True') $s
        Assert 'git: never a problem' ((Get-ExcelPrerequisite -Path $book).Ready -eq $true) 'not Ready'
    }
}

function Invoke-Summary([string]$Book) {
    $json = Join-Path $temp ('summary-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Get-WorkbookSummary.ps1') @('-Path', $Book, '-OutFile', $json)
    $parsed = $null
    $text = ''
    if (Test-Path -LiteralPath $json) {
        $text = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8)
        $parsed = $text | ConvertFrom-Json
    }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Text = $text; Output = $run.Output }
}

Invoke-Case 'workbook-summary' {
    $dir = New-CaseDir 'workbook-summary'
    $other = Join-Path $dir 'other.xlsx'
    [void](New-ExcelWorkbook -Path $other)
    [IO.File]::WriteAllText((Join-Path $dir 'data.csv'), "a,b`r`n1,2`r`n", $utf8)
    $book = Join-Path $dir 'book.xlsm'
    $built = Invoke-Runner $book 'summary-build' @('-Create', '-Save')
    Assert 'workbook built' ($built.Exit -eq 0) ($built.Json | ConvertTo-Json -Compress)
    $ribbon = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', (Join-Path $data 'ribbon\base.xml'))
    Assert 'ribbon embedded' ($ribbon.Exit -eq 0) $ribbon.Output
    $hash = Get-FileHashText $book

    $summary = Invoke-Summary $book
    $detail = $summary.Text
    Assert 'exit 0, no error, path' ($summary.Exit -eq 0 -and $null -eq $summary.Json.Error -and $summary.Json.Path -eq $book) $detail
    # In the order of the workbook. An empty sheet and a chart sheet have no used range; a sheet with a
    # value only in A1 has.
    $sheets = @($summary.Json.Sheets | ForEach-Object { $_.Name + ':' + $_.Kind + ':' + $_.Visible + ':' + $_.UsedRange + ':' + $_.Rows + 'x' + $_.Columns }) -join ','
    # Worksheets.Add puts a sheet before the active one, and a sheet that was hidden is not the active one.
    $expected = 'Hidden:worksheet:hidden::0x0,VeryHidden:worksheet:veryHidden::0x0,Query:worksheet:visible:A1:B2:2x2,' + $S.SummarySingle +
        ':worksheet:visible:A1:1x1,Chart1:chart:visible::0x0,Empty:worksheet:visible::0x0,Data:worksheet:visible:A1:E10:10x5'
    Assert 'sheets: order, kind, visibility, used range' ($sheets -eq $expected) $sheets
    $tables = @($summary.Json.Sheets | ForEach-Object { $sheet = $_; @($sheet.Tables) | ForEach-Object { $sheet.Name + '/' + $_.Name + '@' + $_.Range } }) -join ','
    Assert 'tables, with their sheet' ($tables -eq 'Data/tblData@A1:B3') $tables
    Assert 'sheets: code name and protection' (@($summary.Json.Sheets | Where-Object { -not $_.CodeName -or $_.Protected -ne $false }).Count -eq 0 -and $summary.Json.ProtectStructure -eq $false) $detail
    $names = @($summary.Json.Names | ForEach-Object { $_.Scope + '/' + $_.Name + $_.RefersTo + '/' + $_.Visible } | Sort-Object) -join ','
    Assert 'names: scope, name without the sheet, reference' ($names -eq 'Data/Local=Data!$A$2/True,Query/ExternalData_1=Query!$A$1:$B$2/True,workbook/Rate=Data!$B$2/True') $names
    Assert 'external links' (@($summary.Json.ExternalLinks).Count -eq 1 -and $summary.Json.ExternalLinks[0] -eq $other) $detail
    $connections = @($summary.Json.Connections | ForEach-Object { $_.Name + ':' + $_.Type }) -join ','
    Assert 'connections: name and type' ($connections -eq 'data:text') $connections
    # The connection string holds the path of the CSV file; it must not be in the result.
    Assert 'connections: no connection string in the result' ($summary.Text -notmatch 'data\.csv' -and @($summary.Json.Connections[0].PSObject.Properties).Count -eq 2) $detail
    $components = @($summary.Json.Vba.Components | ForEach-Object { $_.Name + ':' + $_.Kind + ':' + $_.Lines }) -join ','
    Assert 'VBA: components with code only' ($summary.Json.Vba.HasProject -eq $true -and $summary.Json.Vba.Locked -eq $false -and $null -eq $summary.Json.Vba.Error -and $components -eq 'Mod1:standard:3') $detail
    Assert 'customUI entries' ((@($summary.Json.CustomUI) -join ',') -eq 'customUI/customUI14.xml') $detail
    Assert 'workbook unchanged by the summary' ((Get-FileHashText $book) -eq $hash) 'hash changed'

    $plain = Invoke-Summary $other
    Assert 'empty .xlsx: exit 0 and nothing listed' ($plain.Exit -eq 0 -and @($plain.Json.Sheets).Count -eq 1 -and $null -eq $plain.Json.Sheets[0].UsedRange -and
        @($plain.Json.Names).Count -eq 0 -and @($plain.Json.ExternalLinks).Count -eq 0 -and @($plain.Json.Connections).Count -eq 0 -and
        $plain.Json.Vba.HasProject -eq $false -and @($plain.Json.Vba.Components).Count -eq 0 -and @($plain.Json.CustomUI).Count -eq 0) $plain.Text
    # An .xlsb is a zip file as well and can carry a ribbon.
    $binary = Join-Path $dir 'binary.xlsb'
    Use-ExcelWorkbook -Path $other -ScriptBlock { param($xl, $wb) $wb.SaveAs($binary, 50) }  # xlExcel12
    $ribbon = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $binary, '-XmlPath', (Join-Path $data 'ribbon\base.xml'))
    Assert '.xlsb: ribbon embedded' ($ribbon.Exit -eq 0) $ribbon.Output
    $xlsb = Invoke-Summary $binary
    Assert '.xlsb: customUI entries' ($xlsb.Exit -eq 0 -and (@($xlsb.Json.CustomUI) -join ',') -eq 'customUI/customUI14.xml') $xlsb.Text
    $missing = Invoke-Summary (Join-Path $dir 'no-such.xlsm')
    Assert 'missing workbook: exit 1 and Error' ($missing.Exit -eq 1 -and $missing.Json.Error) $missing.Text
}

Invoke-Case 'range' {
    $dir = New-CaseDir 'range'
    $book = Join-Path $dir 'range.xlsx'
    [void](New-ExcelWorkbook -Path $book)
    $r = Use-ExcelWorkbook -Path $book -ScriptBlock {
        param($xl, $wb)
        $sheet = $wb.Worksheets.Item(1)
        $out = [ordered]@{}
        function Get-Failure([scriptblock]$Block) {
            try { & $Block; 'no error' } catch { $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
        }
        # The column offsets of the cells with an apostrophe.
        function Get-Prefix([string]$Address) {
            $cells = @($sheet.Range($Address).Cells)
            @(for ($i = 0; $i -lt $cells.Count; $i++) { if ($cells[$i].PrefixCharacter -eq "'") { $i } }) -join ','
        }

        # Strings that Excel would parse stay text; the others keep their type.
        $json = '{"n":9,"s":"t"}' | ConvertFrom-Json
        # After this, an array set through PowerShell's property syntax fails with InvalidCastException
        # (docs/notes/range-value.md). Write-ExcelRange must still work.
        $sheet.Range('N1').Value2 = $json.n
        $date = [datetime]'2026-11-17 06:00'
        $row1 = @($null, '', [int]7, [long]8, [decimal]1.25, [double]2.5, $true, $date, [psobject]5, $json.n)
        $row2 = @('001', '1/2', '=A1', 'TRUE', "'abc", $S.Text, ' 12 ', '#N/A', 'plain', $json.s)
        $sheet.Range('A1:J2').Value2 = 'old'
        Write-ExcelRange -Range $sheet.Range('A1') -Value @($row1, $row2)
        $out.Mixed = Read-ExcelRange -Range $sheet.Range('A1:J2')
        $out.Prefix = Get-Prefix 'A2:J2'
        $out.DateFormat = [string]$sheet.Range('H1').NumberFormat
        $out.GeneralFormat = [string]$sheet.Range('Z1').NumberFormat

        # The values that make PowerShell's property syntax fail with InvalidCastException.
        foreach ($v in @('a', 1, $true, 2.5)) { Write-ExcelRange -Range $sheet.Range('L1') -Value (, @($v)) }
        Write-ExcelRange -Range $sheet.Range('L2') -Value (, @($json.n))
        Write-ExcelRange -Range $sheet.Range('L3:M3') -Value (, @('x', [double]3))
        $out.Repeated = Read-ExcelRange -Range $sheet.Range('L1:M3')

        # Shapes: one row, one column, one cell, a range of the exact size.
        Write-ExcelRange -Range $sheet.Range('A5') -Value (, @(1, 2, 3))
        Write-ExcelRange -Range $sheet.Range('A6') -Value @(@(4), @(5))
        Write-ExcelRange -Range $sheet.Range('C6:D7') -Value @(@(6, 7), @(8, 9))
        $out.OneRow = Read-ExcelRange -Range $sheet.Range('A5:C5')
        $out.OneColumn = Read-ExcelRange -Range $sheet.Range('A6:A7')
        $out.OneCell = Read-ExcelRange -Range $sheet.Range('A5')
        $out.EmptyCell = Read-ExcelRange -Range $sheet.Range('Z99')
        $out.Exact = Read-ExcelRange -Range $sheet.Range('C6:D7')

        # Formulas as typed input, read back as formulas; error values by name.
        Write-ExcelRange -Range $sheet.Range('A9') -Formula -Value (, @('=SEQUENCE(2)', '=A5*2', '001', '=1/0', '=NA()', '=NOSUCHFN()'))
        $out.Formulas = Read-ExcelRange -Range $sheet.Range('A9:F9') -Formula
        $out.FormulaValues = Read-ExcelRange -Range $sheet.Range('A9:F10')

        # Rejected input leaves the range as it was.
        $sheet.Range('A12:C13').Value2 = 'keep'
        $failures = [ordered]@{
            Ragged = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value @(@(1, 2), @(3)) }
            NotRows = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value @(1, 2, 3) }
            Size = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12:C13') -Value (, @(1, 2)) }
            Char = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value @(@(1, 2), @(3, [char]'q')) }
            Guid = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value (, @([guid]::Empty)) }
            NaN = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value (, @([double]::NaN)) }
            Nested = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value (, @(, @(1, 2))) }
            Long = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12') -Value (, @('x' * 32768)) }
            Edge = Get-Failure { Write-ExcelRange -Range $sheet.Range('A1048576') -Value @(@(1), @(2)) }
            Areas = Get-Failure { Write-ExcelRange -Range $sheet.Range('A12,C12') -Value (, @(1)) }
            ReadAreas = Get-Failure { [void](Read-ExcelRange -Range $sheet.Range('A12,C12')) }
        }
        $out.Failures = $failures
        $out.Kept = Read-ExcelRange -Range $sheet.Range('A12:C13')

        # Excel calculates a string that starts with "=" as soon as it gets it, so the string must already carry its
        # apostrophe in the first write. The arrays given to Excel are recorded by wrapping Set-ComProperty in the module.
        $module = Get-Module XlWright
        $original = & $module { ${function:Set-ComProperty} }
        $recorded = New-Object System.Collections.Generic.List[object]
        $recorder = { param($Object, [string]$Name, $Value) $recorded.Add(@($Value)); & $original $Object $Name $Value }.GetNewClosure()
        & $module { param($body) Set-Item -Path function:script:Set-ComProperty -Value $body } $recorder
        try { Write-ExcelRange -Range $sheet.Range('A15') -Value (, @('=1+1', ([string][char]0xFF1D + '1+1'), 'plain', '001')) }
        finally { & $module { param($body) Set-Item -Path function:script:Set-ComProperty -Value $body } $original }
        $out.Restored = (& $module { ${function:Set-ComProperty} }).ToString() -eq $original.ToString()
        $out.FirstWrite = @(if ($recorded.Count -gt 0) { $recorded[0] })
        $out.HasFormula = @($sheet.Range('A15:D15').Cells | ForEach-Object { [bool]$_.HasFormula })
        $out.Equals = Read-ExcelRange -Range $sheet.Range('A15:D15')
        $out
    }
    $mixed = $r.Mixed
    $row1 = @($mixed[0] | ForEach-Object { if ($null -eq $_) { 'null' } else { $_.GetType().Name + ':' + $_ } }) -join ','
    Assert 'values keep their type; null and empty string clear the cell' ($row1 -eq 'null,null,Double:7,Double:8,Double:1.25,Double:2.5,Boolean:True,Double:46343.25,Double:5,Double:9') $row1
    $row2 = @($mixed[1] | ForEach-Object { if ($null -eq $_) { 'null' } else { $_.GetType().Name + ':' + $_ } }) -join '|'
    $expected = (@('001', '1/2', '=A1', 'TRUE', "'abc", $S.Text, ' 12 ', '#N/A', 'plain', 't') | ForEach-Object { 'String:' + $_ }) -join '|'
    Assert 'strings stay text' ($row2 -eq $expected) $row2
    Assert 'apostrophe only where Excel changed the string' ($r.Prefix -eq '0,1,2,3,4,6,7') $r.Prefix
    Assert 'number format of a date left as it is' ($r.DateFormat -eq $r.GeneralFormat) ($r.DateFormat + ' <> ' + $r.GeneralFormat)
    $repeated = ($r.Repeated | ForEach-Object { @($_) -join ',' }) -join ';'
    Assert 'values of other types and from JSON in a row' ($repeated -eq '2.5,;9,;x,3') $repeated
    Assert 'one row' ((@($r.OneRow).Count -eq 1) -and ((@($r.OneRow[0]) -join ',') -eq '1,2,3')) ($r.OneRow | ConvertTo-Json -Compress)
    Assert 'one column' ((@($r.OneColumn).Count -eq 2) -and $r.OneColumn[0][0] -eq 4 -and $r.OneColumn[1][0] -eq 5) ($r.OneColumn | ConvertTo-Json -Compress)
    Assert 'one cell as one row of one value' ((@($r.OneCell).Count -eq 1) -and (@($r.OneCell[0]).Count -eq 1) -and $r.OneCell[0][0] -eq 1) ($r.OneCell | ConvertTo-Json -Compress)
    Assert 'empty cell as null' ((@($r.EmptyCell).Count -eq 1) -and $null -eq $r.EmptyCell[0][0]) ($r.EmptyCell | ConvertTo-Json -Compress)
    Assert 'range of the exact size' (($r.Exact | ForEach-Object { @($_) -join ',' }) -join ';' -eq '6,7;8,9') ($r.Exact | ConvertTo-Json -Compress)
    $formulas = @($r.Formulas[0]) -join '|'
    Assert 'formulas written as typed input and read back' ($formulas -eq '=SEQUENCE(2)|=A5*2|1|=1/0|=NA()|=NOSUCHFN()') $formulas
    $values = $r.FormulaValues
    $errors = @($values[0][3].Error, $values[0][4].Error, $values[0][5].Error) -join ','
    Assert 'formula spills (Formula2)' ($values[0][0] -eq 1 -and $values[1][0] -eq 2 -and $values[0][1] -eq 2) ($values | ConvertTo-Json -Compress)
    Assert 'error values by name' ($errors -eq '#DIV/0!,#N/A,#NAME?') $errors
    foreach ($name in $r.Failures.Keys) {
        Assert ('rejected: ' + $name) ($r.Failures[$name] -like 'ArgumentException:*') $r.Failures[$name]
    }
    $kept = ($r.Kept | ForEach-Object { @($_) -join ',' }) -join ';'
    Assert 'rejected input wrote nothing' ($kept -eq 'keep,keep,keep;keep,keep,keep') $kept
    $fullWidth = [string][char]0xFF1D + '1+1'
    $first = @($r.FirstWrite) -join '|'
    Assert 'Set-ComProperty restored after recording' ($r.Restored -eq $true) $r.Restored
    Assert 'a string that starts with "=" has its apostrophe in the first write' ($first -eq ("'=1+1|'" + $fullWidth + '|plain|001')) $first
    $equals = (@($r.Equals[0]) -join '|') + ' / ' + (@($r.HasFormula) -join ',')
    Assert 'strings with "=" kept as text, no formula' ($equals -eq ('=1+1|' + $fullWidth + '|plain|001 / False,False,False,False')) $equals

    # Through Invoke-ExcelScript.ps1: the rows and an error value as JSON.
    $run = Invoke-Runner $book 'range-read'
    $j = $run.Json.Result
    Assert 'Invoke-ExcelScript: exit 0' ($run.Exit -eq 0) $run.Output
    Assert 'Invoke-ExcelScript: rows and error value in JSON' ($j[0].Rows[0][0] -eq $S.Text -and $j[0].Rows[0][1].Error -eq '#DIV/0!' -and $null -eq $j[0].Rows[1][0]) ($j | ConvertTo-Json -Depth 6 -Compress)
}

function Invoke-FindError([string]$Book, [string[]]$Extra = @()) {
    $json = Join-Path $temp ('cell-error-' + [guid]::NewGuid().ToString('N') + '.json')
    $run = Invoke-Script (Join-Path $scripts 'Find-CellError.ps1') (@('-Path', $Book, '-OutFile', $json) + $Extra)
    $parsed = $null
    $text = ''
    if (Test-Path -LiteralPath $json) {
        $text = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8)
        $parsed = $text | ConvertFrom-Json
    }
    [pscustomobject]@{ Exit = $run.Exit; Json = $parsed; Text = $text; Output = $run.Output }
}

Invoke-Case 'cell-error' {
    $dir = New-CaseDir 'cell-error'
    $book = Join-Path $dir 'errors.xlsx'
    [void](New-ExcelWorkbook -Path $book)
    Use-ExcelWorkbook -Path $book -Save -ScriptBlock {
        param($xl, $wb)
        $data = $wb.Worksheets.Item(1)
        $data.Name = 'Data'
        # Row 1: formula errors, a formula that spills errors (E1:E2), a blocked spill (G1). Row 2: a constant error.
        Write-ExcelRange -Range $data.Range('A1') -Formula -Value @(
            @('=1/0', '=1+1', '=IFERROR(1/0,0)', $null, '=SEQUENCE(2)/0', $null, '=SEQUENCE(2)'),
            @('#N/A', '=NOSUCHFN()', '=#REF!+1', $null, $null, $null, 'block'))
        $clean = $wb.Worksheets.Add([Type]::Missing, $data)
        $clean.Name = 'Clean'
        Write-ExcelRange -Range $clean.Range('A1') -Formula -Value (, @('=1+1', 'x', '3'))
        $hidden = $wb.Worksheets.Add([Type]::Missing, $clean)
        $hidden.Name = 'Hidden'
        Write-ExcelRange -Range $hidden.Range('B2') -Formula -Value (, @('=NA()'))
        $hidden.Visible = 2  # xlSheetVeryHidden
        $protected = $wb.Worksheets.Add([Type]::Missing, $hidden)
        $protected.Name = 'Protected'
        Write-ExcelRange -Range $protected.Range('A1') -Formula -Value (, @('=1/0'))
        $protected.Protect()
        $offset = $wb.Worksheets.Add([Type]::Missing, $protected)
        $offset.Name = 'Offset'
        # The used range starts at C5; the addresses must still be those of the sheet.
        Write-ExcelRange -Range $offset.Range('C5') -Formula -Value @(@('=1/0', $null), @($null, 'x'))
        # Errors on both sides of the 10000-row blocks that the script reads.
        $long = $wb.Worksheets.Add([Type]::Missing, $offset)
        $long.Name = 'Long'
        foreach ($address in 'A1', 'A10000', 'A10001', 'A20001') {
            Write-ExcelRange -Range $long.Range($address) -Formula -Value (, @('=NA()'))
        }
        $chart = $wb.Charts.Add()
        $chart.Name = 'Chart1'
    }
    $hash = Get-FileHashText $book

    $all = Invoke-FindError $book
    $detail = $all.Text
    Assert 'exit 0, no error, path' ($all.Exit -eq 0 -and $null -eq $all.Json.Error -and $all.Json.Path -eq $book) $detail
    $counts = @($all.Json.Sheets | ForEach-Object { $_.Name + ':' + $_.Count }) -join ','
    Assert 'worksheets in order, chart sheet left out, hidden and protected read' ($counts -eq 'Data:7,Clean:0,Hidden:1,Protected:1,Offset:1,Long:4') $counts
    Assert 'total, not truncated, calculation' ($all.Json.Count -eq 14 -and $all.Json.Truncated -eq $false -and $all.Json.Calculation -eq 'automatic') $detail
    $errors = @($all.Json.Sheets[0].Errors | ForEach-Object { $_.Address + ':' + $_.Error + ':' + $_.Kind + ':' + $_.Formula + ':' + $_.SpillParent }) -join '|'
    $expected = 'A1:#DIV/0!:formula:=1/0:|E1:#DIV/0!:formula:=SEQUENCE(2)/0:|G1:#SPILL!:formula:=SEQUENCE(2):|' +
        'A2:#N/A:constant::|B2:#NAME?:formula:=NOSUCHFN():|C2:#REF!:formula:=#REF!+1:|E2:#DIV/0!:spill::E1'
    Assert 'errors: address, name, kind, formula, spill parent' ($errors -eq $expected) $errors
    $others = @($all.Json.Sheets | Select-Object -Skip 2 | ForEach-Object { $_.Name + '!' + (@($_.Errors | ForEach-Object { $_.Address + ':' + $_.Error }) -join ';') }) -join ','
    Assert 'errors on hidden, protected, offset and long sheets' ($others -eq 'Hidden!B2:#N/A,Protected!A1:#DIV/0!,Offset!C5:#DIV/0!,Long!A1:#N/A;A10000:#N/A;A10001:#N/A;A20001:#N/A') $others

    $few = Invoke-FindError $book @('-MaxErrors', '2')
    $listed = @($few.Json.Sheets | ForEach-Object { @($_.Errors).Count } | Measure-Object -Sum).Sum
    Assert '-MaxErrors: listed, total and Truncated' ($few.Exit -eq 0 -and $listed -eq 2 -and $few.Json.Count -eq 14 -and $few.Json.Truncated -eq $true) $few.Text
    $some = Invoke-FindError $book @('-Sheet', 'Offset,Clean')
    $names = @($some.Json.Sheets | ForEach-Object { $_.Name + ':' + $_.Count }) -join ','
    Assert '-Sheet: only those, in the order of the workbook' ($some.Exit -eq 0 -and $names -eq 'Clean:0,Offset:1' -and $some.Json.Count -eq 1) $some.Text
    $wrong = Invoke-FindError $book @('-Sheet', 'Chart1')
    Assert '-Sheet with a chart sheet: exit 1 and Error' ($wrong.Exit -eq 1 -and $wrong.Json.Error -like '*Chart1*') $wrong.Text
    Assert 'workbook unchanged' ((Get-FileHashText $book) -eq $hash) 'hash changed'
    $missing = Invoke-FindError (Join-Path $dir 'no-such.xlsx')
    Assert 'missing workbook: exit 1 and Error' ($missing.Exit -eq 1 -and $missing.Json.Error) $missing.Text
}

Invoke-Case 'userform' {
    $dir = New-CaseDir 'userform'
    $book = Join-Path $dir 'form.xlsm'
    $c = Invoke-Runner $book 'uf-create' @('-Create', '-Save')
    Assert 'create exit 0' ($c.Exit -eq 0) $c.Output
    Assert 'font assignment from PowerShell fails as documented' ($c.Json.Result[0].FontAssignment -eq 'NullReferenceException') $c.Json.Result[0].FontAssignment
    $export = Invoke-Script (Join-Path $scripts 'Export-Vba.ps1') @('-Path', $book, '-OutDir', (Join-Path $dir 'export'))
    Assert 'export exit 0' ($export.Exit -eq 0) $export.Output
    Assert '.frm and .frx exported' ((Test-Path (Join-Path $dir 'export\forms\frmInfo.frm')) -and (Test-Path (Join-Path $dir 'export\forms\frmInfo.frx'))) 'missing'
    $copy = Join-Path $dir 'copy.xlsm'
    $n = Invoke-Runner $copy 'noop' @('-Create', '-Save')
    Assert 'second workbook created' ($n.Exit -eq 0) $n.Output
    $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $copy, '-SourceDir', (Join-Path $dir 'export'))
    Assert 'import exit 0' ($import.Exit -eq 0) $import.Output
    $d1 = Invoke-Runner $book 'uf-design'
    $d2 = Invoke-Runner $copy 'uf-design'
    $j1 = $d1.Json.Result | ConvertTo-Json -Depth 6 -Compress
    $j2 = $d2.Json.Result | ConvertTo-Json -Depth 6 -Compress
    Assert 'design read back' ($d1.Exit -eq 0 -and $d2.Exit -eq 0) ($d1.Output + $d2.Output)
    Assert 'design identical after import' ($j1 -eq $j2) ($j1 + ' <> ' + $j2)
    $design = $d2.Json.Result[0]
    Assert 'form caption (Japanese)' ($design.Caption -eq $S.FormCaption) $design.Caption
    Assert 'frame holds the check box' ((@($design.FrameChildren) -join ',') -eq 'chkShow') (@($design.FrameChildren) -join ',')
    $f = Invoke-Runner $copy 'uf-font' @('-EnableMacros', '-Save')
    Assert 'font via temporary VBA' ($f.Exit -eq 0) $f.Output
    Assert 'temporary module removed' ((@($f.Json.Result) -notcontains 'TmpSetFont')) (@($f.Json.Result) -join ',')
    $rt = Invoke-Runner $copy 'uf-runtime' @('-EnableMacros')
    $expected = $S.FormCaption + '|' + $S.FormSheetValue + '|True|Meiryo UI'
    Assert 'runtime initialize and font' ($rt.Json.Result[0] -eq $expected) $rt.Json.Result[0]
}

if ($Full) {
    # A compile error shows the VBE window and the error dialog on the screen.
    Invoke-Case 'vba-compile-error' -WatchForeground {
        $dir = New-CaseDir 'vba-compile-error'
        $cases = @(
            @{ Name = 'undeclared'; Column = 5 },
            @{ Name = 'syntax'; Column = 1 })
        foreach ($case in $cases) {
            $book = Join-Path $dir ($case.Name + '.xlsm')
            [void](New-ExcelWorkbook -Path $book)
            $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', (Join-Path $data ('vba-broken\' + $case.Name)))
            Assert ($case.Name + ': import exit 0') ($import.Exit -eq 0) $import.Output
            $hash = Get-FileHashText $book
            $c = Invoke-Compile $book
            $detail = $c.Json | ConvertTo-Json -Compress
            Assert ($case.Name + ': exit 1') ($c.Exit -eq 1) $c.Output
            Assert ($case.Name + ': reported as a compile error, not a failed check') ($c.Json.Compiled -eq $false -and $null -eq $c.Json.Error) $detail
            Assert ($case.Name + ': message read') ([string]$c.Json.CompileError.Message -ne '') $detail
            Assert ($case.Name + ': place read') ($c.Json.CompileError.Module -eq 'Broken' -and $c.Json.CompileError.Line -eq 5 -and $c.Json.CompileError.Column -eq $case.Column) $detail
            Assert ($case.Name + ': workbook unchanged by the check') ((Get-FileHashText $book) -eq $hash) 'hash changed'
        }
    }

    # The dialogs of a macro are on the screen even when Excel is hidden.
    Invoke-Case 'macro-dialog' -WatchForeground {
        $dir = New-CaseDir 'macro-dialog'
        $book = New-BookWithMacros $dir 'vba-macro'

        $r = Use-ExcelWorkbook -Path $book -EnableMacros -ScriptBlock {
            param($xl, $wb)
            $errors = foreach ($name in 'Macros.DivZero', 'Macros.DivZeroSub') {
                try { Invoke-ExcelMacro -Excel $xl -Name $name; 'no error' } catch { $_.Exception.Message }
            }
            # Excel must still be usable after the runtime errors.
            [ordered]@{ Errors = @($errors); After = Invoke-ExcelMacro -Excel $xl -Name 'Macros.AddThree' -ArgumentList 1, 2, 3 }
        }
        Assert 'runtime error in a Function: thrown with the text' ($r.Errors[0] -like "Macro 'Macros.DivZero' stopped with a runtime error: *11*") $r.Errors[0]
        Assert 'runtime error in a Sub: thrown with the text' ($r.Errors[1] -like "Macro 'Macros.DivZeroSub' stopped with a runtime error: *11*") $r.Errors[1]
        Assert 'Excel usable after the runtime errors' ($r.After -eq 6) $r.After

        $m = Invoke-Runner $book 'macro-msgbox' @('-EnableMacros')
        $detail = $m.Json | ConvertTo-Json -Compress
        Assert 'MsgBox: closed and the macro went on' ($m.Exit -eq 0 -and $m.Json.Result[0] -eq 5) $detail
        $closed = @($m.Json.Warnings)
        Assert 'MsgBox: both dialogs reported as warnings' ($closed.Count -eq 2 -and $closed[0] -like '*probe: first message' -and $closed[1] -like '*probe: second message') $detail

        # A compile error leaves the VBE in break mode; it is reset so that Excel stays usable.
        foreach ($kind in 'undeclared', 'syntax') {
            $broken = New-BookWithMacros $dir ('vba-broken\' + $kind)
            $warnings = @()
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $c = Use-ExcelWorkbook -Path $broken -EnableMacros -WarningVariable +warnings -ScriptBlock {
                param($xl, $wb)
                $errors = foreach ($i in 1, 2) {
                    try { Invoke-ExcelMacro -Excel $xl -Name 'Broken.Run'; 'no error' } catch { $_.Exception.Message }
                }
                [ordered]@{
                    Errors = @($errors)
                    After = Invoke-ExcelMacro -Excel $xl -Name 'Good.Seven'
                    EditorVisible = [XlWright.Dialog]::IsEditorVisible([uint32](Get-ExcelProcessId $xl))
                }
            }
            Assert ($kind + ': thrown as a compile error with the text') ($c.Errors[0] -like "Macro 'Broken.Run' stopped with a compile error: ?*") $c.Errors[0]
            Assert ($kind + ': the same when run again') ($c.Errors[1] -eq $c.Errors[0]) $c.Errors[1]
            Assert ($kind + ': Excel usable after the compile errors') ($c.After -eq 7) $c.After
            Assert ($kind + ': the VBE is hidden again') ($c.EditorVisible -eq $false) $c.EditorVisible
            Assert ($kind + ': Excel was closed, not stopped') (@($warnings).Count -eq 0) (@($warnings) -join ' | ')
            Assert ($kind + ': done within 30 s') ($watch.Elapsed.TotalSeconds -lt 30) ([int]$watch.Elapsed.TotalSeconds)
        }

        $stops = @(
            @{ Name = 'Yes/No MsgBox'; Book = $book; Macro = 'Macros.Ask'; Error = '*cannot be closed without choosing an answer: probe: question*' })
        foreach ($stop in $stops) {
            $message = $null
            $watch = [Diagnostics.Stopwatch]::StartNew()
            try {
                Use-ExcelWorkbook -Path $stop.Book -EnableMacros -WarningAction SilentlyContinue -ScriptBlock {
                    param($xl, $wb)
                    Invoke-ExcelMacro -Excel $xl -Name $stop.Macro
                } | Out-Null
            }
            catch { $message = $_.Exception.Message }
            Assert ($stop.Name + ': thrown with the text') ([string]$message -like $stop.Error) $message
            Assert ($stop.Name + ': says its Excel was stopped') ([string]$message -like '*Stopped the Excel instance of this call*') $message
            # Without the watcher the call would only end at the limit of Use-ExcelWorkbook (300 s).
            Assert ($stop.Name + ': stopped within 30 s') ($watch.Elapsed.TotalSeconds -lt 30) ([int]$watch.Elapsed.TotalSeconds)
            $left = Wait-NoExcel 5
            Assert ($stop.Name + ': its Excel was stopped') ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
        }
    }

    # A failing test shows the runtime error dialog on the screen.
    Invoke-Case 'vba-test-fail' -WatchForeground {
        $book = New-BookWithMacros (New-CaseDir 'vba-test-fail') 'vba-test\fail'
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $t = Invoke-VbaTest $book
        $detail = $t.Json | ConvertTo-Json -Compress -Depth 5
        $tests = @($t.Json.Tests)
        Assert 'a failed test: exit 1 and not Ok' ($t.Exit -eq 1 -and $t.Json.Ok -eq $false) $detail
        $results = @($tests | ForEach-Object { $_.Result }) -join ','
        # Test_Broken comes first: a compile error fails only that test and the run goes on.
        Assert 'result of every test' ($results -eq 'failed,passed,failed,passed,passed,failed,not_run') $results
        Assert 'compile error: the test failed with the text of the dialog' ($tests[0].Name -eq 'Test_Broken.Test_Undeclared' -and [string]$tests[0].Message -ne '' -and [string]$tests[0].Message -notlike '*Microsoft Visual Basic*') $tests[0].Message
        Assert 'runtime error: message of the dialog, without the title' ([string]$tests[2].Message -like '*80040001*: expected 3, actual 2' -and [string]$tests[2].Message -notlike '*Microsoft Visual Basic*') $tests[2].Message
        Assert 'passed tests have no message' ($null -eq $tests[1].Message -and $null -eq $tests[3].Message) $detail
        Assert 'MsgBox with OK: closed and reported as a warning' (@($t.Json.Warnings | Where-Object { $_ -like "*'Test_Fail.Test_4_Note': probe: note" }).Count -eq 1) $detail
        Assert 'Yes/No MsgBox: the test failed with the text' ($tests[5].Message -eq 'question') $tests[5].Message
        Assert 'Yes/No MsgBox: Error says that Excel was stopped' ([string]$t.Json.Error -like '*cannot be closed without choosing an answer: probe: question*Stopped the Excel instance of this call*') $t.Json.Error
        Assert 'summary counts' ($t.Json.Summary.Total -eq 7 -and $t.Json.Summary.Passed -eq 3 -and $t.Json.Summary.Failed -eq 3 -and $t.Json.Summary.NotRun -eq 1) $detail
        Assert 'stopped within 30 s' ($watch.Elapsed.TotalSeconds -lt 30) ([int]$watch.Elapsed.TotalSeconds)
    }

    Invoke-Case 'ribbon' -WatchForeground {
        $dir = New-CaseDir 'ribbon'
        $book = New-BookWithVba $dir
        # The sheet tab is a TabItem too: a sheet named like the ribbon tab must not make the tab ambiguous.
        Use-ExcelWorkbook -Path $book -Save -ScriptBlock { param($xl, $wb) $wb.Worksheets.Item(1).Name = $S.Tab }
        $cases = @(
            @{ Xml = 'base.xml'; Exit = 0; Error = $null },
            @{ Xml = 'same-group.xml'; Exit = 0; Error = $null },
            @{ Xml = 'same-tab.xml'; Exit = 1; Error = 'Ambiguous tab*' },
            @{ Xml = 'same-button.xml'; Exit = 1; Error = 'Ambiguous button*' })
        foreach ($case in $cases) {
            $xmlPath = Join-Path $data ('ribbon\' + $case.Xml)
            $add = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', $xmlPath)
            Assert ($case.Xml + ': embed exit 0') ($add.Exit -eq 0) $add.Output
            $zip = [IO.Compression.ZipFile]::OpenRead($book)
            try {
                $reader = New-Object IO.StreamReader($zip.GetEntry('customUI/customUI14.xml').Open())
                try { [xml]$embedded = $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
            finally { $zip.Dispose() }
            [xml]$source = [IO.File]::ReadAllText($xmlPath, [Text.Encoding]::UTF8)
            Assert ($case.Xml + ': embedded XML matches') ($embedded.DocumentElement.OuterXml -ceq $source.DocumentElement.OuterXml) 'differs'
            $json = Join-Path $dir ($case.Xml + '.json')
            $click = Invoke-Script (Join-Path $scripts 'Invoke-RibbonButton.ps1') @('-Path', $book, '-TabLabel', $S.Tab, '-ButtonLabel', $S.Button, '-OutFile', $json)
            $res = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json
            Assert ($case.Xml + ': exit ' + $case.Exit) ($click.Exit -eq $case.Exit) ($res | ConvertTo-Json -Compress -Depth 5)
            if ($case.Error) {
                Assert ($case.Xml + ': stops before clicking') ((-not $res.Clicked) -and ([string]$res.Error -like $case.Error)) ($res | ConvertTo-Json -Compress -Depth 5)
            }
            else {
                Assert ($case.Xml + ': clicked and read the MsgBox') ($res.Clicked -and (@($res.Dialogs[0].Names) -contains $S.Message)) ($res | ConvertTo-Json -Compress -Depth 5)
            }
            if ($case.Xml -eq 'base.xml') {
                # A Yes/No MsgBox ignores the close button: only this script's Excel is stopped.
                $askJson = Join-Path $dir 'ask.json'
                $ask = Invoke-Script (Join-Path $scripts 'Invoke-RibbonButton.ps1') @('-Path', $book, '-TabLabel', $S.Tab, '-ButtonLabel', $S.AskButton, '-OutFile', $askJson)
                $askRes = [IO.File]::ReadAllText($askJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
                Assert 'Yes/No dialog: exit 1' ($ask.Exit -eq 1) ($askRes | ConvertTo-Json -Compress -Depth 5)
                Assert 'Yes/No dialog: reported as not closable' ([string]$askRes.Error -like 'A dialog could not be closed*') $askRes.Error
                $left = Wait-NoExcel 20
                Assert 'Yes/No dialog: its Excel was stopped' ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
            }
        }
    }

    Invoke-Case 'ribbon-layout' -WatchForeground {
        $dir = New-CaseDir 'ribbon-layout'
        $book = New-BookWithVba $dir
        Use-ExcelWorkbook -Path $book -Save -ScriptBlock { param($xl, $wb) $wb.Worksheets.Item(1).Name = $S.Tab }
        $layout = Join-Path $scripts 'Test-RibbonLayout.ps1'

        $noUiJson = Join-Path $dir 'no-ui.json'
        $noUi = Invoke-Script $layout @('-Path', $book, '-OutFile', $noUiJson)
        $noUiRes = [IO.File]::ReadAllText($noUiJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
        Assert 'no customUI: exit 2 with the reason' ($noUi.Exit -eq 2 -and [string]$noUiRes.Error -like '*no customUI*') ($noUiRes | ConvertTo-Json -Compress)

        $add = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', (Join-Path $data 'ribbon\layout.xml'))
        Assert 'layout.xml: embed exit 0' ($add.Exit -eq 0) $add.Output
        $shots = Join-Path $dir 'shots'
        New-Item -ItemType Directory -Path $shots | Out-Null
        $json = Join-Path $dir 'layout.json'
        $run = Invoke-Script $layout @('-Path', $book, '-OutFile', $json, '-ScreenshotDir', $shots)
        $res = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $detail = $res | ConvertTo-Json -Compress -Depth 8
        Assert 'layout.xml: exit 0, no mismatch' ($run.Exit -eq 0 -and @($res.Mismatches).Count -eq 0 -and -not $res.Error) $detail
        $tab = @($res.Actual)[0]
        Assert 'tab placed after the Home tab' ($tab.Before -eq 'TabHome') $tab.Before
        $groups = @($tab.Groups)
        Assert 'built-in group, then the custom groups' ($groups.Count -eq 3 -and $groups[1].Name -eq $S.Group -and $groups[2].Name -eq $S.MenuGroup) (@($groups | ForEach-Object { $_.Name }) -join ',')
        $controls = @($groups[1].Controls)
        $names = @($controls | ForEach-Object { $_.Name }) -join ','
        Assert 'buttons in order, hidden button absent' ($names -eq (@($S.Button, $S.SaveButton, $S.IconButton, $S.OffButton, $S.DynButton) -join ',')) $names
        if ($controls.Count -eq 5) {
            Assert 'large and normal size read' ($controls[0].Size -eq 'large' -and $controls[1].Size -eq 'normal') (($controls | ForEach-Object { $_.Size }) -join ',')
            Assert 'showLabel=false read as icon only' ($controls[2].IconOnly -and -not $controls[1].IconOnly) (($controls | ForEach-Object { $_.IconOnly }) -join ',')
            Assert 'disabled button read' (-not $controls[3].Enabled -and $controls[1].Enabled) (($controls | ForEach-Object { $_.Enabled }) -join ',')
        }
        $dynPath = $S.Tab + '/' + $S.Group + '/' + $S.DynButton
        Assert 'getEnabled listed in Dynamic' (@($res.Dynamic | Where-Object { $_.Path -eq $dynPath -and $_.Item -eq 'Enabled' -and $_.Actual -eq $false }).Count -eq 1) $detail
        Assert 'group with a menu listed in Unsupported' (@($res.Unsupported | Where-Object { $_.Path -eq ($S.Tab + '/' + $S.MenuGroup) }).Count -eq 1) $detail
        Assert 'hidden getLabel button listed as absent, not paired with the next button' (@($res.Dynamic | Where-Object { $_.Path -eq ($S.Tab + '/' + $S.Group + '/ghostButton') -and $_.Item -eq 'Visible' -and $_.Actual -eq $false }).Count -eq 1) $detail
        $extra = @($res.Actual | Where-Object { $_.Tab -eq 'Extra' })
        Assert 'second tab after the Home tab: nearest built-in tab is Home' ($extra.Count -eq 1 -and $extra[0].Before -eq 'TabHome') $detail
        Assert 'hidden built-in group listed in Unsupported' (@($res.Unsupported | Where-Object { $_.Path -eq 'TabHome/GroupFont' }).Count -eq 1) $detail
        Assert 'group position by insertAfterMso listed in Unsupported' (@($res.Unsupported | Where-Object { $_.Path -eq 'TabHome/HomeExtra' }).Count -eq 1) $detail
        $extraGroups = @($extra | ForEach-Object { $_.Groups })
        Assert 'group and button without a label matched' ($extraGroups.Count -eq 2 -and @($extraGroups[1].Controls).Count -eq 1) $detail
        $png = Join-Path $shots 'tab-1.png'
        $pngOk = (Test-Path -LiteralPath $png) -and ((Get-Item -LiteralPath $png).Length -gt 1000)
        if ($pngOk) {
            $head = [IO.File]::ReadAllBytes($png)[0..3]
            $pngOk = ($head[0] -eq 0x89 -and $head[1] -eq 0x50 -and $head[2] -eq 0x4E -and $head[3] -eq 0x47)
        }
        Assert 'screenshot written as PNG' $pngOk $png

        # Excel ignores the whole ribbon when the XML has an unknown attribute: the tab is reported missing.
        $add = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', (Join-Path $data 'ribbon\layout-rejected.xml'))
        Assert 'layout-rejected.xml: embed exit 0' ($add.Exit -eq 0) $add.Output
        $badJson = Join-Path $dir 'rejected.json'
        $bad = Invoke-Script $layout @('-Path', $book, '-OutFile', $badJson, '-FindTimeoutSeconds', '5')
        $badRes = [IO.File]::ReadAllText($badJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $add = Invoke-Script (Join-Path $scripts 'Add-CustomUI.ps1') @('-Path', $book, '-XmlPath', (Join-Path $data 'ribbon\collapse.xml'))
        Assert 'collapse.xml: embed exit 0' ($add.Exit -eq 0) $add.Output
        # Maximized: nothing collapses, and a group holding only a menu with the group label is not taken for collapsed.
        $wideJson = Join-Path $dir 'wide.json'
        $wide = Invoke-Script $layout @('-Path', $book, '-OutFile', $wideJson)
        $wideRes = [IO.File]::ReadAllText($wideJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $wideDetail = $wideRes | ConvertTo-Json -Compress -Depth 8
        Assert 'maximized: nothing collapsed' ($wide.Exit -eq 0 -and @($wideRes.Collapsed).Count -eq 0) $wideDetail
        Assert 'group with a same-named menu is Unsupported, not collapsed' (@($wideRes.Unsupported | Where-Object { $_.Path -eq 'Wide/Menus' }).Count -eq 1) $wideDetail
        # A narrow window: groups collapse, the layout check lists them, and the click names them in the error.
        $narrowJson = Join-Path $dir 'narrow.json'
        $narrow = Invoke-Script $layout @('-Path', $book, '-OutFile', $narrowJson, '-WindowWidth', '420')
        $narrowRes = [IO.File]::ReadAllText($narrowJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
        Assert 'narrow: collapsed groups listed' ($narrow.Exit -eq 0 -and @($narrowRes.Collapsed) -contains 'Wide/G6') ($narrowRes | ConvertTo-Json -Compress -Depth 8)
        $clickJson = Join-Path $dir 'narrow-click.json'
        $click = Invoke-Script (Join-Path $scripts 'Invoke-RibbonButton.ps1') @('-Path', $book, '-TabLabel', 'Wide', '-ButtonLabel', 'G6-4', '-OutFile', $clickJson, '-FindTimeoutSeconds', '5', '-WindowWidth', '420')
        $clickRes = [IO.File]::ReadAllText($clickJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
        Assert 'narrow: click names the collapsed groups' ($click.Exit -eq 1 -and -not $clickRes.Clicked -and [string]$clickRes.Error -like '*collapsed for lack of width:*G6*') ($clickRes | ConvertTo-Json -Compress)
        Assert 'rejected XML: exit 1, tab missing' ($bad.Exit -eq 1 -and @($badRes.Mismatches | Where-Object { $_.Path -eq $S.Tab -and $_.Actual -eq 'missing' }).Count -eq 1) ($badRes | ConvertTo-Json -Compress -Depth 8)
    }

    Invoke-Case 'publish' -WatchForeground {
        $dir = New-CaseDir 'publish'
        $book = Join-Path $dir 'book.xlsm'
        [void](New-ExcelWorkbook -Path $book)
        Copy-Item -LiteralPath (Join-Path $data 'vba') -Destination (Join-Path $dir 'vba') -Recurse
        New-Item -ItemType Directory -Path (Join-Path $dir 'ribbon') | Out-Null
        Copy-Item -LiteralPath (Join-Path $data 'ribbon\layout.xml') -Destination (Join-Path $dir 'ribbon\customUI14.xml')
        $pub =Invoke-Script (Join-Path $scripts 'Publish-Workbook.ps1') @('-Path', $book, '-OutDir', (Join-Path $dir 'dist'))
        Assert 'publish exit 0' ($pub.Exit -eq 0) $pub.Output
        $dist = Join-Path $dir 'dist\book.xlsm'
        foreach ($p in @(@{ Path = $dist; Protection = 1 }, @{ Path = $book; Protection = 0 })) {
            $r = Use-ExcelWorkbook -Path $p.Path -EnableMacros -ScriptBlock {
                param($xl, $wb)
                @{ Protection = $wb.VBProject.Protection; Add = $xl.Run('AddNums', 4, 5) }
            }
            $zip = [IO.Compression.ZipFile]::OpenRead($p.Path)
            try { $ui = $null -ne $zip.GetEntry('customUI/customUI14.xml') } finally { $zip.Dispose() }
            $label = Split-Path -Leaf (Split-Path -Parent $p.Path)
            Assert ($label + ': VBA protection ' + $p.Protection) ($r.Protection -eq $p.Protection) $r.Protection
            Assert ($label + ': macro runs') ($r.Add -eq 9) $r.Add
            Assert ($label + ': ribbon embedded') $ui 'no customUI'
        }
        # The summary of a workbook with a locked VBA project: everything but the components.
        $locked = Invoke-Summary $dist
        Assert 'dist: summary without the VBA components' ($locked.Exit -eq 0 -and $null -eq $locked.Json.Error -and @($locked.Json.Sheets).Count -eq 1 -and
            $locked.Json.Vba.Locked -eq $true -and $null -eq $locked.Json.Vba.Components -and [string]$locked.Json.Vba.Error -like '*locked*' -and
            (@($locked.Json.CustomUI) -join ',') -eq 'customUI/customUI14.xml') $locked.Text
        # The ribbon of the locked copy is built and its callbacks (getEnabled) still run.
        $json = Join-Path $dir 'dist-layout.json'
        $layout = Invoke-Script (Join-Path $scripts 'Test-RibbonLayout.ps1') @('-Path', $dist, '-OutFile', $json)
        $res = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $dynPath = $S.Tab + '/' + $S.Group + '/' + $S.DynButton
        Assert 'dist: ribbon layout matches the XML' ($layout.Exit -eq 0) ($res | ConvertTo-Json -Compress -Depth 8)
        Assert 'dist: getEnabled callback ran' (@($res.Dynamic | Where-Object { $_.Path -eq $dynPath -and $_.Item -eq 'Enabled' -and $_.Actual -eq $false }).Count -eq 1) ($res.Dynamic | ConvertTo-Json -Compress)
    }

    Invoke-Case 'publish-exclude-tests' -WatchForeground {
        $dir = New-CaseDir 'publish-exclude-tests'
        $book = Join-Path $dir 'book.xlsm'
        [void](New-ExcelWorkbook -Path $book)
        Copy-Item -LiteralPath (Join-Path $data 'vba') -Destination (Join-Path $dir 'vba') -Recurse
        # Test_Calc is a test module. Alpha holds tests but is not named Test_*.
        Copy-Item -Path (Join-Path $data 'vba-test\pass\standard\*') -Destination (Join-Path $dir 'vba\standard')
        # The published project is locked, so the modules are probed by running a Sub, not by listing them.
        $probe = {
            param($xl, $wb)
            $run = { param($name) try { [void]$xl.Run($name); 'ran' } catch { 'missing' } }
            @{
                Protection = $wb.VBProject.Protection; Add = $xl.Run('AddNums', 4, 5)
                TestModule = (& $run 'Test_Calc.Test_Add'); Mixed = (& $run 'Alpha.Test_First')
            }
        }
        $publisher = Join-Path $scripts 'Publish-Workbook.ps1'

        $dist = Join-Path $dir 'dist\book.xlsm'
        $pub = Invoke-Script $publisher @('-Path', $book, '-OutDir', (Join-Path $dir 'dist'), '-ExcludeTests')
        $lines = @($pub.Output -split "\r?\n" | Where-Object { $_ })
        Assert '-ExcludeTests: exit 0' ($pub.Exit -eq 0) $pub.Output
        Assert '-ExcludeTests: the output names the removed module and no other' ((@($lines | Where-Object { $_ -like 'removed test module:*' }) -join ',') -eq 'removed test module: Test_Calc') $pub.Output
        Assert '-ExcludeTests: the last line of the output is the published workbook' ($lines[-1] -eq $dist) $pub.Output
        # The host wraps a long warning.
        Assert '-ExcludeTests: warns about the tests left in another module' (($pub.Output -replace '\s+', ' ') -match 'Tests left in the published workbook.*: Alpha\.Test_Second, Alpha\.Test_First') $pub.Output
        $d = Use-ExcelWorkbook -Path $dist -EnableMacros -ScriptBlock $probe
        Assert '-ExcludeTests, dist: locked, macro runs' ($d.Protection -eq 1 -and $d.Add -eq 9) ($d | ConvertTo-Json -Compress)
        Assert '-ExcludeTests, dist: test module removed, other modules kept' ($d.TestModule -eq 'missing' -and $d.Mixed -eq 'ran') ($d | ConvertTo-Json -Compress)
        $w = Use-ExcelWorkbook -Path $book -EnableMacros -ScriptBlock $probe
        Assert '-ExcludeTests, work book: not locked, test module kept' ($w.Protection -eq 0 -and $w.Add -eq 9 -and $w.TestModule -eq 'ran' -and $w.Mixed -eq 'ran') ($w | ConvertTo-Json -Compress)

        $plain = Invoke-Script $publisher @('-Path', $book, '-OutDir', (Join-Path $dir 'dist-plain'))
        Assert 'without -ExcludeTests: exit 0, only the path in the output' ($plain.Exit -eq 0 -and $plain.Output.Trim() -eq (Join-Path $dir 'dist-plain\book.xlsm')) $plain.Output
        $p = Use-ExcelWorkbook -Path (Join-Path $dir 'dist-plain\book.xlsm') -EnableMacros -ScriptBlock $probe
        Assert 'without -ExcludeTests, dist: locked, test module kept' ($p.Protection -eq 1 -and $p.TestModule -eq 'ran' -and $p.Mixed -eq 'ran') ($p | ConvertTo-Json -Compress)
    }
}

# ---------------------------------------------------------------------------------------------
$failedCases = @($results | Where-Object { -not $_.Ok })
$summary = [ordered]@{
    Ok = ($failedCases.Count -eq 0 -and -not $stopped); Full = [bool]$Full; Stopped = $stopped
    TempDir = $temp; Cases = $results.ToArray()
}
if ($OutFile) { [IO.File]::WriteAllText($OutFile, ([pscustomobject]$summary | ConvertTo-Json -Depth 8), $utf8) }
# The failed checks and their details are in the output and -OutFile; the workbooks are kept only on request.
if ($KeepTemp) { Write-Output ('Temporary files kept: ' + $temp) }
else {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temp) { Write-Output ('Could not remove the temporary files (in use?): ' + $temp) }
    elseif (-not $summary.Ok) { Write-Output 'Temporary files removed. Run again with -KeepTemp to inspect them.' }
}
Write-Output ('{0} of {1} cases passed{2}' -f ($results.Count - $failedCases.Count), $results.Count,
    $(if ($stopped) { '; stopped because Excel was left running' } else { '' }))
if ($stopped) { exit 2 }
if ($failedCases.Count -gt 0) { exit 1 }
exit 0
