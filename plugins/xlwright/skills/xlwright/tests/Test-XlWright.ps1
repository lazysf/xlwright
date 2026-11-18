<#
.SYNOPSIS
Checks the xlwright scripts against the real Excel on this machine.

.DESCRIPTION
Every case creates its own workbooks in a new temporary folder and checks explicit expected values.
Nothing outside that folder is changed. The default cases keep Excel hidden. -Full adds the cases
that show Excel and the VBE on the screen (VBA compile error, macro dialogs, ribbon click, ribbon layout, publish with VBA lock, screenshot); do not touch
those windows while they run. These cases also check that Excel never takes the foreground.

Preconditions: no Excel process is running, and AccessVBOM is 1. Otherwise nothing runs (exit 2).
After each case no Excel process may remain. If one does, the run stops without touching it.

-Case runs only the named cases, e.g. -Case timeout,userform.

-Changed runs only the cases that cover the files of the skill changed since HEAD (staged, unstaged and untracked),
by $CasesByFile below. It cannot be combined with -Case. The cases that need -Full are named and skipped without it.

The result is written as UTF-8 JSON to -OutFile and summarized on the console.
Exit code: 0 all passed, 1 a check failed, 2 preconditions not met or the run was stopped.
The temporary folder is removed at the end, whether the cases passed or not, unless -KeepTemp is given.
#>
param(
    [switch]$Full,
    # Run only these cases (names as printed). Cases that need -Full still need it.
    [string[]]$Case,
    # Run only the cases that cover the files changed since HEAD.
    [switch]$Changed,
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

# The cases that cover each file, by its path in the skill folder: the cases that call the script, directly or
# through another script (Publish-Workbook.ps1 calls Import-Vba.ps1, Add-CustomUI.ps1 and Lock-VbaProject.ps1).
# '*' is every case. A file of scripts\
# that is not listed, Test-XlWright.ps1 and anything in tests\data\ also run every case; other files run none.
# When a case or a script is added, add it here as well.
$CasesByFile = [ordered]@{
    'scripts/XlWright.psm1' = '*'
    # Most cases use it to set up their workbooks.
    'scripts/Import-Vba.ps1' = '*'
    'scripts/Invoke-ExcelScript.ps1' = 'invoke-excelscript', 'timeout', 'range', 'userform', 'workbook-summary', 'macro-dialog'
    'scripts/Export-Vba.ps1' = 'cells-and-vba', 'userform'
    'scripts/Test-VbaCompile.ps1' = 'vba-compile', 'vba-compile-error'
    'scripts/Add-CustomUI.ps1' = 'input-rejected', 'workbook-summary', 'ribbon', 'ribbon-layout', 'publish'
    'scripts/Ribbon.psm1' = 'ribbon', 'ribbon-layout'
    'scripts/Invoke-RibbonButton.ps1' = 'ribbon', 'ribbon-layout'
    'scripts/Test-RibbonLayout.ps1' = 'ribbon-layout', 'publish'
    'scripts/Publish-Workbook.ps1' = 'input-rejected', 'publish'
    'scripts/Lock-VbaProject.ps1' = 'input-rejected', 'publish'
    'scripts/Get-WorkbookSummary.ps1' = 'workbook-summary', 'publish'
    'scripts/Find-CellError.ps1' = 'cell-error'
    'scripts/Test-Prerequisite.ps1' = 'prerequisite'
    'scripts/Save-WorkbookScreenshot.ps1' = 'screenshot'
}
$FullOnlyCases = 'vba-compile-error', 'macro-dialog', 'ribbon', 'ribbon-layout', 'publish', 'screenshot'

# Returns the cases that cover the files (paths in the skill folder, with / or \): All when every case is needed.
function Resolve-ChangedCase([string[]]$Files) {
    $all = $false
    $cases = New-Object System.Collections.Generic.List[string]
    foreach ($file in $Files) {
        $path = $file -replace '\\', '/'
        $listed = $CasesByFile[$path]
        if ($listed -eq '*' -or $path -eq 'tests/Test-XlWright.ps1' -or $path -like 'tests/data/*' -or
            ($null -eq $listed -and $path -like 'scripts/*')) { $all = $true }
        elseif ($null -ne $listed) { foreach ($c in $listed) { if (-not $cases.Contains($c)) { $cases.Add($c) } } }
    }
    [pscustomobject]@{ All = $all; Cases = $(if ($all) { @() } else { @($cases | Sort-Object) }) }
}

# The files of the skill changed since HEAD, as paths in the skill folder. Throws when Git cannot tell.
function Get-ChangedFile {
    $ErrorActionPreference = 'Continue'
    $files = foreach ($arguments in @(
            @('diff', '--name-only', '--no-renames', '--relative', 'HEAD'),
            @('ls-files', '--others', '--exclude-standard'))) {
        $output = @(& git -C $skill @arguments 2>&1)
        # stderr comes as error records; a warning there (line endings, for example) is not a file name.
        $messages = @($output | Where-Object { $_ -is [Management.Automation.ErrorRecord] })
        if ($LASTEXITCODE -ne 0) { throw ('git ' + $arguments[0] + ' failed: ' + ($messages -join ' ')) }
        $output | Where-Object { $_ -isnot [Management.Automation.ErrorRecord] }
    }
    @($files | ForEach-Object { [string]$_ } | Where-Object { $_ } | Sort-Object -Unique)
}

if ($Changed) {
    if ($Case) { Write-Output '-Changed cannot be combined with -Case.'; exit 2 }
    try { $files = Get-ChangedFile }
    catch { Write-Output ('-Changed could not list the changed files: ' + $_.Exception.Message); exit 2 }
    $selection = Resolve-ChangedCase $files
    Write-Output ('Changed files of the skill: ' + $(if ($files) { $files -join ', ' } else { 'none' }))
    if ($selection.All) {
        Write-Output ('Cases: all' + $(if ($Full) { '' } else { ' (without -Full, the cases that show Excel are not run)' }))
    }
    else {
        $skipped = @($selection.Cases | Where-Object { -not $Full -and $FullOnlyCases -contains $_ })
        $Case = @($selection.Cases | Where-Object { $skipped -notcontains $_ })
        if ($skipped) { Write-Output ('Not run, these need -Full: ' + ($skipped -join ', ')) }
        if (-not $Case) { Write-Output 'No case to run for these changes. Nothing was run.'; exit 0 }
        Write-Output ('Cases: ' + ($Case -join ', '))
    }
}

if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output 'Excel is running. Close it (or ask the user) before running the tests.'
    exit 2
}
Import-Module $module -Force
# The same reading of AccessVBOM as Test-Prerequisite.ps1, policy keys included.
if (-not (Get-ExcelPrerequisite).AccessVBOM.Enabled) {
    Write-Output 'AccessVBOM is not 1. Ask the user to trust access to the VBA project object model.'
    exit 2
}

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

# A workbook that only sets up a case is built once, by the first case that needs it, and copied after that:
# every start and exit of Excel takes seconds. A build that throws leaves nothing behind, so the next case builds again.
$fixtures = Join-Path $temp 'fixtures'
New-Item -ItemType Directory -Path $fixtures | Out-Null

function Copy-Fixture([string]$Name, [string]$Path, [scriptblock]$Build) {
    $fixture = Join-Path $fixtures $Name
    if (-not (Test-Path -LiteralPath $fixture)) {
        # The extension is kept: New-ExcelWorkbook takes the format from it.
        $building = Join-Path $fixtures ('building-' + $Name)
        Remove-Item -LiteralPath $building -Force -ErrorAction SilentlyContinue
        & $Build $building
        Move-Item -LiteralPath $building -Destination $fixture
    }
    Copy-Item -LiteralPath $fixture -Destination $Path
    $Path
}

function New-EmptyBook([string]$Path) {
    Copy-Fixture ('empty' + [IO.Path]::GetExtension($Path)) $Path { param($p) [void](New-ExcelWorkbook -Path $p) }
}

function New-BookWithVba([string]$Dir) {
    $source = Join-Path $Dir 'vba'
    Copy-Item -LiteralPath (Join-Path $data 'vba') -Destination $source -Recurse
    Copy-Fixture 'vba.xlsm' (Join-Path $Dir 'book.xlsm') {
        param($p)
        [void](New-ExcelWorkbook -Path $p)
        $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $p, '-SourceDir', $source)
        Assert 'Import-Vba exit 0' ($import.Exit -eq 0) $import.Output
        if ($import.Exit -ne 0) { throw 'The workbook with the VBA of tests\data\vba could not be built.' }
    }
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
    $referencesFile = Join-Path $dir 'export\references.json'
    $exportedReferences = if (Test-Path -LiteralPath $referencesFile) { [IO.File]::ReadAllText($referencesFile, $utf8) | ConvertFrom-Json } else { $null }
    $referenceNames = @($exportedReferences.References | ForEach-Object { $_.Name + ' ' + $_.Guid }) -join ','
    Assert 'references.json: the references of a new workbook, without the built-in ones' ($referenceNames -eq 'stdole {00020430-0000-0000-C000-000000000046},Office {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}') $referenceNames

    # Components that are not in the sources stay; Publish-Workbook.ps1 relies on it.
    # A reference in references.json that the workbook lacks is a warning, and only that one.
    $partial = Join-Path $dir 'partial'
    New-Item -ItemType Directory -Path (Join-Path $partial 'standard') | Out-Null
    [IO.File]::WriteAllText((Join-Path $partial 'standard\Extra.bas'), "Attribute VB_Name = ""Extra""`nOption Explicit`n", $utf8)
    [IO.File]::WriteAllText((Join-Path $partial 'references.json'), '{"References": [{"Name": "stdole", "Guid": "{00020430-0000-0000-C000-000000000046}", "Major": 2, "Minor": 0}, {"Name": "Scripting", "Guid": "{420b2830-e718-11cf-893d-00a0c9054228}", "Major": 1, "Minor": 0}]}', $utf8)
    $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', $partial)
    Assert 'partial import: exit 0' ($import.Exit -eq 0) $import.Output
    $missing = @([regex]::Matches($import.Output, 'Reference not in the workbook: (\S+)') | ForEach-Object { $_.Groups[1].Value }) -join ','
    Assert 'partial import: a warning for the missing reference only' ($missing -eq 'Scripting') $import.Output
    $names = Use-ExcelWorkbook -Path $book -ScriptBlock {
        param($xl, $wb)
        @($wb.VBProject.VBComponents | ForEach-Object { $_.Name + ':' + $_.Type } | Sort-Object) -join ','
    }
    Assert 'partial import: the components that are not in the sources stay' ($names -eq 'CalcOptions:2,Extra:1,Sheet1:100,ThisWorkbook:100,Util:1') $names
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
    [void](New-EmptyBook $plain)
    $p = Invoke-Compile $plain
    Assert 'workbook without VBA: compiled' ($p.Exit -eq 0 -and $p.Json.Compiled -eq $true) ($p.Json | ConvertTo-Json -Compress)
}

function New-BookWithMacros([string]$Dir, [string]$Source) {
    $book = Join-Path $Dir ((Split-Path -Leaf $Source) + '.xlsm')
    Copy-Fixture ('macros-' + ($Source -replace '[\\/]', '-') + '.xlsm') $book {
        param($p)
        [void](New-ExcelWorkbook -Path $p)
        $import = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $p, '-SourceDir', (Join-Path $data $Source))
        Assert ($Source + ': import exit 0') ($import.Exit -eq 0) $import.Output
        if ($import.Exit -ne 0) { throw ('The workbook with the VBA of ' + $Source + ' could not be built.') }
    }
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

Invoke-Case 'timeout' {
    $book = Join-Path (New-CaseDir 'timeout') 'book.xlsm'
    [void](New-EmptyBook $book)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    # The limit only has to be longer than starting Excel and opening the workbook (about 1 s).
    $h = Invoke-Runner $book 'hang' @('-EnableMacros', '-TimeoutSeconds', '8')
    $seconds = [int]$watch.Elapsed.TotalSeconds
    Assert 'timeout exit 1' ($h.Exit -eq 1) $h.Exit
    Assert 'timeout message' ([string]$h.Json.Error -like 'Timed out after 8 s*') $h.Json.Error
    Assert 'stopped within 60 s' ($seconds -lt 60) $seconds
}

Invoke-Case 'caller-killed' {
    $dir = New-CaseDir 'caller-killed'
    $book = Join-Path $dir 'book.xlsm'
    [void](New-EmptyBook $book)
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
    [void](New-EmptyBook $book)
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
    # Rejected after the workbook is open; the valid module in the same sources is not saved either.
    $src = Join-Path $dir 'src-no-document'
    New-Item -ItemType Directory -Path (Join-Path $src 'standard'), (Join-Path $src 'documents') -Force | Out-Null
    [IO.File]::WriteAllBytes((Join-Path $src 'standard\Good.bas'), $ascii.GetBytes("Attribute VB_Name = ""Good""`nOption Explicit`n"))
    [IO.File]::WriteAllBytes((Join-Path $src 'documents\NoSuchSheet.cls'), $ascii.GetBytes("Attribute VB_Name = ""NoSuchSheet""`nOption Explicit`n"))
    $r = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', $src)
    Assert 'Import-Vba rejects a document module the workbook does not have' ($r.Exit -ne 0 -and $r.Output -like '*No document module named NoSuchSheet*') $r.Output
    $src = Join-Path $dir 'src-bad-references'
    New-Item -ItemType Directory -Path (Join-Path $src 'standard') -Force | Out-Null
    [IO.File]::WriteAllBytes((Join-Path $src 'standard\Good.bas'), $ascii.GetBytes("Attribute VB_Name = ""Good""`nOption Explicit`n"))
    [IO.File]::WriteAllBytes((Join-Path $src 'references.json'), $ascii.GetBytes('{"References": [{"Name": "Scripting"}]}'))
    $r = Invoke-Script (Join-Path $scripts 'Import-Vba.ps1') @('-Path', $book, '-SourceDir', $src)
    Assert 'Import-Vba rejects a reference without a Guid' ($r.Exit -ne 0 -and $r.Output -like '*Expected References with a Name and a Guid*') $r.Output
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
    [void](New-EmptyBook $book)
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

# The scripts stop an Excel only through this check. A hidden powershell.exe stands in for Excel here.
Invoke-Case 'stop-own-process' {
    $child = Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile', '-Command', 'Start-Sleep -Seconds 60' -WindowStyle Hidden -PassThru
    try {
        $ticks = [XlWright.ExcelProcess]::StartTicks($child.Id)
        Assert 'start ticks of the process' ($ticks -eq $child.StartTime.ToUniversalTime().Ticks) $ticks
        $other = [string][XlWright.ExcelProcess]::StopIfSame($child.Id, $ticks - 1)
        Assert 'another start time: NotSame, not stopped' ($other -eq 'NotSame' -and -not $child.HasExited) ('returned ' + $other)
        $same = [string][XlWright.ExcelProcess]::StopIfSame($child.Id, $ticks)
        Assert 'same PID and start time: Stopped' ($same -eq 'Stopped' -and $child.WaitForExit(5000)) ('returned ' + $same)
        $gone = [string][XlWright.ExcelProcess]::StopIfSame($child.Id, $ticks)
        Assert 'a process that is gone: Gone' ($gone -eq 'Gone') ('returned ' + $gone)
    }
    finally { if (-not $child.HasExited) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue } }

    # A session that loaded the types of another version cannot replace them; the module must refuse to run.
    $stale = '$ErrorActionPreference = ''Stop''; Add-Type -Namespace XlWright -Name Dialog -MemberDefinition ''public static int Old;''; ' +
        'Import-Module ''' + $module + '''; ''imported'''
    $ErrorActionPreference = 'Continue'
    $output = & powershell.exe -NoProfile -Command $stale 2>&1 | Out-String
    $exit = $LASTEXITCODE
    $ErrorActionPreference = 'Stop'
    # The console wraps long lines anywhere, so the text is compared without white space.
    $flat = $output -replace '\s+', ''
    Assert 'old XlWright.Dialog in the session: the import fails with the reason' ($exit -ne 0 -and $flat -notmatch "'?imported'?$" -and
        $flat -like '*XlWright.DialogofanotherversionofXlWright.psm1*newpowershell.exe*') $output
}

# -Changed: which cases the changed files select. No Excel is started.
Invoke-Case 'changed-selection' {
    $show = { param($r) 'All=' + $r.All + ' Cases=' + (@($r.Cases) -join ',') }
    $r = Resolve-ChangedCase @('scripts\Find-CellError.ps1')
    Assert 'a listed script: its cases' ((& $show $r) -eq 'All=False Cases=cell-error') (& $show $r)
    $r = Resolve-ChangedCase @('scripts/Ribbon.psm1', 'scripts/Test-RibbonLayout.ps1')
    Assert 'two scripts: their cases once each, sorted' ((& $show $r) -eq 'All=False Cases=publish,ribbon,ribbon-layout') (& $show $r)
    foreach ($file in 'scripts/XlWright.psm1', 'scripts/New-Thing.ps1', 'tests/Test-XlWright.ps1', 'tests/data/vba/standard/Util.bas') {
        $r = Resolve-ChangedCase @('scripts/Find-CellError.ps1', $file)
        Assert ('every case for ' + $file) ((& $show $r) -eq 'All=True Cases=') (& $show $r)
    }
    $r = Resolve-ChangedCase @('SKILL.md', 'references/vba.md')
    Assert 'documents only: no case' ((& $show $r) -eq 'All=False Cases=') (& $show $r)

    # The table must name only cases that exist, and the cases that need -Full must be those inside if ($Full).
    $text = [IO.File]::ReadAllText($PSCommandPath)
    $names = @([regex]::Matches($text, "(?m)^\s*Invoke-Case '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    $listed = @($CasesByFile.Values | ForEach-Object { $_ } | Where-Object { $_ -ne '*' } | Sort-Object -Unique)
    $unknown = @($listed + $FullOnlyCases | Where-Object { $names -notcontains $_ })
    Assert 'the table names only existing cases' ($unknown.Count -eq 0) ($unknown -join ',')
    $fullBlock = $text.Substring([regex]::Match($text, '(?m)^if \(\$Full\) \{').Index)
    $inFull = @([regex]::Matches($fullBlock, "(?m)^\s*Invoke-Case '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    Assert 'the cases that need -Full are those inside if ($Full)' ((($inFull | Sort-Object) -join ',') -eq (($FullOnlyCases | Sort-Object) -join ',')) ($inFull -join ',')
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
    [void](New-EmptyBook $book)
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
    [void](New-EmptyBook $other)
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
    [void](New-EmptyBook $book)
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
    [void](New-EmptyBook $book)
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
    [void](New-EmptyBook $copy)
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
            $book = New-BookWithMacros $dir ('vba-broken\' + $case.Name)
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
            @{ Name = 'Yes/No MsgBox'; Book = $book; Macro = 'Macros.Ask'; Error = '*cannot be closed without choosing an answer: probe: question*' },
            # WM_CLOSE would close this one by choosing Cancel.
            @{ Name = 'OK/Cancel MsgBox'; Book = $book; Macro = 'Macros.AskOkCancel'; Error = '*cannot be closed without choosing an answer: probe: proceed*' })
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
                # The x button would close an OK/Cancel MsgBox by choosing Cancel, so it is not closed either.
                $okCancelJson = Join-Path $dir 'ok-cancel.json'
                $okCancel = Invoke-Script (Join-Path $scripts 'Invoke-RibbonButton.ps1') @('-Path', $book, '-TabLabel', $S.Tab, '-ButtonLabel', $S.OkCancelButton, '-OutFile', $okCancelJson)
                $okCancelRes = [IO.File]::ReadAllText($okCancelJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
                $detail = $okCancelRes | ConvertTo-Json -Compress -Depth 5
                Assert 'OK/Cancel dialog: exit 1' ($okCancel.Exit -eq 1) $detail
                Assert 'OK/Cancel dialog: an answer to choose, its Excel stopped' ([string]$okCancelRes.Error -like 'A dialog could not be closed without choosing an answer. Stopped the Excel instance of this script*') $detail
                $left = Wait-NoExcel 20
                Assert 'OK/Cancel dialog: no Excel left' ($left.Count -eq 0) ('PIDs ' + ($left -join ','))
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
        [void](New-EmptyBook $book)
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

    Invoke-Case 'screenshot' -WatchForeground {
        $dir = New-CaseDir 'screenshot'
        $book = Join-Path $dir 'tool.xlsx'
        [void](New-EmptyBook $book)
        Use-ExcelWorkbook -Path $book -Save -ScriptBlock {
            param($xl, $wb)
            $main = $wb.Worksheets.Item(1)
            $main.Name = 'Main'
            $main.Range('A1:B2').Interior.Color = 16711935  # magenta
            $far = $wb.Worksheets.Add([Type]::Missing, $main)
            $far.Name = 'Far'
            $far.Range('AA200:AB201').Interior.Color = 65280  # green
            $hidden = $wb.Worksheets.Add([Type]::Missing, $far)
            $hidden.Name = 'Hidden'
            $hidden.Visible = 0  # xlSheetHidden
            $main.Activate()
        }
        $hash = Get-FileHashText $book
        if (-not ('XlWrightTest.Pixels' -as [type])) {
            Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
namespace XlWrightTest {
    public static class Pixels {
        // Pixels close to pure magenta and pure green: the fills of the test sheets. The ribbon icons use red,
        // so red is not used.
        public static int[] MagentaGreen(string file) {
            using (var bmp = new System.Drawing.Bitmap(file)) {
                int magenta = 0, green = 0;
                for (int y = 0; y < bmp.Height; y += 2) for (int x = 0; x < bmp.Width; x += 2) {
                    var c = bmp.GetPixel(x, y);
                    if (c.R > 200 && c.G < 60 && c.B > 200) magenta++;
                    if (c.G > 200 && c.R < 60 && c.B < 60) green++;
                }
                return new int[] { magenta, green };
            }
        }
    }
}
'@
        }
        $shooter = Join-Path $scripts 'Save-WorkbookScreenshot.ps1'
        $shoot = {
            param([string]$Name, [string[]]$Extra)
            $json = Join-Path $dir "$Name.json"
            $png = Join-Path $dir "$Name.png"
            $run = Invoke-Script $shooter (@('-Path', $book, '-OutFile', $json, '-ImageFile', $png) + $Extra)
            $text = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8)
            $colors = $null
            if (Test-Path -LiteralPath $png) { $colors = [XlWrightTest.Pixels]::MagentaGreen($png) }
            [pscustomobject]@{ Exit = $run.Exit; Json = ($text | ConvertFrom-Json); Text = $text; Png = $png; Colors = $colors }
        }
        # The script does not use the clipboard; whatever text it holds stays.
        $clipboard = Get-Clipboard -Format Text -Raw

        $default = & $shoot 'default' @()
        Assert 'default: exit 0, no error, the active sheet from A1' ($default.Exit -eq 0 -and $null -eq $default.Json.Error -and
            $default.Json.Path -eq $book -and $default.Json.Sheet -eq 'Main' -and [string]$default.Json.VisibleRange -like 'A1:*') $default.Text
        Assert 'default: 1280 x 800 PNG' ($default.Json.Width -eq 1280 -and $default.Json.Height -eq 800 -and $default.Json.ImageFile -eq $default.Png -and
            (Test-Path -LiteralPath $default.Png)) $default.Text
        # A hidden Excel window is drawn black: the fill shows only if the window was really shown.
        Assert 'default: the magenta fill of Main drawn, not the green of Far' ($default.Colors -and $default.Colors[0] -gt 500 -and $default.Colors[1] -eq 0) ($default.Colors -join ',')

        $far = & $shoot 'far' @('-Sheet', 'Far', '-Range', 'AA200', '-Width', '900', '-Height', '600')
        Assert '-Sheet, -Range: that sheet, scrolled to the range' ($far.Exit -eq 0 -and $far.Json.Sheet -eq 'Far' -and [string]$far.Json.VisibleRange -like 'AA200:*') $far.Text
        Assert '-Width, -Height: 900 x 600' ($far.Json.Width -eq 900 -and $far.Json.Height -eq 600) $far.Text
        Assert '-Range: the green fill drawn, not the magenta' ($far.Colors -and $far.Colors[1] -gt 500 -and $far.Colors[0] -eq 0) ($far.Colors -join ',')
        # Goto also activates the sheet of the range, so -Sheet is checked without -Range.
        $sheetOnly = & $shoot 'sheet-only' @('-Sheet', 'Far')
        Assert '-Sheet without -Range: Far from A1, the magenta of Main not drawn' ($sheetOnly.Exit -eq 0 -and $sheetOnly.Json.Sheet -eq 'Far' -and
            [string]$sheetOnly.Json.VisibleRange -like 'A1:*' -and $sheetOnly.Colors -and $sheetOnly.Colors[0] -eq 0) ($sheetOnly.Text + ' ' + ($sheetOnly.Colors -join ','))

        foreach ($bad in @(
            @{ Name = 'hidden sheet'; Extra = @('-Sheet', 'Hidden'); Error = '*hidden*Hidden*' },
            @{ Name = 'missing sheet'; Extra = @('-Sheet', 'NoSuch'); Error = '*Not a sheet*NoSuch*Main, Far, Hidden*' },
            @{ Name = 'bad range'; Extra = @('-Range', 'notarange!!'); Error = '*Not a range*notarange!!*' })) {
            $r = & $shoot ($bad.Name -replace ' ', '-') $bad.Extra
            Assert ($bad.Name + ': exit 1 and Error, no image') ($r.Exit -eq 1 -and [string]$r.Json.Error -like $bad.Error -and -not (Test-Path -LiteralPath $r.Png)) $r.Text
        }
        Assert 'clipboard text unchanged' ((Get-Clipboard -Format Text -Raw) -eq $clipboard) 'changed'
        Assert 'workbook unchanged' ((Get-FileHashText $book) -eq $hash) 'hash changed'
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
