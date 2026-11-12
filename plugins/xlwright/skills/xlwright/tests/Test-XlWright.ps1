<#
.SYNOPSIS
Checks the xlwright scripts against the real Excel on this machine.

.DESCRIPTION
Every case creates its own workbooks in a new temporary folder and checks explicit expected values.
Nothing outside that folder is changed. The default cases keep Excel hidden. -Full adds the cases
that show Excel and the VBE on the screen (ribbon click, ribbon layout, publish with VBA lock); do not use the
mouse or keyboard while they run.

Preconditions: no Excel process is running, and AccessVBOM is 1. Otherwise nothing runs (exit 2).
After each case no Excel process may remain. If one does, the run stops without touching it.

-Case runs only the named cases, e.g. -Case timeout,userform.

The result is written as UTF-8 JSON to -OutFile and summarized on the console.
Exit code: 0 all passed, 1 a check failed, 2 preconditions not met or the run was stopped.
The temporary folder is removed when everything passed, unless -KeepTemp is given.
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

function Invoke-Case([string]$Name, [scriptblock]$Body) {
    if ($script:stopped) { return }
    if ($Case -and $Case -notcontains $Name) { return }
    $script:checks = New-Object System.Collections.Generic.List[object]
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $errorText = $null
    try { & $Body } catch { $errorText = $_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage }
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
    Invoke-Case 'ribbon' {
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

    Invoke-Case 'ribbon-layout' {
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
        Assert 'rejected XML: exit 1, tab missing' ($bad.Exit -eq 1 -and @($badRes.Mismatches | Where-Object { $_.Path -eq $S.Tab -and $_.Actual -eq 'missing' }).Count -eq 1) ($badRes | ConvertTo-Json -Compress -Depth 8)
    }

    Invoke-Case 'publish' {
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
        # The ribbon of the locked copy is built and its callbacks (getEnabled) still run.
        $json = Join-Path $dir 'dist-layout.json'
        $layout = Invoke-Script (Join-Path $scripts 'Test-RibbonLayout.ps1') @('-Path', $dist, '-OutFile', $json)
        $res = [IO.File]::ReadAllText($json, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $dynPath = $S.Tab + '/' + $S.Group + '/' + $S.DynButton
        Assert 'dist: ribbon layout matches the XML' ($layout.Exit -eq 0) ($res | ConvertTo-Json -Compress -Depth 8)
        Assert 'dist: getEnabled callback ran' (@($res.Dynamic | Where-Object { $_.Path -eq $dynPath -and $_.Item -eq 'Enabled' -and $_.Actual -eq $false }).Count -eq 1) ($res.Dynamic | ConvertTo-Json -Compress)
    }
}

# ---------------------------------------------------------------------------------------------
$failedCases = @($results | Where-Object { -not $_.Ok })
$summary = [ordered]@{
    Ok = ($failedCases.Count -eq 0 -and -not $stopped); Full = [bool]$Full; Stopped = $stopped
    TempDir = $temp; Cases = $results.ToArray()
}
if ($OutFile) { [IO.File]::WriteAllText($OutFile, ([pscustomobject]$summary | ConvertTo-Json -Depth 8), $utf8) }
if ($summary.Ok -and -not $KeepTemp) {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}
else { Write-Output ('Temporary files kept: ' + $temp) }
Write-Output ('{0} of {1} cases passed{2}' -f ($results.Count - $failedCases.Count), $results.Count,
    $(if ($stopped) { '; stopped because Excel was left running' } else { '' }))
if ($stopped) { exit 2 }
if ($failedCases.Count -gt 0) { exit 1 }
exit 0
