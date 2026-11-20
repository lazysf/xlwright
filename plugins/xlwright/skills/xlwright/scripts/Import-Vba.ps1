<#
.SYNOPSIS
Imports UTF-8 (no BOM), LF .bas / .cls / .frm files and saves the workbook.

.DESCRIPTION
- All sources are checked for lossless CP932 conversion before opening the workbook.
- Temporary files use CP932 and CRLF. VBComponents.Import misreads an LF-only .cls as a standard module.
- An existing standard module, class module, or UserForm with the same name and type is replaced.
- A UserForm's referenced .frx is copied unchanged next to the temporary .frm before import.
- A document module (ThisWorkbook, sheets) cannot be imported. Its code is replaced instead:
  the class header and the leading Attribute lines are dropped and the rest is written to the
  existing module. Attribute lines inside procedures are not handled.
- The standard/, classes/, documents/, and forms/ directories determine the component type.
- Every existing component must match the source type before any code is changed.
- After import, each module's name and type are checked against the file.
- Components that are not in the sources stay in the workbook.
- references.json in SourceDir (written by Export-Vba.ps1), if present, is checked before opening the
  workbook. Each listed reference that the workbook does not have (matched by Guid, the version is not
  compared) is reported as a warning after the import; no reference is added or removed.

The work is done by Import-ExcelVba in XlWright.psm1, which a script block can also call to import, compile
and run macros in one Excel instance.

Output: one line per component: "imported: <name>" or "replaced document module: <name>".
On an error the workbook is not saved.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$SourceDir
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

# Bad sources fail here, before Excel starts. Read-VbaSource is not exported, so it is called in the module scope;
# Import-ExcelVba reads the sources again inside the session.
[void](& (Get-Module XlWright) { param($dir) Read-VbaSource $dir } $SourceDir)

Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
    param($xl, $wb)
    Import-ExcelVba -Workbook $wb -SourceDir $SourceDir
}
