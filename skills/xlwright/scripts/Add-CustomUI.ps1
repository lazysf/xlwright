<#
.SYNOPSIS
Embeds or replaces customUI/customUI14.xml (Office 2010+ RibbonX) in a closed .xlsm package.

.DESCRIPTION
Excel's COM API cannot define a ribbon, so the XML is written straight into the OOXML zip.
The workbook must be closed. If customUI14.xml already exists its content is replaced;
the package-level relationship in _rels/.rels is added only when missing.
The updated package is checked in a sibling temporary file before replacing the workbook.
The new part relies on the default "xml" content type in [Content_Types].xml; the script stops if it is missing.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$XmlPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$partName = 'customUI/customUI14.xml'
$relType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
$relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$Path = (Resolve-Path -LiteralPath $Path).ProviderPath
# Parse first so that malformed XML never reaches the workbook.
[xml]$ui = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $XmlPath).ProviderPath, $utf8)

function Read-PackageXml($Zip, [string]$Part) {
    $entry = $Zip.GetEntry($Part)
    if (!$entry) { throw "Missing package part: $Part" }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-PackageXml($Zip, [string]$Part, [xml]$Document) {
    $old = $Zip.GetEntry($Part)
    if ($old) { $old.Delete() }
    $writer = New-Object System.IO.StreamWriter($Zip.CreateEntry($Part).Open(), $utf8)
    try { $Document.Save($writer) } finally { $writer.Dispose() }
}

function Get-RibbonRelationship([xml]$Relationships) {
    $Relationships.DocumentElement.ChildNodes |
        Where-Object { $_.Type -eq $relType -and $_.Target.TrimStart('/') -eq $partName }
}

$tempPath = Join-Path ([IO.Path]::GetDirectoryName($Path)) ('.customui-' + [guid]::NewGuid() + '.tmp')
try {
    [IO.File]::Copy($Path, $tempPath)
    $zip = [System.IO.Compression.ZipFile]::Open($tempPath, 'Update')
    try {
        $ct = Read-PackageXml $zip '[Content_Types].xml'
        if (-not ($ct.Types.Default | Where-Object { $_.Extension -eq 'xml' })) {
            throw '[Content_Types].xml has no Default for the "xml" extension'
        }

        $replacing = $null -ne $zip.GetEntry($partName)
        $rels = Read-PackageXml $zip '_rels/.rels'
        $exists = @(Get-RibbonRelationship $rels).Count -gt 0
        Write-PackageXml $zip $partName $ui
        if (!$exists) {
            $node = $rels.CreateElement('Relationship', $relNs)
            $id = 'rIdCustomUI14'
            if ($id -in $rels.DocumentElement.ChildNodes.Id) { $id += [guid]::NewGuid().ToString('N') }
            $node.SetAttribute('Id', $id)
            $node.SetAttribute('Type', $relType)
            $node.SetAttribute('Target', $partName)
            [void]$rels.DocumentElement.AppendChild($node)
            Write-PackageXml $zip '_rels/.rels' $rels
        }
    }
    finally { $zip.Dispose() }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    try {
        $written = Read-PackageXml $zip $partName
        $rels = Read-PackageXml $zip '_rels/.rels'
        if ($written.DocumentElement.OuterXml -cne $ui.DocumentElement.OuterXml -or
            @(Get-RibbonRelationship $rels).Count -ne 1) {
            throw 'Could not verify the custom UI part and its package relationship'
        }
    }
    finally { $zip.Dispose() }

    [IO.File]::Replace($tempPath, $Path, [NullString]::Value)
    "customUI14.xml: " + $(if ($replacing) { 'replaced' } else { 'added' })
    "relationship: " + $(if ($exists) { 'kept' } else { 'added' })
}
finally {
    if ([IO.File]::Exists($tempPath)) { [IO.File]::Delete($tempPath) }
}
