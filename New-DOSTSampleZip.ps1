[CmdletBinding()]
param(
    [string]$OutputRoot
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $PSScriptRoot 'samples'
}

function New-DirectoryIfMissing {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path)
    }
}

function Write-SimplePdf {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Text
    )

    $escapedText = $Text.Replace('\', '\\').Replace('(', '\(').Replace(')', '\)')
    $objects = @(
        '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj',
        '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj',
        '3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >> endobj',
        '4 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj'
    )

    $stream = "BT /F1 20 Tf 72 720 Td ($escapedText) Tj ET"
    $objects += "5 0 obj << /Length $($stream.Length) >> stream`n$stream`nendstream endobj"

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('%PDF-1.4')
    $offsets = New-Object System.Collections.Generic.List[int]

    foreach ($object in $objects) {
        $offsets.Add($builder.Length)
        [void]$builder.AppendLine($object)
    }

    $xrefStart = $builder.Length
    [void]$builder.AppendLine('xref')
    [void]$builder.AppendLine("0 $($objects.Count + 1)")
    [void]$builder.AppendLine('0000000000 65535 f ')
    foreach ($offset in $offsets) {
        [void]$builder.AppendLine(('{0:0000000000} 00000 n ' -f $offset))
    }
    [void]$builder.AppendLine("trailer << /Size $($objects.Count + 1) /Root 1 0 R >>")
    [void]$builder.AppendLine('startxref')
    [void]$builder.AppendLine($xrefStart)
    [void]$builder.AppendLine('%%EOF')

    [System.IO.File]::WriteAllText($Path, $builder.ToString(), [System.Text.Encoding]::ASCII)
}

function Write-SampleXlsText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Label
    )

    @(
        "Parent`tLabel`tAmount",
        "$Parent`t$Label`t123.45",
        "$Parent`t$Label extra`t67.89"
    ) | Set-Content -LiteralPath $Path -Encoding ASCII
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$sampleRoot = Join-Path $OutputRoot "DOST_Sample_$stamp"
$payloadRoot = Join-Path $sampleRoot 'payload'
$nestedRoot = Join-Path $sampleRoot 'nested_payload'
New-DirectoryIfMissing -Path $payloadRoot
New-DirectoryIfMissing -Path $nestedRoot

foreach ($parent in @('TD', 'BMO')) {
    $parentPath = Join-Path $payloadRoot $parent
    New-DirectoryIfMissing -Path $parentPath

    "Sample text file for $parent" | Set-Content -LiteralPath (Join-Path $parentPath "$parent-notes.txt") -Encoding UTF8
    Write-SimplePdf -Path (Join-Path $parentPath "$parent-statement.pdf") -Text "$parent statement"
    Write-SampleXlsText -Path (Join-Path $parentPath "$parent-data.xls") -Parent $parent -Label 'outer'

    $pngBytes = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=')
    [System.IO.File]::WriteAllBytes((Join-Path $parentPath "$parent-image.png"), $pngBytes)

    $nestedParentPath = Join-Path $nestedRoot $parent
    New-DirectoryIfMissing -Path $nestedParentPath
    "Nested text file for $parent" | Set-Content -LiteralPath (Join-Path $nestedParentPath "$parent-nested.txt") -Encoding UTF8
    Write-SimplePdf -Path (Join-Path $nestedParentPath "$parent-nested.pdf") -Text "$parent nested PDF"
}

$nestedZip = Join-Path $payloadRoot 'NestedDocuments.zip'
if (Test-Path -LiteralPath $nestedZip) {
    Remove-Item -LiteralPath $nestedZip -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($nestedRoot, $nestedZip)

$finalZip = Join-Path $OutputRoot "DOST_Sample_$stamp.zip"
New-DirectoryIfMissing -Path $OutputRoot
if (Test-Path -LiteralPath $finalZip) {
    Remove-Item -LiteralPath $finalZip -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($payloadRoot, $finalZip)

Write-Host "Created sample zip: $finalZip"
