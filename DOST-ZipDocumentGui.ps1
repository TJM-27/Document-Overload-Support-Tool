<# 
Runs a WinForms GUI that extracts a source zip, recursively expands nested zip
files, groups documents by type, and exports the grouped results.
#>

[CmdletBinding()]
param(
    [string]$InventoryOnlyZip,
    [string]$InventoryOnlyFolder,
    [string]$SmokeExportRoot,
    [switch]$SmokeMergePdf,
    [switch]$SmokeCombineExcel
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:WorkRoot = Join-Path $script:AppRoot 'work'
$script:LibRoot = Join-Path $script:AppRoot 'lib'
$script:Inventory = @()
$script:CurrentExtractRoot = $null
$script:SourceZip = $null
$script:TabData = @{}
$script:ImportWarnings = New-Object System.Collections.Generic.List[string]
$script:ExpandedNestedZipPaths = New-Object 'System.Collections.Generic.HashSet[string]'
$script:ParentGroupingRoot = $null
$script:ParentGroupAliases = @{}
$script:ExportProgressCallback = $null
$script:RunWithExportProgress = $null
$script:ColorGreenDark = [System.Drawing.Color]::FromArgb(19, 78, 52)
$script:ColorGreenMedium = [System.Drawing.Color]::FromArgb(42, 121, 82)
$script:ColorGreenAccent = [System.Drawing.Color]::FromArgb(72, 154, 105)
$script:ColorGreenPale = [System.Drawing.Color]::FromArgb(228, 242, 231)
$script:ColorGreenPaler = [System.Drawing.Color]::FromArgb(242, 249, 243)
$script:ColorGreenMuted = [System.Drawing.Color]::FromArgb(207, 230, 211)
$script:ColorGreenBorder = [System.Drawing.Color]::FromArgb(93, 135, 103)
$script:ColorTextDark = [System.Drawing.Color]::FromArgb(24, 48, 34)

function New-DirectoryIfMissing {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path)
    }
}

function Update-ExportProgress {
    param(
        [Parameter(Mandatory)][string]$Activity,
        [int]$Current = 0,
        [int]$Total = 0
    )

    if ($script:ExportProgressCallback) {
        & $script:ExportProgressCallback $Activity $Current $Total
    }
}

function Format-RemainingTime {
    param([double]$Seconds)

    if ($Seconds -lt 0 -or [double]::IsNaN($Seconds) -or [double]::IsInfinity($Seconds)) {
        return 'calculating...'
    }

    $timeSpan = [TimeSpan]::FromSeconds([Math]::Max(0, [Math]::Round($Seconds)))
    if ($timeSpan.TotalHours -ge 1) {
        return '{0}h {1}m {2}s' -f [int]$timeSpan.TotalHours, $timeSpan.Minutes, $timeSpan.Seconds
    }

    if ($timeSpan.TotalMinutes -ge 1) {
        return '{0}m {1}s' -f $timeSpan.Minutes, $timeSpan.Seconds
    }

    return '{0}s' -f $timeSpan.Seconds
}

function Invoke-ExportWithProgress {
    param(
        [Parameter(Mandatory)][string]$Activity,
        [Parameter(Mandatory)][scriptblock]$Operation
    )

    if ($script:RunWithExportProgress -is [scriptblock]) {
        return & $script:RunWithExportProgress $Activity $Operation
    }

    return & $Operation
}

function Clear-DostWorkFolder {
    param([switch]$PreserveCurrent)

    New-DirectoryIfMissing -Path $script:WorkRoot

    $removedCount = 0
    $errors = New-Object System.Collections.Generic.List[string]
    $preservePath = $null

    if ($PreserveCurrent -and $script:CurrentExtractRoot -and (Test-Path -LiteralPath $script:CurrentExtractRoot)) {
        $preservePath = (Resolve-Path -LiteralPath $script:CurrentExtractRoot).Path.TrimEnd('\')
    }

    foreach ($item in (Get-ChildItem -LiteralPath $script:WorkRoot -Force -ErrorAction SilentlyContinue)) {
        try {
            $itemPath = $item.FullName.TrimEnd('\')
            if ($preservePath -and $itemPath.Equals($preservePath, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            Remove-Item -LiteralPath $item.FullName -Recurse -Force
            $removedCount++
        } catch {
            $errors.Add("Could not remove $($item.FullName): $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        RemovedCount = $removedCount
        Errors = $errors.ToArray()
    }
}

function ConvertTo-SafeName {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$MaxLength = 120,
        [string]$Fallback = 'Item'
    )

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $chars = foreach ($char in $Name.ToCharArray()) {
        if ($invalid -contains $char) { '_' } else { $char }
    }

    $safe = (-join $chars).Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = $Fallback
    }

    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength).Trim()
    }

    return $safe
}

function ConvertTo-SafeWorksheetName {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string[]]$ExistingNames = @()
    )

    $safe = [regex]::Replace($Name, '[:\\/\?\*\[\]]', '_').Trim("'").Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = 'Sheet'
    }

    $baseName = if ($safe.Length -gt 31) { $safe.Substring(0, 31) } else { $safe }
    $candidate = $baseName
    $counter = 2

    while ($ExistingNames -contains $candidate) {
        $suffix = "_$counter"
        $maxBase = 31 - $suffix.Length
        if ($baseName.Length -gt $maxBase) {
            $candidate = $baseName.Substring(0, $maxBase) + $suffix
        } else {
            $candidate = $baseName + $suffix
        }
        $counter++
    }

    return $candidate
}

function New-UniquePath {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$FileName
    )

    New-DirectoryIfMissing -Path $Directory

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    $safeBase = ConvertTo-SafeName -Name $baseName
    $safeExt = ConvertTo-SafeName -Name $extension -Fallback ''
    $candidate = Join-Path $Directory ($safeBase + $safeExt)
    $counter = 2

    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $safeBase, $counter, $safeExt)
        $counter++
    }

    return $candidate
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$Path
    )

    $baseUri = New-Object System.Uri((Resolve-Path -LiteralPath $BasePath).Path.TrimEnd('\') + '\')
    $pathUri = New-Object System.Uri((Resolve-Path -LiteralPath $Path).Path)
    return [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($pathUri).ToString()).Replace('/', '\')
}

function Test-DirectoryHasDirectCombinedExportFiles {
    param([Parameter(Mandatory)][string]$Path)

    $combinedExportExtensions = @('.pdf', '.xls', '.xlsx')
    $directFiles = @(Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue)
    foreach ($file in $directFiles) {
        if ($combinedExportExtensions -contains $file.Extension.ToLowerInvariant()) {
            return $true
        }
    }

    return $false
}

function Get-ParentGroupingRoot {
    param(
        [Parameter(Mandatory)][string]$ExtractRoot,
        [Parameter(Mandatory)][string]$SourceBaseName
    )

    $current = (Resolve-Path -LiteralPath $ExtractRoot).Path

    while ($true) {
        $childDirectories = @(Get-ChildItem -LiteralPath $current -Directory -ErrorAction SilentlyContinue | Sort-Object Name)

        if ($childDirectories.Count -gt 1) {
            if (Test-DirectoryHasDirectCombinedExportFiles -Path $current) {
                $parent = Split-Path -Parent $current
                $extractRootFull = (Resolve-Path -LiteralPath $ExtractRoot).Path

                if ($current.Equals($extractRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $script:ParentGroupAliases[(Split-Path -Leaf $current)] = ConvertTo-SafeName -Name $SourceBaseName
                }

                return $parent
            }

            return $current
        }

        if ($childDirectories.Count -eq 1) {
            $current = $childDirectories[0].FullName
            continue
        }

        return (Resolve-Path -LiteralPath $ExtractRoot).Path
    }
}

function Get-ParentGroup {
    param(
        [Parameter(Mandatory)][string]$GroupingRoot,
        [Parameter(Mandatory)][string]$FilePath
    )

    $groupingRootFull = (Resolve-Path -LiteralPath $GroupingRoot).Path.TrimEnd('\') + '\'
    $fileFull = (Resolve-Path -LiteralPath $FilePath).Path
    if (-not $fileFull.StartsWith($groupingRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return '_Unsorted'
    }

    $relative = Get-RelativePath -BasePath $GroupingRoot -Path $FilePath
    $parts = @($relative -split '[\\/]+' | Where-Object { $_ })
    if ($parts.Count -lt 2) {
        return '_Unsorted'
    }

    if ($script:ParentGroupAliases.ContainsKey($parts[0])) {
        return $script:ParentGroupAliases[$parts[0]]
    }

    return ConvertTo-SafeName -Name $parts[0] -Fallback '_Unsorted'
}

function Get-DocumentType {
    param([Parameter(Mandatory)][string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).TrimStart('.').ToUpperInvariant()
    if ($extension -in @('XLS', 'XLSX')) {
        return 'EXCEL'
    }
    if ([string]::IsNullOrWhiteSpace($extension)) {
        return 'NO_EXTENSION'
    }
    return $extension
}

function Test-ZipReadable {
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [ref]$ErrorMessage
    )

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        [void]$archive.Entries.Count
        return $true
    } catch {
        if ($ErrorMessage) {
            $ErrorMessage.Value = $_.Exception.Message
        }
        return $false
    } finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }
}

function Expand-ZipFileSafe {
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$Destination,
        [switch]$IsSourceZip
    )

    $readError = $null
    if (-not (Test-ZipReadable -ZipPath $ZipPath -ErrorMessage ([ref]$readError))) {
        if ($IsSourceZip) {
            throw "The selected file is not a readable zip archive: $ZipPath`r`n`r`nDetails: $readError"
        }

        throw "Nested zip could not be read: $ZipPath`r`nDetails: $readError"
    }

    New-DirectoryIfMissing -Path $Destination
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
    } catch {
        if ($IsSourceZip) {
            throw "The selected zip opened, but extraction failed: $ZipPath`r`n`r`nDetails: $($_.Exception.Message)"
        }

        throw "Nested zip extraction failed: $ZipPath`r`nDetails: $($_.Exception.Message)"
    }
}

function Expand-NestedZips {
    param([Parameter(Mandatory)][string]$ExtractRoot)

    $processed = New-Object 'System.Collections.Generic.HashSet[string]'
    $expandedCount = 0
    $foundNewZip = $true

    while ($foundNewZip) {
        $foundNewZip = $false
        $zipFiles = Get-ChildItem -LiteralPath $ExtractRoot -Recurse -File -Filter '*.zip' |
            Sort-Object FullName

        foreach ($zipFile in $zipFiles) {
            if ($processed.Contains($zipFile.FullName)) {
                continue
            }

            [void]$processed.Add($zipFile.FullName)
            $destinationName = ConvertTo-SafeName -Name ($zipFile.BaseName + '__unzipped')
            $destination = Join-Path $zipFile.DirectoryName $destinationName
            $suffix = 2
            while (Test-Path -LiteralPath $destination) {
                $destination = Join-Path $zipFile.DirectoryName ("{0}_{1}" -f $destinationName, $suffix)
                $suffix++
            }

            try {
                Expand-ZipFileSafe -ZipPath $zipFile.FullName -Destination $destination
                [void]$script:ExpandedNestedZipPaths.Add($zipFile.FullName)
                $expandedCount++
                $foundNewZip = $true
            } catch {
                Remove-Item -LiteralPath $destination -Recurse -Force -ErrorAction SilentlyContinue
                $script:ImportWarnings.Add($_.Exception.Message)
            }
        }
    }

    return $expandedCount
}

function Initialize-ImportWorkspace {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$SourceBaseName
    )

    New-DirectoryIfMissing -Path $script:WorkRoot
    $script:ImportWarnings.Clear()
    $script:ExpandedNestedZipPaths.Clear()
    $cleanup = Clear-DostWorkFolder
    foreach ($cleanupError in $cleanup.Errors) {
        $script:ImportWarnings.Add($cleanupError)
    }
    $script:CurrentExtractRoot = $null
    $script:ParentGroupingRoot = $null
    $script:ParentGroupAliases = @{}
    $script:Inventory = @()

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $sourceBase = ConvertTo-SafeName -Name $SourceBaseName
    $extractRoot = Join-Path $script:WorkRoot ("{0}_{1}" -f $stamp, $sourceBase)
    New-DirectoryIfMissing -Path $extractRoot

    return $extractRoot
}

function Copy-FolderContents {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$DestinationFolder
    )

    foreach ($item in (Get-ChildItem -LiteralPath $SourceFolder -Force)) {
        Copy-Item -LiteralPath $item.FullName -Destination $DestinationFolder -Recurse -Force
    }
}

function Complete-InventoryFromWorkRoot {
    param(
        [Parameter(Mandatory)][string]$ExtractRoot,
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$SourceType,
        [Parameter(Mandatory)][string]$SourceBaseName
    )

    $nestedCount = Expand-NestedZips -ExtractRoot $ExtractRoot
    $groupingRoot = Get-ParentGroupingRoot -ExtractRoot $ExtractRoot -SourceBaseName $SourceBaseName

    $items = foreach ($file in (Get-ChildItem -LiteralPath $ExtractRoot -Recurse -File | Sort-Object FullName)) {
        if ($script:ExpandedNestedZipPaths.Contains($file.FullName)) {
            continue
        }

        $type = Get-DocumentType -Path $file.FullName
        [pscustomobject]@{
            FileName = $file.Name
            BaseName = $file.BaseName
            Extension = $file.Extension.ToLowerInvariant()
            Type = $type
            ParentGroup = Get-ParentGroup -GroupingRoot $groupingRoot -FilePath $file.FullName
            FullPath = $file.FullName
            RelativePath = Get-RelativePath -BasePath $ExtractRoot -Path $file.FullName
        }
    }

    $script:CurrentExtractRoot = $ExtractRoot
    $script:SourceZip = $SourcePath
    $script:ParentGroupingRoot = $groupingRoot
    $script:Inventory = @($items)

    return [pscustomobject]@{
        SourceType = $SourceType
        ExtractRoot = $ExtractRoot
        ParentGroupingRoot = $groupingRoot
        NestedZipCount = $nestedCount
        ItemCount = $script:Inventory.Count
        WarningCount = $script:ImportWarnings.Count
    }
}

function Import-ZipInventory {
    param([Parameter(Mandatory)][string]$ZipPath)

    if (-not (Test-Path -LiteralPath $ZipPath)) {
        throw "Zip file not found: $ZipPath"
    }

    if ([System.IO.Path]::GetExtension($ZipPath).ToLowerInvariant() -ne '.zip') {
        throw 'Please select a .zip file.'
    }

    $sourceBaseName = [System.IO.Path]::GetFileNameWithoutExtension($ZipPath)
    $extractRoot = Initialize-ImportWorkspace -SourcePath $ZipPath -SourceBaseName $sourceBaseName
    Expand-ZipFileSafe -ZipPath $ZipPath -Destination $extractRoot -IsSourceZip
    return Complete-InventoryFromWorkRoot -ExtractRoot $extractRoot -SourcePath $ZipPath -SourceType 'Zip' -SourceBaseName $sourceBaseName
}

function Import-FolderInventory {
    param([Parameter(Mandatory)][string]$FolderPath)

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        throw "Folder not found: $FolderPath"
    }

    $folder = Get-Item -LiteralPath $FolderPath
    $extractRoot = Initialize-ImportWorkspace -SourcePath $folder.FullName -SourceBaseName $folder.Name
    Copy-FolderContents -SourceFolder $folder.FullName -DestinationFolder $extractRoot
    return Complete-InventoryFromWorkRoot -ExtractRoot $extractRoot -SourcePath $folder.FullName -SourceType 'Folder' -SourceBaseName $folder.Name
}

function Write-ExportLog {
    param(
        [Parameter(Mandatory)][string]$DestinationRoot,
        [Parameter(Mandatory)][string]$ExportName,
        [Parameter(Mandatory)][string[]]$Lines
    )

    New-DirectoryIfMissing -Path $DestinationRoot
    $logPath = Join-Path $DestinationRoot ("{0}_ExportLog_{1}.txt" -f $ExportName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $Lines | Set-Content -LiteralPath $logPath -Encoding UTF8
    return $logPath
}

function Select-Folder {
    param([string]$Description = 'Choose an export destination')

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function Export-OriginalFiles {
    param(
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][string]$TypeName,
        [Parameter(Mandatory)][string]$DestinationRoot
    )

    $log = New-Object System.Collections.Generic.List[string]
    $typeFolder = Join-Path $DestinationRoot (ConvertTo-SafeName -Name $TypeName)
    New-DirectoryIfMissing -Path $typeFolder
    $total = @($Items).Count
    $current = 0
    Update-ExportProgress -Activity "Exporting $TypeName files" -Current 0 -Total $total

    foreach ($item in $Items) {
        $current++
        try {
            $parentFolder = Join-Path $typeFolder (ConvertTo-SafeName -Name $item.ParentGroup -Fallback '_Unsorted')
            $target = New-UniquePath -Directory $parentFolder -FileName $item.FileName
            Copy-Item -LiteralPath $item.FullPath -Destination $target
            $log.Add("COPIED`t$($item.RelativePath)`t$target")
        } catch {
            $log.Add("ERROR`t$($item.RelativePath)`t$($_.Exception.Message)")
        }

        Update-ExportProgress -Activity "Exporting $TypeName files" -Current $current -Total $total
    }

    $logPath = Write-ExportLog -DestinationRoot $typeFolder -ExportName $TypeName -Lines $log.ToArray()
    return [pscustomobject]@{ OutputPath = $typeFolder; LogPath = $logPath }
}

function Test-ExcelAvailable {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        return $true
    } catch {
        return $false
    }
}

function Close-ComObject {
    param($Object)

    if ($null -ne $Object) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        } catch {
            # Ignore COM cleanup failures during shutdown.
        }
    }
}

function New-ExcelTempCopy {
    param(
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][string]$TempRoot,
        [Parameter(Mandatory)][int]$Index
    )

    New-DirectoryIfMissing -Path $TempRoot
    $extension = if ($Item.Extension) { $Item.Extension } else { '.xls' }
    $tempPath = Join-Path $TempRoot ("source_{0:000000}{1}" -f $Index, $extension)
    Copy-Item -LiteralPath $Item.FullPath -Destination $tempPath -Force
    return $tempPath
}

function Copy-WorksheetToWorkbook {
    param(
        [Parameter(Mandatory)]$Excel,
        [Parameter(Mandatory)]$SourceSheet,
        [Parameter(Mandatory)]$TargetSheet,
        [Parameter(Mandatory)][string]$SheetName
    )

    $usedRange = $null

    $TargetSheet.Name = $SheetName

    try {
        $usedRange = $SourceSheet.UsedRange
        if ($null -ne $usedRange) {
            [void]$usedRange.Copy($TargetSheet.Range('A1'))
            try {
                $Excel.CutCopyMode = 0
            } catch {
                # Copy succeeded; clearing Excel's clipboard state is best effort.
            }
        }
    } finally {
        Close-ComObject -Object $usedRange
    }
}

function Export-CombinedExcel {
    param(
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][string]$DestinationRoot
    )

    if (-not (Test-ExcelAvailable)) {
        throw 'Microsoft Excel is required for combined Excel export.'
    }

    $outputRoot = Join-Path $DestinationRoot 'EXCEL_Combined'
    New-DirectoryIfMissing -Path $outputRoot
    $log = New-Object System.Collections.Generic.List[string]
    $tempRoot = Join-Path $script:WorkRoot ("excel-combine-temp_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $sourceIndex = 0
    $total = @($Items).Count
    $current = 0
    Update-ExportProgress -Activity 'Combining Excel files' -Current 0 -Total $total

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        foreach ($group in ($Items | Group-Object ParentGroup | Sort-Object Name)) {
            $parentName = ConvertTo-SafeName -Name $group.Name -Fallback '_Unsorted'
            $outputPath = New-UniquePath -Directory $outputRoot -FileName ($parentName + '.xlsx')
            $combinedWorkbook = $excel.Workbooks.Add()
            $placeholder = $combinedWorkbook.Worksheets.Item(1)
            while ($combinedWorkbook.Worksheets.Count -gt 1) {
                $combinedWorkbook.Worksheets.Item($combinedWorkbook.Worksheets.Count).Delete()
            }
            $createdSheets = New-Object System.Collections.Generic.List[string]
            $nextSheetIsPlaceholder = $true

            foreach ($item in ($group.Group | Sort-Object FileName)) {
                $sourceWorkbook = $null
                $tempSourcePath = $null
                $current++
                try {
                    $sourceIndex++
                    $tempSourcePath = New-ExcelTempCopy -Item $item -TempRoot $tempRoot -Index $sourceIndex
                    $sourceWorkbook = $excel.Workbooks.Open($tempSourcePath, 0, $true)
                    $visibleSheetCount = 0

                    foreach ($sourceSheet in $sourceWorkbook.Worksheets) {
                        $targetSheet = $null
                        try {
                            if ($sourceSheet.Visible -ne -1) {
                                continue
                            }

                            $visibleSheetCount++
                            $desiredSheetName = if ($sourceWorkbook.Worksheets.Count -gt 1) {
                                "$($item.BaseName) - $($sourceSheet.Name)"
                            } else {
                                $item.BaseName
                            }
                            $sheetName = ConvertTo-SafeWorksheetName -Name $desiredSheetName -ExistingNames $createdSheets.ToArray()

                            if ($nextSheetIsPlaceholder) {
                                $targetSheet = $placeholder
                                $nextSheetIsPlaceholder = $false
                            } else {
                                $targetSheet = $combinedWorkbook.Worksheets.Add()
                            }

                            Copy-WorksheetToWorkbook -Excel $excel -SourceSheet $sourceSheet -TargetSheet $targetSheet -SheetName $sheetName
                            $createdSheets.Add($sheetName)
                            $log.Add("ADDED`t$($item.RelativePath)`t$outputPath`t$sheetName`tSource sheet: $($sourceSheet.Name)")
                        } finally {
                            Close-ComObject -Object $targetSheet
                            Close-ComObject -Object $sourceSheet
                        }
                    }

                    if ($visibleSheetCount -eq 0) {
                        $log.Add("SKIPPED`t$($item.RelativePath)`tNo visible worksheet found")
                    }
                } catch {
                    $log.Add("ERROR`t$($item.RelativePath)`t$($_.Exception.Message)")
                } finally {
                    if ($null -ne $sourceWorkbook) {
                        $sourceWorkbook.Close($false)
                        Close-ComObject -Object $sourceWorkbook
                    }
                    Update-ExportProgress -Activity 'Combining Excel files' -Current $current -Total $total
                }
            }

            if ($createdSheets.Count -gt 0) {
                $combinedWorkbook.SaveAs($outputPath, 51)
                $log.Add("SAVED`t$($group.Name)`t$outputPath")
            } else {
                $combinedWorkbook.Close($false)
                $log.Add("SKIPPED`t$($group.Name)`tNo worksheets were exported")
                continue
            }

            $combinedWorkbook.Close($true)
            Close-ComObject -Object $placeholder
            Close-ComObject -Object $combinedWorkbook
        }
    } finally {
        if ($null -ne $excel) {
            $excel.Quit()
            Close-ComObject -Object $excel
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    $logPath = Write-ExportLog -DestinationRoot $outputRoot -ExportName 'EXCEL_Combined' -Lines $log.ToArray()
    return [pscustomobject]@{ OutputPath = $outputRoot; LogPath = $logPath }
}

function Get-NuGetPackageVersion {
    param([Parameter(Mandatory)][string]$PackageId)

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $indexUrl = "https://api.nuget.org/v3-flatcontainer/$($PackageId.ToLowerInvariant())/index.json"
    $index = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing
    return @($index.versions | Where-Object { $_ -notmatch '-' } | Select-Object -Last 1)[0]
}

function Install-NuGetPackageLibrary {
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][string]$DllName
    )

    New-DirectoryIfMissing -Path $script:LibRoot
    $packageRoot = Join-Path $script:LibRoot $PackageId
    $existingDll = Get-ChildItem -LiteralPath $packageRoot -Recurse -File -Filter $DllName -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '\\lib\\' } |
        Sort-Object FullName |
        Select-Object -First 1

    if ($existingDll) {
        return $existingDll.FullName
    }

    $version = Get-NuGetPackageVersion -PackageId $PackageId
    if (-not $version) {
        throw "Could not find a NuGet version for $PackageId."
    }

    $downloadDir = Join-Path $script:LibRoot '_downloads'
    New-DirectoryIfMissing -Path $downloadDir
    New-DirectoryIfMissing -Path $packageRoot

    $nupkgPath = Join-Path $downloadDir "$PackageId.$version.nupkg"
    $downloadUrl = "https://api.nuget.org/v3-flatcontainer/$($PackageId.ToLowerInvariant())/$version/$($PackageId.ToLowerInvariant()).$version.nupkg"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $nupkgPath -UseBasicParsing
    [System.IO.Compression.ZipFile]::ExtractToDirectory($nupkgPath, $packageRoot)

    $dll = Get-ChildItem -LiteralPath $packageRoot -Recurse -File -Filter $DllName |
        Where-Object { $_.FullName -match '\\lib\\' } |
        Sort-Object FullName |
        Select-Object -First 1

    if (-not $dll) {
        throw "Could not locate $DllName in $PackageId $version."
    }

    return $dll.FullName
}

function Ensure-ITextSharpLoaded {
    if ('iTextSharp.text.Document' -as [type]) {
        return
    }

    $packages = @('iTextSharp-LGPL', 'iTextSharp')
    $lastError = $null
    foreach ($package in $packages) {
        try {
            $dllPath = Install-NuGetPackageLibrary -PackageId $package -DllName 'itextsharp.dll'
            Add-Type -Path $dllPath
            return
        } catch {
            $lastError = $_
        }
    }

    throw "Unable to load iTextSharp for PDF export. Last error: $($lastError.Exception.Message)"
}

function Set-PdfCopyOutlines {
    param(
        [Parameter(Mandatory)]$PdfCopy,
        [Parameter(Mandatory)]$Bookmarks
    )

    $method = $PdfCopy.GetType().GetMethod('SetOutlines')
    if ($method) {
        $arguments = New-Object 'object[]' 1
        $arguments[0] = $Bookmarks
        [void]$method.Invoke($PdfCopy, $arguments)
        return
    }

    $property = $PdfCopy.GetType().GetProperty('Outlines')
    if ($property) {
        $property.SetValue($PdfCopy, $Bookmarks, $null)
        return
    }
}

function Export-MergedPdf {
    param(
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][string]$DestinationRoot
    )

    Ensure-ITextSharpLoaded
    $outputRoot = Join-Path $DestinationRoot 'PDF_Merged'
    New-DirectoryIfMissing -Path $outputRoot
    $log = New-Object System.Collections.Generic.List[string]
    $total = @($Items).Count
    $current = 0
    Update-ExportProgress -Activity 'Merging PDF files' -Current 0 -Total $total

    foreach ($group in ($Items | Group-Object ParentGroup | Sort-Object Name)) {
        $parentName = ConvertTo-SafeName -Name $group.Name -Fallback '_Unsorted'
        $outputPath = New-UniquePath -Directory $outputRoot -FileName ($parentName + '.pdf')
        $stream = [System.IO.File]::Create($outputPath)
        $document = New-Object iTextSharp.text.Document
        $copy = New-Object iTextSharp.text.pdf.PdfCopy($document, $stream)
        $bookmarks = New-Object System.Collections.ArrayList
        $pageNumber = 1
        $addedAny = $false

        try {
            $document.Open()
            foreach ($item in ($group.Group | Sort-Object FileName)) {
                $reader = $null
                $current++
                try {
                    $reader = New-Object iTextSharp.text.pdf.PdfReader($item.FullPath)
                    if ($reader.NumberOfPages -lt 1) {
                        $log.Add("SKIPPED`t$($item.RelativePath)`tPDF has no pages")
                        continue
                    }

                    $bookmark = New-Object System.Collections.Hashtable
                    $bookmark['Title'] = [string]$item.BaseName
                    $bookmark['Action'] = 'GoTo'
                    $bookmark['Page'] = [string]"$pageNumber Fit"
                    [void]$bookmarks.Add($bookmark)

                    for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
                        $importedPage = $copy.GetImportedPage($reader, $page)
                        $copy.AddPage($importedPage)
                        $pageNumber++
                    }

                    $copy.FreeReader($reader)
                    $addedAny = $true
                    $log.Add("ADDED`t$($item.RelativePath)`t$outputPath")
                } catch {
                    $log.Add("ERROR`t$($item.RelativePath)`t$($_.Exception.Message)")
                } finally {
                    if ($null -ne $reader) {
                        $reader.Close()
                    }
                    Update-ExportProgress -Activity 'Merging PDF files' -Current $current -Total $total
                }
            }

            if ($addedAny) {
                Set-PdfCopyOutlines -PdfCopy $copy -Bookmarks $bookmarks
            }
        } finally {
            $document.Close()
            $copy.Close()
            $stream.Close()
        }

        if (-not $addedAny) {
            Remove-Item -LiteralPath $outputPath -ErrorAction SilentlyContinue
            $log.Add("SKIPPED`t$($group.Name)`tNo PDFs were merged")
        } else {
            $log.Add("SAVED`t$($group.Name)`t$outputPath")
        }
    }

    $logPath = Write-ExportLog -DestinationRoot $outputRoot -ExportName 'PDF_Merged' -Lines $log.ToArray()
    return [pscustomobject]@{ OutputPath = $outputRoot; LogPath = $logPath }
}

function Show-Info {
    param([string]$Message, [string]$Title = 'DOST Zip Document GUI')
    [void][System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Show-Error {
    param([string]$Message, [string]$Title = 'DOST Zip Document GUI')
    [void][System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

function Set-GreenButtonStyle {
    param([Parameter(Mandatory)][System.Windows.Forms.Button]$Button)

    $Button.UseVisualStyleBackColor = $false
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.BackColor = $script:ColorGreenPale
    $Button.ForeColor = $script:ColorTextDark
    $Button.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $script:ColorGreenBorder
    $Button.FlatAppearance.MouseOverBackColor = $script:ColorGreenMuted
    $Button.FlatAppearance.MouseDownBackColor = $script:ColorGreenAccent
}

function New-RoundedRectanglePath {
    param(
        [Parameter(Mandatory)][System.Drawing.Rectangle]$Rectangle,
        [int]$Radius = 10
    )

    $diameter = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($Rectangle.X, $Rectangle.Y, $diameter, $diameter, 180, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Y, $diameter, $diameter, 270, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($Rectangle.X, $Rectangle.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-DostPaperStackIcon {
    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $outlinePen = New-Object System.Drawing.Pen($script:ColorGreenDark, 1.4)
        $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(85, 0, 0, 0))
        $backBrush = New-Object System.Drawing.SolidBrush($script:ColorGreenMuted)
        $middleBrush = New-Object System.Drawing.SolidBrush($script:ColorGreenPale)
        $frontBrush = New-Object System.Drawing.SolidBrush($script:ColorGreenPaler)
        $linePen = New-Object System.Drawing.Pen($script:ColorGreenAccent, 1)

        try {
            $graphics.FillRectangle($shadowBrush, 9, 7, 17, 21)
            $graphics.FillRectangle($backBrush, 5, 3, 17, 21)
            $graphics.DrawRectangle($outlinePen, 5, 3, 17, 21)
            $graphics.FillRectangle($middleBrush, 8, 6, 17, 21)
            $graphics.DrawRectangle($outlinePen, 8, 6, 17, 21)
            $graphics.FillRectangle($frontBrush, 11, 9, 17, 21)
            $graphics.DrawRectangle($outlinePen, 11, 9, 17, 21)
            $graphics.DrawLine($linePen, 14, 15, 25, 15)
            $graphics.DrawLine($linePen, 14, 19, 25, 19)
            $graphics.DrawLine($linePen, 14, 23, 22, 23)
        } finally {
            $outlinePen.Dispose()
            $shadowBrush.Dispose()
            $backBrush.Dispose()
            $middleBrush.Dispose()
            $frontBrush.Dispose()
            $linePen.Dispose()
        }

        return [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    } finally {
        $graphics.Dispose()
    }
}

function Set-DistinctTabStyle {
    param([Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl)

    $TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
    $TabControl.ItemSize = New-Object System.Drawing.Size(140, 34)
    $TabControl.Padding = New-Object System.Drawing.Point(18, 6)

    $TabControl.Add_DrawItem({
        param($sender, $eventArgs)

        $tab = $sender.TabPages[$eventArgs.Index]
        $isSelected = $eventArgs.Index -eq $sender.SelectedIndex
        $bounds = $eventArgs.Bounds
        $bounds.Inflate(-2, -2)

        $inactiveColors = @(
            [System.Drawing.Color]::FromArgb(228, 242, 231),
            [System.Drawing.Color]::FromArgb(216, 235, 220),
            [System.Drawing.Color]::FromArgb(237, 247, 239),
            [System.Drawing.Color]::FromArgb(207, 230, 211)
        )

        if ($isSelected) {
            $backColor = $script:ColorGreenDark
            $foreColor = [System.Drawing.Color]::White
            $fontStyle = [System.Drawing.FontStyle]::Bold
        } else {
            $backColor = $inactiveColors[$eventArgs.Index % $inactiveColors.Count]
            $foreColor = $script:ColorTextDark
            $fontStyle = [System.Drawing.FontStyle]::Regular
        }

        $brush = New-Object System.Drawing.SolidBrush($backColor)
        $textBrush = New-Object System.Drawing.SolidBrush($foreColor)
        $borderPen = New-Object System.Drawing.Pen($script:ColorGreenBorder)
        $font = New-Object System.Drawing.Font('Segoe UI', 9, $fontStyle)
        $textFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
            [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
            [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor
            [System.Windows.Forms.TextFormatFlags]::SingleLine

        try {
            $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $tabPath = New-RoundedRectanglePath -Rectangle $bounds -Radius 8
            try {
                $eventArgs.Graphics.FillPath($brush, $tabPath)
                $eventArgs.Graphics.DrawPath($borderPen, $tabPath)
            } finally {
                $tabPath.Dispose()
            }
            [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, $tab.Text, $font, $bounds, $foreColor, $textFlags)
        } finally {
            $brush.Dispose()
            $textBrush.Dispose()
            $borderPen.Dispose()
            $font.Dispose()
        }
    })
}

function New-TabPageForType {
    param(
        [Parameter(Mandatory)][string]$TypeName,
        [Parameter(Mandatory)][array]$Items
    )

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = "$TypeName ($($Items.Count))"
    $tab.Tag = $TypeName
    $tab.BackColor = $script:ColorGreenPaler

    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.Height = 42
    $topPanel.BackColor = $script:ColorGreenPale

    $exportFilesButton = New-Object System.Windows.Forms.Button
    $exportFilesButton.Text = 'Export Files...'
    $exportFilesButton.Width = 120
    $exportFilesButton.Height = 28
    $exportFilesButton.Left = 8
    $exportFilesButton.Top = 7
    Set-GreenButtonStyle -Button $exportFilesButton

    $exportCombinedButton = New-Object System.Windows.Forms.Button
    $exportCombinedButton.Width = 160
    $exportCombinedButton.Height = 28
    $exportCombinedButton.Left = 136
    $exportCombinedButton.Top = 7
    Set-GreenButtonStyle -Button $exportCombinedButton
    $exportCombinedButton.Visible = $TypeName -in @('EXCEL', 'PDF')
    if ($TypeName -eq 'EXCEL') {
        $exportCombinedButton.Text = 'Export Combined Excel...'
    } elseif ($TypeName -eq 'PDF') {
        $exportCombinedButton.Text = 'Export Merged PDF...'
    }

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'None'
    $grid.BackgroundColor = $script:ColorGreenPaler
    $grid.GridColor = $script:ColorGreenBorder
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $grid.RowTemplate.Height = 26
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:ColorGreenMedium
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:ColorGreenMedium
    $grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersHeight = 30
    $grid.DefaultCellStyle.BackColor = $script:ColorGreenPaler
    $grid.DefaultCellStyle.ForeColor = $script:ColorTextDark
    $grid.DefaultCellStyle.SelectionBackColor = $script:ColorGreenAccent
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:ColorGreenPale
    [void]$grid.Columns.Add('ParentGroup', 'Parent Folder')
    [void]$grid.Columns.Add('FileName', 'File Name')
    $grid.Columns['ParentGroup'].Width = 180
    $grid.Columns['FileName'].AutoSizeMode = 'Fill'

    foreach ($item in ($Items | Sort-Object ParentGroup, FileName)) {
        $rowIndex = $grid.Rows.Add($item.ParentGroup, $item.FileName)
        $grid.Rows[$rowIndex].Tag = $item
    }

    $exportFilesButton.Add_Click({
        try {
            $destination = Select-Folder
            if (-not $destination) { return }
            $result = Invoke-ExportWithProgress -Activity "Exporting $TypeName files" -Operation {
                Export-OriginalFiles -Items $Items -TypeName $TypeName -DestinationRoot $destination
            }
            Show-Info "Export complete.`r`n`r`nOutput: $($result.OutputPath)`r`nLog: $($result.LogPath)"
        } catch {
            Show-Error $_.Exception.Message
        }
    }.GetNewClosure())

    $exportCombinedButton.Add_Click({
        try {
            $destination = Select-Folder
            if (-not $destination) { return }

            if ($TypeName -eq 'EXCEL') {
                $result = Invoke-ExportWithProgress -Activity 'Combining Excel files' -Operation {
                    Export-CombinedExcel -Items $Items -DestinationRoot $destination
                }
            } elseif ($TypeName -eq 'PDF') {
                $result = Invoke-ExportWithProgress -Activity 'Merging PDF files' -Operation {
                    Export-MergedPdf -Items $Items -DestinationRoot $destination
                }
            } else {
                return
            }

            Show-Info "Export complete.`r`n`r`nOutput: $($result.OutputPath)`r`nLog: $($result.LogPath)"
        } catch {
            Show-Error $_.Exception.Message
        }
    }.GetNewClosure())

    $topPanel.Controls.Add($exportFilesButton)
    $topPanel.Controls.Add($exportCombinedButton)
    $tab.Controls.Add($grid)
    $tab.Controls.Add($topPanel)
    return $tab
}

function Refresh-Tabs {
    param([Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl)

    $TabControl.TabPages.Clear()
    $script:TabData.Clear()

    foreach ($group in ($script:Inventory | Group-Object Type | Sort-Object Name)) {
        $items = @($group.Group)
        $script:TabData[$group.Name] = $items
        [void]$TabControl.TabPages.Add((New-TabPageForType -TypeName $group.Name -Items $items))
    }
}

function Start-DostZipDocumentGui {
    New-DirectoryIfMissing -Path $script:WorkRoot
    New-DirectoryIfMissing -Path $script:LibRoot

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DOST - Document Overload Support Tool'
    $form.Width = 950
    $form.Height = 700
    $form.StartPosition = 'CenterScreen'
    $form.Icon = New-DostPaperStackIcon
    $form.BackColor = $script:ColorGreenPaler

    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.Height = 170
    $topPanel.BackColor = $script:ColorGreenPale

    $bannerLabel = New-Object System.Windows.Forms.Label
    $bannerLabel.Text = 'Document Overload Support Tool'
    $bannerLabel.Left = 10
    $bannerLabel.Top = 8
    $bannerLabel.Width = 900
    $bannerLabel.Height = 28
    $bannerLabel.Anchor = 'Top,Left,Right'
    $bannerLabel.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
    $bannerLabel.ForeColor = $script:ColorGreenDark
    $bannerLabel.BackColor = $script:ColorGreenPale

    $bannerSubLabel = New-Object System.Windows.Forms.Label
    $bannerSubLabel.Text = 'DOST - Zip ingest, document sorting, and export support'
    $bannerSubLabel.Left = 12
    $bannerSubLabel.Top = 38
    $bannerSubLabel.Width = 900
    $bannerSubLabel.Height = 20
    $bannerSubLabel.Anchor = 'Top,Left,Right'
    $bannerSubLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
    $bannerSubLabel.ForeColor = $script:ColorTextDark
    $bannerSubLabel.BackColor = $script:ColorGreenPale

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = 'Select Zip...'
    $selectButton.Width = 110
    $selectButton.Height = 30
    $selectButton.Left = 10
    $selectButton.Top = 70
    Set-GreenButtonStyle -Button $selectButton

    $selectFolderButton = New-Object System.Windows.Forms.Button
    $selectFolderButton.Text = 'Select Folder...'
    $selectFolderButton.Width = 120
    $selectFolderButton.Height = 30
    $selectFolderButton.Left = 130
    $selectFolderButton.Top = 70
    Set-GreenButtonStyle -Button $selectFolderButton

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = 'No zip or folder selected.'
    $sourceLabel.Left = 260
    $sourceLabel.Top = 75
    $sourceLabel.Width = 630
    $sourceLabel.Anchor = 'Top,Left,Right'
    $sourceLabel.ForeColor = $script:ColorTextDark
    $sourceLabel.BackColor = $script:ColorGreenPale

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = 'Select a zip file or folder to begin.'
    $statusLabel.Left = 10
    $statusLabel.Top = 100
    $statusLabel.Width = 880
    $statusLabel.Height = 20
    $statusLabel.Anchor = 'Top,Left,Right'
    $statusLabel.ForeColor = $script:ColorTextDark
    $statusLabel.BackColor = $script:ColorGreenPale

    $exportProgressBar = New-Object System.Windows.Forms.ProgressBar
    $exportProgressBar.Left = 10
    $exportProgressBar.Top = 130
    $exportProgressBar.Width = 300
    $exportProgressBar.Height = 24
    $exportProgressBar.Minimum = 0
    $exportProgressBar.Maximum = 100
    $exportProgressBar.Visible = $false

    $exportProgressLabel = New-Object System.Windows.Forms.Label
    $exportProgressLabel.Left = 320
    $exportProgressLabel.Top = 128
    $exportProgressLabel.Width = 570
    $exportProgressLabel.Height = 28
    $exportProgressLabel.Anchor = 'Top,Left,Right'
    $exportProgressLabel.ForeColor = $script:ColorGreenDark
    $exportProgressLabel.BackColor = $script:ColorGreenPale
    $exportProgressLabel.Visible = $false

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = 'Fill'
    $tabControl.BackColor = $script:ColorGreenPaler
    Set-DistinctTabStyle -TabControl $tabControl

    $script:RunWithExportProgress = {
        param(
            [Parameter(Mandatory)][string]$Activity,
            [Parameter(Mandatory)][scriptblock]$Operation
        )

        $started = Get-Date
        $lastActivity = $Activity

        $exportProgressBar.Value = 0
        $exportProgressBar.Visible = $true
        $exportProgressLabel.Text = "$Activity - starting..."
        $exportProgressLabel.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()

        $script:ExportProgressCallback = {
            param([string]$ProgressActivity, [int]$Current, [int]$Total)

            if (-not [string]::IsNullOrWhiteSpace($ProgressActivity)) {
                $lastActivity = $ProgressActivity
            }

            if ($Total -gt 0) {
                $exportProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
                $percent = [Math]::Min(100, [Math]::Max(0, [int](($Current / $Total) * 100)))
                $exportProgressBar.Value = $percent

                if ($Current -gt 0) {
                    $elapsedSeconds = ((Get-Date) - $started).TotalSeconds
                    $remainingSeconds = ($elapsedSeconds / $Current) * ($Total - $Current)
                    $remainingText = Format-RemainingTime -Seconds $remainingSeconds
                } else {
                    $remainingText = 'calculating...'
                }

                $exportProgressLabel.Text = "$lastActivity - $Current of $Total ($percent%). Remaining: $remainingText"
            } else {
                $exportProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
                $exportProgressLabel.Text = "$lastActivity - working..."
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        try {
            Update-ExportProgress -Activity $Activity -Current 0 -Total 0
            $result = & $Operation
            $exportProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            $exportProgressBar.Value = 100
            $exportProgressLabel.Text = "$Activity - complete."
            [System.Windows.Forms.Application]::DoEvents()
            return $result
        } finally {
            $script:ExportProgressCallback = $null
            Start-Sleep -Milliseconds 250
            $exportProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            $exportProgressBar.Visible = $false
            $exportProgressLabel.Visible = $false
        }
    }

    $loadSource = {
        param(
            [Parameter(Mandatory)][string]$SourceKind,
            [Parameter(Mandatory)][string]$SourcePath
        )

        try {
            $form.UseWaitCursor = $true
            $selectButton.Enabled = $false
            $selectFolderButton.Enabled = $false
            $statusLabel.Text = "Importing $($SourceKind.ToLowerInvariant()). Please wait..."
            [System.Windows.Forms.Application]::DoEvents()

            if ($SourceKind -eq 'Zip') {
                $summary = Import-ZipInventory -ZipPath $SourcePath
            } else {
                $summary = Import-FolderInventory -FolderPath $SourcePath
            }

            Refresh-Tabs -TabControl $tabControl

            $sourceLabel.Text = $SourcePath
            $warningText = if ($summary.WarningCount -gt 0) { " Warnings: $($summary.WarningCount)." } else { '' }
            $groupingFolder = Split-Path -Leaf $summary.ParentGroupingRoot
            $statusLabel.Text = "Loaded $($summary.ItemCount) files from $($summary.SourceType.ToLowerInvariant()). Expanded $($summary.NestedZipCount) nested zip files.$warningText Parent root: $groupingFolder. Work folder: $($summary.ExtractRoot)"

            if ($summary.WarningCount -gt 0) {
                Show-Info "Import completed with warnings:`r`n`r`n$($script:ImportWarnings -join "`r`n`r`n")"
            }
        } catch {
            Show-Error $_.Exception.Message
            $statusLabel.Text = 'Import failed.'
        } finally {
            $selectButton.Enabled = $true
            $selectFolderButton.Enabled = $true
            $form.UseWaitCursor = $false
        }
    }

    $selectButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Zip files (*.zip)|*.zip'
        $dialog.Title = 'Select source zip'

        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        & $loadSource 'Zip' $dialog.FileName
    })

    $selectFolderButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Select source folder'
        $dialog.ShowNewFolderButton = $false

        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }

        & $loadSource 'Folder' $dialog.SelectedPath
    })

    $form.Add_FormClosed({
        try {
            [void](Clear-DostWorkFolder)
        } catch {
            # Cleanup should never block application shutdown.
        }
    })

    $topPanel.Controls.Add($bannerLabel)
    $topPanel.Controls.Add($bannerSubLabel)
    $topPanel.Controls.Add($selectButton)
    $topPanel.Controls.Add($selectFolderButton)
    $topPanel.Controls.Add($sourceLabel)
    $topPanel.Controls.Add($statusLabel)
    $topPanel.Controls.Add($exportProgressBar)
    $topPanel.Controls.Add($exportProgressLabel)
    $form.Controls.Add($tabControl)
    $form.Controls.Add($topPanel)

    [System.Windows.Forms.Application]::EnableVisualStyles()
    [void][System.Windows.Forms.Application]::Run($form)
}

if ($InventoryOnlyZip -or $InventoryOnlyFolder) {
    if ($InventoryOnlyZip -and $InventoryOnlyFolder) {
        throw 'Use either -InventoryOnlyZip or -InventoryOnlyFolder, not both.'
    }

    if ($InventoryOnlyZip) {
        $summary = Import-ZipInventory -ZipPath $InventoryOnlyZip
    } else {
        $summary = Import-FolderInventory -FolderPath $InventoryOnlyFolder
    }

    "SourceType: $($summary.SourceType)"
    "ExtractRoot: $($summary.ExtractRoot)"
    "ParentGroupingRoot: $($summary.ParentGroupingRoot)"
    "NestedZipCount: $($summary.NestedZipCount)"
    "ItemCount: $($summary.ItemCount)"
    "WarningCount: $($summary.WarningCount)"
    if ($summary.WarningCount -gt 0) {
        $script:ImportWarnings | ForEach-Object { "Warning: $_" }
    }
    $script:Inventory |
        Group-Object Type |
        Sort-Object Name |
        ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }

    if ($SmokeExportRoot) {
        foreach ($group in ($script:Inventory | Group-Object Type | Sort-Object Name)) {
            $result = Export-OriginalFiles -Items @($group.Group) -TypeName $group.Name -DestinationRoot $SmokeExportRoot
            "Exported $($group.Name): $($result.OutputPath)"
        }

        if ($SmokeMergePdf) {
            $pdfItems = @($script:Inventory | Where-Object { $_.Type -eq 'PDF' })
            if ($pdfItems.Count -gt 0) {
                $result = Export-MergedPdf -Items $pdfItems -DestinationRoot $SmokeExportRoot
                "Merged PDFs: $($result.OutputPath)"
            }
        }

        if ($SmokeCombineExcel) {
            $excelItems = @($script:Inventory | Where-Object { $_.Type -eq 'EXCEL' })
            if ($excelItems.Count -gt 0) {
                $result = Export-CombinedExcel -Items $excelItems -DestinationRoot $SmokeExportRoot
                "Combined Excel: $($result.OutputPath)"
            }
        }
    }
} else {
    Start-DostZipDocumentGui
}
