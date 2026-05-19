param(
    [string]$InputWorkbook = "$env:USERPROFILE\Downloads\LCCC-NFTs-OPENSEA.xlsx",
    [string]$OutputCsv = "$env:USERPROFILE\Downloads\LCCC-NFTs-OPENSEA-with-page-urls.csv",
    [string]$OutputFolder = "D:\github\ron-english-nfts\light-cult-crypto-club\images",
    [string]$CatalogueBaseUrl = "https://ronenglish-archive.github.io/ron-english-nfts/light-cult-crypto-club/items",
    [string]$OpenSeaContract = "0xbe85fbd182af91290be7293438ae67549638189f",
    [switch]$DownloadImages,
    [switch]$UseLargerImageUrls,
    [int]$MaxRows = 0
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "============================================================"
Write-Host "Light Cult Crypto Club spreadsheet processor v3"
Write-Host "============================================================"

if ($DownloadImages) {
    Write-Host "MODE: GENERATE URL CSV + DOWNLOAD IMAGES"
} else {
    Write-Host "MODE: GENERATE URL CSV ONLY"
}

if ($UseLargerImageUrls) {
    Write-Host "IMAGE URL MODE: remove query strings such as ?w=350"
} else {
    Write-Host "IMAGE URL MODE: keep exact image URLs from spreadsheet"
}

Write-Host ""
Write-Host "Input workbook: $InputWorkbook"
Write-Host "Output CSV:     $OutputCsv"
Write-Host "Image folder:   $OutputFolder"
Write-Host ""

if (!(Test-Path $InputWorkbook)) {
    throw "Input workbook not found: $InputWorkbook"
}

$curlExe = "$env:SystemRoot\System32\curl.exe"
if ($DownloadImages -and !(Test-Path $curlExe)) {
    throw "Could not find curl.exe at $curlExe"
}

New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null

function Get-CleanTokenNumber {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    # Handles:
    # Item #4013 Media
    # #4013
    # 4013
    $m = [regex]::Match($Value, "\#?\s*(\d+)")
    if (!$m.Success) { return $null }

    return [int]$m.Groups[1].Value
}

function Get-DownloadUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return "" }

    $clean = $Url.Trim()

    if ($UseLargerImageUrls) {
        return ($clean -split "\?")[0]
    }

    return $clean
}

function Get-ImageExtensionFromUrl {
    param([string]$Url)

    try {
        $cleanUrl = ($Url -split "\?")[0]
        $ext = [System.IO.Path]::GetExtension($cleanUrl)
        if ($ext -match "^\.(jpg|jpeg|png|webp|gif)$") {
            if ($ext -ieq ".jpeg") { return ".jpg" }
            return $ext.ToLowerInvariant()
        }
    } catch {}

    return ".jpg"
}

function Close-ExcelSafely {
    param($Excel, $Workbook)

    try {
        if ($Workbook) { $Workbook.Close($false) | Out-Null }
    } catch {}

    try {
        if ($Excel) { $Excel.Quit() | Out-Null }
    } catch {}

    try {
        if ($Workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null }
    } catch {}

    try {
        if ($Excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null }
    } catch {}

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Export-CsvSafely {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    try {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        return $Path
    }
    catch {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $folder = Split-Path $Path -Parent
        $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $alt = Join-Path $folder "$base-$stamp.csv"

        Write-Host ""
        Write-Host "Could not write the normal CSV, probably because it is open in Excel."
        Write-Host "Writing a timestamped copy instead:"
        Write-Host "  $alt"
        Write-Host ""

        $Rows | Export-Csv -Path $alt -NoTypeInformation -Encoding UTF8
        return $alt
    }
}

function Download-WithCurl {
    param(
        [string]$Url,
        [string]$Destination
    )

    $tempFile = "$Destination.download"

    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }

    $curlArgs = @(
        "-L",
        "--fail",
        "--silent",
        "--show-error",
        "--retry", "3",
        "--retry-delay", "2",
        "--connect-timeout", "30",
        "--max-time", "180",
        "-A", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome Safari",
        "-H", "Referer: https://opensea.io/",
        "-o", $tempFile,
        $Url
    )

    $output = & $curlExe @curlArgs 2>&1
    $exit = $LASTEXITCODE

    if ($exit -ne 0) {
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
        return "curl failed with exit code $exit. $output"
    }

    if (!(Test-Path $tempFile)) {
        return "curl did not create a file"
    }

    $fileInfo = Get-Item $tempFile
    if ($fileInfo.Length -lt 100) {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return "downloaded file was too small: $($fileInfo.Length) bytes"
    }

    Move-Item -Path $tempFile -Destination $Destination -Force
    return "OK"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($InputWorkbook)
    $sheet = $workbook.Worksheets.Item(1)
    $used = $sheet.UsedRange
    $rowCount = $used.Rows.Count

    if ($MaxRows -gt 0 -and $MaxRows -lt $rowCount) {
        Write-Host "Rows detected: $rowCount"
        Write-Host "TEST MODE: only processing first $MaxRows rows"
        $rowCount = $MaxRows
    } else {
        Write-Host "Rows detected: $rowCount"
    }

    Write-Host "Reading column A as image URL and column B as item label..."
    Write-Host ""

    $results = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $downloaded = 0
    $skippedExisting = 0
    $skippedDuplicate = 0
    $failed = 0
    $missingNumber = 0
    $missingImageUrl = 0
    $failureSamples = New-Object System.Collections.Generic.List[string]

    for ($r = 1; $r -le $rowCount; $r++) {
        $imageUrlRaw = [string]$sheet.Cells.Item($r, 1).Value2
        $itemLabel = [string]$sheet.Cells.Item($r, 2).Value2

        $tokenNumber = Get-CleanTokenNumber -Value $itemLabel

        if ($null -eq $tokenNumber) {
            $missingNumber++
            $results.Add([pscustomobject]@{
                ItemNumber = ""
                OriginalLabel = $itemLabel
                OpenSeaItemURL = ""
                CataloguePageURL = ""
                ImageURLFromSpreadsheet = $imageUrlRaw
                DownloadImageURL = ""
                LocalImageFile = ""
                Status = "Skipped: could not extract item number"
            })
            continue
        }

        $tokenPadded = "{0:D4}" -f $tokenNumber
        $openSeaUrl = "https://opensea.io/item/ethereum/$OpenSeaContract/$tokenNumber"
        $cataloguePageUrl = "$CatalogueBaseUrl/lccc-$tokenPadded.html"
        $downloadUrl = Get-DownloadUrl -Url $imageUrlRaw
        $ext = Get-ImageExtensionFromUrl -Url $downloadUrl
        $localFile = "LCCC-$tokenPadded$ext"
        $destination = Join-Path $OutputFolder $localFile
        $status = "URLs generated"

        if ($DownloadImages) {
            if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
                $missingImageUrl++
                $status = "Skipped image download: no image URL"
                $localFile = ""
            }
            elseif ($seen.ContainsKey($tokenNumber)) {
                $skippedDuplicate++
                $status = "Skipped duplicate token number in spreadsheet"
                $localFile = ""
            }
            elseif (Test-Path $destination) {
                $seen[$tokenNumber] = $true
                $skippedExisting++
                $status = "Skipped: image already exists"
            }
            else {
                $seen[$tokenNumber] = $true
                $downloadResult = Download-WithCurl -Url $downloadUrl -Destination $destination

                if ($downloadResult -eq "OK") {
                    $downloaded++
                    $status = "Downloaded"
                } else {
                    $failed++
                    $status = "Failed download: $downloadResult"
                    $localFile = ""

                    if ($failureSamples.Count -lt 8) {
                        $failureSamples.Add("Item #$tokenNumber - $downloadResult - URL: $downloadUrl")
                    }
                }
            }
        }

        $results.Add([pscustomobject]@{
            ItemNumber = $tokenNumber
            OriginalLabel = $itemLabel
            OpenSeaItemURL = $openSeaUrl
            CataloguePageURL = $cataloguePageUrl
            ImageURLFromSpreadsheet = $imageUrlRaw
            DownloadImageURL = $downloadUrl
            LocalImageFile = $localFile
            Status = $status
        })

        if (($r % 100) -eq 0 -and $DownloadImages) {
            Write-Host "Processed $r of $rowCount rows... downloaded=$downloaded, existing=$skippedExisting, failed=$failed"
        }
        elseif (($r % 500) -eq 0) {
            Write-Host "Processed $r of $rowCount rows..."
        }
    }

    $sortedResults = $results | Sort-Object {
        if ($_.ItemNumber -match "^\d+$") { [int]$_.ItemNumber } else { 999999999 }
    }

    $writtenCsv = Export-CsvSafely -Rows $sortedResults -Path $OutputCsv

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "DONE"
    Write-Host "============================================================"
    Write-Host "Rows processed: $rowCount"
    Write-Host "Rows without item numbers: $missingNumber"
    Write-Host "Rows without image URLs: $missingImageUrl"
    Write-Host "CSV written to:"
    Write-Host "  $writtenCsv"

    if ($DownloadImages) {
        Write-Host ""
        Write-Host "IMAGE DOWNLOAD SUMMARY"
        Write-Host "Downloaded images: $downloaded"
        Write-Host "Skipped existing images: $skippedExisting"
        Write-Host "Skipped duplicate tokens: $skippedDuplicate"
        Write-Host "Failed downloads: $failed"
        Write-Host "Image folder:"
        Write-Host "  $OutputFolder"

        if ($failureSamples.Count -gt 0) {
            Write-Host ""
            Write-Host "FIRST FAILURE SAMPLES"
            foreach ($sample in $failureSamples) {
                Write-Host "  $sample"
            }
        }
    }
    else {
        Write-Host ""
        Write-Host "No images were downloaded because -DownloadImages was not used."
        Write-Host "To test-download only the first 25 rows, run:"
        Write-Host '  & "$env:USERPROFILE\Downloads\process-lccc-opensea-spreadsheet-v3.ps1" -DownloadImages -MaxRows 25'
        Write-Host ""
        Write-Host "To download all images, run:"
        Write-Host '  & "$env:USERPROFILE\Downloads\process-lccc-opensea-spreadsheet-v3.ps1" -DownloadImages'
    }

    Write-Host ""
}
finally {
    Close-ExcelSafely -Excel $excel -Workbook $workbook
}
