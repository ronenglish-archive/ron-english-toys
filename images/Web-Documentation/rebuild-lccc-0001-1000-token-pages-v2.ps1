param(
    [string]$RepoRoot = "D:\github\ron-english-light-cult-0001-1000",
    [int]$StartToken = 1,
    [int]$EndToken = 1000,
    [int]$StartWorkNumber = 1811,
    [string]$PublicBaseUrl = "https://ronenglish-archive.github.io/ron-english-light-cult-0001-1000",
    [string]$NftsHubUrl = "https://ronenglish-archive.github.io/ron-english-nfts/",
    [string]$CatalogueUrl = "https://ronenglish-archive.github.io/ron-english-catalogue-raisonne/",
    [string]$OpenSeaContract = "0xbe85fbd182af91290be7293438ae67549638189f",
    [switch]$Apply
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "============================================================"
Write-Host "Light Cult Crypto Club token-number rebuild v2"
Write-Host "============================================================"
Write-Host "Repo:        $RepoRoot"
Write-Host "Token range: $StartToken-$EndToken"
Write-Host "Work IDs:    RE-$("{0:D6}" -f $StartWorkNumber) onward"
if ($Apply) {
    Write-Host "MODE:        APPLY CHANGES"
} else {
    Write-Host "MODE:        AUDIT ONLY - no HTML pages will be changed"
}
Write-Host ""

if (!(Test-Path $RepoRoot)) {
    throw "Repo root not found: $RepoRoot"
}

$ImagesRoot = Join-Path $RepoRoot "images"
if (!(Test-Path $ImagesRoot)) {
    throw "Images folder not found: $ImagesRoot"
}

$DataRoot = Join-Path $RepoRoot "data"
$GeneratedRoot = Join-Path $DataRoot "generated"
New-Item -ItemType Directory -Force -Path $DataRoot | Out-Null
New-Item -ItemType Directory -Force -Path $GeneratedRoot | Out-Null

function HtmlEncode {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function To-SitePath {
    param([string]$FullPath)

    $relative = $FullPath.Substring($RepoRoot.Length).TrimStart("\", "/")
    return ($relative -replace "\\", "/")
}

function Get-RangeStart {
    param([int]$Token)

    return [int](([math]::Floor(($Token - 1) / 100) * 100) + 1)
}

function Get-RangeEnd {
    param([int]$Token)

    return [int]([math]::Min((Get-RangeStart -Token $Token) + 99, $EndToken))
}

function Get-RangeSlug {
    param([int]$Token)

    $rs = [int](Get-RangeStart -Token $Token)
    $re = [int](Get-RangeEnd -Token $Token)
    return ("{0:D4}-{1:D4}" -f $rs, $re)
}

function Get-RangeTitle {
    param([int]$Token)

    $rs = [int](Get-RangeStart -Token $Token)
    $re = [int](Get-RangeEnd -Token $Token)
    return ("{0}-{1}" -f $rs, $re)
}

function Get-WorkId {
    param([int]$Token)

    $workNumber = $StartWorkNumber + ($Token - $StartToken)
    return "RE-{0:D6}" -f $workNumber
}

function Get-TokenFromFilename {
    param([string]$BaseName)

    # Best case: LCCC-0001, LCCC_0001, Light Cult...0001
    $m = [regex]::Match($BaseName, "(?i)(?:LCCC|Light\s*Cult\s*Crypto\s*Club|LightCultCryptoClub)[^\d]*(\d{1,5})")
    if ($m.Success) {
        return [int]$m.Groups[1].Value
    }

    # General fallback: first standalone number in the filename.
    $m = [regex]::Match($BaseName, "(?<!\d)(\d{1,5})(?!\d)")
    if ($m.Success) {
        return [int]$m.Groups[1].Value
    }

    return $null
}

function Get-FileScore {
    param(
        [System.IO.FileInfo]$File,
        [int]$Token
    )

    $score = 0
    $expectedFolder = Get-RangeSlug -Token $Token

    if ($File.Directory.Name -eq $expectedFolder) { $score += 50 }
    if ($File.BaseName -match ("(?i)^LCCC[-_\s]*0*" + $Token + "$")) { $score += 40 }
    if ($File.BaseName -match ("(?i)0*" + $Token + "$")) { $score += 10 }
    if ($File.Extension -ieq ".jpg" -or $File.Extension -ieq ".jpeg") { $score += 3 }
    if ($File.Extension -ieq ".png") { $score += 2 }
    if ($File.Extension -ieq ".webp") { $score += 1 }

    return $score
}

function Backup-CurrentSite {
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupRoot = "D:\github-backups\ron-english-light-cult-0001-1000\rebuild-light-cult-token-pages-v2-$stamp"

    New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

    Write-Host ""
    Write-Host "Creating backup:"
    Write-Host "  $backupRoot"

    $rootFiles = Get-ChildItem -Path $RepoRoot -File | Where-Object {
        $_.Name -ieq "index.html" -or $_.Name -like "gallery-*.html"
    }

    foreach ($file in $rootFiles) {
        Copy-Item -Path $file.FullName -Destination (Join-Path $backupRoot $file.Name) -Force
    }

    $itemsDir = Join-Path $RepoRoot "items"
    if (Test-Path $itemsDir) {
        Copy-Item -Path $itemsDir -Destination (Join-Path $backupRoot "items") -Recurse -Force
    }

    $manifest = Join-Path $DataRoot "light-cult-0001-1000-manifest.csv"
    if (Test-Path $manifest) {
        New-Item -ItemType Directory -Force -Path (Join-Path $backupRoot "data") | Out-Null
        Copy-Item -Path $manifest -Destination (Join-Path $backupRoot "data\light-cult-0001-1000-manifest.csv") -Force
    }

    return $backupRoot
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Build-TopNav {
    param(
        [string]$Prefix,
        [string]$CurrentRangeHref = ""
    )

    $rangeLink = ""
    if ($CurrentRangeHref -ne "") {
        $rangeLink = "      <a class=""nav-pill"" href=""$CurrentRangeHref"">Back to range</a>`n"
    }

    return @"
    <nav class="nav-pills" aria-label="Site navigation">
$rangeLink      <a class="nav-pill" href="${Prefix}index.html">This repo index</a>
      <a class="nav-pill" href="$NftsHubUrl">NFTs &amp; Digital Works</a>
      <a class="nav-pill" href="$CatalogueUrl">Catalogue Raisonné</a>
    </nav>
"@
}

function Build-IndexHtml {
    param([object[]]$Records)

    $cards = New-Object System.Collections.Generic.List[string]

    for ($rangeStart = $StartToken; $rangeStart -le $EndToken; $rangeStart += 100) {
        $rangeEnd = [int]([math]::Min($rangeStart + 99, $EndToken))
        $rangeSlug = "{0:D4}-{1:D4}" -f $rangeStart, $rangeEnd
        $rangeRecords = @($Records | Where-Object { $_.TokenNumber -ge $rangeStart -and $_.TokenNumber -le $rangeEnd -and $_.Status -eq "OK" })
        $count = $rangeRecords.Count
        $first = $rangeRecords | Select-Object -First 1

        $imgHtml = ""
        if ($first) {
            $imgHtml = "          <img src=""$($first.ImagePath)"" alt=""Light Cult Crypto Club tokens $rangeSlug"" loading=""lazy"">`n"
        }

        $cards.Add(@"
        <a class="range-card" href="gallery-$rangeSlug.html">
$imgHtml          <div class="range-card-body">
            <h2>Tokens $rangeStart-$rangeEnd</h2>
            <p>$count images</p>
          </div>
        </a>
"@)
    }

    $firstWork = Get-WorkId -Token $StartToken
    $lastWork = Get-WorkId -Token $EndToken
    $cardHtml = ($cards -join "`n")

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Light Cult Crypto Club - Tokens $("{0:D4}" -f $StartToken)-$("{0:D4}" -f $EndToken)</title>
  <link rel="stylesheet" href="css/styles.css" />
</head>

<body>
  <main class="wrap">

$(Build-TopNav -Prefix "")

    <header>
      <h1 class="page-title">Light Cult Crypto Club</h1>
      <p class="intro">Token-number archive: Light Cult Crypto Club #$("{0:D4}" -f $StartToken)-#$("{0:D4}" -f $EndToken).</p>
    </header>

    <section class="info-card" aria-label="Archive note">
      <p>
        This repository contains one split section of the Light Cult Crypto Club visual archive.
      </p>
      <p>
        Images are organized by NFT token number. Catalogue Work IDs are retained in order:
        #$("{0:D4}" -f $StartToken) = $firstWork through #$("{0:D4}" -f $EndToken) = $lastWork.
      </p>
    </section>

    <section aria-label="Token range index">
      <h2 class="section-title">Token Indexes</h2>
      <div class="range-grid">
$cardHtml
      </div>
    </section>

    <footer>
      Local manifest: <code>data/light-cult-0001-1000-manifest.csv</code>.
    </footer>

  </main>
</body>
</html>
"@
}

function Build-GalleryHtml {
    param(
        [int]$RangeStart,
        [int]$RangeEnd,
        [object[]]$Records
    )

    $rangeSlug = "{0:D4}-{1:D4}" -f $RangeStart, $RangeEnd
    $figures = New-Object System.Collections.Generic.List[string]

    $rangeRecords = @($Records | Where-Object {
        $_.TokenNumber -ge $RangeStart -and $_.TokenNumber -le $RangeEnd -and $_.Status -eq "OK"
    } | Sort-Object TokenNumber)

    foreach ($r in $rangeRecords) {
        $tokenPadded = "{0:D4}" -f $r.TokenNumber
        $alt = HtmlEncode "Light Cult Crypto Club #$tokenPadded, $($r.WorkID)"
        $figures.Add(@"
        <figure class="image-card">
          <a href="$($r.ItemPage)">
            <img src="$($r.ImagePath)" alt="$alt" loading="lazy">
            <figcaption class="work-id">#$tokenPadded<br><span>$($r.WorkID)</span></figcaption>
          </a>
        </figure>
"@)
    }

    $prevNext = New-Object System.Collections.Generic.List[string]
    $prevStart = $RangeStart - 100
    $nextStart = $RangeStart + 100

    $prevNext.Add('      <a href="index.html">All token ranges in this repo</a>')
    if ($prevStart -ge $StartToken) {
        $prevEnd = [int]([math]::Min($prevStart + 99, $EndToken))
        $prevSlug = "{0:D4}-{1:D4}" -f $prevStart, $prevEnd
        $prevNext.Add("      <a href=""gallery-$prevSlug.html"">Previous tokens</a>")
    }
    if ($nextStart -le $EndToken) {
        $nextEnd = [int]([math]::Min($nextStart + 99, $EndToken))
        $nextSlug = "{0:D4}-{1:D4}" -f $nextStart, $nextEnd
        $prevNext.Add("      <a href=""gallery-$nextSlug.html"">Next tokens</a>")
    }

    $navHtml = ($prevNext -join "`n")
    $figureHtml = ($figures -join "`n")

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Light Cult Crypto Club - Tokens $rangeSlug</title>
  <link rel="stylesheet" href="css/styles.css" />
</head>

<body>
  <main class="wrap">

$(Build-TopNav -Prefix "")

    <header>
      <h1 class="page-title">Light Cult Crypto Club</h1>
      <p class="intro">Token-number archive: Light Cult Crypto Club #$("{0:D4}" -f $RangeStart)-#$("{0:D4}" -f $RangeEnd).</p>
    </header>

    <div class="page-nav" aria-label="Gallery page navigation">
$navHtml
    </div>

    <section class="image-grid" aria-label="Light Cult Crypto Club tokens $rangeSlug">
$figureHtml
    </section>

  </main>
</body>
</html>
"@
}

function Build-ItemHtml {
    param(
        [object]$Record,
        [object[]]$Records
    )

    $token = [int]$Record.TokenNumber
    $tokenPadded = "{0:D4}" -f $token
    $rangeSlug = Get-RangeSlug -Token $token
    $galleryHref = "../gallery-$rangeSlug.html"
    $imagePath = "../$($Record.ImagePath)"
    $title = "Light Cult Crypto Club #$tokenPadded"
    $alt = HtmlEncode "$title, $($Record.WorkID)"
    $openSeaLink = $Record.OpenSeaItemURL

    $prevNext = New-Object System.Collections.Generic.List[string]
    $prev = $Records | Where-Object { $_.Status -eq "OK" -and $_.TokenNumber -lt $token } | Sort-Object TokenNumber -Descending | Select-Object -First 1
    $next = $Records | Where-Object { $_.Status -eq "OK" -and $_.TokenNumber -gt $token } | Sort-Object TokenNumber | Select-Object -First 1

    $prevNext.Add("      <a href=""$galleryHref"">Back to tokens $rangeSlug</a>")
    if ($prev) {
        $prevNext.Add("      <a href=""../$($prev.ItemPage)"">Previous token</a>")
    }
    if ($next) {
        $prevNext.Add("      <a href=""../$($next.ItemPage)"">Next token</a>")
    }

    $pageNav = ($prevNext -join "`n")

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>$title - Ron English Catalogue Raisonné</title>
  <link rel="stylesheet" href="../css/styles.css" />
</head>

<body>
  <main class="wrap">

$(Build-TopNav -Prefix "../" -CurrentRangeHref $galleryHref)

    <header>
      <h1 class="page-title">$title</h1>
      <p class="intro">Light Cult Crypto Club token #$tokenPadded · $($Record.WorkID)</p>
    </header>

    <div class="page-nav" aria-label="Item page navigation">
$pageNav
    </div>

    <section class="info-card" aria-label="Catalogue metadata">
      <figure class="image-card">
        <img src="$imagePath" alt="$alt" loading="eager">
      </figure>

      <ul class="work-meta" data-work-id="$($Record.WorkID)">
        <li><strong>Title:</strong> $title</li>
        <li><strong>Project:</strong> Light Cult Crypto Club</li>
        <li><strong>Token number:</strong> #$tokenPadded</li>
        <li><strong>Work ID:</strong> $($Record.WorkID)</li>
        <li><strong>Image file:</strong> <code>$($Record.ImagePath)</code></li>
        <li><strong>OpenSea record:</strong> <a href="$openSeaLink">$openSeaLink</a></li>
      </ul>
    </section>

  </main>
</body>
</html>
"@
}

Write-Host "Scanning image files..."
$imageExts = @(".jpg", ".jpeg", ".png", ".webp", ".gif")
$imageFiles = Get-ChildItem -Path $ImagesRoot -Recurse -File | Where-Object {
    $imageExts -contains $_.Extension.ToLowerInvariant()
}

Write-Host "Image files found: $($imageFiles.Count)"

$candidatesByToken = @{}

foreach ($file in $imageFiles) {
    $token = Get-TokenFromFilename -BaseName $file.BaseName
    if ($null -eq $token) { continue }

    if ($token -lt $StartToken -or $token -gt $EndToken) { continue }

    if (!$candidatesByToken.ContainsKey($token)) {
        $candidatesByToken[$token] = New-Object System.Collections.Generic.List[object]
    }

    $candidatesByToken[$token].Add([pscustomobject]@{
        TokenNumber = $token
        File = $file
        Score = Get-FileScore -File $file -Token $token
    })
}

$records = New-Object System.Collections.Generic.List[object]
$duplicateRows = New-Object System.Collections.Generic.List[object]

for ($token = $StartToken; $token -le $EndToken; $token++) {
    $tokenPadded = "{0:D4}" -f $token
    $workId = Get-WorkId -Token $token
    $rangeSlug = Get-RangeSlug -Token $token
    $openSeaUrl = "https://opensea.io/item/ethereum/$OpenSeaContract/$token"
    $itemPage = "items/lccc-$tokenPadded.html"
    $publicPageUrl = "$PublicBaseUrl/items/lccc-$tokenPadded.html"

    if ($candidatesByToken.ContainsKey($token)) {
        $candidates = @($candidatesByToken[$token] | Sort-Object Score -Descending)
        $chosen = $candidates | Select-Object -First 1
        $relPath = To-SitePath -FullPath $chosen.File.FullName

        if ($candidates.Count -gt 1) {
            foreach ($c in $candidates) {
                $duplicateRows.Add([pscustomobject]@{
                    TokenNumber = $token
                    WorkID = $workId
                    Score = $c.Score
                    Path = To-SitePath -FullPath $c.File.FullName
                    Chosen = ($c.File.FullName -eq $chosen.File.FullName)
                })
            }
        }

        $expectedFolder = $rangeSlug
        $actualFolder = $chosen.File.Directory.Name
        $folderNote = if ($actualFolder -eq $expectedFolder) { "" } else { "Image is in folder '$actualFolder' but expected '$expectedFolder'" }

        $records.Add([pscustomobject]@{
            TokenNumber = $token
            TokenPadded = $tokenPadded
            Title = "Light Cult Crypto Club #$tokenPadded"
            WorkID = $workId
            Range = $rangeSlug
            ImagePath = $relPath
            ImageFileName = $chosen.File.Name
            ItemPage = $itemPage
            PublicPageURL = $publicPageUrl
            OpenSeaItemURL = $openSeaUrl
            Status = "OK"
            Note = $folderNote
        })
    }
    else {
        $records.Add([pscustomobject]@{
            TokenNumber = $token
            TokenPadded = $tokenPadded
            Title = "Light Cult Crypto Club #$tokenPadded"
            WorkID = $workId
            Range = $rangeSlug
            ImagePath = ""
            ImageFileName = ""
            ItemPage = $itemPage
            PublicPageURL = $publicPageUrl
            OpenSeaItemURL = $openSeaUrl
            Status = "MISSING IMAGE"
            Note = "No local image file found for token #$tokenPadded"
        })
    }
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$auditCsv = Join-Path $GeneratedRoot "lccc-0001-1000-rebuild-audit-$stamp.csv"
$dupesCsv = Join-Path $GeneratedRoot "lccc-0001-1000-duplicates-$stamp.csv"

$records | Export-Csv -Path $auditCsv -NoTypeInformation -Encoding UTF8
if ($duplicateRows.Count -gt 0) {
    $duplicateRows | Export-Csv -Path $dupesCsv -NoTypeInformation -Encoding UTF8
}

$okCount = @($records | Where-Object { $_.Status -eq "OK" }).Count
$missingCount = @($records | Where-Object { $_.Status -ne "OK" }).Count
$folderNoteCount = @($records | Where-Object { $_.Note -like "Image is in folder*" }).Count

Write-Host ""
Write-Host "Audit complete."
Write-Host "OK images:       $okCount"
Write-Host "Missing images:  $missingCount"
Write-Host "Folder warnings: $folderNoteCount"
Write-Host "Audit CSV:"
Write-Host "  $auditCsv"

if ($duplicateRows.Count -gt 0) {
    Write-Host "Duplicate candidate CSV:"
    Write-Host "  $dupesCsv"
}

if (!$Apply) {
    Write-Host ""
    Write-Host "No pages were changed."
    Write-Host "Review the audit CSV. If it looks right, run again with -Apply:"
    Write-Host '  & "$env:USERPROFILE\Downloads\rebuild-lccc-0001-1000-token-pages-v2.ps1" -Apply'
    Write-Host ""
    return
}

$backupRoot = Backup-CurrentSite

Write-Host ""
Write-Host "Writing rebuilt HTML pages..."

$itemsRoot = Join-Path $RepoRoot "items"
New-Item -ItemType Directory -Force -Path $itemsRoot | Out-Null

# Manifest for spreadsheet import later.
$manifestPath = Join-Path $DataRoot "light-cult-0001-1000-manifest.csv"
$records | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8

# Root index.
$indexHtml = Build-IndexHtml -Records $records
Write-Utf8File -Path (Join-Path $RepoRoot "index.html") -Content $indexHtml

# Gallery pages.
for ($rangeStart = $StartToken; $rangeStart -le $EndToken; $rangeStart += 100) {
    $rangeEnd = [int]([math]::Min($rangeStart + 99, $EndToken))
    $rangeSlug = "{0:D4}-{1:D4}" -f $rangeStart, $rangeEnd
    $galleryHtml = Build-GalleryHtml -RangeStart $rangeStart -RangeEnd $rangeEnd -Records $records
    Write-Utf8File -Path (Join-Path $RepoRoot "gallery-$rangeSlug.html") -Content $galleryHtml
}

# Individual item pages for found images only.
$itemPageCount = 0
foreach ($r in ($records | Where-Object { $_.Status -eq "OK" } | Sort-Object TokenNumber)) {
    $itemHtml = Build-ItemHtml -Record $r -Records $records
    Write-Utf8File -Path (Join-Path $RepoRoot $r.ItemPage) -Content $itemHtml
    $itemPageCount++
}

Write-Host ""
Write-Host "============================================================"
Write-Host "DONE"
Write-Host "============================================================"
Write-Host "Backup folder:"
Write-Host "  $backupRoot"
Write-Host "Manifest:"
Write-Host "  $manifestPath"
Write-Host "Audit CSV:"
Write-Host "  $auditCsv"
Write-Host "Gallery pages rebuilt: 10"
Write-Host "Individual item pages written: $itemPageCount"
Write-Host "Missing images: $missingCount"
Write-Host ""
Write-Host "Open locally:"
Write-Host "  file:///$($RepoRoot -replace '\\','/')/index.html"
Write-Host ""
