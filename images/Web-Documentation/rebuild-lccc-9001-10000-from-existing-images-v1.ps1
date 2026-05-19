param(
    [string]$RepoRoot = "D:\github\ron-english-light-cult-9001-10000",
    [int]$StartToken = 9001,
    [int]$EndToken = 10000,
    [int]$StartWorkNumber = 10811,
    [string]$PublicBaseUrl = "https://ronenglish-archive.github.io/ron-english-light-cult-9001-10000",
    [string]$NftsHubUrl = "../ron-english-nfts/index.html",
    [string]$CatalogueUrl = "https://ronenglish-archive.github.io/ron-english-catalogue-raisonne/",
    [string]$OpenSeaContract = "0xbe85fbd182af91290be7293438ae67549638189f",
    [switch]$Apply,
    [switch]$AllowMissing,
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"

$RangeLabel = "{0:D4}-{1:D4}" -f $StartToken, $EndToken
$RepoName = Split-Path $RepoRoot -Leaf

Write-Host ""
Write-Host "============================================================"
Write-Host "Rebuild LCCC 9001-10000 from existing repo images v1"
Write-Host "============================================================"
Write-Host "Repo root:        $RepoRoot"
Write-Host "Token range:      $RangeLabel"
Write-Host "Starting Work ID: RE-$("{0:D6}" -f $StartWorkNumber)"
if ($Apply) {
    Write-Host "MODE:             APPLY"
} else {
    Write-Host "MODE:             AUDIT ONLY"
}
Write-Host ""

if (!(Test-Path $RepoRoot)) {
    throw "Repo root not found: $RepoRoot"
}

$ImagesRoot = Join-Path $RepoRoot "images"
$DataRoot = Join-Path $RepoRoot "data"
$GeneratedRoot = Join-Path $DataRoot "generated"
$ItemsRoot = Join-Path $RepoRoot "items"
$CssRoot = Join-Path $RepoRoot "css"

if (!(Test-Path $ImagesRoot)) {
    throw "Images folder not found: $ImagesRoot"
}

New-Item -ItemType Directory -Force -Path $DataRoot | Out-Null
New-Item -ItemType Directory -Force -Path $GeneratedRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ItemsRoot | Out-Null
New-Item -ItemType Directory -Force -Path $CssRoot | Out-Null

$imageExts = @(".jpg", ".jpeg", ".png", ".webp", ".gif", ".avif")

function Pad4 {
    param([int]$Number)
    return "{0:D4}" -f $Number
}

function Pad6 {
    param([int]$Number)
    return "{0:D6}" -f $Number
}

function HtmlEncode {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function To-RepoRelativePath {
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
    return "$(Pad4 $rs)-$(Pad4 $re)"
}

function Get-WorkId {
    param([int]$Token)
    $workNumber = $StartWorkNumber + ($Token - $StartToken)
    return "RE-$(Pad6 $workNumber)"
}

function Get-TokenFromWorkNumber {
    param([int]$WorkNumber)
    return $StartToken + ($WorkNumber - $StartWorkNumber)
}

function Get-TokenFromFilename {
    param([string]$BaseName)

    # New token-numbered filenames:
    # LCCC-9001, LCCC_9001, LCCC 9001
    $m = [regex]::Match($BaseName, "(?i)^LCCC[-_\s]*0*(\d{1,5})$")
    if ($m.Success) {
        return [int]$m.Groups[1].Value
    }

    # Old Work-ID filenames:
    # RE-010811 = token 9001 when StartWorkNumber is 10811
    $m = [regex]::Match($BaseName, "(?i)^RE[-_\s]*(\d{6})$")
    if ($m.Success) {
        $workNum = [int]$m.Groups[1].Value
        return Get-TokenFromWorkNumber -WorkNumber $workNum
    }

    # Conservative fallback: standalone token number 9001-10000 in filename.
    $m = [regex]::Match($BaseName, "(?<!\d)(9\d{3}|10000)(?!\d)")
    if ($m.Success) {
        return [int]$m.Groups[1].Value
    }

    return $null
}

function Backup-Pages {
    if ($NoBackup) {
        return ""
    }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupRoot = "D:\github-backups\$RepoName\rebuild-from-existing-images-v1-$RangeLabel-$stamp"
    New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

    Get-ChildItem -Path $RepoRoot -File -Filter "*.html" -ErrorAction SilentlyContinue |
        Copy-Item -Destination $backupRoot -Force

    if (Test-Path $ItemsRoot) {
        Copy-Item -Path $ItemsRoot -Destination (Join-Path $backupRoot "items") -Recurse -Force
    }

    if (Test-Path $DataRoot) {
        New-Item -ItemType Directory -Force -Path (Join-Path $backupRoot "data") | Out-Null
        Get-ChildItem -Path $DataRoot -File -Filter "*.csv" -ErrorAction SilentlyContinue |
            Copy-Item -Destination (Join-Path $backupRoot "data") -Force
    }

    return $backupRoot
}

function Build-Css {
    return @"
:root {
  --bg: #ffffff;
  --paper: #ffffff;
  --ink: #151515;
  --muted: #5f6470;
  --line: rgba(15, 23, 42, 0.14);
  --shadow: 0 14px 35px rgba(15, 23, 42, 0.09);
  --shadow-hover: 0 20px 48px rgba(15, 23, 42, 0.14);
  --radius: 22px;
}

* { box-sizing: border-box; }

body {
  margin: 0;
  background: var(--bg);
  color: var(--ink);
  font-family: Georgia, "Times New Roman", serif;
  line-height: 1.55;
}

a { color: inherit; }

.wrap {
  width: min(1280px, calc(100% - 32px));
  margin: 0 auto;
  padding: 32px 0 56px;
}

.nav-pills,
.page-nav {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
  margin: 0 auto 18px;
  font-family: Arial, Helvetica, sans-serif;
}

.nav-pill,
.page-nav a {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 8px 16px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #ffffff;
  text-decoration: none;
  box-shadow: 0 8px 20px rgba(15, 23, 42, 0.05);
}

.nav-pill:hover,
.page-nav a:hover,
.nav-pill.is-active {
  background: #111111;
  color: #ffffff;
  border-color: #111111;
}

header {
  text-align: center;
  margin: 28px auto 28px;
}

.page-title {
  margin: 0 0 14px;
  font-family: Arial, Helvetica, sans-serif;
  font-weight: 900;
  font-size: clamp(2.5rem, 6vw, 5rem);
  line-height: 0.95;
  letter-spacing: -0.055em;
}

.intro {
  width: min(820px, 100%);
  margin: 0 auto;
  color: var(--muted);
  font-size: 1.08rem;
}

.section-title {
  margin: 34px 0 18px;
  text-align: center;
  font-family: Arial, Helvetica, sans-serif;
  font-size: clamp(1.6rem, 3vw, 2.35rem);
  line-height: 1.05;
  letter-spacing: -0.035em;
}

.gallery-stats {
  width: min(760px, 100%);
  margin: 0 auto 26px;
  text-align: center;
  color: var(--muted);
}

.token-grid,
.image-grid,
.range-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(165px, 1fr));
  gap: 18px;
  align-items: stretch;
}

.token-card,
.image-card,
.range-card {
  display: flex;
  flex-direction: column;
  overflow: hidden;
  min-height: 100%;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: var(--paper);
  text-decoration: none;
  box-shadow: var(--shadow);
  transition: transform 160ms ease, box-shadow 160ms ease;
}

.token-card:hover,
.image-card:hover,
.range-card:hover {
  transform: translateY(-3px);
  box-shadow: var(--shadow-hover);
}

.token-card img,
.image-card img,
.range-card img {
  display: block;
  width: 100%;
  height: 190px;
  padding: 12px;
  object-fit: contain;
  background: #ffffff;
}

.token-card-body,
.work-id,
.range-card-body {
  padding: 13px 14px 15px;
  text-align: center;
  font-family: Arial, Helvetica, sans-serif;
  border-top: 1px solid rgba(15, 23, 42, 0.08);
}

.token-title,
.range-card-body h2,
.range-card-body h3 {
  display: block;
  margin: 0 0 4px;
  font-weight: 800;
  font-size: 0.98rem;
  line-height: 1.2;
}

.token-workid,
.work-id span,
.range-card-body p {
  display: block;
  margin: 0;
  color: var(--muted);
  font-size: 0.9rem;
}

.work-layout {
  width: min(1040px, calc(100% - 32px));
  margin: 0 auto 54px;
}

.work-image-card {
  width: min(820px, 100%);
  margin: 24px auto 28px;
  padding: 18px;
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
}

.work-hero {
  display: block;
  max-width: 100%;
  max-height: 78vh;
  width: auto;
  height: auto;
  margin: 0 auto;
  object-fit: contain;
  border-radius: calc(var(--radius) - 8px);
}

.work-meta {
  width: min(760px, 100%);
  margin: 0 auto 30px;
  padding: 0;
  list-style: none;
  text-align: center;
  font-size: 1.06rem;
  background: transparent;
  border: 0;
  box-shadow: none;
}

.work-meta li + li { margin-top: 8px; }
.work-meta strong { font-family: Arial, Helvetica, sans-serif; }

footer {
  margin: 42px auto 0;
  text-align: center;
  color: var(--muted);
  font-family: Arial, Helvetica, sans-serif;
  font-size: 0.92rem;
}

code {
  font-family: "Courier New", monospace;
  font-size: 0.92em;
}

@media (max-width: 720px) {
  .wrap,
  .work-layout {
    width: min(100% - 22px, 1280px);
    padding-top: 22px;
  }

  .token-grid,
  .image-grid,
  .range-grid {
    grid-template-columns: repeat(auto-fit, minmax(135px, 1fr));
    gap: 14px;
  }

  .token-card img,
  .image-card img,
  .range-card img {
    height: 160px;
    padding: 10px;
  }

  .work-image-card { padding: 12px; }
}
"@
}

function Build-TopNav {
    param(
        [string]$Prefix = "",
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

    $ok = @($Records | Where-Object { $_.Status -eq "OK" } | Sort-Object TokenNumber)
    $missing = @($Records | Where-Object { $_.Status -ne "OK" })

    $cards = New-Object System.Collections.Generic.List[string]

    foreach ($r in $ok) {
        $tokenPadded = Pad4 ([int]$r.TokenNumber)
        $alt = HtmlEncode "Light Cult Crypto Club #$tokenPadded"
        $cards.Add(@"
      <a class="token-card" href="$($r.ItemPage)">
        <img src="$($r.ImagePath)" alt="$alt" loading="lazy">
        <span class="token-card-body">
          <span class="token-title">#$tokenPadded</span>
          <span class="token-workid">$($r.WorkID)</span>
        </span>
      </a>
"@)
    }

    $cardsHtml = $cards -join "`n"
    $okCount = $ok.Count
    $missingCount = $missing.Count

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Light Cult Crypto Club - Tokens $RangeLabel</title>
  <link rel="stylesheet" href="css/styles.css" />
</head>

<body>
  <main class="wrap">

$(Build-TopNav -Prefix "")

    <header>
      <h1 class="page-title">Light Cult Crypto Club</h1>
      <p class="intro">Token-number gallery for Light Cult Crypto Club #$RangeLabel.</p>
    </header>

    <section aria-label="Token gallery">
      <h2 class="section-title">Tokens $RangeLabel</h2>
      <p class="gallery-stats">$okCount local images are currently documented in this repository. $missingCount token records are marked as missing in the manifest.</p>

      <div class="token-grid">
$cardsHtml
      </div>
    </section>

    <footer>
      Local manifest: <code>data/light-cult-$RangeLabel-manifest.csv</code>
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

    $rangeSlug = "$(Pad4 $RangeStart)-$(Pad4 $RangeEnd)"
    $rangeRecords = @($Records | Where-Object {
        $_.Status -eq "OK" -and $_.TokenNumber -ge $RangeStart -and $_.TokenNumber -le $RangeEnd
    } | Sort-Object TokenNumber)

    $figures = New-Object System.Collections.Generic.List[string]

    foreach ($r in $rangeRecords) {
        $tokenPadded = Pad4 ([int]$r.TokenNumber)
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

    $navParts = New-Object System.Collections.Generic.List[string]
    $navParts.Add('      <a href="index.html">All tokens in this repo</a>')

    $prevStart = $RangeStart - 100
    if ($prevStart -ge $StartToken) {
        $prevEnd = [int]([math]::Min($prevStart + 99, $EndToken))
        $prevSlug = "$(Pad4 $prevStart)-$(Pad4 $prevEnd)"
        $navParts.Add("      <a href=""gallery-$prevSlug.html"">Previous tokens</a>")
    }

    $nextStart = $RangeStart + 100
    if ($nextStart -le $EndToken) {
        $nextEnd = [int]([math]::Min($nextStart + 99, $EndToken))
        $nextSlug = "$(Pad4 $nextStart)-$(Pad4 $nextEnd)"
        $navParts.Add("      <a href=""gallery-$nextSlug.html"">Next tokens</a>")
    }

    $navHtml = ($navParts -join "`n")
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
      <p class="intro">Token-number gallery: Light Cult Crypto Club #$(Pad4 $RangeStart)-#$(Pad4 $RangeEnd).</p>
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
    $tokenPadded = Pad4 $token
    $rangeSlug = Get-RangeSlug -Token $token
    $title = "Light Cult Crypto Club #$tokenPadded"
    $imageSrc = "../$($Record.ImagePath)"
    $alt = HtmlEncode $title

    $prev = $Records | Where-Object { $_.Status -eq "OK" -and $_.TokenNumber -lt $token } | Sort-Object TokenNumber -Descending | Select-Object -First 1
    $next = $Records | Where-Object { $_.Status -eq "OK" -and $_.TokenNumber -gt $token } | Sort-Object TokenNumber | Select-Object -First 1

    $navParts = New-Object System.Collections.Generic.List[string]
    $navParts.Add("      <a href=""../gallery-$rangeSlug.html"">Back to tokens $rangeSlug</a>")

    if ($prev) {
        $navParts.Add("      <a href=""../$($prev.ItemPage)"">Previous token</a>")
    }

    if ($next) {
        $navParts.Add("      <a href=""../$($next.ItemPage)"">Next token</a>")
    }

    $tokenNav = $navParts -join "`n"

    return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>$title - Ron English Catalogue Raisonné</title>
  <meta name="work-id" content="$($Record.WorkID)" />
  <link rel="stylesheet" href="../css/styles.css" />
</head>

<body>
  <main class="work-layout">

$(Build-TopNav -Prefix "../" -CurrentRangeHref "../gallery-$rangeSlug.html")

    <header>
      <h1 class="page-title">$title</h1>
    </header>

    <figure class="work-image-card">
      <img src="$imageSrc" alt="$alt" class="work-hero" loading="lazy">
    </figure>

    <ul class="work-meta" data-work-id="$($Record.WorkID)">
      <li><strong>Title:</strong> $title</li>
      <li><strong>Project / series:</strong> Light Cult Crypto Club</li>
      <li><strong>Medium:</strong> NFT / digital collectible</li>
      <li><strong>Token number:</strong> #$tokenPadded</li>
      <li><strong>Work ID:</strong> $($Record.WorkID)</li>
      <li><strong>OpenSea record:</strong> <a href="$($Record.OpenSeaItemURL)">$($Record.OpenSeaItemURL)</a></li>
    </ul>

    <div class="page-nav" aria-label="Token navigation">
$tokenNav
    </div>

  </main>
</body>
</html>
"@
}

Write-Host "Scanning existing repo images..."
$imageFiles = Get-ChildItem -Path $ImagesRoot -Recurse -File | Where-Object {
    $imageExts -contains $_.Extension.ToLowerInvariant()
}

Write-Host "Image files found in repo: $($imageFiles.Count)"

$bestImageByToken = @{}
$duplicateRows = New-Object System.Collections.Generic.List[object]
$unmatchedRows = New-Object System.Collections.Generic.List[object]

foreach ($file in $imageFiles) {
    $token = Get-TokenFromFilename -BaseName $file.BaseName

    if ($null -eq $token -or $token -lt $StartToken -or $token -gt $EndToken) {
        $unmatchedRows.Add([pscustomobject]@{
            FileName = $file.Name
            FullPath = $file.FullName
            Reason = "Could not map filename to token $RangeLabel"
        })
        continue
    }

    if (!$bestImageByToken.ContainsKey($token)) {
        $bestImageByToken[$token] = $file
    }
    else {
        $current = $bestImageByToken[$token]
        if ($file.Length -gt $current.Length) {
            $bestImageByToken[$token] = $file
        }

        $duplicateRows.Add([pscustomobject]@{
            TokenNumber = $token
            CandidateA = $current.FullName
            CandidateABytes = $current.Length
            CandidateB = $file.FullName
            CandidateBBytes = $file.Length
            Chosen = $bestImageByToken[$token].FullName
        })
    }
}

$records = New-Object System.Collections.Generic.List[object]

for ($token = $StartToken; $token -le $EndToken; $token++) {
    $tokenPadded = Pad4 $token
    $workId = Get-WorkId -Token $token
    $rangeSlug = Get-RangeSlug -Token $token
    $openSeaUrl = "https://opensea.io/item/ethereum/$OpenSeaContract/$token"
    $itemPage = "items/lccc-$tokenPadded.html"
    $publicPageUrl = "$PublicBaseUrl/items/lccc-$tokenPadded.html"

    if ($bestImageByToken.ContainsKey($token)) {
        $img = $bestImageByToken[$token]
        $relPath = To-RepoRelativePath -FullPath $img.FullName

        $records.Add([pscustomobject]@{
            TokenNumber = $token
            TokenPadded = $tokenPadded
            Title = "Light Cult Crypto Club #$tokenPadded"
            WorkID = $workId
            Range = $rangeSlug
            SourceImagePath = $img.FullName
            SourceBytes = $img.Length
            ImagePath = $relPath
            ImageFileName = $img.Name
            DestinationImagePath = $img.FullName
            ItemPage = $itemPage
            PublicPageURL = $publicPageUrl
            OpenSeaItemURL = $openSeaUrl
            Status = "OK"
            Note = ""
        })
    }
    else {
        $records.Add([pscustomobject]@{
            TokenNumber = $token
            TokenPadded = $tokenPadded
            Title = "Light Cult Crypto Club #$tokenPadded"
            WorkID = $workId
            Range = $rangeSlug
            SourceImagePath = ""
            SourceBytes = ""
            ImagePath = ""
            ImageFileName = ""
            DestinationImagePath = ""
            ItemPage = $itemPage
            PublicPageURL = $publicPageUrl
            OpenSeaItemURL = $openSeaUrl
            Status = "MISSING SOURCE IMAGE"
            Note = "No existing repo image found for token #$tokenPadded"
        })
    }
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$auditCsv = Join-Path $GeneratedRoot "lccc-$RangeLabel-existing-image-audit-$stamp.csv"
$dupesCsv = Join-Path $GeneratedRoot "lccc-$RangeLabel-existing-image-duplicates-$stamp.csv"
$unmatchedCsv = Join-Path $GeneratedRoot "lccc-$RangeLabel-existing-image-unmatched-$stamp.csv"

$records | Export-Csv -Path $auditCsv -NoTypeInformation -Encoding UTF8
$duplicateRows | Export-Csv -Path $dupesCsv -NoTypeInformation -Encoding UTF8
$unmatchedRows | Export-Csv -Path $unmatchedCsv -NoTypeInformation -Encoding UTF8

$okCount = @($records | Where-Object { $_.Status -eq "OK" }).Count
$missingCount = @($records | Where-Object { $_.Status -ne "OK" }).Count

Write-Host ""
Write-Host "Audit complete."
Write-Host "OK existing images:      $okCount"
Write-Host "Missing token images:    $missingCount"
Write-Host "Unmatched image files:   $($unmatchedRows.Count)"
Write-Host "Duplicate token files:   $($duplicateRows.Count)"
Write-Host "Audit CSV:"
Write-Host "  $auditCsv"
Write-Host ""

if (!$Apply) {
    Write-Host "No files were changed."
    Write-Host "If this looks right, apply with:"
    Write-Host '  & "$env:USERPROFILE\Downloads\rebuild-lccc-9001-10000-from-existing-images-v1.ps1" -Apply -AllowMissing'
    Write-Host ""
    return
}

if ($missingCount -gt 0 -and !$AllowMissing) {
    throw "There are $missingCount missing token images. Re-run with -AllowMissing if this is expected."
}

$backupRoot = Backup-Pages
if ($backupRoot -ne "") {
    Write-Host "Backup created:"
    Write-Host "  $backupRoot"
    Write-Host ""
}

Write-Host "Writing manifest and pages..."

$manifestPath = Join-Path $DataRoot "light-cult-$RangeLabel-manifest.csv"
$records | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8

Write-Utf8File -Path (Join-Path $CssRoot "styles.css") -Content (Build-Css)
Write-Utf8File -Path (Join-Path $RepoRoot "index.html") -Content (Build-IndexHtml -Records $records)

# Gallery pages.
$galleryCount = 0
for ($rangeStart = $StartToken; $rangeStart -le $EndToken; $rangeStart += 100) {
    $rangeEnd = [int]([math]::Min($rangeStart + 99, $EndToken))
    $rangeSlug = "$(Pad4 $rangeStart)-$(Pad4 $rangeEnd)"
    $galleryHtml = Build-GalleryHtml -RangeStart $rangeStart -RangeEnd $rangeEnd -Records $records
    Write-Utf8File -Path (Join-Path $RepoRoot "gallery-$rangeSlug.html") -Content $galleryHtml
    $galleryCount++
}

# Item pages.
Get-ChildItem -Path $ItemsRoot -Filter "*.html" -File -ErrorAction SilentlyContinue | Remove-Item -Force

$itemCount = 0
foreach ($r in ($records | Where-Object { $_.Status -eq "OK" } | Sort-Object TokenNumber)) {
    $itemHtml = Build-ItemHtml -Record $r -Records $records
    Write-Utf8File -Path (Join-Path $RepoRoot $r.ItemPage) -Content $itemHtml
    $itemCount++
}

Write-Host ""
Write-Host "============================================================"
Write-Host "DONE"
Write-Host "============================================================"
Write-Host "Manifest:"
Write-Host "  $manifestPath"
Write-Host "Gallery pages rebuilt:         $galleryCount"
Write-Host "Individual item pages written: $itemCount"
Write-Host "Missing token images:          $missingCount"
Write-Host ""
Write-Host "Open locally:"
Write-Host "  file:///$($RepoRoot -replace '\\','/')/index.html"
Write-Host ""
