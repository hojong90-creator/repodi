<#
.SYNOPSIS
  같은 업체로 이미 생성해둔 여러 기간의 광고 보고서(HTML) 안에 내장된 데이터를 모아
  월별(또는 기간별) 추이 보고서 하나로 재구성한다. 원본 CSV/XLSX를 다시 읽지 않고,
  각 보고서 HTML에 내장된 <script id="report-data"> JSON만 사용한다.
.PARAMETER ReportsDir
  기존에 생성된 보고서(.html)들이 들어있는 폴더 경로 ('생성된 보고서' 폴더).
.PARAMETER InputDir
  '보고서 생성용 데이터' 폴더(레포디의 '파일 넣기'). 여기에 보고서(.html)가 있으면 ReportsDir 대신 이 파일들을 모두 쓴다
  (다른 PC에서 받은 보고서, 이름을 바꾼 보고서도 쓸 수 있게. 업체명으로 거르지 않는다).
.PARAMETER ClientName
  묶어서 볼 업체명. 'YYMM_{업체명}_광고_보고서...' (또는 예전 '{업체명}_광고_보고서_...') 형식의 파일만 대상이 된다.
#>
param(
    [Parameter(Mandatory = $true)][string]$ReportsDir,
    [string]$InputDir = "",
    [Parameter(Mandatory = $true)][string]$ClientName,
    [Parameter(Mandatory = $true)][string]$TemplatePath,
    [Parameter(Mandatory = $true)][string]$HtmlOutputPath,
    [string]$AuthorInfo = "",
    [string]$LogoPath = ""
)

$ErrorActionPreference = "Stop"

try {

function Metrics-Combine {
    # 여러 overview 객체(노출수/클릭수/총비용/전환수/전환매출액)를 하나로 합산하고 CTR/CPC/ROAS를 다시 계산한다.
    param($Items)
    $imp = 0.0; $clk = 0.0; $cost = 0.0; $conv = 0.0; $rev = 0.0
    foreach ($it in $Items) {
        if (-not $it) { continue }
        $imp += [double]$it.impressions; $clk += [double]$it.clicks; $cost += [double]$it.cost
        $conv += [double]$it.conversions; $rev += [double]$it.revenue
    }
    $ctr = if ($imp -gt 0) { [math]::Round(($clk / $imp) * 100, 2) } else { 0 }
    $cpc = if ($clk -gt 0) { [math]::Round($cost / $clk) } else { 0 }
    $roas = if ($cost -gt 0) { [math]::Round(($rev / $cost) * 100, 2) } else { 0 }
    [PSCustomObject]@{
        impressions = [int64]$imp; clicks = [int64]$clk; ctr = $ctr
        cost        = [int64]$cost; cpc = $cpc; conversions = [math]::Round($conv, 2)   # 구글 전환은 소수로 나온다
        revenue     = [int64]$rev; roas = $roas
    }
}

$safeClient = ($ClientName -replace '[\\/:\*\?"<>\|]', '').Trim()
if (-not $safeClient) { throw "업체명을 입력해주세요." }

# 대상 파일: 'YYMM_{업체명}_광고_보고서(_N주차).html' 또는 예전 형식 '{업체명}_광고_보고서_....html'
$esc = [regex]::Escape($safeClient)
$namePattern = "^(\d{4}_${esc}_광고_보고서(_\d+주차)?|${esc}_광고_보고서_.+)\.html$"
$files = @()
if ($InputDir -and (Test-Path -LiteralPath $InputDir)) {
    $files = @(Get-ChildItem -LiteralPath $InputDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.html?$' })
}
if ($files.Count -eq 0) {
    $files = @(Get-ChildItem -LiteralPath $ReportsDir -Filter "*.html" -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $namePattern })
}
if ($files.Count -eq 0) {
    throw "'파일 넣기'에 보고서(HTML)가 없고, '$ReportsDir' 폴더에서도 '$safeClient' 업체의 보고서를 찾지 못했습니다. (파일명이 'YYMM_${safeClient}_광고_보고서' 형식이어야 인식됩니다)"
}

$months = @()
foreach ($f in $files) {
    $html = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    if ($html -notmatch '(?s)<script id="report-data" type="application/json">(.*?)</script>') {
        Write-Warning "$($f.Name) : 보고서 데이터를 찾을 수 없어 건너뜁니다."
        continue
    }
    $data = $matches[1] | ConvertFrom-Json

    if (-not $data.period -or $data.period -notmatch '(\d{4})\.(\d{2})\.(\d{2})\.~(\d{4})\.(\d{2})\.(\d{2})\.') {
        Write-Warning "$($f.Name) : 기간 정보를 읽을 수 없어 건너뜁니다."
        continue
    }
    $sy = [int]$matches[1]; $sm = [int]$matches[2]; $sd = [int]$matches[3]
    $ey = [int]$matches[4]; $em = [int]$matches[5]; $ed = [int]$matches[6]
    $startDate = Get-Date -Year $sy -Month $sm -Day $sd
    $endDate = Get-Date -Year $ey -Month $em -Day $ed
    $dayCount = [int]($endDate - $startDate).TotalDays + 1
    # 한 달보다 긴 보고서(예: 1년치 검색어 보고서)를 '파일 넣기'로 함께 넣어도 한 달로 잘못 세지 않게 뺀다
    if ($dayCount -gt 31) {
        Write-Warning "$($f.Name) : 기간이 ${dayCount}일이라 월별 추이에서 뺍니다."
        continue
    }
    $periodLabel = if ($dayCount -ge 25) { "${sy}년 ${sm}월" } else { "${sy}년 ${sm}월 $([math]::Ceiling($sd / 7.0))주차" }

    $combineItems = @($data.overview)
    if ($data.gmarket -and $data.gmarket.available -and $data.gmarket.overview) { $combineItems += $data.gmarket.overview }
    if ($data.auction -and $data.auction.available -and $data.auction.overview) { $combineItems += $data.auction.overview }
    if ($data.elevenst -and $data.elevenst.available -and $data.elevenst.overview) { $combineItems += $data.elevenst.overview }
    if ($data.ohouse -and $data.ohouse.available -and $data.ohouse.overview) { $combineItems += $data.ohouse.overview }
    # 구글, 메타도 합산한다. 달러로 집행한 매체는 환율로 바꾸지 않고 원화 합계와 따로 적는다.
    $usdCost = 0.0; $usdConv = 0.0
    foreach ($ch in @($data.channels)) {
        if ($ch.key -ne 'google' -and $ch.key -ne 'meta') { continue }
        if ($ch.currency -eq 'USD') { $usdCost += [double]$ch.overview.cost; $usdConv += [double]$ch.overview.conversions }
        else { $combineItems += $ch.overview }
    }
    $m = Metrics-Combine $combineItems

    $months += [PSCustomObject]@{
        periodLabel = $periodLabel
        period      = $data.period
        sortKey     = $startDate.ToString("yyyy-MM-dd")
        sourceFile  = $f.Name
        impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
        cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
        revenue     = $m.revenue; roas = $m.roas
        usdCost     = [math]::Round($usdCost, 2); usdConversions = [math]::Round($usdConv, 2)
    }
}

if ($months.Count -eq 0) { throw "유효한 보고서 데이터를 찾지 못했습니다." }

$months = @($months | Sort-Object -Property sortKey)
$total = Metrics-Combine $months
$total | Add-Member -NotePropertyName usdCost -NotePropertyValue ([math]::Round(($months | Measure-Object -Property usdCost -Sum).Sum, 2))
$total | Add-Member -NotePropertyName usdConversions -NotePropertyValue ([math]::Round(($months | Measure-Object -Property usdConversions -Sum).Sum, 2))

$result = [PSCustomObject]@{
    generatedAt = (Get-Date -Format "yyyy-MM-dd HH:mm")
    clientName  = $safeClient
    authorInfo  = $AuthorInfo
    months      = $months
    total       = $total
}

$reportTitle = "${safeClient} 월별 성과 추이"

$jsonForHtml = ($result | ConvertTo-Json -Depth 10 -Compress)
$jsonForHtml = $jsonForHtml.Replace('</', '<\/')

$logoDataUri = "__LOGO_DATA_URI__"
if (-not $LogoPath) {
    $projectRoot = Split-Path $PSScriptRoot -Parent
    $assetsDir = Join-Path $projectRoot "assets"
    if (Test-Path -LiteralPath $assetsDir) {
        $found = Get-ChildItem -LiteralPath $assetsDir -File | Where-Object { $_.Name -match '(?i)(logo|로고)' -and $_.Extension -match '(?i)\.(png|jpg|jpeg)$' } | Select-Object -First 1
        if ($found) { $LogoPath = $found.FullName }
    }
}
if ($LogoPath -and (Test-Path -LiteralPath $LogoPath)) {
    $ext = [System.IO.Path]::GetExtension($LogoPath).ToLower()
    $mime = switch ($ext) { ".png" { "image/png" }; ".jpg" { "image/jpeg" }; ".jpeg" { "image/jpeg" }; default { "image/png" } }
    $bytes = [System.IO.File]::ReadAllBytes($LogoPath)
    $b64 = [System.Convert]::ToBase64String($bytes)
    $logoDataUri = "data:$mime;base64,$b64"
}

$template = Get-Content -LiteralPath $TemplatePath -Raw -Encoding UTF8
$html = $template.Replace('__REPORT_TITLE__', $reportTitle).Replace('__REPORT_DATA_JSON__', $jsonForHtml).Replace('__LOGO_DATA_URI__', $logoDataUri)

$enc = New-Object System.Text.UTF8Encoding($true)
$htmlDir = Split-Path $HtmlOutputPath -Parent
if ($htmlDir -and -not (Test-Path -LiteralPath $htmlDir)) { New-Item -ItemType Directory -Force -Path $htmlDir | Out-Null }
[System.IO.File]::WriteAllText($HtmlOutputPath, $html, $enc)

Write-Output "OK -> $HtmlOutputPath ($($months.Count)개월)"

} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
