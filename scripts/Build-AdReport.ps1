<#
.SYNOPSIS
  네이버(지면별/검색어/시간대별/지역/전환유형/GFA) · 지마켓(상품별/키워드별) · 옥션(AI매출업/파워클릭) · 11번가(광고상품별 일별) · 구글 · 메타 · 오늘의집(캠페인/상품/검색어)
  리포트 CSV/XLSX를 읽어 매체별로 하나의 통합 요약 JSON + HTML로 만든다.
  파일 중 존재하는 것만 사용하며, 없는 매체·항목의 섹션은 보고서에서 자동으로 제외된다.
.PARAMETER InputDir
  CSV/XLSX 파일들이 들어있는 폴더 경로.
.PARAMETER OutputPath
  결과 JSON을 저장할 경로. TemplatePath를 주면 JSON은 만들지 않고, 이 경로의 폴더에 HTML 보고서만 만든다.
#>
param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [string]$TemplatePath,
    [string]$HtmlOutputPath,
    [string]$ClientName = "",
    [string]$AuthorInfo = "",
    [string]$LogoPath = "",
    [string]$VisibleMetrics = "impressions,clicks,ctr,cpc,cost,conversions,revenue,roas",
    # 지정하면, 같은 이름의 보고서가 이미 있을 때 덮어쓰지 않고 이 경로에 대신 저장한 뒤 종료 코드 3으로 끝낸다.
    # (덮어쓸지는 호출한 HTA가 사용자에게 물어서 결정. 원래 저장하려던 경로는 '$PendingHtmlPath.txt'에 UTF-16으로 적어둔다)
    [string]$PendingHtmlPath = ""
)

$ErrorActionPreference = "Stop"

function Get-NaverCsv {
    param([string]$Path)

    $raw = Get-Content -LiteralPath $Path -Encoding UTF8
    if ($raw.Count -lt 2) { throw "$Path : 데이터가 없습니다." }

    # 1행: 리포트 제목 + 기간. 2행: 실제 헤더. 3행부터 데이터.
    $titleLine = $raw[0]
    $period = $null
    if ($titleLine -match '(\d{4}\.\d{2}\.\d{2}\.)~(\d{4}\.\d{2}\.\d{2}\.)') {
        $period = "$($matches[1])~$($matches[2])"
    }

    # 문의, 신청을 전환으로 받는 광고주의 보고서에는 '구매완료 전환수' 대신 '총 전환수'만 있다.
    # 아래 집계가 컬럼 이름 하나만 보도록 여기서 맞춘다. (전환유형 보고서는 '총 전환수' 그대로 읽는다)
    if ($raw[1] -notmatch '구매완료 전환수|전환 유형') { $raw[1] = $raw[1] -replace '(^|,)"?총 전환수"?(?=,|$)', '$1구매완료 전환수' }

    $csvBody = $raw[1..($raw.Count - 1)] -join "`n"
    $rows = $csvBody | ConvertFrom-Csv

    [PSCustomObject]@{
        Period = $period
        Rows   = $rows
    }
}

function Get-NaverPeriod {
    # 네이버 보고서 1행 제목의 기간 ('...보고서(2026.10.01.~2026.10.31.)'). 큰 파일을 통째로 읽지 않으려고 1행만 본다.
    param([string]$Path)
    $sr = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try { $titleLine = $sr.ReadLine() } finally { $sr.Close() }
    if ($titleLine -match '(\d{4}\.\d{2}\.\d{2}\.)~(\d{4}\.\d{2}\.\d{2}\.)') { return "$($matches[1])~$($matches[2])" }
    return $null
}

function Test-GfaHeader([string]$Line) {
    # GFA 성과 보고서는 제목 행 없이 1행이 컬럼이다. '연령 및 성별'은 GFA에만 있는 컬럼이라 있으면 GFA다.
    # 캠페인 단위로만 받은 보고서(연령·성별, 기간 없이 캠페인 이름, 캠페인 ID, 캠페인 목적, 총비용 등. 2026-11-03)는 그 컬럼이 없어서
    # '캠페인 이름' + '총비용' + ('캠페인 목적' 또는 '캠페인 ID')로 알아본다.
    # 메타는 '지출 금액', GFA 상품 성과 보고서는 '네이버쇼핑 상품 ID'가 있어 겹치지 않는다.
    if ($Line -match '연령 및 성별') { return $true }
    return (($Line -match '캠페인 이름') -and ($Line -match '총비용') -and (($Line -match '캠페인 목적') -or ($Line -match '캠페인 ID')) -and
        ($Line -notmatch '지출 금액') -and ($Line -notmatch '네이버쇼핑 상품 ID'))
}

function Get-GfaCsv {
    # GFA 리포트는 제목/기간 행이 따로 없고 1행이 바로 헤더이며, 행마다 '기간' 컬럼(일자)을 갖는다.
    param([string]$Path)

    $raw = Get-Content -LiteralPath $Path -Encoding UTF8
    if ($raw.Count -lt 2) { throw "$Path : 데이터가 없습니다." }

    # 네이버 검색광고와 같은 이유로 '총 전환수'만 있으면 컬럼 이름을 맞춘다.
    if ($raw[0] -notmatch '구매완료 수') { $raw[0] = $raw[0] -replace '(^|,)"?총 전환수"?(?=,|$)', '$1구매완료 수' }

    $csvBody = $raw -join "`n"
    $rows = $csvBody | ConvertFrom-Csv

    $dates = @($rows | ForEach-Object { $_.'기간' } | Where-Object { $_ } | Sort-Object)
    $period = $null
    if ($dates.Count -gt 0) {
        $period = "$($dates[0])~$($dates[-1])"
    }

    [PSCustomObject]@{
        Period = $period
        Rows   = $rows
    }
}

function Read-XlsxRows {
    # Excel 없이 xlsx(OOXML zip)를 직접 읽어 2차원 문자열 배열로 반환한다.
    # (배포 도구라 사용자 PC에 Excel이 있다고 가정할 수 없음)
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Get-ZipEntryXml($zip, $entryName) {
            $entry = $zip.Entries | Where-Object { $_.FullName -eq $entryName }
            if (-not $entry) { return $null }
            $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            try { return [xml]$sr.ReadToEnd() } finally { $sr.Close() }
        }

        $sharedStrings = New-Object System.Collections.Generic.List[string]   # 줄마다 += 하면 배열을 매번 새로 만들어 느려진다
        $ssXml = Get-ZipEntryXml $zip "xl/sharedStrings.xml"
        if ($ssXml) {
            $ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ns.AddNamespace("a", $ssXml.DocumentElement.NamespaceURI)
            foreach ($si in $ssXml.SelectNodes("//a:si", $ns)) {
                $texts = $si.SelectNodes(".//a:t", $ns) | ForEach-Object { $_.InnerText }
                $sharedStrings.Add(($texts -join ""))
            }
        }

        # 첫 번째 워크시트 파일 경로 결정 (workbook.xml.rels 매핑)
        $wbXml = Get-ZipEntryXml $zip "xl/workbook.xml"
        $relsXml = Get-ZipEntryXml $zip "xl/_rels/workbook.xml.rels"
        $sheetEntry = "xl/worksheets/sheet1.xml"
        if ($wbXml -and $relsXml) {
            $wbNs = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
            $wbNs.AddNamespace("a", $wbXml.DocumentElement.NamespaceURI)
            $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            $firstSheet = $wbXml.SelectSingleNode("//a:sheets/a:sheet[1]", $wbNs)
            if ($firstSheet) {
                $rid = $firstSheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                $relsNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
                $relsNs.AddNamespace("a", $relsXml.DocumentElement.NamespaceURI)
                $rel = $relsXml.SelectSingleNode("//a:Relationship[@Id='$rid']", $relsNs)
                if ($rel) { $sheetEntry = "xl/" + $rel.GetAttribute("Target").TrimStart("/") }
            }
        }

        $sheetXml = Get-ZipEntryXml $zip $sheetEntry
        if (-not $sheetXml) { throw "$Path : 워크시트를 읽을 수 없습니다 ($sheetEntry)." }

        $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $sheetNs.AddNamespace("a", $sheetXml.DocumentElement.NamespaceURI)

        function ColLetterToIndex([string]$ref) {
            $letters = ($ref -replace '\d', '')
            $idx = 0
            foreach ($ch in $letters.ToCharArray()) { $idx = $idx * 26 + ([int][char]$ch - [int][char]'A' + 1) }
            return $idx - 1
        }

        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($rowNode in $sheetXml.SelectNodes("//a:sheetData/a:row", $sheetNs)) {
            $rowVals = New-Object System.Collections.Generic.List[string]
            foreach ($c in $rowNode.SelectNodes("a:c", $sheetNs)) {
                $colIdx = ColLetterToIndex $c.GetAttribute("r")
                while ($rowVals.Count -le $colIdx) { $rowVals.Add("") }
                $t = $c.GetAttribute("t")
                $val = ""
                if ($t -eq "inlineStr") {
                    $isNode = $c.SelectSingleNode("a:is", $sheetNs)
                    if ($isNode) { $val = ($isNode.SelectNodes(".//a:t", $sheetNs) | ForEach-Object { $_.InnerText }) -join "" }
                } else {
                    $vNode = $c.SelectSingleNode("a:v", $sheetNs)
                    if ($vNode) {
                        if ($t -eq "s") { $val = $sharedStrings[[int]$vNode.InnerText] }
                        else { $val = $vNode.InnerText }
                    }
                }
                $rowVals[$colIdx] = $val
            }
            $rows.Add(@($rowVals))
        }
        return $rows
    } finally {
        $zip.Dispose()
    }
}

function Get-GmarketXlsx {
    # 지마켓 상품별/키워드별 리포트 xlsx를 읽어 헤더 기준 PSCustomObject 행 목록으로 변환한다.
    # 메타 정보(기간/리포트종류)는 라벨로 찾고, 실제 헤더 행은 '노출 수' 컬럼이 있는 행으로 판별한다
    # (지마켓 리포트는 상단에 빈 줄이 섞여 있어 고정 행 번호로 찾을 수 없음).
    param([string]$Path)

    $rows = Read-XlsxRows -Path $Path
    if ($rows.Count -lt 2) { throw "$Path : 데이터가 없습니다." }

    $period = $null
    $reportType = $null
    $headerRowIdx = -1
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if ($r.Count -ge 2 -and $r[0] -eq '기간') { $period = $r[1] }
        if ($r.Count -ge 2 -and $r[0] -eq '리포트 종류') { $reportType = $r[1] }
        if ($r -contains '노출 수') { $headerRowIdx = $i; break }
    }
    if ($headerRowIdx -lt 0) { throw "$Path : 헤더 행('노출 수' 컬럼)을 찾을 수 없습니다." }

    $headers = $rows[$headerRowIdx]
    $dataRows = New-Object System.Collections.Generic.List[object]   # 줄마다 += 하면 배열을 매번 새로 만들어 줄이 많을 때 느려진다
    for ($i = $headerRowIdx + 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if (-not $r -or $r.Count -eq 0) { continue }
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $key = $headers[$c]
            if (-not $key) { continue }
            $obj[$key] = if ($c -lt $r.Count) { $r[$c] } else { '' }
        }
        $dataRows.Add([PSCustomObject]$obj)
    }

    [PSCustomObject]@{
        Period     = $period
        ReportType = $reportType
        Rows       = $dataRows.ToArray()
    }
}

function Get-AuctionXlsx {
    # 옥션(A옥션) 리포트 xlsx를 읽는다. 지마켓과 달리 메타 정보 없이 1행이 바로 헤더이며,
    # 리포트 종류별로 컬럼 구성이 서로 다르기 때문에 헤더의 고유 컬럼명으로 종류를 구분한다.
    # 반환값이 $null이면 옥션 리포트가 아닌 파일(다른 종류의 xlsx)이라는 뜻이다.
    param([string]$Path)

    $rows = Read-XlsxRows -Path $Path
    if ($rows.Count -lt 1) { throw "$Path : 데이터가 없습니다." }

    $headers = $rows[0]
    $reportType = $null
    if ($headers -contains '판매자기준-광고수익률' -and $headers -contains '상품번호') { $reportType = 'AI매출업_상품별' }
    elseif ($headers -contains '판매자기준-광고수익률' -and $headers -contains '날짜별') { $reportType = 'AI매출업_일별' }   # 날짜마다 한 줄 (2.040)
    elseif ($headers -contains '광고상품번호') { $reportType = '파워클릭_상품별' }
    elseif ($headers -contains '키워드' -and $headers -contains '영역명') { $reportType = '파워클릭_키워드별' }
    elseif ($headers -contains '그룹명' -and $headers -contains '영역명') { $reportType = $null }   # 파워클릭 그룹별은 캐츠온 엑셀에만 쓴다
    elseif ($headers -contains '날짜' -and $headers -contains '영역명') { $reportType = '파워클릭_기간별' }   # 날짜마다 한 줄 (2.040)
    if (-not $reportType) { return $null }

    $dataRows = New-Object System.Collections.Generic.List[object]   # 줄마다 += 하면 배열을 매번 새로 만들어 줄이 많을 때 느려진다
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if (-not $r -or $r.Count -eq 0) { continue }
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $key = $headers[$c]
            if (-not $key) { continue }
            $obj[$key] = if ($c -lt $r.Count) { $r[$c] } else { '' }
        }
        $dataRows.Add([PSCustomObject]$obj)
    }

    [PSCustomObject]@{
        ReportType = $reportType
        Rows       = $dataRows.ToArray()
    }
}

function Get-ElevenstCsv {
    # 11번가 광고상품별(AI캠페인/직접캠페인 등) 일별 리포트. UTF-16 탭 구분 CSV이며 1행 헤더, 그 아래 '합계' 행과 일별 행이 온다.
    # 광고상품명과 기간은 파일 안에 없어 파일명('AI캠페인 보고서_daily_2026-09-01_2026-09-30.csv')에서 얻는다.
    # 반환값이 $null이면 11번가 리포트가 아닌 파일(네이버 CSV 등)이라는 뜻이다.
    param([System.IO.FileInfo]$File)

    # 네이버 통합 보고서처럼 큰 CSV를 통째로 읽지 않도록 헤더 한 줄만 먼저 확인한다 (BOM으로 UTF-16/UTF-8 자동 판별).
    $sr = New-Object System.IO.StreamReader($File.FullName, [System.Text.Encoding]::UTF8, $true)
    try { $headerLine = $sr.ReadLine() } finally { $sr.Close() }
    if (-not $headerLine) { return $null }
    $headers = @($headerLine.Split("`t") | ForEach-Object { $_.Trim() })
    if (-not ($headers -contains '날짜' -and $headers -contains '총 전환 금액')) { return $null }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($line in ([System.IO.File]::ReadAllLines($File.FullName) | Select-Object -Skip 1)) {
        $cells = $line.Split("`t")
        if ($cells[0].Trim() -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }  # '합계' 행 제외
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $headers.Count; $c++) { $obj[$headers[$c]] = if ($c -lt $cells.Count) { $cells[$c].Trim() } else { '' } }
        $rows.Add([PSCustomObject]$obj)
    }

    $period = $null
    if ($File.Name -match '(\d{4})-(\d{2})-(\d{2})_(\d{4})-(\d{2})-(\d{2})') {
        $period = "$($matches[1]).$($matches[2]).$($matches[3]).~$($matches[4]).$($matches[5]).$($matches[6])."
    }
    $adType = ($File.BaseName -replace '\s*보고서.*$', '').Trim()
    if (-not $adType) { $adType = $File.BaseName }

    [PSCustomObject]@{
        AdType = $adType
        Period = $period
        Rows   = $rows.ToArray()
    }
}

function Read-HeadLines {
    # 큰 CSV를 통째로 읽지 않고 앞 몇 줄만 읽는다 (BOM으로 UTF-16/UTF-8 자동 판별).
    param([string]$Path, [int]$Count)
    $sr = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try {
        $lines = @()
        while ($lines.Count -lt $Count -and -not $sr.EndOfStream) { $lines += $sr.ReadLine() }
        return , $lines
    } finally { $sr.Close() }
}

function Get-GoogleCsv {
    # 구글 애즈 보고서. 1행 보고서 이름, 2행 기간, 3행 컬럼이다. 엑셀용 CSV는 탭 구분 UTF-16, 일반 CSV는 쉼표 구분이다.
    # '일' 컬럼 없이 캠페인 단위로만 받아도 읽는다 (일별 차트만 빠진다). 그때 기간은 2행 글에서 읽는다.
    # 반환값이 $null이면 구글 보고서가 아닌 파일이라는 뜻이다.
    param([System.IO.FileInfo]$File)

    $head = Read-HeadLines -Path $File.FullName -Count 6
    $headerIdx = -1
    $delimiter = "`t"
    $hasDay = $false
    for ($i = 0; $i -lt $head.Count -and $headerIdx -lt 0; $i++) {
        foreach ($d in "`t", ",") {
            $cols = @($head[$i].Split($d) | ForEach-Object { $_.Trim().Trim('"') })
            if (($cols -contains '캠페인') -and ($cols -contains '비용')) { $headerIdx = $i; $delimiter = $d; $hasDay = $cols -contains '일'; break }
        }
    }
    if ($headerIdx -lt 0) { return $null }

    $lines = [System.IO.File]::ReadAllLines($File.FullName)
    $all = @($lines[$headerIdx..($lines.Count - 1)] | ConvertFrom-Csv -Delimiter $delimiter)
    if ($hasDay) {
        # 맨 아래에 붙는 '총계' 행은 날짜가 없어서 여기서 걸러진다.
        $rows = @($all | Where-Object { $_.'일' -match '^\d{4}-\d{2}-\d{2}$' })
    } else {
        # 날짜가 없으면 '총계: 계정' 같은 합계 행을 첫 칸 글과 캠페인 칸으로 거른다.
        $rows = @($all | Where-Object { $c = "$($_.'캠페인')".Trim(); $first = "$(@($_.PSObject.Properties)[0].Value)".Trim(); $c -and $c -ne '--' -and $c -notlike '총계*' -and $first -notlike '총계*' })
    }
    # 2행 기간 글 (예: '2026년 10월 1일 - 2026년 10월 31일')
    $period = $null
    for ($i = 0; $i -lt $headerIdx; $i++) {
        if ($head[$i] -match '(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*[-~]\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일') {
            $period = '{0}.{1:D2}.{2:D2}.~{3}.{4:D2}.{5:D2}.' -f $matches[1], [int]$matches[2], [int]$matches[3], $matches[4], [int]$matches[5], [int]$matches[6]
            break
        }
    }
    [PSCustomObject]@{ Rows = $rows; Period = $period }
}

function Get-MetaCsv {
    # 메타 광고 보고서. 1행이 컬럼이고, 비용 컬럼 이름에 통화가 들어 있다 (예: '지출 금액 (USD)').
    # 전환수는 '결과' 컬럼을 쓴다. 반환값이 $null이면 메타 보고서가 아닌 파일이라는 뜻이다.
    param([System.IO.FileInfo]$File)

    $head = Read-HeadLines -Path $File.FullName -Count 1
    if (-not $head.Count) { return $null }
    if ($head[0] -notmatch '캠페인 이름' -or -not ($head[0] -match '지출 금액 \(([A-Za-z]{3})\)')) { return $null }
    $currency = $matches[1]
    $rows = @(Import-Csv -LiteralPath $File.FullName -Encoding UTF8)
    if (-not $rows.Count) { return $null }
    [PSCustomObject]@{
        Currency    = $currency
        CostColumn  = "지출 금액 ($currency)"
        ClickColumn = if ($rows[0].PSObject.Properties.Name -contains '클릭(전체)') { '클릭(전체)' } else { '링크 클릭' }
        Rows        = $rows
    }
}

function Stop-WithMessage {
    # 사용자가 고칠 수 있는 입력 문제. HTA가 이 문구를 그대로 보여준다 (종료 코드 4).
    param([string]$Message)
    if (Test-Path -LiteralPath $naverStage) {
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        Remove-Item -LiteralPath $naverStage -Recurse -Force -ErrorAction SilentlyContinue
    }
    [Console]::Error.WriteLine($Message)
    exit 4
}

function Assert-Currency {
    # 광고비는 보고서 통화 그대로 보여주고 환율로 바꾸지 않는다. 화면이 원화와 달러만 구분하므로 다른 통화는 막는다.
    param([string]$Currency, [string]$MediaName)
    if ($Currency -and $Currency -ne 'KRW' -and $Currency -ne 'USD') { Stop-WithMessage "$MediaName 보고서의 통화($Currency)는 지원하지 않습니다. 원화(KRW)나 달러(USD)로 받아 주세요." }
}

function ToNum {
    param($v)
    if ($null -eq $v -or $v -eq '') { return 0 }
    $s = ($v -replace ',', '').Trim()
    $d = 0.0
    if ([double]::TryParse($s, [ref]$d)) { return $d }
    return 0
}

function Round2($v) { [math]::Round($v, 2) }

function New-Metrics {
    param([double]$Impressions, [double]$Clicks, [double]$Cost, [double]$Conversions, [double]$Revenue)
    $ctr = if ($Impressions -gt 0) { Round2(($Clicks / $Impressions) * 100) } else { 0 }
    $cpc = if ($Clicks -gt 0) { [math]::Round($Cost / $Clicks) } else { 0 }
    $roas = if ($Cost -gt 0) { Round2(($Revenue / $Cost) * 100) } else { 0 }
    [PSCustomObject]@{
        impressions = [int64]$Impressions
        clicks      = [int64]$Clicks
        ctr         = $ctr
        cost        = [int64]$Cost
        cpc         = $cpc
        # 구글 전환은 기여 모델에 따라 소수로 나온다. 정수면 전과 같이 정수로 둔다.
        conversions = if ($Conversions -eq [math]::Truncate($Conversions)) { [int64]$Conversions } else { Round2 $Conversions }
        revenue     = [int64]$Revenue
        roas        = $roas
    }
}

# 네이버 통합(다차원) 보고서는 수만 행이라, 행마다 PowerShell 함수(ToNum/Add-Agg)를 부르면 몇 분씩 걸린다.
# 같은 동작을 컴파일된 C#으로 둬서 대용량 루프(지면별/검색어/요일·일자·기기별)에서만 쓴다.
Add-Type -TypeDefinition @"
using System.Collections;
using System.Globalization;
public static class AdFast {
    public static double ToNum(object v) {
        if (v == null) return 0;
        string s = v.ToString().Replace(",", "").Trim();
        double d;
        return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out d) ? d : 0;
    }
    public static void AddAgg(Hashtable a, double imp, double clk, double cost, double conv, double rev) {
        a["imp"] = (double)a["imp"] + imp; a["clk"] = (double)a["clk"] + clk; a["cost"] = (double)a["cost"] + cost;
        a["conv"] = (double)a["conv"] + conv; a["rev"] = (double)a["rev"] + rev;
    }
    static Hashtable Slot(Hashtable map, string key) {
        Hashtable a = (Hashtable)map[key];
        if (a == null) { a = new Hashtable(); a["imp"] = 0.0; a["clk"] = 0.0; a["cost"] = 0.0; a["conv"] = 0.0; a["rev"] = 0.0; map[key] = a; }
        return a;
    }
    // 캠페인별, 일별, 캠페인별 일별 합계를 한 행씩 쌓는다 (네이버 검색광고, GFA, 구글, 메타 공통).
    public static void AddChannelRow(Hashtable ch, string type, string campaign, string date, double imp, double clk, double cost, double conv, double rev) {
        AddAgg((Hashtable)ch["total"], imp, clk, cost, conv, rev);
        string c = type + "\t" + campaign;
        AddAgg(Slot((Hashtable)ch["camp"], c), imp, clk, cost, conv, rev);
        if (string.IsNullOrEmpty(date)) return;
        AddAgg(Slot((Hashtable)ch["day"], date), imp, clk, cost, conv, rev);
        AddAgg(Slot((Hashtable)ch["campDay"], c + "\t" + date), imp, clk, cost, conv, rev);
    }
}
"@

# 네이버 보고서는 CSV 대신 엑셀(xlsx)로도 내려받을 수 있다. 내용(1행 제목+기간, 2행 헤더, 3행부터 데이터)은 같아서
# xlsx를 같은 모양의 CSV로 바꿔 기존 CSV 처리에 그대로 넘긴다. 수만 행이라 PowerShell XML 대신 C#으로 흘려 읽는다.
# 네이버 보고서가 아니면(1행에 '보고서(YYYY.MM.DD.~YYYY.MM.DD.)'가 없으면) 아무것도 만들지 않고 false를 돌려준다.
Add-Type -ReferencedAssemblies System.IO.Compression, System.IO.Compression.FileSystem, System.Xml -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
public static class NaverXlsx {
    static List<string> SharedStrings(ZipArchive z) {
        List<string> list = new List<string>();
        ZipArchiveEntry e = z.GetEntry("xl/sharedStrings.xml");
        if (e == null) return list;
        using (XmlReader r = XmlReader.Create(e.Open())) {
            StringBuilder sb = null;
            while (!r.EOF) {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "si") {
                    if (r.IsEmptyElement) { list.Add(""); r.Read(); continue; }
                    sb = new StringBuilder();
                } else if (r.NodeType == XmlNodeType.Element && r.LocalName == "rPh") {
                    r.Skip(); continue;   // 일본어 읽기 표시 등은 글자에 넣지 않는다
                } else if (r.NodeType == XmlNodeType.Element && r.LocalName == "t" && sb != null) {
                    sb.Append(r.ReadElementContentAsString()); continue;
                } else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "si") {
                    list.Add(sb == null ? "" : sb.ToString()); sb = null;
                }
                r.Read();
            }
        }
        return list;
    }
    static string FirstSheet(ZipArchive z) {
        string path = "xl/worksheets/sheet1.xml";
        try {
            ZipArchiveEntry wb = z.GetEntry("xl/workbook.xml"), rels = z.GetEntry("xl/_rels/workbook.xml.rels");
            if (wb == null || rels == null) return path;
            XmlDocument w = new XmlDocument(); using (Stream s = wb.Open()) w.Load(s);
            XmlNamespaceManager ns = new XmlNamespaceManager(w.NameTable);
            ns.AddNamespace("a", w.DocumentElement.NamespaceURI);
            XmlElement sheet = (XmlElement)w.SelectSingleNode("//a:sheets/a:sheet[1]", ns);
            if (sheet == null) return path;
            string rid = sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            XmlDocument rd = new XmlDocument(); using (Stream s = rels.Open()) rd.Load(s);
            foreach (XmlElement rel in rd.DocumentElement.ChildNodes) {
                if (rel.GetAttribute("Id") == rid) {
                    string t = rel.GetAttribute("Target").TrimStart('/');
                    return t.StartsWith("xl/") ? t : "xl/" + t;
                }
            }
        } catch { }
        return path;
    }
    static int ColIndex(string cellRef) {
        int idx = 0;
        foreach (char ch in cellRef) { if (ch >= 'A' && ch <= 'Z') idx = idx * 26 + (ch - 'A' + 1); else break; }
        return idx - 1;
    }
    static string Csv(string v) {
        if (v.IndexOfAny(new char[] { ',', '"', '\r', '\n' }) < 0) return v;
        return "\"" + v.Replace("\"", "\"\"") + "\"";
    }
    public static bool ToCsv(string xlsxPath, string csvPath) {
        using (ZipArchive z = ZipFile.OpenRead(xlsxPath)) {
            List<string> ss = SharedStrings(z);
            ZipArchiveEntry sheet = z.GetEntry(FirstSheet(z));
            if (sheet == null) return false;
            List<string[]> rows = new List<string[]>();
            using (XmlReader r = XmlReader.Create(sheet.Open())) {
                List<string> cur = null; int col = -1; string type = null; bool inValue = false;
                StringBuilder val = new StringBuilder();
                while (r.Read()) {
                    if (r.NodeType == XmlNodeType.Element) {
                        if (r.LocalName == "row") { cur = new List<string>(); if (r.IsEmptyElement) { rows.Add(cur.ToArray()); cur = null; } }
                        else if (r.LocalName == "c" && cur != null) {
                            string cr = r.GetAttribute("r");
                            col = cr == null ? cur.Count : ColIndex(cr);
                            type = r.GetAttribute("t");
                            val.Length = 0;
                            while (cur.Count <= col) cur.Add("");
                        }
                        else if ((r.LocalName == "v" || r.LocalName == "t") && cur != null && col >= 0) inValue = !r.IsEmptyElement;
                    } else if ((r.NodeType == XmlNodeType.Text || r.NodeType == XmlNodeType.SignificantWhitespace || r.NodeType == XmlNodeType.Whitespace) && inValue) {
                        val.Append(r.Value);
                    } else if (r.NodeType == XmlNodeType.EndElement) {
                        if (r.LocalName == "v" || r.LocalName == "t") inValue = false;
                        else if (r.LocalName == "c" && cur != null && col >= 0) {
                            string v = val.ToString();
                            if (type == "s") { int i; v = int.TryParse(v, out i) && i < ss.Count ? ss[i] : ""; }
                            cur[col] = v; col = -1;
                        }
                        else if (r.LocalName == "row" && cur != null) { rows.Add(cur.ToArray()); cur = null; }
                    }
                    // 첫 줄만 읽고도 네이버 보고서가 아니면 바로 그만둔다
                    if (rows.Count == 1 && cur == null && !IsNaverTitle(rows[0])) return false;
                }
            }
            if (rows.Count < 2 || !IsNaverTitle(rows[0])) return false;
            using (StreamWriter w = new StreamWriter(csvPath, false, new UTF8Encoding(true))) {
                w.Write("\"" + rows[0][0].Replace("\"", "\"\"") + "\"\r\n");
                for (int i = 1; i < rows.Count; i++) {
                    string[] row = rows[i];
                    if (row.Length == 0) continue;
                    for (int c = 0; c < row.Length; c++) { if (c > 0) w.Write(','); w.Write(Csv(row[c] ?? "")); }
                    w.Write("\r\n");
                }
            }
            return true;
        }
    }
    static bool IsNaverTitle(string[] row) {
        return row.Length > 0 && row[0] != null && Regex.IsMatch(row[0], @"보고서\(\d{4}\.\d{2}\.\d{2}\.~\d{4}\.\d{2}\.\d{2}\.\)");
    }
}
'@

# 오늘의집 CSV 읽기와 묶기. 검색어 보고서는 한 캠페인만 받아도 4만 줄이 넘어서(2026-10, 41,614줄) C#으로 읽고 합친다.
Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
public static class OhouseCsv {
    // 1행이 컬럼(따옴표 없음)이고 2행부터 모든 칸이 따옴표로 싸여 있다. 검색어 칸 안에 줄바꿈이 있을 수 있다.
    // 끝 글자가 따옴표(")인 검색어는 '""' 다음에 닫는 따옴표 없이 쉼표가 온다 (예: "...ㄱ"","자동"). 표준대로 읽으면 칸이 하나씩 밀려
    // 비용 자리에 다른 값이 들어가므로, 따옴표 칸 안에서 '""' 바로 뒤가 쉼표나 줄 끝이면 따옴표 하나를 넣고 칸을 끝낸다.
    public static List<string[]> Read(string path) {
        string s = File.ReadAllText(path, Encoding.UTF8);
        List<string[]> rows = new List<string[]>();
        List<string> cur = new List<string>();
        StringBuilder f = new StringBuilder();
        bool inQ = false;
        int n = s.Length, i = 0;
        while (i < n) {
            char c = s[i];
            if (inQ) {
                if (c == '"') {
                    if (i + 1 < n && s[i + 1] == '"') {
                        f.Append('"');
                        i += 2;
                        if (i >= n || s[i] == ',' || s[i] == '\r' || s[i] == '\n') inQ = false;
                        continue;
                    }
                    inQ = false; i++; continue;
                }
                f.Append(c); i++; continue;
            }
            if (c == '"') { inQ = true; i++; continue; }
            if (c == ',') { cur.Add(f.ToString()); f.Length = 0; i++; continue; }
            if (c == '\r' || c == '\n') {
                if (c == '\r' && i + 1 < n && s[i + 1] == '\n') i++;
                i++;
                cur.Add(f.ToString()); f.Length = 0;
                if (!(cur.Count == 1 && cur[0].Length == 0)) rows.Add(cur.ToArray());   // 빈 줄은 뺀다
                cur = new List<string>();
                continue;
            }
            f.Append(c); i++;
        }
        if (f.Length > 0 || cur.Count > 0) { cur.Add(f.ToString()); rows.Add(cur.ToArray()); }
        return rows;
    }
    // 대소문자를 구분하는 표 (PowerShell의 @{}는 'LED'와 'led' 검색어를 하나로 합쳐 버린다)
    public static Hashtable NewMap() { return new Hashtable(); }
    // 카테고리, 상품(전체) 보고서는 파일 맨 앞에 BOM이 두 번 있어, 읽은 뒤에도 첫 컬럼 이름 앞에 하나가 남는다. 그래서 컬럼 이름에서 BOM을 지운다.
    static int Col(string[] header, string name) {
        for (int i = 0; i < header.Length; i++) if (header[i].Trim().Trim('\uFEFF').Trim('"') == name) return i;
        return -1;
    }
    static int[] Cols(string[] header, string[] names) {
        int[] r = new int[names.Length];
        for (int i = 0; i < names.Length; i++) r[i] = Col(header, names[i]);
        return r;
    }
    static string Cell(string[] row, int i) { return (i >= 0 && i < row.Length) ? row[i] : ""; }
    static string Key(string[] row, int[] cols) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < cols.Length; i++) { if (i > 0) sb.Append('\t'); sb.Append(Cell(row, cols[i])); }
        return sb.ToString();
    }
    public static double Num(string v) {
        if (v == null) return 0;
        string s = v.Replace(",", "").Replace("%", "").Trim();
        double d;
        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0;
    }
    // 한 파일(rows[0]이 컬럼)의 줄을 keyCols 값으로 묶어 map에 더한다. map 값은 imp, clk, cost, conv, rev(New-Agg와 같은 모양)와 metaCols의 첫 줄 값이다.
    // 전환은 '광고 전환 판매 수', 매출은 '광고 전환 매출'(카테고리, 상품(전체) 보고서는 '전환 매출', 둘 다 직접+간접)이다. 오늘의집 '광고 수익률'과 같은 기준.
    // dupCols 값이 앞서 읽은 파일(seen)에 이미 있던 줄은 건너뛴다 (같은 보고서를 두 번 넣은 경우). 숫자가 모두 0인 줄도 건너뛴다.
    public static int Aggregate(List<string[]> rows, string[] keyCols, string[] dupCols, string[] metaCols, Hashtable seen, Hashtable map) {
        if (rows.Count < 1) return 0;
        string[] h = rows[0];
        int[] k = Cols(h, keyCols), d = Cols(h, dupCols), m = Cols(h, metaCols);
        int imp = Col(h, "노출수"), clk = Col(h, "클릭수"), cost = Col(h, "집행 광고비"), rev = Col(h, "광고 전환 매출");
        if (rev < 0) rev = Col(h, "전환 매출");   // 카테고리, 상품(전체) 보고서는 '전환 매출'
        int conv = Col(h, "광고 전환 판매 수");
        if (conv < 0) conv = Col(h, "광고 전환 판매수");   // 검색어 보고서는 띄어쓰기가 다르다
        Hashtable mine = new Hashtable();
        int added = 0;
        for (int r = 1; r < rows.Count; r++) {
            string[] row = rows[r];
            double vi = Num(Cell(row, imp)), vc = Num(Cell(row, clk)), vs = Num(Cell(row, cost)), vn = Num(Cell(row, conv)), vr = Num(Cell(row, rev));
            if (vi == 0 && vc == 0 && vs == 0 && vn == 0 && vr == 0) continue;
            if (d.Length > 0) {
                string dk = Key(row, d);
                if (seen.ContainsKey(dk)) continue;
                mine[dk] = true;
            }
            string key = Key(row, k);
            Hashtable a = (Hashtable)map[key];
            if (a == null) {
                a = new Hashtable();
                a["imp"] = 0.0; a["clk"] = 0.0; a["cost"] = 0.0; a["conv"] = 0.0; a["rev"] = 0.0;
                for (int j = 0; j < m.Length; j++) a[metaCols[j]] = Cell(row, m[j]);
                map[key] = a;
            }
            a["imp"] = (double)a["imp"] + vi; a["clk"] = (double)a["clk"] + vc; a["cost"] = (double)a["cost"] + vs;
            a["conv"] = (double)a["conv"] + vn; a["rev"] = (double)a["rev"] + vr;
            added++;
        }
        foreach (DictionaryEntry e in mine) seen[e.Key] = true;
        return added;
    }
    // map 전체 합계: 노출, 클릭, 비용, 전환, 매출, 전환이 없는 묶음의 비용
    public static double[] Totals(Hashtable map) {
        double[] t = new double[6];
        foreach (Hashtable a in map.Values) {
            t[0] += (double)a["imp"]; t[1] += (double)a["clk"]; t[2] += (double)a["cost"]; t[3] += (double)a["conv"]; t[4] += (double)a["rev"];
            if ((double)a["conv"] == 0) t[5] += (double)a["cost"];
        }
        return t;
    }
    // 클릭이 1회 이상인 묶음의 키
    public static string[] KeysWithClicks(Hashtable map) {
        List<string> r = new List<string>();
        foreach (DictionaryEntry e in map) if ((double)((Hashtable)e.Value)["clk"] > 0) r.Add((string)e.Key);
        return r.ToArray();
    }
    // metaCol 값의 종류 (예: 검색어 보고서에 들어 있는 캠페인 이름)
    public static string[] MetaValues(Hashtable map, string metaCol) {
        Hashtable set = new Hashtable();
        List<string> r = new List<string>();
        foreach (Hashtable a in map.Values) { string v = (a[metaCol] as string) ?? ""; if (!set.ContainsKey(v)) { set[v] = true; r.Add(v); } }
        return r.ToArray();
    }
    // metaCol 값의 종류 수 (예: 검색어 수)
    public static int DistinctMeta(Hashtable map, string metaCol) {
        Hashtable set = new Hashtable();   // HashSet은 System.Core에 있어 Add-Type 기본 참조로는 안 될 수 있다
        foreach (Hashtable a in map.Values) set[(a[metaCol] as string) ?? ""] = true;
        return set.Count;
    }
}
'@

# 네이버 CSV(지면별, 검색어, 키워드, 요일 · 일자 · 기기별, 통합 보고서) 읽기와 묶기 (2.032).
# 예전에는 Get-Content + ConvertFrom-Csv로 줄마다 PowerShell 개체를 만들어 1년치 검색어 보고서처럼 수십만 줄이면 몇 분씩 걸리거나 메모리가 모자랐다.
# 지금은 C#으로 한 줄씩 흘려 읽으면서 바로 묶고, PowerShell은 묶인 결과만 돈다. 시간대별, 지역, 전환유형처럼 작은 보고서는 예전 방식(Get-NaverCsv) 그대로다.
# 묶은 값의 모양은 OhouseCsv.Aggregate와 같다(imp, clk, cost, conv, rev + 첫 줄의 metaCols 값). 그래서 PowerShell에서 Totals, DistinctMeta는 OhouseCsv 것을 같이 쓴다.
Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Text;
public static class NaverCsv {
    // 표준 CSV 한 줄(따옴표 칸 안의 쉼표, 줄바꿈, "" 포함)을 읽는다. 파일 끝이면 null.
    static string[] Next(TextReader r, StringBuilder f) {
        List<string> cur = new List<string>();
        f.Length = 0;
        bool inQ = false, any = false;
        int c;
        while ((c = r.Read()) >= 0) {
            any = true;
            char ch = (char)c;
            if (inQ) {
                if (ch == '"') {
                    if (r.Peek() == '"') { r.Read(); f.Append('"'); }
                    else inQ = false;
                } else f.Append(ch);
                continue;
            }
            if (ch == '"') { inQ = true; continue; }
            if (ch == ',') { cur.Add(f.ToString()); f.Length = 0; continue; }
            if (ch == '\r') { if (r.Peek() == '\n') r.Read(); cur.Add(f.ToString()); return cur.ToArray(); }
            if (ch == '\n') { cur.Add(f.ToString()); return cur.ToArray(); }
            f.Append(ch);
        }
        if (!any) return null;
        cur.Add(f.ToString());
        return cur.ToArray();
    }
    static int Col(string[] h, string name) {
        for (int i = 0; i < h.Length; i++) if (h[i].Trim().Trim('﻿').Trim('"') == name) return i;
        return -1;
    }
    static string Cell(string[] row, int i) { return (i >= 0 && i < row.Length) ? row[i] : ""; }
    // OhouseCsv.Num과 같다 (Add-Type 블록끼리는 서로의 형식을 볼 수 없어 따로 둔다)
    static double Num(string v) {
        if (v == null) return 0;
        string s = v.Replace(",", "").Replace("%", "").Trim();
        double d;
        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0;
    }
    // 2행(컬럼 이름)
    public static string[] Header(string path) {
        using (StreamReader r = new StreamReader(path, Encoding.UTF8, true)) {
            StringBuilder f = new StringBuilder();
            if (Next(r, f) == null) return new string[0];
            string[] h = Next(r, f);
            if (h == null) return new string[0];
            for (int i = 0; i < h.Length; i++) h[i] = h[i].Trim().Trim('﻿');
            return h;
        }
    }
    // 3행부터 keyCols 값으로 묶어 map에 더한다. 전환은 '구매완료 전환수'(없으면 '총 전환수', 문의 · 신청 전환 업체), 매출은 '구매완료 전환매출액(원)'.
    // 숫자가 모두 0인 줄은 건너뛴다. 같은 값이 많은 캠페인, 광고그룹 이름 등은 한 문자열을 같이 써서 메모리를 아낀다. 반환값: 읽은 데이터 줄 수
    public static int Aggregate(string path, string[] keyCols, string[] metaCols, Hashtable map) {
        using (StreamReader r = new StreamReader(path, Encoding.UTF8, true)) {
            StringBuilder f = new StringBuilder();
            if (Next(r, f) == null) return 0;      // 1행: 제목 + 기간
            string[] h = Next(r, f);                // 2행: 컬럼
            if (h == null) return 0;
            int[] k = new int[keyCols.Length], m = new int[metaCols.Length];
            for (int i = 0; i < keyCols.Length; i++) k[i] = Col(h, keyCols[i]);
            for (int i = 0; i < metaCols.Length; i++) m[i] = Col(h, metaCols[i]);
            int imp = Col(h, "노출수"), clk = Col(h, "클릭수"), cost = Col(h, "총비용"), rev = Col(h, "구매완료 전환매출액(원)");
            int conv = Col(h, "구매완료 전환수");
            if (conv < 0) conv = Col(h, "총 전환수");
            Dictionary<string, string> pool = new Dictionary<string, string>();
            StringBuilder kb = new StringBuilder();
            int n = 0;
            string[] row;
            while ((row = Next(r, f)) != null) {
                if (row.Length == 1 && row[0].Length == 0) continue;   // 빈 줄
                n++;
                double vi = Num(Cell(row, imp)), vc = Num(Cell(row, clk)), vs = Num(Cell(row, cost));
                double vn = Num(Cell(row, conv)), vr = Num(Cell(row, rev));
                if (vi == 0 && vc == 0 && vs == 0 && vn == 0 && vr == 0) continue;
                kb.Length = 0;
                for (int i = 0; i < k.Length; i++) { if (i > 0) kb.Append('\t'); kb.Append(Cell(row, k[i])); }
                string key = kb.ToString();
                Hashtable a = (Hashtable)map[key];
                if (a == null) {
                    a = new Hashtable();
                    a["imp"] = 0.0; a["clk"] = 0.0; a["cost"] = 0.0; a["conv"] = 0.0; a["rev"] = 0.0;
                    for (int j = 0; j < m.Length; j++) {
                        string v = Cell(row, m[j]), p;
                        if (pool.TryGetValue(v, out p)) v = p; else pool[v] = v;
                        a[metaCols[j]] = v;
                    }
                    map[key] = a;
                }
                a["imp"] = (double)a["imp"] + vi; a["clk"] = (double)a["clk"] + vc; a["cost"] = (double)a["cost"] + vs;
                a["conv"] = (double)a["conv"] + vn; a["rev"] = (double)a["rev"] + vr;
            }
            return n;
        }
    }
    // 묶은 표를 더 큰 단위(metaCols 값)로 다시 묶는다. 예: 캠페인 · 광고그룹 · 검색어 → 검색어 · 캠페인유형
    public static Hashtable Rollup(Hashtable src, string[] metaCols) {
        Hashtable dst = new Hashtable();
        StringBuilder kb = new StringBuilder();
        foreach (Hashtable a in src.Values) {
            kb.Length = 0;
            for (int i = 0; i < metaCols.Length; i++) { if (i > 0) kb.Append('\t'); kb.Append((a[metaCols[i]] as string) ?? ""); }
            string key = kb.ToString();
            Hashtable b = (Hashtable)dst[key];
            if (b == null) {
                b = new Hashtable();
                b["imp"] = 0.0; b["clk"] = 0.0; b["cost"] = 0.0; b["conv"] = 0.0; b["rev"] = 0.0;
                for (int i = 0; i < metaCols.Length; i++) b[metaCols[i]] = (a[metaCols[i]] as string) ?? "";
                dst[key] = b;
            }
            b["imp"] = (double)b["imp"] + (double)a["imp"]; b["clk"] = (double)b["clk"] + (double)a["clk"]; b["cost"] = (double)b["cost"] + (double)a["cost"];
            b["conv"] = (double)b["conv"] + (double)a["conv"]; b["rev"] = (double)b["rev"] + (double)a["rev"];
        }
        return dst;
    }
    // 노출이 1회 이상인 묶음을 보고서 줄로 만든다 (비용 많은 순, 비용이 같으면 노출 많은 순). 지표 계산은 New-Metrics와 같다.
    // 2.036부터 클릭 1회 이상이 아니라 노출 1회 이상이다 (사용자 요청). 1년치 검색어는 줄이 수십만 개라, PowerShell 개체를 만들어
    // ConvertTo-Json에 넘기면 몇 분씩 걸린다. 그래서 여기서 바로 JSON 글자로 쓴다. count[0]에 줄 수를 넣는다.
    // 모양: {"cols":[항목 이름...],"rows":[[값...],...]}. 줄마다 항목 이름을 되풀이하지 않아 파일이 반쯤 작다(30만 줄 약 60MB -> 30MB). 보고서 화면이 줄 개체로 펼친다.
    // names[i]는 metaCols[i]를 담을 보고서 이름(예: '검색어' -> keyword).
    public static string ShownRowsJson(Hashtable map, string[] metaCols, string[] names, int[] count) {
        List<Hashtable> list = new List<Hashtable>();
        foreach (Hashtable a in map.Values) if ((double)a["imp"] > 0 || (double)a["clk"] > 0 || (double)a["cost"] > 0) list.Add(a);
        list.Sort(delegate(Hashtable x, Hashtable y) {
            int c = ((double)y["cost"]).CompareTo((double)x["cost"]);
            return c != 0 ? c : ((double)y["imp"]).CompareTo((double)x["imp"]);
        });
        StringBuilder sb = new StringBuilder(list.Count * 110 + 200);
        sb.Append("{\"cols\":[");
        for (int i = 0; i < names.Length; i++) { JsonStr(sb, names[i]); sb.Append(','); }
        sb.Append("\"impressions\",\"clicks\",\"ctr\",\"cost\",\"cpc\",\"conversions\",\"revenue\",\"roas\"],\"rows\":[");
        for (int n = 0; n < list.Count; n++) {
            Hashtable a = list[n];
            double imp = (double)a["imp"], clk = (double)a["clk"], cost = (double)a["cost"], conv = (double)a["conv"], rev = (double)a["rev"];
            if (n > 0) sb.Append(',');
            sb.Append('[');
            for (int i = 0; i < metaCols.Length; i++) { JsonStr(sb, (a[metaCols[i]] as string) ?? ""); sb.Append(','); }
            sb.Append(((long)imp).ToString(CultureInfo.InvariantCulture)).Append(',');
            sb.Append(((long)clk).ToString(CultureInfo.InvariantCulture)).Append(',');
            sb.Append(Dbl(imp > 0 ? Math.Round(clk / imp * 100, 2) : 0.0)).Append(',');
            sb.Append(((long)cost).ToString(CultureInfo.InvariantCulture)).Append(',');
            sb.Append(Dbl(clk > 0 ? Math.Round(cost / clk) : 0.0)).Append(',');
            sb.Append(conv == Math.Truncate(conv) ? ((long)conv).ToString(CultureInfo.InvariantCulture) : Dbl(Math.Round(conv, 2))).Append(',');
            sb.Append(((long)rev).ToString(CultureInfo.InvariantCulture)).Append(',');
            sb.Append(Dbl(cost > 0 ? Math.Round(rev / cost * 100, 2) : 0.0));
            sb.Append(']');
        }
        sb.Append("]}");
        count[0] = list.Count;
        return sb.ToString();
    }
    static string Dbl(double v) { return v.ToString("R", CultureInfo.InvariantCulture); }
    // JSON 글자. '<'도 바꿔서 검색어에 '</script>'가 있어도 보고서 HTML이 깨지지 않게 한다
    static void JsonStr(StringBuilder sb, string s) {
        sb.Append('"');
        foreach (char ch in s) {
            switch (ch) {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '<': sb.Append("\\u003c"); break;
                default:
                    if (ch < ' ' || ch == '\u2028' || ch == '\u2029') sb.Append("\\u").Append(((int)ch).ToString("x4"));
                    else sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
    }
}
'@

function New-Agg { @{ imp = 0.0; clk = 0.0; cost = 0.0; conv = 0.0; rev = 0.0 } }

function Add-Agg {
    param($Agg, $Imp, $Clk, $Cost, $Conv, $Rev)
    $Agg.imp += $Imp; $Agg.clk += $Clk; $Agg.cost += $Cost; $Agg.conv += $Conv; $Agg.rev += $Rev
}

function Metrics-FromAgg { param($Agg) New-Metrics -Impressions $Agg.imp -Clicks $Agg.clk -Cost $Agg.cost -Conversions $Agg.conv -Revenue $Agg.rev }

function Metrics-FromChannelAgg {
    # 달러 보고서는 광고비가 소수라서 정수로 자르지 않는다.
    param($Agg, [string]$Currency)
    $m = Metrics-FromAgg $Agg
    if ($Currency -eq 'USD') {
        $m.cost = Round2 $Agg.cost
        $m.cpc = if ($Agg.clk -gt 0) { Round2 ($Agg.cost / $Agg.clk) } else { 0 }
    }
    return $m
}

function New-Channel { @{ total = (New-Agg); camp = @{}; day = @{}; campDay = @{} } }

function Expand-NaverZips {
    # GFA 리포트는 지면별/검색어 등과 달리 CSV가 아니라 zip 압축파일로 다운로드되며,
    # 내부 CSV 파일명이 'result.csv'처럼 리포트 종류를 알 수 없는 이름일 수 있다.
    # 그래서 압축을 풀어 임시 폴더에 두고, 헤더로 리포트 종류를 판별한다.
    param([string]$Dir)
    $zips = Get-ChildItem -LiteralPath $Dir -Filter "*.zip" -File -ErrorAction SilentlyContinue
    $extracted = @()
    if (-not $zips) { return $extracted }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($z in $zips) {
        $destDir = Join-Path ([System.IO.Path]::GetTempPath()) ("naver_report_zip_" + [System.IO.Path]::GetFileNameWithoutExtension($z.Name))
        try {
            if (Test-Path $destDir) { Remove-Item -LiteralPath $destDir -Recurse -Force }
            New-Item -ItemType Directory -Path $destDir | Out-Null
            [System.IO.Compression.ZipFile]::ExtractToDirectory($z.FullName, $destDir)
            $extracted += @(Get-ChildItem -LiteralPath $destDir -Filter "*.csv" -Recurse -File)
        } catch {
            Write-Warning "압축 해제 실패: $($z.Name) - $($_.Exception.Message)"
        }
    }
    return $extracted
}

# ---------- 네이버 보고서를 엑셀(xlsx)로 받은 경우 ----------
# 네이버 xlsx는 같은 이름의 CSV로 바꾸고, 원본 폴더의 나머지 파일(GFA zip, 지마켓/옥션 xlsx 등)과 함께 임시 폴더에 모은다.
# 아래의 파일 찾기는 모두 그 임시 폴더를 입력 폴더로 쓴다. (원본 폴더의 파일은 바꾸지 않는다)
$sourceInputDir = $InputDir
$naverXlsxNames = @()
$naverStage = Join-Path ([System.IO.Path]::GetTempPath()) ("naver_xlsx_" + [guid]::NewGuid().ToString('N'))
foreach ($x in (Get-ChildItem -LiteralPath $InputDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue)) {
    if (-not (Test-Path -LiteralPath $naverStage)) { New-Item -ItemType Directory -Path $naverStage | Out-Null }
    $csvOut = Join-Path $naverStage (([System.IO.Path]::GetFileNameWithoutExtension($x.Name)).Trim() + '.csv')
    try {
        if ([NaverXlsx]::ToCsv($x.FullName, $csvOut)) { $naverXlsxNames += $x.Name }
    } catch {
        Write-Warning "네이버 엑셀 파일 읽기 실패: $($x.Name) - $($_.Exception.Message)"
    }
}
if ($naverXlsxNames.Count -gt 0) {
    foreach ($f in (Get-ChildItem -LiteralPath $InputDir -File)) {
        $dest = Join-Path $naverStage $f.Name
        if ($naverXlsxNames -notcontains $f.Name -and -not (Test-Path -LiteralPath $dest)) { Copy-Item -LiteralPath $f.FullName -Destination $dest }
    }
    $InputDir = $naverStage
}

# ---------- 네이버 CSV 종류는 파일 이름이 아니라 컬럼 이름으로 알아본다 ----------
# 네이버 보고서는 1행이 제목+기간('...보고서(2026.10.01.~2026.10.26.)'), 2행이 컬럼이다.
# (카카오·구글 등 다른 매체 CSV는 1행이 바로 컬럼이라 여기서 걸러진다)
# 여기서는 항목을 하나씩 받은 단독 보고서만 찾고, 지면별+검색어+요일을 함께 받은 통합 보고서는 아래에서 처리한다.
$placementFile = $null; $keywordFile = $null; $hourlyFile = $null; $regionFile = $null; $gfaFile = $null; $dowFile = $null; $convTypeFile = $null
$campaignFile = $null
foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
    $head = @(Get-Content -LiteralPath $csv.FullName -TotalCount 2 -Encoding UTF8)
    if ($head.Count -lt 2) { continue }
    # GFA를 zip이 아니라 압축을 푼 CSV로 넣은 경우: 제목 행 없이 1행이 컬럼이다
    if (Test-GfaHeader $head[0]) { if (-not $gfaFile) { $gfaFile = $csv.FullName }; continue }
    if ($head[0] -notmatch '\d{4}\.\d{2}\.\d{2}\.~\d{4}\.\d{2}\.\d{2}\.') { continue }
    $cols = @($head[1] -split ',' | ForEach-Object { $_.Trim().Trim('"') })
    # 전환유형 보고서는 컬럼이 통합 보고서와 비슷하지만 노출수, 비용이 없어서 합계에 쓰면 안 된다. 따로 읽는다.
    if ($cols -contains '전환 유형') { if (-not $convTypeFile) { $convTypeFile = $csv.FullName }; continue }
    $isTerm = ($cols -contains '검색어') -or ($cols -contains '키워드')
    if (($cols -contains '매체이름') -and $isTerm -and ($cols -contains '요일별')) { continue }  # 통합 보고서
    if ($cols -contains '시간대별') { if (-not $hourlyFile) { $hourlyFile = $csv.FullName }; continue }
    if ($cols -contains '지역') { if (-not $regionFile) { $regionFile = $csv.FullName }; continue }
    if (($cols -contains '매체이름') -and -not $placementFile) { $placementFile = $csv.FullName }
    if ($isTerm -and -not $keywordFile) { $keywordFile = $csv.FullName }
    # 지면(매체이름)·검색어 없이 캠페인 단위로 받은 보고서. 지면별 보고서가 없을 때만 그 자리에 쓴다 (아래)
    if (($cols -contains '캠페인') -and ($cols -contains '노출수') -and -not ($cols -contains '매체이름') -and -not $isTerm -and -not $campaignFile) { $campaignFile = $csv.FullName }
    if ((($cols -contains '요일별') -or ($cols -contains '일별') -or ($cols -contains 'PC/모바일 매체')) -and -not $dowFile) { $dowFile = $csv.FullName }
}

if (-not $gfaFile) {
    # GFA는 파일명이 아니라 zip 안 CSV의 헤더로 식별한다 (Test-GfaHeader).
    foreach ($csv in (Expand-NaverZips -Dir $InputDir)) {
        $headerLine = Get-Content -LiteralPath $csv.FullName -TotalCount 1 -Encoding UTF8
        if (Test-GfaHeader $headerLine) {
            $gfaFile = $csv.FullName
            break
        }
    }
}

# GFA '상품 성과 보고서'도 zip으로 받으며, '네이버쇼핑 상품 ID' 컬럼으로 식별한다. GFA 섹션 안의 상품별 표에만 쓴다.
$gfaProductFile = $null
foreach ($csv in (Expand-NaverZips -Dir $InputDir)) {
    $headerLine = Get-Content -LiteralPath $csv.FullName -TotalCount 1 -Encoding UTF8
    if ($headerLine -match '네이버쇼핑 상품 ID') { $gfaProductFile = $csv.FullName; break }
}

if (-not ($placementFile -and $keywordFile -and $dowFile)) {
    # 요일/일자/기기/검색어/지면별을 한 파일로 합쳐 받은 '다차원 보고서'는 파일명이 매번 다를 수 있어,
    # 헤더에 세 리포트의 식별 컬럼('매체이름'+'검색어' 또는 '키워드'+'요일별')이 모두 있는지로 판별한다.
    # 지면별·검색어·요일_일자_기기별 중 단독 보고서로 못 찾은 항목만 이 통합 파일로 채운다.
    foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
        $headerLine = Get-Content -LiteralPath $csv.FullName -TotalCount 2 -Encoding UTF8 | Select-Object -Last 1
        if ($headerLine -match '매체이름' -and $headerLine -match '검색어|키워드' -and $headerLine -match '요일별' -and $headerLine -notmatch '전환 유형') {
            if (-not $placementFile) { $placementFile = $csv.FullName }
            if (-not $keywordFile)   { $keywordFile = $csv.FullName }
            if (-not $dowFile)       { $dowFile = $csv.FullName }
            break
        }
    }
}

# 네이버는 한 보고서에 '검색어'와 '키워드'를 함께 넣을 수 없어, 둘 다 보려면 통합 보고서를 두 번 받게 된다.
# 두 파일이 모두 있으면 검색어 파일을 기본으로 쓰고, 키워드 파일은 '키워드별 성과' 표에만 추가로 쓴다.
# (두 파일 모두 파일명에 '검색어'가 들어가므로 헤더의 컬럼명으로 구분한다)
$keywordFile2 = $null
$termCsv = $null; $kwdCsv = $null
foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
    $cols = @((Get-Content -LiteralPath $csv.FullName -TotalCount 2 -Encoding UTF8 | Select-Object -Last 1) -split ',')
    if ($cols -contains '전환 유형') { continue }
    if ($cols -contains '검색어') { if (-not $termCsv) { $termCsv = $csv.FullName } }
    elseif ($cols -contains '키워드' -and $cols -contains '캠페인유형') { if (-not $kwdCsv) { $kwdCsv = $csv.FullName } }
}
if ($termCsv -and $kwdCsv) {
    $keywordFile = $termCsv; $keywordFile2 = $kwdCsv
    # 검색어 보고서는 실제 검색어가 있는 노출만 담아서, 검색어 없이 쇼핑 지면에 나간 쇼핑검색 비용이 빠진다
    # (예: 엘비헬스케어 2026-09-16~10-15, 검색어 파일 339만 원 / 키워드·시간대별 파일 503만 원).
    # 그래서 두 파일이 다 있으면 요약 · 지면별 · 요일/일자/기기별 합계는 빠짐없는 키워드 파일로 계산한다.
    if ($placementFile -eq $termCsv) { $placementFile = $kwdCsv }
    if ($dowFile -eq $termCsv) { $dowFile = $kwdCsv }
}

# 캠페인 단위 보고서만 받은 경우(2026-11-03 사용자: 캠페인별 월간 보고서가 대량 필요): 지면별 자리에 넣어
# 요약 합계, 캠페인 유형별, 캠페인별, 일별(일별 컬럼이 있으면)을 만든다. 지면 컬럼이 없으니 지면별 표만 숨긴다.
# 지면별 보고서가 있으면 그쪽을 쓴다 (같은 기간의 두 파일을 함께 넣어도 합계가 두 번 더해지지 않게).
$naverCampaignOnly = $false
if (-not $placementFile -and $campaignFile) { $placementFile = $campaignFile; $naverCampaignOnly = $true }

# 같은 종류의 검색어(또는 키워드) 단독 보고서를 여러 개 넣으면 모두 합쳐 한 보고서로 만든다 (2.033, 사용자: "1년치로 뽑고 싶다".
# 네이버는 한 번에 받을 수 있는 기간이 짧아 달마다, 반달마다 나눠 받는다). 지면 컬럼(매체이름)이 있는 통합 보고서는 지면별 · 요일별에도
# 쓰여서 예전처럼 첫 파일만 쓴다. 기간이 겹치면 같은 광고비가 두 번 더해지므로 멈추고 알려 준다. 빠진 기간은 보고서 주석에 적는다.
$keywordFiles = @()
$keywordGapNote = $null
$keywordCombinedPeriod = $null
if ($keywordFile) {
    $keywordFiles = @($keywordFile)
    $kwHead0 = @([NaverCsv]::Header($keywordFile))
    $kwDimCol = if ($kwHead0 -contains '검색어') { '검색어' } else { '키워드' }
    if ($kwHead0 -notcontains '매체이름') {
        foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
            if ($csv.FullName -eq $keywordFile -or $csv.FullName -eq $keywordFile2) { continue }
            if (-not (Get-NaverPeriod -Path $csv.FullName)) { continue }   # 네이버 보고서(1행 제목 + 기간)만
            $h = @([NaverCsv]::Header($csv.FullName))
            if (($h -contains '전환 유형') -or ($h -contains '매체이름') -or ($h -notcontains $kwDimCol)) { continue }
            if ($kwDimCol -eq '키워드' -and ($h -contains '검색어')) { continue }
            $keywordFiles += $csv.FullName
        }
    }
    if ($keywordFiles.Count -gt 1) {
        $spans = @(foreach ($f in $keywordFiles) {
            $p = Get-NaverPeriod -Path $f
            if ($p -match '(\d{4})\.(\d{2})\.(\d{2})\.~(\d{4})\.(\d{2})\.(\d{2})\.') {
                [PSCustomObject]@{
                    Name  = [System.IO.Path]::GetFileName($f)
                    Text  = $p
                    Start = [datetime]::new([int]$matches[1], [int]$matches[2], [int]$matches[3])
                    End   = [datetime]::new([int]$matches[4], [int]$matches[5], [int]$matches[6])
                }
            }
        })
        $spans = @($spans | Sort-Object -Property Start)
        $gaps = @()
        for ($i = 1; $i -lt $spans.Count; $i++) {
            $prev = $spans[$i - 1]; $cur = $spans[$i]
            if ($cur.Start -le $prev.End) {
                Stop-WithMessage "기간이 겹치는 네이버 $kwDimCol 보고서가 있습니다.`n$($prev.Name) ($($prev.Text))`n$($cur.Name) ($($cur.Text))`n같은 기간이 두 번 더해지지 않게, 겹치는 파일을 빼고 다시 생성해 주세요."
            }
            if ($cur.Start -gt $prev.End.AddDays(1)) {
                $gaps += "$($prev.End.AddDays(1).ToString('yyyy.MM.dd.'))~$($cur.Start.AddDays(-1).ToString('yyyy.MM.dd.'))"
            }
        }
        if ($spans.Count -gt 0) { $keywordCombinedPeriod = "$($spans[0].Start.ToString('yyyy.MM.dd.'))~$(($spans | Sort-Object -Property End)[-1].End.ToString('yyyy.MM.dd.'))" }
        $keywordGapNote = "$kwDimCol 보고서 $($keywordFiles.Count)개를 합쳤습니다." + $(if ($gaps.Count) { " 그 사이에 빠진 기간이 있습니다: $($gaps -join ', ')." } else { '' })
    }
}

$hasPlacement = [bool]$placementFile
$hasKeyword   = [bool]$keywordFile
$hasHourly    = [bool]$hourlyFile
$hasRegion    = [bool]$regionFile
$gfaAvailable = [bool]$gfaFile
$hasDow       = [bool]$dowFile

# ---------- 지마켓 (상품별/키워드별) / 옥션 (AI매출업/파워클릭 상품별/키워드별) xlsx ----------
# 지마켓은 xlsx 내부의 '리포트 종류' 값으로, 옥션은 헤더의 고유 컬럼명으로 구분한다.
# 지마켓 광고센터가 새로 바뀐 뒤에도 예전 광고센터(옥션 광고센터)에서 지마켓 광고를 함께 운영한다(2.038).
# 그곳에서 받은 지마켓 리포트는 옥션 리포트와 컬럼이 같아 내용으로는 구별할 수 없어서, 파일 이름에 '구지마켓'(띄어 써도 됨)이 있으면 지마켓(구 광고센터)으로 본다.
# Get-ReportKinds.ps1, Build-ExcelReport.ps1도 같은 규칙. 규칙을 바꾸면 함께 고칠 것.
$gmarketItem = $null
$gmarketKeyword = $null
$gmarketCampaign = $null
$auctionAI = $null
$auctionGoods = $null
$auctionKeyword = $null
$auctionPC = $null
$auctionAIDaily = $null
$gmOldAI = $null
$gmOldGoods = $null
$gmOldKeyword = $null
$gmOldPC = $null
$gmOldAIDaily = $null
$oldCenterSeen = @{}   # '옥션|AI매출업_상품별' 같은 자리 → 파일 이름. 같은 자리에 파일이 둘이면 하나가 말없이 빠지므로 멈춘다.
foreach ($xlsx in (Get-ChildItem -LiteralPath $InputDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue)) {
    $matched = $false
    try {
        $parsed = Get-GmarketXlsx -Path $xlsx.FullName
        if ($parsed.ReportType -match '상품') { if (-not $gmarketItem) { $gmarketItem = $parsed }; $matched = $true }
        elseif ($parsed.ReportType -match '키워드') { if (-not $gmarketKeyword) { $gmarketKeyword = $parsed }; $matched = $true }
        elseif ($parsed.ReportType -match '캠페인') { if (-not $gmarketCampaign) { $gmarketCampaign = $parsed }; $matched = $true }   # 캠페인별 리포트 (2.039)
    } catch { }
    if ($matched) { continue }
    $aParsed = $null
    try { $aParsed = Get-AuctionXlsx -Path $xlsx.FullName } catch { }
    if (-not $aParsed) { continue }
    $isOldGm = (($xlsx.BaseName -replace '\s', '') -match '구지마켓')
    $slot = $(if ($isOldGm) { '구지마켓' } else { '옥션' }) + '|' + $aParsed.ReportType
    if ($oldCenterSeen.ContainsKey($slot)) {
        $kindName = ($slot -replace '\|', ' ') -replace '_', ' '
        Stop-WithMessage ("같은 종류의 리포트가 두 개 있습니다 ($kindName):`n$($oldCenterSeen[$slot])`n$($xlsx.Name)`n`n" +
            "옥션 광고센터에서 받은 지마켓 리포트라면 파일 이름에 '구지마켓'을 넣어 주세요. 같은 파일을 두 번 넣었다면 하나를 빼 주세요.")
    }
    $oldCenterSeen[$slot] = $xlsx.Name
    switch ($slot) {
        '옥션|AI매출업_상품별'     { $auctionAI = $aParsed }
        '옥션|파워클릭_상품별'     { $auctionGoods = $aParsed }
        '옥션|파워클릭_키워드별'   { $auctionKeyword = $aParsed }
        '옥션|파워클릭_기간별'     { $auctionPC = $aParsed }
        '옥션|AI매출업_일별'       { $auctionAIDaily = $aParsed }
        '구지마켓|AI매출업_상품별'   { $gmOldAI = $aParsed }
        '구지마켓|파워클릭_상품별'   { $gmOldGoods = $aParsed }
        '구지마켓|파워클릭_키워드별' { $gmOldKeyword = $aParsed }
        '구지마켓|파워클릭_기간별'   { $gmOldPC = $aParsed }
        '구지마켓|AI매출업_일별'     { $gmOldAIDaily = $aParsed }
    }
}
$hasGmarketItem = [bool]$gmarketItem
$hasGmarketKeyword = [bool]$gmarketKeyword
$hasGmarketCampaign = [bool]$gmarketCampaign
$gmNewAvailable = ($hasGmarketItem -or $hasGmarketKeyword -or $hasGmarketCampaign)
$gmOldAvailable = [bool]($gmOldAI -or $gmOldGoods -or $gmOldKeyword -or $gmOldPC -or $gmOldAIDaily)
$gmarketAvailable = ($gmNewAvailable -or $gmOldAvailable)

$hasAuctionAI = [bool]$auctionAI
$hasAuctionGoods = [bool]$auctionGoods
$hasAuctionKeyword = [bool]$auctionKeyword
$auctionAvailable = ($hasAuctionAI -or $hasAuctionGoods -or $hasAuctionKeyword -or [bool]$auctionPC -or [bool]$auctionAIDaily)

# ---------- 11번가 (광고상품별 일별 리포트 CSV) ----------
# 네이버 CSV와 확장자가 같아 헤더('날짜' + '총 전환 금액')로 구분한다.
# 같은 광고상품 파일이 중복으로 들어있으면(재다운로드 등) 이중 합산되지 않도록 첫 파일만 쓴다.
$elevenstFiles = @()
foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
    try {
        $sParsed = Get-ElevenstCsv -File $csv
        if ($sParsed -and -not ($elevenstFiles | Where-Object { $_.AdType -eq $sParsed.AdType })) { $elevenstFiles += $sParsed }
    } catch { }
}
$elevenstAvailable = ($elevenstFiles.Count -gt 0)

# ---------- 구글, 메타 (캠페인별 일별 리포트 CSV) ----------
# 이것도 파일 이름이 아니라 컬럼으로 알아본다. 같은 매체 파일이 여러 개면 이중 합산되지 않도록 첫 파일만 쓴다.
$googleReport = $null
$metaReport = $null
foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
    if (-not $googleReport) { $googleReport = Get-GoogleCsv -File $csv }
    if (-not $metaReport) { $metaReport = Get-MetaCsv -File $csv }
}
$googleAvailable = [bool]$googleReport
$metaAvailable = [bool]$metaReport

# ---------- 오늘의집 (보고서 CSV) ----------
# 파일 이름이 아니라 1행 컬럼으로 알아본다. 오늘의집 보고서에는 모두 '집행 광고비'와 '광고 전환 매출'(카테고리, 상품(전체) 보고서는 '전환 매출')이 있고,
# 다른 매체 보고서에는 없다. 종류는 아래 컬럼으로 나눈다 (2026-11-04 받은 2026년 10월 파일 기준):
#   검색어      '키워드'가 있음 ('캠페인_키워드타겟팅통합보고서'). 오늘의집은 실제 검색어를 '키워드' 컬럼에 준다. 캠페인마다 받고 날짜는 없다
#   상품(전체)   '캠페인 명/ID'가 있음. 상품마다 한 줄, 등록된 캠페인이 '이름(ID),이름(ID)'로 한 칸에 들어 있다. 날짜 없음
#   상품        '상품 ID'가 있음 ('캠페인_일별_상품별_성과'). 캠페인마다 따로 받아야 하고 일별로만 받아진다. 여러 개 넣으면 모두 쓴다
#   카테고리     '카테고리'가 있음. 날짜 없음
#   캠페인 일별  '캠페인 ID' + '날짜'
#   캠페인      '캠페인 ID'만 (기간 합계)
#   계정 일별    '날짜'만
$ohKeywordFiles = New-Object System.Collections.Generic.List[string]
$ohProductFiles = New-Object System.Collections.Generic.List[string]
$ohProductTotalFiles = New-Object System.Collections.Generic.List[string]
$ohCategoryFiles = New-Object System.Collections.Generic.List[string]
$ohCampaignDailyFile = $null; $ohCampaignFile = $null; $ohDailyFile = $null
$ohPeriod = $null
foreach ($csv in (Get-ChildItem -LiteralPath $InputDir -Filter "*.csv" -File -ErrorAction SilentlyContinue)) {
    $head = Read-HeadLines -Path $csv.FullName -Count 1
    if (-not $head.Count) { continue }
    $cols = @($head[0].Split(',') | ForEach-Object { $_.Trim().Trim([char]0xFEFF).Trim('"') })   # BOM이 두 번 있는 파일이 있다 (카테고리, 상품(전체))
    if (-not (($cols -contains '집행 광고비') -and (($cols -contains '광고 전환 매출') -or ($cols -contains '전환 매출')))) { continue }
    # 기간은 파일 이름 끝('..._2026-10-01_2026-10-31.csv')에서 읽는다. 집행이 없는 날이 앞뒤에 있어도 받은 기간 그대로 나온다
    if (-not $ohPeriod -and $csv.Name -match '(\d{4})-(\d{2})-(\d{2})_(\d{4})-(\d{2})-(\d{2})') {
        $ohPeriod = "$($matches[1]).$($matches[2]).$($matches[3]).~$($matches[4]).$($matches[5]).$($matches[6])."
    }
    if ($cols -contains '키워드') { $ohKeywordFiles.Add($csv.FullName) }
    elseif ($cols -contains '캠페인 명/ID') { $ohProductTotalFiles.Add($csv.FullName) }
    elseif ($cols -contains '상품 ID') { $ohProductFiles.Add($csv.FullName) }
    elseif ($cols -contains '카테고리') { $ohCategoryFiles.Add($csv.FullName) }
    elseif ($cols -contains '캠페인 ID') {
        if ($cols -contains '날짜') { if (-not $ohCampaignDailyFile) { $ohCampaignDailyFile = $csv.FullName } }
        elseif (-not $ohCampaignFile) { $ohCampaignFile = $csv.FullName }
    }
    elseif ($cols -contains '날짜') { if (-not $ohDailyFile) { $ohDailyFile = $csv.FullName } }
}
$ohouseAvailable = ($ohKeywordFiles.Count -gt 0) -or ($ohProductFiles.Count -gt 0) -or ($ohProductTotalFiles.Count -gt 0) -or ($ohCategoryFiles.Count -gt 0) -or [bool]$ohCampaignDailyFile -or [bool]$ohCampaignFile -or [bool]$ohDailyFile

if (-not ($hasPlacement -or $hasKeyword -or $hasHourly -or $hasRegion -or $gfaAvailable -or $hasDow -or $gmarketAvailable -or $auctionAvailable -or $elevenstAvailable -or $googleAvailable -or $metaAvailable -or $ohouseAvailable)) {
    throw "'$sourceInputDir' 폴더에서 인식 가능한 리포트 파일을 하나도 찾지 못했습니다 (네이버: 지면별/캠페인별/검색어/시간대별/지역/GFA/요일_일자_기기별, 지마켓: 캠페인별/상품별/키워드별, 옥션: AI매출업 상품별/일별, 파워클릭 상품별/키워드별/기간별, 11번가: 광고상품별 일별, 구글, 메타, 오늘의집: 캠페인/상품/카테고리/검색어)."
}

# ---------- 지면별 ----------
$byPlacement = @()
$byCampaignType = @()
$saOverview = $null
$placementPeriod = $null
$naverChannel = New-Channel

if ($hasPlacement) {
    $placementPeriod = Get-NaverPeriod -Path $placementFile
    $byCampaignTypeAgg = @{}
    $saAgg = New-Agg
    # 통합(다차원) 보고서는 같은 지면이 검색어·일자·기기별로 수만 행에 나뉘어 있으므로 지면 단위로 합친다.
    # (지면별 단독 보고서는 지면당 1행이라 결과가 그대로다.) 먼저 C#으로 지면 · 일자 단위까지 묶고, 아래는 묶인 결과만 돈다.
    $placementAgg = @{}
    $placementMeta = @{}
    $plCols = @('캠페인유형', '캠페인', '광고그룹', '검색/콘텐츠 매체', '매체이름', '일별')
    $plMap = [OhouseCsv]::NewMap()
    [void][NaverCsv]::Aggregate($placementFile, $plCols, $plCols, $plMap)

    foreach ($r in $plMap.Values) {
        $imp = $r.imp; $clk = $r.clk; $cost = $r.cost; $conv = $r.conv; $rev = $r.rev

        $pKey = "$($r.'캠페인유형')|$($r.'캠페인')|$($r.'광고그룹')|$($r.'검색/콘텐츠 매체')|$($r.'매체이름')"
        if (-not $placementAgg.ContainsKey($pKey)) {
            $placementAgg[$pKey] = New-Agg
            $placementMeta[$pKey] = $r
        }
        [AdFast]::AddAgg($placementAgg[$pKey], $imp, $clk, $cost, $conv, $rev)

        $ct = $r.'캠페인유형'
        [AdFast]::AddChannelRow($naverChannel, $ct, $r.'캠페인', $r.'일별', $imp, $clk, $cost, $conv, $rev)
        if (-not $byCampaignTypeAgg.ContainsKey($ct)) { $byCampaignTypeAgg[$ct] = New-Agg }
        [AdFast]::AddAgg($byCampaignTypeAgg[$ct], $imp, $clk, $cost, $conv, $rev)
        [AdFast]::AddAgg($saAgg, $imp, $clk, $cost, $conv, $rev)
    }

    $byPlacement = @(foreach ($pKey in $placementAgg.Keys) {
        $r = $placementMeta[$pKey]
        $m = Metrics-FromAgg $placementAgg[$pKey]
        [PSCustomObject]@{
            campaignType = $r.'캠페인유형'
            campaign     = $r.'캠페인'
            adGroup      = $r.'광고그룹'
            network      = $r.'검색/콘텐츠 매체'
            media        = $r.'매체이름'
            impressions  = $m.impressions
            clicks       = $m.clicks
            ctr          = $m.ctr
            cost         = $m.cost
            cpc          = $m.cpc
            conversions  = $m.conversions
            revenue      = $m.revenue
            roas         = $m.roas
        }
    })

    $byPlacement = @($byPlacement | Sort-Object -Property cost -Descending)

    $byCampaignType = @(foreach ($ct in $byCampaignTypeAgg.Keys) {
        $m = Metrics-FromAgg $byCampaignTypeAgg[$ct]
        [PSCustomObject]@{
            type        = $ct
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })

    $saOverview = Metrics-FromAgg $saAgg
}

# ---------- 검색어 (키워드별 집계) ----------
$allKeywords = @()
$topKeywordsByCost = @()
$topKeywordsByConv = @()
$keywordTotalCost = $null
$keywordUniqueCount = 0
$zeroConvSpend = 0
$keywordPeriod = $null
$keywordDimension = '검색어'
$keywordShownCount = 0
# 검색어 · 키워드 표 줄(JSON 글자). 보고서 HTML의 report-rows 칸에 따로 넣는다 (report-data는 작게 두어 화면과 월별 추이가 빨리 읽게)
$keywordRowsJson = $null
$keyword2RowsJson = $null

function Get-KeywordAgg {
    # 검색어(또는 키워드) 보고서를 C#으로 읽어 묶는다. $Column은 '검색어' 또는 '키워드'. $Path가 여러 개면(달마다 나눠 받은 1년치 등) 모두 같은 표에 더한다.
    # 표 줄은 캠페인유형 · 캠페인 · 광고그룹 · 검색어 단위다 (2.032, 사용자: "캠페인, 광고그룹, 키워드(검색어)가 나오게").
    # 보고서에 캠페인, 광고그룹 컬럼이 없으면 그 칸은 비고, 화면에서 빈 컬럼은 숨긴다.
    # 검색어 수, 전환 없는 검색어 비용은 예전처럼 (검색어, 캠페인유형) 단위로 센다 (광고그룹으로 나누면 같은 검색어가 여러 번 세진다).
    param([string[]]$Path, [string]$Column)
    $cols = @('캠페인유형', '캠페인', '광고그룹', $Column)
    $fine = [OhouseCsv]::NewMap()
    foreach ($one in $Path) { [void][NaverCsv]::Aggregate($one, $cols, $cols, $fine) }
    $byKw = [NaverCsv]::Rollup($fine, @($Column, '캠페인유형'))
    $tot = [OhouseCsv]::Totals($byKw)
    # 상위 몇 개로 자르지 않고, 노출이 1회 이상 있었던 줄을 모두 싣는다 (비용순, 2.036). 줄은 JSON 글자로 받는다 (ShownRowsJson 설명)
    $count = New-Object 'int[]' 1
    $json = [NaverCsv]::ShownRowsJson($fine, $cols, @('campaignType', 'campaign', 'adGroup', 'keyword'), $count)
    [PSCustomObject]@{
        RowsJson = $json
        RowCount = $count[0]
        Unique  = [OhouseCsv]::DistinctMeta($byKw, $Column)
        Totals  = $tot   # 노출, 클릭, 비용, 전환, 매출, 전환 없는 검색어 비용
    }
}

$keywordTable2 = $null
$keywordTotals = $null
if ($hasKeyword) {
    $keywordPeriod = Get-NaverPeriod -Path $keywordFile
    # 통합 보고서를 '검색어' 대신 '키워드' 항목으로 내려받은 경우 키워드 기준으로 같은 표를 만든다.
    $kwHeader = [NaverCsv]::Header($keywordFile)
    if (($kwHeader -notcontains '검색어') -and ($kwHeader -contains '키워드')) { $keywordDimension = '키워드' }
    $kwResult = Get-KeywordAgg -Path $keywordFiles -Column $keywordDimension
    if ($keywordCombinedPeriod) { $keywordPeriod = $keywordCombinedPeriod }   # 여러 파일을 합쳤으면 첫 파일 시작 ~ 마지막 파일 끝
    $keywordRowsJson = $kwResult.RowsJson
    $keywordShownCount = $kwResult.RowCount
    $topKeywordsByConv = @()   # 화면에서 쓰지 않는다 (2.032부터 비움)
    $keywordTotals = $kwResult.Totals
    $zeroConvSpend = $keywordTotals[5]
    $keywordUniqueCount = $kwResult.Unique
    $keywordTotalCost = $keywordTotals[2]

    if ($keywordFile2) {
        $kw2 = Get-KeywordAgg -Path $keywordFile2 -Column '키워드'
        $keyword2RowsJson = $kw2.RowsJson
        $keywordTable2 = [PSCustomObject]@{ dimension = '키워드'; rows = @(); totalUniqueKeywords = $kw2.Unique }   # rows는 report-rows 칸에서 채운다
    }
}

# ---------- 시간대별 ----------
$hourlyRows = @()
$hourlyPeriod = $null
$hourlyTotalCost = $null
$hourlyTotalImp = $null

if ($hasHourly) {
    $hourly = Get-NaverCsv -Path $hourlyFile
    $hourlyPeriod = $hourly.Period
    $hourlyRows = @(foreach ($r in $hourly.Rows) {
        $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
        $conv = ToNum $r.'구매완료 전환수'; $rev = ToNum $r.'구매완료 전환매출액(원)'
        $m = New-Metrics -Impressions $imp -Clicks $clk -Cost $cost -Conversions $conv -Revenue $rev
        [PSCustomObject]@{
            hour        = $r.'시간대별'
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $hourlyTotalCost = ($hourlyRows | Measure-Object -Property cost -Sum).Sum
    $hourlyTotalImp  = ($hourlyRows | Measure-Object -Property impressions -Sum).Sum
}

# ---------- 지역별 (파워링크 전용) ----------
$regionRows = @()
$regionPeriod = $null

if ($hasRegion) {
    $region = Get-NaverCsv -Path $regionFile
    $regionPeriod = $region.Period
    $hasAvgRank = $false
    if ($region.Rows.Count -gt 0) {
        $hasAvgRank = [bool]($region.Rows[0].PSObject.Properties.Name -contains '평균노출순위')
    }
    $regionRows = @(foreach ($r in $region.Rows) {
        $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
        $conv = ToNum $r.'구매완료 전환수'; $rev = ToNum $r.'구매완료 전환매출액(원)'
        $m = New-Metrics -Impressions $imp -Clicks $clk -Cost $cost -Conversions $conv -Revenue $rev
        $name = $r.'지역'
        if ($name -eq '-') { $name = '지역 미상' }
        [PSCustomObject]@{
            region      = $name
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc
            avgRank     = if ($hasAvgRank) { ToNum $r.'평균노출순위' } else { $null }
            conversions = $m.conversions; revenue = $m.revenue; roas = $m.roas
        }
    })
    $regionRows = @($regionRows | Sort-Object -Property cost -Descending)
}

# ---------- GFA (있는 경우에만) ----------
$gfaOverview = $null
$gfaByCampaign = @()
$gfaBySegment = @()
$gfaByDate = @()
$gfaPeriod = $null
$gfaPeriodDiffNote = $null

if ($gfaAvailable) {
    $gfa = Get-GfaCsv -Path $gfaFile
    $gfaPeriod = $gfa.Period

    $gfaTotalAgg = New-Agg
    $gfaCampaignAgg = @{}
    $gfaSegmentAgg = @{}
    $gfaDateAgg = @{}
    $gfaChannel = New-Channel

    foreach ($r in $gfa.Rows) {
        $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
        $conv = ToNum $r.'구매완료 수'; $rev = ToNum $r.'구매완료 전환매출액'

        Add-Agg -Agg $gfaTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        [AdFast]::AddChannelRow($gfaChannel, $r.'캠페인 목적', $r.'캠페인 이름', $r.'기간', $imp, $clk, $cost, $conv, $rev)

        $cname = $r.'캠페인 이름'
        if (-not $gfaCampaignAgg.ContainsKey($cname)) {
            $objective = $r.'캠페인 목적'
            if (-not $objective) { $objective = '' }
            $gfaCampaignAgg[$cname] = @{ agg = (New-Agg); objective = $objective; id = $r.'캠페인 ID' }
        }
        Add-Agg -Agg $gfaCampaignAgg[$cname].agg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev

        # '연령 및 성별' 없이 캠페인 단위로만 받은 보고서는 연령·성별 집계를 비워 둔다 (화면에서 그 부분을 숨김)
        $seg = $r.'연령 및 성별'
        if ($seg) {
            if (-not $gfaSegmentAgg.ContainsKey($seg)) { $gfaSegmentAgg[$seg] = New-Agg }
            Add-Agg -Agg $gfaSegmentAgg[$seg] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }

        # '기간'(일자) 컬럼은 GFA 리포트 종류에 따라 없을 수 있다 (예: 캠페인×세그먼트 요약형).
        # 없으면 일별 추이를 만들 수 없으므로 집계에서 제외한다 (아래에서 byDate가 비면 화면에서 해당 섹션을 숨김).
        $d = $r.'기간'
        if ($d) {
            if (-not $gfaDateAgg.ContainsKey($d)) { $gfaDateAgg[$d] = New-Agg }
            Add-Agg -Agg $gfaDateAgg[$d] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
    }

    $gfaOverview = Metrics-FromAgg $gfaTotalAgg

    $gfaByCampaign = @(foreach ($cname in $gfaCampaignAgg.Keys) {
        $entry = $gfaCampaignAgg[$cname]
        $m = Metrics-FromAgg $entry.agg
        [PSCustomObject]@{
            campaign    = $cname
            campaignId  = $entry.id
            objective   = $entry.objective
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $gfaByCampaign = @($gfaByCampaign | Sort-Object -Property cost -Descending)

    $gfaBySegment = @(foreach ($seg in $gfaSegmentAgg.Keys) {
        $m = Metrics-FromAgg $gfaSegmentAgg[$seg]
        [PSCustomObject]@{
            segment     = $seg
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $gfaBySegment = @($gfaBySegment | Sort-Object -Property cost -Descending)

    $gfaByDate = @(foreach ($d in $gfaDateAgg.Keys) {
        $m = Metrics-FromAgg $gfaDateAgg[$d]
        [PSCustomObject]@{
            date        = $d
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $gfaByDate = @($gfaByDate | Sort-Object -Property date)
}

# ---------- 요일 / 일자 / 기기별 ----------
$byDevice = @()
$byDate = @()
$byWeekday = @()
$dowPeriod = $null

if ($hasDow) {
    $dowPeriod = Get-NaverPeriod -Path $dowFile
    $deviceAgg = @{}
    $dateAgg = @{}
    $weekdayAgg = @{}

    # 통합 보고서를 받을 때 기기/일자/요일 항목을 선택하지 않으면 해당 컬럼이 파일에 아예 없다.
    # 있는 컬럼만 집계하고, 없는 항목은 비워둔다(템플릿이 빈 차트는 숨김).
    $dowCols = @([NaverCsv]::Header($dowFile))
    $hasDevCol  = $dowCols -contains 'PC/모바일 매체'
    $hasDateCol = $dowCols -contains '일별'
    $hasWdCol   = $dowCols -contains '요일별'
    # 먼저 C#으로 기기 · 일자 · 요일 단위까지 묶는다 (수만 줄 → 수백 줄)
    $dwKeys = @('PC/모바일 매체', '일별', '요일별')
    $dwMap = [OhouseCsv]::NewMap()
    [void][NaverCsv]::Aggregate($dowFile, $dwKeys, $dwKeys, $dwMap)

    foreach ($r in $dwMap.Values) {
        $imp = $r.imp; $clk = $r.clk; $cost = $r.cost; $conv = $r.conv; $rev = $r.rev

        if ($hasDevCol) {
            $dev = $r.'PC/모바일 매체'
            if (-not $deviceAgg.ContainsKey($dev)) { $deviceAgg[$dev] = New-Agg }
            [AdFast]::AddAgg($deviceAgg[$dev], $imp, $clk, $cost, $conv, $rev)
        }

        if ($hasDateCol) {
            $dt = $r.'일별'
            if (-not $dateAgg.ContainsKey($dt)) { $dateAgg[$dt] = New-Agg }
            [AdFast]::AddAgg($dateAgg[$dt], $imp, $clk, $cost, $conv, $rev)
        }

        if ($hasWdCol) {
            $wd = $r.'요일별'
            if (-not $weekdayAgg.ContainsKey($wd)) { $weekdayAgg[$wd] = New-Agg }
            [AdFast]::AddAgg($weekdayAgg[$wd], $imp, $clk, $cost, $conv, $rev)
        }
    }

    $byDevice = @(foreach ($k in $deviceAgg.Keys) {
        $m = Metrics-FromAgg $deviceAgg[$k]
        [PSCustomObject]@{
            device      = $k
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $byDevice = @($byDevice | Sort-Object -Property cost -Descending)

    $byDate = @(foreach ($k in $dateAgg.Keys) {
        $m = Metrics-FromAgg $dateAgg[$k]
        [PSCustomObject]@{
            date        = $k
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $byDate = @($byDate | Sort-Object -Property date)

    $weekdayOrder = @('일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일')
    $byWeekday = @(foreach ($k in $weekdayOrder) {
        if ($weekdayAgg.ContainsKey($k)) {
            $m = Metrics-FromAgg $weekdayAgg[$k]
            [PSCustomObject]@{
                weekday     = $k
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        }
    })
    # 표준 요일 이름과 다른 값이 있으면(예: 표기 차이) 뒤에 추가로 붙여 데이터 누락을 방지
    foreach ($k in $weekdayAgg.Keys) {
        if ($weekdayOrder -notcontains $k) {
            $m = Metrics-FromAgg $weekdayAgg[$k]
            $byWeekday += [PSCustomObject]@{
                weekday     = $k
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        }
    }

    # 기기/일자/요일 어느 것도 없는 파일이면 이 섹션은 보여줄 게 없으므로 제외한다.
    if (-not ($hasDevCol -or $hasDateCol -or $hasWdCol)) { $hasDow = $false }
}

# ---------- 지마켓 (상품별/키워드별, 네이버와 완전히 별개 매체) ----------
$gmOverview = $null
$gmByCampaignType = @()
$gmByProduct = @()
$gmTopKeywordsByCost = @()
$gmKeywordUniqueCount = 0
$gmZeroConvSpend = 0
$gmPeriod = $null
$gmByCampaign = @()

if ($gmNewAvailable) {
    $gmPeriod = if ($gmarketCampaign) { $gmarketCampaign.Period } elseif ($gmarketItem) { $gmarketItem.Period } else { $gmarketKeyword.Period }

    # 캠페인별 리포트 (2.039, 사용자: "신 지마켓 광고센터에서 캠페인별 합계 데이터도 인식"). 보기 방식이 '합계'면 캠페인마다 한 줄,
    # '날짜별'이면 캠페인 × 날짜 줄이라 캠페인유형 + 캠페인명으로 묶는다. 캠페인 단위가 광고비 전체라서,
    # 이 파일이 있으면 요약 합계 · 캠페인 유형 비교는 이 파일로 만든다(상품별 리포트와 같은 광고비를 두 번 더하지 않게).
    if ($hasGmarketCampaign) {
        $gcAgg = [ordered]@{}
        $gcTypeAgg = @{}
        $gcTotalAgg = New-Agg
        foreach ($r in $gmarketCampaign.Rows) {
            $name = "$($r.'캠페인명')".Trim(); $ct = "$($r.'캠페인타입')".Trim()
            if (-not $name -and -not $ct) { continue }
            if ("$($r.'날짜')" -match '^\s*(합계|총계|전체)') { continue }   # 합계 줄이 있으면 뺀다(샘플에는 없음)
            $imp = ToNum $r.'노출 수'; $clk = ToNum $r.'클릭 수'; $cost = ToNum $r.'광고 비용'
            $conv = ToNum $r.'판매자 전환 수'; $rev = ToNum $r.'판매자 전환 금액'   # 상품별 리포트와 같은 기준(지마켓 '광고수익률'과 같은 계산)
            $key = "$ct`t$name"
            if (-not $gcAgg.Contains($key)) { $gcAgg[$key] = [PSCustomObject]@{ type = $ct; name = $name; agg = (New-Agg) } }
            Add-Agg -Agg $gcAgg[$key].agg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            if (-not $gcTypeAgg.ContainsKey($ct)) { $gcTypeAgg[$ct] = New-Agg }
            Add-Agg -Agg $gcTypeAgg[$ct] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $gcTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
        $gmByCampaign = @(foreach ($e in $gcAgg.Values) {
            $m = Metrics-FromAgg $e.agg
            [PSCustomObject]@{
                campaignType = $e.type
                campaign     = $e.name
                impressions  = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost         = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue      = $m.revenue; roas = $m.roas
            }
        })
        $gmByCampaign = @($gmByCampaign | Sort-Object -Property cost -Descending)
        $gmByCampaignType = @(foreach ($ct in $gcTypeAgg.Keys) {
            $m = Metrics-FromAgg $gcTypeAgg[$ct]
            [PSCustomObject]@{
                type        = $ct
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $gmByCampaignType = @($gmByCampaignType | Sort-Object -Property cost -Descending)
        $gmOverview = Metrics-FromAgg $gcTotalAgg
    }

    if ($hasGmarketItem) {
        $gmCampaignTypeAgg = @{}
        $gmTotalAgg = New-Agg
        $gmList = New-Object System.Collections.Generic.List[object]
        foreach ($r in $gmarketItem.Rows) {
            $imp = ToNum $r.'노출 수'; $clk = ToNum $r.'클릭 수'; $cost = ToNum $r.'광고 비용'
            # '판매자 전환 수/금액'을 매출·전환 기준으로 쓴다 (지마켓 자체 '광고수익률' 컬럼과 동일한 계산식).
            $conv = ToNum $r.'판매자 전환 수'; $rev = ToNum $r.'판매자 전환 금액'

            $m = New-Metrics -Impressions $imp -Clicks $clk -Cost $cost -Conversions $conv -Revenue $rev
            $gmList.Add([PSCustomObject]@{
                campaignType = $r.'캠페인타입'
                campaign     = $r.'캠페인명'
                adGroup      = $r.'그룹명'
                product      = $r.'상품명'
                impressions  = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost         = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue      = $m.revenue; roas = $m.roas
            })

            $ct = $r.'캠페인타입'
            if (-not $gmCampaignTypeAgg.ContainsKey($ct)) { $gmCampaignTypeAgg[$ct] = New-Agg }
            Add-Agg -Agg $gmCampaignTypeAgg[$ct] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $gmTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
        $gmByProduct = @($gmList | Sort-Object -Property cost -Descending)

        if (-not $hasGmarketCampaign) {   # 캠페인별 리포트가 있으면 합계 · 유형 비교는 그쪽 것을 쓴다
        $gmByCampaignType = @(foreach ($ct in $gmCampaignTypeAgg.Keys) {
            $m = Metrics-FromAgg $gmCampaignTypeAgg[$ct]
            [PSCustomObject]@{
                type        = $ct
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $gmByCampaignType = @($gmByCampaignType | Sort-Object -Property cost -Descending)
        $gmOverview = Metrics-FromAgg $gmTotalAgg
        }
    }

    if ($hasGmarketKeyword) {
        $gmKwAgg = @{}
        foreach ($r in $gmarketKeyword.Rows) {
            $k = $r.'키워드'
            $imp = ToNum $r.'노출 수'; $clk = ToNum $r.'클릭 수'; $cost = ToNum $r.'광고 비용'
            $conv = ToNum $r.'판매자 전환 수'; $rev = ToNum $r.'판매자 전환 금액'
            if (-not $gmKwAgg.ContainsKey($k)) { $gmKwAgg[$k] = New-Agg }
            Add-Agg -Agg $gmKwAgg[$k] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
        $gmAllKeywords = @(foreach ($k in $gmKwAgg.Keys) {
            $m = Metrics-FromAgg $gmKwAgg[$k]
            [PSCustomObject]@{
                keyword     = $k
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $gmTopKeywordsByCost = @($gmAllKeywords | Sort-Object -Property cost -Descending | Select-Object -First 30)
        $gmZeroConvSpend = ($gmAllKeywords | Where-Object { $_.conversions -eq 0 } | Measure-Object -Property cost -Sum).Sum
        if (-not $gmZeroConvSpend) { $gmZeroConvSpend = 0 }
        $gmKeywordUniqueCount = $gmKwAgg.Keys.Count

        if (-not $gmOverview) {
            # 상품별 파일이 없으면 키워드별 합계로 개요를 대신 계산한다.
            $gmOverview = New-Metrics -Impressions ($gmAllKeywords | Measure-Object -Property impressions -Sum).Sum -Clicks ($gmAllKeywords | Measure-Object -Property clicks -Sum).Sum -Cost ($gmAllKeywords | Measure-Object -Property cost -Sum).Sum -Conversions ($gmAllKeywords | Measure-Object -Property conversions -Sum).Sum -Revenue ($gmAllKeywords | Measure-Object -Property revenue -Sum).Sum
        }
    }
}

# ---------- 옥션 광고센터 리포트 (AI매출업 + 파워클릭 상품별/키워드별) ----------
# 옥션과 지마켓(구 광고센터, 파일 이름에 '구지마켓')이 같은 양식이라 같은 함수로 묶는다 (2.038).
function ConvertTo-AuctionDate($v) {
    # 옥션 광고센터 '날짜' 칸: 엑셀 날짜 일련번호(45658 = 2025-01-01)이거나 '2025-01-01' 같은 글자. 보고서 날짜 모양 'YYYY.MM.DD.'로 바꾼다.
    $t = "$v".Trim()
    if ($t -match '^\d{5}(\.\d+)?$') { return [DateTime]::FromOADate([double]$t).ToString('yyyy.MM.dd.') }
    if ($t -match '(\d{4})\D(\d{1,2})\D(\d{1,2})') { return '{0}.{1:D2}.{2:D2}.' -f $matches[1], [int]$matches[2], [int]$matches[3] }
    return $null
}

function Get-OldCenterData {
    param($AI, $Goods, $Keyword, $PC, $AIDaily)
    $auOverview = $null
    $auByAdType = @()
    $auByProduct = @()
    $auTopKeywordsByCost = @()
    $auKeywordUniqueCount = 0
    $auZeroConvSpend = 0
    $auHasImpressionGap = $false

    $auTotalAgg = New-Agg
    $auAdTypeAgg = @{}
    $auList = New-Object System.Collections.Generic.List[object]

    if ($AI) {
        $auHasImpressionGap = $true  # AI매출업 리포트에는 노출수 컬럼 자체가 없음
        foreach ($r in $AI.Rows) {
            $imp = 0
            $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
            # '판매자기준' 지표를 쓴다 (옥션 자체 '판매자기준-광고수익률' 컬럼과 동일한 계산식).
            $conv = ToNum $r.'판매자기준-구매수'; $rev = ToNum $r.'판매자기준-구매금액'

            $m = New-Metrics -Impressions $imp -Clicks $clk -Cost $cost -Conversions $conv -Revenue $rev
            $auList.Add([PSCustomObject]@{
                adType      = 'AI매출업'
                group       = $r.'그룹명'
                product     = $r.'상품번호'
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            })
            if (-not $auAdTypeAgg.ContainsKey('AI매출업')) { $auAdTypeAgg['AI매출업'] = New-Agg }
            Add-Agg -Agg $auAdTypeAgg['AI매출업'] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $auTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
    }

    # 일별 데이터 (2.040): 파워클릭 기간별(날짜마다 한 줄, 영역별), AI매출업 일별(날짜마다 한 줄, 노출수 없음).
    # 일별 차트에 쓰고, 같은 광고상품의 상품별 리포트가 없을 때만 합계도 이 파일로 낸다(둘 다 있으면 같은 광고비라 상품별로만 더한다).
    # 둘 다 있으면 차트의 고르기 목록이 'AI매출업', '파워클릭 (영역)'이 된다.
    $pcChannel = $null; $pcPeriod = $null
    if ($PC -or $AIDaily) {
        $pcChannel = New-Channel
        if ($AIDaily) {
            foreach ($r in $AIDaily.Rows) {
                $d = ConvertTo-AuctionDate $r.'날짜별'
                if (-not $d) { continue }   # 합계 줄 등
                $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
                $conv = ToNum $r.'판매자기준-구매수'; $rev = ToNum $r.'판매자기준-구매금액'   # AI매출업 상품별과 같은 '판매자기준'
                [AdFast]::AddChannelRow($pcChannel, 'AI매출업', 'AI매출업', $d, 0, $clk, $cost, $conv, $rev)
                if (-not $AI) {
                    $auHasImpressionGap = $true
                    if (-not $auAdTypeAgg.ContainsKey('AI매출업')) { $auAdTypeAgg['AI매출업'] = New-Agg }
                    Add-Agg -Agg $auAdTypeAgg['AI매출업'] -Imp 0 -Clk $clk -Cost $cost -Conv $conv -Rev $rev
                    Add-Agg -Agg $auTotalAgg -Imp 0 -Clk $clk -Cost $cost -Conv $conv -Rev $rev
                }
            }
        }
        if ($PC) {
            foreach ($r in $PC.Rows) {
                $d = ConvertTo-AuctionDate $r.'날짜'
                if (-not $d) { continue }
                $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
                $conv = ToNum $r.'구매수'; $rev = ToNum $r.'구매금액'
                $area = "$($r.'영역명')".Trim()
                $name = if ($AIDaily) { "파워클릭 ($area)" } else { $area }
                [AdFast]::AddChannelRow($pcChannel, '파워클릭', $name, $d, $imp, $clk, $cost, $conv, $rev)
                if (-not $Goods) {
                    if (-not $auAdTypeAgg.ContainsKey('파워클릭')) { $auAdTypeAgg['파워클릭'] = New-Agg }
                    Add-Agg -Agg $auAdTypeAgg['파워클릭'] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
                    Add-Agg -Agg $auTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
                }
            }
        }
        $days = @($pcChannel.day.Keys | Sort-Object)
        if ($days.Count) {
            $pcPeriod = "$($days[0])~$($days[-1])"
            # 이 리포트는 광고가 돈 날만 담는다(AI매출업 일별 2월 파일이 2/5부터). 한 달 안에서 15일 이상이고 1일에 시작하거나
            # 말일에 끝나면 그달 전체로 본다. 그대로 두면 25일 미만이라 'N주차' 주간 보고서 이름이 붙었다 (2026-11-06).
            $d0 = [DateTime]::ParseExact($days[0], 'yyyy.MM.dd.', $null); $d1 = [DateTime]::ParseExact($days[-1], 'yyyy.MM.dd.', $null)
            $monthEnd = $d0.AddDays(1 - $d0.Day).AddMonths(1).AddDays(-1)
            if ($d0.Year -eq $d1.Year -and $d0.Month -eq $d1.Month -and ($d1 - $d0).TotalDays -ge 14 -and ($d0.Day -eq 1 -or $d1.Date -eq $monthEnd.Date)) {
                $pcPeriod = $d0.ToString('yyyy.MM.') + '01.~' + $monthEnd.ToString('yyyy.MM.dd.')
            }
        } else { $pcChannel = $null }
    }

    if ($Goods) {
        foreach ($r in $Goods.Rows) {
            $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
            $conv = ToNum $r.'구매수'; $rev = ToNum $r.'구매금액'

            $m = New-Metrics -Impressions $imp -Clicks $clk -Cost $cost -Conversions $conv -Revenue $rev
            $auList.Add([PSCustomObject]@{
                adType      = '파워클릭'
                group       = $r.'영역명'
                product     = $r.'광고상품번호'
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            })
            if (-not $auAdTypeAgg.ContainsKey('파워클릭')) { $auAdTypeAgg['파워클릭'] = New-Agg }
            Add-Agg -Agg $auAdTypeAgg['파워클릭'] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $auTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
    }

    $auByProduct = @($auList | Sort-Object -Property cost -Descending)

    $auByAdType = @(foreach ($t in $auAdTypeAgg.Keys) {
        $m = Metrics-FromAgg $auAdTypeAgg[$t]
        [PSCustomObject]@{
            type        = $t
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $auByAdType = @($auByAdType | Sort-Object -Property cost -Descending)

    if ($AI -or $Goods -or $pcChannel) {
        $auOverview = Metrics-FromAgg $auTotalAgg
    }

    if ($Keyword) {
        $auKwAgg = @{}
        foreach ($r in $Keyword.Rows) {
            $k = $r.'키워드'
            $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
            $conv = ToNum $r.'구매수'; $rev = ToNum $r.'구매금액'
            if (-not $auKwAgg.ContainsKey($k)) { $auKwAgg[$k] = New-Agg }
            Add-Agg -Agg $auKwAgg[$k] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
        $auAllKeywords = @(foreach ($k in $auKwAgg.Keys) {
            $m = Metrics-FromAgg $auKwAgg[$k]
            [PSCustomObject]@{
                keyword     = $k
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $auTopKeywordsByCost = @($auAllKeywords | Sort-Object -Property cost -Descending | Select-Object -First 30)
        $auZeroConvSpend = ($auAllKeywords | Where-Object { $_.conversions -eq 0 } | Measure-Object -Property cost -Sum).Sum
        if (-not $auZeroConvSpend) { $auZeroConvSpend = 0 }
        $auKeywordUniqueCount = $auKwAgg.Keys.Count

        if (-not $auOverview) {
            # 상품별 파일이 하나도 없으면 키워드별 합계로 개요를 대신 계산한다.
            $auOverview = New-Metrics -Impressions ($auAllKeywords | Measure-Object -Property impressions -Sum).Sum -Clicks ($auAllKeywords | Measure-Object -Property clicks -Sum).Sum -Cost ($auAllKeywords | Measure-Object -Property cost -Sum).Sum -Conversions ($auAllKeywords | Measure-Object -Property conversions -Sum).Sum -Revenue ($auAllKeywords | Measure-Object -Property revenue -Sum).Sum
        }
    }

    [PSCustomObject]@{
        overview           = $auOverview
        byAdType           = $auByAdType
        byProduct          = $auByProduct
        topKeywordsByCost  = $auTopKeywordsByCost
        keywordUniqueCount = $auKeywordUniqueCount
        zeroConvSpend      = [int64]$auZeroConvSpend
        hasImpressionGap   = $auHasImpressionGap
        pcChannel          = $pcChannel   # 파워클릭 기간별 일별 (아래에서 ConvertTo-ChannelData로 바꿈)
        pcPeriod           = $pcPeriod
        # 상품별 리포트가 있으면 합계는 그쪽이 내서 일별 합과 다를 수 있다 → 'goods'. 일별 파일만으로 합계를 냈으면 'period'
        pcSource           = $(if ($pcChannel -and ($AI -or $Goods)) { 'goods' } elseif ($pcChannel) { 'period' } else { $null })
        dailyKinds         = @(@($(if ($AIDaily) { 'AI매출업 일별' }), $(if ($PC) { '파워클릭 기간별' })) | Where-Object { $_ })
    }
}

# ---------- 옥션 (네이버·지마켓과 완전히 별개 매체) ----------
$auOverview = $null
$auByAdType = @()
$auByProduct = @()
$auTopKeywordsByCost = @()
$auKeywordUniqueCount = 0
$auZeroConvSpend = 0
$auHasImpressionGap = $false
if ($auctionAvailable) {
    $au = Get-OldCenterData -AI $auctionAI -Goods $auctionGoods -Keyword $auctionKeyword -PC $auctionPC -AIDaily $auctionAIDaily
    $auOverview = $au.overview; $auByAdType = $au.byAdType; $auByProduct = $au.byProduct
    $auTopKeywordsByCost = $au.topKeywordsByCost; $auKeywordUniqueCount = $au.keywordUniqueCount
    $auZeroConvSpend = $au.zeroConvSpend; $auHasImpressionGap = $au.hasImpressionGap
}

# ---------- 지마켓 구 광고센터 (옥션 광고센터에서 받은 지마켓 리포트, 파일 이름에 '구지마켓') ----------
# 신 광고센터와 같은 지마켓 탭에 넣는다: 요약 합계 · 캠페인 유형 비교는 두 센터를 더하고, 상품별 · 키워드 표는 센터별로 따로 둔다
# (구 광고센터 상품은 상품번호만 있고, 지표 이름도 달라 한 표로 섞지 않는다). 두 센터는 서로 다른 캠페인을 운영한다고 보고 더한다.
$gmOld = $null
if ($gmOldAvailable) {
    $gmOld = Get-OldCenterData -AI $gmOldAI -Goods $gmOldGoods -Keyword $gmOldKeyword -PC $gmOldPC -AIDaily $gmOldAIDaily
    $oldTypes = @(foreach ($t in $gmOld.byAdType) {
        $c = $t.PSObject.Copy(); $c.type = "구 광고센터 $($t.type)"; $c
    })
    if ($gmOverview -and $gmOld.overview) {
        $a = $gmOverview; $b = $gmOld.overview
        $gmOverview = New-Metrics -Impressions ($a.impressions + $b.impressions) -Clicks ($a.clicks + $b.clicks) -Cost ($a.cost + $b.cost) -Conversions ($a.conversions + $b.conversions) -Revenue ($a.revenue + $b.revenue)
    } elseif ($gmOld.overview) {
        $gmOverview = $gmOld.overview
    }
    $gmByCampaignType = @(@($gmByCampaignType) + $oldTypes | Sort-Object -Property cost -Descending)
}

# ---------- 11번가 (광고상품별 일별 리포트, 다른 매체와 완전히 별개 매체) ----------
$stOverview = $null
$stByAdType = @()
$stByDate = @()
$stPeriod = $null

function DateAgg-ToRows {
    # @{ '2026.09.01.' = Agg } -> 날짜순 일별 행 목록
    param($DateAgg, [string]$Currency = 'KRW')
    $rows = @(foreach ($k in $DateAgg.Keys) {
        $m = Metrics-FromChannelAgg $DateAgg[$k] $Currency
        [PSCustomObject]@{
            date        = $k
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    return , @($rows | Sort-Object -Property date)
}

if ($elevenstAvailable) {
    $stTotalAgg = New-Agg
    $stDateAgg = @{}
    foreach ($f in $elevenstFiles) {
        if (-not $stPeriod) { $stPeriod = $f.Period }
        $typeAgg = New-Agg
        $typeDateAgg = @{}
        foreach ($r in $f.Rows) {
            # 전환은 '총 전환(직접+간접)' 기준 — 11번가 자체 '총 광고수익률' 컬럼과 같은 계산식. 장바구니 전환은 쓰지 않는다.
            $imp = ToNum $r.'노출수'; $clk = ToNum $r.'클릭수'; $cost = ToNum $r.'총비용'
            $conv = ToNum $r.'총 전환수'; $rev = ToNum $r.'총 전환 금액'
            $d = ($r.'날짜' -replace '-', '.') + '.'
            if (-not $stDateAgg.ContainsKey($d)) { $stDateAgg[$d] = New-Agg }
            if (-not $typeDateAgg.ContainsKey($d)) { $typeDateAgg[$d] = New-Agg }
            Add-Agg -Agg $stDateAgg[$d] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $typeDateAgg[$d] -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $typeAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
            Add-Agg -Agg $stTotalAgg -Imp $imp -Clk $clk -Cost $cost -Conv $conv -Rev $rev
        }
        $m = Metrics-FromAgg $typeAgg
        $stByAdType += [PSCustomObject]@{
            type        = $f.AdType
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
            byDate      = (DateAgg-ToRows $typeDateAgg)
        }
    }
    $stByAdType = @($stByAdType | Sort-Object -Property cost -Descending)

    $stByDate = DateAgg-ToRows $stDateAgg
    if (-not $stPeriod -and $stByDate.Count -gt 0) { $stPeriod = "$($stByDate[0].date)~$($stByDate[-1].date)" }
    $stOverview = Metrics-FromAgg $stTotalAgg
}

# ---------- GFA 상품별 (상품 성과 보고서가 있는 경우에만) ----------
# 구매완료 전환이 1건 이상 발생한 상품만 싣는다. 전체 상품 수는 표 설명에 함께 보여준다.
$gfaByProduct = @()
$gfaProductCount = 0
if ($gfaProductFile) {
    $gfaProdAgg = [ordered]@{}
    foreach ($r in (Get-GfaCsv -Path $gfaProductFile).Rows) {
        $key = "$($r.'쇼핑몰 상품 ID')|$($r.'캠페인 ID')"
        if (-not $gfaProdAgg.Contains($key)) { $gfaProdAgg[$key] = @{ agg = (New-Agg); row = $r } }
        Add-Agg -Agg $gfaProdAgg[$key].agg -Imp (ToNum $r.'노출수') -Clk (ToNum $r.'클릭수') -Cost (ToNum $r.'총비용') -Conv (ToNum $r.'구매완료 수') -Rev (ToNum $r.'구매완료 전환매출액')
    }
    $gfaProductCount = $gfaProdAgg.Count
    $gfaByProduct = @(foreach ($key in $gfaProdAgg.Keys) {
        $m = Metrics-FromAgg $gfaProdAgg[$key].agg
        if ($m.conversions -lt 1) { continue }
        [PSCustomObject]@{
            product     = $gfaProdAgg[$key].row.'상품명'
            productId   = $gfaProdAgg[$key].row.'쇼핑몰 상품 ID'
            campaign    = $gfaProdAgg[$key].row.'캠페인 이름'
            impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue     = $m.revenue; roas = $m.roas
        }
    })
    $gfaByProduct = @($gfaByProduct | Sort-Object -Property revenue -Descending)
}

# ---------- 네이버 전환유형 보고서 (있는 경우에만) ----------
# 전환수를 신청 완료, 구매완료처럼 유형별로 나눠 보여준다. 같은 전환이 이미 '총 전환수'에 들어 있으므로 합계에는 더하지 않는다.
$naverConvAgg = @{}
$convDetail = @()
if ($convTypeFile) {
    $detailAgg = @{}
    foreach ($r in (Get-NaverCsv -Path $convTypeFile).Rows) {
        $n = ToNum $r.'총 전환수'
        $k = "$($r.'전환 유형')`t$($r.'캠페인유형')`t$($r.'캠페인')`t$($r.'일별')"
        $naverConvAgg[$k] = [double]$naverConvAgg[$k] + $n
        $dk = "$($r.'전환 유형')`t$($r.'캠페인유형')`t$($r.'캠페인')`t$($r.'광고그룹')`t$($r.'키워드')`t$($r.'PC/모바일 매체')"
        $detailAgg[$dk] = [double]$detailAgg[$dk] + $n
    }
    $convDetail = @(foreach ($dk in $detailAgg.Keys) {
        $p = $dk.Split("`t")
        [PSCustomObject]@{ type = $p[0]; campaignType = $p[1]; campaign = $p[2]; adGroup = $p[3]; keyword = $p[4]; device = $p[5]; conversions = $detailAgg[$dk] }
    })
    $convDetail = @($convDetail | Sort-Object -Property conversions -Descending)
}

# ---------- 구글 (있는 경우에만) ----------
$googleChannel = $null
$googlePeriod = $null
$googleCurrency = 'KRW'
if ($googleAvailable) {
    if ($googleReport.Rows[0].'통화 코드') { $googleCurrency = $googleReport.Rows[0].'통화 코드' }
    Assert-Currency -Currency $googleCurrency -MediaName '구글'
    $googleChannel = New-Channel
    foreach ($r in $googleReport.Rows) {
        $d = if ($r.'일') { ($r.'일' -replace '-', '.') + '.' } else { '' }
        [AdFast]::AddChannelRow($googleChannel, $r.'캠페인 유형', $r.'캠페인', $d, (ToNum $r.'노출수'), (ToNum $r.'클릭수'), (ToNum $r.'비용'), (ToNum $r.'전환'), 0)
    }
    $days = @($googleChannel.day.Keys | Sort-Object)
    if ($days.Count -gt 0) { $googlePeriod = "$($days[0])~$($days[-1])" } else { $googlePeriod = $googleReport.Period }
}

# ---------- 메타 (있는 경우에만) ----------
$metaChannel = $null
$metaPeriod = $null
$metaConvAgg = @{}
if ($metaAvailable) {
    Assert-Currency -Currency $metaReport.Currency -MediaName '메타'
    # 결과가 없는 날은 '결과 유형'이 비어 있어서, 캠페인의 결과 유형을 먼저 모아 둔다.
    $metaTypeOf = @{}
    foreach ($r in $metaReport.Rows) { if ($r.'결과 유형') { $metaTypeOf[$r.'캠페인 이름'] = $r.'결과 유형' } }
    $metaChannel = New-Channel
    foreach ($r in $metaReport.Rows) {
        $name = $r.'캠페인 이름'
        $type = [string]$metaTypeOf[$name]
        $d = if ($r.'일') { ($r.'일' -replace '-', '.') + '.' } else { '' }
        $conv = ToNum $r.'결과'
        [AdFast]::AddChannelRow($metaChannel, $type, $name, $d, (ToNum $r.'노출'), (ToNum $r.($metaReport.ClickColumn)), (ToNum $r.($metaReport.CostColumn)), $conv, 0)
        if ($conv -gt 0) {
            $k = "$($r.'결과 유형')`t$type`t$name`t$d"
            $metaConvAgg[$k] = [double]$metaConvAgg[$k] + $conv
        }
    }
    $days = @($metaChannel.day.Keys | Sort-Object)
    if ($days.Count -gt 0) { $metaPeriod = "$($days[0])~$($days[-1])" }
}

function ConvertTo-ChannelData {
    # 매체 하나의 집계를 보고서에 넣을 모양으로 바꾼다. 통합 화면과 매체별 화면이 같은 데이터를 쓴다.
    # ConvAgg 키는 '전환유형<탭>캠페인유형<탭>캠페인<탭>일자', 값은 전환수다. 유형을 나눠 볼 수 없는 매체는 비워 둔다.
    param([string]$Key, [string]$Name, [string]$Period, $Overview, $Channel, $ConvAgg = @{}, [string]$Currency = 'KRW')

    $campaigns = @(foreach ($k in $Channel.camp.Keys) {
        $p = $k.Split("`t")
        $m = Metrics-FromChannelAgg $Channel.camp[$k] $Currency
        [PSCustomObject]@{
            campaignType = $p[0]
            campaign     = $p[1]
            impressions  = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
            cost         = $m.cost; cpc = $m.cpc; conversions = $m.conversions
            revenue      = $m.revenue; roas = $m.roas
        }
    })
    $campaigns = @($campaigns | Sort-Object -Property cost -Descending)
    $campIndex = @{}
    for ($i = 0; $i -lt $campaigns.Count; $i++) { $campIndex["$($campaigns[$i].campaignType)`t$($campaigns[$i].campaign)"] = $i }

    $daily = @(foreach ($k in $Channel.campDay.Keys) {
        $at = $k.LastIndexOf("`t")
        $m = Metrics-FromChannelAgg $Channel.campDay[$k] $Currency
        [PSCustomObject]@{
            c = $campIndex[$k.Substring(0, $at)]; date = $k.Substring($at + 1)
            impressions = $m.impressions; clicks = $m.clicks; cost = $m.cost; conversions = $m.conversions; revenue = $m.revenue
        }
    })

    $typeTotals = @{}
    foreach ($k in $ConvAgg.Keys) {
        $t = $k.Split("`t")[0]
        $typeTotals[$t] = [double]$typeTotals[$t] + $ConvAgg[$k]
    }
    $convTypes = @($typeTotals.Keys | Sort-Object { -$typeTotals[$_] } | ForEach-Object { [PSCustomObject]@{ type = $_; conversions = $typeTotals[$_] } })
    if ($convTypes.Count -eq 0) { $convTypes = @([PSCustomObject]@{ type = '전환'; conversions = $Overview.conversions }) }
    $typeIndex = @{}
    for ($i = 0; $i -lt $convTypes.Count; $i++) { $typeIndex[$convTypes[$i].type] = $i }
    $convRows = @(foreach ($k in $ConvAgg.Keys) {
        $p = $k.Split("`t")
        $c = $campIndex["$($p[1])`t$($p[2])"]
        if ($null -eq $c) { continue }
        [PSCustomObject]@{ t = $typeIndex[$p[0]]; c = $c; date = $p[3]; n = $ConvAgg[$k] }
    })

    [PSCustomObject]@{
        key           = $Key
        name          = $Name
        period        = $Period
        currency      = $Currency
        overview      = $Overview
        byCampaign    = $campaigns
        byDate        = (DateAgg-ToRows $Channel.day $Currency)
        campaignDaily = $daily
        convTypes     = $convTypes
        convRows      = $convRows
    }
}

# ---------- 오늘의집 (있는 경우에만, 다른 매체와 완전히 별개 매체) ----------
# 합계 · 캠페인별 · 일별은 캠페인 일별 > 캠페인(+계정 일별) > 계정 일별 > 상품(캠페인별) > 상품(전체) > 카테고리 > 검색어 보고서 순으로 하나만 쓴다
# (같은 광고비를 두 번 더하지 않게). 상품(캠페인별)은 '상품별 캠페인 성과'(한 상품이 여러 캠페인에 등록됐을 때 캠페인끼리 비교)에,
# 상품(전체)은 '상품별 성과'와 그 비교의 상품 전체 합계 · 등록 캠페인에, 카테고리는 '카테고리별 성과'에, 검색어는 '검색어별 성과'에 쓴다.
$ohOverview = $null; $ohChannelData = $null; $ohByProduct = @(); $ohKeywords = $null; $ohTotalSource = $null; $ohSources = @()
$ohProductTotals = @(); $ohByCategory = @(); $ohCategoryTotal = $null; $ohProductTotalCost = $null
if ($ohouseAvailable) {
    function ConvertTo-OhDate([string]$s) { if ($s -match '^(\d{4})-(\d{2})-(\d{2})') { return "$($matches[1]).$($matches[2]).$($matches[3])." }; return '' }
    $ohChannel = New-Channel
    $ohCampaignType = @{}   # 캠페인 ID -> 입찰 전략 (상품별 표에 붙인다)

    # 캠페인 일별 또는 캠페인 보고서. 캠페인 유형 자리에는 '입찰 전략'(수동 성과형 입찰 / 스마트 최적화 입찰)을 쓴다 ('유형'은 모두 오하우스애즈라서)
    $ohCampSrc = if ($ohCampaignDailyFile) { $ohCampaignDailyFile } else { $ohCampaignFile }
    if ($ohCampSrc) {
        $map = [OhouseCsv]::NewMap()
        [void][OhouseCsv]::Aggregate([OhouseCsv]::Read($ohCampSrc), @('캠페인 ID', '날짜'), @(), @('캠페인명', '캠페인 ID', '입찰 전략', '날짜'), [OhouseCsv]::NewMap(), $map)
        foreach ($a in $map.Values) {
            $ohCampaignType[$a['캠페인 ID']] = $a['입찰 전략']
            [AdFast]::AddChannelRow($ohChannel, $a['입찰 전략'], $a['캠페인명'], (ConvertTo-OhDate $a['날짜']), $a.imp, $a.clk, $a.cost, $a.conv, $a.rev)
        }
        $ohTotalSource = 'campaign'
        $ohSources += $(if ($ohCampaignDailyFile) { '캠페인 일별' } else { '캠페인' })
    }
    if ($ohDailyFile) {
        $map = [OhouseCsv]::NewMap()
        [void][OhouseCsv]::Aggregate([OhouseCsv]::Read($ohDailyFile), @('날짜'), @(), @('날짜'), [OhouseCsv]::NewMap(), $map)
        if (-not $ohTotalSource) {
            foreach ($a in $map.Values) { [AdFast]::AddChannelRow($ohChannel, '', '', (ConvertTo-OhDate $a['날짜']), $a.imp, $a.clk, $a.cost, $a.conv, $a.rev) }
            $ohChannel.camp.Clear(); $ohChannel.campDay.Clear()   # 캠페인 구분이 없는 파일이라 캠페인별 표는 비운다
            $ohTotalSource = 'account'
            $ohSources += '계정 일별'
        } elseif ($ohChannel.day.Count -eq 0) {
            # 캠페인 보고서에 날짜가 없으면 일별 성과만 계정 일별 보고서로 채운다 (합계에는 더하지 않는다)
            foreach ($a in $map.Values) {
                $d = ConvertTo-OhDate $a['날짜']
                if (-not $d) { continue }
                if (-not $ohChannel.day.ContainsKey($d)) { $ohChannel.day[$d] = New-Agg }
                [AdFast]::AddAgg($ohChannel.day[$d], $a.imp, $a.clk, $a.cost, $a.conv, $a.rev)
            }
            $ohSources += '계정 일별'
        }
    }

    # 상품 보고서 (캠페인마다 하나씩). 같은 날짜 · 캠페인 · 상품 줄이 앞 파일에 있으면 건너뛰어, 같은 파일을 두 번 넣어도 두 번 더하지 않는다.
    if ($ohProductFiles.Count -gt 0) {
        $prodMap = [OhouseCsv]::NewMap(); $prodSeen = [OhouseCsv]::NewMap()
        $prodDayMap = [OhouseCsv]::NewMap(); $prodDaySeen = [OhouseCsv]::NewMap()
        foreach ($f in $ohProductFiles) {
            $rows = [OhouseCsv]::Read($f)
            [void][OhouseCsv]::Aggregate($rows, @('캠페인 ID', '상품 ID'), @('날짜', '캠페인 ID', '상품 ID'), @('상품명', '상품 ID', '캠페인명', '캠페인 ID'), $prodSeen, $prodMap)
            if (-not $ohTotalSource) {
                [void][OhouseCsv]::Aggregate($rows, @('캠페인 ID', '날짜'), @('날짜', '캠페인 ID', '상품 ID'), @('캠페인명', '날짜'), $prodDaySeen, $prodDayMap)
            }
        }
        $ohByProduct = @(foreach ($a in $prodMap.Values) {
            $m = Metrics-FromAgg $a
            [PSCustomObject]@{
                product      = $a['상품명']
                productId    = $a['상품 ID']
                campaign     = $a['캠페인명']
                campaignId   = $a['캠페인 ID']
                campaignType = [string]$ohCampaignType[$a['캠페인 ID']]
                impressions  = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost         = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue      = $m.revenue; roas = $m.roas
            }
        })
        $ohByProduct = @($ohByProduct | Sort-Object -Property cost -Descending)
        if (-not $ohTotalSource) {
            # 캠페인 · 계정 보고서 없이 상품 보고서만 넣은 경우: 넣은 캠페인들의 합계를 쓴다
            foreach ($a in $prodDayMap.Values) { [AdFast]::AddChannelRow($ohChannel, '', $a['캠페인명'], (ConvertTo-OhDate $a['날짜']), $a.imp, $a.clk, $a.cost, $a.conv, $a.rev) }
            $ohTotalSource = 'product'
        }
        $ohSources += "캠페인별 상품(캠페인 $([OhouseCsv]::DistinctMeta($prodMap, '캠페인 ID'))개)"
    }

    # 상품(전체) 보고서. '캠페인 명/ID' 칸은 '허쉬침대프레임(68b7f86b056231530b560f65),#슬립퍼_프레임_허쉬_수동(6a7e...)' 모양이라
    # 끝의 '(24자리 ID)'를 기준으로 캠페인을 나눈다 (캠페인 이름에 쉼표나 괄호가 있어도 되게).
    if ($ohProductTotalFiles.Count -gt 0) {
        $ptMap = [OhouseCsv]::NewMap(); $ptSeen = [OhouseCsv]::NewMap()
        foreach ($f in $ohProductTotalFiles) {
            [void][OhouseCsv]::Aggregate([OhouseCsv]::Read($f), @('상품ID'), @('상품ID'), @('상품명', '상품ID', '캠페인 명/ID'), $ptSeen, $ptMap)
        }
        $ptTot = [OhouseCsv]::Totals($ptMap)
        $ohProductTotalCost = [int64]$ptTot[2]
        $ohProductTotals = @(foreach ($a in $ptMap.Values) {
            $m = Metrics-FromAgg $a
            $raw = [string]$a['캠페인 명/ID']
            $camps = @(foreach ($mt in [regex]::Matches($raw, '\s*(.+?)\(([0-9A-Fa-f]{24})\)\s*(?:,|$)')) { [PSCustomObject]@{ name = $mt.Groups[1].Value; id = $mt.Groups[2].Value } })
            if ($camps.Count -eq 0 -and $raw) { $camps = @([PSCustomObject]@{ name = $raw; id = '' }) }
            [PSCustomObject]@{
                product     = $a['상품명']
                productId   = $a['상품ID']
                campaigns   = $camps
                campaignNames = (($camps | ForEach-Object { $_.name }) -join ', ')
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $ohProductTotals = @($ohProductTotals | Sort-Object -Property cost -Descending)
        if (-not $ohTotalSource) {
            $ohOverview = New-Metrics -Impressions $ptTot[0] -Clicks $ptTot[1] -Cost $ptTot[2] -Conversions $ptTot[3] -Revenue $ptTot[4]
            $ohTotalSource = 'productTotal'
        }
        $ohSources += '상품'
    }

    # 카테고리 보고서
    if ($ohCategoryFiles.Count -gt 0) {
        $catMap = [OhouseCsv]::NewMap(); $catSeen = [OhouseCsv]::NewMap()
        foreach ($f in $ohCategoryFiles) {
            [void][OhouseCsv]::Aggregate([OhouseCsv]::Read($f), @('카테고리'), @('카테고리'), @('카테고리'), $catSeen, $catMap)
        }
        $catTot = [OhouseCsv]::Totals($catMap)
        $ohCategoryTotal = New-Metrics -Impressions $catTot[0] -Clicks $catTot[1] -Cost $catTot[2] -Conversions $catTot[3] -Revenue $catTot[4]
        $ohByCategory = @(foreach ($a in $catMap.Values) {
            $m = Metrics-FromAgg $a
            [PSCustomObject]@{
                category    = $a['카테고리']
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $ohByCategory = @($ohByCategory | Sort-Object -Property cost -Descending)
        if (-not $ohTotalSource) { $ohOverview = $ohCategoryTotal; $ohTotalSource = 'category' }
        $ohSources += '카테고리'
    }

    # 검색어 보고서. 상품을 합쳐 캠페인 · 검색어 · 타겟팅 유형 단위로 묶고, 클릭이 1회 이상인 검색어만 표에 싣는다.
    if ($ohKeywordFiles.Count -gt 0) {
        $kwMap = [OhouseCsv]::NewMap(); $kwSeen = [OhouseCsv]::NewMap()
        foreach ($f in $ohKeywordFiles) {
            [void][OhouseCsv]::Aggregate([OhouseCsv]::Read($f), @('캠페인 ID', '키워드', '키워드 타겟팅 유형'), @('캠페인 ID', '상품 ID', '키워드', '키워드 타겟팅 유형'), @('캠페인명', '키워드', '키워드 타겟팅 유형'), $kwSeen, $kwMap)
        }
        $kwTot = [OhouseCsv]::Totals($kwMap)
        $kwRows = @(foreach ($key in [OhouseCsv]::KeysWithClicks($kwMap)) {
            $a = $kwMap[$key]
            $m = Metrics-FromAgg $a
            [PSCustomObject]@{
                campaign    = $a['캠페인명']
                keyword     = ($a['키워드'] -replace '\s*[\r\n]+\s*', ' ')   # 검색어 안의 줄바꿈은 칸에서 한 줄로
                targeting   = $a['키워드 타겟팅 유형']
                impressions = $m.impressions; clicks = $m.clicks; ctr = $m.ctr
                cost        = $m.cost; cpc = $m.cpc; conversions = $m.conversions
                revenue     = $m.revenue; roas = $m.roas
            }
        })
        $ohKeywords = [PSCustomObject]@{
            rows                = @($kwRows | Sort-Object -Property cost -Descending)
            totalUniqueKeywords = [OhouseCsv]::DistinctMeta($kwMap, '키워드')
            campaigns           = @([OhouseCsv]::MetaValues($kwMap, '캠페인명'))   # 검색어 보고서에 들어 있는 캠페인 (광고비 비교용)
            totalCost           = [int64]$kwTot[2]
            zeroConversionSpend = [int64]$kwTot[5]
        }
        if (-not $ohTotalSource) {
            $ohOverview = New-Metrics -Impressions $kwTot[0] -Clicks $kwTot[1] -Cost $kwTot[2] -Conversions $kwTot[3] -Revenue $kwTot[4]
            $ohTotalSource = 'keyword'
        }
        $ohSources += '검색어'
    }

    if (-not $ohOverview) { $ohOverview = Metrics-FromAgg $ohChannel.total }
    $ohDays = @($ohChannel.day.Keys | Sort-Object)
    if (-not $ohPeriod -and $ohDays.Count -gt 0) { $ohPeriod = "$($ohDays[0])~$($ohDays[-1])" }
    $ohChannelData = ConvertTo-ChannelData -Key 'ohouse' -Name '오늘의집' -Period $ohPeriod -Overview $ohOverview -Channel $ohChannel
}

# ---------- 옥션 광고센터 파워클릭 기간별 → 일별 데이터 (2.040) ----------
# ConvertTo-ChannelData가 위에서 정의된 뒤라 여기서 바꾼다. 화면은 오늘의집과 같은 일별 성과 블록(renderDailyBlock)으로 그린다.
$auDaily = $null; $auPeriod = $null
if ($auctionAvailable -and $au.pcChannel) {
    $auPeriod = $au.pcPeriod
    $auDaily = ConvertTo-ChannelData -Key 'auction' -Name '옥션' -Period $au.pcPeriod -Overview (Metrics-FromAgg $au.pcChannel.total) -Channel $au.pcChannel
}
$gmOldDaily = $null
if ($gmOld -and $gmOld.pcChannel) {
    $gmOldDaily = ConvertTo-ChannelData -Key 'gmold' -Name '구 광고센터' -Period $gmOld.pcPeriod -Overview (Metrics-FromAgg $gmOld.pcChannel.total) -Channel $gmOld.pcChannel
}
# 지마켓 신 광고센터 기간 '2026-10-01 ~ 2026-10-31'을 보고서 기간 모양으로 (지마켓 · 옥션만 넣었을 때 제목의 월, 파일 이름에 쓴다)
$gmPeriodDots = $null
if ($gmPeriod -and "$gmPeriod" -match '(\d{4})\D(\d{2})\D(\d{2})\D+(\d{4})\D(\d{2})\D(\d{2})') {
    $gmPeriodDots = "$($matches[1]).$($matches[2]).$($matches[3]).~$($matches[4]).$($matches[5]).$($matches[6])."
}

# ---------- 대표 기간 결정 (지면별 > 시간대별 > 검색어 > 지역별 순 우선) ----------
$canonicalPeriod = $null
$periodSourceLabel = $null
if ($placementPeriod)      { $canonicalPeriod = $placementPeriod; $periodSourceLabel = "지면별" }
elseif ($hourlyPeriod)     { $canonicalPeriod = $hourlyPeriod;    $periodSourceLabel = "시간대별" }
elseif ($keywordPeriod)    { $canonicalPeriod = $keywordPeriod;   $periodSourceLabel = "검색어" }
elseif ($regionPeriod)     { $canonicalPeriod = $regionPeriod;    $periodSourceLabel = "지역별" }
elseif ($dowPeriod)        { $canonicalPeriod = $dowPeriod;       $periodSourceLabel = "요일_일자_기기별" }
elseif ($gfaPeriod)        { $canonicalPeriod = $gfaPeriod;       $periodSourceLabel = "GFA" }
elseif ($stPeriod)         { $canonicalPeriod = $stPeriod;        $periodSourceLabel = "11번가" }
elseif ($ohPeriod)         { $canonicalPeriod = $ohPeriod;        $periodSourceLabel = "오늘의집" }
elseif ($googlePeriod)     { $canonicalPeriod = $googlePeriod;    $periodSourceLabel = "구글" }
elseif ($metaPeriod)       { $canonicalPeriod = $metaPeriod;      $periodSourceLabel = "메타" }
elseif ($gmPeriodDots)     { $canonicalPeriod = $gmPeriodDots;    $periodSourceLabel = "지마켓" }
elseif ($gmOld -and $gmOld.pcPeriod) { $canonicalPeriod = $gmOld.pcPeriod; $periodSourceLabel = "구지마켓 일별" }
elseif ($auPeriod)         { $canonicalPeriod = $auPeriod;        $periodSourceLabel = "옥션 일별" }

if ($gfaPeriod -and $canonicalPeriod -and ($gfaPeriod -ne $canonicalPeriod) -and $periodSourceLabel -ne "GFA") {
    $gfaPeriodDiffNote = "GFA 리포트 기간($gfaPeriod)이 검색광고 리포트 기간($canonicalPeriod)과 다릅니다. 다운로드 시 기간 설정을 확인해주세요."
}

# 매체끼리 비교하는 보고서라, 기간이 다른 파일이 섞이면 알려 준다.
$channelPeriodNote = $null
$periodDiffs = @()
if ($googlePeriod -and $canonicalPeriod -and $googlePeriod -ne $canonicalPeriod) { $periodDiffs += "구글 $googlePeriod" }
if ($metaPeriod -and $canonicalPeriod -and $metaPeriod -ne $canonicalPeriod) { $periodDiffs += "메타 $metaPeriod" }
if ($gfaPeriod -and $canonicalPeriod -and $gfaPeriod -ne $canonicalPeriod) { $periodDiffs += "GFA $gfaPeriod" }
# ---------- 옥션 AI매출업: 2026년 10월까지는 지마켓 + 옥션 합친 숫자 (2.040) ----------
# 사용자(2026-11-06): AI매출업 리포트는 지마켓과 옥션이 합쳐져 나와 사이트별로 나눌 수 없다. 옥션 탭에만 한 번 넣고 합친 숫자라고 적는다.
# 2026년 11월부터는 AI매출업이 옥션뿐이라 그냥 'AI매출업'. 기준은 보고서 기간 시작일(AI매출업 일별만 넣었으면 그 날짜). 기간을 모르면 옥션으로 본다.
# 이름에 '구지마켓'을 넣은 AI매출업 파일은 사용자가 정한 것이라 예전처럼 지마켓 탭으로 간다.
$auAiCombined = $false
if ($auctionAvailable -and @($auByAdType | Where-Object { $_.type -eq 'AI매출업' }).Count -gt 0 -and "$canonicalPeriod" -match '^(\d{4})\.(\d{2})\.') {
    $auAiCombined = (([int]$matches[1]) * 100 + [int]$matches[2]) -lt 202611
}
if ($auAiCombined) {
    $aiLabel = 'AI매출업 (지마켓+옥션)'
    foreach ($t in $auByAdType) { if ($t.type -eq 'AI매출업') { $t.type = $aiLabel } }
    foreach ($p in $auByProduct) { if ($p.adType -eq 'AI매출업') { $p.adType = $aiLabel } }
    if ($auDaily) { foreach ($c in $auDaily.byCampaign) { if ($c.campaignType -eq 'AI매출업') { $c.campaignType = $aiLabel; $c.campaign = $aiLabel } } }
}

if ($periodDiffs.Count -gt 0) { $channelPeriodNote = "매체별 데이터 기간이 다릅니다 (기준 $canonicalPeriod, $($periodDiffs -join ', ')). 해당 기간에 집행이 없었던 것이 아니라면 파일 기간을 확인해 주세요." }

# ---------- 전체 요약 (검색광고 + GFA 통합) ----------
# 지면별이 없으면 시간대별 -> 검색어 -> 지역별 순으로 "검색광고 전체 합계"를 대신 추정한다.
$saOverviewForTotal = $null
$saTotalSourceLabel = $null
if ($saOverview) {
    $saOverviewForTotal = $saOverview
    $saTotalSourceLabel = "지면별"
} elseif ($hasHourly) {
    $saOverviewForTotal = New-Metrics -Impressions $hourlyTotalImp -Clicks ($hourlyRows | Measure-Object -Property clicks -Sum).Sum -Cost $hourlyTotalCost -Conversions ($hourlyRows | Measure-Object -Property conversions -Sum).Sum -Revenue ($hourlyRows | Measure-Object -Property revenue -Sum).Sum
    $saTotalSourceLabel = "시간대별"
} elseif ($hasKeyword) {
    $saOverviewForTotal = New-Metrics -Impressions $keywordTotals[0] -Clicks $keywordTotals[1] -Cost $keywordTotals[2] -Conversions $keywordTotals[3] -Revenue $keywordTotals[4]
    $saTotalSourceLabel = "검색어"
} elseif ($hasRegion) {
    $saOverviewForTotal = New-Metrics -Impressions ($regionRows | Measure-Object -Property impressions -Sum).Sum -Clicks ($regionRows | Measure-Object -Property clicks -Sum).Sum -Cost ($regionRows | Measure-Object -Property cost -Sum).Sum -Conversions ($regionRows | Measure-Object -Property conversions -Sum).Sum -Revenue ($regionRows | Measure-Object -Property revenue -Sum).Sum
    $saTotalSourceLabel = "지역별(파워링크 전용, 참고치)"
}

$grandAgg = New-Agg
if ($saOverviewForTotal) {
    Add-Agg -Agg $grandAgg -Imp $saOverviewForTotal.impressions -Clk $saOverviewForTotal.clicks -Cost $saOverviewForTotal.cost -Conv $saOverviewForTotal.conversions -Rev $saOverviewForTotal.revenue
}
if ($gfaAvailable) {
    Add-Agg -Agg $grandAgg -Imp $gfaOverview.impressions -Clk $gfaOverview.clicks -Cost $gfaOverview.cost -Conv $gfaOverview.conversions -Rev $gfaOverview.revenue
}
$overview = Metrics-FromAgg $grandAgg

$byCampaignType = @($byCampaignType)
if ($gfaAvailable) {
    $byCampaignType += [PSCustomObject]@{
        type        = "GFA"
        impressions = $gfaOverview.impressions; clicks = $gfaOverview.clicks; ctr = $gfaOverview.ctr
        cost        = $gfaOverview.cost; cpc = $gfaOverview.cpc; conversions = $gfaOverview.conversions
        revenue     = $gfaOverview.revenue; roas = $gfaOverview.roas
    }
}
$byCampaignType = @($byCampaignType | Sort-Object -Property cost -Descending)

# ---------- 정합성 체크 (둘 다 존재하는 소스끼리만 비교) ----------
$costDiffHourly = $null
$impDiffHourly = $null
if ($saOverview -and $hasHourly) {
    $costDiffHourly = [math]::Abs($saOverview.cost - $hourlyTotalCost)
    $impDiffHourly  = [math]::Abs($saOverview.impressions - $hourlyTotalImp)
}
$costDiffKeyword = $null
if ($saOverview -and $hasKeyword) {
    $costDiffKeyword = [math]::Abs($saOverview.cost - $keywordTotalCost)
}

$qualityNotes = @()
if ($hasRegion) { $qualityNotes += "지역별 리포트는 파워링크 캠페인 데이터만 포함합니다 (쇼핑검색 캠페인 제외)." }
if ($keywordGapNote) { $qualityNotes += $keywordGapNote }
if (-not $saOverview -and $saTotalSourceLabel) { $qualityNotes += "지면별 데이터가 없어 전체 요약의 검색광고 합계는 '$saTotalSourceLabel' 데이터를 기준으로 추정했습니다." }
if (-not $hasPlacement -and -not $gfaAvailable) { $qualityNotes += "캠페인 유형별 비교에 사용할 데이터가 없습니다." }

$dataQuality = [PSCustomObject]@{
    placementVsHourlyCostDiff  = $costDiffHourly
    placementVsHourlyImpDiff   = $impDiffHourly
    placementVsKeywordCostDiff = $costDiffKeyword
    channelPeriodWarning       = $channelPeriodNote
    note                       = ($qualityNotes -join ' ')
    gfaPeriodWarning           = $gfaPeriodDiffNote
    saTotalSource              = $saTotalSourceLabel
}

# ---------- 매체 목록 (통합 화면, 구글, 메타 화면과 네이버 캠페인별 표가 쓴다) ----------
$channels = @()
if ($saOverviewForTotal) { $channels += ConvertTo-ChannelData -Key 'naver' -Name '네이버 검색광고' -Period $placementPeriod -Overview $saOverviewForTotal -Channel $naverChannel -ConvAgg $naverConvAgg }
if ($gfaAvailable) { $channels += ConvertTo-ChannelData -Key 'gfa' -Name '네이버 GFA' -Period $gfaPeriod -Overview $gfaOverview -Channel $gfaChannel }
if ($googleAvailable) { $channels += ConvertTo-ChannelData -Key 'google' -Name '구글' -Period $googlePeriod -Currency $googleCurrency -Overview (Metrics-FromChannelAgg $googleChannel.total $googleCurrency) -Channel $googleChannel }
if ($metaAvailable) { $channels += ConvertTo-ChannelData -Key 'meta' -Name '메타' -Period $metaPeriod -Currency $metaReport.Currency -Overview (Metrics-FromChannelAgg $metaChannel.total $metaReport.Currency) -Channel $metaChannel -ConvAgg $metaConvAgg }

# 매출 데이터가 없고 전환만 있으면 문의, 신청을 전환으로 받는 광고주로 본다.
# 이때는 쓸 수 없는 전환매출액, ROAS 대신 전환율(CVR), 전환당 비용(CPA)을 보여준다.
$allOverviews = @($channels | ForEach-Object { $_.overview }) + @($gmOverview, $auOverview, $stOverview, $ohOverview | Where-Object { $_ })
$leadMode = ((($allOverviews | Measure-Object -Property revenue -Sum).Sum -eq 0) -and (($allOverviews | Measure-Object -Property conversions -Sum).Sum -gt 0))

$allMetricKeys = @('impressions','clicks','ctr','cpc','cost','conversions','cvr','cpa','revenue','roas')
$selectedMetricKeys = @($VisibleMetrics -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$visibleMetricsObj = [PSCustomObject]@{}
foreach ($mk in $allMetricKeys) {
    $on = $selectedMetricKeys -contains $mk
    if ($mk -eq 'cvr' -or $mk -eq 'cpa') { $on = $leadMode -and ($selectedMetricKeys -contains 'conversions') }
    if ($leadMode -and ($mk -eq 'revenue' -or $mk -eq 'roas')) { $on = $false }
    $visibleMetricsObj | Add-Member -NotePropertyName $mk -NotePropertyValue $on
}

$naverAvailable = ($hasPlacement -or $hasKeyword -or $hasHourly -or $hasRegion -or $gfaAvailable -or $hasDow)

$result = [PSCustomObject]@{
    generatedAt      = (Get-Date -Format "yyyy-MM-dd HH:mm")
    author           = $AuthorInfo
    visibleMetrics   = $visibleMetricsObj
    leadMode         = $leadMode
    channels         = $channels
    convDetail       = $convDetail
    naverAvailable   = $naverAvailable
    period           = $canonicalPeriod
    periodsBySource  = [PSCustomObject]@{
        placement = $placementPeriod
        keyword   = $keywordPeriod
        hourly    = $hourlyPeriod
        region    = $regionPeriod
        gfa       = $gfaPeriod
        dow       = $dowPeriod
    }
    sections         = [PSCustomObject]@{
        placement = ($hasPlacement -and -not $naverCampaignOnly)
        keyword   = $hasKeyword
        hourly    = $hasHourly
        region    = $hasRegion
        gfa       = $gfaAvailable
        dow       = $hasDow
        compare   = ($byCampaignType.Count -gt 0)
    }
    overview         = $overview
    saOverview       = $saOverview
    byCampaignType   = $byCampaignType
    byPlacement      = $byPlacement
    naverCampaignOnly = $naverCampaignOnly
    topKeywordsByCost = $topKeywordsByCost
    topKeywordsByConv = $topKeywordsByConv
    keywordDimension = $keywordDimension
    keywordTable2    = $keywordTable2
    keywordStats     = [PSCustomObject]@{
        totalUniqueKeywords = $keywordUniqueCount
        shownRows           = $keywordShownCount
        zeroConversionSpend = [int64]$zeroConvSpend
    }
    hourly           = $hourlyRows
    region           = [PSCustomObject]@{
        note = "파워링크 캠페인 전용 데이터 (쇼핑검색 제외)"
        rows = $regionRows
    }
    gfa              = [PSCustomObject]@{
        available  = $gfaAvailable
        period     = $gfaPeriod
        overview   = $gfaOverview
        byCampaign = $gfaByCampaign
        byProduct  = $gfaByProduct
        productCount = $gfaProductCount
        bySegment  = $gfaBySegment
        byDate     = $gfaByDate
    }
    dow              = [PSCustomObject]@{
        available = $hasDow
        period    = $dowPeriod
        byDevice  = $byDevice
        byDate    = $byDate
        byWeekday = $byWeekday
    }
    dataQuality      = $dataQuality
    gmarket          = [PSCustomObject]@{
        available         = $gmarketAvailable
        period            = $gmPeriod
        overview          = $gmOverview
        byCampaignType    = $gmByCampaignType
        byCampaign        = $gmByCampaign      # 신 광고센터 캠페인별 리포트 (2.039)
        overviewSource    = $(if ($hasGmarketCampaign) { 'campaign' } elseif ($hasGmarketItem) { 'item' } elseif ($hasGmarketKeyword) { 'keyword' } else { $null })
        byProduct         = $gmByProduct
        topKeywordsByCost = $gmTopKeywordsByCost
        keywordStats      = [PSCustomObject]@{
            totalUniqueKeywords = $gmKeywordUniqueCount
            zeroConversionSpend = [int64]$gmZeroConvSpend
        }
        hasNewCenter      = $gmNewAvailable
        # 구 광고센터(옥션 광고센터에서 받은 지마켓). overview는 위 합계에 이미 더했다.
        legacy            = $(if ($gmOld) { [PSCustomObject]@{
            overview          = $gmOld.overview
            byAdType          = $gmOld.byAdType
            byProduct         = $gmOld.byProduct
            topKeywordsByCost = $gmOld.topKeywordsByCost
            hasImpressionGap  = $gmOld.hasImpressionGap
            daily             = $gmOldDaily   # 파워클릭 기간별 (2.040)
            pcSource          = $gmOld.pcSource
            dailyKinds        = $gmOld.dailyKinds
            keywordStats      = [PSCustomObject]@{
                totalUniqueKeywords = $gmOld.keywordUniqueCount
                zeroConversionSpend = [int64]$gmOld.zeroConvSpend
            }
        } } else { $null })
    }
    auction          = [PSCustomObject]@{
        available         = $auctionAvailable
        overview          = $auOverview
        byAdType          = $auByAdType
        byProduct         = $auByProduct
        topKeywordsByCost = $auTopKeywordsByCost
        hasImpressionGap  = $auHasImpressionGap
        period            = $auPeriod     # 파워클릭 기간별이 있을 때만 (2.040)
        daily             = $auDaily
        pcSource          = $(if ($auctionAvailable) { $au.pcSource } else { $null })
        dailyKinds        = $(if ($auctionAvailable) { $au.dailyKinds } else { @() })
        aiCombined        = $auAiCombined   # AI매출업이 지마켓 + 옥션 합친 숫자 (2026년 10월까지)
        keywordStats      = [PSCustomObject]@{
            totalUniqueKeywords = $auKeywordUniqueCount
            zeroConversionSpend = [int64]$auZeroConvSpend
        }
    }
    elevenst         = [PSCustomObject]@{
        available = $elevenstAvailable
        period    = $stPeriod
        overview  = $stOverview
        byAdType  = $stByAdType
        byDate    = $stByDate
    }
    ohouse           = [PSCustomObject]@{
        available   = $ohouseAvailable
        period      = $ohPeriod
        overview    = $ohOverview
        totalSource = $ohTotalSource
        sources     = $ohSources
        channel     = $ohChannelData
        byProduct   = $ohByProduct
        productTotals = $ohProductTotals
        productTotalCost = $ohProductTotalCost
        byCategory  = $ohByCategory
        categoryTotal = $ohCategoryTotal
        keywords    = $ohKeywords
    }
}

# 네이버 엑셀을 바꾼 CSV와 복사본은 모든 파일을 다 읽었으므로 지운다 (거래처 데이터가 임시 폴더에 남지 않게)
# (Get-Content -TotalCount로 앞줄만 읽은 파일은 연결이 바로 닫히지 않아, 먼저 정리해야 지워진다)
if ($naverXlsxNames.Count -gt 0 -and (Test-Path -LiteralPath $naverStage)) {
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $naverStage -Recurse -Force -ErrorAction SilentlyContinue
}

$outDir = Split-Path $OutputPath -Parent
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# HTML 보고서를 만들 때는 같은 내용이 HTML 안에 들어가므로 JSON 파일을 따로 남기지 않는다.
if (-not $TemplatePath) {
    $jsonOnly = $result | ConvertTo-Json -Depth 10 -Compress
    if ($keywordRowsJson) { $jsonOnly = $jsonOnly.Substring(0, $jsonOnly.Length - 1) + ',"keywordRows":' + $keywordRowsJson + '}' }   # {"cols":[...],"rows":[[...]]} 모양 (ShownRowsJson)
    [System.IO.File]::WriteAllText($OutputPath, $jsonOnly, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "OK -> $OutputPath"
}

if ($TemplatePath) {
    if (-not $HtmlOutputPath) {
        # 파일명 형식: '2609_업체명_광고_보고서.html' (시작일 기준 연월 YYMM). 주간 보고서는 같은 달에 여러 개라 '_N주차'를 붙인다.
        # 한 달(31일)보다 긴 보고서(예: 1년치 검색어)는 '2511-2610_업체명_광고_보고서.html'처럼 시작 · 끝 연월을 쓴다.
        # 이 이름은 월별 추이가 '생성된 보고서' 폴더에서 찾는 이름 규칙에 맞지 않아, 한 달 보고서로 잘못 들어가지 않는다.
        $yymm = Get-Date -Format "yyMM"
        $weekSuffix = ""
        if ($result.period -match '(\d{4})\.(\d{2})\.(\d{2})\.~(\d{4})\.(\d{2})\.(\d{2})\.') {
            $yymm = $matches[1].Substring(2) + $matches[2]
            $fStart = Get-Date -Year ([int]$matches[1]) -Month ([int]$matches[2]) -Day ([int]$matches[3])
            $fEnd   = Get-Date -Year ([int]$matches[4]) -Month ([int]$matches[5]) -Day ([int]$matches[6])
            $fDays  = [int]($fEnd - $fStart).TotalDays + 1
            if ($fDays -lt 25) { $weekSuffix = "_$([math]::Ceiling([int]$matches[3] / 7.0))주차" }
            elseif ($fDays -gt 31) { $yymm += "-" + $matches[4].Substring(2) + $matches[5] }
        }
        $fileNamePrefix = "${yymm}_"
        if ($ClientName) {
            # 파일명에 못 쓰는 문자 제거 (\ / : * ? " < > |)
            $safeClient = ($ClientName -replace '[\\/:\*\?"<>\|]', '').Trim()
            if ($safeClient) { $fileNamePrefix += "${safeClient}_" }
        }
        $HtmlOutputPath = Join-Path (Split-Path $OutputPath -Parent) "${fileNamePrefix}광고_보고서${weekSuffix}.html"
    }

    $periodLabel = ""
    if ($result.period -match '(\d{4})\.(\d{2})\.(\d{2})\.~(\d{4})\.(\d{2})\.(\d{2})\.') {
        $sy = [int]$matches[1]; $sm = [int]$matches[2]; $sd = [int]$matches[3]
        $ey = [int]$matches[4]; $em = [int]$matches[5]; $ed = [int]$matches[6]
        $startDate = Get-Date -Year $sy -Month $sm -Day $sd
        $endDate   = Get-Date -Year $ey -Month $em -Day $ed
        $dayCount  = [int]($endDate - $startDate).TotalDays + 1
        if ($dayCount -gt 31) {
            # 한 달보다 긴 보고서(예: 1년치)는 제목에 월을 붙이지 않는다. 보고 기간이 제목 아래에 따로 나온다 (2026-11-04 사용자)
            $periodLabel = ""
        } elseif ($dayCount -ge 25) {
            # 월간 보고서: 시작일 기준 월
            $periodLabel = "${sm}월"
        } else {
            # 주간(또는 그 외 단기) 보고서: 시작일 기준 몇 주차
            $weekNum = [math]::Ceiling($sd / 7.0)
            $periodLabel = "${sm}월 ${weekNum}주차"
        }
    }

    $titlePrefix = ""
    if ($periodLabel) { $titlePrefix += "$periodLabel " }
    if ($ClientName)  { $titlePrefix += "$ClientName " }
    # 매체를 드롭다운으로 함께 보여주는 리포트(2개 이상 매체 포함)는 제목에 특정 매체명을 박아두지 않는다.
    # 매체가 하나뿐이면 그 매체 이름을 제목에 명시한다 (예: 옥션만 있으면 "네이버"라고 쓰면 안 됨).
    $mediaCount = @($naverAvailable, $gmarketAvailable, $auctionAvailable, $elevenstAvailable, $googleAvailable, $metaAvailable, $ohouseAvailable | Where-Object { $_ }).Count
    $productLabel = if ($mediaCount -ge 2) {
        "광고 성과 보고서"
    } elseif ($naverAvailable) {
        if ($gfaAvailable) { "네이버 광고(검색광고+GFA) 성과 보고서" } else { "네이버 검색광고 성과 보고서" }
    } elseif ($gmarketAvailable) {
        "지마켓 성과 보고서"
    } elseif ($auctionAvailable) {
        "옥션 성과 보고서"
    } elseif ($elevenstAvailable) {
        "11번가 성과 보고서"
    } elseif ($googleAvailable) {
        "구글 광고 성과 보고서"
    } elseif ($metaAvailable) {
        "메타 광고 성과 보고서"
    } elseif ($ohouseAvailable) {
        "오늘의집 광고 성과 보고서"
    } else {
        "광고 성과 보고서"
    }
    $reportTitle = "${titlePrefix}${productLabel}"

    $jsonForHtml = ($result | ConvertTo-Json -Depth 10 -Compress)
    $jsonForHtml = $jsonForHtml.Replace('</', '<\/')

    # 로고 자동 탐색: 명시적으로 안 주면 프로젝트 루트의 assets 폴더에서 찾는다.
    if (-not $LogoPath) {
        $projectRoot = Split-Path $PSScriptRoot -Parent
        $assetsDir = Join-Path $projectRoot "assets"
        if (Test-Path -LiteralPath $assetsDir) {
            $found = Get-ChildItem -LiteralPath $assetsDir -File | Where-Object { $_.Name -match '(?i)(logo|로고)' -and $_.Extension -match '(?i)\.(png|jpg|jpeg)$' } | Select-Object -First 1
            if ($found) { $LogoPath = $found.FullName }
        }
    }

    $logoDataUri = "__LOGO_DATA_URI__"
    if ($LogoPath -and (Test-Path -LiteralPath $LogoPath)) {
        $ext = [System.IO.Path]::GetExtension($LogoPath).ToLower()
        $mime = switch ($ext) { ".png" { "image/png" }; ".jpg" { "image/jpeg" }; ".jpeg" { "image/jpeg" }; default { "image/png" } }
        $bytes = [System.IO.File]::ReadAllBytes($LogoPath)
        $b64 = [System.Convert]::ToBase64String($bytes)
        $logoDataUri = "data:$mime;base64,$b64"
    }

    $template = Get-Content -LiteralPath $TemplatePath -Raw -Encoding UTF8
    $html = $template.Replace('__REPORT_TITLE__', $reportTitle).Replace('__REPORT_DATA_JSON__', $jsonForHtml).Replace('__LOGO_DATA_URI__', $logoDataUri)
    # 검색어 · 키워드 표 줄은 따로 둔 칸에 (수십만 줄이어도 report-data는 작게). 검색어에 '__LOGO_DATA_URI__' 같은 글자가 있어도 바뀌지 않게 맨 마지막에 넣는다
    $rowsJson = '{"topKeywordsByCost":' + $(if ($keywordRowsJson) { $keywordRowsJson } else { '[]' }) + ',"keywordTable2Rows":' + $(if ($keyword2RowsJson) { $keyword2RowsJson } else { 'null' }) + '}'
    $html = $html.Replace('__REPORT_ROWS_JSON__', $rowsJson)

    $enc = New-Object System.Text.UTF8Encoding($true)
    $htmlDir = Split-Path $HtmlOutputPath -Parent
    if ($htmlDir -and -not (Test-Path -LiteralPath $htmlDir)) { New-Item -ItemType Directory -Force -Path $htmlDir | Out-Null }
    if ($PendingHtmlPath -and (Test-Path -LiteralPath $HtmlOutputPath)) {
        [System.IO.File]::WriteAllText($PendingHtmlPath, $html, $enc)
        [System.IO.File]::WriteAllText("$PendingHtmlPath.txt", $HtmlOutputPath, [System.Text.Encoding]::Unicode)
        Write-Output "EXISTS -> $HtmlOutputPath"
        exit 3
    }
    [System.IO.File]::WriteAllText($HtmlOutputPath, $html, $enc)
    Write-Output "OK -> $HtmlOutputPath"
}
