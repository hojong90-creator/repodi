<#
.SYNOPSIS
  '보고서 생성용 데이터' 폴더의 파일마다 어떤 리포트로 인식되는지 적는다 (레포디 창의 파일 목록에 표시, 2.032).
  -OutFile에 한 줄에 '파일 이름<탭>종류'로 적는다 (UTF-16, HTA가 그대로 읽는다). 인식하지 못한 파일은 종류가 '?'로 시작한다.
.NOTES
  보고서를 만들 때 실제로 종류를 가르는 곳은 Build-AdReport.ps1(HTML 보고서)과 Build-ExcelReport.ps1(캐츠온 엑셀)이다.
  여기는 같은 규칙을 파일 앞 몇 줄만 보고 따라 한다. 그쪽 판별 규칙을 바꾸면 여기도 함께 고칠 것.
  큰 파일(1년치 검색어 등)을 통째로 읽지 않는다. xlsx도 첫 시트의 앞 몇 줄만 읽는다.
#>
param([Parameter(Mandatory = $true)][string]$InputDir, [Parameter(Mandatory = $true)][string]$OutFile)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

function Read-Head([string]$Path, [int]$Count) {
    $sr = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try {
        $lines = New-Object System.Collections.Generic.List[string]
        while ($lines.Count -lt $Count -and -not $sr.EndOfStream) { $lines.Add($sr.ReadLine()) }
        return , $lines.ToArray()
    } finally { $sr.Close() }
}

function Split-CsvLine([string]$Line) {
    # 따옴표 칸 안의 쉼표를 지키며 한 줄을 나눈다 (앞 몇 줄만 보면 되어 간단히)
    $cells = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQ = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inQ) {
            if ($ch -eq '"') { if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$sb.Append('"'); $i++ } else { $inQ = $false } }
            else { [void]$sb.Append($ch) }
        } elseif ($ch -eq '"') { $inQ = $true }
        elseif ($ch -eq ',') { $cells.Add($sb.ToString()); [void]$sb.Clear() }
        else { [void]$sb.Append($ch) }
    }
    $cells.Add($sb.ToString())
    return , @($cells | ForEach-Object { $_.Trim().Trim([char]0xFEFF).Trim() })
}

function Get-NaverKind($Cols) {
    # Build-AdReport.ps1의 네이버 CSV 판별과 같은 순서
    $isTerm = ($Cols -contains '검색어') -or ($Cols -contains '키워드')
    if ($Cols -contains '전환 유형') { return '네이버 전환유형' }
    if (($Cols -contains '매체이름') -and $isTerm -and ($Cols -contains '요일별')) {
        return $(if ($Cols -contains '검색어') { '네이버 통합 (지면 · 검색어 · 요일)' } else { '네이버 통합 (지면 · 키워드 · 요일)' })
    }
    if ($Cols -contains '시간대별') { return '네이버 시간대별' }
    if ($Cols -contains '지역') { return '네이버 지역' }
    if ($Cols -contains '매체이름') { return '네이버 지면별' }
    if ($Cols -contains '검색어') { return '네이버 검색어' }
    if ($Cols -contains '키워드') { return '네이버 키워드' }
    if (($Cols -contains '캠페인') -and ($Cols -contains '노출수')) { return '네이버 캠페인' }
    if (($Cols -contains '요일별') -or ($Cols -contains '일별') -or ($Cols -contains 'PC/모바일 매체')) { return '네이버 요일 · 일자 · 기기' }
    return '? 네이버 (쓸 수 있는 항목 없음)'
}

function Test-Gfa([string]$Line) {
    # Build-AdReport.ps1 Test-GfaHeader와 같다
    if ($Line -match '연령 및 성별') { return $true }
    return (($Line -match '캠페인 이름') -and ($Line -match '총비용') -and (($Line -match '캠페인 목적') -or ($Line -match '캠페인 ID')) -and
        ($Line -notmatch '지출 금액') -and ($Line -notmatch '네이버쇼핑 상품 ID'))
}

function Get-CsvKind([string[]]$Head, [string]$Name) {
    if (-not $Head.Count) { return '? 빈 파일' }
    $first = $Head[0]
    if ($first -match '\d{4}\.\d{2}\.\d{2}\.~\d{4}\.\d{2}\.\d{2}\.') {
        if ($Head.Count -lt 2) { return '? 네이버 (컬럼 줄 없음)' }
        # 달마다 나눠 받은 파일이 겹치거나 빠지지 않았는지 보기 쉽게 기간도 적는다
        $kind = Get-NaverKind (Split-CsvLine $Head[1])
        if ($kind -notlike '?*' -and $first -match '(\d{4}\.\d{2}\.\d{2}\.)~(\d{4}\.\d{2}\.\d{2}\.)') { $kind += " ($($matches[1]) ~ $($matches[2]))" }
        return $kind
    }
    if (Test-Gfa $first) { return 'GFA 성과' }
    if ($first -match '네이버쇼핑 상품 ID') { return 'GFA 상품 성과' }
    $cols = Split-CsvLine $first
    # 오늘의집 (Build-AdReport.ps1의 오늘의집 판별과 같은 순서). 캠페인마다 받는 보고서는 어느 캠페인 것인지 함께 적는다
    if (($cols -contains '집행 광고비') -and (($cols -contains '광고 전환 매출') -or ($cols -contains '전환 매출'))) {
        $camp = ''
        $ci = [array]::IndexOf($cols, '캠페인명')
        if ($ci -ge 0 -and $Head.Count -ge 2) { $row = Split-CsvLine $Head[1]; if ($ci -lt $row.Count -and $row[$ci]) { $camp = " ($($row[$ci]))" } }
        if ($cols -contains '키워드') { return "오늘의집 검색어$camp" }
        if ($cols -contains '캠페인 명/ID') { return '오늘의집 상품별 통합' }
        if ($cols -contains '상품 ID') { return "오늘의집 캠페인별 상품$camp" }
        if ($cols -contains '카테고리') { return '오늘의집 카테고리별 통합' }
        if ($cols -contains '캠페인 ID') { return $(if ($cols -contains '날짜') { '오늘의집 캠페인 일별' } else { '오늘의집 캠페인 (기간 합계)' }) }
        if ($cols -contains '날짜') { return '오늘의집 계정 일별' }
        return '? 오늘의집 (알 수 없는 보고서)'
    }
    $tabCols = @($first.Split("`t") | ForEach-Object { $_.Trim().Trim([char]0xFEFF) })
    if (($tabCols -contains '날짜') -and ($tabCols -contains '총 전환 금액')) {
        $adType = ([System.IO.Path]::GetFileNameWithoutExtension($Name) -replace '\s*보고서.*$', '').Trim()
        return "11번가 $adType"
    }
    if (($first -match '캠페인 이름') -and ($first -match '지출 금액 \([A-Za-z]{3}\)')) { return '메타' }
    foreach ($line in $Head) {
        foreach ($d in "`t", ",") {
            $c = @($line.Split($d) | ForEach-Object { $_.Trim().Trim('"').Trim([char]0xFEFF) })
            if (($c -contains '캠페인') -and ($c -contains '비용')) { return '구글' }
        }
    }
    return '? 인식 못 함'
}

function Read-XlsxHead([string]$Path, [int]$MaxRows) {
    # 첫 시트의 앞 몇 줄만 읽는다. 공유 문자열도 필요한 번호까지만 읽는다 (네이버 1년치 xlsx는 공유 문자열만 수십 MB일 수 있다).
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sheet = $zip.GetEntry('xl/worksheets/sheet1.xml')
        if (-not $sheet) { $sheet = $zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/*.xml' } | Select-Object -First 1 }
        if (-not $sheet) { return , @() }
        $rows = New-Object System.Collections.Generic.List[object]
        $r = [System.Xml.XmlReader]::Create($sheet.Open())
        try {
            $cur = $null; $col = 0; $type = $null
            $ok = $r.Read()
            while ($ok) {
                if ($r.NodeType -eq 'Element') {
                    if ($r.LocalName -eq 'row') { $cur = New-Object System.Collections.Generic.List[object] }
                    elseif ($r.LocalName -eq 'c' -and $cur) {
                        $ref = $r.GetAttribute('r'); $type = $r.GetAttribute('t')
                        $col = 0; if ($ref) { foreach ($ch in ($ref -replace '\d', '').ToCharArray()) { $col = $col * 26 + ([int][char]$ch - 64) }; $col-- } else { $col = $cur.Count }
                        while ($cur.Count -le $col) { $cur.Add('') }
                    }
                    elseif (($r.LocalName -eq 'v' -or $r.LocalName -eq 't') -and $cur) {
                        $v = $r.ReadElementContentAsString()
                        $cur[$col] = if ($type -eq 's') { [int]$v } else { [string]$v }   # 공유 문자열 번호는 int, 나머지는 글자
                        $ok = -not $r.EOF
                        continue   # 이미 다음 칸에 있다
                    }
                } elseif ($r.NodeType -eq 'EndElement' -and $r.LocalName -eq 'row' -and $cur) {
                    $rows.Add($cur); $cur = $null
                    if ($rows.Count -ge $MaxRows) { break }
                }
                $ok = $r.Read()
            }
        } finally { $r.Close() }
        $need = -1
        foreach ($row in $rows) { foreach ($v in $row) { if ($v -is [int] -and $v -gt $need) { $need = $v } } }
        $ss = New-Object System.Collections.Generic.List[string]
        $sst = $zip.GetEntry('xl/sharedStrings.xml')
        if ($need -ge 0 -and $sst) {
            $r = [System.Xml.XmlReader]::Create($sst.Open())
            try {
                # ReadElementContentAsString, Skip은 이미 다음 칸으로 넘어가 있어 Read를 한 번 더 부르지 않는다 (부르면 </si>를 건너뛴다)
                $sb = $null
                $ok = $r.Read()
                while ($ok -and $ss.Count -le $need) {
                    if ($r.NodeType -eq 'Element' -and $r.LocalName -eq 't' -and $sb) { [void]$sb.Append($r.ReadElementContentAsString()); $ok = -not $r.EOF; continue }
                    if ($r.NodeType -eq 'Element' -and $r.LocalName -eq 'rPh') { $r.Skip(); $ok = -not $r.EOF; continue }
                    if ($r.NodeType -eq 'Element' -and $r.LocalName -eq 'si') { $sb = New-Object System.Text.StringBuilder; if ($r.IsEmptyElement) { $ss.Add(''); $sb = $null } }
                    elseif ($r.NodeType -eq 'EndElement' -and $r.LocalName -eq 'si' -and $sb) { $ss.Add($sb.ToString()); $sb = $null }
                    $ok = $r.Read()
                }
            } finally { $r.Close() }
        }
        $out = foreach ($row in $rows) {
            , @(foreach ($v in $row) { if ($v -is [int]) { if ($v -lt $ss.Count) { $ss[$v] } else { '' } } else { [string]$v } })
        }
        return , @($out)
    } finally { $zip.Dispose() }
}

function Get-XlsxKind([string]$Path) {
    $rows = Read-XlsxHead -Path $Path -MaxRows 15
    if (-not $rows.Count) { return '? 엑셀 (시트 없음)' }
    $first = @($rows[0])
    if ($first.Count -gt 0 -and "$($first[0])" -match '보고서\(\d{4}\.\d{2}\.\d{2}\.~\d{4}\.\d{2}\.\d{2}\.\)') {
        if ($rows.Count -lt 2) { return '? 네이버 (컬럼 줄 없음)' }
        return (Get-NaverKind @($rows[1] | ForEach-Object { "$_".Trim() })) + ' (엑셀)'
    }
    foreach ($row in $rows) {
        $row = @($row)
        if ($row.Count -ge 2 -and "$($row[0])" -eq '리포트 종류') { return "지마켓 $($row[1])" }
    }
    if ($first -contains '판매자기준-광고수익률') { return $(if ($first -contains '상품번호') { '옥션 AI매출업 상품별' } else { '옥션 AI매출업 일별 (엑셀 보고서용)' }) }
    if ($first -contains '광고상품번호') { return '옥션 파워클릭 상품별' }
    if (($first -contains '영역명') -and ($first -contains '키워드')) { return '옥션 파워클릭 키워드별' }
    if (($first -contains '영역명') -and ($first -contains '그룹명')) { return '옥션 파워클릭 그룹별 (엑셀 보고서용)' }
    if (($first -contains '영역명') -and ($first -contains '날짜')) { return '옥션 파워클릭 기간별 (엑셀 보고서용)' }
    return '? 인식 못 함'
}

function Get-ZipKind([string]$Path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($e in $zip.Entries) {
            if ($e.FullName -notlike '*.csv') { continue }
            $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8, $true)
            try { $line = $sr.ReadLine() } finally { $sr.Close() }
            if (Test-Gfa $line) { return 'GFA 성과 (zip)' }
            if ($line -match '네이버쇼핑 상품 ID') { return 'GFA 상품 성과 (zip)' }
        }
    } finally { $zip.Dispose() }
    return '? 인식 못 함 (zip)'
}

$out = New-Object System.Collections.Generic.List[string]
foreach ($f in (Get-ChildItem -LiteralPath $InputDir -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
    $kind = '? 인식 못 함'
    try {
        switch -Regex ($f.Extension) {
            '^\.csv$'   { $kind = Get-CsvKind (Read-Head -Path $f.FullName -Count 6) $f.Name; break }
            '^\.xlsx$'  { $kind = Get-XlsxKind $f.FullName; break }
            '^\.zip$'   { $kind = Get-ZipKind $f.FullName; break }
            '^\.html?$' { $kind = '보고서 (월별 추이용)'; break }
        }
    } catch { $kind = '? 읽지 못함' }
    $out.Add("$($f.Name)`t$kind")
}
# HTA(JScript)가 표준 출력을 읽으면 한글이 깨질 수 있어 UTF-16 파일로 넘긴다
# 다 쓴 뒤에 이름을 바꿔, HTA가 쓰는 중인 파일을 읽지 않게 한다
$tmp = "$OutFile.tmp"
[System.IO.File]::WriteAllLines($tmp, $out.ToArray(), [System.Text.Encoding]::Unicode)
Move-Item -LiteralPath $tmp -Destination $OutFile -Force
