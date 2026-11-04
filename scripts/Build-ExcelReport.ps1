# 지마켓 · 옥션 리포트 xlsx를 읽어 '구매전환 리포트' 엑셀 보고서(매체별 1개)를 만든다.
# 서식 · 수식 · 차트는 templates\excel-*.template.xlsx에 들어 있고, 여기서는 표의 행 수를 맞춘 뒤 값만 채운다.
# 차트가 들어간 xlsx를 만들어야 해서 이 스크립트만은 Excel(COM)이 설치된 PC에서 동작한다.
#
# 종료 코드: 0 = 생성 완료(만든 파일 경로를 한 줄씩 출력), 2 = 지마켓/옥션 리포트 파일 없음, 1 = 오류
param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [Parameter(Mandatory = $true)][string]$TemplateDir,
    [string]$ClientName = ""
)
$ErrorActionPreference = 'Stop'

function ToNum($v) {
    if ($null -eq $v) { return 0.0 }
    if ($v -is [double]) { return $v }
    $n = 0.0
    if ([double]::TryParse(("$v" -replace '[,%\s]', ''), [ref]$n)) { return $n }
    return 0.0
}

function ToDate($v) {
    # 옥션 파워클릭은 날짜가 엑셀 일련번호(46266), 나머지는 '2026-09-01' 문자열로 들어온다.
    if ($v -is [double]) { return [DateTime]::FromOADate($v) }
    $d = [DateTime]::MinValue
    if ([DateTime]::TryParse("$v", [ref]$d)) { return $d }
    return $null
}

function Read-Sheet($xl, [string]$Path) {
    # 첫 시트를 행 배열(각 행은 값 배열)로 읽는다.
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    try {
        $v = $wb.Worksheets.Item(1).UsedRange.Value2
        $rows = New-Object System.Collections.Generic.List[object]
        if ($v -is [array]) {
            for ($r = 1; $r -le $v.GetLength(0); $r++) {
                $row = New-Object object[] $v.GetLength(1)
                for ($c = 1; $c -le $v.GetLength(1); $c++) { $row[$c - 1] = $v[$r, $c] }
                $rows.Add($row)
            }
        }
        return , $rows
    } finally { $wb.Close($false) }
}

function To-Records($rows, [int]$headerIdx) {
    # 헤더 행 아래를 '컬럼명 → 값' 해시테이블 목록으로 바꾼다.
    $headers = $rows[$headerIdx]
    $list = New-Object System.Collections.Generic.List[object]
    for ($i = $headerIdx + 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if ($null -eq $r[0] -or "$($r[0])" -eq '') { continue }
        $h = @{}
        for ($c = 0; $c -lt $headers.Count; $c++) { if ($headers[$c]) { $h["$($headers[$c])"] = $r[$c] } }
        $list.Add($h)
    }
    return , $list
}

function Sum-By($records, [string[]]$keyCols, [string[]]$numCols) {
    # keyCols가 같은 행끼리 numCols를 합친다. 결과 행 = 키 값들 + 합계들 (처음 나온 순서 유지).
    $map = [ordered]@{}
    foreach ($rec in $records) {
        $key = ($keyCols | ForEach-Object { "$($rec[$_])" }) -join '|'
        if (-not $map.Contains($key)) {
            $map[$key] = @{ keys = @($keyCols | ForEach-Object { $rec[$_] }); nums = (New-Object double[] $numCols.Count) }
        }
        for ($i = 0; $i -lt $numCols.Count; $i++) { $map[$key].nums[$i] += (ToNum $rec[$numCols[$i]]) }
    }
    return , @($map.Values)
}

function Sort-Rows($items, [int[]]$numOrder) {
    # 완성본과 같은 정렬: 전환금액 → 광고비 → 클릭수 → 노출수 내림차순. numOrder는 nums 안의 위치.
    $props = @($numOrder | ForEach-Object { $i = $_; @{ Expression = [scriptblock]::Create("`$_.nums[$i]"); Descending = $true } })
    return , @($items | Sort-Object -Property $props)
}

function Daily-Rows($records, [string]$dateCol, [string[]]$numCols, [DateTime]$from, [DateTime]$to) {
    # 기간의 모든 날짜를 한 줄씩 만든다 (데이터가 없는 날은 0).
    $map = @{}
    foreach ($rec in $records) {
        $d = ToDate $rec[$dateCol]
        if (-not $d) { continue }
        $k = $d.ToString('yyyyMMdd')
        if (-not $map.ContainsKey($k)) { $map[$k] = (New-Object double[] $numCols.Count) }
        for ($i = 0; $i -lt $numCols.Count; $i++) { $map[$k][$i] += (ToNum $rec[$numCols[$i]]) }
    }
    $out = @()
    for ($d = $from; $d -le $to; $d = $d.AddDays(1)) {
        $k = $d.ToString('yyyyMMdd')
        $nums = if ($map.ContainsKey($k)) { $map[$k] } else { New-Object double[] $numCols.Count }
        $out += , @{ keys = @($d.ToOADate()); nums = $nums }
    }
    return , $out
}

function Fill-Table($ws, [int]$first, $items, [string[]]$keyCols, [string[]]$numCols, [string]$colFrom, [string]$colTo, [bool]$textKeys = $true) {
    # 템플릿의 표는 데이터 행이 2줄(첫 행 · 마지막 행)이다. 그 사이에 행을 넣어 개수를 맞추면
    # 합계 SUM 범위와 차트 범위가 함께 늘어난다. 수식 컬럼은 첫 행의 수식을 아래로 복사하고, 나머지는 값을 넣는다.
    $n = $items.Count
    if ($n -eq 0) { return }
    if ($n -eq 1) { $ws.Range("A$($first + 1)").EntireRow.Delete() | Out-Null }
    if ($n -gt 2) { $ws.Range("A$($first + 1):A$($first + $n - 2)").EntireRow.Insert(-4121, 0) | Out-Null }
    $last = $first + $n - 1

    foreach ($cell in $ws.Range("$colFrom${first}:$colTo$first").Cells) {
        if ($cell.HasFormula) {
            $col = $cell.Address($false, $false) -replace '\d', ''
            $ws.Range("$col${first}:$col$last").FormulaR1C1 = $cell.FormulaR1C1
        }
    }
    for ($k = 0; $k -lt $keyCols.Count + $numCols.Count; $k++) {
        $isKey = $k -lt $keyCols.Count
        $col = if ($isKey) { $keyCols[$k] } else { $numCols[$k - $keyCols.Count] }
        if (-not $col) { continue }
        $arr = New-Object 'object[,]' $n, 1
        for ($i = 0; $i -lt $n; $i++) {
            $arr[$i, 0] = if (-not $isKey) { $items[$i].nums[$k - $keyCols.Count] } elseif ($textKeys) { "$($items[$i].keys[$k])" } else { $items[$i].keys[$k] }
        }
        $rng = $ws.Range("$col${first}:$col$last")
        if ($isKey -and $textKeys) { $rng.NumberFormat = '@' }   # '2.5' 같은 키워드나 상품번호가 숫자로 바뀌지 않게
        $rng.Value2 = $arr
    }
}

function Get-OutputPath([string]$media, [DateTime]$month) {
    # 만든 뒤에 지난달 수치 등을 손으로 채워 넣는 파일이라, 같은 이름이 있으면 덮어쓰지 않고 _v2, _v3...로 저장한다.
    $client = ($ClientName -replace '[\\/:*?"<>|]', '').Trim()
    if (-not $client) { $client = '업체명' }
    $base = "(주)예지솔루션_${media}_구매전환 리포트_${client}_$($month.ToString('yyyyMM'))"
    $path = Join-Path $OutputDir "$base.xlsx"
    for ($v = 2; Test-Path -LiteralPath $path; $v++) { $path = Join-Path $OutputDir "${base}_v$v.xlsx" }
    return $path
}

function Set-MonthLabels($ws, [int]$firstRow, [DateTime]$month) {
    for ($i = 0; $i -lt 3; $i++) {
        $m = $month.AddMonths($i - 2)
        $ws.Range("B$($firstRow + $i)").Value2 = "$($m.Year)년 $($m.Month)월"
    }
}

$xl = $null
$created = @()
try {
    try { $xl = New-Object -ComObject Excel.Application }
    catch { throw "엑셀 보고서를 만들려면 이 PC에 Excel이 설치되어 있어야 합니다." }
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # ---------- 입력 파일 분류 ----------
    # 지마켓은 파일 안의 '리포트 종류' 값으로, 옥션은 1행 헤더의 컬럼명으로 구분한다. 같은 종류가 여러 개면 첫 파일만 쓴다.
    $gm = @{}; $au = @{}
    $gmPeriod = $null
    foreach ($f in (Get-ChildItem -LiteralPath $InputDir -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' })) {
        $rows = Read-Sheet $xl $f.FullName
        if ($rows.Count -lt 2) { continue }
        $h = @($rows[0] | ForEach-Object { "$_" })

        if ($h[0] -eq '기간') {
            $type = $null; $headerIdx = -1
            for ($i = 0; $i -lt [math]::Min($rows.Count, 15); $i++) {
                if ("$($rows[$i][0])" -eq '리포트 종류') { $type = "$($rows[$i][1])" }
                if (@($rows[$i]) -contains '노출 수') { $headerIdx = $i; break }
            }
            if ($headerIdx -lt 0) { continue }
            $key = if ($type -match '캠페인') { 'campaign' } elseif ($type -match '그룹') { 'group' } elseif ($type -match '상품') { 'item' } elseif ($type -match '키워드') { 'keyword' } else { $null }
            if ($key -and -not $gm.ContainsKey($key)) {
                $gm[$key] = To-Records $rows $headerIdx
                if (-not $gmPeriod -and "$($rows[0][1])" -match '(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})') {
                    $gmPeriod = @([DateTime]$matches[1], [DateTime]$matches[2])
                }
            }
            continue
        }

        $key = $null
        if ($h -contains '판매자기준-광고수익률') { $key = if ($h -contains '상품번호') { 'aiGoods' } else { 'aiDaily' } }
        elseif ($h -contains '광고상품번호') { $key = 'goods' }
        elseif ($h -contains '영역명' -and $h -contains '키워드') { $key = 'keyword' }
        elseif ($h -contains '영역명' -and $h -contains '그룹명') { $key = 'group' }
        elseif ($h -contains '영역명' -and $h -contains '날짜') { $key = 'daily' }
        if ($key -and -not $au.ContainsKey($key)) { $au[$key] = To-Records $rows 0 }
    }

    if ($gm.Count -eq 0 -and $au.Count -eq 0) {
        [Console]::Error.WriteLine("지마켓 · 옥션 리포트 파일(xlsx)을 찾지 못했습니다.")
        exit 2
    }
    if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
    $std = '노출 수', '클릭 수', '판매자 전환 수', '광고 비용', '판매자 전환 금액'

    # ---------- 지마켓 ----------
    if ($gm.Count -gt 0) {
        if (-not $gmPeriod) { throw "지마켓 리포트에서 기간을 읽지 못했습니다." }
        $from = $gmPeriod[0]; $to = $gmPeriod[1]
        $wb = $xl.Workbooks.Open((Join-Path $TemplateDir 'excel-gmarket.template.xlsx'), 0, $true)

        # 기간별: 캠페인타입(통합운영형 = AI매출업, 직접운영형, 집중운영형)마다 시트 하나.
        # '기간별(파워클릭)'과 '월간 사이트 요약'의 일별 표는 이 시트들을 더하는 수식이라 날짜만 넣는다.
        $camp = if ($gm.ContainsKey('campaign')) { $gm['campaign'] } else { @() }
        foreach ($pair in @(@('기간별(AI매출업)', '통합운영형'), @('기간별(직접운영형)', '직접운영형'), @('기간별(집중운영형)', '집중운영형'))) {
            $recs = @($camp | Where-Object { "$($_['캠페인타입'])" -eq $pair[1] })
            $items = Daily-Rows $recs '날짜' $std $from $to
            Fill-Table $wb.Worksheets.Item($pair[0]) 23 $items @('B') @('C', 'D', 'E', 'H', 'I') 'B' 'L' $false
        }
        $dates = Daily-Rows @() '날짜' @() $from $to
        Fill-Table $wb.Worksheets.Item('기간별(파워클릭)') 23 $dates @('B') @() 'B' 'L' $false
        $ws = $wb.Worksheets.Item('월간 사이트 요약')
        Fill-Table $ws 32 $dates @('B') @() 'B' 'L' $false
        $ws.Range('B18').Value2 = "지난달 지마켓 현황 ($($from.ToString('yyyy-MM-dd')) ~ $($to.ToString('yyyy-MM-dd')))"
        Set-MonthLabels $ws 23 $from

        if ($gm.ContainsKey('group')) {
            $items = Sort-Rows (Sum-By $gm['group'] @('캠페인명', '그룹명') $std) @(4, 3, 1, 0)
            Fill-Table $wb.Worksheets.Item('그룹별 리포트') 8 $items @('B', 'C') @('D', 'E', 'F', 'I', 'J') 'B' 'M'
        }
        if ($gm.ContainsKey('item')) {
            $items = Sort-Rows (Sum-By $gm['item'] @('캠페인명', '그룹명', '상품명', '상품번호') $std) @(4, 3, 1, 0)
            Fill-Table $wb.Worksheets.Item('상품별 리포트') 8 $items @('B', 'C', 'D', 'E') @('F', 'G', 'H', 'K', 'L') 'B' 'O'
        }
        if ($gm.ContainsKey('keyword')) {
            $items = Sort-Rows (Sum-By $gm['keyword'] @('그룹명', '상품명', '상품번호', '키워드') $std) @(4, 3, 1, 0)
            Fill-Table $wb.Worksheets.Item('키워드별 리포트') 8 $items @('B', 'C', 'D', 'E') @('F', 'G', 'H', 'K', 'L') 'B' 'O'
        }

        $out = Get-OutputPath '지마켓' $from
        $wb.Worksheets.Item(1).Activate()
        $wb.SaveAs($out, 51)
        $wb.Close($false)
        $created += $out
    }

    # ---------- 옥션 ----------
    if ($au.Count -gt 0) {
        # 옥션 리포트에는 기간 정보가 없어 일별 리포트의 날짜 범위를 쓴다. 일별 리포트가 없으면 지난달로 본다.
        $allDates = @()
        foreach ($pair in @(@('daily', '날짜'), @('aiDaily', '날짜별'))) {
            if ($au.ContainsKey($pair[0])) { $allDates += @($au[$pair[0]] | ForEach-Object { ToDate $_[$pair[1]] } | Where-Object { $_ }) }
        }
        if ($allDates.Count -gt 0) {
            $sorted = @($allDates | Sort-Object)
            $from = $sorted[0].Date; $to = $sorted[-1].Date
        } else {
            $from = (Get-Date -Day 1).Date.AddMonths(-1); $to = $from.AddMonths(1).AddDays(-1)
        }
        $pc = '노출수', '클릭수', '구매수', '총비용', '구매금액'
        $ai = '클릭수', '총비용', '광고상품기준-주문수', '광고상품기준-구매수', '광고상품기준-구매금액', '판매자기준-주문수', '판매자기준-구매수', '판매자기준-구매금액'
        $wb = $xl.Workbooks.Open((Join-Path $TemplateDir 'excel-auction.template.xlsx'), 0, $true)

        $ws = $wb.Worksheets.Item('월간 사이트 요약')
        $ws.Range('B20').Value2 = "지난달 옥션 현황 ($($from.ToString('yyyy-MM-dd')) ~ $($to.ToString('yyyy-MM-dd')))"
        Set-MonthLabels $ws 25 $from
        Set-MonthLabels $ws 53 $from

        $recs = if ($au.ContainsKey('daily')) { $au['daily'] } else { @() }
        Fill-Table $wb.Worksheets.Item('기간별(파워클릭)') 23 (Daily-Rows $recs '날짜' $pc $from $to) @('B') @('C', 'D', 'E', 'H', 'I') 'B' 'L' $false
        if ($au.ContainsKey('group')) {
            $items = Sort-Rows (Sum-By $au['group'] @('그룹명') $pc) @(4, 3, 1, 0)
            Fill-Table $wb.Worksheets.Item('그룹별(파워클릭)') 8 $items @('B') @('C', 'D', 'E', 'H', 'I') 'B' 'L'
        }
        if ($au.ContainsKey('goods')) {
            $items = Sort-Rows (Sum-By $au['goods'] @('광고상품번호') $pc) @(4, 3, 1, 0)
            Fill-Table $wb.Worksheets.Item('상품별(파워클릭)') 8 $items @('B') @('C', 'D', 'E', 'H', 'I') 'B' 'L'
        }
        if ($au.ContainsKey('keyword')) {
            # 완성본 기준: 클릭이 1회 이상 있었던 키워드만 싣는다.
            $all = Sum-By $au['keyword'] @('키워드') $pc
            $items = @($all | Where-Object { $_.nums[1] -gt 0 })
            Fill-Table $wb.Worksheets.Item('키워드별(파워클릭)') 8 (Sort-Rows $items @(4, 3, 1, 0)) @('B') @('C', 'D', 'E', 'H', 'I') 'B' 'L'
        }

        $recs = if ($au.ContainsKey('aiDaily')) { $au['aiDaily'] } else { @() }
        Fill-Table $wb.Worksheets.Item('기간별(AI매출업)') 24 (Daily-Rows $recs '날짜별' $ai $from $to) @('B') @('C', 'E', 'F', 'G', 'H', 'I', 'J', 'K') 'B' 'M' $false
        if ($au.ContainsKey('aiGoods')) {
            $items = Sort-Rows (Sum-By $au['aiGoods'] @('판매자 ID', '그룹명', '상품번호') $ai) @(7, 1, 0)
            Fill-Table $wb.Worksheets.Item('상품별(AI매출업)') 9 $items @('B', 'C', 'D') @('E', 'G', 'H', 'I', 'J', 'K', 'L', 'M') 'B' 'O'
            # 그룹별(AI매출업)은 옥션에서 받는 리포트가 따로 없어, 수치는 기간별 합계를 참조하고 이름만 채운다.
            if ($items.Count -gt 0) {
                $wsG = $wb.Worksheets.Item('그룹별(AI매출업)')
                $wsG.Range('B9').Value2 = "$($items[0].keys[0])"
                $wsG.Range('C9').Value2 = "$($items[0].keys[1])"
            }
        }

        $out = Get-OutputPath '옥션' $from
        $wb.Worksheets.Item(1).Activate()
        $wb.SaveAs($out, 51)
        $wb.Close($false)
        $created += $out
    }
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($xl) {
        try { $xl.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
}
$created | ForEach-Object { Write-Output $_ }
exit 0
