<#
.SYNOPSIS
  레포디 자동 업데이트 (2.035). 공개 배포 저장소(hojong90-creator/repodi)의 zip을 받아 이 PC의 레포디 폴더(-AppDir)에 덮어쓴다.
  레포디.hta가 일주일에 한 번 version.txt를 읽고, 지금보다 높을 때만 이 스크립트를 창 없이 돌린다.
  결과는 -OutFile에 'ok' 또는 'fail<탭>이유'로 적는다 (UTF-16, HTA가 그대로 읽는다).
.NOTES
  '보고서 생성용 데이터', '생성된 보고서', '.last_author.txt', '개발용(공유x)'는 덮어쓰지 않는다(배포 저장소에도 없다).
  배포 저장소에서 없어진 파일은 지우지 않는다. 받은 zip의 version.txt가 HTA가 확인한 버전과 다르면 아무것도 바꾸지 않는다.
#>
param(
    [Parameter(Mandatory = $true)][string]$AppDir,
    [Parameter(Mandatory = $true)][string]$Version,
    [Parameter(Mandatory = $true)][string]$OutFile
)

$ErrorActionPreference = "Stop"
$ZipUrl = "https://codeload.github.com/hojong90-creator/repodi/zip/refs/heads/main"

function Write-Result([string]$Text) { [IO.File]::WriteAllText($OutFile, $Text, [Text.Encoding]::Unicode) }

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("repodi_update_" + [guid]::NewGuid().ToString("N"))
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    [void](New-Item -ItemType Directory -Path $tmp)
    $zip = Join-Path $tmp "repodi.zip"
    $wc = New-Object Net.WebClient
    try { $wc.DownloadFile($ZipUrl, $zip) } finally { $wc.Dispose() }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $ex = Join-Path $tmp "x"
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $ex)
    # GitHub zip은 맨 위에 'repodi-main' 폴더 하나가 있다
    $root = @(Get-ChildItem -LiteralPath $ex -Directory)[0]
    if (-not $root -or -not (Test-Path -LiteralPath (Join-Path $root.FullName "레포디.hta"))) { throw "받은 파일에 레포디.hta가 없습니다." }
    $verFile = Join-Path $root.FullName "version.txt"
    $got = if (Test-Path -LiteralPath $verFile) { ([IO.File]::ReadAllText($verFile)).Trim() } else { "" }
    if ($got -ne $Version) { throw "받은 파일의 버전($got)이 확인한 버전($Version)과 다릅니다. 잠시 뒤 다시 시도합니다." }

    $skip = '^(보고서 생성용 데이터|생성된 보고서|개발용\(공유x\))(\\|$)|^\.last_author\.txt$'
    $base = $root.FullName.TrimEnd('\') + '\'
    # 레포디.hta는 맨 마지막에 바꾼다 (도중에 실패해도 창은 예전 버전 그대로 열리게)
    $files = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -File | Sort-Object { $_.Name -eq "레포디.hta" })
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($base.Length)
        if ($rel -match $skip) { continue }
        $dest = Join-Path $AppDir $rel
        $dir = Split-Path $dest -Parent
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
    }
    Write-Result "ok"
} catch {
    Write-Result ("fail`t" + $_.Exception.Message)
} finally {
    try { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
