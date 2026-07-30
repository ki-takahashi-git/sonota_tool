# JEOLデータ 日次同期スクリプト
# ①同期元 → ②対象フォルダ へ一方向ミラー(削除も反映)
$ErrorActionPreference = 'Stop'

$src = 'G:\共有ドライブ\PJ_J社共有\00_JEOL全データ'
$dst = 'G:\共有ドライブ\部署_管理_TSP\00_JEOL全ﾃﾞｰﾀ'
$logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("sync_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

# 古いログを30日分だけ残す
Get-ChildItem $logDir -Filter 'sync_*.log' |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

# 安全チェック: Googleドライブ未マウント/未同期状態でミラーすると
# ②側が誤って全削除されるおそれがあるため、件数が異常に少ない場合は中止する
if (-not (Test-Path $src)) {
    "[$(Get-Date)] ERROR: 同期元が見つかりません: $src (処理を中止しました)" | Out-File $logFile -Encoding UTF8
    exit 1
}
$srcCount = (Get-ChildItem -Path $src -Force -ErrorAction SilentlyContinue | Measure-Object).Count
if ($srcCount -lt 100) {
    "[$(Get-Date)] ERROR: 同期元の件数が異常に少ないです ($srcCount 件)。ドライブ未マウントの可能性があるため中止しました。" | Out-File $logFile -Encoding UTF8
    exit 1
}

robocopy $src $dst /MIR /Z /W:5 /R:3 /MT:16 /NP /LOG:$logFile /TEE /XF Thumbs.db desktop.ini "~$*.*"
$exitCode = $LASTEXITCODE

if ($exitCode -ge 8) {
    Write-Host "同期でエラーが発生しました。終了コード: $exitCode 詳細: $logFile"
} else {
    Write-Host "同期が完了しました。終了コード: $exitCode"
}
exit $exitCode
