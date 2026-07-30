# --- 設定エリア ---
$remotePath = "\\192.168.1.18\各部門\購買Gr"
$driveLetter = "Z:" 
$logPath = "$HOME\Desktop\RenameLog_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
$WhatIfMode = $false  # ★最初は $true でテストしてください。本番実行時は $false にする。

# 特定置換ルール（[ordered] を追加して上から順に確実に実行）
$specificRules = [ordered]@{ 
    "－" = "-"  # StrConvで全角化してしまったハイフンを半角に戻す
    "※" = "注"
    "①" = "_1"  # ユーザー指定の割り当てを維持
    "②" = "_2"
    "③" = "_3"
    "④" = "_4"
    "⑤" = "_5"
    "⑥" = "_6"
    "⑦" = "_7"
    "⑧" = "_8"
    "⑨" = "_9"
    "（" = "("
    "）" = ")"
    "【" = "("
    "】" = ")"
    "［" = "("
    "］" = ")"
    "『" = "("
    "』" = ")"
    "｛" = "("
    "｝" = ")"
    "〈" = "("
    "〉" = ")"
    "＜" = "("
    "＞" = ")"
}

Add-Type -AssemblyName Microsoft.VisualBasic

# --- 実行確認 ---
Write-Host "--- カナ全角・英数半角 ガイドライン完全準拠版 ---" -ForegroundColor Cyan
if ($WhatIfMode) {
    Write-Host "【注意】現在はテストモード（ファイル変更なし）です。" -ForegroundColor Yellow
}
$confirm = Read-Host "変換を開始しますか？ (y/n)"
if ($confirm -ne "y") { exit }

$logList = New-Object System.Collections.Generic.List[PSObject]

# 置換ロジック関数
function Get-NewName($oldName, $isFile, $extension) {
    $raw = $oldName
    if ($isFile) { 
        # 拡張子が含まれていない、または名前そのものが拡張子の場合はそのまま扱う
        if ($oldName.Length -gt $extension.Length) {
            $raw = $oldName.Substring(0, $oldName.Length - $extension.Length) 
        }
    }
    
    # 1. まず全ての半角カナを「全角」に統一する
    $raw = [Microsoft.VisualBasic.Strings]::StrConv($raw, [Microsoft.VisualBasic.VbStrConv]::Wide)
    
    # 2. 英数字だけを「半角」にする
    $convArray = foreach ($char in $raw.ToCharArray()) {
        $val = [int]$char
        if (($val -ge 0xFF10 -and $val -le 0xFF19) -or `
            ($val -ge 0xFF21 -and $val -le 0xFF3A) -or `
            ($val -ge 0xFF41 -and $val -le 0xFF5A)) {
            [Microsoft.VisualBasic.Strings]::StrConv([string]$char, [Microsoft.VisualBasic.VbStrConv]::Narrow)
        } else {
            [string]$char
        }
    }
    $raw = -join $convArray

    # 3. 特定置換（※、カッコ、ハイフン戻し）※「△」の強制変換は削除しました
    foreach ($key in $specificRules.Keys) { $raw = $raw.Replace($key, $specificRules[$key]) }

    # 4. 許可文字の判定パターン（漢字、ひらがな、カタカナ、英数字、一部記号）
    # 【ガイドライン対応】「△」を許可文字に追加し、そのまま残るように修正
    $pattern = '^[ぁ-んァ-ヶa-zA-Z0-9ー()_一-龠々\-△]$'
    
    $newNameArray = foreach ($char in $raw.ToCharArray()) {
        $c = [string]$char
        if ($c -match $pattern) {
            $c
        } else {
            "_"
        }
    }
    $res = -join $newNameArray

    # 5. クリーニング
    while ($res -like "*__*") { $res = $res.Replace("__", "_") }
    
    # 【ガイドライン対応】「先頭に記号は使用禁止」を厳格に適用
    # 名前の先頭にある スペース、アンダーバー、カッコ、ハイフン、△ をすべて自動で削り落とします
    $res = $res -replace '^[ _()\-△]+', ''
    
    # 末尾のアンダーバーを削除
    $res = $res.TrimEnd("_")

    if ([string]::IsNullOrWhiteSpace($res)) { $res = "renamed_" + (Get-Date -Format "ss") }
    if ($isFile) { return $res + $extension } else { return $res }
}

try {
    # --- ネットワーク接続 ---
    if (Test-Path $driveLetter) { 
        Remove-PSDrive -Name $driveLetter.Replace(":","") -ErrorAction SilentlyContinue 
        Start-Sleep -Seconds 1
    }
    New-PSDrive -Name $driveLetter.Replace(":","") -PSProvider FileSystem -Root $remotePath | Out-Null

    $target = $driveLetter + "\"
    $items = Get-ChildItem -Path $target -Recurse | Sort-Object FullName -Descending

    foreach ($item in $items) {
        $isDir = $item.PSIsContainer
        $ext = if ($isDir) { "" } else { $item.Extension }
        $newName = Get-NewName $item.Name (-not $isDir) $ext

        if ($item.Name -ne $newName) {
            $logList.Add([PSCustomObject]@{
                Type     = if ($isDir) { "Folder" } else { "File" }
                Original = $item.FullName
                NewName  = $newName
            })
            
            if ($WhatIfMode) {
                Write-Host "[Test] Rename: $($item.Name) -> $newName" -ForegroundColor Yellow
            } else {
                Write-Host "Rename: $($item.Name) -> $newName"
                Rename-Item -LiteralPath $item.FullName -NewName $newName
            }
        }
    }
    
    # CSV出力
    $logList | Export-Csv -Path $logPath -NoTypeInformation -Encoding Default
    Write-Host "`n完了！CSVを確認してください。 (保存先: $logPath)" -ForegroundColor Green

} catch {
    Write-Host "`nエラー: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if (Test-Path $driveLetter) { 
        Remove-PSDrive -Name $driveLetter.Replace(":","") -ErrorAction SilentlyContinue 
    }
}
pause