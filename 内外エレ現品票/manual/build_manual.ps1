$ErrorActionPreference = "Stop"

$outPath = "C:\Users\ki-takahashi\Desktop\Cloudeプロジェクトフォルダ\内外エレ現品票\manual\現品票作成システム_操作手順マニュアル.docx"
if (Test-Path $outPath) { Remove-Item $outPath -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

# wdBuiltinStyle constants (locale-independent numeric values)
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleTitle = -63
$wdStyleSubtitle = -75

$styleMap = @{
    "Normal"    = $wdStyleNormal
    "Heading 1" = $wdStyleHeading1
    "Heading 2" = $wdStyleHeading2
    "Title"     = $wdStyleTitle
    "Subtitle"  = $wdStyleSubtitle
}

$wdCollapseEnd = 0

function Add-Para {
    param(
        [string]$Text,
        [string]$Style = "Normal",
        [switch]$Bullet
    )
    $rng = $doc.Content
    $rng.Collapse($wdCollapseEnd)
    $rng.InsertAfter($Text + [char]13)
    # A new blank document starts with one empty trailing paragraph. InsertAfter fills that
    # paragraph with our text and creates a new empty trailing paragraph after it. So the
    # paragraph we just wrote is the second-to-last one, not the last.
    $count = $doc.Paragraphs.Count
    $paraRng = $doc.Paragraphs.Item($count - 1).Range
    $paraRng.Style = $styleMap[$Style]
    if ($Bullet) {
        $paraRng.ListFormat.ApplyBulletDefault()
    }
}

# Title
Add-Para -Text "現品票作成システム 操作手順マニュアル" -Style "Title"
Add-Para -Text "対象: 内外エレクトロニクス生産計画から現品票（QRコード付き管理No.シート）を作成・印刷するGoogleスプレッドシート" -Style "Subtitle"
Add-Para -Text ""

# 1. 概要
Add-Para -Text "1. このシステムでできること" -Style "Heading 1"
Add-Para -Text "Googleスプレッドシート「内外エレクトロニクス生産計画」から、NEI注番号をもとに型式・数量を自動取得し、QRコード付きの現品票（管理No.シート）を自動レイアウトして印刷できます。生産計画のデータをコピー＆ペーストする必要はありません。"
Add-Para -Text ""

# 2. 用語
Add-Para -Text "2. 用語・データの場所" -Style "Heading 1"
Add-Para -Text "NEI注番号: 生産計画シートのC列「NEI注文番号」。現品票の下段に印字される番号です（例: JJFK25008650）。B列の「注番」ではない点に注意してください。" -Bullet
Add-Para -Text "型式: 生産計画シートのD列。" -Bullet
Add-Para -Text "連番（管理No.）: 生産計画シートのL列「管理No.」の値をそのまま使用します。自動採番はしません。同じNEI注番号が複数行に分かれて登録されている場合は、それぞれの行のL列の値をすべて取得します（行数がそのまま数量になります）。現品票の中央に大きく表示され、QRコードにも含まれます。印刷時は3桁ゼロ埋め（例: 1→001）で表示されます。" -Bullet
Add-Para -Text ""

# 3. 日常の操作手順
Add-Para -Text "3. 日常の操作手順" -Style "Heading 1"

Add-Para -Text "手順1: 今月の入力シートを開く" -Style "Heading 2"
Add-Para -Text "スプレッドシート上部メニュー「現品票」→「今月の入力シートを開く」を実行します。「入力_202607」のように、月ごとのシートが無ければ自動的に作成されます。"
Add-Para -Text ""

Add-Para -Text "手順2: NEI注番号を入力する" -Style "Heading 2"
Add-Para -Text "入力シートのA列（2行目以降）に、現品票を作りたいNEI注番号を1行に1つずつ入力します。生産計画シートのC列の値をそのまま入力してください。"
Add-Para -Text ""

Add-Para -Text "手順3: 連番を記載" -Style "Heading 2"
Add-Para -Text "メニュー「現品票」→「連番を記載」を実行します。まだB列に連番が無い行すべてに、生産計画シートのL列「管理No.」の値がそのまま取得されます（数量が2以上の場合は「051-058」のような範囲で表示されます）。"
Add-Para -Text "この時点ではまだ印刷用シートは作られません。何度でも実行でき、既に連番が記載済みの行は再度取得されません。" -Bullet
Add-Para -Text ""

Add-Para -Text "手順4: 印刷したい行にチェックを入れる" -Style "Heading 2"
Add-Para -Text "入力シートのC列「印刷対象」にチェックを入れます。今回印刷したい行だけにチェックを入れてください。"
Add-Para -Text "複数行まとめてチェックしたい場合は、範囲をドラッグで選択してからキーボードの[スペースキー]を押すと、選択範囲すべてを一括でON/OFFできます。" -Bullet
Add-Para -Text "メニュー「現品票」→「印刷対象をすべてチェック」「印刷対象をすべて解除」でも一括操作できます。" -Bullet
Add-Para -Text ""

Add-Para -Text "手順5: チェックした行を印刷シートに出力する" -Style "Heading 2"
Add-Para -Text "メニュー「現品票」→「チェックした行を印刷シートに出力」を実行します。チェックが入っている行だけが「印刷」シートに現品票レイアウトとして出力されます。"
Add-Para -Text ""

Add-Para -Text "手順6: 印刷する" -Style "Heading 2"
Add-Para -Text "「印刷」シートを開き、内容を確認してから [ファイル] → [印刷] を選び、以下の設定で印刷します。"
Add-Para -Text "用紙サイズ: A4" -Bullet
Add-Para -Text "ページの向き: 横向き" -Bullet
Add-Para -Text "余白: 上0.5cm・他はなし（0cm）" -Bullet
Add-Para -Text "拡大縮小: 標準（100%）※「幅に合わせる」等の自動スケールは選ばないこと" -Bullet
Add-Para -Text "この設定は一度行えばシートに保存され、次回以降は設定し直す必要はありません。"
Add-Para -Text "拡大縮小を100%以外にすると、14行（2ブロック）ごとの改ページ位置がズレてしまうので注意してください。"
Add-Para -Text ""

# 4. 便利な機能
Add-Para -Text "4. 便利な機能" -Style "Heading 1"

Add-Para -Text "同じ行を後日もう一度印刷したい場合" -Style "Heading 2"
Add-Para -Text "B列に既に連番が入っている行のC列にチェックを入れて「チェックした行を印刷シートに出力」を実行すると、新しい番号を振り直さず、既存の連番のまま再印刷できます。"
Add-Para -Text ""

# 5. 注意点・こんな時は
Add-Para -Text "5. 注意点・こんな時は" -Style "Heading 1"

Add-Para -Text "B列に「エラー:重複」と表示された" -Style "Heading 2"
Add-Para -Text "同じNEI注番号がこの入力シート内に複数回入力されている（コピー＆ペーストミス等）ことを示しています。この状態では連番は記載されません。"
Add-Para -Text "対応: 重複している行のどちらかを削除するか内容を修正し、B列の「エラー:重複」の表示を手動で消してから、「連番を記載」をやり直してください。" -Bullet
Add-Para -Text ""

Add-Para -Text "生産計画シート内に見つからないNEI注番号がある" -Style "Heading 2"
Add-Para -Text "実行結果のメッセージに一覧表示されます。生産計画シート側にまだ登録されていないか、NEI注番号の入力間違いが考えられます。生産計画側を確認し、正しい値に修正してから再実行してください。"
Add-Para -Text ""

Add-Para -Text "印刷シートの作成に時間がかかる" -Style "Heading 2"
Add-Para -Text "QRコード画像の挿入は件数に比例して時間がかかります（Google Sheetsの仕様上、1枚ずつしか挿入できないため）。件数が多いほど時間がかかりますが、正常な動作です。"
Add-Para -Text ""

$doc.SaveAs([ref]$outPath, [ref]16)  # 16 = wdFormatDocumentDefault (.docx)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "Saved: $outPath"
