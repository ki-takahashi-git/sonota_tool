Attribute VB_Name = "単価メンテ設定"
Option Explicit

' 各シートに操作ボタンを自動設置する（1回実行すればOK）
' 事前に「単価メンテ_マクロ.bas」「単価メンテ_CSV.bas」をインポートしておくこと。
Sub ボタン設置()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("抽出")
    ws.Buttons.Delete
    AddBtn ws, "単価検索", "① 単価検索", 300, 3
    AddBtn ws, "クリア", "クリア", 392, 3

    Set ws = ThisWorkbook.Worksheets("単価結果")
    ws.Buttons.Delete
    AddBtn ws, "結果を一括登録へ", "② 結果を一括登録へ", 420, 2

    Set ws = ThisWorkbook.Worksheets("一括登録")
    ws.Buttons.Delete
    AddBtn ws, "CSV生成", "④ CSV生成", 470, 2
    On Error GoTo 0

    MsgBox "各シートにボタンを設置しました。" & vbCrLf & _
           "抽出：①単価検索／クリア　単価結果：②結果を一括登録へ　一括登録：④CSV生成", vbInformation
End Sub

Private Sub AddBtn(ws As Worksheet, ByVal macroName As String, ByVal caption As String, _
                   ByVal x As Single, ByVal y As Single)
    Dim b As Object
    Set b = ws.Buttons.Add(x, y, 150, 24)
    b.OnAction = macroName
    b.Caption = caption
    b.Font.Size = 10
End Sub
