Attribute VB_Name = "modTXSACTExport"
Option Explicit

'====================================================
' 検収発注先ダミー変更ツール
' アクティブシートのA:F列（注番・枝番・分番・アイテムコード・
' ロット名・入庫場所）から、データ区分=Mod、工番=0、
' 実績区分=Q、発注先=1009997を固定値として付加し、
' 10列でCSV出力する
' ※シート上にはデータ区分・工番・実績区分・発注先の列は
'   存在しない（固定値のためCSV出力時にのみ付加する）
' ※ロット名は入力チェック対象外（未入力でもエラーにしない）
'====================================================

Private Const START_ROW As Long = 5
Private Const COL_PORDER As Long = 1    ' 注番 PORDER（シート上のA列）
Private Const COL_PEDA As Long = 2      ' 枝番 PEDA（シート上のB列）
Private Const COL_BUN As Long = 3       ' 分番 BUN（シート上のC列）
Private Const COL_CODE As Long = 4      ' アイテムコード CODE（シート上のD列）
Private Const COL_LOTNAME As Long = 5   ' ロット名 LOTNAME（シート上のE列、チェック対象外）
Private Const COL_HOKAN As Long = 6     ' 入庫場所 HOKAN（シート上のF列）

Private Const FIXED_DKUBU As String = "Mod"       ' データ区分 DKUBU（固定値）
Private Const FIXED_KBAN As String = "0"          ' 工番 KBAN（固定値）
Private Const FIXED_AKUBU As String = "Q"         ' 実績区分 AKUBU（固定値）
Private Const FIXED_VENDOR As String = "1009997"  ' 発注先 VENDOR（固定値）
Private Const DEFAULT_FILE_NAME As String = "TxSlipJitu.csv"

'====================================================
' CSV出力
'====================================================
Sub ExportToCSV()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim errMsg As String
    Dim hasError As Boolean
    Dim saveFileName As Variant
    Dim fNum As Integer
    Dim csvLine As String
    Dim porderBlank As Boolean, pedaBlank As Boolean, bunBlank As Boolean
    Dim codeBlank As Boolean, hokanBlank As Boolean
    Dim missing As String

    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = GetLastRow(ws)

    ' 前回のチェックで付いた色をリセット
    If lastRow >= START_ROW Then
        ws.Range(ws.Cells(START_ROW, "A"), ws.Cells(lastRow, "F")).Interior.ColorIndex = xlColorIndexNone
    End If

    If lastRow < START_ROW Then
        MsgBox "出力するデータがありません。", vbExclamation, "確認"
        Exit Sub
    End If

    ' ---- 入力チェック（値の抜け・空白行チェック）----
    ' ※ロット名（COL_LOTNAME）はチェック対象外
    hasError = False
    errMsg = "以下の行に問題があります。内容を確認してください。" & vbCrLf & vbCrLf

    For r = START_ROW To lastRow
        porderBlank = (Len(Trim(ws.Cells(r, COL_PORDER).Text)) = 0)
        pedaBlank = (Len(Trim(ws.Cells(r, COL_PEDA).Text)) = 0)
        bunBlank = (Len(Trim(ws.Cells(r, COL_BUN).Text)) = 0)
        codeBlank = (Len(Trim(ws.Cells(r, COL_CODE).Text)) = 0)
        hokanBlank = (Len(Trim(ws.Cells(r, COL_HOKAN).Text)) = 0)

        If porderBlank And pedaBlank And bunBlank And codeBlank And hokanBlank Then
            ' 完全な空白行（データの抜け）
            hasError = True
            errMsg = errMsg & r & "行目 : 空白行です（データが抜けています）" & vbCrLf
            ws.Range(ws.Cells(r, "A"), ws.Cells(r, "D")).Interior.Color = RGB(255, 199, 206)
            ws.Cells(r, "F").Interior.Color = RGB(255, 199, 206)
        ElseIf porderBlank Or pedaBlank Or bunBlank Or codeBlank Or hokanBlank Then
            ' 一部だけ未入力
            hasError = True
            missing = ""
            If porderBlank Then missing = missing & "注番 "
            If pedaBlank Then missing = missing & "枝番 "
            If bunBlank Then missing = missing & "分番 "
            If codeBlank Then missing = missing & "アイテムコード "
            If hokanBlank Then missing = missing & "入庫場所 "
            errMsg = errMsg & r & "行目 : " & Trim(missing) & "が未入力です" & vbCrLf
            If porderBlank Then ws.Cells(r, "A").Interior.Color = RGB(255, 199, 206)
            If pedaBlank Then ws.Cells(r, "B").Interior.Color = RGB(255, 199, 206)
            If bunBlank Then ws.Cells(r, "C").Interior.Color = RGB(255, 199, 206)
            If codeBlank Then ws.Cells(r, "D").Interior.Color = RGB(255, 199, 206)
            If hokanBlank Then ws.Cells(r, "F").Interior.Color = RGB(255, 199, 206)
        End If
    Next r

    If hasError Then
        MsgBox errMsg, vbCritical, "入力チェックエラー"
        Exit Sub
    End If

    ' ---- 保存先選択ダイアログ ----
    saveFileName = Application.GetSaveAsFilename( _
        InitialFileName:=DEFAULT_FILE_NAME, _
        FileFilter:="CSVファイル (*.csv), *.csv", _
        Title:="保存先を選択してください")

    If VarType(saveFileName) = vbBoolean Then
        If saveFileName = False Then Exit Sub ' キャンセル
    End If

    ' ---- CSV書き出し（Shift-JIS / CRLF）----
    fNum = FreeFile
    Open saveFileName For Output As #fNum

    ' ヘッダー1行目（日本語見出し）
    Print #fNum, """データ区分"",""注番"",""枝番"",""分番"",""工番"",""アイテムコード"",""ロット名"",""入庫場所"",""実績区分"",""発注先"""
    ' ヘッダー2行目（フィールドコード）
    Print #fNum, """DKUBU"",""PORDER"",""PEDA"",""BUN"",""KBAN"",""CODE"",""LOTNAME"",""HOKAN"",""AKUBU"",""VENDOR"""

    For r = START_ROW To lastRow
        csvLine = """" & FIXED_DKUBU & """," & _
                  """" & EscapeCsv(ws.Cells(r, COL_PORDER).Text) & """," & _
                  FormatNumField(ws.Cells(r, COL_PEDA)) & "," & _
                  FormatNumField(ws.Cells(r, COL_BUN)) & "," & _
                  FIXED_KBAN & "," & _
                  """" & EscapeCsv(ws.Cells(r, COL_CODE).Text) & """," & _
                  """" & EscapeCsv(ws.Cells(r, COL_LOTNAME).Text) & """," & _
                  """" & EscapeCsv(ws.Cells(r, COL_HOKAN).Text) & """," & _
                  """" & FIXED_AKUBU & """," & _
                  """" & FIXED_VENDOR & """"
        Print #fNum, csvLine
    Next r

    Close #fNum

    MsgBox "CSVファイルを出力しました。" & vbCrLf & saveFileName, vbInformation, "完了"

End Sub

'====================================================
' 取消（5行目以降を全クリア）
'====================================================
Sub ClearAllData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ans As VbMsgBoxResult

    Set ws = GetTargetSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = GetLastRow(ws)
    If lastRow < START_ROW Then
        MsgBox "クリアするデータがありません。", vbExclamation, "確認"
        Exit Sub
    End If

    ans = MsgBox(START_ROW & "行目から" & lastRow & "行目までのデータを" & vbCrLf & _
                 "すべて削除します。よろしいですか？", vbYesNo + vbQuestion, "全クリア確認")
    If ans <> vbYes Then Exit Sub

    With ws.Range(ws.Cells(START_ROW, "A"), ws.Cells(lastRow, "F"))
        .ClearContents
        .Interior.ColorIndex = xlColorIndexNone
    End With

    MsgBox "クリアしました。", vbInformation, "完了"

End Sub

'====================================================
' 補助関数
'====================================================
Private Function GetTargetSheet() As Worksheet
    ' ボタンが置かれているシート（実行時のアクティブシート）を対象にする
    Set GetTargetSheet = ActiveSheet
End Function

Private Function GetLastRow(ws As Worksheet) As Long
    ' ロット名は必須項目ではないため、行の存否判定には使用しない
    Dim lp As Long, le As Long, lb As Long, lc As Long, lh As Long
    lp = ws.Cells(ws.Rows.Count, COL_PORDER).End(xlUp).Row
    le = ws.Cells(ws.Rows.Count, COL_PEDA).End(xlUp).Row
    lb = ws.Cells(ws.Rows.Count, COL_BUN).End(xlUp).Row
    lc = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).Row
    lh = ws.Cells(ws.Rows.Count, COL_HOKAN).End(xlUp).Row
    GetLastRow = WorksheetFunction.Max(lp, le, lb, lc, lh, START_ROW - 1)
End Function

Private Function FormatNumField(c As Range) As String
    If Len(Trim(c.Text)) = 0 Then
        FormatNumField = ""
    ElseIf IsNumeric(c.Value) Then
        FormatNumField = CStr(CLng(c.Value))
    Else
        FormatNumField = EscapeCsv(c.Text)
    End If
End Function

Private Function EscapeCsv(s As String) As String
    EscapeCsv = Replace(s, """", """""")
End Function
