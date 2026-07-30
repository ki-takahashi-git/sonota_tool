Attribute VB_Name = "�P�������e"
Option Explicit

' ===== �P���}�X�^�[ �����e�i���X�c�[�� �}�N�� =====
' �Q�l: �R�[�h��񌟍��c�[��Ver1.02�iADO�ڑ������j�𓥏P
' �u���o�v�V�[�g�� ���[�J�[�i�ڃR�[�h(A��) �� �A�C�e���R�[�h(B��) ��\��t�� ��
'  �u�P�������v�����s �� �u�P�����ʁv�V�[�g�Ɍ��s�̒P������\���B
' DB�ւ� SELECT �̂݁i�ǎ���p���[�U�[ tdread�j�B

' --- �ڑ���� ---
Private Const SERVER_IP As String = "YOUR_DB_SERVER_IP"
Private Const DB_NAME As String = "TPICSDB_T"
Private Const DB_USER As String = "YOUR_DB_USER"
Private Const DB_PASS As String = "YOUR_DB_PASSWORD"

Sub �P������()
    Dim conn As Object, rs As Object
    Dim wsIn As Worksheet, wsOut As Worksheet
    Dim lastA As Long, lastB As Long, i As Long, r As Long, outRow As Long
    Dim strConn As String, sql As String, whereClause As String
    Dim dbData As Variant, hasData As Boolean, matchCount As Long

    Set wsIn = ThisWorkbook.Worksheets("���o")
    Set wsOut = ThisWorkbook.Worksheets("�P������")

    lastA = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    lastB = wsIn.Cells(wsIn.Rows.Count, 2).End(xlUp).Row

    If lastA < 3 And lastB < 3 Then
        MsgBox "�u���o�v�V�[�g��3�s�ڂ���A���[�J�[�i�ڃR�[�h �� �A�C�e���R�[�h����͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    If lastA >= 3 And lastB >= 3 Then
        MsgBox "A���B��̗����ɓ��͂�����܂��B�Е��̗񂾂��ɂ��Ă��������B", vbCritical
        Exit Sub
    End If

    ' --- WHERE���g�ݗ��� ---
    whereClause = ""
    If lastA >= 3 Then
        For i = 3 To lastA
            Dim rawA As String: rawA = Trim(CStr(wsIn.Cells(i, 1).Value))
            If rawA <> "" Then
                whereClause = whereClause & "H.MCODE LIKE '" & Replace(Clean�i��(rawA), "'", "''") & "%' OR "
            End If
        Next i
    Else
        For i = 3 To lastB
            Dim rawB As String: rawB = Trim(CStr(wsIn.Cells(i, 2).Value))
            If rawB <> "" Then
                whereClause = whereClause & "H.CODE = '" & Replace(rawB, "'", "''") & "' OR "
            End If
        Next i
    End If
    If whereClause = "" Then Exit Sub
    whereClause = Left(whereClause, Len(whereClause) - 4)

    ' --- DB�ڑ� ---
    strConn = "Provider=SQLOLEDB;Data Source=" & SERVER_IP & _
              ";Initial Catalog=" & DB_NAME & ";User ID=" & DB_USER & _
              ";Password=" & DB_PASS & ";"
    On Error GoTo DBERR
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn
    conn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"

    sql = "SELECT RTRIM(H.CODE) AS CODE, RTRIM(H.MCODE) AS MCODE, RTRIM(H.NAME) AS NAME, " & _
          "RTRIM(ISNULL(S1.BNAME,'')) AS MAKER, RTRIM(ISNULL(T.VENDOR,'')) AS VENDOR, " & _
          "RTRIM(ISNULL(SV.BNAME,'')) AS VENDOR_NAME, ISNULL(CAST(T.TID AS varchar(20)),'') AS TID, " & _
          "ISNULL(CAST(T.TDATE AS varchar(20)),'') AS TDATE, ISNULL(CAST(T.EDATE AS varchar(20)),'') AS EDATE, " & _
          "ISNULL(CAST(T.TVOL AS varchar(30)),'') AS TVOL, ISNULL(CAST(T.PRICE AS varchar(30)),'') AS PRICE, " & _
          "ISNULL(CAST(I.LOTS AS varchar(30)),'') AS MOQ " & _
          "FROM XHEAD H WITH (NOLOCK) " & _
          "LEFT JOIN XSECT S1 WITH (NOLOCK) ON H.MAKER = S1.BUMO " & _
          "LEFT JOIN XTANK T WITH (NOLOCK) ON T.CODE = H.CODE AND T.VENDOR = H.MAINBUMO " & _
          "LEFT JOIN XSECT SV WITH (NOLOCK) ON T.VENDOR = SV.BUMO " & _
          "LEFT JOIN XITEM I WITH (NOLOCK) ON H.CODE = I.CODE AND H.MAINBUMO = I.BUMO " & _
          "WHERE (" & whereClause & ") AND H.CODE NOT LIKE 'TD-%' " & _
          "ORDER BY H.CODE, T.VENDOR, T.TVOL, T.TDATE, T.EDATE"

    Set rs = conn.Execute(sql)
    hasData = False
    If Not rs.EOF Then dbData = rs.GetRows: hasData = True
    rs.Close: conn.Close
    On Error GoTo 0

    ' --- ���ʃV�[�g������ ---
    Dim lastRes As Long: lastRes = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If lastRes >= 3 Then wsOut.Rows("3:" & lastRes).Clear
    outRow = 3
    Application.ScreenUpdating = False

    If lastA >= 3 Then
        ' A��i���[�J�[�i�ڃR�[�h�E�O����v�j���ɏo��
        For i = 3 To lastA
            Dim keyA As String: keyA = Clean�i��(Trim(CStr(wsIn.Cells(i, 1).Value)))
            If keyA <> "" Then
                matchCount = 0
                If hasData Then
                    For r = 0 To UBound(dbData, 2)
                        If CStr(dbData(1, r)) Like keyA & "*" Then
                            If PassFilter(CStr(dbData(8, r)), CStr(dbData(9, r))) Then
                                Call �o��1�s(wsOut, outRow, dbData, r): outRow = outRow + 1: matchCount = matchCount + 1
                            End If
                        End If
                    Next r
                End If
                If matchCount = 0 Then
                    wsOut.Cells(outRow, 2).Value = Trim(CStr(wsIn.Cells(i, 1).Value))
                    wsOut.Cells(outRow, 12).Value = "�Y���Ȃ�"
                    outRow = outRow + 1
                End If
            End If
        Next i
    Else
        ' B��i�A�C�e���R�[�h�E���S��v�j���ɏo��
        For i = 3 To lastB
            Dim keyB As String: keyB = Trim(CStr(wsIn.Cells(i, 2).Value))
            If keyB <> "" Then
                matchCount = 0
                If hasData Then
                    For r = 0 To UBound(dbData, 2)
                        If CStr(dbData(0, r)) = keyB Then
                            If PassFilter(CStr(dbData(8, r)), CStr(dbData(9, r))) Then
                                Call �o��1�s(wsOut, outRow, dbData, r): outRow = outRow + 1: matchCount = matchCount + 1
                            End If
                        End If
                    Next r
                End If
                If matchCount = 0 Then
                    wsOut.Cells(outRow, 1).Value = keyB
                    wsOut.Cells(outRow, 12).Value = "�Y���Ȃ�"
                    outRow = outRow + 1
                End If
            End If
        Next i
    End If

    If outRow > 3 Then
        wsOut.Range("A3:A" & outRow - 1).HorizontalAlignment = xlLeft
        wsOut.Range("G3:G" & outRow - 1).HorizontalAlignment = xlLeft
    End If
    wsOut.Columns("A:L").AutoFit
    Application.ScreenUpdating = True
    wsOut.Select
    MsgBox "�������������܂����i" & (outRow - 3) & " �s�j�B", vbInformation
    Exit Sub

DBERR:
    Application.ScreenUpdating = True
    MsgBox "DB�ڑ��܂��͌����ŃG���[���������܂����B" & vbCrLf & Err.Description, vbCritical
End Sub

' 1�s���u�P�����ʁv�ɏ����o��
Private Sub �o��1�s(ws As Worksheet, ByVal outRow As Long, dbData As Variant, ByVal r As Long)
    Dim tid As String, edate As String, price As String, tvol As String, moq As String
    tid = CStr(dbData(6, r)): edate = CStr(dbData(8, r))
    price = CStr(dbData(10, r)): tvol = CStr(dbData(9, r)): moq = CStr(dbData(11, r))

    ws.Cells(outRow, 1).Value = dbData(0, r)                 ' �A�C�e���R�[�h
    ws.Cells(outRow, 2).Value = dbData(1, r)                 ' ���[�J�[�i��
    ws.Cells(outRow, 3).Value = dbData(2, r)                 ' �i��
    ws.Cells(outRow, 4).Value = dbData(3, r)                 ' ���[�J�[��
    Dim vn As String: vn = CStr(dbData(5, r))
    If vn = "" Then vn = CStr(dbData(4, r))                  ' �d���於(������΃R�[�h)
    ws.Cells(outRow, 5).Value = vn
    ws.Cells(outRow, 6).Value = tid                          ' TID
    ws.Cells(outRow, 7).Value = fmtDate(CStr(dbData(7, r)))  ' �K�p�J�n
    ws.Cells(outRow, 8).Value = fmtDate(edate)              ' �K�p�I��
    ws.Cells(outRow, 9).Value = tvol                         ' �K�p����
    ws.Cells(outRow, 10).Value = price                       ' �����P��
    ws.Cells(outRow, 11).Value = moq                         ' MOQ

    ' �x��
    Dim w As String: w = ""
    If tid = "" Then
        w = "�P�����o�^"
    Else
        If isInf(edate) Then w = w & "������ "
        If Val(price) = 0 Then w = w & "0�~ "
    End If
    ws.Cells(outRow, 12).Value = Trim(w)

    ' �v���ӂ͐Ԏ�
    If tid = "" Or Val(price) = 0 Then ws.Cells(outRow, 10).Font.Color = vbRed
    If isInf(edate) Then ws.Cells(outRow, 8).Font.Color = vbRed
End Sub

' ���[�J�[�i�ڃR�[�h�̐��`�i�^���̊��ʓ��F�ؕ\�L�� < �ȍ~�������j
Private Function Clean�i��(ByVal s As String) As String
    Dim p As Long, kw As Variant, hasKw As Boolean
    p = InStr(s, "<"): If p > 0 Then s = Left(s, p - 1)
    p = InStr(s, "(")
    If p > 0 Then
        hasKw = False
        For Each kw In Array("UL", "ROHS", "CSA", "CE", "TUV", "VDE", "KEMA", "CEE", "PSE")
            If InStr(UCase(s), kw) > 0 Then hasKw = True: Exit For
        Next kw
        If hasKw Then s = Left(s, p - 1)
    End If
    Clean�i�� = Trim(s)
End Function

' ���t(YYYYMMDDS)��\���p�ɐ��`
Private Function fmtDate(ByVal v As String) As String
    Dim d As String, i As Long, ch As String
    d = ""
    For i = 1 To Len(v)
        ch = Mid(v, i, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next i
    If d = "" Then fmtDate = "": Exit Function
    If Left(d, 4) = "9999" Then fmtDate = "������(99/99/99)": Exit Function
    If Len(d) >= 8 Then
        fmtDate = Left(d, 4) & "/" & Mid(d, 5, 2) & "/" & Mid(d, 7, 2)
    Else
        fmtDate = v
    End If
End Function

Private Function isInf(ByVal v As String) As Boolean
    Dim d As String, i As Long, ch As String
    d = ""
    For i = 1 To Len(v)
        ch = Mid(v, i, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next i
    isInf = (Left(d, 4) = "9999")
End Function

' �\�������F�K�p�I����������(99/99/99) ���� �K�p����=0 �̂�
Private Function PassFilter(ByVal edate As String, ByVal tvol As String) As Boolean
    PassFilter = (isInf(edate) And Val(tvol) = 0)
End Function

' ���́E���ʂ̃N���A
Sub �N���A()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Worksheets("���o")
    Set wsOut = ThisWorkbook.Worksheets("�P������")
    wsIn.Range("A3:B10000").ClearContents
    Dim last As Long: last = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If last >= 3 Then wsOut.Rows("3:" & last).Clear
    MsgBox "�N���A���܂����B", vbInformation
End Sub
