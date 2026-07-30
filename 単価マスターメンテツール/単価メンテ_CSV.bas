Attribute VB_Name = "�P�������eCSV"
Option Explicit

' ===== �ꊇ�o�^CSV�����iTPiCS �e�L�X�g�Ǎ��p�j =====
' ���T���v���ɏ������āu�I���p�v�Ɓu�V�K�p�v��2�t�@�C�����o�͂���B
'   �I���p: DKUBU,TID,EDATE ��3��̂݁iEDATE=������̑O���j�� �����ڂ��㏑�����Ȃ�
'   �V�K�p: 28��iDKUBU=New, CODE, VENDOR, MAKERNAME, TDATE, EDATE=99/99/99, TVOL, PRICE ���j
' ���t�� yyyy/m/d �`���A�������� 99/99/99�B�����R�[�h�� UTF-8(BOM)�BDB�� SELECT �̂݁B

Private Const SERVER_IP As String = "YOUR_DB_SERVER_IP"
Private Const DB_NAME As String = "TPICSDB_T"
Private Const DB_USER As String = "YOUR_DB_USER"
Private Const DB_PASS As String = "YOUR_DB_PASSWORD"
Private Const NEWHDR As String = "DKUBU,TID,CODE,VCODE,VENDOR,WCODE,MAKERNAME,TDATE,EDATE,TVOL,TEMA,SPRICE,PRICE,TAXRATE,TAXKUBU,SOUKINGAKU,SOUSUUIN,APRICE,HOKAN,NOTE,SYSNOTE,INPUTDATE,INPUTUSER,KARIKBN,MITSUMORINO,PCAPA,TANI3,IKOUFLAG"

' �u�P�����ʁv���u�ꊇ�o�^�v�֓]�L�iTID�̂���s�̂݁j
Sub ���ʂ��ꊇ�o�^��()
    Dim wsR As Worksheet, wsK As Worksheet
    Set wsR = ThisWorkbook.Worksheets("�P������")
    Set wsK = ThisWorkbook.Worksheets("�ꊇ�o�^")

    Dim cMOQ As Long: cMOQ = Col(wsK, "MOQ")
    If cMOQ = 0 Then
        cMOQ = wsK.Cells(2, wsK.Columns.Count).End(xlToLeft).Column + 1
        wsK.Cells(2, cMOQ - 1).Copy
        wsK.Cells(2, cMOQ).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        wsK.Cells(2, cMOQ).Value = "MOQ"
        wsK.Columns(cMOQ).ColumnWidth = 8
    End If
    Dim cMIT As Long: cMIT = Col(wsK, "����")
    If cMIT = 0 Then
        cMIT = wsK.Cells(2, wsK.Columns.Count).End(xlToLeft).Column + 1
        wsK.Cells(2, cMIT).Value = "����No"
        wsK.Cells(2, cMIT).Font.Bold = True
        wsK.Cells(2, cMIT).Interior.Color = RGB(255, 242, 204)
        wsK.Cells(2, cMIT).HorizontalAlignment = xlCenter
        wsK.Cells(2, cMIT).Borders.LineStyle = xlContinuous
        wsK.Columns(cMIT).ColumnWidth = 14
    End If

    Dim cTID As Long, cCODE As Long, cVEN As Long, cST As Long, cET As Long
    Dim cQTY As Long, cPR As Long, cKBN As Long
    cTID = Col(wsK, "TID"): cCODE = Col(wsK, "�A�C�e���R�[�h"): cVEN = Col(wsK, "������")
    cST = Col(wsK, "���s�K�p�J�n"): cET = Col(wsK, "���s�K�p�I��")
    cQTY = Col(wsK, "�K�p����", "�V"): cPR = Col(wsK, "���s�����P��"): cKBN = Col(wsK, "�敪")

    Dim lastR As Long, lastK As Long, i As Long, o As Long
    lastR = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    lastK = wsK.Cells(wsK.Rows.Count, IIf(cTID > 0, cTID, 1)).End(xlUp).Row
    If lastK >= 3 Then wsK.Rows("3:" & lastK).ClearContents

    o = 3
    For i = 3 To lastR
        Dim tid As String: tid = Trim(CStr(wsR.Cells(i, 6).Value))
        If tid <> "" Then
            If cTID > 0 Then wsK.Cells(o, cTID).Value = tid
            If cCODE > 0 Then wsK.Cells(o, cCODE).Value = wsR.Cells(i, 1).Value
            If cVEN > 0 Then wsK.Cells(o, cVEN).Value = wsR.Cells(i, 5).Value
            If cST > 0 Then wsK.Cells(o, cST).Value = wsR.Cells(i, 7).Value
            If cET > 0 Then wsK.Cells(o, cET).Value = wsR.Cells(i, 8).Value
            If cQTY > 0 Then wsK.Cells(o, cQTY).Value = wsR.Cells(i, 9).Value
            If cPR > 0 Then wsK.Cells(o, cPR).Value = wsR.Cells(i, 10).Value
            If cMOQ > 0 Then wsK.Cells(o, cMOQ).Value = wsR.Cells(i, 11).Value
            If cKBN > 0 Then wsK.Cells(o, cKBN).Value = "���i����"
            o = o + 1
        End If
    Next i
    If o > 3 Then
        If cCODE > 0 Then wsK.Range(wsK.Cells(3, cCODE), wsK.Cells(o - 1, cCODE)).HorizontalAlignment = xlLeft
        If cST > 0 Then wsK.Range(wsK.Cells(3, cST), wsK.Cells(o - 1, cST)).HorizontalAlignment = xlLeft
    End If
    wsK.Select
    MsgBox (o - 3) & " �����u�ꊇ�o�^�v�֓]�L���܂����B����P���E��������L�����āwCSV�����x�����s���Ă��������B", vbInformation
End Sub

Sub CSV����()
    Dim wsK As Worksheet: Set wsK = ThisWorkbook.Worksheets("�ꊇ�o�^")
    Dim cTID As Long, cPRICE As Long, cKD As Long, cKBN As Long, cNQ As Long, cCODE As Long, cVEN As Long
    cTID = Col(wsK, "TID"): cPRICE = Col(wsK, "����P��"): cKD = Col(wsK, "�����")
    cKBN = Col(wsK, "�敪"): cNQ = Col(wsK, "�V�K�p����"): cCODE = Col(wsK, "�A�C�e���R�[�h"): cVEN = Col(wsK, "������")
    Dim cMIT2 As Long: cMIT2 = Col(wsK, "����")
    If cPRICE = 0 Or cKD = 0 Then MsgBox "�u�ꊇ�o�^�v�Ɂw����P���x�w������x�̌��o����������܂���B", vbCritical: Exit Sub

    Dim last As Long
    last = Application.WorksheetFunction.Max( _
             wsK.Cells(wsK.Rows.Count, IIf(cTID > 0, cTID, cPRICE)).End(xlUp).Row, _
             wsK.Cells(wsK.Rows.Count, cPRICE).End(xlUp).Row)
    If last < 3 Then MsgBox "�u�ꊇ�o�^�v�Ƀf�[�^������܂���B", vbExclamation: Exit Sub

    Dim hdr As Variant: hdr = Split(NEWHDR, ",")

    ' --- TID���W ---
    Dim ids As String, i As Long
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 3 To last
        Dim t As String: t = IIf(cTID > 0, Trim(CStr(wsK.Cells(i, cTID).Value)), "")
        If t <> "" And Not seen.Exists(t) Then seen.Add t, True: ids = ids & "'" & Replace(t, "'", "''") & "',"
    Next i

    ' --- XTANK���猻�s�s���擾�i���݂���񂾂��R�s�[�p�Ɏ��j ---
    Dim cur As Object: Set cur = CreateObject("Scripting.Dictionary")  ' TID -> Dictionary(colUpper->value)
    If Len(ids) > 0 Then
        ids = Left(ids, Len(ids) - 1)
        On Error GoTo DBERR
        Dim conn As Object, rs As Object, sql As String
        Set conn = CreateObject("ADODB.Connection")
        conn.Open "Provider=SQLOLEDB;Data Source=" & SERVER_IP & ";Initial Catalog=" & DB_NAME & _
                  ";User ID=" & DB_USER & ";Password=" & DB_PASS & ";"
        conn.Execute "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED"
        ' XTANK�ɑ��݂�����c��
        Dim colset As Object: Set colset = CreateObject("Scripting.Dictionary")
        Set rs = conn.Execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='XTANK'")
        Do While Not rs.EOF
            colset(UCase(Trim(CStr(rs.Fields(0).Value)))) = True
            rs.MoveNext
        Loop
        rs.Close
        ' TID + NEWHDR�̂���XTANK�ɂ���� ���擾
        Dim selp As String, hi As Long, hn As String
        selp = "CAST(TID AS varchar(30)) AS TID"
        For hi = 0 To UBound(hdr)
            hn = Trim(hdr(hi))
            If UCase(hn) <> "TID" And colset.Exists(UCase(hn)) Then
                selp = selp & ", ISNULL(CAST([" & hn & "] AS varchar(255)),'') AS [" & hn & "]"
            End If
        Next hi
        sql = "SELECT " & selp & " FROM XTANK WITH (NOLOCK) WHERE TID IN (" & ids & ")"
        Set rs = conn.Execute(sql)
        Do While Not rs.EOF
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            Dim fj As Long
            For fj = 0 To rs.Fields.Count - 1
                d(UCase(rs.Fields(fj).Name)) = CStr(rs.Fields(fj).Value & "")
            Next fj
            Set cur(CStr(rs.Fields(0).Value & "")) = d
            rs.MoveNext
        Loop
        rs.Close: conn.Close
        On Error GoTo 0
    End If

    ' --- �o�͍s�쐬 ---
    Dim endLines As New Collection, newLines As New Collection
    endLines.Add "DKUBU,TID,EDATE"
    newLines.Add NEWHDR
    Dim madeEnd As Long, madeNew As Long

    For i = 3 To last
        Dim price As String: price = Trim(CStr(wsK.Cells(i, cPRICE).Value))
        Dim dcell As Variant: dcell = wsK.Cells(i, cKD).Value
        If price <> "" And Trim(CStr(dcell)) <> "" Then
            Dim kubun As String: kubun = IIf(cKBN > 0, Trim(CStr(wsK.Cells(i, cKBN).Value)), "")
            Dim isNewOnly As Boolean: isNewOnly = (InStr(kubun, "�V�K") > 0)
            Dim ntvol As String: ntvol = IIf(cNQ > 0, Trim(CStr(wsK.Cells(i, cNQ).Value)), "")
            Dim tid2 As String: tid2 = IIf(cTID > 0, Trim(CStr(wsK.Cells(i, cTID).Value)), "")
            Dim maker As String: maker = IIf(cVEN > 0, Trim(CStr(wsK.Cells(i, cVEN).Value)), "")
            Dim shCODE As String: shCODE = IIf(cCODE > 0, Trim(CStr(wsK.Cells(i, cCODE).Value)), "")

            Dim curd As Object, hasCur As Boolean: hasCur = False
            If tid2 <> "" Then If cur.Exists(tid2) Then Set curd = cur(tid2): hasCur = True

            ' �I���s
            If (Not isNewOnly) And tid2 <> "" Then
                endLines.Add "Mod," & tid2 & "," & prevYmd(dcell)
                madeEnd = madeEnd + 1
            End If

            ' �V�K�s�i���s���R�s�[���A�K�v���ڂ̂ݏ㏑���j
            Dim nf() As String: ReDim nf(UBound(hdr))
            Dim j As Long, nm As String, v As String
            Dim useTVOL As String
            For j = 0 To UBound(hdr)
                nm = UCase(Trim(hdr(j)))
                Select Case nm
                    Case "DKUBU": v = "New"
                    Case "TID": v = ""
                    Case "TDATE": v = ymd(dcell)
                    Case "EDATE": v = "99/99/99"
                    Case "PRICE": v = CleanN(price)
                    Case "SPRICE": v = ""
                    Case "TVOL"
                        useTVOL = ntvol
                        If useTVOL = "" And hasCur And curd.Exists("TVOL") Then useTVOL = curd("TVOL")
                        v = CleanN(useTVOL)
                    Case "SOUKINGAKU", "SOUSUUIN", "APRICE", "NOTE", "SYSNOTE", "INPUTDATE", "INPUTUSER", "KARIKBN", "IKOUFLAG"
                        v = ""
                    Case "CODE"
                        If hasCur And curd.Exists("CODE") Then v = curd("CODE") Else v = shCODE
                    Case "MITSUMORINO"
                        v = IIf(cMIT2 > 0, Trim(CStr(wsK.Cells(i, cMIT2).Value)), "")
                    Case "MAKERNAME"
                        If hasCur And curd.Exists("MAKERNAME") Then v = curd("MAKERNAME") Else v = maker
                    Case Else
                        If hasCur And curd.Exists(nm) Then v = curd(nm) Else v = ""
                End Select
                nf(j) = qCSV(v)
            Next j
            newLines.Add Join(nf, ",")
            madeNew = madeNew + 1
        End If
    Next i
    If madeNew = 0 Then MsgBox "����P���Ɖ���������͂��ꂽ�s������܂���B", vbExclamation: Exit Sub

    ' --- �ۑ��iUTF-8 BOM�A2�t�@�C���j ---
    Dim base As Variant
    base = Application.GetSaveAsFilename(InitialFileName:="tanka_" & Format(Date, "yyyymmdd"), _
             FileFilter:="CSV�t�@�C�� (*.csv), *.csv")
    If base = False Then Exit Sub
    Dim bp As String: bp = CStr(base)
    If LCase(Right(bp, 4)) = ".csv" Then bp = Left(bp, Len(bp) - 4)

    If madeEnd > 0 Then WriteUTF8 bp & "_�I���p.csv", JoinCol(endLines)
    WriteUTF8 bp & "_�V�K�p.csv", JoinCol(newLines)

    MsgBox "CSV���o�͂��܂����B" & vbCrLf & _
           "�I���p: " & madeEnd & " �s �^ �V�K�p: " & madeNew & " �s" & vbCrLf & vbCrLf & _
           "TPiCS�̃e�L�X�g�Ǎ��݂ŁA�I���p���V�K�p �̏��Ŏ�荞��ł��������B", vbInformation
    Exit Sub
DBERR:
    MsgBox "DB�ڑ��G���[: " & Err.Description, vbCritical
End Sub

' ===== �⏕ =====
Private Function newRow(code As String, vendor As String, maker As String, _
                        tdate As String, tvol As String, price As String) As String
    Dim nf(27) As String, z As Long
    For z = 0 To 27: nf(z) = "": Next z
    nf(0) = "New"
    nf(2) = code
    nf(4) = vendor
    nf(6) = maker
    nf(7) = tdate
    nf(8) = "99/99/99"
    nf(9) = tvol
    nf(12) = price
    newRow = Join(nf, ",")
End Function

Private Function qCSV(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Then
        qCSV = """" & Replace(s, """", """""") & """"
    Else
        qCSV = s
    End If
End Function

Private Function JoinCol(c As Collection) As String
    Dim s As String, v As Variant
    For Each v In c: s = s & v & vbCrLf: Next v
    JoinCol = s
End Function

Private Sub WriteUTF8(ByVal path As String, ByVal txt As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "UTF-8": st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

' ���o���𕔕���v�ŒT���ikey�܂�/excl�܂܂Ȃ� �ŏ��̗�j�B�������0�B
Private Function Col(ws As Worksheet, ByVal key As String, Optional ByVal excl As String = "") As Long
    Dim lastC As Long: lastC = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim cc As Long, h As String
    For cc = 1 To lastC
        h = CStr(ws.Cells(2, cc).Value)
        h = Replace(h, " ", ""): h = Replace(h, "�@", ""): h = Replace(h, vbLf, ""): h = Replace(h, vbCr, "")
        If InStr(h, key) > 0 Then
            If excl = "" Or InStr(h, excl) = 0 Then Col = cc: Exit Function
        End If
    Next cc
    Col = 0
End Function

Private Function digits(ByVal v As String) As String
    Dim s As String, i As Long, ch As String
    For i = 1 To Len(v)
        ch = Mid(v, i, 1)
        If ch >= "0" And ch <= "9" Then s = s & ch
    Next i
    digits = s
End Function

Private Function toDt(ByVal v As Variant) As Date
    If IsDate(v) Then toDt = CDate(v): Exit Function
    Dim s As String: s = digits(CStr(v))
    toDt = DateSerial(CLng(Left(s, 4)), CLng(Mid(s, 5, 2)), CLng(Mid(s, 7, 2)))
End Function

Private Function ymd(ByVal v As Variant) As String
    ymd = Format(toDt(v), "yyyy/m/d")
End Function

Private Function prevYmd(ByVal v As Variant) As String
    prevYmd = Format(toDt(v) - 1, "yyyy/m/d")
End Function

Private Function CleanN(ByVal v As String) As String
    Dim s As String: s = Trim(v)
    If s = "" Then CleanN = "": Exit Function
    If Not IsNumeric(s) Then CleanN = s: Exit Function
    Dim d As Double: d = CDbl(s)
    If d = Int(d) Then CleanN = CStr(CLng(d)) Else CleanN = Format(d, "0.############")
End Function
