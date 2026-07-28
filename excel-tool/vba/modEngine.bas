Attribute VB_Name = "modEngine"
Option Explicit

'==============================================================
' PDF解析エンジン
'
'  1) Power Query方式 (Excel 2021 / Microsoft 365) … 推奨
'     PDF内の文字の位置関係を保った表として取り出せるため、
'     「上下動」「左右動」の列を正確に判別できる。
'  2) Word変換方式 (Word 2013以降がインストールされている場合)
'     WordのPDF読み込み機能で表に変換して読み取る代替手段。
'
'  ParsePdf が 1) → 2) の順に自動で試す。
'==============================================================

' 解析中のPDFのヘッダー情報(測定日など)
Private mMeasDate As String
Private mLine As String
Private mCenter As String

' 正規表現(遅延バインディング)
Private mReTime As Object   ' 時刻   9:56:40 / 09:56:40
Private mReKm As Object     ' キロ程 68.504 (小数3桁)
Private mReVal As Object    ' 振幅   0.28   (小数2桁)
Private mReInt As Object    ' 整数
Private mReNum As Object    ' 数値

Private Sub InitRegex()
    If Not mReTime Is Nothing Then Exit Sub
    Set mReTime = CreateObject("VBScript.RegExp")
    mReTime.Pattern = "^\d{1,2}:\d{2}:\d{2}$"
    ' 数値がセル格納時に丸められても拾えるよう、小数桁は幅を持たせる
    ' (キロ程は時刻より右の最初の小数セル、振幅はキロ程より右のセル、と
    '  位置関係で区別するため誤判定しない)
    Set mReKm = CreateObject("VBScript.RegExp")
    mReKm.Pattern = "^\d+(\.\d{1,3})?$"
    Set mReVal = CreateObject("VBScript.RegExp")
    mReVal.Pattern = "^\d\.\d{1,2}$"
    Set mReInt = CreateObject("VBScript.RegExp")
    mReInt.Pattern = "^\d+$"
    Set mReNum = CreateObject("VBScript.RegExp")
    mReNum.Pattern = "^\d+(\.\d+)?$"
End Sub

'--- 入口: PDFを解析して12列の行データ(Collection)を返す ---
Public Function ParsePdf(pdfPath As String, ByRef rowsOut As Collection) As Boolean
    InitRegex
    mMeasDate = "": mLine = "": mCenter = ""
    Set rowsOut = New Collection

    Dim raw As New Collection
    If ParsePdfPowerQuery(pdfPath, raw) Then
        Finalize raw, rowsOut, pdfPath
        ParsePdf = (rowsOut.Count > 0)
        If ParsePdf Then Exit Function
    End If

    Set raw = New Collection
    mMeasDate = "": mLine = "": mCenter = ""
    If ParsePdfWord(pdfPath, raw) Then
        Finalize raw, rowsOut, pdfPath
        ParsePdf = (rowsOut.Count > 0)
    End If
End Function

'==============================================================
' 方式1: Power Query (Pdf.Tables)
'==============================================================
Private Function ParsePdfPowerQuery(pdfPath As String, raw As Collection) As Boolean
    Const QNAME As String = "_一時_PDF取込"
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo fail

    Dim m As String
    m = "let Src = Pdf.Tables(File.Contents(""" & Replace(pdfPath, """", """""") & """)), " & _
        "Pages = Table.SelectRows(Src, each [Kind] = ""Page""), " & _
        "Data = Table.Combine(Pages[Data]) in Data"

    On Error Resume Next
    ThisWorkbook.Queries(QNAME).Delete
    On Error GoTo fail
    ThisWorkbook.Queries.Add QNAME, m

    Set ws = ThisWorkbook.Worksheets.Add
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QNAME, _
        Destination:=ws.Range("A1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM [" & QNAME & "]"
        .BackgroundQuery = False
        .AdjustColumnWidth = False
        .Refresh False
    End With

    Dim grid As Variant
    grid = lo.Range.Value
    Cleanup ws, QNAME
    Set ws = Nothing

    ParsePdfPowerQuery = ParseGrid(grid, raw)
    Exit Function
fail:
    Cleanup ws, QNAME
    ParsePdfPowerQuery = False
End Function

Private Sub Cleanup(ws As Worksheet, qname As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Queries(qname).Delete
    On Error GoTo 0
End Sub

'==============================================================
' 方式2: Word変換
'==============================================================
Private Function ParsePdfWord(pdfPath As String, raw As Collection) As Boolean
    Dim wd As Object, doc As Object
    On Error GoTo fail
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    wd.DisplayAlerts = 0
    Set doc = wd.Documents.Open(FileName:=pdfPath, ConfirmConversions:=False, _
                                ReadOnly:=True, AddToRecentFiles:=False)
    If doc.Tables.Count = 0 Then GoTo fail

    ' ヘッダー情報は本文テキストから取得
    Dim txt As String
    txt = doc.Content.Text
    If mMeasDate = "" Then mMeasDate = ExtractAfter(txt, "測定日：")
    If mLine = "" Then mLine = ExtractAfter(txt, "線別：")
    If mCenter = "" Then mCenter = ExtractAfter(txt, "保線技術センター：")

    Dim t As Long, ok As Boolean
    For t = 1 To doc.Tables.Count
        Dim grid As Variant
        grid = TableToGrid(doc.Tables(t))
        If ParseGrid(grid, raw) Then ok = True
    Next

    doc.Close False
    wd.Quit
    ParsePdfWord = ok
    Exit Function
fail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wd Is Nothing Then wd.Quit
    ParsePdfWord = False
End Function

Private Function TableToGrid(tbl As Object) As Variant
    Dim nR As Long, nC As Long
    nR = tbl.rows.Count
    nC = tbl.Columns.Count
    Dim g() As Variant
    ReDim g(1 To nR, 1 To nC)
    Dim c As Object
    For Each c In tbl.Range.Cells
        Dim s As String
        s = c.Range.Text
        s = Replace(s, Chr(13), " ")
        s = Replace(s, Chr(7), "")
        If c.RowIndex <= nR And c.ColumnIndex <= nC Then
            g(c.RowIndex, c.ColumnIndex) = Trim(s)
        End If
    Next
    TableToGrid = g
End Function

Private Function ExtractAfter(txt As String, key As String) As String
    Dim p As Long, q As Long, s As String
    p = InStr(txt, key)
    If p = 0 Then Exit Function
    s = Mid(txt, p + Len(key), 40)
    For q = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, q, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = "　" Then Exit For
    Next
    ExtractAfter = Left(s, q - 1)
End Function

'==============================================================
' 共通: 表(2次元配列)の解析
'   ・「上下動」「左右動」の見出しセルの列番号を覚えておき、
'     数値がどちらの列に近いかで振り分ける
'   ・データ行は「整数のNO + 時刻 + キロ程(小数3桁)」で判定
'==============================================================
Private Function ParseGrid(grid As Variant, raw As Collection) As Boolean
    If IsEmpty(grid) Then Exit Function
    If Not IsArray(grid) Then Exit Function

    Dim nR As Long, nC As Long
    nR = UBound(grid, 1)
    nC = UBound(grid, 2)

    Dim udCol As Long, saCol As Long
    Dim r As Long, c As Long, added As Long

    For r = 1 To nR
        ' --- 見出し・ヘッダー情報の走査 ---
        For c = 1 To nC
            Dim s As String
            s = CellStr(grid, r, c)
            If s <> "" Then
                If InStr(s, "上下動") > 0 Then udCol = c
                If InStr(s, "左右動") > 0 Then saCol = c
                If InStr(s, "測定日") > 0 And mMeasDate = "" Then mMeasDate = AfterColon(grid, r, c)
                If InStr(s, "線別") > 0 And mLine = "" Then mLine = AfterColon(grid, r, c)
                If InStr(s, "保線技術センター") > 0 And mCenter = "" Then mCenter = AfterColon(grid, r, c)
            End If
        Next

        If udCol = 0 Or saCol = 0 Then GoTo nextRow

        ' --- データ行判定 ---
        Dim noCol As Long, timeCol As Long, kmCol As Long
        noCol = 0: timeCol = 0: kmCol = 0
        For c = 1 To nC
            s = CellStr(grid, r, c)
            If s <> "" Then
                If noCol = 0 And mReInt.Test(s) Then noCol = c
                If timeCol = 0 And mReTime.Test(s) Then timeCol = c
                If kmCol = 0 And mReKm.Test(s) And c > timeCol And timeCol > 0 Then kmCol = c
            End If
        Next
        If noCol = 0 Or timeCol = 0 Or kmCol = 0 Then GoTo nextRow
        If Not (noCol < timeCol And timeCol < kmCol) Then GoTo nextRow

        ' --- 各項目の取り出し ---
        Dim rowNo As Long, spd As Variant, tim As String, km As Double
        Dim ud As Variant, sa As Variant, star As String, biko As String
        rowNo = CLng(CellStr(grid, r, noCol))
        tim = CellStr(grid, r, timeCol)
        km = Val(CellStr(grid, r, kmCol))
        spd = Empty: ud = Empty: sa = Empty: star = "": biko = ""

        For c = 1 To nC
            If c <> noCol And c <> timeCol And c <> kmCol Then
                s = CellStr(grid, r, c)
                If s <> "" Then
                    Dim sv As String
                    sv = Trim(Replace(s, "*", ""))
                    If c > noCol And c < timeCol And mReNum.Test(s) Then
                        spd = Val(s)                       ' 列車速度
                    ElseIf c > kmCol And mReVal.Test(sv) Then
                        If Abs(c - udCol) <= Abs(c - saCol) Then
                            ud = Val(sv)                   ' 上下動
                        Else
                            sa = Val(sv)                   ' 左右動
                        End If
                        If InStr(s, "*") > 0 Then star = "*"
                    ElseIf s = "*" Then
                        star = "*"                         ' 超過マークが単独セルの場合
                    ElseIf c > saCol Then
                        biko = Trim(biko & " " & s)        ' 備考
                    ElseIf c > timeCol And c < kmCol Then
                        biko = Trim("(" & s & ") " & biko) ' キロ程の左の記号(例: W)
                    End If
                End If
            End If
        Next

        Dim a(1 To 9) As Variant
        a(1) = rowNo: a(2) = spd: a(3) = tim: a(4) = km
        a(5) = ud: a(6) = sa: a(7) = star: a(8) = biko
        raw.Add a
        added = added + 1
nextRow:
    Next

    ParseGrid = (added > 0)
End Function

Private Function CellStr(grid As Variant, r As Long, c As Long) As String
    Dim v As Variant
    v = grid(r, c)
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CellStr = Trim(CStr(v))
End Function

'--- 「測定日：2026/06/04」形式、または右隣のセルから値を取り出す ---
Private Function AfterColon(grid As Variant, r As Long, c As Long) As String
    Dim s As String, p As Long
    s = CellStr(grid, r, c)
    p = InStr(s, "：")
    If p = 0 Then p = InStr(s, ":")
    If p > 0 And p < Len(s) Then
        AfterColon = Trim(Mid(s, p + 1))
        If AfterColon <> "" Then Exit Function
    End If
    Dim cc As Long
    For cc = c + 1 To UBound(grid, 2)
        s = CellStr(grid, r, cc)
        If s <> "" Then
            AfterColon = s
            Exit Function
        End If
    Next
End Function

'--- 9列の生データに測定日・ファイル名等を付けて12列にする ---
Private Sub Finalize(raw As Collection, rowsOut As Collection, pdfPath As String)
    Dim fname As String
    fname = Mid(pdfPath, InStrRev(pdfPath, "\") + 1)
    If InStr(fname, "/") > 0 Then fname = Mid(fname, InStrRev(fname, "/") + 1)

    Dim item As Variant
    For Each item In raw
        Dim a(1 To 12) As Variant
        Dim i As Long
        For i = 1 To 8
            a(i) = item(i)
        Next
        a(9) = mMeasDate
        a(10) = mLine
        a(11) = mCenter
        a(12) = fname
        rowsOut.Add a
    Next
End Sub
