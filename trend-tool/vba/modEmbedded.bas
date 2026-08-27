Attribute VB_Name = "modEmbedded"
Option Explicit

'==============================================================
' シートに貼り付けられた(埋め込まれた)PDFの取り出し
'
' 仕組み:
'   Excelブック(.xlsm)は実体がZIP形式で、貼り付けたPDFは
'   ブック内の xl\embeddings フォルダに保存されている。
'   ブックのコピーをZIPとして開き、その中からPDFデータ
'   (%PDF ～ %%EOF)を切り出して一時ファイルに保存する。
'==============================================================

'--- ブック内に埋め込まれた全PDFを一時フォルダに取り出し、パスの一覧を返す ---
Public Function ExtractEmbeddedPdfs() As Collection
    Dim result As New Collection
    Set ExtractEmbeddedPdfs = result

    If ThisWorkbook.Path = "" Then
        MsgBox "先にこのブックを保存してください。", vbExclamation
        Set ExtractEmbeddedPdfs = Nothing
        Exit Function
    End If
    If LCase(Left(ThisWorkbook.FullName, 4)) = "http" Then
        MsgBox "OneDrive/SharePoint上のブックでは「貼り付けたPDFの取込」は使えません。" & vbCrLf & _
               "ローカル(デスクトップ等)に保存し直すか、" & vbCrLf & _
               "「PDFファイルを選択して取込」をご利用ください。", vbExclamation
        Set ExtractEmbeddedPdfs = Nothing
        Exit Function
    End If

    ' 貼り付けたPDFを確実にブック内に保存する
    ThisWorkbook.Save

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmpDir As String
    tmpDir = Environ$("TEMP") & "\動揺PDF取込_" & Format(Now, "yyyymmdd_hhmmss")
    fso.CreateFolder tmpDir

    ' ブックのコピーをZIPとして展開する
    Dim zipPath As String
    zipPath = tmpDir & "\book.zip"
    fso.CopyFile ThisWorkbook.FullName, zipPath

    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    Dim src As Object
    Set src = sh.Namespace(zipPath & "\xl\embeddings")
    If src Is Nothing Then Exit Function        ' 埋め込みオブジェクトなし

    Dim dst As String
    dst = tmpDir & "\emb"
    fso.CreateFolder dst
    sh.Namespace(dst).CopyHere src.Items, 4 + 16

    ' 展開完了を待つ(最大30秒)
    Dim t0 As Single
    t0 = Timer
    Do While fso.GetFolder(dst).Files.Count < src.Items.Count
        DoEvents
        If Timer - t0 > 30 Then Exit Do
    Loop

    ' 各ファイルからPDFデータを切り出す
    Dim f As Object, idx As Long
    For Each f In fso.GetFolder(dst).Files
        Dim pdfPath As String
        pdfPath = CarvePdf(CStr(f.Path), tmpDir, idx)
        If pdfPath <> "" Then
            result.Add pdfPath
            idx = idx + 1
        End If
    Next
End Function

'--- ファイル内の %PDF ～ 最後の %%EOF を切り出して .pdf として保存 ---
Private Function CarvePdf(binPath As String, outDir As String, idx As Long) As String
    On Error GoTo fail
    Dim fnum As Integer
    fnum = FreeFile
    Open binPath For Binary Access Read As #fnum
    If LOF(fnum) < 100 Then
        Close #fnum
        Exit Function
    End If
    Dim bytes() As Byte
    ReDim bytes(0 To LOF(fnum) - 1)
    Get #fnum, , bytes
    Close #fnum

    Dim sigPdf As String, sigEof As String
    sigPdf = StrConv("%PDF", vbFromUnicode)
    sigEof = StrConv("%%EOF", vbFromUnicode)

    Dim p1 As Long
    p1 = InStrB(1, bytes, sigPdf)
    If p1 = 0 Then Exit Function

    ' 最後の %%EOF を探す
    Dim p As Long, p2 As Long
    p = InStrB(p1, bytes, sigEof)
    Do While p > 0
        p2 = p
        p = InStrB(p + 1, bytes, sigEof)
    Loop
    If p2 = 0 Then Exit Function

    Dim endB As Long
    endB = p2 - 1 + LenB(sigEof) + 2            ' 改行分の余裕
    If endB > UBound(bytes) + 1 Then endB = UBound(bytes) + 1

    Dim outLen As Long
    outLen = endB - (p1 - 1)
    Dim outB() As Byte
    ReDim outB(0 To outLen - 1)
    Dim k As Long
    For k = 0 To outLen - 1
        outB(k) = bytes(p1 - 1 + k)
    Next

    Dim outPath As String
    outPath = outDir & "\貼付PDF_" & (idx + 1) & ".pdf"
    fnum = FreeFile
    Open outPath For Binary Access Write As #fnum
    Put #fnum, , outB
    Close #fnum
    CarvePdf = outPath
    Exit Function
fail:
    On Error Resume Next
    Close #fnum
End Function
