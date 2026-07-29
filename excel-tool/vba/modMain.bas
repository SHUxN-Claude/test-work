Attribute VB_Name = "modMain"
Option Explicit

'==============================================================
' 動揺データPDF取込ツール (メイン処理)
'
' シート構成:
'   設定         … 目標値・基準値・黄色強調値
'   PDF貼付      … 取込ボタン / PDF貼り付け場所
'   全データ     … PDFの全行
'   ①基準値超過 … 基準値以上の行
'   ②目標値超過 … 目標値以上～基準値未満の行
'==============================================================

Public Const SHEET_SETTINGS As String = "設定"
Public Const SHEET_PASTE As String = "PDF貼付"
Public Const SHEET_ALL As String = "全データ"
Public Const SHEET_OVER As String = "①基準値超過"
Public Const SHEET_TARGET As String = "②目標値超過"
Public Const DATA_START As Long = 5
Public Const COL_LAST As Long = 12

'--- 初回に1回だけ実行: PDF貼付シートにボタンを設置する ---
Public Sub 初期設定()
    ' .xlsx のまま保存するとマクロが失われるため、先に形式を確認する
    If LCase(Right(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        If MsgBox("このブックがマクロ有効ブック(.xlsm)として保存されていません。" & vbCrLf & _
                  "このまま保存するとマクロが消えてしまいます。" & vbCrLf & vbCrLf & _
                  "F12(名前を付けて保存)で、ファイルの種類を" & vbCrLf & _
                  "「Excel マクロ有効ブック (*.xlsm)」にして保存してください。" & vbCrLf & vbCrLf & _
                  "このまま続けますか?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PASTE)
    Dim b As Object
    For Each b In ws.Buttons
        b.Delete
    Next
    ' 位置はPDF貼付シートの「▼ ここにボタンが表示されます ▼」の直下(20～21行目)
    AddBtn ws, 20, 285, 230, 32, "PDFファイルを選択して取込", "PDFファイル選択取込"
    AddBtn ws, 265, 285, 230, 32, "貼り付けたPDFを取り込む", "貼付PDF取込"
    AddBtn ws, 510, 285, 190, 32, "設定変更後の再抽出", "再抽出"
    ws.Activate
    MsgBox "ボタンを設置しました。" & vbCrLf & vbCrLf & _
           "「PDF貼付」シートの中ほどにボタンが3つ表示されています。", vbInformation
End Sub

Private Sub AddBtn(ws As Worksheet, l As Single, t As Single, _
                   w As Single, h As Single, cap As String, macroName As String)
    Dim b As Button
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = macroName
End Sub

'--- ボタン: ファイル選択ダイアログでPDFを選んで取込 ---
Public Sub PDFファイル選択取込()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "取り込むPDFを選択してください(複数選択可)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        If .Show = 0 Then Exit Sub
    End With
    Dim paths As New Collection, i As Long
    For i = 1 To fd.SelectedItems.Count
        paths.Add fd.SelectedItems(i)
    Next
    ImportPdfFiles paths
End Sub

'--- ボタン: PDF貼付シートに貼られたPDFオブジェクトを取込 ---
Public Sub 貼付PDF取込()
    Dim paths As Collection
    Set paths = ExtractEmbeddedPdfs()
    If paths Is Nothing Then Exit Sub          ' エラーメッセージは表示済み
    If paths.Count = 0 Then
        MsgBox "貼り付けられたPDFが見つかりませんでした。" & vbCrLf & _
               "PDFを貼り付けてブックを保存してから、もう一度実行してください。", vbExclamation
        Exit Sub
    End If
    ImportPdfFiles paths
End Sub

'--- 取込の共通処理 ---
Public Sub ImportPdfFiles(paths As Collection)
    If paths Is Nothing Then Exit Sub
    If paths.Count = 0 Then Exit Sub

    Dim ans As VbMsgBoxResult
    ans = MsgBox("取り込み済みのデータをクリアしてから取り込みますか?" & vbCrLf & _
                 "  [はい]   クリアして取込" & vbCrLf & _
                 "  [いいえ] いまのデータに追記", vbYesNoCancel + vbQuestion)
    If ans = vbCancel Then Exit Sub
    If ans = vbYes Then
        ClearData SHEET_ALL
        ClearData SHEET_OVER
        ClearData SHEET_TARGET
    End If

    Dim p As Variant, okCnt As Long, addCnt As Long, ngList As String
    Application.ScreenUpdating = False
    On Error GoTo done
    For Each p In paths
        Dim rows As Collection
        Set rows = Nothing
        If ParsePdf(CStr(p), rows) Then
            WriteRows ThisWorkbook.Worksheets(SHEET_ALL), rows
            okCnt = okCnt + 1
            addCnt = addCnt + rows.Count
        Else
            ngList = ngList & vbCrLf & "  " & p
        End If
    Next
    再抽出Silent
done:
    Application.ScreenUpdating = True
    Dim msg As String
    msg = okCnt & " 件のPDFから " & addCnt & " 行を取り込みました。"
    If ngList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "取り込めなかったPDF:" & ngList & vbCrLf & _
              "(PDFの様式が異なるか、この環境では読み取りできない可能性があります)"
    End If
    MsgBox msg, IIf(ngList = "", vbInformation, vbExclamation)
End Sub

'--- ボタン: 設定変更後にPDFを読み直さず①②を作り直す ---
Public Sub 再抽出()
    再抽出Silent
    MsgBox "「" & SHEET_OVER & "」「" & SHEET_TARGET & "」を再抽出しました。", vbInformation
End Sub

Public Sub 再抽出Silent()
    Dim tgt As Double, lim As Double
    With ThisWorkbook.Worksheets(SHEET_SETTINGS)
        tgt = Val(.Range("C3").Value)   ' 目標値
        lim = Val(.Range("C4").Value)   ' 基準値
    End With

    ClearData SHEET_OVER
    ClearData SHEET_TARGET

    Dim wsAll As Worksheet
    Set wsAll = ThisWorkbook.Worksheets(SHEET_ALL)
    Dim lastR As Long
    lastR = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row
    If lastR < DATA_START Then Exit Sub

    Dim arr As Variant
    arr = wsAll.Range(wsAll.Cells(DATA_START, 1), wsAll.Cells(lastR, COL_LAST)).Value

    Dim overRows As New Collection, tgtRows As New Collection
    Dim r As Long, v As Double, hasVal As Boolean
    For r = 1 To UBound(arr, 1)
        hasVal = False
        If IsNumber(arr(r, 5)) Then
            v = arr(r, 5): hasVal = True
        ElseIf IsNumber(arr(r, 6)) Then
            v = arr(r, 6): hasVal = True
        End If
        If hasVal Then
            If v >= lim Then
                overRows.Add RowSlice(arr, r)
            ElseIf v >= tgt Then
                tgtRows.Add RowSlice(arr, r)
            End If
        End If
    Next
    WriteRows ThisWorkbook.Worksheets(SHEET_OVER), overRows
    WriteRows ThisWorkbook.Worksheets(SHEET_TARGET), tgtRows
End Sub

Private Function IsNumber(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim(v) = "" Then Exit Function
    End If
    IsNumber = IsNumeric(v)
End Function

Private Function RowSlice(arr As Variant, r As Long) As Variant
    Dim a(1 To COL_LAST) As Variant, i As Long
    For i = 1 To COL_LAST
        a(i) = arr(r, i)
    Next
    RowSlice = a
End Function

'--- データ行の書き込み(罫線・フォント込み) ---
Public Sub WriteRows(ws As Worksheet, rows As Collection)
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub
    Dim startR As Long
    startR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If startR < DATA_START Then startR = DATA_START

    Dim out() As Variant
    ReDim out(1 To rows.Count, 1 To COL_LAST)
    Dim item As Variant, k As Long, i As Long
    For Each item In rows
        k = k + 1
        For i = 1 To COL_LAST
            out(k, i) = item(i)
        Next
    Next
    With ws.Range(ws.Cells(startR, 1), ws.Cells(startR + rows.Count - 1, COL_LAST))
        .Value = out
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(153, 153, 153)
        .Font.Name = "游ゴシック"
        .Font.Size = 10
    End With
End Sub

'--- データ行のクリア(見出し・条件付き書式は残す) ---
Public Sub ClearData(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= DATA_START Then
        With ws.Range(ws.Cells(DATA_START, 1), ws.Cells(lastR, COL_LAST))
            .ClearContents
            .Borders.LineStyle = xlNone
        End With
    End If
End Sub
