Attribute VB_Name = "modTrend"
Option Explicit

'==============================================================
' 動揺データ 年度推移管理ツール
'
' ボタン:
'   PDFファイルを選択して取込 / 貼り付けたPDFを取り込む
'   再計算 / 新年度開始
'
' 中核は 再計算()。取り込んだ全データから毎回グループ分けをやり直すため、
' 「同一箇所とみなす距離」を変えると過去の月も含めて組み替わる。
'==============================================================

Public Const SH_SET As String = "設定"
Public Const SH_PASTE As String = "PDF貼付"
Public Const SH_TREND As String = "推移表"
Public Const SH_CHART As String = "グラフ"
Public Const SH_ALL As String = "全データ"
Public Const SH_OVER As String = "①基準値超過"
Public Const SH_TGT As String = "②目標値超過"
Public Const SH_MEMO As String = "箇所メモ"
Public Const SH_REPAIR As String = "施工記録"
Public Const SH_HIST As String = "取込履歴"
Public Const SH_PREV As String = "前年度実績"

Public Const D_START As Long = 5      ' データ開始行
Public Const ALL_COLS As Long = 13    ' 全データの列数(13列目=月)
Public Const TREND_COLS As Long = 24  ' 推移表の列数
Public Const C_MEMO As Long = 8       ' 推移表 H列 = メモ(手入力)
Public Const C_MON0 As Long = 9       ' 推移表 I列 = 4月
Public Const C_YMAX As Long = 21      ' U列 = 年間最大
Public Const C_PREV As Long = 22      ' V列 = 前年度値
Public Const MAX_ROWS As Long = 20000

'--- グループ(箇所) ---
Private Type TCluster
    line As String
    kind As String
    rep As Double        ' 代表キロ程(最大値が出たときの位置)
    lo As Double
    hi As Double
    biko As String
    maxV As Double
    mon(1 To 12) As Double
End Type

'--- 手入力メモ / 施工記録 ---
Private Type TMemo
    line As String
    km As Double
    text As String
End Type

Private Type TRepair
    line As String
    km As Double
    d As Variant
    text As String
End Type


'==============================================================
' 初期設定(ボタン設置)
'==============================================================
Public Sub 初期設定()
    If LCase(Right(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        If MsgBox("このブックがマクロ有効ブック(.xlsm)として保存されていません。" & vbCrLf & _
                  "このまま保存するとマクロが消えてしまいます。" & vbCrLf & vbCrLf & _
                  "F12(名前を付けて保存)で「Excel マクロ有効ブック (*.xlsm)」" & vbCrLf & _
                  "にして保存してください。" & vbCrLf & vbCrLf & _
                  "このまま続けますか?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_PASTE)
    Dim b As Object
    For Each b In ws.Buttons
        b.Delete
    Next
    AddBtn ws, 20, 285, 210, 32, "PDFファイルを選択して取込", "PDFファイル選択取込"
    AddBtn ws, 240, 285, 200, 32, "貼り付けたPDFを取り込む", "貼付PDF取込"
    AddBtn ws, 450, 285, 120, 32, "再計算", "再計算"
    AddBtn ws, 580, 285, 120, 32, "新年度開始", "新年度開始"
    ws.Activate
    MsgBox "ボタンを設置しました。" & vbCrLf & vbCrLf & _
           "「PDF貼付」シートの中ほどにボタンが4つ表示されています。", vbInformation
End Sub

Private Sub AddBtn(ws As Worksheet, l As Single, t As Single, _
                   w As Single, h As Single, cap As String, macroName As String)
    Dim b As Button
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = macroName
End Sub


'==============================================================
' 取込
'==============================================================
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
    ImportMonthly paths
End Sub

Public Sub 貼付PDF取込()
    Dim paths As Collection
    Set paths = ExtractEmbeddedPdfs()      ' modEmbedded
    If paths Is Nothing Then Exit Sub
    If paths.Count = 0 Then
        MsgBox "貼り付けられたPDFが見つかりませんでした。" & vbCrLf & _
               "PDFを貼り付けてブックを保存してから、もう一度実行してください。", vbExclamation
        Exit Sub
    End If
    ImportMonthly paths
End Sub

'--- PDFを月別に取り込む(同じ月は入れ替え) ---
Private Sub ImportMonthly(paths As Collection)
    Dim p As Variant, okCnt As Long, addCnt As Long, ngList As String
    Application.ScreenUpdating = False
    On Error GoTo done

    For Each p In paths
        Dim rows As Collection
        Set rows = Nothing
        If ParsePdf(CStr(p), rows) Then          ' modEngine
            If rows.Count > 0 Then
                Dim mi As Long
                mi = MonthIdxFromDate(CStr(rows(1)(9)))
                If mi = 0 Then
                    ngList = ngList & vbCrLf & "  " & p & " (測定日が読み取れません)"
                Else
                    RemoveMonth mi
                    AppendAll rows, mi
                    LogHistory mi, rows
                    okCnt = okCnt + 1
                    addCnt = addCnt + rows.Count
                End If
            End If
        Else
            ngList = ngList & vbCrLf & "  " & p
        End If
    Next

    再計算Silent
done:
    Application.ScreenUpdating = True
    Dim msg As String
    msg = okCnt & " 件のPDFから " & addCnt & " 行を取り込みました。"
    If ngList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "取り込めなかったPDF:" & ngList
    End If
    MsgBox msg, IIf(ngList = "", vbInformation, vbExclamation)
End Sub

'--- 「2026/06/04」等から年度内月番号(4月=1 … 3月=12)を得る ---
Private Function MonthIdxFromDate(s As String) As Long
    Dim parts() As String, m As Long
    s = Replace(Replace(s, "-", "/"), ".", "/")
    parts = Split(s, "/")
    If UBound(parts) < 1 Then Exit Function
    m = Val(parts(1))
    If m < 1 Or m > 12 Then Exit Function
    MonthIdxFromDate = ((m - 4 + 12) Mod 12) + 1
End Function

Public Function MonthLabel(mi As Long) As String
    Dim m As Long
    m = ((mi - 1 + 3) Mod 12) + 1
    MonthLabel = m & "月"
End Function

'--- 指定した月の既存行を「全データ」から削除(取り直しの二重計上防止) ---
Private Sub RemoveMonth(mi As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_ALL)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < D_START Then Exit Sub

    Dim arr As Variant
    arr = ws.Range(ws.Cells(D_START, 1), ws.Cells(lastR, ALL_COLS)).Value
    Dim keep() As Variant
    ReDim keep(1 To UBound(arr, 1), 1 To ALL_COLS)
    Dim r As Long, k As Long, c As Long
    Dim lbl As String
    lbl = MonthLabel(mi)
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, ALL_COLS)) <> lbl Then
            k = k + 1
            For c = 1 To ALL_COLS
                keep(k, c) = arr(r, c)
            Next
        End If
    Next

    ws.Range(ws.Cells(D_START, 1), ws.Cells(lastR, ALL_COLS)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(D_START, 1), ws.Cells(D_START + k - 1, ALL_COLS)).Value = keep
    End If

    ' 取込履歴からも該当月を消す
    Set ws = ThisWorkbook.Worksheets(SH_HIST)
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = lastR To D_START Step -1
        If CStr(ws.Cells(r, 2).Value) = lbl Then ws.Rows(r).Delete
    Next
End Sub

'--- 「全データ」へ追記(13列目に月ラベル) ---
Private Sub AppendAll(rows As Collection, mi As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_ALL)
    Dim startR As Long
    startR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If startR < D_START Then startR = D_START

    Dim out() As Variant
    ReDim out(1 To rows.Count, 1 To ALL_COLS)
    Dim item As Variant, k As Long, i As Long
    For Each item In rows
        k = k + 1
        For i = 1 To 12
            out(k, i) = item(i)
        Next
        out(k, ALL_COLS) = MonthLabel(mi)
    Next
    With ws.Range(ws.Cells(startR, 1), ws.Cells(startR + rows.Count - 1, ALL_COLS))
        .Value = out
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(153, 153, 153)
        .Font.Name = "游ゴシック"
        .Font.Size = 10
    End With
End Sub

Private Sub LogHistory(mi As Long, rows As Collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_HIST)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1
    If r < D_START Then r = D_START
    ws.Cells(r, 2).Value = MonthLabel(mi)
    ws.Cells(r, 3).Value = rows(1)(10)          ' 線別
    ws.Cells(r, 4).Value = rows(1)(9)           ' 測定日
    ws.Cells(r, 5).Value = rows(1)(12)          ' 元ファイル
    ws.Cells(r, 6).Value = rows.Count
    ws.Cells(r, 7).Value = Format(Now, "yyyy/mm/dd hh:nn")
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Font.Name = "游ゴシック"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Font.Size = 10
End Sub


'==============================================================
' 再計算(中核)
'==============================================================
Public Sub 再計算()
    再計算Silent
    MsgBox "推移表・グラフを更新しました。", vbInformation
End Sub

Public Sub 再計算Silent()
    Dim tgt As Double, lim As Double, tol As Double
    With ThisWorkbook.Worksheets(SH_SET)
        tgt = Val(.Range("C3").Value)
        lim = Val(.Range("C4").Value)
        tol = Val(.Range("C6").Value) / 1000#      ' m → km
    End With
    If tol <= 0 Then tol = 0.02

    Application.ScreenUpdating = False

    ' --- 0) 手入力メモを先に退避する(順序が重要) ---
    Dim memos() As TMemo, nMemo As Long
    HarvestMemos memos, nMemo, tol

    ' --- 1) 全データを読む ---
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_ALL)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ClearTrend
    If lastR < D_START Then
        WriteMemoSheet memos, nMemo
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(D_START, 1), ws.Cells(lastR, ALL_COLS)).Value

    ' --- 2) 点の一覧を作る ---
    Dim n As Long
    n = UBound(arr, 1)
    Dim pLine() As String, pKind() As String, pKm() As Double
    Dim pVal() As Double, pMi() As Long, pBiko() As String
    ReDim pLine(1 To n): ReDim pKind(1 To n): ReDim pKm(1 To n)
    ReDim pVal(1 To n): ReDim pMi(1 To n): ReDim pBiko(1 To n)

    Dim monthUsed(1 To 12) As Boolean
    Dim r As Long, cnt As Long, v As Double, mi As Long
    For r = 1 To n
        If Len(CStr(arr(r, 4))) > 0 Then
            v = -1
            Dim kind As String
            kind = ""
            If IsNumeric(arr(r, 5)) And Len(CStr(arr(r, 5))) > 0 Then
                v = CDbl(arr(r, 5)): kind = "上下動"
            ElseIf IsNumeric(arr(r, 6)) And Len(CStr(arr(r, 6))) > 0 Then
                v = CDbl(arr(r, 6)): kind = "左右動"
            End If
            mi = MonthIdxFromLabel(CStr(arr(r, ALL_COLS)))
            If v > 0 And mi > 0 Then
                cnt = cnt + 1
                pLine(cnt) = CStr(arr(r, 10))
                pKind(cnt) = kind
                pKm(cnt) = CDbl(arr(r, 4))
                pVal(cnt) = v
                pMi(cnt) = mi
                pBiko(cnt) = CStr(arr(r, 8))
                monthUsed(mi) = True
            End If
        End If
    Next
    If cnt = 0 Then
        WriteMemoSheet memos, nMemo
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- 3) 線別+区分+キロ程 で並べ替え ---
    Dim idx() As Long, key() As String
    ReDim idx(1 To cnt): ReDim key(1 To cnt)
    For r = 1 To cnt
        idx(r) = r
        key(r) = pLine(r) & "|" & pKind(r) & "|" & Format(pKm(r), "0000.000")
    Next
    QuickSortIdx key, idx, 1, cnt

    ' --- 4) 鎖方式でグループ化 ---
    Dim cl() As TCluster, nc As Long
    ReDim cl(1 To cnt)
    Dim i As Long, j As Long, prevKm As Double
    For i = 1 To cnt
        j = idx(i)
        Dim isNew As Boolean
        isNew = True
        If nc > 0 Then
            If cl(nc).line = pLine(j) And cl(nc).kind = pKind(j) Then
                If pKm(j) - prevKm <= tol + 0.0000001 Then isNew = False
            End If
        End If
        If isNew Then
            nc = nc + 1
            cl(nc).line = pLine(j)
            cl(nc).kind = pKind(j)
            cl(nc).lo = pKm(j)
            cl(nc).maxV = -1
            cl(nc).biko = ""
        End If
        cl(nc).hi = pKm(j)
        prevKm = pKm(j)
        If pVal(j) > cl(nc).mon(pMi(j)) Then cl(nc).mon(pMi(j)) = pVal(j)
        If pVal(j) > cl(nc).maxV Then
            cl(nc).maxV = pVal(j)
            cl(nc).rep = pKm(j)
        End If
        If Len(pBiko(j)) > 0 And Len(cl(nc).biko) = 0 Then cl(nc).biko = pBiko(j)
    Next

    ' --- 5) 推移表へ書き出し ---
    Dim rep() As TRepair, nRep As Long
    ReadRepairs rep, nRep
    WriteTrend cl, nc, monthUsed, memos, nMemo, rep, nRep, tol
    WriteMemoSheet memos, nMemo

    ' --- 6) ①② を作り直す ---
    RebuildExtracts arr, n, tgt, lim

    ' --- 7) グラフ ---
    グラフ更新Silent cl, nc, monthUsed, tgt, lim

    Application.ScreenUpdating = True
End Sub

Private Function MonthIdxFromLabel(s As String) As Long
    Dim m As Long
    m = Val(Replace(s, "月", ""))
    If m < 1 Or m > 12 Then Exit Function
    MonthIdxFromLabel = ((m - 4 + 12) Mod 12) + 1
End Function


'==============================================================
' 手入力メモ
'==============================================================
'--- 箇所メモシート + 推移表のメモ列 から拾う ---
Private Sub HarvestMemos(memos() As TMemo, nMemo As Long, tol As Double)
    ReDim memos(1 To 2000)
    nMemo = 0

    ' 正本: 箇所メモシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_MEMO)
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = D_START To lastR
        If Len(CStr(ws.Cells(r, 2).Value)) > 0 And _
           Len(CStr(ws.Cells(r, 4).Value)) > 0 Then
            If nMemo >= UBound(memos) Then Exit For
            nMemo = nMemo + 1
            memos(nMemo).line = CStr(ws.Cells(r, 2).Value)
            memos(nMemo).km = Val(ws.Cells(r, 3).Value)
            memos(nMemo).text = CStr(ws.Cells(r, 4).Value)
        End If
    Next

    ' 推移表に直接入力された分。既にあるものは拾わない(拾うと増殖する)
    Set ws = ThisWorkbook.Worksheets(SH_TREND)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = D_START To lastR
        Dim txt As String, ln As String, km As Double
        txt = CStr(ws.Cells(r, C_MEMO).Value)
        ln = CStr(ws.Cells(r, 2).Value)
        km = Val(ws.Cells(r, 4).Value)
        If Len(txt) > 0 And Len(ln) > 0 Then
            Dim parts() As String, p As Long
            parts = Split(txt, "/")
            For p = 0 To UBound(parts)
                Dim one As String
                one = Trim(parts(p))
                If Len(one) > 0 Then
                    If Not MemoKnown(memos, nMemo, ln, km, one, tol) Then
                        If nMemo >= UBound(memos) Then Exit Sub
                        nMemo = nMemo + 1
                        memos(nMemo).line = ln
                        memos(nMemo).km = km
                        memos(nMemo).text = one
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function MemoKnown(memos() As TMemo, nMemo As Long, ln As String, _
                           km As Double, txt As String, tol As Double) As Boolean
    Dim i As Long
    For i = 1 To nMemo
        If memos(i).line = ln And memos(i).text = txt Then
            If Abs(memos(i).km - km) <= tol Then
                MemoKnown = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub WriteMemoSheet(memos() As TMemo, nMemo As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_MEMO)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastR >= D_START Then
        ws.Range(ws.Cells(D_START, 2), ws.Cells(lastR, 4)).ClearContents
    End If
    Dim i As Long
    For i = 1 To nMemo
        ws.Cells(D_START + i - 1, 2).Value = memos(i).line
        ws.Cells(D_START + i - 1, 3).Value = memos(i).km
        ws.Cells(D_START + i - 1, 4).Value = memos(i).text
    Next
End Sub

Private Sub ReadRepairs(rep() As TRepair, nRep As Long)
    ReDim rep(1 To 500)
    nRep = 0
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_REPAIR)
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = D_START To lastR
        If Len(CStr(ws.Cells(r, 2).Value)) > 0 Then
            nRep = nRep + 1
            rep(nRep).line = CStr(ws.Cells(r, 2).Value)
            rep(nRep).km = Val(ws.Cells(r, 3).Value)
            rep(nRep).d = ws.Cells(r, 4).Value
            rep(nRep).text = CStr(ws.Cells(r, 5).Value)
        End If
    Next
End Sub


'==============================================================
' 推移表の書き出し
'==============================================================
Private Sub ClearTrend()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_TREND)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= D_START Then
        With ws.Range(ws.Cells(D_START, 1), ws.Cells(lastR, TREND_COLS))
            .ClearContents
            .Borders.LineStyle = xlNone
        End With
    End If
End Sub

Private Sub WriteTrend(cl() As TCluster, nc As Long, monthUsed() As Boolean, _
                       memos() As TMemo, nMemo As Long, _
                       rep() As TRepair, nRep As Long, tol As Double)
    If nc = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_TREND)

    Dim out() As Variant
    ReDim out(1 To nc, 1 To TREND_COLS)
    Dim i As Long, m As Long, k As Long

    For i = 1 To nc
        out(i, 1) = i
        out(i, 2) = cl(i).line
        out(i, 3) = cl(i).kind
        out(i, 4) = cl(i).rep
        out(i, 5) = cl(i).lo
        out(i, 6) = cl(i).hi
        out(i, 7) = cl(i).biko

        ' メモ: 範囲±距離に入るものを「 / 」で連結(併合でも失わない)
        Dim memo As String
        memo = ""
        For k = 1 To nMemo
            If memos(k).line = cl(i).line Then
                If memos(k).km >= cl(i).lo - tol And memos(k).km <= cl(i).hi + tol Then
                    If InStr(memo, memos(k).text) = 0 Then
                        If Len(memo) > 0 Then memo = memo & " / "
                        memo = memo & memos(k).text
                    End If
                End If
            End If
        Next
        out(i, C_MEMO) = memo

        For m = 1 To 12
            If cl(i).mon(m) > 0 Then
                out(i, C_MON0 + m - 1) = cl(i).mon(m)
            ElseIf monthUsed(m) Then
                out(i, C_MON0 + m - 1) = "－"       ' 0.15G未満に収まっていた
            End If
        Next
        out(i, C_YMAX) = "=MAX(I" & (D_START + i - 1) & ":T" & (D_START + i - 1) & ")"
        out(i, C_PREV) = PrevYearValue(cl(i).line, cl(i).lo, cl(i).hi, tol)

        For k = 1 To nRep
            If rep(k).line = cl(i).line Then
                If rep(k).km >= cl(i).lo - tol And rep(k).km <= cl(i).hi + tol Then
                    out(i, 23) = rep(k).d
                    out(i, 24) = rep(k).text
                End If
            End If
        Next
    Next

    With ws.Range(ws.Cells(D_START, 1), ws.Cells(D_START + nc - 1, TREND_COLS))
        .Value = out
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(153, 153, 153)
        .Font.Name = "游ゴシック"
        .Font.Size = 10
    End With
End Sub

'--- 前年度実績シートから、同じ位置の年間最大を拾う ---
Private Function PrevYearValue(ln As String, lo As Double, hi As Double, _
                               tol As Double) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_PREV)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim best As Double
    best = -1
    For r = D_START To lastR
        If CStr(ws.Cells(r, 2).Value) = ln Then
            Dim km As Double
            km = Val(ws.Cells(r, 4).Value)
            If km >= lo - tol And km <= hi + tol Then
                If Val(ws.Cells(r, C_YMAX).Value) > best Then
                    best = Val(ws.Cells(r, C_YMAX).Value)
                End If
            End If
        End If
    Next
    If best > 0 Then PrevYearValue = best
End Function


'==============================================================
' ①② の再作成
'==============================================================
Private Sub RebuildExtracts(arr As Variant, n As Long, tgt As Double, lim As Double)
    Dim over() As Variant, tg() As Variant
    ReDim over(1 To n, 1 To ALL_COLS)
    ReDim tg(1 To n, 1 To ALL_COLS)
    Dim nO As Long, nT As Long, r As Long, c As Long, v As Double

    For r = 1 To n
        v = -1
        If IsNumeric(arr(r, 5)) And Len(CStr(arr(r, 5))) > 0 Then
            v = CDbl(arr(r, 5))
        ElseIf IsNumeric(arr(r, 6)) And Len(CStr(arr(r, 6))) > 0 Then
            v = CDbl(arr(r, 6))
        End If
        If v > 0 Then
            If v >= lim Then
                nO = nO + 1
                For c = 1 To ALL_COLS
                    over(nO, c) = arr(r, c)
                Next
            ElseIf v >= tgt Then
                nT = nT + 1
                For c = 1 To ALL_COLS
                    tg(nT, c) = arr(r, c)
                Next
            End If
        End If
    Next
    PutRows SH_OVER, over, nO
    PutRows SH_TGT, tg, nT
End Sub

Private Sub PutRows(sheetName As String, data As Variant, cnt As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= D_START Then
        With ws.Range(ws.Cells(D_START, 1), ws.Cells(lastR, ALL_COLS))
            .ClearContents
            .Borders.LineStyle = xlNone
        End With
    End If
    If cnt = 0 Then Exit Sub

    Dim out() As Variant
    ReDim out(1 To cnt, 1 To ALL_COLS)
    Dim r As Long, c As Long
    For r = 1 To cnt
        For c = 1 To ALL_COLS
            out(r, c) = data(r, c)
        Next
    Next
    With ws.Range(ws.Cells(D_START, 1), ws.Cells(D_START + cnt - 1, ALL_COLS))
        .Value = out
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(153, 153, 153)
        .Font.Name = "游ゴシック"
        .Font.Size = 10
    End With
End Sub


'==============================================================
' グラフ
'==============================================================
Public Sub グラフ更新()
    再計算Silent
End Sub

Private Sub グラフ更新Silent(cl() As TCluster, nc As Long, _
                             monthUsed() As Boolean, tgt As Double, lim As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_CHART)
    Dim sel As Long, j As Long, k As Long

    ' --- 選択箇所(40～43行) ---
    sel = Val(ws.Range("C2").Value)
    If sel < 1 Or sel > nc Then
        sel = 1
        ws.Range("C2").Value = IIf(nc > 0, 1, "")
    End If
    ws.Range(ws.Cells(41, 1), ws.Cells(43, 13)).ClearContents
    If nc > 0 Then
        ws.Cells(41, 1).Value = "No" & sel & " " & cl(sel).line & " " & _
                                Format(cl(sel).rep, "0.000") & "(" & cl(sel).kind & ")"
        For j = 1 To 12
            If cl(sel).mon(j) > 0 Then ws.Cells(41, 1 + j).Value = cl(sel).mon(j)
            ws.Cells(42, 1 + j).Value = tgt
            ws.Cells(43, 1 + j).Value = lim
        Next
    End If
    ws.Cells(42, 1).Value = "目標値"
    ws.Cells(43, 1).Value = "基準値"

    ' --- ワースト5(51～56行): 最新の取込月の値が大きい順 ---
    Dim latest As Long
    For j = 12 To 1 Step -1
        If monthUsed(j) Then
            latest = j
            Exit For
        End If
    Next
    ws.Range(ws.Cells(51, 1), ws.Cells(56, 13)).ClearContents
    If latest > 0 And nc > 0 Then
        Dim topIdx(1 To 5) As Long, topVal(1 To 5) As Double
        Dim i As Long, s As Long
        For i = 1 To nc
            If cl(i).mon(latest) > 0 Then
                For s = 1 To 5
                    If cl(i).mon(latest) > topVal(s) Then
                        Dim m2 As Long
                        For m2 = 5 To s + 1 Step -1
                            topVal(m2) = topVal(m2 - 1)
                            topIdx(m2) = topIdx(m2 - 1)
                        Next
                        topVal(s) = cl(i).mon(latest)
                        topIdx(s) = i
                        Exit For
                    End If
                Next
            End If
        Next
        For k = 1 To 5
            If topIdx(k) > 0 Then
                i = topIdx(k)
                ws.Cells(50 + k, 1).Value = "No" & i & " " & _
                    Format(cl(i).rep, "0.000") & "(" & cl(i).kind & ")"
                For j = 1 To 12
                    If cl(i).mon(j) > 0 Then ws.Cells(50 + k, 1 + j).Value = cl(i).mon(j)
                Next
            Else
                ws.Cells(50 + k, 1).Value = "(ワースト" & k & ")"
            End If
        Next
    End If
    ws.Cells(56, 1).Value = "基準値"
    For j = 1 To 12
        ws.Cells(56, 1 + j).Value = lim
    Next
End Sub


'==============================================================
' 新年度開始
'==============================================================
Public Sub 新年度開始()
    If MsgBox("今年度のデータを「前年度実績」へ退避し、月データをすべて消します。" & vbCrLf & _
              "この操作は元に戻せません。" & vbCrLf & vbCrLf & _
              "※ 先にこのファイルを新年度用の名前でコピー・保存しておいてください。" & vbCrLf & _
              "続けますか?", vbYesNo + vbExclamation) = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Dim src As Worksheet, dst As Worksheet
    Set src = ThisWorkbook.Worksheets(SH_TREND)
    Set dst = ThisWorkbook.Worksheets(SH_PREV)

    Dim lastR As Long
    lastR = dst.Cells(dst.Rows.Count, 1).End(xlUp).Row
    If lastR >= D_START Then
        dst.Range(dst.Cells(D_START, 1), dst.Cells(lastR, TREND_COLS)).ClearContents
    End If

    lastR = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    If lastR >= D_START Then
        dst.Range(dst.Cells(D_START, 1), dst.Cells(D_START + lastR - D_START, TREND_COLS)).Value = _
            src.Range(src.Cells(D_START, 1), src.Cells(lastR, TREND_COLS)).Value
    End If

    ' 月データのクリア(箇所メモ・施工記録は残す)
    ClearSheetData SH_ALL, ALL_COLS
    ClearSheetData SH_OVER, ALL_COLS
    ClearSheetData SH_TGT, ALL_COLS
    ClearSheetData SH_TREND, TREND_COLS
    ClearSheetData SH_HIST, 7

    With ThisWorkbook.Worksheets(SH_SET)
        .Range("C7").Value = Val(.Range("C7").Value) + 1
    End With

    Application.ScreenUpdating = True
    MsgBox "新年度の準備ができました。" & vbCrLf & _
           "「箇所メモ」と「施工記録」は引き継いでいます。", vbInformation
End Sub

Private Sub ClearSheetData(sheetName As String, cols As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastR As Long, c1 As Long
    c1 = IIf(sheetName = SH_HIST, 2, 1)
    lastR = ws.Cells(ws.Rows.Count, c1).End(xlUp).Row
    If lastR >= D_START Then
        With ws.Range(ws.Cells(D_START, c1), ws.Cells(lastR, cols))
            .ClearContents
            .Borders.LineStyle = xlNone
        End With
    End If
End Sub


'==============================================================
' 並べ替え(文字列キー)
'==============================================================
Private Sub QuickSortIdx(key() As String, idx() As Long, lo As Long, hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long
    Dim pivot As String, t As Long
    i = lo: j = hi
    pivot = key(idx((lo + hi) \ 2))
    Do While i <= j
        Do While key(idx(i)) < pivot
            i = i + 1
        Loop
        Do While key(idx(j)) > pivot
            j = j - 1
        Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    QuickSortIdx key, idx, lo, j
    QuickSortIdx key, idx, i, hi
End Sub
