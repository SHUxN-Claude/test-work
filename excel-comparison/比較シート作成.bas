Attribute VB_Name = "Mod比較シート作成"
Option Explicit

'============================================================
' 前回/今回シートから「比較」「抽出_要対応」「チェック」を再作成するマクロ
'
' 使い方:
'   1. このブックの「前回」「今回」シートに検査データを貼り替える
'      (1行目が見出し、2行目以降がデータであること)
'   2. Alt+F8 → 「比較シート作成」を選んで実行
'   3. 「比較」「抽出_要対応」「チェック」シートが作り直される
'      ※実行後は必ず「チェック」シートの警告を確認すること
'
' 判定ルール:
'   ・今回の目標値F=1 の測定値セル → 黄色
'   ・目標値F=1 かつ 変位進み量(|今回-前回|)が10mm以上15mm未満 → オレンジ
'       (30日以内に調査 → 変位進み確認時は60日以内に処置)
'   ・目標値F=1 かつ 変位進み量が15mm以上 → 赤
'       (15日以内に調査 → 変位進み確認時は30日以内に処置)
'============================================================

Private Const SH_PREV As String = "前回"
Private Const SH_CURR As String = "今回"
Private Const SH_OUT As String = "比較"
Private Const SH_EXT As String = "抽出_要対応"
Private Const SH_CHK As String = "チェック"

Private Const TH_LOW As Double = 10#    ' 変位進み量 下側しきい値(mm)
Private Const TH_HIGH As Double = 15#   ' 変位進み量 上側しきい値(mm)

Private Const N_ITEMS As Long = 9
Private Const N_BASE As Long = 12       ' 基本情報の列数

Public Sub 比較シート作成()
    Dim t0 As Double: t0 = Timer
    Dim clrYellow As Long, clrOrange As Long, clrRed As Long
    clrYellow = RGB(255, 242, 168)
    clrOrange = RGB(255, 192, 0)
    clrRed = RGB(255, 107, 107)

    ' ---- 比較項目の定義(見出し名で列を探すので列順が変わっても動く) ----
    Dim itemName(1 To N_ITEMS) As String, itemFlag(1 To N_ITEMS) As String
    itemName(1) = "高低（左）":   itemFlag(1) = "高低（左）目標値F"
    itemName(2) = "高低（右）":   itemFlag(2) = "高低（右）目標値F"
    itemName(3) = "通り（左）":   itemFlag(3) = "通り（左）目標値F"
    itemName(4) = "通り（右）":   itemFlag(4) = "通り（右）目標値F"
    itemName(5) = "水準実測値":   itemFlag(5) = ""                    ' 目標値Fなし(参考表示)
    itemName(6) = "水準":         itemFlag(6) = "水準目標値F"
    itemName(7) = "平面性":       itemFlag(7) = "平面性目標値F"
    itemName(8) = "軌間実測値":   itemFlag(8) = "軌間実測値目標値F"
    itemName(9) = "軌間":         itemFlag(9) = "軌間目標F"

    ' 突き合わせキー(キロ程だけでなく線名等も含めるので複数線区でも安全)
    Dim keyHdr(1 To 5) As String
    keyHdr(1) = "線名コード": keyHdr(2) = "線別コード": keyHdr(3) = "番線番号"
    keyHdr(4) = "本準側コード": keyHdr(5) = "キロ程"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    On Error GoTo FAIL

    ' ---- 前回/今回シートの読み込み ----
    Dim wsP As Worksheet, wsC As Worksheet
    If Not SheetExists(SH_PREV) Or Not SheetExists(SH_CURR) Then
        MsgBox "「" & SH_PREV & "」「" & SH_CURR & "」シートが見つかりません。", vbCritical
        GoTo FIN
    End If
    Set wsP = ThisWorkbook.Worksheets(SH_PREV)
    Set wsC = ThisWorkbook.Worksheets(SH_CURR)

    Dim dP As Variant, dC As Variant
    dP = wsP.UsedRange.Value
    dC = wsC.UsedRange.Value
    If Not IsArray(dP) Or Not IsArray(dC) Then
        MsgBox "前回または今回シートにデータがありません。", vbCritical
        GoTo FIN
    End If

    Dim problems As New Collection

    ' 見出し位置の取得
    Dim hP As Object, hC As Object
    Set hP = HeaderMap(dP): Set hC = HeaderMap(dC)
    Dim need As Variant, i As Long
    need = Array("No", "実績番号", "線名", "線路名", "本準側名", "WB", "キロ程", "測定日")
    For i = LBound(need) To UBound(need)
        If Not hP.Exists(need(i)) Then problems.Add "前回シートに見出し「" & need(i) & "」がありません"
        If Not hC.Exists(need(i)) Then problems.Add "今回シートに見出し「" & need(i) & "」がありません"
    Next
    For i = 1 To 5
        If Not hP.Exists(keyHdr(i)) Then problems.Add "前回シートに見出し「" & keyHdr(i) & "」がありません"
        If Not hC.Exists(keyHdr(i)) Then problems.Add "今回シートに見出し「" & keyHdr(i) & "」がありません"
    Next
    For i = 1 To N_ITEMS
        If Not hP.Exists(itemName(i)) Then problems.Add "前回シートに見出し「" & itemName(i) & "」がありません"
        If Not hC.Exists(itemName(i)) Then problems.Add "今回シートに見出し「" & itemName(i) & "」がありません"
        If itemFlag(i) <> "" Then
            If Not hC.Exists(itemFlag(i)) Then problems.Add "今回シートに見出し「" & itemFlag(i) & "」がありません"
        End If
    Next
    If problems.Count > 0 Then
        Dim msg As String, p As Variant
        For Each p In problems: msg = msg & vbCrLf & "・" & p: Next
        MsgBox "見出しが不足しているため中止しました。" & msg, vbCritical
        GoTo FIN
    End If

    ' ---- 前回データを辞書化 ----
    Dim dictPrev As Object: Set dictPrev = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String, dupPrev As Long, nPrev As Long
    Dim jisPrev As Object: Set jisPrev = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(dP, 1)
        If Trim$(CStr(dP(r, hP("実績番号")))) <> "" Then
            nPrev = nPrev + 1
            jisPrev(Trim$(CStr(dP(r, hP("実績番号"))))) = True
            k = MakeKey(dP, r, hP, keyHdr)
            If dictPrev.Exists(k) Then dupPrev = dupPrev + 1
            dictPrev(k) = r
        End If
    Next

    ' ---- 出力配列の準備 ----
    Dim nCols As Long: nCols = N_BASE
    Dim itemCol(1 To N_ITEMS) As Long   ' 各項目の先頭列(前回値の列)
    For i = 1 To N_ITEMS
        itemCol(i) = nCols + 1
        nCols = nCols + 3                       ' 前回/今回/差分
        If itemFlag(i) <> "" Then nCols = nCols + 1  ' 目標F
    Next

    Dim nCurr As Long
    For r = 2 To UBound(dC, 1)
        If Trim$(CStr(dC(r, hC("実績番号")))) <> "" Then nCurr = nCurr + 1
    Next
    If nCurr = 0 Then
        MsgBox "今回シートにデータ行がありません。", vbCritical
        GoTo FIN
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To nCurr + 1, 1 To nCols)
    outArr(1, 1) = "No": outArr(1, 2) = "線名": outArr(1, 3) = "線路名"
    outArr(1, 4) = "キロ程": outArr(1, 5) = "判定": outArr(1, 6) = "該当項目"
    outArr(1, 7) = "前回実績番号": outArr(1, 8) = "今回実績番号"
    outArr(1, 9) = "前回測定日": outArr(1, 10) = "今回測定日"
    outArr(1, 11) = "WB": outArr(1, 12) = "本準側名"
    For i = 1 To N_ITEMS
        outArr(1, itemCol(i)) = itemName(i) & "前回"
        outArr(1, itemCol(i) + 1) = itemName(i) & "今回"
        outArr(1, itemCol(i) + 2) = itemName(i) & "差分"
        If itemFlag(i) <> "" Then outArr(1, itemCol(i) + 3) = itemName(i) & "目標F"
    Next

    ' 色付け予定セル(行,列,色)
    Dim fillR() As Long, fillC() As Long, fillClr() As Long
    Dim nFill As Long, capFill As Long: capFill = 1024
    ReDim fillR(1 To capFill): ReDim fillC(1 To capFill): ReDim fillClr(1 To capFill)

    Dim extRows As New Collection      ' 抽出対象(出力シートの行番号)
    Dim unmatchedCurr As New Collection, seenCurr As Object
    Set seenCurr = CreateObject("Scripting.Dictionary")
    Dim dupCurr As Long, nonNum As Long
    Dim flagCnt(1 To N_ITEMS) As Long, lowCnt(1 To N_ITEMS) As Long, highCnt(1 To N_ITEMS) As Long
    Dim jisCurr As Object: Set jisCurr = CreateObject("Scripting.Dictionary")

    ' ---- メインループ(今回の行を基準に前回を突き合わせ) ----
    Dim o As Long: o = 1
    Dim pr As Variant, pRow As Long
    Dim pv As Variant, cv As Variant, pf As Variant, cf As Variant
    Dim diff As Variant, flag As Variant, ad As Double
    Dim worst As Long, hitList As String
    For r = 2 To UBound(dC, 1)
        If Trim$(CStr(dC(r, hC("実績番号")))) <> "" Then
            o = o + 1
            jisCurr(Trim$(CStr(dC(r, hC("実績番号"))))) = True
            k = MakeKey(dC, r, hC, keyHdr)
            If seenCurr.Exists(k) Then dupCurr = dupCurr + 1
            seenCurr(k) = True
            pRow = 0
            If dictPrev.Exists(k) Then pRow = dictPrev(k)

            outArr(o, 1) = ToNum(dC(r, hC("No")))
            outArr(o, 2) = dC(r, hC("線名"))
            outArr(o, 3) = dC(r, hC("線路名"))
            outArr(o, 4) = ToNum(dC(r, hC("キロ程")))
            If pRow > 0 Then outArr(o, 7) = ToNum(dP(pRow, hP("実績番号")))
            outArr(o, 8) = ToNum(dC(r, hC("実績番号")))
            If pRow > 0 Then outArr(o, 9) = ToNum(dP(pRow, hP("測定日")))
            outArr(o, 10) = ToNum(dC(r, hC("測定日")))
            outArr(o, 11) = dC(r, hC("WB"))
            outArr(o, 12) = dC(r, hC("本準側名"))

            worst = 0: hitList = ""
            For i = 1 To N_ITEMS
                cv = ToNum(dC(r, hC(itemName(i))))
                If pRow > 0 Then pv = ToNum(dP(pRow, hP(itemName(i)))) Else pv = Empty
                outArr(o, itemCol(i)) = pv
                outArr(o, itemCol(i) + 1) = cv
                ' 数値化できない値の検知(空欄は除く)
                If Not IsEmpty(cv) And Not IsNumeric(cv) Then nonNum = nonNum + 1
                If Not IsEmpty(pv) And Not IsNumeric(pv) Then nonNum = nonNum + 1
                diff = Empty
                If IsNumeric2(pv) And IsNumeric2(cv) Then
                    diff = Application.WorksheetFunction.Round(CDbl(cv) - CDbl(pv), 1)
                End If
                outArr(o, itemCol(i) + 2) = diff
                If itemFlag(i) <> "" Then
                    flag = ToNum(dC(r, hC(itemFlag(i))))
                    outArr(o, itemCol(i) + 3) = flag
                    If IsNumeric2(flag) Then
                        If CDbl(flag) = 1 Then
                            flagCnt(i) = flagCnt(i) + 1
                            AddFill fillR, fillC, fillClr, nFill, capFill, o, itemCol(i) + 1, clrYellow
                            If IsNumeric2(diff) Then
                                ad = Abs(CDbl(diff))
                                If ad >= TH_HIGH Then
                                    highCnt(i) = highCnt(i) + 1
                                    If worst < 2 Then worst = 2
                                    hitList = hitList & IIf(hitList = "", "", "、") & itemName(i)
                                    AddFill fillR, fillC, fillClr, nFill, capFill, o, itemCol(i) + 1, clrRed
                                    AddFill fillR, fillC, fillClr, nFill, capFill, o, itemCol(i) + 2, clrRed
                                ElseIf ad >= TH_LOW Then
                                    lowCnt(i) = lowCnt(i) + 1
                                    If worst < 1 Then worst = 1
                                    hitList = hitList & IIf(hitList = "", "", "、") & itemName(i)
                                    AddFill fillR, fillC, fillClr, nFill, capFill, o, itemCol(i) + 1, clrOrange
                                    AddFill fillR, fillC, fillClr, nFill, capFill, o, itemCol(i) + 2, clrOrange
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If worst = 2 Then
                outArr(o, 5) = "15mm以上"
                AddFill fillR, fillC, fillClr, nFill, capFill, o, 5, clrRed
                extRows.Add o
            ElseIf worst = 1 Then
                outArr(o, 5) = "10mm以上15mm未満"
                AddFill fillR, fillC, fillClr, nFill, capFill, o, 5, clrOrange
                extRows.Add o
            End If
            outArr(o, 6) = hitList
            If pRow = 0 Then
                unmatchedCurr.Add outArr(o, 4)
                outArr(o, 6) = IIf(hitList = "", "(前回データなし)", hitList & "、(前回データなし)")
            End If
        End If
    Next

    ' 前回にあって今回にないキー
    Dim unmatchedPrev As New Collection
    Dim kk As Variant
    For Each kk In dictPrev.Keys
        If Not seenCurr.Exists(CStr(kk)) Then
            unmatchedPrev.Add ToNum(dP(dictPrev(kk), hP("キロ程")))
        End If
    Next

    ' ---- 出力シートの作り直し ----
    DeleteSheetIfExists SH_OUT
    DeleteSheetIfExists SH_EXT
    DeleteSheetIfExists SH_CHK
    Dim wsO As Worksheet, wsE As Worksheet, wsK As Worksheet
    Set wsO = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsO.Name = SH_OUT
    Set wsE = ThisWorkbook.Worksheets.Add(After:=wsO)
    wsE.Name = SH_EXT
    Set wsK = ThisWorkbook.Worksheets.Add(After:=wsE)
    wsK.Name = SH_CHK

    With wsO
        .Range(.Cells(1, 1), .Cells(nCurr + 1, nCols)).Value = outArr
        ' 見出し書式
        With .Range(.Cells(1, 1), .Cells(1, nCols))
            .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242)
        End With
        ' 数値の表示形式(小数1桁)
        .Range(.Cells(2, 4), .Cells(nCurr + 1, 4)).NumberFormat = "0.0"   ' キロ程
        For i = 1 To N_ITEMS
            .Range(.Cells(2, itemCol(i)), .Cells(nCurr + 1, itemCol(i) + 2)).NumberFormat = "0.0"
        Next
        ' 色付け
        For i = 1 To nFill
            .Cells(fillR(i), fillC(i)).Interior.Color = fillClr(i)
        Next
        .Range(.Cells(1, 1), .Cells(nCurr + 1, nCols)).AutoFilter
        .Columns("A:" & ColLetter(nCols)).ColumnWidth = 11
        .Columns("A").ColumnWidth = 6
        .Columns("E").ColumnWidth = 16
        .Columns("F").ColumnWidth = 18
        .Activate
        .Range("G2").Select
        ActiveWindow.FreezePanes = True
    End With

    ' ---- 抽出_要対応シート(書式ごとコピー) ----
    wsO.Rows(1).Copy wsE.Rows(1)
    Dim er As Long: er = 1
    Dim v As Variant
    For Each v In extRows
        er = er + 1
        wsO.Rows(CLng(v)).Copy wsE.Rows(er)
    Next
    Application.CutCopyMode = False
    With wsE
        .Columns("A:" & ColLetter(nCols)).ColumnWidth = 11
        .Columns("A").ColumnWidth = 6
        .Columns("E").ColumnWidth = 16
        .Columns("F").ColumnWidth = 18
        If er > 1 Then .Range(.Cells(1, 1), .Cells(er, nCols)).AutoFilter
        .Activate
        .Range("G2").Select
        ActiveWindow.FreezePanes = True
    End With

    ' ---- チェックシート ----
    Dim cr As Long: cr = 1
    Dim warnClr As Long, okClr As Long
    warnClr = clrRed: okClr = RGB(198, 239, 206)
    With wsK
        .Columns("A").ColumnWidth = 42
        .Columns("B").ColumnWidth = 70
        .Cells(cr, 1).Value = "■ 作成結果チェック": .Cells(cr, 1).Font.Bold = True: cr = cr + 2
        cr = PutChk(wsK, cr, "作成日時", Format$(Now, "yyyy/mm/dd hh:nn:ss"), 0, 0)
        cr = PutChk(wsK, cr, "前回シート 実績番号", Join(DictKeys(jisPrev), ", "), IIf(jisPrev.Count <> 1, warnClr, 0), 0)
        cr = PutChk(wsK, cr, "今回シート 実績番号", Join(DictKeys(jisCurr), ", "), IIf(jisCurr.Count <> 1, warnClr, 0), 0)
        If jisPrev.Count >= 1 And jisCurr.Count >= 1 Then
            If Val(DictKeys(jisPrev)(0)) >= Val(DictKeys(jisCurr)(0)) Then
                cr = PutChk(wsK, cr, "実績番号の新旧", "前回の実績番号が今回以上です。シートが逆の可能性!", warnClr, 0)
            Else
                cr = PutChk(wsK, cr, "実績番号の新旧", "正常(今回の方が新しい)", 0, okClr)
            End If
        End If
        cr = PutChk(wsK, cr, "前回シート データ行数", nPrev, 0, 0)
        cr = PutChk(wsK, cr, "今回シート データ行数", nCurr, 0, 0)
        cr = PutChk(wsK, cr, "比較シート 行数(今回と同数のはず)", nCurr, 0, okClr)
        cr = PutChk(wsK, cr, "前回シートのキー重複", dupPrev, IIf(dupPrev > 0, warnClr, 0), IIf(dupPrev = 0, okClr, 0))
        cr = PutChk(wsK, cr, "今回シートのキー重複", dupCurr, IIf(dupCurr > 0, warnClr, 0), IIf(dupCurr = 0, okClr, 0))
        cr = PutChk(wsK, cr, "今回にあって前回にない箇所", unmatchedCurr.Count, IIf(unmatchedCurr.Count > 0, warnClr, 0), IIf(unmatchedCurr.Count = 0, okClr, 0))
        If unmatchedCurr.Count > 0 Then cr = PutChk(wsK, cr, "  └ 該当キロ程(先頭50件)", JoinCol(unmatchedCurr, 50), 0, 0)
        cr = PutChk(wsK, cr, "前回にあって今回にない箇所", unmatchedPrev.Count, IIf(unmatchedPrev.Count > 0, warnClr, 0), IIf(unmatchedPrev.Count = 0, okClr, 0))
        If unmatchedPrev.Count > 0 Then cr = PutChk(wsK, cr, "  └ 該当キロ程(先頭50件)", JoinCol(unmatchedPrev, 50), 0, 0)
        cr = PutChk(wsK, cr, "数値にできなかった測定値セル", nonNum, IIf(nonNum > 0, warnClr, 0), IIf(nonNum = 0, okClr, 0))
        cr = cr + 1
        .Cells(cr, 1).Value = "■ 検出件数": .Cells(cr, 1).Font.Bold = True: cr = cr + 1
        Dim s1 As String, s2 As String, s3 As String, extTotal As Long
        For i = 1 To N_ITEMS
            If flagCnt(i) > 0 Then s1 = s1 & IIf(s1 = "", "", "、") & itemName(i) & ":" & flagCnt(i) & "件"
            If lowCnt(i) > 0 Then s2 = s2 & IIf(s2 = "", "", "、") & itemName(i) & ":" & lowCnt(i) & "件"
            If highCnt(i) > 0 Then s3 = s3 & IIf(s3 = "", "", "、") & itemName(i) & ":" & highCnt(i) & "件"
        Next
        cr = PutChk(wsK, cr, "今回 目標値F=1 の箇所(項目別)", IIf(s1 = "", "なし", s1), 0, 0)
        cr = PutChk(wsK, cr, "うち 変位進み量 10mm以上15mm未満", IIf(s2 = "", "0件", s2), 0, 0)
        cr = PutChk(wsK, cr, "うち 変位進み量 15mm以上", IIf(s3 = "", "0件", s3), 0, 0)
        cr = PutChk(wsK, cr, "抽出_要対応シートの行数", extRows.Count, 0, 0)
    End With

    wsO.Activate
    wsO.Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "比較シートを作成しました。(" & Format$(Timer - t0, "0.0") & "秒)" & vbCrLf & _
           "要対応(10mm以上): " & extRows.Count & "件" & vbCrLf & _
           "※「チェック」シートで警告がないか確認してください。", vbInformation
    Exit Sub

FAIL:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Exit Sub
FIN:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'---- 補助関数 ----

Private Function HeaderMap(d As Variant) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long, h As String
    For c = 1 To UBound(d, 2)
        h = Trim$(CStr(d(1, c)))
        If h <> "" And Not m.Exists(h) Then m(h) = c
    Next
    Set HeaderMap = m
End Function

Private Function MakeKey(d As Variant, r As Long, h As Object, keyHdr() As String) As String
    Dim i As Long, s As String, v As Variant
    For i = 1 To 5
        v = ToNum(d(r, h(keyHdr(i))))
        s = s & "|" & CStr(v)
    Next
    MakeKey = s
End Function

' 数値化できれば Double、空なら Empty、それ以外は元の値のまま返す
Private Function ToNum(v As Variant) As Variant
    If IsEmpty(v) Or IsError(v) Then
        ToNum = Empty
    ElseIf VarType(v) = vbString Then
        If Trim$(v) = "" Then
            ToNum = Empty
        ElseIf IsNumeric(v) Then
            ToNum = CDbl(v)
        Else
            ToNum = v
        End If
    Else
        ToNum = v
    End If
End Function

Private Function IsNumeric2(v As Variant) As Boolean
    IsNumeric2 = False
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) Then IsNumeric2 = True
End Function

Private Sub AddFill(fillR() As Long, fillC() As Long, fillClr() As Long, _
                    nFill As Long, capFill As Long, r As Long, c As Long, clr As Long)
    nFill = nFill + 1
    If nFill > capFill Then
        capFill = capFill * 2
        ReDim Preserve fillR(1 To capFill)
        ReDim Preserve fillC(1 To capFill)
        ReDim Preserve fillClr(1 To capFill)
    End If
    fillR(nFill) = r: fillC(nFill) = c: fillClr(nFill) = clr
End Sub

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub DeleteSheetIfExists(nm As String)
    If SheetExists(nm) Then ThisWorkbook.Worksheets(nm).Delete
End Sub

Private Function PutChk(ws As Worksheet, r As Long, label As String, v As Variant, _
                        warnClr As Long, okClr As Long) As Long
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = v
    If warnClr <> 0 Then
        ws.Cells(r, 2).Interior.Color = warnClr
        ws.Cells(r, 2).Font.Bold = True
    ElseIf okClr <> 0 Then
        ws.Cells(r, 2).Interior.Color = okClr
    End If
    PutChk = r + 1
End Function

Private Function DictKeys(d As Object) As Variant
    Dim a() As String, i As Long, k As Variant
    ReDim a(0 To d.Count - 1)
    i = 0
    For Each k In d.Keys
        a(i) = CStr(k): i = i + 1
    Next
    DictKeys = a
End Function

Private Function JoinCol(col As Collection, maxN As Long) As String
    Dim s As String, i As Long
    For i = 1 To col.Count
        If i > maxN Then s = s & " …ほか": Exit For
        s = s & IIf(s = "", "", ", ") & CStr(col(i))
    Next
    JoinCol = s
End Function

Private Function ColLetter(c As Long) As String
    ColLetter = Split(Cells(1, c).Address(True, False), "$")(0)
End Function
