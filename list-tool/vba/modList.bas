Attribute VB_Name = "modList"
Option Explicit

'==============================================================
' 列車動揺 整備基準値超過箇所一覧 自動入力ツール
'
' 「下り/上りの動揺データ抽出ツール(.xlsm)」の①基準値超過を読み込み、
' 当月シートへ転記する。駅名・構造物はこのブック内のマスタから判定する。
'
' ボタン:
'   抽出ツールから取込 / 番号・日数を再計算 / このシートをクリア
'   新しい月のシートを作る
'
' 方針:
'   ・セルに数式は入れない(値のみ)。他ブックへコピペしても壊れない
'   ・セル結合もしない。手で行を足し引きしても崩れない
'==============================================================

'--- マスタシート名 ---
Private Const M_TUNNEL As String = "トンネル"
Private Const M_BRIDGE As String = "橋りょう"
Private Const M_CROSS  As String = "踏切"
Private Const M_TURN   As String = "分岐器"
Private Const M_CULV   As String = "開きょ"
Private Const M_RAIL   As String = "レール"
Private Const M_STA    As String = "駅境界"
Private Const SH_TPL   As String = "ひな形（コピー元）"

Private Const LINE_NAME As String = "中央本線"

'--- 一覧シートの列 ---
Private Const C_NO     As Long = 1    ' A 箇所
Private Const C_SITE   As Long = 2    ' B 現場
Private Const C_LINE   As Long = 3    ' C 線名
Private Const C_DIR    As Long = 4    ' D 線別
Private Const C_ST1    As Long = 5    ' E 駅名等(から)
Private Const C_ST2    As Long = 6    ' F 　　　(まで)
Private Const C_TRAMS  As Long = 7    ' G トラムス判定(手入力・空のまま)
Private Const C_KM     As Long = 8    ' H キロ程(車上帳票)
Private Const C_KIND   As Long = 9    ' I 検査種別
Private Const C_METHOD As Long = 10   ' J 検査方法
Private Const C_UD     As Long = 11   ' K 車上帳票 上下動
Private Const C_LR     As Long = 12   ' L 　　　　 左右動
Private Const C_TUD    As Long = 13   ' M TRAMS 上下動(手入力)
Private Const C_TLR    As Long = 14   ' N 　　　 左右動(手入力)
Private Const C_STRUCT As Long = 15   ' O 構造物
Private Const C_DEAL   As Long = 16   ' P 取り扱い(手入力)
Private Const C_STDNO  As Long = 17   ' Q 基準値No(手入力)
Private Const C_PLAN   As Long = 18   ' R 施工予定日(手入力)
Private Const C_INSP   As Long = 19   ' S 検査実施日
Private Const C_TYPE   As Long = 20   ' T 整備期限種別
Private Const C_LIMIT  As Long = 21   ' U 整備実施期限
Private Const C_DONE   As Long = 22   ' V 整備実施日(手入力)
Private Const C_DAYN   As Long = 23   ' W 昼・夜(手入力)
Private Const C_CHK    As Long = 24   ' X 確認者(手入力)
Private Const C_DAYS   As Long = 25   ' Y 日数
Private Const C_JUDGE  As Long = 26   ' Z 整備期限判定
Private Const C_FIXED  As Long = 27   ' AA 確定基準値(手入力・空のまま)
Private Const C_LAST   As Long = 27

Private Const R_DATA   As Long = 5    ' データ開始行
Private Const FONT_NAME As String = "HG丸ｺﾞｼｯｸM-PRO"

'--- 設定(既定値。必要ならここを書き換える) ---
Private Const TOL_STRUCT As Double = 25#   ' 構造物の許容範囲(m)
Private Const TOL_SITE   As Double = 20#   ' 同一現場とみなす距離(m)

'--- 取込行の色 ---
Private Const CLR_NEW  As Long = &HF5E8D8    ' うすい水色(新規)
Private Const CLR_DUP  As Long = &HD0D0FF    ' うすい赤(重複の可能性)

'--- マスタのキャッシュ ---
Private Type TRange
    dir As String
    fr As Double
    to_ As Double
    nm As String
    eki As String
    kind As String
End Type
Private Type TPoint
    dir As String
    km As Double
    lab As String
End Type
Private Type TStation
    dir As String
    fr As Double
    to_ As Double
    nm As String
End Type

Private mRng() As TRange
Private mNRng As Long
Private mPnt() As TPoint
Private mNPnt As Long
Private mSta() As TStation
Private mNSta As Long
Private mLoaded As Boolean


'==============================================================
' 初期設定(ボタン設置 + ひな形作成)
'==============================================================
Public Sub 初期設定()
    If LCase(Right(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        If MsgBox("このブックがマクロ有効ブック(.xlsm)として保存されていません。" & vbCrLf & _
                  "このまま保存するとマクロが消えてしまいます。" & vbCrLf & vbCrLf & _
                  "F12(名前を付けて保存)で「Excel マクロ有効ブック (*.xlsm)」に" & _
                  "してから保存してください。" & vbCrLf & vbCrLf & _
                  "このまま続けますか?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    Dim ws As Worksheet, cnt As Long
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Or ws.Name = SH_TPL Then
            PutButtons ws
            cnt = cnt + 1
        End If
    Next

    If Not SheetExists(SH_TPL) Then MakeTemplate

    MsgBox cnt & " 枚の月シートにボタンを設置しました。" & vbCrLf & _
           "「" & SH_TPL & "」シートも用意しました。" & vbCrLf & vbCrLf & _
           "新しい月は「新しい月のシートを作る」ボタンから作成してください。", vbInformation
End Sub

Private Sub PutButtons(ws As Worksheet)
    Dim b As Object
    For Each b In ws.Buttons
        b.Delete
    Next
    Dim l As Single, t As Single
    l = ws.Cells(1, 33).Left        ' AG列のあたり(AC～AEは選択肢の元データなので避ける)
    t = ws.Cells(1, 1).Top
    AddBtn ws, l, t, 170, 26, "抽出ツールから取込", "一覧に取込"
    AddBtn ws, l, t + 28, 170, 26, "番号・日数を再計算", "番号日数を再計算"
    AddBtn ws, l, t + 56, 170, 26, "このシートをクリア", "シートをクリア"
    AddBtn ws, l, t + 84, 170, 26, "新しい月のシートを作る", "新しい月のシートを作る"
End Sub

Private Sub AddBtn(ws As Worksheet, l As Single, t As Single, w As Single, _
                   h As Single, cap As String, macroName As String)
    Dim b As Button
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = macroName
    b.Font.Size = 10
End Sub

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

'--- シート名が「202605（8月）」の形かどうか ---
Private Function IsMonthSheet(nm As String) As Boolean
    If Len(nm) < 8 Then Exit Function
    If Not IsNumeric(Left(nm, 6)) Then Exit Function
    IsMonthSheet = (InStr(nm, "月") > 0)
End Function


'==============================================================
' ひな形 / 新しい月のシート
'==============================================================
Private Sub MakeTemplate()
    Dim src As Worksheet, ws As Worksheet, latest As String
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then
            If ws.Name > latest Then
                latest = ws.Name
                Set src = ws
            End If
        End If
    Next
    If src Is Nothing Then
        MsgBox "月シートが見つからないため、ひな形を作れませんでした。", vbExclamation
        Exit Sub
    End If
    src.Copy After:=ThisWorkbook.Worksheets(src.Index)
    Set ws = ThisWorkbook.Worksheets(src.Index + 1)
    ws.Name = SH_TPL
    ClearSheet ws, True
    PutButtons ws
End Sub

Public Sub 新しい月のシートを作る()
    If Not SheetExists(SH_TPL) Then MakeTemplate
    If Not SheetExists(SH_TPL) Then Exit Sub

    Dim yr As String, mo As String
    yr = InputBox("年度を入力してください(4月～翌3月が1年度)" & vbCrLf & _
                  "例: 2027", "新しい月のシート", CStr(GuessYear()))
    If yr = "" Then Exit Sub
    If Not IsNumeric(yr) Then
        MsgBox "年度は数字で入力してください。", vbExclamation
        Exit Sub
    End If
    mo = InputBox("暦月を入力してください(1～12)" & vbCrLf & _
                  "例: 4 → 202701（4月）", "新しい月のシート", "4")
    If mo = "" Then Exit Sub
    If Not IsNumeric(mo) Then Exit Sub
    Dim m As Long
    m = CLng(mo)
    If m < 1 Or m > 12 Then
        MsgBox "1～12で入力してください。", vbExclamation
        Exit Sub
    End If

    Dim seq As Long, nm As String
    seq = ((m - 4 + 12) Mod 12) + 1          ' 4月=1 … 3月=12
    nm = yr & Format(seq, "00") & "（" & m & "月）"
    If SheetExists(nm) Then
        MsgBox "「" & nm & "」は既にあります。", vbExclamation
        Exit Sub
    End If

    Dim tpl As Worksheet, ws As Worksheet
    Set tpl = ThisWorkbook.Worksheets(SH_TPL)
    tpl.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set ws = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ws.Name = nm
    ws.Cells(1, 1).Value = yr & "年度　列車動揺　整備基準値超過箇所"
    ClearSheet ws, True
    PutButtons ws
    ws.Activate
    MsgBox "「" & nm & "」を作成しました。", vbInformation
End Sub

Private Function GuessYear() As Long
    Dim ws As Worksheet, latest As String
    For Each ws In ThisWorkbook.Worksheets
        If IsMonthSheet(ws.Name) Then
            If ws.Name > latest Then latest = ws.Name
        End If
    Next
    If latest = "" Then
        GuessYear = Year(Date)
    Else
        GuessYear = CLng(Left(latest, 4)) + 1
    End If
End Function


'==============================================================
' クリア
'==============================================================
Public Sub シートをクリア()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not IsMonthSheet(ws.Name) And ws.Name <> SH_TPL Then
        MsgBox "月シートを開いた状態で実行してください。", vbExclamation
        Exit Sub
    End If
    If MsgBox("「" & ws.Name & "」のデータをすべて消します。" & vbCrLf & _
              "手入力した列も消えます。よろしいですか?", _
              vbYesNo + vbExclamation) = vbNo Then Exit Sub
    ClearSheet ws, True
    MsgBox "クリアしました。", vbInformation
End Sub

Private Sub ClearSheet(ws As Worksheet, alsoHeader As Boolean)
    Dim lastR As Long
    lastR = LastDataRow(ws)
    If lastR >= R_DATA Then
        With ws.Range(ws.Cells(R_DATA, 1), ws.Cells(lastR, C_LAST))
            .ClearContents
            .Interior.Pattern = xlNone
        End With
    End If
    If alsoHeader Then
        ws.Cells(1, 11).ClearContents          ' K1 走行日
        ws.Cells(1, 17).ClearContents          ' Q1 施工期限
    End If
End Sub

Private Function LastDataRow(ws As Worksheet) As Long
    Dim r As Long
    For r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 To R_DATA Step -1
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))) > 0 Then
            LastDataRow = r
            Exit Function
        End If
    Next
    LastDataRow = R_DATA - 1
End Function


'==============================================================
' 取込(メイン)
'==============================================================
Public Sub 一覧に取込()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not IsMonthSheet(ws.Name) Then
        MsgBox "取り込む月のシートを開いた状態で実行してください。" & vbCrLf & _
               "(例:「202606（9月）」)", vbExclamation
        Exit Sub
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "動揺データ抽出ツール(下り・上り)を選んでください(まとめて選択可)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excelブック", "*.xlsm; *.xlsx"
        If .Show = 0 Then Exit Sub
    End With

    LoadMasters

    ' --- 各ファイルを読む ---
    Dim km() As Double, ud() As Double, lr() As Double
    Dim dir_() As String, mth() As String, insp() As Variant
    Dim n As Long
    ReDim km(1 To 500): ReDim ud(1 To 500): ReDim lr(1 To 500)
    ReDim dir_(1 To 500): ReDim mth(1 To 500): ReDim insp(1 To 500)

    Dim i As Long
    For i = 1 To fd.SelectedItems.Count
        If Not ReadTool(CStr(fd.SelectedItems(i)), km, ud, lr, dir_, mth, insp, n) Then
            Exit Sub          ' 中止
        End If
    Next
    If n = 0 Then
        MsgBox "基準値超過の行が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' --- 並べ替え: 下り(昇順) → 上り(昇順) ---
    Dim idx() As Long
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next
    SortRecords idx, dir_, km, n

    ' --- 書き出し(既存行の下に追記) ---
    Application.ScreenUpdating = False
    Dim startR As Long, r As Long, j As Long, newCnt As Long, dupCnt As Long
    startR = LastDataRow(ws) + 1
    If startR < R_DATA Then startR = R_DATA

    For i = 1 To n
        j = idx(i)
        r = startR + i - 1
        WriteRow ws, r, dir_(j), km(j), ud(j), lr(j), mth(j), insp(j)
        Dim dup As Boolean
        dup = IsDuplicate(ws, startR, dir_(j), km(j))
        ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST)).Interior.Color = _
            IIf(dup, CLR_DUP, CLR_NEW)
        If dup Then dupCnt = dupCnt + 1
        newCnt = newCnt + 1
    Next

    ' --- ヘッダー(走行日・施工期限) ---
    If Not IsEmpty(insp(idx(1))) Then ws.Cells(1, 11).Value = insp(idx(1))
    Recalc ws, False
    ws.Cells(1, 17).Value = ws.Cells(R_DATA, C_LIMIT).Value

    Application.ScreenUpdating = True
    Dim msg As String
    msg = newCnt & " 行を取り込みました(うすい水色の行)。"
    If dupCnt > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "そのうち " & dupCnt & " 行は、既にある行と同じ位置でした(うすい赤)。" & vbCrLf & _
              "どちらを残すか見比べて、不要な行を削除してください。" & vbCrLf & _
              "削除したあとに「番号・日数を再計算」を押すと番号が整います。"
    Else
        msg = msg & vbCrLf & vbCrLf & "色は「番号・日数を再計算」で消えます。"
    End If
    MsgBox msg, vbInformation
End Sub

'--- 抽出ツール1ファイルを読む ---
Private Function ReadTool(path As String, km() As Double, ud() As Double, _
                          lr() As Double, dir_() As String, mth() As String, _
                          insp() As Variant, n As Long) As Boolean
    Dim wb As Workbook, ws As Worksheet
    On Error GoTo fail
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    On Error Resume Next
    Set ws = wb.Worksheets("①基準値超過")
    On Error GoTo fail
    If ws Is Nothing Then
        MsgBox "「①基準値超過」シートが見つかりません:" & vbCrLf & path, vbExclamation
        wb.Close False
        ReadTool = True                      ' このファイルは飛ばして続行
        Exit Function
    End If

    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If lastR < R_DATA Then
        wb.Close False
        ReadTool = True
        Exit Function
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(R_DATA, 1), ws.Cells(lastR, 12)).Value

    ' --- 検査方法の判定 + 確認ダイアログ(Wチェック) ---
    Dim fname As String, guess As String, ans As VbMsgBoxResult
    fname = Dir(path)
    If Len(CStr(arr(1, 12))) > 0 Then fname = CStr(arr(1, 12))
    If InStr(fname, "代替検査") > 0 Or InStr(fname, "レールナビ") > 0 Then
        guess = "可搬式"
    Else
        guess = "常設式"
    End If
    ans = MsgBox("【" & Dir(path) & "】" & vbCrLf & vbCrLf & _
                 "線別：" & CStr(arr(1, 10)) & "　測定日：" & CStr(arr(1, 9)) & vbCrLf & _
                 "元PDF：" & fname & vbCrLf & vbCrLf & _
                 "検査方法を「" & guess & "」と判定しました。これでよろしいですか?" & vbCrLf & _
                 "　[はい]　　" & guess & " のまま取り込む" & vbCrLf & _
                 "　[いいえ]　" & IIf(guess = "可搬式", "常設式", "可搬式") & " に変えて取り込む" & vbCrLf & _
                 "　[キャンセル] 取込を中止", _
                 vbYesNoCancel + vbQuestion, "検査方法の確認")
    If ans = vbCancel Then
        wb.Close False
        ReadTool = False
        Exit Function
    End If
    If ans = vbNo Then guess = IIf(guess = "可搬式", "常設式", "可搬式")

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If IsNumeric(arr(r, 4)) And Len(CStr(arr(r, 4))) > 0 Then
            n = n + 1
            km(n) = CDbl(arr(r, 4)) * 1000#
            ud(n) = IIf(IsNumeric(arr(r, 5)) And Len(CStr(arr(r, 5))) > 0, CDbl(arr(r, 5)), 0)
            lr(n) = IIf(IsNumeric(arr(r, 6)) And Len(CStr(arr(r, 6))) > 0, CDbl(arr(r, 6)), 0)
            dir_(n) = Left(Replace(CStr(arr(r, 10)), "線", ""), 1)   ' 下線→下 / 上線→上
            mth(n) = guess
            insp(n) = arr(r, 9)
        End If
    Next
    wb.Close False
    ReadTool = True
    Exit Function
fail:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    MsgBox "読み込めませんでした:" & vbCrLf & path, vbExclamation
    ReadTool = True
End Function

Private Sub SortRecords(idx() As Long, dir_() As String, km() As Double, n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 1 To n - 1
        For j = 1 To n - i
            If SortKey(dir_(idx(j)), km(idx(j))) > SortKey(dir_(idx(j + 1)), km(idx(j + 1))) Then
                t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
            End If
        Next
    Next
End Sub

Private Function SortKey(d As String, km As Double) As Double
    SortKey = IIf(d = "下", 0#, 10000000#) + km
End Function

'--- 1行書き出し ---
Private Sub WriteRow(ws As Worksheet, r As Long, d As String, km As Double, _
                     ud As Double, lr As Double, mth As String, insp As Variant)
    Dim sen As String
    sen = IIf(d = "下", "21", "22")

    ws.Cells(r, C_LINE).Value = LINE_NAME
    ws.Cells(r, C_DIR).Value = d

    Dim s1 As String, s2 As String
    LookupStation km, sen, s1, s2
    ws.Cells(r, C_ST1).Value = s1
    ws.Cells(r, C_ST2).Value = s2

    ws.Cells(r, C_KM).Value = km
    ws.Cells(r, C_KIND).Value = "列車動揺"
    ws.Cells(r, C_METHOD).Value = mth
    If ud > 0 Then ws.Cells(r, C_UD).Value = ud
    If lr > 0 Then ws.Cells(r, C_LR).Value = lr
    ws.Cells(r, C_STRUCT).Value = LookupStruct(km, sen)
    ws.Cells(r, C_INSP).Value = insp
    ws.Cells(r, C_TYPE).Value = 1

    ApplyFormat ws, r
End Sub

'--- 重複判定: 取込前からあった行(startR より上)にだけ同じ位置があるか。
'    今回まとめて取り込んだ行どうし(例 104891と104902)は重複にしない。
Private Function IsDuplicate(ws As Worksheet, startR As Long, d As String, km As Double) As Boolean
    Dim i As Long
    For i = R_DATA To startR - 1
        If CStr(ws.Cells(i, C_DIR).Value) = d Then
            If IsNumeric(ws.Cells(i, C_KM).Value) Then
                If Abs(CDbl(ws.Cells(i, C_KM).Value) - km) <= TOL_SITE Then
                    IsDuplicate = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function


'==============================================================
' 書式
'==============================================================
Private Sub ApplyFormat(ws As Worksheet, r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))
        .Font.Name = FONT_NAME
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
    End With
    ws.Cells(r, C_TRAMS).NumberFormatLocal = "0""k""000.0""m"""
    ws.Cells(r, C_KM).NumberFormatLocal = "0""k""000.0""m"""
    Dim c As Long
    For c = C_UD To C_TLR
        ws.Cells(r, c).NumberFormatLocal = "0.00_ "
        ws.Cells(r, c).Font.Size = 11
    Next
    ws.Cells(r, C_STRUCT).Font.Size = 11
    ws.Cells(r, C_INSP).NumberFormatLocal = "yy/mm/dd"
    ws.Cells(r, C_INSP).Font.Size = 11
    ws.Cells(r, C_LIMIT).NumberFormatLocal = "yy/mm/dd"
    ws.Cells(r, C_DONE).NumberFormatLocal = "yy/mm/dd"
End Sub


'==============================================================
' 番号・日数の再計算(数式を使わずに値で埋める)
'==============================================================
Public Sub 番号日数を再計算()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not IsMonthSheet(ws.Name) And ws.Name <> SH_TPL Then
        MsgBox "月シートを開いた状態で実行してください。", vbExclamation
        Exit Sub
    End If
    Recalc ws, True
    MsgBox "番号・現場・期限・日数を計算し直しました。", vbInformation
End Sub

Private Sub Recalc(ws As Worksheet, askSort As Boolean)
    Dim lastR As Long
    lastR = LastDataRow(ws)
    If lastR < R_DATA Then Exit Sub

    ' --- 並べ替え(下り昇順 → 上り昇順) ---
    If askSort Then
        If MsgBox("下り(キロ程昇順)→上り(キロ程昇順)の順に並べ替えますか?", _
                  vbYesNo + vbQuestion) = vbYes Then
            SortSheet ws, lastR
        End If
        ws.Range(ws.Cells(R_DATA, 1), ws.Cells(lastR, C_LAST)).Interior.Pattern = xlNone
    End If

    Dim r As Long, no As Long, site As Long
    Dim prevDir As String, prevKm As Double, started As Boolean
    For r = R_DATA To lastR
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))) = 0 Then
            ws.Cells(r, C_NO).ClearContents
            ws.Cells(r, C_SITE).ClearContents
        Else
            no = no + 1
            ws.Cells(r, C_NO).Value = no

            ' 現場番号: 同じ線別でキロ程が近いものは同じ現場(先頭行にだけ番号)
            Dim d As String, km As Double, isNew As Boolean
            d = CStr(ws.Cells(r, C_DIR).Value)
            km = 0
            If IsNumeric(ws.Cells(r, C_KM).Value) Then km = CDbl(ws.Cells(r, C_KM).Value)
            isNew = True
            If started Then
                If d = prevDir And Abs(km - prevKm) <= TOL_SITE Then isNew = False
            End If
            If isNew Then
                site = site + 1
                ws.Cells(r, C_SITE).Value = site
            Else
                ws.Cells(r, C_SITE).ClearContents
            End If
            prevDir = d
            prevKm = km
            started = True

            ' 整備実施期限 U = S + 14(種別1) / +29(種別2)
            Dim t As Variant, s As Variant
            t = ws.Cells(r, C_TYPE).Value
            s = ws.Cells(r, C_INSP).Value
            If IsDate(s) And IsNumeric(t) Then
                If CLng(t) = 1 Then
                    ws.Cells(r, C_LIMIT).Value = CDate(s) + 14
                ElseIf CLng(t) = 2 Then
                    ws.Cells(r, C_LIMIT).Value = CDate(s) + 29
                End If
                ws.Cells(r, C_LIMIT).NumberFormatLocal = "yy/mm/dd"
            End If

            ' 日数 Y = V - S + 1、判定 Z
            Dim v As Variant
            v = ws.Cells(r, C_DONE).Value
            If IsDate(s) And IsDate(v) Then
                Dim dy As Long
                dy = CLng(CDate(v) - CDate(s)) + 1
                ws.Cells(r, C_DAYS).Value = dy
                If IsNumeric(t) Then
                    If CLng(t) = 1 Then
                        ws.Cells(r, C_JUDGE).Value = IIf(dy > 15, "×", "OK")
                    ElseIf CLng(t) = 2 Then
                        ws.Cells(r, C_JUDGE).Value = IIf(dy > 30, "×", "OK")
                    End If
                End If
            Else
                ws.Cells(r, C_DAYS).ClearContents
                ws.Cells(r, C_JUDGE).ClearContents
            End If
        End If
    Next
    If lastR >= R_DATA Then ws.Cells(1, 17).Value = ws.Cells(R_DATA, C_LIMIT).Value
End Sub

Private Sub SortSheet(ws As Worksheet, lastR As Long)
    Dim n As Long
    n = lastR - R_DATA + 1
    If n < 2 Then Exit Sub
    Dim arr As Variant
    arr = ws.Range(ws.Cells(R_DATA, 1), ws.Cells(lastR, C_LAST)).Value

    Dim key() As Double, idx() As Long, i As Long, j As Long, t As Long
    ReDim key(1 To n): ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
        Dim d As String, km As Double
        d = CStr(arr(i, C_DIR))
        km = 0
        If IsNumeric(arr(i, C_KM)) Then km = CDbl(arr(i, C_KM))
        key(i) = SortKey(d, km)
        If Len(d) = 0 Then key(i) = 99999999#      ' 空行は最後
    Next
    For i = 1 To n - 1
        For j = 1 To n - i
            If key(idx(j)) > key(idx(j + 1)) Then
                t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
            End If
        Next
    Next

    Dim out() As Variant
    ReDim out(1 To n, 1 To C_LAST)
    Dim c As Long
    For i = 1 To n
        For c = 1 To C_LAST
            out(i, c) = arr(idx(i), c)
        Next
    Next
    ws.Range(ws.Cells(R_DATA, 1), ws.Cells(lastR, C_LAST)).Value = out
End Sub


'==============================================================
' マスタ読み込み
'==============================================================
Private Sub LoadMasters()
    If mLoaded Then Exit Sub
    ReDim mRng(1 To 3000)
    ReDim mPnt(1 To 3000)
    ReDim mSta(1 To 300)
    mNRng = 0: mNPnt = 0: mNSta = 0

    AddRangeMaster M_TUNNEL, 19, "T"
    AddRangeMaster M_BRIDGE, 22, "B"
    AddRangeMaster M_CROSS, 19, "X"
    AddRangeMaster M_TURN, 21, "P"
    AddRangeMaster M_CULV, 21, "C"
    AddRailMaster
    AddStationMaster
    mLoaded = True
End Sub

'--- から/まで を持つマスタ(P=16列, R=18列) ---
Private Sub AddRangeMaster(sheetName As String, nameCol As Long, kind As String)
    If Not SheetExists(sheetName) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 22)).Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, 6)) = LINE_NAME Then
            If IsNumeric(arr(r, 16)) And Len(CStr(arr(r, 16))) > 0 Then
                If mNRng < UBound(mRng) Then
                    mNRng = mNRng + 1
                    mRng(mNRng).dir = CStr(arr(r, 8))
                    mRng(mNRng).fr = CDbl(arr(r, 16))
                    mRng(mNRng).to_ = IIf(IsNumeric(arr(r, 18)) And Len(CStr(arr(r, 18))) > 0, _
                                          CDbl(arr(r, 18)), CDbl(arr(r, 16)))
                    mRng(mNRng).nm = CStr(arr(r, nameCol))
                    mRng(mNRng).eki = CStr(arr(r, 13))
                    mRng(mNRng).kind = kind
                End If
            End If
        End If
    Next
End Sub

'--- レール: AL(38)=伸縮継目→EJ / AR(44)=接着絶縁→IJ ---
Private Sub AddRailMaster()
    If Not SheetExists(M_RAIL) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M_RAIL)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim aF As Variant, aH As Variant, aAL As Variant, aAR As Variant
    aF = ws.Range(ws.Cells(2, 6), ws.Cells(lastR, 6)).Value
    aH = ws.Range(ws.Cells(2, 8), ws.Cells(lastR, 8)).Value
    aAL = ws.Range(ws.Cells(2, 38), ws.Cells(lastR, 38)).Value
    aAR = ws.Range(ws.Cells(2, 44), ws.Cells(lastR, 44)).Value

    Dim r As Long
    For r = 1 To UBound(aF, 1)
        If CStr(aF(r, 1)) = LINE_NAME Then
            If IsNumeric(aAL(r, 1)) And Len(CStr(aAL(r, 1))) > 0 Then
                AddPoint CStr(aH(r, 1)), CDbl(aAL(r, 1)), "EJ"
            End If
            If IsNumeric(aAR(r, 1)) And Len(CStr(aAR(r, 1))) > 0 Then
                AddPoint CStr(aH(r, 1)), CDbl(aAR(r, 1)), "IJ"
            End If
        End If
    Next
End Sub

Private Sub AddPoint(d As String, km As Double, lab As String)
    If mNPnt >= UBound(mPnt) Then Exit Sub
    mNPnt = mNPnt + 1
    mPnt(mNPnt).dir = d
    mPnt(mNPnt).km = km
    mPnt(mNPnt).lab = lab
End Sub

'--- 駅境界: C=3線名 D=4線別 F=6駅名 J=10から L=12まで ---
Private Sub AddStationMaster()
    If Not SheetExists(M_STA) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M_STA)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastR, 12)).Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, 3)) = LINE_NAME Then
            If IsNumeric(arr(r, 10)) And Len(CStr(arr(r, 10))) > 0 Then
                If mNSta < UBound(mSta) Then
                    mNSta = mNSta + 1
                    mSta(mNSta).dir = CStr(arr(r, 4))
                    mSta(mNSta).fr = CDbl(arr(r, 10))
                    mSta(mNSta).to_ = IIf(IsNumeric(arr(r, 12)) And Len(CStr(arr(r, 12))) > 0, _
                                          CDbl(arr(r, 12)), CDbl(arr(r, 10)))
                    mSta(mNSta).nm = CStr(arr(r, 6))
                End If
            End If
        End If
    Next
End Sub


'==============================================================
' 判定
'==============================================================
'--- 駅名: 構内なら (駅名, "構内")、駅間なら (手前の駅, 次の駅) ---
Private Sub LookupStation(km As Double, sen As String, ByRef s1 As String, ByRef s2 As String)
    Dim i As Long
    s1 = "": s2 = ""
    For i = 1 To mNSta
        If mSta(i).dir = sen Then
            If mSta(i).fr <= km And km <= mSta(i).to_ Then
                s1 = mSta(i).nm
                s2 = "構内"
                Exit Sub
            End If
        End If
    Next
    Dim bestPrev As Double, bestNext As Double
    bestPrev = -1E+15: bestNext = 1E+15
    For i = 1 To mNSta
        If mSta(i).dir = sen Then
            If mSta(i).to_ < km And mSta(i).to_ > bestPrev Then
                bestPrev = mSta(i).to_
                s1 = mSta(i).nm
            End If
            If mSta(i).fr > km And mSta(i).fr < bestNext Then
                bestNext = mSta(i).fr
                s2 = mSta(i).nm
            End If
        End If
    Next
End Sub

'--- 構造物: 前後TOL_STRUCTに入るものを「 / 」で連結 ---
Private Function LookupStruct(km As Double, sen As String) As String
    Dim out As String, i As Long, nm As String
    For i = 1 To mNRng
        If mRng(i).dir = sen And Len(mRng(i).nm) > 0 Then
            If (mRng(i).fr - TOL_STRUCT) <= km And km <= (mRng(i).to_ + TOL_STRUCT) Then
                nm = mRng(i).nm
                Select Case mRng(i).kind
                    Case "T"
                        If Right(nm, 1) = "Ｔ" Or Right(nm, 1) = "T" Then
                            nm = Left(nm, Len(nm) - 1) & "トンネル"
                        End If
                    Case "X"
                        If InStr(nm, "載線台") = 0 And InStr(nm, "踏切") = 0 _
                           And InStr(nm, "通路") = 0 Then nm = nm & "踏切"
                    Case "P"
                        nm = mRng(i).eki & ZenAlpha(nm) & "号"
                    Case "C"
                        If nm = "水路" Then nm = "下水きょ"
                End Select
                If InStr(out, nm) = 0 Then
                    If Len(out) > 0 Then out = out & " / "
                    out = out & nm
                End If
            End If
        End If
    Next
    For i = 1 To mNPnt
        If mPnt(i).dir = sen Then
            If Abs(mPnt(i).km - km) <= TOL_STRUCT Then
                If InStr(out, mPnt(i).lab) = 0 Then
                    If Len(out) > 0 Then out = out & " / "
                    out = out & mPnt(i).lab
                End If
            End If
        End If
    Next
    LookupStruct = out
End Function

'--- 英字だけ全角にする(数字は半角のまま): 12A → 12Ａ ---
Private Function ZenAlpha(s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            out = out & StrConv(ch, vbWide)
        Else
            out = out & ch
        End If
    Next
    ZenAlpha = out
End Function
