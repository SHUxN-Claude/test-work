# -*- coding: utf-8 -*-
"""
目標値検索ツール【行丸ごと版】の見出し行自動検出対応
- 元データの1行目が「◯件のデータが検索されました」等の帯行でも動くよう、
  見出し行を先頭10行から自動検出する方式に変更する。
変更点:
  1. 作業用!G10:G19 … 元データ1〜10行目それぞれに「目標」+「F/Ｆ」を含むセルがあるか(0/1)
  2. 作業用!A2      … 見出し行番号 = 最初に1が立つ行 (見つからなければ1)
  3. 作業用!B2:EU2  … 見出しミラーを 元データ!X$1 固定から INDEX(元データ!X:X,$A$2) に変更
  4. 結果まとめ!B4:EU4 … 同上 (作業用!$A$2 を参照)
  5. 結果まとめ!A1/A2 … 空判定・150列超過判定を1行目固定から先頭10行に拡張
  6. workbook.xml   … fullCalcOnLoad="1" (開いたとき全再計算) / calcChain 削除
  7. はじめに(共有文字列) … 使い方①に帯行OKの注記を追記
"""
import zipfile, re, shutil, sys

SRC = '/root/.claude/uploads/09725893-9ee0-5a14-a959-b24990f83ec4/837b4a0a-_________________2026.07.10_________.xlsx'
DST = sys.argv[1] if len(sys.argv) > 1 else '目標値検索ツール_修正版.xlsx'

MIRROR = re.compile(r'IF\(元データ!([A-Z]{1,2})\$1="","",元データ!\1\$1\)')

def fix_sheet4(s):
    # 3. 見出しミラーを可変行参照へ
    s, n = MIRROR.subn(r'IF(INDEX(元データ!\1:\1,$A$2)="","",INDEX(元データ!\1:\1,$A$2))', s)
    assert n == 150, f'sheet4 mirror replaced {n}'
    # 2. A2: 見出し行番号
    a2 = '<c r="A2"><f>IFERROR(MATCH(1,$G$10:$G$19,0),1)</f></c>'
    old = '<row r="2" spans="1:151">'
    assert s.count(old) == 1
    s = s.replace(old, old + a2)
    # 1. G10:G19 見出し行検出ヘルパー (元データ 1〜10行目)
    for i in range(10):
        row_tag = f'<row r="{10+i}" spans="1:151">'
        p = s.find(row_tag)
        assert p >= 0, row_tag
        q = s.find('</row>', p)
        r = i + 1  # 元データの行番号
        g = (f'<c r="G{10+i}"><f>IF(COUNTIFS(元データ!${r}:${r},"*目標*",元データ!${r}:${r},"*F*")'
             f'+COUNTIFS(元データ!${r}:${r},"*目標*",元データ!${r}:${r},"*Ｆ*")&gt;0,1,0)</f></c>')
        s = s[:q] + g + s[q:]
    return s

def fix_sheet3(s):
    # 4. 見出し表示行を可変行参照へ
    s, n = MIRROR.subn(r'IF(INDEX(元データ!\1:\1,作業用!$A$2)="","",INDEX(元データ!\1:\1,作業用!$A$2))', s)
    assert n == 150, f'sheet3 mirror replaced {n}'
    # 5. 空判定・列超過判定を先頭10行で見る
    old1 = 'COUNTA(元データ!$A$1:$ET$1)=0'
    old2 = 'COUNTA(元データ!$EU$1:$XFD$1)&gt;0'
    assert s.count(old1) == 1 and s.count(old2) == 1
    s = s.replace(old1, 'COUNTA(元データ!$A$1:$ET$10)=0')
    s = s.replace(old2, 'COUNTA(元データ!$EU$1:$XFD$10)&gt;0')
    return s

def fix_workbook(s):
    old = '<calcPr calcId="191029"/>'
    assert s.count(old) == 1
    return s.replace(old, '<calcPr calcId="191029" fullCalcOnLoad="1"/>')

def fix_rels(s):
    return re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', s)

def fix_content_types(s):
    return re.sub(r'<Override [^>]*calcChain[^>]*/>', '', s)

def fix_sst(s):
    old = '①  CSVを開いて中身を全部コピーし、「元データ」シートのA1セルに貼り付ける（見出しの行ごと）'
    assert s.count(old) == 1
    new = old + '　※見出しの前に「○件のデータが検索されました」等の行があってもOK（先頭10行以内なら見出し行を自動検出します）'
    return s.replace(old, new)

FIXERS = {
    'xl/worksheets/sheet4.xml': fix_sheet4,
    'xl/worksheets/sheet3.xml': fix_sheet3,
    'xl/workbook.xml': fix_workbook,
    'xl/_rels/workbook.xml.rels': fix_rels,
    '[Content_Types].xml': fix_content_types,
    'xl/sharedStrings.xml': fix_sst,
}

zin = zipfile.ZipFile(SRC)
with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
    for item in zin.infolist():
        if item.filename == 'xl/calcChain.xml':
            continue  # 6. 数式を変えたので削除 (Excelが自動再構築)
        data = zin.read(item.filename)
        if item.filename in FIXERS:
            data = FIXERS[item.filename](data.decode('utf-8')).encode('utf-8')
        zout.writestr(item, data)
print('written:', DST)
