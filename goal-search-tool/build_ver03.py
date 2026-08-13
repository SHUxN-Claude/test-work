# -*- coding: utf-8 -*-
"""
目標値・基準値検索ツール ver.03 ビルド
修正版2 (見出し行自動検出 + 条件付き書式修正済み) をベースに:
  1. 「基準値検索結果まとめ」シート(sheet5) と 計算用「作業用2」シート(sheet6, 非表示) を追加
     - 基準値の検出ルール: 見出しが「基準」と「F/Ｆ」を含む。ただし「確定」を含む列は除外し、
       例外として「正矢量確定」を含む列は対象とする。
       (確定基準値Fは基準値Fのミラーのため対象外。正矢量には素の基準値F列が無いため
        正矢量確定基準値Fを対象にする — 利用者の指定)
     - フラグは目標値と同じく 1 / "1" で該当
  2. 見出し行自動検出(作業用!G10:G19)を「目標」だけでなく「基準」でも反応するよう拡張
  3. 「はじめに」の説明文とタイトルを ver.03 向けに更新、基準値まとめの状態表示も追加
  4. 元データを空にして出荷 (貼り付け前のクリーンな状態)
"""
import zipfile, re, sys

SRC = '目標値検索ツール_修正版2_2026.07.16.xlsx'
DST = sys.argv[1] if len(sys.argv) > 1 else '目標値・基準値検索ツール_ver03.xlsx'
KEEP_DATA = '--keep-data' in sys.argv  # 検証用: 元データを残す

# ---- 基準値フラグ列の判定式 (作業用2 の3行目マスター式) ----
# 基準 かつ (F or Ｆ) かつ (確定を含まない or 正矢量確定を含む)
KIJUN_DETECT = ('IF(B2="",0,IF(ISNUMBER(SEARCH("基準",B2))'
                '*((ISNUMBER(SEARCH("F",B2))+ISNUMBER(SEARCH("Ｆ",B2)))&gt;0)'
                '*(((1-ISNUMBER(SEARCH("確定",B2)))+ISNUMBER(SEARCH("正矢量確定",B2)))&gt;0),1,0))')
MOKUHYO_DETECT = ('IF(B2="",0,IF(ISNUMBER(SEARCH("目標",B2))'
                  '*((ISNUMBER(SEARCH("F",B2))+ISNUMBER(SEARCH("Ｆ",B2)))&gt;0),1,0))')

def make_sheet6(s4):
    """作業用 → 作業用2: フラグ検出を基準値ルールに差し替え。
    3行目は共有数式が5グループ(マスター: B3, AH3, BN3, CT3, DZ3)に分かれているため、
    セル参照を後方参照で受けて5本すべて置換する。"""
    pat = re.compile(
        r'IF\((?P<c>[A-Z]{1,2})2="",0,IF\(ISNUMBER\(SEARCH\("目標",(?P=c)2\)\)'
        r'\*\(\(ISNUMBER\(SEARCH\("F",(?P=c)2\)\)\+ISNUMBER\(SEARCH\("Ｆ",(?P=c)2\)\)\)&gt;0\),1,0\)\)')
    def rep(m):
        c = m.group('c')
        return (f'IF({c}2="",0,IF(ISNUMBER(SEARCH("基準",{c}2))'
                f'*((ISNUMBER(SEARCH("F",{c}2))+ISNUMBER(SEARCH("Ｆ",{c}2)))&gt;0)'
                f'*(((1-ISNUMBER(SEARCH("確定",{c}2)))+ISNUMBER(SEARCH("正矢量確定",{c}2)))&gt;0),1,0))')
    s4, n = pat.subn(rep, s4)
    assert n == 5, f'検出マスター式の置換数が想定外: {n}'
    return s4

def make_sheet5(s3):
    """目標値検索結果まとめ → 基準値検索結果まとめ"""
    s = s3.replace('作業用!', '作業用2!')
    # 表示文言
    s = s.replace('目標値の項目', '基準値の項目')
    s = s.replace('オレンジ＝目標値F系の列見出し', 'オレンジ＝基準値F系の列見出し')
    s = s.replace('赤＝目標値フラグ「1」', '赤＝基準値フラグ「1」')
    s = s.replace('黄＝その目標値の数値', '黄＝その基準値の数値')
    s = s.replace('目標値F系の列が15個を超えています', '基準値F系の列が15個を超えています')
    # 条件付き書式: オレンジ(見出し)・赤(フラグ1) のルールを基準値ルールへ
    old_orange = ('AND(B4&lt;&gt;"",ISNUMBER(SEARCH("目標",B4)),'
                  'OR(ISNUMBER(SEARCH("F",B4)),ISNUMBER(SEARCH("Ｆ",B4))))')
    new_orange = ('AND(B4&lt;&gt;"",ISNUMBER(SEARCH("基準",B4)),'
                  'OR(ISNUMBER(SEARCH("F",B4)),ISNUMBER(SEARCH("Ｆ",B4))),'
                  'OR(NOT(ISNUMBER(SEARCH("確定",B4))),ISNUMBER(SEARCH("正矢量確定",B4))))')
    old_red = ('AND(ISNUMBER(SEARCH("目標",B$4)),'
               'OR(ISNUMBER(SEARCH("F",B$4)),ISNUMBER(SEARCH("Ｆ",B$4))),OR(B5=1,B5="1"))')
    new_red = ('AND(ISNUMBER(SEARCH("基準",B$4)),'
               'OR(ISNUMBER(SEARCH("F",B$4)),ISNUMBER(SEARCH("Ｆ",B$4))),'
               'OR(NOT(ISNUMBER(SEARCH("確定",B$4))),ISNUMBER(SEARCH("正矢量確定",B$4))),'
               'OR(B5=1,B5="1"))')
    assert s.count(old_orange) == 1 and s.count(old_red) == 1, '条件付き書式ルールが見つからない'
    s = s.replace(old_orange, new_orange).replace(old_red, new_red)
    return s

def fix_header_probe(s4):
    """見出し行検出(G10:G19)を「目標」に加えて「基準」でも反応させる"""
    pat = re.compile(
        r'IF\(COUNTIFS\(元データ!\$(\d+):\$\1,"\*目標\*",元データ!\$\1:\$\1,"\*F\*"\)'
        r'\+COUNTIFS\(元データ!\$\1:\$\1,"\*目標\*",元データ!\$\1:\$\1,"\*Ｆ\*"\)&gt;0,1,0\)')
    def rep(m):
        r = m.group(1)
        t = lambda kw, f: f'COUNTIFS(元データ!${r}:${r},"*{kw}*",元データ!${r}:${r},"*{f}*")'
        return ('IF(' + '+'.join([t('目標','F'), t('目標','Ｆ'), t('基準','F'), t('基準','Ｆ')])
                + '&gt;0,1,0)')
    s, n = pat.subn(rep, s4)
    assert n == 10, f'見出し検出ヘルパー置換 {n}'
    return s

def fix_workbook(wb):
    old = '<sheet name="目標値検索結果まとめ" sheetId="3" r:id="rId3"/>'
    assert wb.count(old) == 1
    wb = wb.replace(old, old + '<sheet name="基準値検索結果まとめ" sheetId="5" r:id="rId9"/>')
    old = '<sheet name="作業用" sheetId="4" state="hidden" r:id="rId4"/>'
    assert wb.count(old) == 1
    wb = wb.replace(old, old + '<sheet name="作業用2" sheetId="6" state="hidden" r:id="rId10"/>')
    return wb

def fix_rels(s):
    add = ('<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>'
           '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>')
    return s.replace('</Relationships>', add + '</Relationships>')

def fix_content_types(s):
    tpl = '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    add = tpl % 5 + tpl % 6
    return s.replace('</Types>', add + '</Types>')

def fix_sheet1(s):
    """はじめにシート: 基準値まとめの状態表示行を追加"""
    old = '</c></row></sheetData>'
    assert s.count(old) == 1
    add = '</c></row><row r="22" spans="2:3" ht="18"><c r="C22" s="9" t="str"><f>基準値検索結果まとめ!$A$1</f></c></row></sheetData>'
    return s.replace(old, add)

def fix_sst(s):
    rep = [
        ('目標値検索ツール【行丸ごと版】（マクロなし・自動計算）',
         '目標値・基準値検索ツール【行丸ごと版】ver.03（マクロなし・自動計算）'),
        ('目標値がある行を、元データから1行まるごと抜き出して表示します。',
         '目標値・基準値がある行を、元データからそれぞれのまとめシートに1行まるごと抜き出して表示します。'),
        ('②  「目標値検索結果まとめ」シートを開く',
         '②  「目標値検索結果まとめ」「基準値検索結果まとめ」シートを開く'),
        ('・オレンジ＝「目標」と「F」を含む列の見出し（例：高低（左）目標値F、軌間目標F）',
         '・オレンジ＝検索対象の列見出し。目標値まとめ＝「目標」と「F」を含む列／基準値まとめ＝「基準」と「F」を含む列（確定基準値Fを除く。ただし正矢量確定基準値Fは対象）'),
        ('・赤＝目標値のフラグ「1」が入っているセル',
         '・赤＝フラグ「1」が入っているセル'),
        ('・黄＝その目標値の数値（フラグ列より左の直近の値列。例：高低（右）目標値F → 高低（右））',
         '・黄＝その数値（フラグ列より左の直近の値列。例：高低（右）目標値F → 高低（右）、通り（左）正矢量確定基準値F → 通り（左）実測値）'),
        ('・対応上限：データ 40,000行 × 150列、目標値F系の列 15個、表示できる該当行 1,000行。',
         '・対応上限：データ 40,000行 × 150列、目標値F系・基準値F系の列 各15個、表示できる該当行 各1,000行。'),
        ('・「作業用」シート（非表示）と結果シートの非表示行（3行目）は計算用です。削除しないでください。',
         '・「作業用」「作業用2」シート（非表示）と結果シートの非表示行（3行目）は計算用です。削除しないでください。'),
    ]
    for old, new in rep:
        assert s.count(old) == 1, old[:30]
        s = s.replace(old, new)
    return s

def empty_sheet2(s):
    i = s.find('<sheetData>'); j = s.find('</sheetData>')
    assert i >= 0 and j > i
    s = s[:i] + '<sheetData/>' + s[j+len('</sheetData>'):]
    s = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1"/>', s)
    return s

zin = zipfile.ZipFile(SRC)
s3 = zin.read('xl/worksheets/sheet3.xml').decode('utf-8')
s4 = fix_header_probe(zin.read('xl/worksheets/sheet4.xml').decode('utf-8'))
sheet5 = make_sheet5(s3)
sheet6 = make_sheet6(s4)

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == 'xl/workbook.xml':
            data = fix_workbook(data.decode()).encode()
        elif it.filename == 'xl/_rels/workbook.xml.rels':
            data = fix_rels(data.decode()).encode()
        elif it.filename == '[Content_Types].xml':
            data = fix_content_types(data.decode()).encode()
        elif it.filename == 'xl/worksheets/sheet1.xml':
            data = fix_sheet1(data.decode()).encode()
        elif it.filename == 'xl/worksheets/sheet2.xml' and not KEEP_DATA:
            data = empty_sheet2(data.decode()).encode()
        elif it.filename == 'xl/worksheets/sheet4.xml':
            data = s4.encode()
        elif it.filename == 'xl/sharedStrings.xml':
            data = fix_sst(data.decode()).encode()
        zo.writestr(it, data)
    zo.writestr('xl/worksheets/sheet5.xml', sheet5.encode())
    zo.writestr('xl/worksheets/sheet6.xml', sheet6.encode())
print('written:', DST)
