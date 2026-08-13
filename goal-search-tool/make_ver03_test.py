# -*- coding: utf-8 -*-
"""実CSV(基準値検索の新形式)をver.03の元データに注入したテストブックを作る"""
import csv, io, zipfile

NSW = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def esc(s): return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def colref(n):
    s=''
    while n>0: n,r = divmod(n-1,26); s = chr(65+r)+s
    return s

p = "/root/.claude/uploads/09725893-9ee0-5a14-a959-b24990f83ec4/132c1690-2025________.csv"
text = open(p, encoding='cp932', newline='').read()
raw = list(csv.reader(io.StringIO(text)))

def cellxml(ref, v, forced_text):
    if v == '': return None
    if not forced_text:
        try:
            fv = float(v)
            return f'<c r="{ref}"><v>{v}</v></c>'  # Excelが数値として取り込むセル
        except ValueError:
            pass
    return f'<c r="{ref}" t="inlineStr"><is><t>{esc(v)}</t></is></c>'

def build_sheet2(rows):
    out=[f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         f'<worksheet xmlns="{NSW}"><sheetData>']
    for ri, row in enumerate(rows, start=1):
        cells=[]
        for ci, (v, forced) in enumerate(row, start=1):
            x = cellxml(f'{colref(ci)}{ri}', v, forced)
            if x: cells.append(x)
        if cells: out.append(f'<row r="{ri}">'+''.join(cells)+'</row>')
    out.append('</sheetData></worksheet>')
    return ''.join(out)

# CSV → (値, 強制テキストか) 行列。="x" は Excel ではテキスト "x" になる
rows=[]
for r in raw:
    line=[]
    for c in r:
        c=c.strip()
        if c.startswith('="') and c.endswith('"'): line.append((c[2:-1], True))
        else: line.append((c, False))
    rows.append(line)

def inject(src, dst, rows):
    sheet = build_sheet2(rows)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as zo:
        for it in zin.infolist():
            data = sheet.encode() if it.filename=='xl/worksheets/sheet2.xml' else zin.read(it.filename)
            zo.writestr(it, data)
    print('written', dst)

SRC='目標値・基準値検索ツール_ver03.xlsx'
inject(SRC, 'testC.xlsx', rows)   # そのまま: 全フラグ9 → 該当0のはず

# 変異版: 一部フラグを1へ (行番号は元データ上の行 = リスト index+1)
import copy
m = copy.deepcopy(rows)
def setv(sheet_row, col, val, forced):  # sheet_row: 1-origin
    m[sheet_row-1][col-1] = (val, forced)
setv(13, 22, '1', True)    # 高低（左）基準値F ← 文字"1"    → 基準値のみ
setv(23, 35, '1', True)    # 通り（左）正矢量確定基準値F     → 基準値のみ(黄=通り（左）実測値)
setv(33, 25, '1', True)    # 高低（左）確定基準値F           → どちらにも出ないこと!
setv(43, 53, '1', True)    # 水準目標値F                     → 目標値のみ
setv(53, 28, '1', True)    # 高低（右）基準値F ┬ 同じ行      → 両方に出る
setv(53, 29, '1', True)    # 高低（右）目標値F ┘
setv(63, 67, '1', False)   # 軌間基準値F ← 数値1             → 基準値のみ
inject(SRC, 'testD.xlsx', m)
print('期待値: 基準値まとめ=4行4件(行13,23,53,63) / 目標値まとめ=2行2件(行43,53)')
