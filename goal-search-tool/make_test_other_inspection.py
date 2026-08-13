# -*- coding: utf-8 -*-
"""修正版2の 元データ を、別検査を想定した合成データに差し替えたテストブックを作る"""
import zipfile, sys

NSW = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def esc(s):
    return (str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;'))

def colref(n):
    s=''
    while n>0:
        n,r = divmod(n-1,26)
        s = chr(65+r)+s
    return s

def build_sheet(rows):
    out=[f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         f'<worksheet xmlns="{NSW}"><sheetData>']
    for ri, row in enumerate(rows, start=1):
        if row is None:
            continue
        cells=[]
        for ci, v in enumerate(row, start=1):
            if v is None or v == '':
                continue
            ref = f'{colref(ci)}{ri}'
            if isinstance(v,(int,float)):
                cells.append(f'<c r="{ref}"><v>{v}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{esc(v)}</t></is></c>')
        if cells:
            out.append(f'<row r="{ri}">' + ''.join(cells) + '</row>')
    out.append('</sheetData></worksheet>')
    return ''.join(out)

# ---- 別検査を想定した見出し（元の軌道検測とは全く別の列名）----
HEAD = ['No','設備名','測定日',
        '絶縁抵抗','絶縁抵抗基準値F','絶縁抵抗目標F','絶縁抵抗超過F',          # 半角F・「目標」のみ(値なし)
        '接地抵抗','接地抵抗基準値F','接地抵抗目標値Ｆ','接地抵抗超過F',        # 全角Ｆ
        '電圧','電圧目標値f',                                                  # 小文字f
        '温度','温度目標F','温度目標値2F',                                      # 目標F列が2本 同じ値列を共有
        '判定']
# data rows: (絶縁目標, 接地目標, 電圧目標, 温度目標1, 温度目標2)
DATA = [
    #                     絶縁   基準 目標 超過   接地  基準 目標  超過   電圧 目標   温度 目標1 目標2
    ['1','A変電所','2026-07-01', 12.5, 0,   1,   0,    3.2, 0,  0,    0,    99.1, 0,   21.0, 0,   0,   'OK'],
    ['2','B変電所','2026-07-01', 8.4,  0,   0,   0,    2.1, 0, '1',   0,    98.7, 0,   22.5, 0,   0,   'OK'],   # 文字列の"1"
    ['3','C変電所','2026-07-02', 15.0, 0,   0,   0,    4.0, 0,  0,    0,    97.2, 1,   23.1, 0,   0,   'OK'],   # 小文字f列
    ['4','D変電所','2026-07-02', 6.6,  0,   1,   0,    1.8, 0,  1,    0,    96.0, 0,   24.9, 0,   0,   'NG'],   # 1行に2件
    ['5','E変電所','2026-07-03', 20.1, 0,   0,   0,    5.5, 0,  0,    0,    99.9, 0,   25.3, 1,   0,   'OK'],   # 温度目標1
    ['6','F変電所','2026-07-03', 11.1, 0,   0,   0,    3.3, 0,  0,    0,    98.0, 0,   26.7, 0,   1,   'OK'],   # 温度目標2(重複値列)
    ['7','G変電所','2026-07-04', 13.3, 0,   0,   0,    3.9, 0,  0,    0,    98.5, 0,   20.2, 0,   0,   'OK'],   # 該当なし
]

def rows_with_banner():
    return [['1234件のデータが検索されました'], None, HEAD] + DATA   # 1行目=帯, 2行目=空, 3行目=見出し

def rows_plain():
    return [HEAD] + DATA                                            # 見出しが1行目(従来形式)

def make(src, dst, rows):
    sheet = build_sheet(rows)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as zo:
        for it in zin.infolist():
            data = sheet.encode('utf-8') if it.filename=='xl/worksheets/sheet2.xml' else zin.read(it.filename)
            zo.writestr(it, data)
    print('written', dst)

SRC='目標値検索ツール_修正版2_2026.07.16.xlsx'
make(SRC,'testA.xlsx', rows_with_banner())
make(SRC,'testB.xlsx', rows_plain())

# ---- 期待値をPythonで独立に算出 ----
def expect(rows, hdr_idx):
    head = rows[hdr_idx]
    flag = [i for i,h in enumerate(head) if h and '目標' in h and ('f' in h.lower() or 'Ｆ' in h)]
    print('  flag cols(0-origin):', flag, [head[i] for i in flag])
    nrow=0; nflag=0; hits=[]
    for r,row in enumerate(rows[hdr_idx+1:], start=hdr_idx+2):
        k=sum(1 for i in flag if str(row[i]).strip()=='1')
        if k: nrow+=1; nflag+=k; hits.append((r,k))
    print('  期待: 見出し行=%d 該当行=%d 総件数=%d %s' % (hdr_idx+1, nrow, nflag, hits))
print('testA'); expect(rows_with_banner(), 2)
print('testB'); expect(rows_plain(), 0)
