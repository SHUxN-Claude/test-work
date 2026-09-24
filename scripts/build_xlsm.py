#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一覧ツールの .xlsm を組み立てる。

1) 動作実績のある .xlsm から vbaProject.bin を取り出す
2) モジュールの中身を modList に差し替える(古いコンパイル済みコードは除去)
3) 元の .xlsx を ZIP レベルでコピーし、vbaProject.bin を足して .xlsm にする
   → openpyxl を通さないので、書式・入力規則・印刷設定などは一切変化しない
"""
import shutil
import struct
import sys
import zipfile

import cfb
import ovba


def find_source(stream):
    """モジュールのストリームから圧縮ソース部分を探して (offset, 生ソース) を返す"""
    for off in range(len(stream)):
        if stream[off] != 0x01:
            continue
        try:
            dec = ovba.decompress(stream[off:])
        except Exception:
            continue
        if dec.startswith(b"Attribute VB_Name"):
            return off, dec
    raise ValueError("ソースが見つかりません")


def build_vba(src_bin, out_bin, module_src):
    ents, streams = cfb.read(src_bin)
    by_name = {e.name: e for e in ents}
    sid = {n: e.sid for n, e in by_name.items()}

    # --- すべてのモジュールを「圧縮ソースのみ(TextOffset=0)」にする ---
    docs = ["ThisWorkbook", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
    for nm in docs:
        if nm not in sid:
            continue
        _, source = find_source(streams[sid[nm]])
        streams[sid[nm]] = ovba.compress(source)     # 中身はそのまま、先頭に移す

    streams[sid["modMain"]] = ovba.compress(module_src)
    by_name["modMain"].set_name("modList")           # 7文字→7文字なのでツリー不変
    for nm in ("modEngine", "modEmbedded"):
        stub = (f'Attribute VB_Name = "{nm}"\r\n'
                "'（このツールでは使用しません。削除しても構いません）\r\n").encode("cp932")
        streams[sid[nm]] = ovba.compress(stub)

    # --- dir: MODULEOFFSET を 0 に + 改名 ---
    d = bytearray(ovba.decompress(streams[sid["dir"]]))
    pat = struct.pack("<HI", 0x0031, 4)
    pos, cnt = 0, 0
    while True:
        i = d.find(pat, pos)
        if i < 0:
            break
        struct.pack_into("<I", d, i + 6, 0)
        cnt += 1
        pos = i + 10
    u_old, u_new = "modMain".encode("utf-16-le"), "modList".encode("utf-16-le")
    d = bytearray(bytes(d).replace(b"modMain", b"modList").replace(u_old, u_new))
    streams[sid["dir"]] = ovba.compress(bytes(d))

    # --- PROJECT / PROJECTwm も改名 ---
    streams[sid["PROJECT"]] = streams[sid["PROJECT"]].replace(b"modMain", b"modList")
    streams[sid["PROJECTwm"]] = (streams[sid["PROJECTwm"]]
                                 .replace(b"modMain", b"modList")
                                 .replace(u_old, u_new))

    # --- 古いプロジェクトキャッシュを空に ---
    for e in ents:
        if e.name.startswith("__SRP_"):
            streams[e.sid] = b""

    size = cfb.write(out_bin, ents, streams)
    return cnt, size


CT_OLD = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_NEW = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
REL_NS = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def build_xlsm(src_xlsx, out_xlsm, vba_bin):
    """ZIPを一から作り直さず、元のエントリをそのままコピーして vbaProject を足す"""
    zin = zipfile.ZipFile(src_xlsx)
    names = zin.namelist()
    with zipfile.ZipFile(out_xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                t = data.decode("utf-8")
                if CT_OLD in t:
                    t = t.replace(CT_OLD, CT_NEW)
                # bin拡張子は printerSettings 用の既定が既にあることが多い。
                # 既定より Override が優先されるので、必ず Override で宣言する。
                if "/xl/vbaProject.bin" not in t:
                    t = t.replace(
                        "</Types>",
                        '<Override PartName="/xl/vbaProject.bin" '
                        'ContentType="application/vnd.ms-office.vbaProject"/></Types>', 1)
                data = t.encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels":
                t = data.decode("utf-8")
                if "vbaProject.bin" not in t:
                    t = t.replace(
                        "</Relationships>",
                        f'<Relationship Id="rIdVBAProject" Type="{REL_NS}" '
                        'Target="vbaProject.bin"/></Relationships>')
                data = t.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                t = data.decode("utf-8")
                if "codeName=" not in t:
                    if "<workbookPr" in t:
                        t = t.replace("<workbookPr", '<workbookPr codeName="ThisWorkbook"', 1)
                    else:
                        t = t.replace("<fileVersion",
                                      '<workbookPr codeName="ThisWorkbook"/><fileVersion', 1)
                data = t.encode("utf-8")
            zout.writestr(item, data)
        zout.write(vba_bin, "xl/vbaProject.bin")
    zin.close()
    return names


def main():
    src_xlsm, src_xlsx, out_xlsm, bas = sys.argv[1:5]
    source = open(bas, "rb").read().decode("utf-8").replace("\n", "\r\n").encode("cp932")
    cnt, size = build_vba(src_xlsm and "vbaProject_src.bin", "vbaProject_new.bin", source)
    print(f"vbaProject.bin: {size} bytes (MODULEOFFSET={cnt}箇所を0に)")
    build_xlsm(src_xlsx, out_xlsm, "vbaProject_new.bin")
    print(f"出力: {out_xlsm}")


if __name__ == "__main__":
    main()
