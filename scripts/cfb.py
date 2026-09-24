#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CFB (Compound File Binary / OLE複合ファイル) の最小リーダ・ライタ。

vbaProject.bin を読み、ディレクトリエントリ(名前・ツリー構造)はそのまま保ったまま
ストリームの中身だけ差し替えて書き戻すために使う。
ツリーをExcelが作ったまま維持できるので、壊れにくい。
"""
import struct

SECSZ = 512
MINISZ = 64
MINICUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class Entry:
    """ディレクトリエントリ(128バイト)"""
    __slots__ = ("raw", "name", "etype", "start", "size", "sid")

    def __init__(self, raw, sid):
        self.raw = bytearray(raw)
        self.sid = sid
        nlen = struct.unpack_from("<H", raw, 64)[0]
        self.name = raw[:max(0, nlen - 2)].decode("utf-16-le") if nlen >= 2 else ""
        self.etype = raw[66]
        self.start = struct.unpack_from("<I", raw, 116)[0]
        self.size = struct.unpack_from("<Q", raw, 120)[0]

    def set_name(self, name):
        b = name.encode("utf-16-le") + b"\x00\x00"
        if len(b) > 64:
            raise ValueError("名前が長すぎます")
        self.raw[0:64] = b + b"\x00" * (64 - len(b))
        struct.pack_into("<H", self.raw, 64, len(b))
        self.name = name

    def finish(self, start, size):
        struct.pack_into("<I", self.raw, 116, start)
        struct.pack_into("<Q", self.raw, 120, size)
        return bytes(self.raw)


def read(path):
    """→ (entries, {sid: bytes})"""
    with open(path, "rb") as f:
        data = f.read()
    if data[:8] != SIG:
        raise ValueError("CFBではありません")
    (nfat, dir_start, _, mini_cut, minifat_start, nminifat,
     difat_start, ndifat) = struct.unpack_from("<IIIIIIII", data, 44)

    def sector(n):
        off = SECSZ + n * SECSZ
        return data[off:off + SECSZ]

    # --- DIFAT → FAT ---
    difat = list(struct.unpack_from("<109I", data, 76))
    nxt = difat_start
    while nxt not in (ENDOFCHAIN, FREESECT) and len(difat) < nfat + 109:
        s = sector(nxt)
        difat += list(struct.unpack_from("<127I", s, 0))
        nxt = struct.unpack_from("<I", s, 508)[0]
    fat = []
    for fs in difat[:nfat]:
        if fs == FREESECT:
            continue
        fat += list(struct.unpack_from("<128I", sector(fs), 0))

    def chain(start):
        out, n, guard = [], start, 0
        while n not in (ENDOFCHAIN, FREESECT) and guard < 1000000:
            out.append(n)
            n = fat[n] if n < len(fat) else ENDOFCHAIN
            guard += 1
        return out

    def read_chain(start, size=None):
        buf = b"".join(sector(n) for n in chain(start))
        return buf[:size] if size is not None else buf

    # --- ディレクトリ ---
    dirbuf = read_chain(dir_start)
    entries = []
    for i in range(len(dirbuf) // 128):
        raw = dirbuf[i * 128:(i + 1) * 128]
        if raw[66] == 0:          # 未使用
            entries.append(Entry(raw, i))
            continue
        entries.append(Entry(raw, i))

    # --- ミニFAT / ミニストリーム ---
    minifat = []
    for n in chain(minifat_start):
        minifat += list(struct.unpack_from("<128I", sector(n), 0))
    root = entries[0]
    ministream = read_chain(root.start, root.size)

    def mini_chain(start):
        out, n, guard = [], start, 0
        while n not in (ENDOFCHAIN, FREESECT) and guard < 1000000:
            out.append(n)
            n = minifat[n] if n < len(minifat) else ENDOFCHAIN
            guard += 1
        return out

    streams = {}
    for e in entries:
        if e.etype != 2:          # ストリーム以外
            continue
        if e.size < MINICUTOFF:
            buf = b"".join(ministream[i * MINISZ:(i + 1) * MINISZ]
                           for i in mini_chain(e.start))
        else:
            buf = read_chain(e.start)
        streams[e.sid] = buf[:e.size]
    return entries, streams


def write(path, entries, streams):
    """entries(元のツリーをそのまま)と streams({sid: bytes}) から CFB を組み立てる"""
    sectors = []              # 通常セクタの中身

    def alloc(buf):
        """バイト列を通常セクタに配置し、開始セクタ番号とチェーンを返す"""
        if not buf:
            return ENDOFCHAIN, []
        start = len(sectors)
        for i in range(0, len(buf), SECSZ):
            blk = buf[i:i + SECSZ]
            sectors.append(blk + b"\x00" * (SECSZ - len(blk)))
        return start, list(range(start, len(sectors)))

    # --- ミニストリームに入れるもの / 通常セクタに入れるもの ---
    mini_buf = bytearray()
    mini_alloc = {}           # sid → (開始ミニセクタ, 個数)
    big_alloc = {}
    for e in entries:
        if e.etype != 2:
            continue
        data = streams.get(e.sid, b"")
        if len(data) == 0:
            continue
        if len(data) < MINICUTOFF:
            start = len(mini_buf) // MINISZ
            mini_buf += data
            if len(mini_buf) % MINISZ:
                mini_buf += b"\x00" * (MINISZ - len(mini_buf) % MINISZ)
            mini_alloc[e.sid] = (start, (len(data) + MINISZ - 1) // MINISZ)
        else:
            big_alloc[e.sid] = data

    # 1) 大きいストリーム
    big_start = {}
    for sid, data in big_alloc.items():
        s, _ = alloc(data)
        big_start[sid] = s

    # 2) ミニストリーム本体
    mini_start, _ = alloc(bytes(mini_buf))

    # 3) ミニFAT
    nmini = len(mini_buf) // MINISZ
    minifat = []
    for sid, (st, cnt) in mini_alloc.items():
        for k in range(cnt):
            minifat.append(ENDOFCHAIN if k == cnt - 1 else st + k + 1)
    # 連番で詰めているので、位置順に組み直す
    minifat = [FREESECT] * nmini
    for sid, (st, cnt) in mini_alloc.items():
        for k in range(cnt):
            minifat[st + k] = ENDOFCHAIN if k == cnt - 1 else st + k + 1
    while len(minifat) % 128:
        minifat.append(FREESECT)
    minifat_buf = struct.pack(f"<{len(minifat)}I", *minifat) if minifat else b""
    minifat_start, minifat_chain = alloc(minifat_buf)

    # 4) ディレクトリ
    dirbuf = bytearray()
    for e in entries:
        if e.etype == 5:                       # Root Entry
            dirbuf += e.finish(mini_start if nmini else ENDOFCHAIN, len(mini_buf))
        elif e.etype == 2:
            data = streams.get(e.sid, b"")
            if len(data) == 0:
                dirbuf += e.finish(ENDOFCHAIN, 0)
            elif e.sid in mini_alloc:
                dirbuf += e.finish(mini_alloc[e.sid][0], len(data))
            else:
                dirbuf += e.finish(big_start[e.sid], len(data))
        else:
            dirbuf += e.finish(e.start if e.etype == 1 else ENDOFCHAIN, 0)
    while len(dirbuf) % SECSZ:
        dirbuf += b"\xff" * 4 + b"\x00" * 124 if False else b"\x00" * (SECSZ - len(dirbuf) % SECSZ)
    dir_start, dir_chain = alloc(bytes(dirbuf))

    # 5) FAT(自分自身とDIFATの分も要るので収束させる)
    ndata = len(sectors)
    nfat = 1
    ndifat = 0
    while True:
        total = ndata + nfat + ndifat
        need_fat = (total + 127) // 128
        need_difat = 0 if need_fat <= 109 else ((need_fat - 109) + 126) // 127
        if need_fat == nfat and need_difat == ndifat:
            break
        nfat, ndifat = need_fat, need_difat

    fat_sectors = list(range(ndata, ndata + nfat))
    difat_sectors = list(range(ndata + nfat, ndata + nfat + ndifat))

    fat = [FREESECT] * (nfat * 128)

    def set_chain(chain_list):
        for i, s in enumerate(chain_list):
            fat[s] = ENDOFCHAIN if i == len(chain_list) - 1 else chain_list[i + 1]

    # データ側のチェーンを張り直す
    pos = 0
    for sid, data in big_alloc.items():
        cnt = (len(data) + SECSZ - 1) // SECSZ
        set_chain(list(range(big_start[sid], big_start[sid] + cnt)))
    if nmini:
        cnt = (len(mini_buf) + SECSZ - 1) // SECSZ
        set_chain(list(range(mini_start, mini_start + cnt)))
    if minifat_chain:
        set_chain(minifat_chain)
    set_chain(dir_chain)
    for s in fat_sectors:
        fat[s] = FATSECT
    for s in difat_sectors:
        fat[s] = DIFSECT

    fat_buf = struct.pack(f"<{len(fat)}I", *fat)

    # --- ヘッダ ---
    hdr = bytearray(b"\x00" * SECSZ)
    hdr[0:8] = SIG
    struct.pack_into("<HH", hdr, 24, 0x003E, 0x0003)      # minor, major
    struct.pack_into("<H", hdr, 28, 0xFFFE)               # byte order
    struct.pack_into("<HH", hdr, 30, 9, 6)                # sector shift, mini shift
    struct.pack_into("<I", hdr, 40, 0)                    # dir sector 数(v3は0)
    struct.pack_into("<I", hdr, 44, nfat)
    struct.pack_into("<I", hdr, 48, dir_start)
    struct.pack_into("<I", hdr, 56, MINICUTOFF)
    struct.pack_into("<I", hdr, 60, minifat_start if minifat_buf else ENDOFCHAIN)
    struct.pack_into("<I", hdr, 64, len(minifat_chain))
    struct.pack_into("<I", hdr, 68, difat_sectors[0] if difat_sectors else ENDOFCHAIN)
    struct.pack_into("<I", hdr, 72, ndifat)
    head_difat = fat_sectors[:109] + [FREESECT] * max(0, 109 - len(fat_sectors))
    struct.pack_into("<109I", hdr, 76, *head_difat[:109])

    # --- 出力 ---
    out = bytearray(hdr)
    for s in sectors:
        out += s
    for i in range(nfat):
        out += fat_buf[i * SECSZ:(i + 1) * SECSZ]
    rest = fat_sectors[109:]
    for i, ds in enumerate(difat_sectors):
        part = rest[i * 127:(i + 1) * 127]
        part = part + [FREESECT] * (127 - len(part))
        nxt = difat_sectors[i + 1] if i + 1 < len(difat_sectors) else ENDOFCHAIN
        out += struct.pack("<127I", *part) + struct.pack("<I", nxt)
    with open(path, "wb") as f:
        f.write(bytes(out))
    return len(out)
