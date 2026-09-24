#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MS-OVBA (VBAプロジェクト) の圧縮コンテナ 読み書き。

VBAのソースは「Compressed Container」という独自形式で格納されている。
仕様: [MS-OVBA] 2.4.1 Compression and Decompression
"""
import struct


def decompress(data):
    """Compressed Container → 生バイト列"""
    if not data or data[0] != 0x01:
        raise ValueError("先頭バイトが 0x01 ではありません")
    out = bytearray()
    pos = 1
    while pos + 1 < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        size = (header & 0x0FFF) + 3
        sig = (header >> 12) & 0x07
        flag = (header >> 15) & 0x01
        if sig != 0b011:
            raise ValueError(f"チャンク署名が不正: {sig:#b} (pos={pos-2})")
        end = pos + size - 2
        chunk = data[pos:end]
        if flag == 0:
            out += chunk
        else:
            out += _decompress_chunk(chunk, len(out))
        pos = end
    return bytes(out)


def _decompress_chunk(chunk, dec_start_total):
    dec = bytearray()
    i = 0
    while i < len(chunk):
        flags = chunk[i]
        i += 1
        for bit in range(8):
            if i >= len(chunk):
                break
            if not (flags >> bit) & 1:
                dec.append(chunk[i])
                i += 1
            else:
                token = struct.unpack_from("<H", chunk, i)[0]
                i += 2
                bit_count = _bit_count(len(dec))
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = ((token & ~length_mask & 0xFFFF) >> (16 - bit_count)) + 1
                src = len(dec) - offset
                for k in range(length):
                    dec.append(dec[src + k])
    return bytes(dec)


def _bit_count(diff):
    """CopyTokenHelp: 現在位置からビット幅を決める"""
    n = 4
    while (1 << n) < diff:
        n += 1
    return max(4, min(n, 12))


def compress(data):
    """生バイト列 → Compressed Container

    4096バイトごとのチャンクに分け、各チャンク内で後方一致を探して
    CopyToken に置き換える(仕様どおりの圧縮)。
    """
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        block = data[start:start + 4096]
        body = _compress_chunk(block)
        if len(body) + 2 < 4098:
            header = 0x3000 | 0x8000 | (len(body) + 2 - 3)
            out += struct.pack("<H", header) + body
        else:
            # 圧縮して縮まらない場合は非圧縮チャンク(必ず4096バイト)
            blk = block + b"\x00" * (4096 - len(block))
            out += struct.pack("<H", 0x3000 | 0x0FFF) + blk
    return bytes(out)


def _compress_chunk(block):
    """1チャンク(最大4096バイト)を圧縮"""
    out = bytearray()
    tokens = []          # (is_copy, value)
    pos = 0
    n = len(block)
    while pos < n:
        best_len, best_off = 0, 0
        if pos > 0:
            bit_count = _bit_count(pos)
            max_len = (0xFFFF >> bit_count) + 3
            max_off = 1 << bit_count
            lo = max(0, pos - max_off)
            # 後方一致の最長を探す
            limit = min(max_len, n - pos)
            if limit >= 3:
                for cand in range(pos - 1, lo - 1, -1):
                    if block[cand] != block[pos]:
                        continue
                    ln = 0
                    while ln < limit and block[cand + ln] == block[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                        if ln == limit:
                            break
        if best_len >= 3:
            bit_count = _bit_count(pos)
            length_mask = 0xFFFF >> bit_count
            token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
            tokens.append((True, token))
            pos += best_len
        else:
            tokens.append((False, block[pos]))
            pos += 1

    for i in range(0, len(tokens), 8):
        group = tokens[i:i + 8]
        flags = 0
        for j, (is_copy, _) in enumerate(group):
            if is_copy:
                flags |= (1 << j)
        out.append(flags)
        for is_copy, val in group:
            if is_copy:
                out += struct.pack("<H", val)
            else:
                out.append(val)
    return bytes(out)
