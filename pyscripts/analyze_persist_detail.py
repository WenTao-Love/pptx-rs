#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析 persist 目录和 msoffcrypto 的解密范围。
关键问题：解密后只有前 3 个 record 正确，从第 4 个开始被破坏。
"""
import io
import os
import struct
import olefile
import msoffcrypto

PASSWORD = "pptx-rs-secret"


def parse_record_header(data, pos):
    if pos + 8 > len(data):
        return None
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, pos)
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def analyze_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset)] 列表。"""
    hdr = parse_record_header(data, offset)
    if hdr is None:
        return []
    ver, inst, rec_type, rec_len = hdr
    print(f"  PersistDirectoryAtom: offset={offset}, len={rec_len}", flush=True)

    pd_data = data[offset + 8: offset + 8 + rec_len]
    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry_val = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        print(f"    PersistDirectoryEntry: persistId={persist_id}, cPersist={c_persist}", flush=True)
        pos += 4
        for i in range(c_persist):
            if pos + 4 <= len(pd_data):
                off = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + i, off))
                pos += 4
    return entries


def analyze_original(filepath):
    """分析原始文件的 persist 目录和顶层 record。"""
    print(f"\n=== 分析原始文件: {os.path.basename(filepath)} ===", flush=True)
    ole = olefile.OleFileIO(filepath)
    try:
        with ole.openstream("Current User") as f:
            cu_data = f.read()
        with ole.openstream("PowerPoint Document") as f:
            ppt_data = f.read()

        offset_to_current_edit = struct.unpack_from("<I", cu_data, 16)[0]
        print(f"  offsetToCurrentEdit = {offset_to_current_edit}", flush=True)

        # UserEditAtom
        ue_offset = offset_to_current_edit
        ue_hdr = parse_record_header(ppt_data, ue_offset)
        print(f"  UserEditAtom: offset={ue_offset}, len={ue_hdr[3]}", flush=True)
        offset_persist_dir = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
        print(f"  offsetPersistDirectory = {offset_persist_dir}", flush=True)

        # Persist 目录
        entries = analyze_persist_directory(ppt_data, offset_persist_dir)

        # 列出所有 persist 对象
        print(f"\n  Persist 对象列表 ({len(entries)} 个):", flush=True)
        rt_names = {
            0x03E8: "Document", 0x03EE: "Slide", 0x03F0: "Notes",
            0x03F8: "MainMaster", 0x0FF0: "SlideListWithText",
            0x0FF5: "UserEditAtom", 0x1772: "PersistDirAtom",
            0x2F14: "CryptSession10",
        }
        for pid, off in entries:
            hdr = parse_record_header(ppt_data, off)
            if hdr:
                ver, inst, rec_type, rec_len = hdr
                name = rt_names.get(rec_type, f"0x{rec_type:04X}")
                total = 8 + rec_len
                print(f"    pid={pid:>3} offset={off:>8} type=0x{rec_type:04X}({name:>15}) len={rec_len:>8} total={total:>8}", flush=True)
            else:
                print(f"    pid={pid:>3} offset={off:>8} [无法解析]", flush=True)

        # 计算相邻 persist 对象的偏移差
        print(f"\n  相邻 persist 对象偏移差（msoffcrypto 用此计算加密范围）:", flush=True)
        sorted_entries = sorted(entries, key=lambda x: x[1])
        for i in range(len(sorted_entries)):
            pid, off = sorted_entries[i]
            if i + 1 < len(sorted_entries):
                next_off = sorted_entries[i + 1][1]
                diff = next_off - off
                hdr = parse_record_header(ppt_data, off)
                if hdr:
                    rec_len = hdr[3]
                    total = 8 + rec_len
                    match = "OK" if diff == total else f"MISMATCH (diff={diff}, 8+recLen={total})"
                    print(f"    pid={pid:>3} off={off:>8} next_off={next_off:>8} diff={diff:>8} 8+recLen={total:>8} {match}", flush=True)
                else:
                    print(f"    pid={pid:>3} off={off:>8} [无法解析]", flush=True)
            else:
                # 最后一个
                hdr = parse_record_header(ppt_data, off)
                if hdr:
                    rec_len = hdr[3]
                    total = 8 + rec_len
                    print(f"    pid={pid:>3} off={off:>8} [最后一个] 8+recLen={total:>8}", flush=True)

    finally:
        ole.close()


if __name__ == "__main__":
    test_dir = "_test"
    for fname in sorted(os.listdir(test_dir)):
        if fname.endswith(".ppt"):
            analyze_original(os.path.join(test_dir, fname))
