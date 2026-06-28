#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析原始 .ppt 文件中 PowerPoint Document stream 的顶层 record types。
目的：验证 RT_SLIDE 的正确值（0x03EE vs 0x03F8）。
"""

import sys
import struct
import olefile

# MS-PPT record types
RT_NAMES = {
    0x03E8: "RT_Document",
    0x03EE: "RT_Slide",
    0x03F0: "RT_Notes",
    0x03F2: "RT_Environment",
    0x03F8: "RT_MainMaster",
    0x040B: "RT_DrawingGroup",
    0x040C: "RT_Drawing",
    0x0FF0: "RT_SlideListWithText",
    0x0FF5: "RT_UserEditAtom",
    0x0FF6: "RT_CurrentUserAtom",
    0x1772: "RT_PersistDirectoryAtom",
    0x2F14: "RT_CryptSession10Container",
}


def parse_record_header(data, pos):
    """解析 8 字节 record header。返回 (ver, inst, rec_type, rec_len)。"""
    if pos + 8 > len(data):
        return None
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, pos)
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def analyze_ppt(filepath):
    print(f"\n=== 分析: {filepath} ===", flush=True)
    ole = olefile.OleFileIO(filepath)
    try:
        # olefile.listdir() 返回 [[name1], [name2], ...] 格式
        stream_names = ["/".join(parts) for parts in ole.listdir()]
        print(f"streams: {stream_names}", flush=True)
        if "PowerPoint Document" not in stream_names:
            print("找不到 PowerPoint Document stream", flush=True)
            return
        with ole.openstream("PowerPoint Document") as f:
            data = f.read()
        print(f"PowerPoint Document stream 大小: {len(data)} 字节", flush=True)

        # 遍历顶层 record
        pos = 0
        print("\n顶层 records:", flush=True)
        print(f"{'offset':>10} {'ver':>4} {'inst':>6} {'type':>10} {'len':>10} {'name':>30} {'container':>10}", flush=True)
        slide_count = 0
        master_count = 0
        while pos + 8 <= len(data):
            hdr = parse_record_header(data, pos)
            if hdr is None:
                break
            ver, inst, rec_type, rec_len = hdr
            is_container = ver == 0xF
            name = RT_NAMES.get(rec_type, f"unknown(0x{rec_type:04X})")
            total_len = 8 + rec_len
            print(f"{pos:>10} {ver:>4} {inst:>6} 0x{rec_type:04X} {rec_len:>10} {name:>30} {'YES' if is_container else 'no':>10}", flush=True)

            if rec_type == 0x03EE:
                slide_count += 1
            if rec_type == 0x03F8:
                master_count += 1

            pos += total_len
            if not is_container and rec_len == 0:
                break

        print(f"\n统计: Slide(0x03EE)={slide_count} 个, MainMaster(0x03F8)={master_count} 个", flush=True)
    finally:
        ole.close()


if __name__ == "__main__":
    import os
    test_dir = "_test"
    for fname in os.listdir(test_dir):
        if fname.endswith(".ppt"):
            analyze_ppt(os.path.join(test_dir, fname))
