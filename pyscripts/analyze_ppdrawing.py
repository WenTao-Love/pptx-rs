#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
递归分析原始 .ppt 文件中第一个 Slide 的 PPDrawing 结构，
确认 SpContainer/SpgrContainer/FSP/FOPT/ClientAnchor/ClientTextbox 的正确 record type。
"""
import struct
import olefile

# 已知的 MS-ODRAW record type 名称
RT_NAMES = {
    0xF000: "DggContainer",
    0xF001: "BStoreContainer",
    0xF002: "DgContainer",
    0xF003: "SpgrContainer",
    0xF004: "SpContainer",
    0xF005: "SolverContainer",
    0xF006: "FBSE",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "FSPGR",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF00D: "ClientTextbox",
    0x040C: "PPDrawing",
    0x0F9F: "TextHeaderAtom",
    0x0FA0: "TextCharsAtom",
    0x0FA8: "TextBytesAtom",
    0x0FA1: "StyleTextPropAtom",
}


def parse_header(data, pos):
    if pos + 8 > len(data):
        return None
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, pos)
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def dump_records(data, start, end, indent=0):
    """递归遍历 record，打印结构。"""
    pos = start
    while pos + 8 <= end:
        hdr = parse_header(data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len = hdr
        is_container = ver == 0xF
        name = RT_NAMES.get(rec_type, f"unknown(0x{rec_type:04X})")
        prefix = "  " * indent
        total_len = 8 + rec_len
        print(f"{prefix}@{pos}: ver={ver} inst=0x{inst:03X} type=0x{rec_type:04X}({name}) len={rec_len} total={total_len}", flush=True)

        if is_container and rec_len > 0:
            dump_records(data, pos + 8, pos + total_len, indent + 1)

        pos += total_len
        if not is_container and rec_len == 0:
            break


def analyze_first_slide_ppdrawing(filepath):
    print(f"\n=== 分析 {filepath} 的第一个 Slide PPDrawing ===", flush=True)
    ole = olefile.OleFileIO(filepath)
    try:
        with ole.openstream("PowerPoint Document") as f:
            data = f.read()

        # 找到第一个 Slide (0x03EE)
        pos = 0
        slide_found = False
        while pos + 8 <= len(data):
            hdr = parse_header(data, pos)
            if hdr is None:
                break
            ver, inst, rec_type, rec_len = hdr
            is_container = ver == 0xF
            total_len = 8 + rec_len

            if is_container and rec_type == 0x03EE:
                print(f"找到 Slide @ {pos}, len={rec_len}", flush=True)
                slide_end = pos + total_len
                # 在 Slide 中找 PPDrawing (0x040C)
                inner_pos = pos + 8
                while inner_pos + 8 <= slide_end:
                    inner_hdr = parse_header(data, inner_pos)
                    if inner_hdr is None:
                        break
                    iver, iinst, itype, ilen = inner_hdr
                    if iver == 0xF and itype == 0x040C:
                        print(f"\n找到 PPDrawing @ {inner_pos}, len={ilen}", flush=True)
                        ppd_end = inner_pos + 8 + ilen
                        print("PPDrawing 结构树:", flush=True)
                        dump_records(data, inner_pos + 8, ppd_end, 1)
                        slide_found = True
                        break
                    inner_pos += 8 + ilen
                    if iver != 0xF and ilen == 0:
                        break
                if slide_found:
                    break

            pos += total_len
            if not is_container and rec_len == 0:
                break

        if not slide_found:
            print("未找到 Slide 或 PPDrawing", flush=True)
    finally:
        ole.close()


if __name__ == "__main__":
    import os
    for fname in os.listdir("_test"):
        if fname.endswith(".ppt"):
            analyze_first_slide_ppdrawing(os.path.join("_test", fname))
