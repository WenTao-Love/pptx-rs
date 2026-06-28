#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深入分析原始 .ppt 文件中 SpContainer 和 0xF004 container 的内部结构。

目的：找出 0xF004 container 内部包含什么 record，
以及 PowerPoint 实际使用的形状结构。
"""

import sys
import struct
import olefile

# Record type 常量
RT_SLIDE = 0x03F8
RT_PPDRAWING = 0x040C
RT_SPGR_CONTAINER = 0xF002
RT_SP_CONTAINER = 0xF003
RT_FSP = 0xF007
RT_FOPT = 0xF008
RT_CLIENT_ANCHOR = 0xF00A
RT_CLIENT_DATA = 0xF00B
RT_CLIENT_TEXTBOX = 0xF00D
RT_TEXT_HEADER_ATOM = 0x0F9F
RT_TEXT_CHARS_ATOM = 0x0FA0
RT_TEXT_BYTES_ATOM = 0x0FA8

# 已知的 OfficeArt record type 名称
OFFICEART_TYPES = {
    0xF000: "DggContainer",
    0xF001: "BStoreContainer",
    0xF002: "SpgrContainer",
    0xF003: "SpContainer",
    0xF004: "UnknownF004",
    0xF005: "SolverContainer",
    0xF006: "FDGG",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "ClientTextbox(atom)",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF00C: "FRITContainer",
    0xF00D: "ClientTextbox(container)",
    0xF00E: "TertiaryFOPT",
    0xF00F: "ChildAnchor",
    0xF010: "FSPGR",
    0xF011: "FConnector",
    0xF012: "FDGGBlock",
}


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def type_name(rec_type):
    return OFFICEART_TYPES.get(rec_type, f"Unknown(0x{rec_type:04X})")


def dump_records_recursive(data, offset, end, indent=0):
    """递归 dump 所有 record。"""
    prefix = "  " * indent
    pos = offset
    count = 0
    while pos + 8 <= end and count < 50:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        tname = type_name(rec_type)
        print(f"{prefix}[{pos:6d}] ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} ({tname}) len={rec_len}")

        # 对于 container，递归解析
        if is_container and rec_len > 0:
            dump_records_recursive(data, pos + 8, pos + 8 + rec_len, indent + 1)
        else:
            # 对于 atom，如果是文本类型，打印内容
            if rec_type == RT_TEXT_CHARS_ATOM and rec_len > 0:
                text_data = data[pos + 8:pos + 8 + rec_len]
                try:
                    text = text_data.decode("utf-16-le")
                    print(f"{prefix}  文本: {text!r}")
                except Exception:
                    print(f"{prefix}  文本(原始): {text_data[:40]!r}")
            elif rec_type == RT_TEXT_BYTES_ATOM and rec_len > 0:
                text_data = data[pos + 8:pos + 8 + rec_len]
                try:
                    text = text_data.decode("latin-1")
                    print(f"{prefix}  文本: {text!r}")
                except Exception:
                    print(f"{prefix}  文本(原始): {text_data[:40]!r}")
            elif rec_type == RT_TEXT_HEADER_ATOM and rec_len >= 4:
                tx_type = struct.unpack_from("<I", data, pos + 8)[0]
                print(f"{prefix}  txType: {tx_type}")
            elif rec_type == RT_FSP and rec_len >= 8:
                shape_id = struct.unpack_from("<I", data, pos + 8)[0]
                flags = struct.unpack_from("<I", data, pos + 12)[0]
                print(f"{prefix}  shapeId={shape_id} flags=0x{flags:08X} (inst=0x{inst:X})")
            elif rec_type == RT_FOPT and rec_len > 0:
                num_props = inst
                print(f"{prefix}  FOPT 属性数: {num_props}")
                fopt_pos = pos + 8
                for i in range(min(num_props, 20)):
                    if fopt_pos + 6 > pos + 8 + rec_len:
                        break
                    pid = struct.unpack_from("<H", data, fopt_pos)[0]
                    pval = struct.unpack_from("<I", data, fopt_pos + 2)[0]
                    print(f"{prefix}    0x{pid:04X}: 0x{pval:08X}")
                    fopt_pos += 6
            elif rec_type == RT_CLIENT_ANCHOR and rec_len > 0:
                anchor_data = data[pos + 8:pos + 8 + rec_len]
                values = []
                for i in range(0, len(anchor_data), 2):
                    if i + 2 <= len(anchor_data):
                        values.append(struct.unpack_from("<H", anchor_data, i)[0])
                print(f"{prefix}  anchor values: {values}")

        pos += 8 + rec_len
        count += 1
        if not is_container and rec_len == 0:
            break
    return pos


def find_all_slides(data):
    slides = []
    pos = 0
    while pos + 8 <= len(data):
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SLIDE:
            slides.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return slides


def find_record(data, start, end, target_type, container_only=True):
    """在 [start, end) 范围内找到第一个指定类型的 record。"""
    pos = start
    while pos + 8 <= end:
        h = parse_header(data, pos)
        if h is None:
            return None
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        if rec_type == target_type:
            if not container_only or is_container:
                return pos
        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break
    return None


def analyze_file(filepath):
    print(f"\n{'='*80}")
    print(f"分析文件: {filepath}")
    print(f"{'='*80}")

    ole = olefile.OleFileIO(filepath)
    stream_names = ["/".join(p) for p in ole.listdir()]
    print(f"OLE2 streams: {stream_names}")
    if "PowerPoint Document" not in stream_names:
        print("找不到 PowerPoint Document stream")
        ole.close()
        return
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()

    print(f"PowerPoint Document stream 大小: {len(ppt_data)} 字节")

    slides = find_all_slides(ppt_data)
    print(f"找到 {len(slides)} 个 Slide")

    # 只分析第一个 Slide
    if not slides:
        return
    slide_offset = slides[0]
    print(f"\n--- Slide 1 (offset=0x{slide_offset:X}) ---")

    # 找到 PPDrawing
    h = parse_header(ppt_data, slide_offset)
    slide_end = slide_offset + 8 + h[3]
    ppd_offset = find_record(ppt_data, slide_offset + 8, slide_end, RT_PPDRAWING)
    if ppd_offset is None:
        print("  找不到 PPDrawing")
        return
    print(f"  PPDrawing offset: 0x{ppd_offset:X}")

    # 找到 SpgrContainer
    h = parse_header(ppt_data, ppd_offset)
    ppd_end = ppd_offset + 8 + h[3]
    spgr_offset = find_record(ppt_data, ppd_offset + 8, ppd_end, RT_SPGR_CONTAINER)
    if spgr_offset is None:
        print("  找不到 SpgrContainer")
        return
    print(f"  SpgrContainer offset: 0x{spgr_offset:X}")

    # dump SpgrContainer 的完整结构
    h = parse_header(ppt_data, spgr_offset)
    spgr_end = spgr_offset + 8 + h[3]
    print(f"\n  SpgrContainer 完整结构 (offset=0x{spgr_offset:X}, end=0x{spgr_end:X}):")
    dump_records_recursive(ppt_data, spgr_offset, spgr_end, 2)


def main():
    import os
    test_dir = "_test"
    if not os.path.isdir(test_dir):
        print(f"找不到 {test_dir} 目录")
        return
    for fname in sorted(os.listdir(test_dir)):
        if fname.lower().endswith(".ppt"):
            analyze_file(os.path.join(test_dir, fname))


if __name__ == "__main__":
    main()
