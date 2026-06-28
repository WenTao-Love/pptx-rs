#!/usr/bin/env python3
"""分析原始 .ppt 文件中 Slide 的 PPDrawing 结构，对比水印 SpContainer 差异。

目标：找出水印不显示的根因。
"""

import struct
import sys
from pathlib import Path

try:
    import olefile
except ImportError:
    print("需要 olefile: pip install olefile")
    sys.exit(1)

# Record type 常量
RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_SPGR_CONTAINER = 0xF002
RT_SP_CONTAINER = 0xF003
RT_FSP = 0xF007
RT_FOPT = 0xF008
RT_CLIENT_ANCHOR = 0xF00A
RT_CLIENT_DATA = 0xF00B
RT_CLIENT_TEXTBOX = 0xF00D
RT_FSPGR = 0xF009
RT_FCONNECTOR = 0xF011

# 未知类型 0xF004（原始文件中出现）
RT_UNKNOWN_F004 = 0xF004

RECORD_TYPE_NAMES = {
    0x03E8: "Document",
    0x03EE: "Slide",
    0x03F8: "MainMaster",
    0x040C: "PPDrawing",
    0xF002: "SpgrContainer",
    0xF003: "SpContainer",
    0xF004: "DggContainer(?)",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "FSPGR",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF00D: "ClientTextbox",
    0xF011: "FConnector",
    0x0F9F: "TextHeaderAtom",
    0x0FA0: "TextCharsAtom",
    0x0FA8: "TextRulerAtom",
    0x0FBA: "StyleAtom",
}


def parse_header(data, offset):
    """解析 8 字节 record header，返回 (ver, inst, recType, recLen)。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def type_name(t):
    return RECORD_TYPE_NAMES.get(t, f"0x{t:04X}")


def dump_record_tree(data, offset, end, indent=0):
    """递归打印 record 树。"""
    prefix = "  " * indent
    results = []
    while offset + 8 <= end:
        h = parse_header(data, offset)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if offset + total_len > end + 8:  # 越界保护
            break

        name = type_name(rec_type)
        results.append(
            f"{prefix}offset={offset:>8} ver={ver} inst=0x{inst:03X} type=0x{rec_type:04X}({name}) len={rec_len} container={is_container}"
        )

        if is_container:
            # 递归打印子 record
            child_end = offset + 8 + rec_len
            results.extend(dump_record_tree(data, offset + 8, child_end, indent + 1))

        offset += total_len
        if not is_container and rec_len == 0:
            break
    return results


def find_all_slides(data):
    """找到所有 Slide record 的 offset。"""
    slides = []
    pos = 0
    while pos + 8 <= len(data):
        h = parse_header(data, pos)
        if h is None:
            break
        ver, _, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SLIDE:
            slides.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return slides


def find_ppdrawing(data, slide_offset):
    """在 Slide 中找到 PPDrawing 的 offset。"""
    h = parse_header(data, slide_offset)
    if h is None:
        return None
    _, _, _, slide_len = h
    slide_end = slide_offset + 8 + slide_len
    pos = slide_offset + 8
    while pos + 8 <= slide_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, _, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_PPDRAWING:
            return pos
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return None


def analyze_file(path, label, max_slides=2):
    """分析文件中前 N 个 Slide 的 PPDrawing 结构。"""
    print(f"\n{'='*70}")
    print(f"{label}: {path}")
    print(f"{'='*70}")

    ole = olefile.OleFileIO(path)
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()

    slides = find_all_slides(ppt_data)
    print(f"共找到 {len(slides)} 个 Slide")

    for i, slide_offset in enumerate(slides[:max_slides]):
        print(f"\n--- Slide #{i+1} (offset={slide_offset}) ---")
        ppd_offset = find_ppdrawing(ppt_data, slide_offset)
        if ppd_offset is None:
            print("  找不到 PPDrawing")
            continue

        h = parse_header(ppt_data, ppd_offset)
        _, _, _, ppd_len = h
        ppd_end = ppd_offset + 8 + ppd_len

        print(f"PPDrawing offset={ppd_offset} len={ppd_len}")
        tree = dump_record_tree(ppt_data, ppd_offset, ppd_end, 1)
        for line in tree:
            print(line)


def main():
    test_dir = Path("_test")
    out_dir = Path("_test_out")

    # 分析原始文件
    for ppt in test_dir.glob("*.ppt"):
        analyze_file(ppt, "原始文件", max_slides=2)

    # 分析水印文件
    for ppt in out_dir.glob("wm_*.ppt"):
        analyze_file(ppt, "水印文件", max_slides=2)


if __name__ == "__main__":
    main()
