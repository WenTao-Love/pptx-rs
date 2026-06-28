#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对比原始 .ppt 文件中已有的 SpContainer 结构与水印 SpContainer 的差异。
重点查看带文本框的 SpContainer 的完整结构（包括 ClientData）。
"""

import sys
import struct
import olefile

RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_SPGR_CONTAINER = 0xF003
RT_SP_CONTAINER = 0xF004
RT_FSPGR = 0xF009
RT_FSP = 0xF00A
RT_FOPT = 0xF00B
RT_CLIENT_TEXTBOX = 0xF00D
RT_CLIENT_ANCHOR = 0xF010
RT_CLIENT_DATA = 0xF011
RT_TEXT_HEADER_ATOM = 0x0F9F
RT_TEXT_CHARS_ATOM = 0x0FA0

OFFICEART_TYPES = {
    0xF000: "DggContainer",
    0xF001: "BStoreContainer",
    0xF002: "DgContainer",
    0xF003: "SpgrContainer",
    0xF004: "SpContainer",
    0xF005: "SolverContainer",
    0xF006: "FBSE",
    0xF007: "FSP(old)",
    0xF008: "FOPT(old)",
    0xF009: "FSPGR",
    0xF00A: "FSP",
    0xF00B: "FOPT",
    0xF00D: "ClientTextbox",
    0xF010: "ClientAnchor",
    0xF011: "ClientData",
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


def find_all_spcontainers(data, spgr_offset):
    """在 SpgrContainer 中找到所有 SpContainer（包括第一个 SpgrContainer 子 record）。"""
    spcontainers = []
    h = parse_header(data, spgr_offset)
    if h is None:
        return spcontainers
    _, _, _, spgr_len = h
    spgr_end = spgr_offset + 8 + spgr_len
    pos = spgr_offset + 8
    while pos + 8 <= spgr_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        # 同时收集 SpContainer 和 SpgrContainer（组形状本身）
        if is_container and (rec_type == RT_SP_CONTAINER or rec_type == RT_SPGR_CONTAINER):
            spcontainers.append((pos, rec_type))
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return spcontainers


def dump_spcontainer(data, sp_offset, indent=0, label=""):
    """dump SpContainer 的子 record。"""
    prefix = "  " * indent
    h = parse_header(data, sp_offset)
    if h is None:
        return
    ver, inst, rec_type, rec_len = h
    tname = type_name(rec_type)
    print(f"{prefix}{tname} (offset=0x{sp_offset:X}, ver=0x{ver:X}, inst=0x{inst:03X}, len={rec_len}) {label}")

    pos = sp_offset + 8
    sp_end = sp_offset + 8 + rec_len
    while pos + 8 <= sp_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        tname = type_name(rec_type)
        print(f"{prefix}  [{pos:6d}] ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} ({tname}) len={rec_len}")

        if rec_type == RT_FSP and rec_len >= 8:
            shape_id = struct.unpack_from("<I", data, pos + 8)[0]
            flags = struct.unpack_from("<I", data, pos + 12)[0]
            print(f"{prefix}    FSP: inst=0x{inst:X} (shapeType) shapeId={shape_id} flags=0x{flags:08X}")
        elif rec_type == RT_FSPGR:
            # FSPGR: 4 个 int32 (left, top, right, bottom)
            if rec_len >= 16:
                l, t, r, b = struct.unpack_from("<iiii", data, pos + 8)
                print(f"{prefix}    FSPGR: left={l} top={t} right={r} bottom={b}")
        elif rec_type == RT_FOPT and rec_len > 0:
            num_props = inst
            print(f"{prefix}    FOPT 属性数: {num_props}")
            fopt_pos = pos + 8
            for i in range(min(num_props, 30)):
                if fopt_pos + 6 > pos + 8 + rec_len:
                    break
                pid = struct.unpack_from("<H", data, fopt_pos)[0]
                pval = struct.unpack_from("<I", data, fopt_pos + 2)[0]
                print(f"{prefix}      0x{pid:04X}: 0x{pval:08X}")
                fopt_pos += 6
        elif rec_type == RT_CLIENT_ANCHOR and rec_len > 0:
            anchor_data = data[pos + 8:pos + 8 + rec_len]
            values = []
            for i in range(0, len(anchor_data), 2):
                if i + 2 <= len(anchor_data):
                    values.append(struct.unpack_from("<h", anchor_data, i)[0])
            print(f"{prefix}    anchor values (signed): {values}")
        elif rec_type == RT_CLIENT_DATA and is_container:
            # ClientData 是 container，包含 PPT 特定的 record
            cd_end = pos + 8 + rec_len
            cd_pos = pos + 8
            cd_count = 0
            while cd_pos + 8 <= cd_end and cd_count < 20:
                h2 = parse_header(data, cd_pos)
                if h2 is None:
                    break
                ver2, inst2, rec_type2, rec_len2 = h2
                print(f"{prefix}    ClientData 子项: [{cd_pos:6d}] ver=0x{ver2:X} inst=0x{inst2:03X} type=0x{rec_type2:04X} len={rec_len2}")
                cd_pos += 8 + rec_len2
                cd_count += 1
                if h2[0] != 0xF and rec_len2 == 0:
                    break
        elif rec_type == RT_CLIENT_TEXTBOX and is_container:
            tb_end = pos + 8 + rec_len
            tb_pos = pos + 8
            while tb_pos + 8 <= tb_end:
                h2 = parse_header(data, tb_pos)
                if h2 is None:
                    break
                ver2, inst2, rec_type2, rec_len2 = h2
                tname2 = type_name(rec_type2)
                print(f"{prefix}    [{tb_pos:6d}] ver=0x{ver2:X} inst=0x{inst2:03X} type=0x{rec_type2:04X} ({tname2}) len={rec_len2}")
                if rec_type2 == RT_TEXT_CHARS_ATOM and rec_len2 > 0:
                    text_data = data[tb_pos + 8:tb_pos + 8 + rec_len2]
                    try:
                        text = text_data.decode("utf-16-le")
                        print(f"{prefix}      文本: {text!r}")
                    except Exception:
                        print(f"{prefix}      文本(原始): {text_data[:40]!r}")
                elif rec_type2 == RT_TEXT_HEADER_ATOM and rec_len2 >= 4:
                    tx_type = struct.unpack_from("<I", data, tb_pos + 8)[0]
                    print(f"{prefix}      txType: {tx_type}")
                tb_pos += 8 + rec_len2
                if h2[0] != 0xF and rec_len2 == 0:
                    break

        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break


def analyze_file(filepath, max_slides=3):
    print(f"\n{'='*80}")
    print(f"分析文件: {filepath}")
    print(f"{'='*80}")

    ole = olefile.OleFileIO(filepath)
    stream_names = ["/".join(p) for p in ole.listdir()]
    if "PowerPoint Document" not in stream_names:
        print("找不到 PowerPoint Document stream")
        ole.close()
        return
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()

    print(f"PowerPoint Document stream 大小: {len(ppt_data)} 字节")

    slides = find_all_slides(ppt_data)
    print(f"找到 {len(slides)} 个 Slide")

    # 分析前 max_slides 个 Slide
    for i, slide_offset in enumerate(slides[:max_slides]):
        print(f"\n--- Slide {i+1} (offset=0x{slide_offset:X}) ---")
        h = parse_header(ppt_data, slide_offset)
        slide_end = slide_offset + 8 + h[3]
        ppd_offset = find_record(ppt_data, slide_offset + 8, slide_end, RT_PPDRAWING)
        if ppd_offset is None:
            print("  找不到 PPDrawing")
            continue
        ppd_len = parse_header(ppt_data, ppd_offset)[3]
        ppd_end = ppd_offset + 8 + ppd_len
        # PPDrawing → DgContainer (0xF002) → SpgrContainer (0xF003)
        dg_offset = find_record(ppt_data, ppd_offset + 8, ppd_end, 0xF002)
        if dg_offset is None:
            print("  找不到 DgContainer")
            continue
        dg_len = parse_header(ppt_data, dg_offset)[3]
        dg_end = dg_offset + 8 + dg_len
        spgr_offset = find_record(ppt_data, dg_offset + 8, dg_end, RT_SPGR_CONTAINER)
        if spgr_offset is None:
            print("  找不到 SpgrContainer")
            continue

        spcontainers = find_all_spcontainers(ppt_data, spgr_offset)
        print(f"  SpContainer/SpgrContainer 数量: {len(spcontainers)}")

        # dump 每个 SpContainer
        for j, (sp_offset, sp_type) in enumerate(spcontainers):
            label = "(组形状本身)" if sp_type == RT_SPGR_CONTAINER else f"(形状 {j})"
            print(f"\n  >>> 第 {j+1} 个 {label}:")
            dump_spcontainer(ppt_data, sp_offset, 2, label)


def main():
    import os
    # 分析原始文件
    test_dir = "_test"
    if not os.path.isdir(test_dir):
        print(f"找不到 {test_dir} 目录")
        return
    for fname in sorted(os.listdir(test_dir)):
        if fname.lower().endswith(".ppt"):
            analyze_file(os.path.join(test_dir, fname), max_slides=2)


if __name__ == "__main__":
    main()
