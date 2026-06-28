#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析原始 .ppt 文件中 ClientData (0xF011) 的结构，特别是 PPDrawingBoard (0x1388)。
"""

import sys
import struct
import olefile

RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_SPGR_CONTAINER = 0xF003
RT_SP_CONTAINER = 0xF004
RT_CLIENT_DATA = 0xF011
RT_FSP = 0xF00A
RT_PPDRAWING_BOARD = 0x1388


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


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


def find_spcontainers_with_client_data(data, spgr_offset):
    """在 SpgrContainer 中找到所有带 ClientData 的 SpContainer。"""
    results = []
    h = parse_header(data, spgr_offset)
    if h is None:
        return results
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
        if is_container and rec_type == RT_SP_CONTAINER:
            # 在 SpContainer 中查找 ClientData
            sp_end = pos + 8 + rec_len
            cd_offset = find_record(data, pos + 8, sp_end, RT_CLIENT_DATA)
            if cd_offset is not None:
                results.append((pos, cd_offset))
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return results


def dump_client_data(data, cd_offset, indent=0):
    """dump ClientData 的内容。"""
    prefix = "  " * indent
    h = parse_header(data, cd_offset)
    if h is None:
        return
    ver, inst, rec_type, rec_len = h
    print(f"{prefix}ClientData (offset=0x{cd_offset:X}, ver=0x{ver:X}, inst=0x{inst:03X}, len={rec_len})")

    cd_end = cd_offset + 8 + rec_len
    pos = cd_offset + 8
    while pos + 8 <= cd_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        print(f"{prefix}  [{pos:6d}] ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} len={rec_len}")

        if rec_type == RT_PPDRAWING_BOARD and is_container:
            # PPDrawingBoard container
            pb_end = pos + 8 + rec_len
            pb_pos = pos + 8
            while pb_pos + 8 <= pb_end:
                h2 = parse_header(data, pb_pos)
                if h2 is None:
                    break
                ver2, inst2, rec_type2, rec_len2 = h2
                print(f"{prefix}    [{pb_pos:6d}] ver=0x{ver2:X} inst=0x{inst2:03X} type=0x{rec_type2:04X} len={rec_len2}")

                if rec_type2 == 0x1389 and rec_len2 >= 28:
                    # PPDrawingBoardAtom
                    raw = data[pb_pos + 8:pb_pos + 8 + rec_len2]
                    print(f"{prefix}      原始数据 ({len(raw)} bytes): {raw.hex()}")
                    # 解析字段
                    if len(raw) >= 28:
                        unused1 = struct.unpack_from("<I", raw, 0)[0]
                        unused2 = struct.unpack_from("<I", raw, 4)[0]
                        unused3 = struct.unpack_from("<I", raw, 8)[0]
                        rect_top, rect_left, rect_right, rect_bottom = struct.unpack_from("<hhhh", raw, 12)
                        clip_id = struct.unpack_from("<H", raw, 20)[0]
                        placeholder_id = struct.unpack_from("<H", raw, 22)[0]
                        recolor_id = struct.unpack_from("<H", raw, 24)[0]
                        unused4 = struct.unpack_from("<H", raw, 26)[0]
                        print(f"{prefix}      unused1={unused1} unused2={unused2} unused3={unused3}")
                        print(f"{prefix}      rect: top={rect_top} left={rect_left} right={rect_right} bottom={rect_bottom}")
                        print(f"{prefix}      clipId={clip_id} placeholderId={placeholder_id} recolorId={recolor_id} unused4={unused4}")
                else:
                    raw = data[pb_pos + 8:pb_pos + 8 + rec_len2]
                    print(f"{prefix}      原始数据 ({len(raw)} bytes): {raw.hex()}")

                pb_pos += 8 + rec_len2
                if h2[0] != 0xF and rec_len2 == 0:
                    break
        else:
            raw = data[pos + 8:pos + 8 + rec_len]
            if len(raw) <= 64:
                print(f"{prefix}      原始数据: {raw.hex()}")
            else:
                print(f"{prefix}      原始数据 (前64字节): {raw[:64].hex()}...")

        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break


def main():
    import os
    test_dir = "_test"
    for fname in sorted(os.listdir(test_dir)):
        if not fname.lower().endswith(".ppt"):
            continue
        filepath = os.path.join(test_dir, fname)
        print(f"\n{'='*80}")
        print(f"分析文件: {filepath}")
        print(f"{'='*80}")

        ole = olefile.OleFileIO(filepath)
        ppt_data = ole.openstream("PowerPoint Document").read()
        ole.close()

        slides = find_all_slides(ppt_data)
        print(f"找到 {len(slides)} 个 Slide")

        # 分析前 2 个 Slide 中的 ClientData
        found = 0
        for i, slide_offset in enumerate(slides[:5]):
            if found >= 3:
                break
            h = parse_header(ppt_data, slide_offset)
            slide_end = slide_offset + 8 + h[3]
            ppd_offset = find_record(ppt_data, slide_offset + 8, slide_end, 0x040C)
            if ppd_offset is None:
                continue
            ppd_len = parse_header(ppt_data, ppd_offset)[3]
            ppd_end = ppd_offset + 8 + ppd_len
            dg_offset = find_record(ppt_data, ppd_offset + 8, ppd_end, 0xF002)
            if dg_offset is None:
                continue
            dg_len = parse_header(ppt_data, dg_offset)[3]
            dg_end = dg_offset + 8 + dg_len
            spgr_offset = find_record(ppt_data, dg_offset + 8, dg_end, RT_SPGR_CONTAINER)
            if spgr_offset is None:
                continue

            results = find_spcontainers_with_client_data(ppt_data, spgr_offset)
            for sp_offset, cd_offset in results:
                print(f"\n--- Slide {i+1}, SpContainer at 0x{sp_offset:X} ---")
                dump_client_data(ppt_data, cd_offset, 1)
                found += 1
                if found >= 3:
                    break


if __name__ == "__main__":
    main()
