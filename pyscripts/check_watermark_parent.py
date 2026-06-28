#!/usr/bin/env python3
"""检查水印 SpContainer 的父级 record 类型，确认水印是否真的在 MainMaster 中。"""
import struct
import olefile

RT_MAIN_MASTER = 0x03F8
RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_SP_CONTAINER = 0xF004
RT_SPGR_CONTAINER = 0xF003
RT_DG_CONTAINER = 0xF002
RT_FOPT = 0xF00B
PROP_ROTATION = 0x00BD

def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len

def find_child(data, start, end, rec_type):
    pos = start
    while pos + 8 <= end:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rt, rec_len = hdr
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if rt == rec_type:
            return pos
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return None

def has_rotation_fopt(data, sp_offset):
    ver, inst, rt, rec_len = parse_record_header(data, sp_offset)
    pos = sp_offset + 8
    end = sp_offset + 8 + rec_len
    while pos + 8 <= end:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        cver, cinst, crt, clen = hdr
        if crt == RT_FOPT:
            for i in range(cinst):
                poff = pos + 8 + i * 6
                if poff + 6 > end:
                    break
                prop_id = struct.unpack_from("<H", data, poff)[0]
                prop_val = struct.unpack_from("<I", data, poff + 2)[0]
                if prop_id == PROP_ROTATION and prop_val == 0x002D0000:
                    return True
            return False
        pos += 8 + clen
    return False

def find_watermark_in_ppdrawing(data, ppd_offset):
    ver, inst, rt, rec_len = parse_record_header(data, ppd_offset)
    dg_offset = find_child(data, ppd_offset + 8, ppd_offset + 8 + rec_len, RT_DG_CONTAINER)
    if dg_offset is None:
        return []
    dg_len = parse_record_header(data, dg_offset)[3]
    spgr_offset = find_child(data, dg_offset + 8, dg_offset + 8 + dg_len, RT_SPGR_CONTAINER)
    if spgr_offset is None:
        return []
    spgr_len = parse_record_header(data, spgr_offset)[3]
    results = []
    pos = spgr_offset + 8
    end = spgr_offset + 8 + spgr_len
    while pos + 8 <= end:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rt, rec_len = hdr
        if rt == RT_SP_CONTAINER and has_rotation_fopt(data, pos):
            results.append(pos)
        pos += 8 + rec_len
    return results

def inspect(filepath):
    print(f"\n检查: {filepath}")
    ole = olefile.OleFileIO(filepath)
    ppt = ole.openstream("PowerPoint Document").read()
    ole.close()
    
    pos = 0
    masters = []
    slides = []
    while pos + 8 <= len(ppt):
        hdr = parse_record_header(ppt, pos)
        if hdr is None:
            break
        ver, inst, rt, rec_len = hdr
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if rt == RT_MAIN_MASTER:
            masters.append((pos, rec_len))
        elif rt == RT_SLIDE:
            slides.append((pos, rec_len))
        pos += total_len
        if not is_container and rec_len == 0:
            break
    
    print(f"  顶层 MainMaster 数量: {len(masters)}")
    print(f"  顶层 Slide 数量: {len(slides)}")
    
    all_watermarks = []
    for mpos, mlen in masters:
        ppd = find_child(ppt, mpos + 8, mpos + 8 + mlen, RT_PPDRAWING)
        if ppd is not None:
            for sp in find_watermark_in_ppdrawing(ppt, ppd):
                all_watermarks.append((sp, "MainMaster", mpos))
    for spos, slen in slides:
        ppd = find_child(ppt, spos + 8, spos + 8 + slen, RT_PPDRAWING)
        if ppd is not None:
            for sp in find_watermark_in_ppdrawing(ppt, ppd):
                all_watermarks.append((sp, "Slide", spos))
    
    print(f"  水印 SpContainer 数量: {len(all_watermarks)}")
    for sp_pos, parent_type, parent_pos in all_watermarks:
        print(f"    SpContainer @ {sp_pos}: 父级 {parent_type} @ {parent_pos}")

inspect("_test_out/wm_心理账户理论.ppt")
inspect("_test_out/wm_protected_心理账户理论.ppt.decrypted.ppt")
