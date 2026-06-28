#!/usr/bin/env python3
"""检查原始 .ppt 文件中形状的 FOPT 属性，确认 property ID 用法。"""
import struct
import olefile

RT_PPDRAWING = 0x040C
RT_DG_CONTAINER = 0xF002
RT_SPGR_CONTAINER = 0xF003
RT_SP_CONTAINER = 0xF004
RT_FOPT = 0xF00B
RT_FSP = = 0xF00A

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
        if rt == rec_type:
            return pos
        pos += 8 + rec_len
        if ver != 0xF and rec_len == 0:
            break
    return None

def dump_fopt(data, fopt_offset):
    ver, inst, rt, rec_len = parse_record_header(data, fopt_offset)
    prop_count = inst
    print(f"    FOPT @ {fopt_offset}: prop_count={prop_count}, len={rec_len}")
    for i in range(prop_count):
        poff = fopt_offset + 8 + i * 6
        if poff + 6 > fopt_offset + 8 + rec_len:
            break
        prop_id = struct.unpack_from("<H", data, poff)[0]
        prop_val = struct.unpack_from("<I", data, poff + 2)[0]
        print(f"      prop 0x{prop_id:04X} = 0x{prop_val:08X}")

def inspect_sp_container(data, sp_offset, indent=""):
    ver, inst, rt, rec_len = parse_record_header(data, sp_offset)
    pos = sp_offset + 8
    end = sp_offset + 8 + rec_len
    while pos + 8 <= end:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        cver, cinst, crt, clen = hdr
        if crt == RT_FOPT:
            dump_fopt(data, pos)
        pos += 8 + clen

def inspect_ppdrawing(data, ppd_offset):
    ver, inst, rt, rec_len = parse_record_header(data, ppd_offset)
    dg_offset = find_child(data, ppd_offset + 8, ppd_offset + 8 + rec_len, RT_DG_CONTAINER)
    if dg_offset is None:
        return
    dg_len = parse_record_header(data, dg_offset)[3]
    spgr_offset = find_child(data, dg_offset + 8, dg_offset + 8 + dg_len, RT_SPGR_CONTAINER)
    if spgr_offset is None:
        return
    spgr_len = parse_record_header(data, spgr_offset)[3]
    pos = spgr_offset + 8
    end = spgr_offset + 8 + spgr_len
    count = 0
    while pos + 8 <= end and count < 3:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        cver, cinst, crt, clen = hdr
        if crt == RT_SP_CONTAINER:
            print(f"  SpContainer @ {pos}")
            inspect_sp_container(data, pos)
            count += 1
        pos += 8 + clen

with olefile.OleFileIO("_test/心理账户理论.ppt") as ole:
    ppt = ole.openstream("PowerPoint Document").read()

pos = 0
master_count = 0
slide_count = 0
while pos + 8 <= len(ppt):
    hdr = parse_record_header(ppt, pos)
    if hdr is None:
        break
    ver, inst, rt, rec_len = hdr
    total_len = 8 + rec_len
    if rt == 0x03F8 and master_count < 1:
        print(f"\nMainMaster @ {pos}")
        ppd = find_child(ppt, pos + 8, pos + 8 + rec_len, RT_PPDRAWING)
        if ppd:
            inspect_ppdrawing(ppt, ppd)
        master_count += 1
    elif rt == 0x03EE and slide_count < 1:
        print(f"\nSlide @ {pos}")
        ppd = find_child(ppt, pos + 8, pos + 8 + rec_len, RT_PPDRAWING)
        if ppd:
            inspect_ppdrawing(ppt, ppd)
        slide_count += 1
    pos += total_len
    if ver != 0xF and rec_len == 0:
        break
