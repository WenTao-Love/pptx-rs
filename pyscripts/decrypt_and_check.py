#!/usr/bin/env python3
"""用 msoffcrypto 解密最新生成的 wm_protected 文件，并检查水印是否保留。"""
import io
import struct
import msoffcrypto
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
        total_len = 8 + rec_len
        if rt == rec_type:
            return pos
        pos += total_len
        if ver != 0xF and rec_len == 0:
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

def check_watermarks(ppt):
    pos = 0
    masters = []
    slides = []
    while pos + 8 <= len(ppt):
        hdr = parse_record_header(ppt, pos)
        if hdr is None:
            break
        ver, inst, rt, rec_len = hdr
        total_len = 8 + rec_len
        if rt == RT_MAIN_MASTER:
            masters.append((pos, rec_len))
        elif rt == RT_SLIDE:
            slides.append((pos, rec_len))
        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    
    print(f"  MainMaster: {len(masters)}, Slide: {len(slides)}")
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

def decrypt_and_check(filepath, password="pptx-rs-secret"):
    print(f"\n解密并检查: {filepath}")
    with open(filepath, "rb") as f:
        officefile = msoffcrypto.OfficeFile(f)
        officefile.load_key(password=password)
        decrypted = io.BytesIO()
        officefile.decrypt(decrypted)
        decrypted.seek(0)
        dec_data = decrypted.read()
    
    ole = olefile.OleFileIO(io.BytesIO(dec_data))
    ppt = ole.openstream("PowerPoint Document").read()
    ole.close()
    print(f"  解密后 PowerPoint Document 大小: {len(ppt)}")
    check_watermarks(ppt)

# 检查未加密的加水印文件
print("=" * 70)
with olefile.OleFileIO("_test_out/wm_心理账户理论.ppt") as ole:
    ppt = ole.openstream("PowerPoint Document").read()
print(f"未加密加水印文件 PowerPoint Document 大小: {len(ppt)}")
check_watermarks(ppt)

# 解密最新的水印+加密文件
print("=" * 70)
decrypt_and_check("_test_out/wm_protected_心理账户理论.ppt")

# 解密最新的纯加密文件（应无水印）
print("=" * 70)
decrypt_and_check("_test_out/protected_心理账户理论.ppt")
