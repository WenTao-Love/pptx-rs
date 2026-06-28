#!/usr/bin/env python3
"""检查解密后的水印+加密文件，确认水印结构完整。"""
import struct
import olefile
import io
import msoffcrypto

def parse_record_header(data, offset):
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def find_records(data, rec_type_target, max_depth=6):
    results = []
    def search(offset, end, depth):
        if depth > max_depth:
            return
        pos = offset
        while pos + 8 <= end:
            ver, inst, rec_type, rec_len = parse_record_header(data, pos)
            total = 8 + rec_len
            if rec_type == rec_type_target:
                results.append((pos, ver, inst, rec_len))
            elif ver == 0xF and rec_len > 0:
                search(pos + 8, pos + 8 + rec_len, depth + 1)
            pos += total
    search(0, len(data), 0)
    return results

with open("_test_out/wm_protected_心理账户理论.ppt", "rb") as f:
    officefile = msoffcrypto.OfficeFile(f)
    officefile.load_key(password="pptx-rs-secret")
    decrypted = io.BytesIO()
    officefile.decrypt(decrypted)
    decrypted.seek(0)
    ole = olefile.OleFileIO(decrypted)
    ppt = ole.openstream("PowerPoint Document").read()
    ole.close()

print(f"Decrypted PowerPoint Document size: {len(ppt)}")

sp_containers = find_records(ppt, 0xF004)
print(f"SpContainer count: {len(sp_containers)}")

print("\n查找水印 SpContainer (FOPT 包含 0x00BD rotation):")
for sp_off, sp_ver, sp_inst, sp_len in sp_containers:
    pos = sp_off + 8
    end = sp_off + 8 + sp_len
    while pos + 8 <= end:
        ver, inst, rec_type, rec_len = parse_record_header(ppt, pos)
        if rec_type == 0xF00B:
            for i in range(inst):
                prop_off = pos + 8 + i * 6
                prop_id = struct.unpack_from("<H", ppt, prop_off)[0]
                if prop_id == 0x00BD:
                    print(f"  水印 SpContainer @ {sp_off}: len={sp_len}")
                    print(f"    FOPT @ {pos}: inst={inst}, len={rec_len}")
                    for j in range(inst):
                        p_off = pos + 8 + j * 6
                        p_id = struct.unpack_from("<H", ppt, p_off)[0]
                        p_val = struct.unpack_from("<I", ppt, p_off + 2)[0]
                        print(f"      prop 0x{p_id:04X} = 0x{p_val:08X}")
                    break
        pos += 8 + rec_len
