#!/usr/bin/env python3
"""检查未加密水印文件的结构，确认水印是否正确注入到 MainMaster。"""
import struct
import olefile

def parse_record_header(data, offset):
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def find_records(data, rec_type_target, max_depth=3):
    """递归查找指定类型的 record。"""
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

import sys
path = sys.argv[1] if len(sys.argv) > 1 else "_test_out/wm_心理账户理论.ppt"
ole = olefile.OleFileIO(path)
ppt = ole.openstream("PowerPoint Document").read()

print(f"PowerPoint Document size: {len(ppt)}")

# 查找 MainMaster
main_masters = []
pos = 0
while pos + 8 <= len(ppt):
    ver, inst, rec_type, rec_len = parse_record_header(ppt, pos)
    total = 8 + rec_len
    if rec_type == 0x03F8:
        main_masters.append((pos, rec_len))
    pos += total

print(f"MainMaster count: {len(main_masters)}")

# 查找所有 SpContainer (0xF004)
sp_containers = find_records(ppt, 0xF004, max_depth=6)
print(f"SpContainer count: {len(sp_containers)}")

# 查找所有 FOPT (0xF00B)
fopts = find_records(ppt, 0xF00B)
print(f"FOPT count: {len(fopts)}")

# 查找包含 rotation 属性 (0x00BD) 的 FOPT，即我们的水印
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

ole.close()
