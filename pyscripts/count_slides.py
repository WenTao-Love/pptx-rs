#!/usr/bin/env python3
"""统计原始文件中的 Slide 数量。"""
import struct
import olefile

def parse_record_header(data, offset):
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0xF
    return ver, rec_type, rec_len

ole = olefile.OleFileIO("_test/心理账户理论.ppt")
ppt = ole.openstream("PowerPoint Document").read()

slide_count = 0
main_master_count = 0
pos = 0
while pos + 8 <= len(ppt):
    ver, rec_type, rec_len = parse_record_header(ppt, pos)
    if rec_type == 0x03EE:
        slide_count += 1
    elif rec_type == 0x03F8:
        main_master_count += 1
    pos += 8 + rec_len

print(f"Slide count: {slide_count}")
print(f"MainMaster count: {main_master_count}")
ole.close()
