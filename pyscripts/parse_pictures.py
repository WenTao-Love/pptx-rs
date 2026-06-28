#!/usr/bin/env python3
"""解析原始 Pictures stream 的结构。"""
import struct
import olefile

def parse_pictures(path):
    ole = olefile.OleFileIO(path)
    data = ole.openstream("Pictures").read()
    ole.close()
    print(f"Pictures stream size: {len(data)}")

    offset = 0
    rec_num = 0
    while offset + 8 <= len(data) and rec_num < 20:
        ver_inst = struct.unpack_from("<H", data, offset)[0]
        rec_type = struct.unpack_from("<H", data, offset + 2)[0]
        rlen = struct.unpack_from("<I", data, offset + 4)[0]
        rec_inst = (ver_inst >> 4) & 0x0FFF
        rec_ver = ver_inst & 0xF
        print(f"  record @{offset}: type=0x{rec_type:04X}, inst=0x{rec_inst:03X}, ver={rec_ver}, len={rlen}")

        if rec_type == 0xF007:
            # FBSE
            pos = offset + 8
            bt_win32 = data[pos]
            bt_macos = data[pos+1]
            print(f"    FBSE: btWin32={bt_win32}, btMacOS={bt_macos}")
            # rgbUid at pos+2
            # tag at pos+18 (2 bytes)
            # size at pos+20 (4 bytes)
            # cRef at pos+24
            # foDelay at pos+28
            # unused1 at pos+32
            cb_name = struct.unpack_from("<H", data, pos + 33)[0]
            print(f"    cbName={cb_name}")
            # 36 bytes fixed + cbName
            embedded_start = pos + 36 + cb_name * 2
            if embedded_start < offset + 8 + rlen:
                print(f"    has embedded blip at {embedded_start}, remaining={offset+8+rlen-embedded_start}")
                e_ver_inst = struct.unpack_from("<H", data, embedded_start)[0]
                e_rec_type = struct.unpack_from("<H", data, embedded_start + 2)[0]
                e_rlen = struct.unpack_from("<I", data, embedded_start + 4)[0]
                e_inst = (e_ver_inst >> 4) & 0x0FFF
                print(f"    embedded: type=0x{e_rec_type:04X}, inst=0x{e_inst:03X}, len={e_rlen}")
        elif 0xF01A <= rec_type <= 0xF01F:
            print(f"    Blip type={hex(rec_type)}")

        offset += 8 + rlen
        rec_num += 1

parse_pictures("_test/心理账户理论.ppt")
