#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
查看 .ppt 文件的 DocumentAtom，获取 slide 尺寸。
"""

import struct
import olefile


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def main():
    import os
    test_dir = "_test"
    for fname in sorted(os.listdir(test_dir)):
        if not fname.lower().endswith(".ppt"):
            continue
        filepath = os.path.join(test_dir, fname)
        print(f"\n分析文件: {filepath}")

        ole = olefile.OleFileIO(filepath)
        ppt_data = ole.openstream("PowerPoint Document").read()
        ole.close()

        # DocumentContainer = 0x03E8 (container)? 不对
        # DocumentAtom = 0x03E8? 让我查找
        # 根据 MS-PPT 规范:
        # RT_Document = 0x03E8 (container)
        # RT_DocumentAtom = 0x03E9 (atom)
        # DocumentAtom 结构:
        # - slideSize (8 bytes): cx(4) + cy(4)
        # - notesSize (8 bytes): cx(4) + cy(4)
        # - serverZoom (8 bytes)
        # - ...

        # 遍历顶层 record 查找 Document container (0x03E8)
        pos = 0
        while pos + 8 <= len(ppt_data):
            h = parse_header(ppt_data, pos)
            if h is None:
                break
            ver, inst, rec_type, rec_len = h
            is_container = ver == 0xF
            total_len = 8 + rec_len

            if is_container and rec_type == 0x03E8:
                # Document container
                doc_end = pos + 8 + rec_len
                dp = pos + 8
                print(f"  Document container offset=0x{pos:X}, len={rec_len}")
                while dp + 8 <= doc_end:
                    h2 = parse_header(ppt_data, dp)
                    if h2 is None:
                        break
                    ver2, inst2, rec_type2, rec_len2 = h2
                    print(f"    子 record: offset=0x{dp:X} ver=0x{ver2:X} type=0x{rec_type2:04X} len={rec_len2}")
                    if rec_type2 == 0x03E9 and rec_len2 >= 40:
                        # DocumentAtom
                        # slideSize: cx(4) + cy(4)
                        cx = struct.unpack_from("<I", ppt_data, dp + 8)[0]
                        cy = struct.unpack_from("<I", ppt_data, dp + 12)[0]
                        # notesSize
                        ncx = struct.unpack_from("<I", ppt_data, dp + 16)[0]
                        ncy = struct.unpack_from("<I", ppt_data, dp + 20)[0]
                        print(f"      DocumentAtom:")
                        print(f"        slideSize: cx={cx} cy={cy}")
                        print(f"        notesSize: cx={ncx} cy={ncy}")
                        print(f"        (1/100 mm: slide={cx/100}mm x {cy/100}mm)")
                        print(f"        (1/72 in: slide={cx/72:.2f}in x {cy/72:.2f}in)")
                        print(f"        (1/576 cm: slide={cx/576:.2f}cm x {cy/576:.2f}cm)")
                        # dump 原始数据
                        raw = ppt_data[dp+8:dp+8+rec_len2]
                        print(f"        原始数据: {raw.hex()}")
                    dp += 8 + rec_len2
                    if h2[0] != 0xF and rec_len2 == 0:
                        break
                break

            pos += total_len
            if not is_container and rec_len == 0:
                break


if __name__ == "__main__":
    main()
