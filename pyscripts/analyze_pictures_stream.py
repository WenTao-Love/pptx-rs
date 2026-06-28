#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析原始 .ppt 文件的 Pictures Stream 结构，以及加密后的 Pictures Stream。
"""

import struct
import olefile
import io


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def analyze_pictures(data, label=""):
    """分析 Pictures Stream 的结构。"""
    print(f"\n--- {label} Pictures Stream ({len(data)} bytes) ---")
    print(f"前 64 字节: {data[:64].hex()}")

    offset = 0
    count = 0
    while offset + 8 <= len(data) and count < 10:
        h = parse_header(data, offset)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        total_len = 8 + rec_len
        print(f"\n  Record {count} at offset {offset}:")
        print(f"    ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} len={rec_len}")
        print(f"    header: {data[offset:offset+8].hex()}")
        if rec_type == 0xF007:
            # FBSE
            print(f"    FBSE (File BLIP Store Entry)")
            if rec_len >= 36:
                # BLIB_STORE_ENTRY_PARTS = [1,1,16,2,4,4,4,1,1,1,1] = 36 bytes
                bt_win32 = data[offset+8]
                bt_macos = data[offset+9]
                rgb_uid = data[offset+10:offset+26]
                tag = struct.unpack_from("<H", data, offset+26)[0]
                size = struct.unpack_from("<I", data, offset+28)[0]
                cRef = struct.unpack_from("<I", data, offset+32)[0]
                foDelay = struct.unpack_from("<I", data, offset+36)[0]
                usage = data[offset+40]
                cbName = data[offset+41]
                unused2 = data[offset+42]
                unused3 = data[offset+43]
                print(f"      btWin32={bt_win32} btMacOS={bt_macos}")
                print(f"      rgbUid: {rgb_uid.hex()}")
                print(f"      tag={tag} size={size} cRef={cRef} foDelay={foDelay}")
                print(f"      usage={usage} cbName={cbName} unused2={unused2} unused3={unused3}")
                # cbName 在 offset+41 (1 byte)
                # 但 POI 读取 cbName 的方式不同：cbName = getUShort(pictstream, offset-3)
                # 其中 offset 是 parts 后的位置 = pos+36
                # 所以 cbName 在 pos+33 处读取 2 bytes
                cb_name_poi = struct.unpack_from("<H", data, offset+8+33)[0]
                print(f"      cbName (POI 读取方式, offset+8+33): {cb_name_poi}")
        elif 0xF01A <= rec_type <= 0xF01F:
            # Blip
            blip_names = {
                0xF01A: "EMF", 0xF01B: "WMF", 0xF01C: "PICT",
                0xF01D: "JPEG", 0xF01E: "PNG", 0xF01F: "DIB"
            }
            print(f"    Blip ({blip_names.get(rec_type, 'Unknown')})")
            # rgbUid (16 bytes × rgbUidCnt)
            rgb_uid_cnt = 2 if inst in [0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9] else 1
            print(f"      rgbUidCnt={rgb_uid_cnt} (inst=0x{inst:03X})")
            pos = offset + 8
            for i in range(rgb_uid_cnt):
                uid = data[pos:pos+16]
                print(f"      rgbUid[{i}]: {uid.hex()}")
                pos += 16
            # metafileHeader (34 bytes) or tag (1 byte)
            if rec_type in [0xF01A, 0xF01B, 0xF01C]:
                print(f"      metafileHeader (34 bytes): {data[pos:pos+34].hex()}")
                pos += 34
            else:
                print(f"      tag: {data[pos]:02X}")
                pos += 1
            # blipData (剩余)
            blip_data_len = rec_len - (pos - offset - 8)
            print(f"      blipData: {blip_data_len} bytes")
        else:
            print(f"    Unknown type")

        offset += total_len
        count += 1

    print(f"\n  总共 {count} 个 record (前 10 个)")


def main():
    import os

    # 1. 分析原始文件
    test_dir = "_test"
    for fname in sorted(os.listdir(test_dir)):
        if not fname.lower().endswith(".ppt"):
            continue
        filepath = os.path.join(test_dir, fname)
        print(f"\n{'='*80}")
        print(f"分析文件: {filepath}")
        print(f"{'='*80}")

        ole = olefile.OleFileIO(filepath)
        if ole.exists('Pictures'):
            pics = ole.openstream('Pictures').read()
            analyze_pictures(pics, "原始")
        ole.close()

    # 2. 分析加密文件（解密后）
    import msoffcrypto
    test_out = "_test_out"
    for fname in sorted(os.listdir(test_out)):
        if not fname.startswith("protected_") or not fname.endswith(".ppt"):
            continue
        filepath = os.path.join(test_out, fname)
        print(f"\n{'='*80}")
        print(f"分析加密文件: {filepath}")
        print(f"{'='*80}")

        with open(filepath, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password='pptx-rs-secret')
            out = io.BytesIO()
            office_file.decrypt(out)
            decrypted = out.getvalue()

        ole = olefile.OleFileIO(io.BytesIO(decrypted))
        if ole.exists('Pictures'):
            pics = ole.openstream('Pictures').read()
            analyze_pictures(pics, "解密后")
        ole.close()


if __name__ == "__main__":
    main()
