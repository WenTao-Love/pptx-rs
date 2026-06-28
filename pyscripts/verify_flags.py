#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证加密 .ppt 文件的 flags 值是否正确（应为 0x00000001）。
"""

import sys
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


def verify_file(filepath):
    print(f"\n=== 验证 {filepath} ===")
    ole = olefile.OleFileIO(filepath)
    stream_names = ["/".join(p) for p in ole.listdir()]
    if "PowerPoint Document" not in stream_names:
        print("  找不到 PowerPoint Document stream")
        ole.close()
        return
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()

    # 找到 CurrentUserAtom
    cu_data = olefile.OleFileIO(filepath).openstream("Current User").read()
    h = parse_header(cu_data, 0)
    if h is None or h[2] != 0x0FF6:
        print("  CurrentUserAtom 解析失败")
        return
    ue_offset = struct.unpack_from("<I", cu_data, 16)[0]
    print(f"  UserEditAtom offset: {ue_offset}")

    # 解析 UserEditAtom
    h = parse_header(ppt_data, ue_offset)
    if h is None or h[2] != 0x0FF5:
        print("  UserEditAtom 解析失败")
        return
    ue_len = h[3]
    print(f"  UserEditAtom len: {ue_len} (encrypted={ue_len == 32})")

    # 获取 encryptSessionPersistIdRef
    if ue_len >= 32:
        enc_session_pid = struct.unpack_from("<I", ppt_data, ue_offset + 8 + 28)[0]
        print(f"  encryptSessionPersistIdRef: {enc_session_pid}")
    else:
        print("  文件未加密")
        return

    # 获取 offsetPersistDirectory
    pd_offset = struct.unpack_from("<I", ppt_data, ue_offset + 8 + 12)[0]
    print(f"  PersistDirectoryAtom offset: {pd_offset}")

    # 解析 PersistDirectoryAtom，找到 encryptSession 的 offset
    h = parse_header(ppt_data, pd_offset)
    if h is None or h[2] != 0x1772:
        print("  PersistDirectoryAtom 解析失败")
        return
    pd_data = ppt_data[pd_offset + 8:pd_offset + 8 + h[3]]

    persist_offsets = {}
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                offset = struct.unpack_from("<I", pd_data, pos)[0]
                persist_offsets[persist_id + j] = offset
                pos += 4

    if enc_session_pid not in persist_offsets:
        print(f"  找不到 persistId={enc_session_pid} 的 offset")
        return

    cs_offset = persist_offsets[enc_session_pid]
    print(f"  CryptSession10Container offset: {cs_offset}")

    # 解析 CryptSession10Container
    h = parse_header(ppt_data, cs_offset)
    if h is None or h[2] != 0x2F14:
        print(f"  CryptSession10Container 解析失败 (type=0x{h[2]:04X})")
        return

    cs_data = ppt_data[cs_offset + 8:cs_offset + 8 + h[3]]
    # EncryptionVersionInfo
    v_major = struct.unpack_from("<H", cs_data, 0)[0]
    v_minor = struct.unpack_from("<H", cs_data, 2)[0]
    print(f"  EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}")

    # 外层 flags
    outer_flags = struct.unpack_from("<I", cs_data, 4)[0]
    f_crypto_api = (outer_flags >> 0) & 1
    f_doc_props = (outer_flags >> 1) & 1
    f_external = (outer_flags >> 2) & 1
    f_aes = (outer_flags >> 3) & 1
    f_agile = (outer_flags >> 4) & 1
    print(f"  外层 flags: 0x{outer_flags:08X}")
    print(f"    fCryptoAPI={f_crypto_api}, fDocProps={f_doc_props}, fExternal={f_external}, fAES={f_aes}, fAgile={f_agile}")

    # headerSize
    header_size = struct.unpack_from("<I", cs_data, 8)[0]
    print(f"  headerSize: {header_size}")

    # EncryptionHeader.flags
    eh_flags = struct.unpack_from("<I", cs_data, 12)[0]
    eh_f_crypto_api = (eh_flags >> 0) & 1
    eh_f_external = (eh_flags >> 2) & 1
    print(f"  EncryptionHeader.flags: 0x{eh_flags:08X}")
    print(f"    fCryptoAPI={eh_f_crypto_api}, fExternal={eh_f_external}")

    # 验证
    if outer_flags == 0x00000001 and eh_flags == 0x00000001:
        print("  ✓ flags 值正确（fCryptoAPI=1）")
    else:
        print("  ✗ flags 值错误！应为 0x00000001")
        print(f"    外层 flags: 0x{outer_flags:08X} (期望 0x00000001)")
        print(f"    EH flags:   0x{eh_flags:08X} (期望 0x00000001)")


def main():
    import os
    test_dir = "_test_out"
    for fname in sorted(os.listdir(test_dir)):
        if fname.lower().endswith(".ppt") and ("protected" in fname.lower()):
            verify_file(os.path.join(test_dir, fname))


if __name__ == "__main__":
    main()
