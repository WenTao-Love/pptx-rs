#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 msoffcrypto 解密加密文件，检查解密后的 PowerPoint Document stream 的 record 结构。
比较解密后的结构和原始结构，找出差异。
"""

import struct
import olefile
import io
import msoffcrypto


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def dump_top_records(data, label, max_records=20):
    """dump 顶层 record 结构。"""
    print(f"\n--- {label} ({len(data)} bytes) ---")
    pos = 0
    count = 0
    while pos + 8 <= len(data) and count < max_records:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        c = "C" if is_container else "A"
        print(f"  [{pos:6d}] {c} ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} len={rec_len} (total={total_len})")
        pos += total_len
        count += 1
        if not is_container and rec_len == 0:
            break
    print(f"  ... 共 {count} 个 record (前 {max_records} 个)")


def check_user_edit_atom(data, label):
    """检查 UserEditAtom 和 PersistDirectoryAtom。"""
    print(f"\n--- {label} UserEditAtom/PersistDirectoryAtom ---")
    # CurrentUserAtom 在 Current User stream 中
    # 这里只检查 PowerPoint Document stream 末尾的 UserEditAtom 和 PersistDirectoryAtom

    # 从末尾往前找 UserEditAtom (0x0FF5) 和 PersistDirectoryAtom (0x1772)
    # 遍历所有顶层 record
    pos = 0
    ue_offset = None
    pd_offset = None
    cs_offset = None
    while pos + 8 <= len(data):
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len

        if rec_type == 0x0FF5:  # UserEditAtom
            ue_offset = pos
            print(f"  UserEditAtom at {pos}: len={rec_len}")
            if rec_len >= 28:
                lastSlideID = struct.unpack_from("<I", data, pos + 8)[0]
                version = struct.unpack_from("<I", data, pos + 12)[0]
                minorVersion = struct.unpack_from("<H", data, pos + 16)[0]
                majorVersion = struct.unpack_from("<H", data, pos + 18)[0]
                offsetPersistDirectory = struct.unpack_from("<I", data, pos + 20)[0]
                persistIdSeed = struct.unpack_from("<I", data, pos + 24)[0]
                print(f"    lastSlideId={lastSlideID} version={version}")
                print(f"    major={majorVersion} minor={minorVersion}")
                print(f"    offsetPersistDirectory={offsetPersistDirectory}")
                print(f"    persistIdSeed={persistIdSeed}")
                if rec_len >= 32:
                    encryptSessionPersistIdRef = struct.unpack_from("<I", data, pos + 32)[0]
                    print(f"    encryptSessionPersistIdRef={encryptSessionPersistIdRef}")
                else:
                    print(f"    encryptSessionPersistIdRef: 无（未加密）")

        if rec_type == 0x1772:  # PersistDirectoryAtom
            pd_offset = pos
            print(f"  PersistDirectoryAtom at {pos}: len={rec_len}")
            # 解析 persist entries
            pd_data = data[pos+8:pos+8+rec_len]
            p = 0
            while p + 4 <= len(pd_data):
                entry = struct.unpack_from("<I", pd_data, p)[0]
                persist_id = entry & 0xFFFFF
                c_persist = (entry >> 20) & 0xFFF
                p += 4
                print(f"    persistId={persist_id} cPersist={c_persist}")
                for j in range(c_persist):
                    if p + 4 <= len(pd_data):
                        po = struct.unpack_from("<I", pd_data, p)[0]
                        print(f"      [{persist_id+j}] offset={po}")
                        p += 4

        if rec_type == 0x2F14:  # CryptSession10Container
            cs_offset = pos
            print(f"  CryptSession10Container at {pos}: len={rec_len}")

        pos += total_len
        if not is_container and rec_len == 0:
            break


def main():
    # 1. 读取原始文件
    orig_path = "_test/心理账户理论.ppt"
    ole_orig = olefile.OleFileIO(orig_path)
    orig_ppt = ole_orig.openstream("PowerPoint Document").read()
    ole_orig.close()

    dump_top_records(orig_ppt, "原始 PowerPoint Document")
    check_user_edit_atom(orig_ppt, "原始")

    # 2. 解密加密文件
    enc_path = "_test_out/protected_心理账户理论.ppt"
    with open(enc_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password='pptx-rs-secret')
        out = io.BytesIO()
        office_file.decrypt(out)
        decrypted = out.getvalue()

    ole_dec = olefile.OleFileIO(io.BytesIO(decrypted))
    dec_ppt = ole_dec.openstream("PowerPoint Document").read()
    ole_dec.close()

    dump_top_records(dec_ppt, "解密后 PowerPoint Document")
    check_user_edit_atom(dec_ppt, "解密后")


if __name__ == "__main__":
    main()
