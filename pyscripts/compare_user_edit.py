#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对比原始文件和加密文件的 UserEditAtom 和 CurrentUserAtom 字段。
"""

import struct
import olefile


def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def dump_user_edit_atom(ppt_data, label):
    """Dump UserEditAtom fields."""
    print(f"\n{label}:")

    # 从 Current User 读取 offsetToCurrentEdit
    # (需要从外部传入，这里直接搜索 UserEditAtom)
    # UserEditAtom type = 0x0FF5

    # 搜索所有 UserEditAtom
    offset = 0
    while offset + 8 <= len(ppt_data):
        h = parse_record_header(ppt_data, offset)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        if rec_type == 0x0FF5:
            print(f"  UserEditAtom at offset {offset}:")
            print(f"    RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")
            if rec_len >= 28:
                lastSlideIdRef = struct.unpack_from("<I", ppt_data, offset + 8)[0]
                version = struct.unpack_from("<H", ppt_data, offset + 12)[0]
                minorVersion = ppt_data[offset + 14]
                majorVersion = ppt_data[offset + 15]
                offsetLastEdit = struct.unpack_from("<I", ppt_data, offset + 16)[0]
                offsetPersistDirectory = struct.unpack_from("<I", ppt_data, offset + 20)[0]
                documentRef = struct.unpack_from("<I", ppt_data, offset + 24)[0]
                maxPersistWritten = struct.unpack_from("<I", ppt_data, offset + 28)[0]
                lastViewType = struct.unpack_from("<H", ppt_data, offset + 32)[0] if rec_len >= 30 else None
                unused = struct.unpack_from("<H", ppt_data, offset + 34)[0] if rec_len >= 32 else None
                encryptSessionPersistIdRef = struct.unpack_from("<I", ppt_data, offset + 36)[0] if rec_len >= 36 else None

                print(f"    lastSlideIdRef: 0x{lastSlideIdRef:08X}")
                print(f"    version: 0x{version:04X}")
                print(f"    minorVersion: {minorVersion}")
                print(f"    majorVersion: {majorVersion}")
                print(f"    offsetLastEdit: {offsetLastEdit}")
                print(f"    offsetPersistDirectory: {offsetPersistDirectory}")
                print(f"    documentRef: {documentRef}")
                print(f"    maxPersistWritten: {maxPersistWritten}")
                if lastViewType is not None:
                    print(f"    lastViewType: {lastViewType}")
                if unused is not None:
                    print(f"    unused: {unused}")
                if encryptSessionPersistIdRef is not None:
                    print(f"    encryptSessionPersistIdRef: {encryptSessionPersistIdRef}")
            offset += 8 + rec_len
        else:
            # 跳过这个 record
            offset += 8 + rec_len
            # 如果是 container 类型，需要特殊处理
            # 简化：直接跳过


def dump_current_user(cu_data, label):
    """Dump CurrentUserAtom fields."""
    print(f"\n{label}:")
    h = parse_record_header(cu_data, 0)
    if h is None:
        print("  无法解析 header")
        return
    ver, inst, rec_type, rec_len = h
    print(f"  RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")

    if rec_type != 0x0FF6:
        print(f"  不是 CurrentUserAtom!")
        return

    size = struct.unpack_from("<I", cu_data, 8)[0]
    headerToken = struct.unpack_from("<I", cu_data, 12)[0]
    offsetToCurrentEdit = struct.unpack_from("<I", cu_data, 16)[0]
    print(f"  size: {size}")
    print(f"  headerToken: 0x{headerToken:08X} ({'已加密' if headerToken == 0xF3D1C4DF else '未加密'})")
    print(f"  offsetToCurrentEdit: {offsetToCurrentEdit}")

    if size >= 20:
        docFileVersion = struct.unpack_from("<H", cu_data, 20)[0]
        majorVersion = cu_data[22]
        minorVersion = cu_data[23]
        unused = struct.unpack_from("<I", cu_data, 24)[0]
        ansiUserNameLength = struct.unpack_from("<H", cu_data, 28)[0]
        print(f"  docFileVersion: 0x{docFileVersion:04X}")
        print(f"  majorVersion: {majorVersion}")
        print(f"  minorVersion: {minorVersion}")
        print(f"  unused: 0x{unused:08X}")
        print(f"  ansiUserNameLength: {ansiUserNameLength}")


def main():
    orig_path = "_test/心理账户理论.ppt"
    enc_path = "_test_out/protected_心理账户理论.ppt"

    # 原始文件
    ole = olefile.OleFileIO(orig_path)
    orig_ppt = ole.openstream("PowerPoint Document").read()
    orig_cu = ole.openstream("Current User").read()
    ole.close()

    # 加密文件
    ole = olefile.OleFileIO(enc_path)
    enc_ppt = ole.openstream("PowerPoint Document").read()
    enc_cu = ole.openstream("Current User").read()
    ole.close()

    print(f"原始文件 PowerPoint Document 大小: {len(orig_ppt)}")
    print(f"加密文件 PowerPoint Document 大小: {len(enc_ppt)}")
    print(f"差异: {len(enc_ppt) - len(orig_ppt)}")

    dump_current_user(orig_cu, "原始文件 CurrentUserAtom")
    dump_current_user(enc_cu, "加密文件 CurrentUserAtom")

    # 从 Current User 读取 offsetToCurrentEdit
    orig_offset = struct.unpack_from("<I", orig_cu, 16)[0]
    enc_offset = struct.unpack_from("<I", enc_cu, 16)[0]

    print(f"\n原始文件 offsetToCurrentEdit: {orig_offset}")
    print(f"加密文件 offsetToCurrentEdit: {enc_offset}")

    # 直接读取 UserEditAtom
    print(f"\n--- 原始文件 UserEditAtom (offset={orig_offset}) ---")
    h = parse_record_header(orig_ppt, orig_offset)
    if h:
        ver, inst, rec_type, rec_len = h
        print(f"  RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")
        if rec_type == 0x0FF5 and rec_len >= 28:
            print(f"  lastSlideIdRef: 0x{struct.unpack_from('<I', orig_ppt, orig_offset+8)[0]:08X}")
            print(f"  version: 0x{struct.unpack_from('<H', orig_ppt, orig_offset+12)[0]:04X}")
            print(f"  minorVersion: {orig_ppt[orig_offset+14]}")
            print(f"  majorVersion: {orig_ppt[orig_offset+15]}")
            print(f"  offsetLastEdit: {struct.unpack_from('<I', orig_ppt, orig_offset+16)[0]}")
            print(f"  offsetPersistDirectory: {struct.unpack_from('<I', orig_ppt, orig_offset+20)[0]}")
            print(f"  documentRef: {struct.unpack_from('<I', orig_ppt, orig_offset+24)[0]}")
            print(f"  maxPersistWritten: {struct.unpack_from('<I', orig_ppt, orig_offset+28)[0]}")
            if rec_len >= 30:
                print(f"  lastViewType: {struct.unpack_from('<H', orig_ppt, orig_offset+32)[0]}")
            if rec_len >= 32:
                print(f"  unused: {struct.unpack_from('<H', orig_ppt, orig_offset+34)[0]}")
            if rec_len >= 36:
                print(f"  encryptSessionPersistIdRef: {struct.unpack_from('<I', orig_ppt, orig_offset+36)[0]}")

    print(f"\n--- 加密文件 UserEditAtom (offset={enc_offset}) ---")
    h = parse_record_header(enc_ppt, enc_offset)
    if h:
        ver, inst, rec_type, rec_len = h
        print(f"  RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")
        if rec_type == 0x0FF5 and rec_len >= 28:
            print(f"  lastSlideIdRef: 0x{struct.unpack_from('<I', enc_ppt, enc_offset+8)[0]:08X}")
            print(f"  version: 0x{struct.unpack_from('<H', enc_ppt, enc_offset+12)[0]:04X}")
            print(f"  minorVersion: {enc_ppt[enc_offset+14]}")
            print(f"  majorVersion: {enc_ppt[enc_offset+15]}")
            print(f"  offsetLastEdit: {struct.unpack_from('<I', enc_ppt, enc_offset+16)[0]}")
            print(f"  offsetPersistDirectory: {struct.unpack_from('<I', enc_ppt, enc_offset+20)[0]}")
            print(f"  documentRef: {struct.unpack_from('<I', enc_ppt, enc_offset+24)[0]}")
            print(f"  maxPersistWritten: {struct.unpack_from('<I', enc_ppt, enc_offset+28)[0]}")
            if rec_len >= 30:
                print(f"  lastViewType: {struct.unpack_from('<H', enc_ppt, enc_offset+32)[0]}")
            if rec_len >= 32:
                print(f"  unused: {struct.unpack_from('<H', enc_ppt, enc_offset+34)[0]}")
            if rec_len >= 36:
                print(f"  encryptSessionPersistIdRef: {struct.unpack_from('<I', enc_ppt, enc_offset+36)[0]}")

    # 检查 PersistDirectoryAtom
    print(f"\n--- 原始文件 PersistDirectoryAtom ---")
    orig_pd_offset = struct.unpack_from("<I", orig_ppt, orig_offset + 20)[0]
    h = parse_record_header(orig_ppt, orig_pd_offset)
    if h:
        ver, inst, rec_type, rec_len = h
        print(f"  offset: {orig_pd_offset}")
        print(f"  RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")
        # 解析 persist entries
        pd_data = orig_ppt[orig_pd_offset+8:orig_pd_offset+8+rec_len]
        pos = 0
        while pos + 4 <= len(pd_data):
            entry = struct.unpack_from("<I", pd_data, pos)[0]
            pid = entry & 0xFFFFF
            cpersist = (entry >> 20) & 0xFFF
            pos += 4
            print(f"  persistId={pid} cPersist={cpersist}")
            for j in range(cpersist):
                if pos + 4 <= len(pd_data):
                    po = struct.unpack_from("<I", pd_data, pos)[0]
                    print(f"    [{pid+j}] offset={po}")
                    pos += 4

    print(f"\n--- 加密文件 PersistDirectoryAtom ---")
    enc_pd_offset = struct.unpack_from("<I", enc_ppt, enc_offset + 20)[0]
    h = parse_record_header(enc_ppt, enc_pd_offset)
    if h:
        ver, inst, rec_type, rec_len = h
        print(f"  offset: {enc_pd_offset}")
        print(f"  RecordHeader: ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} recLen={rec_len}")
        # 解析 persist entries
        pd_data = enc_ppt[enc_pd_offset+8:enc_pd_offset+8+rec_len]
        pos = 0
        while pos + 4 <= len(pd_data):
            entry = struct.unpack_from("<I", pd_data, pos)[0]
            pid = entry & 0xFFFFF
            cpersist = (entry >> 20) & 0xFFF
            pos += 4
            print(f"  persistId={pid} cPersist={cpersist}")
            for j in range(cpersist):
                if pos + 4 <= len(pd_data):
                    po = struct.unpack_from("<I", pd_data, pos)[0]
                    print(f"    [{pid+j}] offset={po}")
                    pos += 4


if __name__ == "__main__":
    main()
