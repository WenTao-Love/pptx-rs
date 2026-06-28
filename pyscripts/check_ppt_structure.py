# -*- coding: utf-8 -*-
"""检查 .ppt 加密文件的关键字段，诊断 PowerPoint 无法打开的原因。"""
import sys
import os
import struct
import olefile

TEST_DIR = os.path.join(os.path.dirname(__file__), "_test_out")
ENC_FILE = os.path.join(TEST_DIR, "protected_心理账户理论.ppt")
DEC_FILE = os.path.join(TEST_DIR, "decrypted_ppt.ppt")
ORIG_FILE = os.path.join(os.path.dirname(__file__), "_test", "心理账户理论.ppt")


def read_u32_le(data, offset):
    return struct.unpack_from("<I", data, offset)[0]


def read_u16_le(data, offset):
    return struct.unpack_from("<H", data, offset)[0]


def parse_record_header(data, offset):
    """返回 (ver, inst, recType, recLen)。"""
    ver_inst = read_u16_le(data, offset)
    rec_type = read_u16_le(data, offset + 2)
    rec_len = read_u32_le(data, offset + 4)
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset)] 列表。"""
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    print(f"  PersistDirectoryAtom: type=0x{rec_type:04X}, recLen={rec_len}")
    entries = []
    pos = offset + 8
    end = offset + 8 + rec_len
    while pos + 4 <= end:
        entry = read_u32_le(data, pos)
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        print(f"    entry: persistId={persist_id}, cPersist={c_persist}")
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= end:
                persist_offset = read_u32_le(data, pos)
                entries.append((persist_id + j, persist_offset))
                pos += 4
    return entries


def check_file(path, label):
    print(f"\n=== {label}: {os.path.basename(path)} ===")
    if not os.path.exists(path):
        print(f"  [FAIL] 文件不存在")
        return
    print(f"  文件大小: {os.path.getsize(path)}")

    ole = olefile.OleFileIO(path)
    streams = [s for s in ole.listdir()]
    print(f"  Streams: {streams}")

    # 读取 Current User
    cu = ole.openstream("Current User").read()
    ver, inst, cu_type, cu_len = parse_record_header(cu, 0)
    print(f"  CurrentUserAtom: type=0x{cu_type:04X}, recLen={cu_len}")
    header_token = read_u32_le(cu, 12)
    offset_to_current_edit = read_u32_le(cu, 16)
    print(f"    headerToken=0x{header_token:08X} ({'encrypted' if header_token == 0xF3D1C4DF else 'not encrypted'})")
    print(f"    offsetToCurrentEdit={offset_to_current_edit}")

    # 读取 PowerPoint Document
    ppt = ole.openstream("PowerPoint Document").read()
    print(f"  PowerPoint Document 大小: {len(ppt)}")

    # 解析 UserEditAtom
    ue_offset = offset_to_current_edit
    if ue_offset + 8 > len(ppt):
        print(f"  [FAIL] UserEditAtom offset 超出范围: {ue_offset} > {len(ppt)}")
        return
    ver, inst, ue_type, ue_len = parse_record_header(ppt, ue_offset)
    print(f"  UserEditAtom: type=0x{ue_type:04X}, recLen={ue_len} (offset={ue_offset})")
    if ue_type != 0x0FF5:
        print(f"  [FAIL] 不是 UserEditAtom")
        return

    last_slide_id_ref = read_u32_le(ppt, ue_offset + 8)
    offset_last_edit = read_u32_le(ppt, ue_offset + 16)
    offset_persist_dir = read_u32_le(ppt, ue_offset + 20)
    document_ref = read_u32_le(ppt, ue_offset + 24)
    persist_id_seed = read_u32_le(ppt, ue_offset + 28)
    print(f"    lastSlideIdRef={last_slide_id_ref}")
    print(f"    offsetLastEdit={offset_last_edit}")
    print(f"    offsetPersistDirectory={offset_persist_dir}")
    print(f"    documentRef={document_ref}")
    print(f"    persistIdSeed={persist_id_seed}")

    if ue_len >= 32:
        encrypt_session_pid_ref = read_u32_le(ppt, ue_offset + 36)
        print(f"    encryptSessionPersistIdRef={encrypt_session_pid_ref}")

    # 解析 PersistDirectoryAtom
    pd_offset = offset_persist_dir
    if pd_offset + 8 > len(ppt):
        print(f"  [FAIL] PersistDirectoryAtom offset 超出范围: {pd_offset} > {len(ppt)}")
        return
    entries = parse_persist_directory(ppt, pd_offset)
    print(f"  persist entries ({len(entries)}):")
    for pid, poff in entries[:10]:
        if poff + 8 <= len(ppt):
            _, _, rec_type, rec_len = parse_record_header(ppt, poff)
            print(f"    pid={pid}, offset={poff}, type=0x{rec_type:04X}, recLen={rec_len}")
        else:
            print(f"    pid={pid}, offset={poff} [超出范围!]")

    # 检查最后一个 persist 对象是否是 CryptSession10Container
    if entries:
        last_pid, last_offset = entries[-1]
        if last_offset + 8 <= len(ppt):
            _, _, last_type, last_len = parse_record_header(ppt, last_offset)
            print(f"  最后一个 persist 对象: pid={last_pid}, offset={last_offset}, type=0x{last_type:04X}")
            if last_type == 0x2F14:
                print(f"    [OK] 是 CryptSession10Container")
                # 检查 persistIdSeed 是否正确
                if persist_id_seed <= last_pid:
                    print(f"    [WARN] persistIdSeed ({persist_id_seed}) <= crypt_session_pid ({last_pid})，应该 > {last_pid}")

    ole.close()


def main():
    check_file(ORIG_FILE, "原始文件")
    check_file(ENC_FILE, "加密文件")
    check_file(DEC_FILE, "解密文件")


if __name__ == "__main__":
    main()
