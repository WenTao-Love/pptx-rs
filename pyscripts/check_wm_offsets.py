# -*- coding: utf-8 -*-
"""检查水印 .ppt 文件的 offset 更新是否正确。"""
import sys
import os
import struct
import olefile

TEST_DIR = os.path.join(os.path.dirname(__file__), "_test_out")
WM_FILE = os.path.join(TEST_DIR, "wm_心理账户理论.ppt")
ORIG_FILE = os.path.join(os.path.dirname(__file__), "_test", "心理账户理论.ppt")


def read_u32_le(data, offset):
    return struct.unpack_from("<I", data, offset)[0]


def parse_record_header(data, offset):
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    return ver, rec_type, rec_len


def parse_persist_directory(data, offset):
    ver, rec_type, rec_len = parse_record_header(data, offset)
    entries = []
    pos = offset + 8
    end = offset + 8 + rec_len
    while pos + 4 <= end:
        entry = read_u32_le(data, pos)
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
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

    ole = olefile.OleFileIO(path)
    cu = ole.openstream("Current User").read()
    ppt = ole.openstream("PowerPoint Document").read()

    offset_to_current_edit = read_u32_le(cu, 16)
    print(f"  offsetToCurrentEdit = {offset_to_current_edit}")

    ue_offset = offset_to_current_edit
    ver, ue_type, ue_len = parse_record_header(ppt, ue_offset)
    print(f"  UserEditAtom at offset={ue_offset}, type=0x{ue_type:04X}, recLen={ue_len}")

    if ue_type != 0x0FF5:
        print(f"  [FAIL] 不是 UserEditAtom!")
        ole.close()
        return

    offset_persist_dir = read_u32_le(ppt, ue_offset + 20)
    persist_id_seed = read_u32_le(ppt, ue_offset + 28)
    print(f"  offsetPersistDirectory = {offset_persist_dir}")
    print(f"  persistIdSeed = {persist_id_seed}")

    # 验证 PersistDirectoryAtom 位置
    ver, pd_type, pd_len = parse_record_header(ppt, offset_persist_dir)
    print(f"  PersistDirectoryAtom at offset={offset_persist_dir}, type=0x{pd_type:04X}, recLen={pd_len}")
    if pd_type != 0x1772:
        print(f"  [FAIL] 不是 PersistDirectoryAtom!")
        ole.close()
        return

    # 解析 persist entries
    entries = parse_persist_directory(ppt, offset_persist_dir)
    print(f"  persist entries ({len(entries)}):")

    # 验证每个 persist entry 的 offset 是否指向有效的 record
    ok_count = 0
    fail_count = 0
    for pid, poff in entries:
        if poff + 8 <= len(ppt):
            ver, rec_type, rec_len = parse_record_header(ppt, poff)
            # 检查 record 是否在合理范围内
            record_end = poff + 8 + rec_len
            if record_end <= len(ppt):
                ok_count += 1
            else:
                print(f"    [WARN] pid={pid}, offset={poff}, type=0x{rec_type:04X}, recLen={rec_len}, end={record_end} > stream_len={len(ppt)}")
                fail_count += 1
        else:
            print(f"    [FAIL] pid={pid}, offset={poff} 超出范围!")
            fail_count += 1

    print(f"  persist entries 验证: {ok_count} OK, {fail_count} FAIL")

    # 检查 UserEditAtom 是否在 stream 末尾
    ue_end = ue_offset + 8 + ue_len
    print(f"  UserEditAtom end = {ue_end}, stream len = {len(ppt)}")
    if ue_end == len(ppt):
        print(f"  [OK] UserEditAtom 在 stream 末尾")
    else:
        print(f"  [WARN] UserEditAtom 不在 stream 末尾（差 {len(ppt) - ue_end} 字节）")

    ole.close()


def main():
    check_file(ORIG_FILE, "原始文件")
    check_file(WM_FILE, "水印文件")


if __name__ == "__main__":
    main()
