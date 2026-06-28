#!/usr/bin/env python3
"""检查 persist 目录顺序，对比原始文件和加密文件。

关键假设：PowerPoint 可能期望 persist 目录按 offset 顺序排列，
而不是按 persistId 顺序排列。如果 persist 目录不按 offset 排序，
PowerPoint 可能用错误的解密范围。
"""

import struct
import sys
from pathlib import Path

try:
    import olefile
except ImportError:
    print("需要 olefile: pip install olefile")
    sys.exit(1)


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset)]。"""
    h = parse_header(data, offset)
    if h is None or h[2] != 0x1772:
        return []

    _, _, _, rec_len = h
    pd_data = data[offset + 8 : offset + 8 + rec_len]

    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                poff = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + j, poff))
                pos += 4
    return entries


def analyze_persist_order(path, label):
    """分析 persist 目录顺序。"""
    print(f"\n{'='*70}")
    print(f"{label}: {path}")
    print(f"{'='*70}")

    if not Path(path).exists():
        print(f"  文件不存在")
        return

    ole = olefile.OleFileIO(path)
    ppt_data = ole.openstream("PowerPoint Document").read()
    cu_data = ole.openstream("Current User").read()

    # 找到 UserEditAtom
    ue_offset = struct.unpack_from("<I", cu_data, 16)[0]
    pd_offset = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]

    # 解析 persist 目录
    entries = parse_persist_directory(ppt_data, pd_offset)

    print(f"  persist 目录共 {len(entries)} 个 entry")

    # 检查 offset 是否按 persistId 顺序排列
    offsets_by_pid = [off for _, off in entries]
    offsets_sorted = sorted(offsets_by_pid)

    if offsets_by_pid == offsets_sorted:
        print(f"  ✓ persist 目录按 offset 顺序排列")
    else:
        print(f"  ✗ persist 目录不按 offset 顺序排列")
        print(f"    persistId 顺序的 offset: {offsets_by_pid[:10]}...")
        print(f"    offset 排序后: {offsets_sorted[:10]}...")

    # 打印前 15 个 entry
    print(f"\n  persist 目录 (按 persistId 顺序):")
    for i, (pid, off) in enumerate(entries[:15]):
        h = parse_header(ppt_data, off)
        if h:
            rec_type = h[2]
            rec_len = h[3]
            type_names = {
                0x03E8: "Document",
                0x03EE: "Slide",
                0x03F0: "SlideList",
                0x03F8: "MainMaster",
                0x0FF5: "UserEditAtom",
                0x1772: "PersistDirectoryAtom",
                0x2F14: "CryptSession10Container",
            }
            name = type_names.get(rec_type, f"0x{rec_type:04X}")
            print(f"    persistId={pid:>3}, offset={off:>8}, type=0x{rec_type:04X}({name}), len={rec_len}")

    # 检查相邻 persist 对象的 offset 差
    print(f"\n  相邻 persist 对象的 offset 差 (按 persistId 顺序):")
    for i in range(min(len(entries) - 1, 10)):
        pid1, off1 = entries[i]
        pid2, off2 = entries[i + 1]
        h1 = parse_header(ppt_data, off1)
        if h1:
            rec_len1 = h1[3]
            total_len1 = 8 + rec_len1
            gap = off2 - off1
            expected_gap = total_len1
            status = "✓" if gap == expected_gap else "✗"
            print(f"    persistId={pid1}→{pid2}: offset={off1}→{off2}, gap={gap}, expected={expected_gap} (8+{rec_len1}) {status}")

    # 按 offset 排序后检查
    print(f"\n  相邻 persist 对象的 offset 差 (按 offset 排序):")
    sorted_entries = sorted(entries, key=lambda x: x[1])
    for i in range(min(len(sorted_entries) - 1, 10)):
        pid1, off1 = sorted_entries[i]
        pid2, off2 = sorted_entries[i + 1]
        h1 = parse_header(ppt_data, off1)
        if h1:
            rec_len1 = h1[3]
            total_len1 = 8 + rec_len1
            gap = off2 - off1
            expected_gap = total_len1
            status = "✓" if gap == expected_gap else "✗"
            print(f"    persistId={pid1}→{pid2}: offset={off1}→{off2}, gap={gap}, expected={expected_gap} (8+{rec_len1}) {status}")

    ole.close()


def main():
    test_dir = Path("_test")
    out_dir = Path("_test_out")

    # 检查原始文件
    for ppt in test_dir.glob("*.ppt"):
        analyze_persist_order(ppt, "原始文件")

    # 检查加密文件
    for ppt in out_dir.glob("protected_*.ppt"):
        if "decrypted" in ppt.name:
            continue
        analyze_persist_order(ppt, "加密文件")


if __name__ == "__main__":
    main()
