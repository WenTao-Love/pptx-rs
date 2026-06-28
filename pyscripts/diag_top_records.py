#!/usr/bin/env python3
"""打印所有顶层 record，确认结构。"""

import struct
import olefile


def parse_rh(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def list_top_records(data, max_records=60):
    """列出所有顶层 record。"""
    records = []
    pos = 0
    iterations = 0
    while pos + 8 <= len(data) and iterations < max_records:
        iterations += 1
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        is_container = ver == 0xF
        total_len = 8 + rec_len
        records.append((pos, ver, inst, rec_type, rec_len, is_container))
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return records


# Record type 名称映射
TYPE_NAMES = {
    0x03E8: "Document",
    0x03EE: "Slide",
    0x03F8: "MainMaster",
    0x03F9: "SlideList",
    0x03FF: "SlideListWithText",
    0x040C: "PPDrawing",
    0x0FF5: "UserEditAtom",
    0x0FF6: "CurrentUserAtom",
    0x1772: "PersistDirectoryAtom",
    0x2F14: "CryptSession10Container",
}


def get_type_name(rec_type):
    return TYPE_NAMES.get(rec_type, f"0x{rec_type:04X}")


def analyze(path, label):
    print(f"\n{'='*60}")
    print(f"文件: {path} ({label})")
    print(f"{'='*60}")

    ole = olefile.OleFileIO(path)
    ppt_data = ole.openstream('PowerPoint Document').read()
    ole.close()

    print(f"stream 大小: {len(ppt_data)}")

    records = list_top_records(ppt_data)
    print(f"\n顶层 record ({len(records)} 个):")
    print(f"{'idx':>3} {'offset':>10} {'type':>20} {'ver':>3} {'inst':>5} {'recLen':>10} {'totalLen':>10} {'container':>9} {'end':>10}")
    for i, (pos, ver, inst, rtype, rlen, is_c) in enumerate(records):
        tname = get_type_name(rtype)
        total_len = 8 + rlen
        end = pos + total_len
        c_str = "是" if is_c else "否"
        print(f"{i:>3} 0x{pos:08X} {tname:>20} {ver:>3} {inst:>5} {rlen:>10} {total_len:>10} {c_str:>9} 0x{end:08X}")


if __name__ == "__main__":
    analyze("_test/心理账户理论.ppt", "原始文件")
    analyze("_test_out/wm_心理账户理论.ppt", "加水印后")
