#!/usr/bin/env python3
"""通过字节模式搜索定位水印 FSP。

水印 FSP 的字节序列：
- ver_inst = 0x0CA2 (inst=0xCA=textBox, ver=0x2) → A2 0C
- type = 0xF00A → 0A F0
- len = 8 → 08 00 00 00
搜索模式: A2 0C 0A F0 08 00 00 00
"""

import struct
import olefile

RT_MAIN_MASTER = 0x03F8
RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_SP_CONTAINER = 0xF004
RT_FOPT = 0xF00B


def parse_rh(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def find_top_container(data, target_offset):
    """找到 target_offset 所在的顶层 MainMaster 或 Slide。"""
    pos = 0
    while pos + 8 <= len(data):
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        total_len = 8 + rec_len
        if pos <= target_offset < pos + total_len:
            if rec_type in (RT_MAIN_MASTER, RT_SLIDE):
                return (pos, rec_type)
            return (pos, rec_type)
        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return None


def search_watermark_fsp(data):
    """搜索水印 FSP 的字节模式: A2 0C 0A F0 08 00 00 00。"""
    pattern = bytes([0xA2, 0x0C, 0x0A, 0xF0, 0x08, 0x00, 0x00, 0x00])
    results = []
    start = 0
    while True:
        idx = data.find(pattern, start)
        if idx == -1:
            break
        # 读取 shapeId
        shape_id = struct.unpack_from('<I', data, idx + 8)[0]
        flags = struct.unpack_from('<I', data, idx + 12)[0]
        results.append((idx, shape_id, flags))
        start = idx + 1
    return results


def search_protection_fopt(data):
    """搜索包含 ProtectionBool (0x01C2) 的 FOPT 属性。"""
    # FOPT 属性条目: 2字节 propId + 4字节 value
    # ProtectionBool propId = 0x01C2 → bytes: C2 01
    # 搜索 C2 01 后面跟着 4 字节值
    pattern = bytes([0xC2, 0x01])
    results = []
    start = 0
    while True:
        idx = data.find(pattern, start)
        if idx == -1:
            break
        # 读取 value
        value = struct.unpack_from('<I', data, idx + 2)[0]
        results.append((idx, value))
        start = idx + 1
    return results


def analyze_file(path, label):
    print(f"\n{'='*60}")
    print(f"文件: {path} ({label})")
    print(f"{'='*60}")

    try:
        ole = olefile.OleFileIO(path)
        ppt_data = ole.openstream('PowerPoint Document').read()
        ole.close()
    except Exception as e:
        print(f"读取失败: {e}")
        return

    print(f"stream 大小: {len(ppt_data)}")

    # 搜索水印 FSP
    fsps = search_watermark_fsp(ppt_data)
    print(f"\n水印 FSP (模式 A2 0C 0A F0 08 00 00 00): {len(fsps)} 个")
    for fsp_off, shape_id, flags in fsps:
        container = find_top_container(ppt_data, fsp_off)
        c_info = "未知"
        if container:
            c_off, c_type = container
            type_names = {RT_MAIN_MASTER: "MainMaster", RT_SLIDE: "Slide"}
            c_name = type_names.get(c_type, f"0x{c_type:04X}")
            c_info = f"{c_name} @ 0x{c_off:X}"
        print(f"  FSP @ 0x{fsp_off:X}: shapeId={shape_id}, flags=0x{flags:08X}, 所在容器: {c_info}")

    # 搜索 ProtectionBool 属性
    prots = search_protection_fopt(ppt_data)
    print(f"\nProtectionBool (0x01C2) 属性: {len(prots)} 个")
    for prot_off, value in prots:
        container = find_top_container(ppt_data, prot_off)
        c_info = "未知"
        if container:
            c_off, c_type = container
            type_names = {RT_MAIN_MASTER: "MainMaster", RT_SLIDE: "Slide"}
            c_name = type_names.get(c_type, f"0x{c_type:04X}")
            c_info = f"{c_name} @ 0x{c_off:X}"
        print(f"  ProtectionBool @ 0x{prot_off:X}: value=0x{value:08X}, 所在容器: {c_info}")


if __name__ == "__main__":
    import os

    analyze_file("_test/心理账户理论.ppt", "原始文件")
    analyze_file("_test_out/wm_心理账户理论.ppt", "加水印后")

    dec1 = "_test_out/protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec1):
        analyze_file(dec1, "加密文件解密后")

    dec2 = "_test_out/wm_protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec2):
        analyze_file(dec2, "水印+加密文件解密后")
