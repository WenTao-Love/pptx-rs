#!/usr/bin/env python3
"""精确搜索水印 FSP（shapeId=0x1000, flags=0x280）。

水印 FSP 完整字节序列：
  A2 0C 0A F0 08 00 00 00 00 10 00 00 80 02 00 00
  ver_inst=0x0CA2, type=0xF00A, len=8, shapeId=0x1000, flags=0x280
"""

import struct
import olefile

RT_MAIN_MASTER = 0x03F8
RT_SLIDE = 0x03EE


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
            type_names = {RT_MAIN_MASTER: "MainMaster", RT_SLIDE: "Slide"}
            c_name = type_names.get(rec_type, f"0x{rec_type:04X}")
            return (pos, c_name, rec_type, rec_len)
        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return None


def search_watermark_precise(data):
    """精确搜索水印 FSP（shapeId=0x1000, flags=0x280）。"""
    # 完整水印 FSP 字节模式
    pattern = bytes([
        0xA2, 0x0C,  # ver_inst=0x0CA2 (inst=0xCA, ver=0x2)
        0x0A, 0xF0,  # type=0xF00A
        0x08, 0x00, 0x00, 0x00,  # len=8
        0x00, 0x10, 0x00, 0x00,  # shapeId=0x1000
        0x80, 0x02, 0x00, 0x00,  # flags=0x280
    ])
    results = []
    start = 0
    while True:
        idx = data.find(pattern, start)
        if idx == -1:
            break
        results.append(idx)
        start = idx + 1
    return results


def search_watermark_text(data):
    """搜索水印文本 'pptx-rs 水印' (UTF-16LE)。"""
    text = "pptx-rs 水印"
    text_utf16 = text.encode('utf-16-le')
    results = []
    start = 0
    while True:
        idx = data.find(text_utf16, start)
        if idx == -1:
            break
        results.append(idx)
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

    # 精确搜索水印 FSP
    wm_fsps = search_watermark_precise(ppt_data)
    print(f"\n水印 FSP (shapeId=0x1000, flags=0x280): {len(wm_fsps)} 个")
    for fsp_off in wm_fsps:
        container = find_top_container(ppt_data, fsp_off)
        c_info = "未知"
        if container:
            c_off, c_name, c_type, c_len = container
            c_info = f"{c_name} @ 0x{c_off:X}"
        print(f"  水印 FSP @ 0x{fsp_off:X}, 所在容器: {c_info}")

    # 搜索水印文本
    wm_texts = search_watermark_text(ppt_data)
    print(f"\n水印文本 'pptx-rs 水印': {len(wm_texts)} 个")
    for text_off in wm_texts:
        container = find_top_container(ppt_data, text_off)
        c_info = "未知"
        if container:
            c_off, c_name, c_type, c_len = container
            c_info = f"{c_name} @ 0x{c_off:X}"
        print(f"  水印文本 @ 0x{text_off:X}, 所在容器: {c_info}")


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
