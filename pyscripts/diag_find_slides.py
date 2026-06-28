#!/usr/bin/env python3
"""模拟 Rust find_all_slides 逻辑，确认它找到的是 MainMaster 还是 Slide。"""

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


def find_all_slides_rust_logic(data):
    """模拟 Rust find_all_slides 逻辑。"""
    slides = []
    pos = 0
    iterations = 0
    while pos + 8 <= len(data):
        iterations += 1
        if iterations > 100000:
            print(f"  警告：迭代超过 100000 次，可能死循环，当前 pos=0x{pos:X}")
            break
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        is_container = ver == 0xF
        total_len = 8 + rec_len

        # Rust 代码: if is_container && rec_type == RT_MAIN_MASTER
        if is_container and rec_type == RT_MAIN_MASTER:
            slides.append((pos, rec_type, rec_len))

        pos += total_len
        if not is_container and rec_len == 0:
            break
    return slides


def find_all_slides_correct(data):
    """正确的遍历逻辑：只遍历顶层 container record。"""
    slides = []
    masters = []
    pos = 0
    iterations = 0
    while pos + 8 <= len(data):
        iterations += 1
        if iterations > 100000:
            break
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        is_container = ver == 0xF
        total_len = 8 + rec_len

        if rec_type == RT_MAIN_MASTER:
            masters.append((pos, rec_type, rec_len))
        elif rec_type == RT_SLIDE:
            slides.append((pos, rec_type, rec_len))

        pos += total_len
        if not is_container and rec_len == 0:
            break
    return masters, slides


def analyze(path, label):
    print(f"\n{'='*60}")
    print(f"文件: {path} ({label})")
    print(f"{'='*60}")

    ole = olefile.OleFileIO(path)
    ppt_data = ole.openstream('PowerPoint Document').read()
    ole.close()

    print(f"stream 大小: {len(ppt_data)}")

    # 模拟 Rust find_all_slides 逻辑
    print(f"\n模拟 Rust find_all_slides 逻辑 (查找 RT_MAIN_MASTER=0x03F8):")
    found = find_all_slides_rust_logic(ppt_data)
    print(f"  找到 {len(found)} 个 record:")
    for i, (off, rtype, rlen) in enumerate(found[:10]):
        print(f"    [{i}] offset=0x{off:X}, type=0x{rtype:04X}, len={rlen}")
    if len(found) > 10:
        print(f"    ... 还有 {len(found)-10} 个")

    # 正确的遍历逻辑
    print(f"\n正确遍历顶层 record:")
    masters, slides = find_all_slides_correct(ppt_data)
    print(f"  MainMaster: {len(masters)} 个")
    for off, rtype, rlen in masters:
        print(f"    MainMaster @ 0x{off:X}, len={rlen}")
    print(f"  Slide: {len(slides)} 个")
    if slides:
        print(f"    第一个 Slide @ 0x{slides[0][0]:X}")
        print(f"    最后一个 Slide @ 0x{slides[-1][0]:X}")


if __name__ == "__main__":
    analyze("_test/心理账户理论.ppt", "原始文件")
