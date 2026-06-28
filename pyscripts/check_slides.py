# -*- coding: utf-8 -*-
"""检查 .ppt 文件中所有 Slide 的位置和类型，以及水印插入位置。"""
import os
import struct
import olefile

BASE = os.path.dirname(__file__)
ORIG_FILE = os.path.join(BASE, "_test", "心理账户理论.ppt")
WM_FILE = os.path.join(BASE, "_test_out", "wm_心理账户理论.ppt")


def read_u32_le(data, offset):
    return struct.unpack_from("<I", data, offset)[0]


def read_u16_le(data, offset):
    return struct.unpack_from("<H", data, offset)[0]


def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = read_u16_le(data, offset)
    rec_type = read_u16_le(data, offset + 2)
    rec_len = read_u32_le(data, offset + 4)
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    is_container = (ver == 0xF)
    return (ver, inst, rec_type, rec_len, is_container)


# Slide 类型常量
RT_SLIDE = 0x03F8
RT_PPDRAWING = 0x040C
RT_SLIDE_ATOM = 0x03EE  # SlideAtom


def find_all_slides(data):
    """找到所有 Slide record，返回 [(offset, rec_len)]。"""
    slides = []
    pos = 0
    while pos + 8 <= len(data):
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len, is_container = hdr
        total_len = 8 + rec_len

        if is_container and rec_type == RT_SLIDE:
            slides.append((pos, rec_len))

        pos += total_len
        if not is_container and rec_len == 0:
            break

    return slides


def get_slide_type(data, slide_offset):
    """获取 Slide 的类型。
    SlideAtom (0x03EE) 中的 geomFlags 可以区分 Slide Master 和 Slide。
    实际上，通过检查 Slide 中是否包含特定的子 record 来判断类型。
    """
    hdr = parse_record_header(data, slide_offset)
    slide_end = slide_offset + 8 + hdr[3]

    # 遍历 Slide 的子 record
    pos = slide_offset + 8
    children = []
    has_slide_atom = False
    has_ppdrawing = False
    ppd_offset = None

    while pos + 8 <= slide_end:
        s_hdr = parse_record_header(data, pos)
        if s_hdr is None:
            break
        s_ver, s_inst, s_type, s_len, s_container = s_hdr
        s_total = 8 + s_len

        children.append((s_type, s_len, pos))
        if s_type == RT_SLIDE_ATOM:
            has_slide_atom = True
        elif s_type == RT_PPDRAWING:
            has_ppdrawing = True
            ppd_offset = pos

        pos += s_total
        if not s_container and s_len == 0:
            break

    return has_slide_atom, has_ppdrawing, ppd_offset, children


def check_watermark_in_slide(data, slide_offset):
    """检查 Slide 中是否有水印文本。"""
    hdr = parse_record_header(data, slide_offset)
    slide_end = slide_offset + 8 + hdr[3]

    # 搜索水印文本（UTF-16LE）
    watermark_text = "pptx-rs 水印".encode('utf-16-le')
    slide_data = data[slide_offset:slide_end]

    positions = []
    pos = 0
    while True:
        idx = slide_data.find(watermark_text, pos)
        if idx == -1:
            break
        positions.append(slide_offset + idx)
        pos = idx + 1

    return positions


def main():
    for label, path in [("原始文件", ORIG_FILE), ("水印文件", WM_FILE)]:
        print(f"\n{'=' * 60}")
        print(f"{label}: {os.path.basename(path)}")
        print(f"{'=' * 60}")

        ole = olefile.OleFileIO(path)
        ppt = ole.openstream("PowerPoint Document").read()
        print(f"PowerPoint Document 大小: {len(ppt)}")

        slides = find_all_slides(ppt)
        print(f"找到 {len(slides)} 个 Slide record")

        for i, (slide_offset, slide_len) in enumerate(slides):
            has_slide_atom, has_ppdrawing, ppd_offset, children = get_slide_type(ppt, slide_offset)
            wm_positions = check_watermark_in_slide(ppt, slide_offset)

            print(f"\n  Slide #{i + 1} at offset={slide_offset}, len={slide_len}")
            print(f"    hasSlideAtom={has_slide_atom}, hasPPDrawing={has_ppdrawing}")
            if ppd_offset:
                print(f"    PPDrawing at offset={ppd_offset}")
            print(f"    子 record 类型: {[hex(t) for t, _, _ in children[:10]]}")

            if wm_positions:
                print(f"    [水印] 找到 {len(wm_positions)} 个水印文本: {wm_positions}")
            else:
                print(f"    [水印] 无水印")

        ole.close()


if __name__ == "__main__":
    main()
