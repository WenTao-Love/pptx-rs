#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证 PPT 水印是否正确注入且不可编辑"""

import olefile
import struct
import os
import sys

# Record type 常量
RT_SLIDE = 0x03EE
RT_MAIN_MASTER = 0x03F8
RT_PPDRAWING = 0x040C
RT_SP_CONTAINER = 0xF004
RT_SPGR_CONTAINER = 0xF003
RT_FSP = 0xF00A
RT_FOPT = 0xF00B
RT_CLIENT_TEXTBOX = 0xF00D
RT_TEXT_CHARS_ATOM = 0x0FA0

WATERMARK_TEXT = "pptx-rs 水印"

def parse_record_header(data, offset):
    """解析 8 字节 record header"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)

def find_watermark_in_ppdrawing(data, ppd_offset):
    """在 PPDrawing 中查找水印文本"""
    results = []
    _, _, _, ppd_len = parse_record_header(data, ppd_offset)
    ppd_end = ppd_offset + 8 + ppd_len

    # 递归搜索所有 TextCharsAtom
    def search_records(start, end, depth=0):
        pos = start
        while pos + 8 <= end:
            hdr = parse_record_header(data, pos)
            if hdr is None:
                break
            ver, inst, rec_type, rec_len = hdr
            is_container = ver == 0xF
            total_len = 8 + rec_len

            if rec_type == RT_TEXT_CHARS_ATOM:
                # 读取文本
                text_bytes = data[pos + 8: pos + 8 + rec_len]
                try:
                    text = text_bytes.decode('utf-16-le')
                    if WATERMARK_TEXT in text or '水印' in text or 'watermark' in text.lower():
                        results.append({
                            'offset': pos,
                            'text': text,
                            'depth': depth,
                        })
                except:
                    pass

            if rec_type == RT_FOPT:
                # 检查 FOPT 中的保护属性
                fopt_data = data[pos + 8: pos + 8 + rec_len]
                num_props = inst
                for i in range(num_props):
                    if 6 * (i + 1) > len(fopt_data):
                        break
                    prop_id = struct.unpack_from('<H', fopt_data, i * 6)[0]
                    prop_val = struct.unpack_from('<I', fopt_data, i * 6 + 2)[0]
                    if prop_id == 0x01C2:  # Protection
                        results.append({
                            'type': 'protection',
                            'offset': pos + 8 + i * 6,
                            'value': prop_val,
                            'locked': bool(prop_val & 0x01),
                            'lock_edit': bool(prop_val & 0x04),
                            'lock_select': bool(prop_val & 0x08),
                        })

            if is_container:
                search_records(pos + 8, pos + 8 + rec_len, depth + 1)

            pos += total_len
            if not is_container and rec_len == 0:
                break

    search_records(ppd_offset + 8, ppd_end)
    return results

def analyze_ppt(filepath):
    """分析 PPT 文件中的水印"""
    print(f"\n{'=' * 60}")
    print(f"分析: {os.path.basename(filepath)}")
    print(f"{'=' * 60}")

    if not olefile.isOleFile(filepath):
        print("  [FAIL] 不是有效 OLE2 文件")
        return

    ole = olefile.OleFileIO(filepath)
    # ole.listdir() 返回 [[stream_name], ...]，需要展平检查
    stream_names = ['/'.join(s) for s in ole.listdir()]
    print(f"  Streams: {stream_names}")

    if 'PowerPoint Document' not in stream_names:
        print("  [FAIL] 找不到 PowerPoint Document stream")
        ole.close()
        return

    ppt_data = ole.openstream('PowerPoint Document').read()
    print(f"  PowerPoint Document stream: {len(ppt_data)} bytes")

    # 查找所有 Slide 和 MainMaster
    slides = []
    masters = []
    pos = 0
    while pos + 8 <= len(ppt_data):
        hdr = parse_record_header(ppt_data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len = hdr
        is_container = ver == 0xF
        total_len = 8 + rec_len

        if is_container:
            if rec_type == RT_SLIDE:
                slides.append(pos)
            elif rec_type == RT_MAIN_MASTER:
                masters.append(pos)

        pos += total_len
        if not is_container and rec_len == 0:
            break

    print(f"  找到 {len(slides)} 个 Slide, {len(masters)} 个 MainMaster")

    # 在每个 Slide/Master 中查找 PPDrawing
    watermark_found = False
    for slide_type, offsets in [('Slide', slides), ('MainMaster', masters)]:
        for idx, slide_off in enumerate(offsets):
            _, _, _, slide_len = parse_record_header(ppt_data, slide_off)
            slide_end = slide_off + 8 + slide_len

            # 查找 PPDrawing
            ppos = slide_off + 8
            while ppos + 8 <= slide_end:
                hdr = parse_record_header(ppt_data, ppos)
                if hdr is None:
                    break
                ver, inst, rec_type, rec_len = hdr
                is_container = ver == 0xF
                total_len = 8 + rec_len

                if is_container and rec_type == RT_PPDRAWING:
                    results = find_watermark_in_ppdrawing(ppt_data, ppos)
                    if results:
                        watermark_found = True
                        print(f"\n  [{slide_type} #{idx}] 找到水印相关内容:")
                        for r in results:
                            if r.get('type') == 'protection':
                                print(f"    保护属性: value=0x{r['value']:08X}")
                                print(f"      locked={r['locked']}, lock_edit={r['lock_edit']}, lock_select={r['lock_select']}")
                            else:
                                print(f"    文本: '{r['text']}'")
                    break

                ppos += total_len
                if not is_container and rec_len == 0:
                    break

    if not watermark_found:
        print("\n  [WARN] 未找到水印文本")

    ole.close()

# 分析所有水印文件
test_dir = os.path.expanduser('~/pptx-rs/_test_out')
for f in sorted(os.listdir(test_dir)):
    if f.endswith('.ppt') and 'wm_' in f and not f.endswith('.decrypted'):
        analyze_ppt(os.path.join(test_dir, f))

# 也分析解密后的水印+加密文件
for f in sorted(os.listdir(test_dir)):
    if f.endswith('.decrypted'):
        analyze_ppt(os.path.join(test_dir, f))
