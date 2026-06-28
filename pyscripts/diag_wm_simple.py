#!/usr/bin/env python3
"""简洁诊断：水印注入位置和 FOPT 属性。

关键检查：
1. 水印 SpContainer 是否在 MainMaster 中（而非 Slide）
2. 水印的 FOPT 保护属性是否正确
3. 通过 MSOSPT=0xCA (textBox=202) 和 0x01C2 (ProtectionBool) 识别水印
"""

import struct
import olefile

RT_MAIN_MASTER = 0x03F8
RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_DG_CONTAINER = 0xF002
RT_SPGR_CONTAINER = 0xF003
RT_SP_CONTAINER = 0xF004
RT_FSP = 0xF00A
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


def find_watermark_spcontainers(data):
    """查找所有水印 SpContainer（通过 MSOSPT=0xCA=202 识别）。"""
    results = []
    pos = 0
    while pos + 8 <= len(data):
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        total_len = 8 + rec_len

        # 检查是否是 SpContainer
        if ver == 0xF and rec_type == RT_SP_CONTAINER:
            # 在 SpContainer 中查找 FSP
            sp_end = pos + total_len
            sp_pos = pos + 8
            while sp_pos + 8 <= sp_end:
                fsp_rh = parse_rh(data, sp_pos)
                if fsp_rh is None:
                    break
                fsp_ver, fsp_inst, fsp_type, fsp_len = fsp_rh
                if fsp_type == RT_FSP:
                    # fsp_inst 是 MSOSPT 形状类型
                    msospt = fsp_inst
                    if msospt == 0xCA:  # textBox = 202
                        shape_id = struct.unpack_from('<I', data, sp_pos + 8)[0]
                        results.append((pos, shape_id, msospt, rec_len))
                    break
                sp_pos += 8 + fsp_len

        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return results


def find_container_at(data, target_offset):
    """找到 target_offset 所在的顶层 container 类型和位置。"""
    pos = 0
    while pos + 8 <= len(data):
        rh = parse_rh(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        total_len = 8 + rec_len
        container_end = pos + total_len

        # 检查 target_offset 是否在这个 container 范围内
        if pos <= target_offset < container_end:
            # 如果是 MainMaster 或 Slide，返回
            if rec_type in (RT_MAIN_MASTER, RT_SLIDE):
                type_name = {RT_MAIN_MASTER: "MainMaster", RT_SLIDE: "Slide"}.get(rec_type)
                return (pos, type_name, rec_type, rec_len)
            # 否则继续在子 record 中查找
            # 但我们需要找到顶层 container，所以继续遍历
            # 实际上 target_offset 在这个 container 内，我们需要确认这是顶层
            # 如果 pos == 0 或者这个 record 是顶层，就返回
            # 简化：直接返回这个 container
            type_name = f"0x{rec_type:04X}"
            return (pos, type_name, rec_type, rec_len)

        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return None


def parse_fopt_props(data, fopt_offset, fopt_len):
    """解析 FOPT 属性。"""
    props = []
    pos = fopt_offset + 8
    end = fopt_offset + 8 + fopt_len
    while pos + 6 <= end:
        w = struct.unpack_from('<H', data, pos)[0]
        value = struct.unpack_from('<I', data, pos + 2)[0]
        flags = (w >> 13) & 0x7
        prop_id = w & 0x1FFF
        props.append((prop_id, value, flags))
        pos += 6
    return props


def analyze_watermark_in_file(path, label):
    """分析文件中的水印位置和属性。"""
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

    print(f"PowerPoint Document stream 大小: {len(ppt_data)}")

    # 查找所有水印 SpContainer（MSOSPT=0xCA）
    watermarks = find_watermark_spcontainers(ppt_data)
    print(f"\n找到 {len(watermarks)} 个水印 SpContainer (MSOSPT=0xCA=textBox):")

    for wm_offset, shape_id, msospt, sp_len in watermarks:
        # 找到水印所在的顶层 container
        container = find_container_at(ppt_data, wm_offset)
        container_info = "未知"
        if container:
            c_off, c_name, c_type, c_len = container
            container_info = f"{c_name} @ 0x{c_off:X}"

        print(f"\n  水印 SpContainer @ 0x{wm_offset:X}: shapeId={shape_id}, len={sp_len}")
        print(f"    所在容器: {container_info}")

        # 解析 FOPT
        sp_end = wm_offset + 8 + sp_len
        sp_pos = wm_offset + 8
        while sp_pos + 8 <= sp_end:
            fsp_rh = parse_rh(ppt_data, sp_pos)
            if fsp_rh is None:
                break
            fsp_ver, fsp_inst, fsp_type, fsp_len = fsp_rh
            if fsp_type == RT_FOPT:
                props = parse_fopt_props(ppt_data, sp_pos, fsp_len)
                print(f"    FOPT: numProps={fsp_inst}, len={fsp_len}")
                for pid, val, flags in props:
                    prop_names = {
                        0x00BD: "rotation",
                        0x0180: "fillType",
                        0x01BF: "FillStyleBool",
                        0x01C1: "LineStyleBool",
                        0x01C2: "ProtectionBool",
                    }
                    pname = prop_names.get(pid, f"0x{pid:04X}")
                    print(f"      {pname} (0x{pid:04X}): 0x{val:08X} (flags={flags})")
            sp_pos += 8 + fsp_len

    # 统计顶层 MainMaster 和 Slide 数量
    masters = 0
    slides = 0
    pos = 0
    while pos + 8 <= len(ppt_data):
        rh = parse_rh(ppt_data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        if rec_type == RT_MAIN_MASTER:
            masters += 1
        elif rec_type == RT_SLIDE:
            slides += 1
        pos += 8 + rec_len
        if ver != 0xF and rec_len == 0:
            break
    print(f"\n顶层 record 统计: MainMaster={masters}, Slide={slides}")


if __name__ == "__main__":
    import os

    # 原始文件
    analyze_watermark_in_file("_test/心理账户理论.ppt", "原始文件")

    # 加水印后
    analyze_watermark_in_file("_test_out/wm_心理账户理论.ppt", "加水印后")

    # 加密文件解密后
    dec1 = "_test_out/protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec1):
        analyze_watermark_in_file(dec1, "加密文件解密后")

    dec2 = "_test_out/wm_protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec2):
        analyze_watermark_in_file(dec2, "水印+加密文件解密后")
