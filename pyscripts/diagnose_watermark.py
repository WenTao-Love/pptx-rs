#!/usr/bin/env python3
"""诊断水印注入位置和 FOPT 属性。

对比原始文件和加水印后的文件，确认：
1. 水印是否注入到 Slide Master（MainMaster 0x03F8）而非 Slide
2. 水印 SpContainer 的 FOPT 保护属性是否正确
3. 水印在 PPDrawing 中的层级位置
"""

import struct
import olefile

# Record type 常量
RT_MAIN_MASTER = 0x03F8
RT_SLIDE = 0x03EE
RT_PPDRAWING = 0x040C
RT_DG_CONTAINER = 0xF002
RT_SPGR_CONTAINER = 0xF003
RT_SP_CONTAINER = 0xF004
RT_FSP = 0xF00A
RT_FOPT = 0xF00B
RT_CLIENT_ANCHOR = 0xF010
RT_CLIENT_TEXTBOX = 0xF00D
RT_USER_EDIT_ATOM = 0x0FF5
RT_PERSIST_DIRECTORY_ATOM = 0x1772
RT_CRYPT_SESSION10_CONTAINER = 0x2F14


def parse_record_header(data, offset):
    """解析 8 字节 record header，返回 (ver, inst, recType, recLen)。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def find_top_level_records(data, target_types):
    """遍历顶层 record，返回 [(offset, ver, inst, recType, recLen)]。"""
    results = []
    pos = 0
    while pos + 8 <= len(data):
        rh = parse_record_header(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if rec_type in target_types:
            results.append((pos, ver, inst, rec_type, rec_len))
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return results


def find_child_record(data, parent_offset, parent_len, target_type):
    """在 parent container 中查找指定类型的子 record。"""
    parent_end = parent_offset + 8 + parent_len
    pos = parent_offset + 8
    while pos + 8 <= parent_end:
        rh = parse_record_header(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        if rec_type == target_type:
            return (pos, ver, inst, rec_type, rec_len)
        total_len = 8 + rec_len
        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return None


def find_all_children(data, parent_offset, parent_len, target_type):
    """在 parent container 中查找所有指定类型的子 record。"""
    parent_end = parent_offset + 8 + parent_len
    pos = parent_offset + 8
    results = []
    while pos + 8 <= parent_end:
        rh = parse_record_header(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        if rec_type == target_type:
            results.append((pos, ver, inst, rec_type, rec_len))
        total_len = 8 + rec_len
        pos += total_len
        if ver != 0xF and rec_len == 0:
            break
    return results


def parse_fopt(data, fopt_offset, fopt_len):
    """解析 FOPT 的所有属性，返回 [(propId, value, flags)]。"""
    props = []
    pos = fopt_offset + 8
    end = fopt_offset + 8 + fopt_len
    while pos + 6 <= end:
        w = struct.unpack_from('<H', data, pos)[0]
        value = struct.unpack_from('<I', data, pos + 2)[0]
        # 高 3 位是 flags，低 13 位是 property ID
        flags = (w >> 13) & 0x7
        prop_id = w & 0x1FFF
        props.append((prop_id, value, flags))
        pos += 6
    return props


def analyze_ppdrawing(data, ppd_offset, label=""):
    """分析 PPDrawing 的结构。"""
    rh = parse_record_header(data, ppd_offset)
    if rh is None:
        print(f"  {label}PPDrawing: 无法解析")
        return
    ver, inst, rec_type, rec_len = rh
    print(f"  {label}PPDrawing @ 0x{ppd_offset:X}: ver={ver}, inst={inst}, type=0x{rec_type:X}, len={rec_len}")

    # 找 DgContainer
    dg = find_child_record(data, ppd_offset, rec_len, RT_DG_CONTAINER)
    if dg is None:
        print(f"  {label}  DgContainer: 未找到")
        return
    dg_offset, dg_ver, dg_inst, dg_type, dg_len = dg
    print(f"  {label}  DgContainer @ 0x{dg_offset:X}: len={dg_len}")

    # 找 SpgrContainer
    spgr = find_child_record(data, dg_offset, dg_len, RT_SPGR_CONTAINER)
    if spgr is None:
        print(f"  {label}  SpgrContainer: 未找到")
        return
    spgr_offset, spgr_ver, spgr_inst, spgr_type, spgr_len = spgr
    print(f"  {label}  SpgrContainer @ 0x{spgr_offset:X}: len={spgr_len}")

    # 列出 SpgrContainer 的所有子 record
    spgr_end = spgr_offset + 8 + spgr_len
    pos = spgr_offset + 8
    idx = 0
    while pos + 8 <= spgr_end:
        rh = parse_record_header(data, pos)
        if rh is None:
            break
        ver, inst, rec_type, rec_len = rh
        total_len = 8 + rec_len
        type_name = {
            RT_SPGR_CONTAINER: "SpgrContainer",
            RT_SP_CONTAINER: "SpContainer",
            RT_FSP: "FSP",
            RT_FOPT: "FOPT",
            RT_CLIENT_ANCHOR: "ClientAnchor",
            RT_CLIENT_TEXTBOX: "ClientTextbox",
        }.get(rec_type, f"0x{rec_type:04X}")
        print(f"  {label}    [{idx}] {type_name} @ 0x{pos:X}: ver={ver}, inst={inst}, type=0x{rec_type:X}, len={rec_len}")

        # 如果是 SpContainer，分析其 FSP 和 FOPT
        if rec_type == RT_SP_CONTAINER:
            fsp = find_child_record(data, pos, rec_len, RT_FSP)
            if fsp:
                fsp_off, fsp_v, fsp_inst, fsp_t, fsp_l = fsp
                # FSP: shapeId (4) + flags (4)
                shape_id = struct.unpack_from('<I', data, fsp_off + 8)[0]
                shape_flags = struct.unpack_from('<I', data, fsp_off + 12)[0]
                # inst 是 MSOSPT 形状类型
                msospt = fsp_inst
                print(f"  {label}      FSP: shapeId={shape_id}, MSOSPT={msospt} (0x{msospt:X}), flags=0x{shape_flags:08X}")

            fopt = find_child_record(data, pos, rec_len, RT_FOPT)
            if fopt:
                fopt_off, fopt_v, fopt_inst, fopt_t, fopt_l = fopt
                props = parse_fopt(data, fopt_off, fopt_l)
                print(f"  {label}      FOPT: numProps={fopt_inst}, len={fopt_l}")
                for pid, val, flags in props:
                    prop_names = {
                        0x00BD: "rotation",
                        0x0180: "fillType",
                        0x01BF: "FillStyleBool",
                        0x01C1: "LineStyleBool",
                        0x01C2: "ProtectionBool",
                    }
                    pname = prop_names.get(pid, f"0x{pid:04X}")
                    print(f"  {label}        {pname} (0x{pid:04X}): 0x{val:08X} (flags={flags})")

        pos += total_len
        idx += 1
        if ver != 0xF and rec_len == 0:
            break


def analyze_file(path, label):
    """分析 .ppt 文件的结构。"""
    print(f"\n{'='*60}")
    print(f"分析: {path} ({label})")
    print(f"{'='*60}")

    ole = olefile.OleFileIO(path)
    ppt_data = ole.openstream('PowerPoint Document').read()
    ole.close()

    print(f"PowerPoint Document stream 大小: {len(ppt_data)}")

    # 找所有顶层 MainMaster 和 Slide
    top_records = find_top_level_records(
        ppt_data,
        [RT_MAIN_MASTER, RT_SLIDE, RT_USER_EDIT_ATOM, RT_PERSIST_DIRECTORY_ATOM]
    )

    print(f"\n顶层 record:")
    for off, ver, inst, rtype, rlen in top_records:
        type_names = {
            RT_MAIN_MASTER: "MainMaster",
            RT_SLIDE: "Slide",
            RT_USER_EDIT_ATOM: "UserEditAtom",
            RT_PERSIST_DIRECTORY_ATOM: "PersistDirectoryAtom",
        }
        tname = type_names.get(rtype, f"0x{rtype:04X}")
        print(f"  {tname} @ 0x{off:X}: ver={ver}, inst={inst}, type=0x{rtype:X}, len={rlen}")

    # 分析每个 MainMaster 的 PPDrawing
    masters = [r for r in top_records if r[3] == RT_MAIN_MASTER]
    slides = [r for r in top_records if r[3] == RT_SLIDE]

    print(f"\nMainMaster 数量: {len(masters)}")
    for off, ver, inst, rtype, rlen in masters:
        print(f"\nMainMaster @ 0x{off:X}:")
        ppd = find_child_record(ppt_data, off, rlen, RT_PPDRAWING)
        if ppd:
            ppd_off, ppd_v, ppd_i, ppd_t, ppd_l = ppd
            analyze_ppdrawing(ppt_data, ppd_off, "  ")
        else:
            print("  PPDrawing: 未找到")

    print(f"\nSlide 数量: {len(slides)}")
    # 检查 Slide 中是否有 PPDrawing（水印是否错误注入到 Slide）
    for off, ver, inst, rtype, rlen in slides[:3]:  # 只检查前3个
        print(f"\nSlide @ 0x{off:X}:")
        ppd = find_child_record(ppt_data, off, rlen, RT_PPDRAWING)
        if ppd:
            ppd_off, ppd_v, ppd_i, ppd_t, ppd_l = ppd
            analyze_ppdrawing(ppt_data, ppd_off, "  ")
        else:
            print("  PPDrawing: 未找到")


if __name__ == "__main__":
    # 分析原始文件
    analyze_file("_test/心理账户理论.ppt", "原始文件")

    # 分析加水印后的文件
    analyze_file("_test_out/wm_心理账户理论.ppt", "加水印后")

    # 分析加密后的文件（解密后检查）
    # protected_心理账户理论.ppt 是加密的，需要先解密
    # 但我们可以检查解密后的文件
    import os
    dec_path = "_test_out/protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec_path):
        analyze_file(dec_path, "加密文件解密后")

    dec_path2 = "_test_out/wm_protected_心理账户理论.ppt.decrypted.ppt"
    if os.path.exists(dec_path2):
        analyze_file(dec_path2, "水印+加密文件解密后")
