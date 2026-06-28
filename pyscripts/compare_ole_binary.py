#!/usr/bin/env python3
"""详细对比 OLE2 文件的二进制头部和目录结构。"""

import struct
import os

def read_header(filepath):
    """读取 OLE2 头部。"""
    with open(filepath, "rb") as f:
        data = f.read(512)

    print(f"\n{'='*60}")
    print(f"文件: {filepath} ({os.path.getsize(filepath)} bytes)")
    print(f"{'='*60}")

    # 头部签名
    sig = data[0:8]
    print(f"签名: {sig.hex()} ({'OK' if sig == b'\\xd0\\xcf\\x11\\xe0\\xa1\\xb1\\x1a\\xe1' else 'BAD'})")

    # CLSID
    clsid = data[8:24]
    print(f"头部 CLSID: {clsid.hex()}")

    # 版本
    minor = struct.unpack("<H", data[24:26])[0]
    major = struct.unpack("<H", data[26:28])[0]
    print(f"版本: {major}.{minor}")

    # 字节序
    bo = struct.unpack("<H", data[28:30])[0]
    print(f"字节序: 0x{bo:04X} ({'OK' if bo == 0xFFFE else 'BAD'})")

    # 扇区大小
    sector_shift = struct.unpack("<H", data[30:32])[0]
    mini_shift = struct.unpack("<H", data[32:34])[0]
    print(f"扇区大小: 2^{sector_shift} = {1 << sector_shift}")
    print(f"Mini 扇区大小: 2^{mini_shift} = {1 << mini_shift}")

    # 保留字段
    reserved = data[34:40]
    print(f"保留字段: {reserved.hex()}")

    # 目录扇区数
    dir_sectors = struct.unpack("<I", data[40:44])[0]
    print(f"目录扇区数: {dir_sectors}")

    # FAT 扇区数
    fat_sectors = struct.unpack("<I", data[44:48])[0]
    print(f"FAT 扇区数: {fat_sectors}")

    # 第一个目录扇区 SID
    first_dir_sid = struct.unpack("<i", data[48:52])[0]
    print(f"第一个目录扇区 SID: {first_dir_sid}")

    # 事务签名
    trans_sig = struct.unpack("<I", data[52:56])[0]
    print(f"事务签名: {trans_sig}")

    # Mini stream cutoff
    mini_cutoff = struct.unpack("<I", data[56:60])[0]
    print(f"Mini stream cutoff: {mini_cutoff}")

    # 第一个 mini FAT 扇区 SID
    first_mini_fat_sid = struct.unpack("<i", data[60:64])[0]
    print(f"第一个 mini FAT 扇区 SID: {first_mini_fat_sid}")

    # mini FAT 扇区数
    mini_fat_sectors = struct.unpack("<I", data[64:68])[0]
    print(f"Mini FAT 扇区数: {mini_fat_sectors}")

    # 第一个 DIFAT 扇区 SID
    first_difat_sid = struct.unpack("<i", data[68:72])[0]
    print(f"第一个 DIFAT 扇区 SID: {first_difat_sid}")

    # DIFAT 扇区数
    difat_sectors = struct.unpack("<I", data[72:76])[0]
    print(f"DIFAT 扇区数: {difat_sectors}")

    # DIFAT 数组（前 5 个）
    difat = []
    for i in range(5):
        val = struct.unpack("<I", data[76 + i * 4:80 + i * 4])[0]
        difat.append(val)
    print(f"DIFAT 数组（前5）: {difat}")

    return data


def read_directory(filepath, sector_size=512):
    """读取 OLE2 目录条目。"""
    with open(filepath, "rb") as f:
        data = f.read()

    # 读取头部获取第一个目录扇区
    first_dir_sid = struct.unpack("<i", data[48:52])[0]

    print(f"\n目录条目（从扇区 {first_dir_sid} 开始）:")

    # 读取 FAT
    fat_sectors_count = struct.unpack("<I", data[44:48])[0]
    difat = []
    for i in range(109):
        val = struct.unpack("<I", data[76 + i * 4:80 + i * 4])[0]
        if val != 0xFFFFFFFF:
            difat.append(val)

    # 读取 FAT 数据
    fat = []
    for sid in difat:
        offset = (sid + 1) * sector_size
        for i in range(sector_size // 4):
            val = struct.unpack("<I", data[offset + i * 4:offset + i * 4 + 4])[0]
            fat.append(val)

    # 沿着 FAT 链读取目录扇区
    dir_sectors = []
    sid = first_dir_sid
    while sid != 0xFFFFFFFE and sid < len(fat):
        dir_sectors.append(sid)
        sid = fat[sid]

    # 解析目录条目（每个 128 字节）
    for sec_idx, sec_sid in enumerate(dir_sectors):
        offset = (sec_sid + 1) * sector_size
        for i in range(sector_size // 128):
            entry_offset = offset + i * 128
            if entry_offset + 128 > len(data):
                break
            entry = data[entry_offset:entry_offset + 128]

            # 名称
            name_len = struct.unpack("<H", entry[64:66])[0]
            if name_len == 0:
                continue
            name = entry[0:name_len - 2].decode("utf-16-le", errors="replace")

            # 类型
            obj_type = entry[66]
            color = entry[67]

            # 左/右/子 SID
            left_sid = struct.unpack("<i", entry[68:72])[0]
            right_sid = struct.unpack("<i", entry[72:76])[0]
            child_sid = struct.unpack("<i", entry[76:80])[0]

            # CLSID
            clsid = entry[80:96]

            # 状态 flags
            flags = struct.unpack("<I", entry[96:100])[0]

            # 创建/修改时间
            create_time = entry[100:108]
            mod_time = entry[108:116]

            # 起始扇区
            start_sid = struct.unpack("<i", entry[116:120])[0]

            # 大小
            size = struct.unpack("<I", entry[120:124])[0]

            type_names = {0: "Empty", 1: "Storage", 2: "Stream", 5: "Root"}
            type_name = type_names.get(obj_type, f"Unknown({obj_type})")

            print(f"  [{sec_idx*4+i}] name={name!r}, type={type_name}, "
                  f"start_sid={start_sid}, size={size}, "
                  f"left={left_sid}, right={right_sid}, child={child_sid}")
            if clsid != b'\\x00' * 16:
                print(f"        CLSID: {clsid.hex()}")


def main():
    files = [
        "_test/心理账户理论.ppt",
        "_test_out/protected_心理账户理论.ppt",
    ]

    for f in files:
        if os.path.exists(f):
            read_header(f)
            read_directory(f)
        else:
            print(f"\n[SKIP] 文件不存在: {f}")


if __name__ == "__main__":
    main()
