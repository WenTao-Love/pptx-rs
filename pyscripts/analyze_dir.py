"""深入分析加密 .ppt 文件的目录结构和 mini FAT。"""
import olefile
import struct

def dump_directory(path):
    """Dump OLE2 directory entries and mini FAT."""
    print(f"\n=== {path} ===")
    with open(path, "rb") as f:
        data = f.read()

    # 解析 header
    sector_shift = struct.unpack_from("<H", data, 30)[0]
    sector_size = 1 << sector_shift
    first_dir_sector = struct.unpack_from("<I", data, 48)[0]
    mini_stream_cutoff = struct.unpack_from("<I", data, 56)[0]
    first_mini_fat = struct.unpack_from("<I", data, 60)[0]
    num_mini_fat = struct.unpack_from("<I", data, 64)[0]
    print(f"  sector_size={sector_size}, first_dir_sector={first_dir_sector}")
    print(f"  mini_stream_cutoff={mini_stream_cutoff}")
    print(f"  first_mini_fat={first_mini_fat}, num_mini_fat={num_mini_fat}")

    # 读取 FAT
    difat = []
    for i in range(109):
        val = struct.unpack_from("<I", data, 76 + i * 4)[0]
        if val != 0xFFFFFFFF:
            difat.append(val)
    fat = []
    for sec in difat:
        offset = 512 + sec * sector_size
        for j in range(sector_size // 4):
            fat.append(struct.unpack_from("<I", data, offset + j * 4)[0])

    # 读取目录
    print(f"\n  Directory entries:")
    dir_sector = first_dir_sector
    dir_entries = []
    while dir_sector != 0xFFFFFFFE and dir_sector < 0xFFFFFFFA:
        offset = 512 + dir_sector * sector_size
        for i in range(sector_size // 128):
            entry_offset = offset + i * 128
            if entry_offset + 128 > len(data):
                break
            name_len = struct.unpack_from("<H", data, entry_offset + 64)[0]
            if name_len == 0:
                continue
            name = data[entry_offset:entry_offset + name_len].decode("utf-16-le", errors="replace").rstrip("\x00")
            entry_type = data[entry_offset + 66]
            entry_color = data[entry_offset + 67]
            left_sibling = struct.unpack_from("<I", data, entry_offset + 68)[0]
            right_sibling = struct.unpack_from("<I", data, entry_offset + 72)[0]
            child = struct.unpack_from("<I", data, entry_offset + 76)[0]
            clsid = data[entry_offset + 80:entry_offset + 96]
            start_sector = struct.unpack_from("<I", data, entry_offset + 116)[0]
            stream_size = struct.unpack_from("<Q", data, entry_offset + 120)[0]
            dir_entries.append({
                "name": name,
                "type": entry_type,
                "left": left_sibling,
                "right": right_sibling,
                "child": child,
                "start_sector": start_sector,
                "stream_size": stream_size,
            })
            type_str = {0: "Unknown", 1: "Storage", 2: "Stream", 5: "Root"}.get(entry_type, f"?{entry_type}")
            print(f"    [{len(dir_entries)-1}] {type_str:8} '{name}' start_sector={start_sector} size={stream_size} left={left_sibling} right={right_sibling} child={child}")
        dir_sector = fat[dir_sector] if dir_sector < len(fat) else 0xFFFFFFFE

    # 读取 mini FAT
    print(f"\n  Mini FAT entries (sector {first_mini_fat}):")
    mini_fat_sector = first_mini_fat
    mini_fat = []
    while mini_fat_sector != 0xFFFFFFFE and mini_fat_sector < 0xFFFFFFFA:
        offset = 512 + mini_fat_sector * sector_size
        for j in range(sector_size // 4):
            val = struct.unpack_from("<I", data, offset + j * 4)[0]
            mini_fat.append(val)
        next_sector = fat[mini_fat_sector] if mini_fat_sector < len(fat) else 0xFFFFFFFE
        if next_sector == 0xFFFFFFFE:
            break
        mini_fat_sector = next_sector

    # 打印前 20 个 mini FAT entries
    for i, val in enumerate(mini_fat[:20]):
        if val == 0xFFFFFFFF:
            continue
        val_str = f"0x{val:08X}"
        if val == 0xFFFFFFFE:
            val_str = "ENDOFCHAIN"
        elif val == 0xFFFFFFFF:
            val_str = "FREESECT"
        print(f"    mini_fat[{i}] = {val_str}")

    # 找到 Root Entry 的 start_sector（mini stream 的起始）
    root_entry = None
    for entry in dir_entries:
        if entry["type"] == 5:
            root_entry = entry
            break
    if root_entry:
        print(f"\n  Root Entry: start_sector={root_entry['start_sector']}, size={root_entry['stream_size']}")
        # mini stream 存储在 root entry 的 start_sector 开始的链中
        mini_stream_sectors = []
        sec = root_entry["start_sector"]
        while sec != 0xFFFFFFFE and sec < 0xFFFFFFFA:
            mini_stream_sectors.append(sec)
            sec = fat[sec] if sec < len(fat) else 0xFFFFFFFE
        print(f"    mini stream sectors: {mini_stream_sectors[:10]}{'...' if len(mini_stream_sectors) > 10 else ''}")
        # 每个 mini sector 是 64 字节
        mini_sector_size = 64
        # 检查每个 small stream 的 mini sectors
        for entry in dir_entries:
            if entry["type"] == 2 and entry["stream_size"] < mini_stream_cutoff:
                start_mini_sector = entry["start_sector"]
                size = entry["stream_size"]
                num_mini_sectors = (size + mini_sector_size - 1) // mini_sector_size
                print(f"    Stream '{entry['name']}': start_mini_sector={start_mini_sector}, size={size}, num_mini_sectors={num_mini_sectors}")
                # 遍历 mini FAT 链
                sec = start_mini_sector
                chain = []
                while sec != 0xFFFFFFFE and sec < 0xFFFFFFFA and sec < len(mini_fat):
                    chain.append(sec)
                    sec = mini_fat[sec]
                print(f"      mini FAT chain: {chain}")


# 分析所有文件
files = [
    "_test/心理账户理论.ppt",
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        dump_directory(f)
    except Exception as e:
        print(f"  分析失败: {e}")
        import traceback
        traceback.print_exc()
