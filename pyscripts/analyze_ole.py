"""分析加密 .ppt 文件的 OLE2 结构，找出 PowerPoint 无法打开的原因。"""
import olefile
import struct
import sys

def analyze_ole(path):
    """分析 OLE2 容器结构。"""
    print(f"\n=== 分析 {path} ===")
    try:
        ole = olefile.OleFileIO(path)
    except Exception as e:
        print(f"  打开失败: {e}")
        return

    # 列出所有 streams
    print(f"  Streams:")
    for entry in ole.listdir(streams=True, storages=False):
        name = "/".join(entry)
        size = ole.get_size(name)
        print(f"    {name}: {size} bytes")

    # 读取 header 信息
    print(f"\n  OLE2 Header:")
    print(f"    Sector size: {ole.sectorsize}")
    print(f"    Mini sector size: {ole.minisectorcutoff}")
    print(f"    Num FAT sectors: {ole.num_fat_sectors}")
    print(f"    First mini FAT sector: {ole.first_mini_fat_sector}")
    print(f"    Num mini FAT sectors: {ole.num_mini_fat_sectors}")
    print(f"    First DIFAT sector: {ole.first_difat_sector}")
    print(f"    Num DIFAT sectors: {ole.num_difat_sectors}")

    # 读取 raw header (前 512 字节)
    with open(path, "rb") as f:
        header = f.read(512)

    # 解析 header 字段
    print(f"\n  Raw Header Fields:")
    sig = header[0:8]
    print(f"    Signature: {sig.hex()}")
    sector_shift = struct.unpack_from("<H", header, 30)[0]
    print(f"    Sector shift: {sector_shift} (sector size = {1 << sector_shift})")
    mini_sector_shift = struct.unpack_from("<H", header, 32)[0]
    print(f"    Mini sector shift: {mini_sector_shift}")
    num_fat_sectors = struct.unpack_from("<I", header, 44)[0]
    print(f"    Num FAT sectors: {num_fat_sectors}")
    first_dir_sector = struct.unpack_from("<I", header, 48)[0]
    print(f"    First dir sector: {first_dir_sector}")
    transaction_sig = struct.unpack_from("<I", header, 52)[0]
    print(f"    Transaction sig: 0x{transaction_sig:08X}")
    mini_stream_cutoff = struct.unpack_from("<I", header, 56)[0]
    print(f"    Mini stream cutoff: {mini_stream_cutoff}")
    first_mini_fat = struct.unpack_from("<I", header, 60)[0]
    print(f"    First mini FAT sector: {first_mini_fat}")
    num_mini_fat = struct.unpack_from("<I", header, 64)[0]
    print(f"    Num mini FAT sectors: {num_mini_fat}")
    first_difat = struct.unpack_from("<I", header, 68)[0]
    print(f"    First DIFAT sector: {first_difat}")
    num_difat = struct.unpack_from("<I", header, 72)[0]
    print(f"    Num DIFAT sectors: {num_difat}")

    # 读取 DIFAT 数组
    print(f"\n  DIFAT entries (first 10):")
    for i in range(10):
        val = struct.unpack_from("<I", header, 76 + i * 4)[0]
        if val != 0xFFFFFFFF:
            print(f"    DIFAT[{i}] = {val}")

    # 读取 Current User stream
    if ole.exists("Current User"):
        cu = ole.openstream("Current User").read()
        print(f"\n  Current User stream ({len(cu)} bytes):")
        # CurrentUserAtom 结构
        if len(cu) >= 20:
            ver_inst = struct.unpack_from("<H", cu, 0)[0]
            rec_type = struct.unpack_from("<H", cu, 2)[0]
            rec_len = struct.unpack_from("<I", cu, 4)[0]
            size = struct.unpack_from("<I", cu, 8)[0]
            header_token = struct.unpack_from("<I", cu, 12)[0]
            offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
            print(f"    ver_inst: 0x{ver_inst:04X}")
            print(f"    rec_type: 0x{rec_type:04X}")
            print(f"    rec_len: {rec_len}")
            print(f"    size: {size}")
            print(f"    headerToken: 0x{header_token:08X} ({'ENCRYPTED' if header_token == 0xF3D1C4D0 else 'NOT ENCRYPTED'})")
            print(f"    offsetToCurrentEdit: {offset_to_current_edit}")

    ole.close()


# 分析所有文件
files = [
    "_test/心理账户理论.ppt",
    "_test_out/wm_心理账户理论.ppt",
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        analyze_ole(f)
    except Exception as e:
        print(f"  分析失败: {e}")
        import traceback
        traceback.print_exc()
