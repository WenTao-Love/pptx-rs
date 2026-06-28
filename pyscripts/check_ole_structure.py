#!/usr/bin/env python3
"""检查 OLE2 容器结构，对比原始文件和加密文件。"""

import struct
import sys
from pathlib import Path

try:
    import olefile
except ImportError:
    print("需要 olefile: pip install olefile")
    sys.exit(1)


def check_ole_structure(path, label):
    """检查 OLE2 容器结构。"""
    print(f"\n{'='*70}")
    print(f"{label}: {path}")
    print(f"{'='*70}")

    if not Path(path).exists():
        print(f"  文件不存在")
        return

    ole = olefile.OleFileIO(path)

    # 列出所有 streams
    print(f"\n  Streams:")
    for entry in ole.listdir(streams=True, storages=True):
        path_str = "/".join(entry)
        try:
            size = ole.get_size(path_str)
            print(f"    {path_str}: {size} bytes")
        except Exception:
            print(f"    {path_str}: (storage)")

    # 检查根目录 CLSID
    root = ole.root
    print(f"\n  根目录 CLSID: {root.clsid}")

    # 检查 Current User stream
    if ole.exists("Current User"):
        cu_data = ole.openstream("Current User").read()
        print(f"\n  Current User stream ({len(cu_data)} bytes):")
        print(f"    前 40 字节: {cu_data[:40].hex()}")

        # 解析 CurrentUserAtom
        if len(cu_data) >= 20:
            ver_inst = struct.unpack_from("<H", cu_data, 0)[0]
            rec_type = struct.unpack_from("<H", cu_data, 2)[0]
            rec_len = struct.unpack_from("<I", cu_data, 4)[0]
            size = struct.unpack_from("<I", cu_data, 8)[0]
            header_token = struct.unpack_from("<I", cu_data, 12)[0]
            offset_to_current_edit = struct.unpack_from("<I", cu_data, 16)[0]

            ver = ver_inst & 0x0F
            inst = (ver_inst >> 4) & 0x0FFF

            print(f"    ver={ver}, inst=0x{inst:04X}, type=0x{rec_type:04X}, recLen={rec_len}")
            print(f"    size={size}")
            print(f"    headerToken=0x{header_token:08X} ({'已加密' if header_token == 0xF3D1C4DF else '未加密'})")
            print(f"    offsetToCurrentEdit={offset_to_current_edit}")

    # 检查 PowerPoint Document stream
    if ole.exists("PowerPoint Document"):
        ppt_data = ole.openstream("PowerPoint Document").read()
        print(f"\n  PowerPoint Document stream ({len(ppt_data)} bytes):")
        print(f"    前 16 字节: {ppt_data[:16].hex()}")

    # 检查 Pictures stream
    if ole.exists("Pictures"):
        pic_data = ole.openstream("Pictures").read()
        print(f"\n  Pictures stream ({len(pic_data)} bytes):")
        print(f"    前 16 字节: {pic_data[:16].hex()}")

    # 检查 Summary Information streams
    for name in ["\x05SummaryInformation", "\x05DocumentSummaryInformation"]:
        if ole.exists(name):
            si_data = ole.openstream(name).read()
            print(f"\n  {name} stream ({len(si_data)} bytes):")
            print(f"    前 16 字节: {si_data[:16].hex()}")

    ole.close()


def main():
    test_dir = Path("_test")
    out_dir = Path("_test_out")

    # 检查原始文件
    for ppt in test_dir.glob("*.ppt"):
        check_ole_structure(ppt, "原始文件")

    # 检查加密文件
    for ppt in out_dir.glob("protected_*.ppt"):
        if "decrypted" in ppt.name:
            continue
        check_ole_structure(ppt, "加密文件")

    # 检查解密后的文件
    for ppt in out_dir.glob("*.full_decrypted.ppt"):
        check_ole_structure(ppt, "解密后文件")


if __name__ == "__main__":
    main()
