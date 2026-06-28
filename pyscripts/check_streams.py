"""深入检查加密 .ppt 文件的所有 streams 和结构。"""
import olefile
import struct

def check_streams(path):
    """检查所有 streams 是否能被正确读取。"""
    print(f"\n=== {path} ===")
    ole = olefile.OleFileIO(path)

    # 列出所有 streams
    print("\n  Streams:")
    for entry in ole.listdir(streams=True):
        name = "/".join(entry)
        size = ole.get_size(name)
        print(f"    {name}: {size} bytes")

    # 读取所有 streams
    print("\n  Reading streams:")
    for entry in ole.listdir(streams=True):
        name = "/".join(entry)
        try:
            data = ole.openstream(name).read()
            print(f"    {name}: read {len(data)} bytes, first 16 bytes: {data[:16].hex()}")
        except Exception as e:
            print(f"    {name}: FAILED to read - {e}")

    # 检查 EncryptedSummaryInfo stream
    print("\n  EncryptedSummaryInfo stream:")
    try:
        data = ole.openstream("EncryptedSummaryInfo").read()
        print(f"    size: {len(data)} bytes")
        print(f"    first 32 bytes: {data[:32].hex()}")
        # 解析 Version
        if len(data) >= 4:
            version = struct.unpack_from("<I", data, 0)[0]
            print(f"    Version: 0x{version:08X}")
        # 解析 EncryptionVersionInfo (vMajor, vMinor)
        if len(data) >= 8:
            v_major = struct.unpack_from("<H", data, 0)[0]
            v_minor = struct.unpack_from("<H", data, 2)[0]
            print(f"    vMajor={v_major}, vMinor={v_minor}")
    except Exception as e:
        print(f"    FAILED: {e}")

    # 检查 Current User stream
    print("\n  Current User stream:")
    try:
        data = ole.openstream("Current User").read()
        print(f"    size: {len(data)} bytes")
        # 解析 CurrentUserAtom
        if len(data) >= 20:
            rec_type = struct.unpack_from("<H", data, 2)[0]
            rec_len = struct.unpack_from("<I", data, 4)[0]
            header_token = struct.unpack_from("<I", data, 12)[0]
            offset_to_current_edit = struct.unpack_from("<I", data, 16)[0]
            print(f"    rec_type=0x{rec_type:04X}, rec_len={rec_len}")
            print(f"    header_token=0x{header_token:08X}")
            print(f"    offsetToCurrentEdit={offset_to_current_edit}")
    except Exception as e:
        print(f"    FAILED: {e}")

    # 检查 PowerPoint Document stream 的前 32 字节
    print("\n  PowerPoint Document stream (first 64 bytes):")
    try:
        data = ole.openstream("PowerPoint Document").read()
        print(f"    total size: {len(data)} bytes")
        print(f"    first 64 bytes: {data[:64].hex()}")
    except Exception as e:
        print(f"    FAILED: {e}")

    ole.close()

# 检查所有文件
files = [
    "_test/心理账户理论.ppt",
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        check_streams(f)
    except Exception as e:
        print(f"  分析失败: {e}")
        import traceback
        traceback.print_exc()
