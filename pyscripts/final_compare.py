"""对比 Rust 新版和 msoffcrypto 的 DataSpaces 二进制格式。"""
import olefile

for label, path in [
    ("msoffcrypto", "_test_out/msoffcrypto_encrypted.pptx"),
    ("Rust", "_test_out/protected_文旅IP人设打造抖音短视频运营方案.pptx"),
]:
    print(f"=== {label} ===")
    ole = olefile.OleFileIO(path)
    for stream_path in ole.listdir():
        name = "/".join(stream_path)
        data = ole.openstream(name).read()
        if name in ["EncryptedPackage", "EncryptionInfo"]:
            if name == "EncryptionInfo":
                import struct
                vmaj, vmin = struct.unpack_from("<HH", data, 0)
                print(f"  {name}: {len(data)} bytes, version={vmaj}.{vmin}")
            else:
                print(f"  {name}: {len(data)} bytes")
        else:
            print(f"  {name}: {len(data)} bytes, hex={data.hex()}")
    ole.close()
    print()
