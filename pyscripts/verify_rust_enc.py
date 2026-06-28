"""验证 Rust 生成的加密文件结构是否与 msoffcrypto 一致。"""
import olefile

for label, path in [
    ("msoffcrypto", "_test_out/msoffcrypto_encrypted.pptx"),
    ("Rust", "_test_out/protected_文旅IP人设打造抖音短视频运营方案.pptx"),
]:
    print(f"=== {label} ===")
    try:
        ole = olefile.OleFileIO(path)
        streams = ["/".join(s) for s in ole.listdir()]
        print(f"  Streams ({len(streams)}):")
        for s in sorted(streams):
            size = len(ole.openstream(s).read())
            print(f"    {s}: {size} bytes")
        ole.close()
    except Exception as e:
        print(f"  Error: {e}")
    print()
