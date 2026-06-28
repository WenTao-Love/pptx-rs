"""验证加密的 .ppt 文件是否能被 msoffcrypto 正确解析和解密。"""
import sys
import os
import msoffcrypto
import io

PASSWORD = "pptx-rs-secret"

files = [
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    if not os.path.exists(f):
        print(f"[SKIP] {f} 不存在")
        continue
    print(f"\n=== 验证 {f} ===")
    try:
        with open(f, "rb") as fp:
            office_file = msoffcrypto.OfficeFile(fp)
            print(f"  file format: {type(office_file).__name__}")
            # 尝试加载加密头
            try:
                office_file.load_key(password=PASSWORD)
                print("  load_key: OK")
            except Exception as e:
                print(f"  load_key FAILED: {e}")
                continue
            # 尝试解密
            out_path = f + ".decrypted.ppt"
            try:
                with open(out_path, "wb") as out_fp:
                    office_file.decrypt(out_fp)
                print(f"  decrypt: OK -> {out_path}")
                # 验证解密后的文件能否被 olefile 打开
                import olefile
                try:
                    ole = olefile.OleFileIO(out_path)
                    print(f"  olefile open: OK, streams: {ole.listdir()[:5]}...")
                    ole.close()
                except Exception as e:
                    print(f"  olefile open FAILED: {e}")
            except Exception as e:
                print(f"  decrypt FAILED: {e}")
    except Exception as e:
        print(f"  OPEN FAILED: {e}")

# 也验证一下水印文件（未加密）能否被 olefile 打开
print("\n=== 验证水印文件（未加密） ===")
wm_file = "_test_out/wm_心理账户理论.ppt"
if os.path.exists(wm_file):
    import olefile
    try:
        ole = olefile.OleFileIO(wm_file)
        print(f"  olefile open: OK")
        print(f"  streams: {ole.listdir()}")
        # 读取 PowerPoint Document stream
        if ole.exists("PowerPoint Document"):
            data = ole.openstream("PowerPoint Document").read()
            print(f"  PowerPoint Document size: {len(data)} bytes")
        ole.close()
    except Exception as e:
        print(f"  olefile open FAILED: {e}")
