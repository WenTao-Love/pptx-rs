#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证 pptx-rs 生成的加密 .ppt 文件是否能被 msoffcrypto 解密。

用法（在 WSL 中）：
    python3 -u verify_encrypted_ppt.py
"""

import sys
import io
import os
import traceback

def main():
    try:
        import msoffcrypto
    except ImportError:
        print("ERROR: msoffcrypto-tool 未安装", flush=True)
        sys.exit(1)

    test_dir = "_test_out"
    password = "pptx-rs-secret"

    # 查找所有 protected_*.ppt 和 wm_protected_*.ppt 文件
    files_to_test = []
    if os.path.isdir(test_dir):
        for fname in sorted(os.listdir(test_dir)):
            if fname.endswith(".ppt") and ("protected" in fname):
                files_to_test.append(os.path.join(test_dir, fname))

    if not files_to_test:
        print("没有找到需要验证的加密 .ppt 文件", flush=True)
        return

    print(f"找到 {len(files_to_test)} 个文件需要验证:", flush=True)
    for f in files_to_test:
        print(f"  - {f}", flush=True)
    print("", flush=True)

    for filepath in files_to_test:
        print(f"=== 验证 {filepath} ===", flush=True)
        try:
            with open(filepath, "rb") as f:
                officefile = msoffcrypto.OfficeFile(f)

                # 检查是否加密
                is_enc = officefile.is_encrypted()
                print(f"  is_encrypted: {is_enc}", flush=True)

                if not is_enc:
                    print("  结果: 文件未被识别为加密文件", flush=True)
                    continue

                # 尝试验证密码
                try:
                    officefile.load_key(password=password)
                    print(f"  密码验证: 成功 (password={password})", flush=True)

                    # 尝试解密
                    decrypted = io.BytesIO()
                    officefile.decrypt(decrypted)
                    dec_data = decrypted.getvalue()
                    print(f"  解密成功: 解密后大小={len(dec_data)} 字节", flush=True)

                    # 检查解密后的文件是否是有效的 OLE2 文件
                    # OLE2 文件签名: D0 CF 11 E0 A1 B1 1A E1
                    if dec_data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                        print("  解密后文件签名: 有效 OLE2 签名", flush=True)
                    else:
                        print(f"  解密后文件签名: 无效! 前8字节={dec_data[:8].hex()}", flush=True)

                except Exception as e:
                    print(f"  密码验证/解密失败: {type(e).__name__}: {e}", flush=True)
                    traceback.print_exc()

        except Exception as e:
            print(f"  打开文件失败: {type(e).__name__}: {e}", flush=True)
            traceback.print_exc()

        print("", flush=True)

if __name__ == "__main__":
    main()
