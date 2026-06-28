#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 msoffcrypto 验证加密 .ppt 文件能否被正确解密。
"""

import os
import sys
import io
import msoffcrypto


def verify_file(filepath, password):
    print(f"\n=== 验证 {filepath} ===")
    try:
        with open(filepath, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            print(f"  文件类型: {type(office_file).__name__}")
            # load_key 不返回布尔值，成功时返回 None，失败时抛出异常
            office_file.load_key(password=password)
            print(f"  load_key 成功")
            # 尝试解密
            output = io.BytesIO()
            office_file.decrypt(output)
            decrypted = output.getvalue()
            print(f"  解密成功，大小: {len(decrypted)} 字节")
            # 检查 OLE2 签名
            if decrypted[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                print("  解密后是有效的 OLE2 文件")
            else:
                print(f"  解密后不是 OLE2 文件，签名: {decrypted[:8].hex()}")
            return True
    except Exception as e:
        print(f"  错误: {type(e).__name__}: {e}")
        return False


def main():
    password = "pptx-rs-secret"
    test_dir = "_test_out"
    for fname in sorted(os.listdir(test_dir)):
        if fname.lower().endswith(".ppt") and ("protected" in fname.lower()):
            verify_file(os.path.join(test_dir, fname), password)


if __name__ == "__main__":
    main()
