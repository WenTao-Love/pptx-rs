#!/usr/bin/env python3
"""用 msoffcrypto 加密原始文件，作为结构参考。"""
import msoffcrypto
import io

with open("_test/心理账户理论.ppt", "rb") as f:
    officefile = msoffcrypto.OfficeFile(f)
    officefile.load_key(password="pptx-rs-secret")
    encrypted = io.BytesIO()
    officefile.encrypt(encrypted, "pptx-rs-secret")

with open("_test_out/msoffcrypto_encrypted_心理账户理论.ppt", "wb") as f:
    f.write(encrypted.getvalue())

print("已生成 _test_out/msoffcrypto_encrypted_心理账户理论.ppt")
