#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
比较加密后的 Pictures Stream 和 msoffcrypto 解密后的 Pictures Stream。
检查 msoffcrypto 是否解密了 Pictures Stream。
"""

import struct
import olefile
import io
import msoffcrypto


def main():
    import os

    # 1. 读取加密文件中的 Pictures Stream（密文）
    enc_path = "_test_out/protected_心理账户理论.ppt"
    ole_enc = olefile.OleFileIO(enc_path)
    enc_pics = None
    if ole_enc.exists('Pictures'):
        enc_pics = ole_enc.openstream('Pictures').read()
    ole_enc.close()

    if enc_pics:
        print(f"加密文件中的 Pictures Stream (密文): {len(enc_pics)} bytes")
        print(f"  前 64 字节: {enc_pics[:64].hex()}")
    else:
        print("加密文件中没有 Pictures Stream")

    # 2. 用 msoffcrypto 解密
    with open(enc_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password='pptx-rs-secret')
        out = io.BytesIO()
        office_file.decrypt(out)
        decrypted = out.getvalue()

    ole_dec = olefile.OleFileIO(io.BytesIO(decrypted))
    dec_pics = None
    if ole_dec.exists('Pictures'):
        dec_pics = ole_dec.openstream('Pictures').read()
    ole_dec.close()

    if dec_pics:
        print(f"\nmsoffcrypto 解密后的 Pictures Stream: {len(dec_pics)} bytes")
        print(f"  前 64 字节: {dec_pics[:64].hex()}")
    else:
        print("\nmsoffcrypto 解密后没有 Pictures Stream")

    # 3. 读取原始 Pictures Stream
    orig_path = "_test/心理账户理论.ppt"
    ole_orig = olefile.OleFileIO(orig_path)
    orig_pics = ole_orig.openstream('Pictures').read()
    ole_orig.close()

    print(f"\n原始 Pictures Stream: {len(orig_pics)} bytes")
    print(f"  前 64 字节: {orig_pics[:64].hex()}")

    # 4. 比较
    if enc_pics and dec_pics:
        if enc_pics == dec_pics:
            print("\n加密文件 Pictures == 解密后 Pictures")
            print("→ msoffcrypto 没有解密 Pictures Stream（解密后仍然是密文）")
        else:
            print("\n加密文件 Pictures != 解密后 Pictures")
            print("→ msoffcrypto 解密了 Pictures Stream")
            # 找第一个差异
            min_len = min(len(enc_pics), len(dec_pics))
            for i in range(min_len):
                if enc_pics[i] != dec_pics[i]:
                    print(f"  第一个差异在字节 {i}: 密文={hex(enc_pics[i])} 解密={hex(dec_pics[i])}")
                    break

    if dec_pics and orig_pics:
        if dec_pics == orig_pics:
            print("\n解密后 Pictures == 原始 Pictures (正确!)")
        else:
            print("\n解密后 Pictures != 原始 Pictures (错误!)")
            min_len = min(len(dec_pics), len(orig_pics))
            diff_count = 0
            first_diff = -1
            for i in range(min_len):
                if dec_pics[i] != orig_pics[i]:
                    if first_diff == -1:
                        first_diff = i
                    diff_count += 1
            print(f"  第一个差异在字节 {first_diff}: 原始={hex(orig_pics[first_diff])} 解密={hex(dec_pics[first_diff])}")
            print(f"  总差异字节数: {diff_count} / {min_len}")


if __name__ == "__main__":
    main()
