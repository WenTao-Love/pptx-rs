#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
解密 wm_protected 文件后检查水印文本是否存在。
"""
import io
import os
import struct
import msoffcrypto
import olefile

PASSWORD = "pptx-rs-secret"
WATERMARK_TEXT = "pptx-rs 水印"


def check_watermark_after_decrypt(filepath):
    print(f"\n=== 解密后检查水印: {os.path.basename(filepath)} ===", flush=True)
    with open(filepath, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=PASSWORD)
        out = io.BytesIO()
        office_file.decrypt(out)
        decrypted_data = out.getvalue()

    # 写入临时文件以便用 olefile 解析
    tmp_path = filepath + ".decrypted.ppt"
    with open(tmp_path, "wb") as f:
        f.write(decrypted_data)

    try:
        ole = olefile.OleFileIO(tmp_path)
        with ole.openstream("PowerPoint Document") as f:
            ppt_data = f.read()

        # 搜索水印文本（UTF-16LE）
        watermark_utf16 = WATERMARK_TEXT.encode("utf-16-le")
        count = ppt_data.count(watermark_utf16)
        if count > 0:
            print(f"  [OK] 解密后找到水印文本 '{WATERMARK_TEXT}'，出现 {count} 次", flush=True)
        else:
            print(f"  [FAIL] 解密后未找到水印文本 '{WATERMARK_TEXT}'", flush=True)

        # 检查 Slide 数量
        pos = 0
        slide_count = 0
        while pos + 8 <= len(ppt_data):
            ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", ppt_data, pos)
            ver = ver_inst & 0xF
            is_container = ver == 0xF
            if is_container and rec_type == 0x03EE:
                slide_count += 1
            pos += 8 + rec_len
            if not is_container and rec_len == 0:
                break
        print(f"  解密后 Slide(0x03EE) 数量: {slide_count}", flush=True)

        ole.close()
    finally:
        os.remove(tmp_path)


if __name__ == "__main__":
    test_out = "_test_out"
    for fname in sorted(os.listdir(test_out)):
        if fname.startswith("wm_protected_") and fname.endswith(".ppt"):
            check_watermark_after_decrypt(os.path.join(test_out, fname))
