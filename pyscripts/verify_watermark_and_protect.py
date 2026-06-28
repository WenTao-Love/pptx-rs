#!/usr/bin/env python3
"""验证 .ppt 水印+加密合并文件的正确性。

验证步骤：
1. 检查文件是否被加密（is_encrypted）
2. 用密码解密文件
3. 检查解密后的文件中是否包含水印文本
"""

import os
import sys
import tempfile

try:
    import msoffcrypto
except ImportError:
    print("[ERROR] 请先安装 msoffcrypto-tool: pip install msoffcrypto-tool")
    sys.exit(1)

PASSWORD = "pptx-rs-secret"
WATERMARK_TEXT = "pptx-rs 水印"

def verify_file(filepath):
    """验证单个文件。"""
    print(f"\n{'='*60}")
    print(f"验证文件: {filepath}")
    print(f"{'='*60}")

    if not os.path.exists(filepath):
        print(f"[FAIL] 文件不存在: {filepath}")
        return False

    # 步骤1：检查是否加密
    try:
        with open(filepath, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            is_encrypted = office_file.is_encrypted()
            print(f"[{'OK' if is_encrypted else 'FAIL'}] 文件已加密: {is_encrypted}")
            if not is_encrypted:
                return False
    except Exception as e:
        print(f"[FAIL] 读取文件失败: {e}")
        return False

    # 步骤2：用密码解密
    try:
        with open(filepath, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=PASSWORD)
            with tempfile.NamedTemporaryFile(suffix=".ppt", delete=False) as tmp:
                tmp_path = tmp.name
                office_file.decrypt(tmp)
            print(f"[OK] 密码正确，解密成功")
    except Exception as e:
        print(f"[FAIL] 解密失败: {e}")
        return False

    # 步骤3：检查解密后的文件中是否包含水印
    try:
        with open(tmp_path, "rb") as f:
            decrypted_data = f.read()

        # 水印文本是 UTF-16LE 编码
        watermark_utf16 = WATERMARK_TEXT.encode("utf-16-le")
        count = decrypted_data.count(watermark_utf16)

        if count > 0:
            print(f"[OK] 找到 {count} 个水印文本")
            return True
        else:
            print(f"[FAIL] 未找到水印文本")
            return False
    except Exception as e:
        print(f"[FAIL] 检查水印失败: {e}")
        return False
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def main():
    test_dir = "_test_out"
    if not os.path.exists(test_dir):
        print(f"[FAIL] 测试输出目录不存在: {test_dir}")
        return 1

    success = 0
    total = 0

    for fname in os.listdir(test_dir):
        if not fname.endswith(".ppt"):
            continue
        if not fname.startswith("wm_protected_"):
            continue
        total += 1
        filepath = os.path.join(test_dir, fname)
        if verify_file(filepath):
            success += 1

    print(f"\n{'='*60}")
    print(f"验证完成: {success}/{total} 个文件通过")
    print(f"{'='*60}")
    return 0 if success == total else 1


if __name__ == "__main__":
    sys.exit(main())
