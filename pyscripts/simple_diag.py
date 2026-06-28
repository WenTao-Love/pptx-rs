# -*- coding: utf-8 -*-
"""简化版诊断脚本。"""
import sys
import os

print("STEP 1: 脚本启动", flush=True)
sys.stdout.flush()

BASE = os.path.dirname(os.path.abspath(__file__))
ORIG_FILE = os.path.join(BASE, "_test", "心理账户理论.ppt")
ENC_FILE = os.path.join(BASE, "_test_out", "protected_心理账户理论.ppt")

print(f"STEP 2: 原始文件存在: {os.path.exists(ORIG_FILE)}", flush=True)
print(f"STEP 3: 加密文件存在: {os.path.exists(ENC_FILE)}", flush=True)

try:
    import olefile
    print("STEP 4: olefile 导入成功", flush=True)
except Exception as e:
    print(f"STEP 4 FAIL: olefile 导入失败: {e}", flush=True)
    sys.exit(1)

try:
    import msoffcrypto
    print("STEP 5: msoffcrypto 导入成功", flush=True)
except Exception as e:
    print(f"STEP 5 FAIL: msoffcrypto 导入失败: {e}", flush=True)
    sys.exit(1)

# 解密
print("STEP 6: 开始解密...", flush=True)
try:
    DEC_FILE = os.path.join(BASE, "_test_out", "decrypted_for_compare.ppt")
    with open(ENC_FILE, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password="pptx-rs-secret")
        with open(DEC_FILE, "wb") as out:
            office_file.decrypt(out)
    print(f"STEP 7: 解密成功: {DEC_FILE}", flush=True)
except Exception as e:
    print(f"STEP 7 FAIL: 解密失败: {e}", flush=True)
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 对比
print("STEP 8: 开始对比...", flush=True)
try:
    ole_orig = olefile.OleFileIO(ORIG_FILE)
    ole_dec = olefile.OleFileIO(DEC_FILE)

    ppt_orig = ole_orig.openstream("PowerPoint Document").read()
    ppt_dec = ole_dec.openstream("PowerPoint Document").read()

    print(f"原始文件 PPT 大小: {len(ppt_orig)}", flush=True)
    print(f"解密文件 PPT 大小: {len(ppt_dec)}", flush=True)

    if len(ppt_orig) != len(ppt_dec):
        print(f"大小不同! 差异: {len(ppt_dec) - len(ppt_orig)} 字节", flush=True)

    # 找差异
    min_len = min(len(ppt_orig), len(ppt_dec))
    diff_count = 0
    first_diff = -1
    for i in range(min_len):
        if ppt_orig[i] != ppt_dec[i]:
            if first_diff == -1:
                first_diff = i
            diff_count += 1

    print(f"总差异字节数: {diff_count}", flush=True)
    print(f"第一个差异位置: {first_diff}", flush=True)

    ole_orig.close()
    ole_dec.close()
except Exception as e:
    print(f"STEP 8 FAIL: 对比失败: {e}", flush=True)
    import traceback
    traceback.print_exc()

print("DONE", flush=True)
