#!/usr/bin/env python3
"""比较原始文件和加水印文件的 PowerPoint Document stream，找到差异位置。"""
import olefile

def read_ppt(path):
    ole = olefile.OleFileIO(path)
    ppt = ole.openstream("PowerPoint Document").read()
    ole.close()
    return ppt

orig = read_ppt("_test/心理账户理论.ppt")
wm = read_ppt("_test_out/wm_心理账户理论.ppt")

print(f"orig size: {len(orig)}, wm size: {len(wm)}, diff: {len(wm)-len(orig)}")

# 对齐比较
min_len = min(len(orig), len(wm))
first_diff = None
for i in range(min_len):
    if orig[i] != wm[i]:
        first_diff = i
        break

if first_diff is None:
    print("前 {} 字节相同".format(min_len))
else:
    print(f"第一个不同字节 @ {first_diff}")
    print(f"  orig: {orig[max(0,first_diff-16):first_diff+64].hex()}")
    print(f"  wm:   {wm[max(0,first_diff-16):first_diff+64].hex()}")

# 从后往前找最后一个不同
last_diff = None
for i in range(min_len - 1, -1, -1):
    if orig[i] != wm[i]:
        last_diff = i
        break
print(f"最后一个不同字节 @ {last_diff}")

# 统计不同字节数
diff_count = sum(1 for i in range(min_len) if orig[i] != wm[i])
print(f"总不同字节数: {diff_count}")
