"""比较解密后的文件与原始文件，检查加密/解密过程是否正确。"""
import os
import io
import olefile
import sys

def read_stream(fpath, stream_name):
    with open(fpath, 'rb') as f:
        data = f.read()
    ole = olefile.OleFileIO(io.BytesIO(data))
    for entry in ole.listdir():
        name = '/'.join(entry)
        if name == stream_name:
            return ole.openstream(name).read()
    return None

def compare_files(original_path, decrypted_path, label):
    print(f'=== {label} ===')
    print(f'  原始: {original_path}')
    print(f'  解密: {decrypted_path}')

    if not os.path.exists(original_path):
        print(f'  原始文件不存在')
        return False
    if not os.path.exists(decrypted_path):
        print(f'  解密文件不存在')
        return False

    # 比较 PowerPoint Document stream
    orig_ppt = read_stream(original_path, 'PowerPoint Document')
    decr_ppt = read_stream(decrypted_path, 'PowerPoint Document')

    if orig_ppt is None:
        print(f'  原始文件找不到 PowerPoint Document stream')
        return False
    if decr_ppt is None:
        print(f'  解密文件找不到 PowerPoint Document stream')
        return False

    print(f'  原始 PowerPoint Document: {len(orig_ppt)} bytes')
    print(f'  解密 PowerPoint Document: {len(decr_ppt)} bytes')

    if len(orig_ppt) != len(decr_ppt):
        print(f'  大小不同！差异: {len(decr_ppt) - len(orig_ppt)} bytes')
    else:
        # 比较内容
        diff_count = 0
        first_diff = -1
        for i in range(len(orig_ppt)):
            if orig_ppt[i] != decr_ppt[i]:
                diff_count += 1
                if first_diff < 0:
                    first_diff = i
        if diff_count == 0:
            print(f'  PowerPoint Document stream 完全相同 ✓')
        else:
            print(f'  PowerPoint Document stream 有 {diff_count} 字节不同')
            print(f'  第一个差异位置: {first_diff}')
            # 显示差异上下文
            start = max(0, first_diff - 8)
            end = min(len(orig_ppt), first_diff + 24)
            print(f'  原始 [{start}:{end}]: {orig_ppt[start:end].hex()}')
            print(f'  解密 [{start}:{end}]: {decr_ppt[start:end].hex()}')

    # 比较 Current User stream
    orig_cu = read_stream(original_path, 'Current User')
    decr_cu = read_stream(decrypted_path, 'Current User')
    if orig_cu and decr_cu:
        print(f'  原始 Current User: {len(orig_cu)} bytes')
        print(f'  解密 Current User: {len(decr_cu)} bytes')
        if orig_cu == decr_cu:
            print(f'  Current User stream 完全相同 ✓')
        else:
            diff_count = sum(1 for a, b in zip(orig_cu, decr_cu) if a != b)
            print(f'  Current User stream 有 {diff_count} 字节不同')
            print(f'  原始: {orig_cu.hex()}')
            print(f'  解密: {decr_cu.hex()}')

    # 比较 Pictures stream
    orig_pic = read_stream(original_path, 'Pictures')
    decr_pic = read_stream(decrypted_path, 'Pictures')
    if orig_pic and decr_pic:
        print(f'  原始 Pictures: {len(orig_pic)} bytes')
        print(f'  解密 Pictures: {len(decr_pic)} bytes')
        if orig_pic == decr_pic:
            print(f'  Pictures stream 完全相同 ✓')
        else:
            diff_count = sum(1 for a, b in zip(orig_pic, decr_pic) if a != b)
            print(f'  Pictures stream 有 {diff_count} 字节不同')
            if diff_count > 0:
                # 找第一个差异
                for i in range(min(len(orig_pic), len(decr_pic))):
                    if orig_pic[i] != decr_pic[i]:
                        start = max(0, i - 8)
                        end = min(len(orig_pic), i + 24)
                        print(f'  第一个差异位置: {i}')
                        print(f'  原始 [{start}:{end}]: {orig_pic[start:end].hex()}')
                        print(f'  解密 [{start}:{end}]: {decr_pic[start:end].hex()}')
                        break
    elif orig_pic and not decr_pic:
        print(f'  原始有 Pictures stream，解密文件没有！')
    elif not orig_pic and decr_pic:
        print(f'  原始没有 Pictures stream，解密文件有！')

    print()
    return True

if __name__ == '__main__':
    # 纯加密文件：解密后应与原始文件相同
    compare_files(
        '_test/心理账户理论.ppt',
        '_test_out/protected_心理账户理论.ppt.verified.ppt',
        '纯加密 vs 原始'
    )

    # 水印+加密文件：解密后应与纯水印文件相同
    compare_files(
        '_test_out/wm_心理账户理论.ppt',
        '_test_out/wm_protected_心理账户理论.ppt.verified.ppt',
        '水印+加密 vs 纯水印'
    )
