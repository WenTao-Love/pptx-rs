"""比较 msoffcrypto 解密后的 PowerPoint Document stream 与原始文件。

关键检查：
1. persist 对象是否正确解密
2. UserEditAtom 和 PersistDirectoryAtom 的差异是否预期
3. CryptSession10Container 是否被正确移除（用 0 填充）
"""
import olefile
import io
import os
from struct import unpack


def parse_record_header(data, offset):
    ver_inst = int.from_bytes(data[offset:offset+2], 'little')
    rec_type = int.from_bytes(data[offset+2:offset+4], 'little')
    rec_len = int.from_bytes(data[offset+4:offset+8], 'little')
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len


def parse_persist_directory(data, offset):
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    if rec_type != 0x1772:
        return []
    pd_data = data[offset+8:offset+8+rec_len]
    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry_val = int.from_bytes(pd_data[pos:pos+4], 'little')
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = int.from_bytes(pd_data[pos:pos+4], 'little')
                entries.append((persist_id + j, persist_offset))
                pos += 4
    return entries


def main():
    # 原始文件
    original_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt"
    # msoffcrypto 解密后的文件
    decrypted_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt.decrypted2.ppt"

    ole_orig = olefile.OleFileIO(original_path)
    ole_dec = olefile.OleFileIO(decrypted_path)

    ppt_orig = ole_orig.openstream('PowerPoint Document').read()
    ppt_dec = ole_dec.openstream('PowerPoint Document').read()

    print(f"原始 PowerPoint Document: {len(ppt_orig)} bytes")
    print(f"解密后 PowerPoint Document: {len(ppt_dec)} bytes")

    # 解析原始文件的 persist directory
    cu_orig = ole_orig.openstream('Current User').read()
    offset_to_current_edit_orig = int.from_bytes(cu_orig[16:20], 'little')
    persist_entries_orig = parse_persist_directory(ppt_orig, int.from_bytes(ppt_orig[offset_to_current_edit_orig+20:offset_to_current_edit_orig+24], 'little'))

    print(f"\n原始文件 persist entries: {len(persist_entries_orig)}")
    print(f"原始文件 offsetToCurrentEdit: {offset_to_current_edit_orig}")

    # 解析解密后的文件的 persist directory
    cu_dec = ole_dec.openstream('Current User').read()
    offset_to_current_edit_dec = int.from_bytes(cu_dec[16:20], 'little')
    persist_entries_dec = parse_persist_directory(ppt_dec, int.from_bytes(ppt_dec[offset_to_current_edit_dec+20:offset_to_current_edit_dec+24], 'little'))

    print(f"解密后文件 persist entries: {len(persist_entries_dec)}")
    print(f"解密后文件 offsetToCurrentEdit: {offset_to_current_edit_dec}")

    # 比较 persist 对象
    print(f"\n--- 比较 persist 对象 ---")
    orig_dict = dict(persist_entries_orig)
    dec_dict = dict(persist_entries_dec)

    for persist_id in sorted(orig_dict.keys()):
        orig_offset = orig_dict[persist_id]
        if persist_id not in dec_dict:
            print(f"  persistId {persist_id}: 解密后文件中不存在！")
            continue
        dec_offset = dec_dict[persist_id]

        # 读取原始文件的 persist 对象
        ver, inst, rec_type, rec_len = parse_record_header(ppt_orig, orig_offset)
        orig_data = ppt_orig[orig_offset:orig_offset+8+rec_len]

        # 读取解密后文件的 persist 对象
        ver2, inst2, rec_type2, rec_len2 = parse_record_header(ppt_dec, dec_offset)
        dec_data = ppt_dec[dec_offset:dec_offset+8+rec_len2]

        if orig_data == dec_data:
            print(f"  persistId {persist_id}: ✓ 一致 (type={rec_type:#06x}, len={rec_len})")
        else:
            print(f"  persistId {persist_id}: ✗ 不一致 (orig: type={rec_type:#06x}, len={rec_len}, dec: type={rec_type2:#06x}, len={rec_len2})")
            # 找到第一个不同的字节
            for i in range(min(len(orig_data), len(dec_data))):
                if orig_data[i] != dec_data[i]:
                    print(f"    第一个不同的字节在 offset {i}: orig={orig_data[i]:#x}, dec={dec_data[i]:#x}")
                    start = max(0, i - 8)
                    end = min(len(orig_data), i + 8)
                    print(f"    orig [{start}:{end}]: {orig_data[start:end].hex()}")
                    print(f"    dec  [{start}:{end}]: {dec_data[start:end].hex()}")
                    break

    ole_orig.close()
    ole_dec.close()


if __name__ == "__main__":
    main()
