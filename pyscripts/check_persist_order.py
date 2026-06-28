"""检查 persist 对象的 offset 顺序和 record header 中的 recLen 是否一致。

关键问题：msoffcrypto 解密时使用 directory_items[i+1][1] - offset - 8 计算 recLen，
而不是 record header 中的 recLen。如果 persist 对象不是按 offset 排序的，
这两种计算方式会得到不同的结果，导致解密范围错误。
"""
import olefile
import struct
import sys

def parse_record_header(data, offset):
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def check_persist_order(filepath):
    """检查 persist 对象的 offset 顺序。"""
    print(f'\n{"="*70}')
    print(f'检查 persist 对象顺序: {filepath}')
    print(f'{"="*70}')

    with open(filepath, 'rb') as f:
        ole = olefile.OleFileIO(f)
        cu = ole.openstream('Current User').read()
        ppt = ole.openstream('PowerPoint Document').read()

        # 读取未加密的原始文件
        header_token = struct.unpack_from('<I', cu, 12)[0]
        is_encrypted = (header_token == 0xF3D1C4DF)
        print(f'  headerToken={header_token:#010x}, encrypted={is_encrypted}')

        offset_to_current_edit = struct.unpack_from('<I', cu, 16)[0]
        ue_offset = offset_to_current_edit

        # 读取 UserEditAtom
        ue_rec_len = struct.unpack_from('<I', ppt, ue_offset + 4)[0]
        offset_persist_dir = struct.unpack_from('<I', ppt, ue_offset + 20)[0]

        # 解析 PersistDirectoryAtom
        pd_ver, pd_inst, pd_type, pd_len = parse_record_header(ppt, offset_persist_dir)
        pd_data = ppt[offset_persist_dir + 8 : offset_persist_dir + 8 + pd_len]

        pos = 0
        entries = []
        while pos + 4 <= len(pd_data):
            entry_val = struct.unpack_from('<I', pd_data, pos)[0]
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF
            pos += 4
            for j in range(c_persist):
                if pos + 4 <= len(pd_data):
                    poff = struct.unpack_from('<I', pd_data, pos)[0]
                    entries.append((persist_id + j, poff))
                    pos += 4

        # 按 persistId 顺序检查
        print(f'\n  --- 按 persistId 顺序 ---')
        print(f'  {"pid":>4} {"offset":>10} {"next_off":>10} {"recLen(hdr)":>12} {"recLen(calc)":>12} {"match":>6}')
        for i, (pid, poff) in enumerate(entries):
            if poff + 8 > len(ppt):
                continue
            ver, inst, rec_type, rec_len = parse_record_header(ppt, poff)

            # 计算下一个 persist 对象的 offset（按 persistId 顺序）
            if i + 1 < len(entries):
                next_off = entries[i + 1][1]
                rec_len_calc = next_off - poff - 8
            else:
                next_off = None
                rec_len_calc = None

            match = ''
            if rec_len_calc is not None:
                if rec_len == rec_len_calc:
                    match = 'YES'
                else:
                    match = f'NO (diff={rec_len_calc - rec_len})'

            if is_encrypted and rec_type not in [0x2F14, 0x0FF5, 0x1772]:
                # 加密文件中，record header 被加密，recLen 不可靠
                print(f'  {pid:>4} {poff:>10} {str(next_off):>10} {rec_len:>12} {str(rec_len_calc):>12} {"(encrypted)":>6}')
            else:
                print(f'  {pid:>4} {poff:>10} {str(next_off):>10} {rec_len:>12} {str(rec_len_calc):>12} {match:>6}')

        # 按 offset 排序后检查
        sorted_entries = sorted(entries, key=lambda x: x[1])
        print(f'\n  --- 按 offset 排序后 ---')
        print(f'  {"pid":>4} {"offset":>10} {"next_off":>10} {"recLen(hdr)":>12} {"recLen(calc)":>12} {"match":>6}')
        for i, (pid, poff) in enumerate(sorted_entries):
            if poff + 8 > len(ppt):
                continue
            ver, inst, rec_type, rec_len = parse_record_header(ppt, poff)

            # 计算下一个 persist 对象的 offset（按 offset 排序后）
            if i + 1 < len(sorted_entries):
                next_off = sorted_entries[i + 1][1]
                rec_len_calc = next_off - poff - 8
            else:
                next_off = None
                rec_len_calc = None

            match = ''
            if rec_len_calc is not None:
                if rec_len == rec_len_calc:
                    match = 'YES'
                else:
                    match = f'NO (diff={rec_len_calc - rec_len})'

            if is_encrypted and rec_type not in [0x2F14, 0x0FF5, 0x1772]:
                print(f'  {pid:>4} {poff:>10} {str(next_off):>10} {rec_len:>12} {str(rec_len_calc):>12} {"(encrypted)":>6}')
            else:
                print(f'  {pid:>4} {poff:>10} {str(next_off):>10} {rec_len:>12} {str(rec_len_calc):>12} {match:>6}')

        ole.close()

if __name__ == '__main__':
    # 检查参考文件
    check_persist_order('_test_out/rc4cryptoapi_password.ppt')
    # 检查我们的文件（加密前）
    import os
    for f in os.listdir('_test'):
        if f.endswith('.ppt'):
            check_persist_order(f'_test/{f}')
            break
