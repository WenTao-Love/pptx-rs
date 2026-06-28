"""正确解析加密文件的 UserEditAtom 和 PersistDirectoryAtom，验证结构。"""
import io
import os
import struct
import olefile


def main():
    prot_path = None
    for f in os.listdir('_test_out'):
        if f.startswith('protected_') and f.endswith('.ppt') and '.' not in f[len('protected_'):-4]:
            prot_path = '_test_out/' + f
            break

    with open(prot_path, 'rb') as f:
        enc_data = f.read()

    ole = olefile.OleFileIO(io.BytesIO(enc_data))
    ppt = ole.openstream('PowerPoint Document').read()
    cu = ole.openstream('Current User').read()

    # CurrentUserAtom
    offset_to_current_edit = struct.unpack_from('<I', cu, 16)[0]
    header_token = struct.unpack_from('<I', cu, 12)[0]
    print(f'CurrentUser: headerToken=0x{header_token:08X}, offsetToCurrentEdit={offset_to_current_edit}')

    # UserEditAtom (按 msoffcrypto 的字段顺序)
    ue_offset = offset_to_current_edit
    print(f'\nUserEditAtom at offset {ue_offset}:')
    rec_type = struct.unpack_from('<H', ppt, ue_offset + 2)[0]
    rec_len = struct.unpack_from('<I', ppt, ue_offset + 4)[0]
    print(f'  recType=0x{rec_type:04X}, recLen={rec_len}')

    last_slide_id_ref = struct.unpack_from('<I', ppt, ue_offset + 8)[0]
    version = struct.unpack_from('<H', ppt, ue_offset + 12)[0]
    minor_ver = ppt[ue_offset + 14]
    major_ver = ppt[ue_offset + 15]
    offset_last_edit = struct.unpack_from('<I', ppt, ue_offset + 16)[0]
    offset_persist_dir = struct.unpack_from('<I', ppt, ue_offset + 20)[0]
    doc_persist_id_ref = struct.unpack_from('<I', ppt, ue_offset + 24)[0]
    persist_id_seed = struct.unpack_from('<I', ppt, ue_offset + 28)[0]

    print(f'  lastSlideIdRef={last_slide_id_ref}')
    print(f'  version=0x{version:04X}, minorVersion={minor_ver}, majorVersion={major_ver}')
    print(f'  offsetLastEdit={offset_last_edit}')
    print(f'  offsetPersistDirectory={offset_persist_dir}')
    print(f'  docPersistIdRef={doc_persist_id_ref}')
    print(f'  persistIdSeed={persist_id_seed}')

    if rec_len >= 32:
        # lastView (2 bytes) + unused (2 bytes) at offset 32-35
        last_view = struct.unpack_from('<H', ppt, ue_offset + 32)[0]
        print(f'  lastView={last_view}')
        # encryptSessionPersistIdRef at offset 36
        enc_session_pid = struct.unpack_from('<I', ppt, ue_offset + 36)[0]
        print(f'  encryptSessionPersistIdRef={enc_session_pid}')

    # PersistDirectoryAtom
    pd_offset = offset_persist_dir
    print(f'\nPersistDirectoryAtom at offset {pd_offset}:')
    pd_type = struct.unpack_from('<H', ppt, pd_offset + 2)[0]
    pd_len = struct.unpack_from('<I', ppt, pd_offset + 4)[0]
    print(f'  recType=0x{pd_type:04X}, recLen={pd_len}')

    pd_data = ppt[pd_offset + 8:pd_offset + 8 + pd_len]
    pos = 0
    entry_count = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from('<I', pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        print(f'  Entry {entry_count}: persistId={persist_id}, cPersist={c_persist}')
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                off = struct.unpack_from('<I', pd_data, pos)[0]
                pid = persist_id + j
                # 检查这个 offset 是否指向合理的 record
                if off < len(ppt) and off + 8 <= len(ppt):
                    rt = struct.unpack_from('<H', ppt, off + 2)[0]
                    rl = struct.unpack_from('<I', ppt, off + 4)[0]
                    # 如果 record type 是 0x2F14，说明是 CryptSession10Container
                    if rt == 0x2F14:
                        print(f'    persist[{pid}] -> offset {off} (CryptSession10Container, recLen={rl})')
                    else:
                        print(f'    persist[{pid}] -> offset {off}')
                else:
                    print(f'    persist[{pid}] -> offset {off} (超出范围!)')
                pos += 4
        entry_count += 1

    # 检查 CryptSession10Container
    print(f'\n查找 CryptSession10Container:')
    pattern = b'\x0f\x00\x14\x2f'
    idx = ppt.rfind(pattern)
    if idx >= 0:
        cs_len = struct.unpack_from('<I', ppt, idx + 4)[0]
        print(f'  CryptSession10Container at offset {idx}, recLen={cs_len}, total={8+cs_len}')
        print(f'  stream 末尾: {len(ppt)}')
        print(f'  CS 末尾: {idx + 8 + cs_len}')
        if idx + 8 + cs_len == len(ppt):
            print(f'  ✓ CryptSession10Container 正好在 stream 末尾')
        else:
            print(f'  ✗ CryptSession10Container 后面有多余数据!')

    ole.close()


if __name__ == '__main__':
    main()
