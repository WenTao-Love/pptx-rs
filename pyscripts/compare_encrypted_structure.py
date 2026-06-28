"""比较我们生成的加密文件和 msoffcrypto 测试文件的结构差异。

msoffcrypto 测试文件 rc4cryptoapi_password.ppt 是一个可以正常打开的加密文件。
如果我们的文件结构与它有差异，那可能就是问题所在。
"""
import olefile
from struct import unpack


def parse_record_header(data, offset):
    ver_inst = int.from_bytes(data[offset:offset+2], 'little')
    rec_type = int.from_bytes(data[offset+2:offset+4], 'little')
    rec_len = int.from_bytes(data[offset+4:offset+8], 'little')
    return ver_inst & 0x0F, (ver_inst >> 4) & 0x0FFF, rec_type, rec_len


def parse_crypt_session(data, offset):
    """解析 CryptSession10Container。"""
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    if rec_type != 0x2F14:
        return None
    cs_data = data[offset+8:offset+8+rec_len]
    pos = 0
    v_major, v_minor = unpack('<HH', cs_data[pos:pos+4])
    pos += 4
    outer_flags = unpack('<I', cs_data[pos:pos+4])[0]
    pos += 4
    header_size = unpack('<I', cs_data[pos:pos+4])[0]
    pos += 4
    header = cs_data[pos:pos+header_size]
    pos += header_size
    verifier = cs_data[pos:]

    h_pos = 0
    h_flags = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_size_extra = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_alg_id = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_alg_id_hash = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_key_size = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_provider_type = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_reserved1 = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_reserved2 = unpack('<I', header[h_pos:h_pos+4])[0]; h_pos += 4
    h_csp_name = header[h_pos:].decode('utf-16-le', errors='replace')

    v_pos = 0
    v_salt_size = unpack('<I', verifier[v_pos:v_pos+4])[0]; v_pos += 4
    v_salt = verifier[v_pos:v_pos+16]; v_pos += 16
    v_encrypted_verifier = verifier[v_pos:v_pos+16]; v_pos += 16
    v_verifier_hash_size = unpack('<I', verifier[v_pos:v_pos+4])[0]; v_pos += 4
    v_encrypted_verifier_hash = verifier[v_pos:v_pos+20]; v_pos += 20

    return {
        'rec_ver': ver, 'rec_inst': inst, 'rec_type': rec_type, 'rec_len': rec_len,
        'v_major': v_major, 'v_minor': v_minor, 'outer_flags': outer_flags,
        'header_size': header_size,
        'header': {
            'flags': h_flags, 'sizeExtra': h_size_extra, 'algId': h_alg_id,
            'algIdHash': h_alg_id_hash, 'keySize': h_key_size,
            'providerType': h_provider_type, 'reserved1': h_reserved1,
            'reserved2': h_reserved2, 'cspName': h_csp_name,
        },
        'verifier': {
            'saltSize': v_salt_size, 'salt': v_salt,
            'encryptedVerifier': v_encrypted_verifier,
            'verifierHashSize': v_verifier_hash_size,
            'encryptedVerifierHash': v_encrypted_verifier_hash,
        },
        'total_data_len': rec_len,
    }


def analyze_file(path, label):
    print(f"\n{'='*60}")
    print(f"分析: {label}")
    print(f"路径: {path}")
    print(f"{'='*60}")

    ole = olefile.OleFileIO(path)

    # 列出所有 streams
    print(f"\n--- Streams ---")
    for streams_path in ole.listdir(streams=True, storages=False):
        path_str = '/'.join(streams_path)
        size = ole.get_size(streams_path)
        print(f"  /{path_str}: {size} bytes")

    # 检查 Current User
    cu = ole.openstream('Current User').read()
    print(f"\n--- Current User ---")
    print(f"  size: {len(cu)} bytes")
    header_token = int.from_bytes(cu[12:16], 'little')
    offset_to_current_edit = int.from_bytes(cu[16:20], 'little')
    print(f"  headerToken: {header_token:#010x}")
    print(f"  offsetToCurrentEdit: {offset_to_current_edit} ({offset_to_current_edit:#x})")

    # 检查 PowerPoint Document
    ppt = ole.openstream('PowerPoint Document').read()
    print(f"\n--- PowerPoint Document ---")
    print(f"  size: {len(ppt)} bytes")

    # 读取 UserEditAtom
    ue_ver, ue_inst, ue_type, ue_len = parse_record_header(ppt, offset_to_current_edit)
    print(f"\n  UserEditAtom at {offset_to_current_edit:#x}:")
    print(f"    ver={ue_ver}, inst={ue_inst}, type={ue_type:#06x}, recLen={ue_len}")

    if ue_type == 0x0FF5:
        # 读取 UserEditAtom 字段
        last_slide_id_ref = int.from_bytes(ppt[offset_to_current_edit+8:offset_to_current_edit+12], 'little')
        version = int.from_bytes(ppt[offset_to_current_edit+12:offset_to_current_edit+14], 'little')
        minor_major = int.from_bytes(ppt[offset_to_current_edit+14:offset_to_current_edit+16], 'little')
        offset_last_edit = int.from_bytes(ppt[offset_to_current_edit+16:offset_to_current_edit+20], 'little')
        offset_persist_dir = int.from_bytes(ppt[offset_to_current_edit+20:offset_to_current_edit+24], 'little')
        doc_persist_id_ref = int.from_bytes(ppt[offset_to_current_edit+24:offset_to_current_edit+28], 'little')
        max_persist_written = int.from_bytes(ppt[offset_to_current_edit+28:offset_to_current_edit+32], 'little')
        print(f"    lastSlideIdRef: {last_slide_id_ref}")
        print(f"    version: {version}")
        print(f"    minorVersion/majorVersion: {minor_major:#06x}")
        print(f"    offsetLastEdit: {offset_last_edit} ({offset_last_edit:#x})")
        print(f"    offsetPersistDirectory: {offset_persist_dir} ({offset_persist_dir:#x})")
        print(f"    docPersistIdRef: {doc_persist_id_ref}")
        print(f"    maxPersistWritten: {max_persist_written}")

        if ue_len >= 32:
            last_view = int.from_bytes(ppt[offset_to_current_edit+32:offset_to_current_edit+34], 'little')
            unused = int.from_bytes(ppt[offset_to_current_edit+34:offset_to_current_edit+36], 'little')
            encrypt_session_pid_ref = int.from_bytes(ppt[offset_to_current_edit+36:offset_to_current_edit+40], 'little')
            print(f"    lastView: {last_view}")
            print(f"    unused: {unused}")
            print(f"    encryptSessionPersistIdRef: {encrypt_session_pid_ref}")
        elif ue_len >= 28:
            last_view = int.from_bytes(ppt[offset_to_current_edit+32:offset_to_current_edit+34], 'little')
            unused = int.from_bytes(ppt[offset_to_current_edit+34:offset_to_current_edit+36], 'little')
            print(f"    lastView: {last_view}")
            print(f"    unused: {unused}")
            print(f"    encryptSessionPersistIdRef: (none)")

        # 读取 PersistDirectoryAtom
        pd_ver, pd_inst, pd_type, pd_len = parse_record_header(ppt, offset_persist_dir)
        print(f"\n  PersistDirectoryAtom at {offset_persist_dir:#x}:")
        print(f"    ver={pd_ver}, inst={pd_inst}, type={pd_type:#06x}, recLen={pd_len}")

        # 解析 persist entries
        pd_data = ppt[offset_persist_dir+8:offset_persist_dir+8+pd_len]
        pos = 0
        entries = []
        while pos + 4 <= len(pd_data):
            entry_val = int.from_bytes(pd_data[pos:pos+4], 'little')
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF
            pos += 4
            print(f"    PersistDirectoryEntry: persistId={persist_id}, cPersist={c_persist}")
            for j in range(c_persist):
                if pos + 4 <= len(pd_data):
                    persist_offset = int.from_bytes(pd_data[pos:pos+4], 'little')
                    entries.append((persist_id + j, persist_offset))
                    pos += 4

        # 找到 CryptSession10Container
        if ue_len >= 32:
            encrypt_session_pid_ref = int.from_bytes(ppt[offset_to_current_edit+36:offset_to_current_edit+40], 'little')
            for pid, poff in entries:
                if pid == encrypt_session_pid_ref:
                    print(f"\n  CryptSession10Container at {poff:#x} (persistId={pid}):")
                    cs_info = parse_crypt_session(ppt, poff)
                    if cs_info:
                        print(f"    rec_ver={cs_info['rec_ver']}, rec_inst={cs_info['rec_inst']}, rec_type={cs_info['rec_type']:#06x}, rec_len={cs_info['rec_len']}")
                        print(f"    vMajor={cs_info['v_major']}, vMinor={cs_info['v_minor']}")
                        print(f"    outer_flags={cs_info['outer_flags']:#010x}")
                        print(f"    header_size={cs_info['header_size']}")
                        print(f"    EncryptionHeader:")
                        for k, v in cs_info['header'].items():
                            if k == 'cspName':
                                print(f"      {k}: {v!r}")
                            else:
                                print(f"      {k}: {v:#010x}")
                        print(f"    EncryptionVerifier:")
                        print(f"      saltSize: {cs_info['verifier']['saltSize']}")
                        print(f"      salt: {cs_info['verifier']['salt'].hex()}")
                        print(f"      encryptedVerifier: {cs_info['verifier']['encryptedVerifier'].hex()}")
                        print(f"      verifierHashSize: {cs_info['verifier']['verifierHashSize']}")
                        print(f"      encryptedVerifierHash: {cs_info['verifier']['encryptedVerifierHash'].hex()}")
                    break

    ole.close()


def main():
    # msoffcrypto 测试文件（可以正常打开的加密文件）
    ref_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\rc4cryptoapi_password.ppt"
    # 我们生成的加密文件
    our_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt"

    analyze_file(ref_path, "msoffcrypto 测试文件（参考）")
    analyze_file(our_path, "我们生成的加密文件")


if __name__ == "__main__":
    main()
