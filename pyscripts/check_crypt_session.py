"""检查 PowerPoint Document stream 末尾的 CryptSession10Container。"""
import olefile
import struct

def check_crypt_session(path):
    """检查 CryptSession10Container 的结构。"""
    print(f"\n=== {path} ===")
    ole = olefile.OleFileIO(path)

    ppt = ole.openstream("PowerPoint Document").read()
    print(f"  PowerPoint Document stream size: {len(ppt)} bytes")

    # 读取 CurrentUserAtom 获取 offsetToCurrentEdit
    cu = ole.openstream("Current User").read()
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    print(f"  offsetToCurrentEdit: {offset_to_current_edit}")

    # 读取 UserEditAtom
    ue_offset = offset_to_current_edit
    ue_type = struct.unpack_from("<H", ppt, ue_offset + 2)[0]
    ue_len = struct.unpack_from("<I", ppt, ue_offset + 4)[0]
    print(f"  UserEditAtom: type=0x{ue_type:04X}, rec_len={ue_len}")

    # 读取 offsetPersistDirectory
    pd_offset = struct.unpack_from("<I", ppt, ue_offset + 20)[0]
    print(f"  offsetPersistDirectory: {pd_offset}")

    # 读取 maxPersistWritten
    max_persist = struct.unpack_from("<I", ppt, ue_offset + 28)[0]
    print(f"  maxPersistWritten: {max_persist}")

    # 读取 encryptSessionPersistIdRef (如果存在)
    if ue_len >= 32:
        encrypt_session_ref = struct.unpack_from("<I", ppt, ue_offset + 36)[0]
        print(f"  encryptSessionPersistIdRef: {encrypt_session_ref}")

    # 解析 PersistDirectoryAtom
    pd_type = struct.unpack_from("<H", ppt, pd_offset + 2)[0]
    pd_len = struct.unpack_from("<I", ppt, pd_offset + 4)[0]
    print(f"  PersistDirectoryAtom: type=0x{pd_type:04X}, rec_len={pd_len}")

    # 读取 persist entries
    pd_data_start = pd_offset + 8
    entry_val = struct.unpack_from("<I", ppt, pd_data_start)[0]
    entry_pid = entry_val & 0xFFFFF
    entry_cpersist = (entry_val >> 20) & 0xFFF
    print(f"  PersistDirectory entry: pid={entry_pid}, cPersist={entry_cpersist}")

    # 读取所有 persist offsets
    persist_offsets = []
    for i in range(entry_cpersist):
        offset = struct.unpack_from("<I", ppt, pd_data_start + 4 + i * 4)[0]
        persist_offsets.append((entry_pid + i, offset))
        print(f"    persist[{entry_pid + i}] = offset {offset}")

    # 找到 CryptSession10Container（最后一个 persist 对象）
    if persist_offsets:
        cs_pid, cs_offset = persist_offsets[-1]
        print(f"\n  CryptSession10Container (persist {cs_pid}):")
        print(f"    offset: {cs_offset}")
        if cs_offset + 8 <= len(ppt):
            cs_ver_inst = struct.unpack_from("<H", ppt, cs_offset)[0]
            cs_type = struct.unpack_from("<H", ppt, cs_offset + 2)[0]
            cs_len = struct.unpack_from("<I", ppt, cs_offset + 4)[0]
            print(f"    ver_inst: 0x{cs_ver_inst:04X}")
            print(f"    type: 0x{cs_type:04X}")
            print(f"    len: {cs_len}")
            print(f"    data (first 64 bytes): {ppt[cs_offset + 8:cs_offset + 8 + 64].hex()}")

            # 解析 CryptSession10Container.data
            data_start = cs_offset + 8
            v_major = struct.unpack_from("<H", ppt, data_start)[0]
            v_minor = struct.unpack_from("<H", ppt, data_start + 2)[0]
            print(f"    vMajor={v_major}, vMinor={v_minor}")

            outer_flags = struct.unpack_from("<I", ppt, data_start + 4)[0]
            print(f"    outer_flags: 0x{outer_flags:08X}")

            header_size = struct.unpack_from("<I", ppt, data_start + 8)[0]
            print(f"    headerSize: {header_size}")

            # EncryptionHeader
            header_start = data_start + 12
            flags = struct.unpack_from("<I", ppt, header_start)[0]
            size_extra = struct.unpack_from("<I", ppt, header_start + 4)[0]
            alg_id = struct.unpack_from("<I", ppt, header_start + 8)[0]
            alg_id_hash = struct.unpack_from("<I", ppt, header_start + 12)[0]
            key_size = struct.unpack_from("<I", ppt, header_start + 16)[0]
            provider_type = struct.unpack_from("<I", ppt, header_start + 20)[0]
            print(f"    EncryptionHeader: flags=0x{flags:08X}, sizeExtra={size_extra}, algId=0x{alg_id:08X}, algIdHash=0x{alg_id_hash:08X}, keySize={key_size}, providerType={provider_type}")

            # CSP name
            csp_start = header_start + 28
            csp_name = ppt[csp_start:csp_start + header_size - 28].decode("utf-16-le", errors="replace").rstrip("\x00")
            print(f"    CSPName: '{csp_name}'")

            # EncryptionVerifier
            verifier_start = header_start + header_size
            salt_size = struct.unpack_from("<I", ppt, verifier_start)[0]
            salt = ppt[verifier_start + 4:verifier_start + 4 + 16]
            encrypted_verifier = ppt[verifier_start + 20:verifier_start + 20 + 16]
            verifier_hash_size = struct.unpack_from("<I", ppt, verifier_start + 36)[0]
            encrypted_verifier_hash = ppt[verifier_start + 40:verifier_start + 40 + 20]
            print(f"    EncryptionVerifier: saltSize={salt_size}, verifierHashSize={verifier_hash_size}")
            print(f"      salt: {salt.hex()}")
            print(f"      encryptedVerifier: {encrypted_verifier.hex()}")
            print(f"      encryptedVerifierHash: {encrypted_verifier_hash.hex()}")

    ole.close()

# 检查所有文件
files = [
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        check_crypt_session(f)
    except Exception as e:
        print(f"  分析失败: {e}")
        import traceback
        traceback.print_exc()
