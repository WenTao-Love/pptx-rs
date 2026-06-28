"""分析参考加密 .ppt 文件的结构。"""
import olefile
import struct

def analyze_reference(path):
    """分析参考加密 .ppt 文件的结构。"""
    print(f"\n=== {path} ===")
    ole = olefile.OleFileIO(path)

    # 列出所有 streams
    print("\n  Streams:")
    for entry in ole.listdir(streams=True):
        name = "/".join(entry)
        size = ole.get_size(name)
        print(f"    {name}: {size} bytes")

    # 读取所有 streams
    print("\n  Reading streams:")
    for entry in ole.listdir(streams=True):
        name = "/".join(entry)
        try:
            data = ole.openstream(name).read()
            print(f"    {name}: read {len(data)} bytes, first 16 bytes: {data[:16].hex()}")
        except Exception as e:
            print(f"    {name}: FAILED to read - {e}")

    # 检查 EncryptedSummaryInfo stream
    print("\n  EncryptedSummaryInfo stream:")
    try:
        data = ole.openstream("EncryptedSummaryInfo").read()
        print(f"    size: {len(data)} bytes")
        print(f"    first 64 bytes: {data[:64].hex()}")
        # 解析 Version
        if len(data) >= 4:
            version = struct.unpack_from("<I", data, 0)[0]
            print(f"    Version: 0x{version:08X}")
        # 解析 EncryptionVersionInfo (vMajor, vMinor)
        if len(data) >= 8:
            v_major = struct.unpack_from("<H", data, 0)[0]
            v_minor = struct.unpack_from("<H", data, 2)[0]
            print(f"    vMajor={v_major}, vMinor={v_minor}")
    except Exception as e:
        print(f"    FAILED: {e}")

    # 检查 Current User stream
    print("\n  Current User stream:")
    try:
        data = ole.openstream("Current User").read()
        print(f"    size: {len(data)} bytes")
        # 解析 CurrentUserAtom
        if len(data) >= 20:
            rec_type = struct.unpack_from("<H", data, 2)[0]
            rec_len = struct.unpack_from("<I", data, 4)[0]
            header_token = struct.unpack_from("<I", data, 12)[0]
            offset_to_current_edit = struct.unpack_from("<I", data, 16)[0]
            print(f"    rec_type=0x{rec_type:04X}, rec_len={rec_len}")
            print(f"    header_token=0x{header_token:08X}")
            print(f"    offsetToCurrentEdit={offset_to_current_edit}")
    except Exception as e:
        print(f"    FAILED: {e}")

    # 检查 PowerPoint Document stream
    print("\n  PowerPoint Document stream:")
    try:
        data = ole.openstream("PowerPoint Document").read()
        print(f"    total size: {len(data)} bytes")
        print(f"    first 64 bytes: {data[:64].hex()}")

        # 读取 CurrentUserAtom 获取 offsetToCurrentEdit
        cu = ole.openstream("Current User").read()
        offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
        print(f"    offsetToCurrentEdit: {offset_to_current_edit}")

        # 读取 UserEditAtom
        ue_offset = offset_to_current_edit
        ue_type = struct.unpack_from("<H", data, ue_offset + 2)[0]
        ue_len = struct.unpack_from("<I", data, ue_offset + 4)[0]
        print(f"    UserEditAtom: type=0x{ue_type:04X}, rec_len={ue_len}")

        # 读取 UserEditAtom 字段
        print(f"      lastSlideIdRef: {struct.unpack_from('<I', data, ue_offset + 8)[0]}")
        print(f"      version: {struct.unpack_from('<H', data, ue_offset + 12)[0]}")
        print(f"      minorVersion: {data[ue_offset + 14]}")
        print(f"      majorVersion: {data[ue_offset + 15]}")
        print(f"      offsetLastEdit: {struct.unpack_from('<I', data, ue_offset + 16)[0]}")
        print(f"      offsetPersistDirectory: {struct.unpack_from('<I', data, ue_offset + 20)[0]}")
        print(f"      docPersistIdRef: {struct.unpack_from('<I', data, ue_offset + 24)[0]}")
        print(f"      persistIdSeed: {struct.unpack_from('<I', data, ue_offset + 28)[0]}")
        print(f"      lastView: {struct.unpack_from('<H', data, ue_offset + 32)[0]}")
        if ue_len >= 32:
            print(f"      encryptSessionPersistIdRef: {struct.unpack_from('<I', data, ue_offset + 36)[0]}")

        # 解析 PersistDirectoryAtom
        pd_offset = struct.unpack_from("<I", data, ue_offset + 20)[0]
        pd_type = struct.unpack_from("<H", data, pd_offset + 2)[0]
        pd_len = struct.unpack_from("<I", data, pd_offset + 4)[0]
        print(f"    PersistDirectoryAtom: type=0x{pd_type:04X}, rec_len={pd_len}")

        # 读取 persist entries
        pd_data_start = pd_offset + 8
        entry_val = struct.unpack_from("<I", data, pd_data_start)[0]
        entry_pid = entry_val & 0xFFFFF
        entry_cpersist = (entry_val >> 20) & 0xFFF
        print(f"    PersistDirectory entry: pid={entry_pid}, cPersist={entry_cpersist}")

        # 读取所有 persist offsets
        persist_offsets = []
        for i in range(entry_cpersist):
            offset = struct.unpack_from("<I", data, pd_data_start + 4 + i * 4)[0]
            persist_offsets.append((entry_pid + i, offset))

        # 找到 CryptSession10Container（最后一个 persist 对象）
        if persist_offsets:
            cs_pid, cs_offset = persist_offsets[-1]
            print(f"\n    CryptSession10Container (persist {cs_pid}):")
            print(f"      offset: {cs_offset}")
            if cs_offset + 8 <= len(data):
                cs_ver_inst = struct.unpack_from("<H", data, cs_offset)[0]
                cs_type = struct.unpack_from("<H", data, cs_offset + 2)[0]
                cs_len = struct.unpack_from("<I", data, cs_offset + 4)[0]
                print(f"      ver_inst: 0x{cs_ver_inst:04X}")
                print(f"      type: 0x{cs_type:04X}")
                print(f"      len: {cs_len}")
                print(f"      data (first 64 bytes): {data[cs_offset + 8:cs_offset + 8 + 64].hex()}")

                # 解析 CryptSession10Container.data
                data_start = cs_offset + 8
                v_major = struct.unpack_from("<H", data, data_start)[0]
                v_minor = struct.unpack_from("<H", data, data_start + 2)[0]
                print(f"      vMajor={v_major}, vMinor={v_minor}")

                outer_flags = struct.unpack_from("<I", data, data_start + 4)[0]
                print(f"      outer_flags: 0x{outer_flags:08X}")

                header_size = struct.unpack_from("<I", data, data_start + 8)[0]
                print(f"      headerSize: {header_size}")

                # EncryptionHeader
                header_start = data_start + 12
                flags = struct.unpack_from("<I", data, header_start)[0]
                size_extra = struct.unpack_from("<I", data, header_start + 4)[0]
                alg_id = struct.unpack_from("<I", data, header_start + 8)[0]
                alg_id_hash = struct.unpack_from("<I", data, header_start + 12)[0]
                key_size = struct.unpack_from("<I", data, header_start + 16)[0]
                provider_type = struct.unpack_from("<I", data, header_start + 20)[0]
                print(f"      EncryptionHeader: flags=0x{flags:08X}, sizeExtra={size_extra}, algId=0x{alg_id:08X}, algIdHash=0x{alg_id_hash:08X}, keySize={key_size}, providerType={provider_type}")

                # CSP name
                csp_start = header_start + 28
                csp_name = data[csp_start:csp_start + header_size - 28].decode("utf-16-le", errors="replace").rstrip("\x00")
                print(f"      CSPName: '{csp_name}'")

                # EncryptionVerifier
                verifier_start = header_start + header_size
                salt_size = struct.unpack_from("<I", data, verifier_start)[0]
                salt = data[verifier_start + 4:verifier_start + 4 + 16]
                encrypted_verifier = data[verifier_start + 20:verifier_start + 20 + 16]
                verifier_hash_size = struct.unpack_from("<I", data, verifier_start + 36)[0]
                encrypted_verifier_hash = data[verifier_start + 40:verifier_start + 40 + 20]
                print(f"      EncryptionVerifier: saltSize={salt_size}, verifierHashSize={verifier_hash_size}")
                print(f"        salt: {salt.hex()}")
                print(f"        encryptedVerifier: {encrypted_verifier.hex()}")
                print(f"        encryptedVerifierHash: {encrypted_verifier_hash.hex()}")

    except Exception as e:
        print(f"    FAILED: {e}")
        import traceback
        traceback.print_exc()

    ole.close()

# 分析参考文件
analyze_reference("_test_out/rc4cryptoapi_password.ppt")
