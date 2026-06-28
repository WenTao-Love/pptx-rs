"""深度比较参考加密文件 rc4cryptoapi_password.ppt 与我们生成的加密文件的结构差异。

重点检查：
1. CurrentUserAtom 字段
2. UserEditAtom 字段
3. PersistDirectoryAtom 结构
4. CryptSession10Container 格式
5. OLE2 容器中的 streams 列表
"""
import olefile
import struct
import sys
import os

def parse_record_header(data, offset):
    """解析 8 字节 record header。"""
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def analyze_ppt(filepath):
    """分析加密 PPT 文件的结构。"""
    print(f'\n{"="*70}')
    print(f'分析: {filepath}')
    print(f'{"="*70}')

    if not os.path.exists(filepath):
        print(f'  文件不存在!')
        return

    with open(filepath, 'rb') as f:
        ole = olefile.OleFileIO(f)

        # 1. 列出所有 streams
        print(f'\n--- OLE2 Streams ---')
        for entry in ole.listdir():
            path = '/'.join(entry)
            size = ole.get_size(path)
            print(f'  {path}: {size} bytes')

        # 2. 解析 Current User
        print(f'\n--- Current User Atom ---')
        cu = ole.openstream('Current User')
        cu_data = cu.read()
        ver, inst, rec_type, rec_len = parse_record_header(cu_data, 0)
        print(f'  recVer={ver:#x}, recInstance={inst:#x}, recType={rec_type:#06x}, recLen={rec_len}')
        size = struct.unpack_from('<I', cu_data, 8)[0]
        header_token = struct.unpack_from('<I', cu_data, 12)[0]
        offset_to_current_edit = struct.unpack_from('<I', cu_data, 16)[0]
        print(f'  size={size:#x}')
        print(f'  headerToken={header_token:#010x} (期望 0xF3D1C4DF)')
        print(f'  offsetToCurrentEdit={offset_to_current_edit}')

        # 3. 解析 PowerPoint Document
        print(f'\n--- PowerPoint Document ---')
        ppt = ole.openstream('PowerPoint Document')
        ppt_data = ppt.read()
        print(f'  stream 大小: {len(ppt_data)} bytes')

        # 4. 解析 UserEditAtom
        ue_offset = offset_to_current_edit
        print(f'\n--- UserEditAtom (offset={ue_offset}) ---')
        ver, inst, rec_type, rec_len = parse_record_header(ppt_data, ue_offset)
        print(f'  recVer={ver:#x}, recInstance={inst:#x}, recType={rec_type:#06x}, recLen={rec_len}')
        print(f'  recLen 期望: 32 (0x20) 加密, 28 (0x1C) 未加密')

        if rec_len >= 28:
            last_slide_id_ref = struct.unpack_from('<I', ppt_data, ue_offset + 8)[0]
            version = struct.unpack_from('<H', ppt_data, ue_offset + 12)[0]
            minor_major = struct.unpack_from('<BB', ppt_data, ue_offset + 14)
            offset_last_edit = struct.unpack_from('<I', ppt_data, ue_offset + 16)[0]
            offset_persist_dir = struct.unpack_from('<I', ppt_data, ue_offset + 20)[0]
            doc_persist_id_ref = struct.unpack_from('<I', ppt_data, ue_offset + 24)[0]
            persist_id_seed = struct.unpack_from('<I', ppt_data, ue_offset + 28)[0]
            last_view = struct.unpack_from('<H', ppt_data, ue_offset + 32)[0]
            unused = struct.unpack_from('<H', ppt_data, ue_offset + 34)[0]
            print(f'  lastSlideIdRef={last_slide_id_ref}')
            print(f'  version={version:#x}')
            print(f'  minorVersion={minor_major[0]}, majorVersion={minor_major[1]}')
            print(f'  offsetLastEdit={offset_last_edit}')
            print(f'  offsetPersistDirectory={offset_persist_dir}')
            print(f'  docPersistIdRef={doc_persist_id_ref}')
            print(f'  persistIdSeed={persist_id_seed}')
            print(f'  lastView={last_view}')
            print(f'  unused={unused:#x}')
            if rec_len == 32:
                encrypt_session_pid = struct.unpack_from('<I', ppt_data, ue_offset + 36)[0]
                print(f'  encryptSessionPersistIdRef={encrypt_session_pid}')

        # 5. 解析 PersistDirectoryAtom
        pd_offset = offset_persist_dir
        print(f'\n--- PersistDirectoryAtom (offset={pd_offset}) ---')
        ver, inst, rec_type, rec_len = parse_record_header(ppt_data, pd_offset)
        print(f'  recVer={ver:#x}, recInstance={inst:#x}, recType={rec_type:#06x}, recLen={rec_len}')

        # 解析 persist entries
        pd_data = ppt_data[pd_offset + 8 : pd_offset + 8 + rec_len]
        pos = 0
        entries = []
        while pos + 4 <= len(pd_data):
            entry_val = struct.unpack_from('<I', pd_data, pos)[0]
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF
            pos += 4
            print(f'  PersistDirectoryEntry: persistId={persist_id}, cPersist={c_persist}')
            for j in range(c_persist):
                if pos + 4 <= len(pd_data):
                    poff = struct.unpack_from('<I', pd_data, pos)[0]
                    entries.append((persist_id + j, poff))
                    if j < 3 or j >= c_persist - 1:
                        print(f'    [{persist_id + j}] offset={poff}')
                    elif j == 3:
                        print(f'    ... (省略中间条目)')
                    pos += 4

        print(f'  总 persist 对象数: {len(entries)}')

        # 6. 解析 CryptSession10Container
        if rec_len == 32:
            encrypt_session_pid = struct.unpack_from('<I', ppt_data, ue_offset + 36)[0]
            if encrypt_session_pid < len(entries):
                cs_offset = entries[encrypt_session_pid][1]
                print(f'\n--- CryptSession10Container (offset={cs_offset}) ---')
                ver, inst, rec_type, rec_len = parse_record_header(ppt_data, cs_offset)
                print(f'  recVer={ver:#x}, recInstance={inst:#x}, recType={rec_type:#06x}, recLen={rec_len}')
                print(f'  recType 期望: 0x2F14')

                # 解析 CryptSession10Container.data
                cs_data = ppt_data[cs_offset + 8 : cs_offset + 8 + rec_len]
                if len(cs_data) >= 12:
                    v_major, v_minor = struct.unpack_from('<HH', cs_data, 0)
                    outer_flags = struct.unpack_from('<I', cs_data, 4)[0]
                    header_size = struct.unpack_from('<I', cs_data, 8)[0]
                    print(f'  EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}')
                    print(f'  outer_flags={outer_flags:#010x} (期望 0x0000000C)')
                    print(f'  headerSize={header_size}')

                    if len(cs_data) >= 12 + header_size:
                        # EncryptionHeader
                        eh_off = 12
                        flags = struct.unpack_from('<I', cs_data, eh_off)[0]
                        size_extra = struct.unpack_from('<I', cs_data, eh_off + 4)[0]
                        alg_id = struct.unpack_from('<I', cs_data, eh_off + 8)[0]
                        alg_id_hash = struct.unpack_from('<I', cs_data, eh_off + 12)[0]
                        key_size = struct.unpack_from('<I', cs_data, eh_off + 16)[0]
                        provider_type = struct.unpack_from('<I', cs_data, eh_off + 20)[0]
                        reserved1 = struct.unpack_from('<I', cs_data, eh_off + 24)[0]
                        reserved2 = struct.unpack_from('<I', cs_data, eh_off + 28)[0]
                        csp_name = cs_data[eh_off + 32:eh_off + header_size]
                        try:
                            csp_str = csp_name.decode('utf-16-le').rstrip('\x00')
                        except:
                            csp_str = repr(csp_name[:40])
                        print(f'  EncryptionHeader:')
                        print(f'    flags={flags:#010x}')
                        print(f'    sizeExtra={size_extra}')
                        print(f'    algId={alg_id:#010x} (期望 0x00006801 RC4)')
                        print(f'    algIdHash={alg_id_hash:#010x} (期望 0x00008004 SHA1)')
                        print(f'    keySize={key_size} (期望 128)')
                        print(f'    providerType={provider_type:#010x}')
                        print(f'    reserved1={reserved1}, reserved2={reserved2}')
                        print(f'    cspName="{csp_str}"')

                        # EncryptionVerifier
                        ev_off = 12 + header_size
                        if len(cs_data) >= ev_off + 8 + 16 + 16 + 4 + 20:
                            salt_size = struct.unpack_from('<I', cs_data, ev_off)[0]
                            salt = cs_data[ev_off + 4 : ev_off + 4 + 16]
                            enc_verifier = cs_data[ev_off + 20 : ev_off + 20 + 16]
                            verifier_hash_size = struct.unpack_from('<I', cs_data, ev_off + 36)[0]
                            enc_verifier_hash = cs_data[ev_off + 40 : ev_off + 40 + 20]
                            print(f'  EncryptionVerifier:')
                            print(f'    saltSize={salt_size}')
                            print(f'    salt={salt.hex()}')
                            print(f'    encryptedVerifier={enc_verifier.hex()}')
                            print(f'    verifierHashSize={verifier_hash_size}')
                            print(f'    encryptedVerifierHash={enc_verifier_hash.hex()}')

        # 7. 检查 persist 对象的 record header 是否被加密
        print(f'\n--- Persist 对象 record header 检查 ---')
        for i, (pid, poff) in enumerate(entries):
            if poff >= len(ppt_data) - 8:
                continue
            ver, inst, rec_type, rec_len = parse_record_header(ppt_data, poff)
            # 检查 recType 是否是已知的 record type
            known_types = [0x2F14, 0x0FF5, 0x1772, 0x03EE, 0x03F8, 0x040C, 0x0FF6]
            is_known = rec_type in known_types
            encrypted_mark = '' if is_known else ' (可能已加密)'
            if i < 5 or i >= len(entries) - 3:
                print(f'  [{pid}] offset={poff}, recType={rec_type:#06x}{encrypted_mark}')
            elif i == 5:
                print(f'  ... (省略中间条目)')

        ole.close()

if __name__ == '__main__':
    # 参考文件
    analyze_ppt('_test_out/rc4cryptoapi_password.ppt')
    # 我们的加密文件
    analyze_ppt('_test_out/protected_心理账户理论.ppt')
    analyze_ppt('_test_out/wm_protected_心理账户理论.ppt')
