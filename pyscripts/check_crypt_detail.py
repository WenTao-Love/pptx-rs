"""检查 CryptSession10Container 格式和 msoffcrypto 解密后的 persist 对象。"""
import olefile
import struct
import sys
import os
import msoffcrypto

def parse_record_header(data, offset):
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def check_crypt_session(filepath):
    """检查 CryptSession10Container 格式。"""
    print(f'\n{"="*60}')
    print(f'检查 CryptSession10Container: {filepath}')
    print(f'{"="*60}')

    with open(filepath, 'rb') as f:
        ole = olefile.OleFileIO(f)
        cu = ole.openstream('Current User').read()
        ppt = ole.openstream('PowerPoint Document').read()

        offset_to_current_edit = struct.unpack_from('<I', cu, 16)[0]
        ue_offset = offset_to_current_edit

        # 读取 encryptSessionPersistIdRef
        ue_rec_len = struct.unpack_from('<I', ppt, ue_offset + 4)[0]
        if ue_rec_len == 32:
            encrypt_session_pid = struct.unpack_from('<I', ppt, ue_offset + 36)[0]
        else:
            print(f'  UserEditAtom recLen={ue_rec_len}, 不是加密文件')
            ole.close()
            return

        offset_persist_dir = struct.unpack_from('<I', ppt, ue_offset + 20)[0]

        # 解析 persist directory，构建 persistobjectdirectory
        pd_ver, pd_inst, pd_type, pd_len = parse_record_header(ppt, offset_persist_dir)
        pd_data = ppt[offset_persist_dir + 8 : offset_persist_dir + 8 + pd_len]

        pos = 0
        persist_dir = {}
        while pos + 4 <= len(pd_data):
            entry_val = struct.unpack_from('<I', pd_data, pos)[0]
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF
            pos += 4
            for j in range(c_persist):
                if pos + 4 <= len(pd_data):
                    poff = struct.unpack_from('<I', pd_data, pos)[0]
                    persist_dir[persist_id + j] = poff
                    pos += 4

        # 查找 CryptSession10Container
        cs_offset = persist_dir.get(encrypt_session_pid)
        if cs_offset is None:
            print(f'  找不到 CryptSession10Container (persistId={encrypt_session_pid})')
            ole.close()
            return

        print(f'  CryptSession10Container offset={cs_offset}')
        ver, inst, rec_type, rec_len = parse_record_header(ppt, cs_offset)
        print(f'  recVer={ver:#x}, recInstance={inst:#x}, recType={rec_type:#06x}, recLen={rec_len}')
        print(f'  recType 期望: 0x2F14')

        cs_data = ppt[cs_offset + 8 : cs_offset + 8 + rec_len]
        print(f'  data 长度: {len(cs_data)} bytes')

        if len(cs_data) >= 12:
            v_major, v_minor = struct.unpack_from('<HH', cs_data, 0)
            outer_flags = struct.unpack_from('<I', cs_data, 4)[0]
            header_size = struct.unpack_from('<I', cs_data, 8)[0]
            print(f'  vMajor={v_major} (期望 4)')
            print(f'  vMinor={v_minor} (期望 2)')
            print(f'  outer_flags={outer_flags:#010x} (期望 0x0000000C)')
            print(f'  headerSize={header_size}')

            # EncryptionHeader
            eh_off = 12
            if len(cs_data) >= eh_off + header_size:
                flags = struct.unpack_from('<I', cs_data, eh_off)[0]
                alg_id = struct.unpack_from('<I', cs_data, eh_off + 8)[0]
                alg_id_hash = struct.unpack_from('<I', cs_data, eh_off + 12)[0]
                key_size = struct.unpack_from('<I', cs_data, eh_off + 16)[0]
                csp_name = cs_data[eh_off + 32:eh_off + header_size]
                try:
                    csp_str = csp_name.decode('utf-16-le').rstrip('\x00')
                except:
                    csp_str = repr(csp_name[:40])
                print(f'  EH.flags={flags:#010x}')
                print(f'  EH.algId={alg_id:#010x}')
                print(f'  EH.algIdHash={alg_id_hash:#010x}')
                print(f'  EH.keySize={key_size}')
                print(f'  EH.cspName="{csp_str}"')

                # EncryptionVerifier
                ev_off = 12 + header_size
                if len(cs_data) >= ev_off + 47:
                    salt_size = struct.unpack_from('<I', cs_data, ev_off)[0]
                    salt = cs_data[ev_off + 4 : ev_off + 4 + 16]
                    verifier_hash_size = struct.unpack_from('<I', cs_data, ev_off + 36)[0]
                    print(f'  EV.saltSize={salt_size}')
                    print(f'  EV.verifierHashSize={verifier_hash_size}')

        ole.close()

def check_decrypted_persist(filepath):
    """用 msoffcrypto 解密文件，检查 persist 对象是否被正确解密。"""
    print(f'\n{"="*60}')
    print(f'检查解密后的 persist 对象: {filepath}')
    print(f'{"="*60}')

    out_path = filepath + '.decrypted2.ppt'
    with open(filepath, 'rb') as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password='pptx-rs-secret')
        with open(out_path, 'wb') as out:
            office.decrypt(out)

    with open(out_path, 'rb') as f:
        ole = olefile.OleFileIO(f)
        cu = ole.openstream('Current User').read()
        ppt = ole.openstream('PowerPoint Document').read()

        header_token = struct.unpack_from('<I', cu, 12)[0]
        offset_to_current_edit = struct.unpack_from('<I', cu, 16)[0]
        print(f'  headerToken={header_token:#010x} (期望 0xE391C05F 未加密)')

        ue_offset = offset_to_current_edit
        ue_rec_len = struct.unpack_from('<I', ppt, ue_offset + 4)[0]
        print(f'  UserEditAtom recLen={ue_rec_len} (期望 28 未加密)')

        offset_persist_dir = struct.unpack_from('<I', ppt, ue_offset + 20)[0]

        # 解析 persist directory
        pd_ver, pd_inst, pd_type, pd_len = parse_record_header(ppt, offset_persist_dir)
        pd_data = ppt[offset_persist_dir + 8 : offset_persist_dir + 8 + pd_len]

        pos = 0
        persist_dir = {}
        while pos + 4 <= len(pd_data):
            entry_val = struct.unpack_from('<I', pd_data, pos)[0]
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF
            pos += 4
            for j in range(c_persist):
                if pos + 4 <= len(pd_data):
                    poff = struct.unpack_from('<I', pd_data, pos)[0]
                    persist_dir[persist_id + j] = poff
                    pos += 4

        # 检查每个 persist 对象的 record header 是否有效
        print(f'\n  --- 解密后的 persist 对象 record header ---')
        known_types = [0x2F14, 0x0FF5, 0x1772, 0x03EE, 0x03F8, 0x040C, 0x0FF6,
                       0x03F0, 0x03EF, 0x03F2, 0x03FD, 0x03FE, 0x0410, 0x0411,
                       0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417, 0x0418,
                       0x0419, 0x041A, 0x041B, 0x041C, 0x041D, 0x041E, 0x041F,
                       0x0FF0, 0x0FF1, 0x0FF2, 0x0FF3, 0x0FF4, 0x0FF7, 0x0FF8]
        valid_count = 0
        invalid_count = 0
        for pid in sorted(persist_dir.keys()):
            poff = persist_dir[pid]
            if poff + 8 > len(ppt):
                continue
            ver, inst, rec_type, rec_len = parse_record_header(ppt, poff)
            is_known = rec_type in known_types or (0x03E0 <= rec_type <= 0x0420)
            status = 'OK' if is_known else 'INVALID'
            if is_known:
                valid_count += 1
            else:
                invalid_count += 1
            if pid <= 5 or pid >= max(persist_dir.keys()) - 2 or not is_known:
                print(f'  [{pid}] offset={poff}, recType={rec_type:#06x}, recLen={rec_len} [{status}]')
            elif pid == 6:
                print(f'  ... (省略中间条目)')

        print(f'\n  有效 record header: {valid_count}')
        print(f'  无效 record header: {invalid_count}')

        # 检查 Pictures stream
        try:
            pics = ole.openstream('Pictures').read()
            print(f'\n  Pictures stream: {len(pics)} bytes')
            # 检查第一个 record header
            if len(pics) >= 8:
                ver, inst, rec_type, rec_len = parse_record_header(pics, 0)
                print(f'  Pictures[0]: recType={rec_type:#06x}, recLen={rec_len}')
                # FBSE record type = 0xF007
                if rec_type == 0xF007:
                    print(f'  Pictures[0] 是 FBSE (0xF007) - 有效')
                else:
                    print(f'  Pictures[0] 不是 FBSE - 可能未正确解密')
        except:
            print(f'\n  Pictures stream: 不存在')

        ole.close()

if __name__ == '__main__':
    # 检查参考文件
    check_crypt_session('_test_out/rc4cryptoapi_password.ppt')
    # 检查我们的文件
    check_crypt_session('_test_out/protected_心理账户理论.ppt')

    # 检查解密后的 persist 对象
    check_decrypted_persist('_test_out/protected_心理账户理论.ppt')
