"""独立验证 persist 对象的加密是否正确（使用加水印后的文件作为基准）。

不依赖 msoffcrypto 的有缺陷的解密逻辑。
使用加水印后的文件（wm_心理账户理论.ppt）作为基准，因为 watermark_and_protect_ppt.rs
是先加水印再加密，所以加密前的数据是加水印后的数据。
"""
import olefile
import hashlib
from struct import unpack


PASSWORD = "pptx-rs-secret"
KEY_BITS = 128


def make_key(password, salt, key_bits, block):
    """RC4 CryptoAPI 密钥派生（与 Rust 实现一致）。"""
    password_utf16le = password.encode('utf-16-le')
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    hfinal = hashlib.sha1(h0 + block.to_bytes(4, 'little')).digest()
    key_bytes = key_bits // 8
    if key_bits == 40:
        key = hfinal[:5] + b'\x00' * 11
    else:
        key = hfinal[:key_bytes]
    return key


def rc4_crypt(key, data):
    """RC4 加密/解密（对称）。"""
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) % 256
        s[i], s[j] = s[j], s[i]
    i = 0
    j = 0
    result = bytearray(data)
    for k in range(len(data)):
        i = (i + 1) % 256
        j = (j + s[i]) % 256
        s[i], s[j] = s[j], s[i]
        result[k] ^= s[(s[i] + s[j]) % 256]
    return bytes(result)


def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    """解密一个 persist 对象（与 Rust encrypt_persist_object 反向）。"""
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)
    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        chunk = rc4_crypt(key, data[offset:end])
        result.extend(chunk)
        offset = end
        block += 1
    return bytes(result)


def parse_record_header(data, offset):
    ver_inst = int.from_bytes(data[offset:offset+2], 'little')
    rec_type = int.from_bytes(data[offset+2:offset+4], 'little')
    rec_len = int.from_bytes(data[offset+4:offset+8], 'little')
    return ver_inst & 0x0F, (ver_inst >> 4) & 0x0FFF, rec_type, rec_len


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


def parse_crypt_session(crypt_session_data):
    """解析 CryptSession10Container 的 data 部分。"""
    pos = 0
    v_major, v_minor = unpack('<HH', crypt_session_data[pos:pos+4])
    pos += 4
    outer_flags = unpack('<I', crypt_session_data[pos:pos+4])[0]
    pos += 4
    header_size = unpack('<I', crypt_session_data[pos:pos+4])[0]
    pos += 4
    header = crypt_session_data[pos:pos+header_size]
    pos += header_size
    verifier = crypt_session_data[pos:]

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
        'v_major': v_major, 'v_minor': v_minor, 'outer_flags': outer_flags,
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
    }


def main():
    # 加水印后的文件（作为基准）
    wm_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_心理账户理论.ppt"
    # 加密后的文件
    enc_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt"

    ole_wm = olefile.OleFileIO(wm_path)
    ole_enc = olefile.OleFileIO(enc_path)

    ppt_wm = ole_wm.openstream('PowerPoint Document').read()
    ppt_enc = ole_enc.openstream('PowerPoint Document').read()

    print(f"加水印 PowerPoint Document: {len(ppt_wm)} bytes")
    print(f"加密后 PowerPoint Document: {len(ppt_enc)} bytes")

    # 解析加水印文件的 persist directory
    cu_wm = ole_wm.openstream('Current User').read()
    offset_to_current_edit_wm = int.from_bytes(cu_wm[16:20], 'little')
    offset_persist_dir_wm = int.from_bytes(
        ppt_wm[offset_to_current_edit_wm+20:offset_to_current_edit_wm+24], 'little')
    persist_entries_wm = parse_persist_directory(ppt_wm, offset_persist_dir_wm)

    print(f"\n加水印文件 persist entries: {len(persist_entries_wm)}")
    print(f"加水印文件 offsetToCurrentEdit: {offset_to_current_edit_wm} ({offset_to_current_edit_wm:#x})")

    # 解析加密后文件的 persist directory
    cu_enc = ole_enc.openstream('Current User').read()
    offset_to_current_edit_enc = int.from_bytes(cu_enc[16:20], 'little')
    offset_persist_dir_enc = int.from_bytes(
        ppt_enc[offset_to_current_edit_enc+20:offset_to_current_edit_enc+24], 'little')
    persist_entries_enc = parse_persist_directory(ppt_enc, offset_persist_dir_enc)

    print(f"加密后文件 persist entries: {len(persist_entries_enc)}")

    # 找到 CryptSession10Container
    encrypt_session_pid_ref = int.from_bytes(
        ppt_enc[offset_to_current_edit_enc+8+28:offset_to_current_edit_enc+8+32], 'little')
    crypt_session_offset = None
    for pid, poff in persist_entries_enc:
        if pid == encrypt_session_pid_ref:
            crypt_session_offset = poff
            break

    cs_ver, cs_inst, cs_type, cs_len = parse_record_header(ppt_enc, crypt_session_offset)
    cs_data = ppt_enc[crypt_session_offset+8:crypt_session_offset+8+cs_len]
    cs_info = parse_crypt_session(cs_data)

    salt = cs_info['verifier']['salt']
    key_size = cs_info['header']['keySize']
    if key_size == 0:
        key_size = 0x28
    print(f"\nsalt: {salt.hex()}")
    print(f"keySize: {key_size} bits")

    # 比较 persist 对象
    print(f"\n--- 比较 persist 对象（独立解密 vs 加水印文件）---")
    wm_dict = dict(persist_entries_wm)

    match_count = 0
    mismatch_count = 0
    skip_count = 0

    for pid, wm_offset in sorted(wm_dict.items()):
        ver, inst, rec_type, rec_len = parse_record_header(ppt_wm, wm_offset)

        # 跳过 UserEditAtom 和 PersistDirectoryAtom（不加密）
        if rec_type == 0x0FF5 or rec_type == 0x1772:
            skip_count += 1
            continue

        wm_data = ppt_wm[wm_offset:wm_offset+8+rec_len]
        enc_data = ppt_enc[wm_offset:wm_offset+8+rec_len]

        decrypted = decrypt_persist_object(PASSWORD, salt, key_size, enc_data, pid)

        if decrypted == wm_data:
            match_count += 1
            if pid <= 5 or pid >= 49:
                print(f"  persistId {pid}: ✓ 一致 (type={hex(rec_type)}, len={rec_len}, offset={wm_offset:#x})")
        else:
            mismatch_count += 1
            print(f"  persistId {pid}: ✗ 不一致 (type={hex(rec_type)}, len={rec_len}, offset={wm_offset:#x})")
            for i in range(min(len(wm_data), len(decrypted))):
                if wm_data[i] != decrypted[i]:
                    print(f"    第一个不同的字节在 offset {i}: wm={wm_data[i]:#x}, dec={decrypted[i]:#x}")
                    start = max(0, i - 8)
                    end = min(len(wm_data), i + 8)
                    print(f"    wm  [{start}:{end}]: {wm_data[start:end].hex()}")
                    print(f"    dec [{start}:{end}]: {decrypted[start:end].hex()}")
                    break

    print(f"\n总计: {match_count} 一致, {mismatch_count} 不一致, {skip_count} 跳过")

    if mismatch_count == 0:
        print("\n✓ 所有 persist 对象加密正确！")
    else:
        print(f"\n✗ 有 {mismatch_count} 个 persist 对象加密不正确！")

    ole_wm.close()
    ole_enc.close()


if __name__ == "__main__":
    main()
