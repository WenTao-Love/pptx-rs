"""加密 .ppt 文件（RC4 CryptoAPI），用于验证加密流程的正确性。

加密后用 msoffcrypto 解密验证。
"""
import os
import struct
import io
import shutil
import tempfile
from hashlib import sha1

import olefile
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms as crypto_algorithms

PASSWORD = "pptx-rs-secret"
KEY_SIZE = 128  # bits
SALT_SIZE = 16


def makekey(password, salt, key_length_bits, block):
    """RC4 CryptoAPI 密钥派生。"""
    password_utf16le = password.encode("UTF-16LE")
    h0 = sha1(salt + password_utf16le).digest()
    block_bytes = struct.pack("<I", block)
    h_final = sha1(h0 + block_bytes).digest()
    key_length_bytes = key_length_bits // 8
    if key_length_bits == 40:
        key = h_final[:5] + b"\x00" * 11
    else:
        key = h_final[:key_length_bytes]
    return key


def rc4_crypt(key, data):
    """RC4 加密/解密（对称）。"""
    cipher = Cipher(crypto_algorithms.ARC4(key), mode=None, backend=default_backend())
    encryptor = cipher.encryptor()
    return encryptor.update(data) + encryptor.finalize()


def encrypt_persist_object(password, salt, key_size, data, persist_id):
    """加密一个 persist 对象。"""
    key_size_bytes = key_size // 8
    # blocksize = keySize * ((8 + recLen) // keySize + 1)  -- 未文档化
    total_len = len(data)
    blocksize = key_size_bytes * (total_len // key_size_bytes + 1)

    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        chunk = data[offset:offset + blocksize]
        key = makekey(password, salt, key_size, block)
        result.extend(rc4_crypt(key, chunk))
        offset += blocksize
        block += 1

    return bytes(result)


def parse_record_header(data, offset):
    """解析 8 字节 record header。"""
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, offset)
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def pack_record_header(ver, inst, rec_type, rec_len):
    """打包 8 字节 record header。"""
    ver_inst = (inst << 4) | ver
    return struct.pack("<HHI", ver_inst, rec_type, rec_len)


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset)] 列表。"""
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    assert rec_type == 0x1772, f"Expected PersistDirectoryAtom (0x1772), got 0x{rec_type:04X}"

    pd_data = data[offset + 8:offset + 8 + rec_len]
    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0x1FFFFF
        c_persist = (entry >> 21) & 0x7FF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + j, persist_offset))
                pos += 4

    return entries, rec_len


def build_crypt_session10_container(salt, encrypted_verifier, encrypted_verifier_hash, key_size):
    """构造 CryptSession10Container。

    结构（对齐 msoffcrypto 的 _parse_header_RC4CryptoAPI）：
    - EncryptionVersionInfo: vMajor=2, vMinor=2
    - flags: 0x00000004 (fCryptoAPI=1)
    - headerSize: EncryptionHeader 字节数
    - EncryptionHeader:
        - flags (4 bytes): 0x00000004
        - sizeExtra (4 bytes): 0
        - algId (4 bytes): 0x00006801 (RC4)
        - algIdHash (4 bytes): 0x00008004 (SHA1)
        - keySize (4 bytes): key_size (bits)
        - providerType (4 bytes): 0x00000001
        - reserved1 (4 bytes): 0
        - reserved2 (4 bytes): 0
        - cspName (UTF-16LE): "Microsoft Base Cryptographic Provider v1.0\\0"
    - EncryptionVerifier:
        - saltSize (4 bytes): 16
        - salt (16 bytes)
        - encryptedVerifier (16 bytes)
        - verifierHashSize (4 bytes): 20
        - encryptedVerifierHash (20 bytes)
    """
    # EncryptionVersionInfo: vMajor=2, vMinor=2
    enc_version = struct.pack("<HH", 2, 2)

    # 外层 flags: fCryptoAPI=1, fDocProps=0, fExternal=0
    flags = struct.pack("<I", 0x00000004)

    # EncryptionHeader
    eh_flags = 0x00000004  # fCryptoAPI=1
    size_extra = 0
    alg_id = 0x00006801  # RC4
    alg_id_hash = 0x00008004  # SHA1
    key_size_val = key_size
    provider_type = 0x00000001
    reserved1 = 0
    reserved2 = 0
    csp_name = "Microsoft Base Cryptographic Provider v1.0\0".encode("UTF-16LE")

    header = struct.pack("<IIIIIIII", eh_flags, size_extra, alg_id, alg_id_hash,
                         key_size_val, provider_type, reserved1, reserved2)
    header += csp_name
    header_size = len(header)

    # EncryptionVerifier
    salt_size = SALT_SIZE
    verifier_hash_size = 20  # SHA1 = 20 bytes

    verifier = struct.pack("<I", salt_size)
    verifier += salt
    verifier += encrypted_verifier
    verifier += struct.pack("<I", verifier_hash_size)
    verifier += encrypted_verifier_hash

    # 组装 data
    data = enc_version + flags + struct.pack("<I", header_size) + header + verifier

    # RecordHeader: ver=0xF, inst=0, type=0x2F14
    rh = pack_record_header(0xF, 0, 0x2F14, len(data))

    return rh + data


def encrypt_ppt(input_path, output_path, password):
    """加密 .ppt 文件。"""
    ole = olefile.OleFileIO(input_path)

    # 1. 读取 streams
    cu_data = bytearray(ole.openstream("Current User").read())
    ppt_data = bytearray(ole.openstream("PowerPoint Document").read())

    # 2. 解析 Current User
    cu_ver, cu_inst, cu_type, cu_len = parse_record_header(cu_data, 0)
    assert cu_type == 0x0FF6, f"Expected CurrentUserAtom (0x0FF6), got 0x{cu_type:04X}"
    offset_to_current_edit = struct.unpack_from("<I", cu_data, 16)[0]
    print(f"offsetToCurrentEdit = {offset_to_current_edit}")

    # 3. 解析 UserEditAtom
    ue_offset = offset_to_current_edit
    ue_ver, ue_inst, ue_type, ue_len = parse_record_header(ppt_data, ue_offset)
    assert ue_type == 0x0FF5, f"Expected UserEditAtom (0x0FF5), got 0x{ue_type:04X}"
    print(f"UserEditAtom: recLen={ue_len} (encrypted={ue_len == 0x20})")

    offset_persist_dir = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
    persist_id_seed = struct.unpack_from("<I", ppt_data, ue_offset + 28)[0]
    print(f"offsetPersistDirectory = {offset_persist_dir}, persistIdSeed = {persist_id_seed}")

    # 4. 解析 PersistDirectoryAtom
    persist_entries, pd_len = parse_persist_directory(ppt_data, offset_persist_dir)
    print(f"Persist entries: {len(persist_entries)}")

    # 5. 生成加密参数
    import secrets
    salt = secrets.token_bytes(SALT_SIZE)
    verifier_plain = secrets.token_bytes(16)
    verifier_hash = sha1(verifier_plain).digest()

    # 用 block=0 的 key 加密 verifier
    key_block0 = makekey(password, salt, KEY_SIZE, 0)
    encrypted_verifier = rc4_crypt(key_block0, verifier_plain)
    encrypted_verifier_hash = rc4_crypt(key_block0, verifier_hash)

    # 6. 加密 persist 对象
    # 排序 persist entries by offset
    persist_entries.sort(key=lambda x: x[1])

    for pid, poff in persist_entries:
        ver, inst, rec_type, rec_len = parse_record_header(ppt_data, poff)

        # 跳过 UserEditAtom (0x0FF5) 和 PersistDirectoryAtom (0x1772)
        if rec_type in (0x0FF5, 0x1772):
            print(f"  Skip pid={pid}, type=0x{rec_type:04X} (not encrypted)")
            continue

        # 加密整个 record（header + data）
        record_data = bytes(ppt_data[poff:poff + 8 + rec_len])
        encrypted = encrypt_persist_object(password, salt, KEY_SIZE, record_data, pid)
        ppt_data[poff:poff + 8 + rec_len] = encrypted
        print(f"  Encrypted pid={pid}, type=0x{rec_type:04X}, len={rec_len}")

    # 7. 构造 CryptSession10Container
    crypt_session = build_crypt_session10_container(salt, encrypted_verifier, encrypted_verifier_hash, KEY_SIZE)

    # 8. 在 stream 末尾追加 CryptSession10Container
    crypt_session_offset = len(ppt_data)
    ppt_data.extend(crypt_session)
    print(f"CryptSession10Container at offset {crypt_session_offset}, size={len(crypt_session)}")

    # 9. 修改 UserEditAtom: recLen 28→32, 添加 encryptSessionPersistIdRef
    # CryptSession10Container 的 persistId = 最后一个 persistId + 1
    last_pid = persist_entries[-1][0]
    crypt_session_pid = last_pid + 1

    # 修改 recLen
    new_ue_len = ue_len + 4  # 28 → 32
    struct.pack_into("<I", ppt_data, ue_offset + 4, new_ue_len)

    # 在 UserEditAtom 末尾添加 encryptSessionPersistIdRef
    # UserEditAtom data 从 ue_offset+8 开始，长度为 ue_len
    # encryptSessionPersistIdRef 在 data 的最后 4 字节
    struct.pack_into("<I", ppt_data, ue_offset + 8 + ue_len, crypt_session_pid)
    print(f"UserEditAtom: recLen={ue_len}→{new_ue_len}, encryptSessionPersistIdRef={crypt_session_pid}")

    # 10. 修改 PersistDirectoryAtom: 添加 CryptSession10Container 条目
    # 在第一个 PersistDirectoryEntry 的 rgPersistOffset 末尾添加 offset
    # 同时 cPersist 加 1

    # 读取第一个 PersistDirectoryEntry
    pd_data_start = offset_persist_dir + 8
    entry_val = struct.unpack_from("<I", ppt_data, pd_data_start)[0]
    entry_pid = entry_val & 0x1FFFFF
    entry_cpersist = (entry_val >> 21) & 0x7FF

    # cPersist + 1
    new_entry_val = entry_pid | ((entry_cpersist + 1) << 21)
    struct.pack_into("<I", ppt_data, pd_data_start, new_entry_val)

    # 在 rgPersistOffset 末尾插入 CryptSession10Container 的 offset
    # rgPersistOffset 从 pd_data_start + 4 开始，有 entry_cpersist 个 u32
    insert_pos = pd_data_start + 4 + entry_cpersist * 4
    ppt_data[insert_pos:insert_pos] = struct.pack("<I", crypt_session_offset)

    # 更新 PersistDirectoryAtom 的 recLen
    new_pd_len = pd_len + 4  # 多了一个 offset
    struct.pack_into("<I", ppt_data, offset_persist_dir + 4, new_pd_len)
    print(f"PersistDirectoryAtom: recLen={pd_len}→{new_pd_len}, cPersist={entry_cpersist}→{entry_cpersist+1}")

    # 11. 修改 Current User: headerToken 0xE391C05F → 0xF3D1C4DF
    struct.pack_into("<I", cu_data, 12, 0xF3D1C4DF)
    print(f"CurrentUser: headerToken=0xE391C05F→0xF3D1C4DF")

    # 12. 写回 OLE2 容器
    # 复制原始文件，然后修改 streams
    with tempfile.NamedTemporaryFile(delete=False, suffix='.ppt') as tmp:
        shutil.copyfile(input_path, tmp.name)
        tmp_ole = olefile.OleFileIO(tmp.name, write_mode=True)
        tmp_ole.write_stream("Current User", bytes(cu_data))
        tmp_ole.write_stream("PowerPoint Document", bytes(ppt_data))
        tmp_ole.close()

        shutil.copyfile(tmp.name, output_path)
        os.unlink(tmp.name)

    ole.close()
    print(f"\nEncrypted .ppt saved to: {output_path}")


if __name__ == "__main__":
    os.makedirs("_test_out", exist_ok=True)
    encrypt_ppt(
        "_test/心理账户理论.ppt",
        "_test_out/encrypted_ppt.ppt",
        PASSWORD,
    )
    print(f"\nPassword: {PASSWORD}")

    # 验证：用 msoffcrypto 检查是否加密
    f = open("_test_out/encrypted_ppt.ppt", "rb")
    ms = msoffcrypto.OfficeFile(f) if 'msoffcrypto' in dir() else None
    if ms:
        print(f"msoffcrypto is_encrypted: {ms.is_encrypted()}")
    f.close()
