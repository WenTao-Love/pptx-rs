#!/usr/bin/env python3
"""用正确的解密逻辑（8+recLen）解密 .ppt 文件，测试不同的 blocksize 计算。

MS-OFFCRYPTO 2.3.5.1 规范说 blocksize = keySizeBytes * (totalLength / keySizeBytes + 1)
msoffcrypto 用 keySize（位数）作为 blocksize 基数

测试两种 blocksize 计算，看哪种能正确解密。
"""

import struct
import sys
import hashlib
from pathlib import Path

try:
    import olefile
except ImportError:
    print("需要 olefile: pip install olefile")
    sys.exit(1)

try:
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives.ciphers import Cipher
    try:
        from cryptography.hazmat.decrepit.ciphers.algorithms import ARC4
    except ImportError:
        from cryptography.hazmat.primitives.ciphers.algorithms import ARC4
except ImportError:
    print("需要 cryptography: pip install cryptography")
    sys.exit(1)


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def make_key(password, salt, key_bits, block):
    """RC4 CryptoAPI 密钥派生。"""
    password_utf16le = password.encode("UTF-16LE")
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    hfinal = hashlib.sha1(h0 + struct.pack("<I", block)).digest()
    if key_bits == 40:
        key = hfinal[:5] + b"\x00" * 11
    else:
        key = hfinal[: key_bits // 8]
    return key


def rc4_decrypt(key, data):
    """RC4 解密。"""
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    return decryptor.update(data) + decryptor.finalize()


def decrypt_persist_object(password, salt, key_bits, data, persist_id, blocksize_mode):
    """解密一个 persist 对象。

    blocksize_mode:
      - "keyBits": blocksize = keyBits * (totalLen // keyBits + 1)  (msoffcrypto 方式)
      - "keyBytes": blocksize = keyBytes * (totalLen // keyBytes + 1)  (MS-OFFCRYPTO 规范方式)
    """
    total_len = len(data)
    key_bytes = key_bits // 8

    if blocksize_mode == "keyBits":
        blocksize = key_bits * (total_len // key_bits + 1)
    else:  # keyBytes
        blocksize = key_bytes * (total_len // key_bytes + 1)

    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        chunk = rc4_decrypt(key, data[offset:end])
        result.extend(chunk)
        offset = end
        block += 1
    return bytes(result)


def parse_crypt_session(data, offset):
    """解析 CryptSession10Container，返回 (salt, keySize, encryptedVerifier, encryptedVerifierHash)。"""
    h = parse_header(data, offset)
    if h is None or h[2] != 0x2F14:
        return None

    pos = offset + 8
    # EncryptionVersionInfo
    pos += 4
    # flags
    pos += 4
    # headerSize
    header_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    # EncryptionHeader
    pos += header_size
    # EncryptionVerifier
    salt_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    salt = data[pos : pos + 16]
    pos += 16
    encrypted_verifier = data[pos : pos + 16]
    pos += 16
    verifier_hash_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    encrypted_verifier_hash = data[pos : pos + 20]

    # 读取 keySize from EncryptionHeader
    header_start = offset + 8 + 4 + 4 + 4
    key_size = struct.unpack_from("<I", data, header_start + 16)[0]

    return (salt, key_size, encrypted_verifier, encrypted_verifier_hash)


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset)]。"""
    h = parse_header(data, offset)
    if h is None or h[2] != 0x1772:
        return []

    _, _, _, rec_len = h
    pd_data = data[offset + 8 : offset + 8 + rec_len]

    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                poff = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + j, poff))
                pos += 4
    return entries


def decrypt_ppt(password, encrypted_path, blocksize_mode):
    """解密 .ppt 文件，返回解密后的 PowerPoint Document stream。"""
    ole = olefile.OleFileIO(encrypted_path)
    ppt_data = bytearray(ole.openstream("PowerPoint Document").read())
    cu_data = ole.openstream("Current User").read()

    # 找到 UserEditAtom
    ue_offset = struct.unpack_from("<I", cu_data, 16)[0]
    pd_offset = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
    encrypt_session_pid = struct.unpack_from("<I", ppt_data, ue_offset + 8 + 28)[0]

    # 解析 persist 目录
    entries = parse_persist_directory(ppt_data, pd_offset)

    # 找到 CryptSession10Container
    crypt_session_offset = None
    for pid, off in entries:
        if pid == encrypt_session_pid:
            crypt_session_offset = off
            break

    if crypt_session_offset is None:
        print(f"  找不到 CryptSession10Container")
        ole.close()
        return None

    # 解析 CryptSession10Container
    info = parse_crypt_session(ppt_data, crypt_session_offset)
    if info is None:
        print(f"  解析 CryptSession10Container 失败")
        ole.close()
        return None

    salt, key_bits, enc_verifier, enc_verifier_hash = info

    # 验证密码
    key = make_key(password, salt, key_bits, 0)
    verifier = rc4_decrypt(key, enc_verifier)
    verifier_hash = rc4_decrypt(key, enc_verifier_hash)  # 注意：这里应该用同一个 RC4 流
    # 修正：verifier 和 verifierHash 应该用同一个 RC4 流连续解密
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    verifier = decryptor.update(enc_verifier)
    verifier_hash = decryptor.update(enc_verifier_hash)

    if hashlib.sha1(verifier).digest() == verifier_hash:
        print(f"  密码验证通过 ✓")
    else:
        print(f"  密码验证失败 ✗")
        ole.close()
        return None

    # 解密 persist 对象
    for pid, poff in entries:
        # 跳过 UserEditAtom, PersistDirectoryAtom, CryptSession10Container
        h = parse_header(ppt_data, poff)
        if h is None:
            continue
        rec_type = h[2]
        if rec_type in [0x0FF5, 0x1772, 0x2F14]:
            continue

        # 解密前 8 字节读取 recLen
        key = make_key(password, salt, key_bits, pid)
        # 先解密整个 record（8 + recLen），但我们需要先知道 recLen
        # 正确方法：先解密前 8 字节，读取 recLen，再继续解密剩余部分
        # 但 RC4 是流密码，需要连续解密

        # 方法：用 blocksize 分段解密
        # 第一段：解密前 blocksize 字节（至少 8 字节）
        # 读取 recLen，计算 total_len
        # 继续解密剩余部分

        # 简化方法：假设 blocksize >= 8，先解密第一段，读取 recLen，再解密剩余部分
        key_bytes = key_bits // 8
        if blocksize_mode == "keyBits":
            # 先估算 blocksize，用于解密前 8 字节
            # 但我们不知道 total_len，所以先用一个足够大的 blocksize 解密前 8 字节
            # 实际上，blocksize >= key_bits >= 128 > 8，所以第一段至少 128 字节
            # 先解密前 key_bits 字节
            first_chunk_size = min(key_bits, len(ppt_data) - poff)
            first_key = make_key(password, salt, key_bits, pid)
            first_chunk = rc4_decrypt(first_key, bytes(ppt_data[poff : poff + first_chunk_size]))

            # 读取 recLen
            rec_len = struct.unpack_from("<I", first_chunk, 4)[0]
            total_len = 8 + rec_len

            # 计算实际 blocksize
            blocksize = key_bits * (total_len // key_bits + 1)

            # 重新解密整个 record
            enc_data = bytes(ppt_data[poff : poff + total_len])
            dec_data = decrypt_persist_object(password, salt, key_bits, enc_data, pid, blocksize_mode)
            ppt_data[poff : poff + total_len] = dec_data
        else:  # keyBytes
            # 先解密前 key_bytes 字节
            first_chunk_size = min(key_bytes * 2, len(ppt_data) - poff)  # 至少 16 字节
            first_key = make_key(password, salt, key_bits, pid)
            first_chunk = rc4_decrypt(first_key, bytes(ppt_data[poff : poff + first_chunk_size]))

            # 读取 recLen
            rec_len = struct.unpack_from("<I", first_chunk, 4)[0]
            total_len = 8 + rec_len

            # 重新解密整个 record
            enc_data = bytes(ppt_data[poff : poff + total_len])
            dec_data = decrypt_persist_object(password, salt, key_bits, enc_data, pid, blocksize_mode)
            ppt_data[poff : poff + total_len] = dec_data

    # 解密 Pictures stream
    if ole.exists("Pictures"):
        pic_data = bytearray(ole.openstream("Pictures").read())
        pos = 0
        while pos + 8 <= len(pic_data):
            # 先解密前 8 字节读取 recLen
            key = make_key(password, salt, key_bits, 0)
            first_chunk = rc4_decrypt(key, bytes(pic_data[pos : pos + 8]))
            rec_len = struct.unpack_from("<I", first_chunk, 4)[0]
            total_len = 8 + rec_len

            if pos + total_len > len(pic_data):
                break

            # 解密整个 record（block=0）
            enc_data = bytes(pic_data[pos : pos + total_len])
            # Pictures stream 用 block=0，不分段（或用 keyBits blocksize）
            if blocksize_mode == "keyBits":
                blocksize = key_bits * (total_len // key_bits + 1)
            else:
                blocksize = key_bytes * (total_len // key_bytes + 1)

            result = bytearray()
            offset = 0
            block = 0
            while offset < total_len:
                end = min(offset + blocksize, total_len)
                k = make_key(password, salt, key_bits, block)
                chunk = rc4_decrypt(k, enc_data[offset:end])
                result.extend(chunk)
                offset = end
                block += 1

            pic_data[pos : pos + total_len] = result
            pos += total_len

        print(f"  Pictures stream 解密完成")

    ole.close()
    return bytes(ppt_data)


def verify_decrypted(ppt_data, label):
    """验证解密后的数据结构。"""
    print(f"\n  --- {label} 解密后顶层 records ---")
    pos = 0
    count = 0
    valid = True
    while pos + 8 <= len(ppt_data) and count < 60:
        h = parse_header(ppt_data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len

        # 检查 record 是否合理
        type_names = {
            0x03E8: "Document",
            0x03EE: "Slide",
            0x03F8: "MainMaster",
            0x0FF5: "UserEditAtom",
            0x1772: "PersistDirectoryAtom",
            0x2F14: "CryptSession10Container",
        }
        name = type_names.get(rec_type, f"0x{rec_type:04X}")

        # 检查 rec_len 是否合理
        if rec_len > 10000000:
            print(f"    offset={pos:>8} type=0x{rec_type:04X}({name}) len={rec_len} ⚠️ 异常大")
            valid = False
            break

        print(f"    offset={pos:>8} type=0x{rec_type:04X}({name:>25}) len={rec_len:>10} container={'Y' if is_container else 'N'}")

        pos += total_len
        count += 1
        if not is_container and rec_len == 0:
            break

    print(f"    ... 共 {count} 个顶层 record, stream 总长 {len(ppt_data)}")
    return valid


def main():
    password = "pptx-rs-secret"
    out_dir = Path("_test_out")

    for ppt_path in out_dir.glob("protected_*.ppt"):
        if "decrypted" in ppt_path.name:
            continue

        print(f"\n{'='*70}")
        print(f"文件: {ppt_path}")
        print(f"{'='*70}")

        # 测试 keyBits blocksize
        print(f"\n  [模式1] blocksize = keyBits (msoffcrypto 方式)")
        try:
            dec_data = decrypt_ppt(password, ppt_path, "keyBits")
            if dec_data:
                verify_decrypted(dec_data, "keyBits")
        except Exception as e:
            print(f"  解密失败: {e}")
            import traceback
            traceback.print_exc()

        # 测试 keyBytes blocksize
        print(f"\n  [模式2] blocksize = keyBytes (MS-OFFCRYPTO 规范方式)")
        try:
            dec_data = decrypt_ppt(password, ppt_path, "keyBytes")
            if dec_data:
                verify_decrypted(dec_data, "keyBytes")
        except Exception as e:
            print(f"  解密失败: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()
