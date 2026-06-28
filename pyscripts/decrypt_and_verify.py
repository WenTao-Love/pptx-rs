#!/usr/bin/env python3
"""完整解密 .ppt 文件，保存解密后的文件，检查 Pictures stream 解密是否正确。"""

import struct
import sys
import hashlib
import shutil
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
    password_utf16le = password.encode("UTF-16LE")
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    hfinal = hashlib.sha1(h0 + struct.pack("<I", block)).digest()
    if key_bits == 40:
        key = hfinal[:5] + b"\x00" * 11
    else:
        key = hfinal[: key_bits // 8]
    return key


def rc4_decrypt(key, data):
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    return decryptor.update(data) + decryptor.finalize()


def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    """解密一个 persist 对象，用 keyBits blocksize。"""
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)

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
    h = parse_header(data, offset)
    if h is None or h[2] != 0x2F14:
        return None

    pos = offset + 8
    pos += 4  # EncryptionVersionInfo
    pos += 4  # flags
    header_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    pos += header_size  # EncryptionHeader

    salt_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    salt = data[pos : pos + 16]
    pos += 16
    encrypted_verifier = data[pos : pos + 16]
    pos += 16
    verifier_hash_size = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    encrypted_verifier_hash = data[pos : pos + 20]

    header_start = offset + 8 + 4 + 4 + 4
    key_size = struct.unpack_from("<I", data, header_start + 16)[0]

    return (salt, key_size, encrypted_verifier, encrypted_verifier_hash)


def parse_persist_directory(data, offset):
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


def decrypt_ppt_full(password, encrypted_path, output_path):
    """完整解密 .ppt 文件，保存解密后的文件。"""
    # 复制原始文件
    shutil.copy(encrypted_path, output_path)

    ole = olefile.OleFileIO(output_path, write_mode=True)
    ppt_data = bytearray(ole.openstream("PowerPoint Document").read())
    cu_data = bytearray(ole.openstream("Current User").read())

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

    info = parse_crypt_session(ppt_data, crypt_session_offset)
    salt, key_bits, enc_verifier, enc_verifier_hash = info

    # 验证密码
    key = make_key(password, salt, key_bits, 0)
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    verifier = decryptor.update(enc_verifier)
    verifier_hash = decryptor.update(enc_verifier_hash)

    if hashlib.sha1(verifier).digest() == verifier_hash:
        print(f"  密码验证通过 ✓")
    else:
        print(f"  密码验证失败 ✗")
        ole.close()
        return False

    # 解密 persist 对象
    for pid, poff in entries:
        h = parse_header(ppt_data, poff)
        if h is None:
            continue
        rec_type = h[2]
        if rec_type in [0x0FF5, 0x1772, 0x2F14]:
            continue

        # 先解密前 key_bits 字节读取 recLen
        first_chunk_size = min(key_bits, len(ppt_data) - poff)
        first_key = make_key(password, salt, key_bits, pid)
        first_chunk = rc4_decrypt(first_key, bytes(ppt_data[poff : poff + first_chunk_size]))

        rec_len = struct.unpack_from("<I", first_chunk, 4)[0]
        total_len = 8 + rec_len

        # 重新解密整个 record
        enc_data = bytes(ppt_data[poff : poff + total_len])
        dec_data = decrypt_persist_object(password, salt, key_bits, enc_data, pid)
        ppt_data[poff : poff + total_len] = dec_data

    # 解密 Pictures stream
    if ole.exists("Pictures"):
        pic_data = bytearray(ole.openstream("Pictures").read())
        print(f"  Pictures stream 加密前长度: {len(pic_data)}")

        pos = 0
        record_count = 0
        while pos + 8 <= len(pic_data):
            # 先解密前 8 字节读取 recLen
            key = make_key(password, salt, key_bits, 0)
            first_chunk = rc4_decrypt(key, bytes(pic_data[pos : pos + 8]))
            rec_len = struct.unpack_from("<I", first_chunk, 4)[0]
            total_len = 8 + rec_len

            if pos + total_len > len(pic_data):
                print(f"    Pictures record #{record_count}: offset={pos}, total_len={total_len} 超出范围")
                break

            # 解密整个 record（block=0，用 keyBits blocksize）
            enc_data = bytes(pic_data[pos : pos + total_len])
            dec_data = decrypt_persist_object(password, salt, key_bits, enc_data, 0)
            pic_data[pos : pos + total_len] = dec_data

            # 验证解密后的 record header
            h = parse_header(pic_data, pos)
            if h:
                ver, inst, rec_type, rec_len2 = h
                print(f"    Pictures record #{record_count}: offset={pos}, type=0x{rec_type:04X}, len={rec_len2}")

            pos += total_len
            record_count += 1

        print(f"  Pictures stream 解密完成，共 {record_count} 个 record")

        # 写回 Pictures stream
        ole.write_stream("Pictures", bytes(pic_data))

    # 修改 Current User: headerToken 0xF3D1C4DF → 0xE391C05F
    struct.pack_into("<I", cu_data, 12, 0xE391C05F)
    # UserEditAtom recLen 32 → 28（移除 encryptSessionPersistIdRef）
    # 但这会改变 stream 大小，所以保持 32，只是修改 headerToken

    # 写回 Current User
    ole.write_stream("Current User", bytes(cu_data))

    # 写回 PowerPoint Document
    ole.write_stream("PowerPoint Document", bytes(ppt_data))

    ole.close()

    print(f"  解密后文件保存到: {output_path}")
    return True


def verify_pictures_stream(path):
    """验证解密后的 Pictures stream 结构。"""
    print(f"\n  --- 验证 Pictures stream ---")
    ole = olefile.OleFileIO(path)
    if not ole.exists("Pictures"):
        print(f"  Pictures stream 不存在")
        ole.close()
        return

    pic_data = ole.openstream("Pictures").read()
    print(f"  Pictures stream 长度: {len(pic_data)}")

    pos = 0
    count = 0
    valid = True
    while pos + 8 <= len(pic_data):
        h = parse_header(pic_data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        total_len = 8 + rec_len

        # 检查 record type 是否合理（Pictures stream 中的 record type 通常是 0xF01D, 0xF01E, 0xF01F）
        if rec_type not in [0xF01D, 0xF01E, 0xF01F, 0xF01A, 0xF01B, 0xF01C]:
            print(f"    record #{count}: offset={pos}, type=0x{rec_type:04X} ⚠️ 异常 type")
            valid = False

        if rec_len > 1000000:
            print(f"    record #{count}: offset={pos}, type=0x{rec_type:04X}, len={rec_len} ⚠️ 异常大")
            valid = False
            break

        print(f"    record #{count}: offset={pos}, ver={ver}, inst=0x{inst:04X}, type=0x{rec_type:04X}, len={rec_len}")

        pos += total_len
        count += 1

    print(f"  共 {count} 个 record, 总长度 {pos} (stream 长度 {len(pic_data)})")
    if pos != len(pic_data):
        print(f"  ⚠️ record 总长度 {pos} != stream 长度 {len(pic_data)}")

    ole.close()
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

        output_path = str(ppt_path).replace(".ppt", ".full_decrypted.ppt")

        if decrypt_ppt_full(password, ppt_path, output_path):
            # 验证解密后的文件
            verify_pictures_stream(output_path)

            # 检查解密后的顶层 record
            ole = olefile.OleFileIO(output_path)
            ppt_data = ole.openstream("PowerPoint Document").read()
            print(f"\n  --- 解密后顶层 records ---")
            pos = 0
            count = 0
            while pos + 8 <= len(ppt_data) and count < 60:
                h = parse_header(ppt_data, pos)
                if h is None:
                    break
                ver, inst, rec_type, rec_len = h
                is_container = ver == 0xF
                total_len = 8 + rec_len

                type_names = {
                    0x03E8: "Document",
                    0x03EE: "Slide",
                    0x03F0: "SlideList",
                    0x03F8: "MainMaster",
                    0x0FF5: "UserEditAtom",
                    0x1772: "PersistDirectoryAtom",
                    0x2F14: "CryptSession10Container",
                }
                name = type_names.get(rec_type, f"0x{rec_type:04X}")

                if rec_len > 10000000:
                    print(f"    offset={pos:>8} type=0x{rec_type:04X}({name}) len={rec_len} ⚠️ 异常大")
                    break

                print(f"    offset={pos:>8} type=0x{rec_type:04X}({name:>25}) len={rec_len:>10} container={'Y' if is_container else 'N'}")

                pos += total_len
                count += 1
                if not is_container and rec_len == 0:
                    break

            print(f"    ... 共 {count} 个顶层 record")
            ole.close()


if __name__ == "__main__":
    main()
