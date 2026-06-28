"""用 msoffcrypto 解密我们生成的加密文件，然后检查解密后的 persist 对象是否正确。

关键检查：
1. 解密后的 PowerPoint Document stream 的 persist 对象是否与原始文件一致
2. 解密后的 Pictures stream 是否仍然是加密的（msoffcrypto 不解密 Pictures stream）
3. 用我们自己的解密逻辑解密 Pictures stream，检查是否与原始文件一致
"""
import olefile
import io
import os
import sys
from struct import unpack, pack
from hashlib import sha1
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher

try:
    from cryptography.hazmat.decrepit.ciphers.algorithms import ARC4
except ImportError:
    from cryptography.hazmat.primitives.ciphers.algorithms import ARC4

import msoffcrypto


def make_key(password, salt, key_length, block, algIdHash=0x00008004):
    """RC4 CryptoAPI 密钥派生"""
    password = password.encode("UTF-16LE")
    h0 = sha1(salt + password).digest()
    blockbytes = pack("<I", block)
    hfinal = sha1(h0 + blockbytes).digest()
    if key_length == 40:
        key = hfinal[:5] + b"\x00" * 11
    else:
        key = hfinal[: key_length // 8]
    return key


def rc4_decrypt_block(password, salt, key_size, data, block=0):
    """RC4 解密一个 block"""
    key = make_key(password, salt, key_size, block)
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    return decryptor.update(data) + decryptor.finalize()


def decrypt_pictures_stream(password, salt, key_size, data):
    """解密 Pictures stream（与我们的 encrypt_pictures_stream 反向操作）"""
    result = bytearray(data)
    offset = 0
    while offset + 8 <= len(result):
        # 1. 解密 header (8 bytes)，重置 RC4
        key = make_key(password, salt, key_size, 0)
        cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
        decryptor = cipher.decryptor()
        result[offset:offset+8] = decryptor.update(bytes(result[offset:offset+8])) + decryptor.finalize()

        # 读取 header 字段
        ver_inst = int.from_bytes(result[offset:offset+2], 'little')
        rec_type = int.from_bytes(result[offset+2:offset+4], 'little')
        rlen = int.from_bytes(result[offset+4:offset+8], 'little')
        rec_inst = (ver_inst >> 4) & 0x0FFF

        pos = offset + 8
        end_offset = pos + rlen

        if rec_type == 0xF007:
            # FBSE
            cb_name = int.from_bytes(result[pos+33:pos+35], 'little')
            parts = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1]
            for part in parts:
                key = make_key(password, salt, key_size, 0)
                cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
                decryptor = cipher.decryptor()
                result[pos:pos+part] = decryptor.update(bytes(result[pos:pos+part])) + decryptor.finalize()
                pos += part
            if cb_name > 0:
                key = make_key(password, salt, key_size, 0)
                cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
                decryptor = cipher.decryptor()
                result[pos:pos+cb_name] = decryptor.update(bytes(result[pos:pos+cb_name])) + decryptor.finalize()
                pos += cb_name
            if pos >= end_offset:
                offset = end_offset
                continue
            # 解密嵌入 blip 的 header
            key = make_key(password, salt, key_size, 0)
            cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
            decryptor = cipher.decryptor()
            result[pos:pos+8] = decryptor.update(bytes(result[pos:pos+8])) + decryptor.finalize()
            ver_inst2 = int.from_bytes(result[pos:pos+2], 'little')
            rec_type2 = int.from_bytes(result[pos+2:pos+4], 'little')
            rec_inst2 = (ver_inst2 >> 4) & 0x0FFF
            pos += 8
            # 解密嵌入 blip 的字段
            _decrypt_blip_fields(password, salt, key_size, result, pos, end_offset, rec_type2, rec_inst2)
        else:
            # Blip
            _decrypt_blip_fields(password, salt, key_size, result, pos, end_offset, rec_type, rec_inst)

        offset = end_offset
    return bytes(result)


def _decrypt_blip_fields(password, salt, key_size, data, pos, end_offset, rec_type, rec_inst):
    """解密 Blip 的字段"""
    rgb_uid_cnt = 2 if rec_inst in [0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9] else 1

    for _ in range(rgb_uid_cnt):
        key = make_key(password, salt, key_size, 0)
        cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
        decryptor = cipher.decryptor()
        data[pos:pos+16] = decryptor.update(bytes(data[pos:pos+16])) + decryptor.finalize()
        pos += 16

    next_bytes = 34 if rec_type in [0xF01A, 0xF01B, 0xF01C] else 1
    key = make_key(password, salt, key_size, 0)
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    data[pos:pos+next_bytes] = decryptor.update(bytes(data[pos:pos+next_bytes])) + decryptor.finalize()
    pos += next_bytes

    blip_len = end_offset - pos
    if blip_len > 0:
        key = make_key(password, salt, key_size, 0)
        cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
        decryptor = cipher.decryptor()
        data[pos:pos+blip_len] = decryptor.update(bytes(data[pos:pos+blip_len])) + decryptor.finalize()
        pos += blip_len


def main():
    password = "pptx-rs-secret"

    # 原始文件
    original_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt"
    # 加密文件
    encrypted_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt"
    # 解密后的文件
    decrypted_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt.decrypted2.ppt"

    # 用 msoffcrypto 解密
    print("用 msoffcrypto 解密...")
    with open(encrypted_path, "rb") as f:
        msoffcrypto_file = msoffcrypto.OfficeFile(f)
        msoffcrypto_file.load_key(password=password)
        with open(decrypted_path, "wb") as f2:
            msoffcrypto_file.decrypt(f2)
    print(f"解密完成: {decrypted_path}")

    # 读取原始文件和解密后的文件
    ole_orig = olefile.OleFileIO(original_path)
    ole_dec = olefile.OleFileIO(decrypted_path)

    # 检查 Pictures stream
    print("\n--- 检查 Pictures stream ---")
    if ole_orig.exists('Pictures') and ole_dec.exists('Pictures'):
        pics_orig = ole_orig.openstream('Pictures').read()
        pics_dec = ole_dec.openstream('Pictures').read()
        print(f"原始 Pictures stream: {len(pics_orig)} bytes")
        print(f"解密后 Pictures stream: {len(pics_dec)} bytes")

        # msoffcrypto 不解密 Pictures stream，所以解密后的 Pictures stream 应该仍然是加密的
        if pics_orig == pics_dec:
            print("Pictures stream 完全一致（意外！msoffcrypto 应该不解密 Pictures stream）")
        else:
            print("Pictures stream 不一致（预期：msoffcrypto 不解密 Pictures stream）")

            # 用我们自己的解密逻辑解密 Pictures stream
            # 需要从加密文件中读取 salt
            ole_enc = olefile.OleFileIO(encrypted_path)
            ppt_enc = ole_enc.openstream('PowerPoint Document').read()

            # 找到 CryptSession10Container
            # 读取 Current User stream
            cu_enc = ole_enc.openstream('Current User').read()
            offset_to_current_edit = int.from_bytes(cu_enc[16:20], 'little')
            print(f"offsetToCurrentEdit: {offset_to_current_edit}")

            # 读取 UserEditAtom
            ue_buf = io.BytesIO(ppt_enc[offset_to_current_edit:])
            ue_buf.read(8)  # RecordHeader
            ue_buf.read(4)  # lastSlideIdRef
            ue_buf.read(2)  # version
            ue_buf.read(2)  # minorVersion, majorVersion
            ue_buf.read(4)  # offsetLastEdit
            offset_persist_dir = int.from_bytes(ue_buf.read(4), 'little')
            ue_buf.read(4)  # docPersistIdRef
            ue_buf.read(4)  # maxPersistWritten
            ue_buf.read(2)  # lastView
            ue_buf.read(2)  # unused
            encrypt_session_persist_id_ref = int.from_bytes(ue_buf.read(4), 'little')
            print(f"encryptSessionPersistIdRef: {encrypt_session_persist_id_ref}")

            # 读取 PersistDirectoryAtom
            pd_buf = io.BytesIO(ppt_enc[offset_persist_dir:])
            pd_buf.read(8)  # RecordHeader
            pd_data = pd_buf.read()
            # 解析 persist directory
            persist_dir = {}
            pos = 0
            while pos + 4 <= len(pd_data):
                entry_val = int.from_bytes(pd_data[pos:pos+4], 'little')
                persist_id = entry_val & 0xFFFFF
                c_persist = (entry_val >> 20) & 0xFFF
                pos += 4
                for j in range(c_persist):
                    if pos + 4 <= len(pd_data):
                        persist_offset = int.from_bytes(pd_data[pos:pos+4], 'little')
                        persist_dir[persist_id + j] = persist_offset
                        pos += 4

            cs_offset = persist_dir.get(encrypt_session_persist_id_ref)
            print(f"CryptSession10Container offset: {cs_offset}")

            # 读取 CryptSession10Container
            cs_buf = io.BytesIO(ppt_enc[cs_offset:])
            cs_buf.read(8)  # RecordHeader
            cs_data = cs_buf.read()
            cs_info = io.BytesIO(cs_data)
            cs_info.read(4)  # EncryptionVersionInfo
            cs_info.read(4)  # flags
            header_size = int.from_bytes(cs_info.read(4), 'little')
            cs_info.read(header_size)  # EncryptionHeader

            # EncryptionVerifier
            salt_size = int.from_bytes(cs_info.read(4), 'little')
            salt = cs_info.read(16)
            encrypted_verifier = cs_info.read(16)
            verifier_hash_size = int.from_bytes(cs_info.read(4), 'little')
            encrypted_verifier_hash = cs_info.read(20)

            print(f"salt: {salt.hex()}")
            print(f"keySize: 128")

            # 用我们自己的解密逻辑解密 Pictures stream
            print("\n用我们自己的解密逻辑解密 Pictures stream...")
            pics_enc = ole_enc.openstream('Pictures').read()
            pics_decrypted = decrypt_pictures_stream(password, salt, 128, bytearray(pics_enc))

            # 检查解密后的 Pictures stream 是否与原始文件一致
            if pics_decrypted == pics_orig:
                print("✓ 解密后的 Pictures stream 与原始文件一致！加密逻辑正确！")
            else:
                print("✗ 解密后的 Pictures stream 与原始文件不一致！加密逻辑有误！")
                # 找到第一个不同的字节
                for i in range(min(len(pics_decrypted), len(pics_orig))):
                    if pics_decrypted[i] != pics_orig[i]:
                        print(f"  第一个不同的字节在 offset {i}: 解密={pics_decrypted[i]:#x}, 原始={pics_orig[i]:#x}")
                        # 显示周围的字节
                        start = max(0, i - 8)
                        end = min(len(pics_decrypted), i + 8)
                        print(f"  解密 [{start}:{end}]: {pics_decrypted[start:end].hex()}")
                        print(f"  原始 [{start}:{end}]: {pics_orig[start:end].hex()}")
                        break

            # 检查解密后的 Pictures stream 的 record 结构
            print("\n解密后的 Pictures stream record 结构:")
            offset = 0
            record_count = 0
            while offset + 8 <= len(pics_decrypted):
                ver_inst = int.from_bytes(pics_decrypted[offset:offset+2], 'little')
                rec_type = int.from_bytes(pics_decrypted[offset+2:offset+4], 'little')
                rec_len = int.from_bytes(pics_decrypted[offset+4:offset+8], 'little')
                ver = ver_inst & 0x0F
                inst = (ver_inst >> 4) & 0x0FFF
                print(f"  Record #{record_count} at offset {offset}: ver={ver:#x}, inst={inst:#x}, type={rec_type:#06x}, len={rec_len}")
                record_count += 1
                offset += 8 + rec_len
                if record_count > 20:
                    break

    ole_orig.close()
    ole_dec.close()


if __name__ == "__main__":
    main()
