#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 Python 实现 RC4 CryptoAPI 解密，验证加密逻辑是否正确。
解密 persistId=2 的 record，看看是否能恢复原始数据。
"""

import struct
import olefile
import io
import hashlib


def make_key(password, salt, key_bits, block):
    """RC4 CryptoAPI 密钥派生。"""
    password_utf16le = password.encode('utf-16-le')
    # H0 = SHA1(salt + password_utf16le)
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    # Hfinal = SHA1(H0 + LE32(block))
    hfinal = hashlib.sha1(h0 + struct.pack('<I', block)).digest()
    if key_bits == 40:
        key = hfinal[:5] + b'\x00' * 11
    else:
        key = hfinal[:key_bits // 8]
    return key


class RC4:
    def __init__(self, key):
        S = list(range(256))
        j = 0
        for i in range(256):
            j = (j + S[i] + key[i % len(key)]) % 256
            S[i], S[j] = S[j], S[i]
        self.S = S
        self.i = 0
        self.j = 0

    def process(self, data):
        out = bytearray()
        for byte in data:
            self.i = (self.i + 1) % 256
            self.j = (self.j + self.S[self.i]) % 256
            self.S[self.i], self.S[self.j] = self.S[self.j], self.S[self.i]
            k = self.S[(self.S[self.i] + self.S[self.j]) % 256]
            out.append(byte ^ k)
        return bytes(out)


def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    """解密一个 persist 对象。"""
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)

    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        rc4 = RC4(key)
        chunk = data[offset:end]
        decrypted = rc4.process(chunk)
        result.extend(decrypted)
        offset = end
        block += 1
    return bytes(result)


def parse_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def main():
    password = 'pptx-rs-secret'
    key_bits = 128

    # 1. 读取加密文件和原始文件
    enc_path = "_test_out/protected_心理账户理论.ppt"
    orig_path = "_test/心理账户理论.ppt"

    ole = olefile.OleFileIO(enc_path)
    enc_ppt = ole.openstream("PowerPoint Document").read()
    ole.close()

    ole = olefile.OleFileIO(orig_path)
    orig_ppt = ole.openstream("PowerPoint Document").read()
    ole.close()

    # 2. 读取 CryptSession10Container，获取 salt
    # 从 Current User 读取 offsetToCurrentEdit
    ole = olefile.OleFileIO(enc_path)
    cu = ole.openstream("Current User").read()
    ole.close()

    offsetToCurrentEdit = struct.unpack_from("<I", cu, 16)[0]
    print(f"offsetToCurrentEdit = {offsetToCurrentEdit}")

    # 读取 UserEditAtom
    ue_type = struct.unpack_from("<H", enc_ppt, offsetToCurrentEdit + 2)[0]
    ue_len = struct.unpack_from("<I", enc_ppt, offsetToCurrentEdit + 4)[0]
    print(f"UserEditAtom: type=0x{ue_type:04X} len={ue_len}")

    offsetPersistDirectory = struct.unpack_from("<I", enc_ppt, offsetToCurrentEdit + 20)[0]
    print(f"offsetPersistDirectory = {offsetPersistDirectory}")

    encryptSessionPersistIdRef = struct.unpack_from("<I", enc_ppt, offsetToCurrentEdit + 8 + 28)[0]
    print(f"encryptSessionPersistIdRef = {encryptSessionPersistIdRef}")

    # 解析 PersistDirectoryAtom
    pd_type = struct.unpack_from("<H", enc_ppt, offsetPersistDirectory + 2)[0]
    pd_len = struct.unpack_from("<I", enc_ppt, offsetPersistDirectory + 4)[0]
    print(f"PersistDirectoryAtom: type=0x{pd_type:04X} len={pd_len}")

    pd_data = enc_ppt[offsetPersistDirectory + 8:offsetPersistDirectory + 8 + pd_len]
    persist_entries = {}
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                po = struct.unpack_from("<I", pd_data, pos)[0]
                persist_entries[persist_id + j] = po
                pos += 4

    print(f"persist entries 数量: {len(persist_entries)}")
    print(f"persistId=2 offset={persist_entries.get(2)}")
    print(f"persistId={encryptSessionPersistIdRef} offset={persist_entries.get(encryptSessionPersistIdRef)}")

    # 读取 CryptSession10Container
    cs_offset = persist_entries[encryptSessionPersistIdRef]
    cs_type = struct.unpack_from("<H", enc_ppt, cs_offset + 2)[0]
    cs_len = struct.unpack_from("<I", enc_ppt, cs_offset + 4)[0]
    print(f"\nCryptSession10Container: offset={cs_offset} type=0x{cs_type:04X} len={cs_len}")

    cs_data = enc_ppt[cs_offset + 8:cs_offset + 8 + cs_len]
    # EncryptionVersionInfo
    vMajor = struct.unpack_from("<H", cs_data, 0)[0]
    vMinor = struct.unpack_from("<H", cs_data, 2)[0]
    flags = struct.unpack_from("<I", cs_data, 4)[0]
    headerSize = struct.unpack_from("<I", cs_data, 8)[0]
    print(f"  vMajor={vMajor} vMinor={vMinor} flags=0x{flags:08X} headerSize={headerSize}")

    # EncryptionHeader
    eh = cs_data[12:12 + headerSize]
    eh_flags = struct.unpack_from("<I", eh, 0)[0]
    sizeExtra = struct.unpack_from("<I", eh, 4)[0]
    algId = struct.unpack_from("<I", eh, 8)[0]
    algIdHash = struct.unpack_from("<I", eh, 12)[0]
    keySize = struct.unpack_from("<I", eh, 16)[0]
    providerType = struct.unpack_from("<I", eh, 20)[0]
    print(f"  eh_flags=0x{eh_flags:08X} sizeExtra={sizeExtra}")
    print(f"  algId=0x{algId:08X} algIdHash=0x{algIdHash:08X}")
    print(f"  keySize={keySize} providerType={providerType}")

    # EncryptionVerifier
    ev_start = 12 + headerSize
    ev = cs_data[ev_start:]
    saltSize = struct.unpack_from("<I", ev, 0)[0]
    salt = ev[4:4 + 16]
    encryptedVerifier = ev[20:20 + 16]
    verifierHashSize = struct.unpack_from("<I", ev, 36)[0]
    encryptedVerifierHash = ev[40:40 + 20]
    print(f"  saltSize={saltSize} verifierHashSize={verifierHashSize}")
    print(f"  salt: {salt.hex()}")
    print(f"  encryptedVerifier: {encryptedVerifier.hex()}")

    # 3. 解密 persistId=2 的 record
    # 关键修复：先从原始文件读取 header 获取正确长度（加密后 header 是密文）
    pid = 2
    poff = persist_entries[pid]

    # 从原始文件读取 header（明文）
    orig_h = parse_header(orig_ppt, poff)
    total_len = 8 + orig_h[3]
    print(f"\n解密 persistId={pid} at offset={poff}:")
    print(f"  原始 header: ver=0x{orig_h[0]:X} inst=0x{orig_h[1]:03X} type=0x{orig_h[2]:04X} len={orig_h[3]}")
    print(f"  total_len = {total_len}")

    enc_record = enc_ppt[poff:poff + total_len]
    dec_record = decrypt_persist_object(password, salt, keySize, enc_record, pid)

    dec_h = parse_header(dec_record, 0)
    print(f"  解密后 header: ver=0x{dec_h[0]:X} inst=0x{dec_h[1]:03X} type=0x{dec_h[2]:04X} len={dec_h[3]}")

    # 4. 比较解密结果和原始数据（只比较 record 范围内）
    orig_record = orig_ppt[poff:poff + total_len]
    if len(dec_record) != len(orig_record):
        print(f"  ✗ 长度不一致！解密={len(dec_record)} 原始={len(orig_record)}")
    elif dec_record == orig_record:
        print(f"  ✓ 解密成功！与原始数据完全一致（{total_len} 字节）")
    else:
        print(f"  ✗ 解密失败！与原始数据不一致")
        diff_count = 0
        first_diff = -1
        for i in range(min(len(dec_record), len(orig_record))):
            if dec_record[i] != orig_record[i]:
                diff_count += 1
                if first_diff < 0:
                    first_diff = i
                    print(f"  第一个差异在字节 {i}: 原始={hex(orig_record[i])} 解密={hex(dec_record[i])}")
        print(f"  总差异字节数: {diff_count}/{total_len}")

    # 5. 验证密码
    print(f"\n验证密码:")
    key_block0 = make_key(password, salt, keySize, 0)
    rc4 = RC4(key_block0)
    dec_verifier = rc4.process(encryptedVerifier)
    rc4 = RC4(key_block0)  # 重置 RC4
    # 实际上应该用同一个 RC4 流连续加密
    rc4 = RC4(key_block0)
    dec_verifier = rc4.process(encryptedVerifier)
    dec_verifier_hash = rc4.process(encryptedVerifierHash)

    # 验证：SHA1(dec_verifier) == dec_verifier_hash?
    actual_hash = hashlib.sha1(dec_verifier).digest()
    if actual_hash == dec_verifier_hash:
        print(f"  ✓ 密码验证成功！")
    else:
        print(f"  ✗ 密码验证失败！")
        print(f"    SHA1(dec_verifier): {actual_hash.hex()}")
        print(f"    dec_verifier_hash: {dec_verifier_hash.hex()}")

    # 6. 验证所有 persist 对象的加密/解密
    print(f"\n验证所有 persist 对象:")
    success_count = 0
    fail_count = 0
    skip_count = 0
    for pid in sorted(persist_entries.keys()):
        poff = persist_entries[pid]
        # 从原始文件读取 header
        orig_h = parse_header(orig_ppt, poff)
        if orig_h is None:
            print(f"  pid={pid} offset={poff}: 原始文件 header 读取失败")
            fail_count += 1
            continue
        ver, inst, rec_type, rec_len = orig_h

        # 跳过 UserEditAtom 和 PersistDirectoryAtom（不加密）
        if rec_type == 0x0FF5 or rec_type == 0x1772:
            skip_count += 1
            continue

        # 跳过 CryptSession10Container（加密后新增的）
        if rec_type == 0x2F14:
            skip_count += 1
            continue

        total_len = 8 + rec_len
        if poff + total_len > len(enc_ppt):
            print(f"  pid={pid} offset={poff} type=0x{rec_type:04X} len={rec_len}: 超出加密文件范围")
            fail_count += 1
            continue

        enc_record = enc_ppt[poff:poff + total_len]
        dec_record = decrypt_persist_object(password, salt, keySize, enc_record, pid)
        orig_record = orig_ppt[poff:poff + total_len]

        if dec_record == orig_record:
            success_count += 1
        else:
            fail_count += 1
            diff_count = sum(1 for a, b in zip(dec_record, orig_record) if a != b)
            print(f"  ✗ pid={pid} offset={poff} type=0x{rec_type:04X} len={rec_len}: 失败 ({diff_count}/{total_len} 字节不同)")

    print(f"\n汇总: 成功={success_count} 失败={fail_count} 跳过={skip_count}")


if __name__ == "__main__":
    main()
