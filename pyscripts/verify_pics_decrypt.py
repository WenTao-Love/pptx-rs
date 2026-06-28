#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 Pictures Stream 的加密/解密逻辑。
用 Python 实现 RC4 CryptoAPI 解密 Pictures Stream，与原始数据比较。
"""

import struct
import olefile
import hashlib


def make_key(password, salt, key_bits, block):
    """RC4 CryptoAPI 密钥派生。"""
    password_utf16le = password.encode('utf-16-le')
    h0 = hashlib.sha1(salt + password_utf16le).digest()
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


def decrypt_pic_field(password, salt, key_bits, data, offset, length):
    """解密 Pictures stream 中的一个字段，用 block=0 的 key，重置 RC4 流。"""
    if length == 0:
        return b''
    key = make_key(password, salt, key_bits, 0)
    rc4 = RC4(key)
    chunk = data[offset:offset + length]
    return rc4.process(chunk)


def decrypt_pictures_stream(password, salt, key_bits, data):
    """解密 Pictures stream（与 Rust encrypt_pictures_stream 对应的逆操作）。"""
    result = bytearray(data)
    offset = 0
    while offset + 8 <= len(data):
        # 1. 先解密 header (8 bytes)，重置 RC4 流（block=0）
        dec_header = decrypt_pic_field(password, salt, key_bits, data, offset, 8)
        result[offset:offset + 8] = dec_header

        # 从解密后的 header 读取字段
        ver_inst = struct.unpack_from("<H", dec_header, 0)[0]
        rec_type = struct.unpack_from("<H", dec_header, 2)[0]
        rlen = struct.unpack_from("<I", dec_header, 4)[0]
        rec_inst = (ver_inst >> 4) & 0x0FFF

        pos = offset + 8
        end_offset = pos + rlen

        if rec_type == 0xF007:
            # FBSE (File BLIP Store Entry)
            # 读取 cbName（从解密后的数据）
            # 先解密 BLIB_STORE_ENTRY_PARTS (36 bytes)
            parts = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1]
            for part in parts:
                dec = decrypt_pic_field(password, salt, key_bits, data, pos, part)
                result[pos:pos + part] = dec
                pos += part

            # 读取 cbName（在 pos-3 和 pos-2 处，即 part[8] 和 part[9]）
            # parts 累计: 1+1+16+2+4+4+4+1+1 = 34, 所以 cbName 在 offset+8+33 和 offset+8+34
            cb_name = struct.unpack_from("<H", result, offset + 8 + 33)[0]

            if cb_name > 0:
                dec = decrypt_pic_field(password, salt, key_bits, data, pos, cb_name)
                result[pos:pos + cb_name] = dec
                pos += cb_name

            if pos >= end_offset:
                offset = end_offset
                continue

            # 嵌入 blip：先解密 header (8 bytes)
            dec = decrypt_pic_field(password, salt, key_bits, data, pos, 8)
            result[pos:pos + 8] = dec

            # 从解密后的 header 读取字段
            ver_inst2 = struct.unpack_from("<H", dec, 0)[0]
            rec_type2 = struct.unpack_from("<H", dec, 2)[0]
            rec_inst2 = (ver_inst2 >> 4) & 0x0FFF
            pos += 8

            # 解析 rgbUid + metafileHeader/tag + blipLen
            decrypt_blip_fields(password, salt, key_bits, data, result, pos, end_offset, rec_type2, rec_inst2)
        else:
            # Blip (0xF01A-0xF01F)
            decrypt_blip_fields(password, salt, key_bits, data, result, pos, end_offset, rec_type, rec_inst)

        offset = end_offset

    return bytes(result)


def decrypt_blip_fields(password, salt, key_bits, data, result, pos, end_offset, rec_type, rec_inst):
    """解密 Blip 的字段（rgbUid + metafileHeader/tag + blipLen）。"""
    rgb_uid_cnt = 1
    if rec_inst in (0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9):
        rgb_uid_cnt = 2

    for _ in range(rgb_uid_cnt):
        dec = decrypt_pic_field(password, salt, key_bits, data, pos, 16)
        result[pos:pos + 16] = dec
        pos += 16

    next_bytes = 34 if rec_type in (0xF01A, 0xF01B, 0xF01C) else 1
    dec = decrypt_pic_field(password, salt, key_bits, data, pos, next_bytes)
    result[pos:pos + next_bytes] = dec
    pos += next_bytes

    blip_len = end_offset - pos
    if blip_len > 0:
        dec = decrypt_pic_field(password, salt, key_bits, data, pos, blip_len)
        result[pos:pos + blip_len] = dec
        pos += blip_len


def main():
    password = 'pptx-rs-secret'
    key_bits = 128

    # 读取加密文件
    enc_path = "_test_out/protected_心理账户理论.ppt"
    orig_path = "_test/心理账户理论.ppt"

    # 读取加密文件的 Pictures stream
    ole = olefile.OleFileIO(enc_path)
    has_enc_pics = ole.exists("Pictures")
    if has_enc_pics:
        enc_pics = ole.openstream("Pictures").read()
    else:
        enc_pics = b''
    ole.close()

    # 读取原始文件的 Pictures stream
    ole = olefile.OleFileIO(orig_path)
    has_orig_pics = ole.exists("Pictures")
    if has_orig_pics:
        orig_pics = ole.openstream("Pictures").read()
    else:
        orig_pics = b''
    ole.close()

    print(f"加密文件 Pictures stream: {'存在' if has_enc_pics else '不存在'}")
    if has_enc_pics:
        print(f"  加密后大小: {len(enc_pics)}")
    print(f"原始文件 Pictures stream: {'存在' if has_orig_pics else '不存在'}")
    if has_orig_pics:
        print(f"  原始大小: {len(orig_pics)}")

    if not has_enc_pics or not has_orig_pics:
        print("无法比较：缺少 Pictures stream")
        return

    # 从加密文件读取 salt
    ole = olefile.OleFileIO(enc_path)
    enc_ppt = ole.openstream("PowerPoint Document").read()
    cu = ole.openstream("Current User").read()
    ole.close()

    offsetToCurrentEdit = struct.unpack_from("<I", cu, 16)[0]
    offsetPersistDirectory = struct.unpack_from("<I", enc_ppt, offsetToCurrentEdit + 20)[0]
    encryptSessionPersistIdRef = struct.unpack_from("<I", enc_ppt, offsetToCurrentEdit + 8 + 28)[0]

    # 解析 PersistDirectoryAtom
    pd_len = struct.unpack_from("<I", enc_ppt, offsetPersistDirectory + 4)[0]
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

    # 读取 CryptSession10Container
    cs_offset = persist_entries[encryptSessionPersistIdRef]
    cs_len = struct.unpack_from("<I", enc_ppt, cs_offset + 4)[0]
    cs_data = enc_ppt[cs_offset + 8:cs_offset + 8 + cs_len]
    headerSize = struct.unpack_from("<I", cs_data, 8)[0]
    eh = cs_data[12:12 + headerSize]
    keySize = struct.unpack_from("<I", eh, 16)[0]
    ev_start = 12 + headerSize
    ev = cs_data[ev_start:]
    salt = ev[4:4 + 16]

    print(f"\n加密参数: keySize={keySize} salt={salt.hex()}")

    # 解密 Pictures stream
    print(f"\n解密 Pictures stream...")
    dec_pics = decrypt_pictures_stream(password, salt, keySize, enc_pics)

    # 比较
    if dec_pics == orig_pics:
        print(f"  ✓ Pictures stream 解密成功！与原始数据完全一致（{len(orig_pics)} 字节）")
    else:
        print(f"  ✗ Pictures stream 解密失败！")
        if len(dec_pics) != len(orig_pics):
            print(f"    长度不一致：解密={len(dec_pics)} 原始={len(orig_pics)}")
        diff_count = sum(1 for a, b in zip(dec_pics, orig_pics) if a != b)
        print(f"    差异字节数: {diff_count}/{min(len(dec_pics), len(orig_pics))}")
        # 找前 5 个差异
        diffs = []
        for i in range(min(len(dec_pics), len(orig_pics))):
            if dec_pics[i] != orig_pics[i]:
                diffs.append(i)
                if len(diffs) >= 5:
                    break
        for i in diffs:
            print(f"    字节 {i}: 原始={hex(orig_pics[i])} 解密={hex(dec_pics[i])}")

    # 检查加密前后大小是否一致
    if len(enc_pics) != len(orig_pics):
        print(f"\n  ⚠ 加密后大小变化: {len(orig_pics)} → {len(enc_pics)}")


if __name__ == "__main__":
    main()
