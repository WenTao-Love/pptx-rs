"""验证 .ppt 加密文件的 Pictures Stream：用 RC4 CryptoAPI 按字段解密，与原始比较。"""
import io
import os
import sys
import struct
import hashlib
import olefile


PASSWORD = 'pptx-rs-secret'
KEY_BITS = 128
KEY_SIZE = KEY_BITS // 8  # 16 bytes
SALT_SIZE = 16


class RC4:
    """简单的 RC4 实现。"""

    def __init__(self, key):
        self.S = list(range(256))
        j = 0
        for i in range(256):
            j = (j + self.S[i] + key[i % len(key)]) & 0xFF
            self.S[i], self.S[j] = self.S[j], self.S[i]
        self.i = 0
        self.j = 0

    def process(self, data):
        out = bytearray(data)
        for k in range(len(out)):
            self.i = (self.i + 1) & 0xFF
            self.j = (self.j + self.S[self.i]) & 0xFF
            self.S[self.i], self.S[self.j] = self.S[self.j], self.S[self.i]
            out[k] ^= self.S[(self.S[self.i] + self.S[self.j]) & 0xFF]
        return bytes(out)


def make_key(password, salt, key_bits, block):
    """生成 RC4 密钥：H0 = SHA1(salt + password_utf16le)，Hfinal = SHA1(H0 + LE32(block))。"""
    key_size = key_bits // 8
    password_utf16le = password.encode('utf-16-le')
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    hfinal = hashlib.sha1(h0 + struct.pack('<I', block)).digest()
    return hfinal[:key_size]


def rc4_crypt(key, data):
    """RC4 加密/解密。"""
    cipher = RC4(key)
    return cipher.process(data)


def decrypt_pic_field(password, salt, key_bits, data, offset, length):
    """解密 Pictures stream 中的一个字段，用 block=0 的 key，重置 RC4 流。"""
    if length == 0:
        return
    key = make_key(password, salt, key_bits, 0)
    decrypted = rc4_crypt(key, data[offset:offset + length])
    data[offset:offset + length] = decrypted


def get_stream(data, stream_name):
    """从 OLE2 数据中读取指定 stream 的内容。"""
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        if ole.exists(stream_name):
            return ole.openstream(stream_name).read()
        return None
    finally:
        ole.close()


def get_salt_from_encrypted(encrypted_data):
    """从加密文件中读取 salt（在 CryptSession10Container 中）。

    CryptSession10Container 的 header: ver=0xF, inst=0, type=0x2F14
    即前 4 字节为 0x0F 0x00 0x14 0x2F（小端）。
    """
    ole = olefile.OleFileIO(io.BytesIO(encrypted_data))
    try:
        ppt = ole.openstream('PowerPoint Document').read()
        # 搜索 CryptSession10Container 的 header 模式
        pattern = b'\x0f\x00\x14\x2f'
        idx = ppt.rfind(pattern)  # 从后往前找（通常在末尾）
        if idx < 0:
            return None
        offset = idx
        # 读取 rlen
        rlen = struct.unpack_from('<I', ppt, offset + 4)[0]
        # CryptSession10Container 结构:
        # header(8) + EncryptionVersionInfo(4: vMajor=2,vMinor=2) + flags(4) + headerSize(4)
        # + EncryptionHeader(flags(4)+sizeExtra(4)+algId(4)+algIdHash(4)+keySize(4)+providerType(4)+reserved1(4)+reserved2(4)+CSPName(variable))
        # + salt(16) + verifier(16) + verifierHash(20)
        pos = offset + 8  # 跳过 header
        pos += 4  # EncryptionVersionInfo
        pos += 4  # flags
        header_size = struct.unpack_from('<I', ppt, pos)[0]
        pos += 4  # headerSize 字段本身
        # pos 现在指向 EncryptionHeader 开始，用 headerSize 跳过整个 EncryptionHeader
        pos += header_size
        # 现在 pos 指向 EncryptionVerifier: saltSize(4) + salt(16) + ...
        salt_size = struct.unpack_from('<I', ppt, pos)[0]
        pos += 4  # 跳过 saltSize
        salt = ppt[pos:pos + salt_size]
        if len(salt) < salt_size:
            return None
        return salt
    finally:
        ole.close()


def decrypt_pictures_stream(password, salt, key_bits, data):
    """按字段解密 Pictures Stream，每个字段重置 RC4 流（block=0）。"""
    data = bytearray(data)
    offset = 0
    while offset + 8 <= len(data):
        # 1. 先读取 header 字段（当前是加密状态，需要先解密 header 才能读取）
        # 解密 header
        decrypt_pic_field(password, salt, key_bits, data, offset, 8)

        ver_inst = struct.unpack_from('<H', data, offset)[0]
        rec_type = struct.unpack_from('<H', data, offset + 2)[0]
        rlen = struct.unpack_from('<I', data, offset + 4)[0]
        rec_inst = (ver_inst >> 4) & 0x0FFF

        pos = offset + 8
        end_offset = pos + rlen

        if rec_type == 0xF007:
            # FBSE
            # 先读取 cbName（解密 parts 之前）
            cb_name = struct.unpack_from('<H', data, pos + 33)[0]

            # 解密 BLIB_STORE_ENTRY_PARTS
            parts = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1]
            for part in parts:
                decrypt_pic_field(password, salt, key_bits, data, pos, part)
                pos += part

            # 解密 cbName 字段
            if cb_name > 0:
                decrypt_pic_field(password, salt, key_bits, data, pos, cb_name)
                pos += cb_name

            if pos >= end_offset:
                offset = end_offset
                continue

            # 嵌入 blip：先解密 header，再读取
            decrypt_pic_field(password, salt, key_bits, data, pos, 8)
            ver_inst2 = struct.unpack_from('<H', data, pos)[0]
            rec_type2 = struct.unpack_from('<H', data, pos + 2)[0]
            rec_inst2 = (ver_inst2 >> 4) & 0x0FFF
            pos += 8

            # 解密 blip 字段
            decrypt_blip_fields(password, salt, key_bits, data, pos, end_offset, rec_type2, rec_inst2)
        else:
            # Blip
            decrypt_blip_fields(password, salt, key_bits, data, pos, end_offset, rec_type, rec_inst)

        offset = end_offset
    return bytes(data)


def decrypt_blip_fields(password, salt, key_bits, data, pos, end_offset, rec_type, rec_inst):
    """解密 Blip 的字段。"""
    rgb_uid_cnt = 2 if rec_inst in [0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9] else 1

    for _ in range(rgb_uid_cnt):
        decrypt_pic_field(password, salt, key_bits, data, pos, 16)
        pos += 16

    next_bytes = 34 if rec_type in [0xF01A, 0xF01B, 0xF01C] else 1
    decrypt_pic_field(password, salt, key_bits, data, pos, next_bytes)
    pos += next_bytes

    blip_len = end_offset - pos
    if blip_len > 0:
        decrypt_pic_field(password, salt, key_bits, data, pos, blip_len)
        pos += blip_len


def main():
    # 1. 读取原始文件
    orig_path = None
    for f in os.listdir('_test'):
        if f.endswith('.ppt'):
            orig_path = '_test/' + f
            break
    if not orig_path:
        print('找不到原始 .ppt 文件')
        sys.exit(1)

    with open(orig_path, 'rb') as f:
        orig_data = f.read()
    print('原始文件:', orig_path, '大小:', len(orig_data))

    # 2. 读取加密文件
    prot_path = None
    prefix = sys.argv[1] if len(sys.argv) > 1 else 'protected_'
    for f in os.listdir('_test_out'):
        if f.startswith(prefix) and f.endswith('.ppt') and '.' not in f[len(prefix):-4]:
            prot_path = '_test_out/' + f
            break
    if not prot_path:
        print('找不到', prefix, '文件')
        sys.exit(1)

    with open(prot_path, 'rb') as f:
        enc_data = f.read()
    print('加密文件:', prot_path, '大小:', len(enc_data))

    # 3. 从加密文件中提取 salt
    salt = get_salt_from_encrypted(enc_data)
    if salt is None:
        print('找不到 salt!')
        sys.exit(1)
    print('Salt:', salt.hex())

    # 4. 读取原始和加密的 Pictures Stream
    orig_pics = get_stream(orig_data, 'Pictures')
    enc_pics = get_stream(enc_data, 'Pictures')

    if orig_pics is None:
        print('原始文件没有 Pictures Stream')
        sys.exit(0)
    if enc_pics is None:
        print('加密文件没有 Pictures Stream（异常!）')
        sys.exit(1)

    print('原始 Pictures 大小:', len(orig_pics))
    print('加密 Pictures 大小:', len(enc_pics))

    # 5. 解密加密的 Pictures Stream
    dec_pics = decrypt_pictures_stream(PASSWORD, salt, KEY_BITS, enc_pics)

    # 6. 比较
    if dec_pics == orig_pics:
        print('\n*** Pictures Stream 解密后与原始一致! 加密正确! ***')
    else:
        print('\n*** Pictures Stream 解密后与原始不一致! ***')
        print('  解密后大小:', len(dec_pics), '原始大小:', len(orig_pics))
        min_len = min(len(dec_pics), len(orig_pics))
        diff_count = 0
        first_diff = None
        for i in range(min_len):
            if dec_pics[i] != orig_pics[i]:
                if first_diff is None:
                    first_diff = i
                diff_count += 1
        if first_diff is not None:
            print('  第一个差异在字节', first_diff, ': 原始=', hex(orig_pics[first_diff]), '解密=', hex(dec_pics[first_diff]))
            print('  总差异字节数:', diff_count)
            # 显示前 32 字节的差异
            start = max(0, first_diff - 8)
            end = min(min_len, first_diff + 24)
            print('  原始 [', start, '-', end, ']:', orig_pics[start:end].hex())
            print('  解密 [', start, '-', end, ']:', dec_pics[start:end].hex())


if __name__ == '__main__':
    main()
