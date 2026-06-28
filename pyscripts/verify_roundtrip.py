"""完整 round-trip 验证：解密 PowerPoint Document stream 中的 persist 对象，与原始对比。

验证逻辑：
1. 从 CryptSession10Container 提取 salt
2. 解析 CurrentUser → offsetToCurrentEdit
3. 解析 UserEditAtom → offsetPersistDirectory
4. 解析 PersistDirectoryAtom → persist entries (pid, offset)
5. 对于每个 persist 对象，用 block=pid 的 key 解密
6. 跳过 UserEditAtom 和 PersistDirectoryAtom（不加密）
7. 对比解密后的 persist 对象与原始文件中对应位置的数据
"""
import io
import os
import sys
import struct
import hashlib
import olefile


PASSWORD = 'pptx-rs-secret'
KEY_BITS = 128
KEY_SIZE = KEY_BITS // 8
SALT_SIZE = 16

RT_USER_EDIT_ATOM = 0x0FF5
RT_PERSIST_DIRECTORY_ATOM = 0x1772
RT_CURRENT_USER_ATOM = 0x0FF6
RT_CRYPT_SESSION10_CONTAINER = 0x2F14


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
    """从加密文件中读取 salt。"""
    ole = olefile.OleFileIO(io.BytesIO(encrypted_data))
    try:
        ppt = ole.openstream('PowerPoint Document').read()
        pattern = b'\x0f\x00\x14\x2f'
        idx = ppt.rfind(pattern)
        if idx < 0:
            return None
        offset = idx
        pos = offset + 8  # 跳过 header
        pos += 4  # EncryptionVersionInfo
        pos += 4  # flags
        header_size = struct.unpack_from('<I', ppt, pos)[0]
        pos += 4  # headerSize 字段本身
        pos += header_size  # 跳过整个 EncryptionHeader
        salt_size = struct.unpack_from('<I', ppt, pos)[0]
        pos += 4  # 跳过 saltSize
        salt = ppt[pos:pos + salt_size]
        return salt
    finally:
        ole.close()


def parse_record_header(data, offset):
    """解析 record header，返回 (ver, inst, rec_type, rec_len)。"""
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom，返回 [(persistId, offset), ...]。"""
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    if rec_type != RT_PERSIST_DIRECTORY_ATOM:
        raise ValueError(f'Expected PersistDirectoryAtom, got 0x{rec_type:04X}')
    pd_data = data[offset + 8:offset + 8 + rec_len]
    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from('<I', pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = struct.unpack_from('<I', pd_data, pos)[0]
                entries.append((persist_id + j, persist_offset))
                pos += 4
    return entries


def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    """解密 persist 对象。"""
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)
    result = bytearray(data)
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        cipher = RC4(key)
        chunk = cipher.process(bytes(result[offset:end]))
        result[offset:end] = chunk
        offset = end
        block += 1
    return bytes(result)


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
    for f in os.listdir('_test_out'):
        if f.startswith('protected_') and f.endswith('.ppt') and '.' not in f[len('protected_'):-4]:
            prot_path = '_test_out/' + f
            break
    if not prot_path:
        print('找不到 protected 文件')
        sys.exit(1)

    with open(prot_path, 'rb') as f:
        enc_data = f.read()
    print('加密文件:', prot_path, '大小:', len(enc_data))

    # 3. 提取 salt
    salt = get_salt_from_encrypted(enc_data)
    if salt is None:
        print('找不到 salt!')
        sys.exit(1)
    print('Salt:', salt.hex())

    # 4. 读取原始和加密的 PowerPoint Document stream
    orig_ppt = get_stream(orig_data, 'PowerPoint Document')
    enc_ppt = get_stream(enc_data, 'PowerPoint Document')
    print('原始 PPT Doc 大小:', len(orig_ppt))
    print('加密 PPT Doc 大小:', len(enc_ppt))

    # 5. 从原始文件解析 CurrentUser → offsetToCurrentEdit
    orig_cu = get_stream(orig_data, 'Current User')
    orig_offset_to_current_edit = struct.unpack_from('<I', orig_cu, 16)[0]
    print('原始 offsetToCurrentEdit:', orig_offset_to_current_edit)

    # 6. 从原始文件解析 UserEditAtom
    orig_ue_offset = orig_offset_to_current_edit
    ver, inst, ue_type, ue_len = parse_record_header(orig_ppt, orig_ue_offset)
    if ue_type != RT_USER_EDIT_ATOM:
        print(f'错误：期望 UserEditAtom (0x{RT_USER_EDIT_ATOM:04X})，得到 0x{ue_type:04X}')
        sys.exit(1)
    print(f'原始 UserEditAtom: offset={orig_ue_offset}, recLen={ue_len}')
    orig_offset_persist_dir = struct.unpack_from('<I', orig_ppt, orig_ue_offset + 20)[0]
    print('原始 offsetPersistDirectory:', orig_offset_persist_dir)

    # 7. 从原始文件解析 PersistDirectoryAtom
    persist_entries = parse_persist_directory(orig_ppt, orig_offset_persist_dir)
    print(f'Persist entries: {len(persist_entries)} 个')

    # 8. 解密每个 persist 对象，与原始对比
    # 注意：persist 对象的 offset 在加密过程中不变（in-place 加密）
    # 但 UserEditAtom 的 offset 会 +4（因为 PersistDirectoryAtom 增长了 4 字节）
    print('\n=== 验证 persist 对象 ===')
    all_match = True
    for pid, poff in persist_entries:
        poff = int(poff)
        if poff + 8 > len(orig_ppt):
            print(f'  pid={pid} offset={poff}: 原始文件中超出范围')
            all_match = False
            continue

        # 从原始文件读取 record header（获取 rec_type 和 rec_len）
        ver, inst, rec_type, rec_len = parse_record_header(orig_ppt, poff)

        # 跳过 UserEditAtom 和 PersistDirectoryAtom（不加密）
        if rec_type == RT_USER_EDIT_ATOM or rec_type == RT_PERSIST_DIRECTORY_ATOM:
            print(f'  pid={pid} offset={poff}: 跳过 (type=0x{rec_type:04X})')
            continue

        total_len = 8 + rec_len
        if poff + total_len > len(orig_ppt):
            print(f'  pid={pid} offset={poff} type=0x{rec_type:04X}: 原始文件中 record 超出范围')
            all_match = False
            continue
        if poff + total_len > len(enc_ppt):
            print(f'  pid={pid} offset={poff} type=0x{rec_type:04X}: 加密文件中 record 超出范围')
            all_match = False
            continue

        # 从加密文件中读取对应位置的数据，解密
        enc_record = enc_ppt[poff:poff + total_len]
        dec_record = decrypt_persist_object(PASSWORD, salt, KEY_BITS, enc_record, pid)

        # 与原始对比
        orig_record = orig_ppt[poff:poff + total_len]
        if dec_record == orig_record:
            print(f'  pid={pid} offset={poff} type=0x{rec_type:04X} len={total_len}: 一致 ✓')
        else:
            print(f'  pid={pid} offset={poff} type=0x{rec_type:04X} len={total_len}: 不一致 ✗')
            min_len = min(len(dec_record), len(orig_record))
            for i in range(min_len):
                if dec_record[i] != orig_record[i]:
                    print(f'    第一个差异在字节 {i}: 原始=0x{orig_record[i]:02X} 解密=0x{dec_record[i]:02X}')
                    break
            all_match = False

    if all_match:
        print('\n*** 所有 persist 对象解密后与原始一致! 加密正确! ***')
    else:
        print('\n*** 部分 persist 对象解密后与原始不一致! ***')


if __name__ == '__main__':
    main()
