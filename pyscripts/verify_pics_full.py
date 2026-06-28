"""
验证 Pictures stream 的加密/解密是否正确。
按字段解密 Pictures stream，与原始文件比较。
"""
import olefile
import struct
from hashlib import sha1
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher
try:
    from cryptography.hazmat.decrepit.ciphers.algorithms import ARC4
except ImportError:
    from cryptography.hazmat.primitives.ciphers.algorithms import ARC4

PASSWORD = 'pptx-rs-secret'
KEY_BITS = 128

def make_key(password, salt, key_bits, block):
    password_bytes = password.encode('UTF-16LE')
    h0 = sha1(salt + password_bytes).digest()
    blockbytes = struct.pack('<I', block)
    hfinal = sha1(h0 + blockbytes).digest()
    key_bytes = key_bits // 8
    return hfinal[:key_bytes]

def rc4_process(key, data):
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    return decryptor.update(data) + decryptor.finalize()

def read_u16_le(data, off):
    return struct.unpack_from('<H', data, off)[0]

def read_u32_le(data, off):
    return struct.unpack_from('<I', data, off)[0]

# 读取加密文件的 Pictures stream 和 salt
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\protected_心理账户理论.ppt')
enc_pics = bytearray(o.openstream('Pictures').read())
enc_ppt = o.openstream('PowerPoint Document').read()
enc_cu = o.openstream('Current User').read()
o.close()

# 读取原始文件的 Pictures stream
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt')
orig_pics = o.openstream('Pictures').read()
o.close()

# 从 CryptSession10Container 获取 salt
ue_off = read_u32_le(enc_cu, 16)
encrypt_session_pid = read_u32_le(enc_ppt, ue_off + 36)
# 找到 PersistDirectoryAtom
pd_off = read_u32_le(enc_ppt, ue_off + 20)
pd_rl = read_u32_le(enc_ppt, pd_off + 4)
# 解析 persist directory 获取 CryptSession10Container offset
pos = pd_off + 8
entry_val = read_u32_le(enc_ppt, pos)
persist_id_start = entry_val & 0xFFFFF
c_persist = (entry_val >> 20) & 0xFFF
# CryptSession10Container 是最后一个
cs_off = read_u32_le(enc_ppt, pos + 4 + (encrypt_session_pid - persist_id_start) * 4)
# 读取 salt
cs_data = enc_ppt[cs_off + 8:]
header_size = read_u32_le(cs_data, 8)
ev_off = 12 + header_size
salt = bytes(cs_data[ev_off + 4: ev_off + 20])

print(f"Salt: {salt.hex()}")
print(f"enc Pics len: {len(enc_pics)}")
print(f"orig Pics len: {len(orig_pics)}")

# 按字段解密 Pictures stream
# 与 encrypt_pictures_stream 对应：每个字段重置 RC4 流，block=0
dec_pics = bytearray(enc_pics)
offset = 0
field_count = 0
errors = 0

while offset + 8 <= len(dec_pics):
    # 1. 解密 header (8 bytes)，重置 RC4 流（block=0）
    key = make_key(PASSWORD, salt, KEY_BITS, 0)
    dec_header = rc4_process(key, bytes(dec_pics[offset:offset+8]))
    dec_pics[offset:offset+8] = dec_header
    field_count += 1

    # 读取解密后的 header 字段
    ver_inst = read_u16_le(dec_pics, offset)
    rec_type = read_u16_le(dec_pics, offset + 2)
    rlen = read_u32_le(dec_pics, offset + 4)
    rec_inst = (ver_inst >> 4) & 0x0FFF

    if rlen > len(dec_pics) - offset - 8:
        print(f"  ERROR: rlen {rlen} too large at offset {offset}, recType={hex(rec_type)}")
        break

    pos = offset + 8
    end_offset = pos + rlen

    if rec_type == 0xF007:
        # FBSE
        # 读取 cbName（在解密前读取，但 header 已解密，parts 未解密）
        # 先解密 parts，再读取 cbName
        # BLIB_STORE_ENTRY_PARTS = [1,1,16,2,4,4,4,1,1,1,1] = 36 bytes
        parts = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1]
        for part in parts:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+part]))
            dec_pics[pos:pos+part] = dec
            pos += part
            field_count += 1

        # 读取 cbName（在 parts[8] 和 parts[9] 之间，即 pos-2 处）
        # parts 总共 36 bytes，cbName 在 parts[8] 位置（offset 33-34 from start of parts）
        cb_name = read_u16_le(dec_pics, offset + 8 + 33)

        if cb_name > 0:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+cb_name]))
            dec_pics[pos:pos+cb_name] = dec
            pos += cb_name
            field_count += 1

        if pos >= end_offset:
            offset = end_offset
            continue

        # 嵌入 blip
        # 解密 blip header (8 bytes)
        key = make_key(PASSWORD, salt, KEY_BITS, 0)
        dec = rc4_process(key, bytes(dec_pics[pos:pos+8]))
        dec_pics[pos:pos+8] = dec
        field_count += 1

        # 读取嵌入 blip 的 recType 和 recInst
        ver_inst2 = read_u16_le(dec_pics, pos)
        rec_type2 = read_u16_le(dec_pics, pos + 2)
        rec_inst2 = (ver_inst2 >> 4) & 0x0FFF
        pos += 8

        # 解析 rgbUid + metafileHeader/tag + blipLen
        rgb_uid_cnt = 2 if rec_inst2 in [0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9] else 1

        for _ in range(rgb_uid_cnt):
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+16]))
            dec_pics[pos:pos+16] = dec
            pos += 16
            field_count += 1

        # metafileHeader (34 bytes) for EMF/WMF/PICT, or tag (1 byte) for PNG/JPEG/DIB
        if rec_type2 in [0xF01A, 0xF01B, 0xF01C, 0xF01D, 0xF01E, 0xF01F]:
            # Check if it's a metafile type
            is_metafile = rec_type2 in [0xF01A, 0xF01B, 0xF01C]  # EMF, WMF, PICT
            if is_metafile:
                key = make_key(PASSWORD, salt, KEY_BITS, 0)
                dec = rc4_process(key, bytes(dec_pics[pos:pos+34]))
                dec_pics[pos:pos+34] = dec
                pos += 34
            else:
                key = make_key(PASSWORD, salt, KEY_BITS, 0)
                dec = rc4_process(key, bytes(dec_pics[pos:pos+1]))
                dec_pics[pos:pos+1] = dec
                pos += 1
            field_count += 1

        # blipLen (remaining bytes)
        remaining = end_offset - pos
        if remaining > 0:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+remaining]))
            dec_pics[pos:pos+remaining] = dec
            field_count += 1
    else:
        # Blip (0xF01A-0xF01F)
        rgb_uid_cnt = 2 if rec_inst in [0x217, 0x3D5, 0x46B, 0x543, 0x6E1, 0x6E3, 0x6E5, 0x7A9] else 1

        for _ in range(rgb_uid_cnt):
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+16]))
            dec_pics[pos:pos+16] = dec
            pos += 16
            field_count += 1

        is_metafile = rec_type in [0xF01A, 0xF01B, 0xF01C]
        if is_metafile:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+34]))
            dec_pics[pos:pos+34] = dec
            pos += 34
        else:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+1]))
            dec_pics[pos:pos+1] = dec
            pos += 1
        field_count += 1

        remaining = end_offset - pos
        if remaining > 0:
            key = make_key(PASSWORD, salt, KEY_BITS, 0)
            dec = rc4_process(key, bytes(dec_pics[pos:pos+remaining]))
            dec_pics[pos:pos+remaining] = dec
            field_count += 1

    offset = end_offset

print(f"\nTotal fields decrypted: {field_count}")
print(f"Total records: {offset // 8} (approx)")

# 比较解密后的 Pictures stream 与原始文件
mismatches = 0
for i in range(min(len(dec_pics), len(orig_pics))):
    if dec_pics[i] != orig_pics[i]:
        mismatches += 1
        if mismatches <= 5:
            print(f"  MISMATCH at byte {i}: dec={hex(dec_pics[i])} orig={hex(orig_pics[i])}")

print(f"\nTotal mismatches: {mismatches}")
print(f"dec Pics len: {len(dec_pics)}, orig Pics len: {len(orig_pics)}")

if mismatches == 0 and len(dec_pics) == len(orig_pics):
    print("\n*** PICTURES STREAM DECRYPTED CORRECTLY! ***")
else:
    print("\n*** PICTURES STREAM HAS ERRORS! ***")
