"""
完整模拟 PowerPoint 的解密流程：
1. 读取 Current User → offsetToCurrentEdit
2. 读取 UserEditAtom → encryptSessionPersistIdRef, offsetPersistDirectory
3. 读取 PersistDirectoryAtom → persist object directory
4. 读取 CryptSession10Container → salt, encryptedVerifier, encryptedVerifierHash
5. 验证密码
6. 解密所有 persist 对象（用 record header 中的 recLen）
7. 解密 Pictures stream
8. 检查解密后的内容是否与原始文件完全一致
"""
import olefile
import struct
import hashlib
from hashlib import sha1
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher
try:
    from cryptography.hazmat.decrepit.ciphers.algorithms import ARC4
except ImportError:
    from cryptography.hazmat.primitives.ciphers.algorithms import ARC4

def read_u32_le(data, off):
    return struct.unpack_from('<I', data, off)[0]

def read_u16_le(data, off):
    return struct.unpack_from('<H', data, off)[0]

def parse_record_header(data, off):
    ver_inst = read_u16_le(data, off)
    rec_type = read_u16_le(data, off + 2)
    rec_len = read_u32_le(data, off + 4)
    return ver_inst & 0xF, (ver_inst >> 4) & 0xFFF, rec_type, rec_len

def parse_persist_directory(data, off):
    _, _, rec_type, rec_len = parse_record_header(data, off)
    assert rec_type == 0x1772, f"Expected 0x1772, got {hex(rec_type)}"
    entries = []
    pos = off + 8
    end = pos + rec_len
    while pos < end:
        entry_val = read_u32_le(data, pos)
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        offsets = []
        for i in range(c_persist):
            offsets.append(read_u32_le(data, pos + 4 + i * 4))
        entries.append((persist_id, c_persist, offsets))
        pos += 4 + c_persist * 4
    return entries

def make_key(password, salt, key_bits, block):
    """与 msoffcrypto _makekey 完全一致。"""
    password_bytes = password.encode('UTF-16LE')
    h0 = sha1(salt + password_bytes).digest()
    blockbytes = struct.pack('<I', block)
    hfinal = sha1(h0 + blockbytes).digest()
    key_bytes = key_bits // 8
    if key_bits == 40:
        return hfinal[:5] + b'\x00' * 11
    return hfinal[:key_bytes]

def rc4_decrypt(key, data):
    cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
    decryptor = cipher.decryptor()
    return decryptor.update(data) + decryptor.finalize()

def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    """与我们的 encrypt_persist_object 对应。"""
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)
    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        chunk = bytes(data[offset:end])
        dec = rc4_decrypt(key, chunk)
        result.extend(dec)
        offset = end
        block += 1
    return bytes(result)

PASSWORD = 'pptx-rs-secret'
KEY_BITS = 128

# 读取加密文件
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\protected_心理账户理论.ppt')
enc_ppt = bytearray(o.openstream('PowerPoint Document').read())
enc_cu = o.openstream('Current User').read()
enc_pics = bytearray(o.openstream('Pictures').read()) if o.exists('Pictures') else None
o.close()

# 读取原始文件
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt')
orig_ppt = o.openstream('PowerPoint Document').read()
orig_pics = o.openstream('Pictures').read() if o.exists('Pictures') else None
o.close()

print(f"=== File sizes ===")
print(f"  enc PPT: {len(enc_ppt)}, orig PPT: {len(orig_ppt)}")
print(f"  enc Pics: {len(enc_pics) if enc_pics else 'N/A'}, orig Pics: {len(orig_pics) if orig_pics else 'N/A'}")

# 步骤1：读取 Current User
print(f"\n=== Step 1: Current User ===")
header_token = read_u32_le(enc_cu, 12)
ue_off = read_u32_le(enc_cu, 16)
print(f"  headerToken: {hex(header_token)} (expect 0xF3D1C4DF for encrypted)")
print(f"  offsetToCurrentEdit: {ue_off}")

# 步骤2：读取 UserEditAtom
print(f"\n=== Step 2: UserEditAtom ===")
_, _, rt, rl = parse_record_header(enc_ppt, ue_off)
print(f"  recType: {hex(rt)} (expect 0x0FF5)")
print(f"  recLen: {rl} (expect 32 for encrypted)")
last_slide_id = read_u32_le(enc_ppt, ue_off + 8)
offset_last_edit = read_u32_le(enc_ppt, ue_off + 16)
offset_persist_dir = read_u32_le(enc_ppt, ue_off + 20)
doc_ref = read_u32_le(enc_ppt, ue_off + 24)
max_persist_written = read_u32_le(enc_ppt, ue_off + 28)
encrypt_session_pid = read_u32_le(enc_ppt, ue_off + 36) if rl >= 32 else None
print(f"  offsetLastEdit: {offset_last_edit}")
print(f"  offsetPersistDirectory: {offset_persist_dir}")
print(f"  maxPersistWritten: {max_persist_written}")
print(f"  encryptSessionPersistIdRef: {encrypt_session_pid}")

# 步骤3：读取 PersistDirectoryAtom
print(f"\n=== Step 3: PersistDirectoryAtom ===")
_, _, pd_rt, pd_rl = parse_record_header(enc_ppt, offset_persist_dir)
print(f"  recType: {hex(pd_rt)} (expect 0x1772)")
print(f"  recLen: {pd_rl} (orig was {pd_rl - 4})")
entries = parse_persist_directory(enc_ppt, offset_persist_dir)
persist_dir = {}
for pid, cpersist, offsets in entries:
    print(f"  persistId={pid}, cPersist={cpersist}")
    for i, off in enumerate(offsets):
        persist_dir[pid + i] = off
print(f"  total persist objects: {len(persist_dir)}")

# 步骤4：读取 CryptSession10Container
print(f"\n=== Step 4: CryptSession10Container ===")
cs_off = persist_dir[encrypt_session_pid]
print(f"  offset: {cs_off} (from persistId={encrypt_session_pid})")
_, _, cs_rt, cs_rl = parse_record_header(enc_ppt, cs_off)
print(f"  recType: {hex(cs_rt)} (expect 0x2F14)")
print(f"  recLen: {cs_rl}")
cs_data = enc_ppt[cs_off + 8: cs_off + 8 + cs_rl]
# 解析 CryptSession10Container data
v_major = read_u16_le(cs_data, 0)
v_minor = read_u16_le(cs_data, 2)
flags = read_u32_le(cs_data, 4)
header_size = read_u32_le(cs_data, 8)
print(f"  vMajor: {v_major}, vMinor: {v_minor}")
print(f"  flags: {hex(flags)}")
print(f"  headerSize: {header_size}")
# EncryptionHeader starts at offset 12
eh_flags = read_u32_le(cs_data, 12)
alg_id = read_u32_le(cs_data, 20)
alg_id_hash = read_u32_le(cs_data, 24)
key_size = read_u32_le(cs_data, 28)
print(f"  EH.flags: {hex(eh_flags)}, algId: {hex(alg_id)}, algIdHash: {hex(alg_id_hash)}, keySize: {key_size}")
# EncryptionVerifier starts after header
ev_off = 12 + header_size
salt_size = read_u32_le(cs_data, ev_off)
salt = bytes(cs_data[ev_off + 4: ev_off + 4 + 16])
enc_verifier = bytes(cs_data[ev_off + 20: ev_off + 36])
verifier_hash_size = read_u32_le(cs_data, ev_off + 36)
enc_verifier_hash = bytes(cs_data[ev_off + 40: ev_off + 60])
print(f"  saltSize: {salt_size}")
print(f"  salt: {salt.hex()}")
print(f"  verifierHashSize: {verifier_hash_size}")

# 步骤5：验证密码
print(f"\n=== Step 5: Verify password ===")
key = make_key(PASSWORD, salt, KEY_BITS, 0)
verifier = rc4_decrypt(key, enc_verifier)
verifier_hash = rc4_decrypt(key, enc_verifier_hash)
# 注意：verifier 和 verifier_hash 用同一个 RC4 流连续解密
key = make_key(PASSWORD, salt, KEY_BITS, 0)
cipher = Cipher(ARC4(key), mode=None, backend=default_backend())
decryptor = cipher.decryptor()
verifier = decryptor.update(enc_verifier)
verifier_hash = decryptor.update(enc_verifier_hash)
expected_hash = sha1(verifier).digest()
print(f"  verifier hash match: {verifier_hash == expected_hash}")

# 步骤6：解密所有 persist 对象
# 关键：record header 已被加密，必须从原始文件读取 recLen
print(f"\n=== Step 6: Decrypt persist objects ===")
dec_ppt = bytearray(enc_ppt)
errors = []
mismatches = 0

for pid in sorted(persist_dir.keys()):
    offset = persist_dir[pid]
    # CryptSession10Container 只存在于加密文件中，跳过
    if pid == encrypt_session_pid:
        print(f"  persistId={pid}: CryptSession10Container (skip)")
        continue
    # 从原始文件读取 record header（明文），因为加密文件中的 header 是密文
    _, _, orig_rt, orig_rl = parse_record_header(orig_ppt, offset)
    if orig_rt in [0x0FF5, 0x1772]:
        print(f"  persistId={pid}: {hex(orig_rt)} (skip, not encrypted)")
        continue
    total_len = 8 + orig_rl
    enc_data = bytes(enc_ppt[offset:offset + total_len])
    dec_data = decrypt_persist_object(PASSWORD, salt, KEY_BITS, enc_data, pid)
    # 检查解密后的 record header
    _, _, dec_rt, dec_rl = parse_record_header(dec_data, 0)
    if dec_rt != orig_rt or dec_rl != orig_rl:
        errors.append(f"persistId={pid}: header mismatch! orig=({hex(orig_rt)},{orig_rl}) dec=({hex(dec_rt)},{dec_rl})")
    # 与原始文件比较
    orig_data = orig_ppt[offset:offset + total_len]
    if dec_data != orig_data:
        mismatches += 1
        if mismatches <= 3:
            for i in range(min(len(dec_data), len(orig_data))):
                if dec_data[i] != orig_data[i]:
                    print(f"  persistId={pid} (offset={offset}, recLen={orig_rl}): MISMATCH at byte {i}, dec={hex(dec_data[i])} orig={hex(orig_data[i])}")
                    break
    dec_ppt[offset:offset + total_len] = dec_data

print(f"  total errors: {len(errors)}")
print(f"  total mismatches: {mismatches}")

# 步骤7：检查 UserEditAtom 和 PersistDirectoryAtom 是否与原始文件一致
print(f"\n=== Step 7: Check UserEditAtom & PersistDirectoryAtom ===")
# 原始 UserEditAtom 在 orig_ppt 中的位置
orig_ue_off = read_u32_le(open(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt','rb').read() if False else b'', 0) if False else None
# 直接从原始 CU 读取
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt')
orig_cu = o.openstream('Current User').read()
o.close()
orig_ue_off = read_u32_le(orig_cu, 16)
print(f"  orig UserEditAtom at: {orig_ue_off}")
print(f"  enc  UserEditAtom at: {ue_off} (diff: {ue_off - orig_ue_off})")

# 比较原始和加密文件中的 UserEditAtom（除了 recLen 和 encryptSessionPersistIdRef）
print(f"  orig UserEditAtom lastSlideIdRef: {read_u32_le(orig_ppt, orig_ue_off + 8)}")
print(f"  enc  UserEditAtom lastSlideIdRef: {read_u32_le(enc_ppt, ue_off + 8)}")
print(f"  orig UserEditAtom offsetLastEdit: {read_u32_le(orig_ppt, orig_ue_off + 16)}")
print(f"  enc  UserEditAtom offsetLastEdit: {read_u32_le(enc_ppt, ue_off + 16)}")
print(f"  orig UserEditAtom offsetPersistDirectory: {read_u32_le(orig_ppt, orig_ue_off + 20)}")
print(f"  enc  UserEditAtom offsetPersistDirectory: {read_u32_le(enc_ppt, ue_off + 20)}")

if errors or mismatches:
    print("\n*** VERIFICATION FAILED! ***")
else:
    print("\n*** ALL PERSIST OBJECTS VERIFIED OK! ***")
