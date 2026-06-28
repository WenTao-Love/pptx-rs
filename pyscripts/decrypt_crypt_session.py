#!/usr/bin/env python3
"""解密并显示 CryptSession10Container 内容。"""
import struct
import hashlib
import olefile

def make_key(password, salt, key_bits, block):
    password_utf16le = password.encode('utf-16le')
    h0 = hashlib.sha1(salt + password_utf16le).digest()
    hfinal = hashlib.sha1(h0 + struct.pack('<I', block)).digest()
    key_bytes = key_bits // 8
    if key_bits == 40:
        return hfinal[:5] + b'\x00' * 11
    return hfinal[:key_bytes]

class Rc4:
    def __init__(self, key):
        self.s = list(range(256))
        j = 0
        for i in range(256):
            j = (j + self.s[i] + key[i % len(key)]) % 256
            self.s[i], self.s[j] = self.s[j], self.s[i]
        self.i = 0
        self.j = 0

    def process(self, data):
        out = bytearray(data)
        for idx in range(len(out)):
            self.i = (self.i + 1) % 256
            self.j = (self.j + self.s[self.i]) % 256
            self.s[self.i], self.s[self.j] = self.s[self.j], self.s[self.i]
            k = self.s[(self.s[self.i] + self.s[self.j]) % 256]
            out[idx] ^= k
        return bytes(out)

def decrypt_persist_object(password, salt, key_bits, data, persist_id):
    total_len = len(data)
    blocksize = key_bits * (total_len // key_bits + 1)
    result = bytearray()
    offset = 0
    block = persist_id
    while offset < total_len:
        end = min(offset + blocksize, total_len)
        key = make_key(password, salt, key_bits, block)
        rc4 = Rc4(key)
        result.extend(rc4.process(data[offset:end]))
        offset = end
        block += 1
    return bytes(result)

def parse_crypt_session(data):
    print(f"  header: {data[:8].hex()}")
    v_major = struct.unpack_from("<H", data, 8)[0]
    v_minor = struct.unpack_from("<H", data, 10)[0]
    print(f"  vMajor={v_major}, vMinor={v_minor}")
    flags = struct.unpack_from("<I", data, 12)[0]
    header_size = struct.unpack_from("<I", data, 16)[0]
    print(f"  flags=0x{flags:08X}, headerSize={header_size}")
    eh_flags = struct.unpack_from("<I", data, 20)[0]
    size_extra = struct.unpack_from("<I", data, 24)[0]
    alg_id = struct.unpack_from("<I", data, 28)[0]
    alg_id_hash = struct.unpack_from("<I", data, 32)[0]
    key_size = struct.unpack_from("<I", data, 36)[0]
    provider_type = struct.unpack_from("<I", data, 40)[0]
    print(f"  EncryptionHeader: flags=0x{eh_flags:08X}, sizeExtra={size_extra}, algId=0x{alg_id:08X}, algIdHash=0x{alg_id_hash:08X}, keySize={key_size}, providerType={provider_type}")
    csp_start = 48
    csp_end = csp_start
    while csp_end + 2 <= len(data):
        if data[csp_end] == 0 and data[csp_end+1] == 0:
            break
        csp_end += 2
    csp_name = data[csp_start:csp_end].decode('utf-16le', errors='replace')
    print(f"  CSP name: {csp_name!r}")
    ver_off = csp_end + 2
    salt_size = struct.unpack_from("<I", data, ver_off)[0]
    print(f"  saltSize={salt_size}, verifierSize=16, verifierHashSize=20")
    print(f"  total len={len(data)}")

def inspect(path, password):
    print(f"\n{path}:")
    ole = olefile.OleFileIO(path)
    ppt = ole.openstream("PowerPoint Document").read()
    cu = ole.openstream("Current User").read()
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    ole.close()

    ue_off = offset_to_current_edit
    pd_off = struct.unpack_from("<I", ppt, ue_off + 20)[0]
    entry_val = struct.unpack_from("<I", ppt, pd_off + 8)[0]
    persist_id = entry_val & 0xFFFFF
    c_persist = (entry_val >> 20) & 0xFFF
    cs_pid = persist_id + c_persist - 1
    cs_off = struct.unpack_from("<I", ppt, pd_off + 12 + (c_persist - 1) * 4)[0]
    print(f"  CryptSession10Container: pid={cs_pid}, offset={cs_off}")

    # 读取加密的 record（header + data）
    rec_len = struct.unpack_from("<I", ppt, cs_off + 4)[0]
    encrypted = ppt[cs_off:cs_off + 8 + rec_len]
    print(f"  encrypted len={len(encrypted)}")

    # 需要 salt。从加密数据本身无法直接得到 salt。
    # 这里先用 build_crypt_session10_container 的结构反向读取 salt？
    # 实际上 salt 在 CryptSession10Container 里，但它本身被加密了。
    # 所以我们需要从 Verifier 流程获取 salt，这不可能。
    # 改为：调用 msoffcrypto 解密整个文件后读取 salt？
    # msoffcrypto 不暴露 salt。
    print("  无法直接解密 CryptSession10Container 获取 salt（需要密码验证流程）")

inspect("_test_out/protected_心理账户理论.ppt", "pptx-rs-secret")
inspect("_test_out/rc4cryptoapi_password.ppt", "Password1234_")
