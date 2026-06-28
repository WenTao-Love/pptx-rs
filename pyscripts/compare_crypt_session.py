#!/usr/bin/env python3
"""比较 CryptSession10Container 的内容（解密后）。"""
import struct
import olefile
import io
import msoffcrypto

def get_decrypted_ppt(path, password):
    with open(path, "rb") as f:
        officefile = msoffcrypto.OfficeFile(f)
        officefile.load_key(password=password)
        decrypted = io.BytesIO()
        officefile.decrypt(decrypted)
        decrypted.seek(0)
        ole = olefile.OleFileIO(decrypted)
        ppt = ole.openstream("PowerPoint Document").read()
        ole.close()
        return ppt

def parse_crypt_session(ppt, path):
    print(f"\n{path}:")
    cu_off = 0  # Current User stream 通常从 0 开始？不对，这里不读 Current User
    # 直接从 stream 开头搜索 0x2F14（解密后）
    pos = 0
    while pos + 8 <= len(ppt):
        rec_type = struct.unpack_from("<H", ppt, pos + 2)[0]
        if rec_type == 0x2F14:
            rec_len = struct.unpack_from("<I", ppt, pos + 4)[0]
            data = ppt[pos:pos+8+rec_len]
            print(f"  CryptSession10Container @ {pos}, total_len={8+rec_len}")
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
            return
        rec_len = struct.unpack_from("<I", ppt, pos + 4)[0]
        pos += 8 + rec_len
    print("  未找到 CryptSession10Container")

ppt1 = get_decrypted_ppt("_test_out/protected_心理账户理论.ppt", "pptx-rs-secret")
ppt2 = get_decrypted_ppt("_test_out/rc4cryptoapi_password.ppt", "Password1234_")

parse_crypt_session(ppt1, "protected_心理账户理论.ppt")
parse_crypt_session(ppt2, "rc4cryptoapi_password.ppt")
