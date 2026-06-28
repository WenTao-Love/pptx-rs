#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证生成的 .ppt 文件：
1. 加密文件能否用 msoffcrypto 解密（密码 pptx-rs-secret）
2. 水印文件中是否包含水印文本
3. 加密文件的 CryptSession10Container 结构是否正确（flags、CSPName、keySize）
"""

import sys
import os
import struct
import olefile

PASSWORD = "pptx-rs-secret"
WATERMARK_TEXT = "pptx-rs 水印"
TEST_OUT = "_test_out"


def parse_record_header(data, pos):
    """解析 8 字节 record header。返回 (ver, inst, rec_type, rec_len)。"""
    if pos + 8 > len(data):
        return None
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, pos)
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def find_watermark_text(data):
    """在 PowerPoint Document stream 中搜索水印文本（UTF-16LE）。"""
    watermark_utf16 = WATERMARK_TEXT.encode("utf-16-le")
    count = data.count(watermark_utf16)
    return count


def verify_encryption(filepath):
    """验证加密文件：用 msoffcrypto 尝试解密。"""
    print(f"\n=== 验证加密: {os.path.basename(filepath)} ===", flush=True)
    try:
        import msoffcrypto
        with open(filepath, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if not office_file.is_encrypted():
                print("  [FAIL] 文件未加密", flush=True)
                return False
            try:
                office_file.load_key(password=PASSWORD)
                print(f"  [OK] 密码验证通过（密码: {PASSWORD}）", flush=True)
                # 尝试解密到临时文件
                import io
                out = io.BytesIO()
                office_file.decrypt(out)
                print(f"  [OK] 解密成功，解密后大小: {len(out.getvalue())} 字节", flush=True)
                return True
            except Exception as e:
                print(f"  [FAIL] 密码验证失败: {e}", flush=True)
                return False
    except ImportError:
        print("  [SKIP] msoffcrypto 未安装，跳过解密验证", flush=True)
        return None


def verify_crypt_session(filepath):
    """验证 CryptSession10Container 结构（flags、CSPName、keySize）。"""
    print(f"\n=== 验证 CryptSession 结构: {os.path.basename(filepath)} ===", flush=True)
    ole = olefile.OleFileIO(filepath)
    try:
        with ole.openstream("Current User") as f:
            cu_data = f.read()
        # 检查 headerToken
        header_token = struct.unpack_from("<I", cu_data, 12)[0]
        if header_token == 0xF3D1C4DF:
            print(f"  [OK] headerToken = 0xF3D1C4DF (已加密标记)", flush=True)
        else:
            print(f"  [FAIL] headerToken = 0x{header_token:08X} (期望 0xF3D1C4DF)", flush=True)

        with ole.openstream("PowerPoint Document") as f:
            ppt_data = f.read()

        # 找到 UserEditAtom
        offset_to_current_edit = struct.unpack_from("<I", cu_data, 16)[0]
        ue_offset = offset_to_current_edit
        ver_inst, ue_type, ue_len = struct.unpack_from("<HHI", ppt_data, ue_offset)
        print(f"  UserEditAtom: type=0x{ue_type:04X}, len={ue_len}", flush=True)
        if ue_len == 32:
            print(f"  [OK] UserEditAtom.recLen=32 (加密文件)", flush=True)
        elif ue_len == 28:
            print(f"  [FAIL] UserEditAtom.recLen=28 (未加密)", flush=True)

        # 找到 CryptSession10Container
        # 通过 encryptSessionPersistIdRef 找到
        if ue_len >= 32:
            encrypt_session_ref = struct.unpack_from("<I", ppt_data, ue_offset + 8 + 28)[0]
            print(f"  encryptSessionPersistIdRef = {encrypt_session_ref}", flush=True)

            # 找到 PersistDirectoryAtom
            offset_persist_dir = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
            pd_data = ppt_data[offset_persist_dir + 8:]
            entry_val = struct.unpack_from("<I", pd_data, 0)[0]
            entry_pid = entry_val & 0xFFFFF
            entry_cpersist = (entry_val >> 20) & 0xFFF
            print(f"  PersistDir: persistId={entry_pid}, cPersist={entry_cpersist}", flush=True)

            # 读取所有 persist offsets
            offsets = []
            for i in range(entry_cpersist):
                off = struct.unpack_from("<I", pd_data, 4 + i * 4)[0]
                offsets.append((entry_pid + i, off))

            # 找到 CryptSession10Container 的 offset
            crypt_offset = None
            for pid, off in offsets:
                if pid == encrypt_session_ref:
                    crypt_offset = off
                    break

            if crypt_offset is None:
                print(f"  [FAIL] 找不到 encryptSessionPersistIdRef={encrypt_session_ref} 对应的 persist", flush=True)
            else:
                print(f"  CryptSession10Container offset = {crypt_offset}", flush=True)
                # 解析 CryptSession10Container
                cs_ver_inst, cs_type, cs_len = struct.unpack_from("<HHI", ppt_data, crypt_offset)
                print(f"  CryptSession type=0x{cs_type:04X} (期望 0x2F14), len={cs_len}", flush=True)

                cs_data = ppt_data[crypt_offset + 8: crypt_offset + 8 + cs_len]
                # EncryptionVersionInfo
                v_major, v_minor = struct.unpack_from("<HH", cs_data, 0)
                print(f"  VersionInfo: vMajor={v_major}, vMinor={v_minor}", flush=True)

                # flags
                flags = struct.unpack_from("<I", cs_data, 4)[0]
                print(f"  flags = 0x{flags:08X} (期望 0x0000000C)", flush=True)

                # headerSize
                header_size = struct.unpack_from("<I", cs_data, 8)[0]
                print(f"  headerSize = {header_size}", flush=True)

                # EncryptionHeader
                hdr_start = 12
                eh_flags = struct.unpack_from("<I", cs_data, hdr_start)[0]
                eh_size_extra = struct.unpack_from("<I", cs_data, hdr_start + 4)[0]
                eh_alg_id = struct.unpack_from("<I", cs_data, hdr_start + 8)[0]
                eh_alg_id_hash = struct.unpack_from("<I", cs_data, hdr_start + 12)[0]
                eh_key_size = struct.unpack_from("<I", cs_data, hdr_start + 16)[0]
                eh_provider_type = struct.unpack_from("<I", cs_data, hdr_start + 20)[0]
                print(f"  EH.flags = 0x{eh_flags:08X}", flush=True)
                print(f"  EH.sizeExtra = {eh_size_extra}", flush=True)
                print(f"  EH.algId = 0x{eh_alg_id:08X} (期望 0x00006801)", flush=True)
                print(f"  EH.algIdHash = 0x{eh_alg_id_hash:08X} (期望 0x00008004)", flush=True)
                print(f"  EH.keySize = {eh_key_size} (期望 128)", flush=True)
                print(f"  EH.providerType = 0x{eh_provider_type:08X} (期望 0x00000001)", flush=True)

                # CSPName (从 hdr_start + 32 开始，前面有 8 个 4 字节字段)
                csp_start = hdr_start + 32
                csp_bytes = cs_data[csp_start:]
                # 找到 null 结尾（UTF-16LE 的 \x00\x00）
                null_pos = csp_bytes.find(b"\x00\x00")
                if null_pos >= 0 and null_pos % 2 == 0:
                    csp_name = csp_bytes[:null_pos].decode("utf-16-le")
                    print(f"  EH.cspName = '{csp_name}'", flush=True)
                    if "Enhanced" in csp_name or "Strong" in csp_name:
                        print(f"  [OK] CSPName 支持 128 位密钥", flush=True)
                    elif "Base" in csp_name:
                        print(f"  [FAIL] CSPName 'Base' 只支持 56 位，与 keySize=128 不兼容!", flush=True)
                else:
                    print(f"  [WARN] CSPName 解析失败, null_pos={null_pos}", flush=True)
                    print(f"  CSPName 原始字节前40: {csp_bytes[:40].hex()}", flush=True)
    finally:
        ole.close()


def verify_watermark(filepath):
    """验证水印文件中是否包含水印文本。"""
    print(f"\n=== 验证水印: {os.path.basename(filepath)} ===", flush=True)
    ole = olefile.OleFileIO(filepath)
    try:
        with ole.openstream("PowerPoint Document") as f:
            data = f.read()
        count = find_watermark_text(data)
        if count > 0:
            print(f"  [OK] 找到水印文本 '{WATERMARK_TEXT}'，出现 {count} 次", flush=True)
        else:
            print(f"  [FAIL] 未找到水印文本 '{WATERMARK_TEXT}'", flush=True)
    finally:
        ole.close()


if __name__ == "__main__":
    files = sorted(os.listdir(TEST_OUT))
    for fname in files:
        if not fname.endswith(".ppt"):
            continue
        fpath = os.path.join(TEST_OUT, fname)
        if fname.startswith("protected_"):
            verify_encryption(fpath)
            verify_crypt_session(fpath)
        elif fname.startswith("wm_protected_"):
            verify_watermark(fpath)
            verify_encryption(fpath)
            verify_crypt_session(fpath)
        elif fname.startswith("wm_"):
            verify_watermark(fpath)
