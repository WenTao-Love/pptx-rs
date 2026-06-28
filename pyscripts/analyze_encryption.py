#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""深入分析 PPT 加密结构"""
import olefile
import struct
import os
import sys

def hex_dump(data, offset=0, length=256, prefix='  '):
    """十六进制转储"""
    end = min(offset + length, len(data))
    for i in range(offset, end, 16):
        hex_part = ' '.join(f'{b:02x}' for b in data[i:i+16])
        ascii_part = ''.join(chr(b) if 32 <= b < 127 else '.' for b in data[i:i+16])
        print(f'{prefix}{i:08x}  {hex_part:<48}  {ascii_part}')

def analyze_ole(filename, label):
    print(f"\n{'='*70}")
    print(f"分析: {label}")
    print(f"文件: {filename} ({os.path.getsize(filename)} bytes)")
    print(f"{'='*70}")
    try:
        ole = olefile.OleFileIO(filename, raise_defect=True)
    except Exception as e:
        print(f"OLE 打开失败: {e}")
        # 尝试读取文件头
        with open(filename, 'rb') as f:
            header = f.read(512)
        print("文件头 (前 128 字节):")
        hex_dump(header, 0, 128)
        return None
    print(f"OLE 根目录流:")
    for stream_path in ole.listdir(streams=True, storages=True):
        stream_name = '/'.join(stream_path)
        try:
            data = ole.openstream(stream_path).read()
            print(f"  {stream_name}: {len(data)} bytes")
        except Exception as e:
            print(f"  {stream_name}: [storage or error: {e}]")
    return ole

def analyze_encryption_info(filename):
    """分析 EncryptionInfo 流"""
    print(f"\n{'='*70}")
    print(f"分析加密信息: {filename}")
    print(f"{'='*70}")
    try:
        ole = olefile.OleFileIO(filename)
    except Exception as e:
        print(f"OLE 打开失败: {e}")
        return
    # 检查 EncryptionInfo 流
    if ole.exists('EncryptionInfo'):
        data = ole.openstream('EncryptionInfo').read()
        print(f"\nEncryptionInfo 流 ({len(data)} bytes):")
        hex_dump(data, 0, min(len(data), 256))
        # 解析 EncryptionInfo
        if len(data) >= 8:
            version = struct.unpack('<HH', data[0:4])
            print(f"\n版本: {version[0]}.{version[1]}")
            if version[0] == 0x0002 and version[1] == 0x0002:
                # RC4 CryptoAPI 加密
                print("类型: RC4 CryptoAPI (2.2)")
                if len(data) >= 12:
                    header_len = struct.unpack('<I', data[8:12])[0]
                    print(f"Header 长度: {header_len}")
                    if len(data) >= 12 + header_len:
                        # 读取 EncryptionHeader
                        flags = struct.unpack('<I', data[12:16])[0]
                        size_extra = struct.unpack('<I', data[16:20])[0]
                        alg_id = struct.unpack('<I', data[20:24])[0]
                        alg_id_hash = struct.unpack('<I', data[24:28])[0]
                        key_size = struct.unpack('<I', data[28:32])[0]
                        provider_type = struct.unpack('<I', data[32:36])[0]
                        reserved1 = struct.unpack('<I', data[36:40])[0]
                        reserved2 = struct.unpack('<I', data[40:44])[0]
                        csp_name = data[44:44+header_len-32].decode('utf-16-le', errors='replace')
                        print(f"Flags: 0x{flags:08x}")
                        print(f"Alg ID: 0x{alg_id:08x} ({'RC4' if alg_id == 0x6801 else 'AES' if alg_id == 0x6610 else '未知'})")
                        print(f"Hash Alg: 0x{alg_id_hash:08x} ({'SHA1' if alg_id_hash == 0x8004 else 'MD5' if alg_id_hash == 0x8003 else '未知'})")
                        print(f"Key Size: {key_size} bits")
                        print(f"Provider Type: 0x{provider_type:08x}")
                        print(f"CSP Name: {csp_name!r}")
                        # 读取 EncryptionVerifier
                        verifier_offset = 12 + header_len
                        if len(data) >= verifier_offset + 12:
                            salt_size = struct.unpack('<I', data[verifier_offset:verifier_offset+4])[0]
                            print(f"\nEncryptionVerifier:")
                            print(f"  Salt Size: {salt_size}")
                            salt = data[verifier_offset+4:verifier_offset+4+16]
                            print(f"  Salt (16 bytes): {salt.hex()}")
                            enc_verifier = data[verifier_offset+20:verifier_offset+20+20]
                            print(f"  Encrypted Verifier (20 bytes): {enc_verifier.hex()}")
                            enc_verifier_hash = data[verifier_offset+40:verifier_offset+40+32]
                            print(f"  Encrypted Verifier Hash (32 bytes): {enc_verifier_hash.hex()}")
            elif version[0] == 0x0004:
                print("类型: Agile Encryption (4.x)")
            else:
                print(f"未知版本: {version}")
    else:
        print("\n无 EncryptionInfo 流")
    # 检查 EncryptedSummaryInfo
    if ole.exists('EncryptedSummaryInfo'):
        data = ole.openstream('EncryptedSummaryInfo').read()
        print(f"\nEncryptedSummaryInfo 流 ({len(data)} bytes):")
        hex_dump(data, 0, min(len(data), 128))
    # 检查 EncryptedPackage
    if ole.exists('EncryptedPackage'):
        data = ole.openstream('EncryptedPackage').read()
        print(f"\nEncryptedPackage 流 ({len(data)} bytes)")
        hex_dump(data, 0, min(len(data), 64))
    # 检查 PowerPoint Document
    if ole.exists('PowerPoint Document'):
        data = ole.openstream('PowerPoint Document').read()
        print(f"\nPowerPoint Document 流 ({len(data)} bytes)")
        # 查找水印文本
        wm_text = '机密'.encode('utf-16-le')
        idx = data.find(wm_text)
        print(f"水印文本 '机密' 位置: {idx}")
        if idx >= 0:
            print("水印文本上下文:")
            hex_dump(data, max(0, idx-32), 96)
    ole.close()

if __name__ == '__main__':
    base = '_test_out'
    # 分析原始文件
    orig = '_test/心理账户理论.ppt'
    if os.path.exists(orig):
        analyze_ole(orig, '原始文件')
    # 分析水印文件
    wm_files = [f for f in os.listdir(base) if f.startswith('wm_') and f.endswith('.ppt')]
    for f in wm_files:
        analyze_ole(os.path.join(base, f), '水印文件')
    # 分析加密文件
    prot_files = [f for f in os.listdir(base) if f.startswith('protected_') and f.endswith('.ppt')]
    for f in prot_files:
        analyze_ole(os.path.join(base, f), '加密文件')
        analyze_encryption_info(os.path.join(base, f))
