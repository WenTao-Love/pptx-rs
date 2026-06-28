"""深入检查加密文件的 OLE2 结构和 EncryptedSummaryInfo stream 内容。"""
import olefile
import struct
import sys

def check_ole2_structure(fname):
    """检查 OLE2 结构。"""
    print(f'=== {fname} ===')
    try:
        ole = olefile.OleFileIO(fname)
        print(f'  OLE2 结构正常')
    except Exception as e:
        print(f'  OLE2 结构错误: {e}')
        return

    # 列出所有 stream
    print(f'\n  Streams:')
    for stream in ole.listdir():
        name = '/'.join(stream)
        try:
            data = ole.openstream(stream).read()
            print(f'    {name}: {len(data)} bytes')
        except Exception as e:
            print(f'    {name}: 读取失败 {e}')

    # 检查 EncryptedSummaryInfo stream
    try:
        esi = ole.openstream('EncryptedSummaryInfo').read()
        print(f'\n  EncryptedSummaryInfo stream: {len(esi)} bytes')
        print(f'    hex: {esi[:40].hex(" ")}')

        if len(esi) >= 4:
            v_major, v_minor = struct.unpack_from('<HH', esi, 0)
            print(f'    EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}')

        if len(esi) >= 8:
            flags = struct.unpack_from('<I', esi, 4)[0]
            print(f'    EncryptionHeader.flags: 0x{flags:08X}')

        if len(esi) >= 12:
            size_extra = struct.unpack_from('<I', esi, 8)[0]
            print(f'    EncryptionHeader.sizeExtra: {size_extra}')

        if len(esi) >= 16:
            alg_id = struct.unpack_from('<I', esi, 12)[0]
            print(f'    EncryptionHeader.algId: 0x{alg_id:08X}')

        if len(esi) >= 20:
            alg_id_hash = struct.unpack_from('<I', esi, 16)[0]
            print(f'    EncryptionHeader.algIdHash: 0x{alg_id_hash:08X}')

        if len(esi) >= 24:
            key_size = struct.unpack_from('<I', esi, 20)[0]
            print(f'    EncryptionHeader.keySize: {key_size}')

        if len(esi) >= 28:
            provider_type = struct.unpack_from('<I', esi, 24)[0]
            print(f'    EncryptionHeader.providerType: {provider_type}')

        # EncryptionHeader 总长度 = 4(flags) + 4(sizeExtra) + 4(algId) + 4(algIdHash)
        # + 4(keySize) + 4(providerType) + 4(reserved1) + 4(reserved2) + CSPName(variable)
        # CSPName 是 UTF-16LE 字符串
        if len(esi) >= 36:
            reserved1 = struct.unpack_from('<I', esi, 28)[0]
            reserved2 = struct.unpack_from('<I', esi, 32)[0]
            print(f'    EncryptionHeader.reserved1: {reserved1}')
            print(f'    EncryptionHeader.reserved2: {reserved2}')

            # CSPName 从偏移 36 开始
            csp_name_bytes = esi[36:]
            try:
                # 移除尾部的 null 字符
                if csp_name_bytes.endswith(b'\x00\x00'):
                    csp_name_bytes = csp_name_bytes[:-2]
                csp_name = csp_name_bytes.decode('utf-16-le')
                print(f'    EncryptionHeader.CSPName: "{csp_name}"')
            except:
                print(f'    EncryptionHeader.CSPName: (解码失败) {csp_name_bytes[:40].hex(" ")}')

            # 计算 EncryptionHeader 长度
            header_len = 36 + len(csp_name_bytes) + 2  # +2 for null terminator
            print(f'    EncryptionHeader 长度: {header_len}')

            # EncryptionVerifier 从 header_len 开始
            if len(esi) >= header_len + 4:
                salt_size = struct.unpack_from('<I', esi, header_len)[0]
                print(f'    EncryptionVerifier.saltSize: {salt_size}')

            if len(esi) >= header_len + 20:
                salt = esi[header_len + 4: header_len + 20]
                print(f'    EncryptionVerifier.salt: {salt.hex(" ")}')

            if len(esi) >= header_len + 36:
                encrypted_verifier = esi[header_len + 20: header_len + 36]
                print(f'    EncryptionVerifier.encryptedVerifier: {encrypted_verifier.hex(" ")}')

            if len(esi) >= header_len + 40:
                verifier_hash_size = struct.unpack_from('<I', esi, header_len + 36)[0]
                print(f'    EncryptionVerifier.verifierHashSize: {verifier_hash_size}')

            if len(esi) >= header_len + 60:
                encrypted_verifier_hash = esi[header_len + 40: header_len + 60]
                print(f'    EncryptionVerifier.encryptedVerifierHash: {encrypted_verifier_hash.hex(" ")}')

    except Exception as e:
        print(f'  EncryptedSummaryInfo stream 读取失败: {e}')

    ole.close()

# 检查加密文件
for fname in ['_test_out/protected_心理账户理论.ppt', '_test_out/wm_protected_心理账户理论.ppt']:
    check_ole2_structure(fname)
    print()
