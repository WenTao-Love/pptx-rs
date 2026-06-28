"""全面检查加密文件的结构，对比 msoffcrypto 测试文件。

检查项：
1. CurrentUserAtom: headerToken, offsetToCurrentEdit
2. UserEditAtom: 所有字段
3. PersistDirectoryAtom: 所有字段
4. CryptSession10Container: 所有字段
5. persist 对象的加密状态
"""
import olefile
import io
from struct import unpack, pack
from collections import namedtuple


def parse_record_header(blob):
    getBitSlice = lambda bits, i, w: (bits & (2**w - 1 << i)) >> i
    buf = blob.read(2)
    (val,) = unpack("<H", buf)
    recVer = getBitSlice(val, 0, 4)
    recInstance = getBitSlice(val, 4, 12)
    (recType,) = unpack("<H", blob.read(2))
    (recLen,) = unpack("<I", blob.read(4))
    return recVer, recInstance, recType, recLen


def check_file(ppt_path, label):
    print(f"\n{'='*70}")
    print(f"文件: {ppt_path}")
    print(f"标签: {label}")
    print(f"{'='*70}")

    if not os.path.exists(ppt_path):
        print(f"文件不存在！")
        return

    ole = olefile.OleFileIO(ppt_path)

    # 读取 Current User stream
    cu_data = ole.openstream('Current User').read()
    print(f"\n--- Current User stream ({len(cu_data)} bytes) ---")
    cu_buf = io.BytesIO(cu_data)
    recVer, recInstance, recType, recLen = parse_record_header(cu_buf)
    print(f"RecordHeader: ver={recVer:#x}, inst={recInstance:#x}, type={recType:#06x}, len={recLen}")

    # CurrentUserAtom 字段
    (size,) = unpack("<I", cu_buf.read(4))
    (headerToken,) = unpack("<I", cu_buf.read(4))
    (offsetToCurrentEdit,) = unpack("<I", cu_buf.read(4))
    (lenUserName,) = unpack("<H", cu_buf.read(2))
    (docFileVersion,) = unpack("<H", cu_buf.read(2))
    (majorVersion, minorVersion) = unpack("<BB", cu_buf.read(2))
    print(f"size: {size:#x}")
    print(f"headerToken: {headerToken:#010x} (期望 0xF3D1C4DF 加密 / 0xE391C05F 未加密)")
    print(f"offsetToCurrentEdit: {offsetToCurrentEdit} ({offsetToCurrentEdit:#x})")
    print(f"lenUserName: {lenUserName}")
    print(f"docFileVersion: {docFileVersion:#x}")
    print(f"majorVersion: {majorVersion}, minorVersion: {minorVersion}")

    # 读取 PowerPoint Document stream
    ppt_data = ole.openstream('PowerPoint Document').read()
    print(f"\n--- PowerPoint Document stream ({len(ppt_data)} bytes) ---")

    # 解析 UserEditAtom
    print(f"\n--- UserEditAtom at offset {offsetToCurrentEdit} ---")
    ue_buf = io.BytesIO(ppt_data[offsetToCurrentEdit:])
    recVer, recInstance, recType, recLen = parse_record_header(ue_buf)
    print(f"RecordHeader: ver={recVer:#x}, inst={recInstance:#x}, type={recType:#06x}, len={recLen}")
    print(f"  (期望: type=0x0FF5, len=32 加密 / 28 未加密)")

    if recType == 0x0FF5:
        (lastSlideIdRef,) = unpack("<I", ue_buf.read(4))
        (version,) = unpack("<H", ue_buf.read(2))
        (minorVer, majorVer) = unpack("<BB", ue_buf.read(2))
        (offsetLastEdit,) = unpack("<I", ue_buf.read(4))
        (offsetPersistDirectory,) = unpack("<I", ue_buf.read(4))
        (docPersistIdRef,) = unpack("<I", ue_buf.read(4))
        (maxPersistWritten,) = unpack("<I", ue_buf.read(4))
        (lastView,) = unpack("<H", ue_buf.read(2))
        unused = ue_buf.read(2)
        print(f"lastSlideIdRef: {lastSlideIdRef:#x}")
        print(f"version: {version:#x}, majorVersion: {majorVer}, minorVersion: {minorVer}")
        print(f"offsetLastEdit: {offsetLastEdit:#x}")
        print(f"offsetPersistDirectory: {offsetPersistDirectory} ({offsetPersistDirectory:#x})")
        print(f"docPersistIdRef: {docPersistIdRef}")
        print(f"maxPersistWritten: {maxPersistWritten}")
        print(f"lastView: {lastView:#x}")
        print(f"unused: {unused.hex()}")

        # 读取 encryptSessionPersistIdRef（如果有）
        if recLen == 32:
            (encryptSessionPersistIdRef,) = unpack("<I", ue_buf.read(4))
            print(f"encryptSessionPersistIdRef: {encryptSessionPersistIdRef}")

        # 解析 PersistDirectoryAtom
        print(f"\n--- PersistDirectoryAtom at offset {offsetPersistDirectory} ---")
        pd_buf = io.BytesIO(ppt_data[offsetPersistDirectory:])
        recVer, recInstance, recType, recLen = parse_record_header(pd_buf)
        print(f"RecordHeader: ver={recVer:#x}, inst={recInstance:#x}, type={recType:#06x}, len={recLen}")
        print(f"  (期望: type=0x1772)")

        if recType == 0x1772:
            pd_data = pd_buf.read(recLen)
            pos = 0
            entry_count = 0
            while pos + 4 <= len(pd_data):
                (entry_val,) = unpack("<I", pd_data[pos:pos+4])
                persistId = entry_val & 0xFFFFF
                cPersist = (entry_val >> 20) & 0xFFF
                pos += 4
                print(f"  Entry #{entry_count}: persistId={persistId}, cPersist={cPersist}")
                for j in range(cPersist):
                    if pos + 4 <= len(pd_data):
                        (persist_offset,) = unpack("<I", pd_data[pos:pos+4])
                        print(f"    [{persistId+j}] offset={persist_offset} ({persist_offset:#x})")
                        pos += 4
                entry_count += 1

            # 检查 CryptSession10Container
            if recLen == 32:
                # 加密文件，查找 CryptSession10Container
                # encryptSessionPersistIdRef 指向 persist object directory 中的条目
                # 需要找到 CryptSession10Container 的 offset
                print(f"\n--- 查找 CryptSession10Container ---")

                # 重新解析 persist object directory
                persist_dir = {}
                pos = 0
                while pos + 4 <= len(pd_data):
                    (entry_val,) = unpack("<I", pd_data[pos:pos+4])
                    persistId = entry_val & 0xFFFFF
                    cPersist = (entry_val >> 20) & 0xFFF
                    pos += 4
                    for j in range(cPersist):
                        if pos + 4 <= len(pd_data):
                            (persist_offset,) = unpack("<I", pd_data[pos:pos+4])
                            persist_dir[persistId + j] = persist_offset
                            pos += 4

                if recLen == 32:
                    cs_offset = persist_dir.get(encryptSessionPersistIdRef)
                    if cs_offset:
                        print(f"CryptSession10Container offset: {cs_offset} ({cs_offset:#x})")
                        cs_buf = io.BytesIO(ppt_data[cs_offset:])
                        recVer, recInstance, recType, recLen = parse_record_header(cs_buf)
                        print(f"RecordHeader: ver={recVer:#x}, inst={recInstance:#x}, type={recType:#06x}, len={recLen}")
                        print(f"  (期望: ver=0xF, type=0x2F14)")

                        if recType == 0x2F14:
                            cs_data = cs_buf.read(recLen)
                            cs_info = io.BytesIO(cs_data)

                            # EncryptionVersionInfo
                            (vMajor,) = unpack("<H", cs_info.read(2))
                            (vMinor,) = unpack("<H", cs_info.read(2))
                            print(f"EncryptionVersionInfo: vMajor={vMajor}, vMinor={vMinor}")
                            print(f"  (期望: vMajor=4, vMinor=2)")

                            # flags
                            (flags,) = unpack("<I", cs_info.read(4))
                            print(f"flags: {flags:#010x} (期望 0x0000000C)")

                            # headerSize
                            (headerSize,) = unpack("<I", cs_info.read(4))
                            print(f"headerSize: {headerSize}")

                            # EncryptionHeader
                            (eh_flags,) = unpack("<I", cs_info.read(4))
                            (sizeExtra,) = unpack("<I", cs_info.read(4))
                            (algId,) = unpack("<I", cs_info.read(4))
                            (algIdHash,) = unpack("<I", cs_info.read(4))
                            (keySize,) = unpack("<I", cs_info.read(4))
                            (providerType,) = unpack("<I", cs_info.read(4))
                            (reserved1,) = unpack("<I", cs_info.read(4))
                            (reserved2,) = unpack("<I", cs_info.read(4))
                            cspName = cs_info.read().decode("utf-16le", errors="replace")
                            print(f"EncryptionHeader:")
                            print(f"  flags: {eh_flags:#010x}")
                            print(f"  sizeExtra: {sizeExtra}")
                            print(f"  algId: {algId:#010x} (期望 0x00006801 RC4)")
                            print(f"  algIdHash: {algIdHash:#010x} (期望 0x00008004 SHA1)")
                            print(f"  keySize: {keySize} (期望 128)")
                            print(f"  providerType: {providerType}")
                            print(f"  reserved1: {reserved1}, reserved2: {reserved2}")
                            print(f"  cspName: {cspName!r}")

                            # 注意：EncryptionVerifier 在 EncryptionHeader 之后
                            # 但 cs_info 已经被 cspName 读取完了，需要重新定位
                            # 实际上，cspName 读取了剩余所有字节，所以 EncryptionVerifier 没有被正确解析
                            # 让我重新解析

                            # 重新解析：跳过 EncryptionHeader
                            cs_info2 = io.BytesIO(cs_data)
                            cs_info2.read(4)  # EncryptionVersionInfo
                            cs_info2.read(4)  # flags
                            cs_info2.read(4)  # headerSize
                            cs_info2.read(headerSize)  # EncryptionHeader

                            # EncryptionVerifier
                            (saltSize,) = unpack("<I", cs_info2.read(4))
                            salt = cs_info2.read(16)
                            encryptedVerifier = cs_info2.read(16)
                            (verifierHashSize,) = unpack("<I", cs_info2.read(4))
                            encryptedVerifierHash = cs_info2.read(20)
                            print(f"EncryptionVerifier:")
                            print(f"  saltSize: {saltSize}")
                            print(f"  salt: {salt.hex()}")
                            print(f"  encryptedVerifier: {encryptedVerifier.hex()}")
                            print(f"  verifierHashSize: {verifierHashSize}")
                            print(f"  encryptedVerifierHash: {encryptedVerifierHash.hex()}")

    ole.close()


import os

# 检查 msoffcrypto 测试文件
check_file(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\rc4cryptoapi_password.ppt",
    "msoffcrypto 测试文件（已加密）"
)

# 检查我们生成的加密文件
check_file(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt",
    "我们生成的加密文件"
)

# 检查纯加密文件
check_file(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\protected_心理账户理论.ppt",
    "我们生成的纯加密文件"
)
