"""
用 Python 实现 OOXML Agile Encryption（真正的文件加密，打开需要密码）。
手动构造 OLE2/CFB 二进制格式，无需 olefile 写入支持。
参考：MS-CFB, ECMA-376 Part 4
"""
import struct
import os
import hashlib
import base64

PASSWORD = "pptx-rs-secret"

# === OLE2/CFB 常量 ===
OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
DIFSECT = 0xFFFFFFFC
FATSECT = 0xFFFFFFFD
DIRSECT = 0xFFFFFFFE  # same as ENDOFCHAIN for v3

def derive_key(password, salt, spin_count, key_length, hash_algo="SHA512"):
    """OOXML Agile Encryption key derivation."""
    h = hashlib.new(hash_algo)
    h.update(salt)
    h.update(password.encode("utf-16-le"))
    derived = h.digest()
    for i in range(spin_count):
        h = hashlib.new(hash_algo)
        h.update(derived)
        h.update(struct.pack("<I", i))
        derived = h.digest()
    return derived[:key_length]

def aes_cbc_encrypt(key, iv, data):
    """AES-CBC encryption with PKCS7 padding."""
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.primitives import padding
    padder = padding.PKCS7(128).padder()
    padded = padder.update(data) + padder.finalize()
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv))
    enc = cipher.encryptor()
    return enc.update(padded) + enc.finalize()

def build_ole2(enc_info_bytes, enc_pkg_bytes):
    """构建最小 OLE2/CFB 文件，包含 EncryptionInfo 和 EncryptedPackage 两个 stream。"""
    SECTOR_SIZE = 512
    ENTRIES_PER_FAT = SECTOR_SIZE // 4  # 128

    enc_info_sectors = (len(enc_info_bytes) + SECTOR_SIZE - 1) // SECTOR_SIZE
    enc_pkg_sectors = (len(enc_pkg_bytes) + SECTOR_SIZE - 1) // SECTOR_SIZE
    # 我们需要：FAT sectors + 1 Dir sector + data sectors
    data_sectors = enc_info_sectors + enc_pkg_sectors
    # 先假设 1 个 FAT sector，看够不够
    fat_sectors = 1
    total_needed = fat_sectors + 1 + data_sectors  # FAT + Dir + data
    while total_needed > fat_sectors * ENTRIES_PER_FAT:
        fat_sectors += 1
        total_needed = fat_sectors + 1 + data_sectors

    # Sector layout:
    # Sector 0..fat_sectors-1: FAT
    # Sector fat_sectors: Directory
    # Sector fat_sectors+1..fat_sectors+enc_info_sectors: EncryptionInfo
    # Sector fat_sectors+1+enc_info_sectors..: EncryptedPackage
    dir_sector_id = fat_sectors
    enc_info_start = fat_sectors + 1
    enc_pkg_start = enc_info_start + enc_info_sectors

    # === Header (512 bytes) ===
    header = bytearray(SECTOR_SIZE)
    header[0:8] = OLE2_SIGNATURE
    struct.pack_into("<H", header, 24, 0x003E)  # Minor Version
    struct.pack_into("<H", header, 26, 0x0003)  # Major Version (v3)
    struct.pack_into("<H", header, 28, 0xFFFE)  # Byte Order
    struct.pack_into("<H", header, 30, 0x0009)  # Sector Size
    struct.pack_into("<H", header, 32, 0x0006)  # Mini Sector Size
    struct.pack_into("<I", header, 40, 0)        # Total Dir Sectors (v3=0)
    struct.pack_into("<I", header, 44, fat_sectors)  # Total FAT Sectors
    struct.pack_into("<I", header, 48, dir_sector_id)  # First Dir Sector
    struct.pack_into("<I", header, 52, 0)        # Transaction Signature
    struct.pack_into("<I", header, 56, 0x00001000)  # Mini Stream Cutoff
    struct.pack_into("<I", header, 60, ENDOFCHAIN)  # First Mini FAT Sector
    struct.pack_into("<I", header, 64, 0)        # Total Mini FAT Sectors
    struct.pack_into("<I", header, 68, ENDOFCHAIN)  # First DIFAT Sector
    struct.pack_into("<I", header, 72, 0)        # Total DIFAT Sectors
    # DIFAT array (109 entries)
    for i in range(fat_sectors):
        struct.pack_into("<I", header, 76 + i * 4, i)
    for i in range(fat_sectors, 109):
        struct.pack_into("<I", header, 76 + i * 4, FREESECT)

    # === FAT Sectors ===
    total_entries = fat_sectors * ENTRIES_PER_FAT
    fat_entries = [FREESECT] * total_entries

    # FAT sectors themselves
    for i in range(fat_sectors):
        fat_entries[i] = FATSECT

    # Directory sector
    fat_entries[dir_sector_id] = ENDOFCHAIN

    # EncryptionInfo sectors
    for i in range(enc_info_sectors):
        if i < enc_info_sectors - 1:
            fat_entries[enc_info_start + i] = enc_info_start + i + 1
        else:
            fat_entries[enc_info_start + i] = ENDOFCHAIN

    # EncryptedPackage sectors
    for i in range(enc_pkg_sectors):
        if i < enc_pkg_sectors - 1:
            fat_entries[enc_pkg_start + i] = enc_pkg_start + i + 1
        else:
            fat_entries[enc_pkg_start + i] = ENDOFCHAIN

    # 写入 FAT sectors
    fat_data = bytearray()
    for i, entry in enumerate(fat_entries):
        fat_data.extend(struct.pack("<I", entry))
    # 填充到完整的 FAT sectors
    while len(fat_data) < fat_sectors * SECTOR_SIZE:
        fat_data.extend(b"\x00" * 4)

    # === Directory Sector ===
    dir_sector = bytearray(SECTOR_SIZE)

    # Directory entry format (128 bytes each):
    # 0-63: Name (UTF-16LE, null-terminated)
    # 64-65: Name size in bytes (including null)
    # 66: Object type (0=unknown, 1=storage, 2=stream, 5=root)
    # 67: Color flag (0=red, 1=black)
    # 68-71: Left sibling DID
    # 72-75: Right sibling DID
    # 76-79: Child DID
    # 80-95: CLSID
    # 96-99: State bits
    # 100-107: Creation time
    # 108-115: Modification time
    # 116-119: Starting sector SECID
    # 120-127: Size

    def write_dir_entry(buf, offset, name, obj_type, start_sector, size,
                        left_did=0xFFFFFFFF, right_did=0xFFFFFFFF, child_did=0xFFFFFFFF):
        name_utf16 = name.encode("utf-16-le") + b"\x00\x00"
        name_len = len(name_utf16)
        buf[offset:offset+64] = b"\x00" * 64
        buf[offset:offset+min(name_len, 64)] = name_utf16[:min(name_len, 64)]
        struct.pack_into("<H", buf, offset + 64, name_len)
        buf[offset + 66] = obj_type
        buf[offset + 67] = 1  # black
        struct.pack_into("<I", buf, offset + 68, left_did)
        struct.pack_into("<I", buf, offset + 72, right_did)
        struct.pack_into("<I", buf, offset + 76, child_did)
        struct.pack_into("<I", buf, offset + 116, start_sector)
        struct.pack_into("<I", buf, offset + 120, size)

    # 红黑树结构：Root Entry -> child=1(EncryptionInfo)
    # EncryptionInfo: right=2(EncryptedPackage)
    # EncryptedPackage: 无子节点
    write_dir_entry(dir_sector, 0, "Root Entry", 5, ENDOFCHAIN, 0, child_did=1)
    write_dir_entry(dir_sector, 128, "EncryptionInfo", 2, enc_info_start, len(enc_info_bytes), right_did=2)
    write_dir_entry(dir_sector, 256, "EncryptedPackage", 2, enc_pkg_start, len(enc_pkg_bytes))

    # === Data Sectors ===
    def pad_to_sector(data):
        padded = bytearray(data)
        while len(padded) % SECTOR_SIZE != 0:
            padded.append(0)
        return bytes(padded)

    # Build the complete file
    result = bytearray()
    result.extend(header)          # Header (512 bytes)
    result.extend(fat_data)        # FAT sectors
    result.extend(dir_sector)      # Directory sector
    result.extend(pad_to_sector(enc_info_bytes))  # EncryptionInfo data
    result.extend(pad_to_sector(enc_pkg_bytes))    # EncryptedPackage data

    return bytes(result)

def create_encrypted_pptx(input_path, output_path, password):
    """创建加密 PPTX 文件（OOXML Agile Encryption）。"""
    with open(input_path, "rb") as f:
        package_data = f.read()

    # 参数
    salt_size = 16
    key_bits = 256
    key_length = key_bits // 8
    block_size = 16
    hash_algo = "SHA512"
    hash_size = 64
    spin_count = 100000

    # 生成随机 salt
    salt_value = os.urandom(salt_size)
    key_value_salt = os.urandom(salt_size)
    verifier_salt = os.urandom(salt_size)

    # 派生加密密钥
    derived_key = derive_key(password, key_value_salt, spin_count, key_length, hash_algo)

    # 创建验证器
    verifier_hash_input = verifier_salt
    verifier_hash = hashlib.new(hash_algo, verifier_hash_input).digest()

    # 加密验证器
    iv_zero = bytes(block_size)
    encrypted_verifier_hash_input = aes_cbc_encrypt(derived_key, iv_zero, verifier_hash_input)

    # 验证器哈希的密钥
    h = hashlib.new(hash_algo)
    h.update(derived_key)
    h.update(struct.pack("<I", 0x00000000))
    verifier_key = h.digest()[:key_length]
    encrypted_verifier_hash_value = aes_cbc_encrypt(verifier_key, iv_zero, verifier_hash)

    # 包加密密钥
    h = hashlib.new(hash_algo)
    h.update(derived_key)
    h.update(struct.pack("<I", 0x00000001))
    package_key = h.digest()[:key_length]

    # 加密整个包
    encrypted_package = aes_cbc_encrypt(package_key, iv_zero, package_data)

    # 加密包密钥
    encrypted_key_value = aes_cbc_encrypt(derived_key, iv_zero, package_key)

    # 构建 EncryptionInfo XML
    enc_info = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<EncryptionInfo xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
<EncryptionVersionInfo MajorVersion="4" MinorVersion="4"/>
<KeyData SaltSize="{salt_size}" BlockSize="{block_size}" KeyBits="{key_bits}" HashSize="{hash_size}" CipherAlgorithm="AES" CipherChaining="ChainingModeCBC" HashAlgorithm="{hash_algo}" SaltValue="{base64.b64encode(salt_value).decode()}"/>
<KeyEncryptors>
<KeyEncryptor Uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
<p:KeyEncryptor SaltSize="{salt_size}" BlockSize="{block_size}" KeyBits="{key_bits}" HashSize="{hash_size}" CipherAlgorithm="AES" CipherChaining="ChainingModeCBC" HashAlgorithm="{hash_algo}" SpinCount="{spin_count}" SaltValue="{base64.b64encode(key_value_salt).decode()}" EncryptedVerifierHashInput="{base64.b64encode(encrypted_verifier_hash_input).decode()}" EncryptedVerifierHashValue="{base64.b64encode(encrypted_verifier_hash_value).decode()}" EncryptedKeyValue="{base64.b64encode(encrypted_key_value).decode()}"/>
</KeyEncryptor>
</KeyEncryptors>
</EncryptionInfo>"""

    # 构建 EncryptedPackage stream
    enc_pkg = struct.pack("<I", len(package_data)) + encrypted_package

    # 构建 OLE2/CFB 文件
    ole2_data = build_ole2(enc_info.encode("utf-8"), enc_pkg)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(ole2_data)

    print(f"Created encrypted file: {output_path} ({len(ole2_data)} bytes)")

# 测试
create_encrypted_pptx(
    "_test/文旅IP人设打造抖音短视频运营方案.pptx",
    "_test_out/py_agile_encrypted.pptx",
    PASSWORD
)
