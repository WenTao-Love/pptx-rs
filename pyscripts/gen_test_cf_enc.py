"""
用 Python 实现 OOXML Agile Encryption，使用 compoundfiles 库创建 OLE2/CFB 文件。
"""
import struct
import os
import hashlib
import base64
import compoundfiles

PASSWORD = "pptx-rs-secret"

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

    # 使用 compoundfiles 创建 OLE2/CFB 文件
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with compoundfiles.CompoundFile(output_path, "w") as f:
        f.write("EncryptionInfo", enc_info.encode("utf-8"))
        f.write("EncryptedPackage", enc_pkg)

    print(f"Created encrypted file: {output_path}")

# 测试
create_encrypted_pptx(
    "_test/文旅IP人设打造抖音短视频运营方案.pptx",
    "_test_out/py_cf_encrypted.pptx",
    PASSWORD
)
