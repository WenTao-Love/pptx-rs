//! 为 .pptx 设置真正的文件加密（OOXML Agile Encryption）。
//!
//! 完全对齐 msoffcrypto-python 的加密逻辑，包括：
//! - 子密钥派生使用 blockKey（而非 LE32 selector）
//! - EncryptionInfo 字段使用 password_salt 作为 IV
//! - EncryptedPackage 分段加密（4096字节/段），每段不同 IV
//! - EncryptedPackage 大小头为 u64（8字节）
//!
//! 用法：在 pptx-rs 目录执行
//!   cargo run --example protect_pptx
//! 密码：pptx-rs-secret

use std::io::{Read, Write};

use rand::Rng;
use sha2::{Digest, Sha512};

const PASSWORD: &str = "pptx-rs-secret";

// ============================================================================
// Block Keys（来自 MS-OFFCRYPTO 规范 / msoffcrypto 源码）
// ============================================================================

/// verifierHashInput 的 block key
const BLKKEY_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
/// encryptedVerifierHashValue 的 block key
const BLKKEY_ENCRYPTED_VERIFIER_HASH_VALUE: [u8; 8] =
    [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
/// encryptedKeyValue 的 block key
const BLKKEY_ENCRYPTED_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
/// dataIntegrity HMAC key 的 block key
const BLKKEY_DATA_INTEGRITY_1: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
/// dataIntegrity HMAC value 的 block key
const BLKKEY_DATA_INTEGRITY_2: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

/// EncryptedPackage 分段大小
const SEGMENT_LENGTH: usize = 4096;

// ============================================================================
// DataSpaces 流的原始字节（从 msoffcrypto-python 参考文件提取）
// ============================================================================

/// DataSpaces/Version stream（76 字节）
const DS_VERSION: &[u8] = &[
    0x3c, 0x00, 0x00, 0x00, 0x4d, 0x00, 0x69, 0x00, 0x63, 0x00, 0x72, 0x00, 0x6f, 0x00, 0x73, 0x00,
    0x6f, 0x00, 0x66, 0x00, 0x74, 0x00, 0x2e, 0x00, 0x43, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x74, 0x00,
    0x61, 0x00, 0x69, 0x00, 0x6e, 0x00, 0x65, 0x00, 0x72, 0x00, 0x2e, 0x00, 0x44, 0x00, 0x61, 0x00,
    0x74, 0x00, 0x61, 0x00, 0x53, 0x00, 0x70, 0x00, 0x61, 0x00, 0x63, 0x00, 0x65, 0x00, 0x73, 0x00,
    0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
];

/// DataSpaces/DataSpaceMap stream（112 字节）
const DS_DATASPACEMAP: &[u8] = &[
    0x08, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x68, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x45, 0x00, 0x6e, 0x00, 0x63, 0x00, 0x72, 0x00,
    0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x65, 0x00, 0x64, 0x00, 0x50, 0x00, 0x61, 0x00, 0x63, 0x00,
    0x6b, 0x00, 0x61, 0x00, 0x67, 0x00, 0x65, 0x00, 0x32, 0x00, 0x00, 0x00, 0x53, 0x00, 0x74, 0x00,
    0x72, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x67, 0x00, 0x45, 0x00, 0x6e, 0x00, 0x63, 0x00, 0x72, 0x00,
    0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x69, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x44, 0x00, 0x61, 0x00,
    0x74, 0x00, 0x61, 0x00, 0x53, 0x00, 0x70, 0x00, 0x61, 0x00, 0x63, 0x00, 0x65, 0x00, 0x00, 0x00,
];

/// DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace stream（64 字节）
const DS_STRONGENCRYPTIONDATASPACE: &[u8] = &[
    0x08, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x32, 0x00, 0x00, 0x00, 0x53, 0x00, 0x74, 0x00,
    0x72, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x67, 0x00, 0x45, 0x00, 0x6e, 0x00, 0x63, 0x00, 0x72, 0x00,
    0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x69, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x54, 0x00, 0x72, 0x00,
    0x61, 0x00, 0x6e, 0x00, 0x73, 0x00, 0x66, 0x00, 0x6f, 0x00, 0x72, 0x00, 0x6d, 0x00, 0x00, 0x00,
];

/// DataSpaces/TransformInfo/StrongEncryptionTransform/Primary stream（200 字节）
const DS_PRIMARY: &[u8] = &[
    0x58, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x4c, 0x00, 0x00, 0x00, 0x7b, 0x00, 0x46, 0x00,
    0x46, 0x00, 0x39, 0x00, 0x41, 0x00, 0x33, 0x00, 0x46, 0x00, 0x30, 0x00, 0x33, 0x00, 0x2d, 0x00,
    0x35, 0x00, 0x36, 0x00, 0x45, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x34, 0x00, 0x36, 0x00, 0x31, 0x00,
    0x33, 0x00, 0x2d, 0x00, 0x42, 0x00, 0x44, 0x00, 0x44, 0x00, 0x35, 0x00, 0x2d, 0x00, 0x35, 0x00,
    0x41, 0x00, 0x34, 0x00, 0x31, 0x00, 0x43, 0x00, 0x31, 0x00, 0x44, 0x00, 0x30, 0x00, 0x37, 0x00,
    0x32, 0x00, 0x34, 0x00, 0x36, 0x00, 0x7d, 0x00, 0x4e, 0x00, 0x00, 0x00, 0x4d, 0x00, 0x69, 0x00,
    0x63, 0x00, 0x72, 0x00, 0x6f, 0x00, 0x73, 0x00, 0x6f, 0x00, 0x66, 0x00, 0x74, 0x00, 0x2e, 0x00,
    0x43, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x74, 0x00, 0x61, 0x00, 0x69, 0x00, 0x6e, 0x00, 0x65, 0x00,
    0x72, 0x00, 0x2e, 0x00, 0x45, 0x00, 0x6e, 0x00, 0x63, 0x00, 0x72, 0x00, 0x79, 0x00, 0x70, 0x00,
    0x74, 0x00, 0x69, 0x00, 0x6f, 0x00, 0x6e, 0x00, 0x54, 0x00, 0x72, 0x00, 0x61, 0x00, 0x6e, 0x00,
    0x73, 0x00, 0x66, 0x00, 0x6f, 0x00, 0x72, 0x00, 0x6d, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
    0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00,
];

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let entries = std::fs::read_dir("_test")?;
    let mut processed = 0;

    for entry in entries {
        let entry = entry?;
        let p = entry.path();
        if p.extension().and_then(|e| e.to_str()) != Some("pptx") {
            continue;
        }
        let fname = p.file_name().unwrap().to_string_lossy().to_string();

        match encrypt_pptx(&p, PASSWORD) {
            Ok(encrypted_data) => {
                std::fs::create_dir_all("_test_out")?;
                let out_path = format!("_test_out/protected_{}", fname);
                std::fs::write(&out_path, &encrypted_data)?;
                println!("已加密：{}", out_path);
                processed += 1;
            }
            Err(e) => {
                eprintln!("跳过 {}: {}", fname, e);
            }
        }
    }
    println!("共处理 {} 个 pptx 文件", processed);
    Ok(())
}

/// 使用 OOXML Agile Encryption 加密 PPTX 文件。
///
/// 完全对齐 msoffcrypto-python 的加密逻辑：
/// - 密码派生：H₀ = SHA512(salt + password_utf16le), Hₙ = SHA512(LE32(n) + Hₙ₋₁)
/// - 子密钥派生：key = SHA512(h + blockKey)[:keyBits/8]
/// - EncryptionInfo 字段 IV = password_salt（补齐到块大小）
/// - EncryptedPackage 分段加密，每段 IV = SHA512(keyDataSalt + LE32(segmentIndex))[:16]
fn encrypt_pptx(
    input_path: &std::path::Path,
    password: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let package_data = std::fs::read(input_path)?;

    let salt_size: usize = 16;
    let block_size: usize = 16;
    let key_bits: usize = 256;
    let key_length = key_bits / 8;
    let hash_size: usize = 64;
    let spin_count: u32 = 100000;

    // 生成随机 salt
    let mut rng = rand::thread_rng();
    let key_data_salt: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();
    let password_salt: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();

    // 步骤1：从密码派生基础哈希
    // H₀ = SHA512(salt + password_utf16le)
    // Hₙ = SHA512(LE32(n) + Hₙ₋₁)  注意：迭代器在哈希前面！
    let h = derive_iterated_hash(password, &password_salt, spin_count);

    // 步骤2：用 blockKey 派生各加密密钥
    let key1 = derive_encryption_key(&h, &BLKKEY_VERIFIER_HASH_INPUT, key_length);
    let key2 = derive_encryption_key(&h, &BLKKEY_ENCRYPTED_VERIFIER_HASH_VALUE, key_length);
    let key3 = derive_encryption_key(&h, &BLKKEY_ENCRYPTED_KEY_VALUE, key_length);

    // 步骤3：生成验证器
    // verifierHashInput = 随机16字节，补齐到块大小
    let verifier_hash_input_raw: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();
    let verifier_hash_input =
        resize_buffer(&verifier_hash_input_raw, round_up(salt_size, block_size));

    // IV = password_salt（补齐到块大小，用0x36填充）
    let iv_salt = normalize_key(&password_salt, block_size);

    // encryptedVerifierHashInput = AES-CBC(verifierHashInput, key1, iv_salt)
    let encrypted_verifier_hash_input = aes_cbc_encrypt_raw(&key1, &iv_salt, &verifier_hash_input)?;

    // verifierHash = SHA512(verifierHashInput)，补齐到块大小
    let verifier_hash = Sha512::digest(&verifier_hash_input);
    let verifier_hash_padded = resize_buffer(&verifier_hash, round_up(hash_size, block_size));

    // encryptedVerifierHashValue = AES-CBC(verifierHash, key2, iv_salt)
    let encrypted_verifier_hash_value =
        aes_cbc_encrypt_raw(&key2, &iv_salt, &verifier_hash_padded)?;

    // 步骤4：生成 secret_key（用于加密包和HMAC）
    let secret_key_raw: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();
    let secret_key = normalize_key(&secret_key_raw, key_length);

    // encryptedKeyValue = AES-CBC(secretKey, key3, iv_salt)
    let encrypted_key_value = aes_cbc_encrypt_raw(&key3, &iv_salt, &secret_key)?;

    // 步骤5：加密 EncryptedPackage（分段加密）
    let encrypted_package = encrypt_payload(&package_data, &secret_key, &key_data_salt)?;

    // 步骤6：生成 dataIntegrity（HMAC）
    let (encrypted_hmac_key, encrypted_hmac_value) = generate_integrity_parameter(
        &encrypted_package,
        &secret_key,
        &key_data_salt,
        hash_size,
        block_size,
    )?;

    // 步骤7：构建 EncryptionInfo XML
    let enc_info_xml = build_encryption_info_xml(
        salt_size as u32,
        block_size as u32,
        key_bits as u32,
        hash_size as u32,
        spin_count,
        &key_data_salt,
        &password_salt,
        &encrypted_verifier_hash_input,
        &encrypted_verifier_hash_value,
        &encrypted_key_value,
        &encrypted_hmac_key,
        &encrypted_hmac_value,
    );

    // EncryptionInfo stream 格式：MajorVersion(u16) + MinorVersion(u16) + Reserved(u32=0x40) + XML
    let mut enc_info_stream = Vec::new();
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes());
    enc_info_stream.extend_from_slice(&4u16.to_le_bytes());
    enc_info_stream.extend_from_slice(&0x40u32.to_le_bytes());
    enc_info_stream.extend_from_slice(enc_info_xml.as_bytes());

    // 写入 OLE2/CFB 文件
    let mut buf = std::io::Cursor::new(Vec::new());
    {
        let mut comp = cfb::CompoundFile::create(&mut buf)?;

        comp.create_storage("\u{0006}DataSpaces")?;
        write_stream(&mut comp, "\u{0006}DataSpaces/Version", DS_VERSION)?;
        write_stream(
            &mut comp,
            "\u{0006}DataSpaces/DataSpaceMap",
            DS_DATASPACEMAP,
        )?;
        comp.create_storage("\u{0006}DataSpaces/DataSpaceInfo")?;
        write_stream(
            &mut comp,
            "\u{0006}DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace",
            DS_STRONGENCRYPTIONDATASPACE,
        )?;
        comp.create_storage("\u{0006}DataSpaces/TransformInfo")?;
        comp.create_storage("\u{0006}DataSpaces/TransformInfo/StrongEncryptionTransform")?;
        write_stream(
            &mut comp,
            "\u{0006}DataSpaces/TransformInfo/StrongEncryptionTransform/\u{0006}Primary",
            DS_PRIMARY,
        )?;

        write_stream(&mut comp, "EncryptionInfo", &enc_info_stream)?;
        write_stream(&mut comp, "EncryptedPackage", &encrypted_package)?;

        comp.flush()?;
    }

    Ok(buf.into_inner())
}

/// 写入一个 stream 到 CompoundFile。
fn write_stream<F: Read + Write + std::io::Seek>(
    comp: &mut cfb::CompoundFile<F>,
    path: &str,
    data: &[u8],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut stream = comp.create_stream(path)?;
    stream.write_all(data)?;
    stream.flush()?;
    drop(stream);
    Ok(())
}

/// 从密码派生迭代哈希。
///
/// H₀ = SHA512(salt + password_utf16le)
/// Hₙ = SHA512(LE32(n) + Hₙ₋₁)   注意：迭代器在哈希前面（与 msoffcrypto 一致）
fn derive_iterated_hash(password: &str, salt: &[u8], spin_count: u32) -> Vec<u8> {
    let password_utf16le: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&password_utf16le);
    let mut h = hasher.finalize().to_vec();
    for i in 0u32..spin_count {
        let mut hasher2 = Sha512::new();
        hasher2.update(i.to_le_bytes()); // 迭代器在前
        hasher2.update(&h);
        h = hasher2.finalize().to_vec();
    }
    h
}

/// 用 blockKey 派生加密密钥。
///
/// key = SHA512(h + blockKey)[:key_length]
fn derive_encryption_key(h: &[u8], block_key: &[u8], key_length: usize) -> Vec<u8> {
    let mut hasher = Sha512::new();
    hasher.update(h);
    hasher.update(block_key);
    let derived = hasher.finalize();
    derived[..key_length].to_vec()
}

/// 将缓冲区补齐到指定大小（用 0x36 填充，与 msoffcrypto 的 normalize_key 一致）。
fn normalize_key(key: &[u8], n: usize) -> Vec<u8> {
    if key.len() >= n {
        key[..n].to_vec()
    } else {
        let mut result = key.to_vec();
        result.resize(n, 0x36);
        result
    }
}

/// 将缓冲区补齐到指定大小（用 0x00 填充）。
fn resize_buffer(buf: &[u8], n: usize) -> Vec<u8> {
    if buf.len() >= n {
        buf[..n].to_vec()
    } else {
        let mut result = buf.to_vec();
        result.resize(n, 0);
        result
    }
}

/// 向上取整到块大小的整数倍。
fn round_up(sz: usize, block: usize) -> usize {
    sz.div_ceil(block) * block
}

/// AES-256-CBC 加密（无填充，输入必须是块大小的整数倍）。
fn aes_cbc_encrypt_raw(
    key: &[u8],
    iv: &[u8],
    data: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    use aes::cipher::{BlockEncryptMut, KeyIvInit};
    type Aes256CbcEnc = cbc::Encryptor<aes::Aes256>;
    if data.len() % 16 != 0 {
        return Err(format!("数据长度 {} 不是块大小 16 的整数倍", data.len()).into());
    }
    let mut encryptor = Aes256CbcEnc::new_from_slices(key, iv)?;
    let mut buf = data.to_vec();
    for chunk in buf.chunks_exact_mut(16) {
        let mut block = [0u8; 16];
        block.copy_from_slice(chunk);
        let mut block = aes::cipher::Block::<aes::Aes256>::from(block);
        encryptor.encrypt_block_mut(&mut block);
        chunk.copy_from_slice(&block);
    }
    Ok(buf)
}

/// 分段加密 EncryptedPackage。
///
/// 格式：u64 原始大小 + 分段加密数据
/// 每段 4096 字节，IV = SHA512(keyDataSalt + LE32(segmentIndex))[:16]
/// 最后一段补齐到块大小
fn encrypt_payload(
    package_data: &[u8],
    secret_key: &[u8],
    key_data_salt: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut result = Vec::new();
    // u64 原始大小（先写0，最后更新）
    result.extend_from_slice(&(package_data.len() as u64).to_le_bytes());

    let mut offset = 0;
    let mut segment_index: u32 = 0;

    while offset < package_data.len() {
        let end = std::cmp::min(offset + SEGMENT_LENGTH, package_data.len());
        let chunk = &package_data[offset..end];

        // IV = SHA512(keyDataSalt + LE32(segmentIndex))[:16]
        let mut iv_hasher = Sha512::new();
        iv_hasher.update(key_data_salt);
        iv_hasher.update(segment_index.to_le_bytes());
        let iv_full = iv_hasher.finalize();
        let iv = &iv_full[..16];

        // 补齐到块大小
        let padded_chunk = if chunk.len() % 16 != 0 {
            resize_buffer(chunk, round_up(chunk.len(), 16))
        } else {
            chunk.to_vec()
        };

        let encrypted = aes_cbc_encrypt_raw(secret_key, iv, &padded_chunk)?;
        result.extend_from_slice(&encrypted);

        offset = end;
        segment_index += 1;
    }

    Ok(result)
}

/// 生成 dataIntegrity 的 encryptedHmacKey 和 encryptedHmacValue。
///
/// HMAC key = 随机 hashSize 字节
/// IV1 = SHA512(keyDataSalt + blkKey_dataIntegrity1)[:16]
/// IV2 = SHA512(keyDataSalt + blkKey_dataIntegrity2)[:16]
/// encryptedHmacKey = AES-CBC(hmacKey, secretKey, IV1)
/// encryptedHmacValue = AES-CBC(HMAC-SHA512(encryptedPackage), secretKey, IV2)
fn generate_integrity_parameter(
    encrypted_package: &[u8],
    secret_key: &[u8],
    key_data_salt: &[u8],
    hash_size: usize,
    block_size: usize,
) -> Result<(Vec<u8>, Vec<u8>), Box<dyn std::error::Error>> {
    let mut rng = rand::thread_rng();
    let hmac_key: Vec<u8> = (0..hash_size).map(|_| rng.gen()).collect();

    // IV1 = SHA512(keyDataSalt + blkKey_dataIntegrity1)[:16]
    let mut iv1_hasher = Sha512::new();
    iv1_hasher.update(key_data_salt);
    iv1_hasher.update(BLKKEY_DATA_INTEGRITY_1);
    let iv1 = normalize_key(&iv1_hasher.finalize(), block_size);

    // IV2 = SHA512(keyDataSalt + blkKey_dataIntegrity2)[:16]
    let mut iv2_hasher = Sha512::new();
    iv2_hasher.update(key_data_salt);
    iv2_hasher.update(BLKKEY_DATA_INTEGRITY_2);
    let iv2 = normalize_key(&iv2_hasher.finalize(), block_size);

    // encryptedHmacKey = AES-CBC(hmacKey, secretKey, IV1)
    let hmac_key_padded = resize_buffer(&hmac_key, round_up(hash_size, block_size));
    let encrypted_hmac_key = aes_cbc_encrypt_raw(secret_key, &iv1, &hmac_key_padded)?;

    // HMAC-SHA512(encryptedPackage)
    use hmac::{Hmac, Mac};
    type HmacSha512 = Hmac<Sha512>;
    let mut mac = HmacSha512::new_from_slice(&hmac_key)?;
    mac.update(encrypted_package);
    let hmac_value = mac.finalize().into_bytes();
    let hmac_value_padded = resize_buffer(&hmac_value, round_up(hash_size, block_size));

    // encryptedHmacValue = AES-CBC(hmacValue, secretKey, IV2)
    let encrypted_hmac_value = aes_cbc_encrypt_raw(secret_key, &iv2, &hmac_value_padded)?;

    Ok((encrypted_hmac_key, encrypted_hmac_value))
}

/// 构建 EncryptionInfo XML。
#[allow(clippy::too_many_arguments)]
fn build_encryption_info_xml(
    salt_size: u32,
    block_size: u32,
    key_bits: u32,
    hash_size: u32,
    spin_count: u32,
    key_data_salt: &[u8],
    password_salt: &[u8],
    enc_vhi: &[u8],
    enc_vhv: &[u8],
    enc_kv: &[u8],
    enc_hmac_key: &[u8],
    enc_hmac_value: &[u8],
) -> String {
    use base64::Engine;
    let b64 = |d: &[u8]| base64::engine::general_purpose::STANDARD.encode(d);
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
    <keyData saltSize="{salt_size}" blockSize="{block_size}" keyBits="{key_bits}" hashSize="{hash_size}"
             cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{kds}" />
    <dataIntegrity encryptedHmacKey="{ehk}" encryptedHmacValue="{ehv}" />
    <keyEncryptors>
        <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
            <p:encryptedKey spinCount="{spin_count}" saltSize="{salt_size}" blockSize="{block_size}" keyBits="{key_bits}"
                            hashSize="{hash_size}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"
                            saltValue="{ps}" encryptedVerifierHashInput="{evhi}"
                            encryptedVerifierHashValue="{evhv}" encryptedKeyValue="{ekv}" />
        </keyEncryptor>
    </keyEncryptors>
</encryption>"#,
        salt_size = salt_size,
        block_size = block_size,
        key_bits = key_bits,
        hash_size = hash_size,
        spin_count = spin_count,
        kds = b64(key_data_salt),
        ps = b64(password_salt),
        evhi = b64(enc_vhi),
        evhv = b64(enc_vhv),
        ekv = b64(enc_kv),
        ehk = b64(enc_hmac_key),
        ehv = b64(enc_hmac_value),
    )
}
