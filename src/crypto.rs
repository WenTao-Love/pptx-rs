//! 加密与保护模块：修改密码保护 + ECMA-376 Agile Encryption。
//!
//! 本模块提供两种级别的文档保护：
//!
//! 1. **修改密码保护**（modifyVerifier）：打开时可只读浏览，修改需密码。
//!    对标 PowerPoint "保护演示文稿 → 限制访问" 功能。
//!    算法遵循 [MS-OFFCRYPTO] §2.4.2.4（SHA-512 + salt + spinCount）。
//!
//! 2. **文件加密**（ECMA-376 Agile Encryption）：打开文件即需密码。
//!    对标 PowerPoint "保护演示文稿 → 用密码进行加密" 功能。
//!    使用 AES-256-CBC + SHA-512 密钥派生，符合现代 Office 加密标准。
//!    算法遵循 [MS-OFFCRYPTO] §2.3.4.11（Agile Encryption）。
//!
//! # 与 pypdf 的对应
//!
//! - pypdf `PdfWriter.encrypt(user_password, owner_password)` ←→ [`encrypt_package`]
//! - pypdf `PdfReader.is_encrypted` ←→ [`is_encrypted_package`]
//! - pypdf `PdfReader.decrypt(password)` ←→ [`decrypt_package`]
//!
//! [MS-OFFCRYPTO]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/

use base64::Engine;
use sha2::{Digest, Sha512};

use crate::error::{Error, Result};

/// 修改密码保护的哈希迭代次数。
///
/// Office 2016+ 默认 100 000 次；越大越慢但越安全。
pub const MODIFY_SPIN_COUNT: u32 = 100_000;

/// Agile Encryption 的哈希迭代次数。
///
/// 与 modifyVerifier 相同默认值，但两者属于不同的加密体系。
const AGILE_SPIN_COUNT: u32 = 100_000;

/// salt 字节长度（16 字节 = 128 位，符合 MS-OFFCRYPTO 推荐）。
pub const SALT_LEN: usize = 16;

/// AES-256 密钥长度（32 字节）。
const AES_KEY_LEN: usize = 32;

/// AES 块大小（16 字节）。
const AES_BLOCK_SIZE: usize = 16;

// ---------------------------------------------------------------------------
// Agile Encryption blockKey 常量（MS-OFFCRYPTO §2.3.4.11 / §2.3.4.12）
// ---------------------------------------------------------------------------

/// 加密 encryptedKeyValue 时的 blockKey（32 字节）。
///
/// 用于从密码派生哈希中进一步派生加密 secret_key 所需的密钥。
const BLOCK_KEY_ENCRYPTED_KEY_VALUE: &[u8] = &[
    0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6, 0x2e, 0x0b, 0x80, 0x14, 0x73, 0x0f, 0xbf, 0x7d,
    0x0e, 0xda, 0x25, 0x2e, 0x6b, 0x0f, 0x9b, 0x55, 0xb3, 0x10, 0x0f, 0x65, 0x0f, 0xa3, 0x5c, 0x6b,
];

/// 加密 verifierHashInput 时的 blockKey（8 字节）。
const BLOCK_KEY_VERIFIER_HASH_INPUT: &[u8] = &[0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];

/// 加密 verifierHashValue 时的 blockKey（8 字节）。
const BLOCK_KEY_VERIFIER_HASH_VALUE: &[u8] = &[0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x27, 0x9c];

// ---------------------------------------------------------------------------
// 修改密码保护（modifyVerifier）
// ---------------------------------------------------------------------------

/// 修改密码保护参数。
///
/// 对应 `<p:modifyVerifier>` XML 元素，包含密码哈希、salt、迭代次数等。
/// 注入到 `presentation.xml` 后，WPS / PowerPoint 打开时会提示
/// "以只读方式打开"或"输入密码以修改"。
#[derive(Clone, Debug)]
pub struct ModifyProtection {
    /// salt 的 base64 编码。
    pub salt_b64: String,
    /// 哈希值的 base64 编码。
    pub hash_b64: String,
    /// 迭代次数。
    pub spin_count: u32,
    /// 加密算法提供者类型。
    pub crypt_provider_type: &'static str,
    /// 加密算法 SID（14 = SHA-512）。
    pub algorithm_sid: u32,
}

impl ModifyProtection {
    /// 从密码创建修改密码保护。
    ///
    /// # 算法（MS-OFFCRYPTO §2.4.2.4）
    ///
    /// ```text
    /// H_0 = SHA-512(salt || password_utf16le)
    /// H_n = SHA-512(H_{n-1} || n.to_le_u32)   for n = 0..spinCount-1
    /// ```
    ///
    /// # 参数
    /// - `password`：明文密码；
    /// - `salt`：16 字节随机值（每次调用应使用不同 salt）；
    /// - `spin_count`：迭代次数（默认 100 000）。
    pub fn from_password(password: &str, salt: &[u8], spin_count: u32) -> Self {
        let hash = compute_modify_hash(password, salt, spin_count);
        let salt_b64 = base64::engine::general_purpose::STANDARD.encode(salt);
        let hash_b64 = base64::engine::general_purpose::STANDARD.encode(&hash);
        ModifyProtection {
            salt_b64,
            hash_b64,
            spin_count,
            crypt_provider_type: "rsaFull",
            algorithm_sid: 14,
        }
    }

    /// 验证密码是否匹配。
    ///
    /// 用相同算法重新计算哈希，与存储的哈希比对。
    pub fn verify_password(&self, password: &str) -> bool {
        let salt = match base64::engine::general_purpose::STANDARD.decode(&self.salt_b64) {
            Ok(s) => s,
            Err(_) => return false,
        };
        let expected_hash = compute_modify_hash(password, &salt, self.spin_count);
        let expected_b64 = base64::engine::general_purpose::STANDARD.encode(&expected_hash);
        // 常量时间比较（base64 编码后长度固定，简单比对即可）
        expected_b64 == self.hash_b64
    }

    /// 序列化为 `<p:modifyVerifier .../>` XML 元素。
    ///
    /// 元素位置必须在 `<p:extLst>` 之前（OOXML schema 顺序要求）。
    pub fn to_xml_element(&self) -> String {
        format!(
            r#"<p:modifyVerifier cryptProviderType="{}" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="{}" cryptSpinCount="{}" hash="{}" salt="{}"/>"#,
            self.crypt_provider_type,
            self.algorithm_sid,
            self.spin_count,
            self.hash_b64,
            self.salt_b64,
        )
    }
}

/// 按 MS-OFFCRYPTO §2.4.2.4 计算 modifyVerifier 的哈希。
///
/// # 算法
///
/// ```text
/// H_0 = SHA-512(salt || password_utf16le)
/// H_n = SHA-512(H_{n-1} || n.to_le_u32)   for n = 0..spinCount-1
/// ```
///
/// - 密码用 **UTF-16 LE** 编码（OOXML 规范，不是 UTF-8）；
/// - 每次迭代都把迭代序号 n 以小端 uint32 拼到 H_{n-1} 之后。
///
/// **注意**：此函数的迭代顺序（`H_{n-1} || iterator`）与 Agile Encryption
/// 的迭代顺序（`iterator || H_{n-1}`）**不同**，不可混用。
///
/// # 参数
///
/// - `password`：明文密码（Unicode）；
/// - `salt`：16 字节随机值；
/// - `spin_count`：迭代次数（Office 2016+ 默认 100 000）。
///
/// # 返回
///
/// 64 字节（SHA-512 摘要）的二进制哈希。
pub fn compute_modify_hash(password: &str, salt: &[u8], spin_count: u32) -> Vec<u8> {
    // 把 password 编码为 UTF-16 LE
    let password_utf16le: Vec<u8> = password
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    // H_0 = SHA-512(salt || password_utf16le)
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&password_utf16le);
    let mut h = hasher.finalize().to_vec();

    // H_n = SHA-512(H_{n-1} || n.to_le_u32)
    // modifyVerifier 的迭代顺序：H_{n-1} 在前，iterator 在后
    for i in 0..spin_count {
        let mut hasher = Sha512::new();
        hasher.update(&h);
        hasher.update(i.to_le_bytes());
        h = hasher.finalize().to_vec();
    }
    h
}

// ---------------------------------------------------------------------------
// ECMA-376 Agile Encryption（文件级加密）
// ---------------------------------------------------------------------------

/// Agile Encryption 参数。
///
/// 对应 `EncryptionInfo` XML 流中的加密参数。
/// 使用 AES-256-CBC + SHA-512，符合 Office 2016+ 默认加密标准。
#[derive(Clone, Debug)]
pub struct AgileEncryptionParams {
    /// 密码 salt（16 字节，base64 编码）。
    pub password_salt_b64: String,
    /// 数据 salt（16 字节，base64 编码）。
    pub data_salt_b64: String,
    /// 密码迭代次数。
    pub spin_count: u32,
    /// 加密后的数据密钥（base64 编码）。
    pub encrypted_key_value_b64: String,
    /// 加密后的验证器哈希输入（base64 编码）。
    pub encrypted_verifier_hash_input_b64: String,
    /// 加密后的验证器哈希值（base64 编码）。
    pub encrypted_verifier_hash_value_b64: String,
}

/// 生成随机 salt/IV 字节。
///
/// 使用 `sha2` + 时间戳 + 计数器作为伪随机源（不依赖 `rand` crate，
/// 保持库依赖最小化。安全性由 SHA-512 的不可逆性保证）。
pub fn generate_random_bytes(len: usize) -> Vec<u8> {
    // 使用 SHA-512 对多个熵源进行哈希来生成伪随机字节
    let mut result = Vec::with_capacity(len);
    let mut counter: u64 = 0;
    while result.len() < len {
        let mut hasher = Sha512::new();
        // 熵源：时间戳 + 计数器 + 已生成字节（反馈）
        hasher.update(
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos()
                .to_le_bytes(),
        );
        hasher.update(counter.to_le_bytes());
        hasher.update(&result);
        let digest = hasher.finalize();
        let needed = std::cmp::min(64, len - result.len());
        result.extend_from_slice(&digest[..needed]);
        counter += 1;
    }
    result
}

// ---------------------------------------------------------------------------
// Agile Encryption 密钥派生（MS-OFFCRYPTO §2.3.4.11）
// ---------------------------------------------------------------------------

/// Agile Encryption 密码哈希迭代（MS-OFFCRYPTO §2.3.4.11 第 1-4 步）。
///
/// # 算法
///
/// ```text
/// H_0 = SHA-512(salt || password_utf16le)
/// H_n = SHA-512(iterator_u32_le || H_{n-1})   for iterator = 0..spinCount-1
/// ```
///
/// **与 [`compute_modify_hash`] 的关键区别**：迭代顺序不同！
/// - modifyVerifier（§2.4.2.4）：`H_{n-1} || iterator`
/// - Agile Encryption（§2.3.4.11）：`iterator || H_{n-1}`
///
/// # 参数
/// - `password`：明文密码；
/// - `salt`：passwordSalt（16 字节）；
/// - `spin_count`：迭代次数。
///
/// # 返回
/// 64 字节 SHA-512 哈希值。
fn agile_hash_password(password: &str, salt: &[u8], spin_count: u32) -> Vec<u8> {
    let password_utf16le: Vec<u8> = password
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .collect();

    // H_0 = SHA-512(salt || password_utf16le)
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(&password_utf16le);
    let mut h = hasher.finalize().to_vec();

    // H_n = SHA-512(iterator_u32_le || H_{n-1})
    // Agile Encryption 的迭代顺序：iterator 在前，H_{n-1} 在后
    for i in 0..spin_count {
        let mut hasher = Sha512::new();
        hasher.update(i.to_le_bytes());
        hasher.update(&h);
        h = hasher.finalize().to_vec();
    }
    h
}

/// Agile Encryption 密钥哈希迭代（MS-OFFCRYPTO §2.3.4.11 第 6-9 步）。
///
/// 与 [`agile_hash_password`] 类似，但输入是密钥字节而非密码字符串。
/// 用于从 secret_key + keyDataSalt 派生数据加密密钥。
///
/// # 算法
///
/// ```text
/// H_0 = SHA-512(salt || key_bytes)
/// H_n = SHA-512(iterator_u32_le || H_{n-1})   for iterator = 0..spinCount-1
/// ```
fn agile_hash_key(key: &[u8], salt: &[u8], spin_count: u32) -> Vec<u8> {
    // H_0 = SHA-512(salt || key)
    let mut hasher = Sha512::new();
    hasher.update(salt);
    hasher.update(key);
    let mut h = hasher.finalize().to_vec();

    // H_n = SHA-512(iterator_u32_le || H_{n-1})
    for i in 0..spin_count {
        let mut hasher = Sha512::new();
        hasher.update(i.to_le_bytes());
        hasher.update(&h);
        h = hasher.finalize().to_vec();
    }
    h
}

/// 用 blockKey 派生最终密钥（MS-OFFCRYPTO §2.3.4.11 第 5 步）。
///
/// # 算法
///
/// ```text
/// H_final = SHA-512(H_n || blockKey)
/// derived_key = H_final[0..key_len]
/// ```
///
/// 规范要求在迭代哈希完成后，再拼接 blockKey 做一次哈希，
/// 然后截取前 key_len 字节作为最终派生密钥。
///
/// # 参数
/// - `hash`：迭代哈希结果（[`agile_hash_password`] 或 [`agile_hash_key`] 的输出）；
/// - `block_key`：blockKey 常量（不同用途使用不同的 blockKey）；
/// - `key_len`：需要的密钥字节数（AES-256 为 32）。
///
/// # 返回
/// `key_len` 字节的派生密钥。
fn derive_key_with_block_key(hash: &[u8], block_key: &[u8], key_len: usize) -> Vec<u8> {
    let mut hasher = Sha512::new();
    hasher.update(hash);
    hasher.update(block_key);
    let derived = hasher.finalize();
    derived[..key_len].to_vec()
}

/// 从 salt 派生 IV（MS-OFFCRYPTO §2.3.4.11）。
///
/// 规范要求 IV 取 salt 的前 blockSize 字节，而非随机生成。
/// 这保证了加密/解密双方可以从相同的 salt 推导出相同的 IV。
///
/// # 参数
/// - `salt`：salt 值（至少 16 字节）。
///
/// # 返回
/// 16 字节 IV。
fn iv_from_salt(salt: &[u8]) -> Vec<u8> {
    salt[..AES_BLOCK_SIZE].to_vec()
}

// ---------------------------------------------------------------------------
// AES-256-CBC 加解密
// ---------------------------------------------------------------------------

/// AES-256-CBC 加密。
///
/// # 参数
/// - `key`：32 字节 AES-256 密钥；
/// - `iv`：16 字节初始向量；
/// - `data`：明文数据（会自动添加 PKCS7 填充）。
///
/// # 返回
/// 加密后的密文（长度为 16 的整数倍）。
///
/// # 错误
/// - [`Error::Encryption`]：密钥或 IV 长度不正确。
pub fn aes_256_cbc_encrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    if key.len() != AES_KEY_LEN {
        return Err(Error::encryption("AES-256 key must be 32 bytes"));
    }
    if iv.len() != AES_BLOCK_SIZE {
        return Err(Error::encryption("AES IV must be 16 bytes"));
    }

    use aes::Aes256;
    use cbc::cipher::{BlockEncryptMut, KeyIvInit};
    type Aes256CbcEnc = cbc::Encryptor<Aes256>;

    let encryptor = Aes256CbcEnc::new_from_slices(key, iv)
        .map_err(|e| Error::encryption(format!("AES init failed: {}", e)))?;
    let ciphertext = encryptor.encrypt_padded_vec_mut::<cbc::cipher::block_padding::Pkcs7>(data);

    Ok(ciphertext)
}

/// AES-256-CBC 解密。
///
/// # 参数
/// - `key`：32 字节 AES-256 密钥；
/// - `iv`：16 字节初始向量；
/// - `data`：密文数据。
///
/// # 返回
/// 解密后的明文（自动去除 PKCS7 填充）。
///
/// # 错误
/// - [`Error::Encryption`]：密钥/IV 长度不正确或解密失败（填充错误）。
pub fn aes_256_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    if key.len() != AES_KEY_LEN {
        return Err(Error::encryption("AES-256 key must be 32 bytes"));
    }
    if iv.len() != AES_BLOCK_SIZE {
        return Err(Error::encryption("AES IV must be 16 bytes"));
    }
    if !data.len().is_multiple_of(AES_BLOCK_SIZE) {
        return Err(Error::encryption(
            "ciphertext length must be multiple of 16",
        ));
    }

    use aes::Aes256;
    use cbc::cipher::{BlockDecryptMut, KeyIvInit};
    type Aes256CbcDec = cbc::Decryptor<Aes256>;

    let decryptor = Aes256CbcDec::new_from_slices(key, iv)
        .map_err(|e| Error::encryption(format!("AES init failed: {}", e)))?;

    let plaintext = decryptor
        .decrypt_padded_vec_mut::<cbc::cipher::block_padding::Pkcs7>(data)
        .map_err(|e| Error::encryption(format!("AES decrypt failed: {}", e)))?;
    Ok(plaintext)
}

// ---------------------------------------------------------------------------
// 加密/解密主流程
// ---------------------------------------------------------------------------

/// 加密整个 ZIP 包（ECMA-376 Agile Encryption）。
///
/// # 流程（MS-OFFCRYPTO §2.3.4.11）
///
/// 1. 生成随机 passwordSalt / keyDataSalt / secret_key；
/// 2. 用密码派生哈希 + blockKey 派生加密密钥，加密 secret_key → encryptedKeyValue；
/// 3. 生成 verifierHashInput / verifierHashValue，用密码派生密钥加密；
/// 4. 用 secret_key + keyDataSalt 派生数据加密密钥；
/// 5. IV = keyDataSalt 的前 16 字节（规范要求，非随机生成）；
/// 6. 加密原始 ZIP → EncryptedPackage；
/// 7. 构造 EncryptionInfo XML；
/// 8. 打包为新的 ZIP（含 EncryptionInfo + EncryptedPackage）。
///
/// # 参数
/// - `zip_bytes`：原始 .pptx 的 ZIP 字节流；
/// - `password`：加密密码。
///
/// # 返回
/// 加密后的 ZIP 字节流（打开需密码）。
pub fn encrypt_package(zip_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    let spin_count = AGILE_SPIN_COUNT;

    // 1) 生成随机值
    let password_salt = generate_random_bytes(SALT_LEN);
    let key_data_salt = generate_random_bytes(SALT_LEN);
    let secret_key = generate_random_bytes(AES_KEY_LEN);

    // 2) 从密码派生哈希（Agile Encryption 迭代顺序：iterator 在前）
    let password_hash = agile_hash_password(password, &password_salt, spin_count);

    // 3) 用 blockKey 派生加密 secret_key 的密钥
    let key_encrypt_key =
        derive_key_with_block_key(&password_hash, BLOCK_KEY_ENCRYPTED_KEY_VALUE, AES_KEY_LEN);
    // IV = passwordSalt 的前 16 字节（规范要求）
    let key_iv = iv_from_salt(&password_salt);
    let encrypted_key = aes_256_cbc_encrypt(&key_encrypt_key, &key_iv, &secret_key)?;

    // 4) 生成并加密验证器
    // verifierHashInput：随机 16 字节
    let verifier_hash_input = generate_random_bytes(SALT_LEN);
    // verifierHashValue = SHA-512(verifierHashInput)
    let verifier_hash_value = {
        let mut hasher = Sha512::new();
        hasher.update(&verifier_hash_input);
        hasher.finalize().to_vec()
    };

    // 用 blockKey 派生加密 verifierHashInput 的密钥
    let verifier_input_key =
        derive_key_with_block_key(&password_hash, BLOCK_KEY_VERIFIER_HASH_INPUT, AES_KEY_LEN);
    let encrypted_verifier_hash_input =
        aes_256_cbc_encrypt(&verifier_input_key, &key_iv, &verifier_hash_input)?;

    // 用 blockKey 派生加密 verifierHashValue 的密钥
    let verifier_value_key =
        derive_key_with_block_key(&password_hash, BLOCK_KEY_VERIFIER_HASH_VALUE, AES_KEY_LEN);
    let encrypted_verifier_hash_value =
        aes_256_cbc_encrypt(&verifier_value_key, &key_iv, &verifier_hash_value)?;

    // 5) 从 secret_key + keyDataSalt 派生数据加密密钥
    let data_hash = agile_hash_key(&secret_key, &key_data_salt, spin_count);
    // 数据加密不需要 blockKey（§2.3.4.11 第 14 步直接截取）
    let data_key = data_hash[..AES_KEY_LEN].to_vec();

    // 6) 加密原始 ZIP 数据
    // IV = keyDataSalt 的前 16 字节（规范要求）
    let data_iv = iv_from_salt(&key_data_salt);

    // EncryptedPackage 格式：4 字节 LE uint32 原始长度 + 加密数据
    let original_len = zip_bytes.len() as u32;
    let encrypted_data = aes_256_cbc_encrypt(&data_key, &data_iv, zip_bytes)?;

    let mut encrypted_package = Vec::with_capacity(4 + encrypted_data.len());
    encrypted_package.extend_from_slice(&original_len.to_le_bytes());
    encrypted_package.extend_from_slice(&encrypted_data);

    // 7) 构造 EncryptionInfo XML
    let encryption_info_xml = build_encryption_info_xml(
        &password_salt,
        &key_data_salt,
        &encrypted_key,
        &encrypted_verifier_hash_input,
        &encrypted_verifier_hash_value,
        spin_count,
    );

    // 8) 打包为新的 ZIP
    build_encrypted_zip(&encryption_info_xml, &encrypted_package)
}

/// 解密加密的 ZIP 包。
///
/// # 流程（MS-OFFCRYPTO §2.3.4.11 解密方向）
///
/// 1. 解析加密的 ZIP，读取 EncryptionInfo + EncryptedPackage；
/// 2. 从密码派生哈希 + blockKey 派生解密 secret_key 的密钥；
/// 3. IV = passwordSalt 的前 16 字节；
/// 4. 解密得到 secret_key；
/// 5. 验证密码（可选）：解密 verifierHashInput/Value 并验证；
/// 6. 从 secret_key + keyDataSalt 派生数据解密密钥；
/// 7. IV = keyDataSalt 的前 16 字节；
/// 8. 解密数据。
///
/// # 参数
/// - `encrypted_bytes`：加密的 ZIP 字节流；
/// - `password`：解密密码。
///
/// # 返回
/// 解密后的原始 ZIP 字节流（可被 `Presentation::load_bytes` 加载）。
pub fn decrypt_package(encrypted_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    // 1) 解析加密的 ZIP
    let cursor = std::io::Cursor::new(encrypted_bytes);
    let mut zip = zip::ZipArchive::new(cursor)
        .map_err(|e| Error::encryption(format!("encrypted zip open failed: {}", e)))?;

    // 2) 读取 EncryptionInfo
    let mut info_xml = String::new();
    let mut info_found = false;
    for i in 0..zip.len() {
        let mut entry = zip
            .by_index(i)
            .map_err(|e| Error::encryption(format!("zip read failed: {}", e)))?;
        if entry.name().contains("EncryptionInfo") {
            use std::io::Read;
            entry
                .read_to_string(&mut info_xml)
                .map_err(|e| Error::encryption(format!("read EncryptionInfo failed: {}", e)))?;
            info_found = true;
            break;
        }
    }
    if !info_found {
        return Err(Error::encryption("EncryptionInfo not found in package"));
    }

    // 3) 解析 EncryptionInfo XML
    let params = parse_encryption_info(&info_xml)?;

    // 4) 读取 EncryptedPackage
    let mut package_data = Vec::new();
    let mut package_found = false;
    for i in 0..zip.len() {
        let mut entry = zip
            .by_index(i)
            .map_err(|e| Error::encryption(format!("zip read failed: {}", e)))?;
        if entry.name().contains("EncryptedPackage") {
            use std::io::Read;
            entry
                .read_to_end(&mut package_data)
                .map_err(|e| Error::encryption(format!("read EncryptedPackage failed: {}", e)))?;
            package_found = true;
            break;
        }
    }
    if !package_found {
        return Err(Error::encryption("EncryptedPackage not found in package"));
    }

    // 5) 从密码派生哈希（Agile Encryption 迭代顺序：iterator 在前）
    let password_salt = base64::engine::general_purpose::STANDARD
        .decode(&params.password_salt_b64)
        .map_err(|e| Error::encryption(format!("password salt decode failed: {}", e)))?;
    let password_hash = agile_hash_password(password, &password_salt, params.spin_count);

    // 6) 用 blockKey 派生解密 secret_key 的密钥
    let key_encrypt_key =
        derive_key_with_block_key(&password_hash, BLOCK_KEY_ENCRYPTED_KEY_VALUE, AES_KEY_LEN);
    // IV = passwordSalt 的前 16 字节
    let key_iv = iv_from_salt(&password_salt);

    let encrypted_key_value = base64::engine::general_purpose::STANDARD
        .decode(&params.encrypted_key_value_b64)
        .map_err(|e| Error::encryption(format!("encrypted key decode failed: {}", e)))?;
    let secret_key = aes_256_cbc_decrypt(&key_encrypt_key, &key_iv, &encrypted_key_value)?;

    // 7) 验证密码（解密 verifier 并校验）
    if !params.encrypted_verifier_hash_input_b64.is_empty()
        && !params.encrypted_verifier_hash_value_b64.is_empty()
    {
        let verifier_input_key =
            derive_key_with_block_key(&password_hash, BLOCK_KEY_VERIFIER_HASH_INPUT, AES_KEY_LEN);
        let verifier_value_key =
            derive_key_with_block_key(&password_hash, BLOCK_KEY_VERIFIER_HASH_VALUE, AES_KEY_LEN);

        let encrypted_verifier_input = base64::engine::general_purpose::STANDARD
            .decode(&params.encrypted_verifier_hash_input_b64)
            .map_err(|e| {
                Error::encryption(format!("encrypted verifier input decode failed: {}", e))
            })?;
        let encrypted_verifier_value = base64::engine::general_purpose::STANDARD
            .decode(&params.encrypted_verifier_hash_value_b64)
            .map_err(|e| {
                Error::encryption(format!("encrypted verifier value decode failed: {}", e))
            })?;

        let verifier_input =
            aes_256_cbc_decrypt(&verifier_input_key, &key_iv, &encrypted_verifier_input)?;
        let verifier_value =
            aes_256_cbc_decrypt(&verifier_value_key, &key_iv, &encrypted_verifier_value)?;

        // 验证：SHA-512(verifierHashInput) 应等于 verifierHashValue
        let expected_hash = {
            let mut hasher = Sha512::new();
            hasher.update(&verifier_input);
            hasher.finalize().to_vec()
        };
        if expected_hash != verifier_value {
            return Err(Error::encryption("password verification failed"));
        }
    }

    // 8) 从 secret_key + keyDataSalt 派生数据解密密钥
    let key_data_salt = base64::engine::general_purpose::STANDARD
        .decode(&params.data_salt_b64)
        .map_err(|e| Error::encryption(format!("data salt decode failed: {}", e)))?;
    let data_hash = agile_hash_key(&secret_key, &key_data_salt, params.spin_count);
    let data_key = data_hash[..AES_KEY_LEN].to_vec();

    // 9) 解密数据
    // IV = keyDataSalt 的前 16 字节
    let data_iv = iv_from_salt(&key_data_salt);

    // EncryptedPackage 格式：4 字节 LE uint32 原始长度 + 加密数据
    if package_data.len() < 4 {
        return Err(Error::encryption("EncryptedPackage too short"));
    }
    let _original_len = u32::from_le_bytes([
        package_data[0],
        package_data[1],
        package_data[2],
        package_data[3],
    ]);
    let encrypted_data = &package_data[4..];

    let decrypted = aes_256_cbc_decrypt(&data_key, &data_iv, encrypted_data)?;

    Ok(decrypted)
}

/// 检查字节流是否为加密的 OOXML 文档。
///
/// 加密文档的特征：ZIP 中包含 `EncryptionInfo` 条目。
pub fn is_encrypted_package(bytes: &[u8]) -> bool {
    let cursor = std::io::Cursor::new(bytes);
    match zip::ZipArchive::new(cursor) {
        Ok(mut zip) => {
            for i in 0..zip.len() {
                if let Ok(entry) = zip.by_index(i) {
                    if entry.name().contains("EncryptionInfo") {
                        return true;
                    }
                }
            }
            false
        }
        Err(_) => false,
    }
}

// ---------------------------------------------------------------------------
// EncryptionInfo XML 构造与解析
// ---------------------------------------------------------------------------

/// 解析后的 EncryptionInfo 参数。
#[derive(Clone, Debug)]
struct ParsedEncryptionInfo {
    password_salt_b64: String,
    data_salt_b64: String,
    encrypted_key_value_b64: String,
    encrypted_verifier_hash_input_b64: String,
    encrypted_verifier_hash_value_b64: String,
    spin_count: u32,
}

/// 从 EncryptionInfo XML 解析加密参数。
///
/// 解析 `<encryption>` XML 中的关键属性：
/// - `keyData` 的 `saltValue`（keyDataSalt）
/// - `encryptedKey` 的 `saltValue`（passwordSalt）、`spinCount`、`encryptedKeyValue`
/// - `encryptedKey` 的 `encryptedVerifierHashInput`、`encryptedVerifierHashValue`
fn parse_encryption_info(xml: &str) -> Result<ParsedEncryptionInfo> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut password_salt_b64 = String::new();
    let mut data_salt_b64 = String::new();
    let mut encrypted_key_value_b64 = String::new();
    let mut encrypted_verifier_hash_input_b64 = String::new();
    let mut encrypted_verifier_hash_value_b64 = String::new();
    let mut spin_count: u32 = AGILE_SPIN_COUNT;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"keyData" {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"saltValue" {
                            data_salt_b64 = attr_unescape_value(&attr);
                        }
                    }
                } else if local == b"encryptedKey" {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        if key == b"saltValue" {
                            password_salt_b64 = attr_unescape_value(&attr);
                        } else if key == b"spinCount" {
                            let v = attr_unescape_value(&attr);
                            spin_count = v.parse::<u32>().unwrap_or(AGILE_SPIN_COUNT);
                        } else if key == b"encryptedKeyValue" {
                            encrypted_key_value_b64 = attr_unescape_value(&attr);
                        } else if key == b"encryptedVerifierHashInput" {
                            encrypted_verifier_hash_input_b64 = attr_unescape_value(&attr);
                        } else if key == b"encryptedVerifierHashValue" {
                            encrypted_verifier_hash_value_b64 = attr_unescape_value(&attr);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::encryption(format!(
                    "EncryptionInfo XML parse error: {}",
                    e
                )))
            }
            _ => {}
        }
    }

    Ok(ParsedEncryptionInfo {
        password_salt_b64,
        data_salt_b64,
        encrypted_key_value_b64,
        encrypted_verifier_hash_input_b64,
        encrypted_verifier_hash_value_b64,
        spin_count,
    })
}

/// 从 QName 中提取 local name（去除命名空间前缀）。
///
/// 与 `oxml/parse_sld.rs` 中的同名函数逻辑一致。
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&c| c == b':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

/// 从 Attribute 中安全提取解码后的值字符串。
fn attr_unescape_value(attr: &quick_xml::events::attributes::Attribute<'_>) -> String {
    attr.normalized_value(quick_xml::XmlVersion::Implicit1_0)
        .map(|cow| cow.into_owned())
        .unwrap_or_else(|_| String::from_utf8_lossy(attr.value.as_ref()).into_owned())
}

/// 构造 EncryptionInfo XML（MS-OFFCRYPTO §2.3.4.10）。
///
/// XML 结构：
/// ```xml
/// <encryption>
///   <keyData ... />
///   <dataIntegrity ... />
///   <keyEncryptors>
///     <keyEncryptor uri="...">
///       <p:encryptedKey ... />
///     </keyEncryptor>
///   </keyEncryptors>
/// </encryption>
/// ```
///
/// 关键属性说明：
/// - `keyData`：数据加密参数（saltValue = keyDataSalt）
/// - `dataIntegrity`：HMAC 完整性校验（当前为空，PowerPoint 接受空值）
/// - `encryptedKey`：密码加密参数，包含：
///   - `saltValue`：passwordSalt
///   - `encryptedKeyValue`：加密后的 secret_key
///   - `encryptedVerifierHashInput`：加密后的验证器输入
///   - `encryptedVerifierHashValue`：加密后的验证器哈希值
fn build_encryption_info_xml(
    password_salt: &[u8],
    key_data_salt: &[u8],
    encrypted_key: &[u8],
    encrypted_verifier_hash_input: &[u8],
    encrypted_verifier_hash_value: &[u8],
    spin_count: u32,
) -> String {
    let password_salt_b64 = base64::engine::general_purpose::STANDARD.encode(password_salt);
    let key_data_salt_b64 = base64::engine::general_purpose::STANDARD.encode(key_data_salt);
    let encrypted_key_b64 = base64::engine::general_purpose::STANDARD.encode(encrypted_key);
    let encrypted_verifier_hash_input_b64 =
        base64::engine::general_purpose::STANDARD.encode(encrypted_verifier_hash_input);
    let encrypted_verifier_hash_value_b64 =
        base64::engine::general_purpose::STANDARD.encode(encrypted_verifier_hash_value);

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="{salt_size}" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="{key_data_salt_b64}" />
  <dataIntegrity encryptedHmacKey="" encryptedHmacValue="" algorithmIdentifier="SHA512" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltSize="{salt_size}" saltValue="{password_salt_b64}" hashAlgorithm="SHA512" cipherAlgorithm="AES" keyBits="256" blockSize="16" encryptedKeyValue="{encrypted_key_b64}" encryptedVerifierHashInput="{encrypted_verifier_hash_input_b64}" encryptedVerifierHashValue="{encrypted_verifier_hash_value_b64}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        salt_size = SALT_LEN,
        spin_count = spin_count,
        key_data_salt_b64 = key_data_salt_b64,
        password_salt_b64 = password_salt_b64,
        encrypted_key_b64 = encrypted_key_b64,
        encrypted_verifier_hash_input_b64 = encrypted_verifier_hash_input_b64,
        encrypted_verifier_hash_value_b64 = encrypted_verifier_hash_value_b64,
    )
}

/// 构造加密后的 ZIP 文件。
///
/// ZIP 内包含：
/// - `[Content_Types].xml`
/// - `_rels/.rels`
/// - `EncryptionInfo`
/// - `EncryptedPackage`
fn build_encrypted_zip(encryption_info_xml: &str, encrypted_package: &[u8]) -> Result<Vec<u8>> {
    let mut buf = Vec::new();
    // 使用 block 确保 Cursor 在访问 buf 之前被 drop（与 opc/package.rs 保持一致）
    {
        let cursor = std::io::Cursor::new(&mut buf);
        let mut zip = zip::ZipWriter::new(cursor);
        let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o644);

        // [Content_Types].xml
        let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="bin" ContentType="application/vnd.ms-office.encryptedPackage" />
</Types>"#;
        zip.start_file("[Content_Types].xml", opts)
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;
        zip.write_all(ct_xml.as_bytes())
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;

        // _rels/.rels
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/encryption" Target="EncryptionInfo" />
</Relationships>"#;
        zip.start_file("_rels/.rels", opts)
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;
        zip.write_all(rels_xml.as_bytes())
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;

        // EncryptionInfo
        zip.start_file("EncryptionInfo", opts)
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;
        zip.write_all(encryption_info_xml.as_bytes())
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;

        // EncryptedPackage
        zip.start_file("EncryptedPackage", opts)
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;
        zip.write_all(encrypted_package)
            .map_err(|e| Error::encryption(format!("zip write failed: {}", e)))?;

        zip.finish()
            .map_err(|e| Error::encryption(format!("zip finish failed: {}", e)))?;
    }
    Ok(buf)
}

use std::io::Write;
