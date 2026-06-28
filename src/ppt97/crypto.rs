//! .ppt 文件 RC4 CryptoAPI 加密。
//!
//! 本模块为 PowerPoint 97-2003 二进制格式（`.ppt`）文件实现 RC4 CryptoAPI
//! 加密（与 `.pptx` 的 AES Agile Encryption 完全不同）。加密后的文件可被
//! msoffcrypto-python / PowerPoint / WPS 正确解密打开。
//!
//! # 加密机制概述
//!
//! - **密钥派生**：`H₀ = SHA1(salt + password_utf16le)`，
//!   `Hfinal = SHA1(H₀ + LE32(block))`
//! - 40-bit key：`hfinal[:5] + 11 个 0x00`；其他：`hfinal[:keyBits/8]`
//! - 每个 persist 对象独立加密，`block = persistId`，
//!   `blocksize = keyBits * (totalLen / keyBits + 1)`（与 msoffcrypto 一致）
//! - UserEditAtom 和 PersistDirectoryAtom 不加密
//! - `CryptSession10Container` (0x2F14) 存储加密参数，追加到 stream 中
//! - 加密标记：UserEditAtom.recLen 从 28 → 32（添加 encryptSessionPersistIdRef）
//! - 加密标记：PersistDirectoryAtom 添加 CryptSession10Container 条目
//! - 加密标记：CurrentUserAtom.headerToken 从 0xE391C05F → 0xF3D1C4DF
//!
//! # 规范依据
//!
//! - [MS-OFFCRYPTO] 2.3.5：RC4 CryptoAPI Encryption
//! - [MS-PPT] 2.3.2：CurrentUserAtom / headerToken
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx 不支持 .ppt 二进制格式加密。本模块对标 msoffcrypto-python 的
//! `ppt97.py` 与 Apache POI 的 `HSLFSlideShowEncrypted`。
//!
//! [MS-OFFCRYPTO]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto
//! [MS-PPT]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt

use rand::Rng;
use sha1::{Digest, Sha1};

use crate::error::{Error, Result};
use crate::ppt97::record::{
    parse_persist_directory, parse_record_header, read_u32_le, write_u32_le, RT_CURRENT_USER_ATOM,
    RT_PERSIST_DIRECTORY_ATOM, RT_USER_EDIT_ATOM,
};

// ============================================================================
// 加密常量
// ============================================================================

/// RC4 密钥位数（128 bit）。
///
/// 选择 128 位的原因：现代 Office 默认值，且与 `CSP_NAME` 指定的
/// "Microsoft Enhanced Cryptographic Provider v1.0" 兼容（基础 CSP 仅支持 56 位）。
pub const KEY_SIZE_BITS: u32 = 128;

/// salt 字节数（MS-OFFCRYPTO 规范：16 字节）。
pub const SALT_SIZE: usize = 16;

/// verifier 字节数（MS-OFFCRYPTO 规范：16 字节随机值）。
pub const VERIFIER_SIZE: usize = 16;

/// SHA1 哈希字节数（20 字节）。
pub const SHA1_SIZE: usize = 20;

/// CryptSession10Container record type（MS-PPT 规范：0x2F14）。
pub const RT_CRYPT_SESSION10_CONTAINER: u16 = 0x2F14;

/// 已加密标记（CurrentUserAtom.headerToken）。
///
/// MS-PPT 2.3.2 规范：
/// - 未加密：`0xE391C05F`
/// - 已加密：`0xF3D1C4DF`
///
/// PowerPoint 严格检查此值，错误会导致文件无法打开。
/// 之前错误使用 `0xF3D1C4D0`，参考 msoffcrypto 测试文件
/// `rc4cryptoapi_password.ppt` 确认正确值为 `0xF3D1C4DF`。
pub const HEADER_TOKEN_ENCRYPTED: u32 = 0xF3D1C4DF;

/// CSP 名称（UTF-16LE 编码，以 null 结尾）。
///
/// 选择 "Microsoft Enhanced Cryptographic Provider v1.0" 的原因：
/// 基础 CSP（"Microsoft Base Cryptographic Provider v1.0"）仅支持最多 56 位密钥，
/// 与 `KEY_SIZE_BITS=128` 不兼容。PowerPoint 严格检查 CSP 与 KeySize 的兼容性，
/// 不兼容会拒绝打开。msoffcrypto 忽略 CSPName 所以能验证通过，但 PowerPoint 会失败。
pub const CSP_NAME: &str = "Microsoft Enhanced Cryptographic Provider v1.0\0";

// ============================================================================
// RC4 流密码实现
// ============================================================================

/// RC4 流密码状态。
///
/// RC4 是对称流密码，加密和解密使用相同的操作。
/// 支持连续 `process` 调用，保持内部状态（用于 verifier / verifierHash 连续加密）。
#[derive(Debug)]
struct Rc4 {
    /// 状态数组 S（KSA 后的置换）。
    s: [u8; 256],
    /// 索引 i。
    i: usize,
    /// 索引 j。
    j: usize,
}

impl Rc4 {
    /// 创建 RC4 实例，执行 KSA（Key Scheduling Algorithm）。
    ///
    /// # 参数
    /// - `key`：RC4 密钥字节
    #[allow(clippy::needless_range_loop)]
    fn new(key: &[u8]) -> Self {
        let mut s = [0u8; 256];
        for i in 0..256 {
            s[i] = i as u8;
        }
        let mut j = 0;
        for i in 0..256 {
            j = (j + s[i] as usize + key[i % key.len()] as usize) % 256;
            s.swap(i, j);
        }
        Self { s, i: 0, j: 0 }
    }

    /// 处理数据（PRGA），原地修改。
    ///
    /// # 参数
    /// - `data`：待加密/解密的数据（原地异或 keystream）
    fn process(&mut self, data: &mut [u8]) {
        for byte in data {
            self.i = (self.i + 1) % 256;
            self.j = (self.j + self.s[self.i] as usize) % 256;
            self.s.swap(self.i, self.j);
            let k = self.s[(self.s[self.i] as usize + self.s[self.j] as usize) % 256];
            *byte ^= k;
        }
    }
}

// ============================================================================
// RC4 CryptoAPI 密钥派生
// ============================================================================

/// RC4 CryptoAPI 密钥派生。
///
/// # 算法（MS-OFFCRYPTO 规范）
///
/// ```text
/// H₀ = SHA1(salt + password_utf16le)
/// Hfinal = SHA1(H₀ + LE32(block))
/// 40-bit key: hfinal[:5] + 11 个 0x00
/// 其他:       hfinal[:keyBits/8]
/// ```
///
/// # 参数
/// - `password`：明文密码
/// - `salt`：salt 字节（16 字节）
/// - `key_bits`：密钥位数（如 128）
/// - `block`：block 序号（用于分段加密时不同 block 派生不同密钥）
///
/// # 返回
/// 派生的 RC4 密钥字节序列。
pub fn make_key(password: &str, salt: &[u8], key_bits: u32, block: u32) -> Vec<u8> {
    let password_utf16le: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // H₀ = SHA1(salt + password_utf16le)
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&password_utf16le);
    let h0 = hasher.finalize();

    // Hfinal = SHA1(H₀ + LE32(block))
    let mut hasher = Sha1::new();
    hasher.update(h0);
    hasher.update(block.to_le_bytes());
    let hfinal = hasher.finalize();

    let key_bytes = key_bits as usize / 8;
    if key_bits == 40 {
        // 40-bit key: hfinal[:5] + 11 个 0x00（补齐到 16 字节）
        let mut key = hfinal[..5].to_vec();
        key.resize(16, 0);
        key
    } else {
        hfinal[..key_bytes].to_vec()
    }
}

/// 加密一个 persist 对象。
///
/// 按 `blocksize` 分段，每段用不同 `block` 的 key。block 从 `persist_id` 开始递增。
///
/// # blocksize 计算
///
/// 与 msoffcrypto 一致（虽然 MS-OFFCRYPTO 未文档化）：
/// `blocksize = keyBits * (totalLen / keyBits + 1)`
///
/// 注意：`keyBits` 是位数（128），但 `blocksize` 是字节数。
/// 这意味着 `blocksize` 是 128 字节的倍数，且大于 `totalLen`。
///
/// # 参数
/// - `password`：明文密码
/// - `salt`：salt 字节
/// - `key_bits`：密钥位数
/// - `data`：待加密的 persist 对象数据（header + payload）
/// - `persist_id`：persist 对象的 ID（作为初始 block 编号）
///
/// # 返回
/// 加密后的字节序列（长度与输入相同）。
pub fn encrypt_persist_object(
    password: &str,
    salt: &[u8],
    key_bits: u32,
    data: &mut [u8],
    persist_id: u32,
) -> Vec<u8> {
    let total_len = data.len();
    // 与 msoffcrypto 一致：用 keyBits（128）作为 blocksize 基数，不是 keySizeBytes（16）
    let blocksize = key_bits as usize * (total_len / key_bits as usize + 1);

    let mut result = Vec::with_capacity(total_len);
    let mut offset = 0;
    let mut block = persist_id;

    while offset < total_len {
        let end = std::cmp::min(offset + blocksize, total_len);
        let key = make_key(password, salt, key_bits, block);
        let mut rc4 = Rc4::new(&key);
        let mut chunk = data[offset..end].to_vec();
        rc4.process(&mut chunk);
        result.extend_from_slice(&chunk);
        offset = end;
        block += 1;
    }

    result
}

// ============================================================================
// CryptSession10Container 构造
// ============================================================================

/// 构造 CryptSession10Container。
///
/// # 结构（对齐 msoffcrypto 的 `_parse_header_RC4CryptoAPI`）
///
/// ```text
/// RecordHeader: ver=0xF, inst=0, type=0x2F14
/// EncryptionVersionInfo: vMajor=4, vMinor=2
/// flags: 0x0000000C (fCryptoAPI=1 + fDocProps=1)
/// headerSize: EncryptionHeader 字节数
/// EncryptionHeader:
///     flags (4 bytes): 0x0000000C
///     sizeExtra (4 bytes): 0
///     algId (4 bytes): 0x00006801 (RC4)
///     algIdHash (4 bytes): 0x00008004 (SHA1)
///     keySize (4 bytes): key_bits
///     providerType (4 bytes): 0x00000001
///     reserved1 (4 bytes): 0
///     reserved2 (4 bytes): 0
///     cspName (UTF-16LE): CSP_NAME
/// EncryptionVerifier:
///     saltSize (4 bytes): 16
///     salt (16 bytes)
///     encryptedVerifier (16 bytes)
///     verifierHashSize (4 bytes): 20
///     encryptedVerifierHash (20 bytes)
/// ```
///
/// # 关键参数说明
///
/// - `vMajor=4`：参考 msoffcrypto 测试文件 `rc4cryptoapi_password.ppt` 确认正确值。
///   MS-OFFCRYPTO 2.3.5.1 规范：`vMajor` MUST be 0x0002, 0x0003, or 0x0004。
///   之前错误使用 `vMajor=2`，PowerPoint 严格检查此值，错误会导致文件无法打开。
/// - `flags = 0x0000000C`：`fCryptoAPI(1) + fDocProps(1)`，fDocProps=1 避免需要
///   处理 Summary Info Stream 的加密/移除。
///
/// # 参数
/// - `salt`：salt 字节（16 字节）
/// - `encrypted_verifier`：加密后的 verifier（16 字节）
/// - `encrypted_verifier_hash`：加密后的 verifierHash（20 字节）
/// - `key_bits`：密钥位数
///
/// # 返回
/// 完整的 CryptSession10Container 字节序列（含 RecordHeader）。
pub fn build_crypt_session10_container(
    salt: &[u8],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
    key_bits: u32,
) -> Vec<u8> {
    // EncryptionVersionInfo: vMajor=4, vMinor=2
    let mut version_info = Vec::new();
    version_info.extend_from_slice(&4u16.to_le_bytes());
    version_info.extend_from_slice(&2u16.to_le_bytes());

    // 外层 flags（MS-OFFCRYPTO 2.3.1 EncryptionHeaderFlags 位域定义）：
    //   bit 0-1: Reserved1/2，必须为 0
    //   bit 2 (0x04): fCryptoAPI — RC4 CryptoAPI 加密必须为 1
    //   bit 3 (0x08): fDocProps — 1=文档属性不加密（保留原始 Summary Info Stream）
    //   bit 4 (0x10): fExternal — 必须为 0
    //   bit 5 (0x20): fAES — RC4 时必须为 0
    // 正确值：0x0000000C = fCryptoAPI(1) + fDocProps(1)
    let outer_flags: u32 = 0x0000000C;

    // EncryptionHeader
    let mut header = Vec::new();
    header.extend_from_slice(&outer_flags.to_le_bytes()); // flags
    header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    header.extend_from_slice(&0x00006801u32.to_le_bytes()); // algId (RC4)
    header.extend_from_slice(&0x00008004u32.to_le_bytes()); // algIdHash (SHA1)
    header.extend_from_slice(&key_bits.to_le_bytes()); // keySize
    header.extend_from_slice(&0x00000001u32.to_le_bytes()); // providerType
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved2
                                                   // cspName (UTF-16LE)
    let csp_name_utf16: Vec<u8> = CSP_NAME
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();
    header.extend_from_slice(&csp_name_utf16);

    // EncryptionVerifier
    let mut verifier = Vec::new();
    verifier.extend_from_slice(&(SALT_SIZE as u32).to_le_bytes()); // saltSize
    verifier.extend_from_slice(salt); // salt
    verifier.extend_from_slice(encrypted_verifier); // encryptedVerifier
    verifier.extend_from_slice(&(SHA1_SIZE as u32).to_le_bytes()); // verifierHashSize
    verifier.extend_from_slice(encrypted_verifier_hash); // encryptedVerifierHash

    // 构建 CryptSession10Container 的 data 部分
    let mut data = Vec::new();
    data.extend_from_slice(&version_info);
    data.extend_from_slice(&outer_flags.to_le_bytes()); // 外层 flags
    data.extend_from_slice(&(header.len() as u32).to_le_bytes()); // headerSize
    data.extend_from_slice(&header);
    data.extend_from_slice(&verifier);

    // RecordHeader: ver=0xF, inst=0, type=0x2F14
    let mut result = Vec::new();
    let ver_inst: u16 = 0xF; // inst=0, ver=0xF
    result.extend_from_slice(&ver_inst.to_le_bytes());
    result.extend_from_slice(&RT_CRYPT_SESSION10_CONTAINER.to_le_bytes());
    result.extend_from_slice(&(data.len() as u32).to_le_bytes());
    result.extend_from_slice(&data);

    result
}

// ============================================================================
// Persist 对象重排
// ============================================================================

/// 重排 persist 对象，使其在流中按 persistId 顺序物理排列。
///
/// # 背景
///
/// MS-PPT 加密机制要求 persist 对象在 PowerPoint Document 流中按 persistId 顺序
/// 物理排列。msoffcrypto 和 PowerPoint 解密时使用
/// `recLen = next_offset - offset - 8` 计算每个 persist 对象的长度，
/// 这假设 persist 对象按 persistId 顺序排列。
///
/// 如果 persist 对象未按 persistId 顺序排列（如 WPS 生成的文件），
/// 解密时 `recLen` 计算错误（甚至为负数），解密破坏数据，文件无法打开。
///
/// 本函数在加密前将 persist 对象物理重排为 persistId 顺序，并更新
/// PersistDirectoryAtom 中的偏移量。
///
/// # 参数
/// - `ppt_data`：PowerPoint Document 流数据（可变引用，原地修改）
/// - `persist_entries`：从 PDA 解析的 `(persistId, offset)` 列表
/// - `pd_offset`：PersistDirectoryAtom 在流中的偏移量
///
/// # 返回
/// - 成功：返回更新后的 `(persistId, new_offset)` 列表
///
/// # 错误
/// - [`Error::Ppt97`]：record header 解析失败 / 数据越界
pub fn reorder_persist_objects(
    #[allow(clippy::ptr_arg)] ppt_data: &mut Vec<u8>,
    persist_entries: &[(u32, u32)],
    pd_offset: usize,
) -> Result<Vec<(u32, u32)>> {
    // 1. 读取所有 persist 对象的完整数据（header + data）
    let mut objects: Vec<(u32, Vec<u8>)> = Vec::new();
    for (pid, offset) in persist_entries {
        let offset = *offset as usize;
        let (_, _, rec_type, rec_len) = parse_record_header(ppt_data, offset)?;

        // 跳过 UserEditAtom 和 PersistDirectoryAtom（不移动，不加密）
        if rec_type == RT_USER_EDIT_ATOM || rec_type == RT_PERSIST_DIRECTORY_ATOM {
            continue;
        }

        let total_len = 8 + rec_len as usize;
        if offset + total_len > ppt_data.len() {
            return Err(Error::ppt97(format!(
                "reorder_persist_objects: persist {} (offset {}, len {}) out of range",
                pid, offset, total_len
            )));
        }
        let data = ppt_data[offset..offset + total_len].to_vec();
        objects.push((*pid, data));
    }

    // 2. 按 persistId 排序
    objects.sort_by_key(|(pid, _)| *pid);

    // 3. 找到最小偏移量（重排后的起始位置）
    // 因为 persist 对象在流中是连续的，最小偏移量就是重排后的起始位置
    let min_offset = persist_entries
        .iter()
        .map(|(_, off)| *off)
        .min()
        .unwrap_or(0) as usize;

    // 4. 按排序顺序连续写回，计算新偏移量
    let mut new_offsets: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
    let mut write_pos = min_offset;

    for (pid, data) in &objects {
        if write_pos + data.len() > ppt_data.len() {
            return Err(Error::ppt97(format!(
                "reorder_persist_objects: write_pos {} + len {} out of range",
                write_pos,
                data.len()
            )));
        }
        ppt_data[write_pos..write_pos + data.len()].copy_from_slice(data);
        new_offsets.insert(*pid, write_pos as u32);
        write_pos += data.len();
    }

    // 5. 更新 PDA 中的 rgPersistOffset
    // PDA 格式: [header(8字节)] [entry(4字节: persistId|cPersist)] [rgPersistOffset(cPersist*4字节)]
    let (_, _, _, pd_rec_len) = parse_record_header(ppt_data, pd_offset)?;
    let pd_data_end = pd_offset + 8 + pd_rec_len as usize;
    let mut pos = pd_offset + 8;

    while pos + 4 <= pd_data_end {
        let entry = read_u32_le(ppt_data, pos)?;
        let persist_id = entry & 0xFFFFF;
        let c_persist = (entry >> 20) & 0xFFF;
        pos += 4;

        for j in 0..c_persist {
            let pid = persist_id + j;
            if let Some(&new_offset) = new_offsets.get(&pid) {
                write_u32_le(ppt_data, pos, new_offset)?;
            }
            pos += 4;
        }
    }

    // 6. 构建新的 persist_entries（保持原始顺序，但更新 offset）
    // 对于未移动的对象（如 UE/PDA），保留原始 offset
    let new_entries: Vec<(u32, u32)> = persist_entries
        .iter()
        .map(|(pid, old_off)| {
            let new_off = new_offsets.get(pid).copied().unwrap_or(*old_off);
            (*pid, new_off)
        })
        .collect();

    Ok(new_entries)
}

// ============================================================================
// 加密主流程
// ============================================================================

/// 加密 PowerPoint Document stream 与 Current User stream。
///
/// 完整流程：
/// 1. 解析 CurrentUser → offsetToCurrentEdit
/// 2. 解析 UserEditAtom → offsetPersistDirectory
/// 3. 解析 PersistDirectoryAtom → persist entries
/// 4. **重排 persist 对象**为 persistId 顺序（msoffcrypto/PowerPoint 解密要求）
/// 5. 生成 salt + verifier，用 block=0 的 key 加密 verifier
/// 6. 遍历 persist 对象，跳过 UserEditAtom/PersistDirectoryAtom，加密其余对象
/// 7. 构造 CryptSession10Container 插入到 PDA 之前
/// 8. 修改 UserEditAtom: recLen 28→32，添加 encryptSessionPersistIdRef
/// 9. 修改 PersistDirectoryAtom: cPersist+1，添加 CryptSession10Container offset
/// 10. 修改 CurrentUser: headerToken 0xE391C05F→0xF3D1C4DF，更新 offsetToCurrentEdit
///
/// # 关键设计决策
///
/// - **不加密 Pictures stream**：MS-PPT 2.3.7 规范要求 Pictures stream 必须加密，
///   但 msoffcrypto（参考实现）并未实现 Pictures stream 解密（源码中标注为 TODO）。
///   经实测：加密 Pictures stream 后，WPS 解密时得到乱码，导致文件无法打开。
///   推测 WPS 也不解密 Pictures stream，或期望用不同的加密方式。
///   最稳妥方案：不加密 Pictures stream，保留原始内容。
/// - **不创建 EncryptedSummaryInfo stream**：因 `fDocProps=1`（文档属性不加密），
///   保留原始 SummaryInformation 和 DocumentSummaryInformation streams 即可。
///
/// # 参数
/// - `ppt_data`：PowerPoint Document stream（原地修改）
/// - `cu_data`：Current User stream（原地修改）
/// - `password`：明文密码
///
/// # 错误
/// - [`Error::Ppt97`]：stream 结构损坏 / record 类型不匹配 / 已加密文件重复加密
pub fn encrypt_ppt_stream(
    #[allow(clippy::ptr_arg)] ppt_data: &mut Vec<u8>,
    #[allow(clippy::ptr_arg)] cu_data: &mut Vec<u8>,
    password: &str,
) -> Result<()> {
    // 解析 CurrentUserAtom
    let (_, _, cu_type, _) = parse_record_header(cu_data, 0)?;
    if cu_type != RT_CURRENT_USER_ATOM {
        return Err(Error::ppt97(format!(
            "encrypt: expected CurrentUserAtom (0x{:04X}), got 0x{:04X}",
            RT_CURRENT_USER_ATOM, cu_type
        )));
    }

    // 加密检测（参考 msoffcrypto-python ppt97.py 的 headerToken 判定模式）：
    // CurrentUserAtom.headerToken 位于 offset 12，未加密=0xE391C05F，已加密=0xF3D1C4DF。
    // 若文件已加密，提前返回清晰错误，避免后续解析产生混淆性失败。
    let header_token = read_u32_le(cu_data, 12)?;
    if header_token == HEADER_TOKEN_ENCRYPTED {
        return Err(Error::ppt97(
            "encrypt: file already encrypted (headerToken=0xF3D1C4DF)".to_string(),
        ));
    }

    let offset_to_current_edit = read_u32_le(cu_data, 16)?;

    // 解析 UserEditAtom
    let ue_offset = offset_to_current_edit as usize;
    let (_, _, ue_type, ue_len) = parse_record_header(ppt_data, ue_offset)?;
    if ue_type != RT_USER_EDIT_ATOM {
        return Err(Error::ppt97(format!(
            "encrypt: expected UserEditAtom (0x{:04X}), got 0x{:04X}",
            RT_USER_EDIT_ATOM, ue_type
        )));
    }
    if ue_len != 28 {
        return Err(Error::ppt97(format!(
            "encrypt: file already encrypted or malformed (UserEditAtom.recLen={}, expected 28)",
            ue_len
        )));
    }

    let offset_persist_dir = read_u32_le(ppt_data, ue_offset + 20)?;

    // 解析 PersistDirectoryAtom
    let persist_entries = parse_persist_directory(ppt_data, offset_persist_dir as usize)?;

    // 关键修复：重排 persist 对象为 persistId 顺序
    // msoffcrypto/PowerPoint 解密时使用 `recLen = next_offset - offset - 8` 计算
    // persist 对象长度，这假设 persist 对象按 persistId 顺序物理排列。
    let persist_entries =
        reorder_persist_objects(ppt_data, &persist_entries, offset_persist_dir as usize)?;

    // 生成加密参数
    let mut rng = rand::thread_rng();
    let salt: Vec<u8> = (0..SALT_SIZE).map(|_| rng.gen()).collect();
    let verifier_plain: Vec<u8> = (0..VERIFIER_SIZE).map(|_| rng.gen()).collect();
    let verifier_hash = Sha1::digest(&verifier_plain).to_vec();

    // 用 block=0 的 key 加密 verifier 和 verifierHash
    // 注意：必须使用同一个 RC4 流连续加密（与 msoffcrypto verifypw 一致）
    let key_block0 = make_key(password, &salt, KEY_SIZE_BITS, 0);
    let mut rc4 = Rc4::new(&key_block0);
    let mut encrypted_verifier = verifier_plain.clone();
    rc4.process(&mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash.clone();
    rc4.process(&mut encrypted_verifier_hash);

    // 加密 persist 对象
    // 按 offset 排序，以便计算每个对象的边界
    let mut sorted_entries = persist_entries.clone();
    sorted_entries.sort_by_key(|(_, offset)| *offset);

    for (pid, poff) in &sorted_entries {
        let poff = *poff as usize;
        let (_, _, rec_type, rec_len) = parse_record_header(ppt_data, poff)?;

        // 跳过 UserEditAtom 和 PersistDirectoryAtom（不加密）
        if rec_type == RT_USER_EDIT_ATOM || rec_type == RT_PERSIST_DIRECTORY_ATOM {
            continue;
        }

        // 加密整个 record（header + data）
        let total_len = 8 + rec_len as usize;
        let record_data = &mut ppt_data[poff..poff + total_len];
        let encrypted = encrypt_persist_object(password, &salt, KEY_SIZE_BITS, record_data, *pid);
        record_data.copy_from_slice(&encrypted);
    }

    // 构造 CryptSession10Container
    let crypt_session = build_crypt_session10_container(
        &salt,
        &encrypted_verifier,
        &encrypted_verifier_hash,
        KEY_SIZE_BITS,
    );

    // 步骤1：在 PDA 之前插入 CryptSession10Container
    // 关键修复：CryptSession10Container 必须紧接在 persist 对象之后，不能在 PDA/UserEditAtom 之后。
    // msoffcrypto/PowerPoint 解密时计算 recLen = next_offset - offset - 8，
    // 如果 PDA/UserEditAtom 在最后一个 persist 对象和 CryptSession10Container 之间，
    // recLen 会包含 PDA/UserEditAtom 的字节，导致解密破坏这些非加密数据。
    // 正确布局：[persist objects][CryptSession10Container][PDA][UserEditAtom]
    let crypt_session_offset = offset_persist_dir as u32;
    let crypt_session_len = crypt_session.len();
    ppt_data.splice(
        offset_persist_dir as usize..offset_persist_dir as usize,
        crypt_session.iter().copied(),
    );

    // PDA 和 UserEditAtom 已向后移动 crypt_session_len 字节
    let pd_offset_new = offset_persist_dir as usize + crypt_session_len;
    let ue_offset_after_cs = ue_offset + crypt_session_len;

    // 步骤2：在 PDA 的 rgPersistOffset 末尾插入 4 字节占位符
    // 这会导致 UserEditAtom 向后移动 4 字节
    let pd_data_start = pd_offset_new + 8;
    let entry_val = read_u32_le(ppt_data, pd_data_start)?;
    // persistId: 20 bits, cPersist: 12 bits（MS-PPT 规范）
    let entry_pid = entry_val & 0xFFFFF;
    let entry_cpersist = (entry_val >> 20) & 0xFFF;

    // CryptSession10Container 的 persistId = entry_pid + 原始 cPersist
    let crypt_session_pid = entry_pid + entry_cpersist;

    let insert_pos = pd_data_start + 4 + entry_cpersist as usize * 4;
    ppt_data.splice(insert_pos..insert_pos, std::iter::once(0u8).cycle().take(4));

    // 更新 PersistDirectoryAtom 的 recLen（+4）和 cPersist（+1）
    let (_, _, _, pd_len) = parse_record_header(ppt_data, pd_offset_new)?;
    write_u32_le(ppt_data, pd_offset_new + 4, pd_len + 4)?;
    let new_entry_val = entry_pid | ((entry_cpersist + 1) << 20);
    write_u32_le(ppt_data, pd_data_start, new_entry_val)?;

    // UserEditAtom 又向后移动 4 字节
    let ue_offset_new = ue_offset_after_cs + 4;

    // 步骤3：在 UserEditAtom 末尾添加 4 字节（encryptSessionPersistIdRef 占位符）
    // UserEditAtom 在 stream 末尾，所以直接 extend
    ppt_data.extend_from_slice(&[0u8; 4]);

    // 步骤4：修改 UserEditAtom（在新位置）
    // recLen 28→32，添加 encryptSessionPersistIdRef
    write_u32_le(ppt_data, ue_offset_new + 4, 32)?;
    write_u32_le(ppt_data, ue_offset_new + 8 + 28, crypt_session_pid)?;
    // 更新 maxPersistWritten = crypt_session_pid
    write_u32_le(ppt_data, ue_offset_new + 28, crypt_session_pid)?;
    // 更新 offsetPersistDirectory 指向新的 PDA 位置
    write_u32_le(ppt_data, ue_offset_new + 20, pd_offset_new as u32)?;

    // 步骤5：在 PDA 的 rgPersistOffset 末尾写入 crypt_session_offset
    write_u32_le(ppt_data, insert_pos, crypt_session_offset)?;

    // 步骤6：修改 Current User
    // headerToken 0xE391C05F → 0xF3D1C4DF
    // offsetToCurrentEdit 需要更新为 ue_offset_new
    write_u32_le(cu_data, 12, HEADER_TOKEN_ENCRYPTED)?;
    write_u32_le(cu_data, 16, ue_offset_new as u32)?;

    Ok(())
}
