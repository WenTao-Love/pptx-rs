//! 为 .pptx 同时加水印和加密（OOXML Agile Encryption）。
//!
//! 流程：先加水印（修改 ZIP 中的 slide XML），再加密（加密整个 ZIP 包）。
//!
//! 加密逻辑完全对齐 msoffcrypto-python：
//! - 子密钥派生使用 blockKey（MS-OFFCRYPTO 规范）
//! - EncryptionInfo 字段使用 password_salt 作为 IV
//! - EncryptedPackage 分段加密（4096字节/段），每段不同 IV
//! - EncryptedPackage 大小头为 u64（8字节）
//!
//! 用法：在 pptx-rs 目录执行
//!   cargo run --example watermark_and_protect
//! 密码：pptx-rs-secret

use std::io::{Read, Write};
use std::path::Path;

use pptx::opc::OpcPackage;
use rand::Rng;
use sha2::{Digest, Sha512};

const PASSWORD: &str = "pptx-rs-secret";

// ============================================================================
// 水印配置（参考 python-pptx / Aspose.Slides 的参数化模式）
// ============================================================================

/// 水印配置参数。
///
/// 参考 python-pptx 的 `shapes.add_textbox()` + `font` + `shape.rotation` 组合 API
/// 和 Aspose.Slides 的 `PortionFormat` 参数化模式，将水印的可变属性集中管理。
///
/// 字段对应关系：
/// - `text` ↔ python-pptx `text_frame.text`
/// - `font_size_pt` ↔ python-pptx `font.size = Pt(n)`（OOXML 中单位为 1/100 pt）
/// - `color_hex` ↔ python-pptx `font.color.rgb`（OOXML srgbClr val 属性）
/// - `alpha_percent` ↔ OOXML `a:alpha val`（0-100000，1000 = 1%）
/// - `rotation_deg` ↔ python-pptx `shape.rotation`（OOXML 中单位为 1/60000 度）
struct WatermarkConfig {
    /// 水印文本内容
    text: String,
    /// 字号（磅），OOXML 中转换为 sz 属性（×100）
    font_size_pt: u16,
    /// 文字颜色（十六进制 RGB，如 "BFBFBF"）
    color_hex: String,
    /// 不透明度百分比（0-100），OOXML alpha val = alpha_percent × 1000
    /// （OOXML alpha 语义：0=全透明，100000=不透明；40 表示 40% 不透明度）
    alpha_percent: u8,
    /// 旋转角度（度），OOXML rot = rotation_deg × 60000
    rotation_deg: i32,
}

impl Default for WatermarkConfig {
    fn default() -> Self {
        Self {
            text: "pptx-rs WATERMARK".to_string(),
            font_size_pt: 40,
            color_hex: "BFBFBF".to_string(),
            alpha_percent: 40,
            rotation_deg: -45,
        }
    }
}

/// 构建水印 shape 的 XML。
///
/// **不重新声明 xmlns:p 和 xmlns:a**（已在父元素 <p:sld> 中声明）。
/// 格式与 python-pptx 生成的 textbox 一致，确保 WPS 兼容。
///
/// # 参数
/// - `id`：shape 的唯一 ID（cNvPr id 属性）
/// - `config`：水印配置
fn build_watermark_xml(id: u32, config: &WatermarkConfig) -> String {
    // OOXML 单位转换：
    // - rot: 1度 = 60000（rotation_deg × 60000）
    // - sz: 1pt = 100（font_size_pt × 100）
    // - alpha: 100% = 100000，不透明度 40% → alpha val = 40000
    //   OOXML alpha 语义：val 表示不透明度（0=全透明，100000=不透明）
    let rot_ooxml = config.rotation_deg * 60000;
    let sz_ooxml = config.font_size_pt as u32 * 100;
    let alpha_ooxml = config.alpha_percent as u32 * 1000;

    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="pptx-rs Watermark"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="{rot}"><a:off x="457200" y="2743200"/><a:ext cx="8229600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="{sz}" b="1"><a:solidFill><a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill><a:latin typeface="Calibri"/><a:ea typeface="宋体"/></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#,
        id = id,
        rot = rot_ooxml,
        sz = sz_ooxml,
        color = config.color_hex,
        alpha = alpha_ooxml,
        text = config.text,
    )
}

/// 检测 .pptx 文件是否已加密。
///
/// 参考 Apache POI `POIFSFileSystem.hasPOIFSHeader()` + `DirectoryNode.hasEntry("EncryptedPackage")`
/// 的两层判定模式：
/// 1. 检查文件头是否为 OLE2 复合文档签名（加密的 OOXML 文件会被包进 OLE2 容器）
/// 2. 正常 .pptx 文件是 ZIP 格式（魔数 PK\x03\x04），加密后变为 OLE2 格式（魔数 D0CF11E0）
///
/// # 参数
/// - `path`：文件路径
///
/// # 返回值
/// - `true`：文件已加密（OLE2 容器）
/// - `false`：文件未加密（ZIP 容器）
fn is_pptx_encrypted(path: &Path) -> bool {
    // OLE2 复合文档魔数（MS-CFB 规范）：D0 CF 11 E0 A1 B1 1A E1
    const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    match std::fs::read(path) {
        Ok(data) => data.len() >= 8 && data[..8] == OLE2_MAGIC,
        Err(_) => false,
    }
}

// ============================================================================
// 加密相关常量
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
// DataSpaces 流的原始字节
// ============================================================================
//
// 这些字节是 MS-OFFCRYPTO 规范固定的常量，不是从特定文件提取的"魔法值"。
// 参考 msoffcrypto-python 的 DefaultContent 类（注释明确标注 "Lifted off of
// Herumi/msoffice"），这些流的内容对所有 ECMA-376 加密文档完全相同：
// - Version: "Microsoft.Container.DataSpaces" + 版本 1.1.1
// - DataSpaceMap: 引用 "EncryptedPackage" → "StrongEncryptionDataSpace"
// - StrongEncryptionDataSpace: 引用 "StrongEncryptionTransform"
// - Primary: CLSID {FF9A3F03-56EF-4613-BDD5-5A41C1D07246}
//            + "Microsoft.Container.EncryptionTransform"
//
// 注意：OLE2 容器结构（FAT/MiniFAT/目录红黑树）由 cfb crate 程序化构造，
// 这些流的内容是规范固定的，硬编码是正确做法（与 msoffcrypto-python 一致）。

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
    let mut skipped = 0;

    // 水印配置（参考 python-pptx 参数化模式，集中管理可变属性）
    let config = WatermarkConfig::default();

    for entry in entries {
        let entry = entry?;
        let p = entry.path();
        if p.extension().and_then(|e| e.to_str()) != Some("pptx") {
            continue;
        }
        let fname = p.file_name().unwrap().to_string_lossy().to_string();

        match watermark_and_encrypt(&p, PASSWORD, &config) {
            Ok(encrypted_data) => {
                std::fs::create_dir_all("_test_out")?;
                let out_path = format!("_test_out/wm_protected_{}", fname);
                std::fs::write(&out_path, &encrypted_data)?;
                println!("已加水印+加密：{}", out_path);
                processed += 1;
            }
            Err(e) => {
                eprintln!("跳过 {}: {}", fname, e);
                skipped += 1;
            }
        }
    }
    println!("共处理 {} 个 pptx 文件，跳过 {} 个", processed, skipped);
    Ok(())
}

/// 为 PPTX 文件同时加水印和加密。
///
/// 步骤：
/// 1. 加密检测：若文件已加密（OLE2 容器）则提前返回清晰错误
/// 2. 用 OpcPackage 加载 PPTX，在每张幻灯片注入水印 shape
/// 3. 将修改后的 ZIP 包序列化到内存
/// 4. 用 OOXML Agile Encryption 加密整个 ZIP 包
///
/// # 参数
/// - `input_path`：输入 .pptx 文件路径
/// - `password`：加密密码
/// - `config`：水印配置
///
/// # 错误
/// - 文件已加密（OLE2 容器，魔数 D0CF11E0）
/// - OpcPackage 加载失败
/// - 加密过程失败
fn watermark_and_encrypt(
    input_path: &Path,
    password: &str,
    config: &WatermarkConfig,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    // 加密检测（参考 Apache POI POIFSFileSystem.hasPOIFSHeader 模式）：
    // 正常 .pptx 是 ZIP 容器（PK\x03\x04），加密后变为 OLE2 容器（D0CF11E0）。
    // 提前检测避免后续 OpcPackage::load 产生混淆性失败。
    if is_pptx_encrypted(input_path) {
        return Err("文件已加密，无法重复加密（OLE2 容器）".into());
    }

    // === 第一步：加水印 ===
    let watermarked_zip = add_watermark(input_path, config)?;

    // === 第二步：加密 ===
    let encrypted = encrypt_package(&watermarked_zip, password)?;

    Ok(encrypted)
}

/// 为 PPTX 文件的所有幻灯片添加水印，返回修改后的 ZIP 字节。
///
/// 使用 OpcPackage 底层操作：直接在 slide XML 中注入水印 shape。
/// 水印放在 spTree 末尾（z-order 最顶层），配合半透明效果。
///
/// # 参数
/// - `input_path`：输入 .pptx 文件路径
/// - `config`：水印配置（文本/字号/颜色/透明度/旋转角度）
fn add_watermark(
    input_path: &Path,
    config: &WatermarkConfig,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut pkg = OpcPackage::load(input_path)?;

    // 找到所有 slide*.xml 的 partname
    let mut slide_names: Vec<String> = Vec::new();
    for part in pkg.iter_parts() {
        let n = part.partname.as_str();
        if n.starts_with("/ppt/slides/slide") && n.ends_with(".xml") {
            slide_names.push(n.to_string());
        }
    }
    slide_names.sort();

    // 对每个 slide XML 注入水印
    // shape id 从 9001 开始递增，避免与现有 shape 冲突
    // （参考 python-pptx 的 _next_shape_id 模式，但简化为固定基数）
    let mut max_id: u32 = 9000;
    for name in &slide_names {
        max_id += 1;
        if let Some(part) = pkg.get_part_mut(name) {
            let new_xml = inject_watermark(&part.blob, max_id, config);
            part.blob = new_xml.into_bytes();
        }
    }

    // 序列化到内存
    let buf = pkg.to_bytes()?;
    Ok(buf)
}

/// 在 slide XML 的 `</p:spTree>` 之前注入水印 shape。
///
/// 幂等性：如果已存在水印（按 name 属性 "pptx-rs Watermark" 检测）则跳过，
/// 避免重复注入。参考 Aspose.Slides 的按 shape name 查重模式。
///
/// # 参数
/// - `blob`：slide XML 字节
/// - `id`：shape 唯一 ID（cNvPr id 属性）
/// - `config`：水印配置
fn inject_watermark(blob: &[u8], id: u32, config: &WatermarkConfig) -> String {
    let s = std::str::from_utf8(blob).unwrap_or("").to_string();

    // 幂等性检查：按 name 属性查重（参考 Aspose.Slides 模式）
    if s.contains("pptx-rs Watermark") || s.contains("pptx-rs WATERMARK") {
        return s;
    }

    // 构建水印 shape XML（参数化，参考 python-pptx 的 textbox 生成模式）
    let sp_xml = build_watermark_xml(id, config);

    // 在 </p:spTree> 之前插入水印（z-order 最顶层）
    if let Some(pos) = s.find("</p:spTree>") {
        let mut out = String::with_capacity(s.len() + sp_xml.len());
        out.push_str(&s[..pos]);
        out.push_str(&sp_xml);
        out.push_str(&s[pos..]);
        out
    } else {
        s
    }
}

// ============================================================================
// 加密相关函数
// ============================================================================

/// 使用 OOXML Agile Encryption 加密 ZIP 包数据。
///
/// 完全对齐 msoffcrypto-python 的加密逻辑：
/// - 密码派生：H₀ = SHA512(salt + password_utf16le), Hₙ = SHA512(LE32(n) + Hₙ₋₁)
/// - 子密钥派生：key = SHA512(h + blockKey)[:keyBits/8]
/// - EncryptionInfo 字段 IV = password_salt（补齐到块大小）
/// - EncryptedPackage 分段加密，每段 IV = SHA512(keyDataSalt + LE32(segmentIndex))[:16]
fn encrypt_package(
    package_data: &[u8],
    password: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
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
    let h = derive_iterated_hash(password, &password_salt, spin_count);

    // 步骤2：用 blockKey 派生各加密密钥
    let key1 = derive_encryption_key(&h, &BLKKEY_VERIFIER_HASH_INPUT, key_length);
    let key2 = derive_encryption_key(&h, &BLKKEY_ENCRYPTED_VERIFIER_HASH_VALUE, key_length);
    let key3 = derive_encryption_key(&h, &BLKKEY_ENCRYPTED_KEY_VALUE, key_length);

    // 步骤3：生成验证器
    let verifier_hash_input_raw: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();
    let verifier_hash_input =
        resize_buffer(&verifier_hash_input_raw, round_up(salt_size, block_size));
    let iv_salt = normalize_key(&password_salt, block_size);

    let encrypted_verifier_hash_input = aes_cbc_encrypt_raw(&key1, &iv_salt, &verifier_hash_input)?;
    let verifier_hash = Sha512::digest(&verifier_hash_input);
    let verifier_hash_padded = resize_buffer(&verifier_hash, round_up(hash_size, block_size));
    let encrypted_verifier_hash_value =
        aes_cbc_encrypt_raw(&key2, &iv_salt, &verifier_hash_padded)?;

    // 步骤4：生成 secret_key
    let secret_key_raw: Vec<u8> = (0..salt_size).map(|_| rng.gen()).collect();
    let secret_key = normalize_key(&secret_key_raw, key_length);
    let encrypted_key_value = aes_cbc_encrypt_raw(&key3, &iv_salt, &secret_key)?;

    // 步骤5：加密 EncryptedPackage
    let encrypted_package = encrypt_payload(package_data, &secret_key, &key_data_salt)?;

    // 步骤6：生成 dataIntegrity
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

    // EncryptionInfo stream：MajorVersion(u16) + MinorVersion(u16) + Reserved(u32=0x40) + XML
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
        hasher2.update(i.to_le_bytes());
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
    if !data.len().is_multiple_of(16) {
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
fn encrypt_payload(
    package_data: &[u8],
    secret_key: &[u8],
    key_data_salt: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut result = Vec::new();
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
        let padded_chunk = if !chunk.len().is_multiple_of(16) {
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
fn generate_integrity_parameter(
    encrypted_package: &[u8],
    secret_key: &[u8],
    key_data_salt: &[u8],
    hash_size: usize,
    block_size: usize,
) -> Result<(Vec<u8>, Vec<u8>), Box<dyn std::error::Error>> {
    let mut rng = rand::thread_rng();
    let hmac_key: Vec<u8> = (0..hash_size).map(|_| rng.gen()).collect();

    let mut iv1_hasher = Sha512::new();
    iv1_hasher.update(key_data_salt);
    iv1_hasher.update(BLKKEY_DATA_INTEGRITY_1);
    let iv1 = normalize_key(&iv1_hasher.finalize(), block_size);

    let mut iv2_hasher = Sha512::new();
    iv2_hasher.update(key_data_salt);
    iv2_hasher.update(BLKKEY_DATA_INTEGRITY_2);
    let iv2 = normalize_key(&iv2_hasher.finalize(), block_size);

    let hmac_key_padded = resize_buffer(&hmac_key, round_up(hash_size, block_size));
    let encrypted_hmac_key = aes_cbc_encrypt_raw(secret_key, &iv1, &hmac_key_padded)?;

    use hmac::{Hmac, Mac};
    type HmacSha512 = Hmac<Sha512>;
    let mut mac = HmacSha512::new_from_slice(&hmac_key)?;
    mac.update(encrypted_package);
    let hmac_value = mac.finalize().into_bytes();
    let hmac_value_padded = resize_buffer(&hmac_value, round_up(hash_size, block_size));
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
