//! 同时为 .ppt 文件注入水印和加密。
//!
//! 这是一个 **示例程序**，演示如何使用 [`pptx::ppt97::add_watermark_and_encrypt`]
//! 库 API 一步到位地为 .ppt 文件同时注入水印和加密。实际的逻辑已封装到
//! [`pptx::ppt97`] 模块中。
//!
//! # 处理顺序
//!
//! 必须 **先加水印再加密**：
//! 1. 先注入水印（修改 record 结构，更新 persist 对象 offset）
//! 2. 再设置加密（加密所有 persist 对象，添加 CryptSession10Container）
//!
//! 因为加密后所有 persist 对象被加密，无法直接修改 record 结构；
//! 而水印注入需要修改 record 结构。
//!
//! # 加密参数
//!
//! - 算法：RC4 + SHA1 密钥派生（MS-OFFCRYPTO 规范）
//! - 密钥位数：128 bit
//! - 默认密码：`pptx-rs-secret`
//!
//! # 水印特性
//!
//! - 44pt 大字号，45 度旋转，中灰色（200,200,200）
//! - 作为 MainMaster 母版的背景元素注入（z-order 最低的真正子形状）
//! - 无填充、无边框、锁定不可编辑（FOPT 保护位 0x01C2=0x0D）
//!
//! # 用法
//!
//! 在 pptx-rs 目录执行：
//! ```bash
//! cargo run --example watermark_and_protect_ppt
//! ```
//!
//! 会扫描 `_test/` 目录下所有 `.ppt` 文件，注入水印并加密后输出到
//! `_test_out/wm_protected_<原名>.ppt`。

use std::path::Path;

use pptx::ppt97::{add_watermark_and_encrypt, WatermarkConfig};

/// 默认密码（与 README.md 中记载一致）。
const PASSWORD: &str = "pptx-rs-secret";

/// 批量为 `_test/` 目录下的 .ppt 文件同时注入水印和加密。
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = WatermarkConfig::default();
    let entries = std::fs::read_dir("_test")?;
    let mut processed = 0;
    let mut skipped = 0;

    for entry in entries {
        let entry = entry?;
        let p = entry.path();
        if p.extension().and_then(|e| e.to_str()) != Some("ppt") {
            continue;
        }
        let fname = p.file_name().unwrap().to_string_lossy().to_string();

        match add_watermark_and_encrypt(Path::new(&p), &config, PASSWORD) {
            Ok(data) => {
                std::fs::create_dir_all("_test_out")?;
                let out_path = format!("_test_out/wm_protected_{}", fname);
                std::fs::write(&out_path, &data)?;
                println!("已加水印并加密：{}", out_path);
                processed += 1;
            }
            Err(e) => {
                eprintln!("跳过 {}: {}", fname, e);
                skipped += 1;
            }
        }
    }
    println!("共处理 {} 个 ppt 文件，跳过 {} 个", processed, skipped);
    Ok(())
}
