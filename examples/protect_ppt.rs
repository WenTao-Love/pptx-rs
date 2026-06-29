//! 为 .ppt（PowerPoint 97-2003 二进制格式）文件设置 RC4 CryptoAPI 加密。
//!
//! 这是一个 **示例程序**，演示如何使用 [`pptx_rs::ppt97::encrypt`] 库 API
//! 为 .ppt 文件批量加密。实际的加密逻辑已封装到 [`pptx_rs::ppt97`] 模块中。
//!
//! # 加密参数
//!
//! - 算法：RC4 + SHA1 密钥派生（MS-OFFCRYPTO 规范）
//! - 密钥位数：128 bit
//! - 容器：OLE2/CFB（PowerPoint Document + Current User stream）
//! - 加密结构：CryptSession10Container + PersistDirectoryAtom 更新
//! - 每个 persist 对象独立加密（block=persistId，分段加密）
//! - 默认密码：`pptx-rs-secret`
//!
//! # 用法
//!
//! 在 pptx-rs 目录执行：
//! ```bash
//! cargo run --example protect_ppt
//! ```
//!
//! 会扫描 `_test/` 目录下所有 `.ppt` 文件，加密后输出到 `_test_out/protected_<原名>.ppt`。

use std::path::Path;

use pptx_rs::ppt97::encrypt;

/// 默认密码（与 README.md 中记载一致）。
const PASSWORD: &str = "pptx-rs-secret";

/// 批量为 `_test/` 目录下的 .ppt 文件设置加密。
fn main() -> Result<(), Box<dyn std::error::Error>> {
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

        match encrypt(Path::new(&p), PASSWORD) {
            Ok(encrypted_data) => {
                std::fs::create_dir_all("_test_out")?;
                let out_path = format!("_test_out/protected_{}", fname);
                std::fs::write(&out_path, &encrypted_data)?;
                println!("已加密：{}", out_path);
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
