//! 为 .ppt（PowerPoint 97-2003 二进制格式）文件注入水印。
//!
//! 这是一个 **示例程序**，演示如何使用 [`pptx_rs::ppt97::add_watermark`] 库 API
//! 为 .ppt 文件批量注入水印。实际的水印注入逻辑已封装到
//! [`pptx_rs::ppt97`] 模块中。
//!
//! # 水印特性
//!
//! - 全屏覆盖（ClientAnchor 0,0 到 5760,4320）
//! - 44pt 大字号，45 度旋转，中灰色（200,200,200）
//! - 无填充、无边框、锁定不可编辑（FOPT 保护位 0x01C2=0x0D）
//! - 作为 MainMaster 母版的背景元素注入（z-order 最低的真正子形状）
//!
//! # 用法
//!
//! 在 pptx-rs 目录执行：
//! ```bash
//! cargo run --example watermark_ppt
//! ```
//!
//! 会扫描 `_test/` 目录下所有 `.ppt` 文件，加水印后输出到 `_test_out/wm_<原名>.ppt`。

use std::path::Path;

use pptx_rs::ppt97::{add_watermark, WatermarkConfig};

/// 批量为 `_test/` 目录下的 .ppt 文件注入水印。
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

        match add_watermark(Path::new(&p), &config) {
            Ok(data) => {
                std::fs::create_dir_all("_test_out")?;
                let out_path = format!("_test_out/wm_{}", fname);
                std::fs::write(&out_path, &data)?;
                println!("已加水印：{}", out_path);
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
