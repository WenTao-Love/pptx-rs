//! PowerPoint 97-2003 二进制格式（`.ppt`）支持模块。
//!
//! 本模块为 PowerPoint 97-2003 二进制格式（`.ppt`）文件提供**水印注入**
//! 与 **RC4 CryptoAPI 加密**能力。`.ppt` 文件使用 OLE2/CFB 容器 + 二进制
//! record 树结构，与 `.pptx`（ZIP+XML）完全不同，因此需要独立模块。
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx **不支持** .ppt 97-2003 二进制格式（仅支持 .pptx）。
//! 本模块填补了这一空白，对标 LibreOffice 的 `oox::ole` 和 Apache POI 的
//! `org.apache.poi.hslf`（HSLF = Horrible Slide Layout Format）。
//!
//! # 模块组织
//!
//! - [`record`]：PPT record 树解析工具（基础设施）
//! - [`ole`]：OLE2/CFB 容器操作（`write_stream` / `fix_mini_fat`）
//! - [`watermark`]：水印注入（[`watermark::WatermarkConfig`]）
//! - [`crypto`]：RC4 CryptoAPI 加密
//!
//! # 公共 API
//!
//! - [`add_watermark`]：为 .ppt 文件注入水印
//! - [`encrypt`]：为 .ppt 文件设置 RC4 CryptoAPI 加密
//! - [`add_watermark_and_encrypt`]：同时注入水印和加密
//!
//! # 规范依据
//!
//! - [MS-PPT]：PowerPoint 97-2003 二进制文件格式规范
//! - [MS-ODRAW]：Office Drawing 97-2003 二进制格式（Escher OfficeArt）
//! - [MS-CFB]：Compound File Binary 文件格式（OLE2 容器）
//! - [MS-OFFCRYPTO] 2.3.5：RC4 CryptoAPI Encryption
//!
//! # 示例
//!
//! ```no_run
//! use pptx_rs::ppt97::{add_watermark, encrypt, add_watermark_and_encrypt, watermark::WatermarkConfig};
//! use std::path::Path;
//!
//! // 仅加水印
//! let config = WatermarkConfig::default();
//! let watermarked = add_watermark(Path::new("input.ppt"), &config).unwrap();
//! std::fs::write("watermarked.ppt", &watermarked).unwrap();
//!
//! // 仅加密
//! let encrypted = encrypt(Path::new("input.ppt"), "my-password").unwrap();
//! std::fs::write("encrypted.ppt", &encrypted).unwrap();
//!
//! // 水印 + 加密
//! let both = add_watermark_and_encrypt(
//!     Path::new("input.ppt"),
//!     &config,
//!     "my-password",
//! ).unwrap();
//! std::fs::write("both.ppt", &both).unwrap();
//! ```
//!
//! [MS-PPT]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt
//! [MS-ODRAW]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw
//! [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-cfb
//! [MS-OFFCRYPTO]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto

use std::io::{Cursor, Read};
use std::path::Path;

use cfb::CompoundFile;

use crate::error::{Error, Result};
use crate::ppt97::crypto::encrypt_ppt_stream;
use crate::ppt97::ole::{fix_mini_fat, write_stream};
use crate::ppt97::record::{
    parse_persist_directory, parse_record_header, read_u32_le, write_u32_le, RT_CURRENT_USER_ATOM,
    RT_USER_EDIT_ATOM,
};
use crate::ppt97::watermark::inject_watermark;

// 子模块声明：record / ole / watermark / crypto 为 pub，便于高级用户直接访问底层 API。
pub mod crypto;
pub mod ole;
pub mod record;
pub mod watermark;

// 重新导出最常用的类型，简化用户路径书写。
pub use watermark::WatermarkConfig;

// ============================================================================
// OLE2 stream 读写辅助
// ============================================================================

/// 从 OLE2 容器中读取所有 streams 内容。
///
/// 返回 `Vec<(stream_path, stream_data)>`，stream_path 以 `/` 开头
/// （`cfb` crate 约定）。
///
/// # 参数
/// - `comp`：OLE2 容器引用
///
/// # 返回
/// - 成功：`(path, data)` 列表
///
/// # 错误
/// - [`Error::Ppt97`]：stream 读取失败
fn read_all_streams(comp: &mut CompoundFile<Cursor<Vec<u8>>>) -> Result<Vec<(String, Vec<u8>)>> {
    // 先收集路径（避免 walk() 和 open_stream() 的借用冲突）
    let mut stream_paths: Vec<String> = Vec::new();
    for entry in comp.walk() {
        if entry.is_root() || entry.is_storage() {
            continue;
        }
        if entry.is_stream() {
            stream_paths.push(entry.path().to_string_lossy().to_string());
        }
    }

    let mut streams = Vec::with_capacity(stream_paths.len());
    for path in &stream_paths {
        let mut stream = comp
            .open_stream(path)
            .map_err(|e| Error::ppt97(format!("read_all_streams: open {} failed: {}", path, e)))?;
        let mut data = Vec::new();
        stream
            .read_to_end(&mut data)
            .map_err(|e| Error::ppt97(format!("read_all_streams: read {} failed: {}", path, e)))?;
        streams.push((path.clone(), data));
    }
    Ok(streams)
}

/// 将 streams 按 name 分类为 (Current User, PowerPoint Document, 其他)。
///
/// `cfb` crate 的路径以 `/` 开头，本函数去掉前导 `/` 后匹配 stream 名。
///
/// # 参数
/// - `streams`：原始 (path, data) 列表
///
/// # 返回
/// `(cu_data, ppt_data, other_streams)` 三元组，cu_data / ppt_data 为 None
/// 表示未找到对应 stream。
#[allow(clippy::type_complexity)]
fn classify_streams(
    streams: Vec<(String, Vec<u8>)>,
) -> (Option<Vec<u8>>, Option<Vec<u8>>, Vec<(String, Vec<u8>)>) {
    let mut cu_data: Option<Vec<u8>> = None;
    let mut ppt_data: Option<Vec<u8>> = None;
    let mut other_streams: Vec<(String, Vec<u8>)> = Vec::new();

    for (path, data) in streams {
        let name = path.strip_prefix('/').unwrap_or(&path);
        match name {
            "Current User" => cu_data = Some(data),
            "PowerPoint Document" => ppt_data = Some(data),
            _ => other_streams.push((path, data)),
        }
    }

    (cu_data, ppt_data, other_streams)
}

/// 将 streams 写回 OLE2 容器并落盘，返回最终容器字节。
///
/// 直接修改原始 OLE2 容器（保留原始目录结构、FAT 链、CLSID 等），而非
/// 创建新容器。原因：PowerPoint 严格检查 OLE2 容器结构，新建容器的
/// 目录条目顺序与树结构与原始文件不同会导致 PowerPoint 拒绝打开。
///
/// # 参数
/// - `comp`：可变 OLE2 容器
/// - `cu_data`：Current User stream 内容
/// - `ppt_data`：PowerPoint Document stream 内容
/// - `other_streams`：其他 streams 内容列表
///
/// # 返回
/// - 成功：OLE2 容器字节序列（已应用 `fix_mini_fat` 修复）
///
/// # 错误
/// - [`Error::Ppt97`]：stream 写入 / flush / fix_mini_fat 失败
fn write_back_streams(
    mut comp: CompoundFile<Cursor<Vec<u8>>>,
    cu_data: &[u8],
    ppt_data: &[u8],
    other_streams: &[(String, Vec<u8>)],
) -> Result<Vec<u8>> {
    write_stream(&mut comp, "Current User", cu_data)?;
    write_stream(&mut comp, "PowerPoint Document", ppt_data)?;
    for (path, data) in other_streams {
        write_stream(&mut comp, path, data)?;
    }
    comp.flush()
        .map_err(|e| Error::ppt97(format!("write_back_streams: flush failed: {}", e)))?;

    // 修复 cfb crate 多分配的 mini FAT 扇区（PowerPoint 严格检查 mini FAT 结构）
    // cfb::CompoundFile::into_inner 消耗 self 并返回底层存储 F（非 Result）
    let cursor = comp.into_inner();
    let mut data = cursor.into_inner();
    fix_mini_fat(&mut data)?;
    Ok(data)
}

// ============================================================================
// 公共 API
// ============================================================================

/// 为 .ppt 文件注入水印。
///
/// 完整流程：
/// 1. 读取 OLE2 容器中的所有 streams
/// 2. 解析 CurrentUser / UserEditAtom / PersistDirectoryAtom，获取 persist entries
/// 3. 在 PowerPoint Document stream 的所有 MainMaster 的 PPDrawing 中注入水印 SpContainer
/// 4. 更新所有受影响的 recLen 与 persist 对象 offset
/// 5. 写回 OLE2 容器，应用 mini FAT 修复
///
/// 水印作为 **MainMaster 母版的背景元素**注入到 PPDrawing 的 SpgrContainer 中，
/// 位于"组形状本身"之后（z-order 最低的真正子形状），符合业界水印常识：
/// - 全屏覆盖（ClientAnchor 0,0 → 5760,4320）
/// - 大字号、可配置旋转角度、中灰色（默认）
/// - 无填充、无边框、锁定不可编辑（FOPT 保护位 0x01C2）
/// - 普通视图下不可选中/编辑
///
/// # 参数
/// - `input_path`：输入 .ppt 文件路径
/// - `config`：水印配置（文本、字号、颜色、旋转角度）
///
/// # 返回
/// - 成功：注入水印后的 .ppt 文件字节序列
///
/// # 错误
/// - [`Error::Ppt97`]：找不到必要 stream / OLE2 容器结构损坏 / persist offset 更新失败
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::ppt97::{add_watermark, WatermarkConfig};
///
/// let config = WatermarkConfig {
///     text: "机密".to_string(),
///     ..Default::default()
/// };
/// let data = add_watermark(std::path::Path::new("input.ppt"), &config)?;
/// std::fs::write("watermarked.ppt", &data)?;
/// # Ok::<(), pptx_rs::Error>(())
/// ```
pub fn add_watermark(input_path: &Path, config: &WatermarkConfig) -> Result<Vec<u8>> {
    // 读取原始文件
    let file_data = std::fs::read(input_path)?;
    let cursor = Cursor::new(file_data);
    let mut comp = CompoundFile::open(cursor)
        .map_err(|e| Error::ppt97(format!("add_watermark: open OLE2 failed: {}", e)))?;

    let streams = read_all_streams(&mut comp)?;
    let (cu_data, ppt_data, other_streams) = classify_streams(streams);

    let mut ppt_data = ppt_data
        .ok_or_else(|| Error::ppt97("add_watermark: PowerPoint Document stream not found"))?;
    let mut cu_data =
        cu_data.ok_or_else(|| Error::ppt97("add_watermark: Current User stream not found"))?;

    // ========== 加水印前：读取旧的 offset 信息 ==========
    // 加水印会在 Slide 的 PPDrawing 中插入数据，导致：
    // 1. UserEditAtom 和 PersistDirectoryAtom 自身的位置后移
    // 2. PersistDirectoryAtom 中存储的 persist 对象 offset 过时
    // 因此需要先解析 persist entries，加水印后更新所有 offset。

    // 读取旧的 offsetToCurrentEdit（指向 UserEditAtom）
    let (_, _, cu_type, _) = parse_record_header(&cu_data, 0)?;
    if cu_type != RT_CURRENT_USER_ATOM {
        return Err(Error::ppt97(format!(
            "add_watermark: expected CurrentUserAtom (0x{:04X}), got 0x{:04X}",
            RT_CURRENT_USER_ATOM, cu_type
        )));
    }
    let ue_offset_old = read_u32_le(&cu_data, 16)? as usize;

    // 读取旧的 offsetPersistDirectory（从 UserEditAtom 中）
    let (_, _, ue_type, ue_len) = parse_record_header(&ppt_data, ue_offset_old)?;
    if ue_type != RT_USER_EDIT_ATOM {
        return Err(Error::ppt97(format!(
            "add_watermark: expected UserEditAtom (0x{:04X}), got 0x{:04X}",
            RT_USER_EDIT_ATOM, ue_type
        )));
    }
    if ue_len != 28 {
        return Err(Error::ppt97(format!(
            "add_watermark: file already encrypted or malformed (UserEditAtom.recLen={}, expected 28)",
            ue_len
        )));
    }
    let pd_offset_old = read_u32_le(&ppt_data, ue_offset_old + 20)? as usize;

    // 解析 PersistDirectoryAtom，获取 persist entries（旧 offset）
    let persist_entries = parse_persist_directory(&ppt_data, pd_offset_old)?;

    // ========== 加水印 ==========
    let (total_inserted, _new_entries, pd_offset_new) =
        inject_watermark(&mut ppt_data, config, &persist_entries, pd_offset_old)?;

    // 计算 UserEditAtom 的新 offset
    let ue_offset_new = ue_offset_old + total_inserted;

    // 更新 UserEditAtom 中的 offsetPersistDirectory 字段
    write_u32_le(&mut ppt_data, ue_offset_new + 20, pd_offset_new as u32)?;
    // 更新 CurrentUser 中的 offsetToCurrentEdit 字段
    write_u32_le(&mut cu_data, 16, ue_offset_new as u32)?;

    // 写回 OLE2 容器并修复 mini FAT
    write_back_streams(comp, &cu_data, &ppt_data, &other_streams)
}

/// 为 .ppt 文件设置 RC4 CryptoAPI 加密。
///
/// 完整流程详见 [`crypto::encrypt_ppt_stream`]。
///
/// # 参数
/// - `input_path`：输入 .ppt 文件路径
/// - `password`：明文密码
///
/// # 返回
/// - 成功：加密后的 .ppt 文件字节序列
///
/// # 错误
/// - [`Error::Ppt97`]：找不到必要 stream / 文件已加密 / 加密过程出错
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::ppt97::encrypt;
///
/// let data = encrypt(std::path::Path::new("input.ppt"), "my-password")?;
/// std::fs::write("encrypted.ppt", &data)?;
/// # Ok::<(), pptx_rs::Error>(())
/// ```
pub fn encrypt(input_path: &Path, password: &str) -> Result<Vec<u8>> {
    // 读取原始文件
    let file_data = std::fs::read(input_path)?;
    let cursor = Cursor::new(file_data);
    let mut comp = CompoundFile::open(cursor)
        .map_err(|e| Error::ppt97(format!("encrypt: open OLE2 failed: {}", e)))?;

    let streams = read_all_streams(&mut comp)?;
    let (cu_data, ppt_data, other_streams) = classify_streams(streams);

    let mut ppt_data =
        ppt_data.ok_or_else(|| Error::ppt97("encrypt: PowerPoint Document stream not found"))?;
    let mut cu_data =
        cu_data.ok_or_else(|| Error::ppt97("encrypt: Current User stream not found"))?;

    // 调用加密模块的核心流程
    encrypt_ppt_stream(&mut ppt_data, &mut cu_data, password)?;

    // 写回 OLE2 容器并修复 mini FAT
    write_back_streams(comp, &cu_data, &ppt_data, &other_streams)
}

/// 同时为 .ppt 文件注入水印和加密。
///
/// 完整流程：
/// 1. 先调用 [`add_watermark`] 的内部逻辑注入水印（更新 persist entries）
/// 2. 再调用 [`crypto::encrypt_ppt_stream`] 加密已加水印的 PowerPoint Document stream
///
/// 顺序很重要：必须**先加水印再加密**。因为加密后所有 persist 对象被加密，
/// 无法直接修改 record 结构；而水印注入需要修改 record 结构。
///
/// # 参数
/// - `input_path`：输入 .ppt 文件路径
/// - `config`：水印配置
/// - `password`：加密密码
///
/// # 返回
/// - 成功：注入水印且加密后的 .ppt 文件字节序列
///
/// # 错误
/// - [`Error::Ppt97`]：找不到必要 stream / 加水印失败 / 加密失败
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::ppt97::{add_watermark_and_encrypt, WatermarkConfig};
///
/// let config = WatermarkConfig::default();
/// let data = add_watermark_and_encrypt(
///     std::path::Path::new("input.ppt"),
///     &config,
///     "my-password",
/// )?;
/// std::fs::write("both.ppt", &data)?;
/// # Ok::<(), pptx_rs::Error>(())
/// ```
pub fn add_watermark_and_encrypt(
    input_path: &Path,
    config: &WatermarkConfig,
    password: &str,
) -> Result<Vec<u8>> {
    // 读取原始文件
    let file_data = std::fs::read(input_path)?;
    let cursor = Cursor::new(file_data);
    let mut comp = CompoundFile::open(cursor).map_err(|e| {
        Error::ppt97(format!(
            "add_watermark_and_encrypt: open OLE2 failed: {}",
            e
        ))
    })?;

    let streams = read_all_streams(&mut comp)?;
    let (cu_data, ppt_data, other_streams) = classify_streams(streams);

    let mut ppt_data = ppt_data.ok_or_else(|| {
        Error::ppt97("add_watermark_and_encrypt: PowerPoint Document stream not found")
    })?;
    let mut cu_data = cu_data
        .ok_or_else(|| Error::ppt97("add_watermark_and_encrypt: Current User stream not found"))?;

    // ========== 步骤 1：加水印 ==========
    // 读取旧的 offset 信息
    let (_, _, cu_type, _) = parse_record_header(&cu_data, 0)?;
    if cu_type != RT_CURRENT_USER_ATOM {
        return Err(Error::ppt97(format!(
            "add_watermark_and_encrypt: expected CurrentUserAtom (0x{:04X}), got 0x{:04X}",
            RT_CURRENT_USER_ATOM, cu_type
        )));
    }
    let ue_offset_old = read_u32_le(&cu_data, 16)? as usize;

    let (_, _, ue_type, ue_len) = parse_record_header(&ppt_data, ue_offset_old)?;
    if ue_type != RT_USER_EDIT_ATOM {
        return Err(Error::ppt97(format!(
            "add_watermark_and_encrypt: expected UserEditAtom (0x{:04X}), got 0x{:04X}",
            RT_USER_EDIT_ATOM, ue_type
        )));
    }
    if ue_len != 28 {
        return Err(Error::ppt97(format!(
            "add_watermark_and_encrypt: file already encrypted or malformed (UserEditAtom.recLen={}, expected 28)",
            ue_len
        )));
    }
    let pd_offset_old = read_u32_le(&ppt_data, ue_offset_old + 20)? as usize;

    let persist_entries = parse_persist_directory(&ppt_data, pd_offset_old)?;

    let (total_inserted, _new_entries, pd_offset_new) =
        inject_watermark(&mut ppt_data, config, &persist_entries, pd_offset_old)?;

    let ue_offset_new = ue_offset_old + total_inserted;
    write_u32_le(&mut ppt_data, ue_offset_new + 20, pd_offset_new as u32)?;
    write_u32_le(&mut cu_data, 16, ue_offset_new as u32)?;

    // ========== 步骤 2：加密 ==========
    encrypt_ppt_stream(&mut ppt_data, &mut cu_data, password)?;

    // 写回 OLE2 容器并修复 mini FAT
    write_back_streams(comp, &cu_data, &ppt_data, &other_streams)
}
