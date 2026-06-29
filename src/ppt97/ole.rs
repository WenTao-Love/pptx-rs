//! OLE2/CFB 容器操作工具。
//!
//! 本模块封装 [`cfb`] crate 的容器操作，提供 .ppt 文件（PowerPoint 97-2003
//! 二进制格式）所需的 OLE2（Compound File Binary）容器读写与修复能力。
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx 不支持 .ppt 二进制格式。本模块对标 LibreOffice 的
//! `oox::ole::OleStorage` 和 Apache POI 的 `org.apache.poi.poifs.filesystem`。
//!
//! # 规范依据
//!
//! - [MS-CFB]：Compound File Binary 文件格式规范
//!
//! # 主要功能
//!
//! - [`write_stream`]：写入 stream 到 OLE2 容器（替换已有 stream 内容）
//! - [`fix_mini_fat`]：修复 `cfb` crate 多分配的 mini FAT 扇区
//!
//! [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-cfb

use std::io::{Read, Seek, Write};

use cfb::CompoundFile;

use crate::error::{Error, Result};

// ============================================================================
// OLE2 header 字段偏移（MS-CFB 规范）
// ============================================================================

/// sector_shift 字段偏移（u16）。
const HDR_SECTOR_SIZE_OFF: usize = 30;

/// first_mini_fat_sector 字段偏移（u32）。
const HDR_FIRST_MINI_FAT_OFF: usize = 60;

/// num_mini_fat_sectors 字段偏移（u32）。
const HDR_NUM_MINI_FAT_OFF: usize = 64;

/// DIFAT 数组起始偏移（header 内嵌 109 个 u32 条目）。
const HDR_DIFAT_OFF: usize = 76;

// ============================================================================
// FAT 特殊值（MS-CFB 规范 2.2）
// ============================================================================

/// 最大常规扇区号（大于等于此值的为特殊标记）。
const MAXREGSECT: u32 = 0xFFFFFFFA;

/// 空闲扇区标记。
const FREESECT: u32 = 0xFFFFFFFF;

/// 链结束标记。
const ENDOFCHAIN: u32 = 0xFFFFFFFE;

/// 写入一个 stream 到 OLE2 容器（替换已有 stream 内容）。
///
/// `cfb` crate 的 `create_stream` 会先删除同名 stream 再创建新 stream，
/// 因此本函数等价于"覆盖写入"。写入后会立即 flush，确保数据落盘。
///
/// # 类型参数
/// - `F`：底层可读写可定位的存储类型（如 `std::fs::File` / `std::io::Cursor`）
///
/// # 参数
/// - `comp`：可变的 OLE2 容器
/// - `path`：stream 路径（如 `"PowerPoint Document"` / `"Current User"`）
/// - `data`：待写入的字节内容
///
/// # 错误
/// - [`Error::Ppt97`]：stream 创建 / 写入 / flush 失败
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::ppt97::ole::write_stream;
/// use cfb::CompoundFile;
/// use std::fs::File;
///
/// let mut comp = CompoundFile::open(File::open("input.ppt")?)?;
/// let data = b"new stream content";
/// write_stream(&mut comp, "PowerPoint Document", data)?;
/// comp.flush()?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn write_stream<F: Read + Write + Seek>(
    comp: &mut CompoundFile<F>,
    path: &str,
    data: &[u8],
) -> Result<()> {
    let mut stream = comp.create_stream(path).map_err(|e| {
        Error::ppt97(format!(
            "write_stream: create_stream {} failed: {}",
            path, e
        ))
    })?;
    stream
        .write_all(data)
        .map_err(|e| Error::ppt97(format!("write_stream: write {} failed: {}", path, e)))?;
    stream
        .flush()
        .map_err(|e| Error::ppt97(format!("write_stream: flush {} failed: {}", path, e)))?;
    drop(stream);
    Ok(())
}

/// 修复 `cfb` crate 多分配的 mini FAT 扇区。
///
/// # 背景
///
/// `cfb` crate 的 `create_stream` 在某些情况下会多分配一个 mini FAT 扇区，
/// 但这些多余的扇区所有条目都是 `FREESECT`（未使用）。PowerPoint / WPS 严格
/// 检查 mini FAT 结构，拒绝打开有多余 mini FAT 扇区的文件（表现为：密码验证
/// 通过后立即报错，文件无法打开）。
///
/// # 修复逻辑
///
/// 1. 读取 OLE2 header 中的 `first_mini_fat_sector` 和 `num_mini_fat_sectors`
/// 2. 读取 DIFAT 和 FAT
/// 3. 沿 FAT 链遍历 mini FAT 扇区
/// 4. 如果 mini FAT 链长度 > 1，且第二个及之后的扇区所有条目都是 `FREESECT`：
///    - 将 FAT 中第一个 mini FAT 扇区的 next 指针改为 `ENDOFCHAIN`
///    - 将 FAT 中第二个及之后的 mini FAT 扇区标记为 `FREESECT`
///    - 修改 OLE2 header 中的 `num_mini_fat_sectors` 为实际使用的扇区数
///
/// # 参数
/// - `data`：OLE2 容器数据（原地修改）
///
/// # 错误
/// - [`Error::Ppt97`]：数据过短 / FAT 扇区超出范围 / 无法定位 FAT 条目
pub fn fix_mini_fat(#[allow(clippy::ptr_arg)] data: &mut Vec<u8>) -> Result<()> {
    if data.len() < 512 {
        return Err(Error::ppt97(
            "fix_mini_fat: data too short (less than 512 bytes, not a valid OLE2 container)"
                .to_string(),
        ));
    }

    // 读取 sector_size（2^sector_shift）
    let sector_shift =
        u16::from_le_bytes([data[HDR_SECTOR_SIZE_OFF], data[HDR_SECTOR_SIZE_OFF + 1]]);
    let sector_size = 1usize << sector_shift;
    let entries_per_sector = sector_size / 4; // 每个 sector 容纳的 FAT 条目数

    let first_mini_fat_sector = u32::from_le_bytes([
        data[HDR_FIRST_MINI_FAT_OFF],
        data[HDR_FIRST_MINI_FAT_OFF + 1],
        data[HDR_FIRST_MINI_FAT_OFF + 2],
        data[HDR_FIRST_MINI_FAT_OFF + 3],
    ]);
    let num_mini_fat_sectors = u32::from_le_bytes([
        data[HDR_NUM_MINI_FAT_OFF],
        data[HDR_NUM_MINI_FAT_OFF + 1],
        data[HDR_NUM_MINI_FAT_OFF + 2],
        data[HDR_NUM_MINI_FAT_OFF + 3],
    ]);

    // 只在 num_mini_fat_sectors > 1 时才需要修复
    if num_mini_fat_sectors <= 1 || first_mini_fat_sector >= MAXREGSECT {
        return Ok(());
    }

    // 读取 DIFAT（前 109 个条目在 header 中）
    let mut difat: Vec<u32> = Vec::new();
    for i in 0..109 {
        let off = HDR_DIFAT_OFF + i * 4;
        let val = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
        if val != FREESECT {
            difat.push(val);
        }
    }

    // 读取 FAT
    let mut fat: Vec<u32> = Vec::new();
    for sec in &difat {
        let offset = 512 + (*sec as usize) * sector_size;
        if offset + sector_size > data.len() {
            return Err(Error::ppt97(format!(
                "fix_mini_fat: FAT sector {} out of range (offset {}, data len {})",
                sec,
                offset,
                data.len()
            )));
        }
        for j in 0..entries_per_sector {
            let off = offset + j * 4;
            let val = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
            fat.push(val);
        }
    }

    // 沿 FAT 链遍历 mini FAT 扇区
    let mut mini_fat_chain: Vec<u32> = Vec::new();
    let mut sec = first_mini_fat_sector;
    while sec != ENDOFCHAIN && sec < MAXREGSECT && (sec as usize) < fat.len() {
        mini_fat_chain.push(sec);
        let next = fat[sec as usize];
        if next == ENDOFCHAIN {
            break;
        }
        sec = next;
    }

    if mini_fat_chain.len() <= 1 {
        return Ok(()); // 链长度 <= 1，无需修复
    }

    // 检查第二个及之后的 mini FAT 扇区是否都是 FREESECT
    let mut all_free = true;
    for &s in &mini_fat_chain[1..] {
        let offset = 512 + (s as usize) * sector_size;
        if offset + sector_size > data.len() {
            all_free = false;
            break;
        }
        for j in 0..entries_per_sector {
            let off = offset + j * 4;
            let val = u32::from_le_bytes([data[off], data[off + 1], data[off + 2], data[off + 3]]);
            if val != FREESECT {
                all_free = false;
                break;
            }
        }
        if !all_free {
            break;
        }
    }

    if !all_free {
        return Ok(()); // 第二个扇区有非 FREESECT 条目，不修复
    }

    // 修复步骤 1：将 FAT 中第一个 mini FAT 扇区的 next 指针改为 ENDOFCHAIN
    let first_sec = mini_fat_chain[0];
    let fat_sector_idx = (first_sec as usize) / entries_per_sector;
    let entry_in_sector = (first_sec as usize) % entries_per_sector;
    if fat_sector_idx >= difat.len() {
        return Err(Error::ppt97(format!(
            "fix_mini_fat: cannot locate first mini FAT sector {} in DIFAT",
            first_sec
        )));
    }
    let fat_sector = difat[fat_sector_idx];
    let fat_entry_offset = 512 + (fat_sector as usize) * sector_size + entry_in_sector * 4;
    if fat_entry_offset + 4 > data.len() {
        return Err(Error::ppt97(format!(
            "fix_mini_fat: FAT entry offset {} out of range",
            fat_entry_offset
        )));
    }
    data[fat_entry_offset..fat_entry_offset + 4].copy_from_slice(&ENDOFCHAIN.to_le_bytes());

    // 修复步骤 2：将 FAT 中第二个及之后的 mini FAT 扇区标记为 FREESECT
    for &s in &mini_fat_chain[1..] {
        let fat_sector_idx = (s as usize) / entries_per_sector;
        let entry_in_sector = (s as usize) % entries_per_sector;
        if fat_sector_idx >= difat.len() {
            continue;
        }
        let fat_sector = difat[fat_sector_idx];
        let fat_entry_offset = 512 + (fat_sector as usize) * sector_size + entry_in_sector * 4;
        if fat_entry_offset + 4 > data.len() {
            continue;
        }
        data[fat_entry_offset..fat_entry_offset + 4].copy_from_slice(&FREESECT.to_le_bytes());
    }

    // 修复步骤 3：修改 OLE2 header 中的 num_mini_fat_sectors 为 1
    data[HDR_NUM_MINI_FAT_OFF..HDR_NUM_MINI_FAT_OFF + 4].copy_from_slice(&1u32.to_le_bytes());

    Ok(())
}
