//! PPT 97-2003 二进制 record 树解析工具。
//!
//! 本模块提供 .ppt 文件（PowerPoint 97-2003 二进制格式）底层 record 树的
//! 解析能力，是 [`crate::ppt97`] 模块的水印注入与加密功能的基础设施。
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx **不支持** .ppt 97-2003 二进制格式（仅支持 .pptx）。
//! 本模块填补了这一空白，对标 LibreOffice 的 `oox::ole` 和 Apache POI 的
//! `org.apache.poi.hslf`（HSLF = Horrible Slide Layout Format）。
//!
//! # 规范依据
//!
//! - [MS-PPT]：PowerPoint 97-2003 二进制文件格式规范
//! - [MS-ODRAW]：Office Drawing 97-2003 二进制格式规范（Escher OfficeArt）
//! - [MS-CFB]：Compound File Binary 文件格式规范（OLE2 容器）
//!
//! # record 树结构
//!
//! .ppt 文件的 `PowerPoint Document` stream 是一棵 record 树：
//!
//! ```text
//! record header (8 bytes): verInst(u16) + recType(u16) + recLen(u32)
//! record data (recLen bytes)
//! ```
//!
//! - `ver`（4 bit）：版本号，`0xF` 表示 container（含子 record）
//! - `inst`（12 bit）：实例号，语义随 recType 变化
//! - `recType`（u16）：record 类型码（如 0x03F8 = MainMaster）
//! - `recLen`（u32）：data 字段字节数（不含 header 的 8 字节）
//!
//! [MS-PPT]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt
//! [MS-ODRAW]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw
//! [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-cfb

use crate::error::{Error, Result};

// ============================================================================
// Record 类型常量（MS-PPT 规范）
// ============================================================================

/// MainMaster record type（MS-PPT 规范：RT_MainMaster = 0x03F8）。
///
/// 母版是幻灯片的背景模板，水印注入目标即为此 record 的 PPDrawing 子树。
/// 一个 .ppt 文件可包含多个 MainMaster（每个对应一套版式）。
pub const RT_MAIN_MASTER: u16 = 0x03F8;

/// PPDrawing record type（MS-PPT 规范：RT_PPDrawing = 0x040C）。
///
/// PPDrawing 是 MainMaster / Slide 的子 record，承载 Escher OfficeArt
/// 绘图树（DgContainer → SpgrContainer → SpContainer）。
pub const RT_PPDRAWING: u16 = 0x040C;

/// CurrentUserAtom record type（MS-PPT 规范：RT_CurrentUserAtom = 0x0FF6）。
///
/// 位于 `Current User` stream，记录 offsetToCurrentEdit（指向最新 UserEditAtom）。
pub const RT_CURRENT_USER_ATOM: u16 = 0x0FF6;

/// UserEditAtom record type（MS-PPT 规范：RT_UserEditAtom = 0x0FF5）。
///
/// 记录用户最后一次编辑的元信息：offsetPersistDirectory、maxPersistWritten、
/// encryptSessionPersistIdRef（加密时存在，recLen 由 28 变为 32）。
pub const RT_USER_EDIT_ATOM: u16 = 0x0FF5;

/// PersistDirectoryAtom record type（MS-PPT 规范：RT_PersistDirectoryAtom = 0x1772）。
///
/// persist 目录：记录每个 persistId 到 stream offset 的映射，
/// 是 .ppt 文件"对象寻址"的核心数据结构。
pub const RT_PERSIST_DIRECTORY_ATOM: u16 = 0x1772;

/// 解析 8 字节 record header。
///
/// .ppt record header 结构（MS-PPT 规范）：
///
/// ```text
/// offset  长度  字段      说明
/// 0       2     verInst   ver(低4位) | inst(高12位)
/// 2       2     recType   record 类型码
/// 4       4     recLen    data 字段长度（不含 header 自身）
/// ```
///
/// # 参数
/// - `data`：PowerPoint Document stream 字节
/// - `offset`：record header 起始偏移
///
/// # 返回
/// - 成功：`(ver, inst, rec_type, rec_len)` 四元组
///   - `ver`：低 4 位版本号，`0xF` 表示 container
///   - `inst`：高 12 位实例号
///   - `rec_type`：record 类型码
///   - `rec_len`：data 字段长度
///
/// # 错误
/// - [`Error::Ppt97`]：`offset + 8` 超出 `data` 范围。
pub fn parse_record_header(data: &[u8], offset: usize) -> Result<(u8, u16, u16, u32)> {
    if offset + 8 > data.len() {
        return Err(Error::ppt97(format!(
            "record header parse: offset {} + 8 out of range (data len {})",
            offset,
            data.len()
        )));
    }
    let ver_inst = u16::from_le_bytes([data[offset], data[offset + 1]]);
    let rec_type = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
    let rec_len = u32::from_le_bytes([
        data[offset + 4],
        data[offset + 5],
        data[offset + 6],
        data[offset + 7],
    ]);
    let ver = (ver_inst & 0x0F) as u8;
    let inst = (ver_inst >> 4) & 0x0FFF;
    Ok((ver, inst, rec_type, rec_len))
}

/// 读取小端 u32。
///
/// # 参数
/// - `data`：数据源
/// - `offset`：读取起始位置
///
/// # 返回
/// 小端 u32 值。
///
/// # 错误
/// - [`Error::Ppt97`]：`offset + 4` 超出 `data` 范围。
pub fn read_u32_le(data: &[u8], offset: usize) -> Result<u32> {
    if offset + 4 > data.len() {
        return Err(Error::ppt97(format!(
            "read_u32_le: offset {} + 4 out of range (data len {})",
            offset,
            data.len()
        )));
    }
    Ok(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

/// 写入小端 u32（原地修改）。
///
/// # 参数
/// - `data`：可变数据源
/// - `offset`：写入起始位置
/// - `val`：待写入的 u32 值
///
/// # 错误
/// - [`Error::Ppt97`]：`offset + 4` 超出 `data` 范围。
pub fn write_u32_le(data: &mut [u8], offset: usize, val: u32) -> Result<()> {
    if offset + 4 > data.len() {
        return Err(Error::ppt97(format!(
            "write_u32_le: offset {} + 4 out of range (data len {})",
            offset,
            data.len()
        )));
    }
    let bytes = val.to_le_bytes();
    data[offset..offset + 4].copy_from_slice(&bytes);
    Ok(())
}

/// 解析 PersistDirectoryAtom，返回 persist entries 列表。
///
/// PersistDirectoryAtom 结构（MS-PPT 规范 2.3.3）：
///
/// ```text
/// record header (8 bytes): type=0x1772, len=可变
/// rgPersistDirEntry[]:
///   PersistDirectoryEntry:
///     - persistId (20 bits) + cPersist (12 bits) = 4 bytes
///     - rgPersistOffset[cPersist] (cPersist * 4 bytes): 每个 entry 的 stream offset
/// ```
///
/// 一个 PersistDirectoryAtom 可包含**多个** PersistDirectoryEntry，每个 entry
/// 有自己的 persistId 起始值和 cPersist。persistId 在 entry 内从 `persistId`
/// 开始递增（`persistId + 0`、`persistId + 1`、...、`persistId + cPersist - 1`）。
///
/// # 参数
/// - `data`：PowerPoint Document stream 字节
/// - `offset`：PersistDirectoryAtom 的起始偏移
///
/// # 返回
/// - 成功：`Vec<(persistId, stream_offset)>` 列表
///
/// # 错误
/// - [`Error::Ppt97`]：record type 不匹配 / 数据越界
pub fn parse_persist_directory(data: &[u8], offset: usize) -> Result<Vec<(u32, u32)>> {
    let (_, _, rec_type, rec_len) = parse_record_header(data, offset)?;
    if rec_type != RT_PERSIST_DIRECTORY_ATOM {
        return Err(Error::ppt97(format!(
            "persist directory parse: expected type 0x{:04X}, got 0x{:04X}",
            RT_PERSIST_DIRECTORY_ATOM, rec_type
        )));
    }

    // PersistDirectoryAtom data 范围：header(8) 之后 recLen 字节
    let pd_start = offset + 8;
    let pd_end = pd_start + rec_len as usize;
    if pd_end > data.len() {
        return Err(Error::ppt97(format!(
            "persist directory parse: data out of range (start {}, end {}, data len {})",
            pd_start,
            pd_end,
            data.len()
        )));
    }

    let pd_data = &data[pd_start..pd_end];
    let mut entries = Vec::new();
    let mut pos = 0usize;

    // 遍历 rgPersistDirEntry[]，每个 entry 由 4 字节头 + cPersist 个 4 字节 offset 组成
    while pos + 4 <= pd_data.len() {
        let entry = u32::from_le_bytes([
            pd_data[pos],
            pd_data[pos + 1],
            pd_data[pos + 2],
            pd_data[pos + 3],
        ]);
        // persistId: 20 bits, cPersist: 12 bits（MS-PPT 规范）
        let persist_id = entry & 0xFFFFF;
        let c_persist = (entry >> 20) & 0xFFF;
        pos += 4;

        for j in 0..c_persist {
            if pos + 4 > pd_data.len() {
                return Err(Error::ppt97(format!(
                    "persist directory parse: entry (pid={}, cPersist={}) offset out of range",
                    persist_id, c_persist
                )));
            }
            let persist_offset = u32::from_le_bytes([
                pd_data[pos],
                pd_data[pos + 1],
                pd_data[pos + 2],
                pd_data[pos + 3],
            ]);
            entries.push((persist_id + j, persist_offset));
            pos += 4;
        }
    }

    Ok(entries)
}

/// 在 PowerPoint Document stream 中找到所有 MainMaster record 的 offset。
///
/// 遍历顶层 record，收集所有 `type=0x03F8 (MainMaster)` 的 container record。
///
/// 水印注入目标即为此函数返回的每个 MainMaster 的 PPDrawing 子树。
/// 一个 .ppt 文件可包含多个 MainMaster（每个对应一套版式）。
///
/// # 参数
/// - `data`：PowerPoint Document stream 字节
///
/// # 返回
/// - 成功：所有 MainMaster 的起始 offset 列表（按文件中出现顺序）
///
/// # 错误
/// - [`Error::Ppt97`]：record header 解析失败 / record 越界
pub fn find_main_masters(data: &[u8]) -> Result<Vec<usize>> {
    let mut masters = Vec::new();
    let mut pos = 0usize;
    while pos + 8 <= data.len() {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8usize + rec_len as usize;

        if is_container && rec_type == RT_MAIN_MASTER {
            masters.push(pos);
        }

        // 防御性检查：rec_len 异常时停止遍历，避免死循环
        if total_len == 0 {
            break;
        }
        pos += total_len;
        // 非 container 且 rec_len=0 是终止标记（MS-PPT 规范）
        if !is_container && rec_len == 0 {
            break;
        }
    }
    Ok(masters)
}

/// 在 MainMaster container 中找到 PPDrawing 的 offset。
///
/// MainMaster (container) 的子 record 包括 SlideAtom、Environment、PPDrawing 等。
/// 遍历 MainMaster 的子 record，找到 `type=0x040C (PPDrawing)` 的 container。
///
/// # 参数
/// - `data`：PowerPoint Document stream 字节
/// - `master_offset`：MainMaster record 的起始偏移
///
/// # 返回
/// - 成功：`Some(ppd_offset)` 找到 PPDrawing；`None` 未找到
///
/// # 错误
/// - [`Error::Ppt97`]：record header 解析失败 / MainMaster 越界
pub fn find_ppdrawing_in_master(data: &[u8], master_offset: usize) -> Result<Option<usize>> {
    let (_, _, _, master_len) = parse_record_header(data, master_offset)?;
    let master_end = master_offset + 8 + master_len as usize;

    let mut pos = master_offset + 8;
    while pos + 8 <= master_end {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8usize + rec_len as usize;

        if is_container && rec_type == RT_PPDRAWING {
            return Ok(Some(pos));
        }

        if total_len == 0 {
            break;
        }
        pos += total_len;
        if !is_container && rec_len == 0 {
            break;
        }
    }
    Ok(None)
}
