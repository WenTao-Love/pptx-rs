//! .ppt 文件水印注入。
//!
//! 本模块为 PowerPoint 97-2003 二进制格式（`.ppt`）文件注入水印。
//! 水印作为 **MainMaster 母版的背景元素**注入到 PPDrawing 的 SpgrContainer 中，
//! 位于"组形状本身"之后（z-order 最低的真正子形状），实现真正水印的视觉效果：
//!
//! - 全屏覆盖（ClientAnchor 0,0 → 5760,4320）
//! - 大字号、可配置旋转角度、中灰色（默认）
//! - 无填充、无边框、锁定不可编辑（FOPT 保护位 0x01C2）
//! - 普通视图下不可选中/编辑（符合业界水印常识）
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx 不支持 .ppt 二进制格式。本模块对标 Aspose.Slides 的
//! `ISlideMaster.add_watermark` 和 LibreOffice 的母版背景元素注入。
//!
//! # 规范依据
//!
//! - [MS-ODRAW]：Office Drawing 97-2003 二进制格式（Escher OfficeArt）
//! - [MS-PPT] 2.3.4：PPDrawing / OfficeArtClientTextbox 等
//!
//! # 水印 SpContainer 结构
//!
//! ```text
//! SpContainer (0xF004, container)
//! ├── FSP (0xF00A): 形状属性，inst=0xCA (TextBox)
//! ├── FOPT (0xF00B): 形状选项（无填充、无线条、锁定、旋转）
//! ├── ClientAnchor (0xF010): 形状锚点（全屏覆盖）
//! └── ClientTextbox (0xF00D, container): 文本框
//!     ├── TextHeaderAtom (0x0F9F): 文本类型
//!     ├── TextCharsAtom (0x0FA0): 水印文本（UTF-16LE）
//!     └── StyleTextPropAtom (0x0FA1): 文本样式（字号、颜色）
//! ```
//!
//! [MS-ODRAW]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw
//! [MS-PPT]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt

use crate::error::{Error, Result};
use crate::ppt97::record::{
    find_main_masters, find_ppdrawing_in_master, parse_record_header, read_u32_le, write_u32_le,
    RT_MAIN_MASTER,
};

// ============================================================================
// OfficeArt record type（MS-ODRAW 规范）
// ============================================================================

/// DgContainer record type（Drawing Container）。
const RT_DG_CONTAINER: u16 = 0xF002;

/// SpgrContainer record type（Shape Group Container）。
pub const RT_SPGR_CONTAINER: u16 = 0xF003;

/// SpContainer record type（Shape Container）。
pub const RT_SP_CONTAINER: u16 = 0xF004;

/// FSP record type（File Shape Properties，ver=2，inst=MSOSPT 形状类型）。
pub const RT_FSP: u16 = 0xF00A;

/// FOPT record type（File Option，ver=3，inst=属性个数）。
pub const RT_FOPT: u16 = 0xF00B;

/// ClientTextbox record type（container 版本，ver=0xF）。
pub const RT_CLIENT_TEXTBOX: u16 = 0xF00D;

/// ClientAnchor record type（形状锚点，ver=0，inst=0）。
pub const RT_CLIENT_ANCHOR: u16 = 0xF010;

// ============================================================================
// PPT Text record types
// ============================================================================

/// TextHeaderAtom record type。
const RT_TEXT_HEADER_ATOM: u16 = 0x0F9F;

/// TextCharsAtom record type（UTF-16LE 文本）。
const RT_TEXT_CHARS_ATOM: u16 = 0x0FA0;

/// StyleTextPropAtom record type（文本样式属性）。
const RT_STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;

// ============================================================================
// 水印配置
// ============================================================================

/// 水印配置参数。
///
/// 参考 python-pptx 的 `shapes.add_textbox()` + `font` 组合 API 和
/// Aspose.Slides 的 `PortionFormat` 参数化模式，将水印的可变属性集中管理，
/// 避免硬编码散落在 [`build_watermark_spcontainer`] 各处。
///
/// # 字段对应关系
///
/// - `text` ↔ python-pptx `text_frame.text` / Aspose `add_text_frame(text)`
/// - `font_size_pt` ↔ python-pptx `font.size = Pt(n)` / Aspose `font_height`
/// - `color_rgb` ↔ python-pptx `font.color.rgb` / Aspose `solid_fill_color`
/// - `rotation_deg` ↔ python-pptx `shape.rotation` / Aspose `shape.rotation`
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::ppt97::watermark::WatermarkConfig;
///
/// let config = WatermarkConfig {
///     text: "机密".to_string(),
///     font_size_pt: 72,
///     color_rgb: (180, 180, 180),
///     rotation_deg: -30,
/// };
/// ```
#[derive(Clone, Debug)]
pub struct WatermarkConfig {
    /// 水印文本内容。
    pub text: String,
    /// 字号（磅），常见水印字号 44pt。
    pub font_size_pt: u16,
    /// 文字颜色 (r, g, b)，每个分量 0-255。
    pub color_rgb: (u8, u8, u8),
    /// 旋转角度（度），正值顺时针，负值逆时针。
    pub rotation_deg: i32,
}

impl Default for WatermarkConfig {
    /// 默认水印配置：44pt 中灰色 "pptx-rs 水印"，45 度旋转。
    fn default() -> Self {
        Self {
            text: "pptx-rs 水印".to_string(),
            font_size_pt: 44,
            color_rgb: (200, 200, 200),
            rotation_deg: 45,
        }
    }
}

// ============================================================================
// 水印注入主逻辑
// ============================================================================

/// 在 PowerPoint Document stream 中注入水印。
///
/// 完整流程：
/// 1. 找到所有 MainMaster record
/// 2. 对每个 MainMaster 的 PPDrawing：
///    - 幂等性检查（已存在水印文本则跳过）
///    - 分配唯一 shapeId（避免与已有形状冲突）
///    - 构造水印 SpContainer
///    - 插入到 SpgrContainer 的"组形状本身"之后（z-order 最低的真正形状）
///    - 更新 MainMaster / PPDrawing / DgContainer / SpgrContainer 的 recLen
/// 3. 重新计算所有 persist 对象的 offset（因插入导致后移）
/// 4. 更新 PersistDirectoryAtom 中存储的 offset
///
/// # 参数
/// - `ppt_data`：PowerPoint Document stream（原地修改）
/// - `config`：水印配置
/// - `persist_entries`：旧的 persist entries（注入前的 offset）
/// - `pd_offset_old`：旧的 PersistDirectoryAtom offset
///
/// # 返回
/// - 成功：`(total_inserted, new_entries, pd_offset_new)`
///   - `total_inserted`：总共插入的字节数
///   - `new_entries`：更新后的 persist entries
///   - `pd_offset_new`：新的 PersistDirectoryAtom offset
///
/// # 错误
/// - [`Error::Ppt97`]：record 解析失败 / PPDrawing 结构异常 / shapeId 扫描失败
#[allow(clippy::type_complexity)]
pub fn inject_watermark(
    ppt_data: &mut Vec<u8>,
    config: &WatermarkConfig,
    persist_entries: &[(u32, u32)],
    pd_offset_old: usize,
) -> Result<(usize, Vec<(u32, u32)>, usize)> {
    // 返回类型 (total_inserted, new_entries, pd_offset_new) 较复杂，但语义清晰，
    // 拆分为多个返回值会比封装成结构体更直观（调用方需要分别使用这三个值）。
    let master_offsets = find_main_masters(ppt_data)?;
    let mut insertions: Vec<(usize, usize)> = Vec::new();
    let mut total_inserted = 0;

    // 预计算水印文本的 UTF-16LE 字节序列，用于幂等性检测
    let watermark_utf16: Vec<u8> = config
        .text
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // 扫描所有 MainMaster 中已有的最大 shapeId，避免 ID 冲突
    // 参考 Apache POI XSLFSheet.allocateShapeId 的 BitSet + nextClearBit 模式
    let mut next_shape_id: u32 = find_max_shape_id(ppt_data, &master_offsets)? + 1;
    // 确保 shapeId 不低于常见起始值 0x1000（与 PowerPoint 内部约定一致）
    if next_shape_id < 0x1000 {
        next_shape_id = 0x1000;
    }

    // 从后往前处理 MainMaster，记录所有插入点
    for master_offset in master_offsets.iter().rev() {
        let master_offset = *master_offset;
        if let Some(ppd_offset) = find_ppdrawing_in_master(ppt_data, master_offset)? {
            // 幂等性检查：扫描 PPDrawing 中是否已存在水印文本
            if has_watermark_text(ppt_data, ppd_offset, &watermark_utf16) {
                // 已存在水印，跳过注入（幂等性保证）
                continue;
            }

            let shape_id = next_shape_id;
            next_shape_id += 1;
            let watermark_sp = build_watermark_spcontainer(shape_id, config);

            let (insert_pos, insert_len) =
                inject_watermark_into_ppdrawing(ppt_data, ppd_offset, &watermark_sp)?;

            // 更新 MainMaster 的 recLen
            let (_, _, _, master_len) = parse_record_header(ppt_data, master_offset)?;
            write_u32_le(ppt_data, master_offset + 4, master_len + insert_len as u32)?;

            insertions.push((insert_pos, insert_len));
            total_inserted += insert_len;
        }
    }

    // 计算每个 persist 对象的新 offset
    // 对于旧 offset O，新 offset = O + sum(len for all (pos, len) where pos < O)
    // 因为插入点 pos < O 时，插入发生在 O 之前，O 需要后移
    let mut new_entries = Vec::with_capacity(persist_entries.len());
    for (pid, offset) in persist_entries {
        let mut new_offset = *offset;
        for (pos, len) in &insertions {
            if *pos < new_offset as usize {
                new_offset += *len as u32;
            }
        }
        new_entries.push((*pid, new_offset));
    }

    // 计算 PersistDirectoryAtom 的新 offset
    let mut pd_offset_new = pd_offset_old;
    for (pos, len) in &insertions {
        if *pos < pd_offset_new {
            pd_offset_new += len;
        }
    }

    // 更新 PersistDirectoryAtom 中存储的 offset
    // PersistDirectoryAtom 结构：header(8) + entry(4) + rgPersistOffset(cPersist * 4)
    let pd_data_start = pd_offset_new + 8;

    // 更新每个 persist offset
    for (i, (_, new_offset)) in new_entries.iter().enumerate() {
        let offset_pos = pd_data_start + 4 + i * 4;
        write_u32_le(ppt_data, offset_pos, *new_offset)?;
    }

    Ok((total_inserted, new_entries, pd_offset_new))
}

/// 在 PPDrawing 中注入水印 SpContainer。
///
/// PPDrawing 的完整结构（MS-ODRAW 规范）：
///
/// ```text
/// PPDrawing (container, 0x040C)
///   └── DgContainer (container, 0xF002)
///        └── SpgrContainer (container, 0xF003)
///             ├── SpContainer (0xF004) — 组形状本身（FSP.inst=0, MSOSPT_Min）
///             ├── SpContainer (0xF004) — 其他形状（z-order 从低到高）
///             └── ...
/// ```
///
/// **关键设计**：水印插入到 SpgrContainer 的"组形状本身"之后，成为 z-order
/// 最低的真正子形状（背景层）。这正是 PowerPoint 自带"插入水印"功能的实现方式：
///
/// - 视觉上被其他内容覆盖，符合水印背景特性
/// - 配合 FOPT 锁定属性（0x01C2=0x0D），在普通视图下不可选中/编辑
///
/// # 参数
/// - `data`：PowerPoint Document stream（原地修改）
/// - `ppd_offset`：PPDrawing record 的起始 offset
/// - `watermark_sp`：构造好的水印 SpContainer 字节序列
///
/// # 返回
/// - 成功：`(insert_pos, insert_len)` 插入位置与字节数
///
/// # 错误
/// - [`Error::Ppt97`]：PPDrawing 中找不到 DgContainer / SpgrContainer / 组形状本身
pub fn inject_watermark_into_ppdrawing(
    data: &mut Vec<u8>,
    ppd_offset: usize,
    watermark_sp: &[u8],
) -> Result<(usize, usize)> {
    let (_, _, _, ppd_len) = parse_record_header(data, ppd_offset)?;

    // 第 1 层：在 PPDrawing 中找到 DgContainer (0xF002)
    let ppd_end = ppd_offset + 8 + ppd_len as usize;
    let mut pos = ppd_offset + 8;
    let mut dg_offset = None;
    while pos + 8 <= ppd_end {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8 + rec_len as usize;

        if is_container && rec_type == RT_DG_CONTAINER {
            dg_offset = Some(pos);
            break;
        }

        pos += total_len;
        if !is_container && rec_len == 0 {
            break;
        }
    }

    let dg_offset = dg_offset.ok_or_else(|| {
        Error::ppt97("inject_watermark: DgContainer (0xF002) not found in PPDrawing")
    })?;
    let (_, _, _, dg_len) = parse_record_header(data, dg_offset)?;

    // 第 2 层：在 DgContainer 中找到 SpgrContainer (0xF003)
    let dg_end = dg_offset + 8 + dg_len as usize;
    let mut pos = dg_offset + 8;
    let mut spgr_offset = None;
    while pos + 8 <= dg_end {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8 + rec_len as usize;

        if is_container && rec_type == RT_SPGR_CONTAINER {
            spgr_offset = Some(pos);
            break;
        }

        pos += total_len;
        if !is_container && rec_len == 0 {
            break;
        }
    }

    let spgr_offset = spgr_offset.ok_or_else(|| {
        Error::ppt97("inject_watermark: SpgrContainer (0xF003) not found in DgContainer")
    })?;
    let (_, _, _, spgr_len) = parse_record_header(data, spgr_offset)?;
    let spgr_end = spgr_offset + 8 + spgr_len as usize;

    // 关键设计：把水印插入到 SpgrContainer 的"组形状本身"之后，
    // 成为 z-order 最低的真正子形状（背景层）。
    //
    // MS-ODRAW 规范 2.2.17：SpgrContainer 的第一个 SpContainer 必须是"组形状本身"
    // （FSP.inst=0, MSOSPT_Min=0），描述整个组的属性；后续 SpContainer 才是
    // 真正的子形状，按 z-order 从低到高排列。
    //
    // 之前把水印插到 SpgrContainer 末尾（z-order 最高），导致：
    // 1. 水印在最上层，遮挡其他内容
    // 2. 水印易被选中编辑，违背水印作为背景元素的特性
    // 现在插到"组形状本身"之后，水印成为 z-order 最低的真正形状：
    // - 视觉上被其他内容覆盖，符合水印背景特性
    // - 配合 FOPT 锁定属性（0x01C2=0x0D），在普通视图下不可选中/编辑
    // - 这正是 PowerPoint 自带"插入水印"功能的实现方式
    let mut pos = spgr_offset + 8;
    let mut first_sp_end = None;
    while pos + 8 <= spgr_end {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8 + rec_len as usize;

        // 找到第一个 SpContainer（组形状本身），记录其结束位置
        if is_container && rec_type == RT_SP_CONTAINER {
            first_sp_end = Some(pos + total_len);
            break;
        }

        pos += total_len;
        if !is_container && rec_len == 0 {
            break;
        }
    }

    // 插入位置：第一个 SpContainer 之后（z-order 最低的真正形状位置）
    // 如果找不到第一个 SpContainer（异常情况），退回到末尾插入以保证兼容
    let insert_pos = first_sp_end.unwrap_or(spgr_end);
    data.splice(insert_pos..insert_pos, watermark_sp.iter().copied());

    let insert_len = watermark_sp.len();

    // 更新三层 container 的 recLen（从内到外）：
    // 1. SpgrContainer 的 recLen
    write_u32_le(data, spgr_offset + 4, spgr_len + insert_len as u32)?;
    // 2. DgContainer 的 recLen
    write_u32_le(data, dg_offset + 4, dg_len + insert_len as u32)?;
    // 3. PPDrawing 的 recLen
    write_u32_le(data, ppd_offset + 4, ppd_len + insert_len as u32)?;

    Ok((insert_pos, insert_len))
}

/// 构造水印 SpContainer。
///
/// 结构（MS-ODRAW 规范）：
///
/// ```text
/// SpContainer (0xF004, container)
/// ├── FSP (0xF00A): ver=2, inst=0xCA (MSOSPT_TextBox=202)
/// │   ├── shapeId (u32): 唯一形状 ID
/// │   └── flags (u32): fHaveAnchor + fHaveSpt
/// ├── FOPT (0xF00B): ver=3, inst=属性个数
/// │   ├── 0x00BD (rotation): 旋转角度（16.16 固定点数）
/// │   ├── 0x0180 (fillType): 0 = No fill
/// │   ├── 0x01BF (Fill Style Boolean): fNoFill + fFillOK + fNoFillHitTest
/// │   ├── 0x01C1 (Line Style Boolean): fNoLine + fLineOK + fNoLineDrawDash
/// │   └── 0x01C2 (Protection Boolean): 锁定分组+文本编辑+选择
/// ├── ClientAnchor (0xF010): ver=0, 全屏覆盖 (0,0 → 5760,4320)
/// └── ClientTextbox (0xF00D, container): 文本框
///     ├── TextHeaderAtom (0x0F9F): txType=4 (not body)
///     ├── TextCharsAtom (0x0FA0): 水印文本（UTF-16LE）
///     └── StyleTextPropAtom (0x0FA1): 文本样式（字号、颜色）
/// ```
///
/// # FOPT 保护位详解
///
/// `0x01C2` (Protection Boolean Properties) = `0x0000000D`：
/// - bit 0 (fLockAgainstGrouping)：防止选择、分组、移动
/// - bit 2 (fLockAgainstTextEdit)：防止文本编辑
/// - bit 3 (fLockAgainstSelection)：防止选择
///
/// # 参数
/// - `shape_id`：水印形状的唯一 ID（由调用方确保不冲突）
/// - `config`：水印配置（文本、字号、颜色、旋转角度）
///
/// # 返回
/// 完整的 SpContainer 字节序列（含 record header）。
pub fn build_watermark_spcontainer(shape_id: u32, config: &WatermarkConfig) -> Vec<u8> {
    let mut children = Vec::new();

    // 1. FSP (0xF00A): 形状属性
    // ver=0x2, inst=0xCA (MSOSPT_TextBox=202), type=0xF00A, len=8
    // shapeId (4 bytes) + flags (4 bytes)
    // MSOSPT_TextBox = 202 = 0xCA（MS-ODRAW 规范 2.4.14 MSOSPT 枚举）
    let mut fsp = Vec::new();
    let ver_inst: u16 = (0xCA << 4) | 0x2; // inst=0xCA (textBox=202), ver=0x2
    fsp.extend_from_slice(&ver_inst.to_le_bytes());
    fsp.extend_from_slice(&RT_FSP.to_le_bytes());
    fsp.extend_from_slice(&8u32.to_le_bytes()); // len=8
    fsp.extend_from_slice(&shape_id.to_le_bytes()); // shapeId
                                                    // flags (MS-ODRAW 2.3.1.1 FSP):
                                                    //   bit 7 (0x80): fHaveAnchor — 1 = 形状有 anchor（必须有，否则 PowerPoint 忽略 ClientAnchor）
                                                    //   bit 9 (0x200): fHaveSpt — 1 = 形状有 shape type
                                                    // 正确值：fHaveAnchor + fHaveSpt = 0x80 + 0x200 = 0x280
    fsp.extend_from_slice(&0x00000280u32.to_le_bytes()); // flags: fHaveAnchor + fHaveSpt
    children.extend_from_slice(&fsp);

    // 2. FOPT (0xF00B): 形状属性（无填充、无线条，旋转，锁定不可编辑）
    // FOPT 属性按 property ID 升序排列
    let mut fopt_props = Vec::new();
    // 0x00BD (rotation): 旋转角度（固定点数 16.16 格式，1度 = 65536）
    let rotation_fixed: u32 = ((config.rotation_deg as i64) * 65536) as u32;
    fopt_props.extend_from_slice(&0x00BDu16.to_le_bytes());
    fopt_props.extend_from_slice(&rotation_fixed.to_le_bytes());
    // 0x0180 (fillType): 0 = No fill（明确指定无填充，避免 PowerPoint 使用默认白色背景）
    fopt_props.extend_from_slice(&0x0180u16.to_le_bytes());
    fopt_props.extend_from_slice(&0x00000000u32.to_le_bytes());
    // 0x01BF (Fill Style Boolean Properties): fNoFill + fFillOK + fNoFillHitTest = 0x00000043
    fopt_props.extend_from_slice(&0x01BFu16.to_le_bytes());
    fopt_props.extend_from_slice(&0x00000043u32.to_le_bytes());
    // 0x01C1 (Line Style Boolean Properties): fNoLine + fLineOK + fNoLineDrawDash = 0x00000043
    fopt_props.extend_from_slice(&0x01C1u16.to_le_bytes());
    fopt_props.extend_from_slice(&0x00000043u32.to_le_bytes());
    // 0x01C2 (Protection Boolean Properties): 锁定分组+文本编辑+选择 = 0x0000000D
    fopt_props.extend_from_slice(&0x01C2u16.to_le_bytes());
    fopt_props.extend_from_slice(&0x0000000Du32.to_le_bytes());

    let num_props = fopt_props.len() / 6;
    let mut fopt = Vec::new();
    let ver_inst: u16 = ((num_props as u16) << 4) | 0x3; // inst=num_props, ver=0x3
    fopt.extend_from_slice(&ver_inst.to_le_bytes());
    fopt.extend_from_slice(&RT_FOPT.to_le_bytes());
    fopt.extend_from_slice(&(fopt_props.len() as u32).to_le_bytes());
    fopt.extend_from_slice(&fopt_props);
    children.extend_from_slice(&fopt);

    // 3. ClientAnchor (0xF010): 形状位置
    // ver=0, inst=0, len=8（SmallRectStruct: 4 个 int16: top, left, right, bottom）
    // 单位是 master units（1/576 英寸），slide 标准尺寸 5760 x 4320（10 x 7.5 英寸）
    // 水印覆盖整个 slide 区域，配合旋转和大字号，让水印文字斜向铺满整个幻灯片
    let mut anchor = Vec::new();
    let ver_inst: u16 = 0; // inst=0, ver=0
    anchor.extend_from_slice(&ver_inst.to_le_bytes());
    anchor.extend_from_slice(&RT_CLIENT_ANCHOR.to_le_bytes());
    anchor.extend_from_slice(&8u32.to_le_bytes()); // len=8
    anchor.extend_from_slice(&0i16.to_le_bytes()); // top = 0
    anchor.extend_from_slice(&0i16.to_le_bytes()); // left = 0
    anchor.extend_from_slice(&5760i16.to_le_bytes()); // right = 5760（10 英寸）
    anchor.extend_from_slice(&4320i16.to_le_bytes()); // bottom = 4320（7.5 英寸）
    children.extend_from_slice(&anchor);

    // 4. ClientTextbox (0xF00D, container): 文本框
    let mut textbox_children = Vec::new();

    // 4.1 TextHeaderAtom (0x0F9F): txType=4 (not body)
    let mut text_header = Vec::new();
    let ver_inst: u16 = 0;
    text_header.extend_from_slice(&ver_inst.to_le_bytes());
    text_header.extend_from_slice(&RT_TEXT_HEADER_ATOM.to_le_bytes());
    text_header.extend_from_slice(&4u32.to_le_bytes()); // len=4
    text_header.extend_from_slice(&4u32.to_le_bytes()); // txType=4 (not body)
    textbox_children.extend_from_slice(&text_header);

    // 4.2 TextCharsAtom (0x0FA0): 水印文本（UTF-16LE）
    let text_utf16: Vec<u8> = config
        .text
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();
    let text_char_count = config.text.encode_utf16().count() as u32;
    let mut text_chars = Vec::new();
    let ver_inst: u16 = 0;
    text_chars.extend_from_slice(&ver_inst.to_le_bytes());
    text_chars.extend_from_slice(&RT_TEXT_CHARS_ATOM.to_le_bytes());
    text_chars.extend_from_slice(&(text_utf16.len() as u32).to_le_bytes());
    text_chars.extend_from_slice(&text_utf16);
    textbox_children.extend_from_slice(&text_chars);

    // 4.3 StyleTextPropAtom (0x0FA1): 文本样式（字体大小、颜色）
    // StyleTextPropAtom 结构（MS-PPT 2.9.17 规范）：
    //   lfo (4 bytes): 段落格式 run 数量
    //   rgTextPFRun[lfo]: 段落格式 run
    //   rgTextCFRun[lfo]: 字符格式 run
    // TextPFRun: count(4) + indentLevel(2) + pfFlags(4) = 10 bytes
    // TextCFRun: count(4) + cfFlags(4) + [sz(2) if fSize] + [color(4) if fColor]
    // cfFlags 位定义：
    //   bit 6 (0x40): fSize — 设置字体大小
    //   bit 7 (0x80): fColor — 设置文字颜色
    let mut style_data = Vec::new();
    style_data.extend_from_slice(&1u32.to_le_bytes()); // lfo = 1
    style_data.extend_from_slice(&text_char_count.to_le_bytes()); // count
    style_data.extend_from_slice(&0u16.to_le_bytes()); // indentLevel = 0
    style_data.extend_from_slice(&0u32.to_le_bytes()); // pfFlags = 0
    style_data.extend_from_slice(&text_char_count.to_le_bytes()); // count
    style_data.extend_from_slice(&0x000000C0u32.to_le_bytes()); // cfFlags: fSize + fColor
                                                                // 字号（单位是 1/100 pt）
    let font_size_value = config.font_size_pt as u32 * 100;
    style_data.extend_from_slice(&(font_size_value as u16).to_le_bytes()); // sz
                                                                           // color: ColorIndexStruct (red, green, blue, index)
                                                                           // 半透明效果：使用中灰色模拟半透明（PPT 97-2003 文本透明度需要复杂扩展属性）
    style_data.extend_from_slice(&config.color_rgb.0.to_le_bytes()); // red
    style_data.extend_from_slice(&config.color_rgb.1.to_le_bytes()); // green
    style_data.extend_from_slice(&config.color_rgb.2.to_le_bytes()); // blue
    style_data.extend_from_slice(&0u8.to_le_bytes()); // index = 0 (RGB)

    let mut style_atom = Vec::new();
    let ver_inst: u16 = 0;
    style_atom.extend_from_slice(&ver_inst.to_le_bytes());
    style_atom.extend_from_slice(&RT_STYLE_TEXT_PROP_ATOM.to_le_bytes());
    style_atom.extend_from_slice(&(style_data.len() as u32).to_le_bytes());
    style_atom.extend_from_slice(&style_data);
    textbox_children.extend_from_slice(&style_atom);

    // 组装 ClientTextbox
    let mut client_textbox = Vec::new();
    let ver_inst: u16 = 0xF; // container
    client_textbox.extend_from_slice(&ver_inst.to_le_bytes());
    client_textbox.extend_from_slice(&RT_CLIENT_TEXTBOX.to_le_bytes());
    client_textbox.extend_from_slice(&(textbox_children.len() as u32).to_le_bytes());
    client_textbox.extend_from_slice(&textbox_children);
    children.extend_from_slice(&client_textbox);

    // 组装 SpContainer
    let mut sp_container = Vec::new();
    let ver_inst: u16 = 0xF; // container
    sp_container.extend_from_slice(&ver_inst.to_le_bytes());
    sp_container.extend_from_slice(&RT_SP_CONTAINER.to_le_bytes());
    sp_container.extend_from_slice(&(children.len() as u32).to_le_bytes());
    sp_container.extend_from_slice(&children);

    sp_container
}

/// 扫描所有 MainMaster 的 PPDrawing 中已有的最大 shapeId。
///
/// 参考 Apache POI `XSLFSheet.allocateShapeId()` 的 BitSet 模式：
/// 遍历 SpgrContainer 中所有 FSP record (0xF00A)，读取其 shapeId 字段，
/// 返回最大值。新水印 shape 从 max+1 开始分配，避免 ID 冲突。
///
/// # 参数
/// - `data`：PowerPoint Document stream
/// - `master_offsets`：所有 MainMaster 的 offset 列表
///
/// # 返回
/// 所有 MainMaster 中最大的 shapeId（若无形状返回 0）。
pub fn find_max_shape_id(data: &[u8], master_offsets: &[usize]) -> Result<u32> {
    let mut max_id: u32 = 0;
    for &master_offset in master_offsets {
        if let Some(ppd_offset) = find_ppdrawing_in_master(data, master_offset)? {
            max_id = max_id.max(scan_shape_ids_in_ppdrawing(data, ppd_offset)?);
        }
    }
    Ok(max_id)
}

/// 递归扫描 PPDrawing 子树中所有 FSP record 的 shapeId，返回最大值。
///
/// PPDrawing → DgContainer → SpgrContainer → [SpContainer | SpgrContainer]...
/// 每个 SpContainer 内含一个 FSP (0xF00A)，其前 4 字节数据是 shapeId。
/// SpgrContainer 可嵌套，需递归遍历。
fn scan_shape_ids_in_ppdrawing(data: &[u8], ppd_offset: usize) -> Result<u32> {
    let (_, _, _, ppd_len) = parse_record_header(data, ppd_offset)?;
    let ppd_end = ppd_offset + 8 + ppd_len as usize;
    let mut max_id: u32 = 0;
    let mut pos = ppd_offset + 8;

    while pos + 8 <= ppd_end {
        let (ver, _, rec_type, rec_len) = parse_record_header(data, pos)?;
        let is_container = ver == 0xF;
        let total_len = 8 + rec_len as usize;

        if is_container {
            // 递归扫描 container 子节点
            let child_max = scan_shape_ids_in_ppdrawing(data, pos)?;
            max_id = max_id.max(child_max);
        } else if rec_type == RT_FSP && rec_len >= 8 {
            // FSP record: shapeId 在 header 之后的第 4 字节
            let shape_id = read_u32_le(data, pos + 8)?;
            if shape_id > max_id {
                max_id = shape_id;
            }
        }

        pos += total_len;
        if !is_container && rec_len == 0 {
            break;
        }
    }

    Ok(max_id)
}

/// 检查 PPDrawing 中是否已存在指定水印文本（幂等性检查）。
///
/// 参考 Aspose.Slides 按 shape name/text 查重的幂等性模式：
/// 将水印文本编码为 UTF-16LE（与 TextCharsAtom 0x0FA0 的编码一致），
/// 在 PPDrawing 的数据范围内搜索该字节序列。
///
/// # 参数
/// - `data`：PowerPoint Document stream
/// - `ppd_offset`：PPDrawing record 的起始 offset
/// - `watermark_utf16`：水印文本的 UTF-16LE 字节序列
///
/// # 返回
/// - `true`：PPDrawing 中已存在水印文本（应跳过注入）
/// - `false`：未找到水印文本（可以注入）
pub fn has_watermark_text(data: &[u8], ppd_offset: usize, watermark_utf16: &[u8]) -> bool {
    let (_, _, _, ppd_len) = match parse_record_header(data, ppd_offset) {
        Ok(v) => v,
        Err(_) => return false,
    };
    let ppd_end = ppd_offset + 8 + ppd_len as usize;
    if ppd_end > data.len() {
        return false;
    }
    // 在 PPDrawing 的完整数据范围内搜索水印文本的 UTF-16LE 字节序列
    data[ppd_offset..ppd_end]
        .windows(watermark_utf16.len())
        .any(|w| w == watermark_utf16)
}

/// 抑制未使用常量警告（RT_MAIN_MASTER 在 inject_watermark 中通过 find_main_masters 间接使用）。
#[allow(dead_code)]
const _UNUSED_RT_MAIN_MASTER: u16 = RT_MAIN_MASTER;
