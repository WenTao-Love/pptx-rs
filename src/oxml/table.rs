//! 表格：`<a:tbl>` 嵌入 `<p:graphicFrame>`。
//!
//! OOXML 中"表格"由三层结构组成：
//!
//! ```text
//! <p:graphicFrame>
//!   <p:nvGraphicFramePr><p:cNvPr .../></p:nvGraphicFramePr>
//!   <p:xfrm>...</p:xfrm>
//!   <a:graphic>
//!     <a:graphicData uri="...">
//!       <a:tbl>                  ← 本模块关心的内容
//!         <a:tblPr>...</a:tblPr>
//!         <a:tblGrid>...</a:tblGrid>
//!         <a:tr>...</a:tr>       ← 行
//!           <a:tc>...</a:tc>     ← 单元格
//!         ...
//!       </a:tbl>
//!     </a:graphicData>
//!   </a:graphic>
//! </p:graphicFrame>
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.table.Table` ←→ [`Table`]；
//! - `_Row` / `_Column` / `_Cell` ←→ [`Row`] / [`Col`] / [`Cell`]。
//!
//! # 当前限制
//!
//! - 列宽 / 行高单位均为 EMU（OOXML 标准）。
//! - 表格样式通过 GUID 引用 PowerPoint 内置样式（见 [`TableStyle`]）；
//!   自定义 `<a:tblStyle>` 定义（tableStyles.xml）尚未支持。

use crate::oxml::color::Color;
use crate::oxml::txbody::TextBody;
use crate::units::Emu;

/// 表格列属性。
#[derive(Clone, Debug, Default)]
pub struct Col {
    /// 列宽（EMU）。
    pub width: Emu,
}

/// 表格行属性。
#[derive(Clone, Debug, Default)]
pub struct Row {
    /// 行高（EMU）。
    pub height: Emu,
    /// 行内每个 cell 的内容。
    pub cells: Vec<Cell>,
    /// 标识行头/列头。
    pub header: bool,
}

/// 单元格垂直对齐方式（`<a:tcPr anchor="...">`）。
///
/// 对应 python-pptx `MSO_ANCHOR` 枚举。
#[derive(Clone, Debug, Default, PartialEq)]
pub enum VerticalAnchor {
    /// 顶部对齐（`anchor="t"`）。
    Top,
    /// 中部对齐（`anchor="ctr"`，默认）。
    #[default]
    Middle,
    /// 底部对齐（`anchor="b"`）。
    Bottom,
}

impl VerticalAnchor {
    /// 转为 OOXML 属性值。
    pub fn as_str(&self) -> &'static str {
        match self {
            VerticalAnchor::Top => "t",
            VerticalAnchor::Middle => "ctr",
            VerticalAnchor::Bottom => "b",
        }
    }
}

/// 单元格边框样式（简化版，仅支持实线/虚线/无边框）。
///
/// 对应 `<a:lnL>` / `<a:lnR>` / `<a:lnT>` / `<a:lnB>` 子元素。
#[derive(Clone, Debug, Default)]
pub struct CellBorder {
    /// 边框颜色。`Color::None` 表示不写出颜色（使用主题继承）。
    pub color: Color,
    /// 边框宽度（EMU）。
    pub width: Emu,
    /// 是否无边框（写出 `<a:noFill/>`）。
    pub no_fill: bool,
}

/// 单元格。
#[derive(Clone, Debug, Default)]
pub struct Cell {
    /// 单元格文本体。
    pub text: TextBody,
    /// 单元格填充色。`Color::None` 表示不写出填充。
    pub fill: Color,
    /// 上下左右内边距（EMU）。
    pub margin: (Option<Emu>, Option<Emu>, Option<Emu>, Option<Emu>), // top,left,bottom,right
    /// 跨行数（`rowSpan` 属性，0 或 1 表示不跨行）。
    pub row_span: u32,
    /// 跨列数（`gridSpan` 属性，0 或 1 表示不跨列）。
    pub grid_span: u32,
    /// 水平合并虚拟单元格标记（`hMerge="1"`，被合并方写出）。
    pub h_merge: bool,
    /// 垂直合并虚拟单元格标记（`vMerge="1"`，被合并方写出）。
    pub v_merge: bool,
    /// 垂直对齐方式。
    pub anchor: VerticalAnchor,
    /// 左边框（`<a:lnL>`）。
    pub border_left: Option<CellBorder>,
    /// 右边框（`<a:lnR>`）。
    pub border_right: Option<CellBorder>,
    /// 上边框（`<a:lnT>`）。
    pub border_top: Option<CellBorder>,
    /// 下边框（`<a:lnB>`）。
    pub border_bottom: Option<CellBorder>,
}

/// 表格布尔属性（`<a:tblLook>` 的属性集）。
///
/// 对应 python-pptx `_Table.tbl_look` 的行为，控制表格样式应用范围。
#[derive(Clone, Debug)]
pub struct TableLook {
    /// 样式索引值（默认 `"04A0"`）。
    pub val: String,
    /// 首行特殊格式（`firstRow="1"`）。
    pub first_row: bool,
    /// 末行特殊格式（`lastRow="1"`）。
    pub last_row: bool,
    /// 首列特殊格式（`firstColumn="1"`）。
    pub first_column: bool,
    /// 末列特殊格式（`lastColumn="1"`）。
    pub last_column: bool,
    /// 水平条纹（`noHBand="0"` 表示启用）。
    pub no_h_band: bool,
    /// 垂直条纹（`noVBand="1"` 表示禁用）。
    pub no_v_band: bool,
}

impl Default for TableLook {
    fn default() -> Self {
        // 默认值与 PowerPoint 创建的表格一致
        Self {
            val: "04A0".to_string(),
            first_row: true,
            last_row: false,
            first_column: true,
            last_column: false,
            no_h_band: false,
            no_v_band: true,
        }
    }
}

/// 表格样式引用（`<a:tableStyleId>{GUID}</a:tableStyleId>`）。
///
/// OOXML 中表格样式通过 GUID 引用 `tableStyles.xml` 中定义的 `<a:tblStyle>`，
/// 或引用 PowerPoint 内置样式（内置样式的 GUID 是预定义的，无需在包内定义）。
///
/// # 与 python-pptx 的对应
///
/// - python-pptx `Table.apply_style(style_id)` ←→ [`TableStyle::new`] + [`Table::set_style`]；
/// - python-pptx 默认使用 `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`（Medium Style 2 - Accent 1）。
///
/// # 内置样式
///
/// PowerPoint 内置表格样式的 GUID 是固定的，可通过 [`TableStyle::from_name`] 按名称查找。
/// 常见样式包括 "No Style, Table Grid"、"Medium Style 2 - Accent 1" 等。
///
/// # 示例
///
/// ```
/// use pptx_rs::oxml::table::TableStyle;
///
/// // 按名称查找内置样式
/// let style = TableStyle::from_name("Medium Style 2 - Accent 1").unwrap();
/// assert_eq!(style.style_id(), "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
///
/// // 直接用 GUID 构造
/// let style2 = TableStyle::new("{5940675A-B579-460E-94D1-54222C63F5DA}");
/// assert_eq!(style2.style_id(), "{5940675A-B579-460E-94D1-54222C63F5DA}");
/// ```
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableStyle {
    /// 样式 GUID（如 `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`）。
    style_id: String,
    /// 可选的样式名称（如 "Medium Style 2 - Accent 1"）。
    style_name: Option<String>,
}

impl TableStyle {
    /// 用 GUID 构造表格样式引用。
    ///
    /// # 参数
    /// - `style_id`：样式 GUID 字符串（应包含花括号，如 `{5C22544A-...}`）。
    pub fn new(style_id: impl Into<String>) -> Self {
        Self {
            style_id: style_id.into(),
            style_name: None,
        }
    }

    /// 用 GUID + 名称构造表格样式引用。
    ///
    /// # 参数
    /// - `style_id`：样式 GUID 字符串；
    /// - `style_name`：人类可读的样式名称。
    pub fn with_name(style_id: impl Into<String>, style_name: impl Into<String>) -> Self {
        Self {
            style_id: style_id.into(),
            style_name: Some(style_name.into()),
        }
    }

    /// 按名称查找 PowerPoint 内置表格样式。
    ///
    /// # 参数
    /// - `name`：样式名称（如 "Medium Style 2 - Accent 1"）。
    ///
    /// # 返回值
    /// - 找到时返回 `Some(TableStyle)`；
    /// - 名称不在内置注册表中时返回 `None`。
    ///
    /// # 内置样式列表
    ///
    /// 目前注册的内置样式（可扩展）：
    /// - "No Style, Table Grid"
    /// - "No Style, No Grid"
    /// - "Medium Style 2 - Accent 1"
    /// - "Themed Style 1 - Accent 1"
    pub fn from_name(name: &str) -> Option<Self> {
        builtin_table_style_guid(name).map(|guid| Self::with_name(guid, name))
    }

    /// 取样式 GUID。
    pub fn style_id(&self) -> &str {
        &self.style_id
    }

    /// 取样式名称（可能为 `None`）。
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }
}

/// 查找 PowerPoint 内置表格样式的 GUID。
///
/// 返回 `None` 表示名称不在内置注册表中。
/// 调用方可通过 [`TableStyle::new`] 直接用 GUID 构造。
fn builtin_table_style_guid(name: &str) -> Option<&'static str> {
    // PowerPoint 内置表格样式 GUID 注册表。
    // 这些 GUID 是 OOXML 规范和 PowerPoint 预定义的，无需在 tableStyles.xml 中定义。
    // 参考：ECMA-376 第 20.1.4.2.27 节 tblStyleLst 示例。
    match name {
        // 无样式系列
        "No Style, Table Grid" => Some("{5940675A-B579-460E-94D1-54222C63F5DA}"),
        "No Style, No Grid" => Some("{2D5ABB26-0587-4C30-8999-92F81FD0307C}"),
        // Medium Style 2 系列（python-pptx 默认）
        "Medium Style 2 - Accent 1" => Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"),
        // Themed Style 1 系列（OOXML 规范示例）
        "Themed Style 1 - Accent 1" => Some("{3C2FFA5D-87B4-456A-9821-1D502468CF0F}"),
        _ => None,
    }
}

/// 表格。
#[derive(Clone, Debug, Default)]
pub struct Table {
    pub cols: Vec<Col>,
    pub rows: Vec<Row>,
    /// 表格样式查找属性（控制 tblLook）。
    pub tbl_look: TableLook,
    /// 表格样式引用（`<a:tableStyleId>`）。
    ///
    /// `None` 时不写出 `<a:tableStyleId>` 元素。
    pub table_style: Option<TableStyle>,
}

impl Table {
    /// 写 XML（`<a:tbl>` 嵌在 `<a:graphicData uri="...">` 内）。
    ///
    /// # OOXML 元素顺序
    ///
    /// ```text
    /// <a:tbl>
    ///   <a:tblPr>...</a:tblPr>           ← 必须先写，且不含 `<a:tblGrid>`
    ///     <a:tableStyleId>...</a:tableStyleId>  ← 可选，紧跟 tblPr 属性后
    ///   <a:tblGrid>...</a:tblGrid>       ← 列定义（gridCol 列表）
    ///   <a:tr>...</a:tr>                 ← 行（先 cell 后 close）
    ///     <a:tc>                         ← 单元格
    ///       <a:txBody>...</a:txBody>     ← 文本
    ///       <a:tcPr>                     ← 单元格属性（边距/边框/对齐）
    ///         <a:lnL/>/<a:lnR/>/<a:lnT/>/<a:lnB/>
    ///       </a:tcPr>
    ///     </a:tc>
    ///   </a:tr>
    ///   ...
    /// </a:tbl>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("a:tbl");
        // tblPr：注意 `<a:tblGrid/>` **不能**塞在这里，必须紧跟其后另开块。
        w.open("a:tblPr");
        w.empty_with("a:tblW", &[("w", "0"), ("type", "auto")]);
        // tblLook：从结构体字段生成属性，不再硬编码
        let lk = &self.tbl_look;
        w.empty_with(
            "a:tblLook",
            &[
                ("val", lk.val.as_str()),
                ("firstRow", if lk.first_row { "1" } else { "0" }),
                ("lastRow", if lk.last_row { "1" } else { "0" }),
                ("firstColumn", if lk.first_column { "1" } else { "0" }),
                ("lastColumn", if lk.last_column { "1" } else { "0" }),
                ("noHBand", if lk.no_h_band { "1" } else { "0" }),
                ("noVBand", if lk.no_v_band { "1" } else { "0" }),
            ],
        );
        // tableStyleId：写出表格样式 GUID（可选）
        if let Some(style) = &self.table_style {
            w.leaf("a:tableStyleId", style.style_id());
        }
        w.close("a:tblPr");
        // tblGrid：每个 col 一行
        w.open("a:tblGrid");
        for c in &self.cols {
            w.empty_with("a:gridCol", &[("w", &c.width.value().to_string())]);
        }
        w.close("a:tblGrid");
        // rows
        for r in &self.rows {
            w.open_with("a:tr", &[("h", &r.height.value().to_string())]);
            for c in &r.cells {
                // 构建 tc 的属性列表（gridSpan/rowSpan/hMerge/vMerge）
                let mut tc_attrs: Vec<(&str, String)> = Vec::new();
                if c.grid_span > 1 {
                    tc_attrs.push(("gridSpan", c.grid_span.to_string()));
                }
                if c.row_span > 1 {
                    tc_attrs.push(("rowSpan", c.row_span.to_string()));
                }
                if c.h_merge {
                    tc_attrs.push(("hMerge", "1".to_string()));
                }
                if c.v_merge {
                    tc_attrs.push(("vMerge", "1".to_string()));
                }
                let attr_refs: Vec<(&str, &str)> =
                    tc_attrs.iter().map(|(k, v)| (*k, v.as_str())).collect();
                if attr_refs.is_empty() {
                    w.open("a:tc");
                } else {
                    w.open_with("a:tc", &attr_refs);
                }

                // 文本体
                if c.text.paragraphs.is_empty() {
                    let tb = TextBody::new();
                    tb.write_xml(w);
                } else {
                    c.text.write_xml(w);
                }

                // tcPr：单元格属性（边距、边框、垂直对齐、填充）
                let has_margins = c.margin.0.is_some()
                    || c.margin.1.is_some()
                    || c.margin.2.is_some()
                    || c.margin.3.is_some();
                let has_borders = c.border_left.is_some()
                    || c.border_right.is_some()
                    || c.border_top.is_some()
                    || c.border_bottom.is_some();
                let has_fill = !matches!(c.fill, Color::None);
                let has_anchor = c.anchor != VerticalAnchor::default();

                if has_margins || has_borders || has_fill || has_anchor {
                    // 提前取出所有要序列化的字符串，扩展到块末尾
                    let mart_s = c.margin.0.map(|m| m.value().to_string());
                    let marl_s = c.margin.1.map(|m| m.value().to_string());
                    let marb_s = c.margin.2.map(|m| m.value().to_string());
                    let marr_s = c.margin.3.map(|m| m.value().to_string());
                    let mut tcpr_attrs: Vec<(&str, &str)> = Vec::new();
                    if let Some(s) = &mart_s {
                        tcpr_attrs.push(("marT", s));
                    }
                    if let Some(s) = &marl_s {
                        tcpr_attrs.push(("marL", s));
                    }
                    if let Some(s) = &marb_s {
                        tcpr_attrs.push(("marB", s));
                    }
                    if let Some(s) = &marr_s {
                        tcpr_attrs.push(("marR", s));
                    }
                    if has_anchor {
                        tcpr_attrs.push(("anchor", c.anchor.as_str()));
                    }
                    if tcpr_attrs.is_empty() {
                        w.open("a:tcPr");
                    } else {
                        w.open_with("a:tcPr", &tcpr_attrs);
                    }
                    // 填充色（solidFill）
                    if has_fill {
                        c.fill.write_solid_fill(w);
                    }
                    // 边框：OOXML 顺序为 lnL → lnR → lnT → lnB
                    write_cell_border(w, "a:lnL", c.border_left.as_ref());
                    write_cell_border(w, "a:lnR", c.border_right.as_ref());
                    write_cell_border(w, "a:lnT", c.border_top.as_ref());
                    write_cell_border(w, "a:lnB", c.border_bottom.as_ref());
                    w.close("a:tcPr");
                } else if has_fill {
                    // 仅有填充色时也要写出 tcPr
                    w.open("a:tcPr");
                    c.fill.write_solid_fill(w);
                    w.close("a:tcPr");
                }
                w.close("a:tc");
            }
            w.close("a:tr");
        }
        w.close("a:tbl");
    }

    /// 设置表格样式（按内置样式名称）。
    ///
    /// 对标 python-pptx `Table.apply_style(style_id)`。
    ///
    /// # 参数
    /// - `name`：内置样式名称（如 "Medium Style 2 - Accent 1"）。
    ///
    /// # 返回值
    /// - 成功设置返回 `true`；
    /// - 名称不在内置注册表中返回 `false`（`table_style` 保持不变）。
    ///
    /// # 示例
    /// ```
    /// use pptx_rs::oxml::table::Table;
    ///
    /// let mut t = Table::default();
    /// assert!(t.set_style("Medium Style 2 - Accent 1"));
    /// assert_eq!(
    ///     t.table_style.as_ref().unwrap().style_id(),
    ///     "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    /// );
    /// assert!(!t.set_style("Unknown Style"));
    /// ```
    pub fn set_style(&mut self, name: &str) -> bool {
        if let Some(style) = TableStyle::from_name(name) {
            self.table_style = Some(style);
            true
        } else {
            false
        }
    }

    /// 设置表格样式（按原始 GUID）。
    ///
    /// 用于设置不在内置注册表中的样式（如自定义 tableStyles.xml 中定义的样式）。
    ///
    /// # 参数
    /// - `guid`：样式 GUID 字符串（如 `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`）。
    pub fn set_style_id(&mut self, guid: impl Into<String>) {
        self.table_style = Some(TableStyle::new(guid));
    }

    /// 清除表格样式引用。
    pub fn clear_style(&mut self) {
        self.table_style = None;
    }
}

/// 写出单元格边框子元素。
///
/// `tag` 为 `a:lnL` / `a:lnR` / `a:lnT` / `a:lnB` 之一。
fn write_cell_border(w: &mut super::writer::XmlWriter, tag: &str, border: Option<&CellBorder>) {
    if let Some(b) = border {
        // 提前取出宽度字符串，扩展到函数末尾
        let w_s = b.width.value().to_string();
        if b.no_fill {
            w.open_with(tag, &[("w", w_s.as_str())]);
            w.empty("a:noFill");
            w.close(tag);
        } else if !matches!(b.color, Color::None) {
            w.open_with(tag, &[("w", w_s.as_str())]);
            b.color.write_solid_fill(w);
            w.close(tag);
        } else {
            // 无颜色也无 noFill：仅写空标签
            w.empty_with(tag, &[("w", w_s.as_str())]);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `TableStyle::from_name` 正确查找内置样式。
    #[test]
    fn table_style_from_name_finds_builtin() {
        let style = TableStyle::from_name("Medium Style 2 - Accent 1")
            .expect("Medium Style 2 - Accent 1 应在注册表中");
        assert_eq!(style.style_id(), "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        assert_eq!(style.style_name(), Some("Medium Style 2 - Accent 1"));

        let style2 = TableStyle::from_name("No Style, Table Grid")
            .expect("No Style, Table Grid 应在注册表中");
        assert_eq!(style2.style_id(), "{5940675A-B579-460E-94D1-54222C63F5DA}");
    }

    /// `TableStyle::from_name` 对未知名称返回 `None`。
    #[test]
    fn table_style_from_name_unknown_returns_none() {
        assert!(TableStyle::from_name("Nonexistent Style").is_none());
    }

    /// `TableStyle::new` 用 GUID 构造，`style_name` 为 `None`。
    #[test]
    fn table_style_new_has_no_name() {
        let style = TableStyle::new("{5940675A-B579-460E-94D1-54222C63F5DA}");
        assert_eq!(style.style_id(), "{5940675A-B579-460E-94D1-54222C63F5DA}");
        assert_eq!(style.style_name(), None);
    }

    /// `Table::set_style` 成功设置内置样式。
    #[test]
    fn table_set_style_builtin() {
        let mut t = Table::default();
        assert!(t.set_style("Medium Style 2 - Accent 1"));
        let style = t.table_style.as_ref().expect("style 应已设置");
        assert_eq!(style.style_id(), "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        assert_eq!(style.style_name(), Some("Medium Style 2 - Accent 1"));
    }

    /// `Table::set_style` 对未知名称返回 `false` 且不修改原状态。
    #[test]
    fn table_set_style_unknown_returns_false() {
        let mut t = Table::default();
        t.set_style("Medium Style 2 - Accent 1");
        assert!(!t.set_style("Unknown Style"));
        // 原样式应保持不变
        assert_eq!(
            t.table_style.as_ref().unwrap().style_id(),
            "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
        );
    }

    /// `Table::set_style_id` 用原始 GUID 设置样式。
    #[test]
    fn table_set_style_id_raw_guid() {
        let mut t = Table::default();
        t.set_style_id("{5940675A-B579-460E-94D1-54222C63F5DA}");
        let style = t.table_style.as_ref().expect("style 应已设置");
        assert_eq!(style.style_id(), "{5940675A-B579-460E-94D1-54222C63F5DA}");
        assert_eq!(style.style_name(), None);
    }

    /// `Table::clear_style` 清除样式引用。
    #[test]
    fn table_clear_style() {
        let mut t = Table::default();
        t.set_style("Medium Style 2 - Accent 1");
        assert!(t.table_style.is_some());
        t.clear_style();
        assert!(t.table_style.is_none());
    }

    /// `Table::write_xml` 正确序列化 `<a:tableStyleId>` 元素。
    #[test]
    fn table_write_xml_with_style() {
        let mut t = Table::default();
        t.set_style("Medium Style 2 - Accent 1");
        let mut w = crate::oxml::writer::XmlWriter::new();
        t.write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains("<a:tableStyleId>"));
        assert!(xml.contains("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"));
        assert!(xml.contains("</a:tableStyleId>"));
        // tableStyleId 应在 tblPr 内
        let pr_start = xml.find("<a:tblPr>").unwrap();
        let pr_end = xml.find("</a:tblPr>").unwrap();
        let style_pos = xml.find("<a:tableStyleId>").unwrap();
        let style_end = xml.find("</a:tableStyleId>").unwrap();
        assert!(style_pos > pr_start && style_end < pr_end);
    }

    /// `Table::write_xml` 无样式时不写出 `<a:tableStyleId>`。
    #[test]
    fn table_write_xml_without_style() {
        let t = Table::default();
        let mut w = crate::oxml::writer::XmlWriter::new();
        t.write_xml(&mut w);
        assert!(!w.buf.contains("tableStyleId"));
    }
}
