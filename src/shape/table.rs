//! `Table`：高阶表格。
//!
//! 表格在 OOXML 中嵌套较深（`graphicFrame` > `graphic` > `graphicData` > `tbl`），
//! 本高阶 API 把它们"摊平"为 [`TableShape`]，并以 `(row, col)` 二维索引访问。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.table.Table` ←→ [`TableShape`]；
//! - `_Row` / `_Column` / `_Cell` ←→ `Row` / `Col` / `Cell`（oxml 模型）。
//!
//! # 单元格语义
//!
//! - 每个 cell 持有 [`TextBody`]，可含多段多 Run；
//! - 颜色由 `Cell::fill` 决定（`Color::None` 时不写 `<a:solidFill/>`）；
//! - 边距由 `Cell::margin` 决定（顺序：top, left, bottom, right）。
//!
//! # 限制
//!
//! - 列宽 / 行高必须显式设置（默认全 0）。
//! - 表格样式通过 GUID 引用 PowerPoint 内置样式（见 [`TableStyle`]）。
//!
//! # 示例
//!
//! ```no_run
//! use pptx::shape::TableShape;
//! use pptx::EmuExt;
//! use pptx::Inches;
//!
//! let mut t = TableShape::new(2, 3, Inches(2.0).emu(), Inches(0.5).emu());
//! t.set_cell_text(0, 0, "A1").unwrap();
//! t.set_cell_text(0, 1, "B1").unwrap();
//! ```

use crate::oxml::shape::{Graphic as OxmlGraphic, GraphicFrame as OxmlFrame};
use crate::oxml::table::{Cell as OxmlCell, Col as OxmlCol, Row as OxmlRow, Table as OxmlTable};
use crate::oxml::txbody::TextBody;
use crate::shape::base::Shape;
use crate::units::Emu;

/// 表格（行优先，行内 cell）。
#[derive(Clone, Debug, Default)]
pub struct TableShape {
    /// 内部 oxml 句柄（`GraphicFrame`）。
    pub(crate) frame: OxmlFrame,
}

impl TableShape {
    /// 构造一个指定行列的表格（cell 文本留空）。
    #[allow(clippy::field_reassign_with_default)]
    pub fn new(rows: usize, cols: usize, col_width: Emu, row_height: Emu) -> Self {
        let mut t = OxmlTable::default();
        t.cols = (0..cols).map(|_| OxmlCol { width: col_width }).collect();
        t.rows = (0..rows)
            .map(|_| OxmlRow {
                height: row_height,
                cells: (0..cols).map(|_| OxmlCell::default()).collect(),
                header: false,
            })
            .collect();
        let mut frame = OxmlFrame::default();
        frame.graphic = OxmlGraphic::Table(t);
        TableShape { frame }
    }

    /// 从 oxml Frame 构造。
    pub fn from_frame(frame: OxmlFrame) -> Self {
        TableShape { frame }
    }

    /// 取 oxml Table 引用。
    pub fn table(&self) -> &OxmlTable {
        match &self.frame.graphic {
            OxmlGraphic::Table(t) => t,
            // 不变量被破坏时（frame.graphic 不是 Table）panic——
            // 与库整体"零 panic"约定一致，此处属内部不变量违反，
            // 不走静默忽略路径（与 ChartShape 等"返回 Option"的设计不同，
            // 因为 TableShape 的高阶 API 假设 frame.graphic 一定是 Table）。
            _ => unreachable!("TableShape.table(): frame.graphic 不是 Table 变体"),
        }
    }
    /// 取 oxml Table 可变引用。
    pub fn table_mut(&mut self) -> &mut OxmlTable {
        match &mut self.frame.graphic {
            OxmlGraphic::Table(t) => t,
            _ => unreachable!("TableShape.table_mut(): frame.graphic 不是 Table 变体"),
        }
    }

    /// 行列数。
    pub fn dims(&self) -> (usize, usize) {
        let t = self.table();
        (t.rows.len(), t.cols.len())
    }

    /// **是否**启用首行特殊格式（`firstRow="1"`）。
    ///
    /// 对标 python-pptx `Table.first_row`。
    pub fn first_row(&self) -> bool {
        self.table().tbl_look.first_row
    }
    /// 设置首行特殊格式。
    pub fn set_first_row(&mut self, v: bool) {
        self.table_mut().tbl_look.first_row = v;
    }

    /// **是否**启用末行特殊格式（`lastRow="1"`）。
    ///
    /// 对标 python-pptx `Table.last_row`。
    pub fn last_row(&self) -> bool {
        self.table().tbl_look.last_row
    }
    /// 设置末行特殊格式。
    pub fn set_last_row(&mut self, v: bool) {
        self.table_mut().tbl_look.last_row = v;
    }

    /// **是否**启用首列特殊格式（`firstColumn="1"`）。
    ///
    /// 对标 python-pptx `Table.first_column`。
    pub fn first_column(&self) -> bool {
        self.table().tbl_look.first_column
    }
    /// 设置首列特殊格式。
    pub fn set_first_column(&mut self, v: bool) {
        self.table_mut().tbl_look.first_column = v;
    }

    /// **是否**启用末列特殊格式（`lastColumn="1"`）。
    ///
    /// 对标 python-pptx `Table.last_column`。
    pub fn last_column(&self) -> bool {
        self.table().tbl_look.last_column
    }
    /// 设置末列特殊格式。
    pub fn set_last_column(&mut self, v: bool) {
        self.table_mut().tbl_look.last_column = v;
    }

    /// **是否**启用水平条纹（`noHBand="0"` 表示启用）。
    ///
    /// 对标 python-pptx `Table.horz_banding`。
    /// 注意：OOXML 中 `noHBand="1"` 表示**禁用**，本方法返回取反值（true=启用）。
    pub fn horz_banding(&self) -> bool {
        !self.table().tbl_look.no_h_band
    }
    /// 设置水平条纹。
    pub fn set_horz_banding(&mut self, v: bool) {
        self.table_mut().tbl_look.no_h_band = !v;
    }

    /// **是否**启用垂直条纹（`noVBand="0"` 表示启用）。
    ///
    /// 对标 python-pptx `Table.vert_banding`。
    /// 注意：OOXML 中 `noVBand="1"` 表示**禁用**，本方法返回取反值（true=启用）。
    pub fn vert_banding(&self) -> bool {
        !self.table().tbl_look.no_v_band
    }
    /// 设置垂直条纹。
    pub fn set_vert_banding(&mut self, v: bool) {
        self.table_mut().tbl_look.no_v_band = !v;
    }

    // --------------------- 表格样式 API（TODO-030） ---------------------

    /// 设置表格样式（按内置样式名称）。
    ///
    /// 对标 python-pptx `Table.apply_style(style_id)`。
    ///
    /// # 参数
    /// - `name`：内置样式名称（如 "Medium Style 2 - Accent 1"）。
    ///
    /// # 返回值
    /// - 成功设置返回 `true`；
    /// - 名称不在内置注册表中返回 `false`。
    ///
    /// # 示例
    ///
    /// ```
    /// use pptx::shape::TableShape;
    /// use pptx::Emu;
    ///
    /// let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
    /// assert!(t.set_style("Medium Style 2 - Accent 1"));
    /// assert!(!t.set_style("Unknown Style"));
    /// ```
    pub fn set_style(&mut self, name: &str) -> bool {
        self.table_mut().set_style(name)
    }

    /// 设置表格样式（按原始 GUID）。
    ///
    /// 用于设置不在内置注册表中的样式（如自定义 tableStyles.xml 中定义的样式）。
    ///
    /// # 参数
    /// - `guid`：样式 GUID 字符串（如 `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`）。
    pub fn set_style_id(&mut self, guid: impl Into<String>) {
        self.table_mut().set_style_id(guid);
    }

    /// 取表格样式引用。
    pub fn table_style(&self) -> Option<&crate::oxml::table::TableStyle> {
        self.table().table_style.as_ref()
    }

    /// 清除表格样式引用。
    pub fn clear_style(&mut self) {
        self.table_mut().clear_style();
    }

    /// 取 cell 文本（把多段多 Run 拼起来，段间 `\n`）。
    pub fn cell_text(&self, row: usize, col: usize) -> Option<String> {
        let t = self.table();
        let r = t.rows.get(row)?;
        let c = r.cells.get(col)?;
        let mut s = String::new();
        for p in &c.text.paragraphs {
            if !s.is_empty() {
                s.push('\n');
            }
            for run in &p.runs {
                s.push_str(&run.text);
            }
        }
        Some(s)
    }

    /// 设 cell 文本（自动新建段落 + Run）。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：`row` 或 `col` 越界。
    pub fn set_cell_text(&mut self, row: usize, col: usize, text: &str) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        let c = r
            .cells
            .get_mut(col)
            .ok_or(crate::Error::IndexOutOfRange(col))?;
        let mut tb = TextBody::new();
        let mut p = crate::oxml::txbody::Paragraph::new();
        p.runs.push(crate::oxml::txbody::Run::new(text));
        tb.paragraphs.push(p);
        c.text = tb;
        Ok(())
    }

    /// 取 cell 可变引用。
    ///
    /// 对应 python-pptx 中 `table.cell(row, col)`。**包含**创建空 cell 的能力（越界会
    /// 自动 push），适合"按需扩展"场景。
    pub fn cell_mut(&mut self, row: usize, col: usize) -> &mut OxmlCell {
        let t = self.table_mut();
        // 自动补齐
        while t.rows.len() <= row {
            let ncols = t.cols.len().max(1);
            t.rows.push(OxmlRow {
                height: Emu(0),
                cells: (0..ncols).map(|_| OxmlCell::default()).collect(),
                header: false,
            });
        }
        let r = &mut t.rows[row];
        while r.cells.len() <= col {
            r.cells.push(OxmlCell::default());
        }
        &mut r.cells[col]
    }

    /// 取 cell 不可变引用。
    pub fn cell(&self, row: usize, col: usize) -> Option<&OxmlCell> {
        let t = self.table();
        t.rows.get(row)?.cells.get(col)
    }

    /// 设列宽（覆盖 col 处的宽度）。
    pub fn set_column_width(&mut self, col: usize, w: Emu) -> crate::Result<()> {
        let t = self.table_mut();
        let c = t
            .cols
            .get_mut(col)
            .ok_or(crate::Error::IndexOutOfRange(col))?;
        c.width = w;
        Ok(())
    }

    /// 取列宽。
    pub fn column_width(&self, col: usize) -> Option<Emu> {
        self.table().cols.get(col).map(|c| c.width)
    }

    /// 设行高。
    pub fn set_row_height(&mut self, row: usize, h: Emu) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        r.height = h;
        Ok(())
    }

    /// 取行高。
    pub fn row_height(&self, row: usize) -> Option<Emu> {
        self.table().rows.get(row).map(|r| r.height)
    }

    /// 标记某行为表头（首行加粗等特殊样式）。
    pub fn set_header_row(&mut self, row: usize, is_header: bool) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        r.header = is_header;
        Ok(())
    }

    /// 行数。
    pub fn row_count(&self) -> usize {
        self.table().rows.len()
    }
    /// 列数。
    pub fn column_count(&self) -> usize {
        self.table().cols.len()
    }

    // --------------------- python-pptx 风格扩展 ---------------------

    /// 取整行（不可变）。
    pub fn row(&self, idx: usize) -> Option<&OxmlRow> {
        self.table().rows.get(idx)
    }
    /// 取整行（可变）。
    pub fn row_mut(&mut self, idx: usize) -> Option<&mut OxmlRow> {
        self.table_mut().rows.get_mut(idx)
    }
    /// 整列（不可变）。
    pub fn column(&self, idx: usize) -> Option<&OxmlCol> {
        self.table().cols.get(idx)
    }
    /// 整列（可变）。
    pub fn column_mut(&mut self, idx: usize) -> Option<&mut OxmlCol> {
        self.table_mut().cols.get_mut(idx)
    }

    /// 设 cell 填充色。
    ///
    /// 对应 python-pptx `cell.fill.solid(); cell.fill.fore_color.rgb = ...`。
    pub fn set_cell_fill(
        &mut self,
        row: usize,
        col: usize,
        c: crate::oxml::color::Color,
    ) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        let cidx = r
            .cells
            .get_mut(col)
            .ok_or(crate::Error::IndexOutOfRange(col))?;
        cidx.fill = c;
        Ok(())
    }

    /// 取 cell 填充色（克隆）。
    pub fn cell_fill(&self, row: usize, col: usize) -> Option<crate::oxml::color::Color> {
        self.cell(row, col).map(|c| c.fill.clone())
    }

    /// 设 cell 四向内边距（EMU，顺序 top/left/bottom/right）。
    pub fn set_cell_margins(
        &mut self,
        row: usize,
        col: usize,
        top: Emu,
        left: Emu,
        bottom: Emu,
        right: Emu,
    ) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        let c = r
            .cells
            .get_mut(col)
            .ok_or(crate::Error::IndexOutOfRange(col))?;
        c.margin = (Some(top), Some(left), Some(bottom), Some(right));
        Ok(())
    }

    /// 取 cell 文本帧可变引用（`cell.text_frame` 风格）。
    ///
    /// 通过 `cell_mut` 返回 [`OxmlCell`]，其 `.text` 字段就是 `TextBody`；
    /// 调用方需要进一步 `.text` 访问或用 `TextFrame::from(&mut c.text)` 包装。
    pub fn cell_text_frame_mut(&mut self, row: usize, col: usize) -> &mut TextBody {
        &mut self.cell_mut(row, col).text
    }

    /// 追加新行（返回行索引与可变引用）。
    ///
    /// 列数沿用现有 `cols.len()`。
    pub fn add_row(&mut self, height: Emu) -> &mut OxmlRow {
        let t = self.table_mut();
        let ncols = t.cols.len().max(1);
        t.rows.push(OxmlRow {
            height,
            cells: (0..ncols).map(|_| OxmlCell::default()).collect(),
            header: false,
        });
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = t.rows.len() - 1;
        &mut t.rows[idx]
    }

    /// 追加新列。
    pub fn add_column(&mut self, width: Emu) -> &mut OxmlCol {
        let t = self.table_mut();
        t.cols.push(OxmlCol { width });
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = t.cols.len() - 1;
        &mut t.cols[idx]
    }

    /// 删除指定行。
    ///
    /// 对标 python-pptx 中通过 XML 操作删除 `<a:tr>` 的能力。
    ///
    /// # 参数
    /// - `idx`：行索引（0-based）。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：`idx` 越界。
    pub fn remove_row(&mut self, idx: usize) -> crate::Result<()> {
        let t = self.table_mut();
        if idx >= t.rows.len() {
            return Err(crate::Error::IndexOutOfRange(idx));
        }
        t.rows.remove(idx);
        Ok(())
    }

    /// 删除指定列。
    ///
    /// 同时从每一行中移除对应位置的 cell。
    ///
    /// # 参数
    /// - `idx`：列索引（0-based）。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：`idx` 越界。
    pub fn remove_column(&mut self, idx: usize) -> crate::Result<()> {
        let t = self.table_mut();
        if idx >= t.cols.len() {
            return Err(crate::Error::IndexOutOfRange(idx));
        }
        t.cols.remove(idx);
        // 同步删除每行中对应位置的 cell
        for r in &mut t.rows {
            if idx < r.cells.len() {
                r.cells.remove(idx);
            }
        }
        Ok(())
    }

    /// 合并指定矩形区域的单元格。
    ///
    /// 对标 python-pptx `cell.merge(other_cell)`。
    ///
    /// # 参数
    /// - `row1, col1`：合并区域左上角；
    /// - `row2, col2`：合并区域右下角。
    ///
    /// # 行为
    /// - 左上角 cell 设为合并源（`grid_span` / `row_span`）；
    /// - 区域内其它 cell 设为 `h_merge` / `v_merge` 虚拟单元格；
    /// - 若区域只有 1×1 则为 no-op。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：索引越界；
    /// - [`crate::Error::Other`]：`row2 < row1` 或 `col2 < col1`。
    pub fn merge_cells(
        &mut self,
        row1: usize,
        col1: usize,
        row2: usize,
        col2: usize,
    ) -> crate::Result<()> {
        if row2 < row1 || col2 < col1 {
            return Err(crate::Error::Other(
                "merge_cells: row2/col2 不能小于 row1/col1".into(),
            ));
        }
        let nrows = self.row_count();
        let ncols = self.column_count();
        if row2 >= nrows || col2 >= ncols {
            return Err(crate::Error::IndexOutOfRange(row2.max(col2)));
        }
        let grid_span = (col2 - col1 + 1) as u32;
        let row_span = (row2 - row1 + 1) as u32;
        // 如果只有 1x1，无需合并
        if grid_span == 1 && row_span == 1 {
            return Ok(());
        }
        let t = self.table_mut();
        // 设置合并源（左上角）
        {
            let cell = &mut t.rows[row1].cells[col1];
            cell.grid_span = grid_span;
            cell.row_span = row_span;
            cell.h_merge = false;
            cell.v_merge = false;
        }
        // 设置被合并方（虚拟单元格）
        for r in row1..=row2 {
            for c in col1..=col2 {
                if r == row1 && c == col1 {
                    continue; // 跳过合并源
                }
                let cell = &mut t.rows[r].cells[c];
                cell.grid_span = 1;
                cell.row_span = 1;
                // hMerge 用于同一行内被合并的 cell
                cell.h_merge = r == row1;
                // vMerge 用于跨行被合并的 cell
                cell.v_merge = r != row1;
            }
        }
        Ok(())
    }

    /// 拆分单元格（TODO-029 高阶 API）。
    ///
    /// 把一个由 [`Self::merge_cells`] 合并出的"合并源"单元格还原为
    /// 多个独立单元格。与 `merge_cells` 是逆操作。
    ///
    /// # 参数
    /// - `row, col`：合并源单元格的位置（即合并时的左上角，必须是
    ///   `grid_span > 1` 或 `row_span > 1` 的"真实"单元格，不能是 h_merge/v_merge
    ///   的虚拟单元格）。
    ///
    /// # 行为
    /// - 把 `(row, col)` 的 `grid_span` / `row_span` 重置为 1；
    /// - 把合并区域内的所有虚拟单元格（`h_merge` / `v_merge` 为 true）解除虚拟状态；
    /// - 单元格的文本/填充/边框等属性不会被重置——只会修改合并相关字段。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：`row` / `col` 越界；
    /// - [`crate::Error::Other`]：目标单元格不是合并源（`grid_span == 1 && row_span == 1`），
    ///   或目标单元格是虚拟单元格（`h_merge` / `v_merge` 为 true）。
    pub fn split_cell(&mut self, row: usize, col: usize) -> crate::Result<()> {
        let nrows = self.row_count();
        let ncols = self.column_count();
        if row >= nrows || col >= ncols {
            return Err(crate::Error::IndexOutOfRange(row.max(col)));
        }
        let t = self.table_mut();
        let cell = &t.rows[row].cells[col];
        // 不能拆分虚拟单元格
        if cell.h_merge || cell.v_merge {
            return Err(crate::Error::Other(
                "split_cell: 目标单元格是虚拟单元格，请对合并源调用 split_cell".into(),
            ));
        }
        let grid_span = cell.grid_span;
        let row_span = cell.row_span;
        // 不是合并源
        if grid_span <= 1 && row_span <= 1 {
            return Err(crate::Error::Other(
                "split_cell: 目标单元格不是合并源（grid_span 和 row_span 均为 1）".into(),
            ));
        }
        // 合并区域右下角
        let row2 = row + row_span as usize - 1;
        let col2 = col + grid_span as usize - 1;
        // 重置合并源
        {
            let cell = &mut t.rows[row].cells[col];
            cell.grid_span = 1;
            cell.row_span = 1;
            cell.h_merge = false;
            cell.v_merge = false;
        }
        // 解除虚拟单元格状态
        for r in row..=row2 {
            for c in col..=col2 {
                if r == row && c == col {
                    continue; // 跳过原合并源
                }
                let cell = &mut t.rows[r].cells[c];
                cell.grid_span = 1;
                cell.row_span = 1;
                cell.h_merge = false;
                cell.v_merge = false;
            }
        }
        Ok(())
    }

    /// 设置单元格边框。
    ///
    /// 对标 python-pptx `cell.border_top` / `border_bottom` / `border_left` / `border_right`。
    ///
    /// # 参数
    /// - `row, col`：单元格位置；
    /// - `side`：边框方向（[`BorderSide`]）；
    /// - `width`：边框宽度（EMU）；
    /// - `color`：边框颜色（`Color::None` 表示使用主题继承）；
    /// - `no_fill`：是否写出 `<a:noFill/>`（无填充边框）。
    ///
    /// # 错误
    /// - [`crate::Error::IndexOutOfRange`]：`row` 或 `col` 越界。
    pub fn set_cell_border(
        &mut self,
        row: usize,
        col: usize,
        side: BorderSide,
        width: Emu,
        color: crate::oxml::color::Color,
        no_fill: bool,
    ) -> crate::Result<()> {
        let t = self.table_mut();
        let r = t
            .rows
            .get_mut(row)
            .ok_or(crate::Error::IndexOutOfRange(row))?;
        let c = r
            .cells
            .get_mut(col)
            .ok_or(crate::Error::IndexOutOfRange(col))?;
        let border = crate::oxml::table::CellBorder {
            color,
            width,
            no_fill,
        };
        match side {
            BorderSide::Left => c.border_left = Some(border),
            BorderSide::Right => c.border_right = Some(border),
            BorderSide::Top => c.border_top = Some(border),
            BorderSide::Bottom => c.border_bottom = Some(border),
        }
        Ok(())
    }

    /// 取单元格边框（克隆）。
    pub fn cell_border(
        &self,
        row: usize,
        col: usize,
        side: BorderSide,
    ) -> Option<crate::oxml::table::CellBorder> {
        let c = self.cell(row, col)?;
        match side {
            BorderSide::Left => c.border_left.clone(),
            BorderSide::Right => c.border_right.clone(),
            BorderSide::Top => c.border_top.clone(),
            BorderSide::Bottom => c.border_bottom.clone(),
        }
    }

    // --------------------- 占位符（TODO-007 表格占位符类型化填充） ---------------------

    /// 将本表格形状标记为占位符（TODO-007 表格占位符类型化填充）。
    ///
    /// 写出 XML 时会在 `<p:nvGraphicFramePr>/<p:nvPr>` 内插入
    /// `<p:ph type="tbl" idx="..."/>`，使 PowerPoint 把该 graphicFrame
    /// 识别为表格占位符的填充实例。
    ///
    /// # 参数
    /// - `ph_idx`：占位符索引（对应 `<p:ph idx="..."/>`）。
    /// - `ph_type`：占位符类型字符串（如 `"tbl"` / `"obj"`），`None` 时省略 `type` 属性。
    pub fn set_placeholder(&mut self, ph_idx: u32, ph_type: Option<&str>) {
        self.frame.is_placeholder = true;
        self.frame.ph_idx = Some(ph_idx);
        self.frame.ph_type = ph_type.map(|s| s.to_string());
    }

    /// 清除占位符标记，使本表格形状变为普通 graphicFrame。
    pub fn clear_placeholder(&mut self) {
        self.frame.is_placeholder = false;
        self.frame.ph_idx = None;
        self.frame.ph_type = None;
    }

    /// 是否被标记为占位符。
    pub fn is_placeholder(&self) -> bool {
        self.frame.is_placeholder
    }

    /// 占位符索引（若已标记）。
    pub fn ph_idx(&self) -> Option<u32> {
        self.frame.ph_idx
    }

    /// 占位符类型字符串（若已标记）。
    pub fn ph_type(&self) -> Option<&str> {
        self.frame.ph_type.as_deref()
    }
}

/// 边框方向枚举。
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderSide {
    /// 左边框（`<a:lnL>`）。
    Left,
    /// 右边框（`<a:lnR>`）。
    Right,
    /// 上边框（`<a:lnT>`）。
    Top,
    /// 下边框（`<a:lnB>`）。
    Bottom,
}

impl Shape for TableShape {
    fn id(&self) -> u32 {
        self.frame.id
    }
    fn set_id(&mut self, id: u32) {
        self.frame.id = id;
    }
    fn name(&self) -> &str {
        &self.frame.name
    }
    fn set_name(&mut self, name: String) {
        self.frame.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "table"
    }

    fn left(&self) -> Emu {
        self.frame.properties.xfrm.off_x.unwrap_or_default()
    }
    fn set_left(&mut self, emu: Emu) {
        self.frame.properties.xfrm.off_x = Some(emu);
    }
    fn top(&self) -> Emu {
        self.frame.properties.xfrm.off_y.unwrap_or_default()
    }
    fn set_top(&mut self, emu: Emu) {
        self.frame.properties.xfrm.off_y = Some(emu);
    }
    fn width(&self) -> Emu {
        self.frame.properties.xfrm.ext_cx.unwrap_or_default()
    }
    fn set_width(&mut self, emu: Emu) {
        self.frame.properties.xfrm.ext_cx = Some(emu);
    }
    fn height(&self) -> Emu {
        self.frame.properties.xfrm.ext_cy.unwrap_or_default()
    }
    fn set_height(&mut self, emu: Emu) {
        self.frame.properties.xfrm.ext_cy = Some(emu);
    }

    /// 表格不支持旋转（OOXML 规范）。调用 [`Shape::set_rotation`] 会被忽略。
    fn rotation(&self) -> f64 {
        0.0
    }
    fn set_rotation(&mut self, _deg: f64) {}
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::color::Color;
    use crate::units::{Emu, RGBColor};

    /// `merge_cells` 正确设置 gridSpan/rowSpan/hMerge/vMerge。
    #[test]
    fn merge_cells_sets_attributes() {
        let mut t = TableShape::new(3, 3, Emu(1000), Emu(500));
        // 合并 (0,0) 到 (1,2) —— 2 行 3 列
        t.merge_cells(0, 0, 1, 2).unwrap();
        // 合并源 (0,0)
        let origin = t.cell(0, 0).unwrap();
        assert_eq!(origin.grid_span, 3);
        assert_eq!(origin.row_span, 2);
        assert!(!origin.h_merge);
        assert!(!origin.v_merge);
        // 同行被合并方 (0,1) / (0,2) —— hMerge
        let h1 = t.cell(0, 1).unwrap();
        assert!(h1.h_merge);
        assert!(!h1.v_merge);
        // 跨行被合并方 (1,0) / (1,1) / (1,2) —— vMerge
        let v1 = t.cell(1, 0).unwrap();
        assert!(!v1.h_merge);
        assert!(v1.v_merge);
    }

    /// `merge_cells` 1×1 为 no-op。
    #[test]
    fn merge_cells_1x1_is_noop() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        t.merge_cells(0, 0, 0, 0).unwrap();
        let c = t.cell(0, 0).unwrap();
        assert_eq!(c.grid_span, 0); // 默认值 0
        assert_eq!(c.row_span, 0);
    }

    /// `merge_cells` 越界返回错误。
    #[test]
    fn merge_cells_out_of_bounds() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        assert!(t.merge_cells(0, 0, 5, 5).is_err());
        assert!(t.merge_cells(1, 0, 0, 0).is_err()); // row2 < row1
    }

    /// `set_cell_border` 正确设置边框。
    #[test]
    fn set_cell_border_works() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        t.set_cell_border(
            0,
            0,
            BorderSide::Top,
            Emu(9525),
            Color::RGB(RGBColor::RED),
            false,
        )
        .unwrap();
        let b = t.cell_border(0, 0, BorderSide::Top).unwrap();
        assert_eq!(b.width.value(), 9525);
        assert!(!b.no_fill);
    }

    /// `tblLook` 高阶 API 正确设置和读取布尔属性。
    ///
    /// 这是 TODO-028 的测试：验证 first_row / last_row / first_column / last_column
    /// / horz_banding / vert_banding 的 getter/setter 行为。
    #[test]
    fn tbl_look_boolean_attributes() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        // 默认值（与 PowerPoint 一致）
        assert!(t.first_row(), "默认 firstRow=true");
        assert!(!t.last_row(), "默认 lastRow=false");
        assert!(t.first_column(), "默认 firstColumn=true");
        assert!(!t.last_column(), "默认 lastColumn=false");
        assert!(t.horz_banding(), "默认 horz_banding=true");
        assert!(!t.vert_banding(), "默认 vert_banding=false");

        // 修改值
        t.set_first_row(false);
        t.set_last_row(true);
        t.set_first_column(false);
        t.set_last_column(true);
        t.set_horz_banding(false);
        t.set_vert_banding(true);

        // 验证
        assert!(!t.first_row());
        assert!(t.last_row());
        assert!(!t.first_column());
        assert!(t.last_column());
        assert!(!t.horz_banding());
        assert!(t.vert_banding());

        // 验证底层 oxml 字段也正确更新
        let lk = &t.table().tbl_look;
        assert!(!lk.first_row);
        assert!(lk.last_row);
        assert!(!lk.first_column);
        assert!(lk.last_column);
        assert!(lk.no_h_band, "horz_banding=false → noHBand=true");
        assert!(!lk.no_v_band, "vert_banding=true → noVBand=false");
    }

    /// `remove_row` / `remove_column` 正确删除。
    #[test]
    fn remove_row_and_column() {
        let mut t = TableShape::new(3, 3, Emu(1000), Emu(500));
        t.remove_row(1).unwrap();
        assert_eq!(t.row_count(), 2);
        t.remove_column(1).unwrap();
        assert_eq!(t.column_count(), 2);
        // 每行 cell 数也应同步减少
        assert_eq!(t.table().rows[0].cells.len(), 2);
    }

    /// `remove_row` 越界返回错误。
    #[test]
    fn remove_row_out_of_bounds() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        assert!(t.remove_row(5).is_err());
        assert!(t.remove_column(5).is_err());
    }

    // --------------------- 表格样式测试（TODO-030） ---------------------

    /// `set_style` 正确设置内置样式。
    #[test]
    fn set_style_builtin() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        assert!(t.set_style("Medium Style 2 - Accent 1"));
        let style = t.table_style().expect("style 应已设置");
        assert_eq!(style.style_id(), "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        assert_eq!(style.style_name(), Some("Medium Style 2 - Accent 1"));
    }

    /// `set_style` 对未知名称返回 `false`。
    #[test]
    fn set_style_unknown_returns_false() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        assert!(!t.set_style("Nonexistent Style"));
        assert!(t.table_style().is_none());
    }

    /// `set_style_id` 用原始 GUID 设置样式。
    #[test]
    fn set_style_id_raw_guid() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        t.set_style_id("{5940675A-B579-460E-94D1-54222C63F5DA}");
        let style = t.table_style().expect("style 应已设置");
        assert_eq!(style.style_id(), "{5940675A-B579-460E-94D1-54222C63F5DA}");
    }

    /// `clear_style` 清除样式引用。
    #[test]
    fn clear_style_works() {
        let mut t = TableShape::new(2, 2, Emu(1000), Emu(500));
        t.set_style("Medium Style 2 - Accent 1");
        assert!(t.table_style().is_some());
        t.clear_style();
        assert!(t.table_style().is_none());
    }

    /// 表格样式序列化正确写出 `<a:tableStyleId>`。
    #[test]
    fn table_style_serializes() {
        let mut t = TableShape::new(1, 1, Emu(1000), Emu(500));
        t.set_style("No Style, Table Grid");
        let mut w = crate::oxml::writer::XmlWriter::new();
        t.table().write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains("<a:tableStyleId>"));
        assert!(xml.contains("{5940675A-B579-460E-94D1-54222C63F5DA}"));
        assert!(xml.contains("</a:tableStyleId>"));
    }
}
