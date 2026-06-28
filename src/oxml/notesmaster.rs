//! 备注母版 `<p:notesMaster>` 简化模型（TODO-045）。
//!
//! 备注母版是"主-版-页"三层中的备注页对应母版——所有 notesSlide
//! 默认继承它的形状与样式。对应 python-pptx 中 `NotesMaster` / `NotesMasterPart`。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.parts.notesmaster.NotesMasterPart` ←→ 本模块 [`NotesMaster`] 的 OPC 宿主；
//! - `pptx.NotesMaster` ←→ [`crate::notes_masters::NotesMasterRef`] 高阶句柄。
//!
//! # 当前实现范围
//!
//! 本模块是**极简只读模型**，仅承载解析出的 shapes 列表 + partname/rid 元数据。
//! 完整的 `to_xml` 写路径未实现（写场景由 python-pptx 也较少触达）。
//!
//! # OOXML 结构（参考）
//!
//! ```text
//! <p:notesMaster>
//!   <p:cSld>
//!     <p:spTree>...</p:spTree>          形状树（含占位符：日期/页码/正文等）
//!   </p:cSld>
//!   <p:clrMap .../>                      颜色映射
//!   <p:notesStyle>...</p:notesStyle>     备注文本样式
//! </p:notesMaster>
//! ```

use crate::oxml::shape::Sp;
use crate::oxml::slide::SlideBackground;

/// `<p:notesMaster>` 的内存模型（极简版）。
///
/// 仅承载 `cSld/spTree` 内的形状列表 + 可选背景，便于读取路径还原
/// 已有 .pptx 的备注母版内容。写路径（`to_xml`）暂未实现。
#[derive(Clone, Debug, Default)]
pub struct NotesMaster {
    /// 备注母版中的形状列表（`<p:cSld>/<p:spTree>` 内的 `<p:sp>`）。
    pub shapes: Vec<Sp>,
    /// 备注母版背景（`<p:bg>`，可选）。`None` 表示使用默认背景。
    pub background: Option<SlideBackground>,
}

impl NotesMaster {
    /// 创建一个空的备注母版。
    pub fn new() -> Self {
        Self::default()
    }

    /// 形状数量。
    pub fn len(&self) -> usize {
        self.shapes.len()
    }

    /// 是否无形状。
    pub fn is_empty(&self) -> bool {
        self.shapes.is_empty()
    }
}
