//! # 备注母版（Notes Master）—— 高阶 API（TODO-045）。
//!
//! 对应 OOXML 规范中的 `<p:notesMaster>` 元素，以及 python-pptx 中
//! `NotesMaster` / `NotesMasterPart`。
//!
//! # 概念
//!
//! 备注母版是所有备注页（`<p:notesSlide>`）的"模板"——
//! 它定义了备注页的默认占位符（日期 / 页码 / 正文 / 幻灯片图像）与文本样式。
//! 一个演示文稿通常只有 1 个备注母版。
//!
//! ```text
//!   NotesMaster  (备注母版：定义备注页的默认布局)
//!        ↑ 引用
//!   NotesSlide   (备注页：每张 slide 的演讲者备注)
//! ```
//!
//! # 当前实现范围
//!
//! 本模块是**只读访问**——支持从已存在 `.pptx` 解析出备注母版并查询其形状。
//! 写路径（创建/修改备注母版并持久化）暂未实现，与 python-pptx 行为一致
//! （python-pptx 也仅暴露 `presentation.notes_master` 的只读访问）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.NotesMaster` ←→ [`NotesMasterRef`]；
//! - `presentation.notes_master` ←→ [`crate::presentation::Presentation::notes_master`]。

use std::cell::RefCell;
use std::rc::Rc;

use crate::oxml::notesmaster::NotesMaster as OxmlNotesMaster;
use crate::oxml::shape::Sp as OxmlSp;

/// 单个备注母版引用（高阶句柄）。
///
/// 与 [`crate::slide_masters::SlideMasterRef`] 同样使用 `Rc<RefCell<OxmlNotesMaster>>`
/// 共享 oxml 模型，方便在多视图间共享且通过编译期借用检查。
#[derive(Debug, Clone)]
pub struct NotesMasterRef {
    /// 在所属 [`NotesMasters`] 中的索引。
    #[allow(dead_code)]
    pub(crate) idx: usize,
    /// OPC part 路径（`/ppt/notesMasters/notesMasterN.xml`）。
    pub(crate) partname: String,
    /// 关系 id（在 `presentation.xml.rels` 中使用）。
    pub(crate) rid: String,
    /// 内部 oxml 模型。
    pub(crate) oxml: Rc<RefCell<OxmlNotesMaster>>,
}

impl NotesMasterRef {
    /// 取出 part 路径（如 `/ppt/notesMasters/notesMaster1.xml`）。
    pub fn partname(&self) -> &str {
        &self.partname
    }
    /// 取出关系 id（如 `rIdNotesMaster1`）。
    pub fn rid(&self) -> &str {
        &self.rid
    }

    /// shape 不可变快照（python-pptx `notes_master.shapes` 风格）。
    ///
    /// 返回克隆后的形状列表，调用方可自由操作而不影响母版内部状态。
    pub fn shapes(&self) -> Vec<OxmlSp> {
        self.oxml.borrow().shapes.clone()
    }

    /// shape 可变视图（返回 `RefMut`）。
    ///
    /// 通过 `RefMut` 可以原地修改母版内的形状；修改不会自动持久化到
    /// `.pptx` 文件（写路径暂未实现）。
    pub fn shapes_mut(&self) -> std::cell::RefMut<'_, Vec<OxmlSp>> {
        std::cell::RefMut::map(self.oxml.borrow_mut(), |m| &mut m.shapes)
    }

    /// 备注母版中的形状数量。
    pub fn len(&self) -> usize {
        self.oxml.borrow().shapes.len()
    }

    /// 是否无形状。
    pub fn is_empty(&self) -> bool {
        self.oxml.borrow().shapes.is_empty()
    }

    /// 读取备注母版背景。`None` 表示未设置独立背景。
    pub fn background(&self) -> Option<crate::oxml::slide::SlideBackground> {
        self.oxml.borrow().background.clone()
    }
}

/// 全部备注母版的集合。
///
/// 在 [`crate::presentation::Presentation`] 中由 `notes_masters` 字段持有。
/// 一个演示文稿通常只有 0 或 1 个备注母版。
#[derive(Debug, Default, Clone)]
pub struct NotesMasters {
    pub(crate) items: Vec<NotesMasterRef>,
}

impl NotesMasters {
    /// 新建一个空集合。
    pub fn new() -> Self {
        NotesMasters::default()
    }
    /// 数量。
    pub fn len(&self) -> usize {
        self.items.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }
    /// 遍历所有备注母版引用。
    pub fn iter(&self) -> std::slice::Iter<'_, NotesMasterRef> {
        self.items.iter()
    }
    /// 按索引取不可变引用。
    pub fn get(&self, idx: usize) -> Option<&NotesMasterRef> {
        self.items.get(idx)
    }

    /// 取第一个备注母版（演示文稿通常只有 1 个）。
    ///
    /// `None` 表示该演示文稿无备注母版（`Presentation::new` 创建的空白文档无此 part）。
    pub fn first(&self) -> Option<&NotesMasterRef> {
        self.items.first()
    }

    /// 追加一个备注母版（仅内存模型；写路径暂未实现）。
    pub fn push(&mut self, master: NotesMasterRef) {
        self.items.push(master);
    }
}
