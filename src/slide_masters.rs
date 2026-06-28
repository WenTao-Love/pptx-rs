//! # 幻灯片母版（Slide Master）—— 高阶 API
//!
//! 对应 OOXML 规范中的 `<p:sldMaster>` 元素。
//!
//! # 概念
//!
//! 母版是 PowerPoint "主-版-页"三层中的最上层：
//!
//! ```text
//!   SlideMaster  (母版：定义全局主题、占位符、全局背景)
//!        ↑ 引用
//!   SlideLayout  (版式：基于母版，扩展为不同页面模板)
//!        ↑ 引用
//!   Slide        (页：实际内容)
//! ```
//!
//! # 当前实现范围
//!
//! 本模块是**极简实现**，仅暴露 partname / rid / shapes 等元数据。
//! 完整读取/编辑母版内主题、占位符、形状等是路线图任务。
//!
//! 完整 API 设计可参考 python-pptx 的 `SlideMasters` / `SlideMaster` 类。

use std::cell::RefCell;
use std::rc::Rc;

use crate::oxml::shape::Sp as OxmlSp;
use crate::oxml::slidemaster::SldMaster as OxmlSldMaster;

/// 单个母版引用。
///
/// 与 [`crate::slide_layouts::SlideLayoutRef`] 同样使用 `Rc<RefCell<OxmlSldMaster>>`
/// 共享 oxml 模型，方便母版与版式双向引用时仍能通过编译期借用检查。
#[derive(Debug, Clone)]
pub struct SlideMasterRef {
    /// 在所属 `SlideMasters` 中的索引。
    #[allow(dead_code)]
    pub(crate) idx: usize,
    /// OPC part 路径（`/ppt/slideMasters/slideMasterN.xml`）。
    pub(crate) partname: String,
    /// 关系 id（在 `presentation.xml.rels` 中使用）。
    pub(crate) rid: String,
    /// 内部 oxml 模型。
    pub(crate) oxml: Rc<RefCell<OxmlSldMaster>>,
}

impl SlideMasterRef {
    /// 取出 part 路径（如 `/ppt/slideMasters/slideMaster1.xml`）。
    pub fn partname(&self) -> &str {
        &self.partname
    }
    /// 取出关系 id（如 `rIdMaster1`）。
    pub fn rid(&self) -> &str {
        &self.rid
    }

    /// shape 不可变快照（python-pptx `slide_master.shapes` 风格）。
    pub fn shapes(&self) -> Vec<OxmlSp> {
        self.oxml.borrow().shapes.clone()
    }
    /// shape 可变视图（返回 `RefMut`）。
    pub fn shapes_mut(&self) -> std::cell::RefMut<'_, Vec<OxmlSp>> {
        std::cell::RefMut::map(self.oxml.borrow_mut(), |s| &mut s.shapes)
    }

    /// 占位符列表（母版的占位符会被版式继承）。
    pub fn placeholders(&self) -> Vec<crate::slide_layouts::Placeholder> {
        self.oxml
            .borrow()
            .shapes
            .iter()
            .filter(|s| s.is_placeholder)
            .map(|s| crate::slide_layouts::Placeholder {
                idx: s.ph_idx.unwrap_or(0),
                ph_type: s.ph_type.clone().unwrap_or_else(|| "body".to_string()),
                name: s.name.clone(),
            })
            .collect()
    }

    // --------------------- 背景编辑 API（TODO-049 高阶） ---------------------
    //
    // 对标 python-pptx `slide_master.background`。母版背景会被所有未设置
    // 独立背景的 slide/layout 继承（OOXML 顺序：`<p:cSld>/<p:bg>` 在 `<p:spTree>` 之前）。

    /// 读取母版背景的可变引用。`None` 表示未设置独立背景。
    pub fn background(&self) -> Option<crate::oxml::slide::SlideBackground> {
        self.oxml.borrow().background.clone()
    }

    /// 设置母版背景。`bg = None` 等价于 [`Self::clear_background`]。
    pub fn set_background(&self, bg: Option<crate::oxml::slide::SlideBackground>) {
        self.oxml.borrow_mut().background = bg;
    }

    /// 设置母版背景为纯色（便捷方法）。
    ///
    /// 对标 python-pptx `slide_master.background.fill.solid()` +
    /// `slide_master.background.fill.fore_color.rgb = ...`。
    pub fn set_background_solid(&self, color: crate::oxml::color::Color) {
        self.oxml.borrow_mut().background = Some(crate::oxml::slide::SlideBackground::solid(color));
    }

    /// 清除母版背景（让母版走默认背景）。
    pub fn clear_background(&self) {
        self.oxml.borrow_mut().background = None;
    }

    /// 追加一个 shape 到母版 spTree 末尾。
    ///
    /// 对标 python-pptx `slide_master.shapes._spTree.append(sp)`。
    /// 调用方需自行保证 `sp.id` 在母版内唯一。
    pub fn add_shape(&self, sp: OxmlSp) {
        self.oxml.borrow_mut().shapes.push(sp);
    }

    /// 移除母版中指定 ID 的 shape，返回被移除的 shape。`None` 表示未找到。
    pub fn remove_shape(&self, id: u32) -> Option<OxmlSp> {
        let mut oxml = self.oxml.borrow_mut();
        let pos = oxml.shapes.iter().position(|s| s.id == id)?;
        Some(oxml.shapes.remove(pos))
    }
}

/// 全部母版的集合。
///
/// 在 [`crate::presentation::Presentation`] 中由 `slide_masters` 字段持有。
/// 至少包含 1 个默认母版（由 [`crate::presentation::Presentation::new`] 自动创建）。
#[derive(Debug, Default, Clone)]
pub struct SlideMasters {
    pub(crate) items: Vec<SlideMasterRef>,
}

impl SlideMasters {
    /// 新建一个空集合。
    pub fn new() -> Self {
        SlideMasters::default()
    }
    /// 数量。
    pub fn len(&self) -> usize {
        self.items.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }
    /// 遍历所有母版引用。
    pub fn iter(&self) -> std::slice::Iter<'_, SlideMasterRef> {
        self.items.iter()
    }
    /// 按索引取不可变引用。
    pub fn get(&self, idx: usize) -> Option<&SlideMasterRef> {
        self.items.get(idx)
    }
    /// 按索引取可变引用。
    pub fn get_mut(&mut self, idx: usize) -> Option<&mut SlideMasterRef> {
        self.items.get_mut(idx)
    }

    /// 取一个母版（克隆为不可变轻量句柄 [`SlideMaster`]）。
    ///
    /// 当前实现是空结构体 —— 真正的母版内容编辑需要走 [`SlideMasterRef`]。
    pub fn at(&self, idx: usize) -> Option<SlideMaster> {
        self.items.get(idx).map(|_| SlideMaster {})
    }

    /// 追加一个母版（**仅**内存模型；`presentation::to_opc_package` 会一并写出）。
    pub fn push(&mut self, master: SlideMasterRef) {
        self.items.push(master);
    }
}

/// 母版（不可变视图）。
///
/// 当前为占位空结构体；后续会扩展为包含主题/占位符/背景的完整模型，
/// 类似 python-pptx 的 `SlideMaster` 类。
#[derive(Debug, Clone)]
pub struct SlideMaster {}
