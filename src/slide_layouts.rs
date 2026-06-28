//! # 幻灯片版式（Slide Layout）—— 高阶 API
//!
//! 对应 OOXML 规范中的 `<p:sldLayout>` 元素。
//!
//! # 概念
//!
//! "版式"是 PowerPoint 的核心抽象之一，位于"主-版-页"三层中的中间层：
//!
//! ```text
//!   SlideMaster  (母版：全局默认样式)
//!        ↑ 引用
//!   SlideLayout  (版式：每页的"页面模板"，可指定占位符位置)
//!        ↑ 引用
//!   Slide        (页：实际内容)
//! ```
//!
//! # 当前实现范围
//!
//! 本模块是**极简实现**：
//!
//! - 仅暴露 name / partname / rid / placeholders / shapes 等元数据；
//! - 完整读取/编辑版式内的 placeholder 是路线图任务。
//!
//! 完整 API 设计可参考 python-pptx 的 `SlideLayouts` / `SlideLayout` 类。

use std::cell::RefCell;
use std::rc::Rc;

use crate::oxml::shape::Sp as OxmlSp;
use crate::oxml::slidelayout::SldLayout as OxmlSldLayout;

/// 单个版式的引用。
///
/// 设计上使用 `Rc<RefCell<OxmlSldLayout>>` 而**非**直接拥有，以兼容"版式被母版
/// 与页共享引用"的未来场景。当前实现下每个版式只被一个 [`SlideLayouts`] 拥有。
#[derive(Debug, Clone)]
pub struct SlideLayoutRef {
    /// 在所属 `SlideLayouts` 中的索引（用于 `get(idx)` / `at(idx)`）。
    #[allow(dead_code)]
    pub(crate) idx: usize,
    /// OPC part 路径（`/ppt/slideLayouts/slideLayoutN.xml`）。
    pub(crate) partname: String,
    /// 关系 id（在所属 `SlideMaster` 的 `.rels` 中使用）。
    pub(crate) rid: String,
    /// 内部 oxml 模型（共享，以便母版与页可同时访问）。
    pub(crate) oxml: Rc<RefCell<OxmlSldLayout>>,
}

impl SlideLayoutRef {
    /// 取出 part 路径（如 `/ppt/slideLayouts/slideLayout1.xml`）。
    pub fn partname(&self) -> &str {
        &self.partname
    }
    /// 取出关系 id（如 `rIdLayout1`）。
    pub fn rid(&self) -> &str {
        &self.rid
    }
    /// 取版式名（对应 OOXML 中的 `<p:sldLayout name="..."/>`）。
    pub fn name(&self) -> String {
        self.oxml.borrow().name.clone()
    }
    /// 设置版式名。
    pub fn set_name(&mut self, n: String) {
        self.oxml.borrow_mut().name = n;
    }
    /// 取版式类型（OOXML 中的 `type` 属性，可选；如 `blank` / `title` / `section`）。
    pub fn layout_type(&self) -> String {
        self.oxml.borrow().type_.clone()
    }
    /// 设置版式类型。
    pub fn set_layout_type(&mut self, t: String) {
        self.oxml.borrow_mut().type_ = t;
    }

    /// 不可变 shape 列表（python-pptx `slide_layout.shapes` 风格）。
    ///
    /// 返回的是 `OxmlSp` 句柄的借用，不触发 clone；适合"只读浏览"。
    pub fn shapes(&self) -> Vec<OxmlSp> {
        self.oxml.borrow().shapes.clone()
    }

    /// 可变 shape 列表。
    pub fn shapes_mut(&self) -> std::cell::RefMut<'_, Vec<OxmlSp>> {
        std::cell::RefMut::map(self.oxml.borrow_mut(), |s| &mut s.shapes)
    }

    /// 占位符列表（python-pptx `slide_layout.placeholders` 风格）。
    ///
    /// 当前实现把所有带 `is_placeholder=true` 的 [`OxmlSp`] 视作占位符。
    /// 返回不可变快照，避免借用期跨越。
    pub fn placeholders(&self) -> Vec<Placeholder> {
        self.oxml
            .borrow()
            .shapes
            .iter()
            .filter(|s| s.is_placeholder)
            .map(|s| Placeholder {
                idx: s.ph_idx.unwrap_or(0),
                ph_type: s.ph_type.clone().unwrap_or_else(|| "body".to_string()),
                name: s.name.clone(),
            })
            .collect()
    }

    /// 返回所有占位符在 `shapes` 中的索引列表（TODO-008）。
    ///
    /// 用于配合 [`shapes_mut`] 按索引修改占位符：
    ///
    /// ```no_run
    /// # use pptx::Presentation;
    /// # let mut p = Presentation::new().unwrap();
    /// # let layout = p.slide_layouts().get(0).unwrap().clone();
    /// let indices = layout.placeholder_indices();
    /// let mut shapes = layout.shapes_mut();
    /// for i in indices {
    ///     // 修改第 i 个 shape（它是占位符）
    ///     shapes[i].name = format!("PH {}", i);
    /// }
    /// ```
    pub fn placeholder_indices(&self) -> Vec<usize> {
        self.oxml
            .borrow()
            .shapes
            .iter()
            .enumerate()
            .filter(|(_, s)| s.is_placeholder)
            .map(|(i, _)| i)
            .collect()
    }

    /// 按 `ph_idx` 查找占位符在 `shapes` 中的索引（TODO-008）。
    ///
    /// 返回第一个 `ph_idx` 匹配的 shapes 索引；未找到返回 `None`。
    /// 用于配合 [`shapes_mut`] 修改特定占位符。
    pub fn placeholder_index_by_ph_idx(&self, ph_idx: u32) -> Option<usize> {
        self.oxml
            .borrow()
            .shapes
            .iter()
            .enumerate()
            .find(|(_, s)| s.is_placeholder && s.ph_idx == Some(ph_idx))
            .map(|(i, _)| i)
    }
}

/// 占位符（python-pptx `slide_layout.placeholders[i]` 风格）。
///
/// 仅承载**只读元数据**；如需 mutate 形状本身，请用 [`SlideLayoutRef::shapes_mut`]
/// 再按 `ph_idx` 过滤。
#[derive(Debug, Clone)]
pub struct Placeholder {
    /// 占位符 idx（对应 `<p:ph idx="N"/>`）。
    pub idx: u32,
    /// 占位符类型（对应 `<p:ph type="..."/>`，如 `title` / `body`）。
    pub ph_type: String,
    /// 形状名（cNvPr name）。
    pub name: String,
}

impl Placeholder {
    /// python-pptx 风格 `placeholder.placeholder_format.type` 的简化版。
    pub fn ph_type(&self) -> &str {
        &self.ph_type
    }
    /// python-pptx 风格 `placeholder.placeholder_format.idx`。
    pub fn idx(&self) -> u32 {
        self.idx
    }
    /// 形状名。
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// 全部版式的集合。
///
/// 在 [`crate::presentation::Presentation`] 中由 `slide_layouts` 字段持有。
/// 至少包含 1 个默认版式（由 [`crate::presentation::Presentation::new`] 自动创建）。
#[derive(Debug, Default, Clone)]
pub struct SlideLayouts {
    pub(crate) items: Vec<SlideLayoutRef>,
}

impl SlideLayouts {
    /// 新建一个空集合。
    pub fn new() -> Self {
        SlideLayouts::default()
    }
    /// 数量。
    pub fn len(&self) -> usize {
        self.items.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }
    /// 遍历所有版式引用。
    pub fn iter(&self) -> std::slice::Iter<'_, SlideLayoutRef> {
        self.items.iter()
    }
    /// 按索引取不可变引用。
    pub fn get(&self, idx: usize) -> Option<&SlideLayoutRef> {
        self.items.get(idx)
    }
    /// 按索引取可变引用。
    pub fn get_mut(&mut self, idx: usize) -> Option<&mut SlideLayoutRef> {
        self.items.get_mut(idx)
    }

    /// 取一个版式（克隆为不可变轻量句柄 [`SlideLayout`]）。
    ///
    /// `idx` 越界返回 `None`。该 API 是 python-pptx `slide_layouts[i]` 的等价物。
    pub fn at(&self, idx: usize) -> Option<SlideLayout> {
        self.items.get(idx).map(|r| SlideLayout {
            name: r.oxml.borrow().name.clone(),
            r#type: r.oxml.borrow().type_.clone(),
        })
    }

    /// 追加一个版式（**仅**内存模型；`presentation::to_opc_package` 会一并写出）。
    ///
    /// `partname` 必须以 `/ppt/slideLayouts/slideLayout<N>.xml` 形式。
    pub fn push(&mut self, layout: SlideLayoutRef) {
        self.items.push(layout);
    }

    /// 按索引移除一个版式（TODO-008）。
    ///
    /// 对标 python-pptx `SlideLayouts.remove(layout)`。
    ///
    /// # 注意
    /// 移除版式**不会**自动更新引用该版式的 slide。调用方应确保没有 slide
    /// 仍在引用被移除的版式（可通过 [`crate::presentation::Presentation::slides_using_layout`]
    /// 检查）。
    ///
    /// # 返回
    /// - `Some(SlideLayoutRef)`：被移除的版式；
    /// - `None`：索引越界。
    pub fn remove(&mut self, idx: usize) -> Option<SlideLayoutRef> {
        if idx < self.items.len() {
            Some(self.items.remove(idx))
        } else {
            None
        }
    }

    /// 按关系 id 查找版式索引（TODO-008）。
    ///
    /// 对标 python-pptx `SlideLayouts.index(layout)`。
    /// 返回第一个 `rid` 匹配的索引；未找到返回 `None`。
    pub fn index_of(&self, rid: &str) -> Option<usize> {
        self.items.iter().position(|l| l.rid == rid)
    }

    /// 按名称查找版式（TODO-008）。
    ///
    /// 对标 python-pptx `SlideLayouts.get_by_name(name)`。
    /// 返回第一个 `name` 匹配的版式引用；未找到返回 `None`。
    pub fn get_by_name(&self, name: &str) -> Option<&SlideLayoutRef> {
        self.items.iter().find(|l| l.oxml.borrow().name == name)
    }

    /// 按名称查找版式（可变引用，TODO-008）。
    pub fn get_by_name_mut(&mut self, name: &str) -> Option<&mut SlideLayoutRef> {
        self.items.iter_mut().find(|l| l.oxml.borrow().name == name)
    }
}

/// 版式（不可变视图）。
///
/// 当前为"快照"风格 —— 修改 [`SlideLayoutRef`] 不会反映到此结构。
/// 主要用于"展示给用户当前有哪些版式"的场景。
#[derive(Debug, Clone)]
pub struct SlideLayout {
    /// 版式显示名（如 `Blank` / `Title Slide`）。
    pub name: String,
    /// 版式类型（OOXML 中的 `type` 属性，可选）。
    /// 使用 `r#type` 保留字面量 `type`。
    pub r#type: String,
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use std::cell::RefCell;
    use std::rc::Rc;

    /// 构造一个测试用的 SlideLayoutRef。
    fn make_layout(name: &str, rid: &str) -> SlideLayoutRef {
        let oxml = OxmlSldLayout {
            name: name.to_string(),
            type_: "blank".to_string(),
            shapes: Vec::new(),
        };
        SlideLayoutRef {
            idx: 0,
            partname: format!("/ppt/slideLayouts/{}.xml", name),
            rid: rid.to_string(),
            oxml: Rc::new(RefCell::new(oxml)),
        }
    }

    /// TODO-008：`SlideLayouts::remove` 按索引移除。
    #[test]
    fn layouts_remove_by_index() {
        let mut layouts = SlideLayouts::new();
        layouts.push(make_layout("Blank", "rId1"));
        layouts.push(make_layout("Title", "rId2"));
        layouts.push(make_layout("Content", "rId3"));
        assert_eq!(layouts.len(), 3);

        let removed = layouts.remove(1);
        assert!(removed.is_some());
        assert_eq!(removed.unwrap().rid(), "rId2");
        assert_eq!(layouts.len(), 2);
        assert_eq!(layouts.get(1).unwrap().rid(), "rId3");

        // 越界返回 None
        assert!(layouts.remove(99).is_none());
    }

    /// TODO-008：`SlideLayouts::index_of` 按 rid 查找索引。
    #[test]
    fn layouts_index_of_by_rid() {
        let mut layouts = SlideLayouts::new();
        layouts.push(make_layout("Blank", "rId1"));
        layouts.push(make_layout("Title", "rId2"));

        assert_eq!(layouts.index_of("rId1"), Some(0));
        assert_eq!(layouts.index_of("rId2"), Some(1));
        assert_eq!(layouts.index_of("rId999"), None);
    }

    /// TODO-008：`SlideLayouts::get_by_name` 按名查找。
    #[test]
    fn layouts_get_by_name() {
        let mut layouts = SlideLayouts::new();
        layouts.push(make_layout("Blank", "rId1"));
        layouts.push(make_layout("Title Slide", "rId2"));

        let found = layouts.get_by_name("Title Slide");
        assert!(found.is_some());
        assert_eq!(found.unwrap().rid(), "rId2");

        assert!(layouts.get_by_name("Nonexistent").is_none());
    }

    /// TODO-008：`SlideLayoutRef::placeholder_indices` 返回占位符索引。
    #[test]
    fn layout_placeholder_indices() {
        let mut sp1 = OxmlSp::default();
        sp1.is_placeholder = true;
        sp1.ph_idx = Some(0);
        sp1.ph_type = Some("title".into());

        let mut sp2 = OxmlSp::default();
        sp2.is_placeholder = false; // 非占位符

        let mut sp3 = OxmlSp::default();
        sp3.is_placeholder = true;
        sp3.ph_idx = Some(1);
        sp3.ph_type = Some("body".into());

        let oxml = OxmlSldLayout {
            name: "Test".into(),
            type_: "obj".into(),
            shapes: vec![sp1, sp2, sp3],
        };
        let layout = SlideLayoutRef {
            idx: 0,
            partname: "/ppt/slideLayouts/slideLayout1.xml".into(),
            rid: "rId1".into(),
            oxml: Rc::new(RefCell::new(oxml)),
        };

        let indices = layout.placeholder_indices();
        assert_eq!(indices, vec![0, 2]);

        // 按 ph_idx 查找
        assert_eq!(layout.placeholder_index_by_ph_idx(0), Some(0));
        assert_eq!(layout.placeholder_index_by_ph_idx(1), Some(2));
        assert_eq!(layout.placeholder_index_by_ph_idx(99), None);
    }
}
