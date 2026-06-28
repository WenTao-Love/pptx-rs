//! 演示文稿分组（Section，`<p14:sectionLst>` 扩展元素）。
//!
//! 本模块对应 PowerPoint 2010 引入的"章节分组"功能——把若干 slide
//! 归入命名分组，便于在"大纲视图"中折叠 / 跳转。
//!
//! # OOXML 结构
//!
//! `sectionLst` 是 `<p:presentation>` 内 `<p:extLst>` 的扩展元素，
//! 命名空间为 `p14`（PowerPoint 2010 main）：
//!
//! ```xml
//! <p:presentation>
//!   ...
//!   <p:extLst>
//!     <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E64}">
//!       <p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
//!         <p14:section name="章节一">
//!           <p14:sldIdLst>
//!             <p14:sldId id="256"/>
//!             <p14:sldId id="257"/>
//!           </p14:sldIdLst>
//!         </p14:section>
//!       </p14:sectionLst>
//!     </p:ext>
//!   </p:extLst>
//! </p:presentation>
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx 截至 v1.0 仍未提供 section API，本模块是 pptx-rs 的扩展实现，
//! 参考 OOXML 规范 [MS-PPT] 2.6.4 节。
//!
//! # 序列化约束
//!
//! - **必须**放在 `<p:extLst>` 内的 `<p:ext>` 元素中，且 uri 固定为
//!   `{521415D9-36F7-43E2-AB2F-B90AF26B5E64}`；
//! - `<p:extLst>` 在 `<p:defaultTextStyle>` 之后；
//! - section 的 `<p14:sldId>` 只有 `id` 属性（**无** `r:id`）。

use crate::oxml::writer::XmlWriter;

/// sectionLst 扩展元素的固定 URI（PowerPoint 2010 section 特征 GUID）。
///
/// 该 URI 必须出现在 `<p:ext uri="...">` 属性中，PowerPoint 通过它识别
/// 扩展内容为 section 列表。
pub const SECTION_EXT_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E64}";

/// 单个章节（`<p14:section>`）。
///
/// 一个 section 由 `name` 与若干 slide id 组成。
/// slide id 必须已存在于 `<p:sldIdLst>` 中——section 只是按 id 把它们
/// "逻辑分组"，并不持有额外的关系。
#[derive(Clone, Debug, Default)]
pub struct Section {
    /// 章节名（`name` 属性，必填）。
    pub name: String,
    /// 归入本章节的 slide id 列表（`<p14:sldId id="..."/>`）。
    ///
    /// 这里的 `id` 与 `<p:sldIdLst>/<p:sldId>` 的 `id` 同源——
    /// 由 `Presentation` 在写路径中分配（一般从 256 起递增）。
    pub slide_ids: Vec<u32>,
}

impl Section {
    /// 创建一个空章节（仅指定名称）。
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            slide_ids: Vec::new(),
        }
    }

    /// 向本章节追加一个 slide id。
    pub fn push(&mut self, id: u32) {
        self.slide_ids.push(id);
    }

    /// 写 XML——输出**仅** `<p14:section>` 元素（不含外层 `<p14:sectionLst>`）。
    ///
    /// 调用方（通常是 [`SectionList::write_xml`]）负责提供 `<p14:sectionLst>`
    /// 与 `<p:extLst>/<p:ext>` 外壳。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let attrs: Vec<(&str, &str)> = vec![("name", self.name.as_str())];
        w.open_with("p14:section", &attrs);
        if !self.slide_ids.is_empty() {
            w.open("p14:sldIdLst");
            for id in &self.slide_ids {
                let id_s = id.to_string();
                w.empty_with("p14:sldId", &[("id", id_s.as_str())]);
            }
            w.close("p14:sldIdLst");
        }
        w.close("p14:section");
    }
}

/// 章节列表（`<p14:sectionLst>`）。
///
/// 在 [`crate::oxml::presentation::PresentationRoot`] 中以 `sections` 字段
/// 携带，写路径在 `<p:extLst>` 内展开为 `<p14:sectionLst>`。
///
/// # 序列化
///
/// 空列表时**不**输出任何内容（保持 presentation.xml 干净）。
#[derive(Clone, Debug, Default)]
pub struct SectionList {
    /// 所有章节（按文档顺序）。
    pub items: Vec<Section>,
}

impl SectionList {
    /// 新建空列表。
    pub fn new() -> Self {
        Self::default()
    }

    /// 章节数量。
    pub fn len(&self) -> usize {
        self.items.len()
    }

    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// 追加一个章节。
    pub fn push(&mut self, section: Section) {
        self.items.push(section);
    }

    /// 按 slide id 查询所属章节名。
    ///
    /// 同一个 slide id 理论上不应出现在多个 section 内；
    /// 若被多次追加，返回第一个匹配的章节名。
    pub fn section_name_of(&self, slide_id: u32) -> Option<&str> {
        for s in &self.items {
            if s.slide_ids.contains(&slide_id) {
                return Some(s.name.as_str());
            }
        }
        None
    }

    /// 写 XML——输出**完整** `<p:extLst><p:ext ...><p14:sectionLst>...</p14:sectionLst></p:ext></p:extLst>`。
    ///
    /// 当且仅当列表非空时输出；空列表返回空字符串。
    pub fn write_xml(&self) -> String {
        if self.is_empty() {
            return String::new();
        }
        let mut w = XmlWriter::new();
        w.open("p:extLst");
        let ext_attrs: Vec<(&str, &str)> = vec![("uri", SECTION_EXT_URI)];
        w.open_with("p:ext", &ext_attrs);
        // sectionLst 自身无属性；xmlns:p14 由 presentation 根元素统一声明。
        w.open("p14:sectionLst");
        for s in &self.items {
            s.write_xml(&mut w);
        }
        w.close("p14:sectionLst");
        w.close("p:ext");
        w.close("p:extLst");
        w.into_string()
    }
}
