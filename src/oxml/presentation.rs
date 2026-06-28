//! 演示文稿顶级 XML：`<p:presentation>` / `<p:sldIdLst>` / `<p:sldSz>`。
//!
//! 本文件定义"根 part"——`/ppt/presentation.xml` 对应的 Rust 模型。
//!
//! # 元素结构（OOXML 规范严格顺序）
//!
//! ```text
//! <p:presentation>
//!   <p:sldMasterIdLst>...</p:sldMasterIdLst>   必须存在（可为空）
//!   <p:notesSz .../>                             备注尺寸
//!   <p:sldIdLst>                                 所有 slide 的 ID 列表
//!     <p:sldId id="256" r:id="rId5"/>
//!     ...
//!   </p:sldIdLst>
//!   <p:sldSz cx="..." cy="..."/>                 幻灯片尺寸（EMU）
//!   ...其它可选元素...
//!   <p:defaultTextStyle>...</p:defaultTextStyle> 必须存在
//! </p:presentation>
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.presentation.Presentation` ←→ [`PresentationRoot`]；
//! - `pptx.oxml.presentation.SldId` ←→ [`SlideIdEntry`]。
//!
//! # 序列化要点
//!
//! - **`sldMasterIdLst` 必须存在**——PowerPoint 要求至少一个 `sldMasterId`，即使实际不挂母版。
//! - **`defaultTextStyle` 必须存在**——python-pptx 给出 9 段 `lvlXpPr` 的极简版，
//!   本库直接复用（`DEFAULT_TEXT_STYLE`）。

use std::str::FromStr;

use crate::oxml::ns::{NS_DRAWING_MAIN, NS_PRESENTATION_MAIN};
use crate::oxml::writer::XmlWriter;
use crate::units::Emu;

/// 演示文稿的根 `<p:presentation>` 元素。
#[derive(Clone, Debug)]
pub struct PresentationRoot {
    /// 幻灯片宽度（EMU）。`None` 表示不写出 `<p:sldSz>`。
    pub slide_width: Option<Emu>,
    /// 幻灯片高度（EMU）。`None` 表示不写出 `<p:sldSz>`。
    pub slide_height: Option<Emu>,
    /// 所有 slide 的 ID 列表（`<p:sldIdLst>`）。
    pub slide_ids: Vec<SlideIdEntry>,
    /// 母版 ID 列表（`<p:sldMasterIdLst>`）。
    ///
    /// TODO-001：建模 `sldMasterIdLst`，支持多母版。
    /// 空列表时 `to_xml` 使用默认的单个母版（id=2147483648, rIdP1）。
    pub sld_master_ids: Vec<SldMasterIdEntry>,
    /// 默认文本样式（`<p:defaultTextStyle>`），XML 字符串。`None` 使用内置默认。
    pub default_text_style: Option<String>,
    /// 备注页尺寸（宽, 高，EMU）。`None` 使用默认 6858000 × 9144000。
    pub notes_size: Option<(Emu, Emu)>,
    /// 章节分组（`<p14:sectionLst>` 扩展元素，TODO-039）。
    ///
    /// 空列表时不输出任何 section 相关 XML；
    /// 非空时在 `<p:defaultTextStyle>` 之后追加 `<p:extLst><p:ext ...><p14:sectionLst>...`。
    pub sections: crate::oxml::section::SectionList,
}

impl Default for PresentationRoot {
    fn default() -> Self {
        // PowerPoint 严格要求包含 defaultTextStyle，否则打开时弹错。
        // 这里使用 python-pptx 默认的 9 段 lvlXpPr 极简版。
        Self {
            slide_width: None,
            slide_height: None,
            slide_ids: Vec::new(),
            sld_master_ids: Vec::new(),
            default_text_style: Some(DEFAULT_TEXT_STYLE.to_string()),
            notes_size: None,
            sections: crate::oxml::section::SectionList::default(),
        }
    }
}

/// 单一 slide 引用（`p:sldId` 元素）。
#[derive(Clone, Debug)]
pub struct SlideIdEntry {
    /// slide 在 sldIdLst 中的序号（`id` 属性）。
    pub id: u32,
    /// 关系 id（`r:id` 属性），指向 `ppt/slides/slideN.xml`。
    pub rid: String,
}

/// 单一母版引用（`<p:sldMasterId>` 元素）。
///
/// TODO-001：建模 `sldMasterIdLst` 中的每个条目。
#[derive(Clone, Debug)]
pub struct SldMasterIdEntry {
    /// 母版 ID（`id` 属性，通常 >= 2147483648）。
    pub id: u32,
    /// 关系 id（`r:id` 属性），指向 `ppt/slideMasters/slideMasterN.xml`。
    pub rid: String,
}

impl PresentationRoot {
    /// 写 XML。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::with_decl();
        let root_attrs: Vec<(&str, &str)> = vec![
            ("xmlns:a", NS_DRAWING_MAIN),
            ("xmlns:p", NS_PRESENTATION_MAIN),
            ("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS),
            (
                "xmlns:p14",
                "http://schemas.microsoft.com/office/powerpoint/2010/main",
            ),
        ];
        w.open_with("p:presentation", &root_attrs);

        // 1) sldMasterIdLst（必须存在，列出所有母版）
        w.open("p:sldMasterIdLst");
        if self.sld_master_ids.is_empty() {
            // 默认：单个母版（兼容旧行为）
            w.empty_with("p:sldMasterId", &[("id", "2147483648"), ("r:id", "rIdP1")]);
        } else {
            // TODO-001：使用解析出的母版 ID 列表
            for m in &self.sld_master_ids {
                let id_s = m.id.to_string();
                w.empty_with("p:sldMasterId", &[("id", &id_s), ("r:id", m.rid.as_str())]);
            }
        }
        w.close("p:sldMasterIdLst");

        // 2) sldIdLst（所有 slide）
        w.open("p:sldIdLst");
        for s in &self.slide_ids {
            // let 绑定延长生命周期
            let id_s = s.id.to_string();
            w.empty_with("p:sldId", &[("id", &id_s), ("r:id", s.rid.as_str())]);
        }
        w.close("p:sldIdLst");

        // 3) sldSz
        if let (Some(w_), Some(h)) = (self.slide_width, self.slide_height) {
            let cx_s = w_.value().to_string();
            let cy_s = h.value().to_string();
            w.empty_with("p:sldSz", &[("cx", &cx_s), ("cy", &cy_s)]);
        }
        // 4) notesSz（注释页尺寸：默认 6858000 x 9144000 EMU，约 7.5" x 10"）
        let (nw, nh) = self.notes_size.unwrap_or((Emu(6_858_000), Emu(9_144_000)));
        let ncx = nw.value().to_string();
        let ncy = nh.value().to_string();
        w.empty_with("p:notesSz", &[("cx", &ncx), ("cy", &ncy)]);
        // 5) defaultTextStyle
        if let Some(s) = &self.default_text_style {
            w.raw(s);
        }
        // 6) sectionLst（TODO-039，PowerPoint 2010 扩展）
        // 必须放在 `<p:extLst>` 内的 `<p:ext>` 中，且 uri 固定。
        // 空列表时不输出任何内容。
        let section_xml = self.sections.write_xml();
        if !section_xml.is_empty() {
            w.raw(&section_xml);
        }

        w.close("p:presentation");
        w.into_string()
    }
}

impl FromStr for PresentationRoot {
    type Err = crate::Error;
    fn from_str(_s: &str) -> Result<Self, Self::Err> {
        // 完整解析器仍在路线图，读取走专门函数
        Ok(PresentationRoot::default())
    }
}

/// 默认 `<p:defaultTextStyle>` 内容（与 python-pptx 默认输出对齐）。
/// 9 段 lvlXpPr，每段 18x 字号 / 主题色 / latin/+mn-lt 等。
const DEFAULT_TEXT_STYLE: &str = r##"<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:defaultTextStyle>"##;
