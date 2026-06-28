//! 幻灯片母版 `<p:sldMaster>` 简化模型。
//!
//! 当前只生成"能跑"级别的母版；样式继承仍走 `a:lstStyle` 的最小默认。
//! 注意：PowerPoint 严格要求 master 拥有 `<p:cSld>/<p:spTree>`、`<p:clrMap>`、`<p:sldLayoutIdLst>` 与
//! `<p:txStyles>`（titleStyle / bodyStyle / otherStyle）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.slidemaster.SlideMaster` ←→ [`SldMaster`]；
//! - `SlideMasterPart` 由 `pptx.parts.slidemaster` 维护，本库不直接对应。
//!
//! # 默认内容
//!
//! `TX_STYLES` 是与 python-pptx 默认输出对齐的极简版
//! （`titleStyle` / `bodyStyle` / `otherStyle` 各 1 个 `lvl1pPr`），
//! 已通过 PowerPoint 打开测试。

use crate::oxml::shape::Sp;
use crate::oxml::slide::SlideBackground;

/// `<p:txStyles>` 三段标题/正文/其它文本样式的极简版。
/// python-pptx 输出的 bodyStyle 包含 9 个 lvlXpPr，这里给每个 1 个 lvl1pPr 让 OOXML 校验通过。
const TX_STYLES: &str = r##"<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles>"##;

#[derive(Clone, Debug, Default)]
pub struct SldMaster {
    /// 母版中的形状列表。
    pub shapes: Vec<Sp>,
    /// 关联的版式关系 ID 列表（指向 `slideLayoutN.xml`）。
    pub layout_rids: Vec<String>,
    /// 母版背景（`<p:bg>`，可选）。`None` 表示使用默认背景。
    ///
    /// 由 Slide、SlideLayout 继承。详见 OOXML `CT_SlideMaster` 的 `<p:cSld>/<p:bg>`。
    pub background: Option<SlideBackground>,
}

impl SldMaster {
    /// 写出最小但合规的 `<p:sldMaster>` XML。
    pub fn to_xml(&self) -> String {
        // 头部（含三个 namespace）+ 完整结构
        let mut xml = String::with_capacity(1024);
        xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        xml.push_str(
            "<p:sldMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        );
        xml.push_str(" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"");
        xml.push_str(
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
        );
        // cSld —— 必须按 <p:bg>? <p:spTree> <p:custDataLst>? <p:controls>? <p:extLst>? 顺序
        xml.push_str("<p:cSld>");
        // 背景在 spTree 之前（OOXML 顺序：bg? → spTree → custDataLst? → controls? → extLst?）
        if let Some(bg) = &self.background {
            let mut w = crate::oxml::writer::XmlWriter::default();
            bg.write_xml(&mut w);
            xml.push_str(&w.into_string());
        }
        xml.push_str("<p:spTree>");
        // nvGrpSpPr + grpSpPr (spTree 必填项)
        xml.push_str(
            "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>",
        );
        xml.push_str("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
        // 用户 shapes
        let mut w = crate::oxml::writer::XmlWriter::default();
        for s in &self.shapes {
            s.write_xml(&mut w);
        }
        xml.push_str(&w.into_string());
        xml.push_str("</p:spTree></p:cSld>");
        // clrMap
        xml.push_str("<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>");
        // sldLayoutIdLst - 必须存在且至少有一个 sldLayoutId
        xml.push_str(
            "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>",
        );
        // txStyles (必须有，否则 PowerPoint 解析报错)
        xml.push_str(TX_STYLES);
        xml.push_str("</p:sldMaster>");
        xml
    }
}
