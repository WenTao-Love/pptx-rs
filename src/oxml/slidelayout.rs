//! 幻灯片版式 `<p:sldLayout>`。
//!
//! 母版与版式是 PowerPoint 模板机制的两层抽象：母版定义全局样式，版式
//! 在母版基础上覆盖特定占位符与排版。本库当前只提供"blank"（空白）版式的
//! 极简实现。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.slidelayout.SlideLayout` ←→ [`SldLayout`]；
//! - `pptx.slide.SlideLayout` 高阶对象在 [`crate::slide_layouts`]。
//!
//! # 序列化要点
//!
//! - **`type` 属性必填**（`title` / `blank` / `section` / ...），缺失会报"Invalid OOXML"；
//! - `cSld/@name` 可选，缺失会回退到 `type` 值。

use crate::oxml::shape::Sp;

#[derive(Clone, Debug, Default)]
pub struct SldLayout {
    /// 版式名称（对应 `<p:cSld name="...">`）。
    pub name: String,
    /// 版式类型（`title` / `blank` / ...），对应 `<p:sldLayout type="...">`。
    pub type_: String,
    /// 版式中的占位符形状列表。
    pub shapes: Vec<Sp>,
}

impl SldLayout {
    /// 写出最小但合规的 `<p:sldLayout>` XML。
    pub fn to_xml(&self) -> String {
        // 取出 type（如果为空则用 "blank"）
        let layout_type = if self.type_.is_empty() {
            "blank"
        } else {
            &self.type_
        };
        let name = if self.name.is_empty() {
            layout_type.to_string()
        } else {
            self.name.clone()
        };

        let mut xml = String::with_capacity(1024);
        xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        xml.push_str(
            "<p:sldLayout xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        );
        xml.push_str(" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"");
        xml.push_str(
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        );
        // 关键：type 属性（title / blank / ...）必须存在
        xml.push_str(" type=\"");
        xml.push_str(layout_type);
        xml.push_str("\" preserve=\"1\">");
        // cSld
        xml.push_str("<p:cSld name=\"");
        xml.push_str(&name);
        xml.push_str("\"><p:spTree>");
        // nvGrpSpPr + grpSpPr
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
        // clrMapOvr + transition
        xml.push_str("<p:clrMapOvr bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\"/>");
        xml.push_str("<p:transition/>");
        xml.push_str("</p:sldLayout>");
        xml
    }
}
