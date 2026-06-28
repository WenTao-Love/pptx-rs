//! OLE 对象嵌入（`<p:oleObj>`，TODO-043）。
//!
//! 本模块定义 OLE 对象的强类型模型。OLE 对象在 OOXML 中通过
//! `<p:graphicFrame>` + `<a:graphicData uri=".../ole">` + `<p:oleObj>` 引用一个
//! 独立的 `/ppt/embeddings/oleObjectN.bin` part（CFB 复合文档二进制）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.graphfrm.GraphicFrame` ←→ [`crate::oxml::shape::GraphicFrame`]（承载位置/尺寸）；
//! - `pptx.parts.oleobject.OleObjectPart` ←→ `/ppt/embeddings/oleObjectN.bin` part；
//! - python-pptx 0.6.19+ 的 `shapes.add_ole_object()` ←→ [`crate::slide::ShapesMut::add_ole_object`]。
//!
//! # 写出语义
//!
//! - `OleObject` 序列化时只写出 `<p:oleObj r:id="..."/>` 引用元素 + 可选 `<p:pic>` 图标；
//! - 真正的 OLE 二进制数据由 [`crate::presentation::Presentation::save`] 在
//!   `to_opc_package` 中遍历每张 slide 的 `ole_entries` 写出独立的 `oleObjectN.bin` part；
//! - slide 的 `_rels/slideN.xml.rels` 中会添加 `oleObject` 关系指向该 part。
//!
//! # 限制
//!
//! - 当前仅支持**嵌入**（`<p:embed/>`），不支持链接（`<p:link/>`）；
//! - 图标图片可选（`image_rid` 为空时不写 `<p:pic>`，PowerPoint 会用默认图标）；
//! - 不解析已有 OLE 对象的 `<p:oleObj>` 内容（读路径仅保留 graphicFrame 外壳）。

use crate::oxml::writer::XmlWriter;
use crate::units::Emu;

/// OLE 对象 URI（`<a:graphicData uri="...">`）。
///
/// 固定为 `http://schemas.openxmlformats.org/presentationml/2006/ole`，
/// 与 Chart/Table 的 URI 同级。
pub const OLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

/// OLE 对象模型（`<p:oleObj>`）。
///
/// 一个完整的 OLE 嵌入对象，引用独立的 `/ppt/embeddings/oleObjectN.bin` part。
///
/// # 字段说明
///
/// - `rid`：指向 `oleObjectN.bin` 的关系 id（由 `ShapesMut::add_ole_object` 分配，
///   `to_opc_package` 在 slideN.xml.rels 中注册）。
/// - `image_rid`：指向图标图片的关系 id（与 `image_rid` 共用 slide rels 命名空间）。
///   为空字符串时表示无图标图片，PowerPoint 会用默认图标显示。
/// - `prog_id`：OLE 程序标识符（如 `"Excel.Sheet.12"` / `"Word.Document.12"` /
///   `"Package"`）。PowerPoint 通过 progId 决定双击时调用哪个 OLE 服务器。
/// - `name`：显示名（在 PowerPoint 中作为对象名）。
/// - `show_as_icon`：是否以图标形式显示（`true` 时写出 `showAsIcon="1"`）。
/// - `image_width` / `image_height`：图标显示尺寸（EMU）。
/// - `pic_id` / `pic_name`：图标图片 Pic 形状的 id 与 name。
#[derive(Clone, Debug)]
pub struct OleObject {
    /// 指向 `oleObjectN.bin` 的关系 id。
    pub rid: String,
    /// 指向图标图片的关系 id（空字符串表示无图标）。
    pub image_rid: String,
    /// OLE 程序标识符（如 `"Excel.Sheet.12"`）。
    pub prog_id: String,
    /// 显示名（如 `"Worksheet"` / `"Document"`）。
    pub name: String,
    /// 是否以图标形式显示。
    pub show_as_icon: bool,
    /// 图标宽度（EMU）。
    pub image_width: Emu,
    /// 图标高度（EMU）。
    pub image_height: Emu,
    /// 图标 Pic 形状的 id。
    pub pic_id: u32,
    /// 图标 Pic 形状的 name。
    pub pic_name: String,
}

impl Default for OleObject {
    fn default() -> Self {
        OleObject {
            rid: String::new(),
            image_rid: String::new(),
            prog_id: "Package".to_string(),
            name: "OLE Object".to_string(),
            show_as_icon: true,
            image_width: Emu(914400), // 1 英寸
            image_height: Emu(914400),
            pic_id: 0,
            pic_name: "OLE Icon".to_string(),
        }
    }
}

impl OleObject {
    /// 构造一个指定 progId 与显示名的 OLE 对象（rid/image_rid 留空，由 presentation 层填充）。
    ///
    /// # 参数
    /// - `prog_id`：OLE 程序标识符（如 `"Excel.Sheet.12"` / `"Package"`）。
    /// - `name`：显示名（如 `"Worksheet"`）。
    pub fn new(prog_id: impl Into<String>, name: impl Into<String>) -> Self {
        OleObject {
            prog_id: prog_id.into(),
            name: name.into(),
            ..Default::default()
        }
    }

    /// 写出 `<p:oleObj>` 完整 XML（含 `<p:embed/>` 与可选 `<p:pic>`）。
    ///
    /// # 元素结构
    ///
    /// ```text
    /// <p:oleObj spid="..." name="..." r:id="..." imgW="..." imgH="..." progId="..." showAsIcon="1">
    ///   <p:embed/>
    ///   <p:pic>           ← 仅当 image_rid 非空时写出
    ///     <p:nvPicPr>...
    ///     <p:blipFill>...
    ///     <p:spPr>...
    ///   </p:pic>
    /// </p:oleObj>
    /// ```
    ///
    /// # 注意
    ///
    /// 本方法**只**写出 `<p:oleObj>` 元素本身；外层的 `<a:graphicData uri="...">`
    /// 由 [`crate::oxml::shape::GraphicFrame::write_xml`] 负责包裹。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let spid = format!("_x0000_s{}", self.pic_id);
        let img_w = self.image_width.0.to_string();
        let img_h = self.image_height.0.to_string();
        let show_as_icon = if self.show_as_icon { "1" } else { "0" };
        let rid = if self.rid.is_empty() {
            "rId1"
        } else {
            self.rid.as_str()
        };

        // <p:oleObj ...>
        w.open_with(
            "p:oleObj",
            &[
                ("spid", spid.as_str()),
                ("name", self.name.as_str()),
                ("r:id", rid),
                ("imgW", img_w.as_str()),
                ("imgH", img_h.as_str()),
                ("progId", self.prog_id.as_str()),
                ("showAsIcon", show_as_icon),
            ],
        );
        // <p:embed/>  ← 嵌入对象（与 <p:link/> 对立）
        w.empty("p:embed");
        // <p:pic>  ← 图标图片（可选）
        if !self.image_rid.is_empty() {
            self.write_pic_xml(w);
        }
        w.close("p:oleObj");
    }

    /// 写出图标图片的 `<p:pic>` 元素。
    ///
    /// 包含三个子元素：`<p:nvPicPr>` / `<p:blipFill>` / `<p:spPr>`，
    /// 与普通 Pic 形状结构一致，但 spPr 用固定的矩形几何。
    fn write_pic_xml(&self, w: &mut XmlWriter) {
        let pic_id_s = self.pic_id.to_string();
        let img_rid = self.image_rid.as_str();
        let cx_s = self.image_width.0.to_string();
        let cy_s = self.image_height.0.to_string();

        w.open("p:pic");
        // <p:nvPicPr>
        w.open("p:nvPicPr");
        w.empty_with(
            "p:cNvPr",
            &[("id", pic_id_s.as_str()), ("name", self.pic_name.as_str())],
        );
        w.empty("p:cNvPicPr");
        w.empty("p:nvPr");
        w.close("p:nvPicPr");
        // <p:blipFill>
        w.open("p:blipFill");
        w.empty_with(
            "a:blip",
            &[
                (
                    "xmlns:r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ),
                ("r:embed", img_rid),
            ],
        );
        w.open("a:stretch");
        w.empty("a:fillRect");
        w.close("a:stretch");
        w.close("p:blipFill");
        // <p:spPr>
        w.open("p:spPr");
        w.open("a:xfrm");
        w.empty_with("a:off", &[("x", "0"), ("y", "0")]);
        w.empty_with("a:ext", &[("cx", cx_s.as_str()), ("cy", cy_s.as_str())]);
        w.close("a:xfrm");
        // prstGeom：rect 几何（与 Geometry::Preset(Rectangle, []) 一致）
        w.open_with("a:prstGeom", &[("prst", "rect")]);
        w.empty("a:avLst");
        w.close("a:prstGeom");
        w.close("p:spPr");
        w.close("p:pic");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `new` 正确构造 OleObject：progId/name 自定义，其余字段为默认值。
    #[test]
    fn new_ole_object_basics() {
        let ole = OleObject::new("Excel.Sheet.12", "Worksheet");
        assert_eq!(ole.prog_id, "Excel.Sheet.12");
        assert_eq!(ole.name, "Worksheet");
        assert_eq!(ole.rid, "");
        assert_eq!(ole.image_rid, "");
        assert!(ole.show_as_icon);
        assert_eq!(ole.image_width, Emu(914400));
    }

    /// `write_xml` 在无图标图片时正确省略 `<p:pic>`。
    #[test]
    fn write_xml_without_icon() {
        let mut ole = OleObject::new("Package", "OLE Object");
        ole.rid = "rIdOle1".to_string();
        let mut w = XmlWriter::new();
        ole.write_xml(&mut w);
        let xml = w.into_string();
        assert!(xml.contains("<p:oleObj"));
        assert!(xml.contains("r:id=\"rIdOle1\""));
        assert!(xml.contains("progId=\"Package\""));
        assert!(xml.contains("<p:embed/>"));
        // 无图标时不应出现 <p:pic>
        assert!(!xml.contains("<p:pic>"));
    }

    /// `write_xml` 在有图标图片时正确写出 `<p:pic>`。
    #[test]
    fn write_xml_with_icon() {
        let mut ole = OleObject::new("Word.Document.12", "Document");
        ole.rid = "rIdOle1".to_string();
        ole.image_rid = "rIdImg1".to_string();
        ole.pic_id = 1026;
        let mut w = XmlWriter::new();
        ole.write_xml(&mut w);
        let xml = w.into_string();
        assert!(xml.contains("<p:pic>"));
        assert!(xml.contains("r:embed=\"rIdImg1\""));
        assert!(xml.contains("spid=\"_x0000_s1026\""));
    }

    /// `show_as_icon=false` 时写出 `showAsIcon="0"`。
    #[test]
    fn write_xml_show_as_icon_false() {
        let mut ole = OleObject::new("Package", "Object");
        ole.rid = "rId1".to_string();
        ole.show_as_icon = false;
        let mut w = XmlWriter::new();
        ole.write_xml(&mut w);
        let xml = w.into_string();
        assert!(xml.contains("showAsIcon=\"0\""));
    }
}
