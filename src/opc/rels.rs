//! 关系（`.rels`）模型。
//!
//! 一个 `.rels` 文件是一组 `Relationship` 元素，每个元素表达一个 part
//! 指向另一个 part（以"相对 URI"表达）的链接。OOXML 规范使用 `.rels`
//! 文件把分散的 XML part 组织成有向图。
//!
//! # 文件位置约定
//!
//! 关系文件存放在被描述 part 的"兄弟目录 `_rels/`"中，文件名 `<part>.rels`：
//!
//! ```text
//! ppt/slides/slide1.xml          ←  被描述的 part
//! ppt/slides/_rels/slide1.xml.rels   ←  关系文件
//! ```
//!
//! 根包的关系文件固定为 `/_rels/.rels`。
//!
//! # 解析与序列化
//!
//! - [`Relationships::to_xml`] 序列化为 XML 字符串；
//! - [`Relationships::from_xml`] 解析 XML 字符串为模型。
//!
//! 两者互为逆操作；解析时遇到未知 `Type` URI 会返回 [`Error::Opc`](crate::Error::Opc)
//! 以提示 schema 变更。

use std::borrow::Cow;
use std::collections::BTreeMap;

use super::part::PartName;
use crate::oxml::ns::NS_RELATIONSHIPS;

/// 关系类型（OOXML 标准）。
///
/// 列举 OOXML 规范中"办公文档"命名空间下的常用关系类型；其它类型用
/// [`RelType::Other`] 兜底（传 `&'static str` URI）。
#[derive(Clone, Debug, Eq, PartialEq, Hash)]
pub enum RelType {
    /// `.../relationships/officeDocument`：根指向演示文稿。
    OfficeDocument,
    /// `.../relationships/slide`：幻灯片。
    Slide,
    /// `.../relationships/slideLayout`：幻灯片布局。
    SlideLayout,
    /// `.../relationships/slideMaster`：幻灯片母版。
    SlideMaster,
    /// `.../relationships/theme`：主题。
    Theme,
    /// `.../relationships/image`：图片。
    Image,
    /// `.../relationships/notesSlide`：备注页。
    NotesSlide,
    /// `.../relationships/notesMaster`：备注母版（TODO-045）。
    ///
    /// 对应 `/ppt/notesMasters/notesMasterN.xml`，被 `presentation.xml`
    /// 的 `<p:notesMasterIdLst>` 引用。
    NotesMaster,
    /// `.../relationships/comments`：幻灯片评论（`/ppt/comments/commentN.xml`）。
    Comments,
    /// `.../relationships/commentAuthors`：评论作者列表（`/ppt/commentAuthors.xml`）。
    CommentAuthors,
    /// `.../relationships/hyperlink`：超链接（外链）。
    Hyperlink,
    /// `.../relationships/slide`：幻灯片（备查别名，与 `Slide` 等价）。
    /// 当前未用——保留便于未来其它 part 引用。
    SlideRel,
    /// `.../relationships/chart`：图表（路线图中）。
    Chart,
    /// `.../relationships/customXml`：自定义 XML。
    CustomXml,
    /// `.../relationships/presProps`：演示文稿属性。
    PresProps,
    /// `.../relationships/viewProps`：视图属性。
    ViewProps,
    /// `.../relationships/tableStyles`：表格样式。
    TableStyles,
    /// `.../relationships/oleObject`：OLE 对象嵌入（TODO-043）。
    ///
    /// 对应 `/ppt/embeddings/oleObjectN.bin`，被 slide 的 `<p:oleObj r:id="..."/>` 引用。
    /// 与 `Image` 关系配合使用：`oleObj` 内的 `<p:pic>` 用 `Image` 关系引用图标图片。
    OleObject,
    /// `.../relationships/video`：视频文件（TODO-033）。
    ///
    /// 对应 `/ppt/media/mediaN.mp4`（或其它视频扩展名），被 slide 的
    /// `<p:pic><p:nvPicPr><p:nvPr><a:videoFile r:link="..."/></p:nvPr>` 引用。
    /// 注意：使用 `r:link` 而非 `r:embed`（视频按外部链接方式存储）。
    Video,
    /// `.../relationships/audio`：音频文件（TODO-033）。
    ///
    /// 对应 `/ppt/media/mediaN.mp3`（或其它音频扩展名），被 slide 的
    /// `<p:pic><p:nvPicPr><p:nvPr><a:audioFile r:link="..."/></p:nvPr>` 引用。
    Audio,
    /// `.../relationships/media`：媒体引用（TODO-033，用于 timeline 同步）。
    ///
    /// 在 `<p:timing>` 中通过 `<p:video>` / `<p:audio>` 引用，与 Video/Audio
    /// 关系配合使用。当前库未实现 timing，仅保留 RelType 以支持 round-trip。
    Media,
    /// `.../relationships/diagramData`：SmartArt 数据 part（TODO-037）。
    ///
    /// 对应 `/ppt/diagrams/dataN.xml`，被 slide 的
    /// `<p:graphicFrame><a:graphic><a:graphicData uri="...diagram"><dgm:relIds r:dm="..."/>` 引用。
    /// 与 DiagramLayout / DiagramQuickStyle / DiagramColors 一起构成 SmartArt 完整 round-trip。
    DiagramData,
    /// `.../relationships/diagramLayout`：SmartArt 布局 part（TODO-037）。
    ///
    /// 对应 `/ppt/diagrams/layoutN.xml`，被 `<dgm:relIds r:lo="..."/>` 引用。
    DiagramLayout,
    /// `.../relationships/diagramQuickStyle`：SmartArt 快速样式 part（TODO-037）。
    ///
    /// 对应 `/ppt/diagrams/quickStylesN.xml`，被 `<dgm:relIds r:qs="..."/>` 引用。
    DiagramQuickStyle,
    /// `.../relationships/diagramColors`：SmartArt 颜色 part（TODO-037）。
    ///
    /// 对应 `/ppt/diagrams/colorsN.xml`，被 `<dgm:relIds r:cs="..."/>` 引用。
    DiagramColors,
    /// `.../relationships/package`：嵌入式包关系（TODO-004 Excel 嵌入）。
    ///
    /// OOXML 中通用"嵌入式包"关系类型。在 pptx-rs 当前主要用于 chart 的
    /// `<c:externalData r:id="..."/>` 引用嵌入式 Excel 工作簿
    /// （`/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`）。
    /// 也用于其它嵌入式包（如嵌入的 docx/pptx），但当前库仅支持 xlsx 场景。
    Package,
    /// 任意其它（用字符串表达）。
    Other(String),
}

impl RelType {
    /// 转 URI 字符串。
    pub fn uri(&self) -> Cow<'_, str> {
        match self {
            RelType::OfficeDocument =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
            RelType::Slide =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"),
            RelType::SlideLayout =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"),
            RelType::SlideMaster =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"),
            RelType::Theme =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
            RelType::Image =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
            RelType::NotesSlide =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"),
            RelType::NotesMaster =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"),
            RelType::Comments =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"),
            RelType::CommentAuthors =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"),
            RelType::Hyperlink =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
            RelType::SlideRel =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"),
            RelType::Chart =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
            RelType::CustomXml =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"),
            RelType::PresProps =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"),
            RelType::ViewProps =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"),
            RelType::TableStyles =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"),
            RelType::OleObject =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"),
            RelType::Video =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"),
            RelType::Audio =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"),
            RelType::Media =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/media"),
            RelType::DiagramData =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"),
            RelType::DiagramLayout =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"),
            RelType::DiagramQuickStyle =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"),
            RelType::DiagramColors =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"),
            RelType::Package =>
                Cow::Borrowed("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"),
            RelType::Other(s) => Cow::Borrowed(s.as_str()),
        }
    }
}

impl std::fmt::Display for RelType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.uri())
    }
}

/// 一条关系。
#[derive(Debug, Clone)]
pub struct Relationship {
    /// 关系 ID（如 `rId1`）。
    pub id: String,
    /// 关系类型。
    pub reltype: RelType,
    /// 目标 part name（绝对），或外链（hyperlink）。
    pub target: Target,
    /// 是否为外部资源（hyperlink）。
    pub is_external: bool,
}

impl Relationship {
    /// 创建内部关系。
    pub fn internal(id: impl Into<String>, reltype: RelType, target: PartName) -> Self {
        Relationship {
            id: id.into(),
            reltype,
            target: Target::Internal(target),
            is_external: false,
        }
    }
    /// 创建内部关系（target 用裸字符串，常用于相对路径如 `../theme/theme1.xml`）。
    pub fn internal_str(
        id: impl Into<String>,
        reltype: RelType,
        target: impl Into<String>,
    ) -> Self {
        Relationship {
            id: id.into(),
            reltype,
            target: Target::InternalStr(target.into()),
            is_external: false,
        }
    }
    /// 创建外部关系（hyperlink）。
    pub fn external(id: impl Into<String>, reltype: RelType, url: impl Into<String>) -> Self {
        Relationship {
            id: id.into(),
            reltype,
            target: Target::External(url.into()),
            is_external: true,
        }
    }
}

/// 关系的目标：内部 part / 外部 URL。
#[derive(Debug, Clone)]
pub enum Target {
    /// 内部 part（PartName 形式）。
    Internal(PartName),
    /// 内部 part（裸字符串 target，保留原始相对路径；如 `../theme/theme1.xml`）。
    InternalStr(String),
    /// 外部 URL。
    External(String),
}

impl Target {
    /// 写出 XML 时使用的字面量。
    pub fn as_str(&self) -> &str {
        match self {
            Target::Internal(p) => p.as_str(),
            Target::InternalStr(s) => s.as_str(),
            Target::External(s) => s.as_str(),
        }
    }
}

/// 一个 `.rels` 文件对应的全部关系集合。
///
/// 内部以 `Vec<Relationship>` 保留插入顺序，同时维护 `by_id` 索引
/// 提供 O(log n) 的按 id 查询。
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    items: Vec<Relationship>,
    by_id: BTreeMap<String, usize>,
}

impl Relationships {
    /// 空集合。
    pub fn new() -> Self {
        Relationships::default()
    }

    /// 添加一条关系，自动检查 ID 唯一。
    ///
    /// # 错误
    /// - [`crate::Error::Opc`]：ID 重复。
    pub fn add(&mut self, r: Relationship) -> crate::Result<&Relationship> {
        if self.by_id.contains_key(&r.id) {
            return Err(crate::Error::opc(format!(
                "duplicate relationship id: {}",
                r.id
            )));
        }
        let idx = self.items.len();
        self.by_id.insert(r.id.clone(), idx);
        self.items.push(r);
        Ok(&self.items[idx])
    }

    /// 取所有关系（按插入顺序）。
    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.items.iter()
    }

    /// 数量。
    pub fn len(&self) -> usize {
        self.items.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// 按 ID 找。
    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id).map(|&i| &self.items[i])
    }

    /// 按 reltype 过滤。
    pub fn of_type<'a>(&'a self, t: RelType) -> impl Iterator<Item = &'a Relationship> + 'a {
        self.items.iter().filter(move |r| r.reltype == t)
    }

    /// 转成 XML 字符串。
    pub fn to_xml(&self) -> String {
        let mut s = String::with_capacity(256);
        s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        s.push_str(&format!("<Relationships xmlns=\"{}\">", NS_RELATIONSHIPS));
        for r in &self.items {
            s.push_str(&format!(
                "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"{}/>",
                xml_escape(&r.id),
                xml_escape(&r.reltype.uri()),
                xml_escape(r.target.as_str()),
                if r.is_external {
                    " TargetMode=\"External\""
                } else {
                    ""
                },
            ));
        }
        s.push_str("</Relationships>");
        s
    }

    /// 解析一段 XML，构造关系集合。
    ///
    /// # 错误
    /// - [`crate::Error::Opc`]：元素属性缺失或 Type URI 不在已知清单中；
    /// - [`crate::Error::Xml`]：XML 解析失败。
    pub fn from_xml(xml: &str) -> crate::Result<Self> {
        use quick_xml::events::Event;
        use quick_xml::reader::Reader;

        let mut r = Relationships::new();
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    let name = e.name();
                    if name.as_ref() != b"Relationship" {
                        continue;
                    }
                    let mut id = None;
                    let mut rtype = None;
                    let mut target = None;
                    let mut external = false;
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let v = attr
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .to_string();
                        match key {
                            b"Id" => id = Some(v),
                            b"Type" => rtype = Some(v),
                            b"Target" => target = Some(v),
                            b"TargetMode" if v == "External" => external = true,
                            _ => {}
                        }
                    }
                    let (Some(id), Some(rtype), Some(target)) = (id, rtype, target) else {
                        return Err(crate::Error::opc("malformed <Relationship> element"));
                    };
                    let reltype = match rtype.as_str() {
                        x if x == RelType::OfficeDocument.uri() => RelType::OfficeDocument,
                        x if x == RelType::Slide.uri() => RelType::Slide,
                        x if x == RelType::SlideLayout.uri() => RelType::SlideLayout,
                        x if x == RelType::SlideMaster.uri() => RelType::SlideMaster,
                        x if x == RelType::Theme.uri() => RelType::Theme,
                        x if x == RelType::Image.uri() => RelType::Image,
                        x if x == RelType::NotesSlide.uri() => RelType::NotesSlide,
                        x if x == RelType::NotesMaster.uri() => RelType::NotesMaster,
                        x if x == RelType::Comments.uri() => RelType::Comments,
                        x if x == RelType::CommentAuthors.uri() => RelType::CommentAuthors,
                        x if x == RelType::Hyperlink.uri() => RelType::Hyperlink,
                        x if x == RelType::Chart.uri() => RelType::Chart,
                        x if x == RelType::CustomXml.uri() => RelType::CustomXml,
                        x if x == RelType::PresProps.uri() => RelType::PresProps,
                        x if x == RelType::ViewProps.uri() => RelType::ViewProps,
                        x if x == RelType::TableStyles.uri() => RelType::TableStyles,
                        x if x == RelType::OleObject.uri() => RelType::OleObject,
                        x if x == RelType::Video.uri() => RelType::Video,
                        x if x == RelType::Audio.uri() => RelType::Audio,
                        x if x == RelType::Media.uri() => RelType::Media,
                        x if x == RelType::DiagramData.uri() => RelType::DiagramData,
                        x if x == RelType::DiagramLayout.uri() => RelType::DiagramLayout,
                        x if x == RelType::DiagramQuickStyle.uri() => RelType::DiagramQuickStyle,
                        x if x == RelType::DiagramColors.uri() => RelType::DiagramColors,
                        x if x == RelType::Package.uri() => RelType::Package,
                        // 未知关系类型用 Other 兜底，避免加载真实世界 .pptx 时因
                        // handoutMaster / tags / comment 等未知类型而报错。
                        other => RelType::Other(other.to_string()),
                    };
                    // 保留原始 target 字符串（相对路径如 "../notesSlides/notesSlide1.xml"），
                    // 而非转为 PartName（会错误地加前导 '/' 导致 resolve_relative_partname
                    // 误判为绝对路径）。调用方在需要绝对路径时自行调用 resolve_relative_partname。
                    let rel = if external {
                        Relationship::external(id, reltype, target)
                    } else {
                        Relationship::internal_str(id, reltype, target)
                    };
                    r.add(rel)?;
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::Error::Xml(format!("relationships parse: {e}"))),
                _ => {}
            }
            buf.clear();
        }
        Ok(r)
    }
}

/// XML 字符串属性转义（最小集：覆盖 `& < > " '`）。
///
/// 仅供 OPC 内部使用，故 `pub(crate)`；OOXML 序列化请走
/// [`crate::oxml::writer::XmlWriter`]。
pub(crate) fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::opc::part::new_part_name;

    /// 关系集合的 XML round-trip 测试。
    #[test]
    fn round_trip() {
        let mut r = Relationships::new();
        r.add(Relationship::internal(
            "rId1",
            RelType::Slide,
            new_part_name("/ppt/slides/slide1.xml"),
        ))
        .unwrap();
        r.add(Relationship::internal(
            "rId2",
            RelType::SlideLayout,
            new_part_name("/ppt/slideLayouts/slideLayout1.xml"),
        ))
        .unwrap();
        let xml = r.to_xml();
        let r2 = Relationships::from_xml(&xml).unwrap();
        assert_eq!(r2.len(), 2);
        assert!(r2.get("rId1").is_some());
    }

    // ==================== SmartArt 关系类型测试（TODO-037） ====================

    /// 验证 4 个 diagram 关系类型的 URI 正确。
    #[test]
    fn diagram_reltype_uri_correct() {
        assert_eq!(
            RelType::DiagramData.uri(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
        );
        assert_eq!(
            RelType::DiagramLayout.uri(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
        );
        assert_eq!(
            RelType::DiagramQuickStyle.uri(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
        );
        assert_eq!(
            RelType::DiagramColors.uri(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
        );
    }

    /// 验证 `Relationships::from_xml` 能识别 4 个 diagram 关系类型（而非归为 Other）。
    #[test]
    fn from_xml_recognizes_diagram_reltypes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdDgmData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
<Relationship Id="rIdDgmLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
<Relationship Id="rIdDgmQs1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyles1.xml"/>
<Relationship Id="rIdDgmColors1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
</Relationships>"#;
        let r = Relationships::from_xml(xml).expect("parse ok");
        assert_eq!(r.len(), 4);

        let data = r.get("rIdDgmData1").expect("data rel");
        assert_eq!(data.reltype, RelType::DiagramData);

        let layout = r.get("rIdDgmLayout1").expect("layout rel");
        assert_eq!(layout.reltype, RelType::DiagramLayout);

        let qs = r.get("rIdDgmQs1").expect("qs rel");
        assert_eq!(qs.reltype, RelType::DiagramQuickStyle);

        let colors = r.get("rIdDgmColors1").expect("colors rel");
        assert_eq!(colors.reltype, RelType::DiagramColors);
    }

    /// 验证 4 个 diagram 关系能正确 round-trip（to_xml → from_xml）。
    #[test]
    fn diagram_reltype_round_trip() {
        let mut r = Relationships::new();
        r.add(Relationship::internal_str(
            "rIdDgmData1",
            RelType::DiagramData,
            "../diagrams/data1.xml",
        ))
        .unwrap();
        r.add(Relationship::internal_str(
            "rIdDgmLayout1",
            RelType::DiagramLayout,
            "../diagrams/layout1.xml",
        ))
        .unwrap();
        r.add(Relationship::internal_str(
            "rIdDgmQs1",
            RelType::DiagramQuickStyle,
            "../diagrams/quickStyles1.xml",
        ))
        .unwrap();
        r.add(Relationship::internal_str(
            "rIdDgmColors1",
            RelType::DiagramColors,
            "../diagrams/colors1.xml",
        ))
        .unwrap();

        let xml = r.to_xml();
        let r2 = Relationships::from_xml(&xml).expect("parse ok");
        assert_eq!(r2.len(), 4);
        // 验证 Target 保留原始相对路径
        assert_eq!(
            r2.get("rIdDgmData1").unwrap().target.as_str(),
            "../diagrams/data1.xml"
        );
        assert_eq!(
            r2.get("rIdDgmColors1").unwrap().target.as_str(),
            "../diagrams/colors1.xml"
        );
    }
}
