//! # 演示文稿（Presentation）—— 顶层 API
//!
//! 对标 python-pptx 的 `Presentation` 类，是用户与本库交互的**唯一入口**：
//!
//! - 通过 [`Presentation::new`] / [`Presentation::open`] 创建或加载 `.pptx`；
//! - 通过 [`Presentation::slides_mut`] 增删/编辑幻灯片；
//! - 通过 [`Presentation::save`] / [`Presentation::to_bytes`] 序列化输出。
//!
//! # 内部数据
//!
//! `Presentation` 内部维护三类"集合"：
//!
//! - [`Slides`]：用户实际编辑的幻灯片列表（[`Slide`]）；
//! - [`SlideLayouts`]：版式列表（每个 `Slide` 通过 `rId` 引用其中一个）；
//! - [`SlideMasters`]：母版列表（版式再引用母版，呈现"主-版-页"三层结构）。
//!
//! 这三者**不**与 python-pptx 完全等价（python-pptx 用方法 `slide_layouts[i]` 暴露），
//! 而是直接以集合形式提供，便于将来加入更复杂的工作流。
//!
//! # 在三层架构中的位置
//!
//! `Presentation` 属于**高阶 API 层**。它直接聚合 `Slides` / `SlideLayouts` /
//! `SlideMasters` 三个集合；当用户调用 `save` 时，由 `to_opc_package` 把这些内存模型
//! 序列化为 [`crate::opc::OpcPackage`]，再交给 zip 写出 `.pptx`。
//!
//! # 单元
//!
//! 幻灯片尺寸字段（`width` / `height`）一律使用 [`Emu`]，与 OOXML XML 中
//! `<p:sldSz cx="..." cy="..."/>` 一致。常用换算：1 in = 914 400 EMU，
//! 1 cm = 360 000 EMU。
//!
//! # 示例
//!
//! ```no_run
//! use pptx_rs::Presentation;
//! use pptx_rs::Inches;
//!
//! let mut p = Presentation::new().unwrap();
//! let counter = p.id_counter();
//! let s = p.slides_mut().add_slide(counter).unwrap();
//! s.shapes_mut().add_textbox_with_text(
//!     Inches(1.0), Inches(1.0), Inches(4.0), Inches(1.0),
//!     "Hello, pptx-rs!",
//! ).unwrap();
//! p.save("out.pptx").unwrap();
//! ```
//!
//! （doctest 用 `unwrap` 仅为缩短示例；生产代码应使用 `?` 传播错误。）

use std::cell::Cell;
use std::io::Read;
use std::path::Path;
use std::rc::Rc;

use crate::notes_masters::{NotesMasterRef, NotesMasters};
use crate::opc::package::{ct, rels_partname_for};
use crate::opc::part::{new_part_name, Part, PartName};
use crate::opc::rels::{RelType, Relationship, Relationships};
use crate::opc::OpcPackage;
use crate::oxml::presentation::{PresentationRoot, SlideIdEntry};
use crate::oxml::slidelayout::SldLayout as OxmlSldLayout;
use crate::oxml::slidemaster::SldMaster as OxmlSldMaster;
use crate::oxml::theme::default_theme_xml;
use crate::slide::{Slide, SlideEntry, Slides};
use crate::slide_layouts::{SlideLayoutRef, SlideLayouts};
use crate::slide_masters::{SlideMasterRef, SlideMasters};
use crate::units::Emu;

/// 默认演示文稿宽度（EMU）。对应 10 英寸（4:3 比例）—— 1 in = 914 400 EMU。
pub const DEFAULT_WIDTH_EMU: i64 = 9_144_000;

/// 默认演示文稿高度（EMU）。对应 7.5 英寸 —— 1 in = 914 400 EMU。
pub const DEFAULT_HEIGHT_EMU: i64 = 6_858_000;

/// 演示文稿（内存模型）。
///
/// # 字段语义
///
/// - `slides` / `slide_layouts` / `slide_masters`：三类"主-版-页"集合；
/// - `width` / `height`：画布尺寸，序列化为 `<p:sldSz cx="..." cy="..."/>`；
/// - `id_counter`：共享 ID 计数器（每张 slide / 每个 shape 都会 `next_shape_id`）；
/// - **媒体**（图片等 blob）**按 slide 存储**——参见 `Slide::register_media`。
///   保存时由 `to_opc_package` 遍历 `self.slides` 聚合写入 zip；同一 partname 只写一次。
#[derive(Debug)]
pub struct Presentation {
    /// 幻灯片集合（用户主要操作的子结构）。
    pub(crate) slides: Slides,
    /// 版式集合（与 `slide_masters` 一起构成"主-版-页"分层）。
    pub(crate) slide_layouts: SlideLayouts,
    /// 母版集合。
    pub(crate) slide_masters: SlideMasters,
    /// 备注母版集合（TODO-045）。
    ///
    /// 一个演示文稿通常只有 0 或 1 个备注母版。
    /// `Presentation::new` 创建的空白文档无此 part；
    /// `from_opc` 会从 `presentation.xml.rels` 中的 `NotesMaster` 关系解析。
    pub(crate) notes_masters: NotesMasters,
    /// 主题（`<a:theme>`）。
    ///
    /// TODO-001：从已有 PPTX 解析后存储，写路径使用 `self.theme.to_xml()`
    /// 而非 `default_theme_xml()`，实现 read→save 保真。
    pub(crate) theme: crate::oxml::theme::Theme,
    /// 母版 ID 列表（`<p:sldMasterIdLst>`，从 presentation.xml 解析）。
    ///
    /// TODO-001：用于写路径 `PresentationRoot.sld_master_ids`，支持多母版。
    /// 空列表时写路径使用默认的单个母版。
    pub(crate) sld_master_ids: Vec<(u32, String)>,
    /// 画布宽度（EMU）。
    pub(crate) width: Emu,
    /// 画布高度（EMU）。
    pub(crate) height: Emu,
    /// 共享 shape id 计数器（`Rc<Cell<u32>>` 以便跨多个集合共享）。
    pub(crate) id_counter: Rc<Cell<u32>>,
    /// 文档核心属性（对标 pypdf `DocumentInformation` / OOXML `core-properties`）。
    pub(crate) core_properties: CoreProperties,
    /// 自定义文档属性（`/docProps/custom.xml`，TODO-034）。
    ///
    /// 用户自定义的键值对，在 `to_opc_package` 中非空时序列化为 custom.xml。
    pub(crate) custom_properties: CustomProperties,
    /// 评论作者列表（`/ppt/commentAuthors.xml`，TODO-036）。
    ///
    /// 全局共享的评论作者清单。当任意 slide 有评论时，`to_opc_package`
    /// 会写出 `commentAuthors.xml` 并在 `_rels/.rels` 添加关系。
    pub(crate) comment_authors: crate::oxml::comments::CommentAuthorList,
    /// 修改密码保护参数（`<p:modifyVerifier>`）。
    ///
    /// 设置后，`to_opc_package` 会在 `presentation.xml` 中注入 modifyVerifier 元素，
    /// WPS / PowerPoint 打开时提示"以只读方式打开"或"输入密码以修改"。
    pub(crate) modify_protection: Option<crate::crypto::ModifyProtection>,
    /// 章节分组列表（`<p14:sectionLst>`，TODO-039）。
    ///
    /// 用于把若干 slide 归入命名分组（PowerPoint 大纲视图中的"节"）。
    /// 空列表时不输出 sectionLst 扩展；非空时由 `to_opc_package` 在
    /// `presentation.xml` 的 `<p:extLst>` 内写出。
    pub(crate) sections: crate::oxml::section::SectionList,
}

/// 文档核心属性（对标 pypdf `DocumentInformation` + OOXML `core-properties`）。
///
/// 对应 `docProps/core.xml` 和 `docProps/app.xml` 中的元数据字段。
/// 在 `to_opc_package` 中会序列化为这两个 XML part。
///
/// # 与 pypdf 的对应
///
/// | pypdf `DocumentInformation` | 本结构体字段 | OOXML 路径 |
/// |---|---|---|
/// | `title` | `title` | `dc:title` |
/// | `author` | `creator` | `dc:creator` |
/// | `subject` | `subject` | `dc:subject` |
/// | `creator` | `last_modified_by` | `cp:lastModifiedBy` |
/// | `producer` | `application` | `Application` (app.xml) |
/// | `creation_date` | `created` | `dcterms:created` |
/// | `modification_date` | `modified` | `dcterms:modified` |
/// | `keywords` | `keywords` | `cp:keywords` |
#[derive(Debug, Clone, Default)]
pub struct CoreProperties {
    /// 文档标题。
    pub title: Option<String>,
    /// 文档作者/创建者。
    pub creator: Option<String>,
    /// 文档主题。
    pub subject: Option<String>,
    /// 最后修改者。
    pub last_modified_by: Option<String>,
    /// 创建应用程序（如 "pptx-rs" / "Microsoft Office PowerPoint"）。
    pub application: Option<String>,
    /// 创建时间（ISO 8601 格式，如 "2026-06-14T12:00:00Z"）。
    pub created: Option<String>,
    /// 最后修改时间（ISO 8601 格式）。
    pub modified: Option<String>,
    /// 关键词（逗号分隔）。
    pub keywords: Option<String>,
    /// 分类。
    pub category: Option<String>,
    /// 备注/描述。
    pub description: Option<String>,
    /// 修订号。
    pub revision: Option<String>,
}

/// 自定义属性值类型（`/docProps/custom.xml` 中的 `vt:*` 元素）。
///
/// 对应 OOXML VTypes 命名空间下的几种常用类型。
/// 每个变体序列化为对应的 `<vt:xxx>` 元素。
///
/// # OOXML 结构
///
/// ```text
/// <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Key">
///   <vt:lpwstr>Value</vt:lpwstr>
/// </property>
/// ```
#[derive(Clone, Debug, PartialEq)]
pub enum CustomPropertyValue {
    /// 字符串（`<vt:lpwstr>`）。
    Text(String),
    /// 32 位整数（`<vt:i4>`）。
    Int(i32),
    /// 双精度浮点（`<vt:r8>`）。
    Float(f64),
    /// 布尔（`<vt:bool>`）。
    Bool(bool),
    /// 日期时间（`<vt:filetime>`，ISO 8601 格式字符串）。
    DateTime(String),
}

impl CustomPropertyValue {
    /// 返回对应的 `<vt:xxx>` 元素名。
    fn vt_element(&self) -> &'static str {
        match self {
            CustomPropertyValue::Text(_) => "vt:lpwstr",
            CustomPropertyValue::Int(_) => "vt:i4",
            CustomPropertyValue::Float(_) => "vt:r8",
            CustomPropertyValue::Bool(_) => "vt:bool",
            CustomPropertyValue::DateTime(_) => "vt:filetime",
        }
    }

    /// 返回值的字符串表示（用于序列化到 `<vt:xxx>` 元素文本）。
    fn value_str(&self) -> String {
        match self {
            CustomPropertyValue::Text(s) => s.clone(),
            CustomPropertyValue::Int(i) => i.to_string(),
            CustomPropertyValue::Float(f) => f.to_string(),
            CustomPropertyValue::Bool(b) => if *b { "true" } else { "false" }.to_string(),
            CustomPropertyValue::DateTime(s) => s.clone(),
        }
    }
}

/// 自定义文档属性集合（`/docProps/custom.xml`）。
///
/// 对标 python-pptx 中 `Presentation.custom_properties`（v1.0+）。
/// 存储用户自定义的键值对，PowerPoint 在"文件 → 信息 → 属性 → 高级属性 → 自定义"中显示。
///
/// # 序列化
///
/// `to_opc_package` 会在 `custom_properties` 非空时写出 `/docProps/custom.xml`，
/// 并在 `_rels/.rels` 中添加 `custom-properties` 关系。
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::Presentation;
/// use pptx_rs::presentation::CustomPropertyValue;
///
/// let mut p = Presentation::new().unwrap();
/// p.custom_properties_mut().set("Project", CustomPropertyValue::Text("Demo".to_string()));
/// p.custom_properties_mut().set("Version", CustomPropertyValue::Int(42));
/// p.save("out.pptx").unwrap();
/// ```
#[derive(Debug, Clone, Default)]
pub struct CustomProperties {
    /// 有序键值对列表（保留插入顺序，便于稳定序列化）。
    entries: Vec<(String, CustomPropertyValue)>,
}

impl CustomProperties {
    /// 创建空的自定义属性集合。
    pub fn new() -> Self {
        Self::default()
    }

    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// 返回条目数。
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// 设置（或覆盖）一个自定义属性。
    ///
    /// # 参数
    /// - `name`：属性名（不可为空）；
    /// - `value`：属性值。
    pub fn set(&mut self, name: impl Into<String>, value: CustomPropertyValue) {
        let name = name.into();
        for entry in &mut self.entries {
            if entry.0 == name {
                entry.1 = value;
                return;
            }
        }
        self.entries.push((name, value));
    }

    /// 取指定名称的属性值。
    pub fn get(&self, name: &str) -> Option<&CustomPropertyValue> {
        self.entries.iter().find(|(k, _)| k == name).map(|(_, v)| v)
    }

    /// 移除指定名称的属性。
    ///
    /// 返回被移除的值（若存在）。
    pub fn remove(&mut self, name: &str) -> Option<CustomPropertyValue> {
        let idx = self.entries.iter().position(|(k, _)| k == name)?;
        Some(self.entries.remove(idx).1)
    }

    /// 返回所有条目的迭代器。
    pub fn iter(&self) -> impl Iterator<Item = (&str, &CustomPropertyValue)> {
        self.entries.iter().map(|(k, v)| (k.as_str(), v))
    }

    /// 序列化为 `/docProps/custom.xml` 的 XML 字符串。
    ///
    /// # XML 结构
    ///
    /// ```text
    /// <Properties xmlns="...custom-properties" xmlns:vt="...docPropsVTypes">
    ///   <property fmtid="{D5CDD505-...}" pid="2" name="Key1">
    ///     <vt:lpwstr>Value1</vt:lpwstr>
    ///   </property>
    ///   ...
    /// </Properties>
    /// ```
    ///
    /// `pid` 从 2 开始（1 保留给 SummaryInformation）。
    pub fn to_xml(&self) -> String {
        if self.entries.is_empty() {
            return String::new();
        }
        let mut out = String::with_capacity(256 * self.entries.len());
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        out.push_str("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\"");
        out.push_str(
            " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n",
        );
        for (i, (name, value)) in self.entries.iter().enumerate() {
            let pid = (i + 2) as u32; // pid 从 2 开始
            let vt_elem = value.vt_element();
            let val = xml_escape(&value.value_str());
            out.push_str(&format!(
                "  <property fmtid=\"{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}\" pid=\"{}\" name=\"{}\">\n",
                pid,
                xml_escape(name)
            ));
            out.push_str(&format!("    <{}>{}</{}>\n", vt_elem, val, vt_elem));
            out.push_str("  </property>\n");
        }
        out.push_str("</Properties>");
        out
    }

    /// 从 `/docProps/custom.xml` 的 XML 字符串解析。
    ///
    /// # 参数
    /// - `xml`：custom.xml 的完整内容。
    ///
    /// # 返回值
    /// 解析出的 `CustomProperties`。解析失败的条目会被跳过（不中断整体解析）。
    pub fn from_xml(xml: &str) -> Self {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut props = CustomProperties::new();
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        // 当前正在解析的 <property> 的 name 属性
        let mut cur_name: Option<String> = None;
        // 当前正在解析的 <vt:xxx> 元素名
        let mut cur_vt: Option<String> = None;
        // 当前累积的文本内容
        let mut cur_text = String::new();

        loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    let qname = String::from_utf8_lossy(name.as_ref()).to_string();
                    let local = local_name(name.as_ref());
                    if local == b"property" {
                        // 提取 name 属性
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"name" {
                                cur_name = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if cur_name.is_some() && qname.starts_with("vt:") {
                        // 用完整 QName（含 vt: 前缀）标识 VTypes 元素
                        cur_vt = Some(qname);
                        cur_text.clear();
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"property" {
                        // 自闭合 <property/>：无值，跳过
                        cur_name = None;
                    }
                }
                Ok(Event::Text(t)) => {
                    if cur_vt.is_some() {
                        cur_text.push_str(&t.decode().unwrap_or_default());
                    }
                }
                Ok(Event::End(e)) => {
                    let name = e.name();
                    let qname = String::from_utf8_lossy(name.as_ref()).to_string();
                    let local = local_name(name.as_ref());
                    if qname.starts_with("vt:") && cur_vt.is_some() {
                        // 结束 <vt:xxx> 元素，构造值
                        if let Some(prop_name) = cur_name.take() {
                            let value = parse_vt_value(&cur_vt.take().unwrap(), &cur_text);
                            if let Some(v) = value {
                                props.set(prop_name, v);
                            }
                        }
                        cur_vt = None;
                        cur_text.clear();
                    } else if local == b"property" {
                        cur_name = None;
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }
        props
    }
}

/// XML 特殊字符转义（`&` / `<` / `>` / `"` / `'`）。
fn xml_escape(s: &str) -> String {
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

/// 取命名空间本地名（`a:off` → `off`）。
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

/// 根据 `<vt:xxx>` 元素名和文本内容构造 `CustomPropertyValue`。
fn parse_vt_value(vt_elem: &str, text: &str) -> Option<CustomPropertyValue> {
    let trimmed = text.trim();
    match vt_elem {
        "vt:lpwstr" | "vt:bstr" => Some(CustomPropertyValue::Text(trimmed.to_string())),
        "vt:i4" | "vt:int" | "vt:i2" | "vt:ui1" | "vt:ui2" | "vt:ui4" => {
            trimmed.parse::<i32>().ok().map(CustomPropertyValue::Int)
        }
        "vt:i8" | "vt:ui8" => {
            // 64 位整数截断为 i32（自定义属性中罕见超大值）
            trimmed.parse::<i64>().ok().and_then(|v| {
                if v >= i32::MIN as i64 && v <= i32::MAX as i64 {
                    Some(CustomPropertyValue::Int(v as i32))
                } else {
                    None
                }
            })
        }
        "vt:r4" | "vt:r8" | "vt:decimal" => {
            trimmed.parse::<f64>().ok().map(CustomPropertyValue::Float)
        }
        "vt:bool" => {
            let v = trimmed.eq_ignore_ascii_case("true") || trimmed == "1";
            if v || trimmed.eq_ignore_ascii_case("false") || trimmed == "0" {
                Some(CustomPropertyValue::Bool(v))
            } else {
                None
            }
        }
        "vt:filetime" | "vt:date" => Some(CustomPropertyValue::DateTime(trimmed.to_string())),
        _ => None,
    }
}

/// 媒体条目：保存到 `/ppt/media/<file>` 目录中的二进制资源。
///
/// 对应 python-pptx 中 `SlideShapes.add_picture(...)` 隐式管理的 part。
/// `rid` 是**预先生成的**关系 id（如 `"rIdImg1"`），用于在 `slideN.xml.rels` 中
/// 显式添加 `<Relationship Id="rIdImg1" Type="...image" Target="../media/img.png"/>`。
#[derive(Debug, Clone)]
pub struct MediaEntry {
    /// 媒体 part 路径，例如 `/ppt/media/image1.png`。
    pub partname: PartName,
    /// 媒体 MIME / Office Content-Type（如 `image/png`）。
    pub content_type: String,
    /// 二进制内容（图片像素、嵌入字体字节等）。
    pub blob: Vec<u8>,
    /// 在 slide xml 中引用的关系 id（形如 `rIdImg1`）。
    pub rid: String,
}

/// 图表条目：保存到 `/ppt/charts/chartN.xml` 的 chart part 元数据。
///
/// 对应 python-pptx 中 `SlideShapes.add_chart(...)` 隐式管理的 chart part。
/// 在 `to_opc_package` 阶段，每个 `ChartEntry` 会：
/// 1. 写出独立的 `/ppt/charts/chartN.xml` part（内容为 `Chart::to_xml()`）；
/// 2. 在 `slideN.xml.rels` 中添加 `<Relationship Type=".../chart" Target="../charts/chartN.xml"/>`；
/// 3. 把 `rid` 同步回 `ChartShape.frame.graphic.Chart.rid`，供 `<c:chart r:id="..."/>` 引用。
///
/// # 与 MediaEntry 的差异
/// - MediaEntry 持有 `blob: Vec<u8>`（二进制）；
/// - ChartEntry 持有 `chart: OxmlChart`（强类型模型），写出时调用 `chart.to_xml()`。
#[derive(Debug, Clone)]
pub struct ChartEntry {
    /// chart part 路径，例如 `/ppt/charts/chart1.xml`。
    pub partname: PartName,
    /// 强类型 Chart 模型（包含类型/数据/标题）。
    pub chart: crate::oxml::chart::Chart,
    /// 在 slide xml 中引用的关系 id（形如 `rIdChart1`）。
    pub rid: String,
    /// 嵌入式 Excel 工作簿二进制内容（TODO-004 Excel 嵌入）。
    ///
    /// - `None`：图表数据仅靠 numCache/strCache，不嵌入 Excel；
    /// - `Some(bytes)`：写出时生成 `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`
    ///   part + `chartN.xml.rels` 关系（Type=Package），并在 chart XML 中写入
    ///   `<c:externalData r:id="rIdXlsxN"/>` 引用。PowerPoint 打开图表时会从
    ///   该 xlsx part 读取数据源（"编辑数据" 启动 Excel）。
    ///
    /// 内容应为有效的 `.xlsx` 文件字节流（OOXML SpreadsheetML 包格式）。
    /// 库不校验内容有效性，PowerPoint 会按 zip + XML 解析。
    pub xlsx_blob: Option<Vec<u8>>,
}

/// OLE 对象条目：保存到 `/ppt/embeddings/oleObjectN.bin` 的 OLE part 元数据（TODO-043）。
///
/// 对应 python-pptx 中 `SlideShapes.add_ole_object(...)` 隐式管理的 oleObject part。
/// 在 `to_opc_package` 阶段，每个 `OleEntry` 会：
/// 1. 写出独立的 `/ppt/embeddings/oleObjectN.bin` part（内容为原始 OLE 二进制 blob）；
/// 2. 在 `slideN.xml.rels` 中添加 `<Relationship Type=".../oleObject" Target="../embeddings/oleObjectN.bin"/>`；
/// 3. 把 `rid` 同步回 `OleObjectShape.frame.graphic.OleObject.rid`，供 `<p:oleObj r:id="..."/>` 引用。
///
/// # 与 MediaEntry 的差异
///
/// - MediaEntry 持有图片二进制，用于 `<p:pic>` 的 `<a:blip r:embed="..."/>`；
/// - OleEntry 持有 OLE 复合文档二进制，用于 `<p:oleObj r:id="..."/>`；
/// - 两者 partname 命名空间不同（`/ppt/media/` vs `/ppt/embeddings/`）。
///
/// # 与 ChartEntry 的差异
///
/// - ChartEntry 持有强类型 Chart 模型，写出时调用 `chart.to_xml()` 生成 XML；
/// - OleEntry 持有原始二进制 blob，写出时直接写入字节，**不**经过 XML 序列化。
#[derive(Debug, Clone)]
pub struct OleEntry {
    /// oleObject part 路径，例如 `/ppt/embeddings/oleObject1.bin`。
    pub partname: PartName,
    /// 原始 OLE 二进制数据（CFB 复合文档）。
    pub blob: Vec<u8>,
    /// 在 slide xml 中引用的关系 id（形如 `rIdOle1`）。
    pub rid: String,
}

/// 视频媒体条目：保存到 `/ppt/media/mediaN.mp4` 的视频 part 元数据（TODO-033）。
///
/// 对应 python-pptx 中 `SlideShapes.add_movie(...)` / `add_video(...)` 隐式管理的 media part。
/// 在 `to_opc_package` 阶段，每个 `VideoEntry` 会：
/// 1. 写出独立的 `/ppt/media/mediaN.mp4` part（内容为原始视频二进制 blob）；
/// 2. 在 `slideN.xml.rels` 中添加 `<Relationship Type=".../video" Target="../media/mediaN.mp4"/>`；
/// 3. 把 `rid` 同步回 `Picture.pic.media`（`MediaKind::Video { rid }`），供
///    `<a:videoFile r:link="..."/>` 引用。
///
/// # 与 MediaEntry 的差异
///
/// - MediaEntry 持有**海报帧图片**二进制，关系类型为 `.../image`，用 `r:embed` 引用；
/// - VideoEntry 持有**视频文件**二进制，关系类型为 `.../video`，用 `r:link` 引用；
/// - 两者 partname 都在 `/ppt/media/` 下，但文件名命名空间区分（`imageN.png` vs `mediaN.mp4`）。
///
/// # 与 OleEntry 的差异
///
/// - OleEntry 的 partname 在 `/ppt/embeddings/` 下，关系类型为 `.../oleObject`；
/// - VideoEntry 的 partname 在 `/ppt/media/` 下，关系类型为 `.../video`；
/// - OleEntry 通过 `<p:oleObj r:id="..."/>` 引用（`r:id`），VideoEntry 通过
///   `<a:videoFile r:link="..."/>` 引用（`r:link`）。
#[derive(Debug, Clone)]
pub struct VideoEntry {
    /// 视频 part 路径，例如 `/ppt/media/media1.mp4`。
    pub partname: PartName,
    /// 原始视频二进制数据（如 MP4 字节流）。
    pub blob: Vec<u8>,
    /// 在 slide xml 中引用的关系 id（形如 `rIdVideo1`）。
    ///
    /// 该 rid 会写入 `slideN.xml.rels` 的 `<Relationship Id="rIdVideoN" Type=".../video"/>`，
    /// 同时同步到 `Picture.pic.media`（`MediaKind::Video { rid }`），
    /// 供 `<a:videoFile r:link="rIdVideoN"/>` 引用。
    pub rid: String,
}

/// 音频媒体条目：保存到 `/ppt/media/mediaN.mp3` 的音频 part 元数据（TODO-033）。
///
/// 与 [`VideoEntry`] 结构完全对称，仅媒体类型与 Content-Type 不同。
/// 在 `to_opc_package` 阶段，每个 `AudioEntry` 会：
/// 1. 写出独立的 `/ppt/media/mediaN.mp3` part（内容为原始音频二进制 blob）；
/// 2. 在 `slideN.xml.rels` 中添加 `<Relationship Type=".../audio" Target="../media/mediaN.mp3"/>`；
/// 3. 把 `rid` 同步回 `Picture.pic.media`（`MediaKind::Audio { rid }`），供
///    `<a:audioFile r:link="..."/>` 引用。
#[derive(Debug, Clone)]
pub struct AudioEntry {
    /// 音频 part 路径，例如 `/ppt/media/media1.mp3`。
    pub partname: PartName,
    /// 原始音频二进制数据（如 MP3 字节流）。
    pub blob: Vec<u8>,
    /// 在 slide xml 中引用的关系 id（形如 `rIdAudio1`）。
    pub rid: String,
}

/// SmartArt（diagram）条目：保存到 `/ppt/diagrams/` 下的 4 个 part 元数据（TODO-037）。
///
/// 对应 python-pptx 中 `SlideShapes.add_shape(...)` 隐式管理的 SmartArt parts。
/// 在 OOXML 中，一个 SmartArt 图形由 4 个独立 part 组成：
///
/// | part | 路径 | 关系类型 | `dgm:relIds` 属性 |
/// |---|---|---|---|
/// | data | `/ppt/diagrams/dataN.xml` | `.../diagramData` | `r:dm` |
/// | layout | `/ppt/diagrams/layoutN.xml` | `.../diagramLayout` | `r:lo` |
/// | quickStyle | `/ppt/diagrams/quickStylesN.xml` | `.../diagramQuickStyle` | `r:qs` |
/// | colors | `/ppt/diagrams/colorsN.xml` | `.../diagramColors` | `r:cs` |
///
/// 被 slide 的 `<p:graphicFrame>` 内 `<a:graphicData uri="...diagram"><dgm:relIds .../>` 引用。
///
/// # Round-trip 策略
///
/// 当前实现采用**完整 round-trip**：读路径保留 4 个 part 的原始 XML 字符串，
/// 写路径直接写入保留的 XML（不重新序列化）。这样保证：
/// - 任何 SmartArt 模板/布局/颜色变体都能正确保留；
/// - 不依赖完整的 diagram schema 实现（避免数千行代码）。
///
/// # 与 ChartEntry / OleEntry 的差异
///
/// - ChartEntry 持有强类型 Chart 模型，写出时调用 `chart.to_xml()`；
/// - OleEntry 持有二进制 blob；
/// - DiagramEntry 持有 4 份 XML 字符串，直接写入 zip（不经过序列化）。
#[derive(Debug, Clone)]
pub struct DiagramEntry {
    /// data part 路径，例如 `/ppt/diagrams/data1.xml`。
    pub data_partname: PartName,
    /// layout part 路径，例如 `/ppt/diagrams/layout1.xml`。
    pub layout_partname: PartName,
    /// quickStyle part 路径，例如 `/ppt/diagrams/quickStyles1.xml`。
    pub quick_style_partname: PartName,
    /// colors part 路径，例如 `/ppt/diagrams/colors1.xml`。
    pub colors_partname: PartName,
    /// 原始 data XML 字符串（`<dgm:dataModel>...</dgm:dataModel>`）。
    pub data_xml: String,
    /// 原始 layout XML 字符串（`<dgm:layoutDef>...</dgm:layoutDef>`）。
    pub layout_xml: String,
    /// 原始 quickStyle XML 字符串（`<dgm:styleData>...</dgm:styleData>`）。
    pub quick_style_xml: String,
    /// 原始 colors XML 字符串（`<dgm:colorsDef>...</dgm:colorsDef>`）。
    pub colors_xml: String,
    /// 在 slide xml 中引用 data part 的关系 id（形如 `rIdDgmData1`）。
    ///
    /// 对应 `<dgm:relIds r:dm="rIdDgmData1"/>`。
    pub data_rid: String,
    /// 在 slide xml 中引用 layout part 的关系 id（形如 `rIdDgmLayout1`）。
    ///
    /// 对应 `<dgm:relIds r:lo="rIdDgmLayout1"/>`。
    pub layout_rid: String,
    /// 在 slide xml 中引用 quickStyle part 的关系 id（形如 `rIdDgmQs1`）。
    ///
    /// 对应 `<dgm:relIds r:qs="rIdDgmQs1"/>`。
    pub quick_style_rid: String,
    /// 在 slide xml 中引用 colors part 的关系 id（形如 `rIdDgmColors1`）。
    ///
    /// 对应 `<dgm:relIds r:cs="rIdDgmColors1"/>`。
    pub colors_rid: String,
}

impl DiagramEntry {
    /// 按需解析 data part 为强类型 [`crate::oxml::diagram::DataModel`]（TODO-037）。
    ///
    /// `DiagramEntry` 默认以 `String` blob 持有原始 XML，保证 byte-exact round-trip。
    /// 当调用方需要结构化访问 SmartArt 节点（如查询节点文本、遍历父子关系）时，
    /// 调用本方法触发按需解析。
    ///
    /// # 返回值
    /// - 成功：返回 [`crate::oxml::diagram::DataModel`]，含所有节点与连接。
    /// - 失败：返回 `Error::Xml`（data_xml 畸形时）。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx_rs::presentation::DiagramEntry;
    /// # let entry: DiagramEntry = unimplemented!();
    /// let data_model = entry.data_model().expect("parse data part");
    /// for pt in &data_model.points {
    ///     println!("node {}: {:?}", pt.model_id, pt.text);
    /// }
    /// ```
    pub fn data_model(&self) -> crate::Result<crate::oxml::diagram::DataModel> {
        crate::oxml::diagram::DataModel::parse_from_xml(&self.data_xml)
    }

    /// 按需解析 layout part 为强类型 [`crate::oxml::diagram::LayoutDef`]。
    ///
    /// 返回元数据（uniqueId / title / desc / catLst）+ layoutNode 子树原始 XML。
    pub fn layout_def(&self) -> crate::Result<crate::oxml::diagram::LayoutDef> {
        crate::oxml::diagram::LayoutDef::parse_from_xml(&self.layout_xml)
    }

    /// 按需解析 quickStyle part 为强类型 [`crate::oxml::diagram::QuickStyleDef`]。
    ///
    /// 返回 styleLbl 列表（仅 name + 原始 XML）。
    pub fn quick_style_def(&self) -> crate::Result<crate::oxml::diagram::QuickStyleDef> {
        crate::oxml::diagram::QuickStyleDef::parse_from_xml(&self.quick_style_xml)
    }

    /// 按需解析 colors part 为强类型 [`crate::oxml::diagram::ColorsDef`]。
    ///
    /// 返回元数据 + styleClrLbl 列表（仅 name + 原始 XML）。
    pub fn colors_def(&self) -> crate::Result<crate::oxml::diagram::ColorsDef> {
        crate::oxml::diagram::ColorsDef::parse_from_xml(&self.colors_xml)
    }

    /// 把修改后的 [`crate::oxml::diagram::DataModel`] 写回 `data_xml` 字段（TODO-037 文本节点编辑）。
    ///
    /// 典型用法：调用方先通过 [`DiagramEntry::data_model`] 解析得到 `DataModel`，
    /// 修改节点文本（如 `set_point_text`），再调用本方法把修改写回 `data_xml`，
    /// 后续 `Presentation::save` 会把更新后的 `data_xml` 写入 zip。
    ///
    /// # 参数
    /// - `data_model`：修改后的 DataModel 实例。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx_rs::presentation::DiagramEntry;
    /// # let entry: DiagramEntry = unimplemented!();
    /// let mut dm = entry.data_model().expect("parse");
    /// dm.set_point_text(1, "新文本");
    /// entry.set_data_model(&dm);
    /// ```
    pub fn set_data_model(&mut self, data_model: &crate::oxml::diagram::DataModel) {
        self.data_xml = data_model.to_xml();
    }

    /// 便捷方法：直接修改指定 `model_id` 节点的文本并写回 `data_xml`。
    ///
    /// 内部流程：解析 data_xml → 修改节点文本 → 序列化回 data_xml。
    ///
    /// # 参数
    /// - `model_id`：目标节点 ID；
    /// - `new_text`：新的文本内容。
    ///
    /// # 返回值
    /// - `Ok(true)`：找到节点并成功修改；
    /// - `Ok(false)`：未找到指定 `model_id` 的节点；
    /// - `Err`：data_xml 解析失败（畸形 XML）。
    pub fn set_point_text(
        &mut self,
        model_id: u32,
        new_text: impl Into<String>,
    ) -> crate::Result<bool> {
        let mut dm = self.data_model()?;
        let ok = dm.set_point_text(model_id, new_text);
        if ok {
            self.data_xml = dm.to_xml();
        }
        Ok(ok)
    }
}

impl Presentation {
    /// 新建一个**完全空白**的演示文稿。
    ///
    /// 内部会做两件事：
    ///
    /// 1. 把 `id_counter` 初始化为 `2`（保留 `1` 给"隐藏占位"用途，与 Office 一致）；
    /// 2. 调用 `Presentation::ensure_default_master_and_layout` 创建 1 个
    ///    默认母版 + 1 个空白版式，保证保存出的 `.pptx` 在 PowerPoint 中能正常打开。
    ///
    /// # 错误
    /// 当前实现不会失败；保留 `Result` 是为后续扩展（例如从模板加载 master）留口。
    pub fn new() -> crate::Result<Self> {
        let id_counter = Rc::new(Cell::new(2));
        let mut pres = Presentation {
            slides: Slides::new(),
            slide_layouts: SlideLayouts::new(),
            slide_masters: SlideMasters::new(),
            notes_masters: NotesMasters::new(),
            theme: crate::oxml::theme::Theme::default(),
            sld_master_ids: Vec::new(),
            width: Emu(DEFAULT_WIDTH_EMU),
            height: Emu(DEFAULT_HEIGHT_EMU),
            id_counter,
            core_properties: CoreProperties {
                application: Some("pptx-rs".to_string()),
                ..Default::default()
            },
            custom_properties: CustomProperties::default(),
            comment_authors: crate::oxml::comments::CommentAuthorList::default(),
            modify_protection: None,
            sections: crate::oxml::section::SectionList::default(),
        };
        pres.ensure_default_master_and_layout()?;
        Ok(pres)
    }

    /// 打开一个已存在的 `.pptx` 文件并解析为内存模型。
    ///
    /// 等价于 [`Presentation::load`]，提供"短名 + 详细名"两个入口。
    ///
    /// # 错误
    /// - [`crate::Error::Io`]：文件不存在或权限不足；
    /// - [`crate::Error::Zip`]：zip 损坏；
    /// - [`crate::Error::Xml`]：内部 XML 无法解析（**注意：当前 read 路径尚未完整实现**，
    ///   实际只解析 `[Content_Types].xml`，其它 part 暂未还原到 oxml 模型）。
    pub fn open(path: impl AsRef<Path>) -> crate::Result<Self> {
        Self::load(path.as_ref())
    }

    /// 加载 `.pptx` 的详细入口（与 [`Presentation::open`] 等价）。
    pub fn load(path: &Path) -> crate::Result<Self> {
        let pkg = OpcPackage::load(path)?;
        Self::from_opc(pkg)
    }

    /// 从内存中的 zip 字节流加载。
    ///
    /// 与 [`Presentation::load`] 的区别：本方法**不**触发磁盘 IO，可用于：
    ///
    /// - Web 服务接收 `multipart/form-data` 上传后的字节流；
    /// - 单元测试中 fixture 来自 `include_bytes!` 的场景；
    /// - 通过 `std::io::Cursor` 包装任意 `Read` 来源。
    ///
    /// # 错误
    /// - [`crate::Error::Io`]：`ZipArchive::new` 失败；
    /// - [`crate::Error::Xml`]：`[Content_Types].xml` 解析失败。
    pub fn load_bytes(bytes: &[u8]) -> crate::Result<Self> {
        let mut pkg = OpcPackage::new();
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor)?;
        // 第一步：先读 [Content_Types].xml —— 后续 part 的 Content-Type 都靠它推断。
        let mut ct_xml = String::new();
        zip.by_name("[Content_Types].xml")?
            .read_to_string(&mut ct_xml)?;
        pkg.content_types = crate::opc::package::parse_content_types_public(&ct_xml)?;
        // 第二步：把所有 part 装入 `parts` 表（关系文件统一走 RELATIONSHIPS content-type）。
        for i in 0..zip.len() {
            let mut e = zip.by_index(i)?;
            let name = e.name().to_string();
            if name == "[Content_Types].xml" || e.is_dir() {
                continue;
            }
            let mut blob = Vec::with_capacity(e.size() as usize);
            e.read_to_end(&mut blob)?;
            let ct_str: String = if name.ends_with(".rels") {
                crate::opc::package::ct::RELATIONSHIPS.to_string()
            } else {
                // 用 [Content_Types].xml 推断 partname → Content-Type：
                //   1) 先查 Override（精确 partname 匹配）
                //   2) 再查 Default（扩展名匹配）
                //   3) 都没有则回退 octet-stream
                let partname = format!("/{}", name);
                crate::opc::package::derive_content_type(&pkg.content_types, &partname)
            };
            let partname = format!("/{}", name);
            let p = Part::new(PartName::from_unchecked(partname), ct_str, blob);
            pkg.parts.insert(p.partname.as_str().to_string(), p);
        }
        Self::from_opc(pkg)
    }

    /// 从已构造好的 [`OpcPackage`] 创建 `Presentation`。
    ///
    /// 这是 [`Presentation::load`] / [`Presentation::load_bytes`] 的真正工作函数——
    /// 它会把 zip 中的 part 全部还原为内存模型，从而开启 **read-modify-write** 流程。
    ///
    /// # 还原范围
    ///
    /// | 部分 | 是否还原 | 说明 |
    /// |---|---|---|
    /// | `presentation.xml` → `sldIdLst` | ✅ | 决定 slide 顺序与 sld_id |
    /// | `presentation.xml` → `sldSz` | ✅ | 画布尺寸（回退默认 4:3） |
    /// | `presentation.xml.rels` → slide 关系 | ✅ | rid → slideN.xml partname |
    /// | `slideN.xml` → `Sld`（含 shapes） | ✅ | 走 [`crate::oxml::parse_sld::parse_sld`] |
    /// | `slideN.xml.rels` → layout 关系 | ✅ | 回填 `Slide::layout_rid` |
    /// | `slideN.xml.rels` → notes 关系 | ✅ | 若有则读 notesSlideN.xml |
    /// | `notesSlideN.xml` → 备注 `TextBody` | ✅ | 走 [`crate::oxml::parse_sld::parse_notes`] |
    /// | `notesSlideN.xml` 原始 partname | ✅ | 回填 `Slide::notes_partname`，save 时复用 |
    /// | `notesSlideN.xml.rels` → Slide target | ✅ | 回填 `Slide::notes_slide_rel_target`，save 时复用 |
    /// | `slideMasters` / `slideLayouts` / `theme` | ❌ | 路线图中：本版本以默认 master+layout 占位 |
    ///
    /// # 错误
    /// - [`crate::Error::Xml`]：任意 slide / notes / presentation XML 解析失败；
    /// - [`crate::Error::Opc`]：`presentation.xml.rels` 缺失或关键 rid 未找到。
    fn from_opc(pkg: OpcPackage) -> crate::Result<Self> {
        // 局部类型别名：减少 `diagram_tasks` 元组类型的视觉噪声（clippy::type_complexity）。
        // 字段顺序：(dm_rid, lo_rid, qs_rid, cs_rid, data_partname, layout_partname,
        //           quick_style_partname, colors_partname)
        type DiagramTask = (
            String,
            String,
            String,
            String,
            String,
            String,
            String,
            String,
        );
        // 共享 id_counter —— 后续读出的 slide 直接共用。
        let id_counter = Rc::new(Cell::new(2));
        let mut pres = Presentation {
            slides: Slides::new(),
            slide_layouts: SlideLayouts::new(),
            slide_masters: SlideMasters::new(),
            notes_masters: NotesMasters::new(),
            theme: crate::oxml::theme::Theme::default(),
            sld_master_ids: Vec::new(),
            width: Emu(DEFAULT_WIDTH_EMU),
            height: Emu(DEFAULT_HEIGHT_EMU),
            id_counter: id_counter.clone(),
            core_properties: CoreProperties::default(),
            custom_properties: CustomProperties::default(),
            comment_authors: crate::oxml::comments::CommentAuthorList::default(),
            modify_protection: None,
            sections: crate::oxml::section::SectionList::default(),
        };
        pres.ensure_default_master_and_layout()?;

        // ---------- 1) 读 presentation.xml ----------
        let pres_part = pkg
            .get_part("/ppt/presentation.xml")
            .ok_or_else(|| crate::Error::opc("presentation.xml not found in package"))?;
        let pres_xml = String::from_utf8_lossy(&pres_part.blob).into_owned();
        let (sld_id_list, sld_w, sld_h, sld_master_id_list, sections) =
            crate::oxml::parse_sld::parse_pres_root(&pres_xml)?;
        // TODO-001：存储解析出的 sldMasterIdLst，供写路径使用
        pres.sld_master_ids = sld_master_id_list;
        // TODO-039：存储解析出的 sectionLst，供写路径使用
        pres.sections = sections;
        if let (Some(w), Some(h)) = (sld_w, sld_h) {
            pres.width = w;
            pres.height = h;
        }

        // ---------- 2) 读 presentation.xml.rels 找 slide 关系 ----------
        let pres_rels_part = pkg
            .get_part("/ppt/_rels/presentation.xml.rels")
            .ok_or_else(|| crate::Error::opc("presentation.xml.rels not found"))?;
        let pres_rels_xml = String::from_utf8_lossy(&pres_rels_part.blob).into_owned();
        let pres_rels = Relationships::from_xml(&pres_rels_xml)?;

        // ---------- 2.5) 读取 slideMaster / slideLayout / theme ----------
        // 从 presentation.xml.rels 中找出所有 SlideMaster 关系，逐个解析。
        // 若找到至少一个 master，则清空默认的 master/layout，用解析到的替换。
        let mut parsed_masters: Vec<(String, String, OxmlSldMaster)> = Vec::new(); // (rid, partname, oxml)
        let mut parsed_layouts: Vec<(String, String, OxmlSldLayout)> = Vec::new(); // (rid, partname, oxml)
        for r in pres_rels.iter() {
            if !matches!(r.reltype, RelType::SlideMaster) {
                continue;
            }
            let master_partname =
                resolve_relative_partname("/ppt/presentation.xml", r.target.as_str());
            let master_xml = match pkg.get_part(master_partname.as_str()) {
                Some(p) => String::from_utf8_lossy(&p.blob).into_owned(),
                None => continue,
            };
            let master = match crate::oxml::parse_sld::parse_sld_master(&master_xml) {
                Ok(m) => m,
                Err(_e) => continue,
            };
            // 读 slideMasterN.xml.rels 找 SlideLayout 和 Theme 关系
            let master_rels_path = rels_partname_for(master_partname.as_str());
            if let Some(mrp) = pkg.get_part(master_rels_path.as_str()) {
                let mrp_xml = String::from_utf8_lossy(&mrp.blob).into_owned();
                if let Ok(master_rels) = Relationships::from_xml(&mrp_xml) {
                    for mr in master_rels.iter() {
                        if matches!(mr.reltype, RelType::SlideLayout) {
                            let layout_partname = resolve_relative_partname(
                                master_partname.as_str(),
                                mr.target.as_str(),
                            );
                            let layout_xml = match pkg.get_part(layout_partname.as_str()) {
                                Some(p) => String::from_utf8_lossy(&p.blob).into_owned(),
                                None => continue,
                            };
                            if let Ok(layout) =
                                crate::oxml::parse_sld::parse_sld_layout(&layout_xml)
                            {
                                parsed_layouts.push((mr.id.clone(), layout_partname, layout));
                            }
                        }
                        // Theme 关系：解析并存储到 Presentation（TODO-001：read→save 保真）
                        if matches!(mr.reltype, RelType::Theme) {
                            let theme_partname = resolve_relative_partname(
                                master_partname.as_str(),
                                mr.target.as_str(),
                            );
                            if let Some(tp) = pkg.get_part(theme_partname.as_str()) {
                                let theme_xml = String::from_utf8_lossy(&tp.blob).into_owned();
                                // 解析 theme 并存储到 Presentation（写路径使用 self.theme.to_xml()）
                                if let Ok(theme) = crate::oxml::parse_sld::parse_theme(&theme_xml) {
                                    pres.theme = theme;
                                }
                            }
                        }
                    }
                }
            }
            parsed_masters.push((r.id.clone(), master_partname, master));
        }
        // 若解析到 master/layout，则替换默认的
        if !parsed_masters.is_empty() {
            pres.slide_masters.items.clear();
            pres.slide_layouts.items.clear();
            for (rid, partname, oxml) in parsed_masters {
                let idx = pres.slide_masters.items.len();
                pres.slide_masters.items.push(SlideMasterRef {
                    idx,
                    partname,
                    rid,
                    oxml: Rc::new(std::cell::RefCell::new(oxml)),
                });
            }
            for (rid, partname, oxml) in parsed_layouts {
                let idx = pres.slide_layouts.items.len();
                pres.slide_layouts.items.push(SlideLayoutRef {
                    idx,
                    partname,
                    rid,
                    oxml: Rc::new(std::cell::RefCell::new(oxml)),
                });
            }
        }

        // ---------- 2.6) 读取 notesMaster（TODO-045） ----------
        // 从 presentation.xml.rels 中找出所有 NotesMaster 关系，逐个解析。
        // 与 slideMaster 不同，notesMaster 是可选 part——空白文档无此 part。
        for r in pres_rels.iter() {
            if !matches!(r.reltype, RelType::NotesMaster) {
                continue;
            }
            let nm_partname = resolve_relative_partname("/ppt/presentation.xml", r.target.as_str());
            let nm_xml = match pkg.get_part(nm_partname.as_str()) {
                Some(p) => String::from_utf8_lossy(&p.blob).into_owned(),
                None => continue,
            };
            let nm = match crate::oxml::parse_sld::parse_notes_master(&nm_xml) {
                Ok(m) => m,
                Err(_e) => continue,
            };
            let idx = pres.notes_masters.items.len();
            pres.notes_masters.items.push(NotesMasterRef {
                idx,
                partname: nm_partname,
                rid: r.id.clone(),
                oxml: Rc::new(std::cell::RefCell::new(nm)),
            });
        }

        // ---------- 3) 遍历 sldIdLst，还原每张 slide ----------
        // 维护"已读 layout rid"的最大编号，避免 id_counter 内部冲突。
        let mut max_id_seen: u32 = 2;
        for (sld_id, rid) in &sld_id_list {
            // 找 rid 对应的 target（如 "slides/slide1.xml"）。
            let rel = match pres_rels.get(rid) {
                Some(r) => r,
                None => {
                    // 关系缺失：跳过此 slide，不让单点失败拖垮整份文档。
                    continue;
                }
            };
            // presentation.xml 是 /ppt/presentation.xml，所以 rels 中的相对路径
            // 以 /ppt/ 为基准解析。绝对路径（以 '/' 开头）直接使用。
            let slide_partname =
                resolve_relative_partname("/ppt/presentation.xml", rel.target.as_str());
            // 读 slideN.xml
            let slide_part = match pkg.get_part(slide_partname.as_str()) {
                Some(p) => p,
                None => continue,
            };
            let slide_xml = String::from_utf8_lossy(&slide_part.blob).into_owned();
            let mut sld = match crate::oxml::parse_sld::parse_sld(&slide_xml) {
                Ok(s) => s,
                Err(_e) => {
                    // 单张 slide 解析失败：跳过、记日志。
                    continue;
                }
            };
            // 更新 sld.id
            sld.id = *sld_id;

            // ---------- 4) 读 slideN.xml.rels：layout + notes + image + comments ----------
            let rels_path = rels_partname_for(slide_partname.as_str());
            let mut layout_rid = String::from("rId1");
            let mut notes_rid: Option<String> = None;
            // notes target 的原始相对路径（如 "../notesSlides/notesSlide1.xml"），
            // 在遍历 slide_rels 时一并收集，避免重复解析 rels 文件。
            let mut notes_target_raw: Option<String> = None;
            // comments target 的原始相对路径（如 "../comments/comment1.xml"）。
            let mut comments_rid: Option<String> = None;
            let mut comments_target_raw: Option<String> = None;
            // 收集 image 关系以重建 media_entries（确保 save 时 rels 完整）
            let mut image_rels: Vec<(String, String)> = Vec::new(); // (rid, target_partname)
                                                                    // 收集 SmartArt 4 类 diagram 关系（rid -> 绝对 partname），用于后续构造 DiagramEntry（TODO-037）。
                                                                    // 一个 slide 可能含多个 SmartArt，每个 SmartArt 持有 4 个独立 rels，因此按 rid 建立扁平映射；
                                                                    // 后续根据 sld 中 SmartArtRef 的 4 个 rid 配对成 DiagramEntry。
            let mut diagram_rel_map: std::collections::HashMap<String, String> =
                std::collections::HashMap::new();
            // 收集 chart 关系（rid -> 绝对 partname），用于后续读取 chartN.xml 解析 Chart 模型（TODO-004 读路径）。
            // 一个 slide 可能含多个 chart graphicFrame，每个引用一个独立的 chartN.xml part。
            let mut chart_rel_map: std::collections::HashMap<String, String> =
                std::collections::HashMap::new();
            if let Some(rels_part) = pkg.get_part(rels_path.as_str()) {
                let rels_xml = String::from_utf8_lossy(&rels_part.blob).into_owned();
                if let Ok(slide_rels) = Relationships::from_xml(&rels_xml) {
                    for r in slide_rels.iter() {
                        match r.reltype {
                            RelType::SlideLayout => {
                                layout_rid = r.id.clone();
                            }
                            RelType::NotesSlide => {
                                notes_rid = Some(r.id.clone());
                                // 保留原始 target（相对路径），供后续 resolve_relative_partname 使用
                                notes_target_raw = Some(r.target.as_str().to_string());
                            }
                            RelType::Comments => {
                                comments_rid = Some(r.id.clone());
                                comments_target_raw = Some(r.target.as_str().to_string());
                            }
                            RelType::Image => {
                                let rel_target = r.target.as_str();
                                let abs =
                                    resolve_relative_partname(slide_partname.as_str(), rel_target);
                                image_rels.push((r.id.clone(), abs));
                            }
                            RelType::DiagramData
                            | RelType::DiagramLayout
                            | RelType::DiagramQuickStyle
                            | RelType::DiagramColors => {
                                // 收集 SmartArt 的 4 类关系（TODO-037），按 rid 建立映射。
                                let rel_target = r.target.as_str();
                                let abs =
                                    resolve_relative_partname(slide_partname.as_str(), rel_target);
                                diagram_rel_map.insert(r.id.clone(), abs);
                            }
                            RelType::Chart => {
                                // 收集 chart 关系（TODO-004 读路径），按 rid 建立映射。
                                // 后续根据 graphicFrame.Graphic::Chart.rid 查找 partname，
                                // 读取 chartN.xml 内容调用 Chart::parse_from_xml 还原模型。
                                let rel_target = r.target.as_str();
                                let abs =
                                    resolve_relative_partname(slide_partname.as_str(), rel_target);
                                chart_rel_map.insert(r.id.clone(), abs);
                            }
                            _ => {}
                        }
                    }
                }
            }
            // 把 layout_rid 同步到 sld。
            sld.set_layout_rid(layout_rid.clone());

            // ---------- 5) 处理 notes（如有） ----------
            // 收集本 slide 的 notes 元数据，用于写路径**复用**原始 OPC 关系，
            // 避免 read→save 后 partname / rid 漂移导致外部引用断链。
            //   - `parsed_notes_partname`：解析后的绝对 partname（如 `/ppt/notesSlides/notesSlide1.xml`）；
            //   - `parsed_notes_rels_target`：`notesSlideN.xml.rels` 中 `Slide` 关系的 target。
            let mut parsed_notes_partname: Option<String> = None;
            let mut parsed_notes_rels_target: Option<String> = None;
            if let Some(raw_target) = &notes_target_raw {
                // 从 slide rels 中取 target（相对路径如 "../notesSlides/notesSlide1.xml"），
                // 然后解析成绝对 partname。
                let notes_partname =
                    resolve_relative_partname(slide_partname.as_str(), raw_target.as_str());
                parsed_notes_partname = Some(notes_partname.clone());
                // 读 `notesSlideN.xml.rels`，找出其中指向所属 slide 的 `Slide` 关系 target。
                //   关系文件: /ppt/notesSlides/_rels/notesSlideN.xml.rels
                //   target 形如 "../slides/slide1.xml"。
                let notes_rels_path = rels_partname_for(&notes_partname);
                if let Some(nrp) = pkg.get_part(notes_rels_path.as_str()) {
                    let nrp_xml = String::from_utf8_lossy(&nrp.blob).into_owned();
                    if let Ok(nr) = Relationships::from_xml(&nrp_xml) {
                        for r in nr.iter() {
                            if matches!(r.reltype, RelType::Slide) {
                                parsed_notes_rels_target = Some(r.target.as_str().to_string());
                                break;
                            }
                        }
                    }
                }
                if let Some(np) = pkg.get_part(notes_partname.as_str()) {
                    let notes_xml = String::from_utf8_lossy(&np.blob).into_owned();
                    if let Ok(tb) = crate::oxml::parse_sld::parse_notes(&notes_xml) {
                        sld.notes = Some(tb);
                    }
                }
            }

            // ---------- 5.5) 处理 comments（如有） ----------
            // 收集本 slide 的评论 partname，用于写路径**复用**原始 OPC 关系。
            let mut parsed_comments_partname: Option<String> = None;
            let mut parsed_comments_lst: Option<crate::oxml::comments::CommentList> = None;
            if let Some(raw_target) = &comments_target_raw {
                // 从 slide rels 中取 target（相对路径如 "../comments/comment1.xml"），
                // 然后解析成绝对 partname。
                let comments_partname =
                    resolve_relative_partname(slide_partname.as_str(), raw_target.as_str());
                parsed_comments_partname = Some(comments_partname.clone());
                // 读 commentN.xml 本体，解析出 CommentList
                if let Some(cp) = pkg.get_part(comments_partname.as_str()) {
                    let comments_xml = String::from_utf8_lossy(&cp.blob).into_owned();
                    if let Ok(lst) = crate::oxml::parse_sld::parse_comments(&comments_xml) {
                        parsed_comments_lst = Some(lst);
                    }
                }
            }

            // ---------- 6) 构造 SlideEntry 并推入 ----------
            // 跟新 id_counter 状态：保证后续 shape id 单调递增。
            // 扫描 sld 内最大 sp id（仅 Sp / Pic 等有显式 id 的 shape）。
            for shape in &sld.shapes {
                let sp_id = match shape {
                    crate::oxml::SlideShape::Sp(sp) => sp.id,
                    crate::oxml::SlideShape::Pic(pic) => pic.id,
                    crate::oxml::SlideShape::CxnSp(c) => c.id,
                    crate::oxml::SlideShape::Group(g) => g.id,
                    crate::oxml::SlideShape::GraphicFrame(gf) => gf.id,
                };
                if sp_id > max_id_seen {
                    max_id_seen = sp_id;
                }
            }

            let mut slide = Slide::from_sld(sld, id_counter.clone(), layout_rid);
            if let Some(nrid) = notes_rid {
                slide.set_notes_rid(nrid);
            }
            // 把解析出的 notes 元数据回填到 Slide（写路径将**复用**这些值）：
            //   - partname：保证 read→save 不漂移；
            //   - 反向 rels target：保证 `notesSlideN.xml.rels → slideN.xml` 不断链。
            if let Some(np) = parsed_notes_partname {
                slide.set_notes_partname(np);
            }
            if let Some(nrt) = parsed_notes_rels_target {
                slide.set_notes_slide_rel_target(nrt);
            }
            // 把解析出的 comments 元数据回填到 Slide：
            //   - rid：保证 read→save 不漂移；
            //   - partname：保证 `commentN.xml` 路径稳定。
            if let Some(crid) = comments_rid {
                slide.set_comments_rid(crid);
            }
            if let Some(cp) = parsed_comments_partname {
                slide.set_comments_partname(cp);
            }
            if let Some(clst) = parsed_comments_lst {
                slide.set_comments(Some(clst));
            }

            // ---------- 5.5) 重建 media_entries（从 image_rels 还原） ----------
            // 目的：save 时 to_opc_package 会按 media_entries 写 Image 关系 + media part；
            //      读路径下必须把已读到的 Image 关系补成 MediaEntry，否则保存时会丢图。
            for (rid, target_partname) in &image_rels {
                if let Some(media_part) = pkg.get_part(target_partname.as_str()) {
                    let partname = crate::opc::part::new_part_name(target_partname.as_str());
                    slide.register_media(MediaEntry {
                        partname,
                        content_type: media_part.content_type.clone(),
                        blob: media_part.blob.clone(),
                        rid: rid.clone(),
                    });
                }
            }

            // ---------- 5.6) 处理 SmartArt（TODO-037 round-trip） ----------
            // 遍历 slide.inner.shapes 找 GraphicFrame.Graphic::SmartArt，
            // 根据其 4 个 rid 查 diagram_rel_map 找 target，读取 4 个 part 内容，
            // 构造 DiagramEntry 注入 slide.diagram_entries，保证 read→save 完整 round-trip。
            //
            // **配对策略**：SmartArtRef 在 parse_sld 阶段已提取 4 个 rid（dm/lo/qs/cs），
            // 这里直接根据 rid 查 diagram_rel_map 得到绝对 partname，再读 part 内容。
            // 缺失任一关系则跳过该 SmartArt（保留 slide xml 中的 raw_xml 引用，但 4 个 part 不写）。
            //
            // **两阶段策略**（避免借用冲突，与 5.7 chart 处理一致）：
            // 1. 阶段一（不可变借用）：遍历 shapes 收集 SmartArt 的 4 个 rid + 4 个 partname；
            // 2. 阶段二（可变借用）：逐个读取 part 内容 + 构造 DiagramEntry + 注册。
            // 元组字段顺序：(dm_rid, lo_rid, qs_rid, cs_rid, data_partname, layout_partname,
            //                quick_style_partname, colors_partname)
            let mut diagram_tasks: Vec<DiagramTask> = Vec::new();
            for shape in &slide.inner.shapes {
                if let crate::oxml::SlideShape::GraphicFrame(gf) = shape {
                    if let crate::oxml::shape::Graphic::SmartArt(smart_ref) = &gf.graphic {
                        let (Some(dm_rid), Some(lo_rid), Some(qs_rid), Some(cs_rid)) = (
                            &smart_ref.dm_rid,
                            &smart_ref.lo_rid,
                            &smart_ref.qs_rid,
                            &smart_ref.cs_rid,
                        ) else {
                            // SmartArtRef 的 4 个 rid 未完整提取，跳过（保留 raw_xml 但不构造 entry）。
                            continue;
                        };
                        let (
                            Some(data_partname),
                            Some(layout_partname),
                            Some(quick_style_partname),
                            Some(colors_partname),
                        ) = (
                            diagram_rel_map.get(dm_rid),
                            diagram_rel_map.get(lo_rid),
                            diagram_rel_map.get(qs_rid),
                            diagram_rel_map.get(cs_rid),
                        )
                        else {
                            // 4 个 rels 中有任一缺失（不完整的 SmartArt），跳过。
                            continue;
                        };
                        diagram_tasks.push((
                            dm_rid.clone(),
                            lo_rid.clone(),
                            qs_rid.clone(),
                            cs_rid.clone(),
                            data_partname.clone(),
                            layout_partname.clone(),
                            quick_style_partname.clone(),
                            colors_partname.clone(),
                        ));
                    }
                }
            }
            for (
                dm_rid,
                lo_rid,
                qs_rid,
                cs_rid,
                data_partname,
                layout_partname,
                quick_style_partname,
                colors_partname,
            ) in diagram_tasks
            {
                // 读取 4 个 part 的原始 XML 内容
                let data_xml = pkg
                    .get_part(data_partname.as_str())
                    .map(|p| String::from_utf8_lossy(&p.blob).into_owned())
                    .unwrap_or_default();
                let layout_xml = pkg
                    .get_part(layout_partname.as_str())
                    .map(|p| String::from_utf8_lossy(&p.blob).into_owned())
                    .unwrap_or_default();
                let quick_style_xml = pkg
                    .get_part(quick_style_partname.as_str())
                    .map(|p| String::from_utf8_lossy(&p.blob).into_owned())
                    .unwrap_or_default();
                let colors_xml = pkg
                    .get_part(colors_partname.as_str())
                    .map(|p| String::from_utf8_lossy(&p.blob).into_owned())
                    .unwrap_or_default();
                let entry = DiagramEntry {
                    data_partname: crate::opc::part::new_part_name(data_partname.as_str()),
                    layout_partname: crate::opc::part::new_part_name(layout_partname.as_str()),
                    quick_style_partname: crate::opc::part::new_part_name(
                        quick_style_partname.as_str(),
                    ),
                    colors_partname: crate::opc::part::new_part_name(colors_partname.as_str()),
                    data_xml,
                    layout_xml,
                    quick_style_xml,
                    colors_xml,
                    data_rid: dm_rid,
                    layout_rid: lo_rid,
                    quick_style_rid: qs_rid,
                    colors_rid: cs_rid,
                };
                slide.register_diagram(entry);
            }

            // ---------- 5.7) 处理 chart（TODO-004 读路径） ----------
            // 遍历 slide.inner.shapes 找 GraphicFrame.Graphic::Chart，
            // 根据其 rid 查 chart_rel_map 找 target，读取 chartN.xml 内容，
            // 调用 Chart::parse_from_xml 解析，用解析结果替换占位 Chart 模型。
            //
            // **配对策略**：parse_sld 阶段已在 graphicFrame 内提取 `<c:chart r:id="..."/>` 的 rid，
            // 这里直接根据 rid 查 chart_rel_map 得到绝对 partname，再读 part 内容解析。
            // 缺失关系或解析失败则保留占位 Chart（chart_type=Column, data 空），不阻塞 round-trip。
            //
            // **两阶段策略**（避免借用冲突）：
            // 1. 阶段一（不可变借用）：遍历 shapes 收集 (rid, chart_partname) 对；
            // 2. 阶段二（可变借用）：逐个读取 chartN.xml + 解析 + 替换 graphic + 注册 ChartEntry。
            let mut chart_tasks: Vec<(String, String)> = Vec::new(); // (rid, chart_partname)
            for shape in &slide.inner.shapes {
                if let crate::oxml::SlideShape::GraphicFrame(gf) = shape {
                    if let crate::oxml::shape::Graphic::Chart(chart) = &gf.graphic {
                        let rid = chart.rid.as_str();
                        if rid.is_empty() {
                            continue;
                        }
                        if let Some(chart_partname) = chart_rel_map.get(rid) {
                            chart_tasks.push((rid.to_string(), chart_partname.clone()));
                        }
                    }
                }
            }
            for (rid, chart_partname) in chart_tasks {
                let Some(chart_part) = pkg.get_part(chart_partname.as_str()) else {
                    continue;
                };
                let chart_xml = String::from_utf8_lossy(&chart_part.blob).into_owned();
                let Ok(mut parsed) = crate::oxml::chart::Chart::parse_from_xml(&chart_xml) else {
                    continue;
                };
                parsed.rid = rid.clone();
                // 把解析后的 Chart 同步回 slide 的 graphicFrame（in-place 替换匹配 rid 的 Chart）。
                for s in slide.inner.shapes.iter_mut() {
                    if let crate::oxml::SlideShape::GraphicFrame(gf2) = s {
                        if let crate::oxml::shape::Graphic::Chart(c2) = &mut gf2.graphic {
                            if c2.rid == rid {
                                *c2 = parsed.clone();
                                break;
                            }
                        }
                    }
                }
                // 注册 ChartEntry，供 to_opc_package 写出 chartN.xml part。
                // partname 用解析得到的绝对路径（保留原 partname 避免 rels 漂移）。
                slide.register_chart(ChartEntry {
                    partname: crate::opc::part::new_part_name(chart_partname.as_str()),
                    chart: parsed,
                    rid,
                    xlsx_blob: None,
                });
            }

            let entry = SlideEntry::new(slide, *sld_id, rid.clone(), slide_partname);
            pres.slides.push_entry(entry);
        }
        // 同步 id_counter 到"已见最大 id"——后续 add_* 不会冲撞已读 shape。
        pres.id_counter.set(max_id_seen);

        // ---------- 读 docProps/custom.xml（自定义属性，TODO-034） ----------
        if let Some(custom_part) = pkg.get_part("/docProps/custom.xml") {
            let custom_xml = String::from_utf8_lossy(&custom_part.blob).into_owned();
            pres.custom_properties = CustomProperties::from_xml(&custom_xml);
        }

        // ---------- 读 ppt/commentAuthors.xml（评论作者，TODO-036） ----------
        if let Some(authors_part) = pkg.get_part("/ppt/commentAuthors.xml") {
            let authors_xml = String::from_utf8_lossy(&authors_part.blob).into_owned();
            if let Ok(lst) = crate::oxml::parse_sld::parse_comment_authors(&authors_xml) {
                pres.comment_authors = lst;
            }
        }

        Ok(pres)
    }

    /// 不可变幻灯片集合。
    pub fn slides(&self) -> &Slides {
        &self.slides
    }
    /// 可变幻灯片集合。
    pub fn slides_mut(&mut self) -> &mut Slides {
        &mut self.slides
    }

    /// 取出共享的 `id_counter` 克隆。
    ///
    /// `Slides::add_slide` 接受这个计数器，从而保证跨 slide 的 shape id 不冲突。
    pub fn id_counter(&self) -> std::rc::Rc<std::cell::Cell<u32>> {
        self.id_counter.clone()
    }

    /// 不可变版式集合。
    pub fn slide_layouts(&self) -> &SlideLayouts {
        &self.slide_layouts
    }
    /// 可变版式集合。
    pub fn slide_layouts_mut(&mut self) -> &mut SlideLayouts {
        &mut self.slide_layouts
    }

    /// 按 slide 索引取该 slide 所引用的版式（TODO-007）。
    ///
    /// 对标 python-pptx `slide.slide_layout`。本方法通过匹配
    /// `slide.layout_rid()` 与 `SlideLayoutRef.rid()` 查找。
    ///
    /// # 参数
    /// - `slide_idx`：slide 在 `slides` 集合中的索引。
    ///
    /// # 返回
    /// - `Some(SlideLayoutRef)`：找到匹配的版式（克隆，因 `SlideLayoutRef` 是 `Clone`）；
    /// - `None`：索引越界，或未找到 rid 匹配的版式。
    ///
    /// # 示例
    /// ```no_run
    /// use pptx_rs::Presentation;
    /// let mut p = Presentation::new().unwrap();
    /// let counter = p.id_counter();
    /// let _ = p.slides_mut().add_slide(counter).unwrap();
    /// // 获取第 0 张 slide 所引用的版式
    /// if let Some(layout) = p.layout_for_slide(0) {
    ///     println!("版式名: {}", layout.name());
    /// }
    /// ```
    pub fn layout_for_slide(&self, slide_idx: usize) -> Option<SlideLayoutRef> {
        let entry = self.slides.get(slide_idx)?;
        let layout_rid = entry.sld.layout_rid();
        self.slide_layouts
            .items
            .iter()
            .find(|l| l.rid == layout_rid)
            .cloned()
    }

    /// 获取使用指定版式的所有 slide 索引（TODO-008）。
    ///
    /// 对标 python-pptx `SlideLayout.used_by_slides`。
    ///
    /// # 参数
    /// - `layout_rid`：版式的关系 id（`SlideLayoutRef.rid()`）。
    ///
    /// # 返回
    /// 所有 `layout_rid()` 等于 `layout_rid` 的 slide 索引列表（升序）。
    pub fn slides_using_layout(&self, layout_rid: &str) -> Vec<usize> {
        self.slides
            .iter()
            .enumerate()
            .filter(|(_, entry)| entry.sld.layout_rid() == layout_rid)
            .map(|(i, _)| i)
            .collect()
    }

    /// 不可变母版集合。
    pub fn slide_masters(&self) -> &SlideMasters {
        &self.slide_masters
    }
    /// 可变母版集合。
    pub fn slide_masters_mut(&mut self) -> &mut SlideMasters {
        &mut self.slide_masters
    }

    /// 当前画布宽度（EMU）。
    pub fn slide_width(&self) -> Emu {
        self.width
    }
    /// 当前画布高度（EMU）。
    pub fn slide_height(&self) -> Emu {
        self.height
    }

    /// 显式设置画布尺寸。
    ///
    /// 注意此方法**不**触发任何 `Slide` 内部 shape 的重新布局；shape 仍按各自的
    /// EMU 坐标保存，调用方需自行决定是否缩放。
    pub fn set_slide_size(&mut self, width: Emu, height: Emu) {
        self.width = width;
        self.height = height;
    }

    /// 取文档核心属性（不可变引用）。
    ///
    /// 对标 pypdf `PdfReader.metadata` / python-pptx `Presentation.core_properties`。
    pub fn core_properties(&self) -> &CoreProperties {
        &self.core_properties
    }

    /// 取文档核心属性（可变引用）。
    ///
    /// 对标 pypdf `PdfWriter.add_metadata(infos)`。
    /// 修改后会在 `save` / `to_bytes` 时序列化到 `docProps/core.xml` 和 `docProps/app.xml`。
    pub fn core_properties_mut(&mut self) -> &mut CoreProperties {
        &mut self.core_properties
    }

    /// 取自定义文档属性（不可变引用）。
    ///
    /// 对标 python-pptx `Presentation.custom_properties`（v1.0+）。
    pub fn custom_properties(&self) -> &CustomProperties {
        &self.custom_properties
    }

    /// 取自定义文档属性（可变引用）。
    ///
    /// 修改后会在 `save` / `to_bytes` 时序列化到 `docProps/custom.xml`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx_rs::Presentation;
    /// use pptx_rs::presentation::CustomPropertyValue;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// p.custom_properties_mut().set("Project", CustomPropertyValue::Text("Demo".to_string()));
    /// p.custom_properties_mut().set("Version", CustomPropertyValue::Int(42));
    /// ```
    pub fn custom_properties_mut(&mut self) -> &mut CustomProperties {
        &mut self.custom_properties
    }

    /// 取评论作者列表（不可变引用）。
    ///
    /// 评论作者在 `save` / `to_bytes` 时序列化到 `/ppt/commentAuthors.xml`。
    pub fn comment_authors(&self) -> &crate::oxml::comments::CommentAuthorList {
        &self.comment_authors
    }

    /// 取评论作者列表（可变引用）。
    ///
    /// 修改后会在 `save` / `to_bytes` 时序列化到 `/ppt/commentAuthors.xml`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx_rs::Presentation;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// let author_id = p.comment_authors_mut().get_or_insert_id("张三", "ZS");
    /// ```
    pub fn comment_authors_mut(&mut self) -> &mut crate::oxml::comments::CommentAuthorList {
        &mut self.comment_authors
    }

    /// 取章节分组列表（不可变引用，TODO-039）。
    ///
    /// 章节分组对应 PowerPoint 大纲视图中的"节"功能，在 `presentation.xml`
    /// 的 `<p:extLst>` 内以 `<p14:sectionLst>` 扩展元素持久化。
    ///
    /// # 与 python-pptx 的对应
    ///
    /// python-pptx 截至 v1.0 仍未提供 section API，本方法是 pptx-rs 的扩展。
    pub fn sections(&self) -> &crate::oxml::section::SectionList {
        &self.sections
    }

    /// 取章节分组列表（可变引用，TODO-039）。
    ///
    /// 修改后会在 `save` / `to_bytes` 时序列化到 `presentation.xml` 的
    /// `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E64}">`
    /// 内的 `<p14:sectionLst>`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx_rs::Presentation;
    /// use pptx_rs::oxml::section::Section;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// // 假设已添加 2 张 slide，它们的 sld_id 分别为 256 / 257
    /// let mut s1 = Section::new("引言");
    /// s1.push(256);
    /// s1.push(257);
    /// p.sections_mut().push(s1);
    /// ```
    pub fn sections_mut(&mut self) -> &mut crate::oxml::section::SectionList {
        &mut self.sections
    }

    /// 取备注母版集合（不可变引用，TODO-045）。
    ///
    /// 备注母版是所有备注页的"模板"——定义了备注页的默认占位符与文本样式。
    /// 一个演示文稿通常只有 0 或 1 个备注母版。
    ///
    /// # 与 python-pptx 的对应
    ///
    /// - `pptx.Presentation.notes_master` ←→ `presentation.notes_masters().first()`；
    /// - python-pptx 仅暴露单个 `notes_master` 属性，本库用集合表达以兼容多母版场景。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx_rs::Presentation;
    ///
    /// let p = Presentation::new().unwrap();
    /// // 空白文档无备注母版
    /// assert!(p.notes_masters().is_empty());
    /// ```
    pub fn notes_masters(&self) -> &NotesMasters {
        &self.notes_masters
    }

    /// 取备注母版集合（可变引用，TODO-045）。
    ///
    /// 主要用于在内存中修改已解析出的备注母版形状（写路径暂未实现持久化）。
    pub fn notes_masters_mut(&mut self) -> &mut NotesMasters {
        &mut self.notes_masters
    }

    /// 便捷方法：取第一个备注母版（python-pptx `presentation.notes_master` 风格）。
    ///
    /// `None` 表示该演示文稿无备注母版。
    pub fn notes_master(&self) -> Option<&NotesMasterRef> {
        self.notes_masters.first()
    }

    /// 一次性设置多个核心属性（便捷方法）。
    ///
    /// 对标 pypdf `PdfWriter.add_metadata(infos)` 的"批量设置"语义。
    /// 传入 `None` 的字段**不会**覆盖已有值。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx_rs::Presentation;
    /// let mut p = Presentation::new().unwrap();
    /// p.set_metadata(
    ///     Some("My Presentation"),  // title
    ///     Some("Author Name"),      // creator
    ///     None,                     // subject (keep existing)
    ///     None,                     // keywords
    /// );
    /// ```
    pub fn set_metadata(
        &mut self,
        title: Option<&str>,
        creator: Option<&str>,
        subject: Option<&str>,
        keywords: Option<&str>,
    ) {
        if let Some(t) = title {
            self.core_properties.title = Some(t.to_string());
        }
        if let Some(c) = creator {
            self.core_properties.creator = Some(c.to_string());
        }
        if let Some(s) = subject {
            self.core_properties.subject = Some(s.to_string());
        }
        if let Some(k) = keywords {
            self.core_properties.keywords = Some(k.to_string());
        }
    }

    /// 幻灯片总数。
    ///
    /// 对标 pypdf `PdfReader.num_pages` / python-pptx `len(prs.slides)`。
    pub fn num_slides(&self) -> usize {
        self.slides.len()
    }

    /// 给所有 slide 添加**文本水印**。
    ///
    /// 对标 pypdf `PageObject.merge_page(watermark_page)` 的水印注入模式。
    ///
    /// # 实现原理
    /// 在每张 slide 上添加一个**旋转 + 半透明**的文本框（`p:sp` + `cNvSpPr txBox="1"`），
    /// 位于 slide 中心，字体大小可配置。文本框的 `z-order` 被推到最顶层
    /// （即最后添加的形状），确保水印覆盖在内容之上。
    ///
    /// # 参数
    /// - `text`：水印文本（如 "CONFIDENTIAL" / "DRAFT"）；
    /// - `font_size_pt`：字体大小（磅），默认 36；
    /// - `color`：水印颜色，默认灰色 `RGBColor(0xC0, 0xC0, 0xC0)`；
    /// - `rotation_deg`：旋转角度（度），默认 -30（逆时针 30°）；
    /// - `alpha`：不透明度（0-100000），默认 30_000（30% 不透明 / 70% 透明）；
    /// - `font_name`：字体名称，默认 "Calibri"。
    ///
    /// # 与 pypdf 的差异
    /// - pypdf 用"页面合并"（`merge_page`）实现水印——把水印 PDF 页叠加到目标页；
    /// - 本方法用"形状注入"——直接在 slide XML 中插入一个旋转文本框。
    ///   两种方式在视觉上等价，但 OOXML 不支持"页面合并"语义。
    #[allow(clippy::field_reassign_with_default)]
    pub fn add_watermark(
        &mut self,
        text: &str,
        font_size_pt: Option<f64>,
        color: Option<crate::units::RGBColor>,
        rotation_deg: Option<i32>,
        alpha: Option<i32>,
        font_name: Option<&str>,
    ) -> crate::Result<()> {
        let fs = font_size_pt.unwrap_or(36.0);
        let clr = color.unwrap_or(crate::units::RGBColor(0xC0, 0xC0, 0xC0));
        let rot = rotation_deg.unwrap_or(-30);
        let alpha_val = alpha.unwrap_or(30_000);
        let font = font_name.unwrap_or("Calibri");
        let rot_emu = rot * 60000; // 1° = 60000 EMU 角度单位

        for i in 0..self.slides.len() {
            // 索引 i 来自 0..len，不会越界；使用 ok_or 传播错误以遵守 §5 规则
            let slide = self
                .slides
                .get_mut(i)
                .ok_or(crate::Error::IndexOutOfRange(i))?;
            let sld = &mut slide.sld;
            let id = sld.next_shape_id();

            let mut sp = crate::oxml::shape::Sp::default();
            sp.id = id;
            sp.name = format!("Watermark {}", i + 1);
            sp.c_nv_sp_pr_tx_box = true;

            // 位置：居中（大约在 slide 中心）
            let cx = self.width.0;
            let cy = self.height.0;
            sp.properties.xfrm.off_x = Some(Emu(cx / 4));
            sp.properties.xfrm.off_y = Some(Emu(cy / 4));
            sp.properties.xfrm.ext_cx = Some(Emu(cx / 2));
            sp.properties.xfrm.ext_cy = Some(Emu(cy / 2));
            sp.properties.xfrm.rot = Some(rot_emu);
            sp.properties.geometry = Some(crate::oxml::sppr::Geometry::preset(
                crate::oxml::simpletypes::PresetGeometry::Rectangle,
            ));
            // 水印形状无填充（透明背景），无边框
            sp.properties.fill = crate::oxml::sppr::Fill::None;
            sp.properties.line = None;

            // 文本
            let mut tb = crate::oxml::txbody::TextBody::new();
            let mut para = crate::oxml::txbody::Paragraph::default();
            let mut run = crate::oxml::txbody::Run::default();
            run.text = text.to_string();
            run.properties.size = Some(crate::units::Pt(fs));
            run.properties.color = crate::oxml::color::Color::RGB(clr);
            run.properties.bold = true;
            run.properties.alpha = Some(alpha_val);
            // 设置字体
            run.properties.latin_font = Some(font.to_string());
            run.properties.eastasia_font = Some("宋体".to_string());
            para.runs.push(run);
            // 居中对齐
            para.properties.alignment = Some(crate::oxml::simpletypes::Alignment::Center);
            tb.paragraphs.push(para);
            sp.text = tb;

            sld.inner.shapes.push(crate::oxml::SlideShape::Sp(sp));
        }
        Ok(())
    }

    /// 给所有 slide 添加**图片水印**。
    ///
    /// 对标 pypdf `PageObject.merge_page(watermark_page)` 的水印注入模式。
    ///
    /// # 实现原理
    /// 在每张 slide 上添加一个**半透明**的图片形状（`p:pic` + `a:alphaModFix`），
    /// 位于 slide 中心，大小可配置。图片的 `z-order` 被推到最顶层
    /// （即最后添加的形状），确保水印覆盖在内容之上。
    ///
    /// # 参数
    /// - `image_bytes`：图片二进制数据（PNG / JPG / BMP 等）；
    /// - `ext`：图片扩展名（如 `"png"` / `"jpg"`，不含前导 `.`）；
    /// - `alpha`：不透明度（0-100000），默认 30_000（30% 不透明 / 70% 透明）；
    /// - `left`：水印左上角 x 坐标（EMU），默认居中偏左 1/4 画布宽度；
    /// - `top`：水印左上角 y 坐标（EMU），默认居中偏上 1/4 画布高度；
    /// - `width`：水印宽度（EMU），默认画布宽度的一半；
    /// - `height`：水印高度（EMU），默认画布高度的一半。
    ///
    /// # 与文本水印的差异
    /// - 文本水印（`add_watermark`）用旋转文本框实现，适合"CONFIDENTIAL"等文字；
    /// - 图片水印用半透明图片实现，适合公司 logo / 自定义图案等场景。
    ///
    /// # 错误
    /// - [`crate::Error::Encryption`]：图片字节为空。
    #[allow(clippy::field_reassign_with_default)]
    #[allow(clippy::too_many_arguments)]
    pub fn add_image_watermark(
        &mut self,
        image_bytes: &[u8],
        ext: &str,
        alpha: Option<i32>,
        left: Option<Emu>,
        top: Option<Emu>,
        width: Option<Emu>,
        height: Option<Emu>,
    ) -> crate::Result<()> {
        if image_bytes.is_empty() {
            return Err(crate::Error::encryption("image bytes must not be empty"));
        }
        let alpha_val = alpha.unwrap_or(30_000);
        let cx = self.width.0;
        let cy = self.height.0;
        let wm_left = left.unwrap_or(Emu(cx / 4));
        let wm_top = top.unwrap_or(Emu(cy / 4));
        let wm_width = width.unwrap_or(Emu(cx / 2));
        let wm_height = height.unwrap_or(Emu(cy / 2));

        // 所有 slide 共享同一个 media partname（同一张图片只存一份到 zip）
        let ext_norm = if ext.starts_with('.') {
            ext.to_string()
        } else {
            format!(".{}", ext)
        };
        // 用第一张 slide 的 media_index 分配全局唯一的 partname
        let global_media_idx = self
            .slides
            .get(0)
            .map(|s| s.sld.next_media_index())
            .unwrap_or(1);
        let shared_partname = crate::opc::part::new_part_name(
            format!("/ppt/media/image{}{}", global_media_idx, ext_norm).as_str(),
        );
        let ct = crate::shape::picture::content_type_for(&ext_norm);

        for i in 0..self.slides.len() {
            // 索引 i 来自 0..len，不会越界；使用 ok_or 传播错误以遵守 §5 规则
            let slide = self
                .slides
                .get_mut(i)
                .ok_or(crate::Error::IndexOutOfRange(i))?;
            let id = slide.sld.next_shape_id();
            let rid = slide.sld.allocate_image_rid();

            // 构造 oxml Pic
            let mut pic = crate::oxml::shape::Pic::default();
            pic.id = id;
            pic.name = format!("WatermarkImage {}", i + 1);
            pic.rid = rid.clone();
            pic.alpha = Some(alpha_val);
            pic.fill_mode = crate::oxml::sppr::BlipFillMode::Stretch;
            pic.properties.xfrm.off_x = Some(wm_left);
            pic.properties.xfrm.off_y = Some(wm_top);
            pic.properties.xfrm.ext_cx = Some(wm_width);
            pic.properties.xfrm.ext_cy = Some(wm_height);
            // 图片形状不需要填充和边框
            pic.properties.fill = crate::oxml::sppr::Fill::None;
            pic.properties.line = None;

            // 注册 media 到 slide（所有 slide 共享同一个 partname，to_opc_package 会去重）
            slide.sld.register_media(MediaEntry {
                partname: shared_partname.clone(),
                content_type: ct.to_string(),
                blob: image_bytes.to_vec(),
                rid,
            });

            slide
                .sld
                .inner
                .shapes
                .push(crate::oxml::SlideShape::Pic(pic));
        }
        Ok(())
    }

    /// 设置修改密码保护（打开时可只读浏览，修改需密码）。
    ///
    /// 对标 PowerPoint "保护演示文稿 → 限制访问" 功能。
    ///
    /// # 算法
    /// 使用 SHA-512 + 随机 salt + 100 000 次迭代，符合 MS-OFFCRYPTO §2.4.2.4。
    /// 保护信息注入到 `presentation.xml` 的 `<p:modifyVerifier>` 元素中。
    ///
    /// # 参数
    /// - `password`：修改密码（不能为空）。
    ///
    /// # 错误
    /// - [`crate::Error::Encryption`]：密码为空。
    pub fn set_write_protection(&mut self, password: &str) -> crate::Result<()> {
        if password.is_empty() {
            return Err(crate::Error::encryption("password must not be empty"));
        }
        let salt = crate::crypto::generate_random_bytes(crate::crypto::SALT_LEN);
        self.modify_protection = Some(crate::crypto::ModifyProtection::from_password(
            password,
            &salt,
            crate::crypto::MODIFY_SPIN_COUNT,
        ));
        Ok(())
    }

    /// 验证修改密码是否匹配。
    ///
    /// # 参数
    /// - `password`：待验证的密码。
    ///
    /// # 返回
    /// - `Ok(true)`：密码匹配；
    /// - `Ok(false)`：密码不匹配；
    /// - 若未设置修改保护，返回 `Ok(false)`。
    pub fn verify_write_protection(&self, password: &str) -> crate::Result<bool> {
        match &self.modify_protection {
            Some(mp) => Ok(mp.verify_password(password)),
            None => Ok(false),
        }
    }

    /// 移除修改密码保护。
    pub fn remove_write_protection(&mut self) {
        self.modify_protection = None;
    }

    /// 检查是否设置了修改密码保护。
    pub fn is_write_protected(&self) -> bool {
        self.modify_protection.is_some()
    }

    /// 保存为加密的 `.pptx` 文件（打开文件需密码）。
    ///
    /// 对标 PowerPoint "保护演示文稿 → 用密码进行加密" 功能。
    /// 使用 ECMA-376 Agile Encryption（AES-256-CBC + SHA-512）。
    ///
    /// # 参数
    /// - `path`：输出文件路径；
    /// - `password`：加密密码。
    ///
    /// # 错误
    /// - [`crate::Error::Encryption`]：加密过程失败。
    /// - [`crate::Error::Io`]：文件写入失败。
    pub fn save_encrypted(&self, path: impl AsRef<Path>, password: &str) -> crate::Result<()> {
        let pkg = self.to_opc_package()?;
        let zip_bytes = pkg.to_bytes()?;
        let encrypted = crate::crypto::encrypt_package(&zip_bytes, password)?;
        std::fs::write(path.as_ref(), &encrypted)?;
        Ok(())
    }

    /// 序列化为加密的字节流（打开需密码）。
    ///
    /// 等价于"保存到内存"的加密版本，适用于网络传输等场景。
    pub fn to_encrypted_bytes(&self, password: &str) -> crate::Result<Vec<u8>> {
        let pkg = self.to_opc_package()?;
        let zip_bytes = pkg.to_bytes()?;
        crate::crypto::encrypt_package(&zip_bytes, password)
    }

    /// 打开加密的 `.pptx` 文件。
    ///
    /// # 参数
    /// - `path`：文件路径；
    /// - `password`：解密密码。
    ///
    /// # 错误
    /// - [`crate::Error::Encryption`]：密码错误或加密格式损坏；
    /// - [`crate::Error::Io`]：文件读取失败。
    pub fn open_encrypted(path: impl AsRef<Path>, password: &str) -> crate::Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::load_encrypted_bytes(&bytes, password)
    }

    /// 从加密的字节流加载。
    pub fn load_encrypted_bytes(bytes: &[u8], password: &str) -> crate::Result<Self> {
        let decrypted = crate::crypto::decrypt_package(bytes, password)?;
        Self::load_bytes(&decrypted)
    }

    /// 检查文件是否为加密的 OOXML 文档。
    ///
    /// 加密文档的特征：ZIP 中包含 `EncryptionInfo` 条目。
    pub fn is_encrypted_file(path: impl AsRef<Path>) -> crate::Result<bool> {
        let bytes = std::fs::read(path.as_ref())?;
        Ok(crate::crypto::is_encrypted_package(&bytes))
    }

    /// 检查字节流是否为加密的 OOXML 文档。
    pub fn is_encrypted_bytes(bytes: &[u8]) -> bool {
        crate::crypto::is_encrypted_package(bytes)
    }

    /// 保存到本地 `.pptx` 文件。
    ///
    /// # 错误
    /// 透传 `Presentation::to_opc_package` 与 [`OpcPackage::save`] 的一切错误。
    pub fn save(&self, path: impl AsRef<Path>) -> crate::Result<()> {
        let pkg = self.to_opc_package()?;
        pkg.save(path)
    }

    /// 把整份演示文稿序列化为 zip 字节流。
    ///
    /// 等价于"保存到内存"——适用于网络响应、自动化测试 fixture、邮件附件等场景。
    pub fn to_bytes(&self) -> crate::Result<Vec<u8>> {
        let pkg = self.to_opc_package()?;
        pkg.to_bytes()
    }

    /// 内部入口：把内存模型组装为完整的 [`OpcPackage`]。
    ///
    /// 该方法是 `save` / `to_bytes` 的公共实现，负责按以下顺序构建 part 树：
    ///
    /// 1. `_rels/.rels`（根关系）—— 指向 `ppt/presentation.xml` + `docProps/*`；
    /// 2. `docProps/core.xml` + `docProps/app.xml`（最小化的 Office 文档属性）；
    /// 3. `ppt/theme/theme1.xml`（标准 Office 主题）；
    /// 4. `ppt/slideMasters/slideMaster1.xml` + 关系（指向 theme + slideLayouts）；
    /// 5. `ppt/slideLayouts/slideLayout1.xml` + 关系（指向 slideMaster1）；
    /// 6. `ppt/_rels/presentation.xml.rels` + 全部 slide / media 关系；
    /// 7. `ppt/presentation.xml`（包含 `sldSz` 与 `sldIdLst`）；
    /// 8. 每个 slide 的 `ppt/slides/slideN.xml` 与其 `_rels/slideN.xml.rels`；
    /// 9. 媒体（图片）`ppt/media/*`；
    /// 10. `ppt/presProps.xml` / `ppt/viewProps.xml` / `ppt/tableStyles.xml`（必备辅件）。
    fn to_opc_package(&self) -> crate::Result<OpcPackage> {
        let mut pkg = OpcPackage::new();

        // ---------------- 1) 根 _rels/.rels ----------------
        // 关系中至少需要：指向 presentation.xml 的 OfficeDocument + 指向 docProps 的元数据。
        let mut root_rels = Relationships::new();
        root_rels.add(Relationship::internal(
            "rId1",
            RelType::OfficeDocument,
            new_part_name("/ppt/presentation.xml"),
        ))?;
        root_rels.add(Relationship::internal(
            "rId2",
            crate::opc::rels::RelType::Other(
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties".to_string(),
            ),
            new_part_name("/docProps/core.xml"),
        ))?;
        root_rels.add(Relationship::internal(
            "rId3",
            crate::opc::rels::RelType::Other(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties".to_string(),
            ),
            new_part_name("/docProps/app.xml"),
        ))?;
        // 自定义属性关系（仅当 custom_properties 非空时添加）
        if !self.custom_properties.is_empty() {
            root_rels.add(Relationship::internal(
                "rId4",
                crate::opc::rels::RelType::Other(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties".to_string(),
                ),
                new_part_name("/docProps/custom.xml"),
            ))?;
        }
        let root_rels_xml = root_rels.to_xml();
        let root_rels_part = Part::new(
            new_part_name("/_rels/.rels"),
            ct::RELATIONSHIPS,
            root_rels_xml.into_bytes(),
        );
        pkg.put_part(root_rels_part);

        // ---------------- 2) docProps/core.xml ----------------
        // 极简实现：title/creator/lastModifiedBy 三个字段，PowerPoint 会接受。
        let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>pptx-rs</dc:title>
  <dc:creator>pptx-rs</dc:creator>
  <cp:lastModifiedBy>pptx-rs</cp:lastModifiedBy>
</cp:coreProperties>"#;
        pkg.put_part(Part::new(
            new_part_name("/docProps/core.xml"),
            ct::CORE_PROPS,
            core_xml.as_bytes().to_vec(),
        ));

        // ---------------- 2.1) docProps/app.xml ----------------
        // 标注 Application 与 AppVersion，PowerPoint 用它显示"由 X 产生"。
        let app_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>pptx-rs</Application>
  <AppVersion>1.0</AppVersion>
</Properties>"#;
        pkg.put_part(Part::new(
            new_part_name("/docProps/app.xml"),
            ct::APP_PROPS,
            app_xml.as_bytes().to_vec(),
        ));

        // ---------------- 2.2) docProps/custom.xml（仅当 custom_properties 非空） -----
        // 自定义属性：用户自定义的键值对，PowerPoint 在"文件 → 信息 → 属性"中显示。
        if !self.custom_properties.is_empty() {
            let custom_xml = self.custom_properties.to_xml();
            if !custom_xml.is_empty() {
                pkg.put_part(Part::new(
                    new_part_name("/docProps/custom.xml"),
                    ct::CUSTOM_PROPS,
                    custom_xml.into_bytes(),
                ));
            }
        }

        // ---------------- 3) theme1.xml ----------------
        // TODO-001：使用解析出的 theme（若有），否则用默认 Office 主题。
        // 这样 read→save 能保留原始主题的颜色/字体/格式方案。
        let theme_xml = if self.theme.name.is_empty()
            && self.theme.color_scheme.dk1.is_none()
            && self.theme.font_scheme.major_latin.is_empty()
        {
            // 未解析到 theme（新建的 Presentation），用默认 Office 主题
            default_theme_xml()
        } else {
            self.theme.to_xml()
        };
        pkg.put_part(Part::new(
            new_part_name("/ppt/theme/theme1.xml"),
            ct::THEME,
            theme_xml.into_bytes(),
        ));

        // ---------------- 4) slideMaster1 + 关系 ----------------
        // TODO-001：使用解析出的 master XML（若有），否则用默认。
        // 母版先列 slideLayout，再列 theme —— 与 Office 顺序保持一致以避免兼容性问题。
        let master_xml = if let Some(master_ref) = self.slide_masters.items.first() {
            master_ref.oxml.borrow().to_xml()
        } else {
            OxmlSldMaster::default().to_xml()
        };
        let master_partname = new_part_name("/ppt/slideMasters/slideMaster1.xml");
        let mut master_rels = Relationships::new();
        master_rels.add(Relationship::internal_str(
            "rId1",
            RelType::SlideLayout,
            "../slideLayouts/slideLayout1.xml",
        ))?;
        master_rels.add(Relationship::internal_str(
            "rId2",
            RelType::Theme,
            "../theme/theme1.xml",
        ))?;
        let master_rels_xml = master_rels.to_xml();
        let master_rels_partname = rels_partname_for(master_partname.as_str());
        pkg.put_part(Part::new(
            PartName::from_unchecked(master_rels_partname),
            ct::RELATIONSHIPS,
            master_rels_xml.into_bytes(),
        ));
        pkg.put_part(Part::new(
            master_partname,
            ct::SLIDE_MASTER,
            master_xml.into_bytes(),
        ));

        // ---------------- 5) slideLayout1（空白） + 关系 ----------------
        // TODO-001：使用解析出的 layout XML（若有），否则用默认。
        // 版式关系只指向所属母版。
        let layout_xml = if let Some(layout_ref) = self.slide_layouts.items.first() {
            layout_ref.oxml.borrow().to_xml()
        } else {
            OxmlSldLayout::default().to_xml()
        };
        let layout_partname = new_part_name("/ppt/slideLayouts/slideLayout1.xml");
        let mut layout_rels = Relationships::new();
        layout_rels.add(Relationship::internal_str(
            "rId1",
            RelType::SlideMaster,
            "../slideMasters/slideMaster1.xml",
        ))?;
        let layout_rels_partname = rels_partname_for(layout_partname.as_str());
        pkg.put_part(Part::new(
            PartName::from_unchecked(layout_rels_partname),
            ct::RELATIONSHIPS,
            layout_rels.to_xml().into_bytes(),
        ));
        pkg.put_part(Part::new(
            layout_partname,
            ct::SLIDE_LAYOUT,
            layout_xml.into_bytes(),
        ));

        // ---------------- 6) presentation.xml.rels ----------------
        // 关系至少包括：master、layout、theme、presProps、viewProps、tableStyles。
        // 注意：rid 使用 "rIdP1"-"rIdP6" 前缀，避免与 slide 的 "rIdN" 冲突。
        // （真实 .pptx 中 slide rid 可能从 rId3 开始，与固定 rId3 冲突。）
        let mut pres_rels = Relationships::new();
        pres_rels.add(Relationship::internal_str(
            "rIdP1",
            RelType::SlideMaster,
            "slideMasters/slideMaster1.xml",
        ))?;
        pres_rels.add(Relationship::internal_str(
            "rIdP2",
            RelType::SlideLayout,
            "slideLayouts/slideLayout1.xml",
        ))?;
        pres_rels.add(Relationship::internal_str(
            "rIdP3",
            RelType::Theme,
            "theme/theme1.xml",
        ))?;
        pres_rels.add(Relationship::internal_str(
            "rIdP4",
            crate::opc::rels::RelType::Other(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
                    .to_string(),
            ),
            "presProps.xml",
        ))?;
        pres_rels.add(Relationship::internal_str(
            "rIdP5",
            crate::opc::rels::RelType::Other(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
                    .to_string(),
            ),
            "viewProps.xml",
        ))?;
        pres_rels.add(Relationship::internal_str(
            "rIdP6",
            crate::opc::rels::RelType::Other(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
                    .to_string(),
            ),
            "tableStyles.xml",
        ))?;

        // presentation.xml 根结构：sldSz + sldIdLst + sldMasterIdLst。
        // TODO-001：使用解析出的 sld_master_ids（若有），支持多母版 read→save 保真。
        let sld_master_ids: Vec<crate::oxml::presentation::SldMasterIdEntry> =
            if !self.sld_master_ids.is_empty() {
                self.sld_master_ids
                    .iter()
                    .map(|(id, rid)| crate::oxml::presentation::SldMasterIdEntry {
                        id: *id,
                        rid: rid.clone(),
                    })
                    .collect()
            } else {
                // 默认：空列表，to_xml 会写出默认的单个母版
                Vec::new()
            };
        let mut pres_root = PresentationRoot {
            slide_width: Some(self.width),
            slide_height: Some(self.height),
            slide_ids: Vec::new(),
            sld_master_ids,
            // TODO-039：把 Presentation.sections 透传给 PresentationRoot，
            // 由 PresentationRoot::to_xml 在 <p:extLst> 内输出 <p14:sectionLst>。
            sections: self.sections.clone(),
            ..Default::default()
        };

        // ---------------- 7) 遍历 slide，写 XML + 关系 ----------------
        let mut notes_index: u32 = 0;
        let mut comments_index: u32 = 0;
        // chart 全局索引：保证多 slide 之间的 chartN.xml partname 唯一。
        let mut chart_global_index: u32 = 0;
        // chart 嵌入式 Excel 全局索引：保证多 chart 之间的 Microsoft_Excel_WorksheetN.xlsx partname 唯一（TODO-004 Excel 嵌入）。
        let mut chart_xlsx_global_index: u32 = 0;
        // ole 全局索引：保证多 slide 之间的 oleObjectN.bin partname 唯一（TODO-043）。
        let mut ole_global_index: u32 = 0;
        // video/audio 全局索引：保证多 slide 之间的 mediaN.mp4 / mediaN.mp3 partname 唯一（TODO-033）。
        // 注意：imageN.png 走 media_entries 的去重逻辑，不在此处计数；
        // 视频/音频每次嵌入都生成独立 part（即使多个 slide 引用同一文件也分别写出）。
        let mut video_global_index: u32 = 0;
        let mut audio_global_index: u32 = 0;
        // SmartArt 全局索引：保证多 slide 之间的 dataN.xml / layoutN.xml 等 partname 唯一（TODO-037）。
        let mut diagram_global_index: u32 = 0;
        for (i, entry) in self.slides.iter().enumerate() {
            let partname = PartName::from_unchecked(entry.partname.clone());
            let rid = entry.rid.clone();
            // 给每个 slide 注入 layout_rid（指向 slideLayout1）。
            let mut sld = entry.sld.clone();
            sld.set_layout_rid("rId1".to_string());

            // 该 slide 的 .rels：始终含 layout 关系；含图片再加 Image 关系；含 notes 再加 Notes 关系。
            let mut extra_rels = Relationships::new();
            extra_rels.add(Relationship::internal_str(
                "rId1",
                RelType::SlideLayout,
                "../slideLayouts/slideLayout1.xml",
            ))?;
            // 把"该 slide 用到的 media"挂上来。rId 形如 `rIdImg1` 以便与 layout 区分。
            for media in &entry.sld.media_entries {
                let media_target = format!(
                    "../media/{}",
                    media.partname.as_str().trim_start_matches("/ppt/media/")
                );
                extra_rels.add(Relationship::internal_str(
                    media.rid.clone(),
                    RelType::Image,
                    media_target,
                ))?;
            }
            // 如果该 slide 有 notes，挂上 Notes 关系（slide → notesSlide）。
            //
            // **关键**：写路径**优先复用**读路径保存的原始 OPC 元数据：
            //   - `notes_partname`：从 `Slide.notes_partname` 取（None 时按 `notesSlide{idx}.xml` 分配）；
            //   - 反向 `notesSlideN.xml.rels` 的 Slide target：从 `Slide.notes_slide_rel_target` 取
            //     （None 时按 `../slides/slide{idx+1}.xml` 拼，保证指向真实写入的 slide 文件名）。
            // 这样 read→save→read 后 `notesSlideN.xml` 的 partname / 双向 rid / 反向 target 都保持稳定。
            let mut this_notes_rid: Option<String> = None;
            if let Some(notes_tb) = sld.notes() {
                notes_index += 1;
                // rid 与 target 在 slide 这一侧用 `rIdNotes{N}` + `../notesSlides/<part>` 即可，
                // 因为 slideN.xml.rels 总是**新写**的，无需保留历史 rid。
                let notes_rid = format!("rIdNotes{}", notes_index);
                // notes part 路径：优先复用原始 partname，缺失再按序号分配
                let notes_partname = match entry.sld.notes_partname() {
                    Some(p) => p.to_string(),
                    None => format!("/ppt/notesSlides/notesSlide{}.xml", notes_index),
                };
                // notesSlideN.xml.rels 中 Slide 关系的 target：优先复用原始，缺失再按 slide 编号拼
                let notes_slide_target = match entry.sld.notes_slide_rel_target() {
                    Some(t) => t.to_string(),
                    None => format!("../slides/slide{}.xml", i + 1),
                };
                this_notes_rid = Some(notes_rid.clone());
                // 关系挂在 slide 上 → 指向 notesSlide
                //   Target 是相对于 /ppt/slides/slideN.xml 的相对路径，需从 notes_partname 拆出 basename
                let notes_target_rel = {
                    // 例: "/ppt/notesSlides/notesSlide1.xml" → "../notesSlides/notesSlide1.xml"
                    let trimmed = notes_partname.trim_start_matches("/ppt/");
                    format!("../{}", trimmed)
                };
                extra_rels.add(Relationship::internal_str(
                    notes_rid.clone(),
                    RelType::NotesSlide,
                    notes_target_rel,
                ))?;
                // 同一份 notes 在 presentation.xml.rels 上也要出现一次
                // （PowerPoint 实际不需要，但本库保持 1:1 关系清晰）。
                // 这里的 target 用 presentation.xml.rels 视角的相对路径 `notesSlides/<file>`。
                let pres_target_rel = {
                    let trimmed = notes_partname.trim_start_matches("/ppt/");
                    trimmed.to_string()
                };
                pres_rels.add(Relationship::internal_str(
                    format!("rIdNotesPres{}", notes_index),
                    RelType::NotesSlide,
                    pres_target_rel,
                ))?;
                // notesSlideN.xml.rels：1 个 Slide 关系，指向所属 slideN.xml
                //   target 优先复用原始，缺失再按 i+1 拼。
                let mut notes_rels = Relationships::new();
                notes_rels.add(Relationship::internal_str(
                    "rId1",
                    RelType::Slide,
                    notes_slide_target,
                ))?;
                let notes_rels_partname = rels_partname_for(&notes_partname);
                pkg.put_part(Part::new(
                    PartName::from_unchecked(notes_rels_partname),
                    ct::RELATIONSHIPS,
                    notes_rels.to_xml().into_bytes(),
                ));
                // 写 notesSlideN.xml
                let notes_xml_str = crate::oxml::notes_xml(notes_tb);
                pkg.put_part(Part::new(
                    PartName::from_unchecked(notes_partname),
                    ct::NOTES_SLIDE,
                    notes_xml_str.into_bytes(),
                ));
            }
            // 把 notes_rid 同步回 entry（供后续访问）。
            if let Some(rid) = this_notes_rid {
                sld.set_notes_rid(rid);
            }

            // ---------- 评论（comments）写入 ----------
            // 与 notes 类似：每个有评论的 slide 对应一个 `/ppt/comments/commentN.xml`。
            // 关系挂在 slideN.xml.rels 上（Type=comments）。
            let mut this_comments_rid: Option<String> = None;
            if let Some(comment_lst) = sld.comments() {
                comments_index += 1;
                let comments_rid = format!("rIdComments{}", comments_index);
                // partname：优先复用原始，缺失再按序号分配
                let comments_partname = match sld.comments_partname() {
                    Some(p) => p.to_string(),
                    None => format!("/ppt/comments/comment{}.xml", comments_index),
                };
                this_comments_rid = Some(comments_rid.clone());
                // 关系挂在 slide 上 → 指向 comments
                //   Target 是相对于 /ppt/slides/slideN.xml 的相对路径
                let comments_target_rel = {
                    let trimmed = comments_partname.trim_start_matches("/ppt/");
                    format!("../{}", trimmed)
                };
                extra_rels.add(Relationship::internal_str(
                    comments_rid.clone(),
                    RelType::Comments,
                    comments_target_rel,
                ))?;
                // 写 commentN.xml
                let comments_xml_str = comment_lst.to_xml();
                pkg.put_part(Part::new(
                    PartName::from_unchecked(comments_partname),
                    ct::COMMENTS,
                    comments_xml_str.into_bytes(),
                ));
            }
            if let Some(rid) = this_comments_rid {
                sld.set_comments_rid(rid);
            }

            // ---------- 图表（chart）写入 ----------
            // 与 notes/comments 类似：每个 chart 对应一个独立的 `/ppt/charts/chartN.xml` part。
            // 关系挂在 slideN.xml.rels 上（Type=chart）。
            //
            // **关键**：chart 的 `<c:chart r:id="..."/>` 引用必须与 slideN.xml.rels 中的
            // `<Relationship Id="..."/>` 的 Id 完全一致——这里直接复用 ChartEntry.rid
            // （由 `ShapesMut::add_chart` 分配的 `rIdChartN`）。
            //
            // **partname 重写**：add_chart 时用 slide 局部索引占位，这里用全局索引
            // 重新分配 partname，避免多 slide 之间的 chartN.xml 冲突。
            for chart_entry in &entry.sld.chart_entries {
                chart_global_index += 1;
                let chart_partname = format!("/ppt/charts/chart{}.xml", chart_global_index);
                // 在 slideN.xml.rels 中添加 chart 关系
                // Target 是相对于 /ppt/slides/slideN.xml 的相对路径
                let chart_target_rel = format!("../charts/chart{}.xml", chart_global_index);
                extra_rels.add(Relationship::internal_str(
                    chart_entry.rid.clone(),
                    RelType::Chart,
                    chart_target_rel,
                ))?;

                // 嵌入式 Excel 工作簿（TODO-004 Excel 嵌入）。
                //
                // 若 ChartEntry.xlsx_blob 非空，写出：
                // 1. `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` part（全局索引）；
                // 2. `/ppt/charts/_rels/chartN.xml.rels` part（Type=Package，
                //    Target=`../embeddings/Microsoft_Excel_WorksheetN.xlsx`）；
                // 3. 在 chart 模型上设置 external_data_rid 后重新 to_xml。
                //
                // **关键**：chartN.xml.rels 是 chart part 的**独立关系文件**，
                // 与 slideN.xml.rels 分离（chart part 自己的关系挂在 chart 的 _rels 目录）。
                let mut chart_model = chart_entry.chart.clone();
                if let Some(xlsx_blob) = &chart_entry.xlsx_blob {
                    chart_xlsx_global_index += 1;
                    let xlsx_partname = format!(
                        "/ppt/embeddings/Microsoft_Excel_Worksheet{}.xlsx",
                        chart_xlsx_global_index
                    );
                    let xlsx_target_rel = format!(
                        "../embeddings/Microsoft_Excel_Worksheet{}.xlsx",
                        chart_xlsx_global_index
                    );
                    // chart part 内部的关系 id（与 slideN 的 rIdChartN 命名空间分离）
                    let xlsx_rid = format!("rIdXlsx{}", chart_xlsx_global_index);

                    // 写出 xlsx part
                    pkg.put_part(Part::new(
                        PartName::from_unchecked(xlsx_partname.clone()),
                        ct::SPREADSHEET_XLSX,
                        xlsx_blob.clone(),
                    ));

                    // 构造 chartN.xml.rels part 内容
                    let mut chart_rels = crate::opc::rels::Relationships::new();
                    chart_rels.add(Relationship::internal_str(
                        xlsx_rid.clone(),
                        RelType::Package,
                        xlsx_target_rel,
                    ))?;
                    let chart_rels_partname =
                        format!("/ppt/charts/_rels/chart{}.xml.rels", chart_global_index);
                    pkg.put_part(Part::new(
                        PartName::from_unchecked(chart_rels_partname),
                        ct::RELATIONSHIPS,
                        chart_rels.to_xml().into_bytes(),
                    ));

                    // 在 chart 模型上设置 external_data_rid，使 to_xml 输出 <c:externalData r:id="..."/>
                    chart_model.external_data_rid = Some(xlsx_rid);
                }

                // 写出 chartN.xml part（内容为 Chart::to_xml()，可能含 externalData）
                let chart_xml = chart_model.to_xml();
                pkg.put_part(Part::new(
                    PartName::from_unchecked(chart_partname),
                    ct::CHART,
                    chart_xml.into_bytes(),
                ));
            }

            // ---------- OLE 对象（oleObject）写入 ----------（TODO-043）
            // 与 chart 类似：每个 ole 对应一个独立的 `/ppt/embeddings/oleObjectN.bin` part。
            // 关系挂在 slideN.xml.rels 上（Type=oleObject）。
            //
            // **关键**：oleObj 的 `<p:oleObj r:id="..."/>` 引用必须与 slideN.xml.rels 中的
            // `<Relationship Id="..."/>` 的 Id 完全一致——这里直接复用 OleEntry.rid
            // （由 `ShapesMut::add_ole_object` 分配的 `rIdOleN`）。
            //
            // **partname 重写**：add_ole_object 时用 slide 局部索引占位，这里用全局索引
            // 重新分配 partname，避免多 slide 之间的 oleObjectN.bin 冲突。
            for ole_entry in &entry.sld.ole_entries {
                ole_global_index += 1;
                let ole_partname = format!("/ppt/embeddings/oleObject{}.bin", ole_global_index);
                // 在 slideN.xml.rels 中添加 oleObject 关系
                // Target 是相对于 /ppt/slides/slideN.xml 的相对路径
                let ole_target_rel = format!("../embeddings/oleObject{}.bin", ole_global_index);
                extra_rels.add(Relationship::internal_str(
                    ole_entry.rid.clone(),
                    RelType::OleObject,
                    ole_target_rel,
                ))?;
                // 写出 oleObjectN.bin part（内容为原始 OLE 二进制 blob）
                pkg.put_part(Part::new(
                    PartName::from_unchecked(ole_partname),
                    ct::OLE_OBJECT,
                    ole_entry.blob.clone(),
                ));
            }

            // ---------- 视频（video）写入 ----------（TODO-033）
            // 与 ole 类似：每个 video 对应一个独立的 `/ppt/media/mediaN.mp4` part。
            // 关系挂在 slideN.xml.rels 上（Type=video），Target 用相对路径 `../media/mediaN.mp4`。
            //
            // **关键**：视频用 `r:link` 引用（不是 `r:embed`），所以关系类型是 `.../video`，
            // 而非 `.../image`。`<a:videoFile r:link="rIdVideoN"/>` 的 rIdVideoN 必须与
            // slideN.xml.rels 中的 `<Relationship Id="rIdVideoN"/>` 完全一致。
            //
            // **partname 重写**：add_video 时用 slide 局部索引占位，这里用全局索引
            // 重新分配 partname，避免多 slide 之间的 mediaN.mp4 冲突。
            for video_entry in &entry.sld.video_entries {
                video_global_index += 1;
                let video_partname = format!("/ppt/media/media{}.mp4", video_global_index);
                // 在 slideN.xml.rels 中添加 video 关系
                let video_target_rel = format!("../media/media{}.mp4", video_global_index);
                extra_rels.add(Relationship::internal_str(
                    video_entry.rid.clone(),
                    RelType::Video,
                    video_target_rel,
                ))?;
                // 写出 mediaN.mp4 part（内容为原始视频二进制 blob）
                pkg.put_part(Part::new(
                    PartName::from_unchecked(video_partname),
                    ct::VIDEO_MP4,
                    video_entry.blob.clone(),
                ));
            }

            // ---------- 音频（audio）写入 ----------（TODO-033）
            // 与 video 完全对称，仅媒体类型与 Content-Type 不同。
            for audio_entry in &entry.sld.audio_entries {
                audio_global_index += 1;
                let audio_partname = format!("/ppt/media/media{}.mp3", audio_global_index);
                let audio_target_rel = format!("../media/media{}.mp3", audio_global_index);
                extra_rels.add(Relationship::internal_str(
                    audio_entry.rid.clone(),
                    RelType::Audio,
                    audio_target_rel,
                ))?;
                pkg.put_part(Part::new(
                    PartName::from_unchecked(audio_partname),
                    ct::AUDIO_MP3,
                    audio_entry.blob.clone(),
                ));
            }

            // ---------- SmartArt（diagram）写入 ----------（TODO-037）
            // 与 chart/ole/video 不同：每个 diagram 对应 **4 个**独立 part
            // （`/ppt/diagrams/{data,layout,quickStyles,colors}N.xml`）。
            // 4 个关系都挂在 slideN.xml.rels 上：
            //   - `<Relationship Type=".../diagramData" Target="../diagrams/dataN.xml"/>`
            //   - `<Relationship Type=".../diagramLayout" Target="../diagrams/layoutN.xml"/>`
            //   - `<Relationship Type=".../diagramQuickStyle" Target="../diagrams/quickStylesN.xml"/>`
            //   - `<Relationship Type=".../diagramColors" Target="../diagrams/colorsN.xml"/>`
            //
            // slide xml 的 `<p:graphicFrame>` 内 `<dgm:relIds r:dm="..." r:lo="..." r:qs="..." r:cs="..."/>`
            // 4 个属性分别引用这 4 个 rId。
            //
            // **round-trip**：DiagramEntry 持有 4 份原始 XML 字符串，写出时直接写入 zip
            // （不重新序列化），保证任何 SmartArt 模板都能正确保留。
            //
            // **partname 重写**：add_diagram 时用 slide 局部索引占位，这里用全局索引
            // 重新分配 partname，避免多 slide 之间的 dataN.xml 冲突。
            for diagram_entry in &entry.sld.diagram_entries {
                diagram_global_index += 1;
                let idx = diagram_global_index;

                // 4 个 partname（基于全局索引）
                let data_partname = format!("/ppt/diagrams/data{}.xml", idx);
                let layout_partname = format!("/ppt/diagrams/layout{}.xml", idx);
                let quick_style_partname = format!("/ppt/diagrams/quickStyles{}.xml", idx);
                let colors_partname = format!("/ppt/diagrams/colors{}.xml", idx);

                // 4 个相对 Target 路径（相对于 /ppt/slides/slideN.xml）
                let data_target_rel = format!("../diagrams/data{}.xml", idx);
                let layout_target_rel = format!("../diagrams/layout{}.xml", idx);
                let quick_style_target_rel = format!("../diagrams/quickStyles{}.xml", idx);
                let colors_target_rel = format!("../diagrams/colors{}.xml", idx);

                // 在 slideN.xml.rels 中添加 4 个关系
                extra_rels.add(Relationship::internal_str(
                    diagram_entry.data_rid.clone(),
                    RelType::DiagramData,
                    data_target_rel,
                ))?;
                extra_rels.add(Relationship::internal_str(
                    diagram_entry.layout_rid.clone(),
                    RelType::DiagramLayout,
                    layout_target_rel,
                ))?;
                extra_rels.add(Relationship::internal_str(
                    diagram_entry.quick_style_rid.clone(),
                    RelType::DiagramQuickStyle,
                    quick_style_target_rel,
                ))?;
                extra_rels.add(Relationship::internal_str(
                    diagram_entry.colors_rid.clone(),
                    RelType::DiagramColors,
                    colors_target_rel,
                ))?;

                // 写出 4 个 diagram part（内容为原始 XML 字符串）
                pkg.put_part(Part::new(
                    PartName::from_unchecked(data_partname),
                    ct::DIAGRAM_DATA,
                    diagram_entry.data_xml.clone().into_bytes(),
                ));
                pkg.put_part(Part::new(
                    PartName::from_unchecked(layout_partname),
                    ct::DIAGRAM_LAYOUT,
                    diagram_entry.layout_xml.clone().into_bytes(),
                ));
                pkg.put_part(Part::new(
                    PartName::from_unchecked(quick_style_partname),
                    ct::DIAGRAM_QUICK_STYLE,
                    diagram_entry.quick_style_xml.clone().into_bytes(),
                ));
                pkg.put_part(Part::new(
                    PartName::from_unchecked(colors_partname),
                    ct::DIAGRAM_COLORS,
                    diagram_entry.colors_xml.clone().into_bytes(),
                ));
            }

            let extra_rels_xml = extra_rels.to_xml();
            let extra_rels_partname = rels_partname_for(partname.as_str());
            pkg.put_part(Part::new(
                PartName::from_unchecked(extra_rels_partname),
                ct::RELATIONSHIPS,
                extra_rels_xml.into_bytes(),
            ));

            // slide 本体 XML。
            pkg.put_part(Part::new(partname, ct::SLIDE, sld.to_xml().into_bytes()));

            // 在 presentation.xml.rels 中注册该 slide。
            pres_rels.add(Relationship::internal_str(
                rid.clone(),
                RelType::Slide,
                format!("slides/slide{}.xml", i + 1),
            ))?;
            // 同步进 sldIdLst。
            pres_root.slide_ids.push(SlideIdEntry {
                id: entry.sld_id,
                rid: rid.clone(),
            });
        }

        // ---------- 评论作者列表（commentAuthors.xml）----------
        // 全局共享的作者清单，非空时写出 `/ppt/commentAuthors.xml`，
        // 并在 `presentation.xml.rels` 添加 `commentAuthors` 关系。
        if !self.comment_authors.is_empty() {
            let authors_xml = self.comment_authors.to_xml();
            pkg.put_part(Part::new(
                new_part_name("/ppt/commentAuthors.xml"),
                ct::COMMENT_AUTHORS,
                authors_xml.into_bytes(),
            ));
            pres_rels.add(Relationship::internal_str(
                "rIdCommentAuthors",
                RelType::CommentAuthors,
                "commentAuthors.xml",
            ))?;
        }

        // ---------------- 7) presentation.xml ----------------
        // 若设置了修改密码保护，注入 modifyVerifier 到 presentation.xml。
        let pres_xml = if let Some(ref mp) = self.modify_protection {
            inject_modify_verifier(&pres_root.to_xml(), &mp.to_xml_element())
        } else {
            pres_root.to_xml()
        };
        pkg.put_part(Part::new(
            new_part_name("/ppt/presentation.xml"),
            ct::PRESENTATION,
            pres_xml.into_bytes(),
        ));
        // ---------------- 8) presentation.xml.rels ----------------
        pkg.put_part(Part::new(
            new_part_name("/ppt/_rels/presentation.xml.rels"),
            ct::RELATIONSHIPS,
            pres_rels.to_xml().into_bytes(),
        ));

        // ---------------- 9) 媒体（图片） ----------------
        // 收集所有 slide 的 media 一起写 zip（去重 by partname）。
        use std::collections::BTreeMap;
        let mut seen: BTreeMap<String, ()> = BTreeMap::new();
        for entry in self.slides.iter() {
            for m in &entry.sld.media_entries {
                if seen.insert(m.partname.as_str().to_string(), ()).is_none() {
                    pkg.put_part(Part::new(
                        m.partname.clone(),
                        m.content_type.clone(),
                        m.blob.clone(),
                    ));
                }
            }
        }

        // ---------------- 10) presProps.xml ----------------
        // 仅保留 `<p:showPr/>`，PowerPoint 接受空集。
        let pres_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:showPr/>
</p:presentationPr>"#;
        pkg.put_part(Part::new(
            new_part_name("/ppt/presProps.xml"),
            "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
            pres_props_xml.as_bytes().to_vec(),
        ));

        // ---------------- 11) viewProps.xml ----------------
        // 固定写法：normalViewPr 必填，否则 PowerPoint 会发出修复提示。
        let view_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewProps xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
</p:viewProps>"#;
        pkg.put_part(Part::new(
            new_part_name("/ppt/viewProps.xml"),
            "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
            view_props_xml.as_bytes().to_vec(),
        ));

        // ---------------- 12) tableStyles.xml ----------------
        // 必备辅件 —— 留空表样式列表。
        let table_styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" defStyleStyle=""/>
"#;
        pkg.put_part(Part::new(
            new_part_name("/ppt/tableStyles.xml"),
            "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
            table_styles_xml.as_bytes().to_vec(),
        ));

        Ok(pkg)
    }

    /// 确保演示文稿**至少**有 1 个母版 + 1 个版式。
    ///
    /// # 行为
    ///
    /// - 若 `slide_masters` 为空，则 push 一个默认母版（指向 `slideMaster1.xml`）；
    /// - 若 `slide_layouts` 为空，则 push 一个默认空白版式（指向 `slideLayout1.xml`）。
    ///
    /// 通常由 [`Presentation::new`] / [`Presentation::from_opc`] 调用，
    /// 但也可在 `open` 之后再次手动调用以"补全缺失项"。
    fn ensure_default_master_and_layout(&mut self) -> crate::Result<()> {
        if self.slide_masters.is_empty() {
            self.slide_masters.items.push(SlideMasterRef {
                idx: 0,
                partname: "/ppt/slideMasters/slideMaster1.xml".to_string(),
                rid: "rIdMaster1".to_string(),
                oxml: Rc::new(std::cell::RefCell::new(OxmlSldMaster::default())),
            });
        }
        if self.slide_layouts.is_empty() {
            self.slide_layouts.items.push(SlideLayoutRef {
                idx: 0,
                partname: "/ppt/slideLayouts/slideLayout1.xml".to_string(),
                rid: "rIdLayout1".to_string(),
                oxml: Rc::new(std::cell::RefCell::new(OxmlSldLayout::default())),
            });
        }
        Ok(())
    }
}

/// 把 OPC 关系中的"相对 target"解析为绝对 partname。
///
/// # OPC 相对路径规则
/// 关系文件中 `Target` 相对于**父 part 的目录**（不是相对于 .rels 文件）。
///
/// # 示例
/// - 父 `/ppt/slides/slide1.xml` + 相对 `../notesSlides/notesSlide1.xml`
///   → `/ppt/notesSlides/notesSlide1.xml`
/// - 父 `/ppt/slides/slide1.xml` + 相对 `media/image1.png`
///   → `/ppt/slides/media/image1.png`
/// - 父 `/ppt/slides/slide1.xml` + 相对 `../media/image1.png`
///   → `/ppt/media/image1.png`
fn resolve_relative_partname(parent_partname: &str, rel_target: &str) -> String {
    // 父 partname 必须以 '/' 开头；如果是绝对路径，直接返回。
    if rel_target.starts_with('/') {
        return rel_target.to_string();
    }
    // 父 partname 的目录
    let parent_dir = match parent_partname.rfind('/') {
        Some(i) => &parent_partname[..i],
        None => "",
    };
    // 拼接并标准化
    let mut stack: Vec<&str> = Vec::new();
    let combined = if parent_dir.is_empty() {
        format!("/{}", rel_target)
    } else {
        format!("{}/{}", parent_dir, rel_target)
    };
    for seg in combined.split('/') {
        match seg {
            "" | "." => continue,
            ".." => {
                stack.pop();
            }
            _ => stack.push(seg),
        }
    }
    let mut out = String::from("/");
    out.push_str(&stack.join("/"));
    if out == "/" && !stack.is_empty() {
        // 不会发生，stack 至少 1 段
    }
    out
}

/// 在 presentation.xml 中注入 `<p:modifyVerifier .../>`。
///
/// # 行为
///
/// 1. 若已含 `p:modifyVerifier`，**原样返回**（幂等保护）；
/// 2. 否则把它插到 `<p:extLst` 之前（OOXML schema 顺序要求）；
/// 3. 找不到 extLst 时兜底插到 `</p:presentation>` 之前。
fn inject_modify_verifier(pres_xml: &str, verifier_xml: &str) -> String {
    if pres_xml.contains("p:modifyVerifier") {
        return pres_xml.to_string();
    }
    // 优先级 1：插到 <p:extLst 之前
    if let Some(pos) = pres_xml.find("<p:extLst") {
        let mut out = String::with_capacity(pres_xml.len() + verifier_xml.len());
        out.push_str(&pres_xml[..pos]);
        out.push_str(verifier_xml);
        out.push_str(&pres_xml[pos..]);
        return out;
    }
    // 优先级 2：插到 </p:presentation> 之前
    if let Some(pos) = pres_xml.rfind("</p:presentation>") {
        let mut out = String::with_capacity(pres_xml.len() + verifier_xml.len());
        out.push_str(&pres_xml[..pos]);
        out.push_str(verifier_xml);
        out.push_str(&pres_xml[pos..]);
        return out;
    }
    // 兜底：未修改
    pres_xml.to_string()
}

impl Default for Presentation {
    /// `Default` 等价于 [`Presentation::new`]。
    ///
    /// `Default` trait 签名无法返回 `Result`，因此此处使用 `expect`。
    /// 安全性保证：`Presentation::new()` 当前实现不会返回 `Err`（仅做
    /// 字段初始化），如果未来 `new()` 可能失败，应移除 `Default` impl
    /// 并改为关联常量或工厂方法。
    // 允许 expect：Default trait 签名限制，无法用 ? 传播错误
    #[allow(clippy::expect_used)]
    fn default() -> Self {
        Self::new().expect("Presentation::new() is infallible in current implementation")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::units::Inches;

    /// 验证 `resolve_relative_partname` 处理各种 OOXML 相对路径。
    #[test]
    fn resolve_relative_partname_examples() {
        assert_eq!(
            resolve_relative_partname("/ppt/slides/slide1.xml", "../notesSlides/notesSlide1.xml"),
            "/ppt/notesSlides/notesSlide1.xml"
        );
        assert_eq!(
            resolve_relative_partname("/ppt/slides/slide1.xml", "media/image1.png"),
            "/ppt/slides/media/image1.png"
        );
        assert_eq!(
            resolve_relative_partname("/ppt/slides/slide1.xml", "../media/image1.png"),
            "/ppt/media/image1.png"
        );
        // 绝对路径原样返回
        assert_eq!(
            resolve_relative_partname("/ppt/slides/slide1.xml", "/ppt/notesSlides/n1.xml"),
            "/ppt/notesSlides/n1.xml"
        );
        // 同目录
        assert_eq!(
            resolve_relative_partname("/ppt/slides/slide1.xml", "slide2.xml"),
            "/ppt/slides/slide2.xml"
        );
    }

    /// 完整 read → modify → write 的端到端测试。
    ///
    /// # 流程
    /// 1. 新建一份带 1 张 slide（含文本框 + 备注）的演示文稿；
    /// 2. 序列化为 zip 字节；
    /// 3. 用 `load_bytes` 读回；
    /// 4. 验证：slide 数量、文本框文字、备注文本、画布尺寸都正确还原。
    #[test]
    fn roundtrip_presentation_loads_back() {
        // 1) 新建
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide
            .shapes_mut()
            .add_textbox_with_text(
                Inches(1.0),
                Inches(1.0),
                Inches(4.0),
                Inches(1.0),
                "Hello roundtrip",
            )
            .expect("textbox");
        slide.set_notes_text(Some("speaker note line 1\nline 2"));
        p.set_slide_size(Emu(7_000_000), Emu(5_000_000));

        // 2) 序列化为字节
        let bytes = p.to_bytes().expect("to_bytes ok");
        assert!(!bytes.is_empty());

        // 3) 读回
        let p2 = Presentation::load_bytes(&bytes).expect("load_bytes ok");

        // 4) 验证
        assert_eq!(p2.slides().len(), 1);
        let s = p2.slides().get(0).expect("slide 0");
        // to_opc_package 内部把 inner.layout_rid 强制写为 "rId1"（slide rels 文件的固定 rId），
        // 所以加载后 laytout_rid 应为 "rId1"，与 [Content_Types].xml 的 SlideLayout 关系对应。
        assert_eq!(s.sld.layout_rid(), "rId1");
        // 文本体（直接走内层 oxml，避开高阶 view 的类型复杂度）
        let inner_shapes = &s.sld.inner.shapes;
        assert_eq!(inner_shapes.len(), 1);
        if let crate::oxml::SlideShape::Sp(sp) = &inner_shapes[0] {
            assert_eq!(sp.text.paragraphs.len(), 1);
            assert_eq!(sp.text.paragraphs[0].runs[0].text, "Hello roundtrip");
        } else {
            panic!("expected Sp");
        }
        // 备注
        let notes = s.sld.notes_text().expect("notes present");
        assert!(notes.contains("speaker note line 1"));
        // 画布尺寸
        assert_eq!(p2.slide_width().0, 7_000_000);
        assert_eq!(p2.slide_height().0, 5_000_000);
    }

    /// 验证在 from_opc 之后再次 save 仍能产出可解析的 pptx。
    #[test]
    fn load_then_resave_preserves_content() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide
            .shapes_mut()
            .add_textbox_with_text(Inches(0.5), Inches(0.5), Inches(6.0), Inches(0.8), "stable")
            .expect("textbox");
        let bytes = p.to_bytes().expect("to_bytes");

        // 二次 load + save（写到 Vec<u8>，避免依赖磁盘）
        let p2 = Presentation::load_bytes(&bytes).expect("load 1");
        let bytes2 = p2.to_bytes().expect("to_bytes 2");
        let p3 = Presentation::load_bytes(&bytes2).expect("load 2");

        assert_eq!(p3.slides().len(), 1);
        let s = p3.slides().get(0).expect("slide 0");
        let inner_shapes = &s.sld.inner.shapes;
        if let crate::oxml::SlideShape::Sp(sp) = &inner_shapes[0] {
            assert_eq!(sp.text.paragraphs[0].runs[0].text, "stable");
        } else {
            panic!("expected Sp");
        }
    }

    /// 验证 notes partname + 反向 rels target 在 read→save→read 后**保持稳定**。
    ///
    /// 关键点：
    /// - 第一次 save 后写入的 `notesSlideN.xml` 路径与第二次 load 看到的路径一致；
    /// - `notesSlideN.xml.rels` 中指向 slide 的 target 在二次 save 后不变；
    /// - 备注文本内容在三轮 round-trip 后正确还原。
    #[test]
    fn notes_roundtrip_preserves_partname_and_rels_target() {
        // 1) 构造原始 pptx（含 1 张带 notes 的 slide）
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide
            .shapes_mut()
            .add_textbox_with_text(
                Inches(0.5),
                Inches(0.5),
                Inches(6.0),
                Inches(0.8),
                "roundtrip notes test",
            )
            .expect("textbox");
        slide.set_notes_text(Some("first notes\nline 2"));
        let bytes1 = p.to_bytes().expect("to_bytes 1");

        // 2) load + 修改 notes
        let mut p2 = Presentation::load_bytes(&bytes1).expect("load 1");
        p2.slides_mut()
            .get_mut(0)
            .expect("slide 0")
            .sld
            .set_notes_text(Some("updated notes"));
        let bytes2 = p2.to_bytes().expect("to_bytes 2");

        // 3) load + 校验
        let p3 = Presentation::load_bytes(&bytes2).expect("load 2");
        let s = p3.slides().get(0).expect("slide 0");
        let notes_text = s.sld.notes_text().expect("notes present");
        assert!(notes_text.contains("updated notes"));

        // 4) 解析 bytes2 的 zip，校验 notesSlideN.xml.rels 中指向 slide 的 target 仍是 ../slides/slide1.xml
        let cursor = std::io::Cursor::new(bytes2);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        let mut nrels_xml = String::new();
        zip.by_name("ppt/notesSlides/_rels/notesSlide1.xml.rels")
            .expect("notesSlide1.xml.rels exists")
            .read_to_string(&mut nrels_xml)
            .expect("read notes rels");
        // 反向 target 应该指向 slide1.xml
        assert!(
            nrels_xml.contains("../slides/slide1.xml"),
            "notesSlide1.xml.rels missing slide rels, got: {nrels_xml}"
        );
    }

    /// 验证新增的 slide（无历史 notes 元数据）也能正常序列化。
    ///
    /// 这是为了确认 `notes_partname=None` 的回退路径仍然按 `notesSlide{N}.xml` 分配。
    #[test]
    fn new_slide_with_notes_allocates_fresh_partname() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let s = p.slides_mut().add_slide(counter).expect("add slide");
        s.set_notes_text(Some("brand new"));
        let bytes = p.to_bytes().expect("to_bytes");

        let p2 = Presentation::load_bytes(&bytes).expect("load");
        let s2 = p2.slides().get(0).expect("slide 0");
        let notes = s2.sld.notes_text().expect("notes present");
        assert!(notes.contains("brand new"));

        // 校验 zip 中确实有 notesSlide1.xml
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        assert!(zip.by_name("ppt/notesSlides/notesSlide1.xml").is_ok());
    }

    /// 验证多 slide + 多 notes 时 read→save→read 的关系图谱都正确。
    ///
    /// 这是"补全 24"的主回归测试：确保新加的 partname / 反向 rels target
    /// 字段在多 slide 场景下不串号。
    #[test]
    fn notes_roundtrip_multiple_slides() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        // 3 张 slide：notes 在 1 和 3，2 没有
        // `add_slide` 取得 `counter` 所有权，所以每次循环用 `counter.clone()`。
        for i in 0..3 {
            let s = p
                .slides_mut()
                .add_slide(counter.clone())
                .expect("add slide");
            if i != 1 {
                s.set_notes_text(Some(&format!("note for slide {}", i + 1)));
            }
        }
        let bytes1 = p.to_bytes().expect("to_bytes 1");

        // read
        let mut p2 = Presentation::load_bytes(&bytes1).expect("load");
        // 修改所有有 notes 的 slide
        {
            let sm = p2.slides_mut();
            sm.get_mut(0).unwrap().sld.set_notes_text(Some("edited 1"));
            // skip idx 1
            sm.get_mut(2).unwrap().sld.set_notes_text(Some("edited 3"));
        }
        let bytes2 = p2.to_bytes().expect("to_bytes 2");

        // read again
        let p3 = Presentation::load_bytes(&bytes2).expect("load 2");
        assert_eq!(p3.slides().len(), 3);
        let s0 = p3.slides().get(0).expect("s0");
        let s1 = p3.slides().get(1).expect("s1");
        let s2 = p3.slides().get(2).expect("s2");
        assert_eq!(s0.sld.notes_text().as_deref(), Some("edited 1"));
        assert_eq!(s1.sld.notes_text(), None);
        assert_eq!(s2.sld.notes_text().as_deref(), Some("edited 3"));

        // 校验 zip 中 notesSlide1.xml 和 notesSlide2.xml 都存在
        // （第 2 张 slide 没 notes，所以 notesSlide2 对应 slide3 的 notes）
        let cursor = std::io::Cursor::new(bytes2);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        assert!(zip.by_name("ppt/notesSlides/notesSlide1.xml").is_ok());
        assert!(zip.by_name("ppt/notesSlides/notesSlide2.xml").is_ok());
        assert!(zip
            .by_name("ppt/notesSlides/_rels/notesSlide1.xml.rels")
            .is_ok());
        assert!(zip
            .by_name("ppt/notesSlides/_rels/notesSlide2.xml.rels")
            .is_ok());
    }

    /// 文档测试：明确**记录** `layout_rid` 在 read→save 后会被**强制重置**为 `"rId1"`。
    ///
    /// # 背景
    /// 现实世界中的 .pptx 文件里，`slideN.xml.rels` 中的 layout 关系可能叫
    /// `rId1` / `rId2` / `rId3` 甚至 `rIdLayout`（自定义字符串）。本库**在
    /// `to_opc_package` 中强制把每个 slide 的 layout 关系重写为 `rId1`**，因为：
    /// 1. 0.1.0 默认仅生成一份 `slideLayout1.xml`；
    /// 2. 简化 read 路径，不再需要追踪 layout part 的"原始 rid"；
    /// 3. 节省 1 个字段（`layout_partname`）的状态存储。
    ///
    /// # 含义
    /// - **不要**依赖 `layout_rid()` 在 read→save→read 后保持原始值；
    /// - 若未来需要保留原始 rid，应在 `Slide` 上新增 `layout_rid: Option<String>`
    ///   字段并在 `to_opc_package` 中按"非空则用，空则 rId1"的策略写。
    #[test]
    fn layout_rid_is_overwritten_to_rid1_on_save() {
        // 1) 构造原始 pptx
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide
            .shapes_mut()
            .add_textbox_with_text(
                Inches(0.5),
                Inches(0.5),
                Inches(4.0),
                Inches(0.8),
                "layout_rid reset test",
            )
            .expect("textbox");
        // 故意把 layout_rid 改成奇怪的值，验证 save 会被覆盖
        slide.set_layout_rid("rIdLayout999".to_string());
        let bytes1 = p.to_bytes().expect("to_bytes 1");

        // 2) load + 校验
        let p2 = Presentation::load_bytes(&bytes1).expect("load");
        let s = p2.slides().get(0).expect("slide 0");
        // 关键断言：read 后必为 "rId1"
        assert_eq!(
            s.sld.layout_rid(),
            "rId1",
            "to_opc_package 应把 layout_rid 重置为 rId1（与 slide rels 的固定 rId 对应）"
        );

        // 3) 校验 zip 中 slide1.xml.rels 里 layout 关系的 id 也是 rId1
        let cursor = std::io::Cursor::new(bytes1);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        let mut srels_xml = String::new();
        zip.by_name("ppt/slides/_rels/slide1.xml.rels")
            .expect("slide1.xml.rels exists")
            .read_to_string(&mut srels_xml)
            .expect("read slide rels");
        assert!(
            srels_xml.contains("Id=\"rId1\""),
            "slide1.xml.rels 中 layout 关系应使用 rId1，实际：{srels_xml}"
        );
    }

    // ==================== 自定义文档属性测试（TODO-034） ====================

    /// `CustomProperties::set` / `get` / `remove` 基本 API。
    #[test]
    fn custom_properties_set_get_remove() {
        let mut props = CustomProperties::new();
        assert!(props.is_empty());
        assert_eq!(props.len(), 0);

        props.set("Project", CustomPropertyValue::Text("Demo".to_string()));
        props.set("Version", CustomPropertyValue::Int(42));
        props.set("Active", CustomPropertyValue::Bool(true));
        assert_eq!(props.len(), 3);

        // get
        match props.get("Project") {
            Some(CustomPropertyValue::Text(s)) => assert_eq!(s, "Demo"),
            other => panic!("期望 Text，得到 {:?}", other),
        }
        match props.get("Version") {
            Some(CustomPropertyValue::Int(i)) => assert_eq!(*i, 42),
            other => panic!("期望 Int，得到 {:?}", other),
        }
        assert!(props.get("NonExistent").is_none());

        // 覆盖已有值
        props.set("Version", CustomPropertyValue::Int(100));
        match props.get("Version") {
            Some(CustomPropertyValue::Int(i)) => assert_eq!(*i, 100),
            other => panic!("期望 Int(100)，得到 {:?}", other),
        }
        assert_eq!(props.len(), 3, "覆盖不应增加条目数");

        // remove
        let removed = props.remove("Active");
        assert!(matches!(removed, Some(CustomPropertyValue::Bool(true))));
        assert_eq!(props.len(), 2);
        assert!(props.get("Active").is_none());
    }

    /// `CustomProperties::to_xml` 正确序列化各种值类型。
    #[test]
    fn custom_properties_to_xml() {
        let mut props = CustomProperties::new();
        props.set(
            "Text",
            CustomPropertyValue::Text("Hello & World".to_string()),
        );
        props.set("Int", CustomPropertyValue::Int(42));
        // 注：避开 3.14（clippy::approx_constant 会误判为 π 近似值）。
        props.set("Float", CustomPropertyValue::Float(3.15));
        props.set("Bool", CustomPropertyValue::Bool(true));

        let xml = props.to_xml();
        assert!(xml.contains("<Properties"));
        assert!(xml.contains("custom-properties"));
        // XML 转义
        assert!(xml.contains("Hello &amp; World"));
        // pid 从 2 开始
        assert!(xml.contains("pid=\"2\""));
        assert!(xml.contains("pid=\"3\""));
        // 各值类型
        assert!(xml.contains("<vt:lpwstr>Hello &amp; World</vt:lpwstr>"));
        assert!(xml.contains("<vt:i4>42</vt:i4>"));
        assert!(xml.contains("<vt:r8>3.15</vt:r8>"));
        assert!(xml.contains("<vt:bool>true</vt:bool>"));
        // fmtid
        assert!(xml.contains("fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\""));
    }

    /// `CustomProperties::from_xml` 正确解析各种值类型。
    #[test]
    fn custom_properties_from_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Text">
    <vt:lpwstr>Value</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Int">
    <vt:i4>42</vt:i4>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Bool">
    <vt:bool>false</vt:bool>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Float">
    <vt:r8>3.15</vt:r8>
  </property>
</Properties>"#;
        let props = CustomProperties::from_xml(xml);
        assert_eq!(props.len(), 4);
        match props.get("Text") {
            Some(CustomPropertyValue::Text(s)) => assert_eq!(s, "Value"),
            other => panic!("期望 Text，得到 {:?}", other),
        }
        match props.get("Int") {
            Some(CustomPropertyValue::Int(i)) => assert_eq!(*i, 42),
            other => panic!("期望 Int，得到 {:?}", other),
        }
        match props.get("Bool") {
            Some(CustomPropertyValue::Bool(b)) => assert!(!*b),
            other => panic!("期望 Bool(false)，得到 {:?}", other),
        }
        match props.get("Float") {
            Some(CustomPropertyValue::Float(f)) => assert!((f - 3.15).abs() < 1e-6),
            other => panic!("期望 Float，得到 {:?}", other),
        }
    }

    /// `CustomProperties` 的 XML 往返（to_xml → from_xml）保持一致。
    #[test]
    fn custom_properties_round_trip() {
        let mut props = CustomProperties::new();
        props.set("Key1", CustomPropertyValue::Text("Value1".to_string()));
        props.set("Key2", CustomPropertyValue::Int(123));
        props.set("Key3", CustomPropertyValue::Bool(true));

        let xml = props.to_xml();
        let parsed = CustomProperties::from_xml(&xml);
        assert_eq!(parsed.len(), 3);
        assert_eq!(props.get("Key1"), parsed.get("Key1"));
        assert_eq!(props.get("Key2"), parsed.get("Key2"));
        assert_eq!(props.get("Key3"), parsed.get("Key3"));
    }

    /// 空 `CustomProperties` 不写出 custom.xml。
    #[test]
    fn custom_properties_empty_not_serialized() {
        let p = Presentation::new().expect("new ok");
        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        // 空自定义属性时不应有 custom.xml
        assert!(zip.by_name("docProps/custom.xml").is_err());
    }

    /// 非空 `CustomProperties` 写出 custom.xml 并在 _rels/.rels 中添加关系。
    #[test]
    fn custom_properties_serialized_to_pptx() {
        let mut p = Presentation::new().expect("new ok");
        p.custom_properties_mut()
            .set("Project", CustomPropertyValue::Text("Test".to_string()));
        p.custom_properties_mut()
            .set("Version", CustomPropertyValue::Int(1));

        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");

        // 1) custom.xml 存在
        let mut custom_xml = String::new();
        zip.by_name("docProps/custom.xml")
            .expect("custom.xml 应存在")
            .read_to_string(&mut custom_xml)
            .expect("read custom.xml");
        assert!(custom_xml.contains("name=\"Project\""));
        assert!(custom_xml.contains("<vt:lpwstr>Test</vt:lpwstr>"));
        assert!(custom_xml.contains("name=\"Version\""));
        assert!(custom_xml.contains("<vt:i4>1</vt:i4>"));

        // 2) _rels/.rels 中有 custom-properties 关系
        let mut rels_xml = String::new();
        zip.by_name("_rels/.rels")
            .expect(".rels 应存在")
            .read_to_string(&mut rels_xml)
            .expect("read .rels");
        assert!(rels_xml.contains("custom-properties"), "rels: {}", rels_xml);
        assert!(
            rels_xml.contains("docProps/custom.xml"),
            "rels: {}",
            rels_xml
        );
    }

    /// read→save→read 往返保持自定义属性。
    #[test]
    fn custom_properties_round_trip_through_pptx() {
        let mut p = Presentation::new().expect("new ok");
        p.custom_properties_mut()
            .set("Author", CustomPropertyValue::Text("TestUser".to_string()));
        p.custom_properties_mut()
            .set("Count", CustomPropertyValue::Int(99));

        let bytes = p.to_bytes().expect("to_bytes");
        let p2 = Presentation::load_bytes(&bytes).expect("load");

        assert_eq!(p2.custom_properties().len(), 2);
        match p2.custom_properties().get("Author") {
            Some(CustomPropertyValue::Text(s)) => assert_eq!(s, "TestUser"),
            other => panic!("期望 Text，得到 {:?}", other),
        }
        match p2.custom_properties().get("Count") {
            Some(CustomPropertyValue::Int(i)) => assert_eq!(*i, 99),
            other => panic!("期望 Int，得到 {:?}", other),
        }
    }

    // ==================== 评论测试（TODO-036） ====================

    /// `Slide::add_comment` 基本 API：添加评论后能读回。
    #[test]
    fn slide_add_comment_basic() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let author_id = p.comment_authors_mut().get_or_insert_id("张三", "ZS");
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        let idx = slide.add_comment(
            author_id,
            crate::Emu::new(914400),
            crate::Emu::new(914400),
            "测试评论",
        );
        assert_eq!(idx, 1);
        let comments = slide.comments().expect("comments should exist");
        assert_eq!(comments.len(), 1);
        assert_eq!(comments.comments[0].text, "测试评论");
        assert_eq!(comments.comments[0].author_id, author_id);
        assert_eq!(comments.comments[0].idx, 1);
        assert_eq!(comments.comments[0].pos_x, 914400);
        assert_eq!(comments.comments[0].pos_y, 914400);
    }

    /// 多条评论的 idx 自动递增。
    #[test]
    fn slide_add_comment_multiple_idx_increments() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let author_id = p.comment_authors_mut().get_or_insert_id("作者", "ZZ");
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        let idx1 = slide.add_comment(author_id, 0, 0, "第一条");
        let idx2 = slide.add_comment(author_id, 100, 100, "第二条");
        let idx3 = slide.add_comment(author_id, 200, 200, "第三条");
        assert_eq!(idx1, 1);
        assert_eq!(idx2, 2);
        assert_eq!(idx3, 3);
        assert_eq!(slide.comments().unwrap().len(), 3);
    }

    /// `clear_comments` 清除所有评论。
    #[test]
    fn slide_clear_comments() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let author_id = p.comment_authors_mut().get_or_insert_id("A", "A");
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide.add_comment(author_id, 0, 0, "评论");
        assert!(slide.comments().is_some());
        slide.clear_comments();
        assert!(slide.comments().is_none());
    }

    /// 评论 + 作者序列化到 PPTX 后能读回。
    #[test]
    fn comments_serialized_to_pptx() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let author_id = p.comment_authors_mut().get_or_insert_id("张三", "ZS");
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide.add_comment(
            author_id,
            crate::Emu::new(914400),
            crate::Emu::new(914400),
            "PPTX 评论",
        );

        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");

        // 1) comment1.xml 存在
        let mut cxml = String::new();
        zip.by_name("ppt/comments/comment1.xml")
            .expect("comment1.xml 应存在")
            .read_to_string(&mut cxml)
            .expect("read comment1.xml");
        assert!(cxml.contains("<p:cmLst"), "comment xml: {}", cxml);
        assert!(cxml.contains("PPTX 评论"), "comment text: {}", cxml);
        assert!(cxml.contains("authorId=\""), "comment authorId: {}", cxml);

        // 2) commentAuthors.xml 存在
        let mut axml = String::new();
        zip.by_name("ppt/commentAuthors.xml")
            .expect("commentAuthors.xml 应存在")
            .read_to_string(&mut axml)
            .expect("read commentAuthors.xml");
        assert!(axml.contains("<p:cmAuthorLst"), "authors xml: {}", axml);
        assert!(axml.contains("张三"), "author name: {}", axml);

        // 3) slide1.xml.rels 包含 comments 关系
        let mut srels = String::new();
        zip.by_name("ppt/slides/_rels/slide1.xml.rels")
            .expect("slide1.xml.rels exists")
            .read_to_string(&mut srels)
            .expect("read slide rels");
        assert!(
            srels.contains("comments"),
            "slide rels should contain comments: {}",
            srels
        );
        assert!(
            srels.contains("../comments/comment1.xml"),
            "slide rels target: {}",
            srels
        );

        // 4) presentation.xml.rels 包含 commentAuthors 关系
        let mut prels = String::new();
        zip.by_name("ppt/_rels/presentation.xml.rels")
            .expect("pres rels exists")
            .read_to_string(&mut prels)
            .expect("read pres rels");
        assert!(prels.contains("commentAuthors"), "pres rels: {}", prels);
    }

    /// read→save→read 往返保持评论。
    #[test]
    fn comments_round_trip_through_pptx() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let author_id = p.comment_authors_mut().get_or_insert_id("李四", "LS");
        let slide = p.slides_mut().add_slide(counter).expect("add slide");
        slide.add_comment(
            author_id,
            crate::Emu::new(100),
            crate::Emu::new(200),
            "往返评论",
        );

        let bytes = p.to_bytes().expect("to_bytes");
        let p2 = Presentation::load_bytes(&bytes).expect("load");

        // 校验评论内容
        let s2 = p2.slides().get(0).expect("slide 0");
        let comments = s2.sld.comments().expect("comments should exist");
        assert_eq!(comments.len(), 1);
        let c = &comments.comments[0];
        assert_eq!(c.text, "往返评论");
        assert_eq!(c.pos_x, 100);
        assert_eq!(c.pos_y, 200);
        assert_eq!(c.idx, 1);

        // 校验作者
        let authors = p2.comment_authors();
        assert_eq!(authors.len(), 1);
        assert_eq!(authors.authors[0].name, "李四");
        assert_eq!(authors.authors[0].initials, "LS");
        assert_eq!(authors.authors[0].id, c.author_id);
    }

    /// 空评论不写出 commentN.xml。
    #[test]
    fn no_comments_no_part() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        p.slides_mut().add_slide(counter).expect("add slide");

        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");
        // 无评论时不应有 comment1.xml
        assert!(zip.by_name("ppt/comments/comment1.xml").is_err());
        // 无评论作者时不应有 commentAuthors.xml
        assert!(zip.by_name("ppt/commentAuthors.xml").is_err());
    }

    /// 多 slide 多评论的 partname 稳定性。
    #[test]
    fn comments_multiple_slides_partname_stable() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let aid = p.comment_authors_mut().get_or_insert_id("A", "A");
        let s1 = p.slides_mut().add_slide(counter.clone()).expect("slide 1");
        s1.add_comment(aid, 0, 0, "slide1 评论");

        let s2 = p.slides_mut().add_slide(counter).expect("slide 2");
        s2.add_comment(aid, 0, 0, "slide2 评论");

        let bytes = p.to_bytes().expect("to_bytes");
        let p2 = Presentation::load_bytes(&bytes).expect("load");

        // 校验两个 slide 都有评论
        let s1b = p2.slides().get(0).expect("slide 0");
        let s2b = p2.slides().get(1).expect("slide 1");
        assert_eq!(s1b.sld.comments().unwrap().comments[0].text, "slide1 评论");
        assert_eq!(s2b.sld.comments().unwrap().comments[0].text, "slide2 评论");

        // 再保存一次，验证 partname 不漂移
        let bytes2 = p2.to_bytes().expect("to_bytes 2");
        let cursor = std::io::Cursor::new(bytes2);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip 2");
        assert!(zip.by_name("ppt/comments/comment1.xml").is_ok());
        assert!(zip.by_name("ppt/comments/comment2.xml").is_ok());
    }

    // ==================== SmartArt round-trip 测试（TODO-037） ====================

    /// `Slide::allocate_diagram_rids` 返回 4 个递增的 rId。
    #[test]
    fn slide_allocate_diagram_rids_increments() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");

        let (d1, l1, q1, c1) = slide.allocate_diagram_rids();
        assert_eq!(d1, "rIdDgmData1");
        assert_eq!(l1, "rIdDgmLayout1");
        assert_eq!(q1, "rIdDgmQs1");
        assert_eq!(c1, "rIdDgmColors1");

        let (d2, l2, q2, c2) = slide.allocate_diagram_rids();
        assert_eq!(d2, "rIdDgmData2");
        assert_eq!(l2, "rIdDgmLayout2");
        assert_eq!(q2, "rIdDgmQs2");
        assert_eq!(c2, "rIdDgmColors2");
    }

    /// `Slide::next_diagram_index` 返回递增索引。
    #[test]
    fn slide_next_diagram_index_increments() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");

        assert_eq!(slide.next_diagram_index(), 1);
        assert_eq!(slide.next_diagram_index(), 2);
        assert_eq!(slide.next_diagram_index(), 3);
    }

    /// `Slide::register_diagram` 把 entry 推入 diagram_entries。
    #[test]
    fn slide_register_diagram_stores_entry() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");

        let entry = DiagramEntry {
            data_partname: crate::opc::part::new_part_name("/ppt/diagrams/data1.xml"),
            layout_partname: crate::opc::part::new_part_name("/ppt/diagrams/layout1.xml"),
            quick_style_partname: crate::opc::part::new_part_name("/ppt/diagrams/quickStyles1.xml"),
            colors_partname: crate::opc::part::new_part_name("/ppt/diagrams/colors1.xml"),
            data_xml: "<dgm:dataModel/>".to_string(),
            layout_xml: "<dgm:layoutDef/>".to_string(),
            quick_style_xml: "<dgm:styleData/>".to_string(),
            colors_xml: "<dgm:colorsDef/>".to_string(),
            data_rid: "rIdDgmData1".to_string(),
            layout_rid: "rIdDgmLayout1".to_string(),
            quick_style_rid: "rIdDgmQs1".to_string(),
            colors_rid: "rIdDgmColors1".to_string(),
        };
        slide.register_diagram(entry);

        // 直接访问 slide.inner.diagram_entries（pub(crate)）验证
        assert_eq!(slide.diagram_entries.len(), 1);
        assert_eq!(
            slide.diagram_entries[0].data_partname.as_str(),
            "/ppt/diagrams/data1.xml"
        );
        assert_eq!(slide.diagram_entries[0].data_xml, "<dgm:dataModel/>");
    }

    /// 端到端：`register_diagram` 后 `to_bytes` 应在 zip 中写出 4 个 diagram parts，
    /// 且 `slide1.xml.rels` 含 4 个 diagram 关系。
    #[test]
    fn diagram_parts_written_to_zip_after_register() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();
        let slide = p.slides_mut().add_slide(counter).expect("add slide");

        let entry = DiagramEntry {
            data_partname: crate::opc::part::new_part_name("/ppt/diagrams/data1.xml"),
            layout_partname: crate::opc::part::new_part_name("/ppt/diagrams/layout1.xml"),
            quick_style_partname: crate::opc::part::new_part_name("/ppt/diagrams/quickStyles1.xml"),
            colors_partname: crate::opc::part::new_part_name("/ppt/diagrams/colors1.xml"),
            data_xml: "<?xml version=\"1.0\"?><dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>".to_string(),
            layout_xml: "<?xml version=\"1.0\"?><dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>".to_string(),
            quick_style_xml: "<?xml version=\"1.0\"?><dgm:styleData xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>".to_string(),
            colors_xml: "<?xml version=\"1.0\"?><dgm:colorsDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>".to_string(),
            data_rid: "rIdDgmData1".to_string(),
            layout_rid: "rIdDgmLayout1".to_string(),
            quick_style_rid: "rIdDgmQs1".to_string(),
            colors_rid: "rIdDgmColors1".to_string(),
        };
        slide.register_diagram(entry);

        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");

        // 校验 4 个 diagram parts 都被写入
        assert!(
            zip.by_name("ppt/diagrams/data1.xml").is_ok(),
            "data1.xml missing"
        );
        assert!(
            zip.by_name("ppt/diagrams/layout1.xml").is_ok(),
            "layout1.xml missing"
        );
        assert!(
            zip.by_name("ppt/diagrams/quickStyles1.xml").is_ok(),
            "quickStyles1.xml missing"
        );
        assert!(
            zip.by_name("ppt/diagrams/colors1.xml").is_ok(),
            "colors1.xml missing"
        );

        // 校验 slide1.xml.rels 中含 4 个 diagram 关系
        let mut srels_xml = String::new();
        zip.by_name("ppt/slides/_rels/slide1.xml.rels")
            .expect("slide1.xml.rels exists")
            .read_to_string(&mut srels_xml)
            .expect("read slide rels");
        assert!(
            srels_xml.contains("/diagramData"),
            "slide1.xml.rels 缺 diagramData 关系，实际：{srels_xml}"
        );
        assert!(
            srels_xml.contains("/diagramLayout"),
            "slide1.xml.rels 缺 diagramLayout 关系，实际：{srels_xml}"
        );
        assert!(
            srels_xml.contains("/diagramQuickStyle"),
            "slide1.xml.rels 缺 diagramQuickStyle 关系，实际：{srels_xml}"
        );
        assert!(
            srels_xml.contains("/diagramColors"),
            "slide1.xml.rels 缺 diagramColors 关系，实际：{srels_xml}"
        );
        // 校验 rId 正确
        assert!(srels_xml.contains("rIdDgmData1"), "缺 rIdDgmData1");
        assert!(srels_xml.contains("rIdDgmLayout1"), "缺 rIdDgmLayout1");
        assert!(srels_xml.contains("rIdDgmQs1"), "缺 rIdDgmQs1");
        assert!(srels_xml.contains("rIdDgmColors1"), "缺 rIdDgmColors1");
    }

    /// 端到端：多 slide + 多 SmartArt 时，diagram partname 全局递增不冲突。
    #[test]
    fn diagram_parts_global_index_across_slides() {
        let mut p = Presentation::new().expect("new ok");
        let counter = p.id_counter();

        // slide 1：1 个 SmartArt
        let s1 = p.slides_mut().add_slide(counter.clone()).expect("s1");
        let e1 = DiagramEntry {
            data_partname: crate::opc::part::new_part_name("/ppt/diagrams/data1.xml"),
            layout_partname: crate::opc::part::new_part_name("/ppt/diagrams/layout1.xml"),
            quick_style_partname: crate::opc::part::new_part_name("/ppt/diagrams/quickStyles1.xml"),
            colors_partname: crate::opc::part::new_part_name("/ppt/diagrams/colors1.xml"),
            data_xml: "<dgm:dataModel/>".to_string(),
            layout_xml: "<dgm:layoutDef/>".to_string(),
            quick_style_xml: "<dgm:styleData/>".to_string(),
            colors_xml: "<dgm:colorsDef/>".to_string(),
            data_rid: "rIdDgmData1".to_string(),
            layout_rid: "rIdDgmLayout1".to_string(),
            quick_style_rid: "rIdDgmQs1".to_string(),
            colors_rid: "rIdDgmColors1".to_string(),
        };
        s1.register_diagram(e1);

        // slide 2：1 个 SmartArt（局部 rid 与 slide 1 重复，但全局 partname 应不同）
        let s2 = p.slides_mut().add_slide(counter).expect("s2");
        let e2 = DiagramEntry {
            data_partname: crate::opc::part::new_part_name("/ppt/diagrams/data1.xml"),
            layout_partname: crate::opc::part::new_part_name("/ppt/diagrams/layout1.xml"),
            quick_style_partname: crate::opc::part::new_part_name("/ppt/diagrams/quickStyles1.xml"),
            colors_partname: crate::opc::part::new_part_name("/ppt/diagrams/colors1.xml"),
            data_xml: "<dgm:dataModel/>".to_string(),
            layout_xml: "<dgm:layoutDef/>".to_string(),
            quick_style_xml: "<dgm:styleData/>".to_string(),
            colors_xml: "<dgm:colorsDef/>".to_string(),
            data_rid: "rIdDgmData1".to_string(),
            layout_rid: "rIdDgmLayout1".to_string(),
            quick_style_rid: "rIdDgmQs1".to_string(),
            colors_rid: "rIdDgmColors1".to_string(),
        };
        s2.register_diagram(e2);

        let bytes = p.to_bytes().expect("to_bytes");
        let cursor = std::io::Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(cursor).expect("zip open");

        // slide 1 的 SmartArt 应使用 data1.xml / layout1.xml / ...
        assert!(zip.by_name("ppt/diagrams/data1.xml").is_ok());
        assert!(zip.by_name("ppt/diagrams/layout1.xml").is_ok());
        // slide 2 的 SmartArt 应使用全局递增后的 data2.xml / layout2.xml / ...
        // （to_opc_package 用 diagram_global_index 重新分配 partname）
        assert!(
            zip.by_name("ppt/diagrams/data2.xml").is_ok(),
            "data2.xml missing"
        );
        assert!(
            zip.by_name("ppt/diagrams/layout2.xml").is_ok(),
            "layout2.xml missing"
        );
        assert!(zip.by_name("ppt/diagrams/quickStyles2.xml").is_ok());
        assert!(zip.by_name("ppt/diagrams/colors2.xml").is_ok());
    }
}
