//! `OpcPackage`：包加载/保存。
//!
//! 本文件是 OPC 容器层的入口。它把"加载/保存一个 `.pptx`"和"注册 part +
//! 维护 Content-Types + 维护关系链"三个动作统一到 [`OpcPackage`] 一个类型上。
//!
//! # 加载流程
//!
//! 1. 打开文件 → `zip::ZipArchive`；
//! 2. 读取 `[Content_Types].xml` → `ContentTypes` 模型；
//! 3. 遍历 zip 全部条目，把每个 part 装入 `parts: BTreeMap<PartName, Part>`；
//! 4. 关系文件 `.rels` 同样作为 part 装入（content-type 固定为
//!    `application/vnd...relationships+xml`）。
//!
//! # 保存流程
//!
//! 1. 创建 `zip::ZipWriter`；
//! 2. 写 `[Content_Types].xml`（序列化 [`ContentTypes`]）；
//! 3. 按 partname 顺序遍历 `parts`，逐个写入 zip；
//! 4. 关闭 zip。
//!
//! # 设计取舍
//!
//! - **不维护反向关系索引**：本结构只持有"由 partname 找 part"的正向映射；
//!   反向"由 partname 找关系链"由调用方在 `Relationships` 上自行迭代。
//! - **`BTreeMap` 而非 `HashMap`**：让 `iter_parts()` 总是按 partname 字典序
//!   输出，便于 byte-diff 测试与稳定写入顺序。
//! - **不解析 `.rels`**：关系文件以原始 blob 形式保留，调用方在需要时
//!   调用 `Relationships::from_xml` 显式解析。

use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;

use super::content_types::{ContentTypes, DefaultExt};
use super::part::{Part, PartName};
use super::rels::Relationship;

/// 标准 Content-Type 常量集合。
///
/// 这些常量与 ECMA-376 / OOXML 规范一一对应，使用 `pub const` 暴露
/// 便于在序列化时直接引用而无需硬编码字符串。
pub mod ct {
    /// 根关系 Content-Type：`application/vnd.openxmlformats-package.relationships+xml`。
    pub const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
    /// 主演示文稿：`.../presentationml.presentation.main+xml`。
    pub const PRESENTATION: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    /// 幻灯片：`.../presentationml.slide+xml`。
    pub const SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
    /// 幻灯片布局：`.../presentationml.slideLayout+xml`。
    pub const SLIDE_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
    /// 幻灯片母版：`.../presentationml.slideMaster+xml`。
    pub const SLIDE_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
    /// 主题：`.../theme+xml`。
    pub const THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
    /// 核心属性（`cp:coreProperties`）。
    pub const CORE_PROPS: &str = "application/vnd.openxmlformats-package.core-properties+xml";
    /// 应用属性（`extended-properties`）。
    pub const APP_PROPS: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    /// 自定义属性（`custom-properties`，`/docProps/custom.xml`）。
    ///
    /// 对应 OOXML 中 `Properties` 根元素，命名空间为
    /// `http://schemas.openxmlformats.org/officeDocument/2006/custom-properties`。
    pub const CUSTOM_PROPS: &str =
        "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    /// 备注页：`.../presentationml.notesSlide+xml`。
    ///
    /// python-pptx 中由 `Slide.notes_slide` 隐式管理。
    pub const NOTES_SLIDE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
    /// 备注母版：`.../presentationml.notesMaster+xml`（`/ppt/notesMasters/notesMasterN.xml`，TODO-045）。
    ///
    /// 对应 OOXML 中 `<p:notesMaster>` 根元素，被 `presentation.xml`
    /// 的 `<p:notesMasterIdLst>` 引用。
    pub const NOTES_MASTER: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
    /// 幻灯片评论：`.../presentationml.comments+xml`（`/ppt/comments/commentN.xml`）。
    ///
    /// 对应 OOXML 中 `<p:cmLst>` 根元素，命名空间为
    /// `http://schemas.openxmlformats.org/presentationml/2006/main`。
    pub const COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
    /// 评论作者列表：`.../presentationml.commentAuthors+xml`（`/ppt/commentAuthors.xml`）。
    ///
    /// 全局共享的作者清单，对应 `<p:cmAuthorLst>` 根元素。
    pub const COMMENT_AUTHORS: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
    /// 图表：`.../drawingml.chart+xml`（`/ppt/charts/chartN.xml`）。
    ///
    /// 对应 OOXML 中 `<c:chartSpace>` 根元素，命名空间为
    /// `http://schemas.openxmlformats.org/drawingml/2006/chart`。
    pub const CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    /// OLE 对象：`.../oleobject`（`/ppt/embeddings/oleObjectN.bin`，TODO-043）。
    ///
    /// 对应二进制 OLE 复合文档（CFB），被 slide 的 `<p:oleObj r:id="..."/>` 引用。
    /// Content-Type 固定为 `application/vnd.openxmlformats-officedocument.oleobject`，
    /// 与文件扩展名无关（PowerPoint 通过 progId 区分具体 OLE 服务器）。
    pub const OLE_OBJECT: &str = "application/vnd.openxmlformats-officedocument.oleobject";
    /// 嵌入式 Excel 工作簿：`.../spreadsheetml.sheet`（`/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`，TODO-004 Excel 嵌入）。
    ///
    /// 对应 ECMA-376 SpreadsheetML 包格式，被 chart 的
    /// `<c:externalData r:id="..."/>` 引用。PowerPoint 打开图表时
    /// 会从该 xlsx part 读取数据源（"编辑数据" 会启动 Excel）。
    pub const SPREADSHEET_XLSX: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    /// 视频：`video/mp4`（`/ppt/media/mediaN.mp4`，TODO-033）。
    ///
    /// MP4 是 PowerPoint 最通用的视频格式。其它视频格式（wmv/avi）使用各自 MIME。
    pub const VIDEO_MP4: &str = "video/mp4";
    /// 音频：`audio/mp3`（`/ppt/media/mediaN.mp3`，TODO-033）。
    ///
    /// MP3 是 PowerPoint 最通用的音频格式。其它音频格式（wav/aac）使用各自 MIME。
    pub const AUDIO_MP3: &str = "audio/mp3";
    /// SmartArt 数据：`.../drawingml.diagramData+xml`（`/ppt/diagrams/dataN.xml`，TODO-037）。
    ///
    /// 对应 OOXML 中 `<dgm:dataModel>` 根元素，命名空间为
    /// `http://schemas.openxmlformats.org/drawingml/2006/diagram`。
    /// 被 slide 的 `<p:graphicFrame>` 内 `<dgm:relIds r:dm="..."/>` 引用。
    pub const DIAGRAM_DATA: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
    /// SmartArt 布局：`.../drawingml.diagramLayout+xml`（`/ppt/diagrams/layoutN.xml`，TODO-037）。
    ///
    /// 对应 OOXML 中 `<dgm:layoutDef>` 根元素，定义节点排列算法与拓扑。
    pub const DIAGRAM_LAYOUT: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
    /// SmartArt 快速样式：`.../drawingml.diagramQuickStyle+xml`（`/ppt/diagrams/quickStylesN.xml`，TODO-037）。
    ///
    /// 对应 OOXML 中 `<dgm:styleData>` 根元素，定义节点/连接线视觉样式。
    pub const DIAGRAM_QUICK_STYLE: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml";
    /// SmartArt 颜色：`.../drawingml.diagramColors+xml`（`/ppt/diagrams/colorsN.xml`，TODO-037）。
    ///
    /// 对应 OOXML 中 `<dgm:colorsDef>` 根元素，定义颜色变体映射。
    pub const DIAGRAM_COLORS: &str =
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml";
}

/// 一个 OPC 包。
///
/// 内部包含：
/// - `parts`：所有 part 的字典（key 为 partname 字符串）；
/// - `content_types`：根 `[Content_Types].xml` 的强类型模型。
#[derive(Debug, Clone, Default)]
pub struct OpcPackage {
    /// 所有 part（partname → Part），按 partname 字典序排列。
    pub parts: BTreeMap<String, Part>,
    /// ContentTypes 模型。
    pub content_types: ContentTypes,
}

impl OpcPackage {
    /// 构造空包（含默认 ContentTypes）。
    pub fn new() -> Self {
        OpcPackage {
            parts: BTreeMap::new(),
            content_types: ContentTypes::new_default(),
        }
    }

    /// 放入一个 part（覆盖同名）。
    ///
    /// 在插入前会调用 [`Part::contribute_to`]，由 part 自己决定如何
    /// 更新 Content-Types（添加 override 等）。
    pub fn put_part(&mut self, p: Part) {
        p.contribute_to(&mut self.content_types);
        self.parts.insert(p.partname.as_str().to_string(), p);
    }

    /// 取一个 part 的不可变引用。
    pub fn get_part(&self, partname: &str) -> Option<&Part> {
        self.parts.get(partname)
    }

    /// 取一个 part 的可变引用。
    pub fn get_part_mut(&mut self, partname: &str) -> Option<&mut Part> {
        self.parts.get_mut(partname)
    }

    /// 列出所有 part。
    pub fn iter_parts(&self) -> impl Iterator<Item = &Part> {
        self.parts.values()
    }

    /// part 数量。
    pub fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// 加载 `.pptx`（zip 文件）。
    ///
    /// # 流程
    /// 1. 读取 `[Content_Types].xml` 并解析；
    /// 2. 遍历 zip 全部条目，把非目录、非 `[Content_Types].xml` 的条目
    ///    装入 `parts`。
    ///
    /// # 错误
    /// - `Error::Io`：文件读取失败；
    /// - `Error::Zip`：zip 解压失败；
    /// - `Error::Xml`：`[Content_Types].xml` 解析失败。
    pub fn load(path: impl AsRef<Path>) -> crate::Result<Self> {
        let file = File::open(path.as_ref())?;
        let mut zip = zip::ZipArchive::new(file)?;
        let mut pkg = OpcPackage::new();

        // 1) 先读 [Content_Types].xml
        let mut ct_xml = String::new();
        zip.by_name("[Content_Types].xml")?
            .read_to_string(&mut ct_xml)?;
        pkg.content_types = parse_content_types_public(&ct_xml)?;

        // 2) 读所有 part
        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_string();
            if name == "[Content_Types].xml" {
                continue;
            }
            if entry.is_dir() {
                continue;
            }

            // 关系文件 (单独维护) 仍以 part 形式加入，但 content_type 固定
            let mut blob = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut blob)?;
            let ct = if name.ends_with(".rels") {
                ct::RELATIONSHIPS.to_string()
            } else {
                derive_content_type(&pkg.content_types, &format!("/{}", name))
            };
            let partname = format!("/{}", name);
            let part = Part::new(PartName::from_unchecked(partname), ct, blob);
            pkg.parts.insert(part.partname.as_str().to_string(), part);
        }
        Ok(pkg)
    }

    /// 保存为 `.pptx`。
    ///
    /// 写入顺序：先 `[Content_Types].xml`，再按字典序遍历所有 part。
    /// 所有条目使用 `deflate` 压缩 + `0o644` Unix 权限。
    pub fn save(&self, path: impl AsRef<Path>) -> crate::Result<()> {
        let file = File::create(path.as_ref())?;
        let mut zip = zip::ZipWriter::new(file);
        let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o644);

        // 1) [Content_Types].xml
        zip.start_file("[Content_Types].xml", opts)?;
        zip.write_all(self.content_types.to_xml().as_bytes())?;

        // 2) 其余 part
        for part in self.parts.values() {
            let p = part.partname.to_zip_path();
            zip.start_file(p, opts)?;
            zip.write_all(&part.blob)?;
        }
        zip.finish()?;
        Ok(())
    }

    /// 写回内存 zip（用于测试）。
    pub fn to_bytes(&self) -> crate::Result<Vec<u8>> {
        use std::io::Cursor;
        let mut buf: Vec<u8> = Vec::new();
        {
            let cursor = Cursor::new(&mut buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts: zip::write::FileOptions<'_, ()> = zip::write::FileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            zip.start_file("[Content_Types].xml", opts)?;
            zip.write_all(self.content_types.to_xml().as_bytes())?;
            for part in self.parts.values() {
                zip.start_file(part.partname.to_zip_path(), opts)?;
                zip.write_all(&part.blob)?;
            }
            zip.finish()?;
        }
        Ok(buf)
    }
}

/// 在 ContentTypes 中根据 partname 查找 Content-Type（override 优先）。
///
/// 顺序：先查 `overrides`（按 partname 完全匹配），再查 `defaults`（按扩展名）。
/// 全部未命中则回退到 `application/octet-stream`。
///
/// # 用途
/// - 在 [`OpcPackage::load`] / `Presentation::load_bytes` 之后，
///   对新加入 `parts` 的非 `.rels` 元素推断 content_type；
/// - 在 `to_opc_package` 中也可以用，但通常**不需要**——写路径下
///   `Part::contribute_to` 会主动调用 [`Part`] 的 `put_part` 自动写 override。
pub fn derive_content_type(ct: &ContentTypes, partname: &str) -> String {
    for o in &ct.overrides {
        if o.partname == partname {
            return o.content_type.clone();
        }
    }
    // 用扩展名查 defaults
    let ext = match partname.rfind('.') {
        Some(i) => &partname[i + 1..],
        None => "",
    };
    for d in &ct.defaults {
        if d.extension == ext {
            return d.content_type.clone();
        }
    }
    "application/octet-stream".to_string()
}

/// 极简解析 `[Content_Types].xml`，主要提取 override。
///
/// 完整 schema 还支持 `<Default Extension="..." ContentType="..."/>` 与
/// `<Override PartName="..." ContentType="..."/>` 两种元素，本函数
/// 同时识别两者。
pub fn parse_content_types_public(xml: &str) -> crate::Result<ContentTypes> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;
    let mut ct = ContentTypes::new_default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                match name.as_ref() {
                    b"Default" => {
                        let mut ext = None;
                        let mut ctype = None;
                        for a in e.attributes().flatten() {
                            match a.key.as_ref() {
                                b"Extension" => {
                                    ext = a
                                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .ok()
                                        .map(|v| v.to_string())
                                }
                                b"ContentType" => {
                                    ctype = a
                                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .ok()
                                        .map(|v| v.to_string())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(ext), Some(ctype)) = (ext, ctype) {
                            ct.defaults.push(DefaultExt::new(ext, ctype));
                        }
                    }
                    b"Override" => {
                        let mut partname = None;
                        let mut ctype = None;
                        for a in e.attributes().flatten() {
                            match a.key.as_ref() {
                                b"PartName" => {
                                    partname = a
                                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .ok()
                                        .map(|v| v.to_string())
                                }
                                b"ContentType" => {
                                    ctype = a
                                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .ok()
                                        .map(|v| v.to_string())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(p), Some(c)) = (partname, ctype) {
                            ct.add_override(&p, &c);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("content-types parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(ct)
}

/// 工具：把一个 PartName 映射成 zip 内的"关系文件"路径。
///
/// 严格遵循 OPC 约定：`/ppt/slides/slide1.xml` 的关系文件应在
/// `/ppt/slides/_rels/slide1.xml.rels`。
///
/// # 示例
/// - `/ppt/slides/slide1.xml` → `/ppt/slides/_rels/slide1.xml.rels`
/// - `/word/document.xml` → `/word/_rels/document.xml.rels`
/// - `/foo` → `/_rels/foo.rels`
pub fn rels_partname_for(partname: &str) -> String {
    let p = partname.trim_start_matches('/');
    let last_slash = p.rfind('/');
    match last_slash {
        None => format!("/_rels/{}.rels", p),
        // 父目录 _rels 文件名.rels
        Some(i) => format!("/{}/_rels/{}.rels", &p[..i], &p[i + 1..]),
    }
}

/// 工具：建立"父 part → (id, reltype, target)"链；返回每个 target 的下一个 rId。
///
/// 通过扫描已有关系 ID（如 `rId1` / `rId2` ...）找到最大编号，分配 `max+1`。
/// prefix 必须以字符串形式给出，常用 `rId`。
pub fn next_rid(existing: &[Relationship], prefix: &str) -> String {
    let mut max = 0u32;
    for r in existing {
        if let Some(n) = r.id.strip_prefix(prefix) {
            if let Ok(v) = n.parse::<u32>() {
                if v > max {
                    max = v;
                }
            }
        }
    }
    format!("{}{}", prefix, max + 1)
}
