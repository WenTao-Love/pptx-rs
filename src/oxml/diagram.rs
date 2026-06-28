//! SmartArt（Diagram）OOXML 模型 —— 结构化解析（TODO-037）。
//!
//! 本模块为 SmartArt 图形的 4 个 part（`/ppt/diagrams/*.xml`）提供**结构化解析**能力：
//!
//! | part | 根元素 | 结构体 | 解析深度 |
//! |---|---|---|---|
//! | data | `<dgm:dataModel>` | [`DataModel`] | **完全结构化**（节点 + 连接 + 文本） |
//! | layout | `<dgm:layoutDef>` | [`LayoutDef`] | 半结构化（元数据 + layoutNode 原始 XML） |
//! | quickStyle | `<dgm:styleData>` | [`QuickStyleDef`] | 半结构化（styleLbl 列表 + 原始 XML） |
//! | colors | `<dgm:colorsDef>` | [`ColorsDef`] | 半结构化（元数据 + styleClrLbl 列表） |
//!
//! # 设计要点
//!
//! - **按需解析（lazy parsing）**：[`crate::presentation::DiagramEntry`] 仍以 `String` blob
//!   持有原始 XML，保证 byte-exact round-trip；调用方需要结构化访问时，通过
//!   `DiagramEntry::data_model()` / `layout_def()` 等方法按需触发解析。
//! - **零 panic 设计**：解析失败返回 `Error::Xml`，不阻塞 round-trip。
//! - **不引入 layoutNode 算法树**：layoutNode 拓扑结构复杂（含 alg/forEach 嵌套），
//!   本模块仅保留 layoutNode 整段子树的原始 XML，不展开为强类型树。
//!
//! # 与 python-pptx 的对应
//!
//! python-pptx **不支持** SmartArt（参见 `docs/CHANGELOG.md`），因此本模块是
//! pptx-rs 相对 python-pptx 的扩展能力。
//!
//! # OOXML 命名空间
//!
//! 所有 4 个 part 都使用 `http://schemas.openxmlformats.org/drawingml/2006/diagram`
//! 命名空间（前缀 `dgm`），文本元素使用 `a:` 前缀（DrawingML main）。

use crate::oxml::writer::XmlWriter;

// ============================================================================
// data part: <dgm:dataModel>
// ============================================================================

/// SmartArt 数据模型（data part 的结构化表达）。
///
/// 对应 `/ppt/diagrams/dataN.xml` 的 `<dgm:dataModel>` 根元素。
/// 一个 SmartArt 图形的所有节点（pt）与连接（cxn）都在这里——
/// 这是 SmartArt 的**核心数据**，编辑/查询节点都基于此结构。
///
/// # 结构示例
///
/// ```xml
/// <dgm:dataModel xmlns:dgm="..." xmlns:a="...">
///   <dgm:ptLst>
///     <dgm:pt modelId="0" type="doc"/>
///     <dgm:pt modelId="1" type="par">
///       <dgm:prSet ang="0"/>
///       <dgm:spPr/>
///       <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>根节点</a:t></a:r></a:p></dgm:t>
///     </dgm:pt>
///   </dgm:ptLst>
///   <dgm:cxnLst>
///     <dgm:cxn type="parChld" srcId="1" destId="2"/>
///   </dgm:cxnLst>
/// </dgm:dataModel>
/// ```
#[derive(Clone, Debug, Default)]
pub struct DataModel {
    /// 节点列表（`<dgm:ptLst>/<dgm:pt>`）。
    pub points: Vec<DataModelPoint>,
    /// 连接列表（`<dgm:cxnLst>/<dgm:cxn>`）。
    pub connections: Vec<DataModelConnection>,
}

/// 单个 SmartArt 节点（`<dgm:pt>`）。
///
/// 每个 pt 都有唯一的 `model_id`，通过 `<dgm:cxn>` 建立父子/兄弟关系。
#[derive(Clone, Debug, Default)]
pub struct DataModelPoint {
    /// 节点 ID（`modelId` 属性，必需，同一 dataModel 内唯一）。
    pub model_id: u32,
    /// 节点类型（`type` 属性）。
    ///
    /// OOXML 定义取值：`doc`（文档根）/ `par`（父节点）/ `ch`（子节点）/
    /// `sib`（兄弟节点）/ `prev`（前驱）/ `next`（后继）。
    pub pt_type: Option<String>,
    /// 节点显示文本（从 `<dgm:t>/<a:p>/<a:r>/<a:t>` 提取的首个文本）。
    ///
    /// 若节点无文本（如 `type="doc"` 的虚拟根），此字段为 `None`。
    pub text: Option<String>,
    /// 节点原始 XML（保留 byte-exact，用于 round-trip 兜底）。
    pub raw_xml: String,
}

/// SmartArt 节点连接（`<dgm:cxn>`）。
///
/// 描述两个节点的拓扑关系（父子 / 兄弟）。
#[derive(Clone, Debug, Default)]
pub struct DataModelConnection {
    /// 源节点 ID（`srcId` 属性，必需）。
    pub src_id: u32,
    /// 目标节点 ID（`destId` 属性，必需）。
    pub dest_id: u32,
    /// 连接类型（`type` 属性）。
    ///
    /// OOXML 定义取值：`parChld`（父子）/ `sib`（兄弟）等。
    pub cxn_type: Option<String>,
    /// 连接原始 XML（保留 byte-exact）。
    pub raw_xml: String,
}

impl DataModelPoint {
    /// 设置节点显示文本，同步更新 `text` 字段与 `raw_xml` 中的 `<a:t>` 内容。
    ///
    /// # 工作原理
    ///
    /// 1. 在 `raw_xml` 中查找 `<a:t>...</a:t>` 元素；
    /// 2. 若找到，替换其中的文本内容为 `new_text`（自动 XML 转义）；
    /// 3. 若未找到（如 `type="doc"` 虚拟根节点无 `<dgm:t>` 子元素），仅更新 `text` 字段，
    ///    `raw_xml` 保持不变（虚拟根通常不需要显示文本）。
    ///
    /// # 参数
    /// - `new_text`：新的节点文本内容。
    ///
    /// # 限制
    ///
    /// - 仅替换第一个 `<a:t>` 元素的内容（SmartArt 节点通常只有一个文本运行）；
    /// - 若 `raw_xml` 为空（用户新建的节点），调用本方法后还需要调用
    ///   [`DataModel::to_xml`] 的结构化重建分支才能生成完整 XML。
    pub fn set_text(&mut self, new_text: impl Into<String>) {
        let new_text = new_text.into();
        let escaped = escape_xml_text(&new_text);
        // 在 raw_xml 中查找 <a:t>...</a:t> 并替换内容
        if let Some(start) = self.raw_xml.find("<a:t>") {
            // 查找对应的 </a:t> 闭合标签
            if let Some(end_rel) = self.raw_xml[start..].find("</a:t>") {
                let content_start = start + "<a:t>".len();
                let content_end = start + end_rel;
                let prefix = &self.raw_xml[..content_start];
                let suffix = &self.raw_xml[content_end..];
                self.raw_xml = format!("{}{}{}", prefix, escaped, suffix);
            }
        }
        self.text = Some(new_text);
    }

    /// 清除节点显示文本。
    ///
    /// 同步更新 `text` 字段为 `None` 与 `raw_xml` 中的 `<a:t>` 内容为空字符串。
    /// 若 `raw_xml` 中无 `<a:t>` 元素，仅更新 `text` 字段。
    pub fn clear_text(&mut self) {
        // 在 raw_xml 中查找 <a:t>...</a:t> 并清空内容
        if let Some(start) = self.raw_xml.find("<a:t>") {
            if let Some(end_rel) = self.raw_xml[start..].find("</a:t>") {
                let content_start = start + "<a:t>".len();
                let content_end = start + end_rel;
                let prefix = &self.raw_xml[..content_start];
                let suffix = &self.raw_xml[content_end..];
                self.raw_xml = format!("{}{}", prefix, suffix);
            }
        }
        self.text = None;
    }

    /// 按 model_id 查询节点是否为指定类型。
    ///
    /// 便捷方法：避免外部重复匹配 `pt_type` 字段。
    pub fn is_type(&self, type_str: &str) -> bool {
        self.pt_type.as_deref() == Some(type_str)
    }
}

/// XML 文本转义：把 `&` / `<` / `>` / `'` / `"` 替换为实体引用。
///
/// 用于 `set_text` 写入 `<a:t>` 内容时确保 XML 合法。
fn escape_xml_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '\'' => out.push_str("&apos;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(c),
        }
    }
    out
}

impl DataModel {
    /// 从 `<dgm:dataModel>` XML 字符串解析为强类型 [`DataModel`]。
    ///
    /// # 参数
    /// - `xml`：完整 dataN.xml 内容（含 `<?xml?>` 声明与 `<dgm:dataModel>` 根元素）。
    ///
    /// # 返回值
    /// - 成功：返回 [`DataModel`]，包含所有节点与连接。
    /// - 失败：返回 `Error::Xml`，包含解析错误上下文。
    ///
    /// # 解析策略
    ///
    /// 使用 quick-xml SAX 事件流解析：
    /// - 进入 `<dgm:pt>` 时收集属性（`modelId` / `type`），开始累积 raw_xml；
    /// - 在 `<dgm:pt>` 内的 `<a:t>` 文本事件提取节点文本；
    /// - 离开 `<dgm:pt>` 时把累积的 raw_xml 切片写入 `point.raw_xml`；
    /// - `<dgm:cxn>` 类似处理。
    ///
    /// # 错误
    /// - `Error::Xml`：XML 解析失败（畸形 / 未闭合等）。
    pub fn parse_from_xml(xml: &str) -> crate::Result<DataModel> {
        let _ = xml; // xml 仅用于 Reader::from_str，不直接切片（避免 buffer_position 语义变化）
        let mut points: Vec<DataModelPoint> = Vec::new();
        let mut connections: Vec<DataModelConnection> = Vec::new();

        let mut rd = quick_xml::reader::Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();

        // 当前正在解析的 pt/cxn 缓冲
        let mut cur_pt: Option<DataModelPoint> = None;
        let mut cur_cxn: Option<DataModelConnection> = None;
        // 当前所在元素深度（用于判断是否累积 raw_xml）
        let mut pt_depth: i32 = 0;
        let mut cxn_depth: i32 = 0;
        let mut in_pt_text = false; // 在 <a:t> 内
                                    // 手动累积 raw_xml（避免依赖 buffer_position 的语义变化）
        let mut cur_pt_raw = String::new();
        let mut cur_cxn_raw = String::new();
        // 累积当前 pt 的文本片段
        let mut cur_text_buf = String::new();

        loop {
            match rd.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"pt" && pt_depth == 0 {
                        // 进入顶层 <dgm:pt>
                        pt_depth = 1;
                        let mut pt = DataModelPoint::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"modelId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    pt.model_id = n;
                                }
                            } else if key == b"type" {
                                pt.pt_type = Some(val);
                            }
                        }
                        cur_pt = Some(pt);
                        cur_text_buf.clear();
                        // 开始累积 raw_xml
                        cur_pt_raw.clear();
                        cur_pt_raw.push('<');
                        cur_pt_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_pt_raw.push('>');
                    } else if local == b"cxn" && cxn_depth == 0 {
                        // 进入顶层 <dgm:cxn>
                        cxn_depth = 1;
                        let mut cxn = DataModelConnection::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"srcId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    cxn.src_id = n;
                                }
                            } else if key == b"destId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    cxn.dest_id = n;
                                }
                            } else if key == b"type" {
                                cxn.cxn_type = Some(val);
                            }
                        }
                        cur_cxn = Some(cxn);
                        cur_cxn_raw.clear();
                        cur_cxn_raw.push('<');
                        cur_cxn_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_cxn_raw.push('>');
                    } else {
                        // pt/cxn 内部的子元素 Start：累积到 raw
                        if pt_depth > 0 {
                            cur_pt_raw.push('<');
                            cur_pt_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_pt_raw.push('>');
                            if local == b"t" {
                                in_pt_text = true;
                            }
                            // 嵌套同名元素也增加深度（防止提前结束）
                            if local == b"pt" {
                                pt_depth += 1;
                            }
                        } else if cxn_depth > 0 {
                            cur_cxn_raw.push('<');
                            cur_cxn_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_cxn_raw.push('>');
                            if local == b"cxn" {
                                cxn_depth += 1;
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"pt" && pt_depth == 0 {
                        // 顶层自闭合 <dgm:pt ... />（无子元素，常见于 type="doc" 虚拟根）
                        let mut pt = DataModelPoint::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"modelId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    pt.model_id = n;
                                }
                            } else if key == b"type" {
                                pt.pt_type = Some(val);
                            }
                        }
                        // 自闭合元素的 raw_xml 就是 <tag attrs/>
                        let mut raw = String::from("<");
                        raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        raw.push_str("/>");
                        pt.raw_xml = raw;
                        points.push(pt);
                    } else if local == b"cxn" && cxn_depth == 0 {
                        // 顶层自闭合 <dgm:cxn ... />
                        let mut cxn = DataModelConnection::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"srcId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    cxn.src_id = n;
                                }
                            } else if key == b"destId" {
                                if let Ok(n) = val.parse::<u32>() {
                                    cxn.dest_id = n;
                                }
                            } else if key == b"type" {
                                cxn.cxn_type = Some(val);
                            }
                        }
                        let mut raw = String::from("<");
                        raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        raw.push_str("/>");
                        cxn.raw_xml = raw;
                        connections.push(cxn);
                    } else if pt_depth > 0 {
                        // pt 内部自闭合子元素
                        cur_pt_raw.push('<');
                        cur_pt_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_pt_raw.push_str("/>");
                    } else if cxn_depth > 0 {
                        // cxn 内部自闭合子元素
                        cur_cxn_raw.push('<');
                        cur_cxn_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_cxn_raw.push_str("/>");
                    }
                }
                Ok(quick_xml::events::Event::Text(t)) => {
                    if in_pt_text && pt_depth > 0 {
                        // quick-xml 0.40: BytesText::unescape() 方法已移除，
                        // 改用 quick_xml::escape::unescape 函数（接受 &str）。
                        // BytesText 的 Deref 目标是 [u8]，需要先转成 &str。
                        let text_str = std::str::from_utf8(t.as_ref()).unwrap_or("");
                        let text = quick_xml::escape::unescape(text_str)
                            .unwrap_or_default()
                            .to_string();
                        if !text.is_empty() {
                            cur_text_buf.push_str(&text);
                        }
                    }
                    // 同时累积文本到 raw_xml
                    if pt_depth > 0 {
                        cur_pt_raw.push_str(std::str::from_utf8(t.as_ref()).unwrap_or(""));
                    } else if cxn_depth > 0 {
                        cur_cxn_raw.push_str(std::str::from_utf8(t.as_ref()).unwrap_or(""));
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if pt_depth > 0 {
                        // 累积 End 事件到 pt raw
                        cur_pt_raw.push_str("</");
                        cur_pt_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_pt_raw.push('>');
                        if local == b"t" {
                            in_pt_text = false;
                        } else if local == b"pt" {
                            pt_depth -= 1;
                            if pt_depth == 0 {
                                // 离开顶层 pt：写回 raw_xml + text
                                if let Some(mut pt) = cur_pt.take() {
                                    pt.raw_xml = std::mem::take(&mut cur_pt_raw);
                                    if !cur_text_buf.is_empty() {
                                        pt.text = Some(cur_text_buf.clone());
                                    }
                                    points.push(pt);
                                }
                                cur_text_buf.clear();
                            }
                        }
                    } else if cxn_depth > 0 {
                        cur_cxn_raw.push_str("</");
                        cur_cxn_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_cxn_raw.push('>');
                        if local == b"cxn" {
                            cxn_depth -= 1;
                            if cxn_depth == 0 {
                                if let Some(mut cxn) = cur_cxn.take() {
                                    cxn.raw_xml = std::mem::take(&mut cur_cxn_raw);
                                    connections.push(cxn);
                                }
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(crate::Error::Xml(format!("DataModel parse_from_xml: {e}"))),
                _ => {}
            }
            buf.clear();
        }

        Ok(DataModel {
            points,
            connections,
        })
    }

    /// 把 [`DataModel`] 序列化为 `<dgm:dataModel>` XML 字符串。
    ///
    /// **注意**：本方法用于"从结构化模型重建 XML"场景；如果只是 round-trip，
    /// 应直接使用 [`crate::presentation::DiagramEntry`] 持有的 `data_xml` 字段。
    ///
    /// # 输出结构
    ///
    /// ```xml
    /// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    /// <dgm:dataModel xmlns:dgm="..." xmlns:a="...">
    ///   <dgm:ptLst>
    ///     <!-- 每个 point 的 raw_xml 直接透传（保留 byte-exact） -->
    ///   </dgm:ptLst>
    ///   <dgm:cxnLst>
    ///     <!-- 每个 connection 的 raw_xml 直接透传 -->
    ///   </dgm:cxnLst>
    /// </dgm:dataModel>
    /// ```
    ///
    /// # 设计要点
    ///
    /// 节点/连接的子元素（`<dgm:prSet>` / `<dgm:spPr>` / `<dgm:t>` 等）通过
    /// `raw_xml` 字段直接透传，避免重新序列化时丢失属性。
    ///
    /// # 结构化重建分支
    ///
    /// 当 `raw_xml` 为空但 `text` 字段非空时（用户新建的节点），按 OOXML 顺序
    /// 构造完整 `<dgm:pt>` 结构：
    ///
    /// ```xml
    /// <dgm:pt modelId="..." type="...">
    ///   <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>文本</a:t></a:r></a:p></dgm:t>
    /// </dgm:pt>
    /// ```
    ///
    /// 这是为了支持"从零构造 SmartArt 数据模型并写出"的场景。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::new();
        w.raw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        w.open_with(
            "dgm:dataModel",
            &[
                (
                    "xmlns:dgm",
                    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                ),
                (
                    "xmlns:a",
                    "http://schemas.openxmlformats.org/drawingml/2006/main",
                ),
            ],
        );
        // ptLst
        w.open("dgm:ptLst");
        for pt in &self.points {
            if pt.raw_xml.is_empty() {
                // 无 raw_xml（用户新建）：构造 <dgm:pt>，若有 text 则包含 <dgm:t> 子元素
                let id_s = pt.model_id.to_string();
                let mut attrs: Vec<(&str, &str)> = vec![("modelId", id_s.as_str())];
                if let Some(t) = &pt.pt_type {
                    attrs.push(("type", t.as_str()));
                }
                if let Some(text) = &pt.text {
                    // 有文本：构造完整 <dgm:pt> + <dgm:t>
                    w.open_with("dgm:pt", &attrs);
                    w.open("dgm:t");
                    w.empty("a:bodyPr");
                    w.empty("a:lstStyle");
                    w.open("a:p");
                    w.open("a:r");
                    let escaped = escape_xml_text(text);
                    w.leaf("a:t", escaped.as_str());
                    w.close("a:r");
                    w.close("a:p");
                    w.close("dgm:t");
                    w.close("dgm:pt");
                } else {
                    // 无文本：自闭合 <dgm:pt/>
                    w.empty_with("dgm:pt", &attrs);
                }
            } else {
                // 有 raw_xml：直接透传
                w.raw(&pt.raw_xml);
            }
        }
        w.close("dgm:ptLst");
        // cxnLst（仅当有连接时写出）
        if !self.connections.is_empty() {
            w.open("dgm:cxnLst");
            for cxn in &self.connections {
                if cxn.raw_xml.is_empty() {
                    let src_s = cxn.src_id.to_string();
                    let dst_s = cxn.dest_id.to_string();
                    let mut attrs: Vec<(&str, &str)> =
                        vec![("srcId", src_s.as_str()), ("destId", dst_s.as_str())];
                    if let Some(t) = &cxn.cxn_type {
                        attrs.push(("type", t.as_str()));
                    }
                    w.empty_with("dgm:cxn", &attrs);
                } else {
                    w.raw(&cxn.raw_xml);
                }
            }
            w.close("dgm:cxnLst");
        }
        w.close("dgm:dataModel");
        w.into_string()
    }

    /// 按 `model_id` 查找节点，返回可变引用。
    ///
    /// 便捷方法：避免外部遍历 `points` 列表。未找到返回 `None`。
    pub fn point_mut(&mut self, model_id: u32) -> Option<&mut DataModelPoint> {
        self.points.iter_mut().find(|p| p.model_id == model_id)
    }

    /// 按 `model_id` 查找节点，返回不可变引用。
    pub fn point(&self, model_id: u32) -> Option<&DataModelPoint> {
        self.points.iter().find(|p| p.model_id == model_id)
    }

    /// 设置指定 `model_id` 节点的显示文本（便捷方法）。
    ///
    /// 内部调用 [`DataModelPoint::set_text`]。未找到节点返回 `false`，
    /// 成功修改返回 `true`。
    pub fn set_point_text(&mut self, model_id: u32, new_text: impl Into<String>) -> bool {
        if let Some(pt) = self.point_mut(model_id) {
            pt.set_text(new_text);
            true
        } else {
            false
        }
    }
}

// ============================================================================
// layout part: <dgm:layoutDef>
// ============================================================================

/// SmartArt 布局定义（layout part 的半结构化表达）。
///
/// 对应 `/ppt/diagrams/layoutN.xml` 的 `<dgm:layoutDef>` 根元素。
/// 布局定义描述了 SmartArt 图形的**拓扑算法**（如线性/层次/循环/矩阵），
/// 是 SmartArt 的"模板"。
///
/// # 半结构化策略
///
/// layoutNode 算法树结构复杂（含 alg/forEach 嵌套），本结构仅提取顶层元数据
/// （uniqueId / title / desc / catLst），layoutNode 子树保留原始 XML。
#[derive(Clone, Debug, Default)]
pub struct LayoutDef {
    /// 唯一 ID（`uniqueId` 属性，用于跨 pptx 引用同一布局）。
    pub unique_id: Option<String>,
    /// 标题（`<dgm:title val="..."/>`）。
    pub title: Option<String>,
    /// 描述（`<dgm:desc val="..."/>`）。
    pub desc: Option<String>,
    /// 类别列表（`<dgm:catLst>/<dgm:cat>`）。
    pub categories: Vec<LayoutCategory>,
    /// layoutNode 子树原始 XML（`<dgm:layoutNode>...</dgm:layoutNode>` 整段）。
    ///
    /// 保留 byte-exact，避免重新序列化复杂的算法树。
    pub layout_node_xml: String,
}

/// SmartArt 布局类别（`<dgm:cat>`）。
#[derive(Clone, Debug, Default)]
pub struct LayoutCategory {
    /// 类别类型（`type` 属性，如 `process` / `hierarchy` / `cycle` / `matrix` / `pyramid` 等）。
    pub cat_type: Option<String>,
    /// 优先级（`pri` 属性，数值越小优先级越高）。
    pub priority: Option<i32>,
}

impl LayoutDef {
    /// 从 `<dgm:layoutDef>` XML 字符串解析为 [`LayoutDef`]。
    ///
    /// # 解析策略
    ///
    /// - 提取 `<dgm:layoutDef>` 的 `uniqueId` 属性；
    /// - 提取 `<dgm:title val="..."/>` 与 `<dgm:desc val="..."/>` 的 `val` 属性；
    /// - 提取 `<dgm:catLst>` 下的所有 `<dgm:cat>` 的 `type` 与 `pri` 属性；
    /// - 把 `<dgm:layoutNode>` 整段切片为 `layout_node_xml`（byte-exact）。
    ///
    /// # 错误
    /// - `Error::Xml`：XML 解析失败。
    pub fn parse_from_xml(xml: &str) -> crate::Result<LayoutDef> {
        let _ = xml; // xml 仅用于 Reader::from_str，不直接切片
        let mut layout = LayoutDef::default();

        let mut rd = quick_xml::reader::Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut in_cat_lst = false;
        let mut layout_node_depth: i32 = 0;
        // 手动累积 layoutNode 子树 raw_xml（避免依赖 buffer_position 语义）
        let mut cur_layout_raw = String::new();

        loop {
            match rd.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"layoutDef" {
                        // 提取 uniqueId 属性
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"uniqueId" {
                                layout.unique_id = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"title" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                layout.title = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"desc" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                layout.desc = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"catLst" {
                        in_cat_lst = true;
                    } else if local == b"cat" && in_cat_lst {
                        let mut cat = LayoutCategory::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"type" {
                                cat.cat_type = Some(val);
                            } else if key == b"pri" {
                                if let Ok(n) = val.parse::<i32>() {
                                    cat.priority = Some(n);
                                }
                            }
                        }
                        layout.categories.push(cat);
                    } else if local == b"layoutNode" {
                        layout_node_depth += 1;
                        if layout_node_depth == 1 {
                            // 进入顶层 layoutNode：开始累积 raw_xml
                            cur_layout_raw.clear();
                        }
                        // 累积 Start 事件（quick-xml 0.40: as_ref 不含 `<` 与 `>`）
                        cur_layout_raw.push('<');
                        cur_layout_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_layout_raw.push('>');
                    } else if layout_node_depth > 0 {
                        // layoutNode 内其他子元素 Start
                        cur_layout_raw.push('<');
                        cur_layout_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_layout_raw.push('>');
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"title" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                layout.title = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"desc" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                layout.desc = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"cat" && in_cat_lst {
                        let mut cat = LayoutCategory::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            if key == b"type" {
                                cat.cat_type = Some(val);
                            } else if key == b"pri" {
                                if let Ok(n) = val.parse::<i32>() {
                                    cat.priority = Some(n);
                                }
                            }
                        }
                        layout.categories.push(cat);
                    } else if layout_node_depth > 0 {
                        // layoutNode 内自闭合子元素
                        cur_layout_raw.push('<');
                        cur_layout_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_layout_raw.push_str("/>");
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"catLst" {
                        in_cat_lst = false;
                    } else if local == b"layoutNode" && layout_node_depth > 0 {
                        // 累积 End 事件
                        cur_layout_raw.push_str("</");
                        cur_layout_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_layout_raw.push('>');
                        layout_node_depth -= 1;
                        if layout_node_depth == 0 {
                            // 离开顶层 layoutNode：写回 raw_xml
                            layout.layout_node_xml = std::mem::take(&mut cur_layout_raw);
                        }
                    } else if layout_node_depth > 0 {
                        // layoutNode 内其他子元素 End
                        cur_layout_raw.push_str("</");
                        cur_layout_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_layout_raw.push('>');
                    }
                }
                Ok(quick_xml::events::Event::Text(t)) => {
                    if layout_node_depth > 0 {
                        cur_layout_raw.push_str(std::str::from_utf8(t.as_ref()).unwrap_or(""));
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(crate::Error::Xml(format!("LayoutDef parse_from_xml: {e}"))),
                _ => {}
            }
            buf.clear();
        }

        Ok(layout)
    }

    /// 把 [`LayoutDef`] 序列化为 `<dgm:layoutDef>` XML 字符串。
    ///
    /// layoutNode 子树通过 `layout_node_xml` 直接透传。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::new();
        w.raw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        let mut attrs: Vec<(&str, &str)> = vec![(
            "xmlns:dgm",
            "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        )];
        if let Some(id) = &self.unique_id {
            attrs.push(("uniqueId", id.as_str()));
        }
        w.open_with("dgm:layoutDef", &attrs);
        if let Some(t) = &self.title {
            w.empty_with("dgm:title", &[("val", t.as_str())]);
        }
        if let Some(d) = &self.desc {
            w.empty_with("dgm:desc", &[("val", d.as_str())]);
        }
        if !self.categories.is_empty() {
            w.open("dgm:catLst");
            for cat in &self.categories {
                // p_s 提到 cattrs 之前，保证生命周期覆盖 empty_with 调用
                let p_s = cat.priority.as_ref().map(|p| p.to_string());
                let mut cattrs: Vec<(&str, &str)> = Vec::new();
                if let Some(t) = &cat.cat_type {
                    cattrs.push(("type", t.as_str()));
                }
                if let Some(s) = p_s.as_deref() {
                    cattrs.push(("pri", s));
                }
                w.empty_with("dgm:cat", &cattrs);
            }
            w.close("dgm:catLst");
        }
        // layoutNode 子树直接透传
        if !self.layout_node_xml.is_empty() {
            w.raw(&self.layout_node_xml);
        }
        w.close("dgm:layoutDef");
        w.into_string()
    }
}

// ============================================================================
// quickStyle part: <dgm:styleData>
// ============================================================================

/// SmartArt 快速样式（quickStyle part 的半结构化表达）。
///
/// 对应 `/ppt/diagrams/quickStylesN.xml` 的 `<dgm:styleData>` 根元素。
/// quickStyle 定义了 SmartArt 节点/连接线的视觉样式（填充/边框/效果）。
///
/// # 半结构化策略
///
/// 样式标签 `<dgm:styleLbl>` 内部结构复杂（含 scene3d/sp3d/effectLst/fillLst/lnLst），
/// 本结构仅提取 `name` 属性，body 保留原始 XML。
#[derive(Clone, Debug, Default)]
pub struct QuickStyleDef {
    /// 样式标签列表（`<dgm:styleLbl>`）。
    pub style_labels: Vec<StyleLabel>,
}

/// SmartArt 样式标签（`<dgm:styleLbl>`）。
///
/// quickStyle 与 colors 都使用此结构（colors 中的 `<dgm:styleLbl>` 在 `<dgm:styleClrData>` 内）。
#[derive(Clone, Debug, Default)]
pub struct StyleLabel {
    /// 标签名（`name` 属性，如 `node0` / `node1` / `parTrans` / `sibTrans` 等）。
    pub name: Option<String>,
    /// 标签原始 XML（保留 byte-exact，含所有子元素）。
    pub raw_xml: String,
}

impl QuickStyleDef {
    /// 从 `<dgm:styleData>` XML 字符串解析为 [`QuickStyleDef`]。
    ///
    /// # 解析策略
    ///
    /// 遍历所有 `<dgm:styleLbl>` 元素（无论嵌套深度），提取 `name` 属性 + 整段 raw_xml。
    /// raw_xml 通过手动累积事件字节收集，不依赖 `buffer_position()`，保证 quick-xml 版本兼容。
    ///
    /// # 错误
    /// - `Error::Xml`：XML 解析失败。
    pub fn parse_from_xml(xml: &str) -> crate::Result<QuickStyleDef> {
        let _ = xml; // xml 仅用于 Reader::from_str，不直接切片
        let mut style_labels: Vec<StyleLabel> = Vec::new();

        let mut rd = quick_xml::reader::Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut cur_lbl: Option<StyleLabel> = None;
        let mut lbl_depth: i32 = 0;
        // 手动累积 raw_xml（避免依赖 buffer_position 的语义变化）
        let mut cur_raw = String::new();

        loop {
            match rd.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"styleLbl" {
                        if lbl_depth == 0 {
                            cur_lbl = Some(StyleLabel::default());
                            lbl_depth = 1;
                            cur_raw.clear();
                            // quick-xml 0.40: BytesStart::as_ref() 返回 `tag attrs`（不含 `<` 与 `>`）。
                            // 因此手动补 `<` 与 `>` 还原 Start 事件原始文本。
                            cur_raw.push('<');
                            cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_raw.push('>');
                            // 提取 name 属性
                            for a in e.attributes().flatten() {
                                if a.key.as_ref() == b"name" {
                                    if let Some(lbl) = cur_lbl.as_mut() {
                                        lbl.name = Some(
                                            a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                                .unwrap_or_default()
                                                .to_string(),
                                        );
                                    }
                                }
                            }
                        } else {
                            lbl_depth += 1;
                            // 嵌套的 styleLbl Start 事件也累积
                            cur_raw.push('<');
                            cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_raw.push('>');
                        }
                    } else if lbl_depth > 0 {
                        // styleLbl 内的其他 Start 元素
                        cur_raw.push('<');
                        cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_raw.push('>');
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if lbl_depth > 0 {
                        // 累积 End 事件 </tag>
                        cur_raw.push_str("</");
                        cur_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_raw.push('>');
                    }
                    if local == b"styleLbl" && lbl_depth > 0 {
                        lbl_depth -= 1;
                        if lbl_depth == 0 {
                            if let Some(mut lbl) = cur_lbl.take() {
                                lbl.raw_xml = std::mem::take(&mut cur_raw);
                                style_labels.push(lbl);
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    if lbl_depth > 0 {
                        // 累积 Empty 事件 <tag attrs/>
                        // quick-xml 0.40: BytesStart::as_ref() 不含 `<` 与 `/>`，需手动补全。
                        cur_raw.push('<');
                        cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_raw.push_str("/>");
                    }
                }
                Ok(quick_xml::events::Event::Text(t)) => {
                    if lbl_depth > 0 {
                        cur_raw.push_str(std::str::from_utf8(&t).unwrap_or(""));
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => {
                    return Err(crate::Error::Xml(format!(
                        "QuickStyleDef parse_from_xml: {e}"
                    )))
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(QuickStyleDef { style_labels })
    }

    /// 把 [`QuickStyleDef`] 序列化为 `<dgm:styleData>` XML 字符串。
    ///
    /// styleLbl 子元素通过 `raw_xml` 直接透传。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::new();
        w.raw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        w.open_with(
            "dgm:styleData",
            &[(
                "xmlns:dgm",
                "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            )],
        );
        for lbl in &self.style_labels {
            if lbl.raw_xml.is_empty() {
                let mut attrs: Vec<(&str, &str)> = Vec::new();
                if let Some(n) = &lbl.name {
                    attrs.push(("name", n.as_str()));
                }
                w.empty_with("dgm:styleLbl", &attrs);
            } else {
                w.raw(&lbl.raw_xml);
            }
        }
        w.close("dgm:styleData");
        w.into_string()
    }
}

// ============================================================================
// colors part: <dgm:colorsDef>
// ============================================================================

/// SmartArt 颜色定义（colors part 的半结构化表达）。
///
/// 对应 `/ppt/diagrams/colorsN.xml` 的 `<dgm:colorsDef>` 根元素。
/// colors 定义了 SmartArt 节点的颜色变体映射（基于主题色 accent1-6 的不同组合）。
#[derive(Clone, Debug, Default)]
pub struct ColorsDef {
    /// 唯一 ID（`uniqueId` 属性）。
    pub unique_id: Option<String>,
    /// 标题（`<dgm:title val="..."/>`）。
    pub title: Option<String>,
    /// 描述（`<dgm:desc val="..."/>`）。
    pub desc: Option<String>,
    /// 颜色样式标签列表（`<dgm:styleClrData>/<dgm:styleLbl>`）。
    pub style_color_labels: Vec<StyleLabel>,
}

impl ColorsDef {
    /// 从 `<dgm:colorsDef>` XML 字符串解析为 [`ColorsDef`]。
    ///
    /// # 解析策略
    ///
    /// - 提取 `<dgm:colorsDef>` 的 `uniqueId` 属性；
    /// - 提取 `<dgm:title>` / `<dgm:desc>` 的 `val` 属性；
    /// - 遍历 `<dgm:styleClrData>` 下的所有 `<dgm:styleLbl>`，提取 `name` + 整段 raw_xml。
    ///
    /// # 错误
    /// - `Error::Xml`：XML 解析失败。
    pub fn parse_from_xml(xml: &str) -> crate::Result<ColorsDef> {
        let _ = xml; // xml 仅用于 Reader::from_str，不直接切片
        let mut colors = ColorsDef::default();

        let mut rd = quick_xml::reader::Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();

        let mut cur_lbl: Option<StyleLabel> = None;
        let mut lbl_depth: i32 = 0;
        // 手动累积 styleLbl 子树 raw_xml（避免依赖 buffer_position 语义）
        let mut cur_raw = String::new();

        loop {
            match rd.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"colorsDef" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"uniqueId" {
                                colors.unique_id = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"title" {
                        // title 通常是自闭合 <dgm:title val="..."/>，但若以 Start 形式出现也兼容
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                colors.title = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"desc" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                colors.desc = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"styleLbl" {
                        if lbl_depth == 0 {
                            cur_lbl = Some(StyleLabel::default());
                            lbl_depth = 1;
                            cur_raw.clear();
                            // 累积 Start 事件（quick-xml 0.40: as_ref 不含 `<` 与 `>`）
                            cur_raw.push('<');
                            cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_raw.push('>');
                            for a in e.attributes().flatten() {
                                if a.key.as_ref() == b"name" {
                                    if let Some(lbl) = cur_lbl.as_mut() {
                                        lbl.name = Some(
                                            a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                                .unwrap_or_default()
                                                .to_string(),
                                        );
                                    }
                                }
                            }
                        } else {
                            lbl_depth += 1;
                            // 嵌套 styleLbl 也累积
                            cur_raw.push('<');
                            cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                            cur_raw.push('>');
                        }
                    } else if lbl_depth > 0 {
                        // styleLbl 内其他子元素 Start
                        cur_raw.push('<');
                        cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_raw.push('>');
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if local == b"title" {
                        // <dgm:title val="..."/> 自闭合形式
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                colors.title = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if local == b"desc" {
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"val" {
                                colors.desc = Some(
                                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                        .unwrap_or_default()
                                        .to_string(),
                                );
                            }
                        }
                    } else if lbl_depth > 0 {
                        // styleLbl 内自闭合子元素
                        cur_raw.push('<');
                        cur_raw.push_str(std::str::from_utf8(e.as_ref()).unwrap_or(""));
                        cur_raw.push_str("/>");
                    }
                }
                Ok(quick_xml::events::Event::End(e)) => {
                    // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                    let name = e.name();
                    let local = local_name(name.as_ref());
                    if lbl_depth > 0 {
                        // 累积 End 事件
                        cur_raw.push_str("</");
                        cur_raw.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                        cur_raw.push('>');
                    }
                    if local == b"styleLbl" && lbl_depth > 0 {
                        lbl_depth -= 1;
                        if lbl_depth == 0 {
                            if let Some(mut lbl) = cur_lbl.take() {
                                lbl.raw_xml = std::mem::take(&mut cur_raw);
                                colors.style_color_labels.push(lbl);
                            }
                        }
                    }
                }
                Ok(quick_xml::events::Event::Text(t)) => {
                    if lbl_depth > 0 {
                        cur_raw.push_str(std::str::from_utf8(t.as_ref()).unwrap_or(""));
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => return Err(crate::Error::Xml(format!("ColorsDef parse_from_xml: {e}"))),
                _ => {}
            }
            buf.clear();
        }

        Ok(colors)
    }

    /// 把 [`ColorsDef`] 序列化为 `<dgm:colorsDef>` XML 字符串。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::new();
        w.raw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        let mut attrs: Vec<(&str, &str)> = vec![(
            "xmlns:dgm",
            "http://schemas.openxmlformats.org/drawingml/2006/diagram",
        )];
        if let Some(id) = &self.unique_id {
            attrs.push(("uniqueId", id.as_str()));
        }
        w.open_with("dgm:colorsDef", &attrs);
        if let Some(t) = &self.title {
            w.empty_with("dgm:title", &[("val", t.as_str())]);
        }
        if let Some(d) = &self.desc {
            w.empty_with("dgm:desc", &[("val", d.as_str())]);
        }
        if !self.style_color_labels.is_empty() {
            w.open("dgm:styleClrData");
            for lbl in &self.style_color_labels {
                if lbl.raw_xml.is_empty() {
                    let mut lattrs: Vec<(&str, &str)> = Vec::new();
                    if let Some(n) = &lbl.name {
                        lattrs.push(("name", n.as_str()));
                    }
                    w.empty_with("dgm:styleLbl", &lattrs);
                } else {
                    w.raw(&lbl.raw_xml);
                }
            }
            w.close("dgm:styleClrData");
        }
        w.close("dgm:colorsDef");
        w.into_string()
    }
}

// ============================================================================
// 辅助函数
// ============================================================================

/// 提取元素 local name（去命名空间前缀）。
///
/// 例：`b"dgm:pt"` → `b"pt"`；`b"pt"` → `b"pt"`。
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

// ============================================================================
// 单元测试
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    /// 构造最小 dataModel XML（含虚拟根 + 1 个有文本的节点 + 1 个连接）。
    fn sample_data_model_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\
           <dgm:ptLst>\
             <dgm:pt modelId=\"0\" type=\"doc\"/>\
             <dgm:pt modelId=\"1\" type=\"par\">\
               <dgm:prSet ang=\"0\"/>\
               <dgm:spPr/>\
               <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>根节点</a:t></a:r></a:p></dgm:t>\
             </dgm:pt>\
           </dgm:ptLst>\
           <dgm:cxnLst>\
             <dgm:cxn type=\"parChld\" srcId=\"1\" destId=\"2\"/>\
           </dgm:cxnLst>\
         </dgm:dataModel>"
    }

    /// DataModel 解析：验证节点数 / modelId / type / text / 连接。
    #[test]
    fn data_model_parse_basic() {
        let xml = sample_data_model_xml();
        let dm = DataModel::parse_from_xml(xml).expect("parse dataModel");

        assert_eq!(dm.points.len(), 2, "should have 2 points");
        // 虚拟根（自闭合 <dgm:pt modelId="0" type="doc"/>）
        assert_eq!(dm.points[0].model_id, 0);
        assert_eq!(dm.points[0].pt_type.as_deref(), Some("doc"));
        assert!(dm.points[0].text.is_none(), "doc node has no text");

        // 有文本节点
        assert_eq!(dm.points[1].model_id, 1);
        assert_eq!(dm.points[1].pt_type.as_deref(), Some("par"));
        assert_eq!(dm.points[1].text.as_deref(), Some("根节点"));

        // 连接
        assert_eq!(dm.connections.len(), 1);
        assert_eq!(dm.connections[0].src_id, 1);
        assert_eq!(dm.connections[0].dest_id, 2);
        assert_eq!(dm.connections[0].cxn_type.as_deref(), Some("parChld"));
    }

    /// DataModel 解析：raw_xml 字段非空且含原始标签。
    #[test]
    fn data_model_parse_preserves_raw_xml() {
        let xml = sample_data_model_xml();
        let dm = DataModel::parse_from_xml(xml).expect("parse dataModel");

        // 第二个 pt 的 raw_xml 应包含 <dgm:prSet> 等子元素
        let pt_raw = &dm.points[1].raw_xml;
        assert!(
            pt_raw.contains("<dgm:pt"),
            "raw_xml should contain pt tag: {}",
            pt_raw
        );
        assert!(
            pt_raw.contains("<dgm:prSet"),
            "raw_xml should contain prSet: {}",
            pt_raw
        );
        assert!(
            pt_raw.contains("根节点"),
            "raw_xml should contain text: {}",
            pt_raw
        );

        // 连接的 raw_xml 应包含 <dgm:cxn
        let cxn_raw = &dm.connections[0].raw_xml;
        assert!(cxn_raw.contains("<dgm:cxn"), "cxn raw_xml: {}", cxn_raw);
    }

    /// DataModel round-trip：parse → to_xml → 关键字段仍可识别。
    #[test]
    fn data_model_round_trip() {
        let xml = sample_data_model_xml();
        let dm = DataModel::parse_from_xml(xml).expect("parse");
        let out = dm.to_xml();
        assert!(out.contains("<dgm:dataModel"), "out: {}", out);
        assert!(out.contains("<dgm:ptLst>"), "out: {}", out);
        assert!(out.contains("<dgm:cxnLst>"), "out: {}", out);
        // 节点文本应保留
        assert!(out.contains("根节点"), "out should contain text: {}", out);
        // modelId 应保留
        assert!(out.contains("modelId=\"0\""), "out: {}", out);
        assert!(out.contains("modelId=\"1\""), "out: {}", out);
    }

    /// DataModel 解析空 XML（无 ptLst）应返回空结构而非 panic。
    #[test]
    fn data_model_parse_empty_no_panic() {
        let xml = "<?xml version=\"1.0\"?>\
                   <dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>";
        let dm = DataModel::parse_from_xml(xml).expect("parse empty");
        assert!(dm.points.is_empty());
        assert!(dm.connections.is_empty());
    }

    /// DataModel 解析畸形 XML 应返回 Error::Xml。
    #[test]
    fn data_model_parse_malformed_returns_error() {
        // 注意：quick-xml 0.40 SAX 流式解析器对未闭合标签/裸属性值较宽容（不跟踪标签栈），
        // 需要用真正畸形的 XML 才能触发解析错误。
        // 未闭合注释（`<!--` 后必须有 `-->`）是 quick-xml 0.40 必报错的场景。
        let xml = "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><!-- unclosed comment <dgm:pt modelId=\"1\"/></dgm:dataModel>";
        let result = DataModel::parse_from_xml(xml);
        assert!(result.is_err(), "畸形 XML 应返回错误，实际: {result:?}");
    }

    /// 构造 layoutDef XML（含元数据 + layoutNode 子树）。
    fn sample_layout_def_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" uniqueId=\"process1\">\
           <dgm:title val=\"Process\"/>\
           <dgm:desc val=\"A simple process layout\"/>\
           <dgm:catLst>\
             <dgm:cat type=\"process\" pri=\"1\"/>\
             <dgm:cat type=\"cycle\" pri=\"2\"/>\
           </dgm:catLst>\
           <dgm:layoutNode name=\"root\">\
             <dgm:alg type=\"lin\"/>\
             <dgm:forEach name=\"node\"/>\
           </dgm:layoutNode>\
         </dgm:layoutDef>"
    }

    /// LayoutDef 解析：验证 uniqueId / title / desc / catLst / layoutNode。
    #[test]
    fn layout_def_parse_basic() {
        let xml = sample_layout_def_xml();
        let layout = LayoutDef::parse_from_xml(xml).expect("parse layoutDef");

        assert_eq!(layout.unique_id.as_deref(), Some("process1"));
        assert_eq!(layout.title.as_deref(), Some("Process"));
        assert_eq!(layout.desc.as_deref(), Some("A simple process layout"));
        assert_eq!(layout.categories.len(), 2);
        assert_eq!(layout.categories[0].cat_type.as_deref(), Some("process"));
        assert_eq!(layout.categories[0].priority, Some(1));
        assert_eq!(layout.categories[1].cat_type.as_deref(), Some("cycle"));
        assert_eq!(layout.categories[1].priority, Some(2));

        // layoutNode 子树应保留
        let ln = &layout.layout_node_xml;
        assert!(ln.contains("<dgm:layoutNode"), "layout_node_xml: {}", ln);
        assert!(ln.contains("name=\"root\""), "layout_node_xml: {}", ln);
        assert!(ln.contains("<dgm:alg"), "layout_node_xml: {}", ln);
    }

    /// LayoutDef round-trip：parse → to_xml → 关键字段仍可识别。
    #[test]
    fn layout_def_round_trip() {
        let xml = sample_layout_def_xml();
        let layout = LayoutDef::parse_from_xml(xml).expect("parse");
        let out = layout.to_xml();
        assert!(out.contains("<dgm:layoutDef"), "out: {}", out);
        assert!(out.contains("uniqueId=\"process1\""), "out: {}", out);
        assert!(out.contains("val=\"Process\""), "out: {}", out);
        assert!(out.contains("<dgm:catLst>"), "out: {}", out);
        assert!(out.contains("type=\"process\""), "out: {}", out);
        // layoutNode 子树应保留
        assert!(out.contains("<dgm:layoutNode"), "out: {}", out);
        assert!(out.contains("name=\"root\""), "out: {}", out);
    }

    /// 构造 quickStyle XML（含 2 个 styleLbl）。
    fn sample_quick_style_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <dgm:styleData xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">\
           <dgm:styleLbl name=\"node0\">\
             <dgm:spPr/>\
           </dgm:styleLbl>\
           <dgm:styleLbl name=\"node1\">\
             <dgm:spPr/>\
           </dgm:styleLbl>\
           <dgm:extLst/>\
         </dgm:styleData>"
    }

    /// QuickStyleDef 解析：验证 styleLbl 数量与 name。
    #[test]
    fn quick_style_parse_basic() {
        let xml = sample_quick_style_xml();
        let qs = QuickStyleDef::parse_from_xml(xml).expect("parse quickStyle");

        assert_eq!(qs.style_labels.len(), 2, "should have 2 styleLbl");
        assert_eq!(qs.style_labels[0].name.as_deref(), Some("node0"));
        assert_eq!(qs.style_labels[1].name.as_deref(), Some("node1"));
        // raw_xml 应含 <dgm:styleLbl
        assert!(qs.style_labels[0].raw_xml.contains("<dgm:styleLbl"));
    }

    /// QuickStyleDef round-trip：parse → to_xml → styleLbl 保留。
    #[test]
    fn quick_style_round_trip() {
        let xml = sample_quick_style_xml();
        let qs = QuickStyleDef::parse_from_xml(xml).expect("parse");
        let out = qs.to_xml();
        assert!(out.contains("<dgm:styleData"), "out: {}", out);
        assert!(out.contains("name=\"node0\""), "out: {}", out);
        assert!(out.contains("name=\"node1\""), "out: {}", out);
    }

    /// 构造 colorsDef XML（含元数据 + styleClrData/styleLbl）。
    fn sample_colors_def_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <dgm:colorsDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" uniqueId=\"color1\">\
           <dgm:title val=\"Primary Colors\"/>\
           <dgm:desc val=\"Theme color variants\"/>\
           <dgm:catLst><dgm:cat type=\"primary\" pri=\"1\"/></dgm:catLst>\
           <dgm:styleClrData>\
             <dgm:styleLbl name=\"node0\">\
               <a:effectClrLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:schemeClr val=\"accent1\"/></a:effectClrLst>\
               <a:fillClrLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:schemeClr val=\"accent1\"/></a:fillClrLst>\
             </dgm:styleLbl>\
           </dgm:styleClrData>\
         </dgm:colorsDef>"
    }

    /// ColorsDef 解析：验证元数据 + styleClrLbl。
    #[test]
    fn colors_def_parse_basic() {
        let xml = sample_colors_def_xml();
        let colors = ColorsDef::parse_from_xml(xml).expect("parse colorsDef");

        assert_eq!(colors.unique_id.as_deref(), Some("color1"));
        assert_eq!(colors.title.as_deref(), Some("Primary Colors"));
        assert_eq!(colors.desc.as_deref(), Some("Theme color variants"));
        assert_eq!(colors.style_color_labels.len(), 1);
        assert_eq!(colors.style_color_labels[0].name.as_deref(), Some("node0"));
        // raw_xml 应含 <dgm:styleLbl
        let raw = &colors.style_color_labels[0].raw_xml;
        assert!(raw.contains("<dgm:styleLbl"), "raw: {}", raw);
        assert!(
            raw.contains("accent1"),
            "raw should contain color ref: {}",
            raw
        );
    }

    /// ColorsDef round-trip：parse → to_xml → 关键字段保留。
    #[test]
    fn colors_def_round_trip() {
        let xml = sample_colors_def_xml();
        let colors = ColorsDef::parse_from_xml(xml).expect("parse");
        let out = colors.to_xml();
        assert!(out.contains("<dgm:colorsDef"), "out: {}", out);
        assert!(out.contains("uniqueId=\"color1\""), "out: {}", out);
        assert!(out.contains("val=\"Primary Colors\""), "out: {}", out);
        assert!(out.contains("<dgm:styleClrData>"), "out: {}", out);
        assert!(out.contains("name=\"node0\""), "out: {}", out);
    }

    /// local_name 辅助函数：去命名空间前缀。
    #[test]
    fn local_name_strips_prefix() {
        assert_eq!(local_name(b"dgm:pt"), b"pt");
        assert_eq!(local_name(b"pt"), b"pt");
        assert_eq!(local_name(b"a:t"), b"t");
    }

    // ===================== 文本节点编辑 API 测试（TODO-037 剩余） =====================

    /// `DataModelPoint::set_text` 同步更新 `text` 字段与 `raw_xml` 中的 `<a:t>` 内容。
    #[test]
    fn point_set_text_updates_both_fields() {
        let xml = sample_data_model_xml();
        let mut dm = DataModel::parse_from_xml(xml).expect("parse");
        // 修改 model_id=1 的节点文本
        let pt = dm.point_mut(1).expect("point 1 should exist");
        pt.set_text("新文本");
        // text 字段更新
        assert_eq!(pt.text.as_deref(), Some("新文本"));
        // raw_xml 中的 <a:t> 内容更新
        assert!(
            pt.raw_xml.contains("<a:t>新文本</a:t>"),
            "raw_xml: {}",
            pt.raw_xml
        );
        // 原始文本应被替换
        assert!(
            !pt.raw_xml.contains("根节点"),
            "old text should be replaced: {}",
            pt.raw_xml
        );
    }

    /// `DataModelPoint::set_text` 对无 `<a:t>` 的虚拟根节点（type=doc）只更新 text 字段。
    #[test]
    fn point_set_text_on_doc_node_no_a_t() {
        let xml = sample_data_model_xml();
        let mut dm = DataModel::parse_from_xml(xml).expect("parse");
        let pt = dm.point_mut(0).expect("point 0 should exist");
        // 虚拟根节点无 <a:t>，set_text 仅更新 text 字段
        pt.set_text("doc文本");
        assert_eq!(pt.text.as_deref(), Some("doc文本"));
        // raw_xml 保持不变（无 <a:t> 可替换）
        assert!(
            !pt.raw_xml.contains("doc文本"),
            "raw_xml should not change: {}",
            pt.raw_xml
        );
    }

    /// `DataModelPoint::clear_text` 清空 text 字段与 `<a:t>` 内容。
    #[test]
    fn point_clear_text_empties_both_fields() {
        let xml = sample_data_model_xml();
        let mut dm = DataModel::parse_from_xml(xml).expect("parse");
        let pt = dm.point_mut(1).expect("point 1 should exist");
        assert!(pt.text.is_some());
        pt.clear_text();
        // text 字段为 None
        assert!(pt.text.is_none());
        // raw_xml 中的 <a:t> 内容为空
        assert!(
            pt.raw_xml.contains("<a:t></a:t>"),
            "raw_xml: {}",
            pt.raw_xml
        );
        assert!(!pt.raw_xml.contains("根节点"), "raw_xml: {}", pt.raw_xml);
    }

    /// `DataModelPoint::set_text` 自动 XML 转义特殊字符。
    #[test]
    fn point_set_text_escapes_xml_special_chars() {
        let xml = sample_data_model_xml();
        let mut dm = DataModel::parse_from_xml(xml).expect("parse");
        let pt = dm.point_mut(1).expect("point 1 should exist");
        pt.set_text("a<b>&c\"d'e");
        // raw_xml 中应包含转义后的实体引用
        assert!(
            pt.raw_xml.contains("a&lt;b&gt;&amp;c&quot;d&apos;e"),
            "raw_xml: {}",
            pt.raw_xml
        );
        // text 字段保留原始未转义文本
        assert_eq!(pt.text.as_deref(), Some("a<b>&c\"d'e"));
    }

    /// `DataModel::set_point_text` 便捷方法：按 model_id 修改文本。
    #[test]
    fn data_model_set_point_text_by_id() {
        let xml = sample_data_model_xml();
        let mut dm = DataModel::parse_from_xml(xml).expect("parse");
        // 修改存在的节点
        assert!(dm.set_point_text(1, "节点1新文本"));
        assert_eq!(dm.point(1).unwrap().text.as_deref(), Some("节点1新文本"));
        // 修改不存在的节点返回 false
        assert!(!dm.set_point_text(999, "不存在"));
    }

    /// `DataModel::to_xml` 结构化重建：用户新建节点（raw_xml 为空）+ text 字段非空。
    #[test]
    fn data_model_to_xml_structured_rebuild() {
        let mut dm = DataModel::default();
        // 用户新建一个有文本的节点
        dm.points.push(DataModelPoint {
            model_id: 1,
            pt_type: Some("par".to_string()),
            text: Some("新节点".to_string()),
            raw_xml: String::new(), // 空 raw_xml 触发结构化重建
        });
        let xml = dm.to_xml();
        // 应包含 <dgm:pt modelId="1" type="par">
        assert!(
            xml.contains(r#"<dgm:pt modelId="1" type="par">"#),
            "xml: {}",
            xml
        );
        // 应包含 <dgm:t> 子元素
        assert!(xml.contains("<dgm:t>"), "xml: {}", xml);
        // 应包含 <a:bodyPr/> / <a:lstStyle/>（OOXML 顺序约束）
        assert!(xml.contains("<a:bodyPr/>"), "xml: {}", xml);
        assert!(xml.contains("<a:lstStyle/>"), "xml: {}", xml);
        // 应包含 <a:t>新节点</a:t>
        assert!(xml.contains("<a:t>新节点</a:t>"), "xml: {}", xml);
    }

    /// `DataModel::to_xml` 结构化重建：用户新建节点但无文本（自闭合 <dgm:pt/>）。
    #[test]
    fn data_model_to_xml_structured_rebuild_no_text() {
        let mut dm = DataModel::default();
        dm.points.push(DataModelPoint {
            model_id: 0,
            pt_type: Some("doc".to_string()),
            text: None,
            raw_xml: String::new(),
        });
        let xml = dm.to_xml();
        // 应自闭合 <dgm:pt modelId="0" type="doc"/>
        assert!(
            xml.contains(r#"<dgm:pt modelId="0" type="doc"/>"#),
            "xml: {}",
            xml
        );
        // 不应有 <dgm:t>
        assert!(
            !xml.contains("<dgm:t>"),
            "xml should not have dgm:t: {}",
            xml
        );
    }

    /// `escape_xml_text` 转义所有 5 个特殊字符。
    #[test]
    fn escape_xml_text_all_special_chars() {
        assert_eq!(escape_xml_text("&"), "&amp;");
        assert_eq!(escape_xml_text("<"), "&lt;");
        assert_eq!(escape_xml_text(">"), "&gt;");
        assert_eq!(escape_xml_text("'"), "&apos;");
        assert_eq!(escape_xml_text("\""), "&quot;");
        // 混合
        assert_eq!(
            escape_xml_text("a&b<c>d'e\"f"),
            "a&amp;b&lt;c&gt;d&apos;e&quot;f"
        );
        // 无特殊字符
        assert_eq!(escape_xml_text("普通文本"), "普通文本");
    }

    /// `DataModelPoint::is_type` 便捷查询方法。
    #[test]
    fn point_is_type_query() {
        let xml = sample_data_model_xml();
        let dm = DataModel::parse_from_xml(xml).expect("parse");
        assert!(dm.point(0).unwrap().is_type("doc"));
        assert!(dm.point(1).unwrap().is_type("par"));
        assert!(!dm.point(0).unwrap().is_type("par"));
    }
}
