//! # 单张幻灯片（Slide）—— 高阶 API
//!
//! 对标 python-pptx 中 `pptx.slide.Slide` 类。
//!
//! 在三层架构中处于**高阶 API 层**，直接聚合 [`Shapes`] / [`ShapesMut`]
//! 视图（用于访问形状集合），并通过 `inner: OxmlSld` 桥接到 [`crate::oxml::slide`]
//! 中的 OOXML 模型。
//!
//! # 设计要点
//!
//! - 每个 `Slide` 独立持有一份 `OxmlSld`，**不**跨 slide 共享；
//! - `id_counter` 与 `Presentation` 共享 —— 保证跨 slide 的 shape id 全局唯一；
//! - 形状增删只走 `shapes_mut()`，走完整借用检查（`&mut Slide` 才能 mutate）。
//!
//! # 示例
//!
//! ```no_run
//! use pptx::Presentation;
//! use pptx::Inches;
//!
//! let mut p = Presentation::new().unwrap();
//! let counter = p.id_counter();
//! let s = p.slides_mut().add_slide(counter).unwrap();
//! s.shapes_mut().add_textbox_with_text(
//!     Inches(1.0), Inches(1.0), Inches(4.0), Inches(1.0),
//!     "hello",
//! ).unwrap();
//! ```
//!
//! （doctest 用 `unwrap` 仅为缩短示例；生产代码应使用 `?` 传播错误。）

use std::cell::Cell;
use std::rc::Rc;

use crate::oxml::slide::Sld as OxmlSld;
use crate::oxml::SlideShape as OxmlSlideShape;
use crate::presentation::{AudioEntry, ChartEntry, DiagramEntry, MediaEntry, OleEntry, VideoEntry};
use crate::shape::base::Shape;
use crate::shape::freeform::Freeform;
use crate::shape::picture::Picture;
use crate::shape::{
    AutoShape, ChartShape, Connector, Group, OleObjectShape, ShapeKind, SmartArtShape, TableShape,
    TextBox,
};

/// 从 [`TextBody`] 中提取纯文本（段落间 `\n`）。
///
/// 辅助函数，供 [`Slide::extract_text`] 使用。
fn extract_textbody_text(tb: &crate::oxml::txbody::TextBody) -> String {
    let mut s = String::new();
    let mut first = true;
    for p in &tb.paragraphs {
        if !first {
            s.push('\n');
        }
        first = false;
        for r in &p.runs {
            s.push_str(&r.text);
        }
    }
    s
}
use crate::units::{Emu, EmuExt};

/// 幻灯片的高阶包装。**直接拥有** oxml [`OxmlSld`]。
///
/// 该类型既"包装"oxml 模型，又把 id 计数器共享给 `Presentation`，
/// 是"高阶 API ↔ OOXML 模型"之间的唯一桥梁。
#[derive(Clone, Debug)]
pub struct Slide {
    /// 共享 oxml 模型。每个 Slide 自己持有一份，不共享。
    pub(crate) inner: OxmlSld,
    /// 共享 ID 分配计数器（与 [`crate::presentation::Presentation`] 共享）。
    pub(crate) id_counter: Rc<Cell<u32>>,
    /// 本 slide 内的图片关系 id 计数器（`rIdImg1` / `rIdImg2` ...）。
    pub(crate) image_rid_counter: Rc<Cell<u32>>,
    /// 本 slide 注册到 Presentation 的媒体条目列表（保存时统一写入 zip）。
    pub(crate) media_entries: Vec<MediaEntry>,
    /// 媒体索引计数器。
    pub(crate) media_index_counter: Rc<Cell<u32>>,
    /// 本 slide 注册到 Presentation 的图表条目列表（保存时统一写入 zip）。
    ///
    /// 每个 [`ChartEntry`] 对应一个 `/ppt/charts/chartN.xml` part，
    /// 由 `ShapesMut::add_chart` 在创建图表时调用 `register_chart` 注入。
    pub(crate) chart_entries: Vec<ChartEntry>,
    /// 图表索引计数器（用于生成 `chart{N}.xml` 文件名）。
    pub(crate) chart_index_counter: Rc<Cell<u32>>,
    /// 本 slide 内的图表关系 id 计数器（`rIdChart1` / `rIdChart2` ...）。
    pub(crate) chart_rid_counter: Rc<Cell<u32>>,
    /// 引用的版式（默认 `rIdLayout1`）。
    #[allow(dead_code)]
    pub(crate) layout_rid: String,
    /// 备注 part 关系 id（指向 `ppt/notesSlides/notesSlideN.xml`）。
    /// 当 `notes` 为 `Some` 时由 `to_opc_package` 注入 `rIdNotesN`。
    pub(crate) notes_rid: Option<String>,
    /// **包内**：备注 part 路径（`/ppt/notesSlides/notesSlideN.xml`）。
    ///
    /// 用于在 read-modify-write 循环中**保留**原始 partname，
    /// 避免重新分配时导致 partname 漂移、外部引用断链。
    /// - `None`：由 `to_opc_package` 按 `notesSlide{idx}.xml` 分配；
    /// - `Some(_)`：写路径直接使用，**禁止重命名**。
    pub(crate) notes_partname: Option<String>,
    /// **包内**：`notesSlideN.xml.rels` 中指向所属 slide 的 `Target`（如 `../slides/slide1.xml`）。
    ///
    /// 由 [`crate::presentation::Presentation::from_opc`] 读 `notesSlideN.xml.rels`
    /// 中的 `Slide` 关系后回填。保存时优先复用，缺失再按 `../slides/slide{idx+1}.xml` 拼。
    pub(crate) notes_slide_rel_target: Option<String>,
    /// 该 slide 的评论列表（`<p:cmLst>`）。
    ///
    /// `None` 表示该 slide 没有评论；`Some(CommentList)` 即使为空也会写出 `commentN.xml`。
    /// 由 `Presentation::to_opc_package` 在打包阶段写入 `/ppt/comments/commentN.xml`。
    pub(crate) comments: Option<crate::oxml::comments::CommentList>,
    /// **包内**：评论 part 路径（`/ppt/comments/commentN.xml`）。
    ///
    /// 用于 read-modify-write 循环中保留原始 partname，避免漂移。
    pub(crate) comments_partname: Option<String>,
    /// **包内**：评论 part 的关系 id（`rIdCommentsN`），由 `to_opc_package` 注入。
    pub(crate) comments_rid: Option<String>,
    /// 本 slide 注册到 Presentation 的 OLE 对象条目列表（保存时统一写出 oleObjectN.bin）。
    ///
    /// 每个 [`OleEntry`] 对应一个 `/ppt/embeddings/oleObjectN.bin` part，
    /// 由 `ShapesMut::add_ole_object` 在创建 OLE 对象时调用 `register_ole` 注入。
    pub(crate) ole_entries: Vec<OleEntry>,
    /// OLE 对象索引计数器（用于生成 `oleObject{N}.bin` 文件名）。
    pub(crate) ole_index_counter: Rc<Cell<u32>>,
    /// 本 slide 内的 OLE 关系 id 计数器（`rIdOle1` / `rIdOle2` ...）。
    pub(crate) ole_rid_counter: Rc<Cell<u32>>,
    /// 本 slide 注册到 Presentation 的视频条目列表（保存时统一写出 mediaN.mp4，TODO-033）。
    ///
    /// 每个 [`VideoEntry`] 对应一个 `/ppt/media/mediaN.mp4` part，
    /// 由 `ShapesMut::add_video` 在创建视频形状时调用 `register_video` 注入。
    pub(crate) video_entries: Vec<VideoEntry>,
    /// 视频索引计数器（用于生成 `media{N}.mp4` 文件名）。
    pub(crate) video_index_counter: Rc<Cell<u32>>,
    /// 本 slide 内的视频关系 id 计数器（`rIdVideo1` / `rIdVideo2` ...）。
    pub(crate) video_rid_counter: Rc<Cell<u32>>,
    /// 本 slide 注册到 Presentation 的音频条目列表（保存时统一写出 mediaN.mp3，TODO-033）。
    ///
    /// 每个 [`AudioEntry`] 对应一个 `/ppt/media/mediaN.mp3` part，
    /// 由 `ShapesMut::add_audio` 在创建音频形状时调用 `register_audio` 注入。
    pub(crate) audio_entries: Vec<AudioEntry>,
    /// 音频索引计数器（用于生成 `media{N}.mp3` 文件名）。
    pub(crate) audio_index_counter: Rc<Cell<u32>>,
    /// 本 slide 内的音频关系 id 计数器（`rIdAudio1` / `rIdAudio2` ...）。
    pub(crate) audio_rid_counter: Rc<Cell<u32>>,
    /// 本 slide 注册到 Presentation 的 SmartArt 条目列表（保存时统一写出 4 个 diagramN.xml，TODO-037）。
    ///
    /// 每个 [`DiagramEntry`] 对应 4 个 `/ppt/diagrams/{data,layout,quickStyles,colors}N.xml` part，
    /// 由 `ShapesMut::add_diagram` 在创建 SmartArt 图形时调用 `register_diagram` 注入。
    pub(crate) diagram_entries: Vec<DiagramEntry>,
    /// SmartArt 索引计数器（用于生成 `data{N}.xml` 等文件名）。
    pub(crate) diagram_index_counter: Rc<Cell<u32>>,
    /// 本 slide 内的 SmartArt 关系 id 计数器（`rIdDgmData1` / `rIdDgmLayout1` / ...）。
    pub(crate) diagram_rid_counter: Rc<Cell<u32>>,
}

impl Slide {
    /// 包内构造：构造一个**空白** slide。
    ///
    /// 仅供 `crate::slide::Slides::add_slide` 调用；公开 API 应使用
    /// [`crate::presentation::Presentation::slides_mut`] 间接获取。
    pub(crate) fn blank(id_counter: Rc<Cell<u32>>) -> Self {
        Slide {
            inner: OxmlSld::default(),
            id_counter,
            image_rid_counter: Rc::new(Cell::new(0)),
            media_entries: Vec::new(),
            media_index_counter: Rc::new(Cell::new(0)),
            chart_entries: Vec::new(),
            chart_index_counter: Rc::new(Cell::new(0)),
            chart_rid_counter: Rc::new(Cell::new(0)),
            layout_rid: "rId1".to_string(),
            notes_rid: None,
            // 新建 slide 默认没有历史 partname：由 to_opc_package 按 notes_index 分配
            notes_partname: None,
            notes_slide_rel_target: None,
            comments: None,
            comments_partname: None,
            comments_rid: None,
            ole_entries: Vec::new(),
            ole_index_counter: Rc::new(Cell::new(0)),
            ole_rid_counter: Rc::new(Cell::new(0)),
            video_entries: Vec::new(),
            video_index_counter: Rc::new(Cell::new(0)),
            video_rid_counter: Rc::new(Cell::new(0)),
            audio_entries: Vec::new(),
            audio_index_counter: Rc::new(Cell::new(0)),
            audio_rid_counter: Rc::new(Cell::new(0)),
            diagram_entries: Vec::new(),
            diagram_index_counter: Rc::new(Cell::new(0)),
            diagram_rid_counter: Rc::new(Cell::new(0)),
        }
    }

    /// 包内构造：从一个**已解析的** [`OxmlSld`] 还原为 `Slide`。
    ///
    /// 典型用途：[`crate::presentation::Presentation::from_opc`] 在
    /// 读路径里把 `slideN.xml` 解析为 `OxmlSld` 后，**直接接管**为 Slide，
    /// 跳过 `blank` 的空壳构造。
    ///
    /// # 参数
    /// - `inner`：从 `slideN.xml` 解析得到的 `OxmlSld`；
    /// - `id_counter`：必须与所属 `Presentation` 共享（保持 shape id 全局唯一）；
    /// - `layout_rid`：从 `slideN.xml.rels` 中查到的 `SlideLayout` 关系 id。
    pub(crate) fn from_sld(inner: OxmlSld, id_counter: Rc<Cell<u32>>, layout_rid: String) -> Self {
        Slide {
            inner,
            id_counter,
            // 已加载的 slide 不再分配新 image rId（由 `parse_sld` 直接接管 pic.rid）。
            image_rid_counter: Rc::new(Cell::new(0)),
            media_entries: Vec::new(),
            media_index_counter: Rc::new(Cell::new(0)),
            chart_entries: Vec::new(),
            chart_index_counter: Rc::new(Cell::new(0)),
            chart_rid_counter: Rc::new(Cell::new(0)),
            layout_rid,
            notes_rid: None,
            // 这两个字段由 `Presentation::from_opc` 解析出 notes 后**单独**回填，
            // 构造器阶段无法获取 rels 信息。
            notes_partname: None,
            notes_slide_rel_target: None,
            comments: None,
            comments_partname: None,
            comments_rid: None,
            // OLE 嵌入：读路径当前不解析已有 oleObj，所以 from_sld 阶段为空。
            ole_entries: Vec::new(),
            ole_index_counter: Rc::new(Cell::new(0)),
            ole_rid_counter: Rc::new(Cell::new(0)),
            // 视频/音频：读路径当前不解析已有 videoFile/audioFile，所以 from_sld 阶段为空。
            video_entries: Vec::new(),
            video_index_counter: Rc::new(Cell::new(0)),
            video_rid_counter: Rc::new(Cell::new(0)),
            audio_entries: Vec::new(),
            audio_index_counter: Rc::new(Cell::new(0)),
            audio_rid_counter: Rc::new(Cell::new(0)),
            // SmartArt：读路径当前不解析已有 graphicFrame/diagram，所以 from_sld 阶段为空。
            // 写路径由 register_diagram 注入。
            diagram_entries: Vec::new(),
            diagram_index_counter: Rc::new(Cell::new(0)),
            diagram_rid_counter: Rc::new(Cell::new(0)),
        }
    }

    /// 把当前 slide 序列化为 XML 字符串。
    ///
    /// # 用途
    /// - 调试：直接 `println!("{}", slide.to_xml())`；
    /// - 测试 fixture：与已知 snapshot 对比；
    /// - 自定义输出管道：把 XML 与 zip 步骤解耦。
    pub fn to_xml(&self) -> String {
        self.inner.to_xml()
    }

    /// 不可变形状集合视图。
    pub fn shapes(&self) -> Shapes<'_> {
        Shapes { slide: self }
    }
    /// 可变形状集合视图。
    pub fn shapes_mut(&mut self) -> ShapesMut<'_> {
        ShapesMut { slide: self }
    }

    /// 内部 ID（在所属 `Presentation` 内的 sldIdLst 序号，与 shape id 独立）。
    pub fn internal_id(&self) -> u32 {
        self.inner.id
    }
    /// 取 layout 关系 id（指向 `ppt/slideLayouts/slideLayoutN.xml`）。
    pub fn layout_rid(&self) -> String {
        self.inner.layout_rid.clone()
    }
    /// 设置 layout 关系 id。
    pub fn set_layout_rid(&mut self, rid: String) {
        self.inner.layout_rid = rid;
    }

    /// 取 slide 名（对应 `p:sld/p:cSld/@name`，空字符串表示未命名）。
    ///
    /// 对标 python-pptx `Slide.name`。
    pub fn name(&self) -> &str {
        &self.inner.name
    }
    /// 设置 slide 名。`None` 或空字符串等价于"移除名字"。
    pub fn set_name(&mut self, name: Option<&str>) {
        self.inner.name = name.unwrap_or("").to_string();
    }

    /// 备注文本（speaker notes）拼成单字符串。
    ///
    /// 返回 `None` 表示当前 slide 没有任何 `<p:notes>` 部分。
    /// 多段落以 `\n` 拼接，与 python-pptx 行为一致。
    pub fn notes_text(&self) -> Option<String> {
        self.inner.notes.as_ref().map(|tb| {
            let mut s = String::new();
            let mut first = true;
            for p in &tb.paragraphs {
                if !first {
                    s.push('\n');
                }
                first = false;
                for r in &p.runs {
                    s.push_str(&r.text);
                }
            }
            s
        })
    }

    /// **是否**存在 notes slide。
    ///
    /// 对标 python-pptx `Slide.has_notes_slide`。
    /// 与 [`Self::notes_text`] 不同：后者读取 notes 内容，前者只判断"是否创建了
    /// notes 容器"——一旦写过 notes，本值即**持续为 true**，直至显式 `set_notes_text(None)`。
    pub fn has_notes_slide(&self) -> bool {
        self.inner.notes.is_some()
    }

    /// **是否**继承 master 背景。
    ///
    /// 对标 python-pptx `Slide.follow_master_background`。
    ///
    /// # 实现语义
    /// - `inner.background` 为 `None`：未设置独立背景，渲染时回退到 master → 返回 `true`；
    /// - `inner.background` 为 `Some(SlideBackground::Reference { idx=1001, scheme_color="bg1" })`：
    ///   显式引用 master 背景 → 返回 `true`；
    /// - `inner.background` 为 `Some(SlideBackground::Property(_))`：已设置独立背景 → 返回 `false`；
    /// - 其它 `Reference`（非 bg1/1001）：视为"引用主题背景样式"，仍算继承 → 返回 `true`。
    pub fn follow_master_background(&self) -> bool {
        match &self.inner.background {
            None => true,
            Some(crate::oxml::slide::SlideBackground::Reference(_)) => true,
            Some(crate::oxml::slide::SlideBackground::Property(_)) => false,
        }
    }

    /// 设置"是否继承 master 背景"。
    ///
    /// - `v = true`：清空独立背景（`inner.background = None`），渲染时回退到 master；
    /// - `v = false`：若当前已是独立背景则保留；否则写入一个"占位"的纯白背景，
    ///   后续可通过 [`Self::set_background_solid`] 修改颜色。
    ///
    /// 对标 python-pptx `Slide.follow_master_background = True/False`。
    pub fn set_follow_master_background(&mut self, v: bool) {
        if v {
            // 清空独立背景，回退到 master
            self.inner.background = None;
        } else if self.follow_master_background() {
            // 当前是继承状态，切换为独立背景：写入一个默认纯白背景占位
            self.inner.background = Some(crate::oxml::slide::SlideBackground::solid(
                crate::oxml::color::Color::RGB(crate::units::RGBColor::WHITE),
            ));
        }
        // 若已经是独立背景且 v=false，则保持不变
    }

    /// 设置**纯色**背景（写出 `<p:bg><p:bgPr><a:solidFill>...</a:solidFill></p:bgPr></p:bg>`）。
    ///
    /// 对标 python-pptx `slide.background.fill.solid(); slide.background.fill.fore_color.rgb = ...`。
    ///
    /// # 参数
    /// - `color`：填充颜色（`Color::RGB` / `Color::Scheme` / `Color::Preset`）；
    ///   `Color::None` 等价于 [`Self::clear_background`]。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::{Presentation, Inches, RGBColor};
    /// # use pptx::oxml::color::Color;
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// s.set_background_solid(Color::RGB(RGBColor::RED));
    /// ```
    pub fn set_background_solid(&mut self, color: crate::oxml::color::Color) {
        if matches!(color, crate::oxml::color::Color::None) {
            self.clear_background();
            return;
        }
        self.inner.background = Some(crate::oxml::slide::SlideBackground::solid(color));
    }

    /// 清空独立背景（回退到继承 master 背景）。
    ///
    /// 等价于 `set_follow_master_background(true)`。
    pub fn clear_background(&mut self) {
        self.inner.background = None;
    }

    /// 提取 slide 中所有文本内容（纯文本，不含格式）。
    ///
    /// 对标 pypdf `PageObject.extract_text()` / python-pptx `Shape.text_frame.text`。
    /// 遍历 slide 上所有形状的文本体，把每个 `Run` 的 `text` 拼接成单字符串，
    /// 段落间以 `\n` 分隔，形状间以 `\n\n` 分隔。
    ///
    /// # 与 pypdf 的差异
    /// - pypdf 的 `extract_text()` 按 PDF 内容流顺序提取，可能乱序；
    /// - 本方法按 slide XML 中的形状声明顺序提取，与 PowerPoint 中阅读顺序一致。
    pub fn extract_text(&self) -> String {
        let mut out = String::new();
        let mut first_shape = true;
        for sh in &self.inner.shapes {
            let tb = match sh {
                OxmlSlideShape::Sp(sp) => &sp.text,
                OxmlSlideShape::Pic(_) => continue,
                OxmlSlideShape::CxnSp(_) => continue,
                OxmlSlideShape::Group(grp) => {
                    // 递归提取 group 内的文本
                    let mut grp_text = String::new();
                    let mut first = true;
                    for child in &grp.children {
                        if let crate::oxml::shape::GroupChild::Sp(sp) = child {
                            if !first {
                                grp_text.push_str("\n\n");
                            }
                            first = false;
                            grp_text.push_str(&extract_textbody_text(&sp.text));
                        }
                    }
                    if grp_text.is_empty() {
                        continue;
                    }
                    if !first_shape {
                        out.push_str("\n\n");
                    }
                    first_shape = false;
                    out.push_str(&grp_text);
                    continue;
                }
                OxmlSlideShape::GraphicFrame(_) => continue,
            };
            let text = extract_textbody_text(tb);
            if text.is_empty() {
                continue;
            }
            if !first_shape {
                out.push_str("\n\n");
            }
            first_shape = false;
            out.push_str(&text);
        }
        out
    }

    /// 深拷贝当前 slide（分配新 id / rid / partname 由 `Slides` 在插入时处理）。
    ///
    /// 对标 pypdf `PdfWriter.clone_page_from_reader`。
    /// 返回的 `Slide` 与原 slide **完全独立**——修改克隆体不影响原件。
    pub fn clone_slide(&self) -> Slide {
        self.clone()
    }

    /// 设置标题占位符的文本（TODO-007）。
    ///
    /// 对标 python-pptx `slide.shapes.title.text = "..."`。
    ///
    /// 查找策略与 [`Shapes::title`] 一致：优先 `ph_type == "title"` / `"ctrTitle"`，
    /// 其次 `ph_idx == 0`。找到后**替换**其文本体为单段落单 Run。
    ///
    /// # 返回
    /// - `true`：找到标题占位符并已设置；
    /// - `false`：未找到标题占位符。
    pub fn set_title_text(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    let is_title = sp
                        .ph_type
                        .as_deref()
                        .map(|t| t == "title" || t == "ctrTitle")
                        .unwrap_or(false)
                        || sp.ph_idx == Some(0);
                    if is_title {
                        let mut tb = crate::oxml::txbody::TextBody::new();
                        tb.set_text(text);
                        sp.text = tb;
                        return true;
                    }
                }
            }
        }
        false
    }

    /// 取标题占位符的文本（TODO-007）。
    ///
    /// 对标 python-pptx `slide.shapes.title.text`。
    /// 未找到标题占位符时返回 `None`。
    pub fn title_text(&self) -> Option<String> {
        for sh in &self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    let is_title = sp
                        .ph_type
                        .as_deref()
                        .map(|t| t == "title" || t == "ctrTitle")
                        .unwrap_or(false)
                        || sp.ph_idx == Some(0);
                    if is_title {
                        return Some(extract_textbody_text(&sp.text));
                    }
                }
            }
        }
        None
    }

    /// 向正文占位符**追加**一个段落（TODO-007）。
    ///
    /// 对标 python-pptx `slide.placeholders[1].text_frame.add_paragraph()`。
    ///
    /// 查找策略：优先 `ph_type == "body"`，其次 `ph_idx == 1`。
    /// 找到后在文本体末尾追加一个新段落（单 Run，文本为 `text`）。
    ///
    /// # 返回
    /// - `true`：找到正文占位符并已追加；
    /// - `false`：未找到正文占位符。
    pub fn append_body_paragraph(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    let is_body = sp
                        .ph_type
                        .as_deref()
                        .map(|t| t == "body" || t == "obj")
                        .unwrap_or(false)
                        || sp.ph_idx == Some(1);
                    if is_body {
                        let r = crate::oxml::txbody::Run {
                            text: text.to_string(),
                            ..Default::default()
                        };
                        let mut p = crate::oxml::txbody::Paragraph::default();
                        p.runs.push(r);
                        sp.text.paragraphs.push(p);
                        return true;
                    }
                }
            }
        }
        false
    }

    /// 设置正文占位符的文本（**替换**全部段落，TODO-007）。
    ///
    /// 对标 python-pptx `slide.placeholders[1].text_frame.text = "..."`。
    ///
    /// # 返回
    /// - `true`：找到正文占位符并已设置；
    /// - `false`：未找到正文占位符。
    pub fn set_body_text(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    let is_body = sp
                        .ph_type
                        .as_deref()
                        .map(|t| t == "body" || t == "obj")
                        .unwrap_or(false)
                        || sp.ph_idx == Some(1);
                    if is_body {
                        let mut tb = crate::oxml::txbody::TextBody::new();
                        tb.set_text(text);
                        sp.text = tb;
                        return true;
                    }
                }
            }
        }
        false
    }

    /// 取正文占位符的文本（TODO-007）。
    ///
    /// 对标 python-pptx `slide.placeholders[1].text_frame.text`。
    /// 未找到正文占位符时返回 `None`。
    pub fn body_text(&self) -> Option<String> {
        for sh in &self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    let is_body = sp
                        .ph_type
                        .as_deref()
                        .map(|t| t == "body" || t == "obj")
                        .unwrap_or(false)
                        || sp.ph_idx == Some(1);
                    if is_body {
                        return Some(extract_textbody_text(&sp.text));
                    }
                }
            }
        }
        None
    }

    /// 设置页脚占位符的文本（TODO-007 剩余小项）。
    ///
    /// 对标 python-pptx `slide.placeholders[footer_idx].text_frame.text = "..."`。
    ///
    /// 查找策略：仅按 `ph_type == "ftr"` 匹配（不按 `ph_idx` 回退，因为页脚
    /// 占位符的 idx 在不同版式中取值不一）。找到后**替换**其文本体为单段落单 Run。
    ///
    /// # 返回
    /// - `true`：找到页脚占位符并已设置；
    /// - `false`：未找到页脚占位符（页脚占位符需由版式/母版提供）。
    pub fn set_footer_text(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("ftr") {
                    let mut tb = crate::oxml::txbody::TextBody::new();
                    tb.set_text(text);
                    sp.text = tb;
                    return true;
                }
            }
        }
        false
    }

    /// 取页脚占位符的文本（TODO-007 剩余小项）。
    ///
    /// 未找到页脚占位符时返回 `None`。
    pub fn footer_text(&self) -> Option<String> {
        for sh in &self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("ftr") {
                    return Some(extract_textbody_text(&sp.text));
                }
            }
        }
        None
    }

    /// 设置日期占位符的文本（TODO-007 剩余小项）。
    ///
    /// 对标 python-pptx `slide.placeholders[dt_idx].text_frame.text = "..."`。
    ///
    /// 查找策略：仅按 `ph_type == "dt"` 匹配。找到后**替换**其文本体。
    ///
    /// # 注意
    /// PowerPoint 默认会让日期占位符显示"自动更新日期"；一旦显式设置文本，
    /// 会覆盖自动日期。如需恢复自动日期，请重新从版式继承占位符。
    ///
    /// # 返回
    /// - `true`：找到日期占位符并已设置；
    /// - `false`：未找到日期占位符。
    pub fn set_date_text(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("dt") {
                    let mut tb = crate::oxml::txbody::TextBody::new();
                    tb.set_text(text);
                    sp.text = tb;
                    return true;
                }
            }
        }
        false
    }

    /// 取日期占位符的文本（TODO-007 剩余小项）。
    ///
    /// 未找到日期占位符时返回 `None`。
    pub fn date_text(&self) -> Option<String> {
        for sh in &self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("dt") {
                    return Some(extract_textbody_text(&sp.text));
                }
            }
        }
        None
    }

    /// 设置幻灯片编号占位符的文本（TODO-007 剩余小项）。
    ///
    /// 对标 python-pptx `slide.placeholders[sldNum_idx].text_frame.text = "..."`。
    ///
    /// 查找策略：仅按 `ph_type == "sldNum"` 匹配。找到后**替换**其文本体。
    ///
    /// # 注意
    /// 与日期占位符类似，PowerPoint 默认会自动渲染当前页码；显式设置文本
    /// 会覆盖自动页码。
    ///
    /// # 返回
    /// - `true`：找到编号占位符并已设置；
    /// - `false`：未找到编号占位符。
    pub fn set_slide_number_text(&mut self, text: &str) -> bool {
        for sh in &mut self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("sldNum") {
                    let mut tb = crate::oxml::txbody::TextBody::new();
                    tb.set_text(text);
                    sp.text = tb;
                    return true;
                }
            }
        }
        false
    }

    /// 取幻灯片编号占位符的文本（TODO-007 剩余小项）。
    ///
    /// 未找到编号占位符时返回 `None`。
    pub fn slide_number_text(&self) -> Option<String> {
        for sh in &self.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_type.as_deref() == Some("sldNum") {
                    return Some(extract_textbody_text(&sp.text));
                }
            }
        }
        None
    }

    /// 取得 slide 背景的**高阶视图**（只读）。
    ///
    /// 对标 python-pptx `Slide.background`。
    ///
    /// # 实现说明
    /// 返回的 [`SlideBackground`] 句柄仅提供**读取**能力（如 `fill_type()`）。
    /// 若需修改背景，请使用以下方法：
    /// - [`Self::set_background_solid`]：设置纯色背景；
    /// - [`Self::clear_background`]：清空独立背景；
    /// - [`Self::set_follow_master_background`]：切换继承/独立。
    pub fn background(&self) -> SlideBackground<'_> {
        SlideBackground { slide: self }
    }

    /// 读取幻灯片过渡（`<p:transition>`）。
    ///
    /// 对标 python-pptx `slide.transition`（python-pptx 实际只暴露底层元素，本 API 返回结构体）。
    ///
    /// 返回 `Some(&Transition)` 表示该幻灯片已设置过渡；`None` 表示未设置（遵循 PowerPoint 默认行为）。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::Presentation;
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// if let Some(t) = s.transition() {
    ///     println!("speed = {:?}", t.speed);
    /// }
    /// ```
    pub fn transition(&self) -> Option<&crate::oxml::slide::Transition> {
        self.inner.transition.as_ref()
    }

    /// 设置幻灯片过渡（`<p:transition>`）。
    ///
    /// 对标 python-pptx 中通过 `slide.element.transition` 操作过渡的方式，本 API 接收结构体直接覆盖。
    ///
    /// 若传入 `TransitionType::None`，等价于 [`Self::clear_transition`]。
    ///
    /// # 参数
    /// - `transition`：过渡配置（速度/换片方式/类型）
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::{Presentation, Transition, TransitionSpeed, TransitionType};
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// let t = Transition {
    ///     speed: TransitionSpeed::Slow,
    ///     advance_click: true,
    ///     advance_after_ms: Some(5000),
    ///     transition_type: TransitionType::Fade { thru_blk: false },
    /// };
    /// s.set_transition(t);
    /// ```
    pub fn set_transition(&mut self, transition: crate::oxml::slide::Transition) {
        if matches!(
            transition.transition_type,
            crate::oxml::slide::TransitionType::None
        ) {
            self.inner.transition = None;
        } else {
            self.inner.transition = Some(transition);
        }
    }

    /// 清除幻灯片过渡（等价于删除 `<p:transition>` 元素）。
    ///
    /// 清除后该幻灯片将使用 PowerPoint 默认的"无过渡"行为。
    pub fn clear_transition(&mut self) {
        self.inner.transition = None;
    }

    /// 备注文本体（`TextBody`）的可选引用。
    ///
    /// 对标 python-pptx `slide.notes_slide.notes_text_frame`。
    /// 当未设置备注时返回 `None`。
    pub fn notes(&self) -> Option<&crate::oxml::txbody::TextBody> {
        self.inner.notes.as_ref()
    }
    /// 备注文本体（`TextBody`）的可变引用。
    pub fn notes_mut(&mut self) -> Option<&mut crate::oxml::txbody::TextBody> {
        self.inner.notes.as_mut()
    }
    /// 直接覆盖备注文本体。`None` 表示删除备注。
    pub fn set_notes(&mut self, tb: Option<crate::oxml::txbody::TextBody>) {
        self.inner.notes = tb;
    }

    /// 备注 part 的关系 id（`rIdNotesN`），由 `Presentation::save` 在打包时注入。
    pub fn notes_rid(&self) -> Option<&str> {
        self.notes_rid.as_deref()
    }
    /// 显式设置备注 part 的关系 id（**仅供 `Presentation::save` 内部使用**）。
    pub(crate) fn set_notes_rid(&mut self, rid: String) {
        self.notes_rid = Some(rid);
    }

    /// 取备注 part 路径（仅供 `Presentation::save` 内部使用）。
    pub(crate) fn notes_partname(&self) -> Option<&str> {
        self.notes_partname.as_deref()
    }
    /// 设置备注 part 路径（**仅供 `Presentation::from_opc` 内部使用**）。
    ///
    /// 由读路径在解析 `slideN.xml.rels` 找到 `NotesSlide` 关系后回填，
    /// 保证 save 时复用原始 partname。
    pub(crate) fn set_notes_partname(&mut self, partname: String) {
        self.notes_partname = Some(partname);
    }

    /// 取 `notesSlideN.xml.rels` 中的 `Slide` 关系 target（仅供内部）。
    pub(crate) fn notes_slide_rel_target(&self) -> Option<&str> {
        self.notes_slide_rel_target.as_deref()
    }
    /// 设置 `notesSlideN.xml.rels` 中的 `Slide` 关系 target（**仅供内部**）。
    ///
    /// 由读路径在解析 `notesSlideN.xml.rels` 找到 `Slide` 关系后回填。
    pub(crate) fn set_notes_slide_rel_target(&mut self, target: String) {
        self.notes_slide_rel_target = Some(target);
    }

    /// 设置备注文本（**整体替换**）。
    ///
    /// 若 `text` 为 `None`，则删除 notes；否则按 `\n` 切分为多段。
    /// 备注的持久化由 [`crate::presentation::Presentation::save`] 在打包阶段
    /// 写入 `/ppt/notesSlides/notesSlideN.xml`；本方法仅设置内存模型。
    pub fn set_notes_text(&mut self, text: Option<&str>) {
        if let Some(t) = text {
            let mut tb = crate::oxml::txbody::TextBody::new();
            for line in t.split('\n') {
                let mut p = crate::oxml::txbody::Paragraph::new();
                p.runs.push(crate::oxml::txbody::Run::new(line));
                tb.paragraphs.push(p);
            }
            self.inner.notes = Some(tb);
        } else {
            self.inner.notes = None;
        }
    }

    // ==================== 评论 API ====================

    /// 返回该 slide 的评论列表（不可变）。
    ///
    /// `None` 表示该 slide 没有评论 part。
    pub fn comments(&self) -> Option<&crate::oxml::comments::CommentList> {
        self.comments.as_ref()
    }

    /// 返回该 slide 的评论列表（可变）。
    pub fn comments_mut(&mut self) -> Option<&mut crate::oxml::comments::CommentList> {
        self.comments.as_mut()
    }

    /// 直接覆盖评论列表。`None` 表示删除评论。
    pub fn set_comments(&mut self, lst: Option<crate::oxml::comments::CommentList>) {
        self.comments = lst;
    }

    /// 添加一条评论。
    ///
    /// 这是便捷 API：自动维护评论索引（`idx` 在该 slide 内递增），
    /// 并确保 `comments` 字段为 `Some`。
    ///
    /// # 参数
    /// - `author_id`：作者 ID（需在 `Presentation.comment_authors` 中存在）；
    /// - `pos_x` / `pos_y`：评论锚点坐标（EMU）；
    /// - `text`：评论正文。
    ///
    /// # 返回值
    /// 新评论在该 slide 中的 `idx`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx::Presentation;
    /// use pptx::Inches;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// let counter = p.id_counter();
    /// // 先注册作者，拿到 author_id（必须在借用 slides_mut 之前完成）
    /// let author_id = p.comment_authors_mut().get_or_insert_id("张三", "ZS");
    /// let slide = p.slides_mut().add_slide(counter).unwrap();
    /// let idx = slide.add_comment(author_id, Inches(1.0), Inches(1.0), "评论内容");
    /// ```
    pub fn add_comment<X: crate::units::EmuExt, Y: crate::units::EmuExt>(
        &mut self,
        author_id: u32,
        pos_x: X,
        pos_y: Y,
        text: impl Into<String>,
    ) -> u32 {
        let lst = self
            .comments
            .get_or_insert_with(crate::oxml::comments::CommentList::new);
        // idx 在该 slide 内递增（从 1 开始）
        let next_idx = lst
            .comments
            .iter()
            .map(|c| c.idx)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        let c = crate::oxml::comments::Comment::new(
            author_id,
            next_idx,
            pos_x.emu().value(),
            pos_y.emu().value(),
            text,
        );
        lst.push(c);
        next_idx
    }

    /// 清除该 slide 的所有评论。
    pub fn clear_comments(&mut self) {
        self.comments = None;
    }

    /// 评论 part 路径（仅供 `Presentation::save` 内部使用）。
    pub(crate) fn comments_partname(&self) -> Option<&str> {
        self.comments_partname.as_deref()
    }

    /// 设置评论 part 路径（**仅供 `Presentation::from_opc` 内部使用**）。
    pub(crate) fn set_comments_partname(&mut self, partname: String) {
        self.comments_partname = Some(partname);
    }

    /// 评论 part 的关系 id（仅供内部）。
    #[allow(dead_code)]
    pub(crate) fn comments_rid(&self) -> Option<&str> {
        self.comments_rid.as_deref()
    }

    /// 设置评论 part 的关系 id（**仅供 `Presentation::save` 内部使用**）。
    pub(crate) fn set_comments_rid(&mut self, rid: String) {
        self.comments_rid = Some(rid);
    }

    /// 分配一个**全局唯一**的 shape id（在所属 `Presentation` 内）。
    ///
    /// 由 `ShapesMut` 的 `add_*` 流程调用；用户**不应**直接调用。
    pub(crate) fn next_shape_id(&self) -> u32 {
        let v = self.id_counter.get() + 1;
        self.id_counter.set(v);
        v
    }

    /// 分配一个**全局唯一**的图片关系 id（`rIdImg1` / `rIdImg2` ...）。
    ///
    /// 由 [`crate::slide::ShapesMut::add_picture`] 内部调用。
    pub(crate) fn allocate_image_rid(&self) -> String {
        let v = self.image_rid_counter.get() + 1;
        self.image_rid_counter.set(v);
        format!("rIdImg{}", v)
    }

    /// 分配一个**全局唯一**的媒体索引。
    pub(crate) fn next_media_index(&self) -> u32 {
        let v = self.media_index_counter.get() + 1;
        self.media_index_counter.set(v);
        v
    }

    /// 把媒体条目注册到本 slide（保存时统一写 zip）。
    pub(crate) fn register_media(&mut self, entry: MediaEntry) {
        self.media_entries.push(entry);
    }

    /// 分配一个**本 slide 内**唯一的图表关系 id（`rIdChart1` / `rIdChart2` ...）。
    ///
    /// 由 [`ShapesMut::add_chart`] 内部调用，用于在 `slideN.xml.rels` 中
    /// 显式添加 `<Relationship Type=".../chart" Target="../charts/chartN.xml"/>`。
    pub(crate) fn allocate_chart_rid(&self) -> String {
        let v = self.chart_rid_counter.get() + 1;
        self.chart_rid_counter.set(v);
        format!("rIdChart{}", v)
    }

    /// 分配一个图表 part 索引（用于生成 `chart{N}.xml` 文件名）。
    ///
    /// **注意**：该计数器是 slide 局部的；`to_opc_package` 在打包阶段会
    /// 用一个**全局**递增的 `chart_index` 重新分配 partname，避免多 slide
    /// 之间的索引冲突。本方法仅用于在 `add_chart` 时占位。
    pub(crate) fn next_chart_index(&self) -> u32 {
        let v = self.chart_index_counter.get() + 1;
        self.chart_index_counter.set(v);
        v
    }

    /// 把图表条目注册到本 slide（保存时统一写出 chartN.xml part + rels）。
    pub(crate) fn register_chart(&mut self, entry: ChartEntry) {
        self.chart_entries.push(entry);
    }

    /// 分配一个**本 slide 内**唯一的 OLE 关系 id（`rIdOle1` / `rIdOle2` ...）。
    ///
    /// 由 [`ShapesMut::add_ole_object`] 内部调用，用于在 `slideN.xml.rels` 中
    /// 显式添加 `<Relationship Type=".../oleObject" Target="../embeddings/oleObjectN.bin"/>`。
    pub(crate) fn allocate_ole_rid(&self) -> String {
        let v = self.ole_rid_counter.get() + 1;
        self.ole_rid_counter.set(v);
        format!("rIdOle{}", v)
    }

    /// 分配一个 OLE part 索引（用于生成 `oleObject{N}.bin` 文件名）。
    ///
    /// **注意**：该计数器是 slide 局部的；`to_opc_package` 在打包阶段会
    /// 用一个**全局**递增的 `ole_global_index` 重新分配 partname，避免多 slide
    /// 之间的索引冲突。本方法仅用于在 `add_ole_object` 时占位。
    pub(crate) fn next_ole_index(&self) -> u32 {
        let v = self.ole_index_counter.get() + 1;
        self.ole_index_counter.set(v);
        v
    }

    /// 把 OLE 对象条目注册到本 slide（保存时统一写出 oleObjectN.bin part + rels）。
    pub(crate) fn register_ole(&mut self, entry: OleEntry) {
        self.ole_entries.push(entry);
    }

    /// 分配一个**本 slide 内**唯一的视频关系 id（`rIdVideo1` / `rIdVideo2` ...，TODO-033）。
    ///
    /// 由 [`ShapesMut::add_video`] 内部调用，用于在 `slideN.xml.rels` 中
    /// 显式添加 `<Relationship Type=".../video" Target="../media/mediaN.mp4"/>`。
    pub(crate) fn allocate_video_rid(&self) -> String {
        let v = self.video_rid_counter.get() + 1;
        self.video_rid_counter.set(v);
        format!("rIdVideo{}", v)
    }

    /// 分配一个视频 part 索引（用于生成 `media{N}.mp4` 文件名，TODO-033）。
    ///
    /// **注意**：该计数器是 slide 局部的；`to_opc_package` 在打包阶段会
    /// 用一个**全局**递增的 `video_global_index` 重新分配 partname，避免多 slide
    /// 之间的索引冲突。本方法仅用于在 `add_video` 时占位。
    pub(crate) fn next_video_index(&self) -> u32 {
        let v = self.video_index_counter.get() + 1;
        self.video_index_counter.set(v);
        v
    }

    /// 把视频条目注册到本 slide（保存时统一写出 mediaN.mp4 part + rels，TODO-033）。
    pub(crate) fn register_video(&mut self, entry: VideoEntry) {
        self.video_entries.push(entry);
    }

    /// 分配一个**本 slide 内**唯一的音频关系 id（`rIdAudio1` / `rIdAudio2` ...，TODO-033）。
    ///
    /// 由 [`ShapesMut::add_audio`] 内部调用，用于在 `slideN.xml.rels` 中
    /// 显式添加 `<Relationship Type=".../audio" Target="../media/mediaN.mp3"/>`。
    pub(crate) fn allocate_audio_rid(&self) -> String {
        let v = self.audio_rid_counter.get() + 1;
        self.audio_rid_counter.set(v);
        format!("rIdAudio{}", v)
    }

    /// 分配一个音频 part 索引（用于生成 `media{N}.mp3` 文件名，TODO-033）。
    ///
    /// **注意**：该计数器是 slide 局部的；`to_opc_package` 在打包阶段会
    /// 用一个**全局**递增的 `audio_global_index` 重新分配 partname，避免多 slide
    /// 之间的索引冲突。本方法仅用于在 `add_audio` 时占位。
    pub(crate) fn next_audio_index(&self) -> u32 {
        let v = self.audio_index_counter.get() + 1;
        self.audio_index_counter.set(v);
        v
    }

    /// 把音频条目注册到本 slide（保存时统一写出 mediaN.mp3 part + rels，TODO-033）。
    pub(crate) fn register_audio(&mut self, entry: AudioEntry) {
        self.audio_entries.push(entry);
    }

    // --------------------- SmartArt（TODO-037） ---------------------

    /// 分配一个 SmartArt part 索引（用于生成 `data{N}.xml` / `layout{N}.xml` 等文件名，TODO-037）。
    ///
    /// **注意**：该计数器是 slide 局部的；`to_opc_package` 在打包阶段会
    /// 用一个**全局**递增的 `diagram_global_index` 重新分配 partname，避免多 slide
    /// 之间的索引冲突。本方法仅用于在 `add_diagram` 时占位。
    pub(crate) fn next_diagram_index(&self) -> u32 {
        let v = self.diagram_index_counter.get() + 1;
        self.diagram_index_counter.set(v);
        v
    }

    /// 分配 4 个 SmartArt 关系 id（`rIdDgmData1` / `rIdDgmLayout1` / `rIdDgmQs1` / `rIdDgmColors1`，TODO-037）。
    ///
    /// 由 `ShapesMut::add_diagram` 内部调用，用于在 `slideN.xml.rels` 中
    /// 显式添加 4 个关系：`diagramData` / `diagramLayout` / `diagramQuickStyle` / `diagramColors`。
    ///
    /// 返回值顺序固定为 `(data_rid, layout_rid, quick_style_rid, colors_rid)`。
    pub(crate) fn allocate_diagram_rids(&self) -> (String, String, String, String) {
        let v = self.diagram_rid_counter.get() + 1;
        self.diagram_rid_counter.set(v);
        (
            format!("rIdDgmData{}", v),
            format!("rIdDgmLayout{}", v),
            format!("rIdDgmQs{}", v),
            format!("rIdDgmColors{}", v),
        )
    }

    /// 把 SmartArt 条目注册到本 slide（保存时统一写出 4 个 diagramN.xml part + rels，TODO-037）。
    pub(crate) fn register_diagram(&mut self, entry: DiagramEntry) {
        self.diagram_entries.push(entry);
    }
}

/// 不可变形状视图。
///
/// 通过 `slide.shapes()` 获取，提供 `len` / `is_empty` / `iter` / `get` 等只读 API。
#[derive(Debug)]
pub struct Shapes<'a> {
    slide: &'a Slide,
}

impl<'a> Shapes<'a> {
    /// 形状数量。
    pub fn len(&self) -> usize {
        self.slide.inner.shapes.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.slide.inner.shapes.is_empty()
    }

    /// 遍历所有形状（克隆为 `ShapeKind` 枚举）。
    ///
    /// 返回的 `ShapeKind` 是高阶句柄，与原 `Slide` 拥有独立生命周期；
    /// 适合"先收集再逐个处理"的链式风格。
    pub fn iter(&self) -> impl Iterator<Item = ShapeKind> + 'a {
        let v: Vec<ShapeKind> = self
            .slide
            .inner
            .shapes
            .iter()
            .map(crate::shape::wrap)
            .collect();
        v.into_iter()
    }

    /// 按索引取一个形状（克隆为 `ShapeKind`）。
    pub fn get(&self, idx: usize) -> Option<ShapeKind> {
        self.slide.inner.shapes.get(idx).map(crate::shape::wrap)
    }

    /// 取**标题占位符**（若有）。
    ///
    /// 对标 python-pptx `Slide.shapes.title`。
    ///
    /// # 判定策略
    /// 按以下顺序查找第一个命中的形状（顺序与 python-pptx 略有差异，但行为更直观）：
    /// 1. `ph_type == "title"` 或 `ph_type == "ctrTitle"`（居中标题）；
    /// 2. `ph_idx == 0` 的占位符（OOXML 默认 idx=0 即为标题位）；
    /// 3. `name` 含 "title"（不区分大小写）的占位符（兼容手工制作的 slide）。
    ///
    /// 返回 `None` 表示该 slide 没有标题占位符。
    pub fn title(&self) -> Option<ShapeKind> {
        for sh in &self.slide.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    if let Some(t) = &sp.ph_type {
                        if t == "title" || t == "ctrTitle" {
                            return Some(crate::shape::wrap(sh));
                        }
                    }
                    if sp.ph_idx == Some(0) {
                        return Some(crate::shape::wrap(sh));
                    }
                }
                if sp.name.to_ascii_lowercase().contains("title") {
                    return Some(crate::shape::wrap(sh));
                }
            }
        }
        None
    }

    /// 取所有**占位符**形状（按 `ph_idx` 升序排列；同 idx 多个按原顺序）。
    ///
    /// 对标 python-pptx `Slide.placeholders` / `Slide.shapes.placeholders`。
    pub fn placeholders(&self) -> Vec<ShapeKind> {
        let mut out: Vec<(Option<u32>, usize, ShapeKind)> = Vec::new();
        for (i, sh) in self.slide.inner.shapes.iter().enumerate() {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder {
                    out.push((sp.ph_idx, i, crate::shape::wrap(sh)));
                }
            }
        }
        // 按 (ph_idx, i) 排序——保证稳定的遍历顺序
        out.sort_by_key(|(idx, i, _)| (*idx, *i));
        out.into_iter().map(|(_, _, s)| s).collect()
    }

    /// 按 `ph_idx` 查占位符。
    ///
    /// 对标 python-pptx `slide.placeholders[idx]` 的 `__getitem__` 语义。
    /// 返回第一个 `ph_idx == idx` 的占位符；没有则返回 `None`。
    pub fn placeholder(&self, idx: u32) -> Option<ShapeKind> {
        for sh in &self.slide.inner.shapes {
            if let OxmlSlideShape::Sp(sp) = sh {
                if sp.is_placeholder && sp.ph_idx == Some(idx) {
                    return Some(crate::shape::wrap(sh));
                }
            }
        }
        None
    }

    /// 取所有占位符，并从给定版式继承位置/尺寸/填充/边框。
    ///
    /// 对标 python-pptx `slide.placeholders` 的继承语义：当 slide 占位符
    /// 未显式设置 xfrm / fill / line 时，从 layout 中同 `ph_idx` 的占位符继承。
    ///
    /// # 参数
    /// - `layout`：所属 Presentation 中的版式引用（通过 `Presentation::layout_for_slide` 获取）。
    ///
    /// # 返回
    /// 返回的 `ShapeKind::Placeholder` 携带**继承后**的属性快照（clone），
    /// 修改返回值**不会**回写到 slide。如需修改 slide 上的占位符，请用
    /// [`ShapesMut::placeholder_mut`]。
    ///
    /// # 继承规则
    /// 1. 若 slide 占位符 `xfrm.is_empty()`，从 layout 占位符继承 xfrm；
    /// 2. 若 slide 占位符 `fill == Fill::Inherit`，从 layout 占位符继承 fill；
    /// 3. 若 slide 占位符 `line == None`，从 layout 占位符继承 line；
    /// 4. `ph_type` / `ph_idx` 保持 slide 自身值不变。
    pub fn placeholders_inherited(
        &self,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> Vec<ShapeKind> {
        let layout_spans = layout.oxml.borrow();
        let mut out: Vec<(Option<u32>, usize, ShapeKind)> = Vec::new();
        for (i, sh) in self.slide.inner.shapes.iter().enumerate() {
            match sh {
                OxmlSlideShape::Sp(sp) if sp.is_placeholder => {
                    // 在 layout 中查找匹配的占位符（按 ph_idx 优先，其次 ph_type）
                    let layout_ph = find_layout_placeholder(&layout_spans.shapes, sp);
                    let merged = if let Some(lph) = layout_ph {
                        let mut cloned = sp.clone();
                        inherit_placeholder_from_layout(&mut cloned, lph);
                        crate::shape::ShapeKind::Placeholder(crate::shape::PlaceholderShape(
                            crate::shape::AutoShape::from_sp(cloned),
                        ))
                    } else {
                        crate::shape::wrap(sh)
                    };
                    out.push((sp.ph_idx, i, merged));
                }
                // TODO-007：识别 Pic 占位符（图片占位符填充）。
                // Pic 占位符的位置/尺寸在 add_picture_to_placeholder 时已从 layout 继承，
                // 此处不再二次合并，直接以 ShapeKind::Picture 返回。
                OxmlSlideShape::Pic(pic) if pic.is_placeholder => {
                    out.push((pic.ph_idx, i, crate::shape::wrap(sh)));
                }
                // TODO-007：识别 GraphicFrame 占位符（图表/表格占位符填充）。
                // GraphicFrame 占位符的位置/尺寸在 add_chart_to_placeholder /
                // add_table_to_placeholder 时已从 layout 继承，此处不再二次合并，
                // 直接以 ShapeKind::Chart / ShapeKind::Table 返回。
                OxmlSlideShape::GraphicFrame(gf) if gf.is_placeholder => {
                    out.push((gf.ph_idx, i, crate::shape::wrap(sh)));
                }
                _ => {}
            }
        }
        out.sort_by_key(|(idx, i, _)| (*idx, *i));
        out.into_iter().map(|(_, _, s)| s).collect()
    }

    /// 按 `ph_idx` 查占位符，并从给定版式继承位置/尺寸/填充/边框。
    ///
    /// 与 [`Shapes::placeholders_inherited`] 的区别：仅返回指定 idx 的占位符。
    pub fn placeholder_inherited(
        &self,
        idx: u32,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> Option<ShapeKind> {
        let _layout_spans = layout.oxml.borrow();
        for sh in &self.slide.inner.shapes {
            match sh {
                OxmlSlideShape::Sp(sp) if sp.is_placeholder && sp.ph_idx == Some(idx) => {
                    let layout_ph = find_layout_placeholder(&_layout_spans.shapes, sp);
                    let merged = if let Some(lph) = layout_ph {
                        let mut cloned = sp.clone();
                        inherit_placeholder_from_layout(&mut cloned, lph);
                        crate::shape::ShapeKind::Placeholder(crate::shape::PlaceholderShape(
                            crate::shape::AutoShape::from_sp(cloned),
                        ))
                    } else {
                        crate::shape::wrap(sh)
                    };
                    return Some(merged);
                }
                // TODO-007：识别 Pic 占位符（图片占位符填充）。
                OxmlSlideShape::Pic(pic) if pic.is_placeholder && pic.ph_idx == Some(idx) => {
                    return Some(crate::shape::wrap(sh));
                }
                // TODO-007：识别 GraphicFrame 占位符（图表/表格占位符填充）。
                OxmlSlideShape::GraphicFrame(gf) if gf.is_placeholder && gf.ph_idx == Some(idx) => {
                    return Some(crate::shape::wrap(sh));
                }
                _ => {}
            }
        }
        None
    }
}

/// 在 layout 的 shapes 中查找与 slide 占位符匹配的占位符。
///
/// 匹配规则（与 python-pptx 对齐）：
/// 1. `ph_idx` 相同 → 命中（最常见情况）；
/// 2. `ph_type` 相同且 `ph_idx` 都为 None → 命中（兼容无 idx 的占位符）。
fn find_layout_placeholder<'a>(
    layout_shapes: &'a [crate::oxml::shape::Sp],
    slide_sp: &crate::oxml::shape::Sp,
) -> Option<&'a crate::oxml::shape::Sp> {
    // 优先按 ph_idx 匹配
    if let Some(idx) = slide_sp.ph_idx {
        for lsp in layout_shapes {
            if lsp.is_placeholder && lsp.ph_idx == Some(idx) {
                return Some(lsp);
            }
        }
    }
    // 回退：按 ph_type 匹配（当 slide 占位符有 type 但无 idx 时）
    if let Some(pht) = &slide_sp.ph_type {
        for lsp in layout_shapes {
            if lsp.is_placeholder && lsp.ph_type.as_deref() == Some(pht.as_str()) {
                return Some(lsp);
            }
        }
    }
    None
}

/// 把 layout 占位符的属性继承到 slide 占位符（原地修改）。
///
/// 继承规则见 [`Shapes::placeholders_inherited`] 文档。
fn inherit_placeholder_from_layout(
    slide_sp: &mut crate::oxml::shape::Sp,
    layout_sp: &crate::oxml::shape::Sp,
) {
    // 1. xfrm 继承：slide 占位符未设置位置/尺寸时，从 layout 继承
    if slide_sp.properties.xfrm.is_empty() {
        slide_sp.properties.xfrm = layout_sp.properties.xfrm;
    }
    // 2. fill 继承：slide 占位符 fill 为 Inherit 时，从 layout 继承
    if matches!(slide_sp.properties.fill, crate::oxml::sppr::Fill::Inherit) {
        slide_sp.properties.fill = layout_sp.properties.fill.clone();
    }
    // 3. line 继承：slide 占位符 line 为 None 时，从 layout 继承
    if slide_sp.properties.line.is_none() {
        slide_sp.properties.line = layout_sp.properties.line.clone();
    }
    // 4. 几何继承：slide 占位符 geometry 为 None 时，从 layout 继承
    if slide_sp.properties.geometry.is_none() {
        slide_sp.properties.geometry = layout_sp.properties.geometry.clone();
    }
}

/// 可变形状视图。
///
/// 通过 `slide.shapes_mut()` 获取，是 `add_*` / `remove` 唯一入口。
#[derive(Debug)]
pub struct ShapesMut<'a> {
    slide: &'a mut Slide,
}

impl<'a> ShapesMut<'a> {
    /// 形状数量。
    pub fn len(&self) -> usize {
        self.slide.inner.shapes.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.slide.inner.shapes.is_empty()
    }

    /// 遍历（**只读克隆**）。
    pub fn iter(&self) -> impl Iterator<Item = ShapeKind> + '_ {
        let v: Vec<ShapeKind> = self
            .slide
            .inner
            .shapes
            .iter()
            .map(crate::shape::wrap)
            .collect();
        v.into_iter()
    }

    /// 按索引取（克隆）。
    pub fn get(&self, idx: usize) -> Option<ShapeKind> {
        self.slide.inner.shapes.get(idx).map(crate::shape::wrap)
    }

    /// 按索引移除一个形状。
    ///
    /// 返回被移除形状的高阶克隆（`None` 表示 idx 越界）。
    /// 移除后，**后续 shape 的索引会前移**。
    pub fn remove(&mut self, idx: usize) -> Option<ShapeKind> {
        if idx < self.slide.inner.shapes.len() {
            let v = self.slide.inner.shapes.remove(idx);
            Some(crate::shape::wrap(&v))
        } else {
            None
        }
    }

    /// 从版式占位符定义创建一个新的 slide 占位符（TODO-007）。
    ///
    /// 对标 python-pptx "从模板创建 slide 后填充占位符"的工作流。
    ///
    /// 本方法会：
    /// 1. 在 layout 的 shapes 中查找 `ph_idx` 匹配的占位符；
    /// 2. 克隆该 layout 占位符作为 slide 占位符的初始状态（继承位置/尺寸/格式）；
    /// 3. 分配新的 shape id，清空文本内容（等待调用方填充）；
    /// 4. 追加到 slide 的 shapes 末尾。
    ///
    /// # 参数
    /// - `ph_idx`：要创建的占位符 idx（对应 `<p:ph idx="N"/>`）；
    /// - `layout`：所属 Presentation 中的版式引用。
    ///
    /// # 返回
    /// - `Ok(AutoShape)`：创建成功，返回占位符的高阶视图；
    /// - `Err(Error::Other)`：layout 中未找到 `ph_idx` 匹配的占位符。
    pub fn add_placeholder_from_layout(
        &mut self,
        ph_idx: u32,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> crate::Result<crate::shape::AutoShape> {
        let layout_borrow = layout.oxml.borrow();
        // 在 layout 中查找匹配 ph_idx 的占位符
        let mut found: Option<crate::oxml::shape::Sp> = None;
        for lsp in &layout_borrow.shapes {
            if lsp.is_placeholder && lsp.ph_idx == Some(ph_idx) {
                found = Some(lsp.clone());
                break;
            }
        }
        let mut sp = found.ok_or_else(|| {
            crate::Error::Other(format!("layout 中未找到 ph_idx={} 的占位符", ph_idx))
        })?;
        // 分配新 id，清空文本（等待调用方填充）
        sp.id = self.slide.next_shape_id();
        sp.text = crate::oxml::txbody::TextBody::new();
        let auto = crate::shape::AutoShape::from_sp(sp.clone());
        self.slide.inner.shapes.push(OxmlSlideShape::Sp(sp));
        Ok(auto)
    }

    /// 添加一个**空**文本框。
    ///
    /// 文本内容留空，需要后续 `text_frame_mut().set_text(...)`。
    /// 位置/尺寸参数接受任意实现 [`EmuExt`] 的类型（`Inches` / `Cm` / `Emu` / ...）。
    pub fn add_textbox<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<TextBox> {
        let mut tb = TextBox::new(format!("TextBox {}", self.slide.inner.shapes.len() + 1));
        tb.set_left(left.emu());
        tb.set_top(top.emu());
        tb.set_width(width.emu());
        tb.set_height(height.emu());
        tb.set_id(self.slide.next_shape_id());
        let sp = tb.shape.sp.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Sp(sp));
        Ok(tb)
    }

    /// 添加一个文本框，**并直接写入文本**（一步完成，避免 clone-sp 与 sp 失同步）。
    ///
    /// 比 `add_textbox` + 后续 `set_text` 略快，且不易写错。
    pub fn add_textbox_with_text<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        left: L,
        top: T,
        width: W,
        height: H,
        text: &str,
    ) -> crate::Result<TextBox> {
        let mut tb = TextBox::new(format!("TextBox {}", self.slide.inner.shapes.len() + 1));
        tb.set_left(left.emu());
        tb.set_top(top.emu());
        tb.set_width(width.emu());
        tb.set_height(height.emu());
        tb.set_id(self.slide.next_shape_id());
        tb.set_text(text);
        // set_text 修改的是 tb.shape.sp.text；现在再 clone 推入
        let sp = tb.shape.sp.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Sp(sp));
        Ok(tb)
    }

    /// 从本地路径添加一张图片。
    ///
    /// 自动推断 Content-Type 与扩展名（png / jpg / gif / bmp / svg 等）。
    ///
    /// 内部会自动：
    /// - 分配全局唯一的 `rIdImgN` 关系 id；
    /// - 注册 `/ppt/media/imageN.<ext>` part（保存时被写入 zip）；
    /// - 在 `slideN.xml.rels` 中添加 `Image` 关系。
    pub fn add_picture<P: AsRef<std::path::Path>, L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        path: P,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<Picture> {
        let mut pic = Picture::from_path(path.as_ref())?;
        pic.set_left(left.emu());
        pic.set_top(top.emu());
        pic.set_width(width.emu());
        pic.set_height(height.emu());
        pic.set_id(self.slide.next_shape_id());
        // 给 pic 设置默认 name（避免 cNvPr name 为空，PowerPoint 严格模式下会报警）
        if pic.name().is_empty() {
            pic.set_name(format!("Picture {}", self.slide.inner.shapes.len() + 1));
        }
        // 分配图片关系 id（rIdImgN 命名空间，与 layout 区分）
        let rid = self.slide.allocate_image_rid();
        // 同步写到 oxml
        pic.pic_mut().rid = rid.clone();
        // 注册 media 到 Presentation
        let partname = crate::opc::part::new_part_name(
            format!(
                "/ppt/media/image{}.{}",
                self.slide.next_media_index(),
                pic.ext.trim_start_matches('.')
            )
            .as_str(),
        );
        let ct = crate::shape::picture::content_type_for(&pic.ext);
        let blob = pic.blob.clone().unwrap_or_default();
        self.slide.register_media(MediaEntry {
            partname,
            content_type: ct.to_string(),
            blob,
            rid: rid.clone(),
        });
        let oxml_pic = pic.pic.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Pic(oxml_pic));
        Ok(pic)
    }

    /// 添加一张图片并标记为占位符填充（TODO-007 高阶 API）。
    ///
    /// 与 [`add_picture`](Self::add_picture) 的区别：
    /// - 本方法创建的 `<p:pic>` 会带 `<p:ph type="pic" idx="N"/>` 标记，
    ///   PowerPoint 会把它识别为"占位符填充"而非自由图片；
    /// - 位置 / 尺寸**自动从版式占位符继承**（`layout` 中 `ph_idx == ph_idx` 的占位符），
    ///   调用方无需手动指定 left/top/width/height；
    /// - 若版式中找不到匹配的占位符，回退到 `left=0, top=0, width=9144000, height=6858000`
    ///   （默认 10" × 7.5"），并仍然写出占位符标记。
    ///
    /// # 参数
    /// - `ph_idx`：占位符 idx（对应版式中 `<p:ph idx="N"/>` 的 N）；
    /// - `path`：图片文件路径；
    /// - `layout`：当前 slide 关联的版式引用（用于继承位置/尺寸）。
    ///
    /// # 返回值
    /// - 成功：返回 [`Picture`]（已设置占位符标记）；
    /// - 失败：文件读取失败返回 [`crate::Error::Io`]。
    ///
    /// # 与 python-pptx 的对应
    ///
    /// python-pptx 中 `slide.shapes.add_picture(path, ...)` 不会自动绑定占位符；
    /// 用户需要先 `slide.placeholders[idx]` 取占位符再 `placeholder.insert_picture(path)`。
    /// 本方法把两步合并为一个原子操作，语义更清晰。
    ///
    /// # 示例
    /// ```no_run
    /// use pptx::Presentation;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// // 先取版式（避免与 slides_mut 的可变借用冲突）
    /// let layout = p.slide_layouts().get(0).cloned().unwrap();
    /// let counter = p.id_counter();
    /// let s = p.slides_mut().add_slide(counter).unwrap();
    /// // 假设版式有 idx=10 的图片占位符
    /// let pic = s.shapes_mut().add_picture_to_placeholder(10, "logo.png", &layout).unwrap();
    /// ```
    pub fn add_picture_to_placeholder<P: AsRef<std::path::Path>>(
        &mut self,
        ph_idx: u32,
        path: P,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> crate::Result<Picture> {
        // 从版式中查找匹配的占位符，取其位置/尺寸
        let (left, top, width, height) = {
            let layout_ref = layout.oxml.borrow();
            let mut found = None;
            for lsp in layout_ref.shapes.iter() {
                if lsp.is_placeholder && lsp.ph_idx == Some(ph_idx) {
                    found = Some((
                        lsp.properties.xfrm.off_x.unwrap_or_default(),
                        lsp.properties.xfrm.off_y.unwrap_or_default(),
                        lsp.properties
                            .xfrm
                            .ext_cx
                            .unwrap_or(crate::units::Emu(9_144_000)),
                        lsp.properties
                            .xfrm
                            .ext_cy
                            .unwrap_or(crate::units::Emu(6_858_000)),
                    ));
                    break;
                }
            }
            found.unwrap_or((
                crate::units::Emu::default(),
                crate::units::Emu::default(),
                crate::units::Emu(9_144_000),
                crate::units::Emu(6_858_000),
            ))
        };
        // 复用 add_picture 创建基础图片
        let mut pic = self.add_picture(path, left, top, width, height)?;
        // 标记为占位符（type="pic"）
        pic.set_placeholder(ph_idx, Some("pic"));
        // 同步到 oxml（add_picture 已经 push 了 oxml_pic，需要更新最后一个 Pic）
        if let Some(OxmlSlideShape::Pic(last_pic)) = self.slide.inner.shapes.last_mut() {
            last_pic.is_placeholder = true;
            last_pic.ph_idx = Some(ph_idx);
            last_pic.ph_type = Some("pic".to_string());
        }
        Ok(pic)
    }

    /// 添加一个图表并绑定到指定占位符（TODO-007 图表占位符类型化填充）。
    ///
    /// 与 [`Self::add_chart`] 的区别：本方法会从 `layout` 中查找 `ph_idx` 对应的
    /// 占位符，继承其位置/尺寸，并把生成的 graphicFrame 标记为占位符
    /// （写出 `<p:ph type="chart" idx="..."/>`）。
    ///
    /// # 参数
    /// - `ph_idx`：占位符索引（对应版式中 `<p:ph idx="..."/>`）。
    /// - `chart_type`：图表类型（柱/条/线/饼）。
    /// - `data`：图表数据（类别 + 系列 + 可选标题）。
    /// - `layout`：所属 Presentation 中的版式引用。
    ///
    /// # 返回值
    /// - 成功：返回 [`ChartShape`]（已设置占位符标记 + 继承的位置/尺寸）。
    ///
    /// # 与 python-pptx 的对应
    ///
    /// python-pptx 中通过 `slide.placeholders[idx].insert_chart(...)` 实现图表占位符填充；
    /// 本方法把"取占位符 + 创建图表 + 绑定"合并为一个原子操作。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::*;
    /// # use pptx::oxml::chart::{ChartData, ChartSeries, ChartCategory, ChartType};
    /// # let mut prs = Presentation::new().unwrap();
    /// # let layout = prs.slide_layouts().get(0).cloned().unwrap();
    /// # let counter = prs.id_counter();
    /// # let s = prs.slides_mut().add_slide(counter).unwrap();
    /// let mut data = ChartData::default();
    /// data.categories = vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")];
    /// data.series = vec![ChartSeries::new("Sales", vec![10.0, 20.0])];
    /// // 假设版式有 idx=12 的图表占位符
    /// let chart = s.shapes_mut().add_chart_to_placeholder(
    ///     12, ChartType::Column, data, &layout,
    /// ).unwrap();
    /// ```
    pub fn add_chart_to_placeholder(
        &mut self,
        ph_idx: u32,
        chart_type: crate::oxml::chart::ChartType,
        data: crate::oxml::chart::ChartData,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> crate::Result<ChartShape> {
        // 从版式中查找匹配的占位符，取其位置/尺寸
        let (left, top, width, height) = {
            let layout_ref = layout.oxml.borrow();
            let mut found = None;
            for lsp in layout_ref.shapes.iter() {
                if lsp.is_placeholder && lsp.ph_idx == Some(ph_idx) {
                    found = Some((
                        lsp.properties.xfrm.off_x.unwrap_or_default(),
                        lsp.properties.xfrm.off_y.unwrap_or_default(),
                        lsp.properties
                            .xfrm
                            .ext_cx
                            .unwrap_or(crate::units::Emu(9_144_000)),
                        lsp.properties
                            .xfrm
                            .ext_cy
                            .unwrap_or(crate::units::Emu(6_858_000)),
                    ));
                    break;
                }
            }
            found.unwrap_or((
                crate::units::Emu::default(),
                crate::units::Emu::default(),
                crate::units::Emu(9_144_000),
                crate::units::Emu(6_858_000),
            ))
        };
        // 复用 add_chart 创建基础图表
        let mut chart = self.add_chart(chart_type, data, left, top, width, height)?;
        // 标记为占位符（type="chart"）
        chart.set_placeholder(ph_idx, Some("chart"));
        // 同步到 oxml（add_chart 已经 push 了 oxml GraphicFrame，需要更新最后一个）
        if let Some(OxmlSlideShape::GraphicFrame(last_gf)) = self.slide.inner.shapes.last_mut() {
            last_gf.is_placeholder = true;
            last_gf.ph_idx = Some(ph_idx);
            last_gf.ph_type = Some("chart".to_string());
        }
        Ok(chart)
    }

    /// 添加一个表格并绑定到指定占位符（TODO-007 表格占位符类型化填充）。
    ///
    /// 与 [`Self::add_table`] 的区别：本方法会从 `layout` 中查找 `ph_idx` 对应的
    /// 占位符，继承其位置/尺寸，并把生成的 graphicFrame 标记为占位符
    /// （写出 `<p:ph type="tbl" idx="..."/>`）。
    ///
    /// # 参数
    /// - `ph_idx`：占位符索引（对应版式中 `<p:ph idx="..."/>`）。
    /// - `rows` / `cols`：表格行列数。
    /// - `layout`：所属 Presentation 中的版式引用。
    ///
    /// # 返回值
    /// - 成功：返回 [`TableShape`]（已设置占位符标记 + 继承的位置/尺寸）。
    ///
    /// # 与 python-pptx 的对应
    ///
    /// python-pptx 中通过 `slide.placeholders[idx].insert_table(rows, cols)` 实现表格占位符填充；
    /// 本方法把"取占位符 + 创建表格 + 绑定"合并为一个原子操作。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::*;
    /// # let mut prs = Presentation::new().unwrap();
    /// # let layout = prs.slide_layouts().get(0).cloned().unwrap();
    /// # let counter = prs.id_counter();
    /// # let s = prs.slides_mut().add_slide(counter).unwrap();
    /// // 假设版式有 idx=14 的表格占位符
    /// let tbl = s.shapes_mut().add_table_to_placeholder(
    ///     14, 3, 4, &layout,
    /// ).unwrap();
    /// ```
    pub fn add_table_to_placeholder(
        &mut self,
        ph_idx: u32,
        rows: usize,
        cols: usize,
        layout: &crate::slide_layouts::SlideLayoutRef,
    ) -> crate::Result<TableShape> {
        // 从版式中查找匹配的占位符，取其位置/尺寸
        let (left, top, width, height) = {
            let layout_ref = layout.oxml.borrow();
            let mut found = None;
            for lsp in layout_ref.shapes.iter() {
                if lsp.is_placeholder && lsp.ph_idx == Some(ph_idx) {
                    found = Some((
                        lsp.properties.xfrm.off_x.unwrap_or_default(),
                        lsp.properties.xfrm.off_y.unwrap_or_default(),
                        lsp.properties
                            .xfrm
                            .ext_cx
                            .unwrap_or(crate::units::Emu(9_144_000)),
                        lsp.properties
                            .xfrm
                            .ext_cy
                            .unwrap_or(crate::units::Emu(6_858_000)),
                    ));
                    break;
                }
            }
            found.unwrap_or((
                crate::units::Emu::default(),
                crate::units::Emu::default(),
                crate::units::Emu(9_144_000),
                crate::units::Emu(6_858_000),
            ))
        };
        // 复用 add_table 创建基础表格
        let mut tbl = self.add_table(rows, cols, left, top, width, height)?;
        // 标记为占位符（type="tbl"）
        tbl.set_placeholder(ph_idx, Some("tbl"));
        // 同步到 oxml（add_table 已经 push 了 oxml GraphicFrame，需要更新最后一个）
        if let Some(OxmlSlideShape::GraphicFrame(last_gf)) = self.slide.inner.shapes.last_mut() {
            last_gf.is_placeholder = true;
            last_gf.ph_idx = Some(ph_idx);
            last_gf.ph_type = Some("tbl".to_string());
        }
        Ok(tbl)
    }

    /// 添加一个**自选图形**（预设几何）。
    ///
    /// 几何形如 [`crate::oxml::simpletypes::PresetGeometry::Rectangle`] / `Ellipse` /
    /// `RightArrow` / ...，完整列表见 OOXML 规范。
    pub fn add_shape<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        geometry: crate::oxml::simpletypes::PresetGeometry,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<AutoShape> {
        let name = format!("Shape {}", self.slide.inner.shapes.len() + 1);
        let mut s = AutoShape::new(name, geometry);
        s.set_left(left.emu());
        s.set_top(top.emu());
        s.set_width(width.emu());
        s.set_height(height.emu());
        s.set_id(self.slide.next_shape_id());
        let sp = s.sp.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Sp(sp));
        Ok(s)
    }

    /// 添加一个**连接器**。
    ///
    /// # 参数
    /// - `connector_type`：连接器几何（`MSO_CONNECTOR_TYPE`，默认 `MsoConnectorType::Straight`）；
    /// - `begin_x` / `begin_y`：起点 EMU 坐标；
    /// - `end_x` / `end_y`：终点 EMU 坐标；
    ///
    /// # 与 python-pptx 的对应
    /// 对应 `SlideShapes.add_connector(connector_type, begin_x, begin_y, end_x, end_y)`。
    /// 在 pptx-rs 中所有长度参数都实现 [`EmuExt`]，可直接传 `Inches(...)`。
    ///
    /// # 修订历史
    /// 早期签名采用 4 参版（仅 `begin_x, begin_y, end_x, end_y`），后被改写为
    /// `(begin_x, _begin_y_ignored, end_x, _end_y_ignored)`，但实现却把 y 强制为 0，
    /// 导致连接器始终在 y=0 —— 现已**恢复**为 4 参版本并正确使用 y 坐标。
    pub fn add_connector<BX: EmuExt, BY: EmuExt, EX: EmuExt, EY: EmuExt>(
        &mut self,
        connector_type: crate::oxml::simpletypes::MsoConnectorType,
        begin_x: BX,
        begin_y: BY,
        end_x: EX,
        end_y: EY,
    ) -> crate::Result<Connector> {
        let mut c = Connector::new_with_type(
            format!("Connector {}", self.slide.inner.shapes.len() + 1),
            connector_type,
        );
        c.set_left(Emu(0));
        c.set_top(Emu(0));
        c.set_width(Emu(0));
        c.set_height(Emu(0));
        // ✅ 使用真实传入的 y 坐标（不再被强制为 0）
        c.set_begin(crate::units::EmuPoint(
            begin_x.emu().value(),
            begin_y.emu().value(),
        ));
        c.set_end(crate::units::EmuPoint(
            end_x.emu().value(),
            end_y.emu().value(),
        ));
        c.set_id(self.slide.next_shape_id());
        let cx = c.cxn.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::CxnSp(cx));
        Ok(c)
    }

    /// 添加一个**连接器**（完整 4 端点 + 类型版本）。
    ///
    /// 真正对齐 python-pptx `add_connector(connector_type, begin_x, begin_y, end_x, end_y)`。
    ///
    /// # 弃用说明
    /// 此方法自 0.1.x 起**已与 [`Self::add_connector`] 合并**——后者已恢复 4 端点签名。
    /// 保留仅为**兼容**既有调用方；新代码请直接使用 [`Self::add_connector`]。
    #[deprecated(
        since = "0.1.0",
        note = "已与 add_connector 合并；请改用 add_connector(connector_type, begin_x, begin_y, end_x, end_y)"
    )]
    pub fn add_connector_geom<BL: EmuExt, BT: EmuExt, EL: EmuExt, ET: EmuExt>(
        &mut self,
        connector_type: crate::oxml::simpletypes::MsoConnectorType,
        begin_x: BL,
        begin_y: BT,
        end_x: EL,
        end_y: ET,
    ) -> crate::Result<Connector> {
        self.add_connector(connector_type, begin_x, begin_y, end_x, end_y)
    }

    /// 添加一个**自由形**（已完成 [`crate::shape::freeform::FreeformBuilder`] 流程）。
    ///
    /// # 当前实现
    /// 0.1.0 的 `Freeform::build` 内部退化为 `AutoShape` + 矩形 prstGeom——
    /// 真正的 `a:custGeom` 描述（折线/曲线/闭合）将在 0.2.0 接入。
    /// 因此**当前** `add_freeform` 在视觉上与 `add_shape(Rectangle, ...)` 一致，
    /// 但 `Freeform` 的 `points()` 仍可被读取。
    pub fn add_freeform(
        &mut self,
        mut freeform: Freeform,
        left: Emu,
        top: Emu,
        width: Emu,
        height: Emu,
    ) -> crate::Result<Freeform> {
        freeform.set_left(left);
        freeform.set_top(top);
        freeform.set_width(width);
        freeform.set_height(height);
        freeform.set_id(self.slide.next_shape_id());
        // 内部 AutoShape 句柄——推入 sp 列表
        let sp = freeform.shape.sp.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Sp(sp));
        Ok(freeform)
    }

    /// 把 idx 处的形状移到末尾（z-order 顶层）。
    pub fn move_to_front(&mut self, idx: usize) {
        if idx < self.slide.inner.shapes.len() {
            let s = self.slide.inner.shapes.remove(idx);
            self.slide.inner.shapes.push(s);
        }
    }

    /// 把 idx 处的形状移到首位置（z-order 底层）。
    pub fn move_to_back(&mut self, idx: usize) {
        if idx < self.slide.inner.shapes.len() {
            let s = self.slide.inner.shapes.remove(idx);
            self.slide.inner.shapes.insert(0, s);
        }
    }

    /// 把 idx 处的形状向上移动一级（z-order 提升）。
    ///
    /// 对标 python-pptx 中通过 XML 操作调整形状顺序的能力。
    /// 若 idx 已是顶层（最后一个）或越界，则为 no-op。
    pub fn move_up(&mut self, idx: usize) {
        let len = self.slide.inner.shapes.len();
        if idx < len.saturating_sub(1) {
            // 与后一个交换位置
            self.slide.inner.shapes.swap(idx, idx + 1);
        }
    }

    /// 把 idx 处的形状向下移动一级（z-order 降低）。
    ///
    /// 对标 python-pptx 中通过 XML 操作调整形状顺序的能力。
    /// 若 idx 已是底层（第一个）或越界，则为 no-op。
    pub fn move_down(&mut self, idx: usize) {
        if idx > 0 && idx < self.slide.inner.shapes.len() {
            // 与前一个交换位置
            self.slide.inner.shapes.swap(idx, idx - 1);
        }
    }

    /// 找某个 ShapeKind 的索引（克隆比较）。
    ///
    /// 等价 python-pptx 中 `SlideShapes.index(shape)`。
    pub fn index(&self, kind: &ShapeKind) -> Option<usize> {
        // 用名字 + 位置比较（最稳定可观察）。
        let target_name = kind.name().to_string();
        for (i, s) in self.slide.inner.shapes.iter().enumerate() {
            if crate::shape::name_of(s) == target_name {
                return Some(i);
            }
        }
        None
    }

    /// 添加一个**空**组合（后续通过 `Group::children_mut` 加入子形状）。
    #[allow(clippy::field_reassign_with_default)]
    pub fn add_group<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<Group> {
        let name = format!("Group {}", self.slide.inner.shapes.len() + 1);
        let mut g = crate::oxml::shape::Group::default();
        g.id = self.slide.next_shape_id();
        g.name = name;
        g.off = (left.emu(), top.emu());
        g.ext = (width.emu(), height.emu());
        let grp = Group::from_group(g.clone());
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::Group(Box::new(g)));
        Ok(grp)
    }

    /// 添加一个**等分**表格。
    ///
    /// 行高 = `height / rows`，列宽 = `width / cols`。如需不等分请改用
    /// `TableShape::set_row_height` / `set_col_width`。
    pub fn add_table<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        rows: usize,
        cols: usize,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<TableShape> {
        let col_w = Emu(width.emu().value() / (cols as i64).max(1));
        let row_h = Emu(height.emu().value() / (rows as i64).max(1));
        let mut tbl = TableShape::new(rows, cols, col_w, row_h);
        tbl.set_left(left.emu());
        tbl.set_top(top.emu());
        tbl.set_width(width.emu());
        tbl.set_height(height.emu());
        tbl.set_id(self.slide.next_shape_id());
        let frame = tbl.frame.clone();
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::GraphicFrame(frame));
        Ok(tbl)
    }

    /// 添加一个图表（TODO-004 基础图表支持）。
    ///
    /// 当前支持 4 种图表类型：柱状图（[`ChartType::Column`]）、条形图（[`ChartType::Bar`]）、
    /// 折线图（[`ChartType::Line`]）、饼图（[`ChartType::Pie`]）。数据通过 `<c:numCache>`
    /// 内嵌，不依赖嵌入 Excel。
    ///
    /// # 参数
    /// - `chart_type`：图表类型。
    /// - `data`：图表数据（类别 + 系列 + 可选标题）。
    /// - `left` / `top` / `width` / `height`：图表在 slide 中的位置与尺寸。
    ///
    /// # 内部行为
    ///
    /// 1. 创建一个 [`ChartShape`]，设置几何 + id；
    /// 2. 分配一个**本 slide 内**唯一的关系 id `rIdChartN`；
    /// 3. 把 `rIdChartN` 同步写入 `ChartShape.frame.graphic.Chart.rid`，
    ///    供 `<c:chart r:id="rIdChartN"/>` 引用；
    /// 4. 注册 [`ChartEntry`] 到本 slide（保存时由 `to_opc_package` 写出独立的
    ///    `/ppt/charts/chartN.xml` part + `slideN.xml.rels` 关系）；
    /// 5. 把 `ChartShape.frame` 作为 `GraphicFrame` 推入 `spTree`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # use pptx::oxml::chart::{ChartData, ChartSeries, ChartCategory, ChartType};
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let mut data = ChartData::default();
    /// data.categories = vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")];
    /// data.series = vec![ChartSeries::new("Sales", vec![10.0, 20.0])];
    /// data.title = Some("Revenue".to_string());
    /// let _chart = slide.shapes_mut().add_chart(
    ///     ChartType::Column, data,
    ///     Inches(1.0), Inches(2.0), Inches(8.0), Inches(4.0),
    /// ).unwrap();
    /// ```
    pub fn add_chart<L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt>(
        &mut self,
        chart_type: crate::oxml::chart::ChartType,
        data: crate::oxml::chart::ChartData,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<ChartShape> {
        let mut chart = ChartShape::new(chart_type, data);
        chart.set_left(left.emu());
        chart.set_top(top.emu());
        chart.set_width(width.emu());
        chart.set_height(height.emu());
        chart.set_id(self.slide.next_shape_id());
        if chart.name().is_empty() {
            chart.set_name(format!("Chart {}", self.slide.inner.shapes.len() + 1));
        }
        // 分配本 slide 内唯一的关系 id（rIdChartN 命名空间，与 image/notes/comments 区分）
        let rid = self.slide.allocate_chart_rid();
        chart.set_rid(rid.clone());
        // 注册 chart part 元数据；to_opc_package 阶段会用全局索引重新分配 partname，
        // 这里先用 slide 局部索引占位（实际 partname 在打包时确定）。
        let local_idx = self.slide.next_chart_index();
        let partname =
            crate::opc::part::new_part_name(format!("/ppt/charts/chart{}.xml", local_idx).as_str());
        // 取出 chart 的强类型模型克隆一份挂到 ChartEntry。
        // 注意：ChartShape.frame.graphic 与 ChartEntry.chart 是两份独立数据，
        // 修改 ChartShape 后需要重新调用 `chart()` 同步——但本 API 在创建时一次性
        // 注入，后续用户修改 chart 数据需要重新 save 才能反映到 chartN.xml。
        // 后续可考虑用 Rc<RefCell<Chart>> 共享，0.2.x 暂保持 clone 语义。
        let oxml_chart = chart
            .chart()
            .cloned()
            .unwrap_or_else(crate::oxml::chart::Chart::default);
        self.slide.register_chart(ChartEntry {
            partname,
            chart: oxml_chart,
            rid: rid.clone(),
            xlsx_blob: None,
        });
        let frame = chart.frame.clone();
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::GraphicFrame(frame));
        Ok(chart)
    }

    /// 添加一个**带嵌入式 Excel 工作簿**的图表（TODO-004 Excel 嵌入）。
    ///
    /// 与 [`Self::add_chart`] 的区别：本方法额外接受 `xlsx_blob` 参数（有效的
    /// `.xlsx` 文件字节流），保存时会在 `.pptx` 内生成
    /// `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` part + chart part 的
    /// 独立关系文件 `_rels/chartN.xml.rels`（Type=Package），并在 chart XML
    /// 中写入 `<c:externalData r:id="rIdXlsxN"/>` 引用。PowerPoint 打开图表时
    /// 会从该 xlsx part 读取数据源，"编辑数据" 会启动 Excel。
    ///
    /// # 参数
    /// - `chart_type`：图表类型。
    /// - `data`：图表数据（类别 + 系列 + 可选标题）。**仍会**写入 numCache/strCache，
    ///   即使 PowerPoint 不打开 xlsx 也能渲染图表。
    /// - `xlsx_blob`：有效的 `.xlsx` 文件字节流（OOXML SpreadsheetML 包格式）。
    ///   库不校验内容有效性，PowerPoint 会按 zip + XML 解析。
    /// - `left` / `top` / `width` / `height`：图表在 slide 中的位置与尺寸。
    ///
    /// # 返回值
    /// - 成功：返回 [`ChartShape`] 高阶句柄。
    ///
    /// # 内部行为
    ///
    /// 1. 与 [`Self::add_chart`] 一致地构造 `ChartShape` + 分配 `rIdChartN`；
    /// 2. 注册 [`ChartEntry`]，`xlsx_blob` 字段填入 `Some(xlsx_blob.to_vec())`；
    /// 3. 保存时由 `to_opc_package` 写出 xlsx part + chart rels + externalData 引用。
    ///
    /// # 关键约束
    ///
    /// - **xlsx 内容应由调用方保证有效**：库不解析 xlsx，PowerPoint 打开时
    ///   若 xlsx 损坏会报错。
    /// - **xlsx 内数据应与 `data` 一致**：PowerPoint 优先用 numCache 渲染，
    ///   但用户"编辑数据"时会以 xlsx 为准。若两者不一致，编辑后图表会变化。
    /// - **xlsx_rid 由 presentation 层自动分配**：用户无需手动设置。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # use pptx::oxml::chart::{ChartData, ChartSeries, ChartCategory, ChartType};
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// # let xlsx_bytes: Vec<u8> = vec![]; // 实际场景应为有效 .xlsx 文件内容
    /// let mut data = ChartData::default();
    /// data.categories = vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")];
    /// data.series = vec![ChartSeries::new("Sales", vec![10.0, 20.0])];
    /// let _chart = slide.shapes_mut().add_chart_with_excel(
    ///     ChartType::Column, data, xlsx_bytes,
    ///     Inches(1.0), Inches(2.0), Inches(8.0), Inches(4.0),
    /// ).unwrap();
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn add_chart_with_excel<L, T, W, H>(
        &mut self,
        chart_type: crate::oxml::chart::ChartType,
        data: crate::oxml::chart::ChartData,
        xlsx_blob: Vec<u8>,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<ChartShape>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
    {
        let mut chart = ChartShape::new(chart_type, data);
        chart.set_left(left.emu());
        chart.set_top(top.emu());
        chart.set_width(width.emu());
        chart.set_height(height.emu());
        chart.set_id(self.slide.next_shape_id());
        if chart.name().is_empty() {
            chart.set_name(format!("Chart {}", self.slide.inner.shapes.len() + 1));
        }
        let rid = self.slide.allocate_chart_rid();
        chart.set_rid(rid.clone());
        let local_idx = self.slide.next_chart_index();
        let partname =
            crate::opc::part::new_part_name(format!("/ppt/charts/chart{}.xml", local_idx).as_str());
        let oxml_chart = chart
            .chart()
            .cloned()
            .unwrap_or_else(crate::oxml::chart::Chart::default);
        // 关键差异：xlsx_blob 字段填入 Some(...)，触发 to_opc_package 写出
        // xlsx part + chart rels + externalData 引用。
        self.slide.register_chart(ChartEntry {
            partname,
            chart: oxml_chart,
            rid: rid.clone(),
            xlsx_blob: Some(xlsx_blob),
        });
        let frame = chart.frame.clone();
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::GraphicFrame(frame));
        Ok(chart)
    }

    /// 在当前幻灯片上嵌入一个 OLE 对象（TODO-043）。
    ///
    /// 对标 python-pptx 0.6.19+ 的 `shapes.add_ole_object()`。把指定文件
    /// 作为 OLE 复合文档嵌入到 `.pptx` 中，PowerPoint 双击时会调用对应
    /// OLE 服务器（由 `prog_id` 决定）打开编辑。
    ///
    /// # 参数
    /// - `path`：OLE 文件路径（如 `.xls` / `.doc` / `.bin`）。
    ///   文件内容会以原始字节写入 `/ppt/embeddings/oleObjectN.bin`。
    /// - `prog_id`：OLE 程序标识符（如 `"Excel.Sheet.12"` / `"Word.Document.12"` /
    ///   `"Package"`）。PowerPoint 通过 progId 决定双击时调用哪个 OLE 服务器。
    /// - `name`：显示名（在 PowerPoint 中作为对象名，如 `"Worksheet"`）。
    /// - `left` / `top` / `width` / `height`：OLE 对象在 slide 上的位置与尺寸（EMU）。
    ///
    /// # 返回值
    /// - 成功：返回 [`OleObjectShape`] 高阶句柄，可继续调整 `show_as_icon` /
    ///   `set_image_rid`（图标图片）等属性。
    /// - 失败：返回 [`crate::Error::Io`]（文件读取失败）。
    ///
    /// # 内部流程
    ///
    /// 1. 读取文件内容为 `Vec<u8>`；
    /// 2. 创建一个 [`OleObjectShape`]，设置几何 + id + progId + name；
    /// 3. 分配一个**本 slide 内**唯一的关系 id `rIdOleN`；
    /// 4. 把 `rIdOleN` 同步写入 `OleObjectShape.frame.graphic.OleObject.rid`，
    ///    供 `<p:oleObj r:id="rIdOleN"/>` 引用；
    /// 5. 注册 [`OleEntry`] 到本 slide（保存时由 `to_opc_package` 写出独立的
    ///    `/ppt/embeddings/oleObjectN.bin` part + `slideN.xml.rels` 关系）；
    /// 6. 把 `OleObjectShape.frame` 作为 `GraphicFrame` 推入 `spTree`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let _ole = slide.shapes_mut().add_ole_object(
    ///     "data.xlsx", "Excel.Sheet.12", "Worksheet",
    ///     Inches(1.0), Inches(2.0), Inches(4.0), Inches(3.0),
    /// ).unwrap();
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn add_ole_object<L, T, W, H, P>(
        &mut self,
        path: P,
        prog_id: &str,
        name: &str,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<OleObjectShape>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
        P: AsRef<std::path::Path>,
    {
        // 1) 读取 OLE 文件二进制内容（io::Error 由 `#[from]` 自动转为 Error::Io）
        let blob = std::fs::read(path.as_ref())?;

        // 2) 创建 OleObjectShape
        let mut ole = OleObjectShape::new(prog_id, name);
        ole.set_left(left.emu());
        ole.set_top(top.emu());
        ole.set_width(width.emu());
        ole.set_height(height.emu());
        ole.set_id(self.slide.next_shape_id());
        if ole.name().is_empty() {
            ole.set_name(format!("OLE Object {}", self.slide.inner.shapes.len() + 1));
        }
        // 图标默认尺寸与对象尺寸一致（PowerPoint 默认行为）
        ole.set_icon_size(width.emu(), height.emu());
        // 图标 Pic 形状的 id 与 OleObjectShape 的 id 区分（避免重复）
        ole.set_pic_id_name(self.slide.next_shape_id(), "OLE Icon");

        // 3) 分配本 slide 内唯一的 OLE 关系 id
        let rid = self.slide.allocate_ole_rid();
        ole.set_rid(rid.clone());

        // 4) 注册 OLE part 元数据（to_opc_package 阶段会用全局索引重新分配 partname）
        let local_idx = self.slide.next_ole_index();
        let partname = crate::opc::part::new_part_name(
            format!("/ppt/embeddings/oleObject{}.bin", local_idx).as_str(),
        );
        self.slide.register_ole(OleEntry {
            partname,
            blob,
            rid: rid.clone(),
        });

        // 5) 推入 spTree
        let frame = ole.frame.clone();
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::GraphicFrame(frame));
        Ok(ole)
    }

    /// 在当前幻灯片上创建一个 **SmartArt** 图形（从 4 份原始 XML 字符串，TODO-037 创建 API）。
    ///
    /// 这是"逃生舱"入口，适合用户已有 4 份 diagram XML（如从其它 `.pptx`
    /// 复制、或手工构造）的场景。若希望通过结构化模型程序化构建，请使用
    /// [`Self::add_smartart`]。
    ///
    /// # 参数
    /// - `data_xml`：`<dgm:dataModel>` 完整 XML（对应 `/ppt/diagrams/dataN.xml`）。
    /// - `layout_xml`：`<dgm:layoutDef>` 完整 XML（对应 `/ppt/diagrams/layoutN.xml`）。
    /// - `quick_style_xml`：`<dgm:styleData>` 完整 XML（对应 `/ppt/diagrams/quickStylesN.xml`）。
    /// - `colors_xml`：`<dgm:colorsDef>` 完整 XML（对应 `/ppt/diagrams/colorsN.xml`）。
    /// - `left` / `top` / `width` / `height`：SmartArt 在 slide 上的位置与尺寸（EMU）。
    ///
    /// # 返回值
    /// - 成功：返回 [`SmartArtShape`] 高阶句柄，可继续调整位置 / 占位符等属性。
    /// - 失败：返回 [`crate::Error`]（目前为不可失败，但保留 `Result` 以便未来扩展）。
    ///
    /// # 内部流程
    ///
    /// 1. 调用 [`Slide::allocate_diagram_rids`] 分配本 slide 内唯一的 4 个关系 id
    ///    （`rIdDgmDataN` / `rIdDgmLayoutN` / `rIdDgmQsN` / `rIdDgmColorsN`）；
    /// 2. 用 4 个 rid 构造 [`SmartArtShape`]（内部通过 [`crate::oxml::shape::SmartArtRef::from_rids`]
    ///    生成 `<a:graphicData><dgm:relIds .../></a:graphicData>`）；
    /// 3. 调用 [`Slide::next_diagram_index`] 取局部索引，构造 4 个 partname 占位
    ///    （`to_opc_package` 阶段会用全局索引重新分配，覆盖此处的占位）；
    /// 4. 注册 [`DiagramEntry`] 到本 slide（保存时由 `to_opc_package` 写出 4 个
    ///    diagram part + slide rels 的 4 个关系）；
    /// 5. 推入 `spTree`。
    ///
    /// # 关键约束
    ///
    /// - **partname 占位**：本方法生成的 4 个 partname 仅用于 `DiagramEntry` 字段
    ///   占位，`to_opc_package` 写出时会用**全局索引**重新分配（避免多 slide 冲突）。
    /// - **rid 全程不变**：4 个 rid 一旦分配就贯穿 slide XML / slideN.xml.rels /
    ///   diagram parts 引用链，不会被重写。
    /// - **XML 透传**：4 份 XML 字符串**原样**写入 zip part，不做任何解析或重建
    ///   （byte-exact round-trip 友好）。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # use pptx::shape::Shape;
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let data_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    /// <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
    ///                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    ///   <dgm:ptLst><dgm:pt modelId="1" type="doc"/></dgm:ptLst>
    /// </dgm:dataModel>"#;
    /// // ... 其余 3 份 XML 省略
    /// # let layout_xml = "<dgm:layoutDef/>";
    /// # let quick_style_xml = "<dgm:styleData/>";
    /// # let colors_xml = "<dgm:colorsDef/>";
    /// let sa = slide.shapes_mut().add_smartart_from_xml(
    ///     data_xml, layout_xml, quick_style_xml, colors_xml,
    ///     Inches(1.0), Inches(1.0), Inches(8.0), Inches(4.0),
    /// ).unwrap();
    /// assert_eq!(sa.shape_type(), "smart_art");
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn add_smartart_from_xml<L, T, W, H>(
        &mut self,
        data_xml: &str,
        layout_xml: &str,
        quick_style_xml: &str,
        colors_xml: &str,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<SmartArtShape>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
    {
        // 1) 分配 4 个关系 id（本 slide 内唯一）
        let (dm_rid, lo_rid, qs_rid, cs_rid) = self.slide.allocate_diagram_rids();

        // 2) 构造 SmartArtShape（内部生成 raw_xml）
        let mut sa = SmartArtShape::from_rids(&dm_rid, &lo_rid, &qs_rid, &cs_rid);
        sa.set_left(left.emu());
        sa.set_top(top.emu());
        sa.set_width(width.emu());
        sa.set_height(height.emu());
        sa.set_id(self.slide.next_shape_id());
        if sa.name().is_empty() {
            sa.set_name(format!("SmartArt {}", self.slide.inner.shapes.len() + 1));
        }

        // 3) 取局部索引，构造 4 个 partname 占位（to_opc_package 阶段会重写）
        let local_idx = self.slide.next_diagram_index();
        let data_partname = crate::opc::part::new_part_name(
            format!("/ppt/diagrams/data{}.xml", local_idx).as_str(),
        );
        let layout_partname = crate::opc::part::new_part_name(
            format!("/ppt/diagrams/layout{}.xml", local_idx).as_str(),
        );
        let quick_style_partname = crate::opc::part::new_part_name(
            format!("/ppt/diagrams/quickStyles{}.xml", local_idx).as_str(),
        );
        let colors_partname = crate::opc::part::new_part_name(
            format!("/ppt/diagrams/colors{}.xml", local_idx).as_str(),
        );

        // 4) 注册 DiagramEntry
        self.slide.register_diagram(DiagramEntry {
            data_partname,
            layout_partname,
            quick_style_partname,
            colors_partname,
            data_xml: data_xml.to_string(),
            layout_xml: layout_xml.to_string(),
            quick_style_xml: quick_style_xml.to_string(),
            colors_xml: colors_xml.to_string(),
            data_rid: dm_rid.clone(),
            layout_rid: lo_rid.clone(),
            quick_style_rid: qs_rid.clone(),
            colors_rid: cs_rid.clone(),
        });

        // 5) 推入 spTree
        let frame = sa.frame.clone();
        self.slide
            .inner
            .shapes
            .push(OxmlSlideShape::GraphicFrame(frame));
        Ok(sa)
    }

    /// 在当前幻灯片上创建一个 **SmartArt** 图形（从结构化模型，TODO-037 创建 API）。
    ///
    /// 这是高阶友好入口，适合程序化构建 SmartArt。内部调用 4 个结构化模型的
    /// `to_xml()` 生成 XML 字符串，再委托给 [`Self::add_smartart_from_xml`]。
    ///
    /// # 参数
    /// - `data_model`：数据模型（`<dgm:dataModel>`，含 points + connections）。
    /// - `layout_def`：布局定义（`<dgm:layoutDef>`，含 layoutNode 子树）。
    /// - `quick_style`：样式定义（`<dgm:styleData>`，含 styleLbl 列表）。
    /// - `colors`：颜色定义（`<dgm:colorsDef>`，含 styleClrLbl 列表）。
    /// - `left` / `top` / `width` / `height`：位置与尺寸（EMU）。
    ///
    /// # 返回值
    /// - 成功：返回 [`SmartArtShape`] 高阶句柄。
    /// - 失败：返回 [`crate::Error`]（来自 4 个 `to_xml()` 调用，目前不会失败）。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// use pptx::oxml::diagram::{DataModel, DataModelPoint, LayoutDef, QuickStyleDef, ColorsDef};
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let mut dm = DataModel::default();
    /// dm.points.push(DataModelPoint {
    ///     model_id: 1,
    ///     pt_type: Some("doc".into()),
    ///     text: Some("根节点".into()),
    ///     ..Default::default()
    /// });
    /// let sa = slide.shapes_mut().add_smartart(
    ///     dm,
    ///     LayoutDef::default(),
    ///     QuickStyleDef::default(),
    ///     ColorsDef::default(),
    ///     Inches(1.0), Inches(1.0), Inches(8.0), Inches(4.0),
    /// ).unwrap();
    /// assert_eq!(sa.dm_rid().starts_with("rIdDgmData"), true);
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn add_smartart<L, T, W, H>(
        &mut self,
        data_model: crate::oxml::diagram::DataModel,
        layout_def: crate::oxml::diagram::LayoutDef,
        quick_style: crate::oxml::diagram::QuickStyleDef,
        colors: crate::oxml::diagram::ColorsDef,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<SmartArtShape>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
    {
        let data_xml = data_model.to_xml();
        let layout_xml = layout_def.to_xml();
        let quick_style_xml = quick_style.to_xml();
        let colors_xml = colors.to_xml();
        self.add_smartart_from_xml(
            &data_xml,
            &layout_xml,
            &quick_style_xml,
            &colors_xml,
            left,
            top,
            width,
            height,
        )
    }

    /// 在当前幻灯片上嵌入一个**视频**形状（TODO-033）。
    ///
    /// 对标 python-pptx 0.6.19+ 的 `shapes.add_movie()`。把指定视频文件
    /// 嵌入到 `.pptx` 中，PowerPoint 双击视频区域时会播放该视频。
    ///
    /// # 参数
    /// - `video_path`：视频文件路径（如 `.mp4`）。文件内容会以原始字节
    ///   写入 `/ppt/media/mediaN.mp4`。
    /// - `poster_path`：海报帧图片路径（视频未播放时显示的静态画面，如 `.png` / `.jpg`）。
    ///   若为 `None`，PowerPoint 会显示空白占位（推荐传入代表视频首帧的图片）。
    /// - `left` / `top` / `width` / `height`：视频形状在 slide 上的位置与尺寸（EMU）。
    ///
    /// # 返回值
    /// - 成功：返回 [`Picture`] 高阶句柄（已标记为视频形状），可继续调整裁剪 /
    ///   填充模式等图片属性。
    /// - 失败：返回 [`crate::Error::Io`]（视频或海报文件读取失败）。
    ///
    /// # 内部流程
    ///
    /// 1. 读取视频文件内容为 `Vec<u8>`；
    /// 2. 读取海报帧图片内容（若 `poster_path` 为 `None`，用 1x1 透明 PNG 占位）；
    /// 3. 手动构造 [`Picture`] 并注册海报帧 `MediaEntry`（`rIdImgN` + `imageN.png` part）；
    /// 4. 分配**本 slide 内**唯一的视频关系 id `rIdVideoN`；
    /// 5. 调用 `pic.set_video(rIdVideoN)` 把图片标记为视频形状（写出 `<a:videoFile r:link="rIdVideoN"/>`）；
    /// 6. 注册 [`VideoEntry`] 到本 slide（保存时由 `to_opc_package` 写出独立的
    ///    `/ppt/media/mediaN.mp4` part + `slideN.xml.rels` Video 关系）。
    ///
    /// # 关键约束
    ///
    /// - **r:embed vs r:link**：海报帧图片用 `r:embed`（嵌入图片），视频文件用 `r:link`（外部链接方式）；
    /// - **partname 命名**：视频文件 `/ppt/media/mediaN.mp4`，海报帧图片 `/ppt/media/imageN.png`，
    ///   两者命名空间不同但同在 `/ppt/media/` 目录下；
    /// - **关系类型**：视频用 `.../relationships/video`，海报帧用 `.../relationships/image`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let _video = slide.shapes_mut().add_video(
    ///     "intro.mp4", Some("poster.png"),
    ///     Inches(1.0), Inches(1.0), Inches(6.0), Inches(4.0),
    /// ).unwrap();
    /// ```
    pub fn add_video<L, T, W, H, P>(
        &mut self,
        video_path: P,
        poster_path: Option<P>,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<Picture>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
        P: AsRef<std::path::Path>,
    {
        // 1) 读取视频文件二进制内容
        let video_blob = std::fs::read(video_path.as_ref())?;

        // 2) 取海报帧图片：若调用方未提供，用一个 1x1 透明 PNG 占位。
        //    （PowerPoint 在没有海报帧时会显示空白，这里给出最小合法 PNG 避免渲染异常）
        let poster_pic = if let Some(poster) = poster_path.as_ref() {
            Picture::from_path(poster.as_ref())?
        } else {
            // 1x1 透明 PNG（67 字节，标准 base64: iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/IQ7AAAAAElFTkSuQmCC）
            const TRANSPARENT_PNG_1X1: &[u8] = &[
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
                0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
                0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
                0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
                0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
            ];
            Picture::from_bytes(TRANSPARENT_PNG_1X1.to_vec(), ".png")
        };

        // 3) 直接构造 Picture（不走 add_picture，因为它强制从文件路径读取，
        //    而本方法的 poster_path 可能为 None，需要用内置占位 PNG）
        let mut pic = poster_pic;
        pic.set_left(left.emu());
        pic.set_top(top.emu());
        pic.set_width(width.emu());
        pic.set_height(height.emu());
        pic.set_id(self.slide.next_shape_id());
        if pic.name().is_empty() {
            pic.set_name(format!("Video {}", self.slide.inner.shapes.len() + 1));
        }
        // 4a) 海报帧图片关系 id（rIdImgN 命名空间）
        let poster_rid = self.slide.allocate_image_rid();
        pic.pic_mut().rid = poster_rid.clone();
        // 4b) 注册海报帧图片 MediaEntry（imageN.png part + Image 关系）
        let poster_partname = crate::opc::part::new_part_name(
            format!(
                "/ppt/media/image{}.{}",
                self.slide.next_media_index(),
                pic.ext.trim_start_matches('.')
            )
            .as_str(),
        );
        let poster_ct = crate::shape::picture::content_type_for(&pic.ext);
        let poster_blob = pic.blob.clone().unwrap_or_default();
        self.slide.register_media(MediaEntry {
            partname: poster_partname,
            content_type: poster_ct.to_string(),
            blob: poster_blob,
            rid: poster_rid.clone(),
        });

        // 5) 分配视频关系 id（rIdVideoN 命名空间，与海报帧 image 区分）
        let video_rid = self.slide.allocate_video_rid();
        pic.set_video(video_rid.clone());

        // 6) 注册 VideoEntry（to_opc_package 阶段会用全局索引重新分配 partname）
        let local_idx = self.slide.next_video_index();
        let video_partname =
            crate::opc::part::new_part_name(format!("/ppt/media/media{}.mp4", local_idx).as_str());
        self.slide.register_video(VideoEntry {
            partname: video_partname,
            blob: video_blob,
            rid: video_rid.clone(),
        });

        // 7) 推入 spTree
        let oxml_pic = pic.pic.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Pic(oxml_pic));
        Ok(pic)
    }

    /// 在当前幻灯片上嵌入一个**音频**形状（TODO-033）。
    ///
    /// 对标 python-pptx 0.6.19+ 的 `shapes.add_audio()`。把指定音频文件
    /// 嵌入到 `.pptx` 中，PowerPoint 双击音频形状时会播放该音频。
    ///
    /// # 参数
    /// - `audio_path`：音频文件路径（如 `.mp3`）。文件内容会以原始字节
    ///   写入 `/ppt/media/mediaN.mp3`。
    /// - `poster_path`：海报帧图片路径（音频未播放时显示的图标，如 `.png` / `.jpg`）。
    ///   若为 `None`，PowerPoint 会显示默认音频图标（推荐传入代表音频的图标图片）。
    /// - `left` / `top` / `width` / `height`：音频形状在 slide 上的位置与尺寸（EMU）。
    ///
    /// # 返回值
    /// - 成功：返回 [`Picture`] 高阶句柄（已标记为音频形状）；
    /// - 失败：返回 [`crate::Error::Io`]（音频或海报文件读取失败）。
    ///
    /// # 与 [`Self::add_video`] 的差异
    /// - `add_video` 写出 `<a:videoFile r:link="..."/>`，partname 后缀 `.mp4`；
    /// - `add_audio` 写出 `<a:audioFile r:link="..."/>`，partname 后缀 `.mp3`；
    /// - 关系类型分别为 `.../video` 与 `.../audio`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::*;
    /// # let mut prs = Presentation::new().unwrap();
    /// # let counter = prs.id_counter();
    /// # let slide = prs.slides_mut().add_slide(counter).unwrap();
    /// let _audio = slide.shapes_mut().add_audio(
    ///     "bgm.mp3", Some("speaker.png"),
    ///     Inches(1.0), Inches(1.0), Inches(2.0), Inches(2.0),
    /// ).unwrap();
    /// ```
    pub fn add_audio<L, T, W, H, P>(
        &mut self,
        audio_path: P,
        poster_path: Option<P>,
        left: L,
        top: T,
        width: W,
        height: H,
    ) -> crate::Result<Picture>
    where
        L: EmuExt,
        T: EmuExt,
        W: EmuExt,
        H: EmuExt,
        P: AsRef<std::path::Path>,
    {
        // 1) 读取音频文件二进制内容
        let audio_blob = std::fs::read(audio_path.as_ref())?;

        // 2) 取海报帧图片（同 add_video 流程）
        let poster_pic = if let Some(poster) = poster_path.as_ref() {
            Picture::from_path(poster.as_ref())?
        } else {
            const TRANSPARENT_PNG_1X1: &[u8] = &[
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
                0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
                0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78,
                0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
                0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
            ];
            Picture::from_bytes(TRANSPARENT_PNG_1X1.to_vec(), ".png")
        };

        // 3) 直接构造 Picture（不走 add_picture，避免文件读取限制）
        let mut pic = poster_pic;
        pic.set_left(left.emu());
        pic.set_top(top.emu());
        pic.set_width(width.emu());
        pic.set_height(height.emu());
        pic.set_id(self.slide.next_shape_id());
        if pic.name().is_empty() {
            pic.set_name(format!("Audio {}", self.slide.inner.shapes.len() + 1));
        }
        // 4a) 海报帧图片关系 id
        let poster_rid = self.slide.allocate_image_rid();
        pic.pic_mut().rid = poster_rid.clone();
        // 4b) 注册海报帧图片 MediaEntry
        let poster_partname = crate::opc::part::new_part_name(
            format!(
                "/ppt/media/image{}.{}",
                self.slide.next_media_index(),
                pic.ext.trim_start_matches('.')
            )
            .as_str(),
        );
        let poster_ct = crate::shape::picture::content_type_for(&pic.ext);
        let poster_blob = pic.blob.clone().unwrap_or_default();
        self.slide.register_media(MediaEntry {
            partname: poster_partname,
            content_type: poster_ct.to_string(),
            blob: poster_blob,
            rid: poster_rid.clone(),
        });

        // 5) 分配音频关系 id
        let audio_rid = self.slide.allocate_audio_rid();
        pic.set_audio(audio_rid.clone());

        // 6) 注册 AudioEntry
        let local_idx = self.slide.next_audio_index();
        let audio_partname =
            crate::opc::part::new_part_name(format!("/ppt/media/media{}.mp3", local_idx).as_str());
        self.slide.register_audio(AudioEntry {
            partname: audio_partname,
            blob: audio_blob,
            rid: audio_rid.clone(),
        });

        // 7) 推入 spTree
        let oxml_pic = pic.pic.clone();
        self.slide.inner.shapes.push(OxmlSlideShape::Pic(oxml_pic));
        Ok(pic)
    }
}

/// Slide 的"高阶背景"句柄（对标 python-pptx `_Background`）。
///
/// # 能力
/// - **只读视图**：通过 `fill_type()` 查询当前背景填充类型；
/// - **写入入口**：通过 `Slide` 上的 `set_background_solid` / `clear_background` /
///   `set_follow_master_background` 方法修改背景（这些方法直接写入 oxml 模型，
///   序列化时产出 `<p:cSld><p:bg>...</p:bg></p:cSld>`）。
///
/// # 与 python-pptx 的差异
/// python-pptx 中 `slide.background.fill.solid()` 是链式写入；本库改为
/// `slide.set_background_solid(color)` 直接写入，避免可变句柄的生命周期问题。
#[derive(Debug)]
pub struct SlideBackground<'a> {
    /// 内部只读引用。
    slide: &'a Slide,
}

impl<'a> SlideBackground<'a> {
    /// 当前背景填充类型。
    ///
    /// # 返回值
    /// - `MsoFillType::Inherit`：未设置独立背景（`inner.background` 为 `None` 或 `Reference`）；
    /// - `MsoFillType::Solid`：已设置纯色背景（`inner.background` 为 `Property` 且 `solid_fill` 非 `None`）；
    /// - 其它类型暂未支持，遇到时返回 `Inherit` 作为兜底。
    pub fn fill_type(&self) -> crate::oxml::simpletypes::MsoFillType {
        use crate::oxml::slide::SlideBackground as OxmlBg;
        match &self.slide.inner.background {
            None => crate::oxml::simpletypes::MsoFillType::Inherit,
            Some(OxmlBg::Reference(_)) => crate::oxml::simpletypes::MsoFillType::Inherit,
            Some(OxmlBg::Property(p)) => {
                if matches!(p.solid_fill, crate::oxml::color::Color::None) {
                    crate::oxml::simpletypes::MsoFillType::Inherit
                } else {
                    crate::oxml::simpletypes::MsoFillType::Solid
                }
            }
        }
    }
}

/// 幻灯片 ID（指向 `sldIdLst` 中的条目）。
///
/// 当前实现未暴露给用户（仅作类型占位），留作后续 read-modify-write 流程使用。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Hash, Default)]
pub struct SlideId(
    /// 幻灯片 ID 值（对应 `sldIdLst` 中的 `id` 属性）。
    pub u32,
);

/// 关系 id 引用（指向 `ppt/slides/slideN.xml`）。
///
/// 当前实现未暴露给用户（同 [`SlideId`]）。
#[derive(Clone, Debug, Eq, PartialEq, Hash, Default)]
pub struct SlideRef(
    /// 关系 id（如 `rId10`，指向 `ppt/slides/slideN.xml`）。
    pub String,
);

/// `Slides` —— [`crate::presentation::Presentation`] 上的幻灯片集合。
///
/// # 内部表示
///
/// 仅持有一个 `Vec<SlideEntry>`，每个 entry 包含：
///
/// - `sld`：高阶 [`Slide`]；
/// - `sld_id`：`sldIdLst` 中用到的 id（一般从 `256` 开始递增）；
/// - `rid`：与 `presentation.xml.rels` 中的 `<Relationship Id="..."/>` 对应；
/// - `partname`：相对 zip 根的 part 路径（如 `/ppt/slides/slide1.xml`）。
#[derive(Debug, Default)]
pub struct Slides {
    pub(crate) slides: Vec<SlideEntry>,
}

/// 单张幻灯片在 `Slides` 集合内的"条目"（含 oxml 与 OPC 元数据）。
#[derive(Debug, Clone)]
pub(crate) struct SlideEntry {
    /// 高阶 Slide。
    pub sld: Slide,
    /// `sldIdLst` 中用到的 id。
    pub sld_id: u32,
    /// 关系 id（指向 `ppt/slides/slideN.xml`）。
    pub rid: String,
    /// slide 文件名（`/ppt/slides/slideN.xml`）。
    pub partname: String,
}

impl SlideEntry {
    /// **包内**构造一个 [`SlideEntry`]。
    ///
    /// 三个 OPC 字段由 [`crate::presentation::Presentation::from_opc`] 在
    /// 解析 `presentation.xml.rels` + `sldIdLst` 后填入。
    pub(crate) fn new(sld: Slide, sld_id: u32, rid: String, partname: String) -> Self {
        SlideEntry {
            sld,
            sld_id,
            rid,
            partname,
        }
    }
}

impl Slides {
    /// 新建一个空集合。
    pub fn new() -> Self {
        Slides::default()
    }

    /// 数量。
    pub fn len(&self) -> usize {
        self.slides.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.slides.is_empty()
    }

    /// 遍历所有 entry。
    #[allow(private_interfaces)]
    pub fn iter(&self) -> std::slice::Iter<'_, SlideEntry> {
        self.slides.iter()
    }

    /// 按下标取不可变 entry。
    #[allow(private_interfaces)]
    pub fn get(&self, idx: usize) -> Option<&SlideEntry> {
        self.slides.get(idx)
    }

    /// 按下标取可变 entry。
    #[allow(private_interfaces)]
    pub fn get_mut(&mut self, idx: usize) -> Option<&mut SlideEntry> {
        self.slides.get_mut(idx)
    }

    /// 添加一个新 slide（空白）。
    ///
    /// 返回新 slide 的可变引用，调用方可继续 `shapes_mut().add_*(...)`。
    /// `id_counter` 必须从 [`crate::presentation::Presentation::id_counter`] 传入。
    ///
    /// 默认使用第一个 layout（索引 0）。如需指定版式，调用
    /// [`Slides::add_slide_with_layout`]。
    pub fn add_slide(&mut self, id_counter: Rc<Cell<u32>>) -> crate::Result<&mut Slide> {
        self.add_slide_with_layout(id_counter, 0)
    }

    /// 添加一个新 slide，并指定使用的版式索引（对应 `SlideLayouts[i]`）。
    ///
    /// # 参数
    /// - `id_counter`：全局 shape id 计数器（必须与所属 Presentation 共享）；
    /// - `layout_idx`：版式索引。`<= 0` 时退化为 `0`；越界会被钳制为最后一个。
    ///
    /// # 行为
    /// - 新 slide 的 `layout_rid` 形如 `rIdLayout<N>`；
    /// - 在保存时 `presentation.xml` 的 `<p:sldIdLst/>` 内会按调用顺序排列。
    pub fn add_slide_with_layout(
        &mut self,
        id_counter: Rc<Cell<u32>>,
        layout_idx: usize,
    ) -> crate::Result<&mut Slide> {
        let next_no = self.slides.len() + 1;
        let mut sld = Slide::blank(id_counter);
        sld.set_layout_rid(format!("rIdLayout{}", layout_idx + 1));
        let entry = SlideEntry {
            sld,
            sld_id: 256 + next_no as u32,
            rid: format!("rId{}", 10 + next_no),
            partname: format!("/ppt/slides/slide{}.xml", next_no),
        };
        // 先 push，然后用最后位置的索引获取可变引用，避免在借用时再借 self
        self.slides.push(entry);
        let last = self.slides.len() - 1;
        Ok(&mut self.slides[last].sld)
    }

    /// **包内**：把一个外部构造的 [`SlideEntry`] 推入集合末尾。
    ///
    /// 主要服务于 [`crate::presentation::Presentation::from_opc`]：
    /// 在从 zip 读出所有 slide 后，按 `sldIdLst` 顺序依次推入，
    /// 保证后续 `save` 时 `sldIdLst` 与 `slide_ids` 顺序一致。
    pub(crate) fn push_entry(&mut self, entry: SlideEntry) {
        self.slides.push(entry);
    }

    /// 按下标移除。
    ///
    /// 对标 python-pptx `Presentation.slides._sldIdLst.remove(index)`。
    /// 返回被移除的 `Slide`；越界返回 `None`。
    pub fn remove(&mut self, idx: usize) -> Option<Slide> {
        if idx < self.slides.len() {
            let removed = self.slides.remove(idx);
            // 重新分配 sld_id / rid / partname，避免空号
            self.reindex();
            Some(removed.sld)
        } else {
            None
        }
    }

    /// 将幻灯片从 `from_idx` 移动到 `to_idx`，实现重排序。
    ///
    /// 对标 python-pptx 中通过 XML 操作调整 `sldIdLst` 子元素顺序的能力。
    ///
    /// # 参数
    /// - `from_idx`：源位置索引（0-based）；
    /// - `to_idx`：目标位置索引（0-based，基于移除前的长度）。
    ///
    /// # 行为
    /// - 越界时返回 `Err(IndexOutOfRange)`；
    /// - `from_idx == to_idx` 时为 no-op；
    /// - 移动后所有 `sld_id` / `rid` / `partname` 会重新分配（调用 [`reindex`]）。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::Presentation;
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # p.slides_mut().add_slide(counter.clone()).unwrap();
    /// # p.slides_mut().add_slide(counter.clone()).unwrap();
    /// # p.slides_mut().add_slide(counter).unwrap();
    /// // 将第 0 张移到末尾
    /// p.slides_mut().move_slide(0, 2).unwrap();
    /// ```
    pub fn move_slide(&mut self, from_idx: usize, to_idx: usize) -> crate::Result<()> {
        let len = self.slides.len();
        if from_idx >= len {
            return Err(crate::Error::IndexOutOfRange(from_idx));
        }
        // to_idx 允许等于 len（表示移到末尾），但实际插入位置不超过 len-1
        if to_idx > len {
            return Err(crate::Error::IndexOutOfRange(to_idx));
        }
        if from_idx == to_idx || from_idx + 1 == to_idx {
            // no-op：移到自身或紧邻后方
            return Ok(());
        }
        // 取出元素
        let entry = self.slides.remove(from_idx);
        // 计算实际插入位置：如果 from_idx < to_idx，由于已移除一个元素，目标位置需 -1
        let insert_at = if to_idx > from_idx {
            to_idx - 1
        } else {
            to_idx
        };
        self.slides.insert(insert_at, entry);
        // 重新分配 sld_id / rid / partname
        self.reindex();
        Ok(())
    }

    /// 按给定索引顺序批量重排幻灯片。
    ///
    /// 对标 python-pptx 中通过 XML 操作批量调整 `sldIdLst` 子元素顺序的能力。
    ///
    /// # 参数
    /// - `indices`：新顺序的索引列表（基于重排前的位置）。长度必须等于当前幻灯片数，
    ///   且每个索引在 `[0, len)` 范围内、不重复、全覆盖。
    ///
    /// # 行为
    /// - 验证失败时返回 `Err` 且**不修改**当前顺序；
    /// - 重排后所有 `sld_id` / `rid` / `partname` 会重新分配（调用 [`reindex`]）。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::Presentation;
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # p.slides_mut().add_slide(counter.clone()).unwrap();
    /// # p.slides_mut().add_slide(counter.clone()).unwrap();
    /// # p.slides_mut().add_slide(counter).unwrap();
    /// // 反转顺序：[0,1,2] -> [2,1,0]
    /// p.slides_mut().reorder(&[2, 1, 0]).unwrap();
    /// ```
    pub fn reorder(&mut self, indices: &[usize]) -> crate::Result<()> {
        let len = self.slides.len();
        if indices.len() != len {
            return Err(crate::Error::Other(format!(
                "reorder: 索引数 {} 与幻灯片数 {} 不匹配",
                indices.len(),
                len
            )));
        }
        // 验证：所有索引在范围内且不重复
        let mut seen = vec![false; len];
        for &i in indices {
            if i >= len {
                return Err(crate::Error::IndexOutOfRange(i));
            }
            if seen[i] {
                return Err(crate::Error::Other(format!("reorder: 索引 {i} 重复出现")));
            }
            seen[i] = true;
        }
        // 执行重排：按 indices 顺序收集 entries
        let mut new_slides: Vec<SlideEntry> = Vec::with_capacity(len);
        // 先取出所有 entries（避免部分借用问题）
        let old_slides = std::mem::take(&mut self.slides);
        for &i in indices {
            // safety: 已验证 i < len 且 old_slides.len() == len
            new_slides.push(old_slides[i].clone());
        }
        self.slides = new_slides;
        // 重新分配 sld_id / rid / partname
        self.reindex();
        Ok(())
    }

    /// 在指定位置插入一个空白 slide。
    ///
    /// 对标 pypdf `PdfWriter.insert_page(index)` / python-pptx 中通过
    /// XML 操作在 `sldIdLst` 中间插入条目的能力。
    ///
    /// # 参数
    /// - `id_counter`：全局 shape id 计数器；
    /// - `index`：插入位置（0 = 最前，`len` = 末尾，等价于 `add_slide`）。
    ///
    /// # 注意
    /// 插入后所有 `sld_id` / `rid` / `partname` 会**重新分配**——
    /// 因为 OOXML 的 `sldIdLst` 要求 id 单调递增。调用方不应依赖
    /// 插入前的 `sld_id` / `rid` 值。
    pub fn insert_slide(
        &mut self,
        id_counter: Rc<Cell<u32>>,
        index: usize,
    ) -> crate::Result<&mut Slide> {
        let clamped = index.min(self.slides.len());
        let mut sld = Slide::blank(id_counter);
        sld.set_layout_rid("rIdLayout1".to_string());
        let entry = SlideEntry {
            sld,
            sld_id: 0, // 占位；reindex 会重算
            rid: String::new(),
            partname: String::new(),
        };
        self.slides.insert(clamped, entry);
        // 重新分配 sld_id / rid / partname
        self.reindex();
        let last = clamped.min(self.slides.len() - 1);
        Ok(&mut self.slides[last].sld)
    }

    /// 克隆一个已有 slide 到指定位置。
    ///
    /// 对标 pypdf `PdfWriter.clone_page_from_reader` + `insert_page`。
    /// 深拷贝源 slide 的所有形状、文本、备注等，但**分配新的** id / rid / partname。
    ///
    /// # 参数
    /// - `src_idx`：源 slide 索引；
    /// - `insert_at`：目标位置（`len` = 末尾）。
    pub fn clone_slide(&mut self, src_idx: usize, insert_at: usize) -> crate::Result<&mut Slide> {
        if src_idx >= self.slides.len() {
            return Err(crate::Error::IndexOutOfRange(src_idx));
        }
        let clamped = insert_at.min(self.slides.len());
        let cloned_sld = self.slides[src_idx].sld.clone();
        let entry = SlideEntry {
            sld: cloned_sld,
            sld_id: 0,
            rid: String::new(),
            partname: String::new(),
        };
        self.slides.insert(clamped, entry);
        self.reindex();
        let last = clamped.min(self.slides.len() - 1);
        Ok(&mut self.slides[last].sld)
    }

    /// 从另一个 `Slides` 集合追加所有 slide（深拷贝）。
    ///
    /// 对标 pypdf `PdfWriter.append_pages_from_reader`。
    /// 每个源 slide 被深拷贝后追加到当前集合末尾。
    ///
    /// # 注意
    /// - 源 slide 的 `id_counter` **不会被**共享——新 slide 使用
    ///   当前集合的 `id_counter`；
    /// - 追加后所有 `sld_id` / `rid` / `partname` 重新分配。
    pub fn append_slides_from(&mut self, other: &Slides) {
        for entry in &other.slides {
            let cloned = entry.sld.clone();
            self.slides.push(SlideEntry {
                sld: cloned,
                sld_id: 0,
                rid: String::new(),
                partname: String::new(),
            });
        }
        self.reindex();
    }

    /// 重新分配所有 slide 的 `sld_id` / `rid` / `partname`。
    ///
    /// 在 `insert_slide` / `clone_slide` / `remove` / `append_slides_from`
    /// 之后调用，确保 `sldIdLst` 中的 id 单调递增、partname 不冲突。
    fn reindex(&mut self) {
        for (i, entry) in self.slides.iter_mut().enumerate() {
            let no = i + 1;
            entry.sld_id = 256 + no as u32;
            entry.rid = format!("rId{}", 10 + no);
            entry.partname = format!("/ppt/slides/slide{}.xml", no);
        }
    }

    /// 按引用找到第一个匹配的 slide 索引。
    ///
    /// 形如 python-pptx 中 `slides.index(slide)`；**未找到返回 None**（不抛异常），
    /// 与 python-pptx 略不同——但更符合 Rust 的"零异常"习惯。
    pub fn index_of(&self, sld: &Slide) -> Option<usize> {
        // 比较内部 oxml 引用即可（每个 slide 独立持有 OxmlSld）
        self.slides
            .iter()
            .position(|e| std::ptr::eq(&e.sld.inner, &sld.inner))
    }
}

#[cfg(test)]
#[allow(clippy::field_reassign_with_default)]
mod tests {
    use super::*;
    use crate::oxml::simpletypes::PresetGeometry;
    use crate::units::Inches;

    /// 在 slide 上手工添加一个**标题占位符** sp，验证 `Shapes::title()` 能找到它。
    #[test]
    fn shapes_title_finds_title_placeholder() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let mut sp = crate::oxml::shape::Sp::default();
        sp.id = 2;
        sp.name = "Title 1".into();
        sp.is_placeholder = true;
        sp.ph_idx = Some(0);
        sp.ph_type = Some("title".into());
        sp.text = crate::oxml::txbody::TextBody::new();
        s.inner.shapes.push(OxmlSlideShape::Sp(sp));
        let _t = s.shapes().title().expect("title exists");
    }

    /// 验证 `Shapes::placeholders()` 收集占位符 + 按 idx 排序。
    #[test]
    fn shapes_placeholders_sorted_by_idx() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        for (i, idx) in [3u32, 0, 5].iter().copied().enumerate() {
            let mut sp = crate::oxml::shape::Sp::default();
            sp.id = (i + 10) as u32;
            sp.name = format!("PH {idx}");
            sp.is_placeholder = true;
            sp.ph_idx = Some(idx);
            sp.text = crate::oxml::txbody::TextBody::new();
            s.inner.shapes.push(OxmlSlideShape::Sp(sp));
        }
        let phs = s.shapes().placeholders();
        assert_eq!(phs.len(), 3);
        // 第 0 个应是 idx=0
        if let crate::shape::ShapeKind::AutoShape(a) = &phs[0] {
            assert_eq!(a.sp().name, "PH 0");
        } else {
            panic!("expected AutoShape");
        }
    }

    /// TODO-007：`set_title_text` / `title_text` 基本流程。
    #[test]
    fn placeholder_title_text_set_and_get() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let mut sp = crate::oxml::shape::Sp::default();
        sp.id = 2;
        sp.name = "Title 1".into();
        sp.is_placeholder = true;
        sp.ph_idx = Some(0);
        sp.ph_type = Some("title".into());
        sp.text = crate::oxml::txbody::TextBody::new();
        s.inner.shapes.push(OxmlSlideShape::Sp(sp));

        // 初始无文本
        assert_eq!(s.title_text(), Some("".to_string()));
        // 设置标题
        assert!(s.set_title_text("Hello Title"));
        assert_eq!(s.title_text(), Some("Hello Title".to_string()));
    }

    /// TODO-007：`set_title_text` 未找到标题占位符时返回 false。
    #[test]
    fn placeholder_title_text_not_found() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        // 无任何占位符
        assert!(!s.set_title_text("test"));
        assert_eq!(s.title_text(), None);
    }

    /// TODO-007：`append_body_paragraph` / `set_body_text` / `body_text`。
    #[test]
    fn placeholder_body_text_append_and_set() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let mut sp = crate::oxml::shape::Sp::default();
        sp.id = 3;
        sp.name = "Content Placeholder 1".into();
        sp.is_placeholder = true;
        sp.ph_idx = Some(1);
        sp.ph_type = Some("body".into());
        sp.text = crate::oxml::txbody::TextBody::new();
        s.inner.shapes.push(OxmlSlideShape::Sp(sp));

        // 追加段落
        assert!(s.append_body_paragraph("第一段"));
        assert!(s.append_body_paragraph("第二段"));
        assert_eq!(s.body_text(), Some("第一段\n第二段".to_string()));

        // 替换全部
        assert!(s.set_body_text("替换文本"));
        assert_eq!(s.body_text(), Some("替换文本".to_string()));
    }

    /// TODO-007：占位符继承——从 layout 继承 xfrm / fill / line。
    #[test]
    fn placeholder_inheritance_from_layout() {
        use crate::oxml::sppr::{Fill, Transform};
        use crate::units::Emu;
        use std::cell::RefCell;
        use std::rc::Rc;

        // 构造 layout：含一个 title 占位符（idx=0），带完整 xfrm
        let mut layout_sp = crate::oxml::shape::Sp::default();
        layout_sp.is_placeholder = true;
        layout_sp.ph_idx = Some(0);
        layout_sp.ph_type = Some("title".into());
        layout_sp.properties.xfrm = Transform {
            off_x: Some(Emu(457200)),
            off_y: Some(Emu(457200)),
            ext_cx: Some(Emu(8229600)),
            ext_cy: Some(Emu(1143000)),
            rot: None,
            flip_h: false,
            flip_v: false,
        };
        layout_sp.properties.fill =
            Fill::Solid(crate::oxml::color::Color::RGB(crate::units::RGBColor::RED));
        let layout_oxml = crate::oxml::slidelayout::SldLayout {
            name: "Title Slide".into(),
            type_: "title".into(),
            shapes: vec![layout_sp],
        };
        let layout_ref = crate::slide_layouts::SlideLayoutRef {
            idx: 0,
            partname: "/ppt/slideLayouts/slideLayout1.xml".into(),
            rid: "rIdLayout1".into(),
            oxml: Rc::new(RefCell::new(layout_oxml)),
        };

        // 构造 slide：含一个 title 占位符（idx=0），但 xfrm 为空、fill 为 Inherit
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let mut slide_sp = crate::oxml::shape::Sp::default();
        slide_sp.id = 2;
        slide_sp.name = "Title 1".into();
        slide_sp.is_placeholder = true;
        slide_sp.ph_idx = Some(0);
        slide_sp.ph_type = Some("title".into());
        slide_sp.text = crate::oxml::txbody::TextBody::new();
        // xfrm 为空（应继承）、fill 为 Inherit（应继承）、line 为 None（应继承）
        s.inner.shapes.push(OxmlSlideShape::Sp(slide_sp));

        // 调用 placeholders_inherited
        let phs = s.shapes().placeholders_inherited(&layout_ref);
        assert_eq!(phs.len(), 1);
        if let crate::shape::ShapeKind::Placeholder(p) = &phs[0] {
            // 验证 xfrm 已继承
            assert_eq!(p.0.sp().properties.xfrm.off_x, Some(Emu(457200)));
            assert_eq!(p.0.sp().properties.xfrm.ext_cx, Some(Emu(8229600)));
            // 验证 fill 已继承
            assert!(matches!(
                &p.0.sp().properties.fill,
                Fill::Solid(crate::oxml::color::Color::RGB(_))
            ));
        } else {
            panic!("expected Placeholder");
        }

        // 验证原 slide 上的占位符**未被修改**（继承是 clone 后的快照）
        if let OxmlSlideShape::Sp(orig) = &s.inner.shapes[0] {
            assert!(orig.properties.xfrm.is_empty(), "原 slide 占位符不应被修改");
        } else {
            panic!("expected Sp");
        }
    }

    /// TODO-007：`add_placeholder_from_layout` 从 layout 创建新占位符。
    #[test]
    fn add_placeholder_from_layout_creates_new() {
        use std::cell::RefCell;
        use std::rc::Rc;

        // 构造 layout：含一个 body 占位符（idx=1）
        let mut layout_sp = crate::oxml::shape::Sp::default();
        layout_sp.is_placeholder = true;
        layout_sp.ph_idx = Some(1);
        layout_sp.ph_type = Some("body".into());
        layout_sp.name = "Content Placeholder 1".into();
        let layout_oxml = crate::oxml::slidelayout::SldLayout {
            name: "Title and Content".into(),
            type_: "obj".into(),
            shapes: vec![layout_sp],
        };
        let layout_ref = crate::slide_layouts::SlideLayoutRef {
            idx: 0,
            partname: "/ppt/slideLayouts/slideLayout1.xml".into(),
            rid: "rIdLayout1".into(),
            oxml: Rc::new(RefCell::new(layout_oxml)),
        };

        // 在 slide 上创建占位符
        let mut s = Slide::blank(Rc::new(Cell::new(10)));
        let auto = s
            .shapes_mut()
            .add_placeholder_from_layout(1, &layout_ref)
            .expect("创建成功");
        // 验证：新占位符继承了 layout 的 ph_idx / ph_type
        assert!(auto.sp().is_placeholder);
        assert_eq!(auto.sp().ph_idx, Some(1));
        assert_eq!(auto.sp().ph_type.as_deref(), Some("body"));
        // 验证：分配了新 id（>10）
        assert!(auto.sp().id > 10);
        // 验证：文本为空
        assert!(auto.sp().text.paragraphs.is_empty() || auto.sp().text.paragraphs.len() == 1);
    }

    /// TODO-007：`add_placeholder_from_layout` 未找到匹配占位符时返回错误。
    #[test]
    fn add_placeholder_from_layout_not_found() {
        use std::cell::RefCell;
        use std::rc::Rc;

        let layout_oxml = crate::oxml::slidelayout::SldLayout::default();
        let layout_ref = crate::slide_layouts::SlideLayoutRef {
            idx: 0,
            partname: "/ppt/slideLayouts/slideLayout1.xml".into(),
            rid: "rIdLayout1".into(),
            oxml: Rc::new(RefCell::new(layout_oxml)),
        };

        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let result = s.shapes_mut().add_placeholder_from_layout(99, &layout_ref);
        assert!(result.is_err());
    }

    /// `has_notes_slide` / `follow_master_background` / `name` 同款对齐。
    #[test]
    fn slide_metadata_apis() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        // 初始：无 notes
        assert!(!s.has_notes_slide());
        // 默认遵循 master 背景
        assert!(s.follow_master_background());
        // 切换为独立背景
        s.set_follow_master_background(false);
        assert!(!s.follow_master_background(), "切换后应不遵循 master");
        // 切回继承
        s.set_follow_master_background(true);
        assert!(s.follow_master_background(), "切回后应遵循 master");
        // name
        assert_eq!(s.name(), "");
        s.set_name(Some("intro"));
        assert_eq!(s.name(), "intro");
        s.set_name(None);
        assert_eq!(s.name(), "");
        // notes 触发 has_notes_slide
        s.set_notes_text(Some("hello"));
        assert!(s.has_notes_slide());
    }

    /// 验证纯色背景写入 oxml 模型并正确序列化为 `<p:bg>`。
    #[test]
    fn slide_background_solid_writes_xml() {
        use crate::oxml::color::Color;
        use crate::oxml::simpletypes::MsoFillType;
        use crate::units::RGBColor;

        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        // 初始：继承 master
        assert_eq!(s.background().fill_type(), MsoFillType::Inherit);
        assert!(s.follow_master_background());

        // 设置红色纯色背景
        s.set_background_solid(Color::RGB(RGBColor::RED));
        assert_eq!(s.background().fill_type(), MsoFillType::Solid);
        assert!(!s.follow_master_background());

        // 序列化检查：XML 中应包含 <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/>
        let xml = s.to_xml();
        assert!(xml.contains("<p:bg>"), "应写出 <p:bg> 元素");
        assert!(xml.contains("<p:bgPr>"), "应写出 <p:bgPr> 元素");
        assert!(xml.contains("FF0000"), "应写出红色 srgbClr");
        // bg 必须在 spTree 之前
        let bg_pos = xml.find("<p:bg>").expect("bg exists");
        let sptree_pos = xml.find("<p:spTree>").expect("spTree exists");
        assert!(bg_pos < sptree_pos, "<p:bg> 必须在 <p:spTree> 之前");

        // 清空背景
        s.clear_background();
        assert_eq!(s.background().fill_type(), MsoFillType::Inherit);
        assert!(s.follow_master_background());
        let xml2 = s.to_xml();
        assert!(!xml2.contains("<p:bg>"), "清空后不应再有 <p:bg>");
    }

    /// 验证 `set_background_solid(None)` 等价于 `clear_background`。
    #[test]
    fn slide_background_solid_none_clears() {
        use crate::oxml::color::Color;
        use crate::units::RGBColor;

        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        s.set_background_solid(Color::RGB(RGBColor::BLUE));
        assert!(!s.follow_master_background());
        // Color::None 等价于清空
        s.set_background_solid(Color::None);
        assert!(s.follow_master_background());
    }

    /// 验证 `set_follow_master_background(false)` 写入占位背景。
    #[test]
    fn slide_background_follow_master_toggle() {
        use crate::oxml::simpletypes::MsoFillType;

        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        // 初始继承
        assert!(s.follow_master_background());
        // 切换为独立：写入白色占位
        s.set_follow_master_background(false);
        assert!(!s.follow_master_background());
        assert_eq!(s.background().fill_type(), MsoFillType::Solid);
        // 再次 set(false) 应保持不变（不覆盖已有独立背景）
        s.set_follow_master_background(false);
        assert!(!s.follow_master_background());
        // 切回继承
        s.set_follow_master_background(true);
        assert!(s.follow_master_background());
    }

    /// `add_connector` 必须正确使用 4 端点（包括 y 坐标）。
    ///
    /// 早期版本会把 y 强制为 0，导出的 PPT 中连接器无法显示正确的倾斜。
    #[test]
    fn add_connector_uses_y_coords() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let c = s
            .shapes_mut()
            .add_connector(
                crate::oxml::simpletypes::MsoConnectorType::Straight,
                Inches(1.0),
                Inches(2.0),
                Inches(3.0),
                Inches(4.0),
            )
            .expect("add connector");
        // begin/end 已正确设置（不依赖 xfrm）
        let b = c.begin().expect("begin");
        let e = c.end().expect("end");
        assert_eq!(b.0, Inches(1.0).emu().value());
        assert_eq!(b.1, Inches(2.0).emu().value());
        assert_eq!(e.0, Inches(3.0).emu().value());
        assert_eq!(e.1, Inches(4.0).emu().value());
    }

    /// `add_shape` 后形状的 `id` / `name` 应正常设置。
    #[test]
    fn add_shape_preserves_id_and_name() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        let sh = s
            .shapes_mut()
            .add_shape(
                PresetGeometry::Rectangle,
                Inches(0.5),
                Inches(0.5),
                Inches(3.0),
                Inches(2.0),
            )
            .expect("add shape");
        assert!(sh.id() > 0, "shape id 应已分配");
        assert!(
            sh.name().contains("Shape") || sh.name().contains("Rectangle"),
            "shape name 应反映类型，实际：{}",
            sh.name()
        );
    }

    /// `move_slide` 重排序：将第 0 张移到末尾，验证顺序变化与 reindex。
    #[test]
    fn move_slide_reorders_correctly() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        // 添加 3 张 slide，分别设置不同的 name 以便区分
        for i in 0..3 {
            let s = slides.add_slide(counter.clone()).unwrap();
            s.set_name(Some(&format!("slide{}", i)));
        }
        assert_eq!(slides.len(), 3);
        // 移动前：[slide0, slide1, slide2]
        assert_eq!(slides.get(0).unwrap().sld.name(), "slide0");
        // 将第 0 张移到末尾（to_idx=3 表示移到 len 位置）
        slides.move_slide(0, 3).unwrap();
        // 移动后：[slide1, slide2, slide0]
        assert_eq!(slides.get(0).unwrap().sld.name(), "slide1");
        assert_eq!(slides.get(1).unwrap().sld.name(), "slide2");
        assert_eq!(slides.get(2).unwrap().sld.name(), "slide0");
        // reindex 后 sld_id 应单调递增
        assert_eq!(slides.get(0).unwrap().sld_id, 257);
        assert_eq!(slides.get(1).unwrap().sld_id, 258);
        assert_eq!(slides.get(2).unwrap().sld_id, 259);
    }

    /// `move_slide` 越界返回错误。
    #[test]
    fn move_slide_out_of_bounds_returns_error() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        slides.add_slide(counter.clone()).unwrap();
        // from_idx 越界
        assert!(slides.move_slide(5, 0).is_err());
        // to_idx 越界
        assert!(slides.move_slide(0, 5).is_err());
    }

    /// `move_slide` 同位置为 no-op。
    #[test]
    fn move_slide_same_index_is_noop() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        slides.add_slide(counter.clone()).unwrap();
        slides.add_slide(counter.clone()).unwrap();
        let before = slides.get(0).unwrap().sld_id;
        slides.move_slide(0, 0).unwrap();
        assert_eq!(slides.get(0).unwrap().sld_id, before);
    }

    /// `remove` 删除后 reindex 保证 sld_id 连续。
    #[test]
    fn remove_reindexes_remaining_slides() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        for i in 0..3 {
            let s = slides.add_slide(counter.clone()).unwrap();
            s.set_name(Some(&format!("slide{}", i)));
        }
        // 删除中间一张
        let removed = slides.remove(1).unwrap();
        assert_eq!(removed.name(), "slide1");
        // 剩余 2 张，sld_id 应重新分配为 257/258
        assert_eq!(slides.len(), 2);
        assert_eq!(slides.get(0).unwrap().sld.name(), "slide0");
        assert_eq!(slides.get(1).unwrap().sld.name(), "slide2");
        assert_eq!(slides.get(0).unwrap().sld_id, 257);
        assert_eq!(slides.get(1).unwrap().sld_id, 258);
    }

    /// `reorder` 批量重排幻灯片顺序。
    ///
    /// 这是 TODO-021 的测试。
    #[test]
    fn reorder_batch_reorders_slides() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        for i in 0..4 {
            let s = slides.add_slide(counter.clone()).unwrap();
            s.set_name(Some(&format!("slide{}", i)));
        }
        // 反转顺序：[0,1,2,3] -> [3,2,1,0]
        slides.reorder(&[3, 2, 1, 0]).unwrap();
        assert_eq!(slides.get(0).unwrap().sld.name(), "slide3");
        assert_eq!(slides.get(1).unwrap().sld.name(), "slide2");
        assert_eq!(slides.get(2).unwrap().sld.name(), "slide1");
        assert_eq!(slides.get(3).unwrap().sld.name(), "slide0");
        // sld_id 应重新分配为 257/258/259/260
        assert_eq!(slides.get(0).unwrap().sld_id, 257);
        assert_eq!(slides.get(3).unwrap().sld_id, 260);
    }

    /// `reorder` 索引数不匹配时返回错误。
    ///
    /// 这是 TODO-021 的测试。
    #[test]
    fn reorder_length_mismatch_returns_error() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        for _ in 0..3 {
            slides.add_slide(counter.clone()).unwrap();
        }
        // 索引数不足
        assert!(slides.reorder(&[0, 1]).is_err());
        // 索引数过多
        assert!(slides.reorder(&[0, 1, 2, 3]).is_err());
    }

    /// `reorder` 索引重复或越界时返回错误。
    ///
    /// 这是 TODO-021 的测试。
    #[test]
    fn reorder_duplicate_or_out_of_bounds_returns_error() {
        let counter = Rc::new(Cell::new(2));
        let mut slides = Slides::new();
        for _ in 0..3 {
            slides.add_slide(counter.clone()).unwrap();
        }
        // 重复索引
        assert!(slides.reorder(&[0, 0, 1]).is_err());
        // 越界索引
        assert!(slides.reorder(&[0, 1, 5]).is_err());
    }

    /// `move_up` 把 idx 处形状与后一个交换（z-order 提升）。
    ///
    /// 这是 TODO-025 的测试。
    #[test]
    fn shapes_mut_move_up_swaps_with_next() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        // 依次添加 3 个文本框，文本分别为 "A" / "B" / "C"
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(2.0), Inches(1.0), "A")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(2.0), Inches(2.0), Inches(1.0), "B")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(3.0), Inches(2.0), Inches(1.0), "C")
            .unwrap();
        // 初始顺序：A B C
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "B");
        assert_eq!(shape_text_at(&s, 2), "C");
        // 把 idx=0 上移一级 → B A C
        s.shapes_mut().move_up(0);
        assert_eq!(shape_text_at(&s, 0), "B");
        assert_eq!(shape_text_at(&s, 1), "A");
        assert_eq!(shape_text_at(&s, 2), "C");
    }

    /// `move_up` 对最后一个形状为 no-op。
    ///
    /// 这是 TODO-025 的测试。
    #[test]
    fn shapes_mut_move_up_on_last_is_noop() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(2.0), Inches(1.0), "A")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(2.0), Inches(2.0), Inches(1.0), "B")
            .unwrap();
        // 对最后一个（idx=1）上移：no-op
        s.shapes_mut().move_up(1);
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "B");
        // 越界 idx 也应为 no-op
        s.shapes_mut().move_up(999);
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "B");
    }

    /// `move_down` 把 idx 处形状与前一个交换（z-order 降低）。
    ///
    /// 这是 TODO-025 的测试。
    #[test]
    fn shapes_mut_move_down_swaps_with_prev() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(2.0), Inches(1.0), "A")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(2.0), Inches(2.0), Inches(1.0), "B")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(3.0), Inches(2.0), Inches(1.0), "C")
            .unwrap();
        // 初始顺序：A B C
        // 把 idx=2 下移一级 → A C B
        s.shapes_mut().move_down(2);
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "C");
        assert_eq!(shape_text_at(&s, 2), "B");
    }

    /// `move_down` 对第一个形状为 no-op。
    ///
    /// 这是 TODO-025 的测试。
    #[test]
    fn shapes_mut_move_down_on_first_is_noop() {
        let mut s = Slide::blank(Rc::new(Cell::new(0)));
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(2.0), Inches(1.0), "A")
            .unwrap();
        s.shapes_mut()
            .add_textbox_with_text(Inches(1.0), Inches(2.0), Inches(2.0), Inches(1.0), "B")
            .unwrap();
        // 对第一个（idx=0）下移：no-op
        s.shapes_mut().move_down(0);
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "B");
        // 越界 idx 也应为 no-op
        s.shapes_mut().move_down(999);
        assert_eq!(shape_text_at(&s, 0), "A");
        assert_eq!(shape_text_at(&s, 1), "B");
    }

    /// 辅助：取 slide 上第 idx 个形状的纯文本（仅支持 Sp 文本框）。
    fn shape_text_at(s: &Slide, idx: usize) -> String {
        match &s.inner.shapes[idx] {
            OxmlSlideShape::Sp(sp) => {
                let mut out = String::new();
                for p in &sp.text.paragraphs {
                    for r in &p.runs {
                        out.push_str(&r.text);
                    }
                }
                out
            }
            _ => String::new(),
        }
    }
}
