//! 单个 slide XML：`<p:sld>`。
//!
//! 每一张幻灯片对应一个 part：`/ppt/slides/slideN.xml`。
//!
//! # 元素结构
//!
//! ```text
//! <p:sld>
//!   <p:cSld>
//!     <p:spTree>
//!       <p:nvGrpSpPr>...</p:nvGrpSpPr>   spTree 必填项
//!       <p:grpSpPr>...</p:grpSpPr>
//!       <p:sp>...</p:sp>                ← 形状
//!       <p:pic>...</p:pic>
//!       <p:grpSp>...</p:grpSp>
//!       <p:cxnSp>...</p:cxnSp>
//!       <p:graphicFrame>...</p:graphicFrame>
//!     </p:spTree>
//!   </p:cSld>
//!   <p:clrMapOvr>...</p:clrMapOvr>      可选的颜色映射覆盖
//!   <p:transition>...</p:transition>     可选的过渡
//!   <p:timing>...</p:timing>            可选的时序
//! </p:sld>
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.slide.Slide` ←→ [`Sld`]；
//! - `pptx.shapes.shapetree` 内部枚举 ←→ [`SlideShape`]。

use crate::oxml::ns::{NS_DRAWING_MAIN, NS_PRESENTATION_MAIN};
use crate::oxml::shape::{Connector, GraphicFrame, Group, Pic, Sp};
use crate::oxml::txbody::TextBody;
use crate::oxml::writer::XmlWriter;

/// 一张幻灯片。
#[derive(Clone, Debug, Default)]
pub struct Sld {
    /// 内部 ID（`id="256"` 等）。
    pub id: u32,
    /// slide 关系 id（指向 `slideLayoutN.xml`）。
    pub layout_rid: String,
    /// 用户可读的 slide 名（对应 `<p:sld>/<p:cSld>/@name`，可选）。
    ///
    /// 对标 python-pptx `Slide.name`。空字符串表示未命名。
    /// 序列化时若非空，写到 `<p:cSld name="...">`。
    pub name: String,
    /// 幻灯片背景（`<p:bg>`）。`None` 表示遵循母版背景。
    pub background: Option<SlideBackground>,
    /// 幻灯片中的形状列表。
    pub shapes: Vec<SlideShape>,
    /// 备注（notes）。
    pub notes: Option<TextBody>,
    /// 幻灯片过渡（`<p:transition>`，可选）。
    pub transition: Option<Transition>,
    /// 形状树（spTree）末尾的扩展列表。
    pub ext_lst: Option<crate::oxml::shape::ExtensionList>,
}

/// 过渡速度。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum TransitionSpeed {
    /// 慢速。
    Slow,
    /// 中速（默认）。
    #[default]
    Medium,
    /// 快速。
    Fast,
}

impl TransitionSpeed {
    /// 转 OOXML 属性值。
    pub fn as_str(&self) -> &'static str {
        match self {
            TransitionSpeed::Slow => "slow",
            TransitionSpeed::Medium => "med",
            TransitionSpeed::Fast => "fast",
        }
    }
}

/// 过渡类型（`<p:transition>` 的子元素）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum TransitionType {
    /// 淡入淡出（`<p:fade thruBlk="1|0"/>`）。
    Fade { thru_blk: bool },
    /// 推入（`<p:push dir="..."/>`）。
    Push { dir: TransitionDirection },
    /// 擦除（`<p:wipe dir="..."/>`）。
    Wipe { dir: TransitionDirection },
    /// 分割（`<p:split orient="...|..." dir="..."/>`）。
    Split {
        orient: SplitOrientation,
        dir: TransitionDirection,
    },
    /// 覆盖（`<p:cover dir="..."/>`）。
    Cover { dir: TransitionDirection },
    /// 拉出（`<p:pull dir="..."/>`）。
    Pull { dir: TransitionDirection },
    /// 切割（`<p:cut thruBlk="1|0"/>`）。
    Cut { thru_blk: bool },
    /// 缩放（`<p:zoom dir="..."/>`）。
    Zoom { dir: TransitionDirection },
    /// 变形（`<p:morph option="byObject|byWord|byChar"/>`）。
    Morph { option: MorphOption },
    /// 无过渡（默认）。
    #[default]
    None,
}

/// 过渡方向（`dir` 属性）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum TransitionDirection {
    /// 向左（`dir="l"`）。
    Left,
    /// 向右（`dir="r"`，默认）。
    #[default]
    Right,
    /// 向上（`dir="u"`）。
    Up,
    /// 向下（`dir="d"`）。
    Down,
    /// 左上（`dir="lu"`）。
    LeftUp,
    /// 左下（`dir="ld"`）。
    LeftDown,
    /// 右上（`dir="ru"`）。
    RightUp,
    /// 右下（`dir="rd"`）。
    RightDown,
}

impl TransitionDirection {
    /// 转 OOXML 属性值。
    pub fn as_str(&self) -> &'static str {
        match self {
            TransitionDirection::Left => "l",
            TransitionDirection::Right => "r",
            TransitionDirection::Up => "u",
            TransitionDirection::Down => "d",
            TransitionDirection::LeftUp => "lu",
            TransitionDirection::LeftDown => "ld",
            TransitionDirection::RightUp => "ru",
            TransitionDirection::RightDown => "rd",
        }
    }
}

/// 分割方向（`orient` 属性）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum SplitOrientation {
    /// 水平（`orient="horz"`，默认）。
    #[default]
    Horizontal,
    /// 垂直（`orient="vert"`）。
    Vertical,
}

impl SplitOrientation {
    /// 转 OOXML 属性值。
    pub fn as_str(&self) -> &'static str {
        match self {
            SplitOrientation::Horizontal => "horz",
            SplitOrientation::Vertical => "vert",
        }
    }
}

/// Morph 选项（`option` 属性）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum MorphOption {
    /// 按对象（默认）。
    #[default]
    ByObject,
    /// 按单词。
    ByWord,
    /// 按字符。
    ByChar,
}

impl MorphOption {
    /// 转 OOXML 属性值。
    pub fn as_str(&self) -> &'static str {
        match self {
            MorphOption::ByObject => "byObject",
            MorphOption::ByWord => "byWord",
            MorphOption::ByChar => "byChar",
        }
    }
}

/// 幻灯片过渡（`<p:transition>`）。
///
/// 对应 OOXML 中 `<p:sld>` 内 `<p:cSld>` 之后的 `<p:transition>` 元素。
#[derive(Clone, Debug, Default)]
pub struct Transition {
    /// 过渡速度。
    pub speed: TransitionSpeed,
    /// 是否点击换片（`advClick="1|0"`，默认 true）。
    pub advance_click: bool,
    /// 自动换片时间（毫秒，`advTm="..."`）。`None` 表示不自动换片。
    pub advance_after_ms: Option<u32>,
    /// 过渡类型。
    pub transition_type: TransitionType,
}

impl Transition {
    /// 写出 `<p:transition>` 元素。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        attrs.push(("spd", self.speed.as_str()));
        if !self.advance_click {
            attrs.push(("advClick", "0"));
        }
        let adv_tm_s;
        if let Some(ms) = self.advance_after_ms {
            adv_tm_s = ms.to_string();
            attrs.push(("advTm", adv_tm_s.as_str()));
        }
        w.open_with("p:transition", &attrs);
        // 子元素：过渡类型
        match &self.transition_type {
            TransitionType::Fade { thru_blk } => {
                if *thru_blk {
                    w.empty_with("p:fade", &[("thruBlk", "1")]);
                } else {
                    w.empty("p:fade");
                }
            }
            TransitionType::Push { dir } => {
                w.empty_with("p:push", &[("dir", dir.as_str())]);
            }
            TransitionType::Wipe { dir } => {
                w.empty_with("p:wipe", &[("dir", dir.as_str())]);
            }
            TransitionType::Split { orient, dir } => {
                w.empty_with(
                    "p:split",
                    &[("orient", orient.as_str()), ("dir", dir.as_str())],
                );
            }
            TransitionType::Cover { dir } => {
                w.empty_with("p:cover", &[("dir", dir.as_str())]);
            }
            TransitionType::Pull { dir } => {
                w.empty_with("p:pull", &[("dir", dir.as_str())]);
            }
            TransitionType::Cut { thru_blk } => {
                if *thru_blk {
                    w.empty_with("p:cut", &[("thruBlk", "1")]);
                } else {
                    w.empty("p:cut");
                }
            }
            TransitionType::Zoom { dir } => {
                w.empty_with("p:zoom", &[("dir", dir.as_str())]);
            }
            TransitionType::Morph { option } => {
                w.empty_with("p:morph", &[("option", option.as_str())]);
            }
            TransitionType::None => {}
        }
        w.close("p:transition");
    }
}

/// 幻灯片背景（对应 `<p:bg>`）。
///
/// OOXML 中 `<p:bg>` 是 `<p:cSld>` 的第一个子元素（在 `<p:spTree>` 之前）。
/// 支持两种模式：
/// - `BackgroundProperty`：显式背景属性（`<p:bgPr>`），如纯色填充；
/// - `BackgroundReference`：引用主题背景（`<p:bgRef>`），idx 指向主题中的背景样式。
#[derive(Clone, Debug)]
pub enum SlideBackground {
    /// 显式背景属性（`<p:bgPr>`）。
    Property(BackgroundProperty),
    /// 背景引用（`<p:bgRef idx="...">`）。
    Reference(BackgroundReference),
}

/// 显式背景属性（`<p:bgPr>` 内的填充）。
#[derive(Clone, Debug, Default)]
pub struct BackgroundProperty {
    /// 纯色填充颜色。`Color::None` 表示不写出 `<a:solidFill>`。
    pub solid_fill: crate::oxml::color::Color,
}

/// 背景引用（`<p:bgRef idx="...">`）。
#[derive(Clone, Debug, Default)]
pub struct BackgroundReference {
    /// 背景样式索引（如 `1001` = bg1, `1002` = bg2）。
    pub idx: u32,
    /// 方案颜色（如 `bg1` / `bg2` / `tx1`）。
    pub scheme_color: String,
}

impl SlideBackground {
    /// 创建一个纯色背景。
    pub fn solid(color: crate::oxml::color::Color) -> Self {
        SlideBackground::Property(BackgroundProperty { solid_fill: color })
    }

    /// 创建一个遵循母版的背景引用（`idx=1001`，`schemeClr=bg1`）。
    pub fn follow_master() -> Self {
        SlideBackground::Reference(BackgroundReference {
            idx: 1001,
            scheme_color: "bg1".to_string(),
        })
    }

    /// 写出 `<p:bg>` XML。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        w.open("p:bg");
        match self {
            SlideBackground::Property(p) => {
                w.open("p:bgPr");
                if !matches!(p.solid_fill, crate::oxml::color::Color::None) {
                    p.solid_fill.write_solid_fill(w);
                }
                // bgPr 必须以 empty 的 bgPr 结尾属性收尾（OOXML 要求 a:effectLst 可选）
                w.close("p:bgPr");
            }
            SlideBackground::Reference(r) => {
                let idx_s = r.idx.to_string();
                w.open_with("p:bgRef", &[("idx", idx_s.as_str())]);
                w.empty_with("a:schemeClr", &[("val", r.scheme_color.as_str())]);
                w.close("p:bgRef");
            }
        }
        w.close("p:bg");
    }
}

#[derive(Clone, Debug)]
pub enum SlideShape {
    Sp(Sp),
    Pic(Pic),
    CxnSp(Connector),
    Group(Box<Group>),
    GraphicFrame(GraphicFrame),
}

impl Sld {
    /// 写 XML。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::with_decl();
        let attrs: Vec<(&str, &str)> = vec![
            ("xmlns:a", NS_DRAWING_MAIN),
            ("xmlns:p", NS_PRESENTATION_MAIN),
            ("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS),
        ];
        w.open_with("p:sld", &attrs);
        // cSld
        // name 是可选属性；空字符串 / 仅空白时**不**写出，避免脏数据进入文件
        if !self.name.trim().is_empty() {
            w.open_with("p:cSld", &[("name", self.name.as_str())]);
        } else {
            w.open("p:cSld");
        }
        // bg：背景（可选，必须在 spTree 之前）
        if let Some(bg) = &self.background {
            bg.write_xml(&mut w);
        }
        // spTree
        w.open("p:spTree");
        // nvGrpSpPr：spTree 必填项；cNvPr/cNvGrpSpPr/nvPr 三个子元素按规范顺序输出。
        w.open("p:nvGrpSpPr");
        w.empty_with("p:cNvPr", &[("id", "1"), ("name", "")]);
        w.empty("p:cNvGrpSpPr");
        w.empty("p:nvPr");
        w.close("p:nvGrpSpPr");
        // grpSpPr：spTree 必填项；这里写空 xfrm 占位（OOXML 允许 grpSpPr 为空）。
        w.open("p:grpSpPr");
        w.empty("a:xfrm");
        w.close("p:grpSpPr");
        for shape in &self.shapes {
            match shape {
                SlideShape::Sp(s) => s.write_xml(&mut w),
                SlideShape::Pic(p) => p.write_xml(&mut w),
                SlideShape::CxnSp(c) => c.write_xml(&mut w),
                SlideShape::Group(g) => g.write_xml(&mut w),
                SlideShape::GraphicFrame(g) => g.write_xml(&mut w),
            }
        }
        // spTree 末尾的 extLst（CT_GroupShape 允许最后包含一个 extLst）
        if let Some(ext) = &self.ext_lst {
            ext.write_xml(&mut w);
        }
        w.close("p:spTree");
        w.close("p:cSld");
        // clrMapOvr
        w.empty("p:clrMapOvr");
        // transition（TODO-020：幻灯片过渡，可选，必须在 clrMapOvr 之后、timing 之前）
        if let Some(tr) = &self.transition {
            tr.write_xml(&mut w);
        }
        // timing 暂略
        w.close("p:sld");
        w.into_string()
    }

    /// 设置 layout 关系 id（指向 `ppt/slideLayouts/slideLayoutN.xml`）。
    ///
    /// 由 `Presentation::from_opc` 在 read 路径里
    /// 从 `slideN.xml.rels` 解析出 SlideLayout 关系后回填。
    pub fn set_layout_rid(&mut self, rid: String) {
        self.layout_rid = rid;
    }
}

/// 写一段**备注页** XML（`<p:notes>`）。
///
/// 备注页与 slide 结构类似，但内部只有一个 placeholder（`type="body"`）
/// 承载 [`TextBody`]。OOXML 规范要求 placeholder id 显式标注（常用 `idx="1"`）。
///
/// # 与 python-pptx 的对应
///
/// python-pptx 中 `slide.notes_slide.notes_text_frame.text = "..."` 最终会
/// 序列化为本函数产出的格式。
///
/// # 元素顺序（OOXML 规范）
///
/// ```text
/// <p:notes>
///   <p:cSld>
///     <p:spTree>
///       <p:nvGrpSpPr/>
///       <p:grpSpPr/>
///       <p:sp>           ← placeholder body
///         <p:nvSpPr>
///           <p:cNvPr id="2" name="Notes Placeholder"/>
///           <p:cNvSpPr txBox="1"/>
///           <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
///         </p:nvSpPr>
///         <p:spPr><a:xfrm/><a:prstGeom prst="rect"/></p:spPr>
///         <p:txBody>...   ← TextBody 内容
///       </p:sp>
///     </p:spTree>
///   </p:cSld>
///   <p:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/>
/// </p:notes>
/// ```
pub fn notes_xml(tb: &TextBody) -> String {
    let mut w = XmlWriter::with_decl();
    let attrs: Vec<(&str, &str)> = vec![
        ("xmlns:a", NS_DRAWING_MAIN),
        ("xmlns:p", NS_PRESENTATION_MAIN),
        ("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS),
    ];
    w.open_with("p:notes", &attrs);
    // cSld
    w.open("p:cSld");
    // spTree
    w.open("p:spTree");
    w.empty_with("p:nvGrpSpPr", &[]);
    w.empty_with("p:grpSpPr", &[]);
    // 备注占位符 sp
    w.open("p:sp");
    w.open("p:nvSpPr");
    w.empty_with("p:cNvPr", &[("id", "2"), ("name", "Notes Placeholder")]);
    w.empty_with("p:cNvSpPr", &[("txBox", "1")]);
    w.open("p:nvPr");
    w.empty_with("p:ph", &[("type", "body"), ("idx", "1")]);
    w.close("p:nvPr");
    w.close("p:nvSpPr");
    // spPr：占位符占满整张备注页（EMU：默认 6858000 x 9144000）
    w.open("p:spPr");
    w.open("a:xfrm");
    w.empty_with("a:off", &[("x", "0"), ("y", "0")]);
    w.empty_with("a:ext", &[("cx", "6858000"), ("cy", "9144000")]);
    w.close("a:xfrm");
    w.open_with("a:prstGeom", &[("prst", "rect")]);
    w.empty("a:avLst");
    w.close("a:prstGeom");
    w.close("p:spPr");
    // txBody
    tb.write_xml(&mut w);
    w.close("p:sp");
    w.close("p:spTree");
    w.close("p:cSld");
    // clrMapOvr
    w.empty("p:clrMapOvr");
    w.close("p:notes");
    w.into_string()
}
