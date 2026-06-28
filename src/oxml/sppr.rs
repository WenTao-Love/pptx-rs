//! 形状属性：`<p:spPr>`（几何、变换、填充、线）。
//!
//! `spPr` 是所有视觉形状（sp/pic/grpSp/cxnSp）共享的元素；它包含：
//!
//! - `<a:xfrm>`：位置、尺寸、旋转、翻转。
//! - `<a:prstGeom>` 或 `<a:custGeom>`：几何。
//! - `<a:solidFill>` / 其它填充。
//! - `<a:ln>`：线。
//!
//! # 设计要点
//!
//! - 所有几何量都用 [`Emu`] 表示，序列化时按 OOXML 整数规范输出；
//! - 旋转单位为"60000 分之一度"（即 `rot=5400000` 表示 90°），由 `Transform::rot` 直接承载；
//! - 填充/边框未设置时**不**写出 XML（让 PowerPoint 走母版默认）；
//! - [`ShapeProperties`] 是 `Sp` / `Pic` / `Group` / `Connector` 共享的字段。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.shapes.shared.ShapeProperties` ←→ [`ShapeProperties`]；
//! - python-pptx 的 `BaseShapeElement` 风格未直接移植（避免 OOP 复杂性）；
//!   本库以"组合 + 共享字段"达到同样效果。

use std::str::FromStr;

use crate::oxml::color::Color;
use crate::oxml::simpletypes::PresetGeometry;
use crate::units::Emu;

/// 仿射变换。
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Transform {
    pub off_x: Option<Emu>,
    pub off_y: Option<Emu>,
    pub ext_cx: Option<Emu>,
    pub ext_cy: Option<Emu>,
    /// 顺时针旋转（60000 分之一度）。
    pub rot: Option<i32>,
    /// 水平翻转。
    pub flip_h: bool,
    /// 垂直翻转。
    pub flip_v: bool,
}

impl Transform {
    /// 写一段 XML。
    ///
    /// # 元素结构（OOXML）
    ///
    /// ```text
    /// <a:xfrm rot="..." flipH="1" flipV="1">     ← 属性可省略
    ///   <a:off x="..." y="..."/>                 ← 可选
    ///   <a:ext cx="..." cy="..."/>               ← 可选
    /// </a:xfrm>
    /// ```
    ///
    /// # 错误模式
    ///
    /// - 早期版本曾把属性错误地写成嵌套的 `<a:xfrm/>` + `</a:xfrm>`，导致
    ///   PowerPoint 报 "Invalid OOXML"。本方法已**合并**到**同一个**外层
    ///   `<a:xfrm>` 开始标签上。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        if self.is_empty() {
            return;
        }
        // 提前取出所有要序列化的字符串，扩展到函数末尾
        let rot_s = self.rot.map(|v| v.to_string());
        let off_x_s = self.off_x.map(|v| v.value().to_string());
        let off_y_s = self.off_y.map(|v| v.value().to_string());
        let ext_cx_s = self.ext_cx.map(|v| v.value().to_string());
        let ext_cy_s = self.ext_cy.map(|v| v.value().to_string());

        // 统一在外层 <a:xfrm> 上带属性——避免历史 bug 出现的"嵌套 a:xfrm"双标签。
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if self.flip_h {
            attrs.push(("flipH", "1"));
        }
        if self.flip_v {
            attrs.push(("flipV", "1"));
        }
        if let Some(s) = &rot_s {
            attrs.push(("rot", s.as_str()));
        }
        w.open_with("a:xfrm", &attrs);
        if let (Some(xs), Some(ys)) = (off_x_s.as_ref(), off_y_s.as_ref()) {
            w.empty_with("a:off", &[("x", xs.as_str()), ("y", ys.as_str())]);
        }
        if let (Some(xs), Some(ys)) = (ext_cx_s.as_ref(), ext_cy_s.as_ref()) {
            w.empty_with("a:ext", &[("cx", xs.as_str()), ("cy", ys.as_str())]);
        }
        w.close("a:xfrm");
    }

    /// 是否所有字段都为空。
    pub fn is_empty(&self) -> bool {
        self.off_x.is_none()
            && self.off_y.is_none()
            && self.ext_cx.is_none()
            && self.ext_cy.is_none()
            && self.rot.is_none()
            && !self.flip_h
            && !self.flip_v
    }
}

/// 渐变类型（`<a:lin>` / `<a:path>`）。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum GradientType {
    /// 线性渐变（`<a:lin ang="..." scaled="..."/>`）。
    /// `ang` 单位为 1/60000 度（0 = 向右，5400000 = 向下）。
    Linear(i32),
    /// 路径渐变（`<a:path path="..."/>`）。
    Path(GradientPath),
}

/// 路径渐变形状。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum GradientPath {
    /// 圆形（`path="circle"`）。
    Circle,
    /// 矩形（`path="rect"`）。
    Rect,
    /// 形状轮廓（`path="shape"`）。
    Shape,
}

impl GradientPath {
    /// 转为 OOXML 属性值。
    pub fn as_str(self) -> &'static str {
        match self {
            GradientPath::Circle => "circle",
            GradientPath::Rect => "rect",
            GradientPath::Shape => "shape",
        }
    }
}

/// 渐变光轨（`<a:gs pos="...">`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct GradientStop {
    /// 位置（0-100000，单位 1/1000 个百分点）。
    pub pos: u32,
    /// 颜色。
    pub color: Color,
}

/// 渐变填充（`<a:gradFill>`）。
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GradientFill {
    /// 光轨列表（至少 2 个）。
    pub stops: Vec<GradientStop>,
    /// 渐变类型（线性/路径）。
    pub gradient_type: GradientType,
    /// 是否翻转（`flip="none|tx|ty|xy|yx"`，默认 none）。
    pub flip: Option<String>,
    /// 是否与形状一起旋转（`rotWithShape="1|0"`）。
    pub rot_with_shape: Option<bool>,
}

/// 图案填充（`<a:pattFill>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PatternFill {
    /// 预置图案类型（如 `"pct5"` / `"horz"` / `"vert"` / `"cross"` 等）。
    pub prst: String,
    /// 前景色（`<a:fgClr>`）。
    pub fg_color: Color,
    /// 背景色（`<a:bgClr>`）。
    pub bg_color: Color,
}

/// 图片填充模式（`<a:stretch>` / `<a:tile>`）。
///
/// 对应 OOXML `CT_BlipFillProperties` 中的 `stretch` / `tile` 子元素。
/// 用于控制图片在形状区域内的填充方式。
///
/// # OOXML 结构
/// - `<a:stretch><a:fillRect/></a:stretch>`：拉伸填充（默认）
/// - `<a:tile tx="..." ty="..." sx="..." sy="..." flip="..." algn="..."/>`：平铺填充
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum BlipFillMode {
    /// 拉伸填充（`<a:stretch><a:fillRect/></a:stretch>`）。
    ///
    /// 图片被拉伸以填充整个形状区域。这是 PowerPoint 的默认模式。
    #[default]
    Stretch,
    /// 平铺填充（`<a:tile .../>`）。
    ///
    /// 图片按指定间距、缩放、翻转和对齐方式平铺重复。
    Tile {
        /// 水平偏移（EMU）。`None` 表示不写出 `tx` 属性。
        tx: Option<i64>,
        /// 垂直偏移（EMU）。`None` 表示不写出 `ty` 属性。
        ty: Option<i64>,
        /// 水平缩放（千分比，100000 = 100%）。`None` 表示不写出 `sx` 属性。
        sx: Option<i32>,
        /// 垂直缩放（千分比，100000 = 100%）。`None` 表示不写出 `sy` 属性。
        sy: Option<i32>,
        /// 翻转模式（`"none"` / `"x"` / `"y"` / `"xy"`）。`None` 表示不写出 `flip` 属性。
        flip: Option<String>,
        /// 对齐方式（`"tl"` / `"t"` / `"tr"` / `"l"` / `"ctr"` / `"r"` / `"bl"` / `"b"` / `"br"`）。
        /// `None` 表示不写出 `algn` 属性。
        algn: Option<String>,
    },
    /// 无填充模式（不写出 `<a:stretch>` 或 `<a:tile>`）。
    ///
    /// 罕见但合法：图片仅按原始尺寸放置在形状左上角。
    None,
}

/// 填充。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum Fill {
    /// 无填充（`<a:noFill/>`）。
    None,
    /// 实色。
    Solid(Color),
    /// 图片（`blipFill`），由 `blip_rid` 引用媒体。
    Blip {
        /// 关系 id（指向 `imageN.png`）。
        rid: String,
        /// 填充模式（拉伸/平铺/无）。
        mode: BlipFillMode,
    },
    /// 渐变填充（`<a:gradFill>`）。
    Gradient(GradientFill),
    /// 图案填充（`<a:pattFill>`）。
    Pattern(PatternFill),
    /// 继承（不写）。
    #[default]
    Inherit,
}

impl BlipFillMode {
    /// 写 XML。
    ///
    /// 根据 `BlipFillMode` 变体写出对应的 `<a:stretch>` / `<a:tile>` 元素，
    /// 或什么都不写（`None` 变体）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        match self {
            BlipFillMode::Stretch => {
                w.open("a:stretch");
                w.empty("a:fillRect");
                w.close("a:stretch");
            }
            BlipFillMode::Tile {
                tx,
                ty,
                sx,
                sy,
                flip,
                algn,
            } => {
                // 把所有属性值先取到块外，扩展生命周期
                let tx_s = tx.map(|v| v.to_string());
                let ty_s = ty.map(|v| v.to_string());
                let sx_s = sx.map(|v| v.to_string());
                let sy_s = sy.map(|v| v.to_string());
                let mut attrs: Vec<(&str, &str)> = Vec::new();
                if let Some(s) = &tx_s {
                    attrs.push(("tx", s));
                }
                if let Some(s) = &ty_s {
                    attrs.push(("ty", s));
                }
                if let Some(s) = &sx_s {
                    attrs.push(("sx", s));
                }
                if let Some(s) = &sy_s {
                    attrs.push(("sy", s));
                }
                if let Some(f) = flip {
                    attrs.push(("flip", f));
                }
                if let Some(a) = algn {
                    attrs.push(("algn", a));
                }
                w.empty_with("a:tile", &attrs);
            }
            BlipFillMode::None => {
                // 不写出任何填充模式元素
            }
        }
    }
}

impl Fill {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        match self {
            Fill::None => {
                w.empty("a:noFill");
            }
            Fill::Solid(c) => c.write_solid_fill(w),
            Fill::Blip { rid, mode } => {
                w.open_with(
                    "a:blipFill",
                    &[("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS)],
                );
                // 写出 <a:blip r:embed="..."/>（自闭合）
                // 注意：此前实现错误地 w.open("a:blip") 后又 w.empty_with("a:blip", ...)
                // 生成嵌套的无效 XML <a:blip><a:blip .../></a:blip>，已修复为单次 empty_with。
                w.empty_with("a:blip", &[("r:embed", rid.as_str())]);
                // 写出填充模式
                mode.write_xml(w);
                w.close("a:blipFill");
            }
            Fill::Gradient(g) => {
                // <a:gradFill flip="..." rotWithShape="...">
                let flip_s = g.flip.as_deref();
                let rws_s = g.rot_with_shape.map(|b| if b { "1" } else { "0" });
                let mut attrs: Vec<(&str, &str)> = Vec::new();
                if let Some(f) = flip_s {
                    attrs.push(("flip", f));
                }
                if let Some(r) = rws_s {
                    attrs.push(("rotWithShape", r));
                }
                if attrs.is_empty() {
                    w.open("a:gradFill");
                } else {
                    w.open_with("a:gradFill", &attrs);
                }
                // gsLst：光轨列表
                w.open("a:gsLst");
                for stop in &g.stops {
                    let pos_s = stop.pos.to_string();
                    w.open_with("a:gs", &[("pos", pos_s.as_str())]);
                    stop.color.write_solid_fill(w);
                    w.close("a:gs");
                }
                w.close("a:gsLst");
                // 渐变类型
                match &g.gradient_type {
                    GradientType::Linear(ang) => {
                        let ang_s = ang.to_string();
                        w.empty_with("a:lin", &[("ang", ang_s.as_str()), ("scaled", "1")]);
                    }
                    GradientType::Path(p) => {
                        w.empty_with("a:path", &[("path", p.as_str())]);
                    }
                }
                w.close("a:gradFill");
            }
            Fill::Pattern(p) => {
                w.open_with("a:pattFill", &[("prst", p.prst.as_str())]);
                // fgClr
                w.open("a:fgClr");
                p.fg_color.write_solid_fill(w);
                w.close("a:fgClr");
                // bgClr
                w.open("a:bgClr");
                p.bg_color.write_solid_fill(w);
                w.close("a:bgClr");
                w.close("a:pattFill");
            }
            Fill::Inherit => { /* noop */ }
        }
    }
}

/// 箭头类型（`<a:headEnd>` / `<a:tailEnd>` 的 type 属性）。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq)]
pub enum ArrowType {
    /// 无箭头（`type="none"`）。
    #[default]
    None,
    /// 箭头（`type="triangle"`）。
    Triangle,
    /// 箭头（stealth，`type="stealth"`）。
    Stealth,
    /// 菱形（`type="diamond"`）。
    Diamond,
    /// 椭圆（`type="oval"`）。
    Oval,
    /// 开放箭头（`type="arrow"`）。
    Arrow,
}

impl ArrowType {
    /// 转为 OOXML 属性值。
    pub fn as_str(self) -> &'static str {
        match self {
            ArrowType::None => "none",
            ArrowType::Triangle => "triangle",
            ArrowType::Stealth => "stealth",
            ArrowType::Diamond => "diamond",
            ArrowType::Oval => "oval",
            ArrowType::Arrow => "arrow",
        }
    }
}

/// 箭头尺寸（宽度/长度）。
///
/// OOXML 中 `w` 和 `len` 属性均取以下枚举值。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq)]
pub enum ArrowSize {
    /// 小（`"sm"`）。
    Small,
    /// 中等（`"med"`，默认）。
    #[default]
    Medium,
    /// 大（`"lg"`）。
    Large,
}

impl ArrowSize {
    /// 转为 OOXML 属性值。
    pub fn as_str(self) -> &'static str {
        match self {
            ArrowSize::Small => "sm",
            ArrowSize::Medium => "med",
            ArrowSize::Large => "lg",
        }
    }
}

/// 线条端点箭头（`<a:headEnd>` / `<a:tailEnd>`）。
#[derive(Copy, Clone, Debug, Default)]
pub struct ArrowHead {
    /// 箭头类型。
    pub arrow_type: ArrowType,
    /// 箭头宽度。
    pub width: ArrowSize,
    /// 箭头长度。
    pub length: ArrowSize,
}

/// 线条连接类型（`<a:round>` / `<a:miter>` / `<a:bevel>`）。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq)]
pub enum LineJoin {
    /// 圆角连接（`<a:round/>`，默认）。
    #[default]
    Round,
    /// 尖角连接（`<a:miter lim="..."/>`）。
    ///
    /// `lim` 为斜接限制（1/1000 度，如 800000 表示 800 度）。
    Miter(i32),
    /// 斜角连接（`<a:bevel/>`）。
    Bevel,
}

/// 边框（`<a:ln>`）。
#[derive(Clone, Debug, Default)]
pub struct Line {
    pub width: Option<Emu>, // EMU
    pub color: Color,
    pub dash: Option<Dash>,
    pub cap: Option<String>,
    pub compound: Option<String>,
    pub no_fill: bool,
    /// 起点箭头（`<a:headEnd>`）。
    pub head_end: Option<ArrowHead>,
    /// 终点箭头（`<a:tailEnd>`）。
    pub tail_end: Option<ArrowHead>,
    /// 连接类型（`<a:round>` / `<a:miter>` / `<a:bevel>`）。
    pub join: Option<LineJoin>,
    /// 渐变/图案填充（`<a:gradFill>` / `<a:pattFill>`）。
    ///
    /// - `Fill::Inherit`（默认）：使用 `color` / `no_fill` 字段（solidFill 或 noFill）；
    /// - `Fill::Gradient` / `Fill::Pattern`：写出对应的渐变/图案填充，**忽略** `color`。
    pub fill: Fill,
}

/// 线型虚实（枚举子集）。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum Dash {
    Solid,
    Dash,
    DashDot,
    Dot,
    LgDash,
    LgDashDot,
    LgDashDotDot,
    SysDash,
    SysDashDot,
    SysDashDotDot,
    SysDot,
}

impl Dash {
    pub fn as_str(self) -> &'static str {
        match self {
            Dash::Solid => "solid",
            Dash::Dash => "dash",
            Dash::DashDot => "dashDot",
            Dash::Dot => "dot",
            Dash::LgDash => "lgDash",
            Dash::LgDashDot => "lgDashDot",
            Dash::LgDashDotDot => "lgDashDotDot",
            Dash::SysDash => "sysDash",
            Dash::SysDashDot => "sysDashDot",
            Dash::SysDashDotDot => "sysDashDotDot",
            Dash::SysDot => "sysDot",
        }
    }
}

impl FromStr for Dash {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "solid" => Dash::Solid,
            "dash" => Dash::Dash,
            "dashDot" => Dash::DashDot,
            "dot" => Dash::Dot,
            "lgDash" => Dash::LgDash,
            "lgDashDot" => Dash::LgDashDot,
            "lgDashDotDot" => Dash::LgDashDotDot,
            "sysDash" => Dash::SysDash,
            "sysDashDot" => Dash::SysDashDot,
            "sysDashDotDot" => Dash::SysDashDotDot,
            "sysDot" => Dash::SysDot,
            _ => return Err(()),
        })
    }
}

impl Line {
    /// 写 XML。
    ///
    /// # 元素结构
    ///
    /// ```text
    /// <a:ln w="..." cap="..." cmpd="...">     ← 属性可省略
    ///   <a:noFill/>                            ← 或 <a:solidFill><a:srgbClr .../></a:solidFill>
    ///   <a:prstDash val="dash"/>               ← 可选
    /// </a:ln>
    /// ```
    ///
    /// # 行为
    ///
    /// - **总是**写出 `<a:ln>` 外壳（即便 width/color 都没设置）—— PowerPoint
    ///   期望"显式无边框"也是 `<a:ln><a:noFill/></a:ln>` 的形态；
    /// - 当 [`Line::no_fill`] 为真时，写 `<a:noFill/>`；否则写 solidFill。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        // 提前取出 width 字符串，扩展到函数末尾
        let w_s = self.width.map(|v| v.value().to_string());
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if let Some(s) = &w_s {
            attrs.push(("w", s.as_str()));
        }
        if let Some(c) = &self.cap {
            attrs.push(("cap", c.as_str()));
        }
        if let Some(c) = &self.compound {
            attrs.push(("cmpd", c.as_str()));
        }
        w.open_with("a:ln", &attrs);
        if self.no_fill {
            w.empty("a:noFill");
        } else {
            // 优先使用 fill 字段（渐变/图案填充）；否则回退到 solidFill
            match &self.fill {
                Fill::Gradient(_) | Fill::Pattern(_) => self.fill.write_xml(w),
                _ => self.color.write_solid_fill(w),
            }
        }
        if let Some(d) = &self.dash {
            w.open("a:prstDash");
            w.empty_with("a:prst", &[("val", d.as_str())]);
            w.close("a:prstDash");
        }
        // OOXML 顺序：headEnd → tailEnd → join（ECMA-376 §20.1.8.46-49）
        if let Some(h) = &self.head_end {
            w.empty_with(
                "a:headEnd",
                &[
                    ("type", h.arrow_type.as_str()),
                    ("w", h.width.as_str()),
                    ("len", h.length.as_str()),
                ],
            );
        }
        if let Some(t) = &self.tail_end {
            w.empty_with(
                "a:tailEnd",
                &[
                    ("type", t.arrow_type.as_str()),
                    ("w", t.width.as_str()),
                    ("len", t.length.as_str()),
                ],
            );
        }
        if let Some(j) = &self.join {
            match j {
                LineJoin::Round => {
                    w.empty("a:round");
                }
                LineJoin::Miter(lim) => {
                    let lim_s = lim.to_string();
                    w.empty_with("a:miter", &[("lim", lim_s.as_str())]);
                }
                LineJoin::Bevel => {
                    w.empty("a:bevel");
                }
            }
        }
        w.close("a:ln");
    }
}

/// 阴影效果（`<a:outerShdw>` / `<a:innerShdw>`）。
///
/// 对应 python-pptx 中 `ShadowFormat` 的底层支撑。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ShadowEffect {
    /// 阴影方向（1/60000 度，0 = 向右，2700000 = 向下，5400000 = 向左，8100000 = 向上）。
    pub dir: i32,
    /// 距离（EMU）。
    pub dist: i64,
    /// 模糊半径（EMU）。
    pub blur_rad: i64,
    /// 阴影颜色。
    pub color: Color,
    /// 是否随形状旋转（仅 outerShdw）。
    pub rot_with_shape: Option<bool>,
}

/// 发光效果（`<a:glow>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct GlowEffect {
    /// 发光半径（EMU）。
    pub rad: i64,
    /// 发光颜色。
    pub color: Color,
}

/// 柔化边缘效果（`<a:softEdge>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SoftEdgeEffect {
    /// 柔化半径（EMU）。
    pub rad: i64,
}

/// 反射效果（`<a:reflection>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ReflectionEffect {
    /// 模糊半径（EMU）。
    pub blur_rad: Option<i64>,
    /// 起始透明度（0-100000）。
    pub st_a: Option<i32>,
    /// 起始位置（0-100000）。
    pub st_pos: Option<i32>,
    /// 结束透明度（0-100000）。
    pub end_a: Option<i32>,
    /// 结束位置（0-100000）。
    pub end_pos: Option<i32>,
    /// 距离（EMU）。
    pub dist: Option<i64>,
    /// 方向（1/60000 度）。
    pub dir: Option<i32>,
    /// 是否随形状旋转。
    pub rot_with_shape: Option<bool>,
}

/// 效果列表（`<a:effectLst>`）。
///
/// 对应 OOXML 中 `<p:spPr>` 内 `<a:effectLst>` 元素。
/// 按 OOXML 顺序：blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct EffectList {
    /// 外阴影（`<a:outerShdw>`，最常用）。
    pub outer_shadow: Option<ShadowEffect>,
    /// 内阴影（`<a:innerShdw>`）。
    pub inner_shadow: Option<ShadowEffect>,
    /// 发光（`<a:glow>`）。
    pub glow: Option<GlowEffect>,
    /// 柔化边缘（`<a:softEdge>`）。
    pub soft_edge: Option<SoftEdgeEffect>,
    /// 反射（`<a:reflection>`）。
    pub reflection: Option<ReflectionEffect>,
}

impl EffectList {
    /// 是否所有效果都为 None。
    pub fn is_empty(&self) -> bool {
        self.outer_shadow.is_none()
            && self.inner_shadow.is_none()
            && self.glow.is_none()
            && self.soft_edge.is_none()
            && self.reflection.is_none()
    }

    /// 写出 `<a:effectLst>...</a:effectLst>`。
    ///
    /// 若 `is_empty()` 为 true，则**不**写出（调用方应先判断）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        if self.is_empty() {
            return;
        }
        w.open("a:effectLst");
        // 按 OOXML 顺序：glow → innerShdw → outerShdw → reflection → softEdge
        if let Some(g) = &self.glow {
            let rad_s = g.rad.to_string();
            w.open_with("a:glow", &[("rad", rad_s.as_str())]);
            g.color.write_solid_fill(w);
            w.close("a:glow");
        }
        if let Some(s) = &self.inner_shadow {
            let dir_s = s.dir.to_string();
            let dist_s = s.dist.to_string();
            let blur_s = s.blur_rad.to_string();
            w.open_with(
                "a:innerShdw",
                &[
                    ("blurRad", blur_s.as_str()),
                    ("dist", dist_s.as_str()),
                    ("dir", dir_s.as_str()),
                ],
            );
            s.color.write_solid_fill(w);
            w.close("a:innerShdw");
        }
        if let Some(s) = &self.outer_shadow {
            let dir_s = s.dir.to_string();
            let dist_s = s.dist.to_string();
            let blur_s = s.blur_rad.to_string();
            let mut attrs: Vec<(&str, &str)> = vec![
                ("blurRad", blur_s.as_str()),
                ("dist", dist_s.as_str()),
                ("dir", dir_s.as_str()),
            ];
            let rot_s;
            if let Some(r) = s.rot_with_shape {
                rot_s = if r { "1".to_string() } else { "0".to_string() };
                attrs.push(("rotWithShape", rot_s.as_str()));
            }
            w.open_with("a:outerShdw", &attrs);
            s.color.write_solid_fill(w);
            w.close("a:outerShdw");
        }
        if let Some(r) = &self.reflection {
            // 用 owned String 收集属性值，避免生命周期问题
            let mut attrs: Vec<(String, String)> = Vec::new();
            if let Some(v) = r.blur_rad {
                attrs.push(("blurRad".to_string(), v.to_string()));
            }
            if let Some(v) = r.st_a {
                attrs.push(("stA".to_string(), v.to_string()));
            }
            if let Some(v) = r.st_pos {
                attrs.push(("stPos".to_string(), v.to_string()));
            }
            if let Some(v) = r.end_a {
                attrs.push(("endA".to_string(), v.to_string()));
            }
            if let Some(v) = r.end_pos {
                attrs.push(("endPos".to_string(), v.to_string()));
            }
            if let Some(v) = r.dist {
                attrs.push(("dist".to_string(), v.to_string()));
            }
            if let Some(v) = r.dir {
                attrs.push(("dir".to_string(), v.to_string()));
            }
            if let Some(v) = r.rot_with_shape {
                attrs.push((
                    "rotWithShape".to_string(),
                    if v { "1".to_string() } else { "0".to_string() },
                ));
            }
            let refs: Vec<(&str, &str)> = attrs
                .iter()
                .map(|(k, v)| (k.as_str(), v.as_str()))
                .collect();
            w.empty_with("a:reflection", &refs);
        }
        if let Some(s) = &self.soft_edge {
            let rad_s = s.rad.to_string();
            w.empty_with("a:softEdge", &[("rad", rad_s.as_str())]);
        }
        w.close("a:effectLst");
    }
}

// ===== TODO-024：自定义几何（custGeom）=====

/// 几何矩形（`<a:rect l="..." t="..." r="..." b="..."/>`）。
///
/// 用于 `<a:custGeom>` 内部的内嵌区域定义，值可以是百分比或 EMU。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct GeomRect {
    /// 左边界（`l="..."`）。
    pub l: String,
    /// 上边界（`t="..."`）。
    pub t: String,
    /// 右边界（`r="..."`）。
    pub r: String,
    /// 下边界（`b="..."`）。
    pub b: String,
}

/// 路径段（`<a:moveTo>` / `<a:lnTo>` / `<a:cubicBezTo>` / `<a:quadBezTo>` / `<a:arcTo>` / `<a:close/>`）。
///
/// 对应 OOXML `<a:path>` 内的子元素，描述自由路径的绘制步骤。
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PathSegment {
    /// 移动到（`<a:moveTo><a:pt x="..." y="..."/></a:moveTo>`）。
    MoveTo { x: i64, y: i64 },
    /// 直线到（`<a:lnTo><a:pt x="..." y="..."/></a:lnTo>`）。
    LineTo { x: i64, y: i64 },
    /// 三次贝塞尔曲线（`<a:cubicBezTo>`，含 3 个控制点）。
    CubicBezTo {
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        x3: i64,
        y3: i64,
    },
    /// 二次贝塞尔曲线（`<a:quadBezTo>`，含 2 个控制点）。
    QuadBezTo { x1: i64, y1: i64, x2: i64, y2: i64 },
    /// 弧线（`<a:arcTo wR="..." hR="..." stAng="..." swAng="..."/>`）。
    ///
    /// - `w_r` / `h_r`：椭圆半径（EMU）；
    /// - `st_ang` / `sw_ang`：起始角 / 扫掠角（1/60000 度）。
    ArcTo {
        w_r: i64,
        h_r: i64,
        st_ang: i32,
        sw_ang: i32,
    },
    /// 关闭路径（`<a:close/>`）。
    Close,
}

impl PathSegment {
    /// 写出路径段 XML 到 writer。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        match self {
            PathSegment::MoveTo { x, y } => {
                let xs = x.to_string();
                let ys = y.to_string();
                w.open("a:moveTo");
                w.empty_with("a:pt", &[("x", xs.as_str()), ("y", ys.as_str())]);
                w.close("a:moveTo");
            }
            PathSegment::LineTo { x, y } => {
                let xs = x.to_string();
                let ys = y.to_string();
                w.open("a:lnTo");
                w.empty_with("a:pt", &[("x", xs.as_str()), ("y", ys.as_str())]);
                w.close("a:lnTo");
            }
            PathSegment::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => {
                let x1s = x1.to_string();
                let y1s = y1.to_string();
                let x2s = x2.to_string();
                let y2s = y2.to_string();
                let x3s = x3.to_string();
                let y3s = y3.to_string();
                w.open("a:cubicBezTo");
                w.empty_with("a:pt", &[("x", x1s.as_str()), ("y", y1s.as_str())]);
                w.empty_with("a:pt", &[("x", x2s.as_str()), ("y", y2s.as_str())]);
                w.empty_with("a:pt", &[("x", x3s.as_str()), ("y", y3s.as_str())]);
                w.close("a:cubicBezTo");
            }
            PathSegment::QuadBezTo { x1, y1, x2, y2 } => {
                let x1s = x1.to_string();
                let y1s = y1.to_string();
                let x2s = x2.to_string();
                let y2s = y2.to_string();
                w.open("a:quadBezTo");
                w.empty_with("a:pt", &[("x", x1s.as_str()), ("y", y1s.as_str())]);
                w.empty_with("a:pt", &[("x", x2s.as_str()), ("y", y2s.as_str())]);
                w.close("a:quadBezTo");
            }
            PathSegment::ArcTo {
                w_r,
                h_r,
                st_ang,
                sw_ang,
            } => {
                let wr_s = w_r.to_string();
                let hr_s = h_r.to_string();
                let st_s = st_ang.to_string();
                let sw_s = sw_ang.to_string();
                w.empty_with(
                    "a:arcTo",
                    &[
                        ("wR", wr_s.as_str()),
                        ("hR", hr_s.as_str()),
                        ("stAng", st_s.as_str()),
                        ("swAng", sw_s.as_str()),
                    ],
                );
            }
            PathSegment::Close => {
                w.empty("a:close");
            }
        }
    }
}

/// 路径（`<a:path>`）。
///
/// 一条路径由宽度/高度和一组 [`PathSegment`] 组成。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Path {
    /// 宽度（EMU，`w="..."`）。
    pub width: i64,
    /// 高度（EMU，`h="..."`）。
    pub height: i64,
    /// 填充模式（`fill="none|norm|darken|darkenLess|lighten|lightenLess"`，可选）。
    pub fill: Option<String>,
    /// 描边模式（`stroke="none|norm"`，可选）。
    pub stroke: Option<String>,
    /// 路径段列表。
    pub segments: Vec<PathSegment>,
}

impl Path {
    /// 写出 `<a:path>` 元素到 writer。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let w_s = self.width.to_string();
        let h_s = self.height.to_string();
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        attrs.push(("w", w_s.as_str()));
        attrs.push(("h", h_s.as_str()));
        if let Some(f) = &self.fill {
            attrs.push(("fill", f.as_str()));
        }
        if let Some(s) = &self.stroke {
            attrs.push(("stroke", s.as_str()));
        }
        w.open_with("a:path", &attrs);
        for seg in &self.segments {
            seg.write_xml(w);
        }
        w.close("a:path");
    }
}

/// 自定义几何（`<a:custGeom>`）。
///
/// 对应 OOXML 中 `<a:custGeom>` 元素，包含路径列表和可选的填充/描边/内嵌区域。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CustomGeometry {
    /// 是否允许填充（`<a:fill>`，默认 "norm"）。`None` 表示不写出。
    pub fill: Option<String>,
    /// 是否允许描边（`<a:stroke>`，默认 "norm"）。`None` 表示不写出。
    pub stroke: Option<String>,
    /// 内嵌区域（`<a:rect l="..." t="..." r="..." b="..."/>`，可选）。
    pub rect: Option<GeomRect>,
    /// 路径列表（`<a:pathLst>`）。
    pub path_list: Vec<Path>,
}

impl CustomGeometry {
    /// 写出 `<a:custGeom>` 元素到 writer。
    ///
    /// # 元素顺序（OOXML 规范）
    ///
    /// ```text
    /// <a:custGeom>
    ///   <a:avLst/>          ← 可选，调整手柄列表（暂不支持）
    ///   <a:fill>...</a:fill> ← 可选
    ///   <a:stroke>...</a:stroke> ← 可选
    ///   <a:rect .../>        ← 可选
    ///   <a:pathLst>
    ///     <a:path>...</a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("a:custGeom");
        // avLst（调整手柄列表，暂不支持，写空）
        w.empty("a:avLst");
        // fill（可选）
        if let Some(f) = &self.fill {
            w.leaf("a:fill", f.as_str());
        }
        // stroke（可选）
        if let Some(s) = &self.stroke {
            w.leaf("a:stroke", s.as_str());
        }
        // rect（可选）
        if let Some(r) = &self.rect {
            w.empty_with(
                "a:rect",
                &[
                    ("l", r.l.as_str()),
                    ("t", r.t.as_str()),
                    ("r", r.r.as_str()),
                    ("b", r.b.as_str()),
                ],
            );
        }
        // pathLst
        w.open("a:pathLst");
        for p in &self.path_list {
            p.write_xml(w);
        }
        w.close("a:pathLst");
        w.close("a:custGeom");
    }
}

/// 调整值（`<a:gd name="..." fmla="val <value>"/>`）。
///
/// 对应 python-pptx `Adjustment`，控制预设形状的调整手柄（如圆角矩形的圆角半径）。
///
/// # OOXML 结构
///
/// ```text
/// <a:avLst>
///   <a:gd name="adj" fmla="val 16667"/>   ← 调整值，16667 = 16.667%
/// </a:avLst>
/// ```
///
/// # 值的含义
///
/// 调整值以 1/100000 为单位（即 `100000` = 100%）。`effective_value()` 返回归一化的
/// 0.0-1.0 浮点值，`raw_value` 保存原始整数。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct AdjustmentValue {
    /// 调整值名称（如 "adj"、"adj1"、"adj2"）。
    pub name: String,
    /// 原始值（从 `fmla="val <value>"` 中提取的数值，单位 1/100000）。
    pub raw_value: i64,
}

impl AdjustmentValue {
    /// 用名称和原始值构造。
    ///
    /// # 参数
    /// - `name`：调整值名称（如 "adj"）；
    /// - `raw_value`：原始值（单位 1/100000，如 16667 表示 16.667%）。
    pub fn new(name: impl Into<String>, raw_value: i64) -> Self {
        Self {
            name: name.into(),
            raw_value,
        }
    }

    /// 用名称和归一化值（0.0-1.0）构造。
    ///
    /// # 参数
    /// - `name`：调整值名称；
    /// - `value`：归一化值（0.0-1.0，会被转换为 1/100000 单位）。
    pub fn from_normalized(name: impl Into<String>, value: f64) -> Self {
        Self {
            name: name.into(),
            raw_value: (value * 100000.0).round() as i64,
        }
    }

    /// 归一化值（0.0-1.0，即 `raw_value / 100000.0`）。
    ///
    /// 对应 python-pptx `Adjustment.effective_value`。
    pub fn effective_value(&self) -> f64 {
        self.raw_value as f64 / 100000.0
    }

    /// 写出 `<a:gd name="..." fmla="val ..."/>` 元素。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let val_s = self.raw_value.to_string();
        w.empty_with(
            "a:gd",
            &[
                ("name", self.name.as_str()),
                ("fmla", &format!("val {}", val_s)),
            ],
        );
    }
}

/// 几何类型（`<a:prstGeom>` 或 `<a:custGeom>`）。
///
/// 对标 python-pptx 的 `Geometry` 概念，统一表达预设几何和自定义几何。
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Geometry {
    /// 预设几何（`<a:prstGeom prst="...">`）。
    ///
    /// 第二个元素是调整值列表（`<a:avLst>`），控制形状的调整手柄。
    Preset(PresetGeometry, Vec<AdjustmentValue>),
    /// 自定义几何（`<a:custGeom>`）。
    Custom(CustomGeometry),
}

impl Default for Geometry {
    fn default() -> Self {
        Geometry::Preset(PresetGeometry::Rectangle, Vec::new())
    }
}

impl Geometry {
    /// 创建无调整值的预设几何。
    ///
    /// 等价于 `Geometry::Preset(prst, Vec::new())`。
    pub fn preset(prst: PresetGeometry) -> Self {
        Geometry::Preset(prst, Vec::new())
    }

    /// 写出几何 XML（`<a:prstGeom>` 或 `<a:custGeom>`）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        match self {
            Geometry::Preset(p, adjustments) => {
                w.open_with("a:prstGeom", &[("prst", p.as_str())]);
                // avLst：调整值列表
                if adjustments.is_empty() {
                    w.empty("a:avLst");
                } else {
                    w.open("a:avLst");
                    for adj in adjustments {
                        adj.write_xml(w);
                    }
                    w.close("a:avLst");
                }
                w.close("a:prstGeom");
            }
            Geometry::Custom(c) => {
                c.write_xml(w);
            }
        }
    }
}

/// 形状属性 `<p:spPr>`。
#[derive(Clone, Debug, Default)]
pub struct ShapeProperties {
    /// 仿射变换（位置/尺寸/旋转/翻转）。
    pub xfrm: Transform,
    /// 几何（预设或自定义）。`None` 默认为 `Preset(Rectangle)`。
    ///
    /// TODO-024：从 `Option<PresetGeometry>` 改为 `Option<Geometry>`，
    /// 支持自定义几何（`<a:custGeom>`）。
    pub geometry: Option<Geometry>,
    /// 填充（实色/图片/无/继承）。
    pub fill: Fill,
    /// 边框（`<a:ln>`）。`None` 表示不写出边框。
    pub line: Option<Line>,
    /// 效果列表（`<a:effectLst>`，可选）。
    pub effects: Option<EffectList>,
    /// 三维场景（`<a:scene3d>`，可选，TODO-050）。
    ///
    /// 定义相机与光照；与 `sp3d` 共同表达 3D 效果。
    pub scene3d: Option<Scene3d>,
    /// 形状 3D 属性（`<a:sp3d>`，可选，TODO-050）。
    ///
    /// 定义形状本身的拉伸/棱台/材质。
    pub sp3d: Option<Sp3d>,
    /// 旋转角度（度数，正向顺时针）。
    pub rot_deg: Option<f64>,
}

impl ShapeProperties {
    /// 写 XML。`tag` 一般为 `"p:spPr"`。
    ///
    /// 按 OOXML 规范，`<p:spPr>` 内部必须按以下顺序排列：
    ///
    /// 1. `<a:xfrm>`（可选）：位置/尺寸/旋转；
    /// 2. `<a:prstGeom>` / `<a:custGeom>`：几何；
    /// 3. 填充相关（`<a:noFill>` / `<a:solidFill>` / `<a:gradFill>` / `<a:blipFill>` / `<a:pattFill>`）；
    /// 4. `<a:ln>`：边框；
    /// 5. 效果（`<a:effectLst>` 等）；
    /// 6. `<a:scene3d>` / `<a:sp3d>`：三维效果；
    /// 7. `<a:extLst>`：扩展。
    ///
    /// 顺序错误会导致 PowerPoint 弹出"Invalid OOXML"对话框。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter, tag: &str) {
        w.open(tag);
        // xfrm
        if !self.xfrm.is_empty() {
            self.xfrm.write_xml(w);
        }
        // 几何（prstGeom 或 custGeom，TODO-024）
        let geom = self.geometry.clone().unwrap_or_default();
        geom.write_xml(w);
        // fill
        self.fill.write_xml(w);
        // ln
        if let Some(ln) = &self.line {
            ln.write_xml(w);
        }
        // effectLst（TODO-011：形状效果）
        if let Some(effects) = &self.effects {
            effects.write_xml(w);
        }
        // scene3d（TODO-050：三维场景）
        if let Some(scene) = &self.scene3d {
            scene.write_xml(w);
        }
        // sp3d（TODO-050：形状 3D 属性）
        if let Some(sp3d) = &self.sp3d {
            sp3d.write_xml(w);
        }
        w.close(tag);
    }

    /// **仅**写 xfrm 元素（`<p:xfrm><a:xfrm>...</a:xfrm></p:xfrm>`）。
    ///
    /// 仅 [`crate::oxml::shape::GraphicFrame`] 使用 —— OOXML 中
    /// `<p:graphicFrame>` 的子元素顺序是：
    ///
    /// ```text
    /// <p:graphicFrame>
    ///   <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
    ///   <p:xfrm>...</p:xfrm>          ← 本方法
    ///   <a:graphic>...</a:graphic>
    ///   <p:extLst>...</p:extLst>      ← 可选
    /// </p:graphicFrame>
    /// ```
    ///
    /// 与 [`ShapeProperties::write_xml`] 的区别是：本方法**不**输出 prstGeom /
    /// 填充 / 边框（这些是 `<p:spPr>` 的职责），只输出位置 / 尺寸变换。
    pub fn write_xfrm_only(&self, w: &mut super::writer::XmlWriter) {
        if self.xfrm.is_empty() {
            // 仍然输出空 <p:xfrm/> 以保证 OOXML 顺序稳定
            w.empty("p:xfrm");
            return;
        }
        w.open("p:xfrm");
        self.xfrm.write_xml(w);
        w.close("p:xfrm");
    }

    // --------------------- 形状效果 API（TODO-011 高阶） ---------------------
    //
    // 以下便捷方法封装 `<a:effectLst>` 的常用操作，对标 python-pptx 的
    // `shape.shadow` / `shape.glow` 等高阶接口。底层逻辑都通过 `effects`
    // 字段（`Option<EffectList>`）承接，序列化时由 `write_xml` 自动按
    // OOXML 顺序输出在 `<a:ln>` 之后。

    /// 读取效果列表（`<a:effectLst>`）。`None` 表示未设置。
    pub fn effects(&self) -> Option<&EffectList> {
        self.effects.as_ref()
    }

    /// 读取效果列表的可变引用。若未设置，自动初始化为空 `EffectList`。
    ///
    /// 调用此方法后即使不设置任何效果，也会写出空 `<a:effectLst/>`（PowerPoint 兼容）。
    /// 若要避免写出空元素，请用 [`Self::clear_effects`]。
    pub fn effects_mut(&mut self) -> &mut EffectList {
        self.effects.get_or_insert_with(EffectList::default)
    }

    /// 设置外阴影（`<a:outerShdw>`）。覆盖既有外阴影，保留其他效果。
    ///
    /// 对标 python-pptx `shape.shadow.inherit = False` + `shape.shadow.outerShadow`。
    ///
    /// # 参数
    /// - `shadow`：阴影配置（方向/距离/模糊半径/颜色）
    pub fn set_outer_shadow(&mut self, shadow: ShadowEffect) {
        self.effects_mut().outer_shadow = Some(shadow);
    }

    /// 设置内阴影（`<a:innerShdw>`）。覆盖既有内阴影，保留其他效果。
    pub fn set_inner_shadow(&mut self, shadow: ShadowEffect) {
        self.effects_mut().inner_shadow = Some(shadow);
    }

    /// 设置发光（`<a:glow>`）。覆盖既有发光，保留其他效果。
    pub fn set_glow(&mut self, glow: GlowEffect) {
        self.effects_mut().glow = Some(glow);
    }

    /// 设置柔化边缘（`<a:softEdge>`）。覆盖既有柔边，保留其他效果。
    pub fn set_soft_edge(&mut self, rad: i64) {
        self.effects_mut().soft_edge = Some(SoftEdgeEffect { rad });
    }

    /// 设置反射（`<a:reflection>`）。覆盖既有反射，保留其他效果。
    pub fn set_reflection(&mut self, reflection: ReflectionEffect) {
        self.effects_mut().reflection = Some(reflection);
    }

    /// 清除外阴影。
    pub fn clear_outer_shadow(&mut self) {
        if let Some(e) = self.effects.as_mut() {
            e.outer_shadow = None;
        }
    }

    /// 清除内阴影。
    pub fn clear_inner_shadow(&mut self) {
        if let Some(e) = self.effects.as_mut() {
            e.inner_shadow = None;
        }
    }

    /// 清除发光。
    pub fn clear_glow(&mut self) {
        if let Some(e) = self.effects.as_mut() {
            e.glow = None;
        }
    }

    /// 清除柔化边缘。
    pub fn clear_soft_edge(&mut self) {
        if let Some(e) = self.effects.as_mut() {
            e.soft_edge = None;
        }
    }

    /// 清除反射。
    pub fn clear_reflection(&mut self) {
        if let Some(e) = self.effects.as_mut() {
            e.reflection = None;
        }
    }

    /// 清除所有效果（删除整个 `<a:effectLst>` 元素）。
    pub fn clear_effects(&mut self) {
        self.effects = None;
    }
}

// ====================================================================
// 高阶 Fill / Line 视图（python-pptx 风格）
// ====================================================================

use crate::oxml::color::ColorFormat;

/// 填充高阶视图（`pptx.dml.fill.FillFormat`）。
///
/// # 与 python-pptx 的对应
///
/// - `pptx.dml.fill.FillFormat` ←→ [`FillFormat`]；
/// - `shape.fill.solid()` + `shape.fill.fore_color.rgb = ...` ←→
///   `fill_format.solid().set_rgb(...)`。
///
/// # 设计要点
///
/// - **借用 + 透明代理**：构造时传入 `&mut Fill`；所有写都走底层；
/// - **零分配**：颜色写入走 [`ColorFormat`] 的借用；
/// - **类型安全**：通过 [`super::simpletypes::MsoFillType`] 表达"当前填充类型"。
#[derive(Debug)]
pub struct FillFormat<'a> {
    /// 底层 [`Fill`] 引用。
    fill: &'a mut Fill,
}

impl<'a> FillFormat<'a> {
    /// 构造一个 fill 视图。
    pub fn new(fill: &'a mut Fill) -> Self {
        FillFormat { fill }
    }

    /// 底层 fill 不可变引用。
    pub fn fill(&self) -> &Fill {
        self.fill
    }
    /// 底层 fill 可变引用。
    pub fn fill_mut(&mut self) -> &mut Fill {
        self.fill
    }

    /// 当前填充类型（python-pptx `fill.type`）。
    pub fn fill_type(&self) -> super::simpletypes::MsoFillType {
        match self.fill {
            Fill::None => super::simpletypes::MsoFillType::Background,
            Fill::Solid(_) => super::simpletypes::MsoFillType::Solid,
            Fill::Blip { .. } => super::simpletypes::MsoFillType::Picture,
            Fill::Gradient(_) => super::simpletypes::MsoFillType::Gradient,
            Fill::Pattern(_) => super::simpletypes::MsoFillType::Pattern,
            Fill::Inherit => super::simpletypes::MsoFillType::Inherit,
        }
    }

    /// 切到**实色**模式并返回 [`ColorFormat`] 代理。
    ///
    /// 对应 python-pptx 中 `fill.solid()` 后再 `fill.fore_color.rgb = ...`。
    /// 调用本方法会把 `Fill` 切到 `Solid(Color::None)`，后续 `set_rgb` /
    /// `set_theme` 等会原地更新 `Color`。
    pub fn solid(&mut self) -> ColorFormat<'_> {
        // 如果不是 Solid，先切到 Solid(None)
        if !matches!(self.fill, Fill::Solid(_)) {
            *self.fill = Fill::Solid(Color::None);
        }
        match self.fill {
            Fill::Solid(c) => ColorFormat::new(c),
            _ => unreachable!("just set to Solid"),
        }
    }

    /// 切到**无填充**。
    ///
    /// 对应 python-pptx 中 `fill.background()`。但通常我们用 [`FillFormat::clear`] 更直白。
    pub fn set_none(&mut self) {
        *self.fill = Fill::None;
    }

    /// 切到图片填充。
    ///
    /// # 参数
    /// - `rid`：图片关系 id（形如 `rIdImg1`）。
    /// - `mode`：填充模式（拉伸/平铺/无）。使用 [`BlipFillMode::Stretch`] 为默认拉伸。
    pub fn set_picture(&mut self, rid: impl Into<String>, mode: BlipFillMode) {
        *self.fill = Fill::Blip {
            rid: rid.into(),
            mode,
        };
    }

    /// 重置（继承主题默认）。
    pub fn clear(&mut self) {
        *self.fill = Fill::Inherit;
    }

    /// 便捷：直接设成 sRGB 实色。
    pub fn set_solid_rgb(&mut self, c: impl Into<crate::units::RGBColor>) {
        *self.fill = Fill::Solid(Color::RGB(c.into()));
    }
    /// 便捷：直接设成主题色实色。
    pub fn set_solid_theme(&mut self, t: super::simpletypes::MsoThemeColorIndex) {
        if let Some(s) = t.as_str() {
            if let Ok(sc) = s.parse::<crate::oxml::color::SchemeColor>() {
                *self.fill = Fill::Solid(Color::Scheme(sc));
            }
        }
    }

    // ===== 渐变填充 API（TODO-009）=====

    /// 切到**渐变**模式并返回 `&mut GradientFill` 供进一步配置。
    ///
    /// 对应 python-pptx 中 `fill.gradient()`。调用本方法会把 `Fill` 切到
    /// `Gradient(GradientFill { stops: vec![], gradient_type: Linear(0), .. })`，
    /// 调用方随后通过返回的引用添加光轨、设置角度等。
    ///
    /// # 示例
    /// ```ignore
    /// use pptx::oxml::{Fill, GradientStop, GradientType};
    /// use pptx::units::RGBColor;
    /// use pptx::oxml::color::Color;
    ///
    /// let mut fill = Fill::Inherit;
    /// let mut fmt = FillFormat::new(&mut fill);
    /// let g = fmt.gradient();
    /// g.stops.push(GradientStop { pos: 0, color: Color::RGB(RGBColor::new(0xFF, 0x00, 0x00)) });
    /// g.stops.push(GradientStop { pos: 100000, color: Color::RGB(RGBColor::new(0x00, 0x00, 0xFF)) });
    /// g.gradient_type = GradientType::Linear(2_700_000); // 向下
    /// ```
    pub fn gradient(&mut self) -> &mut GradientFill {
        if !matches!(self.fill, Fill::Gradient(_)) {
            *self.fill = Fill::Gradient(GradientFill {
                stops: Vec::new(),
                gradient_type: GradientType::Linear(0),
                flip: None,
                rot_with_shape: None,
            });
        }
        match &mut self.fill {
            Fill::Gradient(g) => g,
            _ => unreachable!("just set to Gradient"),
        }
    }

    /// 便捷：直接设成**线性渐变**。
    ///
    /// # 参数
    /// - `stops`：光轨列表（至少 2 个）。
    /// - `angle`：角度（1/60000 度，0 = 向右，5400000 = 向下）。
    pub fn set_gradient_linear(&mut self, stops: Vec<GradientStop>, angle: i32) {
        *self.fill = Fill::Gradient(GradientFill {
            stops,
            gradient_type: GradientType::Linear(angle),
            flip: None,
            rot_with_shape: None,
        });
    }

    /// 便捷：直接设成**路径渐变**。
    ///
    /// # 参数
    /// - `stops`：光轨列表。
    /// - `path`：路径形状（`Circle` / `Rect` / `Shape`）。
    pub fn set_gradient_path(&mut self, stops: Vec<GradientStop>, path: GradientPath) {
        *self.fill = Fill::Gradient(GradientFill {
            stops,
            gradient_type: GradientType::Path(path),
            flip: None,
            rot_with_shape: None,
        });
    }

    // ===== 图案填充 API（TODO-010）=====

    /// 切到**图案**模式并返回 `&mut PatternFill` 供进一步配置。
    ///
    /// 调用本方法会把 `Fill` 切到 `Pattern(PatternFill::default())`，
    /// 调用方随后通过返回的引用设置 prst / fg_color / bg_color。
    pub fn pattern(&mut self) -> &mut PatternFill {
        if !matches!(self.fill, Fill::Pattern(_)) {
            *self.fill = Fill::Pattern(PatternFill {
                prst: String::new(),
                fg_color: Color::None,
                bg_color: Color::None,
            });
        }
        match &mut self.fill {
            Fill::Pattern(p) => p,
            _ => unreachable!("just set to Pattern"),
        }
    }

    /// 便捷：直接设成图案填充。
    ///
    /// # 参数
    /// - `prst`：预置图案类型（如 `"pct5"` / `"horz"` / `"vert"` / `"cross"`）。
    /// - `fg`：前景色。
    /// - `bg`：背景色。
    pub fn set_pattern(&mut self, prst: impl Into<String>, fg: Color, bg: Color) {
        *self.fill = Fill::Pattern(PatternFill {
            prst: prst.into(),
            fg_color: fg,
            bg_color: bg,
        });
    }
}

impl<'a> From<&'a mut Fill> for FillFormat<'a> {
    fn from(f: &'a mut Fill) -> Self {
        FillFormat::new(f)
    }
}

/// 线框高阶视图（`pptx.dml.line.LineFormat`）。
///
/// # 与 python-pptx 的对应
///
/// - `pptx.dml.line.LineFormat` ←→ [`LineFormat`]；
/// - `line.color.rgb = ...` / `line.width = Pt(1)` / `line.dash_style = MSO_LINE.DASH`
///   ←→ [`LineFormat::color`] / [`LineFormat::set_width`] / [`LineFormat::set_dash_style`]
#[derive(Debug)]
pub struct LineFormat<'a> {
    /// 底层 [`Line`] 引用。
    line: &'a mut Line,
}

impl<'a> LineFormat<'a> {
    /// 构造一个 line 视图。
    pub fn new(line: &'a mut Line) -> Self {
        LineFormat { line }
    }

    /// 底层 line 不可变引用。
    pub fn line(&self) -> &Line {
        self.line
    }
    /// 底层 line 可变引用。
    pub fn line_mut(&mut self) -> &mut Line {
        self.line
    }

    /// 颜色 [`ColorFormat`] 代理。
    pub fn color(&mut self) -> ColorFormat<'_> {
        ColorFormat::new(&mut self.line.color)
    }

    /// 读取宽度（EMU）。
    pub fn width(&self) -> Option<crate::units::Emu> {
        self.line.width
    }
    /// 设置宽度（EMU）。通常 `Pt(1.0).emu()` 即可。
    pub fn set_width(&mut self, w: crate::units::Emu) {
        self.line.width = Some(w);
    }

    /// 读取线型。
    pub fn dash_style(&self) -> Option<Dash> {
        self.line.dash
    }
    /// 设置线型（用 [`super::simpletypes::MsoLineDashStyle`]，自动转 [`Dash`]）。
    pub fn set_dash_style(&mut self, s: super::simpletypes::MsoLineDashStyle) {
        self.line.dash = Some(s.into());
    }

    /// 设为无填充（透明线框）。
    pub fn set_no_fill(&mut self) {
        self.line.no_fill = true;
        self.line.color = crate::oxml::color::Color::None;
    }

    /// 便捷：设宽度为磅值。
    pub fn set_width_pt(&mut self, pt: crate::units::Pt) {
        self.line.width = Some(crate::units::Emu((pt.value() * 12_700.0) as i64));
    }
}

impl<'a> From<&'a mut Line> for LineFormat<'a> {
    fn from(l: &'a mut Line) -> Self {
        LineFormat::new(l)
    }
}

// ===================== 三维效果（TODO-050） =====================
//
// OOXML 中 `<a:scene3d>` / `<a:sp3d>` 位于 `<p:spPr>` 的 effectLst 之后、extLst 之前。
// - scene3d 定义相机与光照（场景级 3D）
// - sp3d 定义形状本身的 3D 拉伸/棱台/材质（形状级 3D）
//
// 典型 XML 结构（来自 Office 默认 fmtScheme 第三个 effectStyle）：
// ```xml
// <a:scene3d>
//   <a:camera prst="orthographicFront">
//     <a:rot lat="0" lon="0" rev="0"/>
//   </a:camera>
//   <a:lightRig rig="threePt" dir="t">
//     <a:rot lat="0" lon="0" rev="1200000"/>
//   </a:lightRig>
// </a:scene3d>
// <a:sp3d>
//   <a:bevelT w="63500" h="25400"/>
// </a:sp3d>
// ```
//
// 所有角度单位为"60000 分之一度"（OOXML ST_Angle），与 Transform::rot 一致。

/// 三维旋转（`<a:rot>`，OOXML CT_SphereCoords）。
///
/// 三个角度均以 1/60000 度为单位：
/// - `lat`：纬度（-90°~90°，即 -5400000~5400000）
/// - `lon`：经度（0°~360°，即 0~21600000）
/// - `rev`：滚转（沿视线轴的旋转）
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Rotation3d {
    /// 纬度（1/60000 度）。
    pub lat: i32,
    /// 经度（1/60000 度）。
    pub lon: i32,
    /// 滚转（1/60000 度）。
    pub rev: i32,
}

impl Rotation3d {
    /// 写出 `<a:rot lat="..." lon="..." rev="..."/>`（自闭合）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let lat = self.lat.to_string();
        let lon = self.lon.to_string();
        let rev = self.rev.to_string();
        w.empty_with(
            "a:rot",
            &[
                ("lat", lat.as_str()),
                ("lon", lon.as_str()),
                ("rev", rev.as_str()),
            ],
        );
    }
}

/// 相机预设类型（OOXML ST_PresetCameraType，常用子集）。
///
/// 完整列表参见 ECMA-376 Part 1 §20.1.10.13。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum CameraPreset {
    /// `orthographicFront`（默认，正交前视图）。
    #[default]
    OrthographicFront,
    /// `isometricOffAxis1` ~ `isometricOffAxis4`：等轴侧视图。
    IsometricOffAxis1,
    IsometricOffAxis2,
    IsometricOffAxis3,
    IsometricOffAxis4,
    /// `isometricLeftDown` / `isometricLeftUp` / `isometricRightDown` / `isometricRightUp`。
    IsometricLeftDown,
    IsometricLeftUp,
    IsometricRightDown,
    IsometricRightUp,
    /// `perspectiveFront`：透视前视图。
    PerspectiveFront,
    /// `perspectiveLeft` / `perspectiveRight`：透视侧视图。
    PerspectiveLeft,
    PerspectiveRight,
    /// 其它未显式枚举的预设（保留原始字符串）。
    Other(String),
}

impl CameraPreset {
    /// 转 OOXML 字符串。
    pub fn as_str(&self) -> &str {
        match self {
            CameraPreset::OrthographicFront => "orthographicFront",
            CameraPreset::IsometricOffAxis1 => "isometricOffAxis1",
            CameraPreset::IsometricOffAxis2 => "isometricOffAxis2",
            CameraPreset::IsometricOffAxis3 => "isometricOffAxis3",
            CameraPreset::IsometricOffAxis4 => "isometricOffAxis4",
            CameraPreset::IsometricLeftDown => "isometricLeftDown",
            CameraPreset::IsometricLeftUp => "isometricLeftUp",
            CameraPreset::IsometricRightDown => "isometricRightDown",
            CameraPreset::IsometricRightUp => "isometricRightUp",
            CameraPreset::PerspectiveFront => "perspectiveFront",
            CameraPreset::PerspectiveLeft => "perspectiveLeft",
            CameraPreset::PerspectiveRight => "perspectiveRight",
            CameraPreset::Other(s) => s.as_str(),
        }
    }

    /// 从字符串解析（不区分大小写、未识别则归入 `Other`）。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Self {
        match s {
            "orthographicFront" => CameraPreset::OrthographicFront,
            "isometricOffAxis1" => CameraPreset::IsometricOffAxis1,
            "isometricOffAxis2" => CameraPreset::IsometricOffAxis2,
            "isometricOffAxis3" => CameraPreset::IsometricOffAxis3,
            "isometricOffAxis4" => CameraPreset::IsometricOffAxis4,
            "isometricLeftDown" => CameraPreset::IsometricLeftDown,
            "isometricLeftUp" => CameraPreset::IsometricLeftUp,
            "isometricRightDown" => CameraPreset::IsometricRightDown,
            "isometricRightUp" => CameraPreset::IsometricRightUp,
            "perspectiveFront" => CameraPreset::PerspectiveFront,
            "perspectiveLeft" => CameraPreset::PerspectiveLeft,
            "perspectiveRight" => CameraPreset::PerspectiveRight,
            other => CameraPreset::Other(other.to_string()),
        }
    }
}

/// 相机（`<a:camera>`，OOXML CT_Camera）。
///
/// 定义观察形状的虚拟相机：预设类型、视野（FOV）、缩放、可选旋转。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Camera {
    /// 相机预设类型（默认 `orthographicFront`）。
    pub preset: CameraPreset,
    /// 视野角度（1/60000 度，0 表示使用预设默认）。
    pub fov: i32,
    /// 缩放（百分比 * 1000，100000 = 100%）。
    pub zoom: i32,
    /// 可选旋转（覆盖预设默认视角）。
    pub rotation: Option<Rotation3d>,
}

impl Camera {
    /// 写出 `<a:camera prst="..." fov="..." zoom="...">...</a:camera>`。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let prst = self.preset.as_str();
        let mut attrs: Vec<(&str, &str)> = vec![("prst", prst)];
        let fov_s;
        if self.fov != 0 {
            fov_s = self.fov.to_string();
            attrs.push(("fov", fov_s.as_str()));
        }
        let zoom_s;
        if self.zoom != 0 && self.zoom != 100000 {
            zoom_s = self.zoom.to_string();
            attrs.push(("zoom", zoom_s.as_str()));
        }
        if let Some(rot) = &self.rotation {
            w.open_with("a:camera", &attrs);
            rot.write_xml(w);
            w.close("a:camera");
        } else {
            w.empty_with("a:camera", &attrs);
        }
    }
}

/// 光照设备预设类型（OOXML ST_LightRigType，常用子集）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum LightRigType {
    /// `balanced`（平衡光，默认）。
    #[default]
    Balanced,
    /// `bright`：明亮。
    Bright,
    /// `chilly`：冷光。
    Chilly,
    /// `contrasting`：对比光。
    Contrasting,
    /// `flat`：平光。
    Flat,
    /// `harsh`：强光。
    Harsh,
    /// `morning`：晨光。
    Morning,
    /// `soft`：柔光。
    Soft,
    /// `sunrise`：日出光。
    Sunrise,
    /// `sunset`：日落光。
    Sunset,
    /// `threePt`：三点光（Office 默认 effectStyle 中使用）。
    ThreePt,
    /// `twoPt`：两点光。
    TwoPt,
    /// 其它未显式枚举的光照类型（保留原始字符串）。
    Other(String),
}

impl LightRigType {
    /// 转 OOXML 字符串。
    pub fn as_str(&self) -> &str {
        match self {
            LightRigType::Balanced => "balanced",
            LightRigType::Bright => "bright",
            LightRigType::Chilly => "chilly",
            LightRigType::Contrasting => "contrasting",
            LightRigType::Flat => "flat",
            LightRigType::Harsh => "harsh",
            LightRigType::Morning => "morning",
            LightRigType::Soft => "soft",
            LightRigType::Sunrise => "sunrise",
            LightRigType::Sunset => "sunset",
            LightRigType::ThreePt => "threePt",
            LightRigType::TwoPt => "twoPt",
            LightRigType::Other(s) => s.as_str(),
        }
    }

    /// 从字符串解析。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Self {
        match s {
            "balanced" => LightRigType::Balanced,
            "bright" => LightRigType::Bright,
            "chilly" => LightRigType::Chilly,
            "contrasting" => LightRigType::Contrasting,
            "flat" => LightRigType::Flat,
            "harsh" => LightRigType::Harsh,
            "morning" => LightRigType::Morning,
            "soft" => LightRigType::Soft,
            "sunrise" => LightRigType::Sunrise,
            "sunset" => LightRigType::Sunset,
            "threePt" => LightRigType::ThreePt,
            "twoPt" => LightRigType::TwoPt,
            other => LightRigType::Other(other.to_string()),
        }
    }
}

/// 光照方向（OOXML ST_LightRigDirection）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum LightRigDirection {
    /// `tl`：左上。
    TopLeft,
    /// `t`：上（默认，Office effectStyle 中使用）。
    #[default]
    Top,
    /// `tr`：右上。
    TopRight,
    /// `l`：左。
    Left,
    /// `r`：右。
    Right,
    /// `bl`：左下。
    BottomLeft,
    /// `b`：下。
    Bottom,
    /// `br`：右下。
    BottomRight,
}

impl LightRigDirection {
    /// 转 OOXML 字符串。
    pub fn as_str(&self) -> &str {
        match self {
            LightRigDirection::TopLeft => "tl",
            LightRigDirection::Top => "t",
            LightRigDirection::TopRight => "tr",
            LightRigDirection::Left => "l",
            LightRigDirection::Right => "r",
            LightRigDirection::BottomLeft => "bl",
            LightRigDirection::Bottom => "b",
            LightRigDirection::BottomRight => "br",
        }
    }

    /// 从字符串解析。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Self {
        match s {
            "tl" => LightRigDirection::TopLeft,
            "t" => LightRigDirection::Top,
            "tr" => LightRigDirection::TopRight,
            "l" => LightRigDirection::Left,
            "r" => LightRigDirection::Right,
            "bl" => LightRigDirection::BottomLeft,
            "b" => LightRigDirection::Bottom,
            "br" => LightRigDirection::BottomRight,
            _ => LightRigDirection::Top,
        }
    }
}

/// 光照设备（`<a:lightRig>`，OOXML CT_LightRig）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LightRig {
    /// 光照类型（默认 `balanced`）。
    pub rig: LightRigType,
    /// 光照方向（默认 `t`）。
    pub dir: LightRigDirection,
    /// 可选旋转（覆盖预设默认）。
    pub rotation: Option<Rotation3d>,
}

impl LightRig {
    /// 写出 `<a:lightRig rig="..." dir="...">...</a:lightRig>`。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let rig = self.rig.as_str();
        let dir = self.dir.as_str();
        if let Some(rot) = &self.rotation {
            w.open_with("a:lightRig", &[("rig", rig), ("dir", dir)]);
            rot.write_xml(w);
            w.close("a:lightRig");
        } else {
            w.empty_with("a:lightRig", &[("rig", rig), ("dir", dir)]);
        }
    }
}

/// 三维场景（`<a:scene3d>`，OOXML CT_Scene3D）。
///
/// 包含相机与光照，定义观察形状的虚拟 3D 环境。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Scene3d {
    /// 相机（必填，默认 `orthographicFront`）。
    pub camera: Camera,
    /// 光照设备（必填，默认 `balanced` + `t`）。
    pub light_rig: LightRig,
    /// 三维场景背景（`<a:backdrop>`，可选，TODO-050）。
    ///
    /// 定义 6 个背景平面（地板/墙壁/左/右/顶/底），启用的平面会渲染为可见的背景面。
    /// `None` 表示不写出 `<a:backdrop>` 元素（与 Office 默认行为一致）。
    pub backdrop: Option<Backdrop>,
}

impl Scene3d {
    /// 写出 `<a:scene3d>...</a:scene3d>`。
    ///
    /// # 元素顺序（OOXML CT_Scene3D）
    ///
    /// ```text
    /// <a:scene3d>
    ///   <a:camera>...</a:camera>        ← 必填
    ///   <a:lightRig>...</a:lightRig>    ← 必填
    ///   <a:backdrop>...</a:backdrop>    ← 可选（TODO-050）
    /// </a:scene3d>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("a:scene3d");
        self.camera.write_xml(w);
        self.light_rig.write_xml(w);
        if let Some(b) = &self.backdrop {
            b.write_xml(w);
        }
        w.close("a:scene3d");
    }
}

/// 三维点（`<a:anchor>`，OOXML CT_Point3D）。
///
/// 用于 [`Backdrop`] 的锚点位置。三个坐标均以 EMU 为单位。
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Point3d {
    /// X 坐标（EMU）。
    pub x: i32,
    /// Y 坐标（EMU）。
    pub y: i32,
    /// Z 坐标（EMU）。
    pub z: i32,
}

impl Point3d {
    /// 写出 `<a:anchor x="..." y="..." z="..."/>`（自闭合）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let xs = self.x.to_string();
        let ys = self.y.to_string();
        let zs = self.z.to_string();
        w.empty_with(
            "a:anchor",
            &[("x", xs.as_str()), ("y", ys.as_str()), ("z", zs.as_str())],
        );
    }
}

/// 三维场景背景（`<a:backdrop>`，OOXML CT_Backdrop，TODO-050）。
///
/// 定义 3D 场景中的 6 个背景平面，每个平面可独立启用/禁用。
/// 启用的平面会渲染为可见的背景面（如地板/墙壁），用于营造空间感。
///
/// # OOXML 元素顺序
///
/// ```text
/// <a:backdrop>
///   <a:anchor x="..." y="..." z="..."/>   ← 可选：锚点位置
///   <a:floor/>                            ← 可选：地板平面
///   <a:wall/>                             ← 可选：后墙平面
///   <a:l/>                                ← 可选：左平面
///   <a:r/>                                ← 可选：右平面
///   <a:t/>                                ← 可选：顶平面
///   <a:b/>                                ← 可选：底平面
/// </a:backdrop>
/// ```
///
/// # 与 python-pptx 的对应
///
/// python-pptx 不支持 backdrop 编辑；本结构是 pptx-rs 的扩展能力。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Backdrop {
    /// 锚点位置（`<a:anchor>`）。
    ///
    /// 定义背景平面集合的参考原点（EMU 单位）。`None` 表示不写出 `<a:anchor>`。
    pub anchor: Option<Point3d>,
    /// 是否启用地板平面（`<a:floor/>`）。
    pub floor: bool,
    /// 是否启用后墙平面（`<a:wall/>`）。
    pub wall: bool,
    /// 是否启用左平面（`<a:l/>`）。
    pub left: bool,
    /// 是否启用右平面（`<a:r/>`）。
    pub right: bool,
    /// 是否启用顶平面（`<a:t/>`）。
    pub top: bool,
    /// 是否启用底平面（`<a:b/>`）。
    pub bottom: bool,
}

impl Backdrop {
    /// 写出 `<a:backdrop>...</a:backdrop>`。
    ///
    /// 仅写出启用（`true`）的平面元素，未启用的平面不写出。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("a:backdrop");
        if let Some(a) = &self.anchor {
            a.write_xml(w);
        }
        if self.floor {
            w.empty("a:floor");
        }
        if self.wall {
            w.empty("a:wall");
        }
        if self.left {
            w.empty("a:l");
        }
        if self.right {
            w.empty("a:r");
        }
        if self.top {
            w.empty("a:t");
        }
        if self.bottom {
            w.empty("a:b");
        }
        w.close("a:backdrop");
    }
}

/// 棱台（`<a:bevelT>` / `<a:bevelB>`，OOXML CT_Bevel）。
///
/// 宽高均以 EMU 为单位（典型值 63500 = 5pt）。
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Bevel {
    /// 棱台宽度（EMU）。
    pub w: i32,
    /// 棱台高度（EMU）。
    pub h: i32,
}

impl Bevel {
    /// 写出 `<a:bevelT w="..." h="..."/>`（自闭合，tag 由调用方指定）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter, tag: &str) {
        let w_s = self.w.to_string();
        let h_s = self.h.to_string();
        w.empty_with(tag, &[("w", w_s.as_str()), ("h", h_s.as_str())]);
    }
}

/// 形状材质预设（OOXML ST_PresetMaterialType，常用子集）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum MaterialPreset {
    /// `warmMatte`（暖色哑光，默认）。
    #[default]
    WarmMatte,
    /// `clear`：透明。
    Clear,
    /// `dkEdge`：暗边缘。
    DarkEdge,
    /// `flat`：平面。
    Flat,
    /// `legacyMatte`：旧版哑光。
    LegacyMatte,
    /// `legacyMetallic`：旧版金属。
    LegacyMetallic,
    /// `legacyPlastic`：旧版塑料。
    LegacyPlastic,
    /// `legacyWireframe`：旧版线框。
    LegacyWireframe,
    /// `matte`：哑光。
    Matte,
    /// `metallic`：金属。
    Metallic,
    /// `plastic`：塑料。
    Plastic,
    /// `powder`：粉末。
    Powder,
    /// `softEdge`：柔边。
    SoftEdge,
    /// `softmetal`：软金属。
    SoftMetal,
    /// 其它未显式枚举的材质（保留原始字符串）。
    Other(String),
}

impl MaterialPreset {
    /// 转 OOXML 字符串。
    pub fn as_str(&self) -> &str {
        match self {
            MaterialPreset::WarmMatte => "warmMatte",
            MaterialPreset::Clear => "clear",
            MaterialPreset::DarkEdge => "dkEdge",
            MaterialPreset::Flat => "flat",
            MaterialPreset::LegacyMatte => "legacyMatte",
            MaterialPreset::LegacyMetallic => "legacyMetallic",
            MaterialPreset::LegacyPlastic => "legacyPlastic",
            MaterialPreset::LegacyWireframe => "legacyWireframe",
            MaterialPreset::Matte => "matte",
            MaterialPreset::Metallic => "metallic",
            MaterialPreset::Plastic => "plastic",
            MaterialPreset::Powder => "powder",
            MaterialPreset::SoftEdge => "softEdge",
            MaterialPreset::SoftMetal => "softmetal",
            MaterialPreset::Other(s) => s.as_str(),
        }
    }

    /// 从字符串解析。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Self {
        match s {
            "warmMatte" => MaterialPreset::WarmMatte,
            "clear" => MaterialPreset::Clear,
            "dkEdge" => MaterialPreset::DarkEdge,
            "flat" => MaterialPreset::Flat,
            "legacyMatte" => MaterialPreset::LegacyMatte,
            "legacyMetallic" => MaterialPreset::LegacyMetallic,
            "legacyPlastic" => MaterialPreset::LegacyPlastic,
            "legacyWireframe" => MaterialPreset::LegacyWireframe,
            "matte" => MaterialPreset::Matte,
            "metallic" => MaterialPreset::Metallic,
            "plastic" => MaterialPreset::Plastic,
            "powder" => MaterialPreset::Powder,
            "softEdge" => MaterialPreset::SoftEdge,
            "softmetal" => MaterialPreset::SoftMetal,
            other => MaterialPreset::Other(other.to_string()),
        }
    }
}

/// 形状 3D 属性（`<a:sp3d>`，OOXML CT_Shape3D）。
///
/// 定义形状本身的 3D 拉伸、棱台、材质。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Sp3d {
    /// 拉伸高度（EMU，典型值 38100 = 3pt）。
    pub extrusion_h: i32,
    /// 轮廓宽度（EMU）。
    pub contour_w: i32,
    /// 材质预设（默认 `warmMatte`）。
    pub prst_material: MaterialPreset,
    /// 顶部棱台（`<a:bevelT>`）。
    pub bevel_top: Option<Bevel>,
    /// 底部棱台（`<a:bevelB>`）。
    pub bevel_bottom: Option<Bevel>,
    /// 拉伸颜色（`<a:extrusionClr>`，`None` 表示不写出）。
    pub extrusion_color: Option<Color>,
    /// 轮廓颜色（`<a:contourClr>`，`None` 表示不写出）。
    pub contour_color: Option<Color>,
}

impl Sp3d {
    /// 写出 `<a:sp3d>...</a:sp3d>`。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        let eh_s;
        if self.extrusion_h != 0 {
            eh_s = self.extrusion_h.to_string();
            attrs.push(("extrusionH", eh_s.as_str()));
        }
        let cw_s;
        if self.contour_w != 0 {
            cw_s = self.contour_w.to_string();
            attrs.push(("contourW", cw_s.as_str()));
        }
        // prstMaterial 仅在非默认时输出（避免覆盖默认值）
        let pm_s;
        if !matches!(self.prst_material, MaterialPreset::WarmMatte) {
            pm_s = self.prst_material.as_str().to_string();
            attrs.push(("prstMaterial", pm_s.as_str()));
        }
        let has_children = self.bevel_top.is_some()
            || self.bevel_bottom.is_some()
            || self.extrusion_color.is_some()
            || self.contour_color.is_some();
        if !has_children && attrs.is_empty() {
            w.empty("a:sp3d");
            return;
        }
        if !has_children {
            w.empty_with("a:sp3d", &attrs);
            return;
        }
        w.open_with("a:sp3d", &attrs);
        if let Some(b) = &self.bevel_top {
            b.write_xml(w, "a:bevelT");
        }
        if let Some(b) = &self.bevel_bottom {
            b.write_xml(w, "a:bevelB");
        }
        if let Some(c) = &self.extrusion_color {
            w.open("a:extrusionClr");
            c.write_solid_fill(w);
            w.close("a:extrusionClr");
        }
        if let Some(c) = &self.contour_color {
            w.open("a:contourClr");
            c.write_solid_fill(w);
            w.close("a:contourClr");
        }
        w.close("a:sp3d");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::color::Color;
    use crate::oxml::writer::XmlWriter;
    use crate::units::RGBColor;

    // --------------------- 调整值测试（TODO-038） ---------------------

    /// `AdjustmentValue::new` 正确设置字段。
    #[test]
    fn adjustment_value_new() {
        let av = AdjustmentValue::new("adj", 16667);
        assert_eq!(av.name, "adj");
        assert_eq!(av.raw_value, 16667);
    }

    /// `AdjustmentValue::from_normalized` 正确转换归一化值。
    #[test]
    fn adjustment_value_from_normalized() {
        let av = AdjustmentValue::from_normalized("adj", 0.5);
        assert_eq!(av.raw_value, 50000);
        let av2 = AdjustmentValue::from_normalized("adj1", 0.25);
        assert_eq!(av2.raw_value, 25000);
    }

    /// `AdjustmentValue::effective_value` 正确返回归一化值。
    #[test]
    fn adjustment_value_effective_value() {
        let av = AdjustmentValue::new("adj", 16667);
        assert!((av.effective_value() - 0.16667).abs() < 0.0001);
    }

    /// `AdjustmentValue::write_xml` 正确序列化 `<a:gd>` 元素。
    #[test]
    fn adjustment_value_write_xml() {
        let av = AdjustmentValue::new("adj", 16667);
        let mut w = XmlWriter::new();
        av.write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains("<a:gd"), "xml: {}", xml);
        assert!(xml.contains(r#"name="adj""#), "xml: {}", xml);
        assert!(xml.contains(r#"fmla="val 16667""#), "xml: {}", xml);
    }

    /// `Geometry::Preset` 带调整值正确序列化 `<a:avLst>`。
    #[test]
    fn geometry_preset_with_adjustments_write_xml() {
        let geom = Geometry::Preset(
            PresetGeometry::RoundRectangle,
            vec![AdjustmentValue::new("adj", 16667)],
        );
        let mut w = XmlWriter::new();
        geom.write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains(r#"prst="roundRect""#), "xml: {}", xml);
        assert!(xml.contains("<a:avLst>"), "xml: {}", xml);
        assert!(xml.contains(r#"fmla="val 16667""#), "xml: {}", xml);
        assert!(xml.contains("</a:avLst>"), "xml: {}", xml);
    }

    /// `Geometry::Preset` 无调整值时写出空 `<a:avLst/>`。
    #[test]
    fn geometry_preset_no_adjustments_write_xml() {
        let geom = Geometry::preset(PresetGeometry::Rectangle);
        let mut w = XmlWriter::new();
        geom.write_xml(&mut w);
        assert!(w.buf.contains("<a:avLst/>"), "xml: {}", w.buf);
    }

    /// `Geometry::preset` 辅助方法创建无调整值的预设几何。
    #[test]
    fn geometry_preset_helper() {
        let geom = Geometry::preset(PresetGeometry::Ellipse);
        match geom {
            Geometry::Preset(p, adj) => {
                assert_eq!(p, PresetGeometry::Ellipse);
                assert!(adj.is_empty());
            }
            _ => panic!("应为 Preset 变体"),
        }
    }

    // --------------------- 其他测试 ---------------------

    /// 验证 `FillFormat::gradient()` 切到渐变模式并返回可变引用。
    #[test]
    fn fill_format_gradient_switches_mode() {
        let mut fill = Fill::Inherit;
        let mut fmt = FillFormat::new(&mut fill);
        let g = fmt.gradient();
        g.stops.push(GradientStop {
            pos: 0,
            color: Color::RGB(RGBColor(0xFF, 0x00, 0x00)),
        });
        g.stops.push(GradientStop {
            pos: 100_000,
            color: Color::RGB(RGBColor(0x00, 0x00, 0xFF)),
        });
        g.gradient_type = GradientType::Linear(2_700_000);
        assert!(matches!(fill, Fill::Gradient(_)));
        if let Fill::Gradient(g) = &fill {
            assert_eq!(g.stops.len(), 2);
            assert_eq!(g.stops[0].pos, 0);
            assert_eq!(g.stops[1].pos, 100_000);
            assert!(matches!(g.gradient_type, GradientType::Linear(2_700_000)));
        }
    }

    /// 验证 `FillFormat::set_gradient_linear` 便捷方法。
    #[test]
    fn fill_format_set_gradient_linear() {
        let mut fill = Fill::Inherit;
        let mut fmt = FillFormat::new(&mut fill);
        let stops = vec![
            GradientStop {
                pos: 0,
                color: Color::RGB(RGBColor(0xFF, 0x00, 0x00)),
            },
            GradientStop {
                pos: 100_000,
                color: Color::RGB(RGBColor(0x00, 0xFF, 0x00)),
            },
        ];
        fmt.set_gradient_linear(stops, 5_400_000);
        assert!(matches!(fill, Fill::Gradient(_)));
        if let Fill::Gradient(g) = &fill {
            assert_eq!(g.stops.len(), 2);
            assert!(matches!(g.gradient_type, GradientType::Linear(5_400_000)));
        }
    }

    /// 验证 `FillFormat::set_gradient_path` 便捷方法。
    #[test]
    fn fill_format_set_gradient_path() {
        let mut fill = Fill::Inherit;
        let mut fmt = FillFormat::new(&mut fill);
        let stops = vec![GradientStop {
            pos: 50_000,
            color: Color::RGB(RGBColor(0x80, 0x80, 0x80)),
        }];
        fmt.set_gradient_path(stops, GradientPath::Circle);
        if let Fill::Gradient(g) = &fill {
            assert!(matches!(
                g.gradient_type,
                GradientType::Path(GradientPath::Circle)
            ));
        }
    }

    /// 验证 `FillFormat::pattern()` 切到图案模式并返回可变引用。
    #[test]
    fn fill_format_pattern_switches_mode() {
        let mut fill = Fill::Inherit;
        let mut fmt = FillFormat::new(&mut fill);
        let p = fmt.pattern();
        p.prst = "horz".to_string();
        p.fg_color = Color::RGB(RGBColor(0xFF, 0x00, 0x00));
        p.bg_color = Color::RGB(RGBColor(0xFF, 0xFF, 0xFF));
        assert!(matches!(fill, Fill::Pattern(_)));
        if let Fill::Pattern(p) = &fill {
            assert_eq!(p.prst, "horz");
        }
    }

    /// 验证 `FillFormat::set_pattern` 便捷方法。
    #[test]
    fn fill_format_set_pattern() {
        let mut fill = Fill::Inherit;
        let mut fmt = FillFormat::new(&mut fill);
        fmt.set_pattern(
            "cross",
            Color::RGB(RGBColor(0x00, 0x00, 0x00)),
            Color::RGB(RGBColor(0xFF, 0xFF, 0xFF)),
        );
        if let Fill::Pattern(p) = &fill {
            assert_eq!(p.prst, "cross");
            assert!(matches!(p.fg_color, Color::RGB(_)));
            assert!(matches!(p.bg_color, Color::RGB(_)));
        } else {
            panic!("应为 Pattern 变体");
        }
    }

    /// 验证 `fill_type()` 对渐变和图案返回正确的枚举值。
    #[test]
    fn fill_format_fill_type_for_gradient_and_pattern() {
        let mut fill = Fill::Gradient(GradientFill {
            stops: vec![],
            gradient_type: GradientType::Linear(0),
            flip: None,
            rot_with_shape: None,
        });
        let fmt = FillFormat::new(&mut fill);
        assert_eq!(
            fmt.fill_type(),
            crate::oxml::simpletypes::MsoFillType::Gradient
        );

        let mut fill = Fill::Pattern(PatternFill::default());
        let fmt = FillFormat::new(&mut fill);
        assert_eq!(
            fmt.fill_type(),
            crate::oxml::simpletypes::MsoFillType::Pattern
        );
    }

    /// 验证 `EffectList::is_empty` 和 `write_xml`（外阴影）。
    #[test]
    fn effect_list_outer_shadow_serialization() {
        let mut effects = EffectList::default();
        assert!(effects.is_empty());
        effects.outer_shadow = Some(ShadowEffect {
            dir: 2_700_000,
            dist: 38100,
            blur_rad: 40000,
            color: Color::RGB(RGBColor(0x00, 0x00, 0x00)),
            rot_with_shape: Some(false),
        });
        assert!(!effects.is_empty());

        let mut w = XmlWriter::new();
        effects.write_xml(&mut w);
        let buf = &w.buf;
        assert!(buf.contains("<a:effectLst>"));
        assert!(buf.contains("<a:outerShdw"));
        assert!(buf.contains("dir=\"2700000\""));
        assert!(buf.contains("dist=\"38100\""));
        assert!(buf.contains("blurRad=\"40000\""));
        assert!(buf.contains("rotWithShape=\"0\""));
        assert!(buf.contains("<a:srgbClr val=\"000000\""));
        assert!(buf.contains("</a:effectLst>"));
    }

    /// 验证 `EffectList` 发光和柔化边缘序列化。
    #[test]
    fn effect_list_glow_and_soft_edge() {
        let effects = EffectList {
            glow: Some(GlowEffect {
                rad: 50000,
                color: Color::RGB(RGBColor(0xFF, 0x00, 0xFF)),
            }),
            soft_edge: Some(SoftEdgeEffect { rad: 25000 }),
            ..Default::default()
        };

        let mut w = XmlWriter::new();
        effects.write_xml(&mut w);
        let buf = &w.buf;
        assert!(buf.contains("<a:glow rad=\"50000\">"));
        assert!(buf.contains("<a:srgbClr val=\"FF00FF\""));
        assert!(buf.contains("<a:softEdge rad=\"25000\"/>"));
    }

    /// 验证 `EffectList` 反射效果序列化（仅属性，无子元素）。
    #[test]
    fn effect_list_reflection_serialization() {
        let effects = EffectList {
            reflection: Some(ReflectionEffect {
                blur_rad: Some(50000),
                st_a: Some(52000),
                st_pos: Some(0),
                end_a: Some(30000),
                end_pos: Some(50000),
                dist: Some(38100),
                dir: Some(5_400_000),
                rot_with_shape: None,
            }),
            ..Default::default()
        };

        let mut w = XmlWriter::new();
        effects.write_xml(&mut w);
        let buf = &w.buf;
        assert!(buf.contains("<a:reflection"));
        assert!(buf.contains("blurRad=\"50000\""));
        assert!(buf.contains("stA=\"52000\""));
        assert!(buf.contains("endPos=\"50000\""));
        assert!(buf.contains("dist=\"38100\""));
        assert!(buf.contains("/>")); // 自闭合
    }

    /// 验证空 `EffectList` 不写出任何 XML。
    #[test]
    fn effect_list_empty_writes_nothing() {
        let effects = EffectList::default();
        let mut w = XmlWriter::new();
        effects.write_xml(&mut w);
        assert!(w.buf.is_empty());
    }

    // --------------------- TODO-046: 图片填充模式测试 ---------------------

    /// 验证 `BlipFillMode::Stretch` 序列化。
    #[test]
    fn blip_fill_mode_stretch_serialize() {
        let mode = BlipFillMode::Stretch;
        let mut w = XmlWriter::new();
        mode.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:stretch>"), "应输出 a:stretch，实际: {s}");
        assert!(s.contains("<a:fillRect/>"), "应输出 a:fillRect，实际: {s}");
        assert!(s.contains("</a:stretch>"), "应关闭 a:stretch，实际: {s}");
    }

    /// 验证 `BlipFillMode::Tile` 序列化（带全部属性）。
    #[test]
    fn blip_fill_mode_tile_serialize_full() {
        let mode = BlipFillMode::Tile {
            tx: Some(914400),
            ty: Some(457200),
            sx: Some(100_000),
            sy: Some(50_000),
            flip: Some("x".to_string()),
            algn: Some("tl".to_string()),
        };
        let mut w = XmlWriter::new();
        mode.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:tile"), "应输出 a:tile，实际: {s}");
        assert!(s.contains("tx=\"914400\""), "应包含 tx，实际: {s}");
        assert!(s.contains("ty=\"457200\""), "应包含 ty，实际: {s}");
        assert!(s.contains("sx=\"100000\""), "应包含 sx，实际: {s}");
        assert!(s.contains("sy=\"50000\""), "应包含 sy，实际: {s}");
        assert!(s.contains("flip=\"x\""), "应包含 flip，实际: {s}");
        assert!(s.contains("algn=\"tl\""), "应包含 algn，实际: {s}");
        assert!(s.contains("/>"), "应为自闭合标签，实际: {s}");
    }

    /// 验证 `BlipFillMode::Tile` 序列化（仅部分属性）。
    #[test]
    fn blip_fill_mode_tile_serialize_partial() {
        let mode = BlipFillMode::Tile {
            tx: None,
            ty: None,
            sx: Some(200_000),
            sy: None,
            flip: None,
            algn: Some("ctr".to_string()),
        };
        let mut w = XmlWriter::new();
        mode.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:tile"), "应输出 a:tile，实际: {s}");
        assert!(!s.contains("tx="), "不应包含 tx，实际: {s}");
        assert!(!s.contains("ty="), "不应包含 ty，实际: {s}");
        assert!(s.contains("sx=\"200000\""), "应包含 sx，实际: {s}");
        assert!(!s.contains("sy="), "不应包含 sy，实际: {s}");
        assert!(s.contains("algn=\"ctr\""), "应包含 algn，实际: {s}");
    }

    /// 验证 `BlipFillMode::None` 不写出任何 XML。
    #[test]
    fn blip_fill_mode_none_serialize() {
        let mode = BlipFillMode::None;
        let mut w = XmlWriter::new();
        mode.write_xml(&mut w);
        assert!(w.buf.is_empty(), "None 模式不应写出任何 XML");
    }

    /// 验证 `Fill::Blip` 使用 `BlipFillMode::Stretch` 序列化。
    #[test]
    fn fill_blip_with_stretch_serialize() {
        let fill = Fill::Blip {
            rid: "rIdImg1".to_string(),
            mode: BlipFillMode::Stretch,
        };
        let mut w = XmlWriter::new();
        fill.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:blipFill"), "应输出 a:blipFill，实际: {s}");
        assert!(
            s.contains("r:embed=\"rIdImg1\""),
            "应包含 r:embed，实际: {s}"
        );
        assert!(s.contains("<a:stretch>"), "应输出 a:stretch，实际: {s}");
        // BUG-001 回归检测：<a:blip 标签应只出现 1 次（不应有嵌套的双重标签）
        // 注意：不能用 matches("<a:blip")，因为它会同时匹配 <a:blipFill
        let blip_count = s.matches("<a:blip ").count()
            + s.matches("<a:blip/>").count()
            + s.matches("<a:blip>").count();
        assert_eq!(
            blip_count, 1,
            "应仅输出 1 个 <a:blip 标签，实际 {} 次（可能存在双重标签 bug）: {s}",
            blip_count
        );
        // 同时验证不会出现嵌套的 <a:blip><a:blip
        assert!(
            !s.contains("<a:blip><a:blip"),
            "检测到嵌套的双重 <a:blip> 标签（BUG-001 回归）: {s}"
        );
    }

    /// 验证 `Fill::Blip` 使用 `BlipFillMode::Tile` 序列化。
    #[test]
    fn fill_blip_with_tile_serialize() {
        let fill = Fill::Blip {
            rid: "rIdImg2".to_string(),
            mode: BlipFillMode::Tile {
                tx: Some(100),
                ty: None,
                sx: None,
                sy: None,
                flip: Some("xy".to_string()),
                algn: None,
            },
        };
        let mut w = XmlWriter::new();
        fill.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:blipFill"), "应输出 a:blipFill，实际: {s}");
        assert!(
            s.contains("r:embed=\"rIdImg2\""),
            "应包含 r:embed，实际: {s}"
        );
        assert!(s.contains("<a:tile"), "应输出 a:tile，实际: {s}");
        assert!(s.contains("tx=\"100\""), "应包含 tx，实际: {s}");
        assert!(s.contains("flip=\"xy\""), "应包含 flip，实际: {s}");
    }

    /// 验证 `BlipFillMode` 默认值为 `Stretch`。
    #[test]
    fn blip_fill_mode_default_is_stretch() {
        let mode = BlipFillMode::default();
        assert!(matches!(mode, BlipFillMode::Stretch));
    }

    // --------------------- TODO-050: 三维效果测试 ---------------------

    /// 验证 `Rotation3d` 序列化为自闭合 `<a:rot/>`。
    #[test]
    fn rotation_3d_serialize() {
        let rot = Rotation3d {
            lat: 0,
            lon: 0,
            rev: 1200000,
        };
        let mut w = XmlWriter::new();
        rot.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:rot"), "应输出 a:rot，实际: {s}");
        assert!(s.contains("lat=\"0\""), "应包含 lat，实际: {s}");
        assert!(s.contains("lon=\"0\""), "应包含 lon，实际: {s}");
        assert!(s.contains("rev=\"1200000\""), "应包含 rev，实际: {s}");
        assert!(s.contains("/>"), "应为自闭合，实际: {s}");
    }

    /// 验证 `Camera` 序列化（默认预设、无旋转）。
    #[test]
    fn camera_default_serialize() {
        let camera = Camera::default();
        let mut w = XmlWriter::new();
        camera.write_xml(&mut w);
        let s = &w.buf;
        assert!(
            s.contains("<a:camera prst=\"orthographicFront\""),
            "应包含默认 prst，实际: {s}"
        );
        // 默认 fov=0/zoom=0 不应输出
        assert!(!s.contains("fov="), "默认 fov 不应输出，实际: {s}");
        assert!(!s.contains("zoom="), "默认 zoom 不应输出，实际: {s}");
    }

    /// 验证 `Camera` 带旋转的序列化。
    #[test]
    fn camera_with_rotation_serialize() {
        let camera = Camera {
            preset: CameraPreset::PerspectiveFront,
            fov: 3600000,
            // 注意：zoom=100000 是 OOXML 默认值，write_xml 会省略；
            // 这里用 200000 测试非默认值的输出。
            zoom: 200000,
            rotation: Some(Rotation3d {
                lat: 30,
                lon: 45,
                rev: 0,
            }),
        };
        let mut w = XmlWriter::new();
        camera.write_xml(&mut w);
        let s = &w.buf;
        assert!(
            s.contains("<a:camera prst=\"perspectiveFront\""),
            "应包含 perspectiveFront，实际: {s}"
        );
        assert!(s.contains("fov=\"3600000\""), "应包含 fov，实际: {s}");
        assert!(s.contains("zoom=\"200000\""), "应包含 zoom，实际: {s}");
        assert!(s.contains("<a:rot"), "应包含 a:rot 子元素，实际: {s}");
        assert!(s.contains("</a:camera>"), "应关闭 a:camera，实际: {s}");
    }

    /// 验证 `Scene3d` 完整序列化（相机 + 光照）。
    #[test]
    fn scene3d_full_serialize() {
        let scene = Scene3d {
            camera: Camera {
                preset: CameraPreset::OrthographicFront,
                fov: 0,
                zoom: 0,
                rotation: Some(Rotation3d {
                    lat: 0,
                    lon: 0,
                    rev: 0,
                }),
            },
            light_rig: LightRig {
                rig: LightRigType::ThreePt,
                dir: LightRigDirection::Top,
                rotation: Some(Rotation3d {
                    lat: 0,
                    lon: 0,
                    rev: 1200000,
                }),
            },
            backdrop: None,
        };
        let mut w = XmlWriter::new();
        scene.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:scene3d>"), "应输出 a:scene3d，实际: {s}");
        assert!(
            s.contains("<a:camera prst=\"orthographicFront\">"),
            "应包含 camera 子元素，实际: {s}"
        );
        assert!(
            s.contains("<a:lightRig rig=\"threePt\" dir=\"t\">"),
            "应包含 lightRig 子元素，实际: {s}"
        );
        assert!(s.contains("</a:scene3d>"), "应关闭 a:scene3d，实际: {s}");
    }

    /// 验证 `Sp3d` 默认序列化（仅 prstMaterial 默认值，输出空 `<a:sp3d/>`）。
    #[test]
    fn sp3d_default_serialize() {
        let sp3d = Sp3d::default();
        let mut w = XmlWriter::new();
        sp3d.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:sp3d"), "应输出 a:sp3d，实际: {s}");
        // 默认 warmMatte 不输出 prstMaterial
        assert!(
            !s.contains("prstMaterial="),
            "默认 warmMatte 不应输出，实际: {s}"
        );
    }

    /// 验证 `Sp3d` 带棱台和颜色的序列化。
    #[test]
    fn sp3d_with_bevel_and_colors_serialize() {
        let sp3d = Sp3d {
            extrusion_h: 38100,
            contour_w: 12700,
            prst_material: MaterialPreset::Metallic,
            bevel_top: Some(Bevel { w: 63500, h: 25400 }),
            bevel_bottom: None,
            extrusion_color: Some(Color::RGB(crate::units::RGBColor::BLACK)),
            contour_color: Some(Color::RGB(crate::units::RGBColor::WHITE)),
        };
        let mut w = XmlWriter::new();
        sp3d.write_xml(&mut w);
        let s = &w.buf;
        assert!(
            s.contains("extrusionH=\"38100\""),
            "应包含 extrusionH，实际: {s}"
        );
        assert!(
            s.contains("contourW=\"12700\""),
            "应包含 contourW，实际: {s}"
        );
        assert!(
            s.contains("prstMaterial=\"metallic\""),
            "应包含 prstMaterial，实际: {s}"
        );
        assert!(
            s.contains("<a:bevelT w=\"63500\" h=\"25400\"/>"),
            "应包含 bevelT，实际: {s}"
        );
        assert!(
            s.contains("<a:extrusionClr>"),
            "应包含 extrusionClr，实际: {s}"
        );
        assert!(s.contains("<a:contourClr>"), "应包含 contourClr，实际: {s}");
        assert!(s.contains("</a:sp3d>"), "应关闭 a:sp3d，实际: {s}");
    }

    /// 验证 `ShapeProperties` 中 scene3d/sp3d 在 effectLst 之后输出。
    #[test]
    fn shape_properties_with_3d_serialize() {
        let sp = ShapeProperties {
            scene3d: Some(Scene3d::default()),
            sp3d: Some(Sp3d::default()),
            ..Default::default()
        };
        let mut w = XmlWriter::new();
        sp.write_xml(&mut w, "p:spPr");
        let s = &w.buf;
        // 验证 scene3d 与 sp3d 都被输出
        assert!(s.contains("<a:scene3d>"), "应输出 scene3d，实际: {s}");
        assert!(s.contains("<a:sp3d"), "应输出 sp3d，实际: {s}");
        // 验证顺序：scene3d 在 sp3d 之前
        let scene_pos = s.find("<a:scene3d>").expect("scene3d 应存在");
        let sp3d_pos = s.find("<a:sp3d").expect("sp3d 应存在");
        assert!(
            scene_pos < sp3d_pos,
            "scene3d 应在 sp3d 之前，实际 scene3d@{scene_pos} sp3d@{sp3d_pos}"
        );
    }

    /// 验证 `CameraPreset::parse` 解析已知预设。
    #[test]
    fn camera_preset_from_str() {
        assert!(matches!(
            CameraPreset::parse("orthographicFront"),
            CameraPreset::OrthographicFront
        ));
        assert!(matches!(
            CameraPreset::parse("perspectiveFront"),
            CameraPreset::PerspectiveFront
        ));
        // 未知值归入 Other
        match CameraPreset::parse("customUnknown") {
            CameraPreset::Other(s) => assert_eq!(s, "customUnknown"),
            other => panic!("未知预设应归入 Other，实际: {other:?}"),
        }
    }

    /// 验证 `MaterialPreset::parse` 解析已知材质。
    #[test]
    fn material_preset_from_str() {
        assert!(matches!(
            MaterialPreset::parse("warmMatte"),
            MaterialPreset::WarmMatte
        ));
        assert!(matches!(
            MaterialPreset::parse("metallic"),
            MaterialPreset::Metallic
        ));
        // 未知值归入 Other
        match MaterialPreset::parse("futureMaterial") {
            MaterialPreset::Other(s) => assert_eq!(s, "futureMaterial"),
            other => panic!("未知材质应归入 Other，实际: {other:?}"),
        }
    }

    // ===================== TODO-050 backdrop 背景元素测试 =====================

    /// 验证 `Backdrop` 默认序列化（无平面启用，仅空 backdrop）。
    #[test]
    fn backdrop_default_serialize() {
        let bd = Backdrop::default();
        let mut w = XmlWriter::new();
        bd.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:backdrop>"), "应输出 a:backdrop，实际: {s}");
        assert!(s.contains("</a:backdrop>"), "应关闭 a:backdrop，实际: {s}");
        // 默认所有平面未启用，不应输出 a:floor/a:wall/a:l/a:r/a:t/a:b
        // 注意：用 <a:xxx/> 精确匹配，避免 <a:b 误匹配 <a:backdrop>
        assert!(!s.contains("<a:floor/>"), "默认不应输出 floor: {s}");
        assert!(!s.contains("<a:wall/>"), "默认不应输出 wall: {s}");
        assert!(!s.contains("<a:l/>"), "默认不应输出 l: {s}");
        assert!(!s.contains("<a:r/>"), "默认不应输出 r: {s}");
        assert!(!s.contains("<a:t/>"), "默认不应输出 t: {s}");
        assert!(!s.contains("<a:b/>"), "默认不应输出 b: {s}");
    }

    /// 验证 `Backdrop` 带锚点 + 全平面启用的序列化。
    #[test]
    fn backdrop_with_anchor_and_all_planes_serialize() {
        let bd = Backdrop {
            anchor: Some(Point3d {
                x: 100000,
                y: 200000,
                z: 300000,
            }),
            floor: true,
            wall: true,
            left: true,
            right: true,
            top: true,
            bottom: true,
        };
        let mut w = XmlWriter::new();
        bd.write_xml(&mut w);
        let s = &w.buf;
        // 锚点
        assert!(
            s.contains("<a:anchor x=\"100000\" y=\"200000\" z=\"300000\"/>"),
            "应包含 anchor，实际: {s}"
        );
        // 6 个平面（按 OOXML 顺序）
        assert!(s.contains("<a:floor/>"), "应包含 floor: {s}");
        assert!(s.contains("<a:wall/>"), "应包含 wall: {s}");
        assert!(s.contains("<a:l/>"), "应包含 l: {s}");
        assert!(s.contains("<a:r/>"), "应包含 r: {s}");
        assert!(s.contains("<a:t/>"), "应包含 t: {s}");
        assert!(s.contains("<a:b/>"), "应包含 b: {s}");
        // 验证顺序：anchor → floor → wall → l → r → t → b
        // 注意：用 <a:xxx/> 精确匹配，避免 <a:b 误匹配 <a:backdrop>
        let anchor_pos = s.find("<a:anchor").expect("anchor 应存在");
        let floor_pos = s.find("<a:floor/>").expect("floor 应存在");
        let wall_pos = s.find("<a:wall/>").expect("wall 应存在");
        let l_pos = s.find("<a:l/>").expect("l 应存在");
        let r_pos = s.find("<a:r/>").expect("r 应存在");
        let t_pos = s.find("<a:t/>").expect("t 应存在");
        let b_pos = s.find("<a:b/>").expect("b 应存在");
        assert!(anchor_pos < floor_pos, "anchor 应在 floor 之前");
        assert!(floor_pos < wall_pos, "floor 应在 wall 之前");
        assert!(wall_pos < l_pos, "wall 应在 l 之前");
        assert!(l_pos < r_pos, "l 应在 r 之前");
        assert!(r_pos < t_pos, "r 应在 t 之前");
        assert!(t_pos < b_pos, "t 应在 b 之前");
    }

    /// 验证 `Backdrop` 仅启用部分平面（floor + wall）。
    #[test]
    fn backdrop_partial_planes_serialize() {
        let bd = Backdrop {
            floor: true,
            wall: true,
            ..Default::default()
        };
        let mut w = XmlWriter::new();
        bd.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:floor/>"), "应包含 floor: {s}");
        assert!(s.contains("<a:wall/>"), "应包含 wall: {s}");
        // 注意：用 <a:xxx/> 精确匹配，避免 <a:b 误匹配 <a:backdrop>
        assert!(!s.contains("<a:l/>"), "不应包含 l: {s}");
        assert!(!s.contains("<a:r/>"), "不应包含 r: {s}");
        assert!(!s.contains("<a:t/>"), "不应包含 t: {s}");
        assert!(!s.contains("<a:b/>"), "不应包含 b: {s}");
    }

    /// 验证 `Scene3d` 带 backdrop 的完整序列化。
    #[test]
    fn scene3d_with_backdrop_serialize() {
        let scene = Scene3d {
            camera: Camera::default(),
            light_rig: LightRig::default(),
            backdrop: Some(Backdrop {
                floor: true,
                wall: true,
                ..Default::default()
            }),
        };
        let mut w = XmlWriter::new();
        scene.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:scene3d>"), "应包含 scene3d: {s}");
        assert!(s.contains("<a:backdrop>"), "应包含 backdrop: {s}");
        assert!(s.contains("<a:floor/>"), "应包含 floor: {s}");
        assert!(s.contains("<a:wall/>"), "应包含 wall: {s}");
        // 验证顺序：camera → lightRig → backdrop
        let cam_pos = s.find("<a:camera").expect("camera 应存在");
        let rig_pos = s.find("<a:lightRig").expect("lightRig 应存在");
        let bd_pos = s.find("<a:backdrop>").expect("backdrop 应存在");
        assert!(cam_pos < rig_pos, "camera 应在 lightRig 之前");
        assert!(rig_pos < bd_pos, "lightRig 应在 backdrop 之前");
    }

    /// 验证 `Scene3d` 无 backdrop 时不输出 backdrop 元素（向后兼容）。
    #[test]
    fn scene3d_without_backdrop_no_element() {
        let scene = Scene3d::default();
        let mut w = XmlWriter::new();
        scene.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("<a:scene3d>"), "应包含 scene3d: {s}");
        assert!(!s.contains("<a:backdrop>"), "无 backdrop 时不应输出: {s}");
    }

    /// 验证 `Point3d` 序列化坐标值。
    #[test]
    fn point3d_serialize() {
        let p = Point3d {
            x: -100000,
            y: 0,
            z: 500000,
        };
        let mut w = XmlWriter::new();
        p.write_xml(&mut w);
        let s = &w.buf;
        assert!(s.contains("x=\"-100000\""), "应包含 x: {s}");
        assert!(s.contains("y=\"0\""), "应包含 y: {s}");
        assert!(s.contains("z=\"500000\""), "应包含 z: {s}");
    }
}
