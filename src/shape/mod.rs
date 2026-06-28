//! 高阶形状：AutoShape / Picture / Group / Connector / Table / TextBox / Freeform。
//!
//! 对应 python-pptx 中 `pptx.shapes.*` 各类。
//!
//! 设计：所有形状都**借用**所属 Slide 上的 `spTree` 元素。本 crate 在第一版
//! 采用了"对 oxml 类型做轻包装 + 直接修改"的策略——即一个 `AutoShape` 实质
//! 上持有一个 `Rc<RefCell<Sp>>` 句柄；高阶方法直接 mutate 内部 oxml 模型。
//!
//! # 模块构成
//!
//! - [`base`]：所有形状共享的 trait [`Shape`]；
//! - [`autoshape`]：自选图形（矩形/椭圆/箭头/...）；
//! - [`textbox`]：纯文本框；
//! - [`picture`]：图片；
//! - [`group`]：组合 + 递归子形状；
//! - [`connector`]：连接器（直线/折线/曲线）；
//! - [`table`]：高阶表格（封装 `<a:tbl>`）；
//! - [`chartshape`]：高阶图表（封装 `<p:graphicFrame>` + `<c:chart>` 引用）；
//! - [`oleshape`]：高阶 OLE 对象（封装 `<p:graphicFrame>` + `<p:oleObj>` 引用）；
//! - [`freeform`]：手绘自由形（custGeom 极简版）。
//!
//! # 在三层架构中的位置
//!
//! 本模块是 `shape::*` —— `oxml::shape` 的"高阶薄包装"。它不引入新数据，
//! 全部为 `pub(crate) sp: OxmlSp` 等字段的引用/可变访问。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.shapetree.SlideShapes` ←→ [`crate::slide::Shapes`] / [`crate::slide::ShapesMut`]；
//! - `pptx.shapes.autoshape.Shape` ←→ [`AutoShape`]；
//! - `pptx.shapes.textbox.TextBox` ←→ [`TextBox`]；
//! - `pptx.shapes.picture.Picture` ←→ [`Picture`]；
//! - `pptx.shapes.group.Group` ←→ [`Group`]；
//! - `pptx.shapes.connector.Connector` ←→ [`Connector`]；
//! - `pptx.table.Table` ←→ [`TableShape`]。

pub mod autoshape;
pub mod base;
pub mod chartshape;
pub mod connector;
pub mod freeform;
pub mod group;
pub mod oleshape;
pub mod picture;
pub mod smartartshape;
pub mod table;
pub mod textbox;

pub use autoshape::AutoShape;
pub use base::Shape;
pub use chartshape::ChartShape;
pub use connector::Connector;
pub use freeform::{Freeform, FreeformBuilder, Point};
pub use group::{Group, GroupChild};
pub use oleshape::OleObjectShape;
pub use picture::Picture;
pub use smartartshape::SmartArtShape;
pub use table::{BorderSide, TableShape};
pub use textbox::TextBox;

// 重新导出常用枚举，路径短一点。
pub use crate::oxml::simpletypes::{
    Alignment, Cap, MsoAnchor, MsoAutoSize, MsoConnectorType, MsoShapeType, PresetGeometry,
    TextDirection, TextWrapping, Underline,
};

use crate::oxml::SlideShape as OxmlSlideShape;

/// 统一的形状枚举（`Slide.shapes().get(i)` 拿到的是这个）。
///
/// 该枚举**仅**作为"返回类型"——一旦拿到具体子类型，调用方可以转回
/// [`AutoShape`] / [`Picture`] / ... 等高阶 API。
#[derive(Clone, Debug)]
pub enum ShapeKind {
    /// 文本框。
    TextBox(TextBox),
    /// 自选图形（矩形、椭圆、箭头、…）。
    AutoShape(AutoShape),
    /// 图片。
    Picture(Picture),
    /// 组合。
    Group(Group),
    /// 连接器。
    Connector(Connector),
    /// 表格。
    Table(TableShape),
    /// 图表（TODO-004）。
    Chart(ChartShape),
    /// OLE 对象（TODO-043）。
    OleObject(OleObjectShape),
    /// SmartArt 图形（TODO-037 创建 API）。
    SmartArt(SmartArtShape),
    /// 占位符（标题、正文…）。
    Placeholder(PlaceholderShape),
}

/// 占位符形状（暂时借用 AutoShape 表示 + 携带 ph 信息）。
///
/// 占位符与 [`AutoShape`] 的区别仅在于"是否带 `<p:ph>` 元素"，所以本类型
/// 直接包装 `AutoShape`；在序列化时 `AutoShape::write_xml`（实为 oxml `Sp::write_xml`）
/// 会自动按 `is_placeholder` 字段决定是否写 `<p:ph>`。
#[derive(Clone, Debug)]
pub struct PlaceholderShape(
    /// 内部包装的自选形状。
    pub AutoShape,
);

impl ShapeKind {
    /// 形状类型字符串（与 python-pptx 的 `shape_type` 对齐）。
    pub fn shape_type(&self) -> &'static str {
        match self {
            ShapeKind::TextBox(_) => "text_box",
            ShapeKind::AutoShape(_) => "auto_shape",
            ShapeKind::Picture(_) => "picture",
            ShapeKind::Group(_) => "group",
            ShapeKind::Connector(_) => "connector",
            ShapeKind::Table(_) => "table",
            ShapeKind::Chart(_) => "chart",
            ShapeKind::OleObject(_) => "ole_object",
            ShapeKind::SmartArt(_) => "smart_art",
            ShapeKind::Placeholder(_) => "placeholder",
        }
    }

    /// 形状名（与 python-pptx 的 `shape.name` 对齐）。
    pub fn name(&self) -> &str {
        match self {
            ShapeKind::TextBox(t) => t.name(),
            ShapeKind::AutoShape(s) => s.name(),
            ShapeKind::Picture(p) => p.name(),
            ShapeKind::Group(g) => g.name(),
            ShapeKind::Connector(c) => c.name(),
            ShapeKind::Table(t) => t.name(),
            ShapeKind::Chart(c) => c.name(),
            ShapeKind::OleObject(o) => o.name(),
            ShapeKind::SmartArt(s) => s.name(),
            ShapeKind::Placeholder(p) => p.0.name(),
        }
    }

    /// 形状 ID（与 python-pptx 的 `shape.shape_id` 对齐）。
    pub fn shape_id(&self) -> u32 {
        match self {
            ShapeKind::TextBox(t) => t.id(),
            ShapeKind::AutoShape(s) => s.id(),
            ShapeKind::Picture(p) => p.id(),
            ShapeKind::Group(g) => g.id(),
            ShapeKind::Connector(c) => c.id(),
            ShapeKind::Table(t) => t.id(),
            ShapeKind::Chart(c) => c.id(),
            ShapeKind::OleObject(o) => o.id(),
            ShapeKind::SmartArt(s) => s.id(),
            ShapeKind::Placeholder(p) => p.0.id(),
        }
    }

    /// 左边坐标（EMU）。
    pub fn left(&self) -> crate::units::Emu {
        match self {
            ShapeKind::TextBox(t) => t.left(),
            ShapeKind::AutoShape(s) => s.left(),
            ShapeKind::Picture(p) => p.left(),
            ShapeKind::Group(g) => g.left(),
            ShapeKind::Connector(c) => c.left(),
            ShapeKind::Table(t) => t.left(),
            ShapeKind::Chart(c) => c.left(),
            ShapeKind::OleObject(o) => o.left(),
            ShapeKind::SmartArt(s) => s.left(),
            ShapeKind::Placeholder(p) => p.0.left(),
        }
    }
    /// 顶边坐标（EMU）。
    pub fn top(&self) -> crate::units::Emu {
        match self {
            ShapeKind::TextBox(t) => t.top(),
            ShapeKind::AutoShape(s) => s.top(),
            ShapeKind::Picture(p) => p.top(),
            ShapeKind::Group(g) => g.top(),
            ShapeKind::Connector(c) => c.top(),
            ShapeKind::Table(t) => t.top(),
            ShapeKind::Chart(c) => c.top(),
            ShapeKind::OleObject(o) => o.top(),
            ShapeKind::SmartArt(s) => s.top(),
            ShapeKind::Placeholder(p) => p.0.top(),
        }
    }
    /// 宽度（EMU）。
    pub fn width(&self) -> crate::units::Emu {
        match self {
            ShapeKind::TextBox(t) => t.width(),
            ShapeKind::AutoShape(s) => s.width(),
            ShapeKind::Picture(p) => p.width(),
            ShapeKind::Group(g) => g.width(),
            ShapeKind::Connector(c) => c.width(),
            ShapeKind::Table(t) => t.width(),
            ShapeKind::Chart(c) => c.width(),
            ShapeKind::OleObject(o) => o.width(),
            ShapeKind::SmartArt(s) => s.width(),
            ShapeKind::Placeholder(p) => p.0.width(),
        }
    }
    /// 高度（EMU）。
    pub fn height(&self) -> crate::units::Emu {
        match self {
            ShapeKind::TextBox(t) => t.height(),
            ShapeKind::AutoShape(s) => s.height(),
            ShapeKind::Picture(p) => p.height(),
            ShapeKind::Group(g) => g.height(),
            ShapeKind::Connector(c) => c.height(),
            ShapeKind::Table(t) => t.height(),
            ShapeKind::Chart(c) => c.height(),
            ShapeKind::OleObject(o) => o.height(),
            ShapeKind::SmartArt(s) => s.height(),
            ShapeKind::Placeholder(p) => p.0.height(),
        }
    }
}

/// 从 oxml SlideShape 中**取名字**（不构造 [`ShapeKind`]）。
///
/// 在 `SlideShapes::index(shape)` 等"按名字定位"场景下使用，避免无谓 clone。
pub fn name_of(oxml: &OxmlSlideShape) -> &str {
    match oxml {
        OxmlSlideShape::Sp(sp) => &sp.name,
        OxmlSlideShape::Pic(p) => &p.name,
        OxmlSlideShape::CxnSp(c) => &c.name,
        OxmlSlideShape::Group(g) => &g.name,
        OxmlSlideShape::GraphicFrame(g) => &g.name,
    }
}

/// 把 oxml SlideShape 转换为高阶 [`ShapeKind`]。
///
/// 在 [`crate::slide::Shapes::get`] 中调用；属于"边界转换"，调用方不需要直接使用。
pub fn wrap(oxml: &OxmlSlideShape) -> ShapeKind {
    match oxml {
        OxmlSlideShape::Sp(sp) => {
            if sp.c_nv_sp_pr_tx_box {
                ShapeKind::TextBox(TextBox::from_sp(sp.clone()))
            } else {
                // 占位符与普通自选图形统一用 AutoShape 表示；
                // is_placeholder 仅影响序列化时是否写出 <p:ph> 元素。
                ShapeKind::AutoShape(AutoShape::from_sp(sp.clone()))
            }
        }
        OxmlSlideShape::Pic(p) => ShapeKind::Picture(Picture::from_pic(p.clone())),
        OxmlSlideShape::CxnSp(c) => ShapeKind::Connector(Connector::from_cxn(c.clone())),
        OxmlSlideShape::Group(g) => {
            // 简化: 子形状全部转成 AutoShape / Picture / Connector（递归 group）
            let ng = Group {
                group: (**g).clone(),
            };
            ShapeKind::Group(ng)
        }
        OxmlSlideShape::GraphicFrame(g) => match &g.graphic {
            crate::oxml::shape::Graphic::Table(_) => {
                ShapeKind::Table(TableShape::from_frame(g.clone()))
            }
            crate::oxml::shape::Graphic::Chart(_) => {
                ShapeKind::Chart(ChartShape::from_frame(g.clone()))
            }
            crate::oxml::shape::Graphic::OleObject(_) => {
                ShapeKind::OleObject(OleObjectShape::from_frame(g.clone()))
            }
            crate::oxml::shape::Graphic::SmartArt(_) => {
                ShapeKind::SmartArt(SmartArtShape::from_frame(g.clone()))
            }
        },
    }
}
