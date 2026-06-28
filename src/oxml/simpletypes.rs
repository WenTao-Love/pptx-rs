//! 简单 OOXML 类型。
//!
//! 包括文本对齐、字体、颜色等"枚举"型属性。
//!
//! 本模块与 python-pptx 中 `pptx.enum.text`、`pptx.enum.shapes` 等
//! 子包对标，但本库倾向于把"枚举"集中在 `simpletypes` 一处，便于查找。
//!
//! # 设计要点
//!
//! - **全部派生 `Copy + Eq + Hash`**：便于 `match` / `BTreeMap` 索引。
//! - **`as_str()` 返回 `&'static str`**：零分配序列化。
//! - **`FromStr` 实现**：解析时把 OOXML 字面量转回枚举。
//! - **未知值兜底**：用 [`PresetGeometry::Other`] 等变体吸收未识别值，
//!   避免解析阶段直接 panic；与 python-pptx 的容错策略一致。

use std::str::FromStr;

use crate::oxml::color::SchemeColor;
use crate::oxml::sppr::Dash;

/// 段落水平对齐方式（`algn` 属性）。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum Alignment {
    /// 左对齐。
    #[default]
    Left,
    /// 居中。
    Center,
    /// 右对齐。
    Right,
    /// 两端对齐。
    Justify,
    /// 分散对齐。
    Distribute,
}

impl Alignment {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            Alignment::Left => "l",
            Alignment::Center => "ctr",
            Alignment::Right => "r",
            Alignment::Justify => "just",
            Alignment::Distribute => "dist",
        }
    }
}

impl FromStr for Alignment {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "l" => Alignment::Left,
            "ctr" => Alignment::Center,
            "r" => Alignment::Right,
            "just" => Alignment::Justify,
            "dist" => Alignment::Distribute,
            _ => return Err(()),
        })
    }
}

/// 几何预设形状名（`prstGeom prst="..."`）。
///
/// 仅包含 python-pptx 文档过的常用子集，足以应付大部分 case。
/// 全部 200+ 形状见 ECMA-376 Part 1, §19.3 起的 `ST_PresetShape`。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum PresetGeometry {
    /// 矩形。
    Rectangle,
    /// 圆角矩形。
    RoundRectangle,
    /// 椭圆。
    Ellipse,
    /// 直线。
    Line,
    /// 直角三角形。
    RightTriangle,
    /// 菱形。
    Diamond,
    /// 平行四边形。
    Parallelogram,
    /// 梯形。
    Trapezoid,
    /// 六边形。
    Hexagon,
    /// 八边形。
    Octagon,
    /// 五角星。
    Star5,
    /// 六角星。
    Star6,
    /// 十角星。
    Star10,
    /// 普通箭头。
    Arrow,
    /// 右箭头。
    RightArrow,
    /// 左箭头。
    LeftArrow,
    /// 上箭头。
    UpArrow,
    /// 下箭头。
    DownArrow,
    /// 弯箭头。
    BentArrow,
    /// 弧形右箭头。
    CurvedRightArrow,
    /// 矩形标注。
    Callout1,
    /// 圆角矩形标注。
    Callout2,
    /// 云形。
    Cloud,
    /// 心形。
    Heart,
    /// 闪电。
    LightningBolt,
    /// 太阳。
    Sun,
    /// 月亮。
    Moon,
    /// 笑脸。
    SmileyFace,
    /// 甜甜圈。
    Donut,
    /// 立方体。
    Cube,
    /// 圆柱。
    Can,
    /// 斜面。
    Bevel,
    /// 禁止符号。
    NoSmoking,
    /// 弧块。
    BlockArc,
    /// 加号。
    Plus,
    /// 减号。
    Minus,
    /// 牌匾。
    Plaque,
    /// 波浪。
    Wave,
    /// 饼图。
    Pie,
    /// 斜面 2。
    Bevel2,
    /// 折角。
    FoldedCorner,
    /// 五边形。
    Pentagon,
    /// 人字形。
    Chevron,
    /// 右方括号。
    RightBracket,
    /// 左方括号。
    LeftBracket,
    /// 双方括号。
    DoubleBracket,
    /// 直线连接器。
    StraightConnector1,
    /// 折线连接器（2 段，无调整点）。
    BentConnector2,
    /// 折线连接器（3 段，1 个调整点）。
    BentConnector3,
    /// 折线连接器（4 段，2 个调整点）。
    BentConnector4,
    /// 折线连接器（5 段，3 个调整点）。
    BentConnector5,
    /// 曲线连接器（2 段，无调整点）。
    CurvedConnector2,
    /// 曲线连接器（3 段，1 个调整点）。
    CurvedConnector3,
    /// 曲线连接器（4 段，2 个调整点）。
    CurvedConnector4,
    /// 曲线连接器（5 段，3 个调整点）。
    CurvedConnector5,
    /// 任意未识别形状的兜底（实际写出 `prst="rect"`）。
    Other,
}

impl PresetGeometry {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        use PresetGeometry::*;
        match self {
            Rectangle => "rect",
            RoundRectangle => "roundRect",
            Ellipse => "ellipse",
            Line => "line",
            RightTriangle => "rtTriangle",
            Diamond => "diamond",
            Parallelogram => "parallelogram",
            Trapezoid => "trapezoid",
            Hexagon => "hexagon",
            Octagon => "octagon",
            Star5 => "star5",
            Star6 => "star6",
            Star10 => "star10",
            Arrow => "arrow",
            RightArrow => "rightArrow",
            LeftArrow => "leftArrow",
            UpArrow => "upArrow",
            DownArrow => "downArrow",
            BentArrow => "bentArrow",
            CurvedRightArrow => "curvedRightArrow",
            Callout1 => "wedgeRectCallout",
            Callout2 => "wedgeRoundRectCallout",
            Cloud => "cloud",
            Heart => "heart",
            LightningBolt => "lightningBolt",
            Sun => "sun",
            Moon => "moon",
            SmileyFace => "smileyFace",
            Donut => "donut",
            Cube => "cube",
            Can => "can",
            Bevel => "bevel",
            NoSmoking => "noSmoking",
            BlockArc => "blockArc",
            Plus => "mathPlus",
            Minus => "mathMinus",
            Plaque => "plaque",
            Wave => "wave",
            Pie => "pie",
            Bevel2 => "bevel2",
            FoldedCorner => "foldedCorner",
            Pentagon => "homePlate",
            Chevron => "chevron",
            RightBracket => "rightBracket",
            LeftBracket => "leftBracket",
            DoubleBracket => "doubleBracket",
            StraightConnector1 => "straightConnector1",
            BentConnector2 => "bentConnector2",
            BentConnector3 => "bentConnector3",
            BentConnector4 => "bentConnector4",
            BentConnector5 => "bentConnector5",
            CurvedConnector2 => "curvedConnector2",
            CurvedConnector3 => "curvedConnector3",
            CurvedConnector4 => "curvedConnector4",
            CurvedConnector5 => "curvedConnector5",
            Other => "rect",
        }
    }
}

impl FromStr for PresetGeometry {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        use PresetGeometry::*;
        Ok(match s {
            "rect" => Rectangle,
            "roundRect" => RoundRectangle,
            "ellipse" => Ellipse,
            "line" => Line,
            "rtTriangle" => RightTriangle,
            "diamond" => Diamond,
            "parallelogram" => Parallelogram,
            "trapezoid" => Trapezoid,
            "hexagon" => Hexagon,
            "octagon" => Octagon,
            "star5" => Star5,
            "star6" => Star6,
            "star10" => Star10,
            "arrow" => Arrow,
            "rightArrow" => RightArrow,
            "leftArrow" => LeftArrow,
            "upArrow" => UpArrow,
            "downArrow" => DownArrow,
            "bentArrow" => BentArrow,
            "curvedRightArrow" => CurvedRightArrow,
            "wedgeRectCallout" => Callout1,
            "wedgeRoundRectCallout" => Callout2,
            "cloud" => Cloud,
            "heart" => Heart,
            "lightningBolt" => LightningBolt,
            "sun" => Sun,
            "moon" => Moon,
            "smileyFace" => SmileyFace,
            "donut" => Donut,
            "cube" => Cube,
            "can" => Can,
            "bevel" => Bevel,
            "noSmoking" => NoSmoking,
            "blockArc" => BlockArc,
            "mathPlus" => Plus,
            "mathMinus" => Minus,
            "plaque" => Plaque,
            "wave" => Wave,
            "pie" => Pie,
            "bevel2" => Bevel2,
            "foldedCorner" => FoldedCorner,
            "homePlate" => Pentagon,
            "chevron" => Chevron,
            "rightBracket" => RightBracket,
            "leftBracket" => LeftBracket,
            "doubleBracket" => DoubleBracket,
            "straightConnector1" => StraightConnector1,
            "bentConnector2" => BentConnector2,
            "bentConnector3" => BentConnector3,
            "bentConnector4" => BentConnector4,
            "bentConnector5" => BentConnector5,
            "curvedConnector2" => CurvedConnector2,
            "curvedConnector3" => CurvedConnector3,
            "curvedConnector4" => CurvedConnector4,
            "curvedConnector5" => CurvedConnector5,
            _ => Other,
        })
    }
}

/// 文本下划线样式（`a:rPr u="..."`）。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum Underline {
    /// 无下划线。
    None,
    /// 单线下划线。
    Single,
    /// 双线下划线。
    Double,
    /// 加粗下划线。
    Heavy,
    /// 点状下划线。
    Dotted,
    /// 虚线下划线。
    Dashed,
    /// 波浪下划线。
    Wavy,
}

impl Underline {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            Underline::None => "none",
            Underline::Single => "sng",
            Underline::Double => "dbl",
            Underline::Heavy => "heavy",
            Underline::Dotted => "dotted",
            Underline::Dashed => "dash",
            Underline::Wavy => "wavy",
        }
    }
}

impl FromStr for Underline {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "none" => Underline::None,
            "sng" => Underline::Single,
            "dbl" => Underline::Double,
            "heavy" => Underline::Heavy,
            "dotted" => Underline::Dotted,
            "dash" => Underline::Dashed,
            "wavy" => Underline::Wavy,
            _ => return Err(()),
        })
    }
}

/// 端帽样式（`a:ln cap="..."`）。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum Cap {
    /// 平头（`flat`）。
    Flat,
    /// 小头（`small`）。
    Small,
    /// 大头（`all`）。
    All,
}

impl Cap {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            Cap::Flat => "flat",
            Cap::Small => "small",
            Cap::All => "all",
        }
    }
}

/// 段落属性方向（竖排/横排，`a:bodyPr vert="..."`）。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum TextDirection {
    /// 水平。
    Horizontal,
    /// 竖排。
    Vertical,
    /// 东亚竖排（旋转 270°）。
    VerticalEastAsia,
}

impl TextDirection {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            TextDirection::Horizontal => "horz",
            TextDirection::Vertical => "vert",
            TextDirection::VerticalEastAsia => "vert270",
        }
    }
}

/// 文本框自动调整策略（`MSO_AUTO_SIZE`）。
///
/// 对标 python-pptx 中 `pptx.enum.text.MSO_AUTO_SIZE`。
/// 序列化时分别落到 `<a:bodyPr>` 内的 `<a:normAutofit>` / `<a:spAutoFit>` 子元素。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum MsoAutoSize {
    /// 不自动调整（不写 autofit 子元素）。
    #[default]
    None,
    /// 文字溢出时调整形状（`spAutoFit`）。
    ShapeToFitText,
    /// 文字溢出时缩小字号（`normAutofit`）。
    TextToFitShape,
}

impl MsoAutoSize {
    /// 转 OOXML 字面量（`a:bodyPr` 直接属性或子元素标签）。
    ///
    /// - `None` → 无输出；
    /// - `ShapeToFitText` → `<a:spAutoFit/>`；
    /// - `TextToFitShape` → `<a:normAutofit/>`。
    pub fn tag_name(self) -> Option<&'static str> {
        match self {
            MsoAutoSize::None => None,
            MsoAutoSize::ShapeToFitText => Some("a:spAutoFit"),
            MsoAutoSize::TextToFitShape => Some("a:normAutofit"),
        }
    }
}

/// 文本框垂直对齐（`MSO_ANCHOR`）。
///
/// 对标 python-pptx 中 `pptx.enum.text.MSO_ANCHOR`。
/// 序列化时落到 `<a:bodyPr anchor="...">`。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum MsoAnchor {
    /// 顶端对齐。
    #[default]
    Top,
    /// 居中对齐。
    Middle,
    /// 底端对齐。
    Bottom,
}

impl MsoAnchor {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoAnchor::Top => "t",
            MsoAnchor::Middle => "ctr",
            MsoAnchor::Bottom => "b",
        }
    }
}

impl FromStr for MsoAnchor {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "t" => MsoAnchor::Top,
            "ctr" => MsoAnchor::Middle,
            "b" => MsoAnchor::Bottom,
            _ => return Err(()),
        })
    }
}

/// 连接器类型（`MSO_CONNECTOR_TYPE`）。
///
/// 对标 python-pptx 中 `pptx.enum.shapes.MSO_CONNECTOR_TYPE`。
/// 序列化时落到 `<a:prstGeom prst="...">` 的 `prst` 属性。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoConnectorType {
    /// 直线（`straightConnector1`）。
    Straight,
    /// 单折线（`bentConnector2`）。
    Elbow,
    /// 曲线（`curvedConnector2`）。
    Curve,
    /// 弯曲连接器 3 段（`bentConnector3`，带 1 个调整点）。
    BentConnector3,
    /// 弯曲连接器 4 段（`bentConnector4`，带 2 个调整点）。
    BentConnector4,
    /// 弯曲连接器 5 段（`bentConnector5`，带 3 个调整点）。
    BentConnector5,
    /// 曲线连接器 3 段（`curvedConnector3`）。
    CurvedConnector3,
    /// 曲线连接器 4 段（`curvedConnector4`）。
    CurvedConnector4,
    /// 曲线连接器 5 段（`curvedConnector5`）。
    CurvedConnector5,
}

impl MsoConnectorType {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoConnectorType::Straight => "straightConnector1",
            MsoConnectorType::Elbow => "bentConnector2",
            MsoConnectorType::Curve => "curvedConnector2",
            MsoConnectorType::BentConnector3 => "bentConnector3",
            MsoConnectorType::BentConnector4 => "bentConnector4",
            MsoConnectorType::BentConnector5 => "bentConnector5",
            MsoConnectorType::CurvedConnector3 => "curvedConnector3",
            MsoConnectorType::CurvedConnector4 => "curvedConnector4",
            MsoConnectorType::CurvedConnector5 => "curvedConnector5",
        }
    }
}

impl FromStr for MsoConnectorType {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "straightConnector1" => MsoConnectorType::Straight,
            "bentConnector2" => MsoConnectorType::Elbow,
            "curvedConnector2" => MsoConnectorType::Curve,
            "bentConnector3" => MsoConnectorType::BentConnector3,
            "bentConnector4" => MsoConnectorType::BentConnector4,
            "bentConnector5" => MsoConnectorType::BentConnector5,
            "curvedConnector3" => MsoConnectorType::CurvedConnector3,
            "curvedConnector4" => MsoConnectorType::CurvedConnector4,
            "curvedConnector5" => MsoConnectorType::CurvedConnector5,
            _ => return Err(()),
        })
    }
}

/// 形状类型（`MSO_SHAPE_TYPE`）。
///
/// 对标 python-pptx 中 `pptx.enum.shapes.MSO_SHAPE_TYPE`。
/// 本枚举用于 `Shape::shape_type` 的返回值。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoShapeType {
    /// 自由形状（`AUTO_SHAPE` / `FREE_FORM`）。
    AutoShape,
    /// 文本框。
    TextBox,
    /// 占位符。
    Placeholder,
    /// 图片。
    Picture,
    /// 组合。
    Group,
    /// 连接器。
    Connector,
    /// 表格（GraphicFrame）。
    GraphicFrame,
}

impl MsoShapeType {
    /// 转小写字符串（与 [`crate::shape::ShapeKind::shape_type`] 对齐）。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoShapeType::AutoShape => "auto_shape",
            MsoShapeType::TextBox => "text_box",
            MsoShapeType::Placeholder => "placeholder",
            MsoShapeType::Picture => "picture",
            MsoShapeType::Group => "group",
            MsoShapeType::Connector => "connector",
            MsoShapeType::GraphicFrame => "table",
        }
    }
}

/// 自动换行模式（`<a:bodyPr wrap="...">`）。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum TextWrapping {
    /// 不自动换行（`wrap="none"`）。
    None,
    /// 方形换行（`wrap="square"`）——OOXML 默认行为。
    #[default]
    Square,
}

impl TextWrapping {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            TextWrapping::None => "none",
            TextWrapping::Square => "square",
        }
    }
}

/// 制表位对齐类型（`<a:tab algn="...">`）。
///
/// 对标 python-pptx `pptx.enum.text.PP_ALIGN` 在制表位场景下的子集。
/// OOXML 中 `<a:tab>` 元素的 `algn` 属性取值：`l` / `ctr` / `r` / `dec`。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum TabAlignment {
    /// 左对齐制表位（`algn="l"`）——默认。
    #[default]
    Left,
    /// 居中对齐制表位（`algn="ctr"`）。
    Center,
    /// 右对齐制表位（`algn="r"`）。
    Right,
    /// 小数点对齐制表位（`algn="dec"`），常用于数字列表。
    Decimal,
}

impl TabAlignment {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            TabAlignment::Left => "l",
            TabAlignment::Center => "ctr",
            TabAlignment::Right => "r",
            TabAlignment::Decimal => "dec",
        }
    }
}

impl FromStr for TabAlignment {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "l" => TabAlignment::Left,
            "ctr" => TabAlignment::Center,
            "r" => TabAlignment::Right,
            "dec" => TabAlignment::Decimal,
            _ => return Err(()),
        })
    }
}

// ====================================================================
// 以下枚举对标 python-pptx `pptx.enum.*`，是 0.1.x 末段补全：
//   - MsoFillType / MsoThemeColorIndex / MsoColorType / MsoLineDashStyle
//   - PpAlign / PpPlaceholderType / MsoShapeType 完整版
// 主要为"python-pptx 风格 API"提供枚举来源；序列化时仍走 [`as_str`]
// 返回的字面量。
// ====================================================================

/// 填充类型（`MSO_FILL_TYPE`）。
///
/// 对标 python-pptx 中 `pptx.enum.dml.MSO_FILL_TYPE`。
/// 主要用于 `FillFormat.type` 的只读视图。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoFillType {
    /// 背景自动。
    Background,
    /// 渐变。
    Gradient,
    /// 图案。
    Pattern,
    /// 图片。
    Picture,
    /// 实色。
    Solid,
    /// 继承。
    Inherit,
    /// 自动（由 Office 决定）。
    Mixed,
}

impl MsoFillType {
    /// 转 OOXML 字面量（仅用于调试 / 诊断）。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoFillType::Background => "background",
            MsoFillType::Gradient => "gradient",
            MsoFillType::Pattern => "pattern",
            MsoFillType::Picture => "picture",
            MsoFillType::Solid => "solid",
            MsoFillType::Inherit => "inherit",
            MsoFillType::Mixed => "mixed",
        }
    }
}

/// 主题色索引（`MSO_THEME_COLOR_INDEX`）。
///
/// 对标 python-pptx 中 `pptx.enum.dml.MSO_THEME_COLOR_INDEX`。
/// 与 [`SchemeColor`] 区别：本枚举是"运行时颜色索引"，`SchemeColor`
/// 是"OOXML 字面量"——两者一一对应。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoThemeColorIndex {
    /// 背景 1（Not scheme color）。
    NotThemeColor,
    /// 背景 1。
    Background1,
    /// 背景 2。
    Background2,
    /// 文本 1。
    Text1,
    /// 文本 2。
    Text2,
    /// 强调色 1。
    Accent1,
    /// 强调色 2。
    Accent2,
    /// 强调色 3。
    Accent3,
    /// 强调色 4。
    Accent4,
    /// 强调色 5。
    Accent5,
    /// 强调色 6。
    Accent6,
    /// 超链接。
    Hyperlink,
    /// 已访问超链接。
    FollowedHyperlink,
    /// 浅色 1。
    Light1,
    /// 浅色 2。
    Light2,
    /// 深色 1。
    Dark1,
    /// 深色 2。
    Dark2,
    /// 混合。
    Mixed,
}

impl MsoThemeColorIndex {
    /// 转 OOXML 字面量（与 [`SchemeColor::as_str`] 一致）。
    pub fn as_str(self) -> Option<&'static str> {
        match self {
            MsoThemeColorIndex::NotThemeColor => None,
            MsoThemeColorIndex::Background1 => Some("bg1"),
            MsoThemeColorIndex::Background2 => Some("bg2"),
            MsoThemeColorIndex::Text1 => Some("tx1"),
            MsoThemeColorIndex::Text2 => Some("tx2"),
            MsoThemeColorIndex::Accent1 => Some("accent1"),
            MsoThemeColorIndex::Accent2 => Some("accent2"),
            MsoThemeColorIndex::Accent3 => Some("accent3"),
            MsoThemeColorIndex::Accent4 => Some("accent4"),
            MsoThemeColorIndex::Accent5 => Some("accent5"),
            MsoThemeColorIndex::Accent6 => Some("accent6"),
            MsoThemeColorIndex::Hyperlink => Some("hlink"),
            MsoThemeColorIndex::FollowedHyperlink => Some("folHlink"),
            MsoThemeColorIndex::Light1 => Some("lt1"),
            MsoThemeColorIndex::Light2 => Some("lt2"),
            MsoThemeColorIndex::Dark1 => Some("dk1"),
            MsoThemeColorIndex::Dark2 => Some("dk2"),
            MsoThemeColorIndex::Mixed => None,
        }
    }

    /// 从 [`SchemeColor`] 映射回来（与 `as_str` 互逆）。
    pub fn from_scheme(c: SchemeColor) -> Self {
        match c {
            SchemeColor::Background1 => MsoThemeColorIndex::Background1,
            SchemeColor::Background2 => MsoThemeColorIndex::Background2,
            SchemeColor::Text1 => MsoThemeColorIndex::Text1,
            SchemeColor::Text2 => MsoThemeColorIndex::Text2,
            SchemeColor::Accent1 => MsoThemeColorIndex::Accent1,
            SchemeColor::Accent2 => MsoThemeColorIndex::Accent2,
            SchemeColor::Accent3 => MsoThemeColorIndex::Accent3,
            SchemeColor::Accent4 => MsoThemeColorIndex::Accent4,
            SchemeColor::Accent5 => MsoThemeColorIndex::Accent5,
            SchemeColor::Accent6 => MsoThemeColorIndex::Accent6,
            SchemeColor::Hlink => MsoThemeColorIndex::Hyperlink,
            SchemeColor::FolHlink => MsoThemeColorIndex::FollowedHyperlink,
            SchemeColor::Lt1 => MsoThemeColorIndex::Light1,
            SchemeColor::Lt2 => MsoThemeColorIndex::Light2,
            SchemeColor::Dk1 => MsoThemeColorIndex::Dark1,
            SchemeColor::Dk2 => MsoThemeColorIndex::Dark2,
        }
    }
}

/// 颜色类型（`MSO_COLOR_TYPE`）。
///
/// 对标 python-pptx 中 `pptx.enum.dml.MSO_COLOR_TYPE`。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoColorType {
    /// 系统窗口主题色。
    SystemColor,
    /// sRGB 颜色。
    Rgb,
    /// 预设颜色。
    PresetColor,
    /// 主题色（schemeClr）。
    SchemeColor,
    /// 混合。
    Mixed,
}

impl MsoColorType {
    /// 转字符串。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoColorType::SystemColor => "system",
            MsoColorType::Rgb => "rgb",
            MsoColorType::PresetColor => "preset",
            MsoColorType::SchemeColor => "scheme",
            MsoColorType::Mixed => "mixed",
        }
    }
}

/// 段落水平对齐（`PP_ALIGN`）——python-pptx 风格别名。
///
/// 对标 python-pptx 中 `pptx.enum.text.PP_ALIGN`。本枚举与 [`Alignment`]
/// 1:1 对应，但**仅**作为面向用户的"python-pptx 同名"别名。
pub type PpAlign = Alignment;

/// 占位符类型（`PP_PLACEHOLDER`）。
///
/// 对标 python-pptx 中 `pptx.enum.shapes.PP_PLACEHOLDER`。
/// 注意本枚举只覆盖**常见**的占位符类型；其它取值统一收为 [`PpPlaceholderType::Other`]。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum PpPlaceholderType {
    /// 标题。
    Title,
    /// 正文。
    Body,
    /// 中心文本。
    CenterTitle,
    /// 子标题。
    Subtitle,
    /// 日期。
    Date,
    /// 幻灯片编号。
    SlideNumber,
    /// 页脚。
    Footer,
    /// 页眉。
    Header,
    /// 对象（表格/图/...）。
    Object,
    /// 图表。
    Chart,
    /// 表格。
    Table,
    /// 剪贴画。
    ClipArt,
    /// 组织结构图。
    OrgChart,
    /// 媒体剪辑。
    MediaClip,
    /// 图片占位符（`pic`，TODO-007）。
    ///
    /// 对应 PowerPoint "图片占位符"——版式中的 `<p:ph type="pic"/>`，
    /// 用户点击后弹出"插入图片"对话框。
    Picture,
    /// 其它（任意未识别的 `type` 字符串）。
    Other,
}

impl PpPlaceholderType {
    /// 转 OOXML `<p:ph type="...">` 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            PpPlaceholderType::Title => "title",
            PpPlaceholderType::Body => "body",
            PpPlaceholderType::CenterTitle => "ctrTitle",
            PpPlaceholderType::Subtitle => "subTitle",
            PpPlaceholderType::Date => "dt",
            PpPlaceholderType::SlideNumber => "sldNum",
            PpPlaceholderType::Footer => "ftr",
            PpPlaceholderType::Header => "hdr",
            PpPlaceholderType::Object => "obj",
            PpPlaceholderType::Chart => "chart",
            PpPlaceholderType::Table => "tbl",
            PpPlaceholderType::ClipArt => "clipArt",
            PpPlaceholderType::OrgChart => "orgChart",
            PpPlaceholderType::MediaClip => "media",
            PpPlaceholderType::Picture => "pic",
            // Other 回落到 body——这是 OOXML 规范的默认 placeholder 类型，
            // 也是 PowerPoint 在用户新增占位符时实际写入的值。
            PpPlaceholderType::Other => "body",
        }
    }

    /// 从 OOXML `<p:ph type="...">` 字面量解析（TODO-007）。
    ///
    /// 与 [`as_str`](Self::as_str) 互逆；未识别的字符串回落到 [`Other`](Self::Other)。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Self {
        match s {
            "title" => PpPlaceholderType::Title,
            "body" => PpPlaceholderType::Body,
            "ctrTitle" => PpPlaceholderType::CenterTitle,
            "subTitle" => PpPlaceholderType::Subtitle,
            "dt" => PpPlaceholderType::Date,
            "sldNum" => PpPlaceholderType::SlideNumber,
            "ftr" => PpPlaceholderType::Footer,
            "hdr" => PpPlaceholderType::Header,
            "obj" => PpPlaceholderType::Object,
            "chart" => PpPlaceholderType::Chart,
            "tbl" => PpPlaceholderType::Table,
            "clipArt" => PpPlaceholderType::ClipArt,
            "orgChart" => PpPlaceholderType::OrgChart,
            "media" => PpPlaceholderType::MediaClip,
            "pic" => PpPlaceholderType::Picture,
            _ => PpPlaceholderType::Other,
        }
    }
}

/// 线型虚实（`MSO_LINE_DASH_STYLE`）。
///
/// 对标 python-pptx 中 `pptx.enum.dml.MSO_LINE_DASH_STYLE`。
/// 与 [`Dash`] 关系：`Dash` 是"已实现序列化"的精简子集，
/// `MsoLineDashStyle` 是"完整 ECMA-376"枚举——两者一一对应。
#[derive(Copy, Clone, Debug, Eq, PartialEq)]
pub enum MsoLineDashStyle {
    Solid,
    Dash,
    DashDot,
    DashDotDot,
    Dot,
    LongDash,
    LongDashDot,
    LongDashDotDot,
    RoundDot,
    SysDash,
    SysDashDot,
    SysDashDotDot,
    SysDot,
    Mixed,
}

impl MsoLineDashStyle {
    /// 转 OOXML `<a:prstDash val="...">` 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            MsoLineDashStyle::Solid => "solid",
            MsoLineDashStyle::Dash => "dash",
            MsoLineDashStyle::DashDot => "dashDot",
            MsoLineDashStyle::DashDotDot => "dashDotDot",
            MsoLineDashStyle::Dot => "dot",
            MsoLineDashStyle::LongDash => "lgDash",
            MsoLineDashStyle::LongDashDot => "lgDashDot",
            MsoLineDashStyle::LongDashDotDot => "lgDashDotDot",
            MsoLineDashStyle::RoundDot => "sysDot",
            MsoLineDashStyle::SysDash => "sysDash",
            MsoLineDashStyle::SysDashDot => "sysDashDot",
            MsoLineDashStyle::SysDashDotDot => "sysDashDotDot",
            MsoLineDashStyle::SysDot => "sysDot",
            MsoLineDashStyle::Mixed => "solid",
        }
    }
}

impl From<MsoLineDashStyle> for Dash {
    fn from(s: MsoLineDashStyle) -> Self {
        match s {
            MsoLineDashStyle::Solid => Dash::Solid,
            MsoLineDashStyle::Dash => Dash::Dash,
            MsoLineDashStyle::DashDot => Dash::DashDot,
            MsoLineDashStyle::DashDotDot => Dash::LgDashDotDot,
            MsoLineDashStyle::Dot => Dash::Dot,
            MsoLineDashStyle::LongDash => Dash::LgDash,
            MsoLineDashStyle::LongDashDot => Dash::LgDashDot,
            MsoLineDashStyle::LongDashDotDot => Dash::LgDashDotDot,
            MsoLineDashStyle::RoundDot => Dash::SysDot,
            MsoLineDashStyle::SysDash => Dash::SysDash,
            MsoLineDashStyle::SysDashDot => Dash::SysDashDot,
            MsoLineDashStyle::SysDashDotDot => Dash::SysDashDotDot,
            MsoLineDashStyle::SysDot => Dash::SysDot,
            MsoLineDashStyle::Mixed => Dash::Solid,
        }
    }
}
