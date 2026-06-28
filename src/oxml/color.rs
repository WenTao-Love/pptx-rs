//! 颜色：`a:srgbClr` / `a:schemeClr` / `a:prstClr` 统一表达。
//!
//! OOXML 中颜色出现在多处（填充、文本、边框、效果、表格单元格背景等），
//! 但所有"颜色值"最终都落在三种元素上：
//!
//! - `<a:srgbClr val="RRGGBB"/>`：绝对 sRGB 颜色；
//! - `<a:schemeClr val="..."/>`：主题色（间接引用 `theme1.xml`）；
//! - `<a:prstClr val="..."/>`：147 个预设颜色之一。
//!
//! 本模块用 [`Color`] 枚举统一表达这三种 + "无颜色"。

use std::str::FromStr;

use crate::units::RGBColor;

/// 一个颜色（`a:solidFill` 内的颜色）。
///
/// `#[default]` 选 [`Color::None`] 是为了与"未设置"语义对齐。
#[derive(Clone, Debug, Default, PartialEq, Eq, Hash)]
pub enum Color {
    /// 主题颜色（引用主题/母版中的 schemeClr）。
    Scheme(SchemeColor),
    /// 预设颜色（`prstClr`）。
    Preset(PresetColor),
    /// sRGB 颜色。
    RGB(RGBColor),
    /// 暂未填充。
    #[default]
    None,
}

impl Color {
    /// 写一段 XML（写到指定 tag，例如 `a:srgbClr`）。
    ///
    /// `Color::None` 不写任何字节。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter, tag: &str) {
        match self {
            Color::None => {} // 不写
            Color::Scheme(s) => {
                w.open_with(tag, &[("val", s.as_str())]);
                // 可选 lumMod/lumOff/shade/tint 暂略
                w.close(tag);
            }
            Color::Preset(p) => {
                w.open_with(tag, &[("val", p.as_str())]);
                w.close(tag);
            }
            Color::RGB(c) => {
                w.open(tag);
                w.empty_with(
                    "a:srgbClr",
                    &[("val", &format!("{:02X}{:02X}{:02X}", c.0, c.1, c.2))],
                );
                w.close(tag);
            }
        }
    }

    /// 写为 `<a:solidFill><a:??>...</a:??></a:solidFill>`。
    ///
    /// `Color::None` 写为 `<a:noFill/>`（语义上"无填充"）。
    pub fn write_solid_fill(&self, w: &mut super::writer::XmlWriter) {
        self.write_solid_fill_with_alpha(w, None);
    }

    /// 写为 `<a:solidFill><a:??>...<a:alpha val="..."/></a:??></a:solidFill>`。
    ///
    /// 与 [`Self::write_solid_fill`] 相同，但额外在颜色元素内写入 `<a:alpha>` 子元素。
    ///
    /// # 参数
    /// - `alpha`：透明度值（0-100000），`None` 表示不写 alpha。
    ///   - 0 = 完全不透明
    ///   - 100000 = 完全透明
    ///   - 30000 = 30% 不透明（70% 透明），常用于水印
    pub fn write_solid_fill_with_alpha(
        &self,
        w: &mut super::writer::XmlWriter,
        alpha: Option<i32>,
    ) {
        match self {
            Color::None => {
                w.empty("a:noFill");
            }
            Color::RGB(c) => {
                w.open("a:solidFill");
                let val_s = format!("{:02X}{:02X}{:02X}", c.0, c.1, c.2);
                if alpha.is_some() {
                    w.open_with("a:srgbClr", &[("val", val_s.as_str())]);
                    if let Some(a) = alpha {
                        w.empty_with("a:alpha", &[("val", a.to_string().as_str())]);
                    }
                    w.close("a:srgbClr");
                } else {
                    w.empty_with("a:srgbClr", &[("val", val_s.as_str())]);
                }
                w.close("a:solidFill");
            }
            Color::Scheme(s) => {
                w.open("a:solidFill");
                if alpha.is_some() {
                    w.open_with("a:schemeClr", &[("val", s.as_str())]);
                    if let Some(a) = alpha {
                        w.empty_with("a:alpha", &[("val", a.to_string().as_str())]);
                    }
                    w.close("a:schemeClr");
                } else {
                    w.empty_with("a:schemeClr", &[("val", s.as_str())]);
                }
                w.close("a:solidFill");
            }
            Color::Preset(p) => {
                w.open("a:solidFill");
                if alpha.is_some() {
                    w.open_with("a:prstClr", &[("val", p.as_str())]);
                    if let Some(a) = alpha {
                        w.empty_with("a:alpha", &[("val", a.to_string().as_str())]);
                    }
                    w.close("a:prstClr");
                } else {
                    w.empty_with("a:prstClr", &[("val", p.as_str())]);
                }
                w.close("a:solidFill");
            }
        }
    }
}

/// 主题色（`schemeClr val`）。
///
/// 完整列表见 ECMA-376 Part 1, §20.1.2.3.22。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Hash)]
pub enum SchemeColor {
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
    Hlink,
    /// 已访问超链接。
    FolHlink,
    /// 浅色 1。
    Lt1,
    /// 浅色 2。
    Lt2,
    /// 深色 1。
    Dk1,
    /// 深色 2。
    Dk2,
}

impl SchemeColor {
    /// 转 OOXML 字面量。
    pub fn as_str(self) -> &'static str {
        match self {
            SchemeColor::Background1 => "bg1",
            SchemeColor::Background2 => "bg2",
            SchemeColor::Text1 => "tx1",
            SchemeColor::Text2 => "tx2",
            SchemeColor::Accent1 => "accent1",
            SchemeColor::Accent2 => "accent2",
            SchemeColor::Accent3 => "accent3",
            SchemeColor::Accent4 => "accent4",
            SchemeColor::Accent5 => "accent5",
            SchemeColor::Accent6 => "accent6",
            SchemeColor::Hlink => "hlink",
            SchemeColor::FolHlink => "folHlink",
            SchemeColor::Lt1 => "lt1",
            SchemeColor::Lt2 => "lt2",
            SchemeColor::Dk1 => "dk1",
            SchemeColor::Dk2 => "dk2",
        }
    }
}

impl FromStr for SchemeColor {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            "bg1" => SchemeColor::Background1,
            "bg2" => SchemeColor::Background2,
            "tx1" => SchemeColor::Text1,
            "tx2" => SchemeColor::Text2,
            "accent1" => SchemeColor::Accent1,
            "accent2" => SchemeColor::Accent2,
            "accent3" => SchemeColor::Accent3,
            "accent4" => SchemeColor::Accent4,
            "accent5" => SchemeColor::Accent5,
            "accent6" => SchemeColor::Accent6,
            "hlink" => SchemeColor::Hlink,
            "folHlink" => SchemeColor::FolHlink,
            "lt1" => SchemeColor::Lt1,
            "lt2" => SchemeColor::Lt2,
            "dk1" => SchemeColor::Dk1,
            "dk2" => SchemeColor::Dk2,
            _ => return Err(()),
        })
    }
}

/// 预设颜色（`prstClr val`）。
///
/// 完整列表见 ECMA-376 Part 1, §20.1.2.3.23。
/// 本枚举仅暴露 ECMA-376 中规定的 147 个命名颜色；与 python-pptx
/// 的 `MSO_THEME_COLOR` 不同，**这是颜色值而非主题色**。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Hash)]
pub enum PresetColor {
    AliceBlue,
    AntiqueWhite,
    Aqua,
    Aquamarine,
    Azure,
    Beige,
    Bisque,
    Black,
    BlanchedAlmond,
    Blue,
    BlueViolet,
    Brown,
    BurlyWood,
    CadetBlue,
    Chartreuse,
    Chocolate,
    Coral,
    CornflowerBlue,
    Cornsilk,
    Crimson,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGoldenrod,
    DarkGray,
    DarkGreen,
    DarkGrey,
    DarkKhaki,
    DarkMagenta,
    DarkOliveGreen,
    DarkOrange,
    DarkOrchid,
    DarkRed,
    DarkSalmon,
    DarkSeaGreen,
    DarkSlateBlue,
    DarkSlateGray,
    DarkSlateGrey,
    DarkTurquoise,
    DarkViolet,
    DeepPink,
    DeepSkyBlue,
    DimGray,
    DimGrey,
    DodgerBlue,
    Firebrick,
    FloralWhite,
    ForestGreen,
    Fuchsia,
    Gainsboro,
    GhostWhite,
    Gold,
    Goldenrod,
    Gray,
    Green,
    GreenYellow,
    Grey,
    Honeydew,
    HotPink,
    IndianRed,
    Indigo,
    Ivory,
    Khaki,
    Lavender,
    LavenderBlush,
    LawnGreen,
    LemonChiffon,
    LightBlue,
    LightCoral,
    LightCyan,
    LightGoldenrodYellow,
    LightGray,
    LightGreen,
    LightGrey,
    LightPink,
    LightSalmon,
    LightSeaGreen,
    LightSkyBlue,
    LightSlateGray,
    LightSlateGrey,
    LightSteelBlue,
    LightYellow,
    Lime,
    LimeGreen,
    Linen,
    Magenta,
    Maroon,
    MediumAquamarine,
    MediumBlue,
    MediumOrchid,
    MediumPurple,
    MediumSeaGreen,
    MediumSlateBlue,
    MediumSpringGreen,
    MediumTurquoise,
    MediumVioletRed,
    MidnightBlue,
    MintCream,
    MistyRose,
    Moccasin,
    NavajoWhite,
    Navy,
    OldLace,
    Olive,
    OliveDrab,
    Orange,
    OrangeRed,
    Orchid,
    PaleGoldenrod,
    PaleGreen,
    PaleTurquoise,
    PaleVioletRed,
    PapayaWhip,
    PeachPuff,
    Peru,
    Pink,
    Plum,
    PowderBlue,
    Purple,
    Red,
    RosyBrown,
    RoyalBlue,
    SaddleBrown,
    Salmon,
    SandyBrown,
    SeaGreen,
    SeaShell,
    Sienna,
    Silver,
    SkyBlue,
    SlateBlue,
    SlateGray,
    SlateGrey,
    Snow,
    SpringGreen,
    SteelBlue,
    Tan,
    Teal,
    Thistle,
    Tomato,
    Turquoise,
    Violet,
    Wheat,
    White,
    WhiteSmoke,
    Yellow,
    YellowGreen,
}

impl PresetColor {
    /// 转 OOXML 字面量（camelCase）。
    pub fn as_str(self) -> &'static str {
        // 用 serde-like 派生? 直接展开更稳
        match self {
            PresetColor::AliceBlue => "aliceBlue",
            PresetColor::AntiqueWhite => "antiqueWhite",
            PresetColor::Aqua => "aqua",
            PresetColor::Aquamarine => "aquamarine",
            PresetColor::Azure => "azure",
            PresetColor::Beige => "beige",
            PresetColor::Bisque => "bisque",
            PresetColor::Black => "black",
            PresetColor::BlanchedAlmond => "blanchedAlmond",
            PresetColor::Blue => "blue",
            PresetColor::BlueViolet => "blueViolet",
            PresetColor::Brown => "brown",
            PresetColor::BurlyWood => "burlyWood",
            PresetColor::CadetBlue => "cadetBlue",
            PresetColor::Chartreuse => "chartreuse",
            PresetColor::Chocolate => "chocolate",
            PresetColor::Coral => "coral",
            PresetColor::CornflowerBlue => "cornflowerBlue",
            PresetColor::Cornsilk => "cornsilk",
            PresetColor::Crimson => "crimson",
            PresetColor::Cyan => "cyan",
            PresetColor::DarkBlue => "darkBlue",
            PresetColor::DarkCyan => "darkCyan",
            PresetColor::DarkGoldenrod => "darkGoldenrod",
            PresetColor::DarkGray => "darkGray",
            PresetColor::DarkGreen => "darkGreen",
            PresetColor::DarkGrey => "darkGrey",
            PresetColor::DarkKhaki => "darkKhaki",
            PresetColor::DarkMagenta => "darkMagenta",
            PresetColor::DarkOliveGreen => "darkOliveGreen",
            PresetColor::DarkOrange => "darkOrange",
            PresetColor::DarkOrchid => "darkOrchid",
            PresetColor::DarkRed => "darkRed",
            PresetColor::DarkSalmon => "darkSalmon",
            PresetColor::DarkSeaGreen => "darkSeaGreen",
            PresetColor::DarkSlateBlue => "darkSlateBlue",
            PresetColor::DarkSlateGray => "darkSlateGray",
            PresetColor::DarkSlateGrey => "darkSlateGrey",
            PresetColor::DarkTurquoise => "darkTurquoise",
            PresetColor::DarkViolet => "darkViolet",
            PresetColor::DeepPink => "deepPink",
            PresetColor::DeepSkyBlue => "deepSkyBlue",
            PresetColor::DimGray => "dimGray",
            PresetColor::DimGrey => "dimGrey",
            PresetColor::DodgerBlue => "dodgerBlue",
            PresetColor::Firebrick => "firebrick",
            PresetColor::FloralWhite => "floralWhite",
            PresetColor::ForestGreen => "forestGreen",
            PresetColor::Fuchsia => "fuchsia",
            PresetColor::Gainsboro => "gainsboro",
            PresetColor::GhostWhite => "ghostWhite",
            PresetColor::Gold => "gold",
            PresetColor::Goldenrod => "goldenrod",
            PresetColor::Gray => "gray",
            PresetColor::Green => "green",
            PresetColor::GreenYellow => "greenYellow",
            PresetColor::Grey => "grey",
            PresetColor::Honeydew => "honeydew",
            PresetColor::HotPink => "hotPink",
            PresetColor::IndianRed => "indianRed",
            PresetColor::Indigo => "indigo",
            PresetColor::Ivory => "ivory",
            PresetColor::Khaki => "khaki",
            PresetColor::Lavender => "lavender",
            PresetColor::LavenderBlush => "lavenderBlush",
            PresetColor::LawnGreen => "lawnGreen",
            PresetColor::LemonChiffon => "lemonChiffon",
            PresetColor::LightBlue => "lightBlue",
            PresetColor::LightCoral => "lightCoral",
            PresetColor::LightCyan => "lightCyan",
            PresetColor::LightGoldenrodYellow => "lightGoldenrodYellow",
            PresetColor::LightGray => "lightGray",
            PresetColor::LightGreen => "lightGreen",
            PresetColor::LightGrey => "lightGrey",
            PresetColor::LightPink => "lightPink",
            PresetColor::LightSalmon => "lightSalmon",
            PresetColor::LightSeaGreen => "lightSeaGreen",
            PresetColor::LightSkyBlue => "lightSkyBlue",
            PresetColor::LightSlateGray => "lightSlateGray",
            PresetColor::LightSlateGrey => "lightSlateGrey",
            PresetColor::LightSteelBlue => "lightSteelBlue",
            PresetColor::LightYellow => "lightYellow",
            PresetColor::Lime => "lime",
            PresetColor::LimeGreen => "limeGreen",
            PresetColor::Linen => "linen",
            PresetColor::Magenta => "magenta",
            PresetColor::Maroon => "maroon",
            PresetColor::MediumAquamarine => "mediumAquamarine",
            PresetColor::MediumBlue => "mediumBlue",
            PresetColor::MediumOrchid => "mediumOrchid",
            PresetColor::MediumPurple => "mediumPurple",
            PresetColor::MediumSeaGreen => "mediumSeaGreen",
            PresetColor::MediumSlateBlue => "mediumSlateBlue",
            PresetColor::MediumSpringGreen => "mediumSpringGreen",
            PresetColor::MediumTurquoise => "mediumTurquoise",
            PresetColor::MediumVioletRed => "mediumVioletRed",
            PresetColor::MidnightBlue => "midnightBlue",
            PresetColor::MintCream => "mintCream",
            PresetColor::MistyRose => "mistyRose",
            PresetColor::Moccasin => "moccasin",
            PresetColor::NavajoWhite => "navajoWhite",
            PresetColor::Navy => "navy",
            PresetColor::OldLace => "oldLace",
            PresetColor::Olive => "olive",
            PresetColor::OliveDrab => "oliveDrab",
            PresetColor::Orange => "orange",
            PresetColor::OrangeRed => "orangeRed",
            PresetColor::Orchid => "orchid",
            PresetColor::PaleGoldenrod => "paleGoldenrod",
            PresetColor::PaleGreen => "paleGreen",
            PresetColor::PaleTurquoise => "paleTurquoise",
            PresetColor::PaleVioletRed => "paleVioletRed",
            PresetColor::PapayaWhip => "papayaWhip",
            PresetColor::PeachPuff => "peachPuff",
            PresetColor::Peru => "peru",
            PresetColor::Pink => "pink",
            PresetColor::Plum => "plum",
            PresetColor::PowderBlue => "powderBlue",
            PresetColor::Purple => "purple",
            PresetColor::Red => "red",
            PresetColor::RosyBrown => "rosyBrown",
            PresetColor::RoyalBlue => "royalBlue",
            PresetColor::SaddleBrown => "saddleBrown",
            PresetColor::Salmon => "salmon",
            PresetColor::SandyBrown => "sandyBrown",
            PresetColor::SeaGreen => "seaGreen",
            PresetColor::SeaShell => "seaShell",
            PresetColor::Sienna => "sienna",
            PresetColor::Silver => "silver",
            PresetColor::SkyBlue => "skyBlue",
            PresetColor::SlateBlue => "slateBlue",
            PresetColor::SlateGray => "slateGray",
            PresetColor::SlateGrey => "slateGrey",
            PresetColor::Snow => "snow",
            PresetColor::SpringGreen => "springGreen",
            PresetColor::SteelBlue => "steelBlue",
            PresetColor::Tan => "tan",
            PresetColor::Teal => "teal",
            PresetColor::Thistle => "thistle",
            PresetColor::Tomato => "tomato",
            PresetColor::Turquoise => "turquoise",
            PresetColor::Violet => "violet",
            PresetColor::Wheat => "wheat",
            PresetColor::White => "white",
            PresetColor::WhiteSmoke => "whiteSmoke",
            PresetColor::Yellow => "yellow",
            PresetColor::YellowGreen => "yellowGreen",
        }
    }
}

impl FromStr for PresetColor {
    type Err = ();
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        // 简化: 接受所有 147 种预设名; 用 lower-case 比较
        let all = [
            ("aliceBlue", PresetColor::AliceBlue),
            ("antiqueWhite", PresetColor::AntiqueWhite),
            ("aqua", PresetColor::Aqua),
            ("aquamarine", PresetColor::Aquamarine),
            ("azure", PresetColor::Azure),
            ("beige", PresetColor::Beige),
            ("bisque", PresetColor::Bisque),
            ("black", PresetColor::Black),
            ("blanchedAlmond", PresetColor::BlanchedAlmond),
            ("blue", PresetColor::Blue),
            ("blueViolet", PresetColor::BlueViolet),
            ("brown", PresetColor::Brown),
            ("burlyWood", PresetColor::BurlyWood),
            ("cadetBlue", PresetColor::CadetBlue),
            ("chartreuse", PresetColor::Chartreuse),
            ("chocolate", PresetColor::Chocolate),
            ("coral", PresetColor::Coral),
            ("cornflowerBlue", PresetColor::CornflowerBlue),
            ("cornsilk", PresetColor::Cornsilk),
            ("crimson", PresetColor::Crimson),
            ("cyan", PresetColor::Cyan),
            ("darkBlue", PresetColor::DarkBlue),
            ("darkCyan", PresetColor::DarkCyan),
            ("darkGoldenrod", PresetColor::DarkGoldenrod),
            ("darkGray", PresetColor::DarkGray),
            ("darkGreen", PresetColor::DarkGreen),
            ("darkGrey", PresetColor::DarkGrey),
            ("darkKhaki", PresetColor::DarkKhaki),
            ("darkMagenta", PresetColor::DarkMagenta),
            ("darkOliveGreen", PresetColor::DarkOliveGreen),
            ("darkOrange", PresetColor::DarkOrange),
            ("darkOrchid", PresetColor::DarkOrchid),
            ("darkRed", PresetColor::DarkRed),
            ("darkSalmon", PresetColor::DarkSalmon),
            ("darkSeaGreen", PresetColor::DarkSeaGreen),
            ("darkSlateBlue", PresetColor::DarkSlateBlue),
            ("darkSlateGray", PresetColor::DarkSlateGray),
            ("darkSlateGrey", PresetColor::DarkSlateGrey),
            ("darkTurquoise", PresetColor::DarkTurquoise),
            ("darkViolet", PresetColor::DarkViolet),
            ("deepPink", PresetColor::DeepPink),
            ("deepSkyBlue", PresetColor::DeepSkyBlue),
            ("dimGray", PresetColor::DimGray),
            ("dimGrey", PresetColor::DimGrey),
            ("dodgerBlue", PresetColor::DodgerBlue),
            ("firebrick", PresetColor::Firebrick),
            ("floralWhite", PresetColor::FloralWhite),
            ("forestGreen", PresetColor::ForestGreen),
            ("fuchsia", PresetColor::Fuchsia),
            ("gainsboro", PresetColor::Gainsboro),
            ("ghostWhite", PresetColor::GhostWhite),
            ("gold", PresetColor::Gold),
            ("goldenrod", PresetColor::Goldenrod),
            ("gray", PresetColor::Gray),
            ("green", PresetColor::Green),
            ("greenYellow", PresetColor::GreenYellow),
            ("grey", PresetColor::Grey),
            ("honeydew", PresetColor::Honeydew),
            ("hotPink", PresetColor::HotPink),
            ("indianRed", PresetColor::IndianRed),
            ("indigo", PresetColor::Indigo),
            ("ivory", PresetColor::Ivory),
            ("khaki", PresetColor::Khaki),
            ("lavender", PresetColor::Lavender),
            ("lavenderBlush", PresetColor::LavenderBlush),
            ("lawnGreen", PresetColor::LawnGreen),
            ("lemonChiffon", PresetColor::LemonChiffon),
            ("lightBlue", PresetColor::LightBlue),
            ("lightCoral", PresetColor::LightCoral),
            ("lightCyan", PresetColor::LightCyan),
            ("lightGoldenrodYellow", PresetColor::LightGoldenrodYellow),
            ("lightGray", PresetColor::LightGray),
            ("lightGreen", PresetColor::LightGreen),
            ("lightGrey", PresetColor::LightGrey),
            ("lightPink", PresetColor::LightPink),
            ("lightSalmon", PresetColor::LightSalmon),
            ("lightSeaGreen", PresetColor::LightSeaGreen),
            ("lightSkyBlue", PresetColor::LightSkyBlue),
            ("lightSlateGray", PresetColor::LightSlateGray),
            ("lightSlateGrey", PresetColor::LightSlateGrey),
            ("lightSteelBlue", PresetColor::LightSteelBlue),
            ("lightYellow", PresetColor::LightYellow),
            ("lime", PresetColor::Lime),
            ("limeGreen", PresetColor::LimeGreen),
            ("linen", PresetColor::Linen),
            ("magenta", PresetColor::Magenta),
            ("maroon", PresetColor::Maroon),
            ("mediumAquamarine", PresetColor::MediumAquamarine),
            ("mediumBlue", PresetColor::MediumBlue),
            ("mediumOrchid", PresetColor::MediumOrchid),
            ("mediumPurple", PresetColor::MediumPurple),
            ("mediumSeaGreen", PresetColor::MediumSeaGreen),
            ("mediumSlateBlue", PresetColor::MediumSlateBlue),
            ("mediumSpringGreen", PresetColor::MediumSpringGreen),
            ("mediumTurquoise", PresetColor::MediumTurquoise),
            ("mediumVioletRed", PresetColor::MediumVioletRed),
            ("midnightBlue", PresetColor::MidnightBlue),
            ("mintCream", PresetColor::MintCream),
            ("mistyRose", PresetColor::MistyRose),
            ("moccasin", PresetColor::Moccasin),
            ("navajoWhite", PresetColor::NavajoWhite),
            ("navy", PresetColor::Navy),
            ("oldLace", PresetColor::OldLace),
            ("olive", PresetColor::Olive),
            ("oliveDrab", PresetColor::OliveDrab),
            ("orange", PresetColor::Orange),
            ("orangeRed", PresetColor::OrangeRed),
            ("orchid", PresetColor::Orchid),
            ("paleGoldenrod", PresetColor::PaleGoldenrod),
            ("paleGreen", PresetColor::PaleGreen),
            ("paleTurquoise", PresetColor::PaleTurquoise),
            ("paleVioletRed", PresetColor::PaleVioletRed),
            ("papayaWhip", PresetColor::PapayaWhip),
            ("peachPuff", PresetColor::PeachPuff),
            ("peru", PresetColor::Peru),
            ("pink", PresetColor::Pink),
            ("plum", PresetColor::Plum),
            ("powderBlue", PresetColor::PowderBlue),
            ("purple", PresetColor::Purple),
            ("red", PresetColor::Red),
            ("rosyBrown", PresetColor::RosyBrown),
            ("royalBlue", PresetColor::RoyalBlue),
            ("saddleBrown", PresetColor::SaddleBrown),
            ("salmon", PresetColor::Salmon),
            ("sandyBrown", PresetColor::SandyBrown),
            ("seaGreen", PresetColor::SeaGreen),
            ("seaShell", PresetColor::SeaShell),
            ("sienna", PresetColor::Sienna),
            ("silver", PresetColor::Silver),
            ("skyBlue", PresetColor::SkyBlue),
            ("slateBlue", PresetColor::SlateBlue),
            ("slateGray", PresetColor::SlateGray),
            ("slateGrey", PresetColor::SlateGrey),
            ("snow", PresetColor::Snow),
            ("springGreen", PresetColor::SpringGreen),
            ("steelBlue", PresetColor::SteelBlue),
            ("tan", PresetColor::Tan),
            ("teal", PresetColor::Teal),
            ("thistle", PresetColor::Thistle),
            ("tomato", PresetColor::Tomato),
            ("turquoise", PresetColor::Turquoise),
            ("violet", PresetColor::Violet),
            ("wheat", PresetColor::Wheat),
            ("white", PresetColor::White),
            ("whiteSmoke", PresetColor::WhiteSmoke),
            ("yellow", PresetColor::Yellow),
            ("yellowGreen", PresetColor::YellowGreen),
        ];
        for (k, v) in &all {
            if *k == s {
                return Ok(*v);
            }
        }
        Err(())
    }
}

// ====================================================================
// 高阶颜色 API：对标 python-pptx 中 `pptx.dml.color.ColorFormat`。
// ====================================================================

/// 颜色高阶视图（`pptx.dml.color.ColorFormat`）。
///
/// 与 [`Color`] 区别：
///
/// - [`Color`] 是"OOXML 模型层"表达，**只关心值**；
/// - [`ColorFormat`] 是"高阶 API 层"包装，**关心"颜色如何被使用"**——
///   例如字体前景色、字体背景高亮、形状填充、形状边框等，引用同一个
///   [`Color`] 时通过 `&mut Color` 共享。
///
/// # 与 python-pptx 的对应
///
/// - `pptx.dml.color.ColorFormat` ←→ [`ColorFormat`]；
/// - `font.color.rgb = RGBColor(...)` ←→ `color_format.set_rgb(...)`。
///
/// # 设计要点
///
/// - **借用 + 透明代理**：构造时传入 `&mut Color`，**所有写操作直接修改
///   原值**；这与 python-pptx 的 `font.color.rgb = X` 行为一致——
///   `font.color` 是个 proxy，实际写回的是底层 `a:srgbClr` 元素。
/// - **零分配**：除读取时的临时 `String` 外不触发堆分配。
/// - **类型安全**：所有 setter 接受 [`Color`] / `RGBColor` / [`SchemeColor`]
///   / [`PresetColor`] 任一；`From<impl Into<Color>>` 自动转换。
#[derive(Debug)]
pub struct ColorFormat<'a> {
    /// 借用底层颜色（OXML 模型层）。所有写都走这个引用。
    color: &'a mut Color,
    /// 父上下文是"什么角色"——决定 type / 行为（仅供 hint / 调试）。
    role: ColorRole,
}

/// 颜色在文档中的角色（仅 hint，不参与序列化）。
///
/// python-pptx 中用 `color.type` 反映"是前景 / 背景 / 文本"等角色；
/// 本枚举用于 [`ColorFormat::role`] 的返回值。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum ColorRole {
    /// 前景 / 文本。
    #[default]
    Foreground,
    /// 背景。
    Background,
    /// 主题色（schemeClr）。
    Theme,
    /// 未知 / 不关心。
    Unknown,
}

impl<'a> ColorFormat<'a> {
    /// 构造一个颜色格式视图。
    ///
    /// # 参数
    /// - `color`：底层颜色的可变引用；
    /// - `role`：角色（默认 [`ColorRole::Foreground`]）。
    pub fn new(color: &'a mut Color) -> Self {
        ColorFormat {
            color,
            role: ColorRole::Foreground,
        }
    }
    /// 构造时指定角色。
    pub fn with_role(color: &'a mut Color, role: ColorRole) -> Self {
        ColorFormat { color, role }
    }

    /// 角色（python-pptx `ColorFormat.type` 的简化版）。
    pub fn role(&self) -> ColorRole {
        self.role
    }

    /// 当前颜色的"类型"（python-pptx `color.type`）。
    pub fn color_type(&self) -> super::simpletypes::MsoColorType {
        match self.color {
            Color::None => super::simpletypes::MsoColorType::Mixed,
            Color::Scheme(_) => super::simpletypes::MsoColorType::SchemeColor,
            Color::Preset(_) => super::simpletypes::MsoColorType::PresetColor,
            Color::RGB(_) => super::simpletypes::MsoColorType::Rgb,
        }
    }

    /// 是否已设置颜色（不等于 [`Color::None`]）。
    pub fn is_set(&self) -> bool {
        !matches!(self.color, Color::None)
    }

    // --------- 读 ---------

    /// 取内部颜色克隆。
    pub fn color(&self) -> Color {
        self.color.clone()
    }
    /// 若当前是 RGB，取出 sRGB；否则返回 `None`。
    pub fn rgb(&self) -> Option<RGBColor> {
        match &*self.color {
            Color::RGB(c) => Some(*c),
            _ => None,
        }
    }
    /// 若当前是 theme color，取出 scheme 枚举；否则 `None`。
    pub fn theme_color(&self) -> Option<super::simpletypes::MsoThemeColorIndex> {
        match &*self.color {
            Color::Scheme(s) => Some(super::simpletypes::MsoThemeColorIndex::from_scheme(*s)),
            _ => None,
        }
    }

    /// 亮度调整（python-pptx 风格）。
    ///
    /// 返回当前颜色按 `lumMod` / `lumOff` 调整后的"近似"亮度。
    /// **注意**：本方法只计算并返回 f32 (0.0..=1.0)，**不**修改底层
    /// `Color`。如需在 XML 中表达亮度调整，请在 [`Color`] 上加修饰子元素
    /// （路线图 0.2.0）。
    pub fn brightness(&self) -> f32 {
        match &*self.color {
            Color::RGB(c) => {
                // 用 sRGB → Y (ITU-R BT.601) 亮度公式
                let y = 0.299 * c.0 as f32 + 0.587 * c.1 as f32 + 0.114 * c.2 as f32;
                y / 255.0
            }
            Color::Scheme(s) => match s {
                // 给主题色一个粗略亮度估计（仅供 UI 调试用，非严格计算）
                SchemeColor::Background1 | SchemeColor::Text1 | SchemeColor::Lt1 => 1.0,
                SchemeColor::Background2 | SchemeColor::Text2 | SchemeColor::Lt2 => 0.9,
                SchemeColor::Dk1 | SchemeColor::Dk2 => 0.1,
                _ => 0.5,
            },
            Color::Preset(p) => preset_brightness(*p),
            Color::None => 0.0,
        }
    }

    // --------- 写 ---------

    /// 直接覆盖为任意 [`Color`]。对应 `color = some_color`。
    pub fn set(&mut self, c: impl Into<Color>) {
        *self.color = c.into();
    }

    /// 设为 sRGB 颜色（对应 python-pptx `color.rgb = RGBColor(r, g, b)`）。
    pub fn set_rgb(&mut self, c: impl Into<RGBColor>) {
        *self.color = Color::RGB(c.into());
    }

    /// 设为预设颜色。
    pub fn set_preset(&mut self, p: PresetColor) {
        *self.color = Color::Preset(p);
    }

    /// 设为 schemeClr 主题色。
    pub fn set_theme(&mut self, t: super::simpletypes::MsoThemeColorIndex) {
        if let Some(s) = t.as_str() {
            // 用字面量反查回 SchemeColor
            if let Ok(sc) = s.parse::<SchemeColor>() {
                *self.color = Color::Scheme(sc);
            }
        }
    }

    /// 重置为未设置。
    pub fn clear(&mut self) {
        *self.color = Color::None;
    }
}

/// 预设颜色亮度估算（粗略）。
fn preset_brightness(p: PresetColor) -> f32 {
    use PresetColor::*;
    // 把"亮色"统一判定为亮度 1，"深色"为 0
    match p {
        White | WhiteSmoke | Snow | GhostWhite | Azure | Ivory | MintCream | Honeydew
        | FloralWhite | AliceBlue | Lavender | LavenderBlush | MistyRose | SeaShell | Linen
        | OldLace | Cornsilk | LemonChiffon | LightYellow | Beige | PapayaWhip | BlanchedAlmond
        | Bisque | PeachPuff | NavajoWhite | Wheat | LightGoldenrodYellow | AntiqueWhite
        | Gainsboro | LightGray | LightGrey | PowderBlue | LightCyan | LightBlue | LightPink
        | LightCoral | LightSalmon | LightSkyBlue | LightSteelBlue | LightSeaGreen
        | PaleGoldenrod | PaleGreen | PaleTurquoise | PaleVioletRed => 0.9,
        Black | MidnightBlue | Navy | DarkBlue | DarkRed | DarkGreen | DarkMagenta | DarkViolet
        | Indigo | DarkSlateBlue | DarkSlateGray | DarkSlateGrey | DarkCyan | DarkOliveGreen
        | DarkKhaki | Maroon | Purple | DarkGoldenrod | DarkOrchid | Firebrick | SaddleBrown
        | Sienna | Brown => 0.1,
        _ => 0.5,
    }
}

impl<'a> From<&'a mut Color> for ColorFormat<'a> {
    fn from(c: &'a mut Color) -> Self {
        ColorFormat::new(c)
    }
}
