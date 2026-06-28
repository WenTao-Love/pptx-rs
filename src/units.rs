//! 单位、基本类型与颜色。
//!
//! 本模块集中管理 `pptx-rs` 中所有"非 OOXML 强类型但又频繁使用"的小型数据载体。
//! 它是 `crate::lib` 中"高阶 API 层"与"OPC / OOXML 模型层"之间的语义桥梁。
//!
//! # PowerPoint 的几何单位
//!
//! OOXML 中所有几何属性（`<a:off>` / `<a:ext>` / 行高 / 列宽 / 边框 / 缩进）都是
//! **EMU（English Metric Unit）** 整数。换算关系：
//!
//! | 单位      | EMU            | 备注                           |
//! | --------- | -------------- | ------------------------------ |
//! | 1 inch    | 914 400 EMU    | 1 in = 914 400                |
//! | 1 cm      | 360 000 EMU    | 1 cm = 360 000                |
//! | 1 pt      | 12 700 EMU     | 1 pt = 12 700                 |
//! | 1 sp      | 6 350 EMU      | 1 pt = 20 sp（OpenType 子像素）|
//!
//! 库内统一以 `i64` 形式的 EMU 作内部表示，理由：
//!
//! 1. **无浮点漂移**：在缩放 / 旋转 / 多次叠加时仍能保持整型精度。
//! 2. **互操作零成本**：与 python-pptx / Open XML SDK / LibreOffice 完全一致。
//! 3. **可读性**：用户阅读 `Emu(914400)` 与 `Inches(1.0)` 时是等价的。
//!
//! # 类型分层
//!
//! - [`Emu`]：绝对内部单位，整数 NewType。
//! - [`Pt`] / [`Inches`] / [`Cm`]：外部友好单位，浮点 NewType。
//! - [`EmuPoint`]：2D 点（x, y），以 EMU 整数存储，主要给连接器/手绘 freeform 用。
//! - [`EmuExt`]：扩展 trait，统一从 `f64` / `f32` / `i32` / `i64` / `Pt` / `Inches` / `Cm` 转换到 `Emu`。
//! - [`RGBColor`]：sRGB 颜色三字节 0-255，对应 OOXML `<a:srgbClr val="RRGGBB"/>`。
//!
//! # 设计哲学
//!
//! - **不暴露 `f64` 形参**：高阶 API 的"位置 / 大小"参数应接受 NewType 或显式单位
//!   转换。例如 `add_textbox(Inches(1.0), Inches(1.0), Inches(8.0), Inches(1.0))`，
//!   让"单位混用"在编译期就被拒绝。
//! - **浮点语义**：`f64` 在 `EmuExt` 中默认解释为"英寸"——这是最常用的隐含单位；
//!   若需其它语义请显式包装 `Pt(...)` / `Cm(...)`。
//! - **颜色不归此模块管**：主题色（schemeClr）、系统色（sysClr）、preset 颜色
//!   见 [`crate::oxml::color`]。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.util.Pt` / `pptx.util.Inches` / `pptx.util.Cm` / `pptx.util.Emu` ←→ 同名 NewType。
//! - `pptx.util.Length` 的 `.x` / `.y` 抽象 → 本库通过 [`EmuPoint`] 暴露。
//! - `pptx.dml.color.RGBColor` ←→ [`RGBColor`]（构造时 `(r, g, b)` 元组也支持）。
//!
//! # 示例
//!
//! ```
//! use pptx::{Emu, Inches, Pt, Cm, EmuExt, RGBColor};
//!
//! // 隐式单位转换
//! let w: Emu = Inches(8.5).emu();
//! assert_eq!(w.value(), (8.5 * 914_400.0) as i64);
//!
//! // 跨单位混算
//! let h = Pt(72.0).emu() + Inches(1.0).emu();
//!
//! // 颜色
//! let red: RGBColor = (255, 0, 0).into();
//! assert_eq!(red, RGBColor::RED);
//! ```

use std::fmt;

/// EMU（English Metric Unit），所有内部几何计算都使用 `i64` 形式的 EMU。
///
/// EMU 在 OOXML 中是绝对整数，可保证精度无损；库内任何"宽 / 高 / 偏移 / 行高 / 列宽"
/// 字段均以本类型为内部表示。NewType 包装的 `i64` 同时：
///
/// - 提供 `value()` / `new()` 与 f64 单位的双向换算；
/// - 暴露 `+` / `-` / `*` 等运算符便于几何计算；
/// - 实现 `Copy + Default + Ord`，可直接作为 `BTreeMap` key。
#[derive(Copy, Clone, Default, Eq, PartialEq, Ord, PartialOrd, Hash)]
pub struct Emu(
    /// EMU 数值（可为负）。
    pub i64,
);

impl Emu {
    /// 0 EMU。用于默认值与"无偏移"语义。
    pub const ZERO: Emu = Emu(0);

    /// 直接以 i64 构造。
    ///
    /// # 参数
    /// - `v`：EMU 数值（可为负）。
    ///
    /// # 示例
    /// ```
    /// use pptx::Emu;
    /// let off = Emu::new(914_400);
    /// assert_eq!(off.value(), 914_400);
    /// ```
    #[inline]
    pub const fn new(v: i64) -> Self {
        Emu(v)
    }

    /// 取内部 i64 值。
    #[inline]
    pub const fn value(self) -> i64 {
        self.0
    }

    /// 转为英寸（`f64`）。
    #[inline]
    pub const fn inches(self) -> f64 {
        self.0 as f64 / 914_400.0
    }

    /// 转为磅（`f64`）。
    #[inline]
    pub const fn pt(self) -> f64 {
        self.0 as f64 / 12_700.0
    }

    /// 转为厘米（`f64`）。
    #[inline]
    pub const fn cm(self) -> f64 {
        self.0 as f64 / 360_000.0
    }
}

impl fmt::Debug for Emu {
    /// Debug 输出同时给出原始 EMU 与换算后的英寸，便于日志阅读。
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Emu({} = {} in)", self.0, self.inches())
    }
}

impl fmt::Display for Emu {
    /// Display 形式为 `"<n> EMU"`，便于错误消息中嵌入。
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{} EMU", self.0)
    }
}

// 算术运算符：让 Emu 参与几何计算时无需 .0 拆箱。
impl std::ops::Add for Emu {
    type Output = Emu;
    #[inline]
    fn add(self, rhs: Emu) -> Emu {
        Emu(self.0 + rhs.0)
    }
}
impl std::ops::Sub for Emu {
    type Output = Emu;
    #[inline]
    fn sub(self, rhs: Emu) -> Emu {
        Emu(self.0 - rhs.0)
    }
}
impl std::ops::Neg for Emu {
    type Output = Emu;
    #[inline]
    fn neg(self) -> Emu {
        Emu(-self.0)
    }
}
impl std::ops::Mul<i64> for Emu {
    type Output = Emu;
    #[inline]
    fn mul(self, rhs: i64) -> Emu {
        Emu(self.0 * rhs)
    }
}

/// 单位转换 trait，封装"任意类型 → [`Emu`]"的语义。
///
/// 通过 `EmuExt`，任意实现了该 trait 的类型都能 `.emu()` 转为 `Emu`。
/// 这样 `Inches(1.0).emu() + Pt(72.0).emu()` 这类混合表达式可以一行写完，
/// 而不必每次手写 `Emu::new(...)`。
///
/// 各类实现的隐含语义：
///
/// - `Emu` / `i64` / `i32` / `u32`：原值就是 EMU；
/// - `f64` / `f32`：解释为"英寸"（最常用的隐含单位）；
/// - `Pt`：磅（1 pt = 12 700 EMU）；
/// - `Inches`：英寸（1 in = 914 400 EMU）；
/// - `Cm`：厘米（1 cm = 360 000 EMU）。
pub trait EmuExt: Copy {
    /// 转为 [`Emu`]。具体换算由实现决定。
    fn emu(self) -> Emu;
}

impl EmuExt for Emu {
    /// 恒等映射。
    #[inline]
    fn emu(self) -> Emu {
        self
    }
}

impl EmuExt for i64 {
    /// 原值视为 EMU。
    #[inline]
    fn emu(self) -> Emu {
        Emu(self)
    }
}

impl EmuExt for i32 {
    /// 原值视为 EMU（提升到 i64）。
    #[inline]
    fn emu(self) -> Emu {
        Emu(self as i64)
    }
}

impl EmuExt for u32 {
    /// 原值视为 EMU（提升到 i64）。
    #[inline]
    fn emu(self) -> Emu {
        Emu(self as i64)
    }
}

impl EmuExt for f64 {
    /// `f64` 当作**英寸**处理。
    #[inline]
    fn emu(self) -> Emu {
        Emu((self * 914_400.0) as i64)
    }
}

impl EmuExt for f32 {
    /// `f32` 当作**英寸**处理。
    #[inline]
    fn emu(self) -> Emu {
        Emu((self as f64 * 914_400.0) as i64)
    }
}

/// 磅：1 pt = 12700 EMU。
///
/// 文本大小（`sz`）、行距（`lnSpc`）、缩进（`indent`）等文本相关属性
/// 优先以 `Pt` 为单位。包装为 NewType 后编译器可阻止与 `Inches` 误混。
#[derive(Copy, Clone, Debug, Default, PartialEq, PartialOrd)]
pub struct Pt(
    /// 磅值（f64）。
    pub f64,
);

impl Pt {
    /// 构造磅值。
    #[inline]
    pub const fn new(v: f64) -> Self {
        Pt(v)
    }

    /// 取内部 f64。
    #[inline]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl EmuExt for Pt {
    /// 1 pt = 12 700 EMU。
    #[inline]
    fn emu(self) -> Emu {
        Emu((self.0 * 12_700.0) as i64)
    }
}

/// 英寸：1 inch = 914400 EMU。
///
/// 形状的 `cx` / `cy` / `x` / `y` 等几何属性最常用的语义单位。
#[derive(Copy, Clone, Debug, Default, PartialEq, PartialOrd)]
pub struct Inches(
    /// 英寸值（f64）。
    pub f64,
);

impl Inches {
    /// 构造英寸值。
    #[inline]
    pub const fn new(v: f64) -> Self {
        Inches(v)
    }

    /// 取内部 f64。
    #[inline]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl EmuExt for Inches {
    /// 1 in = 914 400 EMU。
    #[inline]
    fn emu(self) -> Emu {
        Emu((self.0 * 914_400.0) as i64)
    }
}

/// 厘米：1 cm = 360000 EMU。
///
/// 仅作为欧式使用习惯的可选入口，与 Inches / Pt 完全等价（换算即可）。
#[derive(Copy, Clone, Debug, Default, PartialEq, PartialOrd)]
pub struct Cm(
    /// 厘米值（f64）。
    pub f64,
);

impl Cm {
    /// 构造厘米值。
    #[inline]
    pub const fn new(v: f64) -> Self {
        Cm(v)
    }

    /// 取内部 f64。
    #[inline]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl EmuExt for Cm {
    /// 1 cm = 360 000 EMU。
    #[inline]
    fn emu(self) -> Emu {
        Emu((self.0 * 360_000.0) as i64)
    }
}

/// 2D 几何点（EMU 整数坐标）。
///
/// 主要给以下场景使用：
/// - **连接器**的 begin / end；
/// - **freeform** 的内部路径点（`a:cubicBezTo` / `a:lineTo` / `a:moveTo`）；
/// - **占位符锚点**（`a:off x="" y=""`）。
///
/// # 与 python-pptx 的对应
///
/// python-pptx 中没有 `EmuPoint` 这样的直接类型；最接近的是
/// `(x: int, y: int)` 元组。本库用 NewType 是为了类型安全与可读性。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq, Hash)]
pub struct EmuPoint(
    /// x 坐标（EMU）。
    pub i64,
    /// y 坐标（EMU）。
    pub i64,
);

impl EmuPoint {
    /// 构造新点。
    #[inline]
    pub const fn new(x: i64, y: i64) -> Self {
        EmuPoint(x, y)
    }
    /// 取 x 坐标。
    #[inline]
    pub const fn x(self) -> i64 {
        self.0
    }
    /// 取 y 坐标。
    #[inline]
    pub const fn y(self) -> i64 {
        self.1
    }
    /// 转为 `(x, y)` 元组。
    #[inline]
    pub const fn as_tuple(self) -> (i64, i64) {
        (self.0, self.1)
    }
}

impl From<(i64, i64)> for EmuPoint {
    #[inline]
    fn from((x, y): (i64, i64)) -> Self {
        EmuPoint(x, y)
    }
}
impl From<EmuPoint> for (i64, i64) {
    #[inline]
    fn from(p: EmuPoint) -> Self {
        (p.0, p.1)
    }
}

/// sRGB 颜色（三字节 0-255）。
///
/// 对应 OOXML 中的 `<a:srgbClr val="RRGGBB"/>` 元素。
/// 注意：本类型仅表达"绝对 RGB 颜色"，**不**含透明度（透明度由
/// `<a:alpha val="..."/>` 子元素承载，见 [`crate::oxml::color`]）。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq, Hash)]
pub struct RGBColor(
    /// 红色分量（0-255）。
    pub u8,
    /// 绿色分量（0-255）。
    pub u8,
    /// 蓝色分量（0-255）。
    pub u8,
);

impl RGBColor {
    /// 黑色 `#000000`。
    pub const BLACK: RGBColor = RGBColor(0, 0, 0);
    /// 白色 `#FFFFFF`。
    pub const WHITE: RGBColor = RGBColor(255, 255, 255);
    /// 红色 `#FF0000`。
    pub const RED: RGBColor = RGBColor(255, 0, 0);
    /// 绿色 `#00FF00`。
    pub const GREEN: RGBColor = RGBColor(0, 255, 0);
    /// 蓝色 `#0000FF`。
    pub const BLUE: RGBColor = RGBColor(0, 0, 255);
}

/// 允许 `(r, g, b)` 三元组直接 `into()` 转为 [`RGBColor`]，方便函数签名紧凑。
impl From<(u8, u8, u8)> for RGBColor {
    #[inline]
    fn from((r, g, b): (u8, u8, u8)) -> Self {
        RGBColor(r, g, b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 主要单位换算 round-trip 测试。
    #[test]
    fn unit_conversion() {
        assert_eq!(Pt(72.0).emu().value(), 72 * 12_700);
        assert_eq!(Inches(1.0).emu().value(), 914_400);
        assert_eq!(Cm(1.0).emu().value(), 360_000);
        assert!((Emu(914_400).inches() - 1.0).abs() < 1e-9);
        assert!((Emu(12_700).pt() - 1.0).abs() < 1e-9);
    }

    /// RGBColor 基本转换测试。
    #[test]
    fn rgb() {
        let c: RGBColor = (10, 20, 30).into();
        assert_eq!(c, RGBColor(10, 20, 30));
    }
}
