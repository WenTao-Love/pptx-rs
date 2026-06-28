//! `Shape`：所有形状的"基类"包装。
//!
//! 本文件定义高阶形状的统一接口 [`Shape`]，由 `AutoShape` / `TextBox` /
//! `Picture` / `Group` / `Connector` / `TableShape` 各自实现。
//!
//! # 设计要点
//!
//! - **位置 / 尺寸 / 旋转** 等通用属性被提到 trait 层，避免在每个形状上重复定义；
//! - **类型相关**的属性（如 `Picture::set_stretch`、`Connector::cxn`）保留在
//!   具体类型的 `impl` 块中；
//! - trait 方法均**只读 / 可变**访问 `self`，不返回额外句柄，
//!   避免破坏"借用 oxml 模型"的不变式。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.base.BaseShape` ←→ [`Shape`] trait；
//! - python-pptx 把通用属性都挂在 `BaseShape` 上，python 端可以直接调用；
//!   本库因 Rust trait 限制需要在每个 impl 上"转发"方法。

use crate::units::Emu;

/// 高阶形状的抽象接口。
///
/// 调用方拿到一个 `Shape`（trait object）或具体类型时，可通过此 trait
/// 访问"位置 / 尺寸 / 旋转"等通用属性。具体类型特有属性需用具体类型
/// 句柄访问。
pub trait Shape {
    /// 形状唯一 ID（在所属 slide 内）。
    ///
    /// ID 由 `add_*` 流程自动分配，调用方一般不应手动设置。
    fn id(&self) -> u32;
    /// 设置 ID。
    fn set_id(&mut self, id: u32);
    /// 形状名。
    fn name(&self) -> &str;
    /// 设置形状名。
    fn set_name(&mut self, name: String);
    /// 形状类型（如 `"text_box"` / `"picture"` / ...）。
    ///
    /// 返回值与 [`crate::shape::ShapeKind::shape_type`] 保持一致。
    fn shape_type(&self) -> &'static str;

    /// 左上角 x（EMU）。
    fn left(&self) -> Emu;
    /// 设置左上角 x。
    fn set_left(&mut self, emu: Emu);
    /// 左上角 y（EMU）。
    fn top(&self) -> Emu;
    /// 设置左上角 y。
    fn set_top(&mut self, emu: Emu);
    /// 宽（EMU）。
    fn width(&self) -> Emu;
    /// 设置宽。
    fn set_width(&mut self, emu: Emu);
    /// 高（EMU）。
    fn height(&self) -> Emu;
    /// 设置高。
    fn set_height(&mut self, emu: Emu);

    /// 旋转角度（度数，正向顺时针）。
    ///
    /// 与 OOXML 规范的"60000 分之一度"通过 `set_rotation` 自动转换；
    /// 读取时返回度数。
    fn rotation(&self) -> f64;
    /// 设置旋转角度。
    fn set_rotation(&mut self, deg: f64);
}
