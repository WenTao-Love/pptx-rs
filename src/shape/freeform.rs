//! `Freeform` / `FreeformBuilder`：手绘自由形（`<a:custGeom>`）。
//!
//! 第一版只暴露"加路径点 → 关闭"的最小接口，足够画折线/多边形。
//! 写入时把每个点用 `<a:moveTo>` / `<a:lnTo>` 表达，最终通过
//! [`crate::slide::ShapesMut::add_freeform`] 添加到 slide。
//!
//! # 与 python-pptx 的对应
//!
//! - python-pptx **未提供** 等价 API；
//! - 本库把 Freeform 落地为 [`AutoShape`] + `<a:custGeom>`（自定义几何），
//!   `FreeformBuilder::build` 会把累积的路径点序列化为 `<a:pathLst>`。
//!
//! # 设计要点
//!
//! - [`FreeformBuilder`] 是**流式 API**：`move_to` / `line_to` 链式调用；
//! - `build` 时把 builder 转为 [`Freeform`]（不可再改）；
//! - 路径坐标系：所有点的 EMU 坐标直接作为 `<a:pt x="..." y="..."/>` 输出；
//!   `<a:path>` 的 `w` / `h` 属性取路径边界框尺寸，让 PowerPoint 按比例缩放。
//!
//! # 示例
//!
//! ```no_run
//! use pptx::shape::FreeformBuilder;
//! use pptx::EmuExt;
//! use pptx::Inches;
//!
//! let mut builder = FreeformBuilder::new();
//! builder.move_to(Inches(1.0).emu(), Inches(1.0).emu());
//! builder.line_to(Inches(3.0).emu(), Inches(1.0).emu());
//! builder.line_to(Inches(2.0).emu(), Inches(2.0).emu());
//! builder.close();
//! let _f = builder.build("triangle");
//! ```

use crate::oxml::shape::Sp as OxmlSp;
use crate::oxml::sppr::{CustomGeometry, GeomRect, Geometry, Path, PathSegment, ShapeProperties};
use crate::oxml::txbody::TextBody;
use crate::shape::autoshape::AutoShape;
use crate::shape::base::Shape;
use crate::units::Emu;

/// 一个 2D 点。
#[derive(Copy, Clone, Debug, Default)]
pub struct Point {
    /// x 坐标（EMU）。
    pub x: Emu,
    /// y 坐标（EMU）。
    pub y: Emu,
}

/// `FreeformBuilder`：累积路径点。
#[derive(Clone, Debug, Default)]
pub struct FreeformBuilder {
    points: Vec<Point>,
    auto_close: bool,
}

impl FreeformBuilder {
    /// 新建。
    pub fn new() -> Self {
        FreeformBuilder {
            points: Vec::new(),
            auto_close: false,
        }
    }

    /// 在 `(x, y)` 处下笔（相当于 SVG `M`）。
    pub fn move_to(&mut self, x: Emu, y: Emu) -> &mut Self {
        self.points.push(Point { x, y });
        self
    }

    /// 连接到 `(x, y)`（相当于 SVG `L`）。
    pub fn line_to(&mut self, x: Emu, y: Emu) -> &mut Self {
        self.points.push(Point { x, y });
        self
    }

    /// 闭合路径（首尾相连，相当于 SVG `Z`）。
    pub fn close(&mut self) -> &mut Self {
        self.auto_close = true;
        self
    }

    /// 取所有点。
    pub fn points(&self) -> &[Point] {
        &self.points
    }

    /// 构造一个 [`Freeform`]`（一个用 `<a:custGeom>` 描述的 AutoShape）。
    ///
    /// # 行为
    ///
    /// 1. 把累积的点序列化为 `<a:pathLst>`，首点为 `<a:moveTo>`，后续点为 `<a:lnTo>`；
    /// 2. 若调用过 [`Self::close`]，追加 `<a:close/>` 段；
    /// 3. `<a:path>` 的 `w` / `h` 取路径边界框尺寸（最小 x/y 到最大 x/y）；
    /// 4. `spPr.xfrm` 的位置默认为路径最小点，尺寸为边界框尺寸——调用方可后续
    ///    通过 `set_left/set_top/set_width/set_height` 覆盖。
    ///
    /// # 边界情况
    ///
    /// - 点数 < 2：仍会构造一个空路径，PowerPoint 会按形状的 xfrm 尺寸渲染空白区域；
    /// - 单点（仅 `move_to`）：等价于空路径，几何上不可见。
    #[allow(clippy::field_reassign_with_default)]
    pub fn build(self, name: impl Into<String>) -> Freeform {
        let mut sp = OxmlSp::default();
        sp.id = 0;
        sp.name = name.into();
        sp.properties = ShapeProperties::default();
        sp.text = TextBody::new();

        // 把点序列化为路径段：首点 moveTo，后续点 lnTo，最后按需 close。
        let mut segments: Vec<PathSegment> = Vec::with_capacity(self.points.len() + 1);
        for (i, p) in self.points.iter().enumerate() {
            let x = p.x.value();
            let y = p.y.value();
            if i == 0 {
                segments.push(PathSegment::MoveTo { x, y });
            } else {
                segments.push(PathSegment::LineTo { x, y });
            }
        }
        if self.auto_close {
            segments.push(PathSegment::Close);
        }

        // 计算路径边界框，作为 <a:path w="..." h="..."> 的取值。
        // PowerPoint 会按 w/h 比例把 path 缩放到 spPr.xfrm 的尺寸。
        let (path_w, path_h, min_x, min_y) = compute_bbox(&self.points);

        let path = Path {
            width: path_w,
            height: path_h,
            fill: None,
            stroke: None,
            segments,
        };
        let geom = CustomGeometry {
            fill: None,
            stroke: None,
            // 内嵌区域设为边界框，让 PowerPoint 文本布局有合理边界。
            rect: Some(GeomRect {
                l: min_x.to_string(),
                t: min_y.to_string(),
                r: (min_x + path_w).to_string(),
                b: (min_y + path_h).to_string(),
            }),
            path_list: vec![path],
        };
        sp.properties.geometry = Some(Geometry::Custom(geom));

        // 同步设置 xfrm 的位置和尺寸，让形状默认就能显示出路径形状。
        // 调用方仍可通过 set_left/set_top/set_width/set_height 覆盖。
        sp.properties.xfrm.off_x = Some(Emu(min_x));
        sp.properties.xfrm.off_y = Some(Emu(min_y));
        sp.properties.xfrm.ext_cx = Some(Emu(path_w));
        sp.properties.xfrm.ext_cy = Some(Emu(path_h));

        Freeform {
            shape: AutoShape::from_sp(sp),
        }
    }
}

/// 计算点列表的边界框，返回 `(width, height, min_x, min_y)`。
///
/// 空列表返回 `(0, 0, 0, 0)`。
fn compute_bbox(points: &[Point]) -> (i64, i64, i64, i64) {
    if points.is_empty() {
        return (0, 0, 0, 0);
    }
    let mut min_x = points[0].x.value();
    let mut min_y = points[0].y.value();
    let mut max_x = min_x;
    let mut max_y = min_y;
    for p in &points[1..] {
        let x = p.x.value();
        let y = p.y.value();
        if x < min_x {
            min_x = x;
        }
        if x > max_x {
            max_x = x;
        }
        if y < min_y {
            min_y = y;
        }
        if y > max_y {
            max_y = y;
        }
    }
    (max_x - min_x, max_y - min_y, min_x, min_y)
}

/// 自由形。
#[derive(Clone, Debug, Default)]
pub struct Freeform {
    /// 内部包装的 [`AutoShape`]。
    pub(crate) shape: AutoShape,
}

impl Freeform {
    /// 借用内部 AutoShape（便于复用 `set_outer_shadow` / `set_fill` 等高阶 API）。
    pub fn as_shape(&self) -> &AutoShape {
        &self.shape
    }
    /// 借用内部 AutoShape 的可变引用。
    pub fn as_shape_mut(&mut self) -> &mut AutoShape {
        &mut self.shape
    }
}

impl Shape for Freeform {
    fn id(&self) -> u32 {
        self.shape.id()
    }
    fn set_id(&mut self, id: u32) {
        self.shape.set_id(id);
    }
    fn name(&self) -> &str {
        self.shape.name()
    }
    fn set_name(&mut self, name: String) {
        self.shape.set_name(name);
    }
    fn shape_type(&self) -> &'static str {
        "freeform"
    }
    fn left(&self) -> Emu {
        self.shape.left()
    }
    fn set_left(&mut self, emu: Emu) {
        self.shape.set_left(emu);
    }
    fn top(&self) -> Emu {
        self.shape.top()
    }
    fn set_top(&mut self, emu: Emu) {
        self.shape.set_top(emu);
    }
    fn width(&self) -> Emu {
        self.shape.width()
    }
    fn set_width(&mut self, emu: Emu) {
        self.shape.set_width(emu);
    }
    fn height(&self) -> Emu {
        self.shape.height()
    }
    fn set_height(&mut self, emu: Emu) {
        self.shape.set_height(emu);
    }
    fn rotation(&self) -> f64 {
        self.shape.rotation()
    }
    fn set_rotation(&mut self, deg: f64) {
        self.shape.set_rotation(deg);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::writer::XmlWriter;
    use crate::EmuExt;
    use crate::Inches;

    /// `build` 输出 custGeom 而非 prstGeom=rect（TODO-024 回归检测）。
    #[test]
    fn build_outputs_cust_geom() {
        let mut b = FreeformBuilder::new();
        b.move_to(Inches(1.0).emu(), Inches(1.0).emu())
            .line_to(Inches(3.0).emu(), Inches(1.0).emu())
            .line_to(Inches(2.0).emu(), Inches(2.0).emu())
            .close();
        let f = b.build("triangle");
        let mut w = XmlWriter::new();
        f.shape.sp().properties.write_xml(&mut w, "p:spPr");
        let xml = &w.buf;
        assert!(
            xml.contains("<a:custGeom>"),
            "must output custGeom, xml: {}",
            xml
        );
        assert!(
            !xml.contains("<a:prstGeom"),
            "must not output prstGeom, xml: {}",
            xml
        );
        assert!(xml.contains("<a:moveTo>"), "xml: {}", xml);
        assert!(xml.contains("<a:lnTo>"), "xml: {}", xml);
        assert!(xml.contains("<a:close/>"), "xml: {}", xml);
    }

    /// `build` 在空路径时不 panic，仍能写出 custGeom。
    #[test]
    fn build_with_empty_points() {
        let b = FreeformBuilder::new();
        let f = b.build("empty");
        let mut w = XmlWriter::new();
        f.shape.sp().properties.write_xml(&mut w, "p:spPr");
        let xml = &w.buf;
        assert!(xml.contains("<a:custGeom>"), "xml: {}", xml);
    }

    /// 边界框计算正确（两点形成一条对角线）。
    #[test]
    fn bbox_two_points() {
        let pts = [
            Point {
                x: Inches(1.0).emu(),
                y: Inches(2.0).emu(),
            },
            Point {
                x: Inches(4.0).emu(),
                y: Inches(6.0).emu(),
            },
        ];
        let (w, h, min_x, min_y) = compute_bbox(&pts);
        assert_eq!(w, Inches(3.0).emu().value());
        assert_eq!(h, Inches(4.0).emu().value());
        assert_eq!(min_x, Inches(1.0).emu().value());
        assert_eq!(min_y, Inches(2.0).emu().value());
    }
}
