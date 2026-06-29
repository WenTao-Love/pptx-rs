//! `AutoShape`：自选图形（矩形、椭圆、箭头、…）。
//!
//! [`AutoShape`] 是本库"几何 + 文本"的通用形状——文本框、矩形、椭圆、箭头
//! 等在 python-pptx 中也都被归为 `Shape`（区别于 `TextBox` 是早期 OOXML 的
//! 历史残留）。本库把 `TextBox` 单独拆出，主要为了类型可读性。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.autoshape.Shape` ←→ [`AutoShape`]；
//! - `pptx.shapes.textbox.TextBox` ←→ [`crate::shape::TextBox`]。
//!
//! # 几何与位置
//!
//! - 几何由 [`PresetGeometry`] 决定（`rect` / `ellipse` / `arrow` / ...）；
//! - 位置 / 尺寸 / 旋转在 [`crate::oxml::sppr::Transform`] 上；
//! - 文本由 [`TextBody`] 承载，可含多段多 Run。
//!
//! # 示例
//!
//! ```no_run
//! use pptx_rs::shape::{AutoShape, Shape};  // 引入 Shape trait 以使用 set_* 方法
//! use pptx_rs::oxml::simpletypes::PresetGeometry;
//! use pptx_rs::EmuExt;
//! use pptx_rs::Inches;
//!
//! let mut s = AutoShape::new("MyRect", PresetGeometry::Rectangle);
//! s.set_left(Inches(1.0).emu());
//! s.set_top(Inches(1.0).emu());
//! s.set_width(Inches(3.0).emu());
//! s.set_height(Inches(2.0).emu());
//! let _ = s.text_frame_mut();
//! ```

use crate::oxml::color::Color;
use crate::oxml::shape::Sp as OxmlSp;
use crate::oxml::simpletypes::PresetGeometry;
use crate::oxml::sppr::{Fill, Geometry, ShapeProperties};
use crate::oxml::txbody::TextBody;
use crate::shape::base::Shape;
use crate::units::Emu;

/// 自选形状（带或不带文本都可以）。
#[derive(Clone, Debug, Default)]
pub struct AutoShape {
    /// 内部 oxml 句柄（`pub(crate)` 供 `slide::ShapesMut` mutate）。
    pub(crate) sp: OxmlSp,
}

impl AutoShape {
    /// 构造一个指定几何形状（不设置位置/尺寸，由调用方决定）。
    #[allow(clippy::field_reassign_with_default)]
    pub fn new(name: impl Into<String>, geometry: PresetGeometry) -> Self {
        let mut sp = OxmlSp::default();
        sp.id = 0;
        sp.name = name.into();
        sp.properties.geometry = Some(Geometry::preset(geometry));
        sp.text = TextBody::new();
        AutoShape { sp }
    }

    /// 从 oxml [`OxmlSp`] 构造包装。
    pub fn from_sp(sp: OxmlSp) -> Self {
        AutoShape { sp }
    }

    /// 取出内部 oxml 引用。
    pub fn sp(&self) -> &OxmlSp {
        &self.sp
    }
    /// 取出内部 oxml 可变引用。
    pub fn sp_mut(&mut self) -> &mut OxmlSp {
        &mut self.sp
    }

    /// 文本体不可变引用。
    pub fn text_frame(&self) -> &TextBody {
        &self.sp.text
    }
    /// 文本体可变引用。
    pub fn text_frame_mut(&mut self) -> &mut TextBody {
        &mut self.sp.text
    }

    /// 形状属性不可变引用。
    pub fn properties(&self) -> &ShapeProperties {
        &self.sp.properties
    }
    /// 形状属性可变引用。
    pub fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.sp.properties
    }

    /// 设填充。
    pub fn set_fill(&mut self, fill: Fill) {
        self.sp.properties.fill = fill;
    }

    /// 便捷方法：设填充为指定 RGB 颜色。
    ///
    /// 对应 python-pptx 中
    /// `shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(r, g, b)`。
    pub fn set_fill_color(&mut self, c: impl Into<Color>) {
        self.sp.properties.fill = Fill::Solid(c.into());
    }

    /// 便捷方法：设描边颜色。
    pub fn set_stroke_color(&mut self, c: impl Into<Color>) {
        let mut line = self.sp.properties.line.clone().unwrap_or_default();
        line.color = c.into();
        self.sp.properties.line = Some(line);
    }

    /// 便捷方法：设描边宽度（EMU）。通常 `Pt(1.0).emu()`。
    pub fn set_stroke_width(&mut self, w: Emu) {
        let mut line = self.sp.properties.line.clone().unwrap_or_default();
        line.width = Some(w);
        self.sp.properties.line = Some(line);
    }

    // --------------------- 形状效果 API（TODO-011 高阶） ---------------------

    /// 设置外阴影（`<a:outerShdw>`）。覆盖既有外阴影，保留其他效果。
    ///
    /// 对标 python-pptx `shape.shadow.inherit = False` + `shape.shadow.outerShadow`。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx_rs::{Presentation, ShadowEffect, RGBColor, Color, Pt};
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// # let mut sh = s.shapes_mut().add_textbox(Pt(10.0), Pt(10.0), Pt(100.0), Pt(20.0)).unwrap();
    /// sh.set_outer_shadow(ShadowEffect {
    ///     dir: 2_700_000,        // 向下（1/60000 度）
    ///     dist: 38_100,          // 3pt（EMU）
    ///     blur_rad: 38_100,
    ///     color: Color::RGB(RGBColor(0x80, 0x80, 0x80)),
    ///     rot_with_shape: None,
    /// });
    /// ```
    pub fn set_outer_shadow(&mut self, shadow: crate::oxml::sppr::ShadowEffect) {
        self.sp.properties.set_outer_shadow(shadow);
    }

    /// 设置内阴影（`<a:innerShdw>`）。
    pub fn set_inner_shadow(&mut self, shadow: crate::oxml::sppr::ShadowEffect) {
        self.sp.properties.set_inner_shadow(shadow);
    }

    /// 设置发光（`<a:glow>`）。
    pub fn set_glow(&mut self, glow: crate::oxml::sppr::GlowEffect) {
        self.sp.properties.set_glow(glow);
    }

    /// 设置柔化边缘（`<a:softEdge>`）。
    pub fn set_soft_edge(&mut self, rad: i64) {
        self.sp.properties.set_soft_edge(rad);
    }

    /// 设置反射（`<a:reflection>`）。
    pub fn set_reflection(&mut self, reflection: crate::oxml::sppr::ReflectionEffect) {
        self.sp.properties.set_reflection(reflection);
    }

    /// 清除所有效果（删除整个 `<a:effectLst>` 元素）。
    pub fn clear_effects(&mut self) {
        self.sp.properties.clear_effects();
    }

    // --------------------- 三维效果 API（TODO-050 高阶） ---------------------
    //
    // 对标 PowerPoint "形状效果 → 三维旋转 / 三维格式"。
    // oxml 层 Scene3d/Sp3d 模型已在 sppr.rs 完成，这里仅暴露便捷方法。
    // 角度参数统一使用"度"（用户直觉），内部转换为 OOXML 的 1/60000 度。

    /// 设置三维旋转（`<a:rot lat="..." lon="..." rev="..."/>`）。
    ///
    /// 三个角度均以**度**为单位（用户直觉），内部转换为 1/60000 度（OOXML ST_Angle）。
    /// 若 `scene3d` 不存在则自动创建默认 Scene3d（正交前视图 + 平衡光）。
    ///
    /// # 参数
    /// - `lat_deg`：纬度（-90°~90°，正值抬头）
    /// - `lon_deg`：经度（0°~360°，正值右转）
    /// - `rev_deg`：滚转（沿视线轴的旋转，正值顺时针）
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx_rs::{Presentation, Pt};
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// # let mut sh = s.shapes_mut().add_textbox(Pt(10.0), Pt(10.0), Pt(100.0), Pt(20.0)).unwrap();
    /// sh.set_3d_rotation(30.0, 45.0, 0.0); // 抬头 30° + 右转 45°
    /// ```
    pub fn set_3d_rotation(&mut self, lat_deg: f64, lon_deg: f64, rev_deg: f64) {
        let scene = self
            .sp
            .properties
            .scene3d
            .get_or_insert_with(crate::oxml::sppr::Scene3d::default);
        scene.camera.rotation = Some(crate::oxml::sppr::Rotation3d {
            lat: (lat_deg * 60_000.0) as i32,
            lon: (lon_deg * 60_000.0) as i32,
            rev: (rev_deg * 60_000.0) as i32,
        });
    }

    /// 设置三维拉伸高度与可选拉伸颜色（`<a:sp3d extrusionH="...">` + `<a:extrusionClr>`）。
    ///
    /// # 参数
    /// - `height_emu`：拉伸高度（EMU，38100 = 3pt）
    /// - `color`：拉伸颜色（`None` 表示不写出 `<a:extrusionClr>`）
    pub fn set_3d_extrusion(&mut self, height_emu: i32, color: Option<crate::oxml::color::Color>) {
        let sp3d = self
            .sp
            .properties
            .sp3d
            .get_or_insert_with(crate::oxml::sppr::Sp3d::default);
        sp3d.extrusion_h = height_emu;
        sp3d.extrusion_color = color;
    }

    /// 设置三维棱台（`<a:bevelT>` + `<a:bevelB>`）。
    ///
    /// # 参数
    /// - `top_w` / `top_h`：顶部棱台的宽高（EMU，63500 = 5pt）
    /// - `bottom_w` / `bottom_h`：底部棱台的宽高（EMU）
    pub fn set_3d_bevel(&mut self, top_w: i32, top_h: i32, bottom_w: i32, bottom_h: i32) {
        let sp3d = self
            .sp
            .properties
            .sp3d
            .get_or_insert_with(crate::oxml::sppr::Sp3d::default);
        sp3d.bevel_top = Some(crate::oxml::sppr::Bevel { w: top_w, h: top_h });
        sp3d.bevel_bottom = Some(crate::oxml::sppr::Bevel {
            w: bottom_w,
            h: bottom_h,
        });
    }

    /// 设置三维材质预设（`<a:sp3d prstMaterial="...">`）。
    pub fn set_3d_material(&mut self, material: crate::oxml::sppr::MaterialPreset) {
        let sp3d = self
            .sp
            .properties
            .sp3d
            .get_or_insert_with(crate::oxml::sppr::Sp3d::default);
        sp3d.prst_material = material;
    }

    /// 清除所有三维效果（删除 `scene3d` + `sp3d`）。
    pub fn clear_3d(&mut self) {
        self.sp.properties.scene3d = None;
        self.sp.properties.sp3d = None;
    }

    /// 取三维场景引用（camera + lightRig）。`None` 表示未设置 3D 场景。
    pub fn scene_3d(&self) -> Option<&crate::oxml::sppr::Scene3d> {
        self.sp.properties.scene3d.as_ref()
    }

    /// 取三维场景可变引用（可直接操作 camera / lightRig）。
    pub fn scene_3d_mut(&mut self) -> &mut Option<crate::oxml::sppr::Scene3d> {
        &mut self.sp.properties.scene3d
    }

    /// 取形状三维属性引用（拉伸/棱台/材质）。
    pub fn sp_3d(&self) -> Option<&crate::oxml::sppr::Sp3d> {
        self.sp.properties.sp3d.as_ref()
    }

    /// 取形状三维属性可变引用。
    pub fn sp_3d_mut(&mut self) -> &mut Option<crate::oxml::sppr::Sp3d> {
        &mut self.sp.properties.sp3d
    }

    // --------------------- 形状锁定 API（TODO-027 高阶） ---------------------

    /// 读取形状锁定（`<a:spLocks>`）。`None` 表示未设置。
    pub fn locks(&self) -> Option<&crate::oxml::shape::ShapeLocks> {
        self.sp.locks.as_ref()
    }

    /// 读取形状锁定的可变引用。若未设置，自动初始化为空 `ShapeLocks`（无任何锁定）。
    pub fn locks_mut(&mut self) -> &mut crate::oxml::shape::ShapeLocks {
        self.sp.locks.get_or_insert_with(Default::default)
    }

    /// 便捷锁定：禁止选中（`noSelect="1"`）。
    pub fn lock_select(&mut self, locked: bool) {
        self.locks_mut().no_select = locked;
    }

    /// 便捷锁定：禁止移动（`noMove="1"`）。
    pub fn lock_move(&mut self, locked: bool) {
        self.locks_mut().no_move = locked;
    }

    /// 便捷锁定：禁止缩放（`noResize="1"`）。
    pub fn lock_resize(&mut self, locked: bool) {
        self.locks_mut().no_resize = locked;
    }

    /// 便捷锁定：禁止旋转（`noRot="1"`）。
    pub fn lock_rotate(&mut self, locked: bool) {
        self.locks_mut().no_rot = locked;
    }

    /// 便捷锁定：禁止组合（`noGrp="1"`）。
    pub fn lock_group(&mut self, locked: bool) {
        self.locks_mut().no_grp = locked;
    }

    /// 清除所有锁定。
    pub fn clear_locks(&mut self) {
        self.sp.locks = None;
    }

    /// 统一锁定入口：按 [`crate::oxml::shape::LockType`] 设置指定锁定（TODO-027 高阶 API）。
    ///
    /// 对标 python-pptx 风格 `shape.set_lock(MSO_SHAPE_LOCK_TYPE.Select, True)`。
    /// 比 `lock_select` / `lock_move` 等具名方法更通用，可覆盖所有 12 种锁定类型。
    ///
    /// # 参数
    /// - `lock_type`：锁定类型枚举；
    /// - `locked`：`true` 启用，`false` 解除。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx_rs::{Presentation, Pt, LockType};
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// # let mut sh = s.shapes_mut().add_textbox(Pt(10.0), Pt(10.0), Pt(100.0), Pt(20.0)).unwrap();
    /// sh.set_lock(LockType::Select, true);
    /// sh.set_lock(LockType::ChangeAspect, true);
    /// assert!(sh.locks().unwrap().get_lock(LockType::Select));
    /// ```
    pub fn set_lock(&mut self, lock_type: crate::oxml::shape::LockType, locked: bool) {
        self.locks_mut().set_lock(lock_type, locked);
    }

    // --------------------- ShapeStyle API（TODO-006 高阶） ---------------------

    /// 读取主题样式引用（`<p:style>`）。`None` 表示未设置。
    pub fn style(&self) -> Option<&crate::oxml::shape::ShapeStyle> {
        self.sp.style.as_ref()
    }

    /// 设置主题样式引用（`<p:style>`）。
    ///
    /// 对标 python-pptx `shape.style` 的底层元素引用方式。
    ///
    /// # 参数
    /// - `style`：主题样式（line_ref / fill_ref / effect_ref / font_ref）
    pub fn set_style(&mut self, style: crate::oxml::shape::ShapeStyle) {
        self.sp.style = Some(style);
    }

    /// 清除主题样式引用。
    pub fn clear_style(&mut self) {
        self.sp.style = None;
    }

    /// 便捷方法：设文本（**单段单 Run**）。多段请用 `text_frame_mut().add_paragraph`。
    pub fn set_text(&mut self, t: impl Into<String>) {
        let s: String = t.into();
        self.sp.text.set_text(&s);
    }

    /// 文本快捷访问：返回首段首 Run 的文本（多段时只反映首段）。
    pub fn text(&self) -> String {
        self.sp.text.text()
    }

    // --------------------- 调整手柄 API（TODO-038） ---------------------

    /// 取调整值列表（不可变）。
    ///
    /// 对标 python-pptx `shape.adjustments`。
    pub fn adjustments(&self) -> &[crate::oxml::sppr::AdjustmentValue] {
        match &self.sp.properties.geometry {
            Some(Geometry::Preset(_, adj)) => adj,
            _ => &[],
        }
    }

    /// 取调整值列表（可变）。
    ///
    /// 如果几何不是 `Preset` 变体，返回空 Vec 的可变引用（实际是临时变量）。
    /// 调用方应确保几何是 `Preset` 类型后再调用。
    pub fn adjustments_mut(&mut self) -> &mut Vec<crate::oxml::sppr::AdjustmentValue> {
        // 确保 geometry 是 Preset 变体
        if !matches!(self.sp.properties.geometry, Some(Geometry::Preset(..))) {
            self.sp.properties.geometry = Some(Geometry::default());
        }
        match &mut self.sp.properties.geometry {
            Some(Geometry::Preset(_, adj)) => adj,
            _ => unreachable!(),
        }
    }

    /// 设置指定索引的调整值（归一化 0.0-1.0）。
    ///
    /// 对标 python-pptx `shape.adjustments[0] = 0.15`。
    ///
    /// # 参数
    /// - `idx`：调整值索引（0-based）；
    /// - `value`：归一化值（0.0-1.0）。
    ///
    /// # 行为
    /// - 若 `idx` 超出现有列表长度，自动追加新条目；
    /// - 新条目的 `name` 默认为 `"adj"` （首个）或 `"adjN"`（后续）。
    pub fn set_adjustment(&mut self, idx: usize, value: f64) {
        let adj = self.adjustments_mut();
        while adj.len() <= idx {
            let name = if adj.is_empty() {
                "adj".to_string()
            } else {
                format!("adj{}", adj.len() + 1)
            };
            adj.push(crate::oxml::sppr::AdjustmentValue::from_normalized(
                &name, 0.0,
            ));
        }
        adj[idx].raw_value = (value * 100000.0).round() as i64;
    }

    /// 取指定索引的调整值归一化值（0.0-1.0）。
    ///
    /// 对标 python-pptx `shape.adjustments[0]`。
    pub fn adjustment_value(&self, idx: usize) -> Option<f64> {
        self.adjustments().get(idx).map(|a| a.effective_value())
    }
}

impl Shape for AutoShape {
    fn id(&self) -> u32 {
        self.sp.id
    }
    fn set_id(&mut self, id: u32) {
        self.sp.id = id;
    }
    fn name(&self) -> &str {
        &self.sp.name
    }
    fn set_name(&mut self, name: String) {
        self.sp.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "auto_shape"
    }

    fn left(&self) -> Emu {
        self.sp.properties.xfrm.off_x.unwrap_or_default()
    }
    fn set_left(&mut self, emu: Emu) {
        self.sp.properties.xfrm.off_x = Some(emu);
    }
    fn top(&self) -> Emu {
        self.sp.properties.xfrm.off_y.unwrap_or_default()
    }
    fn set_top(&mut self, emu: Emu) {
        self.sp.properties.xfrm.off_y = Some(emu);
    }
    fn width(&self) -> Emu {
        self.sp.properties.xfrm.ext_cx.unwrap_or_default()
    }
    fn set_width(&mut self, emu: Emu) {
        self.sp.properties.xfrm.ext_cx = Some(emu);
    }
    fn height(&self) -> Emu {
        self.sp.properties.xfrm.ext_cy.unwrap_or_default()
    }
    fn set_height(&mut self, emu: Emu) {
        self.sp.properties.xfrm.ext_cy = Some(emu);
    }

    /// 旋转角度（度数，正向顺时针）。
    fn rotation(&self) -> f64 {
        self.sp.properties.rot_deg.unwrap_or(0.0)
    }
    /// 设置旋转角度。OOXML 内部单位是"60000 分之一度"，
    /// 本方法在读写时自动转换。
    fn set_rotation(&mut self, deg: f64) {
        self.sp.properties.rot_deg = Some(deg);
        // OOXML 中 rot 单位是 1/60000 度
        let rot = (deg * 60_000.0) as i32;
        self.sp.properties.xfrm.rot = Some(rot);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::writer::XmlWriter;

    /// `set_adjustment` 正确设置首个调整值（TODO-038）。
    #[test]
    fn set_adjustment_first_value() {
        let mut s = AutoShape::new("RoundRect", PresetGeometry::RoundRectangle);
        s.set_adjustment(0, 0.25);
        assert_eq!(s.adjustments().len(), 1);
        assert!((s.adjustment_value(0).unwrap() - 0.25).abs() < 1e-6);
    }

    /// `set_adjustment` 自动追加多个调整值（TODO-038）。
    #[test]
    fn set_adjustment_multiple_values() {
        let mut s = AutoShape::new("Shape", PresetGeometry::Rectangle);
        s.set_adjustment(0, 0.5);
        s.set_adjustment(1, 0.3);
        assert_eq!(s.adjustments().len(), 2);
        assert!((s.adjustment_value(0).unwrap() - 0.5).abs() < 1e-6);
        assert!((s.adjustment_value(1).unwrap() - 0.3).abs() < 1e-6);
    }

    /// `adjustment_value` 对越界索引返回 `None`（TODO-038）。
    #[test]
    fn adjustment_value_out_of_bounds_returns_none() {
        let s = AutoShape::new("Rect", PresetGeometry::Rectangle);
        assert!(s.adjustment_value(0).is_none());
    }

    /// 调整值在序列化后正确写出 `<a:gd>` 元素（TODO-038）。
    #[test]
    fn adjustment_serializes_to_xml() {
        let mut s = AutoShape::new("RoundRect", PresetGeometry::RoundRectangle);
        s.set_adjustment(0, 0.16667);
        let mut w = XmlWriter::new();
        s.sp.properties.geometry.as_ref().unwrap().write_xml(&mut w);
        let xml = &w.buf;
        assert!(
            xml.contains("<a:prstGeom prst=\"roundRect\">"),
            "xml: {}",
            xml
        );
        assert!(
            xml.contains("<a:gd name=\"adj\" fmla=\"val 16667\"/>"),
            "xml: {}",
            xml
        );
    }
}
