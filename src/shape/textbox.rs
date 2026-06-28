//! `TextBox`：纯文本框（`<p:sp>` + `prstGeom="rect"` + `txBody`）。
//!
//! 与 [`AutoShape`] 区别在于：[`TextBox`] 强制 `prstGeom=rect`、bodyPr
//! 默认无填充无外框，更接近"传统文本框"语义。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.textbox.TextBox` ←→ [`TextBox`]；
//! - `Slide.shapes.add_textbox(left, top, width, height)` 返回 [`TextBox`]。
//!
//! # 文本替换语义
//!
//! [`TextBox::set_text`] 的语义是"替换全部段落"——每行（`\n` 分隔）变成一段；
//! 旧的字体/颜色属性会被丢弃。如需保留属性，请直接 mutate
//! [`TextBox::text_frame_mut`]。
//!
//! # 示例
//!
//! ```no_run
//! use pptx::shape::TextBox;
//!
//! let mut tb = TextBox::new("MyTextBox");
//! tb.set_text("Hello\nWorld");
//! assert_eq!(tb.text(), "Hello\nWorld");
//! ```

use crate::oxml::shape::Sp as OxmlSp;
use crate::oxml::simpletypes::PresetGeometry;
use crate::oxml::txbody::{Paragraph, TextBody};
use crate::shape::autoshape::AutoShape;
use crate::shape::base::Shape;
use crate::units::Emu;

/// 文本框（python-pptx 的 `shapes.add_textbox` 对应）。
#[derive(Clone, Debug, Default)]
pub struct TextBox {
    /// 内部包装的 [`AutoShape`]。
    pub(crate) shape: AutoShape,
}

impl TextBox {
    /// 新建一个文本框。
    pub fn new(name: impl Into<String>) -> Self {
        let mut s = AutoShape::new(name, PresetGeometry::Rectangle);
        // 默认 textbox 的 bodyPr 是无填充、无外框
        s.sp_mut().text = TextBody::new();
        TextBox { shape: s }
    }

    /// 从 oxml [`OxmlSp`] 构造。
    pub fn from_sp(sp: OxmlSp) -> Self {
        TextBox {
            shape: AutoShape::from_sp(sp),
        }
    }

    /// 文本帧不可变引用。
    pub fn text_frame(&self) -> &TextBody {
        &self.shape.sp.text
    }
    /// 文本帧可变引用。
    pub fn text_frame_mut(&mut self) -> &mut TextBody {
        &mut self.shape.sp_mut().text
    }

    /// 替换全部文本（每行一个段落）。
    ///
    /// 替换后旧的 Run 属性（字体、颜色、加粗）会被清除；如需保留请直接操作
    /// [`TextBox::text_frame_mut`]。
    pub fn set_text(&mut self, text: &str) -> &mut TextBody {
        let tb = self.text_frame_mut();
        tb.paragraphs.clear();
        for line in text.split('\n') {
            let mut p = Paragraph::new();
            p.runs.push(crate::oxml::txbody::Run::new(line));
            tb.paragraphs.push(p);
        }
        tb
    }

    /// 取文本（把全部段落拼起来，行间 `\n`）。
    pub fn text(&self) -> String {
        let mut out = String::new();
        let mut first = true;
        for p in &self.text_frame().paragraphs {
            if !first {
                out.push('\n');
            }
            first = false;
            for r in &p.runs {
                out.push_str(&r.text);
            }
        }
        out
    }

    /// 单词换行便捷方法。
    pub fn set_word_wrap(&mut self, v: bool) {
        self.text_frame_mut().set_word_wrap(v);
    }

    /// 自动调整便捷方法。
    pub fn set_auto_size(&mut self, v: crate::oxml::simpletypes::MsoAutoSize) {
        self.text_frame_mut().set_auto_size(v);
    }

    /// 垂直对齐便捷方法。
    pub fn set_vertical_anchor(&mut self, v: crate::oxml::simpletypes::MsoAnchor) {
        self.text_frame_mut().set_vertical_anchor(v);
    }

    // --------------------- 委派 AutoShape 的效果/锁定/样式 API ---------------------
    //
    // `TextBox` 内部就是 `AutoShape` + `prstGeom=rect` + `txBox=1`，因此所有
    // 形状效果、锁定、主题样式相关的高阶 API 都直接转发到内部 AutoShape。
    // 这样用户对 `TextBox` 也能用 `set_outer_shadow` / `lock_select` / `set_style`
    // 等便捷方法，与 `AutoShape` 行为完全一致。

    /// 设置外阴影（`<a:outerShdw>`）。详见 [`crate::shape::AutoShape::set_outer_shadow`]。
    pub fn set_outer_shadow(&mut self, shadow: crate::oxml::sppr::ShadowEffect) {
        self.shape.set_outer_shadow(shadow);
    }

    /// 设置内阴影（`<a:innerShdw>`）。详见 [`crate::shape::AutoShape::set_inner_shadow`]。
    pub fn set_inner_shadow(&mut self, shadow: crate::oxml::sppr::ShadowEffect) {
        self.shape.set_inner_shadow(shadow);
    }

    /// 设置发光（`<a:glow>`）。详见 [`crate::shape::AutoShape::set_glow`]。
    pub fn set_glow(&mut self, glow: crate::oxml::sppr::GlowEffect) {
        self.shape.set_glow(glow);
    }

    /// 设置柔化边缘（`<a:softEdge>`）。详见 [`crate::shape::AutoShape::set_soft_edge`]。
    pub fn set_soft_edge(&mut self, rad: i64) {
        self.shape.set_soft_edge(rad);
    }

    /// 设置反射（`<a:reflection>`）。详见 [`crate::shape::AutoShape::set_reflection`]。
    pub fn set_reflection(&mut self, reflection: crate::oxml::sppr::ReflectionEffect) {
        self.shape.set_reflection(reflection);
    }

    /// 清除所有效果。详见 [`crate::shape::AutoShape::clear_effects`]。
    pub fn clear_effects(&mut self) {
        self.shape.clear_effects();
    }

    /// 设置三维旋转。详见 [`crate::shape::AutoShape::set_3d_rotation`]。
    pub fn set_3d_rotation(&mut self, lat_deg: f64, lon_deg: f64, rev_deg: f64) {
        self.shape.set_3d_rotation(lat_deg, lon_deg, rev_deg);
    }

    /// 设置三维拉伸。详见 [`crate::shape::AutoShape::set_3d_extrusion`]。
    pub fn set_3d_extrusion(&mut self, height_emu: i32, color: Option<crate::oxml::color::Color>) {
        self.shape.set_3d_extrusion(height_emu, color);
    }

    /// 设置三维棱台。详见 [`crate::shape::AutoShape::set_3d_bevel`]。
    pub fn set_3d_bevel(&mut self, top_w: i32, top_h: i32, bottom_w: i32, bottom_h: i32) {
        self.shape.set_3d_bevel(top_w, top_h, bottom_w, bottom_h);
    }

    /// 设置三维材质预设。详见 [`crate::shape::AutoShape::set_3d_material`]。
    pub fn set_3d_material(&mut self, material: crate::oxml::sppr::MaterialPreset) {
        self.shape.set_3d_material(material);
    }

    /// 清除所有三维效果。详见 [`crate::shape::AutoShape::clear_3d`]。
    pub fn clear_3d(&mut self) {
        self.shape.clear_3d();
    }

    /// 取三维场景引用。详见 [`crate::shape::AutoShape::scene_3d`]。
    pub fn scene_3d(&self) -> Option<&crate::oxml::sppr::Scene3d> {
        self.shape.scene_3d()
    }

    /// 取三维场景可变引用。详见 [`crate::shape::AutoShape::scene_3d_mut`]。
    pub fn scene_3d_mut(&mut self) -> &mut Option<crate::oxml::sppr::Scene3d> {
        self.shape.scene_3d_mut()
    }

    /// 取形状三维属性引用。详见 [`crate::shape::AutoShape::sp_3d`]。
    pub fn sp_3d(&self) -> Option<&crate::oxml::sppr::Sp3d> {
        self.shape.sp_3d()
    }

    /// 取形状三维属性可变引用。详见 [`crate::shape::AutoShape::sp_3d_mut`]。
    pub fn sp_3d_mut(&mut self) -> &mut Option<crate::oxml::sppr::Sp3d> {
        self.shape.sp_3d_mut()
    }

    /// 读取形状锁定（`<a:spLocks>`）。详见 [`crate::shape::AutoShape::locks`]。
    pub fn locks(&self) -> Option<&crate::oxml::shape::ShapeLocks> {
        self.shape.locks()
    }

    /// 读取形状锁定的可变引用。详见 [`crate::shape::AutoShape::locks_mut`]。
    pub fn locks_mut(&mut self) -> &mut crate::oxml::shape::ShapeLocks {
        self.shape.locks_mut()
    }

    /// 便捷锁定：禁止选中。详见 [`crate::shape::AutoShape::lock_select`]。
    pub fn lock_select(&mut self, locked: bool) {
        self.shape.lock_select(locked);
    }

    /// 便捷锁定：禁止移动。详见 [`crate::shape::AutoShape::lock_move`]。
    pub fn lock_move(&mut self, locked: bool) {
        self.shape.lock_move(locked);
    }

    /// 便捷锁定：禁止缩放。详见 [`crate::shape::AutoShape::lock_resize`]。
    pub fn lock_resize(&mut self, locked: bool) {
        self.shape.lock_resize(locked);
    }

    /// 便捷锁定：禁止旋转。详见 [`crate::shape::AutoShape::lock_rotate`]。
    pub fn lock_rotate(&mut self, locked: bool) {
        self.shape.lock_rotate(locked);
    }

    /// 便捷锁定：禁止组合。详见 [`crate::shape::AutoShape::lock_group`]。
    pub fn lock_group(&mut self, locked: bool) {
        self.shape.lock_group(locked);
    }

    /// 清除所有锁定。详见 [`crate::shape::AutoShape::clear_locks`]。
    pub fn clear_locks(&mut self) {
        self.shape.clear_locks();
    }

    /// 统一锁定入口：按 [`crate::LockType`] 设置指定锁定。
    /// 详见 [`crate::shape::AutoShape::set_lock`]。
    pub fn set_lock(&mut self, lock_type: crate::oxml::shape::LockType, locked: bool) {
        self.shape.set_lock(lock_type, locked);
    }

    /// 读取主题样式引用（`<p:style>`）。详见 [`crate::shape::AutoShape::style`]。
    pub fn style(&self) -> Option<&crate::oxml::shape::ShapeStyle> {
        self.shape.style()
    }

    /// 设置主题样式引用。详见 [`crate::shape::AutoShape::set_style`]。
    pub fn set_style(&mut self, style: crate::oxml::shape::ShapeStyle) {
        self.shape.set_style(style);
    }

    /// 清除主题样式引用。详见 [`crate::shape::AutoShape::clear_style`]。
    pub fn clear_style(&mut self) {
        self.shape.clear_style();
    }
}

impl Shape for TextBox {
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
        "text_box"
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
