//! `OleObjectShape`：高阶 OLE 对象形状（TODO-043）。
//!
//! OLE 对象在 OOXML 中通过 `<p:graphicFrame>` + `<a:graphicData uri=".../ole">`
//! + `<p:oleObj r:id="..."/>` 引用一个独立的 `/ppt/embeddings/oleObjectN.bin` part。
//!   本高阶 API 把 graphicFrame 包装为 [`OleObjectShape`]，提供 rid / image_rid /
//!   prog_id / name / show_as_icon 的便捷访问，并让 [`Shape`] trait 直接可用。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.parts.oleobject.OleObjectPart` ←→ [`crate::oxml::ole::OleObject`]（数据模型）；
//! - `pptx.shapes.graphfrm.GraphicFrame` ←→ 本 [`OleObjectShape`]（承载位置/尺寸 + 引用）。
//!
//! # 写出语义
//!
//! - `OleObjectShape` 序列化时只写出 `<p:graphicFrame>` + 引用元素 `<p:oleObj r:id="..."/>`；
//! - 真正的 OLE 二进制数据由 [`crate::presentation::Presentation::save`] 在
//!   `to_opc_package` 中遍历每张 slide 的 `ole_entries` 写出独立的 `oleObjectN.bin` part；
//! - slide 的 `_rels/slideN.xml.rels` 中会添加 `oleObject` 关系指向该 part。
//!
//! # 限制
//!
//! - 当前仅支持嵌入（`<p:embed/>`），不支持链接（`<p:link/>`）；
//! - 图标图片可选（`image_rid` 为空时 PowerPoint 用默认图标）；
//! - 读取已有 OLE 对象的 graphicFrame 时，**仅保留外壳**，不解析 `<p:oleObj>` 内容。

use crate::oxml::ole::OleObject as OxmlOleObject;
use crate::oxml::shape::{Graphic as OxmlGraphic, GraphicFrame as OxmlFrame};
use crate::shape::base::Shape;
use crate::units::Emu;

/// 高阶 OLE 对象形状（承载 `<p:graphicFrame>` + `<p:oleObj r:id="..."/>` 引用）。
///
/// 通过 [`OleObjectShape::ole`] / [`OleObjectShape::ole_mut`] 访问 OLE 数据模型；
/// 通过 [`Shape`] trait 方法（`left` / `top` / `width` / `height`）调整位置与尺寸。
///
/// # 内部不变量
///
/// `frame.graphic` 始终保持为 `Graphic::OleObject(_)`。本类型所有便捷方法
/// （`rid` / `set_rid` / `image_rid` / `prog_id` / `name` / `show_as_icon`）
/// 在不变量被破坏时**静默忽略**或返回默认值，绝不 panic——
/// 这与库整体"零 panic"约定一致（参见 `.trae/rules/project_rules.md` §5）。
#[derive(Clone, Debug, Default)]
pub struct OleObjectShape {
    /// 内部 oxml 句柄（`GraphicFrame`，承载 `Graphic::OleObject`）。
    pub(crate) frame: OxmlFrame,
}

impl OleObjectShape {
    /// 构造一个指定 progId 与显示名的 OLE 对象形状（rid/image_rid 留空，由 presentation 层填充）。
    ///
    /// # 参数
    /// - `prog_id`：OLE 程序标识符（如 `"Excel.Sheet.12"` / `"Package"`）。
    /// - `name`：显示名（如 `"Worksheet"`）。
    pub fn new(prog_id: impl Into<String>, name: impl Into<String>) -> Self {
        let ole = OxmlOleObject::new(prog_id, name);
        let frame = OxmlFrame {
            graphic: OxmlGraphic::OleObject(ole),
            ..Default::default()
        };
        OleObjectShape { frame }
    }

    /// 从 oxml Frame 构造（通常用于读取已有 OLE 对象时）。
    pub fn from_frame(frame: OxmlFrame) -> Self {
        OleObjectShape { frame }
    }

    /// 取内部 oxml OleObject 引用。
    ///
    /// 返回 `None` 仅在内部不变量被外部错误破坏时（`frame.graphic` 不是 `OleObject` 变体）。
    pub fn ole(&self) -> Option<&OxmlOleObject> {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => Some(o),
            _ => None,
        }
    }

    /// 取内部 oxml OleObject 可变引用。
    pub fn ole_mut(&mut self) -> Option<&mut OxmlOleObject> {
        match &mut self.frame.graphic {
            OxmlGraphic::OleObject(o) => Some(o),
            _ => None,
        }
    }

    /// 当前关联的 OLE 关系 id（指向 `/ppt/embeddings/oleObjectN.bin`）。
    ///
    /// 新建时为空字符串；由 `Presentation::to_opc_package` 在写出 ole part 时填充。
    /// 不变量被破坏时返回空字符串。
    pub fn rid(&self) -> &str {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => &o.rid,
            _ => "",
        }
    }

    /// 设置 OLE 关系 id（一般由 presentation 层自动调用，用户无需直接设置）。
    /// 不变量被破坏时静默忽略。
    pub fn set_rid(&mut self, rid: impl Into<String>) {
        if let Some(o) = self.ole_mut() {
            o.rid = rid.into();
        }
    }

    /// 当前关联的图标图片关系 id（指向 `/ppt/media/imageN.{ext}`）。
    ///
    /// 空字符串表示无图标图片，PowerPoint 会用默认图标显示。
    /// 不变量被破坏时返回空字符串。
    pub fn image_rid(&self) -> &str {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => &o.image_rid,
            _ => "",
        }
    }

    /// 设置图标图片关系 id。不变量被破坏时静默忽略。
    pub fn set_image_rid(&mut self, rid: impl Into<String>) {
        if let Some(o) = self.ole_mut() {
            o.image_rid = rid.into();
        }
    }

    /// OLE 程序标识符（如 `"Excel.Sheet.12"`）。
    ///
    /// 不变量被破坏时返回 `"Package"` 作为兜底。
    pub fn prog_id(&self) -> &str {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => &o.prog_id,
            _ => "Package",
        }
    }

    /// 修改 OLE 程序标识符。不变量被破坏时静默忽略。
    pub fn set_prog_id(&mut self, prog_id: impl Into<String>) {
        if let Some(o) = self.ole_mut() {
            o.prog_id = prog_id.into();
        }
    }

    /// 显示名（如 `"Worksheet"` / `"Document"`）。
    ///
    /// 不变量被破坏时返回空字符串。
    pub fn ole_name(&self) -> &str {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => &o.name,
            _ => "",
        }
    }

    /// 修改显示名。不变量被破坏时静默忽略。
    pub fn set_ole_name(&mut self, name: impl Into<String>) {
        if let Some(o) = self.ole_mut() {
            o.name = name.into();
        }
    }

    /// **是否**以图标形式显示。
    ///
    /// 不变量被破坏时返回 `true`（与默认值一致）。
    pub fn show_as_icon(&self) -> bool {
        match &self.frame.graphic {
            OxmlGraphic::OleObject(o) => o.show_as_icon,
            _ => true,
        }
    }

    /// 设置是否以图标形式显示。不变量被破坏时静默忽略。
    pub fn set_show_as_icon(&mut self, show: bool) {
        if let Some(o) = self.ole_mut() {
            o.show_as_icon = show;
        }
    }

    /// 设置图标显示尺寸（EMU）。
    /// 不变量被破坏时静默忽略。
    pub fn set_icon_size(&mut self, width: Emu, height: Emu) {
        if let Some(o) = self.ole_mut() {
            o.image_width = width;
            o.image_height = height;
        }
    }

    /// 设置图标 Pic 形状的 id 与 name（用于 `<p:oleObj spid="...">` 与 `<p:cNvPr name="...">`）。
    /// 不变量被破坏时静默忽略。
    pub fn set_pic_id_name(&mut self, id: u32, name: impl Into<String>) {
        if let Some(o) = self.ole_mut() {
            o.pic_id = id;
            o.pic_name = name.into();
        }
    }

    /// 将本 OLE 对象形状标记为占位符。
    ///
    /// 写出 XML 时会在 `<p:nvGraphicFramePr>/<p:nvPr>` 内插入
    /// `<p:ph type="obj" idx="..."/>`，使 PowerPoint 把该 graphicFrame
    /// 识别为对象占位符的填充实例。
    ///
    /// # 参数
    /// - `ph_idx`：占位符索引（对应 `<p:ph idx="..."/>`）。
    /// - `ph_type`：占位符类型字符串（如 `"obj"`），`None` 时省略 `type` 属性。
    pub fn set_placeholder(&mut self, ph_idx: u32, ph_type: Option<&str>) {
        self.frame.is_placeholder = true;
        self.frame.ph_idx = Some(ph_idx);
        self.frame.ph_type = ph_type.map(|s| s.to_string());
    }

    /// 清除占位符标记，使本 OLE 对象形状变为普通 graphicFrame。
    pub fn clear_placeholder(&mut self) {
        self.frame.is_placeholder = false;
        self.frame.ph_idx = None;
        self.frame.ph_type = None;
    }

    /// 是否被标记为占位符。
    pub fn is_placeholder(&self) -> bool {
        self.frame.is_placeholder
    }

    /// 占位符索引（若已标记）。
    pub fn ph_idx(&self) -> Option<u32> {
        self.frame.ph_idx
    }

    /// 占位符类型字符串（若已标记）。
    pub fn ph_type(&self) -> Option<&str> {
        self.frame.ph_type.as_deref()
    }
}

impl Shape for OleObjectShape {
    fn id(&self) -> u32 {
        self.frame.id
    }
    fn set_id(&mut self, id: u32) {
        self.frame.id = id;
    }
    fn name(&self) -> &str {
        &self.frame.name
    }
    fn set_name(&mut self, name: String) {
        self.frame.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "ole_object"
    }

    fn left(&self) -> Emu {
        self.frame.properties.xfrm.off_x.unwrap_or_default()
    }
    fn set_left(&mut self, emu: Emu) {
        self.frame.properties.xfrm.off_x = Some(emu);
    }
    fn top(&self) -> Emu {
        self.frame.properties.xfrm.off_y.unwrap_or_default()
    }
    fn set_top(&mut self, emu: Emu) {
        self.frame.properties.xfrm.off_y = Some(emu);
    }
    fn width(&self) -> Emu {
        self.frame.properties.xfrm.ext_cx.unwrap_or_default()
    }
    fn set_width(&mut self, emu: Emu) {
        self.frame.properties.xfrm.ext_cx = Some(emu);
    }
    fn height(&self) -> Emu {
        self.frame.properties.xfrm.ext_cy.unwrap_or_default()
    }
    fn set_height(&mut self, emu: Emu) {
        self.frame.properties.xfrm.ext_cy = Some(emu);
    }

    /// OLE 对象不支持旋转（OOXML 规范）。调用 [`Shape::set_rotation`] 会被忽略。
    fn rotation(&self) -> f64 {
        0.0
    }
    fn set_rotation(&mut self, _deg: f64) {}
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `new` 正确构造 OleObjectShape：progId/name 自定义，rid/image_rid 默认空。
    #[test]
    fn new_ole_shape_basics() {
        let s = OleObjectShape::new("Excel.Sheet.12", "Worksheet");
        assert_eq!(s.prog_id(), "Excel.Sheet.12");
        assert_eq!(s.ole_name(), "Worksheet");
        assert_eq!(s.rid(), "");
        assert_eq!(s.image_rid(), "");
        assert!(s.show_as_icon());
    }

    /// `set_rid` / `set_image_rid` 同步到内部 oxml OleObject。
    #[test]
    fn set_rids_propagate() {
        let mut s = OleObjectShape::new("Package", "Object");
        s.set_rid("rIdOle1");
        s.set_image_rid("rIdImg1");
        assert_eq!(s.rid(), "rIdOle1");
        assert_eq!(s.image_rid(), "rIdImg1");
    }

    /// `set_show_as_icon(false)` 后 `show_as_icon()` 返回 false。
    #[test]
    fn set_show_as_icon() {
        let mut s = OleObjectShape::new("Package", "Object");
        s.set_show_as_icon(false);
        assert!(!s.show_as_icon());
    }

    /// `set_placeholder` 标记后 `is_placeholder()` 返回 true。
    #[test]
    fn set_placeholder_works() {
        let mut s = OleObjectShape::new("Package", "Object");
        assert!(!s.is_placeholder());
        s.set_placeholder(0, Some("obj"));
        assert!(s.is_placeholder());
        assert_eq!(s.ph_idx(), Some(0));
        assert_eq!(s.ph_type(), Some("obj"));
        s.clear_placeholder();
        assert!(!s.is_placeholder());
    }
}
