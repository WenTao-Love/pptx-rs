//! `SmartArtShape`：高阶 SmartArt 形状（TODO-037 创建 API）。
//!
//! SmartArt（diagram）在 OOXML 中通过 `<p:graphicFrame>` +
//! `<a:graphicData uri=".../diagram">` + `<dgm:relIds r:dm=".." r:lo=".." r:qs=".." r:cs=".."/>`
//! 引用 4 个独立的 `/ppt/diagrams/{data,layout,quickStyles,colors}N.xml` part。
//! 本高阶 API 把 graphicFrame 包装为 [`SmartArtShape`]，提供 4 个关系 id 的
//! 便捷访问与设置，并让 [`Shape`] trait 直接可用。
//!
//! # 与 python-pptx 的对应
//!
//! - python-pptx 当前**未提供** SmartArt 创建 API（截至 0.6.23）；
//! - 本类型参考 `ChartShape` / `OleObjectShape` 的设计模式自研。
//!
//! # 写出语义
//!
//! - `SmartArtShape` 序列化时只写出 `<p:graphicFrame>` + `<a:graphicData>` 内的
//!   `<dgm:relIds r:dm=".." r:lo=".." r:qs=".." r:cs=".."/>` 引用元素；
//! - 真正的 4 个 diagram part（data / layout / quickStyles / colors）由
//!   [`crate::presentation::Presentation::save`] 在 `to_opc_package` 中遍历每张
//!   slide 的 `diagram_entries` 写出独立 part；
//! - slide 的 `_rels/slideN.xml.rels` 中会添加 4 个关系（DiagramData /
//!   DiagramLayout / DiagramQuickStyle / DiagramColors）指向对应 part。
//!
//! # 创建路径
//!
//! - [`ShapesMut::add_smartart`](crate::slide::ShapesMut::add_smartart)：从结构化
//!   模型（`DataModel` + `LayoutDef` + `QuickStyleDef` + `ColorsDef`）创建；
//! - [`ShapesMut::add_smartart_from_xml`](crate::slide::ShapesMut::add_smartart_from_xml)：
//!   从 4 份原始 XML 字符串创建（round-trip 友好的"逃生舱"入口）。

use crate::oxml::shape::{Graphic as OxmlGraphic, GraphicFrame as OxmlFrame, SmartArtRef};
use crate::shape::base::Shape;
use crate::units::Emu;

/// 高阶 SmartArt 形状（承载 `<p:graphicFrame>` + `<dgm:relIds>` 引用）。
///
/// 通过 [`SmartArtShape::smart_art`] / [`SmartArtShape::smart_art_mut`] 访问
/// 内部 [`SmartArtRef`]（含 4 个关系 id 与原始 XML）；通过 [`Shape`] trait 方法
/// （`left` / `top` / `width` / `height`）调整位置与尺寸。
///
/// # 内部不变量
///
/// `frame.graphic` 始终保持为 `Graphic::SmartArt(_)`。本类型所有便捷方法
/// （`dm_rid` / `set_dm_rid` / `lo_rid` / `set_lo_rid` / ...）在不变量被破坏时
/// **静默忽略**或返回空字符串，绝不 panic——
/// 这与库整体"零 panic"约定一致（参见 `.trae/rules/project_rules.md` §5）。
#[derive(Clone, Debug)]
pub struct SmartArtShape {
    /// 内部 oxml 句柄（`GraphicFrame`，承载 `Graphic::SmartArt`）。
    pub(crate) frame: OxmlFrame,
}

impl SmartArtShape {
    /// 构造一个 SmartArt 形状（4 个关系 id 留空，由 presentation 层填充）。
    ///
    /// 内部调用 [`SmartArtRef::from_rids`] 构造初始 `raw_xml`。
    /// 后续通过 [`SmartArtShape::set_dm_rid`] 等 setter 更新任一 rid 时，
    /// `raw_xml` 会**整体重新生成**（4 个 rid 同步刷新）。
    pub fn new() -> Self {
        let smart = SmartArtRef::from_rids("", "", "", "");
        let frame = OxmlFrame {
            graphic: OxmlGraphic::SmartArt(smart),
            ..Default::default()
        };
        SmartArtShape { frame }
    }

    /// 从 4 个关系 id 构造 SmartArt 形状（直接初始化路径）。
    pub fn from_rids(dm_rid: &str, lo_rid: &str, qs_rid: &str, cs_rid: &str) -> Self {
        let smart = SmartArtRef::from_rids(dm_rid, lo_rid, qs_rid, cs_rid);
        let frame = OxmlFrame {
            graphic: OxmlGraphic::SmartArt(smart),
            ..Default::default()
        };
        SmartArtShape { frame }
    }

    /// 从 oxml Frame 构造（通常用于读取已有 SmartArt 时）。
    pub fn from_frame(frame: OxmlFrame) -> Self {
        SmartArtShape { frame }
    }

    /// 取内部 oxml SmartArtRef 引用。
    ///
    /// 返回 `None` 仅在内部不变量被外部错误破坏时（`frame.graphic` 不是 `SmartArt` 变体）。
    pub fn smart_art(&self) -> Option<&SmartArtRef> {
        match &self.frame.graphic {
            OxmlGraphic::SmartArt(s) => Some(s),
            _ => None,
        }
    }

    /// 取内部 oxml SmartArtRef 可变引用。
    pub fn smart_art_mut(&mut self) -> Option<&mut SmartArtRef> {
        match &mut self.frame.graphic {
            OxmlGraphic::SmartArt(s) => Some(s),
            _ => None,
        }
    }

    /// 数据模型关系 id（`r:dm`，指向 `/ppt/diagrams/diagramDataN.xml`）。
    ///
    /// 不变量被破坏时返回空字符串。
    pub fn dm_rid(&self) -> &str {
        self.smart_art()
            .and_then(|s| s.dm_rid.as_deref())
            .unwrap_or("")
    }

    /// 布局定义关系 id（`r:lo`，指向 `/ppt/diagrams/diagramLayoutN.xml`）。
    pub fn lo_rid(&self) -> &str {
        self.smart_art()
            .and_then(|s| s.lo_rid.as_deref())
            .unwrap_or("")
    }

    /// 样式关系 id（`r:qs`，指向 `/ppt/diagrams/diagramQuickStyleN.xml`）。
    pub fn qs_rid(&self) -> &str {
        self.smart_art()
            .and_then(|s| s.qs_rid.as_deref())
            .unwrap_or("")
    }

    /// 颜色关系 id（`r:cs`，指向 `/ppt/diagrams/diagramColorsN.xml`）。
    pub fn cs_rid(&self) -> &str {
        self.smart_art()
            .and_then(|s| s.cs_rid.as_deref())
            .unwrap_or("")
    }

    /// 设置数据模型关系 id（`r:dm`）。
    ///
    /// **重要**：任一 rid 变更都会触发 `raw_xml` 整体重建（4 个 rid 同步刷新），
    /// 因为 `raw_xml` 内的 `<dgm:relIds>` 元素同时承载 4 个属性。
    /// 不变量被破坏时静默忽略。
    pub fn set_dm_rid(&mut self, rid: impl Into<String>) {
        if let Some(s) = self.smart_art_mut() {
            s.dm_rid = Some(rid.into());
            rebuild_raw_xml(s);
        }
    }

    /// 设置布局定义关系 id（`r:lo`）。语义同 [`Self::set_dm_rid`]。
    pub fn set_lo_rid(&mut self, rid: impl Into<String>) {
        if let Some(s) = self.smart_art_mut() {
            s.lo_rid = Some(rid.into());
            rebuild_raw_xml(s);
        }
    }

    /// 设置样式关系 id（`r:qs`）。语义同 [`Self::set_dm_rid`]。
    pub fn set_qs_rid(&mut self, rid: impl Into<String>) {
        if let Some(s) = self.smart_art_mut() {
            s.qs_rid = Some(rid.into());
            rebuild_raw_xml(s);
        }
    }

    /// 设置颜色关系 id（`r:cs`）。语义同 [`Self::set_dm_rid`]。
    pub fn set_cs_rid(&mut self, rid: impl Into<String>) {
        if let Some(s) = self.smart_art_mut() {
            s.cs_rid = Some(rid.into());
            rebuild_raw_xml(s);
        }
    }

    /// 一次性设置全部 4 个关系 id（比逐个 setter 高效：仅重建一次 raw_xml）。
    pub fn set_all_rids(
        &mut self,
        dm_rid: impl Into<String>,
        lo_rid: impl Into<String>,
        qs_rid: impl Into<String>,
        cs_rid: impl Into<String>,
    ) {
        if let Some(s) = self.smart_art_mut() {
            s.dm_rid = Some(dm_rid.into());
            s.lo_rid = Some(lo_rid.into());
            s.qs_rid = Some(qs_rid.into());
            s.cs_rid = Some(cs_rid.into());
            rebuild_raw_xml(s);
        }
    }

    /// 将本 SmartArt 形状标记为占位符。
    ///
    /// 写出 XML 时会在 `<p:nvGraphicFramePr>/<p:nvPr>` 内插入
    /// `<p:ph type="..." idx="..."/>`。
    pub fn set_placeholder(&mut self, ph_idx: u32, ph_type: Option<&str>) {
        self.frame.is_placeholder = true;
        self.frame.ph_idx = Some(ph_idx);
        self.frame.ph_type = ph_type.map(|s| s.to_string());
    }

    /// 清除占位符标记。
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

impl Default for SmartArtShape {
    fn default() -> Self {
        Self::new()
    }
}

/// 根据 SmartArtRef 当前的 4 个 rid 重建 raw_xml。
///
/// 内部委托 [`SmartArtRef::from_rids`]，保证写路径一致。
/// 4 个 rid 中任一为 `None` 时用空字符串兜底。
fn rebuild_raw_xml(s: &mut SmartArtRef) {
    let dm = s.dm_rid.as_deref().unwrap_or("");
    let lo = s.lo_rid.as_deref().unwrap_or("");
    let qs = s.qs_rid.as_deref().unwrap_or("");
    let cs = s.cs_rid.as_deref().unwrap_or("");
    s.raw_xml = SmartArtRef::from_rids(dm, lo, qs, cs).raw_xml;
}

impl Shape for SmartArtShape {
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
        "smart_art"
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

    /// SmartArt 不支持旋转（OOXML 规范）。调用 [`Shape::set_rotation`] 会被忽略。
    fn rotation(&self) -> f64 {
        0.0
    }
    fn set_rotation(&mut self, _deg: f64) {}
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `new` 构造的 SmartArtShape：4 个 rid 默认空，但 raw_xml 已初始化。
    #[test]
    fn new_smartart_shape_basics() {
        let s = SmartArtShape::new();
        assert_eq!(s.dm_rid(), "");
        assert_eq!(s.lo_rid(), "");
        assert_eq!(s.qs_rid(), "");
        assert_eq!(s.cs_rid(), "");
        // raw_xml 非空（包含 graphicData 外壳 + dgm:relIds 空属性）
        assert!(s.smart_art().unwrap().raw_xml.contains("dgm:relIds"));
    }

    /// `from_rids` 直接初始化 4 个 rid。
    #[test]
    fn from_rids_works() {
        let s = SmartArtShape::from_rids("rId1", "rId2", "rId3", "rId4");
        assert_eq!(s.dm_rid(), "rId1");
        assert_eq!(s.lo_rid(), "rId2");
        assert_eq!(s.qs_rid(), "rId3");
        assert_eq!(s.cs_rid(), "rId4");
        let xml = &s.smart_art().unwrap().raw_xml;
        assert!(xml.contains("r:dm=\"rId1\""));
        assert!(xml.contains("r:lo=\"rId2\""));
        assert!(xml.contains("r:qs=\"rId3\""));
        assert!(xml.contains("r:cs=\"rId4\""));
    }

    /// 单个 setter 触发 raw_xml 整体重建（4 个 rid 同步刷新）。
    #[test]
    fn set_single_rid_rebuilds_raw_xml() {
        let mut s = SmartArtShape::from_rids("rId1", "rId2", "rId3", "rId4");
        s.set_dm_rid("rIdDmNew");
        assert_eq!(s.dm_rid(), "rIdDmNew");
        // 其余 3 个 rid 保持不变
        assert_eq!(s.lo_rid(), "rId2");
        assert_eq!(s.qs_rid(), "rId3");
        assert_eq!(s.cs_rid(), "rId4");
        // raw_xml 同步刷新
        let xml = &s.smart_art().unwrap().raw_xml;
        assert!(xml.contains("r:dm=\"rIdDmNew\""));
        assert!(xml.contains("r:lo=\"rId2\""));
    }

    /// `set_all_rids` 一次性更新 4 个 rid（仅重建一次 raw_xml）。
    #[test]
    fn set_all_rids_works() {
        let mut s = SmartArtShape::new();
        s.set_all_rids("a", "b", "c", "d");
        assert_eq!(s.dm_rid(), "a");
        assert_eq!(s.lo_rid(), "b");
        assert_eq!(s.qs_rid(), "c");
        assert_eq!(s.cs_rid(), "d");
        let xml = &s.smart_art().unwrap().raw_xml;
        assert!(xml.contains("r:dm=\"a\""));
        assert!(xml.contains("r:cs=\"d\""));
    }

    /// `set_placeholder` / `clear_placeholder` 标记占位符。
    #[test]
    fn set_placeholder_works() {
        let mut s = SmartArtShape::new();
        assert!(!s.is_placeholder());
        s.set_placeholder(0, Some("obj"));
        assert!(s.is_placeholder());
        assert_eq!(s.ph_idx(), Some(0));
        assert_eq!(s.ph_type(), Some("obj"));
        s.clear_placeholder();
        assert!(!s.is_placeholder());
        assert_eq!(s.ph_idx(), None);
    }

    /// Shape trait 基本几何操作生效。
    #[test]
    fn shape_trait_geometry() {
        let mut s = SmartArtShape::new();
        s.set_left(crate::Emu(100));
        s.set_top(crate::Emu(200));
        s.set_width(crate::Emu(300));
        s.set_height(crate::Emu(400));
        assert_eq!(s.left(), crate::Emu(100));
        assert_eq!(s.top(), crate::Emu(200));
        assert_eq!(s.width(), crate::Emu(300));
        assert_eq!(s.height(), crate::Emu(400));
        // 旋转始终为 0
        s.set_rotation(45.0);
        assert_eq!(s.rotation(), 0.0);
    }
}
