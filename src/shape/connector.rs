//! `Connector`：连接器（直线/折线/曲线）。
//!
//! 连接器在 OOXML 中是一等公民——`p:cxnSp` 与 `p:sp` 平级，但用途上
//! 专门表达"两点之间"的几何关系。PowerPoint 允许在连接器上"挂接"形状，
//! 使其在源/目标形状移动时自动跟随（流程图常用）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.connector.Connector` ←→ [`Connector`]；
//! - `Slide.shapes.add_connector(connector_type, begin_x, begin_y, end_x, end_y)`
//!   返回 [`Connector`]（本库当前未实现 `add_connector` 高阶方法，
//!   仅暴露 `Connector::new` 供高级用户手动构造）。
//!
//! # 几何
//!
//! 连接器几何通常是 `prstGeom=line` / `straightConnector1` 等；通过
//! [`Connector::cxn_mut`] 可直接 mutate 内部 oxml 字段。
//!
//! # 限制
//!
//! - **未实现** 自动连接源/目标形状的"挂接点"语义（`stCxn` / `endCxn`）；
//! - 弯曲连接器（curvedConnector）序列化后 PowerPoint 端能识别但本库未专门测试。

use crate::oxml::shape::Connector as OxmlCxn;
use crate::oxml::simpletypes::MsoConnectorType;
use crate::oxml::simpletypes::PresetGeometry;
use crate::oxml::sppr::Geometry;
use crate::shape::base::Shape;
use crate::units::{Emu, EmuPoint};

/// 连接器（直线 / 折线 / 曲线）。
#[derive(Clone, Debug, Default)]
pub struct Connector {
    /// 内部 oxml 句柄。
    pub(crate) cxn: OxmlCxn,
}

impl Connector {
    /// 从 oxml [`OxmlCxn`] 构造包装。
    pub fn from_cxn(c: OxmlCxn) -> Self {
        Connector { cxn: c }
    }

    /// 新建一个连接器（默认名为 `name`，其余字段由调用方填充）。
    pub fn new(name: impl Into<String>) -> Self {
        Connector {
            cxn: OxmlCxn {
                name: name.into(),
                ..Default::default()
            },
        }
    }

    /// 新建一个**指定类型**的连接器。
    ///
    /// 对应 python-pptx 中 `Connector(connector_type, begin_x, begin_y, end_x, end_y)`
    /// 的"类型"部分。几何由 `connector_type` 决定。
    ///
    /// # 修订历史
    /// 早期版本**不会**自动计算 xfrm（bounding box），导致 begin/end 在序列化时
    /// 被丢弃 / 位置跑到 (0, 0)。现已**自动**根据 begin/end 计算 xfrm：
    /// - off_x = min(begin_x, end_x)
    /// - off_y = min(begin_y, end_y)
    /// - ext_cx = |end_x - begin_x|
    /// - ext_cy = |end_y - begin_y|
    ///
    /// # 注意
    /// 若 begin/end 还未设置，xfrm 保持为 0/0/0/0；可在调用 [`Self::set_begin`]
    /// 与 [`Self::set_end`] 之后用 [`Self::recompute_xfrm`] 重算。
    pub fn new_with_type(name: impl Into<String>, connector_type: MsoConnectorType) -> Self {
        let mut cxn = OxmlCxn {
            name: name.into(),
            ..Default::default()
        };
        cxn.connector_type = Some(connector_type);
        // 把 connector_type 同步到 spPr.geometry（OOXML 中 xfrm 只放变换，几何在 spPr 直接子级）。
        cxn.properties.geometry =
            Some(Geometry::preset(connector_type_to_geometry(connector_type)));
        Connector { cxn }
    }

    /// 重新计算 xfrm 边界盒（基于 begin/end）。
    ///
    /// 若 begin/end 都没设置，xfrm 会被重置为 0/0/0/0。
    ///
    /// # OOXML 语义
    /// `p:spPr/a:xfrm` 表达"此连接器在 slide 上的占位矩形"；`a:off` 是
    /// 该矩形的左上角，`a:ext` 是宽高。**`p:cxnSp` 本身没有 begin/end 元素**，
    /// 端点是用 `<a:xfrm>` 内嵌 `<a:off x=0 y=0/><a:ext cx=... cy=.../>` +
    /// `prstGeom` 的 `prst="line"`（或 bent/curved）配合"反转坐标系"实现的。
    ///
    /// 简化策略：把 begin/end 转成 (offset, extent) 的 bounding box 写入 xfrm，
    /// 由 PowerPoint 端按 line 几何反算端点。这样**位置信息不丢**。
    pub fn recompute_xfrm(&mut self) {
        match (self.cxn.begin, self.cxn.end) {
            (Some((bx, by)), Some((ex, ey))) => {
                let min_x = bx.value().min(ex.value());
                let min_y = by.value().min(ey.value());
                let max_x = bx.value().max(ex.value());
                let max_y = by.value().max(ey.value());
                self.cxn.properties.xfrm.off_x = Some(Emu(min_x));
                self.cxn.properties.xfrm.off_y = Some(Emu(min_y));
                self.cxn.properties.xfrm.ext_cx = Some(Emu(max_x - min_x));
                self.cxn.properties.xfrm.ext_cy = Some(Emu(max_y - min_y));
            }
            _ => {
                self.cxn.properties.xfrm.off_x = Some(Emu(0));
                self.cxn.properties.xfrm.off_y = Some(Emu(0));
                self.cxn.properties.xfrm.ext_cx = Some(Emu(0));
                self.cxn.properties.xfrm.ext_cy = Some(Emu(0));
            }
        }
    }

    /// 取出 oxml 引用。
    pub fn cxn(&self) -> &OxmlCxn {
        &self.cxn
    }
    /// 取出 oxml 可变引用。
    pub fn cxn_mut(&mut self) -> &mut OxmlCxn {
        &mut self.cxn
    }

    /// 形状属性（spPr）不可变引用。
    pub fn properties(&self) -> &crate::oxml::sppr::ShapeProperties {
        &self.cxn.properties
    }
    /// 形状属性（spPr）可变引用（python-pptx 风格，方便 `LineFormat::from`）。
    pub fn properties_mut(&mut self) -> &mut crate::oxml::sppr::ShapeProperties {
        &mut self.cxn.properties
    }

    /// 起点坐标（EMU）。
    pub fn begin(&self) -> Option<EmuPoint> {
        self.cxn.begin.map(|(x, y)| EmuPoint(x.value(), y.value()))
    }
    /// 设置起点坐标。
    ///
    /// **副作用**：会**自动**重算 xfrm 边界盒（除非你显式调用了
    /// `Self::set_no_xfrm_recompute` 关闭）。这是为了与 OOXML 序列化要求一致——
    /// `p:cxnSp` 的端点位置在 xfrm 的 off/ext 中体现。
    pub fn set_begin(&mut self, p: EmuPoint) {
        self.cxn.begin = Some((Emu(p.0), Emu(p.1)));
        self.recompute_xfrm();
    }

    /// 终点坐标（EMU）。
    pub fn end(&self) -> Option<EmuPoint> {
        self.cxn.end.map(|(x, y)| EmuPoint(x.value(), y.value()))
    }
    /// 设置终点坐标。**自动**重算 xfrm。参见 [`Self::set_begin`]。
    pub fn set_end(&mut self, p: EmuPoint) {
        self.cxn.end = Some((Emu(p.0), Emu(p.1)));
        self.recompute_xfrm();
    }

    /// 连接器类型。
    pub fn connector_type(&self) -> Option<MsoConnectorType> {
        self.cxn.connector_type
    }
    /// 设置连接器类型。
    pub fn set_connector_type(&mut self, t: MsoConnectorType) {
        self.cxn.connector_type = Some(t);
        self.cxn.properties.geometry = Some(Geometry::preset(connector_type_to_geometry(t)));
    }

    /// 起点挂接。
    pub fn begin_connection(&self) -> Option<(u32, u32)> {
        self.cxn.st_cxn
    }
    /// 设置起点挂接。
    pub fn set_begin_connection(&mut self, shape_id: u32, idx: u32) {
        self.cxn.st_cxn = Some((shape_id, idx));
    }

    /// 终点挂接。
    pub fn end_connection(&self) -> Option<(u32, u32)> {
        self.cxn.end_cxn
    }
    /// 设置终点挂接。
    pub fn set_end_connection(&mut self, shape_id: u32, idx: u32) {
        self.cxn.end_cxn = Some((shape_id, idx));
    }
}

/// 把 `MsoConnectorType` 映射到对应的 `PresetGeometry` 几何。
///
/// 几何直接落在 `spPr/prstGeom` 上（OOXML 不允许 xfrm 内嵌 prstGeom）。
fn connector_type_to_geometry(t: MsoConnectorType) -> PresetGeometry {
    match t {
        MsoConnectorType::Straight => PresetGeometry::Line,
        MsoConnectorType::Elbow => PresetGeometry::BentConnector2,
        MsoConnectorType::Curve => PresetGeometry::CurvedConnector2,
        MsoConnectorType::BentConnector3 => PresetGeometry::BentConnector3,
        MsoConnectorType::BentConnector4 => PresetGeometry::BentConnector4,
        MsoConnectorType::BentConnector5 => PresetGeometry::BentConnector5,
        MsoConnectorType::CurvedConnector3 => PresetGeometry::CurvedConnector3,
        MsoConnectorType::CurvedConnector4 => PresetGeometry::CurvedConnector4,
        MsoConnectorType::CurvedConnector5 => PresetGeometry::CurvedConnector5,
    }
}

impl Shape for Connector {
    fn id(&self) -> u32 {
        self.cxn.id
    }
    fn set_id(&mut self, id: u32) {
        self.cxn.id = id;
    }
    fn name(&self) -> &str {
        &self.cxn.name
    }
    fn set_name(&mut self, name: String) {
        self.cxn.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "connector"
    }

    fn left(&self) -> Emu {
        self.cxn.properties.xfrm.off_x.unwrap_or_default()
    }
    fn set_left(&mut self, emu: Emu) {
        self.cxn.properties.xfrm.off_x = Some(emu);
    }
    fn top(&self) -> Emu {
        self.cxn.properties.xfrm.off_y.unwrap_or_default()
    }
    fn set_top(&mut self, emu: Emu) {
        self.cxn.properties.xfrm.off_y = Some(emu);
    }
    fn width(&self) -> Emu {
        self.cxn.properties.xfrm.ext_cx.unwrap_or_default()
    }
    fn set_width(&mut self, emu: Emu) {
        self.cxn.properties.xfrm.ext_cx = Some(emu);
    }
    fn height(&self) -> Emu {
        self.cxn.properties.xfrm.ext_cy.unwrap_or_default()
    }
    fn set_height(&mut self, emu: Emu) {
        self.cxn.properties.xfrm.ext_cy = Some(emu);
    }

    fn rotation(&self) -> f64 {
        self.cxn.properties.rot_deg.unwrap_or(0.0)
    }
    fn set_rotation(&mut self, deg: f64) {
        self.cxn.properties.rot_deg = Some(deg);
        let rot = (deg * 60_000.0) as i32;
        self.cxn.properties.xfrm.rot = Some(rot);
    }
}
