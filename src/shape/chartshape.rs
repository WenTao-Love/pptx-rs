//! `ChartShape`：高阶图表。
//!
//! 图表在 OOXML 中通过 `<p:graphicFrame>` + `<a:graphicData uri="...chart">`
//! + `<c:chart r:id="..."/>` 引用一个独立的 `/ppt/charts/chartN.xml` part。
//!   本高阶 API 把 graphicFrame 包装为 [`ChartShape`]，提供类型 / 数据 / rid
//!   的便捷访问，并让 [`Shape`] trait 直接可用。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.chart.Chart` ←→ [`crate::oxml::chart::Chart`]（数据/类型/标题）；
//! - `pptx.shapes.graphfrm.GraphicFrame` ←→ 本 [`ChartShape`]（承载位置/尺寸 + 引用）。
//!
//! # 写出语义
//!
//! - `ChartShape` 序列化时只写出 `<p:graphicFrame>` + 引用元素 `<c:chart r:id="..."/>`；
//! - 真正的 chart 数据由 [`crate::presentation::Presentation::save`] 在
//!   `to_opc_package` 中遍历每张 slide 的 `chart_entries` 写出独立的 `chartN.xml` part；
//! - slide 的 `_rels/slideN.xml.rels` 中会添加 `chart` 关系指向该 part。
//!
//! # 限制
//!
//! - 当前仅支持柱/条/线/饼 4 种类型（见 [`crate::oxml::chart::ChartType`]）；
//! - 数据通过 `<c:numCache>` / `<c:strCache>` 内嵌，不依赖嵌入 Excel；
//! - 读取已有图表的 graphicFrame 时，**仅保留 r:id 引用**，不解析 chartN.xml 内容。

use crate::oxml::chart::{Chart as OxmlChart, ChartData, ChartSeries, ChartType, DataLabels};
use crate::oxml::shape::{Graphic as OxmlGraphic, GraphicFrame as OxmlFrame};
use crate::shape::base::Shape;
use crate::units::Emu;

/// 高阶图表形状（承载 `<p:graphicFrame>` + `<c:chart r:id="..."/>` 引用）。
///
/// 通过 [`ChartShape::chart`] / [`ChartShape::chart_mut`] 访问图表数据；
/// 通过 [`Shape`] trait 方法（`left` / `top` / `width` / `height`）调整位置与尺寸。
///
/// # 内部不变量
///
/// `frame.graphic` 始终保持为 `Graphic::Chart(_)`。本类型所有便捷方法
/// （`chart_type` / `set_chart_type` / `data` / `data_mut` / `rid` / `set_rid`）
/// 在不变量被破坏时**静默忽略**或返回 `None`，绝不 panic——
/// 这与库整体"零 panic"约定一致（参见 `.trae/rules/project_rules.md` §5）。
#[derive(Clone, Debug, Default)]
pub struct ChartShape {
    /// 内部 oxml 句柄（`GraphicFrame`，承载 `Graphic::Chart`）。
    pub(crate) frame: OxmlFrame,
}

impl ChartShape {
    /// 构造一个指定类型与数据的图表形状（rid 留空，由 presentation 层填充）。
    ///
    /// # 参数
    /// - `chart_type`：图表类型（柱/条/线/饼）。
    /// - `data`：图表数据（类别 + 系列 + 可选标题）。
    pub fn new(chart_type: ChartType, data: ChartData) -> Self {
        let chart = OxmlChart::new(chart_type, data);
        let frame = OxmlFrame {
            graphic: OxmlGraphic::Chart(chart),
            ..Default::default()
        };
        ChartShape { frame }
    }

    /// 从 oxml Frame 构造（通常用于读取已有图表时）。
    pub fn from_frame(frame: OxmlFrame) -> Self {
        ChartShape { frame }
    }

    /// 取内部 oxml Chart 引用。
    ///
    /// 返回 `None` 仅在内部不变量被外部错误破坏时（`frame.graphic` 不是 `Chart` 变体）。
    pub fn chart(&self) -> Option<&OxmlChart> {
        match &self.frame.graphic {
            OxmlGraphic::Chart(c) => Some(c),
            _ => None,
        }
    }

    /// 取内部 oxml Chart 可变引用。
    pub fn chart_mut(&mut self) -> Option<&mut OxmlChart> {
        match &mut self.frame.graphic {
            OxmlGraphic::Chart(c) => Some(c),
            _ => None,
        }
    }

    /// 图表类型（便捷访问）。
    ///
    /// 若内部不变量被破坏（`frame.graphic` 不是 `Chart`），返回 [`ChartType::Column`] 作为兜底。
    pub fn chart_type(&self) -> ChartType {
        self.chart()
            .map(|c| c.chart_type)
            .unwrap_or(ChartType::Column)
    }

    /// 修改图表类型（保留数据不变）。不变量被破坏时静默忽略。
    pub fn set_chart_type(&mut self, chart_type: ChartType) {
        if let Some(c) = self.chart_mut() {
            c.chart_type = chart_type;
        }
    }

    /// 图表数据引用。
    pub fn data(&self) -> Option<&ChartData> {
        self.chart().map(|c| &c.data)
    }

    /// 图表数据可变引用。
    pub fn data_mut(&mut self) -> Option<&mut ChartData> {
        self.chart_mut().map(|c| &mut c.data)
    }

    /// 当前关联的关系 id（指向 `/ppt/charts/chartN.xml`）。
    ///
    /// 新建时为空字符串；由 `Presentation::to_opc_package` 在写出 chart part 时填充。
    /// 不变量被破坏时返回空字符串。
    pub fn rid(&self) -> &str {
        match &self.frame.graphic {
            OxmlGraphic::Chart(c) => &c.rid,
            _ => "",
        }
    }

    /// 设置关系 id（一般由 presentation 层自动调用，用户无需直接设置）。
    /// 不变量被破坏时静默忽略。
    pub fn set_rid(&mut self, rid: impl Into<String>) {
        if let Some(c) = self.chart_mut() {
            c.rid = rid.into();
        }
    }

    /// 将本图表形状标记为占位符（TODO-007 图表占位符类型化填充）。
    ///
    /// 写出 XML 时会在 `<p:nvGraphicFramePr>/<p:nvPr>` 内插入
    /// `<p:ph type="chart" idx="..."/>`，使 PowerPoint 把该 graphicFrame
    /// 识别为图表占位符的填充实例。
    ///
    /// # 参数
    /// - `ph_idx`：占位符索引（对应 `<p:ph idx="..."/>`）。
    /// - `ph_type`：占位符类型字符串（如 `"chart"` / `"obj"`），`None` 时省略 `type` 属性。
    pub fn set_placeholder(&mut self, ph_idx: u32, ph_type: Option<&str>) {
        self.frame.is_placeholder = true;
        self.frame.ph_idx = Some(ph_idx);
        self.frame.ph_type = ph_type.map(|s| s.to_string());
    }

    /// 清除占位符标记，使本图表形状变为普通 graphicFrame。
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

    // ===================== 数据标签便捷 API（TODO-004 数据标签） =====================

    /// 图表级数据标签引用（`<c:dLbls>`，写在图表元素内、`<c:ser>` 之前）。
    ///
    /// 系列级 [`ChartShape::series_data_labels`] 会覆盖此处的图表级配置。
    /// 返回 `None` 仅在内部不变量被破坏时（`frame.graphic` 不是 `Chart` 变体）。
    pub fn data_labels(&self) -> Option<&DataLabels> {
        self.chart().and_then(|c| c.data.data_labels.as_ref())
    }

    /// 设置图表级数据标签。
    ///
    /// 传入 `None` 清除现有配置。不变量被破坏时静默忽略。
    ///
    /// # 参数
    /// - `labels`：数据标签配置；`None` 清除，`Some(dl)` 写出 `<c:dLbls>`（若非空）。
    pub fn set_data_labels(&mut self, labels: Option<DataLabels>) {
        if let Some(c) = self.chart_mut() {
            c.data.data_labels = labels;
        }
    }

    /// 系列级数据标签引用（按系列索引访问）。
    ///
    /// 系列级配置会覆盖图表级 [`ChartShape::data_labels`]。返回 `None` 的情况：
    /// - 内部不变量被破坏；
    /// - 索引越界；
    /// - 该系列未设置系列级数据标签（继承图表级）。
    pub fn series_data_labels(&self, series_idx: usize) -> Option<&DataLabels> {
        self.chart()
            .and_then(|c| c.data.series.get(series_idx))
            .and_then(|s| s.data_labels.as_ref())
    }

    /// 设置指定系列的数据标签（覆盖图表级配置）。
    ///
    /// # 参数
    /// - `series_idx`：系列索引（越界时静默忽略）；
    /// - `labels`：数据标签配置；`None` 清除该系列的系列级配置（继承图表级）。
    pub fn set_series_data_labels(&mut self, series_idx: usize, labels: Option<DataLabels>) {
        if let Some(c) = self.chart_mut() {
            if let Some(s) = c.data.series.get_mut(series_idx) {
                s.data_labels = labels;
            }
        }
    }

    // ===================== 次坐标轴便捷 API（TODO-004 次坐标轴） =====================

    /// 查询指定系列是否绑定到次坐标轴（右侧 Y 轴）。
    ///
    /// 返回 `false` 的情况：
    /// - 内部不变量被破坏；
    /// - 索引越界；
    /// - 该系列绑定主轴（`secondary_axis=false`）。
    ///
    /// **注意**：饼图/散点图/气泡图不支持次坐标轴，写出时该字段被忽略。
    pub fn is_series_secondary(&self, series_idx: usize) -> bool {
        self.chart()
            .and_then(|c| c.data.series.get(series_idx))
            .map(|s| s.secondary_axis)
            .unwrap_or(false)
    }

    /// 设置指定系列是否绑定到次坐标轴（右侧 Y 轴）。
    ///
    /// 典型场景：主轴显示柱形图，次轴显示折线图（双轴组合图）。
    ///
    /// # 参数
    /// - `series_idx`：系列索引（越界时静默忽略）；
    /// - `secondary`：`true` 绑定次轴，`false` 绑定主轴。
    ///
    /// # 限制
    ///
    /// 饼图/散点图/气泡图不支持次坐标轴（OOXML 规范约束），
    /// `to_xml` 写出时会忽略该字段。
    pub fn set_series_secondary(&mut self, series_idx: usize, secondary: bool) {
        if let Some(c) = self.chart_mut() {
            if let Some(s) = c.data.series.get_mut(series_idx) {
                s.secondary_axis = secondary;
            }
        }
    }

    /// 添加一个新系列到图表数据末尾，返回新系列的索引。
    ///
    /// 便捷方法：避免外部构造 `ChartData` 字面量。不变量被破坏时返回 `None`。
    pub fn push_series(&mut self, series: ChartSeries) -> Option<usize> {
        if let Some(c) = self.chart_mut() {
            c.data.series.push(series);
            Some(c.data.series.len() - 1)
        } else {
            None
        }
    }
}

impl Shape for ChartShape {
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
        "chart"
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

    /// 图表不支持旋转（OOXML 规范）。调用 [`Shape::set_rotation`] 会被忽略。
    fn rotation(&self) -> f64 {
        0.0
    }
    fn set_rotation(&mut self, _deg: f64) {}
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::chart::{ChartCategory, ChartSeries};

    /// `new` 正确构造 ChartShape：类型/数据/rid 默认值均正确。
    #[test]
    fn new_chart_shape_basics() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
            series: vec![ChartSeries::new("S", vec![1.0, 2.0])],
            title: Some("T".to_string()),
            data_labels: None,
        };
        let mut s = ChartShape::new(ChartType::Column, data);
        assert_eq!(s.chart_type(), ChartType::Column);
        assert_eq!(s.chart().unwrap().data.categories.len(), 2);
        assert_eq!(s.chart().unwrap().data.series.len(), 1);
        assert_eq!(s.rid(), "");
        // 切换类型不丢数据
        s.set_chart_type(ChartType::Line);
        assert_eq!(s.chart_type(), ChartType::Line);
        assert_eq!(s.chart().unwrap().data.series[0].values, vec![1.0, 2.0]);
    }

    /// `set_rid` 同步到内部 oxml Chart。
    #[test]
    fn set_rid_propagates() {
        let mut s = ChartShape::new(ChartType::Pie, ChartData::default());
        s.set_rid("rIdChart1");
        assert_eq!(s.rid(), "rIdChart1");
    }

    /// `Shape` trait 的几何 setter 正确写入 frame.properties.xfrm。
    #[test]
    fn shape_trait_geometry() {
        let mut s = ChartShape::new(ChartType::Bar, ChartData::default());
        s.set_left(Emu(100));
        s.set_top(Emu(200));
        s.set_width(Emu(300));
        s.set_height(Emu(400));
        assert_eq!(s.left(), Emu(100));
        assert_eq!(s.top(), Emu(200));
        assert_eq!(s.width(), Emu(300));
        assert_eq!(s.height(), Emu(400));
        // 旋转被忽略
        s.set_rotation(45.0);
        assert_eq!(s.rotation(), 0.0);
    }

    /// `to_xml` 通过 frame 序列化包含 `<c:chart r:id="..."/>` 引用。
    #[test]
    fn frame_xml_contains_chart_reference() {
        let mut s = ChartShape::new(ChartType::Column, ChartData::default());
        s.set_rid("rIdChart7");
        let mut w = crate::oxml::writer::XmlWriter::new();
        s.frame.write_xml(&mut w);
        let xml = w.into_string();
        assert!(xml.contains("<p:graphicFrame>"), "xml: {}", xml);
        assert!(xml.contains("<c:chart"), "xml: {}", xml);
        assert!(xml.contains("r:id=\"rIdChart7\""), "xml: {}", xml);
    }

    // ===================== 数据标签便捷 API 测试 =====================

    /// `set_data_labels` / `data_labels` 往返：设置后能正确读取。
    #[test]
    fn data_labels_setter_and_getter() {
        let mut s = ChartShape::new(ChartType::Column, ChartData::default());
        assert!(s.data_labels().is_none());
        // 设置图表级数据标签
        s.set_data_labels(Some(DataLabels::show_values()));
        assert_eq!(s.data_labels().unwrap().show_val, Some(true));
        // 清除
        s.set_data_labels(None);
        assert!(s.data_labels().is_none());
    }

    /// `set_series_data_labels` / `series_data_labels` 按索引访问。
    #[test]
    fn series_data_labels_by_index() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A")],
            series: vec![
                ChartSeries::new("S1", vec![1.0]),
                ChartSeries::new("S2", vec![2.0]),
            ],
            title: None,
            data_labels: None,
        };
        let mut s = ChartShape::new(ChartType::Column, data);
        // 初始无系列级 dLbls
        assert!(s.series_data_labels(0).is_none());
        // 为第 1 个系列设置
        s.set_series_data_labels(0, Some(DataLabels::show_values()));
        assert_eq!(s.series_data_labels(0).unwrap().show_val, Some(true));
        // 第 2 个系列仍为 None
        assert!(s.series_data_labels(1).is_none());
        // 越界访问返回 None
        assert!(s.series_data_labels(99).is_none());
        // 越界设置静默忽略（不 panic）
        s.set_series_data_labels(99, Some(DataLabels::show_values()));
    }

    // ===================== 次坐标轴便捷 API 测试 =====================

    /// `set_series_secondary` / `is_series_secondary` 按索引访问。
    #[test]
    fn series_secondary_axis_setter_and_getter() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A")],
            series: vec![
                ChartSeries::new("S1", vec![1.0]),
                ChartSeries::new("S2", vec![2.0]),
            ],
            title: None,
            data_labels: None,
        };
        let mut s = ChartShape::new(ChartType::Column, data);
        // 初始所有系列绑定主轴
        assert!(!s.is_series_secondary(0));
        assert!(!s.is_series_secondary(1));
        // 把第 2 个系列绑定到次轴
        s.set_series_secondary(1, true);
        assert!(!s.is_series_secondary(0));
        assert!(s.is_series_secondary(1));
        // 越界访问返回 false（不 panic）
        assert!(!s.is_series_secondary(99));
        // 越界设置静默忽略
        s.set_series_secondary(99, true);
    }

    /// `push_series` 添加新系列到末尾，返回索引。
    #[test]
    fn push_series_appends_to_end() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A")],
            series: vec![ChartSeries::new("S1", vec![1.0])],
            title: None,
            data_labels: None,
        };
        let mut s = ChartShape::new(ChartType::Column, data);
        assert_eq!(s.data().unwrap().series.len(), 1);
        let idx = s.push_series(ChartSeries::new("S2", vec![2.0]));
        assert_eq!(idx, Some(1));
        assert_eq!(s.data().unwrap().series.len(), 2);
        assert_eq!(s.data().unwrap().series[1].name, "S2");
    }
}
