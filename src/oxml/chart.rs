//! 图表（Chart）OOXML 模型 —— 最小可用集（TODO-004）。
//!
//! 本模块定义 `<c:chartSpace>` / `<c:chart>` / `<c:plotArea>` 等元素的强类型模型，
//! 当前支持**柱状图 / 条形图 / 折线图 / 饼图 / 散点图 / 面积图**六种类型，且数据采用
//! `<c:numCache>` 内嵌缓存方式（**不依赖**嵌入 Excel）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.chart.Chart` ←→ [`Chart`]；
//! - `pptx.chart.ChartData` ←→ [`ChartData`]；
//! - `pptx.chart.Series` ←→ [`ChartSeries`]；
//! - `pptx.chart.Category` ←→ [`ChartCategory`]；
//! - `pptx.enum.chart.XL_CHART_TYPE` ←→ [`ChartType`]。
//!
//! # 设计要点
//!
//! - 数据**全部内嵌**在 chartN.xml 中（`<c:numCache>` / `<c:strCache>`），
//!   避免引入嵌入式 Excel 计算链；
//! - 关系 `rId` 指向 `/ppt/charts/chartN.xml` part；
//! - 类型枚举覆盖 6 种常用集，未来扩展到雷达/气泡等只需追加变体。

use crate::oxml::writer::XmlWriter;

/// 图表类型。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Hash)]
pub enum ChartType {
    /// 簇状柱形图（`barChart` with `barDir="col"`）。
    Column,
    /// 簇状条形图（`barChart` with `barDir="bar"`）。
    Bar,
    /// 折线图（`lineChart`）。
    Line,
    /// 饼图（`pieChart`）。
    Pie,
    /// 散点图（`scatterChart`）。系列需提供 `x_values`，使用两个 valAx。
    Scatter,
    /// 面积图（`areaChart`）。
    Area,
    /// 雷达图（`radarChart`）。每个类别一个辐射轴，多系列叠加。
    Radar,
    /// 气泡图（`bubbleChart`）。系列需提供 `x_values` + `bubble_sizes`，使用两个 valAx。
    Bubble,
}

impl ChartType {
    /// 返回 OOXML `<c:chartSpace>` 内 `plotArea` 下的图表元素名（如 `c:barChart`）。
    pub fn chart_element(self) -> &'static str {
        match self {
            ChartType::Column | ChartType::Bar => "c:barChart",
            ChartType::Line => "c:lineChart",
            ChartType::Pie => "c:pieChart",
            ChartType::Scatter => "c:scatterChart",
            ChartType::Area => "c:areaChart",
            ChartType::Radar => "c:radarChart",
            ChartType::Bubble => "c:bubbleChart",
        }
    }

    /// `barChart` 的 `barDir` 属性值。仅对 Column/Bar 有意义。
    pub fn bar_dir(self) -> &'static str {
        match self {
            ChartType::Column => "col",
            ChartType::Bar => "bar",
            _ => "",
        }
    }

    /// 是否为散点图（特殊处理：使用 c:xVal/c:yVal + 两个 valAx）。
    pub fn is_scatter(self) -> bool {
        matches!(self, ChartType::Scatter)
    }

    /// 是否为 X-Y 坐标图（散点图或气泡图）。
    ///
    /// 这两类都使用 `c:xVal` / `c:yVal` + 两个 `valAx`，区别仅在于气泡图额外有 `c:bubbleSize`。
    pub fn is_xy_chart(self) -> bool {
        matches!(self, ChartType::Scatter | ChartType::Bubble)
    }

    /// 是否为气泡图（额外需要 `bubble_sizes` 字段）。
    pub fn is_bubble(self) -> bool {
        matches!(self, ChartType::Bubble)
    }
}

/// 数据标签位置（`<c:dLblPos val="..."/>`）。
///
/// OOXML ST_DLblPos 枚举，控制数据标签相对于数据点的显示位置。
/// 不同图表类型支持的位置子集不同（参见 ECMA-376 第 21.2.2.40 节）。
#[derive(Copy, Clone, Debug, Default, Eq, PartialEq, Hash)]
pub enum LabelPosition {
    /// 最佳位置（由 PowerPoint 自动选择）。适用于大多数图表类型。
    #[default]
    BestFit,
    /// 数据点上方（柱形/折线/面积图）。
    Above,
    /// 数据点下方（柱形/折线/面积图）。
    Below,
    /// 数据点中心（柱形/饼图）。
    Center,
    /// 基线（柱形图底部）。
    Base,
    /// 数据点内侧（柱形/饼图，标签在柱内）。
    InsideBase,
    /// 数据点内侧端（柱形图顶部内侧）。
    InsideEnd,
    /// 数据点外侧端（柱形图顶部外侧）。
    OutsideEnd,
    /// 左侧（折线图/散点图）。
    Left,
    /// 右侧（折线图/散点图）。
    Right,
    /// 饼图扇区外侧。
    OutsideEndPie,
    /// 饼图扇区内侧。
    InsideEndPie,
}

impl LabelPosition {
    /// 转 OOXML `<c:dLblPos val="..."/>` 属性值。
    pub fn as_str(self) -> &'static str {
        match self {
            LabelPosition::BestFit => "bestFit",
            LabelPosition::Above => "t",
            LabelPosition::Below => "b",
            LabelPosition::Center => "ctr",
            LabelPosition::Base => "inBase",
            LabelPosition::InsideBase => "inBase",
            LabelPosition::InsideEnd => "inEnd",
            LabelPosition::OutsideEnd => "outEnd",
            LabelPosition::Left => "l",
            LabelPosition::Right => "r",
            LabelPosition::OutsideEndPie => "outEnd",
            LabelPosition::InsideEndPie => "inEnd",
        }
    }

    /// 从 OOXML 属性值解析（不区分大小写），未知值返回 `None`。
    ///
    /// 注：方法名为 `parse` 而非 `from_str`，以避免与 `std::str::FromStr` trait 冲突。
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "bestFit" => Some(LabelPosition::BestFit),
            "t" => Some(LabelPosition::Above),
            "b" => Some(LabelPosition::Below),
            "ctr" => Some(LabelPosition::Center),
            "inBase" => Some(LabelPosition::Base),
            "inEnd" => Some(LabelPosition::InsideEnd),
            "outEnd" => Some(LabelPosition::OutsideEnd),
            "l" => Some(LabelPosition::Left),
            "r" => Some(LabelPosition::Right),
            _ => None,
        }
    }
}

/// 数据标签配置（`<c:dLbls>`），对应 OOXML CT_DLbls。
///
/// 控制数据点数值/类别/系列名等标签的显示。可挂在图表级（`<c:barChart>` 内）
/// 或系列级（`<c:ser>` 内）——系列级会覆盖图表级。
///
/// # 字段语义
///
/// 所有 `show_*` 字段为 `Option<bool>`：
/// - `None`：不写出该 `<c:showXxx>` 元素（继承上层或 PowerPoint 默认）；
/// - `Some(true)`：写出 `<c:showXxx val="1"/>`；
/// - `Some(false)`：写出 `<c:showXxx val="0"/>`。
///
/// # 与 python-pptx 的对应
///
/// - python-pptx `plot.data_labels` ←→ 图表级 `ChartData.data_labels`；
/// - python-pptx `series.data_labels` ←→ 系列级 `ChartSeries.data_labels`。
#[derive(Clone, Debug, Default)]
pub struct DataLabels {
    /// 显示数值（`<c:showVal val="1"/>`）。最常用的数据标签。
    pub show_val: Option<bool>,
    /// 显示类别名（`<c:showCatName val="1"/>`）。
    pub show_cat_name: Option<bool>,
    /// 显示系列名（`<c:showSerName val="1"/>`）。
    pub show_ser_name: Option<bool>,
    /// 显示图例键（`<c:showLegendKey val="1"/>`，图例色块前缀）。
    pub show_legend_key: Option<bool>,
    /// 显示百分比（`<c:showPercent val="1"/>`，主要用于饼图）。
    pub show_percent: Option<bool>,
    /// 显示气泡尺寸（`<c:showBubbleSize val="1"/>`，仅气泡图）。
    pub show_bubble_size: Option<bool>,
    /// 标签位置（`<c:dLblPos val="..."/>`）。
    pub position: Option<LabelPosition>,
    /// 分隔符（`<c:separator val="..."/>`，如 `", "` / `"\n"`）。
    pub separator: Option<String>,
    /// 数字格式（`<c:numFmt formatCode="..." sourceLinked="0"/>`）。
    pub num_fmt: Option<String>,
}

impl DataLabels {
    /// 新建一个仅显示数值的数据标签配置（最常用场景）。
    pub fn show_values() -> Self {
        DataLabels {
            show_val: Some(true),
            ..Default::default()
        }
    }

    /// 新建一个显示百分比的饼图数据标签配置。
    pub fn show_percent_pie() -> Self {
        DataLabels {
            show_percent: Some(true),
            ..Default::default()
        }
    }

    /// 写出 `<c:dLbls>` 元素到 XmlWriter。
    ///
    /// OOXML 元素顺序约束（ECMA-376 21.2.2.40）：
    /// `numFmt` → `spacing` → `position` → `separator` → `showLegendKey` →
    /// `showVal` → `showCatName` → `showSerName` → `showPercent` →
    /// `showBubbleSize` → `dLbl` → `dLblPos` → ...
    ///
    /// 实际 PowerPoint 实践中常用顺序：`showVal` → `showCatName` → ... → `dLblPos` → `separator` → `numFmt`。
    /// 本实现采用 PowerPoint 实践顺序以保证兼容性。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        w.open("c:dLbls");
        if let Some(b) = self.show_legend_key {
            w.empty_with("c:showLegendKey", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(b) = self.show_val {
            w.empty_with("c:showVal", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(b) = self.show_cat_name {
            w.empty_with("c:showCatName", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(b) = self.show_ser_name {
            w.empty_with("c:showSerName", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(b) = self.show_percent {
            w.empty_with("c:showPercent", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(b) = self.show_bubble_size {
            w.empty_with("c:showBubbleSize", &[("val", if b { "1" } else { "0" })]);
        }
        if let Some(pos) = self.position {
            w.empty_with("c:dLblPos", &[("val", pos.as_str())]);
        }
        if let Some(sep) = &self.separator {
            w.empty_with("c:separator", &[("val", sep.as_str())]);
        }
        if let Some(fmt) = &self.num_fmt {
            w.empty_with(
                "c:numFmt",
                &[("formatCode", fmt.as_str()), ("sourceLinked", "0")],
            );
        }
        w.close("c:dLbls");
    }

    /// 是否所有字段都为 `None`（即写出后只有空 `<c:dLbls/>`，无意义）。
    pub fn is_empty(&self) -> bool {
        self.show_val.is_none()
            && self.show_cat_name.is_none()
            && self.show_ser_name.is_none()
            && self.show_legend_key.is_none()
            && self.show_percent.is_none()
            && self.show_bubble_size.is_none()
            && self.position.is_none()
            && self.separator.is_none()
            && self.num_fmt.is_none()
    }
}

/// 图表数据：一维类别轴 + 多个数据系列。
#[derive(Clone, Debug, Default)]
pub struct ChartData {
    /// 类别（X 轴标签或饼图扇区标签）。
    ///
    /// **散点图忽略此字段**——散点图的 X 坐标由 `ChartSeries.x_values` 提供。
    pub categories: Vec<ChartCategory>,
    /// 数据系列列表（每个系列一条曲线/一组柱子/一个饼图扇区集合）。
    pub series: Vec<ChartSeries>,
    /// 图表标题。`None` 表示无标题。
    pub title: Option<String>,
    /// 图表级数据标签（`<c:dLbls>`，写在 `<c:barChart>` 等图表元素内、`<c:ser>` 之前）。
    ///
    /// 系列级 [`ChartSeries::data_labels`] 会覆盖此处的图表级配置。
    /// `None` 表示不写出图表级 `<c:dLbls>`（PowerPoint 使用默认：不显示标签）。
    pub data_labels: Option<DataLabels>,
}

/// 单个类别（X 轴标签或饼图扇区标签）。
#[derive(Clone, Debug, Default)]
pub struct ChartCategory {
    /// 类别名称。
    pub name: String,
}

impl ChartCategory {
    /// 新建一个类别。
    pub fn new(name: impl Into<String>) -> Self {
        ChartCategory { name: name.into() }
    }
}

/// 数据系列（一条折线 / 一组柱子 / 一组饼图扇区 / 一组散点 / 一组气泡）。
#[derive(Clone, Debug, Default)]
pub struct ChartSeries {
    /// 系列名称（图例上显示）。
    pub name: String,
    /// 数值列表，长度应与 `ChartData.categories` 一致。
    ///
    /// 对散点图/气泡图，此字段表示 **Y 坐标**值，需配合 `x_values` 使用。
    pub values: Vec<f64>,
    /// 散点图/气泡图的 **X 坐标**值列表。仅 `ChartType::Scatter` / `Bubble` 使用，长度应与 `values` 一致。
    ///
    /// `None` 或空表示该系列非 X-Y 坐标图；散点图系列应通过 [`ChartSeries::new_scatter`]
    /// 构造以同时提供 `x_values` 与 `values`（Y 坐标）。
    pub x_values: Option<Vec<f64>>,
    /// 气泡图的**气泡尺寸**列表。仅 `ChartType::Bubble` 使用，长度应与 `values` 一致。
    ///
    /// `None` 表示非气泡图；气泡图系列应通过 [`ChartSeries::new_bubble`] 构造。
    pub bubble_sizes: Option<Vec<f64>>,
    /// 系列级数据标签（`<c:dLbls>`，写在 `<c:ser>` 内 `<c:val>` / `<c:yVal>` 之后）。
    ///
    /// 系列级配置会覆盖图表级 [`ChartData::data_labels`]。`None` 表示继承图表级配置。
    pub data_labels: Option<DataLabels>,
    /// 是否将该系列绑定到**次坐标轴**（右侧 Y 轴）。
    ///
    /// - `false`（默认）：系列绑定主轴（axId=222222222）。
    /// - `true`：系列绑定次轴（axId=444444444），`Chart::to_xml` 会额外写出
    ///   次轴定义（含 `<c:crosses val="max"/>`），PowerPoint 将其显示在右侧。
    ///
    /// 典型场景：主轴显示柱形图，次轴显示折线图（双轴组合图）。
    /// **注意**：饼图/散点图/气泡图不支持次坐标轴（写出时该字段被忽略）。
    pub secondary_axis: bool,
}

impl ChartSeries {
    /// 新建一个数据系列（非散点图/气泡图）。
    ///
    /// `values` 为 Y 值列表（柱/条/线/面积/雷达图）或扇区数值（饼图）。
    pub fn new(name: impl Into<String>, values: Vec<f64>) -> Self {
        ChartSeries {
            name: name.into(),
            values,
            x_values: None,
            bubble_sizes: None,
            data_labels: None,
            secondary_axis: false,
        }
    }

    /// 新建一个散点图数据系列（提供 X 与 Y 坐标）。
    ///
    /// `x_values` 与 `y_values` 长度应一致；不一致时以较短的为准（写出时按 `y_values` 长度迭代）。
    ///
    /// # 参数
    /// - `name`：系列名称（图例显示）。
    /// - `x_values`：X 坐标列表。
    /// - `y_values`：Y 坐标列表（与 [`ChartSeries::values`] 同字段）。
    pub fn new_scatter(name: impl Into<String>, x_values: Vec<f64>, y_values: Vec<f64>) -> Self {
        ChartSeries {
            name: name.into(),
            values: y_values,
            x_values: Some(x_values),
            bubble_sizes: None,
            data_labels: None,
            secondary_axis: false,
        }
    }

    /// 新建一个气泡图数据系列（提供 X、Y 坐标与气泡尺寸）。
    ///
    /// 三个列表长度应一致；不一致时以 `y_values` 长度为准迭代。
    ///
    /// # 参数
    /// - `name`：系列名称（图例显示）。
    /// - `x_values`：X 坐标列表。
    /// - `y_values`：Y 坐标列表（与 [`ChartSeries::values`] 同字段）。
    /// - `bubble_sizes`：气泡尺寸列表（决定每个气泡的面积）。
    pub fn new_bubble(
        name: impl Into<String>,
        x_values: Vec<f64>,
        y_values: Vec<f64>,
        bubble_sizes: Vec<f64>,
    ) -> Self {
        ChartSeries {
            name: name.into(),
            values: y_values,
            x_values: Some(x_values),
            bubble_sizes: Some(bubble_sizes),
            data_labels: None,
            secondary_axis: false,
        }
    }
}

/// 一个完整的图表，引用 chart part。
#[derive(Clone, Debug)]
pub struct Chart {
    /// 图表类型。
    pub chart_type: ChartType,
    /// 图表数据。
    pub data: ChartData,
    /// 关系 id（指向 `/ppt/charts/chartN.xml`）。
    ///
    /// 由 `Presentation::to_opc_package` 在写出 chart part 时自动分配。
    /// 高阶 API 创建时通常为空字符串，由 presentation 层填充。
    pub rid: String,
    /// 嵌入式 Excel 工作簿的关系 id（TODO-004 Excel 嵌入）。
    ///
    /// 指向 `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`，写出时在
    /// `</c:chart>` 之后、`</c:chartSpace>` 之前生成
    /// `<c:externalData r:id="..."><c:autoUpdate val="0"/></c:externalData>`。
    ///
    /// - `None`：不写出 `<c:externalData>`（图表数据仅靠 numCache/strCache）。
    /// - `Some(rid)`：写出 `<c:externalData r:id="{rid}"/>`，PowerPoint 打开
    ///   图表时会从 xlsx part 读取数据源（"编辑数据" 启动 Excel）。
    ///
    /// 由 `Presentation::to_opc_package` 在写出 xlsx part 时自动分配；
    /// 高阶 API 创建时通常为 `None`，由 presentation 层填充。
    pub external_data_rid: Option<String>,
}

impl Chart {
    /// 新建一个图表。`rid` 默认为空字符串，由 presentation 层在写出时填充。
    pub fn new(chart_type: ChartType, data: ChartData) -> Self {
        Chart {
            chart_type,
            data,
            rid: String::new(),
            external_data_rid: None,
        }
    }

    /// 从 `<c:chartSpace>` XML 字符串解析为强类型 [`Chart`]。
    ///
    /// 详见模块级独立函数 [`parse_from_xml`]（本关联函数是其委托）。
    ///
    /// # 参数
    /// - `xml`：完整 chartN.xml 内容。
    ///
    /// # 错误
    /// - `Error::Xml`：XML 解析失败。
    pub fn parse_from_xml(xml: &str) -> crate::Result<Chart> {
        parse_from_xml(xml)
    }
}

impl Default for Chart {
    /// 默认图表：`ChartType::Column` + 空 `ChartData` + 空 rid。
    fn default() -> Self {
        Chart {
            chart_type: ChartType::Column,
            data: ChartData::default(),
            rid: String::new(),
            external_data_rid: None,
        }
    }
}

impl Chart {
    /// 写出 `<c:chartSpace>` 完整 XML（chartN.xml 内容）。
    ///
    /// # 元素结构（最小可用集）
    ///
    /// ```text
    /// <c:chartSpace xmlns:c=... xmlns:a=... xmlns:r=...>
    ///   <c:chart>
    ///     <c:title>...?         ← 可选
    ///     <c:plotArea>
    ///       <c:layout/>         ← 自动布局
    ///       <c:barChart|c:lineChart|c:pieChart>
    ///         <c:barDir>...     ← 仅 barChart
    ///         <c:grouping val="clustered"/>
    ///         <c:ser>           ← 一个或多个系列
    ///           <c:idx val="0"/>
    ///           <c:order val="0"/>
    ///           <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>...</c:strCache></c:strRef></c:tx>
    ///           <c:cat><c:strRef>...类别标签...</c:strRef></c:cat>
    ///           <c:val><c:numRef><c:numCache>...数值...</c:numCache></c:numRef></c:val>
    ///         </c:ser>
    ///         <c:axId val="..."/>  ← 1~2 个 axId
    ///       </c:barChart>
    ///       <c:catAx>...</c:catAx>     ← 类别轴（饼图无）
    ///       <c:valAx>...</c:valAx>     ← 数值轴（饼图无）
    ///     </c:plotArea>
    ///   </c:chart>
    /// </c:chartSpace>
    /// ```
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::new();
        w.raw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        w.open_with(
            "c:chartSpace",
            &[
                (
                    "xmlns:c",
                    "http://schemas.openxmlformats.org/drawingml/2006/chart",
                ),
                (
                    "xmlns:a",
                    "http://schemas.openxmlformats.org/drawingml/2006/main",
                ),
                (
                    "xmlns:r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ),
            ],
        );
        w.open("c:chart");
        // 标题
        if let Some(title) = &self.data.title {
            w.open("c:title");
            w.open("a:tx");
            w.open("a:rich");
            w.open("a:bodyPr");
            w.close("a:bodyPr");
            w.open("a:lstStyle");
            w.close("a:lstStyle");
            w.open("a:p");
            w.open("a:pPr");
            w.empty_with("a:defRPr", &[("sz", "1400"), ("b", "1")]);
            w.close("a:pPr");
            w.leaf("a:t", title.as_str());
            w.close("a:p");
            w.close("a:rich");
            w.close("a:tx");
            w.empty_with("c:overlay", &[("val", "0")]);
            w.close("c:title");
        }
        // plotArea
        w.open("c:plotArea");
        w.open("c:layout");
        w.close("c:layout");

        let chart_elem = self.chart_type.chart_element();
        w.open(chart_elem);
        if matches!(self.chart_type, ChartType::Column | ChartType::Bar) {
            w.empty_with("c:barDir", &[("val", self.chart_type.bar_dir())]);
            w.empty_with("c:grouping", &[("val", "clustered")]);
        }
        // 面积图 grouping（standard = 普通面积图，不堆叠）
        if matches!(self.chart_type, ChartType::Area) {
            w.empty_with("c:grouping", &[("val", "standard")]);
        }
        // 散点图 scatterStyle（lineMarker = 带标记的折线散点）
        if matches!(self.chart_type, ChartType::Scatter) {
            w.empty_with("c:scatterStyle", &[("val", "lineMarker")]);
        }
        // 雷达图 radarStyle（marker = 带数据标记的雷达图）
        if matches!(self.chart_type, ChartType::Radar) {
            w.empty_with("c:radarStyle", &[("val", "marker")]);
        }
        // 气泡图 bubbleScale（100 = 默认缩放比例）
        if matches!(self.chart_type, ChartType::Bubble) {
            w.empty_with("c:bubbleScale", &[("val", "100")]);
        }
        // varColors
        if matches!(self.chart_type, ChartType::Pie) {
            w.empty_with("c:varyColors", &[("val", "1")]);
        }

        // 图表级数据标签（<c:dLbls>，OOXML 顺序：在 varyColors 之后、<c:ser> 之前）。
        // 系列级 ChartSeries.data_labels 会覆盖此处的图表级配置。
        if let Some(dl) = &self.data.data_labels {
            if !dl.is_empty() {
                dl.write_xml(&mut w);
            }
        }

        // 系列列表
        for (idx, s) in self.data.series.iter().enumerate() {
            let idx_s = idx.to_string();
            w.open("c:ser");
            w.empty_with("c:idx", &[("val", idx_s.as_str())]);
            w.empty_with("c:order", &[("val", idx_s.as_str())]);
            // 系列名（c:tx）
            w.open("c:tx");
            w.open("c:strRef");
            w.leaf("c:f", &format!("Sheet1!${}$1", col_letter(idx + 1)));
            w.open("c:strCache");
            w.open_with("c:pt", &[("idx", "0")]);
            w.leaf("c:v", s.name.as_str());
            w.close("c:pt");
            w.close("c:strCache");
            w.close("c:strRef");
            w.close("c:tx");
            if self.chart_type.is_xy_chart() {
                // 散点图/气泡图：c:xVal / c:yVal（均为数值）
                // c:xVal
                w.open("c:xVal");
                w.open("c:numRef");
                w.leaf("c:f", "Sheet1!$A$2:$A$100");
                w.open("c:numCache");
                w.empty_with("c:formatCode", &[("val", "General")]);
                if let Some(xs) = &s.x_values {
                    for (i, v) in xs.iter().enumerate() {
                        let i_s = i.to_string();
                        let v_s = format_f64(*v);
                        w.open_with("c:pt", &[("idx", i_s.as_str())]);
                        w.leaf("c:v", v_s.as_str());
                        w.close("c:pt");
                    }
                }
                w.close("c:numCache");
                w.close("c:numRef");
                w.close("c:xVal");
                // c:yVal
                w.open("c:yVal");
                w.open("c:numRef");
                w.leaf(
                    "c:f",
                    &format!(
                        "Sheet1!${}$2:${}$100",
                        col_letter(idx + 1),
                        col_letter(idx + 1)
                    ),
                );
                w.open("c:numCache");
                w.empty_with("c:formatCode", &[("val", "General")]);
                for (i, v) in s.values.iter().enumerate() {
                    let i_s = i.to_string();
                    let v_s = format_f64(*v);
                    w.open_with("c:pt", &[("idx", i_s.as_str())]);
                    w.leaf("c:v", v_s.as_str());
                    w.close("c:pt");
                }
                w.close("c:numCache");
                w.close("c:numRef");
                w.close("c:yVal");
                // 气泡图额外：c:bubbleSize（气泡尺寸，决定每个气泡的面积）
                if self.chart_type.is_bubble() {
                    w.open("c:bubbleSize");
                    w.open("c:numRef");
                    w.leaf(
                        "c:f",
                        &format!(
                            "Sheet1!${}$2:${}$100",
                            col_letter(idx + 2),
                            col_letter(idx + 2)
                        ),
                    );
                    w.open("c:numCache");
                    w.empty_with("c:formatCode", &[("val", "General")]);
                    if let Some(sizes) = &s.bubble_sizes {
                        for (i, v) in sizes.iter().enumerate() {
                            let i_s = i.to_string();
                            let v_s = format_f64(*v);
                            w.open_with("c:pt", &[("idx", i_s.as_str())]);
                            w.leaf("c:v", v_s.as_str());
                            w.close("c:pt");
                        }
                    }
                    w.close("c:numCache");
                    w.close("c:numRef");
                    w.close("c:bubbleSize");
                }
            } else {
                // 非散点图：c:cat（类别标签） + c:val（数值）
                // 类别标签（c:cat）—— 仅 bar/line/area 有
                if !matches!(self.chart_type, ChartType::Pie) && !self.data.categories.is_empty() {
                    w.open("c:cat");
                    w.open("c:strRef");
                    w.leaf("c:f", "Sheet1!$A$2:$A$100");
                    w.open("c:strCache");
                    for (i, cat) in self.data.categories.iter().enumerate() {
                        let i_s = i.to_string();
                        w.open_with("c:pt", &[("idx", i_s.as_str())]);
                        w.leaf("c:v", cat.name.as_str());
                        w.close("c:pt");
                    }
                    w.close("c:strCache");
                    w.close("c:strRef");
                    w.close("c:cat");
                } else if matches!(self.chart_type, ChartType::Pie)
                    && !self.data.categories.is_empty()
                {
                    // 饼图：类别标签写在 c:cat 里
                    w.open("c:cat");
                    w.open("c:strRef");
                    w.leaf("c:f", "Sheet1!$A$2:$A$100");
                    w.open("c:strCache");
                    for (i, cat) in self.data.categories.iter().enumerate() {
                        let i_s = i.to_string();
                        w.open_with("c:pt", &[("idx", i_s.as_str())]);
                        w.leaf("c:v", cat.name.as_str());
                        w.close("c:pt");
                    }
                    w.close("c:strCache");
                    w.close("c:strRef");
                    w.close("c:cat");
                }
                // 数值（c:val）
                w.open("c:val");
                w.open("c:numRef");
                w.leaf(
                    "c:f",
                    &format!(
                        "Sheet1!${}$2:${}$100",
                        col_letter(idx + 1),
                        col_letter(idx + 1)
                    ),
                );
                w.open("c:numCache");
                w.empty_with("c:formatCode", &[("val", "General")]);
                for (i, v) in s.values.iter().enumerate() {
                    let i_s = i.to_string();
                    let v_s = format_f64(*v);
                    w.open_with("c:pt", &[("idx", i_s.as_str())]);
                    w.leaf("c:v", v_s.as_str());
                    w.close("c:pt");
                }
                w.close("c:numCache");
                w.close("c:numRef");
                w.close("c:val");
            }
            // 系列级数据标签（<c:dLbls>，OOXML 顺序：在 <c:val>/<c:yVal> 之后、</c:ser> 之前）。
            // 覆盖图表级 ChartData.data_labels 配置。
            if let Some(dl) = &s.data_labels {
                if !dl.is_empty() {
                    dl.write_xml(&mut w);
                }
            }
            w.close("c:ser");
        }

        // 轴 ID（饼图无轴；散点图/气泡图两个 valAx）。
        //
        // 次坐标轴支持（TODO-004 次坐标轴）：当某系列 `secondary_axis=true` 时，
        // 该系列额外引用次轴 axId=444444444；图表级 axId 列表需额外包含次轴 id，
        // 并在下方轴定义段写出第三个轴（次 valAx，含 `<c:crosses val="max"/>`）。
        //
        // 饼图/散点图/气泡图不支持次坐标轴（OOXML 规范约束），secondary_axis 字段被忽略。
        let has_secondary = !self.chart_type.is_xy_chart()
            && !matches!(self.chart_type, ChartType::Pie)
            && self.data.series.iter().any(|s| s.secondary_axis);
        if matches!(self.chart_type, ChartType::Pie) {
            // 饼图无轴
        } else if self.chart_type.is_xy_chart() {
            // 散点图/气泡图：两个 valAx（不支持次坐标轴）
            w.empty_with("c:axId", &[("val", "111111111")]);
            w.empty_with("c:axId", &[("val", "222222222")]);
        } else {
            // 柱/条/线/面积/雷达：catAx + valAx
            w.empty_with("c:axId", &[("val", "111111111")]);
            w.empty_with("c:axId", &[("val", "222222222")]);
            // 次坐标轴：当存在 secondary_axis=true 系列时，额外引用次 valAx（axId=444444444）。
            // 次轴定义在下方轴定义段写出，含 <c:crosses val="max"/> 标记。
            if has_secondary {
                w.empty_with("c:axId", &[("val", "444444444")]);
            }
        }
        w.close(chart_elem);

        // 轴定义（饼图跳过）
        if !matches!(self.chart_type, ChartType::Pie) {
            if self.chart_type.is_xy_chart() {
                // 散点图/气泡图：两个 valAx（X 轴底部 + Y 轴左侧）
                // valAx（X 轴）
                w.open("c:valAx");
                w.empty_with("c:axId", &[("val", "111111111")]);
                w.empty_with("c:scaling", &[("orientation", "minMax")]);
                w.empty_with("c:delete", &[("val", "0")]);
                w.empty_with("c:axPos", &[("val", "b")]);
                w.empty_with("c:crossAx", &[("val", "222222222")]);
                w.close("c:valAx");
                // valAx（Y 轴）
                w.open("c:valAx");
                w.empty_with("c:axId", &[("val", "222222222")]);
                w.empty_with("c:scaling", &[("orientation", "minMax")]);
                w.empty_with("c:delete", &[("val", "0")]);
                w.empty_with("c:axPos", &[("val", "l")]);
                w.empty_with("c:crossAx", &[("val", "111111111")]);
                w.close("c:valAx");
            } else {
                // 柱/条/线/面积：catAx + valAx
                // catAx（主类别轴，底部）
                w.open("c:catAx");
                w.empty_with("c:axId", &[("val", "111111111")]);
                w.empty_with("c:scaling", &[("orientation", "minMax")]);
                w.empty_with("c:delete", &[("val", "0")]);
                w.empty_with("c:axPos", &[("val", "b")]);
                w.empty_with("c:crossAx", &[("val", "222222222")]);
                w.close("c:catAx");
                // valAx（主数值轴，左侧）
                w.open("c:valAx");
                w.empty_with("c:axId", &[("val", "222222222")]);
                w.empty_with("c:scaling", &[("orientation", "minMax")]);
                w.empty_with("c:delete", &[("val", "0")]);
                w.empty_with("c:axPos", &[("val", "l")]);
                w.empty_with("c:crossAx", &[("val", "111111111")]);
                w.close("c:valAx");
                // 次数值轴（右侧，仅当 has_secondary 时写出）。
                //
                // <c:crosses val="max"/> 是 PowerPoint 识别次轴的关键元素：
                // 表示该轴在主轴的最大值处交叉（即显示在右侧/顶部）。
                // axPos="r" 表示右侧；crossAx 指向主 catAx（111111111）。
                if has_secondary {
                    w.open("c:valAx");
                    w.empty_with("c:axId", &[("val", "444444444")]);
                    w.empty_with("c:scaling", &[("orientation", "minMax")]);
                    w.empty_with("c:delete", &[("val", "0")]);
                    w.empty_with("c:axPos", &[("val", "r")]);
                    w.empty_with("c:crossAx", &[("val", "111111111")]);
                    w.empty_with("c:crosses", &[("val", "max")]);
                    w.close("c:valAx");
                }
            }
        }
        w.close("c:plotArea");

        // legend
        w.open("c:legend");
        w.empty_with("c:legendPos", &[("val", "r")]);
        w.empty_with("c:overlay", &[("val", "0")]);
        w.close("c:legend");

        w.empty_with("c:plotVisOnly", &[("val", "1")]);
        w.empty_with("c:dispBlanksAs", &[("val", "gap")]);
        w.close("c:chart");

        // 嵌入式 Excel 工作簿引用（TODO-004 Excel 嵌入）。
        //
        // OOXML 顺序约束：`<c:externalData>` 必须在 `<c:chart>` 之后、
        // `<c:chartSpace>` 关闭之前。PowerPoint 读取该元素后会从对应
        // xlsx part 加载数据源（"编辑数据" 启动 Excel）。
        //
        // `<c:autoUpdate val="0"/>` 表示不自动同步（用户手动编辑后才更新），
        // 这是 PowerPoint 默认行为；val="1" 会在打开时自动同步，可能卡顿。
        if let Some(ext_rid) = &self.external_data_rid {
            w.open_with("c:externalData", &[("r:id", ext_rid.as_str())]);
            w.empty_with("c:autoUpdate", &[("val", "0")]);
            w.close("c:externalData");
        }

        w.close("c:chartSpace");
        w.into_string()
    }
}

/// 把 1-based 索引转为 Excel 列字母（1=A, 2=B, ..., 27=AA）。
fn col_letter(n: usize) -> String {
    let mut s = String::new();
    let mut n = n;
    while n > 0 {
        n -= 1;
        let ch = (b'A' + (n % 26) as u8) as char;
        s.insert(0, ch);
        n /= 26;
    }
    s
}

/// 格式化 f64 为 OOXML 兼容字符串（整数不输出小数点）。
fn format_f64(v: f64) -> String {
    if v.is_nan() || v.is_infinite() {
        return "0".to_string();
    }
    if v.fract() == 0.0 && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{}", v)
    }
}

/// 从 `<c:chartSpace>` XML 解析回 [`Chart`] 模型（TODO-004 读路径）。
///
/// 与 [`Chart::to_xml`] 对称的解析函数，用于 `Presentation::from_opc` 阶段
/// 读取 `/ppt/charts/chartN.xml` part 内容后还原 Chart 结构。
///
/// # 解析策略
///
/// - **chart_type**：根据 `<c:plotArea>` 下的图表元素名（`c:barChart` / `c:lineChart` /
///   `c:pieChart` / `c:scatterChart` / `c:areaChart` / `c:radarChart` / `c:bubbleChart`）
///   判断；对 `c:barChart` 额外读取 `<c:barDir val="col|bar">` 区分 Column / Bar。
/// - **title**：从 `<c:title><a:tx><a:rich>...<a:t>文本</a:t>` 提取第一个文本节点。
/// - **categories**：从**第一个** `<c:ser>` 的 `<c:cat><c:strRef><c:strCache>` 提取类别标签
///   （所有系列共享同一类别轴，取第一个即可；散点图/气泡图无 `c:cat`，类别列表为空）。
/// - **series**：遍历每个 `<c:ser>`，提取：
///   - `name`：`<c:tx><c:strRef><c:strCache><c:pt><c:v>`
///   - `values`：`<c:val><c:numRef><c:numCache><c:pt><c:v>`（Y 值）
///   - `x_values`：`<c:xVal><c:numRef><c:numCache><c:pt><c:v>`（仅散点图/气泡图）
///   - `bubble_sizes`：`<c:bubbleSize><c:numRef><c:numCache><c:pt><c:v>`（仅气泡图）
///
/// # 容错
///
/// - 缺失 `<c:plotArea>` 或其中无已知图表元素 → 返回 `Column` 占位类型（避免阻塞 round-trip）；
/// - 单个系列的 numCache/strCache 解析失败 → 该系列用空数据填充，不中断整体解析；
/// - 不解析 `<c:axId>` / `<c:catAx>` / `<c:valAx>` / `<c:legend>` 等装饰元素
///   （写出时由 [`Chart::to_xml`] 按图表类型自动重建）。
///
/// # 参数
/// - `xml`：完整的 `<c:chartSpace>` XML 字符串（chartN.xml 内容）。
///
/// # 返回值
/// - 成功：返回 [`Chart`]，`rid` 字段为空字符串（由调用方根据 slide rels 填充）；
/// - 失败：返回 `Error::Xml`，包含解析错误上下文。
///
/// # 示例
///
/// ```no_run
/// # use pptx::oxml::chart::{Chart, ChartType};
/// let xml = std::fs::read_to_string("chart1.xml").unwrap();
/// let chart = Chart::parse_from_xml(&xml).expect("parse chart");
/// println!("type: {:?}, series: {}", chart.chart_type, chart.data.series.len());
/// ```
pub fn parse_from_xml(xml: &str) -> crate::Result<Chart> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut chart_type = ChartType::Column;
    let mut title: Option<String> = None;
    let mut series: Vec<ChartSeries> = Vec::new();
    let mut categories: Vec<ChartCategory> = Vec::new();

    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    // 当前所在的图表元素名（如 c:barChart / c:lineChart），用于判断是否已进入 plotArea 的图表区域。
    let mut in_chart_elem: Option<&'static str> = None;
    // 是否在 <c:plotArea> 内（用于检测次坐标轴的 <c:crosses val="max"/> 标记，
    // 该标记位于 plotArea 下的 <c:valAx> 子元素内，而非主图表元素 barChart/lineChart 内）。
    let mut in_plot_area = false;
    // barChart 的 barDir 值（col/bar），用于区分 Column / Bar。
    let mut bar_dir: Option<String> = None;
    // 是否在 c:title 内的 a:t 元素中（用于精确提取标题文本，避免误取 a:defRPr 等元素的属性文本）。
    let mut in_title_text = false;
    // 当前正在解析的系列缓冲区。
    let mut cur_ser: Option<ChartSeries> = None;
    // 当前系列的字段上下文：正在读取的字段（tx/cat/val/xVal/yVal/bubbleSize）。
    let mut ser_field: Option<SerField> = None;
    // 是否在 numCache/strCache 内（决定 c:pt 的 c:v 是数值还是字符串）。
    let mut in_cache = false;
    // 嵌入式 Excel 工作簿的关系 id（从 <c:externalData r:id="..."> 提取）。
    let mut external_data_rid: Option<String> = None;
    // 临时缓冲：当前 cache 内的值列表。
    let mut cur_values: Vec<f64> = Vec::new();
    let mut cur_strings: Vec<String> = Vec::new();

    // ===== 数据标签（<c:dLbls>）解析上下文 =====
    //
    // dLbls 可出现在两个位置：
    //   1. 图表级：在 <c:barChart> 等图表元素内、<c:ser> 之前（plot 级配置）；
    //   2. 系列级：在 <c:ser> 内 <c:val>/<c:yVal> 之后（覆盖图表级）。
    //
    // 通过 `dl_target` 区分当前 dLbls 归属：进入 c:ser 之前是 Chart，进入后是 Series。
    let mut chart_dl: Option<DataLabels> = None;
    let mut series_dl: Option<DataLabels> = None;
    // dLbls 解析目标：None=不在 dLbls 内, Some(Chart)=图表级, Some(Series)=系列级。
    let mut dl_target: Option<DlTarget> = None;
    // 是否检测到次坐标轴（<c:crosses val="max"/> 是 PowerPoint 识别次轴的关键标记）。
    // 若解析到次轴定义，所有非散点/非饼图系列的 secondary_axis 字段置为 true。
    let mut has_secondary_axis = false;

    /// 系列字段枚举（用于 `ser_field` 上下文）。
    #[derive(Copy, Clone, Debug, Eq, PartialEq)]
    enum SerField {
        /// `<c:tx>` 系列名。
        Name,
        /// `<c:cat>` 类别标签（仅取第一个系列的）。
        Cat,
        /// `<c:val>` 数值（Y 坐标）。
        Val,
        /// `<c:xVal>` X 坐标（散点/气泡图）。
        XVal,
        /// `<c:yVal>` Y 坐标（散点/气泡图，与 Val 同语义但元素名不同）。
        YVal,
        /// `<c:bubbleSize>` 气泡尺寸。
        BubbleSize,
    }

    /// 数据标签归属（图表级 vs 系列级）。
    #[derive(Copy, Clone, Debug, Eq, PartialEq)]
    enum DlTarget {
        /// 图表级 dLbls（写在 <c:barChart> 等图表元素内、<c:ser> 之前）。
        Chart,
        /// 系列级 dLbls（写在 <c:ser> 内 <c:val>/<c:yVal> 之后，覆盖图表级）。
        Series,
    }

    /// 从 local name 字节判断是否为已知图表元素，返回 (元素名, 对应 ChartType)。
    fn chart_elem_for(local: &[u8]) -> Option<(&'static str, ChartType)> {
        match local {
            b"barChart" => Some(("barChart", ChartType::Column)), // Column/Bar 由 barDir 后续区分
            b"lineChart" => Some(("lineChart", ChartType::Line)),
            b"pieChart" => Some(("pieChart", ChartType::Pie)),
            b"scatterChart" => Some(("scatterChart", ChartType::Scatter)),
            b"areaChart" => Some(("areaChart", ChartType::Area)),
            b"radarChart" => Some(("radarChart", ChartType::Radar)),
            b"bubbleChart" => Some(("bubbleChart", ChartType::Bubble)),
            _ => None,
        }
    }

    /// 提取元素 local name（去命名空间前缀）。
    fn local_name_quick(name: &[u8]) -> &[u8] {
        match name.iter().position(|&b| b == b':') {
            Some(i) => &name[i + 1..],
            None => name,
        }
    }

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                // 提取 <c:externalData r:id="..."> 的关系 id（TODO-004 Excel 嵌入）。
                //
                // OOXML 中 externalData 通常是 Start + <c:autoUpdate/> + End 形式，
                // 这里在 Start 事件中提取 r:id 属性即可。
                if local == b"externalData" {
                    for a in e.attributes().flatten() {
                        // 兼容三种写法：
                        // - `r:id="..."`（标准 OOXML， relationships 命名空间前缀）
                        // - `ns0:id="..."` / `xxx:id="..."`（其他命名空间前缀 + id 本地名）
                        // - 裸 `id="..."`（无命名空间前缀，部分工具不严格加 r:）
                        let key = a.key.as_ref();
                        if key == b"r:id" || key == b"id" || key.ends_with(b":id") {
                            external_data_rid = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                }
                // 进入 <c:plotArea>：标记后续 crosses 检测生效
                if local == b"plotArea" {
                    in_plot_area = true;
                }
                // 进入图表元素（plotArea 下）
                if let Some((elem, ct)) = chart_elem_for(local) {
                    in_chart_elem = Some(elem);
                    chart_type = ct;
                    // barChart 需要后续 barDir 区分 Column/Bar，此处先占位 Column
                }
                // barDir 属性（区分 Column / Bar）
                if local == b"barDir" && in_chart_elem.is_some() {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            bar_dir = Some(v);
                        }
                    }
                }
                // 进入 c:title 内的 a:t 元素（精确匹配，避免误取 a:defRPr 等元素的属性文本）
                if local == b"t" {
                    in_title_text = true;
                }
                // 进入 c:ser
                if local == b"ser" && in_chart_elem.is_some() {
                    cur_ser = Some(ChartSeries::default());
                    // 重置系列级 dLbls 缓冲（每个系列独立）。
                    series_dl = None;
                }
                // 进入 c:dLbls：根据当前是否在 c:ser 内决定归属。
                //
                // - 在 c:ser 内 → 系列级（覆盖图表级）；
                // - 不在 c:ser 内但在图表元素内 → 图表级。
                if local == b"dLbls" {
                    if cur_ser.is_some() {
                        dl_target = Some(DlTarget::Series);
                        series_dl = Some(DataLabels::default());
                    } else if in_chart_elem.is_some() {
                        dl_target = Some(DlTarget::Chart);
                        chart_dl = Some(DataLabels::default());
                    }
                }
                // 系列字段上下文
                if cur_ser.is_some() {
                    match local {
                        b"tx" => ser_field = Some(SerField::Name),
                        b"cat" => ser_field = Some(SerField::Cat),
                        b"val" => ser_field = Some(SerField::Val),
                        b"xVal" => ser_field = Some(SerField::XVal),
                        b"yVal" => ser_field = Some(SerField::YVal),
                        b"bubbleSize" => ser_field = Some(SerField::BubbleSize),
                        _ => {}
                    }
                }
                // 进入 numCache / strCache
                if matches!(local, b"numCache" | b"strCache") {
                    in_cache = true;
                    cur_values.clear();
                    cur_strings.clear();
                }
                // 检测次坐标轴的 <c:crosses val="max"/> 标记。
                //
                // PowerPoint 通过该元素识别次坐标轴（显示在右侧）。
                // 解析到该标记后，所有非散点/非饼图系列的 secondary_axis 字段在
                // 解析末尾被置为 true（写出时由 to_xml 重建次轴定义）。
                //
                // 注意：crosses 元素位于 <c:valAx> 内（plotArea 下，但不在主图表元素
                // barChart/lineChart 内），因此用 in_plot_area 而非 in_chart_elem 判断。
                if local == b"crosses" && in_plot_area {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if v == "max" {
                                has_secondary_axis = true;
                            }
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                // barDir 自闭合形式
                if local == b"barDir" && in_chart_elem.is_some() {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            bar_dir = Some(v);
                        }
                    }
                }
                // c:pt 自闭合（无子元素，但通常 c:pt 是 Start + c:v + End）
                // c:v 自闭合形式（罕见但兼容）
                if local == b"v" && in_cache {
                    if let Some(v) = e.attributes().flatten().find(|a| a.key.as_ref() == b"val") {
                        let s = v
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .to_string();
                        if let Ok(f) = s.parse::<f64>() {
                            cur_values.push(f);
                        }
                        cur_strings.push(s);
                    }
                }
                // ===== 数据标签子元素解析（仅在 dLbls 上下文内） =====
                //
                // 这些子元素均以自闭合形式出现（<c:showVal val="1"/>），
                // 统一在 Empty 事件中处理。归属由 `dl_target` 决定。
                if let Some(target) = dl_target.as_ref() {
                    // 取目标 dLbls 缓冲的可变引用（图表级或系列级）。
                    let dl_buf: &mut Option<DataLabels> = if *target == DlTarget::Chart {
                        &mut chart_dl
                    } else {
                        &mut series_dl
                    };
                    let dl = dl_buf.get_or_insert_with(DataLabels::default);
                    match local {
                        b"showVal" => {
                            dl.show_val = parse_bool_val(&e);
                        }
                        b"showCatName" => {
                            dl.show_cat_name = parse_bool_val(&e);
                        }
                        b"showSerName" => {
                            dl.show_ser_name = parse_bool_val(&e);
                        }
                        b"showLegendKey" => {
                            dl.show_legend_key = parse_bool_val(&e);
                        }
                        b"showPercent" => {
                            dl.show_percent = parse_bool_val(&e);
                        }
                        b"showBubbleSize" => {
                            dl.show_bubble_size = parse_bool_val(&e);
                        }
                        b"dLblPos" => {
                            if let Some(v) = attr_val(&e, "val") {
                                dl.position = LabelPosition::parse(&v);
                            }
                        }
                        b"separator" => {
                            if let Some(v) = attr_val(&e, "val") {
                                dl.separator = Some(v);
                            }
                        }
                        b"numFmt" => {
                            if let Some(fc) = attr_val(&e, "formatCode") {
                                dl.num_fmt = Some(fc);
                            }
                        }
                        _ => {}
                    }
                }
                // c:crosses 自闭合形式（部分工具输出自闭合的 <c:crosses val="max"/>）。
                // 同 Start 分支：用 in_plot_area 判断（crosses 在 <c:valAx> 内，不在主图表元素内）。
                if local == b"crosses" && in_plot_area {
                    if let Some(v) = attr_val(&e, "val") {
                        if v == "max" {
                            has_secondary_axis = true;
                        }
                    }
                }
            }
            Ok(Event::Text(t)) => {
                // 标题文本（c:title 内的第一个 a:t 元素的文本内容）
                if in_title_text && title.is_none() {
                    // quick-xml 0.40: BytesText::unescape() 方法已移除，
                    // 改用 quick_xml::escape::unescape 函数（接受 &str）。
                    // BytesText 的 Deref 目标是 [u8]，需要先转成 &str。
                    let text_str = std::str::from_utf8(t.as_ref()).unwrap_or("");
                    let text = quick_xml::escape::unescape(text_str)
                        .unwrap_or_default()
                        .to_string();
                    if !text.is_empty() {
                        title = Some(text);
                    }
                }
                // cache 内的文本（c:v 的文本内容）
                if in_cache {
                    let text_str = std::str::from_utf8(t.as_ref()).unwrap_or("");
                    let text = quick_xml::escape::unescape(text_str)
                        .unwrap_or_default()
                        .to_string();
                    if !text.is_empty() {
                        if let Ok(f) = text.parse::<f64>() {
                            cur_values.push(f);
                        }
                        cur_strings.push(text);
                    }
                }
            }
            Ok(Event::End(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                // 离开 <c:plotArea>：清空 in_plot_area 标志
                if local == b"plotArea" {
                    in_plot_area = false;
                }
                // 离开图表元素
                if let Some(elem) = in_chart_elem {
                    if local == elem.as_bytes() {
                        // 最终确定 chart_type（barChart + barDir=bar → Bar）
                        if elem == "barChart" {
                            if let Some(dir) = &bar_dir {
                                if dir == "bar" {
                                    chart_type = ChartType::Bar;
                                }
                            }
                        }
                        in_chart_elem = None;
                    }
                }
                // 离开 a:t 元素
                if local == b"t" {
                    in_title_text = false;
                }
                // 离开 c:dLbls：清空 dl_target 上下文（缓冲已由子元素 Empty 事件填充）。
                if local == b"dLbls" {
                    dl_target = None;
                }
                // 离开 numCache/strCache：把缓冲写入当前字段
                if matches!(local, b"numCache" | b"strCache") && cur_ser.is_some() {
                    let field = ser_field.take();
                    if let Some(f) = field {
                        match f {
                            SerField::Name => {
                                // 系列名取第一个字符串
                                if let Some(s) = cur_strings.first() {
                                    if let Some(s_obj) = cur_ser.as_mut() {
                                        s_obj.name = s.clone();
                                    }
                                }
                                ser_field = None;
                            }
                            SerField::Cat => {
                                // 类别标签：仅第一个系列的 cat 写入 categories
                                if categories.is_empty() {
                                    for s in &cur_strings {
                                        categories.push(ChartCategory::new(s.clone()));
                                    }
                                }
                                ser_field = None;
                            }
                            SerField::Val | SerField::YVal => {
                                if let Some(s_obj) = cur_ser.as_mut() {
                                    s_obj.values = cur_values.clone();
                                }
                                ser_field = None;
                            }
                            SerField::XVal => {
                                if let Some(s_obj) = cur_ser.as_mut() {
                                    s_obj.x_values = Some(cur_values.clone());
                                }
                                ser_field = None;
                            }
                            SerField::BubbleSize => {
                                if let Some(s_obj) = cur_ser.as_mut() {
                                    s_obj.bubble_sizes = Some(cur_values.clone());
                                }
                                ser_field = None;
                            }
                        }
                    }
                    in_cache = false;
                    cur_values.clear();
                    cur_strings.clear();
                }
                // 离开 c:ser：把当前系列推入 series 列表，并写入系列级 dLbls。
                if local == b"ser" && cur_ser.is_some() {
                    if let Some(mut s_obj) = cur_ser.take() {
                        // 系列级数据标签：覆盖图表级配置。
                        if let Some(dl) = series_dl.take() {
                            if !dl.is_empty() {
                                s_obj.data_labels = Some(dl);
                            }
                        }
                        series.push(s_obj);
                    }
                    ser_field = None;
                    // 清理系列级 dLbls 上下文，避免泄漏到下一个系列。
                    series_dl = None;
                    if dl_target == Some(DlTarget::Series) {
                        dl_target = None;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("chart parse_from_xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    // 若解析到次坐标轴（<c:crosses val="max"/>），把所有非散点/非饼图系列的
    // `secondary_axis` 字段置为 true。这是简化策略：OOXML 没有显式的"系列-轴绑定"元素，
    // 实际绑定通过图表级 axId 列表 + 次轴 crosses=max 隐式表达。
    //
    // 对散点图/气泡图/饼图忽略（OOXML 规范约束：这些类型不支持次坐标轴）。
    if has_secondary_axis && !chart_type.is_xy_chart() && !matches!(chart_type, ChartType::Pie) {
        for s in series.iter_mut() {
            s.secondary_axis = true;
        }
    }

    let data = ChartData {
        categories,
        series,
        title,
        data_labels: chart_dl.filter(|dl| !dl.is_empty()),
    };
    Ok(Chart {
        chart_type,
        data,
        rid: String::new(),
        external_data_rid,
    })
}

/// 从 `<c:showXxx val="1|0|true|false"/>` 提取 bool 值。
///
/// OOXML 规范要求 `val` 属性值为 "0" 或 "1"（XML Schema boolean），
/// 但部分工具（如 LibreOffice）会输出 "true" / "false"，这里一并兼容。
fn parse_bool_val(e: &quick_xml::events::BytesStart<'_>) -> Option<bool> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"val" {
            let v = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default();
            return match v.as_ref() {
                "1" | "true" | "True" => Some(true),
                "0" | "false" | "False" => Some(false),
                _ => None,
            };
        }
    }
    None
}

/// 从元素提取指定属性值（如 `val` / `formatCode`），未找到返回 `None`。
///
/// 兼容带命名空间前缀的写法（如 `r:val`），用 local name 后缀匹配。
fn attr_val(e: &quick_xml::events::BytesStart<'_>, key: &str) -> Option<String> {
    let key_bytes = key.as_bytes();
    // 构造 ":key" 字节序列，用于后缀匹配（如匹配 "r:val" 中的 ":val"）。
    let suffix: Vec<u8> = std::iter::once(b':')
        .chain(key_bytes.iter().copied())
        .collect();
    for a in e.attributes().flatten() {
        let k = a.key.as_ref();
        if k == key_bytes || k.ends_with(&suffix) {
            return Some(
                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string(),
            );
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 柱状图最小 XML 包含必要元素。
    #[test]
    fn column_chart_minimal_xml() {
        let data = ChartData {
            categories: vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")],
            series: vec![ChartSeries::new("Sales", vec![10.0, 20.0])],
            title: Some("Revenue".to_string()),
            data_labels: None,
        };
        let chart = Chart::new(ChartType::Column, data);
        let xml = chart.to_xml();
        assert!(xml.contains("<c:chartSpace"), "xml: {}", xml);
        assert!(xml.contains("<c:barChart>"), "xml: {}", xml);
        assert!(xml.contains("barDir"), "xml: {}", xml);
        assert!(xml.contains("val=\"col\""), "xml: {}", xml);
        assert!(xml.contains("<c:ser>"), "xml: {}", xml);
        assert!(xml.contains("<c:catAx>"), "xml: {}", xml);
        assert!(xml.contains("<c:valAx>"), "xml: {}", xml);
        assert!(xml.contains("Revenue"), "xml: {}", xml);
    }

    /// 饼图不写出 catAx/valAx。
    #[test]
    fn pie_chart_no_axes() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
            series: vec![ChartSeries::new("X", vec![1.0, 2.0])],
            title: None,
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Pie, data).to_xml();
        assert!(xml.contains("<c:pieChart>"), "xml: {}", xml);
        assert!(!xml.contains("<c:catAx>"), "xml: {}", xml);
        assert!(!xml.contains("<c:valAx>"), "xml: {}", xml);
        assert!(xml.contains("varyColors"), "xml: {}", xml);
    }

    /// 折线图使用 lineChart 元素。
    #[test]
    fn line_chart_uses_line_chart_element() {
        let data = ChartData {
            categories: vec![ChartCategory::new("X1"), ChartCategory::new("X2")],
            series: vec![ChartSeries::new("Y", vec![1.5, 2.5])],
            title: None,
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Line, data).to_xml();
        assert!(xml.contains("<c:lineChart>"), "xml: {}", xml);
        assert!(!xml.contains("<c:barChart>"), "xml: {}", xml);
        // 浮点值应正确序列化
        assert!(xml.contains("1.5"), "xml: {}", xml);
        assert!(xml.contains("2.5"), "xml: {}", xml);
    }

    /// 散点图使用 scatterChart 元素 + c:xVal/c:yVal + 两个 valAx。
    #[test]
    fn scatter_chart_uses_scatter_elements() {
        let data = ChartData {
            categories: vec![], // 散点图忽略 categories
            series: vec![ChartSeries::new_scatter(
                "Series1",
                vec![1.0, 2.0, 3.0], // X
                vec![2.0, 4.0, 6.0], // Y
            )],
            title: Some("Scatter".to_string()),
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Scatter, data).to_xml();
        assert!(xml.contains("<c:scatterChart>"), "xml: {}", xml);
        assert!(xml.contains("scatterStyle"), "xml: {}", xml);
        assert!(xml.contains("<c:xVal>"), "xml: {}", xml);
        assert!(xml.contains("<c:yVal>"), "xml: {}", xml);
        // 散点图不应有 catAx
        assert!(!xml.contains("<c:catAx>"), "xml: {}", xml);
        // 散点图应有两个 valAx
        let val_ax_count = xml.matches("<c:valAx>").count();
        assert_eq!(
            val_ax_count, 2,
            "expected 2 valAx, got {}: {}",
            val_ax_count, xml
        );
        // 散点图不应有 c:cat
        assert!(!xml.contains("<c:cat>"), "xml: {}", xml);
        // X 坐标值应正确序列化
        assert!(xml.contains(">1<"), "xml: {}", xml);
        assert!(xml.contains(">3<"), "xml: {}", xml);
    }

    /// 面积图使用 areaChart 元素 + grouping=standard + catAx + valAx。
    #[test]
    fn area_chart_uses_area_elements() {
        let data = ChartData {
            categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
            series: vec![ChartSeries::new("S1", vec![10.0, 20.0])],
            title: None,
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Area, data).to_xml();
        assert!(xml.contains("<c:areaChart>"), "xml: {}", xml);
        assert!(xml.contains("grouping"), "xml: {}", xml);
        assert!(xml.contains("val=\"standard\""), "xml: {}", xml);
        assert!(xml.contains("<c:catAx>"), "xml: {}", xml);
        assert!(xml.contains("<c:valAx>"), "xml: {}", xml);
        // 面积图应有 c:cat（类别标签）
        assert!(xml.contains("<c:cat>"), "xml: {}", xml);
    }

    /// 雷达图使用 radarChart 元素 + radarStyle=marker + catAx + valAx。
    #[test]
    fn radar_chart_uses_radar_elements() {
        let data = ChartData {
            categories: vec![
                ChartCategory::new("速度"),
                ChartCategory::new("力量"),
                ChartCategory::new("技巧"),
            ],
            series: vec![ChartSeries::new("选手A", vec![8.0, 6.0, 9.0])],
            title: Some("Radar".to_string()),
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Radar, data).to_xml();
        assert!(xml.contains("<c:radarChart>"), "xml: {}", xml);
        assert!(xml.contains("radarStyle"), "xml: {}", xml);
        assert!(xml.contains("val=\"marker\""), "xml: {}", xml);
        // 雷达图使用 catAx + valAx（与柱/线相同）
        assert!(xml.contains("<c:catAx>"), "xml: {}", xml);
        assert!(xml.contains("<c:valAx>"), "xml: {}", xml);
        // 雷达图有 c:cat（类别标签）
        assert!(xml.contains("<c:cat>"), "xml: {}", xml);
        // 雷达图不应有散点图元素
        assert!(!xml.contains("<c:xVal>"), "xml: {}", xml);
        assert!(!xml.contains("<c:bubbleSize>"), "xml: {}", xml);
    }

    /// 气泡图使用 bubbleChart 元素 + bubbleScale + c:xVal/c:yVal/c:bubbleSize + 两个 valAx。
    #[test]
    fn bubble_chart_uses_bubble_elements() {
        let data = ChartData {
            categories: vec![],
            series: vec![ChartSeries::new_bubble(
                "S1",
                vec![1.0, 2.0, 3.0],    // x_values
                vec![10.0, 20.0, 30.0], // y_values
                vec![5.0, 15.0, 25.0],  // bubble_sizes
            )],
            title: Some("Bubble".to_string()),
            data_labels: None,
        };
        let xml = Chart::new(ChartType::Bubble, data).to_xml();
        assert!(xml.contains("<c:bubbleChart>"), "xml: {}", xml);
        assert!(xml.contains("bubbleScale"), "xml: {}", xml);
        assert!(xml.contains("val=\"100\""), "xml: {}", xml);
        // 气泡图使用 c:xVal / c:yVal / c:bubbleSize
        assert!(xml.contains("<c:xVal>"), "xml: {}", xml);
        assert!(xml.contains("<c:yVal>"), "xml: {}", xml);
        assert!(xml.contains("<c:bubbleSize>"), "xml: {}", xml);
        // 气泡图使用两个 valAx（无 catAx）
        assert!(!xml.contains("<c:catAx>"), "xml: {}", xml);
        assert!(xml.contains("<c:valAx>"), "xml: {}", xml);
        // 验证气泡尺寸值被正确写出
        assert!(xml.contains(">5<"), "xml: {}", xml);
        assert!(xml.contains(">25<"), "xml: {}", xml);
    }

    /// col_letter 正确转换。
    #[test]
    fn col_letter_basic() {
        assert_eq!(col_letter(1), "A");
        assert_eq!(col_letter(2), "B");
        assert_eq!(col_letter(26), "Z");
        assert_eq!(col_letter(27), "AA");
        assert_eq!(col_letter(28), "AB");
    }

    /// format_f64 整数无小数点。
    #[test]
    fn format_f64_integer() {
        assert_eq!(format_f64(10.0), "10");
        assert_eq!(format_f64(1.5), "1.5");
        assert_eq!(format_f64(f64::NAN), "0");
    }

    // ===================== TODO-004 图表读路径测试 =====================

    /// 柱状图 round-trip：to_xml → parse_from_xml → 字段一致。
    ///
    /// 覆盖：Column 类型识别（barDir=col）、标题、类别、单系列数值。
    #[test]
    fn parse_column_chart_round_trip() {
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![
                    ChartCategory::new("Q1"),
                    ChartCategory::new("Q2"),
                    ChartCategory::new("Q3"),
                ],
                series: vec![ChartSeries::new("Sales", vec![10.0, 20.0, 30.0])],
                title: Some("Revenue".to_string()),
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse column chart");

        assert_eq!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.data.title.as_deref(), Some("Revenue"));
        assert_eq!(parsed.data.categories.len(), 3);
        assert_eq!(parsed.data.categories[0].name, "Q1");
        assert_eq!(parsed.data.categories[2].name, "Q3");
        assert_eq!(parsed.data.series.len(), 1);
        assert_eq!(parsed.data.series[0].name, "Sales");
        assert_eq!(parsed.data.series[0].values, vec![10.0, 20.0, 30.0]);
    }

    /// 条形图 round-trip：barDir=bar 区分 Bar 与 Column。
    ///
    /// 这是 parse_from_xml 的关键边界：barChart 元素本身不区分柱/条，
    /// 必须依靠 barDir 属性后续修正 chart_type。
    #[test]
    fn parse_bar_chart_distinguishes_bar_from_column() {
        let original = Chart::new(
            ChartType::Bar,
            ChartData {
                categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
                series: vec![ChartSeries::new("S1", vec![1.0, 2.0])],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse bar chart");

        // 必须识别为 Bar 而非 Column
        assert_eq!(parsed.chart_type, ChartType::Bar);
        assert_ne!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.data.series.len(), 1);
        assert_eq!(parsed.data.series[0].values, vec![1.0, 2.0]);
    }

    /// 折线图 round-trip：lineChart 元素识别 + 浮点数保留。
    #[test]
    fn parse_line_chart_round_trip() {
        let original = Chart::new(
            ChartType::Line,
            ChartData {
                categories: vec![ChartCategory::new("X1"), ChartCategory::new("X2")],
                series: vec![ChartSeries::new("Y", vec![1.5, 2.5])],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse line chart");

        assert_eq!(parsed.chart_type, ChartType::Line);
        assert_eq!(parsed.data.categories.len(), 2);
        assert_eq!(parsed.data.series[0].name, "Y");
        // 浮点值必须正确解析
        assert_eq!(parsed.data.series[0].values, vec![1.5, 2.5]);
    }

    /// 饼图 round-trip：pieChart 元素识别 + categories 保留（饼图也写 c:cat）。
    #[test]
    fn parse_pie_chart_round_trip() {
        let original = Chart::new(
            ChartType::Pie,
            ChartData {
                categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
                series: vec![ChartSeries::new("X", vec![1.0, 2.0])],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse pie chart");

        assert_eq!(parsed.chart_type, ChartType::Pie);
        // 饼图 to_xml 也写 c:cat（line 654-669），所以解析后 categories 应保留
        assert_eq!(parsed.data.categories.len(), 2);
        assert_eq!(parsed.data.categories[0].name, "A");
        assert_eq!(parsed.data.categories[1].name, "B");
        assert_eq!(parsed.data.series.len(), 1);
        assert_eq!(parsed.data.series[0].values, vec![1.0, 2.0]);
    }

    /// 散点图 round-trip：scatterChart 元素 + c:xVal/c:yVal 解析。
    ///
    /// 散点图无 c:cat，使用 c:xVal / c:yVal 提供坐标。
    #[test]
    fn parse_scatter_chart_round_trip() {
        let original = Chart::new(
            ChartType::Scatter,
            ChartData {
                categories: vec![],
                series: vec![ChartSeries::new_scatter(
                    "Series1",
                    vec![1.0, 2.0, 3.0], // X
                    vec![2.0, 4.0, 6.0], // Y
                )],
                title: Some("Scatter".to_string()),
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse scatter chart");

        assert_eq!(parsed.chart_type, ChartType::Scatter);
        assert_eq!(parsed.data.title.as_deref(), Some("Scatter"));
        assert_eq!(parsed.data.series.len(), 1);
        let s = &parsed.data.series[0];
        assert_eq!(s.name, "Series1");
        // Y 坐标写入 values
        assert_eq!(s.values, vec![2.0, 4.0, 6.0]);
        // X 坐标写入 x_values
        assert_eq!(s.x_values.as_deref(), Some(&[1.0, 2.0, 3.0][..]));
        // 散点图无 bubble_sizes
        assert!(s.bubble_sizes.is_none());
    }

    /// 气泡图 round-trip：bubbleChart 元素 + c:xVal/c:yVal/c:bubbleSize 解析。
    #[test]
    fn parse_bubble_chart_round_trip() {
        let original = Chart::new(
            ChartType::Bubble,
            ChartData {
                categories: vec![],
                series: vec![ChartSeries::new_bubble(
                    "S1",
                    vec![1.0, 2.0, 3.0],    // x_values
                    vec![10.0, 20.0, 30.0], // y_values
                    vec![5.0, 15.0, 25.0],  // bubble_sizes
                )],
                title: Some("Bubble".to_string()),
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse bubble chart");

        assert_eq!(parsed.chart_type, ChartType::Bubble);
        assert_eq!(parsed.data.series.len(), 1);
        let s = &parsed.data.series[0];
        assert_eq!(s.name, "S1");
        assert_eq!(s.values, vec![10.0, 20.0, 30.0]);
        assert_eq!(s.x_values.as_deref(), Some(&[1.0, 2.0, 3.0][..]));
        assert_eq!(s.bubble_sizes.as_deref(), Some(&[5.0, 15.0, 25.0][..]));
    }

    /// 雷达图 round-trip：radarChart 元素识别 + 类别标签保留。
    #[test]
    fn parse_radar_chart_round_trip() {
        let original = Chart::new(
            ChartType::Radar,
            ChartData {
                categories: vec![
                    ChartCategory::new("速度"),
                    ChartCategory::new("力量"),
                    ChartCategory::new("技巧"),
                ],
                series: vec![ChartSeries::new("选手A", vec![8.0, 6.0, 9.0])],
                title: Some("Radar".to_string()),
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse radar chart");

        assert_eq!(parsed.chart_type, ChartType::Radar);
        assert_eq!(parsed.data.categories.len(), 3);
        assert_eq!(parsed.data.categories[0].name, "速度");
        assert_eq!(parsed.data.series[0].values, vec![8.0, 6.0, 9.0]);
    }

    /// 多系列柱状图 round-trip：验证系列顺序与名称。
    #[test]
    fn parse_multi_series_column_chart() {
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
                series: vec![
                    ChartSeries::new("S1", vec![1.0, 2.0]),
                    ChartSeries::new("S2", vec![3.0, 4.0]),
                    ChartSeries::new("S3", vec![5.0, 6.0]),
                ],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse multi-series chart");

        assert_eq!(parsed.data.series.len(), 3);
        assert_eq!(parsed.data.series[0].name, "S1");
        assert_eq!(parsed.data.series[1].name, "S2");
        assert_eq!(parsed.data.series[2].name, "S3");
        assert_eq!(parsed.data.series[2].values, vec![5.0, 6.0]);
    }

    /// 无标题图表 round-trip：title 字段为 None。
    #[test]
    fn parse_chart_without_title() {
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("A")],
                series: vec![ChartSeries::new("S", vec![1.0])],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse no-title chart");

        assert!(parsed.data.title.is_none());
    }

    /// 解析空 chartSpace（无 plotArea 子元素）应返回默认 Column 类型而非 panic。
    ///
    /// 验证零 panic 设计：解析失败不阻塞 round-trip。
    #[test]
    fn parse_empty_chart_space_no_panic() {
        let xml = "<?xml version=\"1.0\"?>\
                   <c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\
                     <c:chart></c:chart>\
                   </c:chartSpace>";
        let parsed = Chart::parse_from_xml(xml).expect("parse empty chartSpace");
        // 无 plotArea → chart_type 保持默认 Column
        assert_eq!(parsed.chart_type, ChartType::Column);
        assert!(parsed.data.series.is_empty());
        assert!(parsed.data.categories.is_empty());
    }

    /// 解析畸形 XML 应返回 Error::Xml 而非 panic。
    #[test]
    fn parse_malformed_xml_returns_error() {
        // 注意：quick-xml 0.40 SAX 流式解析器对未闭合标签较宽容（不跟踪标签栈），
        // 需要用真正畸形的 XML 才能触发解析错误。
        // 未闭合注释（`<!--` 后必须有 `-->`）是 quick-xml 0.40 必报错的场景。
        let xml = "<c:chartSpace><!-- unclosed comment <c:chart/></c:chartSpace>";
        let result = Chart::parse_from_xml(xml);
        assert!(result.is_err(), "malformed xml should error");
    }

    /// external_data_rid 写出：to_xml 在 `</c:chart>` 之后、`</c:chartSpace>` 之前
    /// 生成 `<c:externalData r:id="..."><c:autoUpdate val="0"/></c:externalData>`（TODO-004 Excel 嵌入）。
    #[test]
    fn to_xml_writes_external_data() {
        let mut chart = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("A")],
                series: vec![ChartSeries::new("S", vec![1.0])],
                title: None,
                data_labels: None,
            },
        );
        chart.external_data_rid = Some("rIdXlsx1".to_string());
        let xml = chart.to_xml();
        // 元素顺序：c:chart 关闭后紧跟 c:externalData
        assert!(
            xml.contains(r#"<c:externalData r:id="rIdXlsx1">"#),
            "应输出 c:externalData 开标签"
        );
        assert!(
            xml.contains(r#"<c:autoUpdate val="0"/>"#),
            "应输出 autoUpdate=0"
        );
        assert!(xml.contains("</c:externalData>"), "应闭合 c:externalData");
        // 验证顺序：c:chart 关闭在 c:externalData 之前
        let pos_chart_close = xml.find("</c:chart>").expect("c:chart 关闭存在");
        let pos_ext = xml.find("<c:externalData").expect("c:externalData 存在");
        assert!(
            pos_chart_close < pos_ext,
            "c:externalData 必须在 </c:chart> 之后"
        );
        let pos_chart_space_close = xml.rfind("</c:chartSpace>").expect("c:chartSpace 关闭存在");
        assert!(
            pos_ext < pos_chart_space_close,
            "c:externalData 必须在 </c:chartSpace> 之前"
        );
    }

    /// external_data_rid = None 时不输出 c:externalData（保持原有行为）。
    #[test]
    fn to_xml_omits_external_data_when_none() {
        let chart = Chart::new(
            ChartType::Line,
            ChartData {
                categories: vec![ChartCategory::new("A")],
                series: vec![ChartSeries::new("S", vec![1.0])],
                title: None,
                data_labels: None,
            },
        );
        let xml = chart.to_xml();
        assert!(
            !xml.contains("c:externalData"),
            "None 时不应输出 c:externalData"
        );
    }

    /// parse_from_xml 对称还原 external_data_rid（TODO-004 Excel 嵌入读路径）。
    #[test]
    fn parse_external_data_rid_round_trip() {
        let mut original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")],
                series: vec![ChartSeries::new("Sales", vec![10.0, 20.0])],
                title: Some("Revenue".into()),
                data_labels: None,
            },
        );
        original.external_data_rid = Some("rIdXlsx1".into());
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse round-trip");
        assert_eq!(parsed.chart_type, ChartType::Column);
        assert_eq!(parsed.external_data_rid, Some("rIdXlsx1".to_string()));
    }

    /// parse_from_xml 兼容 r:id 简写为 id 的写法（部分工具不严格加 r: 前缀）。
    #[test]
    fn parse_external_data_rid_bare_id_form() {
        let xml = "<?xml version=\"1.0\"?>\
                   <c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" \
                                 xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
                     <c:chart><c:plotArea><c:barChart><c:barDir val=\"col\"/>\
                       <c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>\
                         <c:val><c:numCache><c:formatCode>General</c:formatCode>\
                           <c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt>\
                         </c:numCache></c:val>\
                     </c:ser></c:barChart></c:plotArea></c:chart>\
                     <c:externalData id=\"rIdBare\"><c:autoUpdate val=\"0\"/></c:externalData>\
                   </c:chartSpace>";
        let parsed = Chart::parse_from_xml(xml).expect("parse bare id form");
        assert_eq!(parsed.external_data_rid, Some("rIdBare".to_string()));
    }

    // ===================== 数据标签（dLbls）测试 =====================

    /// `DataLabels::show_values()` 便捷构造器：仅 show_val=Some(true)。
    #[test]
    fn data_labels_show_values_constructor() {
        let dl = DataLabels::show_values();
        assert_eq!(dl.show_val, Some(true));
        assert_eq!(dl.show_cat_name, None);
        assert_eq!(dl.show_ser_name, None);
        assert!(!dl.is_empty());
    }

    /// `DataLabels::show_percent_pie()` 便捷构造器：仅 show_percent=Some(true)。
    #[test]
    fn data_labels_show_percent_pie_constructor() {
        let dl = DataLabels::show_percent_pie();
        assert_eq!(dl.show_percent, Some(true));
        assert_eq!(dl.show_val, None);
        assert!(!dl.is_empty());
    }

    /// `DataLabels::is_empty()` 对空配置返回 true。
    #[test]
    fn data_labels_is_empty_for_default() {
        let dl = DataLabels::default();
        assert!(dl.is_empty());
    }

    /// `LabelPosition` as_str / from_str 往返。
    #[test]
    fn label_position_round_trip() {
        for pos in [
            LabelPosition::BestFit,
            LabelPosition::Above,
            LabelPosition::Below,
            LabelPosition::Center,
            LabelPosition::InsideEnd,
            LabelPosition::OutsideEnd,
            LabelPosition::Left,
            LabelPosition::Right,
        ] {
            let s = pos.as_str();
            // 未知字符串返回 None（不应误识别）
            let parsed = LabelPosition::parse(s).expect("known position should parse");
            assert_eq!(parsed, pos, "position round-trip failed for {:?}", pos);
        }
        // 未知字符串返回 None
        assert!(LabelPosition::parse("unknown").is_none());
    }

    /// 图表级 `<c:dLbls>` round-trip：to_xml → parse_from_xml → 字段一致。
    ///
    /// 覆盖：show_val / show_cat_name / show_ser_name / show_legend_key / position / separator / numFmt。
    #[test]
    fn parse_chart_level_data_labels_round_trip() {
        let dl = DataLabels {
            show_val: Some(true),
            show_cat_name: Some(true),
            show_ser_name: Some(false),
            show_legend_key: Some(false),
            show_percent: None,
            show_bubble_size: None,
            position: Some(LabelPosition::OutsideEnd),
            separator: Some(", ".to_string()),
            num_fmt: Some("0.00".to_string()),
        };
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")],
                series: vec![ChartSeries::new("Sales", vec![10.0, 20.0])],
                title: None,
                data_labels: Some(dl.clone()),
            },
        );
        let xml = original.to_xml();
        // 验证 XML 包含 dLbls 子元素
        assert!(xml.contains("<c:dLbls>"), "xml: {}", xml);
        assert!(xml.contains(r#"<c:showVal val="1"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<c:showCatName val="1"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<c:showSerName val="0"/>"#), "xml: {}", xml);
        assert!(
            xml.contains(r#"<c:showLegendKey val="0"/>"#),
            "xml: {}",
            xml
        );
        assert!(xml.contains(r#"<c:dLblPos val="outEnd"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<c:separator val=", "/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"formatCode="0.00""#), "xml: {}", xml);

        // round-trip：解析后字段一致
        let parsed = Chart::parse_from_xml(&xml).expect("parse chart with dLbls");
        let parsed_dl = parsed
            .data
            .data_labels
            .expect("chart-level dLbls should parse");
        assert_eq!(parsed_dl.show_val, Some(true));
        assert_eq!(parsed_dl.show_cat_name, Some(true));
        assert_eq!(parsed_dl.show_ser_name, Some(false));
        assert_eq!(parsed_dl.show_legend_key, Some(false));
        assert_eq!(parsed_dl.show_percent, None);
        assert_eq!(parsed_dl.position, Some(LabelPosition::OutsideEnd));
        assert_eq!(parsed_dl.separator.as_deref(), Some(", "));
        assert_eq!(parsed_dl.num_fmt.as_deref(), Some("0.00"));
    }

    /// 系列级 `<c:dLbls>` round-trip：覆盖图表级配置。
    #[test]
    fn parse_series_level_data_labels_round_trip() {
        let series_dl = DataLabels {
            show_val: Some(true),
            show_percent: Some(true),
            position: Some(LabelPosition::InsideEnd),
            ..Default::default()
        };
        let original = Chart::new(
            ChartType::Pie,
            ChartData {
                categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
                series: vec![{
                    let mut s = ChartSeries::new("X", vec![1.0, 2.0]);
                    s.data_labels = Some(series_dl.clone());
                    s
                }],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        // 饼图 dLbls 应在 <c:ser> 内（系列级），而非图表级
        assert!(xml.contains("<c:dLbls>"), "xml: {}", xml);
        assert!(xml.contains(r#"<c:showVal val="1"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<c:showPercent val="1"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<c:dLblPos val="inEnd"/>"#), "xml: {}", xml);

        let parsed = Chart::parse_from_xml(&xml).expect("parse pie with series dLbls");
        assert_eq!(parsed.chart_type, ChartType::Pie);
        // 图表级 dLbls 应为 None（pie 未设置图表级）
        assert!(parsed.data.data_labels.is_none());
        // 系列级 dLbls 应正确解析
        let s = &parsed.data.series[0];
        let parsed_dl = s.data_labels.as_ref().expect("series dLbls should parse");
        assert_eq!(parsed_dl.show_val, Some(true));
        assert_eq!(parsed_dl.show_percent, Some(true));
        assert_eq!(parsed_dl.position, Some(LabelPosition::InsideEnd));
    }

    /// 空 `<c:dLbls/>`（无子元素）不应写入 data_labels 字段（保持 None）。
    #[test]
    fn parse_empty_dlbls_is_ignored() {
        // 构造一个图表级 dLbls 但所有字段都 None（is_empty=true）
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("A")],
                series: vec![ChartSeries::new("S", vec![1.0])],
                title: None,
                data_labels: Some(DataLabels::default()), // is_empty = true
            },
        );
        let xml = original.to_xml();
        // 空的 dLbls 不应输出（to_xml 检查 is_empty）
        assert!(
            !xml.contains("<c:dLbls>"),
            "empty dLbls should not be written: {}",
            xml
        );
        // 解析后 data_labels 也应为 None
        let parsed = Chart::parse_from_xml(&xml).expect("parse");
        assert!(parsed.data.data_labels.is_none());
    }

    // ===================== 次坐标轴测试 =====================

    /// `to_xml` 在 series.secondary_axis=true 时输出次坐标轴定义（含 crosses=max）。
    #[test]
    fn to_xml_writes_secondary_axis() {
        let s2 = {
            let mut s = ChartSeries::new("Secondary", vec![1.0, 2.0]);
            s.secondary_axis = true;
            s
        };
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")],
                series: vec![s2],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        // 应包含次轴 axId=444444444
        assert!(xml.contains(r#"<c:axId val="444444444"/>"#), "xml: {}", xml);
        // 应包含次轴定义（crosses=max）
        assert!(xml.contains(r#"<c:crosses val="max"/>"#), "xml: {}", xml);
        // 应有 2 个 valAx（主 valAx + 次 valAx）+ 1 个 catAx
        let val_ax_count = xml.matches("<c:valAx>").count();
        assert_eq!(
            val_ax_count, 2,
            "expected 2 valAx (primary + secondary): {}",
            xml
        );
        assert!(xml.contains("<c:catAx>"), "xml: {}", xml);
    }

    /// 次坐标轴 round-trip：parse_from_xml 解析 crosses=max 后，
    /// 所有非散点/非饼图系列的 secondary_axis 字段置为 true。
    #[test]
    fn parse_secondary_axis_round_trip() {
        let original = Chart::new(
            ChartType::Column,
            ChartData {
                categories: vec![ChartCategory::new("Q1"), ChartCategory::new("Q2")],
                series: vec![{
                    let mut s = ChartSeries::new("S", vec![10.0, 20.0]);
                    s.secondary_axis = true;
                    s
                }],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        let parsed = Chart::parse_from_xml(&xml).expect("parse secondary axis chart");
        // 解析后所有非散点/非饼图系列的 secondary_axis 应为 true
        assert_eq!(parsed.data.series.len(), 1);
        assert!(
            parsed.data.series[0].secondary_axis,
            "secondary_axis should be true after round-trip"
        );
    }

    /// 饼图不支持次坐标轴：secondary_axis 字段被忽略，不输出次轴定义。
    #[test]
    fn pie_chart_ignores_secondary_axis() {
        let original = Chart::new(
            ChartType::Pie,
            ChartData {
                categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
                series: vec![{
                    let mut s = ChartSeries::new("X", vec![1.0, 2.0]);
                    s.secondary_axis = true; // 应被忽略
                    s
                }],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        // 饼图不应输出次轴 axId
        assert!(
            !xml.contains("444444444"),
            "pie should not write secondary axId: {}",
            xml
        );
        assert!(
            !xml.contains(r#"<c:crosses val="max"/>"#),
            "pie should not write crosses=max: {}",
            xml
        );
    }

    /// 散点图不支持次坐标轴：secondary_axis 字段被忽略。
    #[test]
    fn scatter_chart_ignores_secondary_axis() {
        let mut s = ChartSeries::new_scatter("S", vec![1.0, 2.0], vec![3.0, 4.0]);
        s.secondary_axis = true; // 应被忽略
        let original = Chart::new(
            ChartType::Scatter,
            ChartData {
                categories: vec![],
                series: vec![s],
                title: None,
                data_labels: None,
            },
        );
        let xml = original.to_xml();
        assert!(
            !xml.contains("444444444"),
            "scatter should not write secondary axId: {}",
            xml
        );
        assert!(
            !xml.contains(r#"<c:crosses val="max"/>"#),
            "scatter should not write crosses=max: {}",
            xml
        );
    }
}
