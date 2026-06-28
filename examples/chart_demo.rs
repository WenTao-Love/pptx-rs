//! # 端到端示例：在幻灯片中添加图表（TODO-004 基础图表 + 进阶图表支持）
//!
//! 演示 `pptx-rs` 的图表 API：
//!
//! 1. 新建演示文稿；
//! 2. 添加两张幻灯片；
//! 3. 第一张：柱状图、折线图、饼图各一个；
//! 4. 第二张：散点图、面积图各一个；
//! 5. 保存为 `chart_demo.pptx`。
//!
//! # 运行方式
//!
//! ```bash
//! cargo run --example chart_demo
//! ```
//!
//! # 预期输出
//!
//! - `chart_demo.pptx` 包含 2 张幻灯片；
//! - 第一张有 3 个图表（柱/线/饼），第二张有 2 个图表（散点/面积）；
//! - 每个图表通过 `<c:chart r:id="..."/>` 引用独立的 `/ppt/charts/chartN.xml` part；
//! - 数据通过 `<c:numCache>` 内嵌，PowerPoint / WPS 打开后可直接编辑数据。

use pptx::oxml::chart::{ChartCategory, ChartData, ChartSeries, ChartType};
use pptx::{Inches, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // ---------- 1) 新建演示文稿 ----------
    let mut prs = Presentation::new()?;

    // ---------- 2) 第一张幻灯片：柱/线/饼 ----------
    let counter = prs.id_counter();
    let slide1 = prs.slides_mut().add_slide(counter)?;

    // 添加柱状图
    let col_data = ChartData {
        categories: vec![
            ChartCategory::new("Q1"),
            ChartCategory::new("Q2"),
            ChartCategory::new("Q3"),
            ChartCategory::new("Q4"),
        ],
        series: vec![
            ChartSeries::new("Sales", vec![10.0, 20.0, 15.0, 25.0]),
            ChartSeries::new("Cost", vec![5.0, 8.0, 7.0, 10.0]),
        ],
        title: Some("季度销售对比".to_string()),
        data_labels: None,
    };
    let _col = slide1.shapes_mut().add_chart(
        ChartType::Column,
        col_data,
        Inches(0.5),
        Inches(0.5),
        Inches(4.5),
        Inches(3.0),
    )?;

    // 添加折线图
    let line_data = ChartData {
        categories: vec![
            ChartCategory::new("Jan"),
            ChartCategory::new("Feb"),
            ChartCategory::new("Mar"),
            ChartCategory::new("Apr"),
            ChartCategory::new("May"),
        ],
        series: vec![ChartSeries::new(
            "Temperature",
            vec![5.0, 8.0, 12.0, 18.0, 22.0],
        )],
        title: Some("月度温度趋势".to_string()),
        data_labels: None,
    };
    let _line = slide1.shapes_mut().add_chart(
        ChartType::Line,
        line_data,
        Inches(5.2),
        Inches(0.5),
        Inches(4.5),
        Inches(3.0),
    )?;

    // 添加饼图
    let pie_data = ChartData {
        categories: vec![
            ChartCategory::new("Product A"),
            ChartCategory::new("Product B"),
            ChartCategory::new("Product C"),
        ],
        series: vec![ChartSeries::new("Share", vec![40.0, 35.0, 25.0])],
        title: Some("市场份额".to_string()),
        data_labels: None,
    };
    let _pie = slide1.shapes_mut().add_chart(
        ChartType::Pie,
        pie_data,
        Inches(2.5),
        Inches(3.8),
        Inches(5.0),
        Inches(3.5),
    )?;

    // ---------- 3) 第二张幻灯片：散点/面积 ----------
    let counter = prs.id_counter();
    let slide2 = prs.slides_mut().add_slide(counter)?;

    // 添加散点图（X-Y 坐标对）
    // 散点图忽略 categories，X 坐标由 series.x_values 提供
    let scatter_data = ChartData {
        categories: vec![],
        series: vec![ChartSeries::new_scatter(
            "实验数据",
            vec![1.0, 2.0, 3.0, 4.0, 5.0],  // X
            vec![2.1, 3.9, 6.2, 7.8, 10.1], // Y
        )],
        title: Some("散点图：X-Y 关系".to_string()),
        data_labels: None,
    };
    let _scatter = slide2.shapes_mut().add_chart(
        ChartType::Scatter,
        scatter_data,
        Inches(0.5),
        Inches(0.5),
        Inches(4.5),
        Inches(3.0),
    )?;

    // 添加面积图
    let area_data = ChartData {
        categories: vec![
            ChartCategory::new("2019"),
            ChartCategory::new("2020"),
            ChartCategory::new("2021"),
            ChartCategory::new("2022"),
            ChartCategory::new("2023"),
        ],
        series: vec![
            ChartSeries::new("Revenue", vec![100.0, 120.0, 150.0, 180.0, 210.0]),
            ChartSeries::new("Profit", vec![20.0, 28.0, 40.0, 55.0, 70.0]),
        ],
        title: Some("年度营收面积图".to_string()),
        data_labels: None,
    };
    let _area = slide2.shapes_mut().add_chart(
        ChartType::Area,
        area_data,
        Inches(5.2),
        Inches(0.5),
        Inches(4.5),
        Inches(3.0),
    )?;

    // ---------- 4) 保存 ----------
    prs.save("chart_demo.pptx")?;
    println!(
        "已生成 chart_demo.pptx（{} 张幻灯片，5 个图表：柱/线/饼/散点/面积）",
        prs.slides().len()
    );

    Ok(())
}
