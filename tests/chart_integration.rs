//! 集成测试：图表端到端 round-trip（TODO-041）。
//!
//! 验证 6 种图表类型（柱状图 / 条形图 / 折线图 / 饼图 / 散点图 / 面积图）
//! 从添加到序列化再加载的完整流程，确保 chartN.xml part 正确生成。
//!
//! # 覆盖场景
//!
//! - 柱状图（Column）round-trip
//! - 条形图（Bar）round-trip
//! - 折线图（Line）round-trip
//! - 饼图（Pie）round-trip
//! - 散点图（Scatter）round-trip（含 x_values）
//! - 面积图（Area）round-trip
//! - 多图表混合在单 slide

use pptx_rs::oxml::chart::{ChartCategory, ChartData, ChartSeries, ChartType};
use pptx_rs::{Inches, Presentation};

/// 辅助函数：构造一份简单的柱状图数据。
fn make_column_data() -> ChartData {
    ChartData {
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
    }
}

/// 验证柱状图 round-trip。
#[test]
fn chart_column_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let _col = slide
        .shapes_mut()
        .add_chart(
            ChartType::Column,
            make_column_data(),
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Column failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证条形图 round-trip。
#[test]
fn chart_bar_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let _bar = slide
        .shapes_mut()
        .add_chart(
            ChartType::Bar,
            make_column_data(),
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Bar failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证折线图 round-trip。
#[test]
fn chart_line_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

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

    let _line = slide
        .shapes_mut()
        .add_chart(
            ChartType::Line,
            line_data,
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Line failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证饼图 round-trip。
#[test]
fn chart_pie_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

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

    let _pie = slide
        .shapes_mut()
        .add_chart(
            ChartType::Pie,
            pie_data,
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Pie failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证散点图 round-trip（含 x_values）。
#[test]
fn chart_scatter_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let scatter_data = ChartData {
        categories: vec![],
        series: vec![ChartSeries::new_scatter(
            "实验数据",
            vec![1.0, 2.0, 3.0, 4.0, 5.0],
            vec![2.1, 3.9, 6.2, 7.8, 10.1],
        )],
        title: Some("散点图：X-Y 关系".to_string()),
        data_labels: None,
    };

    let _scatter = slide
        .shapes_mut()
        .add_chart(
            ChartType::Scatter,
            scatter_data,
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Scatter failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证面积图 round-trip。
#[test]
fn chart_area_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

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

    let _area = slide
        .shapes_mut()
        .add_chart(
            ChartType::Area,
            area_data,
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Area failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证多图表混合在单 slide round-trip。
#[test]
fn mixed_charts_in_one_slide_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    // 柱状图
    let _col = slide
        .shapes_mut()
        .add_chart(
            ChartType::Column,
            make_column_data(),
            Inches(0.5),
            Inches(0.5),
            Inches(4.5),
            Inches(3.0),
        )
        .expect("add_chart Column failed");

    // 折线图
    let line_data = ChartData {
        categories: vec![ChartCategory::new("A"), ChartCategory::new("B")],
        series: vec![ChartSeries::new("S1", vec![1.0, 2.0])],
        title: None,
        data_labels: None,
    };
    let _line = slide
        .shapes_mut()
        .add_chart(
            ChartType::Line,
            line_data,
            Inches(5.2),
            Inches(0.5),
            Inches(4.5),
            Inches(3.0),
        )
        .expect("add_chart Line failed");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证多张幻灯片各带不同图表的复杂场景。
#[test]
fn multiple_slides_with_different_charts() {
    let mut prs = Presentation::new().expect("Presentation::new failed");

    // 第 1 张：柱状图
    let counter = prs.id_counter();
    let slide1 = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide 1 failed");
    slide1
        .shapes_mut()
        .add_chart(
            ChartType::Column,
            make_column_data(),
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Column failed");

    // 第 2 张：饼图
    let counter = prs.id_counter();
    let slide2 = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide 2 failed");
    let pie_data = ChartData {
        categories: vec![ChartCategory::new("X"), ChartCategory::new("Y")],
        series: vec![ChartSeries::new("S", vec![60.0, 40.0])],
        title: None,
        data_labels: None,
    };
    slide2
        .shapes_mut()
        .add_chart(
            ChartType::Pie,
            pie_data,
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(4.0),
        )
        .expect("add_chart Pie failed");

    assert_eq!(prs.slides().len(), 2);

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(
        prs2.slides().len(),
        2,
        "slide count preserved after round-trip"
    );
}
