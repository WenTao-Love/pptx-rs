//! 性能基准测试：基础保存场景（TODO-040）。
//!
//! 建立 pptx-rs 在常见场景下的性能基线，便于后续优化对比。
//! 运行方式：`cargo bench --bench save_pptx`。
//!
//! # 覆盖场景
//!
//! - 创建空 Presentation
//! - 保存空 Presentation 到 bytes
//! - 保存带 1 张幻灯片 + 文本框的 Presentation（hello_pptx 级别）
//! - 保存带多种形状的 Presentation（textbox + autoshape + table）
//! - 保存→读取→保存 round-trip
//!
//! # 设计原则
//!
//! - 每个 bench 函数都对应一个真实使用场景；
//! - `b.iter` 内部避免不必要的 clone，测量纯库开销；
//! - 使用 `black_box` 防止编译器优化掉结果。

use criterion::{black_box, criterion_group, criterion_main, Criterion};
use pptx_rs::{Inches, Presentation};

/// 测量 `Presentation::new()` 的开销（OPC 容器初始化 + 默认主题 + 默认母版）。
fn bench_new_presentation(c: &mut Criterion) {
    c.bench_function("new_presentation", |b| {
        b.iter(|| {
            let prs = Presentation::new().expect("Presentation::new failed");
            black_box(prs);
        });
    });
}

/// 测量保存空 Presentation 到 bytes 的开销（序列化 + zip 压缩）。
fn bench_save_empty(c: &mut Criterion) {
    c.bench_function("save_empty_to_bytes", |b| {
        let prs = Presentation::new().expect("Presentation::new failed");
        b.iter(|| {
            let bytes = prs.to_bytes().expect("to_bytes failed");
            black_box(bytes);
        });
    });
}

/// 测量 hello_pptx 级别的保存开销（1 张幻灯片 + 1 个文本框）。
fn bench_save_hello(c: &mut Criterion) {
    c.bench_function("save_hello_pptx", |b| {
        let mut prs = Presentation::new().expect("Presentation::new failed");
        let counter = prs.id_counter();
        let slide = prs
            .slides_mut()
            .add_slide(counter)
            .expect("add_slide failed");
        slide
            .shapes_mut()
            .add_textbox_with_text(
                Inches(1.0),
                Inches(1.0),
                Inches(8.0),
                Inches(1.0),
                "Hello, pptx-rs!",
            )
            .expect("add_textbox failed");
        b.iter(|| {
            let bytes = prs.to_bytes().expect("to_bytes failed");
            black_box(bytes);
        });
    });
}

/// 测量带多种形状的保存开销（1 张幻灯片 + 文本框 + 自选形状 + 表格）。
fn bench_save_with_shapes(c: &mut Criterion) {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");
    // 文本框
    slide
        .shapes_mut()
        .add_textbox_with_text(
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(1.0),
            "标题文本",
        )
        .expect("add_textbox failed");
    // 自选形状（矩形）
    slide
        .shapes_mut()
        .add_shape(
            pptx_rs::PresetGeometry::Rectangle,
            Inches(1.0),
            Inches(2.5),
            Inches(3.0),
            Inches(1.5),
        )
        .expect("add_shape failed");
    // 表格
    slide
        .shapes_mut()
        .add_table(4, 3, Inches(1.0), Inches(4.5), Inches(8.0), Inches(3.0))
        .expect("add_table failed");
    c.bench_function("save_with_shapes", |b| {
        b.iter(|| {
            let bytes = prs.to_bytes().expect("to_bytes failed");
            black_box(bytes);
        });
    });
}

/// 测量 round-trip 开销（保存→读取→保存）。
///
/// 这个场景反映"打开已有 PPTX 并重新保存"的典型工作流性能。
fn bench_round_trip(c: &mut Criterion) {
    // 先构造一个有内容的 Presentation 并保存到 bytes
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");
    slide
        .shapes_mut()
        .add_textbox_with_text(
            Inches(1.0),
            Inches(1.0),
            Inches(8.0),
            Inches(1.0),
            "round-trip 测试文本",
        )
        .expect("add_textbox failed");
    let original_bytes = prs.to_bytes().expect("first to_bytes failed");

    c.bench_function("round_trip_save_load_save", |b| {
        b.iter(|| {
            // 1. 从 bytes 读取
            let prs = Presentation::load_bytes(&original_bytes).expect("load_bytes failed");
            // 2. 重新保存
            let bytes = prs.to_bytes().expect("second to_bytes failed");
            black_box(bytes);
        });
    });
}

criterion_group!(
    benches,
    bench_new_presentation,
    bench_save_empty,
    bench_save_hello,
    bench_save_with_shapes,
    bench_round_trip,
);
criterion_main!(benches);
