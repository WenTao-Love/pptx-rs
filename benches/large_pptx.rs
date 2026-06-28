//! 性能基准测试：大型 PPTX 场景（TODO-040）。
//!
//! 建立 pptx-rs 在 100+ slides 场景下的性能基线，定位"大文件"瓶颈。
//! 运行方式：`cargo bench --bench large_pptx`。
//!
//! # 覆盖场景
//!
//! - 生成 100 张幻灯片并保存到 bytes
//! - 生成 500 张幻灯片并保存到 bytes
//! - 100 张幻灯片 round-trip（保存→读取→保存）
//!
//! # 设计原则
//!
//! - 每张幻灯片包含 1 个文本框 + 1 个自选形状，模拟真实使用场景；
//! - 使用 `black_box` 防止编译器优化；
//! - 500 slides 场景可能较慢，建议用 `--bench large_pptx -- --sample-size 10` 降低采样数。

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion};
use pptx::{Inches, Presentation, PresetGeometry};

/// 构造一个包含 `n` 张幻灯片的 Presentation，每张幻灯片含 1 文本框 + 1 矩形。
///
/// 用于为基准测试准备数据，本身不被直接测量。
fn build_presentation(n: usize) -> Presentation {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    for i in 0..n {
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
                &format!("幻灯片 #{}", i + 1),
            )
            .expect("add_textbox failed");
        // 矩形自选形状
        slide
            .shapes_mut()
            .add_shape(
                PresetGeometry::Rectangle,
                Inches(1.0),
                Inches(2.5),
                Inches(3.0),
                Inches(1.5),
            )
            .expect("add_shape failed");
    }
    prs
}

/// 测量"构造 + 保存"大型 PPTX 的开销。
///
/// 分别测试 100 / 500 张幻灯片，观察规模扩展性。
fn bench_save_large(c: &mut Criterion) {
    let mut group = c.benchmark_group("save_large_pptx");
    // 大文件场景降低采样数，避免基准测试运行过久
    group.sample_size(10);

    for n in [100usize, 500].iter() {
        group.bench_with_input(BenchmarkId::from_parameter(n), n, |b, &n| {
            b.iter(|| {
                let prs = build_presentation(n);
                let bytes = prs.to_bytes().expect("to_bytes failed");
                black_box(bytes);
            });
        });
    }
    group.finish();
}

/// 测量"仅保存"（已构造好 Presentation）的开销，隔离序列化 + zip 压缩成本。
fn bench_serialize_only(c: &mut Criterion) {
    let mut group = c.benchmark_group("serialize_only");
    group.sample_size(10);

    // 预先构造好 Presentation，bench 内只测 to_bytes
    for n in [100usize, 500].iter() {
        let prs = build_presentation(*n);
        group.bench_with_input(BenchmarkId::from_parameter(n), &prs, |b, prs| {
            b.iter(|| {
                let bytes = prs.to_bytes().expect("to_bytes failed");
                black_box(bytes);
            });
        });
    }
    group.finish();
}

/// 测量 100 张幻灯片 round-trip（保存→读取→保存）的开销。
///
/// 这个场景反映"打开大型 PPTX 并重新保存"的工作流性能，是真实用户最关心的指标。
fn bench_round_trip_large(c: &mut Criterion) {
    let mut group = c.benchmark_group("round_trip_large");
    group.sample_size(10);

    // 预先构造 100 张幻灯片并保存为 bytes
    let prs_100 = build_presentation(100);
    let bytes_100 = prs_100.to_bytes().expect("prepare to_bytes failed");

    group.bench_function("100_slides", |b| {
        b.iter(|| {
            // 1. 从 bytes 读取
            let prs = Presentation::load_bytes(&bytes_100).expect("load_bytes failed");
            // 2. 重新保存
            let bytes = prs.to_bytes().expect("second to_bytes failed");
            black_box(bytes);
        });
    });

    group.finish();
}

criterion_group!(
    benches,
    bench_save_large,
    bench_serialize_only,
    bench_round_trip_large,
);
criterion_main!(benches);
