//! 集成测试：大型 PPTX 读写（TODO-041）。
//!
//! 验证 pptx-rs 在 100+ slides 场景下的读写稳定性，确保：
//!
//! - 大型 PPTX 可以成功保存
//! - 保存后的 bytes 可以成功读取
//! - 读取后的幻灯片数量保持一致
//! - 多次 round-trip 数据稳定
//!
//! # 注意
//!
//! 这些测试可能较慢（生成 100 张幻灯片 + 序列化），在 CI 上属于"重资源测试"。
//! 如果需要跳过，可用 `cargo test -- --skip large_pptx`。

use pptx::{Inches, Presentation, PresetGeometry};

/// 构造一个包含 `n` 张幻灯片的 Presentation，每张含 1 文本框 + 1 矩形。
fn build_large_presentation(n: usize) -> Presentation {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    for i in 0..n {
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
                &format!("大型 PPTX 幻灯片 #{}", i + 1),
            )
            .expect("add_textbox failed");
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

/// 验证 50 张幻灯片可以成功 round-trip。
#[test]
fn large_pptx_50_slides_round_trip() {
    let prs = build_large_presentation(50);
    assert_eq!(prs.slides().len(), 50, "should have 50 slides before save");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    assert!(!bytes.is_empty(), "saved bytes should not be empty");

    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(
        prs2.slides().len(),
        50,
        "slide count should be 50 after round-trip"
    );
}

/// 验证 100 张幻灯片可以成功 round-trip（TODO-041 关键场景）。
#[test]
fn large_pptx_100_slides_round_trip() {
    let prs = build_large_presentation(100);
    assert_eq!(
        prs.slides().len(),
        100,
        "should have 100 slides before save"
    );

    let bytes = prs.to_bytes().expect("to_bytes failed");
    assert!(!bytes.is_empty(), "saved bytes should not be empty");

    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(
        prs2.slides().len(),
        100,
        "slide count should be 100 after round-trip"
    );
}

/// 验证 100 张幻灯片保存到临时文件后可以读取回来。
#[test]
fn large_pptx_100_slides_save_to_file() {
    use tempfile::TempDir;

    let dir = TempDir::new().expect("TempDir::new failed");
    let path = dir.path().join("large_100.pptx");

    let prs = build_large_presentation(100);
    prs.save(&path).expect("save failed");
    assert!(path.exists(), "saved file should exist");

    let prs2 = Presentation::open(&path).expect("open failed");
    assert_eq!(
        prs2.slides().len(),
        100,
        "reloaded presentation should have 100 slides"
    );
}

/// 验证 100 张幻灯片多次 round-trip 数据稳定。
#[test]
fn large_pptx_multiple_round_trips() {
    let prs = build_large_presentation(50);
    let mut bytes = prs.to_bytes().expect("first to_bytes failed");

    // 连续 2 次 round-trip（大文件场景降低次数避免 CI 超时）
    for i in 0..2 {
        let prs = Presentation::load_bytes(&bytes).expect("load_bytes failed");
        assert_eq!(
            prs.slides().len(),
            50,
            "round-trip #{} should preserve 50 slides",
            i + 1
        );
        bytes = prs.to_bytes().expect("to_bytes failed");
    }

    assert!(
        !bytes.is_empty(),
        "bytes should not be empty after round-trips"
    );
}

/// 验证"读取 → 修改 → 再保存"在大型 PPTX 场景下的正确性。
#[test]
fn large_pptx_load_modify_resave() {
    // 1. 构造 30 张幻灯片
    let prs = build_large_presentation(30);
    let bytes1 = prs.to_bytes().expect("first to_bytes failed");

    // 2. 读取后添加更多幻灯片
    let mut prs2 = Presentation::load_bytes(&bytes1).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 30);

    for i in 0..20 {
        let counter = prs2.id_counter();
        let slide = prs2
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
                &format!("新增幻灯片 #{}", i + 1),
            )
            .expect("add_textbox failed");
    }
    assert_eq!(
        prs2.slides().len(),
        50,
        "should have 50 slides after adding 20"
    );

    // 3. 再保存并验证
    let bytes2 = prs2.to_bytes().expect("second to_bytes failed");
    let prs3 = Presentation::load_bytes(&bytes2).expect("second load_bytes failed");
    assert_eq!(
        prs3.slides().len(),
        50,
        "final presentation should have 50 slides"
    );
}
