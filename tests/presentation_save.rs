//! 集成测试：Presentation 全流程（TODO-041）。
//!
//! 验证 Presentation 的"新建 → 添加内容 → 保存 → 读取 → 验证 → 修改 → 再保存"
//! 完整工作流，确保 round-trip 不丢失数据。
//!
//! # 覆盖场景
//!
//! - 新建空 Presentation 并保存到 bytes
//! - 从 bytes 读取 Presentation 并验证结构
//! - 添加幻灯片 + 文本框后 round-trip
//! - 多张幻灯片 round-trip（验证 slide 数量保持）
//! - 保存到临时文件并读取

use pptx_rs::{Inches, Presentation};
use tempfile::TempDir;

/// 验证新建的空 Presentation 可以成功保存到 bytes。
#[test]
fn new_presentation_saves_to_bytes() {
    let prs = Presentation::new().expect("Presentation::new failed");
    let bytes = prs.to_bytes().expect("to_bytes failed");
    // 一个最小的 .pptx 至少包含 [Content_Types].xml + presentation.xml + 默认主题
    assert!(!bytes.is_empty(), "saved bytes should not be empty");
    // 验证是合法的 zip（PK 头）
    assert_eq!(
        &bytes[0..2],
        b"PK",
        "saved bytes should start with PK zip signature"
    );
}

/// 验证从 bytes 读取 Presentation 后，幻灯片数量保持一致。
#[test]
fn round_trip_preserves_slide_count() {
    // 1. 构造一个有 3 张幻灯片的 Presentation
    let mut prs = Presentation::new().expect("Presentation::new failed");
    for i in 0..3 {
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
                &format!("Slide {}", i + 1),
            )
            .expect("add_textbox failed");
    }
    assert_eq!(prs.slides().len(), 3, "should have 3 slides before save");

    // 2. 保存到 bytes
    let bytes = prs.to_bytes().expect("to_bytes failed");

    // 3. 从 bytes 读取
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");

    // 4. 验证幻灯片数量保持
    assert_eq!(
        prs2.slides().len(),
        3,
        "slide count should be preserved after round-trip"
    );
}

/// 验证保存到临时文件后可以读取回来。
#[test]
fn save_to_temp_file_and_reload() {
    let dir = TempDir::new().expect("TempDir::new failed");
    let path = dir.path().join("test_presentation.pptx");

    // 1. 新建并保存
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
            "临时文件测试",
        )
        .expect("add_textbox failed");
    prs.save(&path).expect("save failed");
    assert!(path.exists(), "saved file should exist");

    // 2. 从文件读取
    let prs2 = Presentation::open(&path).expect("open failed");
    assert_eq!(
        prs2.slides().len(),
        1,
        "reloaded presentation should have 1 slide"
    );
}

/// 验证"读取 → 修改 → 再保存"工作流。
#[test]
fn load_modify_resave_workflow() {
    // 1. 构造初始 Presentation（1 张幻灯片）
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let _ = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");
    let bytes1 = prs.to_bytes().expect("first to_bytes failed");
    assert_eq!(prs.slides().len(), 1);

    // 2. 读取后添加第二张幻灯片
    let mut prs2 = Presentation::load_bytes(&bytes1).expect("load_bytes failed");
    assert_eq!(
        prs2.slides().len(),
        1,
        "loaded presentation should have 1 slide"
    );
    let counter2 = prs2.id_counter();
    let _ = prs2
        .slides_mut()
        .add_slide(counter2)
        .expect("add_slide failed");
    assert_eq!(
        prs2.slides().len(),
        2,
        "should have 2 slides after adding one"
    );

    // 3. 再保存
    let bytes2 = prs2.to_bytes().expect("second to_bytes failed");
    assert!(!bytes2.is_empty());

    // 4. 再读取验证
    let prs3 = Presentation::load_bytes(&bytes2).expect("second load_bytes failed");
    assert_eq!(
        prs3.slides().len(),
        2,
        "final presentation should have 2 slides"
    );
}

/// 验证空 Presentation round-trip 不丢失基本结构。
#[test]
fn empty_presentation_round_trip() {
    let prs = Presentation::new().expect("Presentation::new failed");
    let bytes = prs.to_bytes().expect("to_bytes failed");

    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    // 空 Presentation 应该有 0 张幻灯片
    assert_eq!(
        prs2.slides().len(),
        0,
        "empty presentation should have 0 slides after round-trip"
    );
}

/// 验证连续多次 round-trip 数据稳定性（save→load→save→load→save）。
#[test]
fn multiple_round_trips_are_stable() {
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
            "多重 round-trip 稳定性测试",
        )
        .expect("add_textbox failed");

    let mut bytes = prs.to_bytes().expect("first to_bytes failed");

    // 连续 3 次 round-trip
    for i in 0..3 {
        let prs = Presentation::load_bytes(&bytes).expect("load_bytes failed");
        assert_eq!(
            prs.slides().len(),
            1,
            "round-trip #{} should preserve 1 slide",
            i + 1
        );
        bytes = prs.to_bytes().expect("to_bytes failed");
    }

    assert!(
        !bytes.is_empty(),
        "bytes should not be empty after 3 round-trips"
    );
}
