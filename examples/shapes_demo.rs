//! 演示对齐 python-pptx 主要 API 的一次性 e2e：AutoShape / TextBox / Picture /
//! Connector / Group / Table / Freeform / Notes。
//!
//! 输出一个能被 PowerPoint 与 WPS 打开的 .pptx。
//!
//! 跑法（Rust 1.75+）：
//! ```text
//! cargo run --example shapes_demo
//! ```

use pptx_rs::shape::MsoConnectorType;
use pptx_rs::shape::PresetGeometry;
use pptx_rs::{EmuExt, Inches, Pt, RGBColor};
use pptx_rs::{MsoShapeType, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1) 准备一个空 presentation
    let mut prs = Presentation::new()?;
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide_with_layout(counter, 6 /* blank */)?;

    // 2) AutoShape：矩形 + 五边形
    let mut rect = slide.shapes_mut().add_shape(
        PresetGeometry::Rectangle,
        Inches(1.0),
        Inches(1.0),
        Inches(2.0),
        Inches(1.0),
    )?;
    rect.set_text("rectangle");
    rect.set_fill_color(pptx_rs::Color::RGB(RGBColor(200, 50, 50)));
    rect.set_stroke_color(RGBColor::BLACK);
    rect.set_stroke_width(Pt(1.0).emu());

    let mut pent = slide.shapes_mut().add_shape(
        PresetGeometry::Pentagon,
        Inches(4.0),
        Inches(1.0),
        Inches(2.0),
        Inches(1.0),
    )?;
    pent.set_text("Step 1");
    pent.set_fill_color(pptx_rs::Color::RGB(RGBColor(50, 50, 200)));

    // 3) TextBox：多段多 Run
    let mut tb = slide.shapes_mut().add_textbox_with_text(
        Inches(1.0),
        Inches(3.0),
        Inches(5.0),
        Inches(1.5),
        "这是第一段",
    )?;
    {
        let p = tb.text_frame_mut().add_paragraph();
        p.add_run_with_text("这是第二段（加粗）").set_bold(true);
        p.add_line_break();
        p.add_run_with_text("软回车后继续");
        p.set_alignment(pptx_rs::Alignment::Center);
    }
    tb.set_word_wrap(true);

    // 4) Connector：折线 from (1,5) to (5,6)
    slide.shapes_mut().add_connector(
        MsoConnectorType::Elbow,
        Inches(1.0),
        Inches(5.0),
        Inches(5.0),
        Inches(6.0),
    )?;

    // 5) Picture（占位：用最小 1x1 红色 PNG，保存能打开）
    let png_bytes = minimal_png();
    let dir = std::env::temp_dir();
    let path = dir.join("__pptx_rs_demo_red.png");
    std::fs::write(&path, &png_bytes)?;
    slide
        .shapes_mut()
        .add_picture(&path, Inches(7.0), Inches(1.0), Inches(1.0), Inches(1.0))?;
    let _ = std::fs::remove_file(&path);

    // 6) Table：3 行 3 列
    let mut tbl =
        slide
            .shapes_mut()
            .add_table(3, 3, Inches(1.0), Inches(7.0), Inches(6.0), Inches(1.5))?;
    tbl.set_cell_text(0, 0, "A1")?;
    tbl.set_cell_text(0, 1, "B1")?;
    tbl.set_cell_text(0, 2, "C1")?;
    tbl.set_header_row(0, true)?;
    tbl.set_column_width(0, Inches(2.0).emu())?;
    tbl.set_column_width(1, Inches(2.0).emu())?;
    tbl.set_column_width(2, Inches(2.0).emu())?;
    tbl.set_row_height(0, Inches(0.5).emu())?;

    // 7) Notes
    slide.set_notes_text(Some("这是本张 slide 的演讲者备注\n第二行"));

    // 8) 输出
    let out = std::env::current_dir()?.join("shapes_demo.pptx");
    prs.save(&out)?;
    println!("saved -> {}", out.display());
    Ok(())
}

/// 生成最小 1×1 红色 PNG 字节。
///
/// 避免对外部 crate 的 image 依赖——本例只要有"能加载的图片"即可。
fn minimal_png() -> Vec<u8> {
    // 1x1 红色 PNG (8-bit, palette index)
    // Pre-encoded bytes
    vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77,
        0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT
        0x08, 0x99, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x5B, 0x39, 0xBA,
        0x9C, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND
        0xAE, 0x42, 0x60, 0x82,
    ]
}

/// 静默兼容：使用 MsoShapeType 但不强制调用
#[allow(dead_code)]
fn _ensure_mso_in_scope(t: MsoShapeType) -> &'static str {
    t.as_str()
}
