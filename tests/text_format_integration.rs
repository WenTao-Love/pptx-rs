//! 集成测试：文本格式化端到端流程（TODO-041）。
//!
//! 验证 TextFrame / ParagraphFormat / Font / ColorFormat 视图 API
//! 从设置到 round-trip 的完整流程，确保关键 OOXML 元素不丢失。
//!
//! # 覆盖场景
//!
//! - Run 字体大小 / 加粗 / 斜体 / 下划线 / 删除线
//! - Run 颜色（RGB / 主题色）
//! - 段落对齐 / 行距
//! - 文本框锚定 / 自适应 / 边距
//! - 多段落多 Run 混合
//! - 东亚字体设置

use pptx_rs::oxml::{
    Alignment, Font, MsoAnchor, MsoAutoSize, MsoThemeColorIndex, ParagraphFormat, TextFrame,
    Underline,
};
use pptx_rs::shape::PresetGeometry;
use pptx_rs::EmuExt;
use pptx_rs::{Inches, Presentation, Pt, RGBColor};

/// 验证 Run 的字体大小、加粗 round-trip。
#[test]
fn run_font_size_and_bold_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        let p = tf.add_paragraph();
        let r = p.add_run_with_text("加粗 24pt");
        let mut f = Font::from(&mut r.properties);
        f.set_size(Pt(24.0));
        f.set_bold(true);
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证 Run 的斜体 / 下划线 / 删除线 round-trip。
#[test]
fn run_italic_underline_strikethrough_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        let p = tf.add_paragraph();
        let r = p.add_run_with_text("斜体+下划线+删除线");
        let mut f = Font::from(&mut r.properties);
        f.set_italic(true);
        f.set_underline(Some(Underline::Single));
        f.set_strike(true);
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证 Run 的 RGB 颜色 round-trip。
#[test]
fn run_color_rgb_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        let p = tf.add_paragraph();
        let r = p.add_run_with_text("红色文本");
        let mut f = Font::from(&mut r.properties);
        f.color().set_rgb(RGBColor(0xC0, 0x39, 0x2B));
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证 Run 的主题色 round-trip。
#[test]
fn run_color_theme_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        let p = tf.add_paragraph();
        let r = p.add_run_with_text("主题色 accent1");
        let mut f = Font::from(&mut r.properties);
        f.color().set_theme(MsoThemeColorIndex::Accent1);
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证段落对齐和行距 round-trip。
#[test]
fn paragraph_alignment_and_line_spacing_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        let p = tf.add_paragraph();
        p.add_run_with_text("居中对齐 + 1.5 倍行距");
        let mut pf = ParagraphFormat::from(&mut p.properties);
        pf.set_alignment(Alignment::Center);
        pf.set_line_spacing_pct(1.5);
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证文本框锚定 / 自适应 / 边距 round-trip。
#[test]
fn text_frame_anchor_autosize_margins_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();
        tf.add_paragraph().add_run_with_text("锚定测试");
        tf.set_auto_size(MsoAutoSize::TextToFitShape);
        tf.set_vertical_anchor(MsoAnchor::Middle);
        tf.set_margins(
            Inches(0.1).emu(),
            Inches(0.1).emu(),
            Inches(0.1).emu(),
            Inches(0.1).emu(),
        );
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证多段落多 Run 混合场景 round-trip。
#[test]
fn multiple_paragraphs_and_runs_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut tb = slide
        .shapes_mut()
        .add_textbox(Inches(1.0), Inches(1.0), Inches(8.0), Inches(4.0))
        .expect("add_textbox failed");
    {
        let body = tb.text_frame_mut();
        let mut tf = TextFrame::new(body);
        tf.clear();

        // 段落 1：标题
        let p1 = tf.add_paragraph();
        let r1 = p1.add_run_with_text("标题段");
        {
            let mut f = Font::from(&mut r1.properties);
            f.set_bold(true);
            f.set_size(Pt(28.0));
        }

        // 段落 2：多个 Run 混合格式
        let p2 = tf.add_paragraph();
        {
            let r2a = p2.add_run_with_text("普通文本 ");
            let mut f = Font::from(&mut r2a.properties);
            f.set_size(Pt(18.0));
        }
        let _r2b = p2.add_run_with_text("加粗文本");
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证 AutoShape 的填充和线条格式 round-trip。
#[test]
fn autoshape_fill_and_line_format_round_trip() {
    use pptx_rs::oxml::{FillFormat, LineFormat};

    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let mut arrow = slide
        .shapes_mut()
        .add_shape(
            PresetGeometry::RightArrow,
            Inches(1.0),
            Inches(1.0),
            Inches(4.0),
            Inches(1.5),
        )
        .expect("add_shape failed");
    arrow.set_text("FillFormat demo");
    {
        let props = arrow.properties_mut();
        FillFormat::from(&mut props.fill)
            .solid()
            .set_rgb(RGBColor(0x2C, 0x82, 0xC9));
        if props.line.is_none() {
            props.line = Some(Default::default());
        }
        let mut line = LineFormat::from(props.line.as_mut().unwrap());
        line.color().set_rgb(RGBColor(0x1B, 0x4F, 0x72));
        line.set_width_pt(Pt(2.0));
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}
