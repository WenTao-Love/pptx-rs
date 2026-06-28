//! 演示对齐 python-pptx `text.text` 风格 API：TextFrame / ParagraphFormat / Font / ColorFormat。
//!
//! 跑法（Rust 1.75+）：
//! ```text
//! cargo run --example text_format_demo
//! ```

use pptx::oxml::Alignment;
use pptx::oxml::ColorFormat;
use pptx::oxml::FillFormat;
use pptx::oxml::Font;
use pptx::oxml::LineFormat;
use pptx::oxml::MsoAnchor;
use pptx::oxml::MsoAutoSize;
use pptx::oxml::ParagraphFormat;
use pptx::oxml::TextFrame;
use pptx::shape::MsoConnectorType;
use pptx::shape::PresetGeometry;
use pptx::Presentation;
use pptx::{EmuExt, Inches, Pt, RGBColor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 1) 新建空演示文稿
    let mut prs = Presentation::new()?;
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide_with_layout(counter, 0)?;

    // 2) 加一个文本框并走 TextFrame 视图
    let mut tb =
        slide
            .shapes_mut()
            .add_textbox(Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.5))?;
    {
        let body = tb.text_frame_mut();
        // TextFrame 直接借用 TextBody
        let mut tf = TextFrame::new(body);
        // 清空 + 走 view 加段、加 Run
        tf.clear();
        let p1 = tf.add_paragraph();
        p1.add_run_with_text("标题：pptx-rs TextFrame 演示")
            .set_bold(true);
        // 用 view 设对齐
        let mut pf = ParagraphFormat::from(&mut p1.properties);
        pf.set_alignment(Alignment::Center);
        pf.set_line_spacing_pct(1.2);

        let p2 = tf.add_paragraph();
        let r2 = p2.add_run_with_text("第二段：彩色 + 斜体 + 24pt");
        {
            let mut f = Font::from(&mut r2.properties);
            f.set_size(Pt(24.0));
            f.set_italic(true);
            f.color().set_rgb(RGBColor(0xC0, 0x39, 0x2B));
        }
        // 第三段：主题色
        let p3 = tf.add_paragraph();
        let r3 = p3.add_run_with_text("第三段：主题色 accent1 + 下划线");
        {
            let mut f = Font::from(&mut r3.properties);
            f.set_underline(Some(pptx::oxml::Underline::Single));
            f.color().set_theme(pptx::oxml::MsoThemeColorIndex::Accent1);
        }
        // 文本框级属性（autofit / 垂直对齐）
        tf.set_auto_size(MsoAutoSize::TextToFitShape);
        tf.set_vertical_anchor(MsoAnchor::Middle);
        tf.set_margins(
            Inches(0.1).emu(),
            Inches(0.1).emu(),
            Inches(0.1).emu(),
            Inches(0.1).emu(),
        );
    }

    // 3) AutoShape：演示 FillFormat / LineFormat
    let mut arrow = slide.shapes_mut().add_shape(
        PresetGeometry::RightArrow,
        Inches(1.0),
        Inches(4.0),
        Inches(2.5),
        Inches(1.0),
    )?;
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
        line.set_dash_style(pptx::oxml::MsoLineDashStyle::Dash);
    }

    // 4) 加一个 connector，让文件不是空
    slide.shapes_mut().add_connector(
        MsoConnectorType::Straight,
        Inches(1.0),
        Inches(6.0),
        Inches(8.0),
        Inches(6.0),
    )?;

    // 5) 写入文件
    let out = std::env::current_dir()?.join("text_format_demo.pptx");
    prs.save(&out)?;
    println!("saved -> {}", out.display());

    // 6) 验证 ColorFormat 借用语义（只在 release 也可调）
    #[allow(unused_variables)]
    let mut c: pptx::oxml::color::Color = RGBColor(0x10, 0x20, 0x30).into();
    {
        let mut cf = ColorFormat::new(&mut c);
        cf.set_theme(pptx::oxml::MsoThemeColorIndex::Accent3);
    }
    assert!(matches!(c, pptx::oxml::color::Color::Scheme(_)));

    Ok(())
}
