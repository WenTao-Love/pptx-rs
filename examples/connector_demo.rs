//! 演示连接器（直线/折线/曲线）和带 `add_connector_geom` 的几何风格。
//!
//! 跑法（Rust 1.75+）：
//! ```text
//! cargo run --example connector_demo
//! ```

use pptx::oxml::LineFormat;
use pptx::oxml::MsoLineDashStyle;
use pptx::shape::{MsoConnectorType, Shape};
use pptx::Presentation;
use pptx::{Inches, Pt, RGBColor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut prs = Presentation::new()?;
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide_with_layout(counter, 0)?;

    // 1) 直线
    let mut c1 = slide.shapes_mut().add_connector(
        MsoConnectorType::Straight,
        Inches(1.0),
        Inches(1.0),
        Inches(8.0),
        Inches(1.0),
    )?;
    c1.set_name("Straight".to_string());
    {
        let props = c1.properties_mut();
        if props.line.is_none() {
            props.line = Some(Default::default());
        }
        let mut lf = LineFormat::from(props.line.as_mut().unwrap());
        lf.color().set_rgb(RGBColor(0x55, 0x55, 0x55));
        lf.set_width_pt(Pt(2.0));
    }

    // 2) 单折线
    let mut c2 = slide.shapes_mut().add_connector(
        MsoConnectorType::Elbow,
        Inches(1.0),
        Inches(2.0),
        Inches(8.0),
        Inches(2.0),
    )?;
    c2.set_name("Elbow".to_string());
    {
        let props = c2.properties_mut();
        if props.line.is_none() {
            props.line = Some(Default::default());
        }
        let mut lf = LineFormat::from(props.line.as_mut().unwrap());
        lf.color().set_rgb(RGBColor(0x2C, 0x82, 0xC9));
        lf.set_width_pt(Pt(1.5));
        lf.set_dash_style(MsoLineDashStyle::Dash);
    }

    // 3) 曲线
    let mut c3 = slide.shapes_mut().add_connector(
        MsoConnectorType::Curve,
        Inches(1.0),
        Inches(3.0),
        Inches(8.0),
        Inches(3.0),
    )?;
    c3.set_name("Curve".to_string());
    {
        let props = c3.properties_mut();
        if props.line.is_none() {
            props.line = Some(Default::default());
        }
        let mut lf = LineFormat::from(props.line.as_mut().unwrap());
        lf.color().set_rgb(RGBColor(0xED, 0x7D, 0x31));
        lf.set_width_pt(Pt(1.0));
        lf.set_dash_style(MsoLineDashStyle::DashDot);
    }

    // 4) 三段折线
    let mut c4 = slide.shapes_mut().add_connector(
        MsoConnectorType::BentConnector3,
        Inches(1.0),
        Inches(4.0),
        Inches(8.0),
        Inches(4.0),
    )?;
    c4.set_name("Bent3".to_string());
    {
        let props = c4.properties_mut();
        if props.line.is_none() {
            props.line = Some(Default::default());
        }
        let mut lf = LineFormat::from(props.line.as_mut().unwrap());
        lf.color().set_rgb(RGBColor(0xC0, 0x39, 0x2B));
        lf.set_width_pt(Pt(1.5));
        lf.set_dash_style(MsoLineDashStyle::Dot);
    }

    let out = std::env::current_dir()?.join("connector_demo.pptx");
    prs.save(&out)?;
    println!("saved -> {}", out.display());
    Ok(())
}
