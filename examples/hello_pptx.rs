//! # 端到端示例：创建并保存一个完整的 `.pptx` 文件
//!
//! 该示例演示了 `pptx-rs` 顶层 API 的核心用法：
//!
//! 1. 新建空白演示文稿；
//! 2. 添加多张幻灯片；
//! 3. 在幻灯片上添加文本框、自选形状；
//! 4. 设置文本样式（字号、加粗、颜色）；
//! 5. 把内存模型保存为本地 `.pptx` 文件。
//!
//! # 运行方式
//!
//! 在 `pptx-rs` 目录下执行：
//!
//! ```bash
//! cargo run --example hello_pptx
//! ```
//!
//! 运行后会在当前目录生成 `hello.pptx`（约 3~5 KB），可被 PowerPoint / WPS / Keynote 直接打开。
//!
//! # 预期输出
//!
//! - `hello.pptx` 包含 2 张幻灯片；
//! - 第 1 张幻灯片：1 个文本框（"Hello, pptx-rs!"，36pt 加粗蓝色）+ 1 个红色椭圆；
//! - 第 2 张幻灯片：1 个文本框（"第二张幻灯片"）。
//!
//! # 设计要点
//!
//! - **位置/尺寸单位**：用 [`pptx_rs::Inches`] 表示 1 英寸 = 914 400 EMU；
//!   也可用 [`pptx_rs::Cm`] / [`pptx_rs::Pt`] / [`pptx_rs::Emu`]。所有单位都实现了
//!   `EmuExt` trait，可在 API 边界自由混用。
//! - **id_counter**：`Presentation::id_counter()` 返回的 `Rc<Cell<u32>>` **必须**
//!   透传给 `slides_mut().add_slide(counter)`，否则所有 shape 共享同一 id（PowerPoint 会报"Invalid shape id"）。
//! - **可变性**：所有 mutation 走 `*_mut()` 入口，借用检查在编译期保证。

use pptx_rs::{Inches, Presentation, Pt, RGBColor};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // ---------- 1) 新建演示文稿 ----------
    // 此时内部会创建：默认母版 / 默认版式 / 默认主题。
    let mut prs = Presentation::new()?;

    // ---------- 2) 添加第一张幻灯片 ----------
    // 必须透传 id_counter，让后续 shape id 不会重复。
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide(counter)?;

    // ---------- 3) 在幻灯片上添加一个文本框（带文字） ----------
    // 位置/尺寸用 Inches；返回的 TextBox 可继续调整样式。
    let mut tb = slide.shapes_mut().add_textbox_with_text(
        Inches(1.0),
        Inches(1.0),
        Inches(8.0),
        Inches(1.0),
        "Hello, pptx-rs!",
    )?;

    // ---------- 4) 在文本框里设置字号、颜色 ----------
    // text_frame_mut() 拿到 p:txBody 的可变引用；进一步访问 paragraphs / runs。
    if let Some(p) = tb.text_frame_mut().paragraphs.first_mut() {
        if let Some(r) = p.runs.first_mut() {
            r.properties.size = Some(Pt(36.0));
            r.properties.bold = true;
            // RGBColor(r, g, b) 三个 0..=255 的分量；通过 .into() 转成 OOXML 的 a:srgbClr。
            r.properties.color = RGBColor(0x1F, 0x6F, 0xEB).into();
        }
    }

    // ---------- 5) 添加一个椭圆自选形状 ----------
    // PresetGeometry 是枚举，覆盖 OOXML 全部预设几何（rect/ellipse/arrow/...）。
    let mut shape = slide.shapes_mut().add_shape(
        pptx_rs::oxml::simpletypes::PresetGeometry::Ellipse,
        Inches(1.0),
        Inches(3.0),
        Inches(3.0),
        Inches(2.0),
    )?;
    // Fill::Solid 包一个 Color；此处用 RGBColor 的 From 实现直接 .into()。
    shape.set_fill(pptx_rs::oxml::sppr::Fill::Solid(
        RGBColor(0xE7, 0x4C, 0x3C).into(),
    ));

    // ---------- 6) 添加第二张幻灯片 + 文本框 ----------
    // 重新取 id_counter（或者沿用上面那个都行 —— Cell 内部 +1）。
    let counter2 = prs.id_counter();
    let slide2 = prs.slides_mut().add_slide(counter2)?;
    let tb2 = slide2.shapes_mut().add_textbox_with_text(
        Inches(2.0),
        Inches(2.0),
        Inches(6.0),
        Inches(1.0),
        "第二张幻灯片",
    )?;

    // 静默未使用变量警告
    let _ = tb2;

    // ---------- 7) 保存到本地 .pptx ----------
    // save() 内部：内存模型 → OpcPackage → zip → 写文件。
    prs.save("hello.pptx")?;
    println!("已生成 hello.pptx（包含 {} 张幻灯片）", prs.slides().len());

    Ok(())
}
