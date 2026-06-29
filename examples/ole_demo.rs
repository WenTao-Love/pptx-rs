//! # 端到端示例：在幻灯片中嵌入 OLE 对象（TODO-043）
//!
//! 演示 `pptx-rs` 的 OLE 嵌入 API：
//!
//! 1. 新建演示文稿；
//! 2. 添加一张幻灯片；
//! 3. 生成一个示例 OLE 数据文件（简单的二进制 blob）；
//! 4. 通过 `add_ole_object` 把文件嵌入到幻灯片；
//! 5. 保存为 `ole_demo.pptx`。
//!
//! # 运行方式
//!
//! ```bash
//! cargo run --example ole_demo
//! ```
//!
//! # 预期输出
//!
//! - `ole_demo.pptx` 包含 1 张幻灯片；
//! - 幻灯片上有 1 个 OLE 对象（progId="Package"）；
//! - OLE 对象通过 `<p:oleObj r:id="rIdOle1"/>` 引用独立的 `/ppt/embeddings/oleObject1.bin` part；
//! - slide1.xml.rels 中含 `<Relationship Type=".../oleObject" Target="../embeddings/oleObject1.bin"/>`。
//!
//! # 注意
//!
//! 本示例为了演示 API 流程，使用一个简单的二进制 blob 作为 OLE 数据。
//! 实际使用时，应传入真实的 OLE 复合文档文件（如 `.xlsx` / `.docx` / `.pdf`），
//! 并设置对应的 `prog_id`（如 `"Excel.Sheet.12"` / `"Word.Document.12"` / `"Package"`）。
//! PowerPoint 双击 OLE 对象时会调用对应 OLE 服务器打开编辑。

use pptx_rs::{Inches, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // ---------- 1) 新建演示文稿 ----------
    let mut prs = Presentation::new()?;

    // ---------- 2) 添加幻灯片 ----------
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide(counter)?;

    // ---------- 3) 生成示例 OLE 数据文件 ----------
    // 实际使用时，应替换为真实的 OLE 文件路径（如 "data.xlsx"）。
    // 这里写一个简单的二进制文件作为演示数据。
    let ole_data_path = "ole_sample_data.bin";
    let sample_data = b"This is a sample OLE blob for pptx-rs demo.\n\
                       Replace this with a real OLE compound file (.xlsx/.docx) in production.\n\
                       progId=\"Package\" tells PowerPoint to use the shell Packager.";
    std::fs::write(ole_data_path, sample_data)?;

    // ---------- 4) 嵌入 OLE 对象 ----------
    // progId="Package" 是通用的 OLE 容器，适用于任意文件类型。
    // PowerPoint 双击时会调用 Windows shell 的 Packager 打开。
    let _ole = slide.shapes_mut().add_ole_object(
        ole_data_path,
        "Package",
        "Sample OLE Object",
        Inches(2.0),
        Inches(2.0),
        Inches(4.0),
        Inches(3.0),
    )?;

    // 可选：调整 OLE 对象属性
    // ole.set_show_as_icon(true);  // 默认就是 true
    // ole.set_icon_size(Inches(2.0).emu(), Inches(2.0).emu());

    // ---------- 5) 保存 ----------
    prs.save("ole_demo.pptx")?;

    // 清理临时文件
    let _ = std::fs::remove_file(ole_data_path);

    println!(
        "已生成 ole_demo.pptx（{} 张幻灯片，1 个 OLE 对象 progId=\"Package\"）",
        prs.slides().len()
    );
    println!("OLE 二进制数据写入 /ppt/embeddings/oleObject1.bin");
    println!("slide1.xml.rels 中含 oleObject 关系，slide1.xml 中含 <p:oleObj r:id=\"rIdOle1\"/>");

    Ok(())
}
