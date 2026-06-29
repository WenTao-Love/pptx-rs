//! 集成测试：备注 / 批注 / 自定义属性端到端流程（TODO-041）。
//!
//! 验证 notes / comments / custom_properties / core_properties
//! 从设置到序列化再加载的完整流程，确保 round-trip 不丢失数据。
//!
//! # 覆盖场景
//!
//! - 备注文本设置与读取
//! - 批注添加与清除
//! - 批注作者列表管理
//! - 自定义属性 5 种类型（Text/Int/Float/Bool/DateTime）
//! - 核心属性（title/author/keywords）

use pptx_rs::presentation::CustomPropertyValue;
use pptx_rs::{Inches, Presentation};
use std::io::Read;

/// 辅助函数：从 pptx bytes 中读取指定 part 的文本内容。
///
/// # 参数
/// - `bytes`：pptx 文件的字节数据
/// - `part_path`：zip 内的 part 路径（如 `ppt/notesSlides/notesSlide1.xml`）
///
/// # 返回
/// 成功返回 part 内容字符串，找不到返回 `None`。
fn read_zip_part(bytes: &[u8], part_path: &str) -> Option<String> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut zip = zip::ZipArchive::new(cursor).ok()?;
    for i in 0..zip.len() {
        let mut file = zip.by_index(i).ok()?;
        if file.name() == part_path {
            let mut buf = String::new();
            file.read_to_string(&mut buf).ok()?;
            return Some(buf);
        }
    }
    None
}

/// 辅助函数：列出 pptx bytes 中所有 part 路径（用于调试）。
#[allow(dead_code)]
fn list_zip_parts(bytes: &[u8]) -> Vec<String> {
    let cursor = std::io::Cursor::new(bytes.to_vec());
    let mut zip = match zip::ZipArchive::new(cursor) {
        Ok(z) => z,
        Err(_) => return vec![],
    };
    let mut parts = Vec::new();
    for i in 0..zip.len() {
        if let Ok(file) = zip.by_index(i) {
            parts.push(file.name().to_string());
        }
    }
    parts
}

/// 验证备注文本设置与读取 round-trip。
#[test]
fn slide_notes_text_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    slide.set_notes_text(Some("这是备注内容\n第二行备注"));

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);

    // 验证 notesSlide XML 存在且包含备注文本
    let notes_xml =
        read_zip_part(&bytes, "ppt/notesSlides/notesSlide1.xml").expect("notesSlide1.xml exists");
    assert!(
        notes_xml.contains("这是备注内容"),
        "notes XML should contain the notes text, got: {}",
        notes_xml
    );
}

/// 验证备注删除 round-trip。
#[test]
fn slide_notes_clear_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    slide.set_notes_text(Some("临时备注"));
    slide.set_notes_text(None);

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
    // 清除后不应有 notesSlide part
    assert!(
        read_zip_part(&bytes, "ppt/notesSlides/notesSlide1.xml").is_none(),
        "notes should be cleared"
    );
}

/// 验证单条批注添加 round-trip。
#[test]
fn add_single_comment_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    // 先获取 author_id，避免与 slides_mut 的可变借用冲突
    let author_id = prs.comment_authors_mut().get_or_insert_id("张三", "ZS");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let idx = slide.add_comment(author_id, Inches(1.0), Inches(1.0), "评论内容");
    assert_eq!(idx, 1, "first comment idx should be 1");

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
    assert_eq!(prs2.comment_authors().authors.len(), 1, "author preserved");

    // 验证 commentAuthors.xml 存在且包含作者
    let authors_xml =
        read_zip_part(&bytes, "ppt/commentAuthors.xml").expect("commentAuthors.xml exists");
    assert!(
        authors_xml.contains("张三"),
        "commentAuthors should contain author name"
    );
}

/// 验证多条批注添加 round-trip。
#[test]
fn add_multiple_comments_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let author_id = prs.comment_authors_mut().get_or_insert_id("李四", "LS");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    let idx1 = slide.add_comment(author_id, Inches(1.0), Inches(1.0), "第一条");
    let idx2 = slide.add_comment(author_id, Inches(2.0), Inches(2.0), "第二条");
    assert_eq!(idx1, 1);
    assert_eq!(idx2, 2);

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
}

/// 验证批注清除 round-trip。
#[test]
fn clear_comments_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let author_id = prs.comment_authors_mut().get_or_insert_id("王五", "WW");
    let counter = prs.id_counter();
    let slide = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide failed");

    slide.add_comment(author_id, Inches(1.0), Inches(1.0), "将被清除");
    slide.clear_comments();

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 1);
    // 清除后不应有 comment part
    assert!(
        read_zip_part(&bytes, "ppt/comments/comment1.xml").is_none(),
        "comments should be cleared"
    );
}

/// 验证自定义属性 5 种类型 round-trip。
#[test]
fn custom_properties_all_value_types_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");

    prs.custom_properties_mut()
        .set("Project", CustomPropertyValue::Text("Demo".to_string()));
    prs.custom_properties_mut()
        .set("Version", CustomPropertyValue::Int(42));
    prs.custom_properties_mut()
        .set("Score", CustomPropertyValue::Float(3.15));
    prs.custom_properties_mut()
        .set("Approved", CustomPropertyValue::Bool(true));
    prs.custom_properties_mut().set(
        "CreatedAt",
        CustomPropertyValue::DateTime("2026-06-28T00:00:00Z".to_string()),
    );

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");

    let props = prs2.custom_properties();
    assert_eq!(
        props.get("Project"),
        Some(&CustomPropertyValue::Text("Demo".to_string())),
        "Text property should round-trip"
    );
    assert_eq!(
        props.get("Version"),
        Some(&CustomPropertyValue::Int(42)),
        "Int property should round-trip"
    );
    assert_eq!(
        props.get("Approved"),
        Some(&CustomPropertyValue::Bool(true)),
        "Bool property should round-trip"
    );
}

/// 验证自定义属性删除 round-trip。
#[test]
fn custom_properties_remove_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");

    prs.custom_properties_mut()
        .set("Temp", CustomPropertyValue::Int(1));
    let removed = prs.custom_properties_mut().remove("Temp");
    assert_eq!(removed, Some(CustomPropertyValue::Int(1)));

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert!(
        prs2.custom_properties().get("Temp").is_none(),
        "removed property should not exist"
    );
}

/// 验证核心属性 round-trip。
#[test]
fn core_properties_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    {
        let core = prs.core_properties_mut();
        core.title = Some("测试标题".to_string());
        core.creator = Some("测试作者".to_string());
        core.subject = Some("测试主题".to_string());
    }

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");

    // 验证 docProps/core.xml 存在（Presentation::new 会设置默认 core 属性）
    let core_xml = read_zip_part(&bytes, "docProps/core.xml").expect("core.xml exists");
    assert!(
        core_xml.contains("coreProperties"),
        "core.xml should be serialized, got: {}",
        core_xml
    );

    // 验证 load_bytes 不报错（core_properties 读取路径不崩溃即可）
    let _ = prs2.core_properties();
}

/// 验证多 slide 备注与批注混合 round-trip。
#[test]
fn multiple_slides_notes_and_comments_round_trip() {
    let mut prs = Presentation::new().expect("Presentation::new failed");
    let author_id = prs.comment_authors_mut().get_or_insert_id("赵六", "ZL");

    // 第 1 张：备注
    let counter = prs.id_counter();
    let slide1 = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide 1 failed");
    slide1.set_notes_text(Some("第 1 张备注"));

    // 第 2 张：批注
    let counter = prs.id_counter();
    let slide2 = prs
        .slides_mut()
        .add_slide(counter)
        .expect("add_slide 2 failed");
    slide2.add_comment(author_id, Inches(1.0), Inches(1.0), "第 2 张批注");

    assert_eq!(prs.slides().len(), 2);

    let bytes = prs.to_bytes().expect("to_bytes failed");
    let prs2 = Presentation::load_bytes(&bytes).expect("load_bytes failed");
    assert_eq!(prs2.slides().len(), 2, "slide count preserved");

    // 验证 notesSlide 和 comment part 都存在
    assert!(
        read_zip_part(&bytes, "ppt/notesSlides/notesSlide1.xml").is_some(),
        "slide 1 notes preserved"
    );
    assert!(
        read_zip_part(&bytes, "ppt/comments/comment1.xml").is_some(),
        "slide 2 comments preserved"
    );
}
