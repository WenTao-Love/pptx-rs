//! 幻灯片评论（`<p:cmLst>` / `<p:cmAuthorLst>`）。
//!
//! 对标 python-pptx 的 `Slide.comments`（v0.6.21+）与 `Comment` / `CommentAuthor`。
//!
//! # OOXML 结构
//!
//! ## 评论列表（`/ppt/comments/commentN.xml`）
//!
//! ```text
//! <p:cmLst xmlns:a="..." xmlns:p="..." xmlns:r="...">
//!   <p:cm authorId="0" dt="2024-01-01T12:00:00Z" idx="1">
//!     <p:pos x="100" y="100"/>
//!     <p:text>评论正文</p:text>
//!   </p:cm>
//! </p:cmLst>
//! ```
//!
//! ## 作者列表（`/ppt/commentAuthors.xml`，全局共享）
//!
//! ```text
//! <p:cmAuthorLst xmlns:p="...">
//!   <p:cmAuthor id="0" name="张三" initials="ZS"/>
//! </p:cmAuthorLst>
//! ```
//!
//! # 在三层架构中的位置
//!
//! 本模块属于 **OOXML 模型层**：负责把 `<p:cmLst>` / `<p:cmAuthorLst>` 序列化为 XML，
//! 以及从 XML 解析回内存模型。高阶 API（`Slide::add_comment` 等）在 `slide.rs` 中。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.comments.Comment` ←→ [`Comment`]；
//! - `pptx.comments.CommentAuthor` ←→ [`CommentAuthor`]；
//! - `_CommentAuthors` 集合 ←→ [`CommentAuthorList`]。

use crate::oxml::ns::{NS_DRAWING_MAIN, NS_DRAWING_RELS, NS_PRESENTATION_MAIN};
use crate::oxml::writer::XmlWriter;

/// 一条评论（`<p:cm>`）。
///
/// # 字段
/// - `author_id`：作者 ID（指向 [`CommentAuthorList`] 中的作者）；
/// - `date_time`：评论时间（ISO 8601 格式字符串，如 `"2024-01-01T12:00:00Z"`）；
/// - `idx`：评论在该 slide 中的唯一索引（PowerPoint 用此关联批注锚点）；
/// - `pos_x` / `pos_y`：评论锚点在 slide 上的坐标（EMU）；
/// - `text`：评论正文。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Comment {
    /// 作者 ID（对应 `<p:cmAuthor id="...">`）。
    pub author_id: u32,
    /// 评论时间（ISO 8601 字符串）。
    pub date_time: String,
    /// 评论索引（`<p:cm idx="...">`）。
    pub idx: u32,
    /// 锚点 X 坐标（EMU）。
    pub pos_x: i64,
    /// 锚点 Y 坐标（EMU）。
    pub pos_y: i64,
    /// 评论正文。
    pub text: String,
}

impl Comment {
    /// 创建一条新评论。
    ///
    /// # 参数
    /// - `author_id`：作者 ID；
    /// - `idx`：评论索引；
    /// - `pos_x` / `pos_y`：锚点坐标（EMU）；
    /// - `text`：评论正文。
    pub fn new(author_id: u32, idx: u32, pos_x: i64, pos_y: i64, text: impl Into<String>) -> Self {
        Comment {
            author_id,
            date_time: String::new(),
            idx,
            pos_x,
            pos_y,
            text: text.into(),
        }
    }

    /// 序列化为 `<p:cm>` 元素（不含外层 `<p:cmLst>`）。
    ///
    /// # XML 结构
    ///
    /// ```text
    /// <p:cm authorId="0" dt="..." idx="1">
    ///   <p:pos x="100" y="100"/>
    ///   <p:text>评论正文</p:text>
    /// </p:cm>
    /// ```
    pub fn write_xml(&self, w: &mut XmlWriter) {
        w.open_with(
            "p:cm",
            &[
                ("authorId", &self.author_id.to_string()),
                ("dt", &self.date_time),
                ("idx", &self.idx.to_string()),
            ],
        );
        w.empty_with(
            "p:pos",
            &[
                ("x", &self.pos_x.to_string()),
                ("y", &self.pos_y.to_string()),
            ],
        );
        w.open("p:text");
        w.text(&self.text);
        w.close("p:text");
        w.close("p:cm");
    }
}

/// 评论列表（`<p:cmLst>`），对应 `/ppt/comments/commentN.xml`。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CommentList {
    /// 评论条目（按插入顺序保留）。
    pub comments: Vec<Comment>,
}

impl CommentList {
    /// 创建空列表。
    pub fn new() -> Self {
        Self::default()
    }

    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.comments.is_empty()
    }

    /// 返回条目数。
    pub fn len(&self) -> usize {
        self.comments.len()
    }

    /// 追加一条评论。
    pub fn push(&mut self, c: Comment) {
        self.comments.push(c);
    }

    /// 序列化为完整的 `commentN.xml` 文档字符串。
    ///
    /// # XML 结构
    ///
    /// ```text
    /// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    /// <p:cmLst xmlns:a="..." xmlns:p="..." xmlns:r="...">
    ///   <p:cm ...>...</p:cm>
    /// </p:cmLst>
    /// ```
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::with_decl();
        let attrs: Vec<(&str, &str)> = vec![
            ("xmlns:a", NS_DRAWING_MAIN),
            ("xmlns:p", NS_PRESENTATION_MAIN),
            ("xmlns:r", NS_DRAWING_RELS),
        ];
        w.open_with("p:cmLst", &attrs);
        for c in &self.comments {
            c.write_xml(&mut w);
        }
        w.close("p:cmLst");
        w.into_string()
    }
}

/// 一位评论作者（`<p:cmAuthor>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CommentAuthor {
    /// 作者 ID（全局唯一，`<p:cmAuthor id="...">`）。
    pub id: u32,
    /// 作者显示名。
    pub name: String,
    /// 作者缩写（姓名首字母）。
    pub initials: String,
}

impl CommentAuthor {
    /// 创建一位作者。
    pub fn new(id: u32, name: impl Into<String>, initials: impl Into<String>) -> Self {
        CommentAuthor {
            id,
            name: name.into(),
            initials: initials.into(),
        }
    }

    /// 序列化为 `<p:cmAuthor>` 元素。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        w.empty_with(
            "p:cmAuthor",
            &[
                ("id", &self.id.to_string()),
                ("name", &self.name),
                ("initials", &self.initials),
            ],
        );
    }
}

/// 评论作者列表（`<p:cmAuthorLst>`），对应 `/ppt/commentAuthors.xml`。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CommentAuthorList {
    /// 作者条目。
    pub authors: Vec<CommentAuthor>,
}

impl CommentAuthorList {
    /// 创建空列表。
    pub fn new() -> Self {
        Self::default()
    }

    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.authors.is_empty()
    }

    /// 返回条目数。
    pub fn len(&self) -> usize {
        self.authors.len()
    }

    /// 追加一位作者。
    pub fn push(&mut self, a: CommentAuthor) {
        self.authors.push(a);
    }

    /// 按 ID 查找作者。
    pub fn get_by_id(&self, id: u32) -> Option<&CommentAuthor> {
        self.authors.iter().find(|a| a.id == id)
    }

    /// 按名字查找作者 ID。若不存在则分配新 ID 并插入，返回新 ID。
    ///
    /// 用于 `Slide::add_comment` 时自动维护作者列表。
    pub fn get_or_insert_id(&mut self, name: &str, initials: &str) -> u32 {
        if let Some(a) = self.authors.iter().find(|a| a.name == name) {
            return a.id;
        }
        let next_id = self
            .authors
            .iter()
            .map(|a| a.id)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        self.authors
            .push(CommentAuthor::new(next_id, name, initials));
        next_id
    }

    /// 序列化为完整的 `commentAuthors.xml` 文档字符串。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::with_decl();
        let attrs: Vec<(&str, &str)> = vec![("xmlns:p", NS_PRESENTATION_MAIN)];
        w.open_with("p:cmAuthorLst", &attrs);
        for a in &self.authors {
            a.write_xml(&mut w);
        }
        w.close("p:cmAuthorLst");
        w.into_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn comment_write_xml_basic() {
        let c = Comment::new(0, 1, 100, 200, "Hello");
        let mut w = XmlWriter::new();
        c.write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains("<p:cm authorId=\"0\""));
        assert!(xml.contains("idx=\"1\""));
        assert!(xml.contains("<p:pos x=\"100\" y=\"200\"/>"));
        assert!(xml.contains("<p:text>Hello</p:text>"));
    }

    #[test]
    fn comment_list_to_xml() {
        let mut lst = CommentList::new();
        lst.push(Comment::new(0, 1, 100, 200, "First"));
        lst.push(Comment::new(1, 2, 300, 400, "Second"));
        let xml = lst.to_xml();
        assert!(xml.contains("<?xml"));
        assert!(xml.contains("<p:cmLst"));
        assert!(xml.contains("xmlns:a="));
        assert!(xml.contains("xmlns:p="));
        assert!(xml.contains("First"));
        assert!(xml.contains("Second"));
    }

    #[test]
    fn comment_list_empty_to_xml() {
        let lst = CommentList::new();
        let xml = lst.to_xml();
        assert!(xml.contains("<p:cmLst"));
        assert!(xml.contains("</p:cmLst>"));
        // 空列表不应包含 <p:cm
        assert!(!xml.contains("<p:cm "));
    }

    #[test]
    fn comment_author_write_xml() {
        let a = CommentAuthor::new(0, "张三", "ZS");
        let mut w = XmlWriter::new();
        a.write_xml(&mut w);
        let xml = &w.buf;
        assert!(xml.contains("<p:cmAuthor id=\"0\""));
        assert!(xml.contains("name=\"张三\""));
        assert!(xml.contains("initials=\"ZS\""));
    }

    #[test]
    fn comment_author_list_to_xml() {
        let mut lst = CommentAuthorList::new();
        lst.push(CommentAuthor::new(0, "张三", "ZS"));
        lst.push(CommentAuthor::new(1, "李四", "LS"));
        let xml = lst.to_xml();
        assert!(xml.contains("<?xml"));
        assert!(xml.contains("<p:cmAuthorLst"));
        assert!(xml.contains("张三"));
        assert!(xml.contains("李四"));
    }

    #[test]
    fn comment_author_get_or_insert() {
        let mut lst = CommentAuthorList::new();
        let id1 = lst.get_or_insert_id("张三", "ZS");
        assert_eq!(id1, 1);
        let id2 = lst.get_or_insert_id("李四", "LS");
        assert_eq!(id2, 2);
        // 已存在的作者返回原 ID
        let id3 = lst.get_or_insert_id("张三", "ZS");
        assert_eq!(id3, 1);
        assert_eq!(lst.len(), 2);
    }

    #[test]
    fn comment_author_get_by_id() {
        let mut lst = CommentAuthorList::new();
        lst.push(CommentAuthor::new(0, "张三", "ZS"));
        lst.push(CommentAuthor::new(1, "李四", "LS"));
        assert_eq!(lst.get_by_id(0).unwrap().name, "张三");
        assert_eq!(lst.get_by_id(1).unwrap().name, "李四");
        assert!(lst.get_by_id(2).is_none());
    }
}
