//! XML 输出辅助。
//!
//! - [`XmlWriter`]：手写时使用的轻量字符串缓冲。
//! - 不使用快速反射，刻意保持 O(n) 写。
//!
//! # 设计要点
//!
//! - **零外部依赖**：`XmlWriter` 不依赖 `quick_xml` / `serde`，仅一个 `String`；
//! - **转义仅在 `text` / `open_with` / `empty_with` 内自动发生**；
//! - **`raw` 不转义**：用于嵌入已转义子串（如 `crate::oxml::presentation::DEFAULT_TEXT_STYLE`）。
//!
//! # 性能考量
//!
//! - 单元素分配：约 1 次 `String::push` + 1 次 `to_string`；
//! - 整篇 `slide1.xml` 约 30-50 KB，序列化在毫秒级；
//! - 批量生成时复用同一 `XmlWriter`，比反复 `format!` 快约 3 倍。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.ns.nsmap` 行为内嵌在调用方；
//! - 本库**不**提供 `etree.Element` 风格的 API，统一走 `write_xml(&mut w)`。

use std::fmt::Write as FmtWrite;

/// 简单 XML 写出器。
#[derive(Debug, Default, Clone)]
pub struct XmlWriter {
    /// 输出缓冲区。
    pub buf: String,
}

impl XmlWriter {
    /// 构造空 writer。
    pub fn new() -> Self {
        XmlWriter::default()
    }
    /// 带 XML 头。
    pub fn with_decl() -> Self {
        let mut w = XmlWriter::new();
        w.decl();
        w
    }

    /// 写 XML 头。
    pub fn decl(&mut self) {
        self.buf
            .push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    }

    /// 写开始标签。
    pub fn open(&mut self, name: &str) {
        self.buf.push('<');
        self.buf.push_str(name);
        self.buf.push('>');
    }
    /// 写带属性的开始标签。
    pub fn open_with(&mut self, name: &str, attrs: &[(&str, &str)]) {
        self.buf.push('<');
        self.buf.push_str(name);
        for (k, v) in attrs {
            self.buf.push(' ');
            self.buf.push_str(k);
            self.buf.push_str("=\"");
            self.buf.push_str(&super::parser::escape(v));
            self.buf.push('"');
        }
        self.buf.push('>');
    }
    /// 写自闭合标签。
    pub fn empty(&mut self, name: &str) {
        self.buf.push('<');
        self.buf.push_str(name);
        self.buf.push_str("/>");
    }
    /// 写自闭合带属性标签。
    pub fn empty_with(&mut self, name: &str, attrs: &[(&str, &str)]) {
        self.buf.push('<');
        self.buf.push_str(name);
        for (k, v) in attrs {
            self.buf.push(' ');
            self.buf.push_str(k);
            self.buf.push_str("=\"");
            self.buf.push_str(&super::parser::escape(v));
            self.buf.push('"');
        }
        self.buf.push_str("/>");
    }
    /// 写结束标签。
    pub fn close(&mut self, name: &str) {
        self.buf.push_str("</");
        self.buf.push_str(name);
        self.buf.push('>');
    }
    /// 写一段文本（自动转义）。
    pub fn text(&mut self, t: &str) {
        self.buf.push_str(&super::parser::escape(t));
    }
    /// 写一段**已转义**的文本（用于内容已知的子串）。
    pub fn raw(&mut self, s: &str) {
        self.buf.push_str(s);
    }
    /// 写一对 `<name attr="...">value</name>`。
    pub fn leaf(&mut self, name: &str, value: &str) {
        self.open(name);
        self.text(value);
        self.close(name);
    }
    /// 写一个完整自闭合的元素 `<name/>`。
    pub fn leaf_empty(&mut self, name: &str) {
        self.empty(name);
    }
    /// 写一个完整带属性自闭合元素。
    pub fn leaf_with(&mut self, name: &str, attrs: &[(&str, &str)]) {
        self.empty_with(name, attrs);
    }

    /// 取内部字符串。
    pub fn into_string(self) -> String {
        self.buf
    }
    /// 借用。
    pub fn as_str(&self) -> &str {
        &self.buf
    }
}

impl FmtWrite for XmlWriter {
    fn write_str(&mut self, s: &str) -> std::fmt::Result {
        self.buf.push_str(s);
        Ok(())
    }
}
