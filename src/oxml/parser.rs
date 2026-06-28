//! XML 解析与序列化基础。
//!
//! 我们的策略是 **不** 用 serde 派生（OOXML 命名空间混乱、属性多），而是用
//! quick-xml 直接走 SAX 风格流式 API，专注读写一致性。
//!
//! 本文件提供：
//!
//! - 字符串 ↔ 数字 / 布尔 / 枚举 的转换；
//! - 属性读取的便利函数；
//! - XML 转义 / 反转义；
//! - 简易的"开始/结束元素"过滤器。
//!
//! # 读取策略
//!
//! 读取大文件（如 `theme1.xml`）时**不**用 `Event::Start/End` 嵌套扫描完整
//! 树形结构——一方面性能差，另一方面没必要。本库只对"有用属性"做精确匹配，
//! 其余部分用 [`collect_inner_text`] 工具一次吞掉。
//!
//! # 写入策略
//!
//! 序列化时**不**经过本模块——所有写操作走 [`super::writer::XmlWriter`]。
//! 这里的 `escape` / `render_attrs` / `render_start` / `render_empty` / `render_end`
//! 是为兼容老路径与测试用例保留的"基础字符串工具"。

use std::collections::HashMap;
use std::str::FromStr;

use quick_xml::events::BytesStart;
use quick_xml::reader::Reader;

use crate::units::Emu;

/// XML 字符串属性转义。
pub fn escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
    out
}

/// XML 字符串反转义（解析时用）。
pub fn unescape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut i = 0;
    let bytes = s.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'&' {
            if let Some(end) = s[i..].find(';') {
                let entity = &s[i + 1..i + end];
                let replacement = match entity {
                    "amp" => Some('&'),
                    "lt" => Some('<'),
                    "gt" => Some('>'),
                    "quot" => Some('"'),
                    "apos" => Some('\''),
                    _ => None,
                };
                if let Some(c) = replacement {
                    out.push(c);
                    i += end + 1;
                    continue;
                }
            }
        }
        out.push(bytes[i] as char);
        i += 1;
    }
    out
}

/// 元素属性集合：本地名（去前缀）→ 字符串。
/// 使用 `Vec<u8>` 拥有键的所有权，避免借用迭代器临时值。
pub type AttrMap = HashMap<Vec<u8>, String>;

/// 从 `BytesStart` 提取属性集合（值自动反转义）。
pub fn attrs_of(e: &BytesStart<'_>) -> AttrMap {
    let mut m = HashMap::new();
    for a in e.attributes().flatten() {
        let k = a.key.as_ref().to_vec();
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .map(|c| c.to_string())
            .unwrap_or_default();
        m.insert(k, v);
    }
    m
}

/// 取出某个属性，缺失返回 None。
pub fn get_attr<'a>(m: &'a AttrMap, key: &[u8]) -> Option<&'a String> {
    m.get(key)
}

/// 取出某个属性，缺失时用默认值。
pub fn get_attr_or<'a>(m: &'a AttrMap, key: &[u8], default: &'a str) -> &'a str {
    m.get(key).map(|s| s.as_str()).unwrap_or(default)
}

/// `i64` 解析。
pub fn parse_i64(s: &str) -> Option<i64> {
    s.parse().ok()
}
/// `u32` 解析。
pub fn parse_u32(s: &str) -> Option<u32> {
    s.parse().ok()
}
/// `u16` 解析。
pub fn parse_u16(s: &str) -> Option<u16> {
    s.parse().ok()
}
/// `u8` 解析。
pub fn parse_u8(s: &str) -> Option<u8> {
    s.parse().ok()
}
/// `f64` 解析（解析失败时返回 None 而不是 NaN）。
pub fn parse_f64(s: &str) -> Option<f64> {
    s.parse().ok()
}
/// 布尔解析：`true`/`1`/`on` 视为真；其余视为假（`0`/`false`/`off`/空）。
pub fn parse_bool(s: &str) -> bool {
    matches!(s, "1" | "true" | "on" | "True")
}

/// 任意 FromStr 类型的属性解析：缺失返回 None。
pub fn parse_attr<T: FromStr>(m: &AttrMap, key: &[u8]) -> Option<T> {
    m.get(key).and_then(|v| v.parse().ok())
}

/// 任意 FromStr 类型的属性解析：缺失时使用默认值。
pub fn parse_attr_or<T: FromStr>(m: &AttrMap, key: &[u8], default: T) -> T {
    m.get(key).and_then(|v| v.parse().ok()).unwrap_or(default)
}

/// 把 i64 序列化为十进制字符串。
pub fn i64_to_str(v: i64) -> String {
    v.to_string()
}

/// 写一个 EMU 值（i64 → 十进制字符串）。
pub fn emu_str(v: Emu) -> String {
    v.value().to_string()
}

/// 写一个可选 EMU 值（不写就不输出）。
pub fn emu_opt(v: Option<Emu>) -> String {
    v.map(|x| x.value().to_string()).unwrap_or_default()
}

/// 把字符串集合渲染为 `"k1=\"v1\" k2=\"v2\""`（不带前后尖括号）。
pub fn render_attrs(pairs: &[(&str, &str)]) -> String {
    let mut s = String::new();
    for (k, v) in pairs {
        s.push(' ');
        s.push_str(k);
        s.push_str("=\"");
        s.push_str(&escape(v));
        s.push('"');
    }
    s
}

/// 把一个 `BytesStart` 渲染为完整的开标签（仅写 start 事件，不写子节点）。
pub fn render_start(name: &[u8], attrs: &[(&str, &str)]) -> String {
    let mut s = String::from("<");
    s.push_str(std::str::from_utf8(name).unwrap_or("?"));
    for (k, v) in attrs {
        s.push(' ');
        s.push_str(k);
        s.push_str("=\"");
        s.push_str(&escape(v));
        s.push('"');
    }
    s.push('>');
    s
}

/// `render_start` 的自闭合版本。
pub fn render_empty(name: &[u8], attrs: &[(&str, &str)]) -> String {
    let mut s = String::from("<");
    s.push_str(std::str::from_utf8(name).unwrap_or("?"));
    for (k, v) in attrs {
        s.push(' ');
        s.push_str(k);
        s.push_str("=\"");
        s.push_str(&escape(v));
        s.push('"');
    }
    s.push_str("/>");
    s
}

/// 写一个闭标签。
pub fn render_end(name: &[u8]) -> String {
    let mut s = String::from("</");
    s.push_str(std::str::from_utf8(name).unwrap_or("?"));
    s.push('>');
    s
}

/// 收集一个 XML Reader 中**当前开始标签到对应结束标签**之间的内容到字符串。
/// 用于 `<a:rPr>...</a:rPr>` 这种"完整子元素"提取。
pub fn collect_inner_text<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    name: &[u8],
) -> crate::Result<String> {
    use quick_xml::events::Event;
    let mut depth = 0usize;
    let mut out = String::new();
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == name {
                    depth = 1;
                    continue;
                }
                if depth > 0 {
                    depth += 1;
                }
            }
            Ok(Event::End(_e)) => {
                if depth > 0 {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(out);
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if depth > 0 {
                    // quick-xml 0.40 移除 `unescape()`，使用 `decode()`
                    out.push_str(&t.decode().unwrap_or_default());
                }
            }
            Ok(Event::CData(c)) => {
                if depth > 0 {
                    out.push_str(std::str::from_utf8(&c).unwrap_or(""));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Err(crate::Error::oxml(format!(
        "unexpected EOF when collecting <{}>",
        std::str::from_utf8(name).unwrap_or("?")
    )))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escape_unescape_round() {
        let s = "a < b & \"c\" 'd' > z";
        let e = escape(s);
        let u = unescape(&e);
        assert_eq!(u, s);
    }

    #[test]
    fn render_attrs_works() {
        let s = render_attrs(&[("a", "1"), ("b", "x\"y")]);
        assert_eq!(s, " a=\"1\" b=\"x&quot;y\"");
    }
}
