//! Office 主题 `<a:theme>` 标准版。
//!
//! PowerPoint 强制要求一份合法的 theme1.xml，否则会被 Office/WPS 拒绝打开。
//! 这里直接采用 python-pptx 默认输出的 Office 主题，保留完整结构与所有字体脚本。
//! 出于二进制大小考虑，省略了 latin/ea/cs 之外的 `<a:font>` 列表（实际 Office 主题里全是它们）。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.theme.Theme` ←→ [`Theme`] 结构体（v0.2 起支持结构化解析）；
//! - `pptx.parts.theme.ThemePart._element` 在加载时由 [`parse_theme`] 解析为 [`Theme`]；
//! - 写路径使用 [`Theme::to_xml`]（结构化序列化）或 [`default_theme_xml`]（完整 Office 主题 XML）。
//!
//! # 字体脚本表
//!
//! `THEME_XML` 完整列出了 30+ 个 script（`Jpan` / `Hang` / `Hans` / `Hant` / ...），
//! 缺失任何一个都会导致 PowerPoint 警告"主题不完整"。

use crate::oxml::ns::NS_DRAWING_MAIN;
use crate::oxml::writer::XmlWriter;

#[allow(dead_code)]
const THEME_NAME: &str = "Office Theme";

/// 结构化主题模型（对应 `<a:theme>`）。
///
/// 解析 `name` / `color_scheme` / `font_scheme` / `format_scheme` 四个核心字段。
/// `<a:objectDefaults>` / `<a:extraClrSchemeLst>` 等暂不解析（写路径用 [`default_theme_xml`]）。
#[derive(Clone, Debug, Default)]
pub struct Theme {
    /// 主题名（`<a:theme name="Office Theme">`）。
    pub name: String,
    /// 颜色方案（`<a:clrScheme>`）。
    pub color_scheme: ColorScheme,
    /// 字体方案（`<a:fontScheme>`）。
    pub font_scheme: FontScheme,
    /// 格式方案（`<a:fmtScheme>`）。
    pub format_scheme: FormatScheme,
}

/// 颜色方案（对应 `<a:clrScheme>`）。
///
/// 包含 12 个颜色槽位：dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink。
#[derive(Clone, Debug, Default)]
pub struct ColorScheme {
    /// 颜色方案名（`<a:clrScheme name="Office">`）。
    pub name: String,
    /// 暗色 1（通常为黑色/文字色）。
    pub dk1: Option<ThemeColor>,
    /// 亮色 1（通常为白色/背景色）。
    pub lt1: Option<ThemeColor>,
    /// 暗色 2。
    pub dk2: Option<ThemeColor>,
    /// 亮色 2。
    pub lt2: Option<ThemeColor>,
    /// 强调色 1-6。
    pub accent1: Option<ThemeColor>,
    pub accent2: Option<ThemeColor>,
    pub accent3: Option<ThemeColor>,
    pub accent4: Option<ThemeColor>,
    pub accent5: Option<ThemeColor>,
    pub accent6: Option<ThemeColor>,
    /// 超链接颜色。
    pub hlink: Option<ThemeColor>,
    /// 已访问超链接颜色。
    pub fol_hlink: Option<ThemeColor>,
}

/// 字体方案（对应 `<a:fontScheme>`）。
///
/// 解析 majorFont / minorFont 的 latin / ea / cs typeface。
#[derive(Clone, Debug, Default)]
pub struct FontScheme {
    /// 字体方案名（`<a:fontScheme name="Office">`）。
    pub name: String,
    /// 主标题字体 latin（`<a:majorFont><a:latin typeface="..."/>`）。
    pub major_latin: String,
    /// 主标题字体 东亚（`<a:majorFont><a:ea typeface="..."/>`）。
    pub major_ea: String,
    /// 主标题字体 复杂文种（`<a:majorFont><a:cs typeface="..."/>`）。
    pub major_cs: String,
    /// 正文字体 latin（`<a:minorFont><a:latin typeface="..."/>`）。
    pub minor_latin: String,
    /// 正文字体 东亚（`<a:minorFont><a:ea typeface="..."/>`）。
    pub minor_ea: String,
    /// 正文字体 复杂文种（`<a:minorFont><a:cs typeface="..."/>`）。
    pub minor_cs: String,
}

/// 格式方案（对应 `<a:fmtScheme>`）。
///
/// FormatScheme 包含 fillStyleLst / lnStyleLst / effectStyleLst / bgFillStyleLst，
/// 每个都含有复杂的嵌套元素（gradFill / ln / effectLst / scene3d / sp3d 等）。
///
/// # 结构化策略
///
/// 采用**适度结构化**：保留 raw_xml 用于 round-trip，同时把 4 个 style 列表
/// 拆分为 `fill_styles` / `line_styles` / `effect_styles` / `bg_fill_styles`，
/// 每个元素是对应子元素的原始 XML 字符串（如 `<a:gradFill>...</a:gradFill>`）。
/// 这样用户可以查询/替换某个 style，而无需完整解析每个 fill/ln/effect 的内部结构。
#[derive(Clone, Debug, Default)]
pub struct FormatScheme {
    /// 格式方案名（`<a:fmtScheme name="Office">`）。
    pub name: String,
    /// 原始 XML 内容（`<a:fmtScheme>` 内部的子元素 XML）。
    ///
    /// 为空时 [`Theme::to_xml`] 会使用默认的 Office 格式方案。
    pub raw_xml: String,
    /// 填充样式列表（`<a:fillStyleLst>` 内的每个 fill 元素 XML）。
    ///
    /// 通常 3 个：solidFill / gradFill / gradFill。
    pub fill_styles: Vec<String>,
    /// 线条样式列表（`<a:lnStyleLst>` 内的每个 `<a:ln>` 元素 XML）。
    ///
    /// 通常 3 个。
    pub line_styles: Vec<String>,
    /// 效果样式列表（`<a:effectStyleLst>` 内的每个 `<a:effectStyle>` 元素 XML）。
    ///
    /// 通常 3 个。
    pub effect_styles: Vec<String>,
    /// 背景填充样式列表（`<a:bgFillStyleLst>` 内的每个 fill 元素 XML）。
    ///
    /// 通常 2 个。
    pub bg_fill_styles: Vec<String>,
}

/// 主题颜色值（对应 `<a:srgbClr>` 或 `<a:sysClr>`）。
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub enum ThemeColor {
    /// sRGB 颜色（`<a:srgbClr val="RRGGBB"/>`）。
    Srgb(String),
    /// 系统颜色（`<a:sysClr val="windowText" lastClr="000000"/>`）。
    Sys(String, String),
    /// 无颜色（解析失败或空槽位）。
    #[default]
    None,
}

impl ThemeColor {
    /// 写出对应的 XML 元素到 writer。
    ///
    /// # 元素结构
    ///
    /// - `Srgb` → `<a:srgbClr val="RRGGBB"/>`
    /// - `Sys` → `<a:sysClr val="..." lastClr="..."/>`
    /// - `None` → 不写出任何元素
    pub fn write_xml(&self, w: &mut XmlWriter) {
        match self {
            ThemeColor::Srgb(val) => {
                w.empty_with("a:srgbClr", &[("val", val.as_str())]);
            }
            ThemeColor::Sys(val, last_clr) => {
                w.empty_with(
                    "a:sysClr",
                    &[("val", val.as_str()), ("lastClr", last_clr.as_str())],
                );
            }
            ThemeColor::None => { /* 不写出 */ }
        }
    }
}

impl ColorScheme {
    /// 写出 `<a:clrScheme>` 元素到 writer。
    ///
    /// # 元素顺序（OOXML 规范）
    ///
    /// dk1 → lt1 → dk2 → lt2 → accent1 → accent2 → accent3 → accent4 →
    /// accent5 → accent6 → hlink → folHlink
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let name = if self.name.is_empty() {
            "Office"
        } else {
            self.name.as_str()
        };
        w.open_with("a:clrScheme", &[("name", name)]);
        // 按 OOXML 规范顺序写出 12 个颜色槽位
        self.write_slot(w, "a:dk1", &self.dk1);
        self.write_slot(w, "a:lt1", &self.lt1);
        self.write_slot(w, "a:dk2", &self.dk2);
        self.write_slot(w, "a:lt2", &self.lt2);
        self.write_slot(w, "a:accent1", &self.accent1);
        self.write_slot(w, "a:accent2", &self.accent2);
        self.write_slot(w, "a:accent3", &self.accent3);
        self.write_slot(w, "a:accent4", &self.accent4);
        self.write_slot(w, "a:accent5", &self.accent5);
        self.write_slot(w, "a:accent6", &self.accent6);
        self.write_slot(w, "a:hlink", &self.hlink);
        self.write_slot(w, "a:folHlink", &self.fol_hlink);
        w.close("a:clrScheme");
    }

    /// 写出单个颜色槽位（如 `<a:dk1><a:srgbClr val="..."/></a:dk1>`）。
    fn write_slot(&self, w: &mut XmlWriter, tag: &str, color: &Option<ThemeColor>) {
        w.open(tag);
        if let Some(c) = color {
            c.write_xml(w);
        }
        w.close(tag);
    }
}

impl FontScheme {
    /// 写出 `<a:fontScheme>` 元素到 writer。
    ///
    /// # 元素结构
    ///
    /// ```text
    /// <a:fontScheme name="...">
    ///   <a:majorFont>
    ///     <a:latin typeface="..."/>
    ///     <a:ea typeface="..."/>
    ///     <a:cs typeface="..."/>
    ///   </a:majorFont>
    ///   <a:minorFont>
    ///     <a:latin typeface="..."/>
    ///     <a:ea typeface="..."/>
    ///     <a:cs typeface="..."/>
    ///   </a:minorFont>
    /// </a:fontScheme>
    /// ```
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let name = if self.name.is_empty() {
            "Office"
        } else {
            self.name.as_str()
        };
        w.open_with("a:fontScheme", &[("name", name)]);
        // majorFont
        w.open("a:majorFont");
        w.empty_with("a:latin", &[("typeface", self.major_latin.as_str())]);
        w.empty_with("a:ea", &[("typeface", self.major_ea.as_str())]);
        w.empty_with("a:cs", &[("typeface", self.major_cs.as_str())]);
        w.close("a:majorFont");
        // minorFont
        w.open("a:minorFont");
        w.empty_with("a:latin", &[("typeface", self.minor_latin.as_str())]);
        w.empty_with("a:ea", &[("typeface", self.minor_ea.as_str())]);
        w.empty_with("a:cs", &[("typeface", self.minor_cs.as_str())]);
        w.close("a:minorFont");
        w.close("a:fontScheme");
    }
}

impl FormatScheme {
    /// 写出 `<a:fmtScheme>` 元素到 writer。
    ///
    /// 优先级：结构化字段（fill_styles 等）> raw_xml > 默认 Office 格式方案。
    /// 若结构化字段非空，按 OOXML 顺序输出 4 个 style 列表；否则回退到 raw_xml。
    pub fn write_xml(&self, w: &mut XmlWriter) {
        let name = if self.name.is_empty() {
            "Office"
        } else {
            self.name.as_str()
        };
        // 判断是否使用结构化字段
        let has_structured = !self.fill_styles.is_empty()
            || !self.line_styles.is_empty()
            || !self.effect_styles.is_empty()
            || !self.bg_fill_styles.is_empty();

        if has_structured {
            // 使用结构化字段按 OOXML 顺序输出
            w.open_with("a:fmtScheme", &[("name", name)]);
            if !self.fill_styles.is_empty() {
                w.open("a:fillStyleLst");
                for s in &self.fill_styles {
                    w.raw(s);
                }
                w.close("a:fillStyleLst");
            }
            if !self.line_styles.is_empty() {
                w.open("a:lnStyleLst");
                for s in &self.line_styles {
                    w.raw(s);
                }
                w.close("a:lnStyleLst");
            }
            if !self.effect_styles.is_empty() {
                w.open("a:effectStyleLst");
                for s in &self.effect_styles {
                    w.raw(s);
                }
                w.close("a:effectStyleLst");
            }
            if !self.bg_fill_styles.is_empty() {
                w.open("a:bgFillStyleLst");
                for s in &self.bg_fill_styles {
                    w.raw(s);
                }
                w.close("a:bgFillStyleLst");
            }
            w.close("a:fmtScheme");
        } else if self.raw_xml.is_empty() {
            // 使用默认的 Office 格式方案（从 THEME_XML 中提取）
            w.raw(DEFAULT_FMT_SCHEME);
        } else {
            // 回退到 raw_xml
            w.open_with("a:fmtScheme", &[("name", name)]);
            w.raw(&self.raw_xml);
            w.close("a:fmtScheme");
        }
    }

    /// 从 raw_xml 解析填充 4 个结构化字段（fill_styles / line_styles / effect_styles / bg_fill_styles）。
    ///
    /// 解析后结构化字段非空，后续 [`write_xml`](Self::write_xml) 会优先使用结构化字段。
    /// 若 raw_xml 为空，此方法不做任何操作。
    ///
    /// # 解析策略
    ///
    /// 使用简单字符串查找定位 4 个 `<a:xxxStyleLst>` 容器，然后收集每个容器的
    /// 直接子元素 XML。不解析子元素的内部结构（如 gradFill 的 gsLst），保留原始 XML。
    pub fn parse_from_raw_xml(&mut self) {
        if self.raw_xml.is_empty() {
            return;
        }
        self.fill_styles = collect_style_lst_children(&self.raw_xml, "fillStyleLst");
        self.line_styles = collect_style_lst_children(&self.raw_xml, "lnStyleLst");
        self.effect_styles = collect_style_lst_children(&self.raw_xml, "effectStyleLst");
        self.bg_fill_styles = collect_style_lst_children(&self.raw_xml, "bgFillStyleLst");
    }

    /// 填充样式数量。
    pub fn fill_style_count(&self) -> usize {
        self.fill_styles.len()
    }

    /// 线条样式数量。
    pub fn line_style_count(&self) -> usize {
        self.line_styles.len()
    }

    /// 效果样式数量。
    pub fn effect_style_count(&self) -> usize {
        self.effect_styles.len()
    }

    /// 背景填充样式数量。
    pub fn bg_fill_style_count(&self) -> usize {
        self.bg_fill_styles.len()
    }
}

/// 从 XML 字符串中收集指定 `<a:xxxLst>` 容器的直接子元素 XML。
///
/// 例如 `collect_style_lst_children(xml, "fillStyleLst")` 会找到 `<a:fillStyleLst>...</a:fillStyleLst>`，
/// 返回其内部每个直接子元素的完整 XML（如 `<a:solidFill>...</a:solidFill>`）。
///
/// # 算法
///
/// 使用简单的括号配对算法：在容器内遇到 Start 事件开始收集，遇到对应的 End 事件结束收集。
/// 自闭合 Empty 事件直接收集。不递归到子元素（只收集直接子元素）。
fn collect_style_lst_children(xml: &str, container_local: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut rd = quick_xml::Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();

    // 状态机：先找容器，再收集子元素
    enum State {
        Seeking,
        InContainer,
    }
    let mut state = State::Seeking;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                match state {
                    State::Seeking => {
                        if local == container_local.as_bytes() {
                            state = State::InContainer;
                        }
                    }
                    State::InContainer => {
                        // 收集完整子元素（含开闭标签）
                        if let Ok(inner) = collect_full_element_str(&mut rd, e.into_owned()) {
                            result.push(inner);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                if let State::InContainer = state {
                    if local == container_local.as_bytes() {
                        state = State::Seeking;
                    }
                }
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let name = e.name();
                let local = local_name_quick(name.as_ref());
                if let State::InContainer = state {
                    // 自闭合子元素，重构为完整 XML 字符串
                    // 注意：用完整 name（含命名空间前缀）而非 local，保证 round-trip
                    let _ = local; // local 仅用于上面的状态判断保留
                    let mut s = String::new();
                    s.push('<');
                    s.push_str(std::str::from_utf8(name.as_ref()).unwrap_or(""));
                    for a in e.attributes().flatten() {
                        s.push(' ');
                        s.push_str(std::str::from_utf8(a.key.as_ref()).unwrap_or(""));
                        s.push_str("=\"");
                        s.push_str(
                            a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .as_ref(),
                        );
                        s.push('"');
                    }
                    s.push_str("/>");
                    result.push(s);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    result
}

/// 取 XML 元素名的 local part（去掉命名空间前缀）。
fn local_name_quick(name: &[u8]) -> &[u8] {
    // 处理 "a:fillStyleLst" → "fillStyleLst"
    match name.iter().position(|&b| b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

/// 收集从当前 Start 事件到对应 End 事件的完整 XML 字符串（含开闭标签）。
///
/// 与 parse_sld.rs 中的 `collect_full_element` 功能相同，但返回 String 而非写入 buffer。
fn collect_full_element_str(
    rd: &mut quick_xml::Reader<&[u8]>,
    start: quick_xml::events::BytesStart<'_>,
) -> quick_xml::Result<String> {
    let mut result = String::new();
    // 写入开标签
    // 把 start.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
    let start_name = start.name();
    let name = std::str::from_utf8(start_name.as_ref()).unwrap_or("");
    result.push('<');
    result.push_str(name);
    for a in start.attributes().flatten() {
        result.push(' ');
        result.push_str(std::str::from_utf8(a.key.as_ref()).unwrap_or(""));
        result.push_str("=\"");
        result.push_str(
            a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default()
                .as_ref(),
        );
        result.push('"');
    }
    result.push('>');

    let target = start.name();
    let mut depth = 1;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let e_name = e.name();
                let n = std::str::from_utf8(e_name.as_ref()).unwrap_or("");
                result.push('<');
                result.push_str(n);
                for a in e.attributes().flatten() {
                    result.push(' ');
                    result.push_str(std::str::from_utf8(a.key.as_ref()).unwrap_or(""));
                    result.push_str("=\"");
                    result.push_str(
                        a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .as_ref(),
                    );
                    result.push('"');
                }
                result.push('>');
                if e.name() == target {
                    depth += 1;
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let e_name = e.name();
                let n = std::str::from_utf8(e_name.as_ref()).unwrap_or("");
                result.push_str("</");
                result.push_str(n);
                result.push('>');
                if e.name() == target {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let e_name = e.name();
                let n = std::str::from_utf8(e_name.as_ref()).unwrap_or("");
                result.push('<');
                result.push_str(n);
                for a in e.attributes().flatten() {
                    result.push(' ');
                    result.push_str(std::str::from_utf8(a.key.as_ref()).unwrap_or(""));
                    result.push_str("=\"");
                    result.push_str(
                        a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .as_ref(),
                    );
                    result.push('"');
                }
                result.push_str("/>");
            }
            Ok(quick_xml::events::Event::Text(t)) => {
                // quick-xml 0.40: BytesText::unescape() 方法已移除，
                // 改用 quick_xml::escape::unescape 函数（接受 &str）。
                // BytesText 的 Deref 目标是 [u8]，需要先转成 &str。
                let text_str = std::str::from_utf8(t.as_ref()).unwrap_or("");
                result.push_str(&quick_xml::escape::unescape(text_str).unwrap_or_default());
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(e) => return Err(e),
            _ => {}
        }
        buf.clear();
    }
    Ok(result)
}

impl Theme {
    /// 从结构化模型序列化为完整 theme XML。
    ///
    /// 与 [`default_theme_xml`] 不同，此方法根据 [`Theme`] 结构体的字段值
    /// 生成 XML，允许用户修改颜色方案、字体方案等。
    ///
    /// # 注意
    ///
    /// - 字体脚本表（`<a:font script="..." typeface="..."/>`）不在此方法中生成，
    ///   因为它们不影响 PowerPoint 的核心解析。如需完整字体脚本，使用 [`default_theme_xml`]。
    /// - `FormatScheme` 若 `raw_xml` 为空，使用默认 Office 格式方案。
    pub fn to_xml(&self) -> String {
        let mut w = XmlWriter::with_decl();
        let theme_name = if self.name.is_empty() {
            "Office Theme"
        } else {
            self.name.as_str()
        };
        w.open_with(
            "a:theme",
            &[("xmlns:a", NS_DRAWING_MAIN), ("name", theme_name)],
        );
        // themeElements
        w.open("a:themeElements");
        self.color_scheme.write_xml(&mut w);
        self.font_scheme.write_xml(&mut w);
        self.format_scheme.write_xml(&mut w);
        w.close("a:themeElements");
        // objectDefaults（简化：空）
        w.empty("a:objectDefaults");
        // extraClrSchemeLst（简化：空）
        w.empty("a:extraClrSchemeLst");
        w.close("a:theme");
        w.into_string()
    }
}

/// 返回完整且经过验证的 Office 主题 XML（与 python-pptx 默认输出对齐）。
///
/// 这是写路径的默认实现，保证 PowerPoint/WPS 能正确打开。
/// 如需根据 [`Theme`] 结构体生成自定义 XML，使用 [`Theme::to_xml`]。
pub fn default_theme_xml() -> String {
    THEME_XML.to_string()
}

/// 默认的 `<a:fmtScheme>` 元素（从完整 Office 主题中提取）。
///
/// 包含 fillStyleLst / lnStyleLst / effectStyleLst / bgFillStyleLst。
const DEFAULT_FMT_SCHEME: &str = r#"<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme>"#;

/// 完整且经过验证的 Office 主题 XML（与 python-pptx 默认输出对齐）。
const THEME_XML: &str = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:spDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style></a:spDef><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/></a:theme>"##;

#[cfg(test)]
mod tests {
    use super::*;

    /// 验证 `Theme::to_xml` 能生成合法的 theme XML。
    #[test]
    fn theme_to_xml_generates_valid_xml() {
        let mut theme = Theme {
            name: "Test Theme".to_string(),
            ..Default::default()
        };
        theme.color_scheme.name = "Test".to_string();
        theme.color_scheme.dk1 = Some(ThemeColor::Srgb("000000".to_string()));
        theme.color_scheme.lt1 = Some(ThemeColor::Srgb("FFFFFF".to_string()));
        theme.font_scheme.major_latin = "Arial".to_string();
        theme.font_scheme.minor_latin = "Calibri".to_string();

        let xml = theme.to_xml();
        // 验证关键元素存在
        assert!(xml.contains(r#"name="Test Theme""#), "theme name missing");
        assert!(
            xml.contains(r#"<a:clrScheme name="Test">"#),
            "clrScheme missing"
        );
        assert!(
            xml.contains(r#"<a:dk1><a:srgbClr val="000000"/></a:dk1>"#),
            "dk1 missing"
        );
        assert!(
            xml.contains(r#"<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>"#),
            "lt1 missing"
        );
        assert!(
            xml.contains(r#"<a:latin typeface="Arial"/>"#),
            "major latin missing"
        );
        assert!(
            xml.contains(r#"<a:latin typeface="Calibri"/>"#),
            "minor latin missing"
        );
        // 验证 fmtScheme 存在（默认）
        assert!(xml.contains("<a:fmtScheme"), "fmtScheme missing");
    }

    /// 验证 `ThemeColor::write_xml` 对 Srgb 类型的序列化。
    #[test]
    fn theme_color_srgb_write_xml() {
        let mut w = XmlWriter::new();
        let color = ThemeColor::Srgb("FF0000".to_string());
        color.write_xml(&mut w);
        assert_eq!(w.buf, r#"<a:srgbClr val="FF0000"/>"#);
    }

    /// 验证 `ThemeColor::write_xml` 对 Sys 类型的序列化。
    #[test]
    fn theme_color_sys_write_xml() {
        let mut w = XmlWriter::new();
        let color = ThemeColor::Sys("windowText".to_string(), "000000".to_string());
        color.write_xml(&mut w);
        assert_eq!(w.buf, r#"<a:sysClr val="windowText" lastClr="000000"/>"#);
    }

    /// 验证 `ThemeColor::write_xml` 对 None 类型不输出任何内容。
    #[test]
    fn theme_color_none_write_xml() {
        let mut w = XmlWriter::new();
        let color = ThemeColor::None;
        color.write_xml(&mut w);
        assert_eq!(w.buf, "");
    }

    /// 验证 `ColorScheme::write_xml` 按 OOXML 顺序写出 12 个颜色槽位。
    #[test]
    fn color_scheme_write_xml_order() {
        let scheme = ColorScheme {
            name: "Test".to_string(),
            dk1: Some(ThemeColor::Srgb("000000".to_string())),
            lt1: Some(ThemeColor::Srgb("FFFFFF".to_string())),
            accent1: Some(ThemeColor::Srgb("4F81BD".to_string())),
            ..Default::default()
        };

        let mut w = XmlWriter::new();
        scheme.write_xml(&mut w);
        let xml = &w.buf;

        // 验证顺序：dk1 在 lt1 之前，lt1 在 accent1 之前
        let dk1_pos = xml.find("<a:dk1>").unwrap();
        let lt1_pos = xml.find("<a:lt1>").unwrap();
        let accent1_pos = xml.find("<a:accent1>").unwrap();
        assert!(dk1_pos < lt1_pos, "dk1 should come before lt1");
        assert!(lt1_pos < accent1_pos, "lt1 should come before accent1");
    }

    /// 验证 `FontScheme::write_xml` 写出 majorFont 和 minorFont。
    #[test]
    fn font_scheme_write_xml() {
        let fs = FontScheme {
            name: "Test".to_string(),
            major_latin: "Calibri Light".to_string(),
            major_ea: "宋体".to_string(),
            minor_latin: "Calibri".to_string(),
            minor_ea: "宋体".to_string(),
            ..Default::default()
        };

        let mut w = XmlWriter::new();
        fs.write_xml(&mut w);
        let xml = &w.buf;

        assert!(
            xml.contains(r#"<a:latin typeface="Calibri Light"/>"#),
            "major latin missing"
        );
        assert!(
            xml.contains(r#"<a:ea typeface="宋体"/>"#),
            "ea typeface missing"
        );
        assert!(
            xml.contains(r#"<a:latin typeface="Calibri"/>"#),
            "minor latin missing"
        );
    }

    /// 验证 `default_theme_xml` 返回非空且包含关键元素。
    #[test]
    fn default_theme_xml_is_valid() {
        let xml = default_theme_xml();
        assert!(!xml.is_empty());
        assert!(xml.contains("<a:theme"));
        assert!(xml.contains("<a:clrScheme"));
        assert!(xml.contains("<a:fontScheme"));
        assert!(xml.contains("<a:fmtScheme"));
    }

    // ===================== TODO-005：FormatScheme 结构化解析测试 =====================

    /// 构造一份简化的 fmtScheme raw_xml，用于测试结构化解析。
    ///
    /// 包含完整的 4 个 style 列表容器，每个容器有 1-2 个子元素。
    fn sample_fmt_scheme_raw_xml() -> &'static str {
        r#"<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000"/></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>"#
    }

    /// 验证 `parse_from_raw_xml` 能正确拆分 4 个 style 列表。
    ///
    /// 每个列表的子元素数量与原始 XML 中一致：
    /// - fillStyleLst: 2 个（solidFill + gradFill）
    /// - lnStyleLst: 2 个（两个 ln）
    /// - effectStyleLst: 1 个（effectStyle）
    /// - bgFillStyleLst: 1 个（solidFill）
    #[test]
    fn fmt_scheme_parse_from_raw_xml_splits_four_lists() {
        let mut fmt = FormatScheme {
            name: "Test".to_string(),
            raw_xml: sample_fmt_scheme_raw_xml().to_string(),
            ..Default::default()
        };

        // 解析前为空
        assert_eq!(fmt.fill_style_count(), 0);
        assert_eq!(fmt.line_style_count(), 0);
        assert_eq!(fmt.effect_style_count(), 0);
        assert_eq!(fmt.bg_fill_style_count(), 0);

        fmt.parse_from_raw_xml();

        // 解析后正确拆分
        assert_eq!(fmt.fill_style_count(), 2, "fillStyleLst 应有 2 个子元素");
        assert_eq!(fmt.line_style_count(), 2, "lnStyleLst 应有 2 个子元素");
        assert_eq!(
            fmt.effect_style_count(),
            1,
            "effectStyleLst 应有 1 个子元素"
        );
        assert_eq!(
            fmt.bg_fill_style_count(),
            1,
            "bgFillStyleLst 应有 1 个子元素"
        );
    }

    /// 验证解析后的子元素 XML 内容正确（含完整开闭标签与属性）。
    #[test]
    fn fmt_scheme_parsed_children_content_correct() {
        let mut fmt = FormatScheme {
            raw_xml: sample_fmt_scheme_raw_xml().to_string(),
            ..Default::default()
        };
        fmt.parse_from_raw_xml();

        // 第一个 fill 子元素是 solidFill
        let fill0 = &fmt.fill_styles[0];
        assert!(
            fill0.starts_with("<a:solidFill>"),
            "fill[0] 应以 <a:solidFill> 开头"
        );
        assert!(
            fill0.contains(r#"<a:schemeClr val="phClr"/>"#),
            "fill[0] 应包含 schemeClr"
        );
        assert!(
            fill0.ends_with("</a:solidFill>"),
            "fill[0] 应以 </a:solidFill> 结尾"
        );

        // 第二个 fill 子元素是 gradFill（带属性 rotWithShape="1"）
        let fill1 = &fmt.fill_styles[1];
        assert!(
            fill1.starts_with(r#"<a:gradFill rotWithShape="1">"#),
            "fill[1] 应包含属性"
        );

        // 第一个 ln 子元素带 w/cap 属性
        let ln0 = &fmt.line_styles[0];
        assert!(
            ln0.starts_with(r#"<a:ln w="9525" cap="flat">"#),
            "ln[0] 应带属性"
        );

        // effectStyle 子元素嵌套 outerShdw
        let eff0 = &fmt.effect_styles[0];
        assert!(eff0.contains("<a:outerShdw"), "effect[0] 应包含 outerShdw");
    }

    /// 验证结构化 `write_xml` 输出正确的 `<a:fillStyleLst>` 等容器。
    #[test]
    fn fmt_scheme_structured_write_xml_outputs_containers() {
        let mut fmt = FormatScheme {
            name: "Office".to_string(),
            raw_xml: sample_fmt_scheme_raw_xml().to_string(),
            ..Default::default()
        };
        fmt.parse_from_raw_xml();

        let mut w = XmlWriter::new();
        fmt.write_xml(&mut w);
        let xml = &w.buf;

        // 验证 4 个容器都存在且按 OOXML 顺序
        let fill_pos = xml.find("<a:fillStyleLst>").expect("fillStyleLst missing");
        let ln_pos = xml.find("<a:lnStyleLst>").expect("lnStyleLst missing");
        let eff_pos = xml
            .find("<a:effectStyleLst>")
            .expect("effectStyleLst missing");
        let bg_pos = xml
            .find("<a:bgFillStyleLst>")
            .expect("bgFillStyleLst missing");

        // 顺序：fillStyleLst → lnStyleLst → effectStyleLst → bgFillStyleLst
        assert!(fill_pos < ln_pos, "fillStyleLst 应在 lnStyleLst 之前");
        assert!(ln_pos < eff_pos, "lnStyleLst 应在 effectStyleLst 之前");
        assert!(eff_pos < bg_pos, "effectStyleLst 应在 bgFillStyleLst 之前");

        // 验证根元素带 name 属性
        assert!(
            xml.contains(r#"<a:fmtScheme name="Office">"#),
            "fmtScheme 应带 name 属性"
        );

        // 验证容器内有子元素内容（不是空容器）
        assert!(xml.contains("<a:solidFill>"), "应包含 solidFill 子元素");
        assert!(xml.contains("<a:ln "), "应包含 ln 子元素");
    }

    /// 验证查询方法（fill_style_count 等）在解析后返回正确数量。
    #[test]
    fn fmt_scheme_count_methods_after_parse() {
        let mut fmt = FormatScheme {
            raw_xml: sample_fmt_scheme_raw_xml().to_string(),
            ..Default::default()
        };
        fmt.parse_from_raw_xml();

        assert_eq!(fmt.fill_style_count(), 2);
        assert_eq!(fmt.line_style_count(), 2);
        assert_eq!(fmt.effect_style_count(), 1);
        assert_eq!(fmt.bg_fill_style_count(), 1);
    }

    /// 验证 `parse_from_raw_xml` 在 raw_xml 为空时不做任何操作。
    #[test]
    fn fmt_scheme_parse_empty_raw_xml_is_noop() {
        let mut fmt = FormatScheme::default();
        // raw_xml 默认为空
        fmt.parse_from_raw_xml();

        assert_eq!(fmt.fill_style_count(), 0);
        assert_eq!(fmt.line_style_count(), 0);
        assert_eq!(fmt.effect_style_count(), 0);
        assert_eq!(fmt.bg_fill_style_count(), 0);
    }

    /// 验证 `write_xml` 在结构化字段为空但 raw_xml 非空时回退到 raw_xml。
    #[test]
    fn fmt_scheme_write_xml_falls_back_to_raw_xml() {
        let fmt = FormatScheme {
            name: "Office".to_string(),
            raw_xml: r#"<a:customChild/>"#.to_string(),
            ..Default::default()
        };
        // 不调用 parse_from_raw_xml，结构化字段保持为空

        let mut w = XmlWriter::new();
        fmt.write_xml(&mut w);
        let xml = &w.buf;

        // 应输出 <a:fmtScheme name="Office"> + raw_xml + </a:fmtScheme>
        assert!(
            xml.contains(r#"<a:fmtScheme name="Office">"#),
            "应使用 raw_xml 路径"
        );
        assert!(xml.contains("<a:customChild/>"), "应包含 raw_xml 内容");
    }

    /// 验证 `write_xml` 在 raw_xml 与结构化字段都为空时使用默认 Office 格式方案。
    #[test]
    fn fmt_scheme_write_xml_uses_default_when_both_empty() {
        let fmt = FormatScheme::default();
        // name / raw_xml / 4 个结构化字段全部为空

        let mut w = XmlWriter::new();
        fmt.write_xml(&mut w);
        let xml = &w.buf;

        // 应使用 DEFAULT_FMT_SCHEME（直接 raw 输出，无 a:fmtScheme 包裹）
        assert!(xml.contains("<a:fmtScheme"), "应包含默认 fmtScheme");
        assert!(
            xml.contains("<a:fillStyleLst>"),
            "默认方案应含 fillStyleLst"
        );
        assert!(xml.contains("<a:lnStyleLst>"), "默认方案应含 lnStyleLst");
        assert!(
            xml.contains("<a:effectStyleLst>"),
            "默认方案应含 effectStyleLst"
        );
        assert!(
            xml.contains("<a:bgFillStyleLst>"),
            "默认方案应含 bgFillStyleLst"
        );
    }

    /// 验证默认 Office 主题 XML 中的 fmtScheme 能被正确结构化解析。
    ///
    /// 这是 round-trip 测试：默认方案的 fillStyleLst 应有 3 个元素，
    /// lnStyleLst 应有 3 个，effectStyleLst 应有 3 个，bgFillStyleLst 应有 3 个。
    #[test]
    fn fmt_scheme_parse_default_office_format_scheme() {
        // 从默认 Office 主题 XML 中提取 fmtScheme 内部内容
        let theme_xml = default_theme_xml();
        let fmt_start = theme_xml
            .find("<a:fmtScheme")
            .expect("默认主题应包含 <a:fmtScheme");
        let fmt_end = theme_xml
            .find("</a:fmtScheme>")
            .expect("默认主题应包含 </a:fmtScheme>");
        // 提取 <a:fmtScheme> 到 </a:fmtScheme> 之间的内部内容（不含根元素）
        let inner_start = theme_xml[fmt_start..].find('>').unwrap() + fmt_start + 1;
        let fmt_inner = &theme_xml[inner_start..fmt_end];

        let mut fmt = FormatScheme {
            raw_xml: fmt_inner.to_string(),
            ..Default::default()
        };
        fmt.parse_from_raw_xml();

        // 默认 Office 主题：每个 style 列表都有 3 个子元素
        assert_eq!(
            fmt.fill_style_count(),
            3,
            "默认 fillStyleLst 应有 3 个子元素"
        );
        assert_eq!(fmt.line_style_count(), 3, "默认 lnStyleLst 应有 3 个子元素");
        assert_eq!(
            fmt.effect_style_count(),
            3,
            "默认 effectStyleLst 应有 3 个子元素"
        );
        assert_eq!(
            fmt.bg_fill_style_count(),
            3,
            "默认 bgFillStyleLst 应有 3 个子元素"
        );
    }

    /// 验证结构化字段可被直接修改（替换某个 fill style）。
    #[test]
    fn fmt_scheme_structured_fields_are_mutable() {
        let mut fmt = FormatScheme {
            raw_xml: sample_fmt_scheme_raw_xml().to_string(),
            ..Default::default()
        };
        fmt.parse_from_raw_xml();

        // 替换第一个 fill style 为自定义 XML
        fmt.fill_styles[0] = r#"<a:noFill/>"#.to_string();

        let mut w = XmlWriter::new();
        fmt.write_xml(&mut w);
        let xml = &w.buf;

        // 验证替换后的内容出现在输出中
        assert!(
            xml.contains("<a:noFill/>"),
            "替换后的 fill style 应出现在输出"
        );
        // 验证其他 fill style 仍然存在
        assert!(xml.contains("<a:gradFill"), "其他 fill style 应保留");
    }

    /// 验证自闭合子元素（Empty 事件）能被正确收集为完整 XML 字符串。
    #[test]
    fn fmt_scheme_collect_self_closing_children() {
        // raw_xml 中 fillStyleLst 内只有自闭合子元素
        let raw = r#"<a:fillStyleLst><a:noFill/><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>"#;
        let children = collect_style_lst_children(raw, "fillStyleLst");
        assert_eq!(children.len(), 2, "应有 2 个子元素");
        // 第一个是自闭合，应被重构为 <a:noFill/>
        assert_eq!(children[0], "<a:noFill/>", "自闭子元素应重构为完整 XML");
        // 第二个是带子元素的
        assert!(
            children[1].starts_with("<a:solidFill>"),
            "第二个子元素应为 solidFill"
        );
    }

    /// 验证 `local_name_quick` 正确处理带命名空间前缀的元素名。
    #[test]
    fn local_name_quick_handles_namespaced_elements() {
        assert_eq!(local_name_quick(b"a:fillStyleLst"), b"fillStyleLst");
        assert_eq!(local_name_quick(b"fillStyleLst"), b"fillStyleLst");
        assert_eq!(local_name_quick(b"a:ln"), b"ln");
    }

    /// 验证 `collect_style_lst_children` 在容器不存在时返回空列表。
    #[test]
    fn collect_style_lst_children_returns_empty_when_container_missing() {
        let raw = r#"<a:otherLst><a:child/></a:otherLst>"#;
        let children = collect_style_lst_children(raw, "fillStyleLst");
        assert!(children.is_empty(), "容器不存在时应返回空列表");
    }
}
