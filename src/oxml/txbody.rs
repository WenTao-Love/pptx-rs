//! 文本体（`<p:txBody>`）：段落、Run、字体属性等。
//!
//! 对应 python-pptx 中 `text.py`/`txbody.py`/`font.py` 三层对象的合并。
//!
//! # 元素结构（OOXML 规范）
//!
//! ```text
//! <p:txBody>
//!   <a:bodyPr .../>            文本框属性
//!   <a:p>                       段落
//!     <a:pPr .../>              段落属性
//!     <a:r>                     Run（可多个）
//!       <a:rPr .../>            Run 属性
//!       <a:t>文本</a:t>
//!     </a:r>
//!     <a:endParaRPr .../>       段落末尾默认属性
//!   </a:p>
//! </p:txBody>
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.text.text._Paragraph` ←→ [`Paragraph`]；
//! - `pptx.text.text._Run` ←→ [`Run`]；
//! - `pptx.text.text._ParagraphFormat` ←→ [`ParagraphProperties`]；
//! - `pptx.text.text._Font` ←→ [`RunProperties`]；
//! - `pptx.text.textframe.TextFrame` ←→ [`TextBody`]。
//!
//! # 序列化约束
//!
//! - Run 属性若全部为默认（无 size/bold/color/...）则**不**写出 `<a:rPr/>`；
//! - 段落若整段为空也必须输出（PowerPoint 期望至少一个 `<a:p>`）；
//! - `<a:t>` 内的文本必须用 [`super::writer::XmlWriter::text`] 自动转义。

use crate::oxml::color::Color;
use crate::oxml::simpletypes::{
    Alignment, MsoAnchor, MsoAutoSize, TabAlignment, TextWrapping, Underline,
};
use crate::units::{Emu, Pt, RGBColor};

/// 制表位（`<a:tab pos="..." algn="..."/>`）。
///
/// 对标 python-pptx `_TabStop` 对象。
/// 一个制表位由位置（EMU）和对齐类型（左/居中/右/小数点）决定。
#[derive(Copy, Clone, Debug, Default)]
pub struct TabStop {
    /// 制表位位置（EMU）。
    pub pos: Emu,
    /// 对齐类型。
    pub alignment: TabAlignment,
}

/// 段落水平缩进 / 悬挂缩进（EMU）。
#[derive(Copy, Clone, Debug, Default)]
pub struct Indent {
    /// 左边缩进（EMU）。
    pub left: Option<Emu>,
    /// 右边缩进（EMU）。
    pub right: Option<Emu>,
    /// 首行缩进（EMU）。
    pub first_line: Option<Emu>,
    /// 自定义悬挂量（in 100ths of a line）。
    pub hanging: Option<i32>,
}

/// 项目符号样式（`<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` 等）。
///
/// 对应 python-pptx 中 `_ParagraphFormat.bullet` 的详细控制。
#[derive(Clone, Debug, Default)]
pub enum BulletStyle {
    /// 无项目符号（`<a:buNone/>`）。
    #[default]
    None,
    /// 自定义字符项目符号（`<a:buChar char="..."/>`）。
    Char {
        /// 项目符号字符（如 `"•"` / `"▪"` / `"→"`）。
        char: String,
    },
    /// 自动编号（`<a:buAutoNum type="..." startAt="..."/>`）。
    AutoNum {
        /// 编号类型（如 `"arabicPeriod"` / `"alphaLcParenR"` / `"romanLcParenBoth"`）。
        auto_num_type: String,
        /// 起始编号（可选）。
        start_at: Option<u32>,
    },
}

/// 段落属性 `<a:pPr>`。
#[derive(Clone, Debug, Default)]
pub struct ParagraphProperties {
    /// 水平对齐方式（`<a:pPr algn="...">`）。`None` 表示不写出 algn 属性。
    pub alignment: Option<Alignment>,
    /// 缩进（左/右/首行/悬挂）。
    pub indent: Indent,
    /// 行距（固定值，百分之一磅）。与 `line_spacing_pct` 互斥。`None` 表示不写出。
    pub line_spacing: Option<i32>,
    /// 行距（百分比，1000 = 100%）。与 `line_spacing` 互斥。`None` 表示不写出。
    pub line_spacing_pct: Option<i32>,
    /// 段前间距（EMU）。`None` 表示不写出。
    pub space_before: Option<Emu>,
    /// 段后间距（EMU）。`None` 表示不写出。
    pub space_after: Option<Emu>,
    /// 是否项目符号列表（兼容旧字段，建议使用 `bullet_style`）。
    pub bullet: bool,
    /// 项目符号详细样式（`<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>`）。
    pub bullet_style: Option<BulletStyle>,
    /// 段落级别（0-8）。
    pub level: u8,
    /// 默认 Run 属性（缺省时应用于段落末尾）。
    pub default_run_properties: Option<RunProperties>,
    /// 制表位列表（`<a:tabLst><a:tab .../></a:tabLst>`）。
    ///
    /// 对标 python-pptx `_ParagraphFormat.tab_stops`。
    /// 空列表表示不写出 `<a:tabLst>`。
    pub tab_stops: Vec<TabStop>,
}

impl ParagraphProperties {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        if self.alignment.is_none()
            && self.indent.left.is_none()
            && self.indent.right.is_none()
            && self.indent.first_line.is_none()
            && self.indent.hanging.is_none()
            && self.line_spacing.is_none()
            && self.line_spacing_pct.is_none()
            && self.space_before.is_none()
            && self.space_after.is_none()
            && !self.bullet
            && self.bullet_style.is_none()
            && self.level == 0
            && self.default_run_properties.is_none()
            && self.tab_stops.is_empty()
        {
            return;
        }
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if let Some(a) = self.alignment {
            attrs.push(("algn", a.as_str()));
        }
        // 元组赋值：先取出所有要用的字符串到块外，扩展生命周期
        let lvl_s = if self.level != 0 {
            Some(self.level.to_string())
        } else {
            None
        };
        if let Some(s) = &lvl_s {
            attrs.push(("lvl", s));
        }
        w.open_with("a:pPr", &attrs);
        // indent
        if self.indent.left.is_some()
            || self.indent.right.is_some()
            || self.indent.first_line.is_some()
            || self.indent.hanging.is_some()
        {
            // 把所有字符串先取到块外
            let l_s = self.indent.left.map(|v| v.value().to_string());
            let r_s = self.indent.right.map(|v| v.value().to_string());
            let first_s = self.indent.first_line.map(|v| v.value().to_string());
            let hang_s = self.indent.hanging.map(|v| v.to_string());
            let mut iattrs: Vec<(&str, &str)> = Vec::new();
            if let Some(s) = &l_s {
                iattrs.push(("l", s));
            }
            if let Some(s) = &r_s {
                iattrs.push(("r", s));
            }
            if let Some(s) = &first_s {
                iattrs.push(("firstLine", s));
            }
            if let Some(s) = &hang_s {
                iattrs.push(("hanging", s));
            }
            w.empty_with("a:indent", &iattrs);
        }
        if self.line_spacing.is_some() || self.line_spacing_pct.is_some() {
            // 注意：OOXML 中 <a:lnSpc> 内只能含一个子元素（<a:spcPct> 或 <a:spcPts>），
            // 因此当同时设置 pct 和 pts 时必须**分别**输出两个 <a:lnSpc> 块，
            // 而不能把多个子元素塞到同一个 <a:lnSpc> 里。
            if let Some(p) = self.line_spacing_pct {
                let s = p.to_string();
                w.open("a:lnSpc");
                w.empty_with("a:spcPct", &[("val", &s)]);
                w.close("a:lnSpc");
            }
            if let Some(sp) = self.line_spacing {
                let s = sp.to_string();
                w.open("a:lnSpc");
                w.empty_with("a:spcPts", &[("val", &s)]);
                w.close("a:lnSpc");
            }
        }
        if self.space_before.is_some() || self.space_after.is_some() {
            // 把所有字符串先取到块外
            let sb_s = self.space_before.map(|v| v.value().to_string());
            let sa_s = self.space_after.map(|v| v.value().to_string());
            let mut sattrs: Vec<(&str, &str)> = Vec::new();
            if let Some(s) = &sb_s {
                sattrs.push(("spcBef", s));
            }
            if let Some(s) = &sa_s {
                sattrs.push(("spcAft", s));
            }
            w.empty_with("a:spcBef", &sattrs);
        }
        // 项目符号样式（OOXML 顺序：buNone/buChar/buAutoNum 在 defRPr 之前）
        if let Some(bs) = &self.bullet_style {
            match bs {
                BulletStyle::None => {
                    w.empty("a:buNone");
                }
                BulletStyle::Char { char } => {
                    w.empty_with("a:buChar", &[("char", char.as_str())]);
                }
                BulletStyle::AutoNum {
                    auto_num_type,
                    start_at,
                } => {
                    if let Some(sa) = start_at {
                        let sa_s = sa.to_string();
                        w.empty_with(
                            "a:buAutoNum",
                            &[("type", auto_num_type.as_str()), ("startAt", sa_s.as_str())],
                        );
                    } else {
                        w.empty_with("a:buAutoNum", &[("type", auto_num_type.as_str())]);
                    }
                }
            }
        }
        // 制表位列表（OOXML 顺序：tabLst 在 bullet 之后、defRPr 之前）
        if !self.tab_stops.is_empty() {
            w.open("a:tabLst");
            for tab in &self.tab_stops {
                let pos_s = tab.pos.value().to_string();
                w.empty_with(
                    "a:tab",
                    &[("pos", pos_s.as_str()), ("algn", tab.alignment.as_str())],
                );
            }
            w.close("a:tabLst");
        }
        if let Some(rpr) = &self.default_run_properties {
            rpr.write_xml(w, "a:defRPr");
        }
        w.close("a:pPr");
    }
}

/// 超链接（`<a:hlinkClick>` / `<a:hlinkHover>`）。
///
/// 对应 python-pptx `Hyperlink` 对象。通过 `r:id` 引用 `rels` 中的目标 URL。
#[derive(Clone, Debug, Default)]
pub struct Hyperlink {
    /// 关系 ID（`r:id`，指向 `.rels` 中的 target URL）。
    pub rid: Option<String>,
    /// 鼠标悬停提示（`tooltip` 属性）。
    pub tooltip: Option<String>,
    /// 动作类型（`action` 属性，如 `"ppaction://hlinksldjump"` 跳转幻灯片）。
    pub action: Option<String>,
    /// 是否无效的超链接（仅写出 `<a:hlinkClick/>` 无属性，用于继承）。
    pub invalid: bool,
}

impl Hyperlink {
    /// 创建一个指向 URL 的超链接。
    pub fn new(rid: impl Into<String>) -> Self {
        Self {
            rid: Some(rid.into()),
            ..Default::default()
        }
    }

    /// 创建一个跳转幻灯片的动作超链接。
    pub fn new_slide_jump() -> Self {
        Self {
            action: Some("ppaction://hlinksldjump".to_string()),
            ..Default::default()
        }
    }
}

/// Run 属性 `<a:rPr>`。
#[derive(Clone, Debug, Default)]
pub struct RunProperties {
    /// 字号（Pt）。`None` 表示走主题继承。
    pub size: Option<Pt>,
    /// 是否加粗。
    pub bold: bool,
    /// 是否斜体。
    pub italic: bool,
    /// 下划线样式。`None` 表示不下划线。
    pub underline: Option<Underline>,
    /// 单删除线。
    pub strike: bool,
    /// 双删除线。
    pub strike_dbl: bool,
    /// 文本颜色。`Color::None` 表示走主题继承。
    pub color: Color,
    /// 高亮背景色。`None` 表示不高亮。
    pub highlight: Option<Color>,
    /// 字体名（拉丁）。
    pub latin_font: Option<String>,
    /// 东亚字体。
    pub eastasia_font: Option<String>,
    /// 复杂脚本字体。
    pub cs_font: Option<String>,
    /// baseline 偏移（百分比，正=上标，负=下标）。
    pub baseline: Option<i32>,
    /// 字距（百分之一磅）。
    pub kerning: Option<i32>,
    /// 字符间距。
    pub spc: Option<i32>,
    /// 大写（small caps / all caps）。
    pub caps: Caps,
    /// `lang`：英文/中文/...语言。
    pub lang: Option<String>,
    /// 透明度（0-100000，0=不透明，100000=完全透明）。
    ///
    /// 对标 DrawingML `<a:solidFill><a:srgbClr val="..."><a:alpha val="30000"/></a:srgbClr></a:solidFill>`。
    /// 常用值：`30_000` = 30% 不透明（70% 透明），用于水印等场景。
    pub alpha: Option<i32>,
    /// 点击超链接（`<a:hlinkClick>`）。
    pub hlink_click: Option<Hyperlink>,
    /// 悬停超链接（`<a:hlinkHover>`）。
    pub hlink_hover: Option<Hyperlink>,
}

#[derive(Copy, Clone, Debug, Eq, PartialEq, Default)]
pub enum Caps {
    #[default]
    None,
    Small,
    All,
}

impl Caps {
    pub fn as_str(self) -> &'static str {
        match self {
            Caps::None => "none",
            Caps::Small => "small",
            Caps::All => "all",
        }
    }
}

impl RunProperties {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter, tag: &str) {
        // 提前取出所有要序列化的字符串，扩展到函数末尾
        let sz_s = self
            .size
            .map(|sz| ((sz.value() * 100.0) as i32).to_string());
        let baseline_s = self.baseline.map(|v| v.to_string());
        let kerning_s = self.kerning.map(|v| v.to_string());
        let spc_s = self.spc.map(|v| v.to_string());

        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if let Some(s) = &sz_s {
            attrs.push(("sz", s));
        }
        if self.bold {
            attrs.push(("b", "1"));
        }
        if self.italic {
            attrs.push(("i", "1"));
        }
        if let Some(u) = self.underline {
            attrs.push(("u", u.as_str()));
        }
        if self.strike {
            attrs.push(("strike", "sngStrike"));
        }
        if self.strike_dbl {
            attrs.push(("strike", "dblStrike"));
        }
        if self.caps != Caps::None {
            attrs.push(("cap", self.caps.as_str()));
        }
        if let Some(s) = &baseline_s {
            attrs.push(("baseline", s));
        }
        if let Some(s) = &kerning_s {
            attrs.push(("kern", s));
        }
        if let Some(s) = &spc_s {
            attrs.push(("spc", s));
        }
        if let Some(l) = &self.lang {
            attrs.push(("lang", l));
        }
        // 严格 OOXML 顺序: <a:ln> <a:noFill>/<a:solidFill> ... <a:highlight> <a:latin> ...
        w.open_with(tag, &attrs);
        if !matches!(self.color, Color::None) {
            self.color.write_solid_fill_with_alpha(w, self.alpha);
        } else if self.alpha.is_some() {
            // 有 alpha 但无颜色——写一个空 solidFill + alpha（罕见但合法）
            w.open("a:solidFill");
            w.empty_with("a:srgbClr", &[("val", "000000")]);
            if let Some(a) = self.alpha {
                w.empty_with("a:alpha", &[("val", a.to_string().as_str())]);
            }
            w.close("a:srgbClr");
            w.close("a:solidFill");
        }
        if let Some(h) = &self.highlight {
            h.write_solid_fill(w);
        }
        if let Some(latin) = &self.latin_font {
            w.empty_with("a:latin", &[("typeface", latin)]);
        }
        if let Some(ea) = &self.eastasia_font {
            w.empty_with("a:ea", &[("typeface", ea)]);
        }
        if let Some(cs) = &self.cs_font {
            w.empty_with("a:cs", &[("typeface", cs)]);
        }
        // 超链接（OOXML 顺序：hlinkClick → hlinkHover，在字体之后）
        if let Some(h) = &self.hlink_click {
            let mut hattrs: Vec<(&str, &str)> = Vec::new();
            if let Some(rid) = &h.rid {
                hattrs.push(("r:id", rid.as_str()));
            }
            if let Some(tip) = &h.tooltip {
                hattrs.push(("tooltip", tip.as_str()));
            }
            if let Some(act) = &h.action {
                hattrs.push(("action", act.as_str()));
            }
            if hattrs.is_empty() {
                w.empty("a:hlinkClick");
            } else {
                w.empty_with("a:hlinkClick", &hattrs);
            }
        }
        if let Some(h) = &self.hlink_hover {
            let mut hattrs: Vec<(&str, &str)> = Vec::new();
            if let Some(rid) = &h.rid {
                hattrs.push(("r:id", rid.as_str()));
            }
            if let Some(tip) = &h.tooltip {
                hattrs.push(("tooltip", tip.as_str()));
            }
            if hattrs.is_empty() {
                w.empty("a:hlinkHover");
            } else {
                w.empty_with("a:hlinkHover", &hattrs);
            }
        }
        w.close(tag);
    }

    /// 从一段 XML 属性中解析（用于读取时）。`src` 来自父级解析器，限于 rPr 属性。
    #[doc(hidden)]
    pub fn from_attrs_unused(_attrs: &super::parser::AttrMap) -> Self {
        // 暂留接口；读路径走 [`crate::oxml::parse_sld::parse_run_properties`]
        RunProperties::default()
    }

    // --------------------- 高阶便捷 API ---------------------

    /// 克隆出 `RunProperties`（含同名字段）。
    pub fn copy(&self) -> Self {
        self.clone()
    }
}

/// 一个 Run：`<a:r>` + 文本。
#[derive(Clone, Debug, Default)]
pub struct Run {
    /// 文本内容。若以 `"\n"` 开头（且不包含其他非空白字符）则被识别为换行 Run。
    pub text: String,
    /// Run 属性。
    pub properties: RunProperties,
}

impl Run {
    /// 构造一个普通 Run。
    pub fn new(text: impl Into<String>) -> Self {
        Run {
            text: text.into(),
            properties: RunProperties::default(),
        }
    }
    /// 构造一个**换行** Run（等价于"软回车"，OOXML 表达为 `<a:br/>`）。
    pub fn line_break() -> Self {
        Run {
            text: String::from("\n"),
            properties: RunProperties::default(),
        }
    }

    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        // 软换行 Run：写出 `<a:br/>`，**不**套 `<a:t>`
        if self.text == "\n" {
            if !is_default_rpr(&self.properties) {
                // 即使有属性也要写出，但 OOXML 规范 `<a:br/>` 在 rPr 内——这里给最简版本
                w.empty("a:br");
            } else {
                w.empty("a:br");
            }
            return;
        }
        w.open("a:r");
        if !is_default_rpr(&self.properties) {
            self.properties.write_xml(w, "a:rPr");
        }
        w.open("a:t");
        w.text(&self.text);
        w.close("a:t");
        w.close("a:r");
    }

    // --------------------- 高阶 API（对齐 python-pptx _Run） ---------------------

    /// 取 Run 文本。
    pub fn text(&self) -> &str {
        &self.text
    }
    /// 设置 Run 文本（**不**支持 `"\n"`，换行请用 [`Run::line_break`]）。
    pub fn set_text(&mut self, t: impl Into<String>) {
        self.text = t.into();
    }

    /// Run 字体大小（`Pt`）。
    pub fn size(&self) -> Option<Pt> {
        self.properties.size
    }
    /// 设置 Run 字体大小。
    pub fn set_size(&mut self, v: Pt) {
        self.properties.size = Some(v);
    }

    /// 是否加粗。
    pub fn bold(&self) -> bool {
        self.properties.bold
    }
    /// 设置加粗。
    pub fn set_bold(&mut self, v: bool) {
        self.properties.bold = v;
    }

    /// 是否斜体。
    pub fn italic(&self) -> bool {
        self.properties.italic
    }
    /// 设置斜体。
    pub fn set_italic(&mut self, v: bool) {
        self.properties.italic = v;
    }

    /// 颜色便捷访问（拷贝）。
    pub fn color(&self) -> Color {
        self.properties.color.clone()
    }
    /// 设置颜色（接受 `RGBColor` / `SchemeColor` / `PresetColor` 任一）。
    pub fn set_color(&mut self, c: impl Into<Color>) {
        self.properties.color = c.into();
    }

    /// 字体名（拉丁 / 主体）。
    pub fn font_name(&self) -> Option<&str> {
        self.properties.latin_font.as_deref()
    }
    /// 设置字体名。
    pub fn set_font_name(&mut self, name: impl Into<String>) {
        self.properties.latin_font = Some(name.into());
    }

    /// 东亚字体名（`<a:ea typeface="..."/>`，对应中文/日文/韩文字体）。
    ///
    /// 与 [`Run::font_name`]（拉丁字体）配对使用：PowerPoint 根据字符脚本
    /// 自动切换 latin/ea 字体。例如同一 Run 内 "Hello 你好" 会用 latin 字体
    /// 渲染 "Hello"，用 ea 字体渲染 "你好"。
    pub fn eastasia_name(&self) -> Option<&str> {
        self.properties.eastasia_font.as_deref()
    }
    /// 设置东亚字体名。
    ///
    /// # 参数
    /// - `name`：东亚字体名称（如 `"宋体"` / `"Microsoft YaHei"` / `"MS Mincho"`）。
    pub fn set_eastasia_name(&mut self, name: impl Into<String>) {
        self.properties.eastasia_font = Some(name.into());
    }

    /// 复杂脚本字体名（`<a:cs typeface="..."/>`，对应阿拉伯语/希伯来语/泰语等）。
    ///
    /// 复杂脚本需要双向排版（RTL）或连字处理，PowerPoint 用 cs 字体渲染这类字符。
    pub fn complex_script_name(&self) -> Option<&str> {
        self.properties.cs_font.as_deref()
    }
    /// 设置复杂脚本字体名。
    ///
    /// # 参数
    /// - `name`：复杂脚本字体名称（如 `"Arial"` / `"Tahoma"` / `"Traditional Arabic"`）。
    pub fn set_complex_script_name(&mut self, name: impl Into<String>) {
        self.properties.cs_font = Some(name.into());
    }

    /// 下划线便捷访问。
    pub fn underline(&self) -> Option<Underline> {
        self.properties.underline
    }
    /// 设置下划线。
    pub fn set_underline(&mut self, v: Underline) {
        self.properties.underline = Some(v);
    }

    /// 是否删除线。
    pub fn strike(&self) -> bool {
        self.properties.strike
    }
    /// 设置删除线。
    pub fn set_strike(&mut self, v: bool) {
        self.properties.strike = v;
    }

    /// 是否双删除线（TODO-017 高阶 API）。
    pub fn double_strike(&self) -> bool {
        self.properties.strike_dbl
    }
    /// 设置双删除线（TODO-017 高阶 API）。
    pub fn set_double_strike(&mut self, v: bool) {
        self.properties.strike_dbl = v;
    }

    /// 取高亮背景色（TODO-018 高阶 API）。
    ///
    /// 返回 `None` 表示未设置高亮。
    pub fn highlight(&self) -> Option<&Color> {
        self.properties.highlight.as_ref()
    }
    /// 设置高亮背景色（TODO-018 高阶 API）。
    ///
    /// 传入 `None` 清除高亮。
    pub fn set_highlight(&mut self, color: Option<Color>) {
        self.properties.highlight = color;
    }
    /// 清除高亮（TODO-018 高阶 API 便捷方法）。
    pub fn clear_highlight(&mut self) {
        self.properties.highlight = None;
    }

    /// 取点击超链接（TODO-026 高阶 API）。
    pub fn hlink_click(&self) -> Option<&Hyperlink> {
        self.properties.hlink_click.as_ref()
    }
    /// 设置点击超链接（TODO-026 高阶 API）。
    pub fn set_hlink_click(&mut self, hl: Hyperlink) {
        self.properties.hlink_click = Some(hl);
    }
    /// 清除点击超链接（TODO-026 高阶 API）。
    pub fn clear_hlink_click(&mut self) {
        self.properties.hlink_click = None;
    }

    /// 取悬停超链接（TODO-026 高阶 API）。
    pub fn hlink_hover(&self) -> Option<&Hyperlink> {
        self.properties.hlink_hover.as_ref()
    }
    /// 设置悬停超链接（TODO-026 高阶 API）。
    pub fn set_hlink_hover(&mut self, hl: Hyperlink) {
        self.properties.hlink_hover = Some(hl);
    }
    /// 清除悬停超链接（TODO-026 高阶 API）。
    pub fn clear_hlink_hover(&mut self) {
        self.properties.hlink_hover = None;
    }

    /// 设置外部 URL 超链接（TODO-026 高阶 API 便捷方法）。
    ///
    /// 对标 python-pptx `run.hyperlink.address = url`。
    ///
    /// # 参数
    /// - `rid`：关系 id（指向 slide `.rels` 中注册的外部 URL）。
    ///   **注意**：本方法只设置 run 的 `hlinkClick`，不负责创建 OPC 关系。
    ///   调用方需自行在 slide 的 `.rels` 中注册该 URL 并取得 rid。
    /// - `tooltip`：可选悬停提示文本；`None` 表示不写出 tooltip 属性。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::oxml::txbody::Run;
    /// let mut run = Run::new("点击");
    /// run.set_hyperlink("rIdHlink1", Some("打开链接"));
    /// ```
    pub fn set_hyperlink(&mut self, rid: impl Into<String>, tooltip: Option<&str>) {
        let mut hl = Hyperlink::new(rid);
        if let Some(t) = tooltip {
            hl.tooltip = Some(t.to_string());
        }
        self.properties.hlink_click = Some(hl);
    }

    /// 设置跳转幻灯片动作超链接（TODO-026 高阶 API 便捷方法）。
    ///
    /// 对标 python-pptx `run.hyperlink.action = "ppaction://hlinksldjump"`。
    ///
    /// # 说明
    /// 仅设置 `action = ppaction://hlinksldjump`，具体跳转目标由 slide `.rels`
    /// 中的关系决定（调用方需自行注册）。若要同时绑定目标 slide，需配合
    /// `Hyperlink::new(rid)` 手动设置 rid + action。
    pub fn set_slide_jump(&mut self) {
        self.properties.hlink_click = Some(Hyperlink::new_slide_jump());
    }

    /// 取 `Font` 高阶视图（用于批量设置字体属性、超链接等）。
    ///
    /// 对标 python-pptx `run.font`。
    pub fn font(&mut self) -> Font<'_> {
        Font::new(&mut self.properties)
    }
}

fn is_default_rpr(p: &RunProperties) -> bool {
    p.size.is_none()
        && !p.bold
        && !p.italic
        && p.underline.is_none()
        && !p.strike
        && !p.strike_dbl
        && matches!(p.color, Color::None)
        && p.highlight.is_none()
        && p.latin_font.is_none()
        && p.eastasia_font.is_none()
        && p.cs_font.is_none()
        && p.baseline.is_none()
        && p.kerning.is_none()
        && p.spc.is_none()
        && p.caps == Caps::None
        && p.lang.is_none()
        && p.alpha.is_none()
        && p.hlink_click.is_none()
        && p.hlink_hover.is_none()
}

/// 字段类型（`<a:fld type="...">`）。
///
/// 对标 python-pptx 中字段类型的常用子集。
/// OOXML 中 `type` 属性的取值较多，这里枚举最常见的几种。
#[derive(Clone, Debug, Default, PartialEq)]
pub enum FieldType {
    /// 幻灯片编号（`type="slidenum"`）。
    #[default]
    SlideNumber,
    /// 日期时间（`type="datetime"`），使用默认格式。
    DateTime,
    /// 日期时间格式 1（`type="datetime1"`，如 `1/1/2024`）。
    DateTime1,
    /// 日期时间格式 2（`type="datetime2"`，如 `Monday, January 1, 2024`）。
    DateTime2,
    /// 日期时间格式 3（`type="datetime3"`，如 `January 1, 2024`）。
    DateTime3,
    /// 页脚（`type="footer"`）。
    Footer,
    /// 自定义字段类型（任意字符串）。
    Custom(String),
}

impl FieldType {
    /// 转 OOXML `type` 属性字面量。
    pub fn as_str(&self) -> &str {
        match self {
            FieldType::SlideNumber => "slidenum",
            FieldType::DateTime => "datetime",
            FieldType::DateTime1 => "datetime1",
            FieldType::DateTime2 => "datetime2",
            FieldType::DateTime3 => "datetime3",
            FieldType::Footer => "footer",
            FieldType::Custom(s) => s.as_str(),
        }
    }

    /// 从 OOXML `type` 属性字面量构造。
    pub fn from_str_value(s: &str) -> Self {
        match s {
            "slidenum" => FieldType::SlideNumber,
            "datetime" => FieldType::DateTime,
            "datetime1" => FieldType::DateTime1,
            "datetime2" => FieldType::DateTime2,
            "datetime3" => FieldType::DateTime3,
            "footer" => FieldType::Footer,
            other => FieldType::Custom(other.to_string()),
        }
    }
}

/// 字段（`<a:fld>`）。
///
/// 对标 python-pptx 中字段对象。字段是段落中特殊的"动态文本"，
/// 如幻灯片编号、日期时间、页脚等，由 PowerPoint 在渲染时自动填充。
///
/// # OOXML 结构
/// ```xml
/// <a:fld id="{GUID}" type="slidenum">
///   <a:rPr .../>
///   <a:t>1</a:t>
/// </a:fld>
/// ```
#[derive(Clone, Debug, Default)]
pub struct Field {
    /// 字段 GUID（`id` 属性）。PowerPoint 用此关联字段实例。
    pub id: String,
    /// 字段类型（`type` 属性）。
    pub field_type: FieldType,
    /// Run 属性（`<a:rPr>`）。
    pub properties: RunProperties,
    /// 字段文本（`<a:t>`）——渲染前的占位文本。
    pub text: String,
}

impl Field {
    /// 创建一个指定类型的字段。
    ///
    /// # 参数
    /// - `field_type`：字段类型；
    /// - `text`：占位文本（如幻灯片编号用 `"1"`，日期用 `"1/1/2024"`）。
    pub fn new(field_type: FieldType, text: impl Into<String>) -> Self {
        Field {
            id: format!("{{{}}}", uuid_like()),
            field_type,
            properties: RunProperties::default(),
            text: text.into(),
        }
    }

    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open_with(
            "a:fld",
            &[("id", self.id.as_str()), ("type", self.field_type.as_str())],
        );
        self.properties.write_xml(w, "a:rPr");
        w.open("a:t");
        w.text(&self.text);
        w.close("a:t");
        w.close("a:fld");
    }
}

/// 生成一个简易的 GUID 风格字符串（非标准 UUID，仅用于字段 id 占位）。
///
/// PowerPoint 对 `id` 属性的格式要求是 `{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}`，
/// 但实际上只要是大括号包裹的唯一字符串即可被正确识别。
/// 这里用时间戳 + 计数器生成，避免引入 uuid 依赖。
fn uuid_like() -> String {
    use std::sync::atomic::{AtomicU64, Ordering};
    static COUNTER: AtomicU64 = AtomicU64::new(1);
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    let ts = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64)
        .unwrap_or(0);
    format!("{:016x}-{:016x}", ts, n)
}

/// 段落中的"结束段落"Run（`</a:p>` 之前可附带 `a:endParaRPr`）。
#[derive(Clone, Debug, Default)]
pub struct Paragraph {
    /// 段落属性。
    pub properties: ParagraphProperties,
    /// 段落中的 Run 序列（按出现顺序）。
    pub runs: Vec<Run>,
    /// 段落中的字段序列（`<a:fld>`）。
    ///
    /// 字段在 Run 之后序列化。OOXML 允许 `<a:r>` 和 `<a:fld>` 交错出现，
    /// 但本实现简化为 runs 先于 fields 输出。
    pub fields: Vec<Field>,
    /// 段落末尾的属性（`a:endParaRPr`）。
    pub end_properties: Option<RunProperties>,
}

impl Paragraph {
    /// 新建一个空段落。
    pub fn new() -> Self {
        Paragraph::default()
    }

    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("a:p");
        self.properties.write_xml(w);
        for r in &self.runs {
            r.write_xml(w);
        }
        for f in &self.fields {
            f.write_xml(w);
        }
        if let Some(rpr) = &self.end_properties {
            rpr.write_xml(w, "a:endParaRPr");
        }
        w.close("a:p");
    }

    // --------------------- 高阶 API（对齐 python-pptx _Paragraph） ---------------------

    /// 新增一个空 Run，返回其可变引用。
    ///
    /// 对应 python-pptx 中 `paragraph.add_run()`。
    pub fn add_run(&mut self) -> &mut Run {
        self.runs.push(Run::default());
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = self.runs.len() - 1;
        &mut self.runs[idx]
    }

    /// 新增一个**带文本**的 Run，返回其可变引用。
    ///
    /// 对应 python-pptx 中 `paragraph.add_run(text)`。
    pub fn add_run_with_text(&mut self, text: impl Into<String>) -> &mut Run {
        self.runs.push(Run::new(text));
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = self.runs.len() - 1;
        &mut self.runs[idx]
    }

    /// 追加一个**换行符 Run**（`<a:br/>`）。
    ///
    /// python-pptx 中由 `paragraph.add_run().text = "\n"` 触发；在
    /// 本库中换行被建模为独立 Run——这样既保留属性语义又简化序列化。
    pub fn add_line_break(&mut self) -> &mut Run {
        self.runs.push(Run {
            text: String::from("\n"),
            properties: RunProperties::default(),
        });
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = self.runs.len() - 1;
        &mut self.runs[idx]
    }

    /// 清空段落中的全部 Run（保留段落属性）。
    ///
    /// 对应 python-pptx 中对 `_Paragraph` 重新赋值的常见用法。
    pub fn clear_runs(&mut self) {
        self.runs.clear();
    }

    /// 新增一个字段（`<a:fld>`），返回其可变引用。
    ///
    /// 对标 python-pptx 中通过 `paragraph.add_field()` 添加幻灯片编号/日期等动态字段。
    ///
    /// # 参数
    /// - `field_type`：字段类型（如 [`FieldType::SlideNumber`]）；
    /// - `text`：占位文本（如 `"1"`）。
    pub fn add_field(&mut self, field_type: FieldType, text: impl Into<String>) -> &mut Field {
        self.fields.push(Field::new(field_type, text));
        let idx = self.fields.len() - 1;
        &mut self.fields[idx]
    }

    /// 清空段落中的全部字段。
    pub fn clear_fields(&mut self) {
        self.fields.clear();
    }

    /// 替换为单段单 Run 的纯文本（保留段落属性）。
    ///
    /// 与 `TextBox::set_text` 不同，**不**按 `\n` 切分多段；
    /// 适合"修改 Run 文本"场景。
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Run {
        self.runs.clear();
        self.runs.push(Run::new(text));
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = self.runs.len() - 1;
        &mut self.runs[idx]
    }

    /// 取段落全部 Run 文本拼接（不含 `\n`）。
    pub fn text(&self) -> String {
        let mut out = String::new();
        for r in &self.runs {
            // 把 Run 内含的 "\n" 视作软换行 → 真实换行
            out.push_str(&r.text);
        }
        out
    }

    /// 水平对齐（便捷访问）。
    pub fn alignment(&self) -> Option<Alignment> {
        self.properties.alignment
    }
    /// 设置水平对齐。
    pub fn set_alignment(&mut self, v: Alignment) {
        self.properties.alignment = Some(v);
    }

    /// 段落级别（0-8）。
    pub fn level(&self) -> u8 {
        self.properties.level
    }
    /// 设置段落级别。
    pub fn set_level(&mut self, lvl: u8) {
        self.properties.level = lvl;
    }

    /// 行距（点数，固定值）。与 `set_line_spacing_pct` 互斥。
    pub fn line_spacing(&self) -> Option<Pt> {
        // OOXML 内部用 EMU（百分之一磅 1pt = 12700 EMU），对外用 Pt
        self.properties
            .line_spacing
            .map(|emu| Pt(emu as f64 / 12_700.0))
    }
    /// 设置行距为**固定点数**（1 pt = 12700 EMU）。
    pub fn set_line_spacing(&mut self, v: Pt) {
        let emu = (v.value() * 12_700.0) as i32;
        self.properties.line_spacing = Some(emu);
        self.properties.line_spacing_pct = None;
    }

    /// 行距（百分比，1.0 = 100%）。与 `set_line_spacing` 互斥。
    pub fn line_spacing_pct(&self) -> Option<f32> {
        self.properties.line_spacing_pct.map(|v| v as f32 / 1000.0)
    }
    /// 设置行距为**倍数**（1.0 = 100% = `1000`，1.5 = 150% = `1500`）。
    pub fn set_line_spacing_pct(&mut self, v: f32) {
        // python-pptx 的 line_spacing = 1.5 ⇒ 150000；这里用 v * 1000 与 OOXML 百分位对齐
        self.properties.line_spacing_pct = Some((v * 1000.0) as i32);
        self.properties.line_spacing = None;
    }

    /// 段前间距（EMU）。
    pub fn space_before(&self) -> Option<Emu> {
        self.properties.space_before
    }
    /// 设置段前间距。
    pub fn set_space_before(&mut self, emu: Emu) {
        self.properties.space_before = Some(emu);
    }

    /// 段后间距（EMU）。
    pub fn space_after(&self) -> Option<Emu> {
        self.properties.space_after
    }
    /// 设置段后间距。
    pub fn set_space_after(&mut self, emu: Emu) {
        self.properties.space_after = Some(emu);
    }

    /// 缩进便捷访问。
    pub fn indent(&self) -> Indent {
        self.properties.indent
    }
    /// 设置缩进（一次性给 4 个字段）。
    pub fn set_indent(
        &mut self,
        left: Option<Emu>,
        right: Option<Emu>,
        first_line: Option<Emu>,
        hanging: Option<i32>,
    ) {
        self.properties.indent = Indent {
            left,
            right,
            first_line,
            hanging,
        };
    }

    // --------------------- 段落末尾属性（endParaRPr，TODO-047） ---------------------

    /// 取段落末尾属性（`<a:endParaRPr>`）的不可变引用。
    ///
    /// `endParaRPr` 用于指定段落末尾（最后一个 Run 之后）的默认 Run 属性。
    /// PowerPoint 在用户把光标移到段落末尾时，会用此属性渲染后续输入的文本。
    ///
    /// 对应 OOXML 元素 `<a:endParaRPr>`，结构与 `<a:rPr>` 完全一致。
    pub fn end_para_rpr(&self) -> Option<&RunProperties> {
        self.end_properties.as_ref()
    }

    /// 取段落末尾属性的可变引用。
    pub fn end_para_rpr_mut(&mut self) -> Option<&mut RunProperties> {
        self.end_properties.as_mut()
    }

    /// 设置段落末尾属性（`<a:endParaRPr>`）。
    ///
    /// 若段落已有 `endParaRPr`，会被覆盖。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::oxml::txbody::{Paragraph, RunProperties};
    /// # use pptx::Pt;
    /// let mut p = Paragraph::new();
    /// let mut rpr = RunProperties::default();
    /// rpr.size = Some(Pt(24.0));
    /// rpr.latin_font = Some("Calibri".to_string());
    /// p.set_end_para_rpr(rpr);
    /// assert!(p.end_para_rpr().is_some());
    /// ```
    pub fn set_end_para_rpr(&mut self, rpr: RunProperties) {
        self.end_properties = Some(rpr);
    }

    /// 清除段落末尾属性（删除 `<a:endParaRPr>`）。
    pub fn clear_end_para_rpr(&mut self) {
        self.end_properties = None;
    }
}

/// 文本体 `<p:txBody>`。
///
/// 对应 python-pptx 的 `TextFrame` 类：提供段落列表 + 自动调整 / 边距 /
/// 垂直对齐 / 换行等"文本框级"属性。
#[derive(Clone, Debug, Default)]
pub struct TextBody {
    pub body_properties: Option<BodyProperties>,
    pub paragraphs: Vec<Paragraph>,
}

impl TextBody {
    pub fn new() -> Self {
        TextBody::default()
    }

    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:txBody");
        if let Some(bp) = &self.body_properties {
            bp.write_xml(w);
        }
        for p in &self.paragraphs {
            p.write_xml(w);
        }
        w.close("p:txBody");
    }

    // --------------------- 高阶 API（对齐 python-pptx TextFrame） ---------------------

    /// 取所有文本（段间 `\n`）。
    pub fn text(&self) -> String {
        let mut out = String::new();
        for (i, p) in self.paragraphs.iter().enumerate() {
            if i > 0 {
                out.push('\n');
            }
            for r in &p.runs {
                out.push_str(&r.text);
            }
        }
        out
    }

    /// 取首个段落（`TextFrame.paragraphs[0]` 的等价物）。
    pub fn first_paragraph(&self) -> Option<&Paragraph> {
        self.paragraphs.first()
    }

    /// 取首个段落的可变引用。
    pub fn first_paragraph_mut(&mut self) -> Option<&mut Paragraph> {
        self.paragraphs.first_mut()
    }

    /// 取第 idx 个段落的不可变引用（越界返回 None）。
    pub fn paragraph(&self, idx: usize) -> Option<&Paragraph> {
        self.paragraphs.get(idx)
    }

    /// 取第 idx 个段落的可变引用（越界返回 None）。
    pub fn paragraph_mut(&mut self, idx: usize) -> Option<&mut Paragraph> {
        self.paragraphs.get_mut(idx)
    }

    /// 新增一个空段落并返回其可变引用。
    ///
    /// 对应 `TextFrame.add_paragraph()`。
    pub fn add_paragraph(&mut self) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::new());
        // push 后直接按索引取最后一个元素，避免 expect
        let idx = self.paragraphs.len() - 1;
        &mut self.paragraphs[idx]
    }

    /// 新增一个**带文本**的段落，返回其可变引用。
    ///
    /// 文本按 `\n` 切分为多段，与 python-pptx 的 `TextFrame.text = "a\nb"` 行为对齐。
    pub fn add_paragraph_with_text(&mut self, text: &str) -> &mut Paragraph {
        for line in text.split('\n') {
            let mut p = Paragraph::new();
            p.runs.push(Run::new(line));
            self.paragraphs.push(p);
        }
        // split('\n') 至少产生一个元素（空字符串也会产生一段），所以 paragraphs 非空
        let idx = self.paragraphs.len() - 1;
        &mut self.paragraphs[idx]
    }

    /// 移除第 idx 个段落。
    pub fn remove_paragraph(&mut self, idx: usize) -> Option<Paragraph> {
        if idx < self.paragraphs.len() {
            Some(self.paragraphs.remove(idx))
        } else {
            None
        }
    }

    /// 清空段落（保留一个空段，对应 python-pptx `TextFrame.clear()` 语义）。
    pub fn clear(&mut self) {
        self.paragraphs.clear();
        self.paragraphs.push(Paragraph::new());
    }

    /// 整体替换文本（按 `\n` 切分为多段，丢弃原段落属性）。
    ///
    /// 对应 python-pptx `text_frame.text = "a\nb"`。
    pub fn set_text(&mut self, text: &str) {
        self.paragraphs.clear();
        for line in text.split('\n') {
            let mut p = Paragraph::new();
            p.runs.push(Run::new(line));
            self.paragraphs.push(p);
        }
    }

    // --------------------- BodyProperties 便捷访问 ---------------------

    /// 取/建 `body_properties`。
    fn ensure_body_properties(&mut self) -> &mut BodyProperties {
        if self.body_properties.is_none() {
            self.body_properties = Some(BodyProperties::default());
        }
        // 安全：上方 if 确保了 body_properties 为 Some；使用 match 避免 expect
        match &mut self.body_properties {
            Some(bp) => bp,
            None => unreachable!("body_properties was just initialized above"),
        }
    }

    /// 自动调整策略（`TextFrame.auto_size`）。
    pub fn auto_size(&self) -> MsoAutoSize {
        self.body_properties
            .as_ref()
            .and_then(|b| b.auto_size())
            .unwrap_or(MsoAutoSize::None)
    }
    /// 设置自动调整策略。
    pub fn set_auto_size(&mut self, v: MsoAutoSize) {
        self.ensure_body_properties().set_auto_size(v);
    }

    /// 垂直对齐（`TextFrame.vertical_anchor`）。
    pub fn vertical_anchor(&self) -> Option<MsoAnchor> {
        self.body_properties.as_ref().and_then(|b| b.anchor)
    }
    /// 设置垂直对齐。
    pub fn set_vertical_anchor(&mut self, v: MsoAnchor) {
        self.ensure_body_properties().anchor = Some(v);
    }

    /// 是否自动换行（`TextFrame.word_wrap`）。
    ///
    /// 返回 `None` 表示走默认值 / 继承。
    pub fn word_wrap(&self) -> Option<bool> {
        self.body_properties.as_ref().and_then(|b| match b.wrap {
            Some(TextWrapping::Square) => Some(true),
            Some(TextWrapping::None) => Some(false),
            None => None,
        })
    }
    /// 设置是否自动换行。
    pub fn set_word_wrap(&mut self, v: bool) {
        self.ensure_body_properties().wrap = Some(if v {
            TextWrapping::Square
        } else {
            TextWrapping::None
        });
    }

    /// 左边距（EMU）。
    pub fn margin_left(&self) -> Option<Emu> {
        self.body_properties
            .as_ref()
            .and_then(|b| b.insets.as_ref().map(|i| i.left))
    }
    /// 右边距。
    pub fn margin_right(&self) -> Option<Emu> {
        self.body_properties
            .as_ref()
            .and_then(|b| b.insets.as_ref().map(|i| i.right))
    }
    /// 上边距。
    pub fn margin_top(&self) -> Option<Emu> {
        self.body_properties
            .as_ref()
            .and_then(|b| b.insets.as_ref().map(|i| i.top))
    }
    /// 下边距。
    pub fn margin_bottom(&self) -> Option<Emu> {
        self.body_properties
            .as_ref()
            .and_then(|b| b.insets.as_ref().map(|i| i.bottom))
    }
    /// 一次性设置四向边距。
    pub fn set_margins(&mut self, left: Emu, top: Emu, right: Emu, bottom: Emu) {
        let bp = self.ensure_body_properties();
        bp.insets = Some(Inset {
            left,
            top,
            right,
            bottom,
        });
    }
    /// 设置左边距。
    pub fn set_margin_left(&mut self, emu: Emu) {
        let bp = self.ensure_body_properties();
        let i = bp.insets.get_or_insert(Inset::default());
        i.left = emu;
    }
    /// 设置右边距。
    pub fn set_margin_right(&mut self, emu: Emu) {
        let bp = self.ensure_body_properties();
        let i = bp.insets.get_or_insert(Inset::default());
        i.right = emu;
    }
    /// 设置上边距。
    pub fn set_margin_top(&mut self, emu: Emu) {
        let bp = self.ensure_body_properties();
        let i = bp.insets.get_or_insert(Inset::default());
        i.top = emu;
    }
    /// 设置下边距。
    pub fn set_margin_bottom(&mut self, emu: Emu) {
        let bp = self.ensure_body_properties();
        let i = bp.insets.get_or_insert(Inset::default());
        i.bottom = emu;
    }

    /// 文本列数（`numCol` 属性）。
    ///
    /// 返回 `None` 表示未设置（默认单列）；`Some(1)` 等价于单列。
    pub fn num_cols(&self) -> Option<u32> {
        self.body_properties.as_ref().and_then(|b| b.num_cols)
    }
    /// 设置文本列数。
    pub fn set_num_cols(&mut self, count: u32) {
        let bp = self.ensure_body_properties();
        bp.num_cols = Some(count);
    }
    /// 列间距（`spcCol` 属性，EMU）。
    pub fn col_spacing(&self) -> Option<Emu> {
        self.body_properties.as_ref().and_then(|b| b.col_spacing)
    }
    /// 设置列间距。
    pub fn set_col_spacing(&mut self, emu: Emu) {
        let bp = self.ensure_body_properties();
        bp.col_spacing = Some(emu);
    }
}

/// 文本体属性 `<a:bodyPr>`。
#[derive(Clone, Debug, Default)]
pub struct BodyProperties {
    /// 文本与边界框的内边距。
    pub insets: Option<Inset>,
    /// 文字方向。
    pub vertical: Option<String>,
    /// 文字方向（rotation）。
    pub rotation: Option<i32>,
    /// 自动换行（`wrap="square"/"none"`）。
    pub wrap: Option<TextWrapping>,
    /// 文本框在溢出时的行为。
    pub sp_auto_fit: bool,
    /// 形状自适应文字。
    pub norm_autofit: bool,
    /// 锚定（top/ctr/b）。
    pub anchor: Option<MsoAnchor>,
    /// 水平对齐（l/ctr/just/dist）——仅占位/组合时使用。
    pub anchor_ctr: bool,
    /// 文本列数（`numCol` 属性，1 表示单列即默认）。
    ///
    /// 对应 python-pptx `TextFrame.word_wrap` 相关的多列布局。
    /// 多列时文本从左到右依次填充各列。
    pub num_cols: Option<u32>,
    /// 列间距（`spcCol` 属性，EMU）。
    ///
    /// 仅当 `num_cols > 1` 时有效。
    pub col_spacing: Option<Emu>,
}

/// 文本框内边距（`<a:bodyPr lIns tIns rIns bIns>`）。
///
/// 在 `<a:bodyPr>` 元素上以 4 个独立属性表示，单位 EMU。
/// python-pptx 中 `TextFrame.margin_left/right/top/bottom` 即对应这 4 个值。
#[derive(Copy, Clone, Debug, Default)]
pub struct Inset {
    /// 左内边距（EMU）。
    pub left: Emu,
    /// 上内边距（EMU）。
    pub top: Emu,
    /// 右内边距（EMU）。
    pub right: Emu,
    /// 下内边距（EMU）。
    pub bottom: Emu,
}

impl BodyProperties {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        // 提前取出所有要序列化的字符串，扩展到函数末尾
        let l_s = self.insets.as_ref().map(|i| i.left.value().to_string());
        let t_s = self.insets.as_ref().map(|i| i.top.value().to_string());
        let r_s = self.insets.as_ref().map(|i| i.right.value().to_string());
        let b_s = self.insets.as_ref().map(|i| i.bottom.value().to_string());
        let rot_s = self.rotation.map(|v| v.to_string());
        let wrap_s = self.wrap.map(|v| v.as_str());
        let anchor_s = self.anchor.map(|v| v.as_str());
        let numcol_s = self.num_cols.map(|v| v.to_string());
        let spccol_s = self.col_spacing.map(|v| v.value().to_string());

        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if let Some(s) = &l_s {
            attrs.push(("lIns", s));
        }
        if let Some(s) = &t_s {
            attrs.push(("tIns", s));
        }
        if let Some(s) = &r_s {
            attrs.push(("rIns", s));
        }
        if let Some(s) = &b_s {
            attrs.push(("bIns", s));
        }
        if let Some(v) = &self.vertical {
            attrs.push(("vert", v));
        }
        if let Some(s) = &rot_s {
            attrs.push(("rot", s));
        }
        if let Some(s) = &wrap_s {
            attrs.push(("wrap", s));
        }
        if let Some(s) = &anchor_s {
            attrs.push(("anchor", s));
        }
        if let Some(s) = &numcol_s {
            attrs.push(("numCol", s));
        }
        if let Some(s) = &spccol_s {
            attrs.push(("spcCol", s));
        }
        w.open_with("a:bodyPr", &attrs);
        // 自动调整子元素：优先按 sp_auto_fit / norm_autofit 标志位写出（保留兼容路径）
        if self.sp_auto_fit {
            w.empty("a:spAutoFit");
        } else if self.norm_autofit {
            w.empty("a:normAutofit");
        }
        w.close("a:bodyPr");
    }

    /// 取 [`MsoAutoSize`] 视图（合并 `sp_auto_fit` / `norm_autofit` 两个 bool）。
    pub fn auto_size(&self) -> Option<MsoAutoSize> {
        if self.sp_auto_fit {
            Some(MsoAutoSize::ShapeToFitText)
        } else if self.norm_autofit {
            Some(MsoAutoSize::TextToFitShape)
        } else {
            None
        }
    }
    /// 设置 [`MsoAutoSize`] 视图（同时清空另一个标志位以保持互斥）。
    pub fn set_auto_size(&mut self, v: MsoAutoSize) {
        match v {
            MsoAutoSize::None => {
                self.sp_auto_fit = false;
                self.norm_autofit = false;
            }
            MsoAutoSize::ShapeToFitText => {
                self.sp_auto_fit = true;
                self.norm_autofit = false;
            }
            MsoAutoSize::TextToFitShape => {
                self.sp_auto_fit = false;
                self.norm_autofit = true;
            }
        }
    }
}

impl From<RGBColor> for Color {
    fn from(c: RGBColor) -> Self {
        Color::RGB(c)
    }
}

// ====================================================================
// 高阶 Font 视图（python-pptx 风格 `pptx.text.text.Font`）
// ====================================================================

use crate::oxml::color::ColorFormat;

/// 字体高阶视图（`pptx.text.text.Font`）。
///
/// # 与 python-pptx 的对应
///
/// - `pptx.text.text.Font` ←→ [`Font`]；
/// - `font.bold = True` / `font.size = Pt(24)` / `font.color.rgb = RGBColor(...)`
///   ←→ [`Font::set_bold`] / [`Font::set_size`] / [`Font::color`].set_rgb()。
///
/// # 设计要点
///
/// - **借用 + 透明代理**：构造时传入 `&mut RunProperties`；
/// - **零分配**：颜色走 [`ColorFormat`] 借用；
/// - **可空值语义**：`name` / `size` / `lang` 等字段用 `Option<T>`，
///   `None` 表示"走主题继承"——与 python-pptx `None` 一致。
#[derive(Debug)]
pub struct Font<'a> {
    /// 底层 [`RunProperties`] 引用。
    rpr: &'a mut RunProperties,
}

impl<'a> Font<'a> {
    /// 构造。
    pub fn new(rpr: &'a mut RunProperties) -> Self {
        Font { rpr }
    }
    /// 底层 rpr 不可变引用。
    pub fn rpr(&self) -> &RunProperties {
        self.rpr
    }
    /// 底层 rpr 可变引用。
    pub fn rpr_mut(&mut self) -> &mut RunProperties {
        self.rpr
    }

    /// 颜色 [`ColorFormat`] 代理。
    pub fn color(&mut self) -> ColorFormat<'_> {
        ColorFormat::new(&mut self.rpr.color)
    }

    /// 字号（Pt）。
    pub fn size(&self) -> Option<Pt> {
        self.rpr.size
    }
    /// 设置字号。
    pub fn set_size(&mut self, v: Pt) {
        self.rpr.size = Some(v);
    }
    /// 清空字号（走主题继承）。
    pub fn clear_size(&mut self) {
        self.rpr.size = None;
    }

    /// 加粗。
    pub fn bold(&self) -> bool {
        self.rpr.bold
    }
    /// 设置加粗。
    pub fn set_bold(&mut self, v: bool) {
        self.rpr.bold = v;
    }

    /// 斜体。
    pub fn italic(&self) -> bool {
        self.rpr.italic
    }
    /// 设置斜体。
    pub fn set_italic(&mut self, v: bool) {
        self.rpr.italic = v;
    }

    /// 删除线。
    pub fn strike(&self) -> bool {
        self.rpr.strike
    }
    /// 设置删除线。
    pub fn set_strike(&mut self, v: bool) {
        self.rpr.strike = v;
    }

    /// 双删除线。
    ///
    /// 对标 python-pptx `font._rPr.attrib['strike'] == 'dblStrike'`。
    /// 当为 `true` 时，写出 `strike="dblStrike"`；普通删除线请用 [`Self::set_strike`]。
    pub fn double_strike(&self) -> bool {
        self.rpr.strike_dbl
    }
    /// 设置双删除线。
    pub fn set_double_strike(&mut self, v: bool) {
        self.rpr.strike_dbl = v;
    }

    /// 高亮色。
    ///
    /// 对标 python-pptx `Font.highlight_color`（v0.6.21+）。
    /// 返回 `None` 表示未设置高亮。
    pub fn highlight(&self) -> Option<&Color> {
        self.rpr.highlight.as_ref()
    }
    /// 设置高亮色。
    ///
    /// 传入 `None` 清除高亮；传入 `Color::None` 也清除高亮。
    pub fn set_highlight(&mut self, color: Option<Color>) {
        match color {
            None => self.rpr.highlight = None,
            Some(Color::None) => self.rpr.highlight = None,
            Some(c) => self.rpr.highlight = Some(c),
        }
    }

    /// 下划线。
    pub fn underline(&self) -> Option<Underline> {
        self.rpr.underline
    }
    /// 设置下划线（`None` 表示清除）。
    pub fn set_underline(&mut self, v: Option<Underline>) {
        self.rpr.underline = v;
    }

    /// 字体名（拉丁）。
    pub fn name(&self) -> Option<&str> {
        self.rpr.latin_font.as_deref()
    }
    /// 设置字体名。
    pub fn set_name(&mut self, n: impl Into<String>) {
        self.rpr.latin_font = Some(n.into());
    }
    /// 清空字体名（走主题继承）。
    pub fn clear_name(&mut self) {
        self.rpr.latin_font = None;
    }

    /// 东亚字体。
    pub fn eastasia_name(&self) -> Option<&str> {
        self.rpr.eastasia_font.as_deref()
    }
    /// 设置东亚字体。
    pub fn set_eastasia_name(&mut self, n: impl Into<String>) {
        self.rpr.eastasia_font = Some(n.into());
    }
    /// 清空东亚字体（走主题继承）。
    ///
    /// 对应 OOXML 中删除 `<a:ea>` 元素，PowerPoint 将使用主题的
    /// `minorFont.ea` / `majorFont.ea` 作为回退。
    pub fn clear_eastasia_name(&mut self) {
        self.rpr.eastasia_font = None;
    }

    /// 复杂脚本字体。
    pub fn complex_script_name(&self) -> Option<&str> {
        self.rpr.cs_font.as_deref()
    }
    /// 设置复杂脚本字体。
    pub fn set_complex_script_name(&mut self, n: impl Into<String>) {
        self.rpr.cs_font = Some(n.into());
    }
    /// 清空复杂脚本字体（走主题继承）。
    ///
    /// 对应 OOXML 中删除 `<a:cs>` 元素，PowerPoint 将使用主题的
    /// `minorFont.cs` / `majorFont.cs` 作为回退。
    pub fn clear_complex_script_name(&mut self) {
        self.rpr.cs_font = None;
    }

    /// baseline 偏移（百分比，正=上标，负=下标）。
    pub fn baseline(&self) -> Option<i32> {
        self.rpr.baseline
    }
    /// 设置 baseline。
    pub fn set_baseline(&mut self, v: i32) {
        self.rpr.baseline = Some(v);
    }

    /// 字符间距（百分之一磅）。
    pub fn spacing(&self) -> Option<i32> {
        self.rpr.spc
    }
    /// 设置字符间距。
    pub fn set_spacing(&mut self, v: i32) {
        self.rpr.spc = Some(v);
    }

    // ===== 超链接 API（TODO-026）=====

    /// 取点击超链接（`<a:hlinkClick>`）。
    pub fn hlink_click(&self) -> Option<&Hyperlink> {
        self.rpr.hlink_click.as_ref()
    }
    /// 设置点击超链接（直接传入 [`Hyperlink`]）。
    pub fn set_hlink_click(&mut self, hl: Hyperlink) {
        self.rpr.hlink_click = Some(hl);
    }
    /// 清除点击超链接。
    pub fn clear_hlink_click(&mut self) {
        self.rpr.hlink_click = None;
    }

    /// 取悬停超链接（`<a:hlinkHover>`）。
    pub fn hlink_hover(&self) -> Option<&Hyperlink> {
        self.rpr.hlink_hover.as_ref()
    }
    /// 设置悬停超链接。
    pub fn set_hlink_hover(&mut self, hl: Hyperlink) {
        self.rpr.hlink_hover = Some(hl);
    }
    /// 清除悬停超链接。
    pub fn clear_hlink_hover(&mut self) {
        self.rpr.hlink_hover = None;
    }

    /// 便捷方法：设置一个指向 URL 的点击超链接。
    ///
    /// # 参数
    /// - `rid`：关系 ID（指向 `.rels` 中的目标 URL）；
    /// - `tooltip`：可选的鼠标悬停提示。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx::oxml::txbody::{Run, Font};
    /// # let mut run = Run::new("链接文本");
    /// run.font().set_hyperlink("rId1", Some("点击访问"));
    /// ```
    pub fn set_hyperlink(&mut self, rid: impl Into<String>, tooltip: Option<&str>) {
        let mut hl = Hyperlink::new(rid);
        if let Some(t) = tooltip {
            hl.tooltip = Some(t.to_string());
        }
        self.rpr.hlink_click = Some(hl);
    }

    /// 便捷方法：设置一个跳转幻灯片的动作超链接。
    ///
    /// 对应 OOXML `action="ppaction://hlinksldjump"`。
    pub fn set_slide_jump(&mut self) {
        self.rpr.hlink_click = Some(Hyperlink::new_slide_jump());
    }
}

impl<'a> From<&'a mut RunProperties> for Font<'a> {
    fn from(r: &'a mut RunProperties) -> Self {
        Font::new(r)
    }
}

// ====================================================================
// 高阶 ParagraphFormat 视图（python-pptx 风格 `pptx.text.text._ParagraphFormat`）
// ====================================================================

/// 段落格式高阶视图（`pptx.text.text._ParagraphFormat`）。
///
/// # 与 python-pptx 的对应
///
/// - `paragraph.alignment = PP_ALIGN.CENTER` ←→ [`ParagraphFormat::set_alignment`]；
/// - `paragraph.line_spacing = 1.5` ←→ [`ParagraphFormat::set_line_spacing`]；
/// - `paragraph.space_before = Pt(6)` ←→ [`ParagraphFormat::set_space_before`]。
///
/// # 设计要点
///
/// - **借用 + 透明代理**：构造时传入 `&mut ParagraphProperties`；
/// - **零分配**：纯字段操作；
/// - **互斥语义**：`set_line_spacing` / `set_line_spacing_pct` 互斥；
///   后调用的会清空前者。
#[derive(Debug)]
pub struct ParagraphFormat<'a> {
    /// 底层 [`ParagraphProperties`] 引用。
    ppr: &'a mut ParagraphProperties,
}

impl<'a> ParagraphFormat<'a> {
    /// 构造。
    pub fn new(ppr: &'a mut ParagraphProperties) -> Self {
        ParagraphFormat { ppr }
    }
    /// 底层 ppr 不可变引用。
    pub fn ppr(&self) -> &ParagraphProperties {
        self.ppr
    }
    /// 底层 ppr 可变引用。
    pub fn ppr_mut(&mut self) -> &mut ParagraphProperties {
        self.ppr
    }

    /// 水平对齐。
    pub fn alignment(&self) -> Option<Alignment> {
        self.ppr.alignment
    }
    /// 设置水平对齐。
    pub fn set_alignment(&mut self, v: Alignment) {
        self.ppr.alignment = Some(v);
    }
    /// 清除水平对齐（走默认）。
    pub fn clear_alignment(&mut self) {
        self.ppr.alignment = None;
    }

    /// 段落级别（0-8）。
    pub fn level(&self) -> u8 {
        self.ppr.level
    }
    /// 设置段落级别。
    pub fn set_level(&mut self, lvl: u8) {
        self.ppr.level = lvl;
    }

    /// 行距（固定值，Pt）。
    pub fn line_spacing(&self) -> Option<Pt> {
        self.ppr.line_spacing.map(|emu| Pt(emu as f64 / 12_700.0))
    }
    /// 设置行距为**固定点数**（与 `set_line_spacing_pct` 互斥）。
    pub fn set_line_spacing(&mut self, v: Pt) {
        let emu = (v.value() * 12_700.0) as i32;
        self.ppr.line_spacing = Some(emu);
        self.ppr.line_spacing_pct = None;
    }
    /// 行距（倍数，1.0 = 100%）。
    pub fn line_spacing_pct(&self) -> Option<f32> {
        self.ppr.line_spacing_pct.map(|v| v as f32 / 1000.0)
    }
    /// 设置行距为**倍数**（与 `set_line_spacing` 互斥）。
    pub fn set_line_spacing_pct(&mut self, v: f32) {
        self.ppr.line_spacing_pct = Some((v * 1000.0) as i32);
        self.ppr.line_spacing = None;
    }
    /// 清除行距。
    pub fn clear_line_spacing(&mut self) {
        self.ppr.line_spacing = None;
        self.ppr.line_spacing_pct = None;
    }

    /// 段前（EMU）。
    pub fn space_before(&self) -> Option<Emu> {
        self.ppr.space_before
    }
    /// 设置段前。
    pub fn set_space_before(&mut self, emu: Emu) {
        self.ppr.space_before = Some(emu);
    }
    /// 段后（EMU）。
    pub fn space_after(&self) -> Option<Emu> {
        self.ppr.space_after
    }
    /// 设置段后。
    pub fn set_space_after(&mut self, emu: Emu) {
        self.ppr.space_after = Some(emu);
    }

    /// 一次性设置缩进。
    pub fn set_indent(
        &mut self,
        left: Option<Emu>,
        right: Option<Emu>,
        first_line: Option<Emu>,
        hanging: Option<i32>,
    ) {
        self.ppr.indent = Indent {
            left,
            right,
            first_line,
            hanging,
        };
    }

    /// 缩进（EMU 视图）。
    pub fn indent(&self) -> Indent {
        self.ppr.indent
    }

    // --------- 项目符号样式（TODO-014） ---------

    /// 项目符号样式。
    ///
    /// 返回 `None` 表示未设置（走默认继承）；`Some(BulletStyle::None)` 表示显式无项目符号。
    pub fn bullet_style(&self) -> Option<&BulletStyle> {
        self.ppr.bullet_style.as_ref()
    }

    /// 设置自定义字符项目符号（`<a:buChar char="..."/>`）。
    ///
    /// 对标 python-pptx `paragraph.font` + bullet 字符设置。
    /// 常用字符：`"•"` / `"▪"` / `"→"` / `"○"` / `"♦"`。
    pub fn set_bullet_char(&mut self, ch: impl Into<String>) {
        self.ppr.bullet = true;
        self.ppr.bullet_style = Some(BulletStyle::Char { char: ch.into() });
    }

    /// 设置自动编号项目符号（`<a:buAutoNum type="..." startAt="..."/>`）。
    ///
    /// 对标 python-pptx 编号列表。
    /// 常用 type：`"arabicPeriod"` (1.) / `"alphaLcParenR"` (a)) / `"romanLcParenBoth"` ((i))。
    ///
    /// # 参数
    /// - `auto_num_type`：编号类型字面量。
    /// - `start_at`：起始编号（可选，`None` 表示从 1 开始）。
    pub fn set_bullet_numbered(&mut self, auto_num_type: impl Into<String>, start_at: Option<u32>) {
        self.ppr.bullet = true;
        self.ppr.bullet_style = Some(BulletStyle::AutoNum {
            auto_num_type: auto_num_type.into(),
            start_at,
        });
    }

    /// 清除项目符号（`<a:buNone/>`）。
    pub fn clear_bullet(&mut self) {
        self.ppr.bullet = false;
        self.ppr.bullet_style = Some(BulletStyle::None);
    }

    /// 是否有项目符号（`bullet=true` 且 `bullet_style` 非 `None`）。
    pub fn has_bullet(&self) -> bool {
        self.ppr.bullet
            && self
                .ppr
                .bullet_style
                .as_ref()
                .map(|bs| !matches!(bs, BulletStyle::None))
                .unwrap_or(false)
    }

    // --------- 制表位（TODO-015） ---------

    /// 制表位列表（不可变）。
    ///
    /// 对标 python-pptx `_ParagraphFormat.tab_stops`。
    pub fn tab_stops(&self) -> &[TabStop] {
        &self.ppr.tab_stops
    }

    /// 添加一个制表位（`<a:tab pos="..." algn="..."/>`）。
    ///
    /// 对标 python-pptx `paragraph.tab_stops.add_tab_stop(position, alignment)`。
    ///
    /// # 参数
    /// - `pos`：制表位位置（EMU）；
    /// - `alignment`：对齐类型（左/居中/右/小数点）。
    pub fn add_tab_stop(&mut self, pos: Emu, alignment: TabAlignment) {
        self.ppr.tab_stops.push(TabStop { pos, alignment });
    }

    /// 清除所有制表位。
    pub fn clear_tab_stops(&mut self) {
        self.ppr.tab_stops.clear();
    }
}

impl<'a> From<&'a mut ParagraphProperties> for ParagraphFormat<'a> {
    fn from(p: &'a mut ParagraphProperties) -> Self {
        ParagraphFormat::new(p)
    }
}

// ====================================================================
// 高阶 TextFrame 视图（python-pptx 风格 `pptx.text.textframe.TextFrame`）
// ====================================================================

/// 文本框高阶视图（`pptx.text.textframe.TextFrame`）。
///
/// # 与 python-pptx 的对应
///
/// - `shape.text_frame` ←→ `text_frame_mut().as_text_frame()`（薄包装）；
/// - `text_frame.text` ←→ [`TextFrame::text_getter`] / [`TextFrame::set_text`]；
/// - `text_frame.paragraphs[0]` ←→ [`TextFrame::paragraphs`]；
/// - `text_frame.add_paragraph()` ←→ [`TextFrame::add_paragraph`]；
/// - `text_frame.word_wrap` ←→ [`TextFrame::word_wrap`] / [`TextFrame::set_word_wrap`]；
/// - `text_frame.auto_size` ←→ [`TextFrame::auto_size`] / [`TextFrame::set_auto_size`]；
/// - `text_frame.vertical_anchor` ←→ [`TextFrame::vertical_anchor`] / [`TextFrame::set_vertical_anchor`]；
/// - `text_frame.margin_left/...` ←→ [`TextFrame::margin_left`] / ... / [`TextFrame::set_margins`]。
///
/// # 设计要点
///
/// - **借用**：构造时传入 `&mut TextBody`；
/// - **零分配**：所有方法均不触发堆分配；
/// - **互斥语义**：[`TextFrame::add_paragraph`] 不会自动 trim，已有段落不会被移除。
#[derive(Debug)]
pub struct TextFrame<'a> {
    /// 底层 [`TextBody`] 引用。
    body: &'a mut TextBody,
}

impl<'a> TextFrame<'a> {
    /// 构造。
    pub fn new(body: &'a mut TextBody) -> Self {
        TextFrame { body }
    }

    /// 底层 body 不可变引用。
    pub fn body(&self) -> &TextBody {
        self.body
    }
    /// 底层 body 可变引用。
    pub fn body_mut(&mut self) -> &mut TextBody {
        self.body
    }

    // --------- 段落集合 ---------

    /// 段落数。
    pub fn len(&self) -> usize {
        self.body.paragraphs.len()
    }
    /// 是否为空。
    pub fn is_empty(&self) -> bool {
        self.body.paragraphs.is_empty()
    }
    /// 不可变段落迭代器。
    pub fn paragraphs(&self) -> std::slice::Iter<'_, Paragraph> {
        self.body.paragraphs.iter()
    }
    /// 可变段落迭代器。
    pub fn paragraphs_mut(&mut self) -> std::slice::IterMut<'_, Paragraph> {
        self.body.paragraphs.iter_mut()
    }
    /// 按下标取不可变段落。
    pub fn paragraph(&self, idx: usize) -> Option<&Paragraph> {
        self.body.paragraphs.get(idx)
    }
    /// 按下标取可变段落。
    pub fn paragraph_mut(&mut self, idx: usize) -> Option<&mut Paragraph> {
        self.body.paragraphs.get_mut(idx)
    }
    /// 首个段落。
    pub fn first_paragraph(&self) -> Option<&Paragraph> {
        self.body.paragraphs.first()
    }
    /// 首个段落的可变引用。
    pub fn first_paragraph_mut(&mut self) -> Option<&mut Paragraph> {
        self.body.paragraphs.first_mut()
    }

    /// 新增空段落（python-pptx `text_frame.add_paragraph()`）。
    pub fn add_paragraph(&mut self) -> &mut Paragraph {
        self.body.add_paragraph()
    }

    /// 新增带文本段落（按 `\n` 切分为多段），返回**最后**一段的可变引用。
    pub fn add_paragraph_with_text(&mut self, text: &str) -> &mut Paragraph {
        self.body.add_paragraph_with_text(text)
    }

    /// 按下标移除段落。
    pub fn remove_paragraph(&mut self, idx: usize) -> Option<Paragraph> {
        self.body.remove_paragraph(idx)
    }

    /// 清空段落（保留一个空段，对应 python-pptx `text_frame.clear()`）。
    pub fn clear(&mut self) {
        self.body.clear();
    }

    // --------- 文本便捷 ---------

    /// 整体取文本（段间 `\n`，与 python-pptx `text_frame.text` 一致）。
    pub fn text_getter(&self) -> String {
        self.body.text()
    }

    /// 整体替换文本（按 `\n` 切分；旧段落属性会被丢弃）。
    pub fn set_text(&mut self, text: &str) {
        self.body.set_text(text);
    }

    // --------- BodyProperties 便捷 ---------

    /// 自动调整策略。
    pub fn auto_size(&self) -> MsoAutoSize {
        self.body.auto_size()
    }
    /// 设置自动调整策略。
    pub fn set_auto_size(&mut self, v: MsoAutoSize) {
        self.body.set_auto_size(v);
    }

    /// 垂直对齐。
    pub fn vertical_anchor(&self) -> Option<MsoAnchor> {
        self.body.vertical_anchor()
    }
    /// 设置垂直对齐。
    pub fn set_vertical_anchor(&mut self, v: MsoAnchor) {
        self.body.set_vertical_anchor(v);
    }

    /// 是否自动换行（None = 走默认）。
    pub fn word_wrap(&self) -> Option<bool> {
        self.body.word_wrap()
    }
    /// 设置是否自动换行。
    pub fn set_word_wrap(&mut self, v: bool) {
        self.body.set_word_wrap(v);
    }

    /// 左边距。
    pub fn margin_left(&self) -> Option<Emu> {
        self.body.margin_left()
    }
    /// 右边距。
    pub fn margin_right(&self) -> Option<Emu> {
        self.body.margin_right()
    }
    /// 上边距。
    pub fn margin_top(&self) -> Option<Emu> {
        self.body.margin_top()
    }
    /// 下边距。
    pub fn margin_bottom(&self) -> Option<Emu> {
        self.body.margin_bottom()
    }
    /// 一次性设置四向边距。
    pub fn set_margins(&mut self, l: Emu, t: Emu, r: Emu, b: Emu) {
        self.body.set_margins(l, t, r, b);
    }
    /// 设置单边边距（left）。
    pub fn set_margin_left(&mut self, emu: Emu) {
        self.body.set_margin_left(emu);
    }
    /// 设置单边边距（right）。
    pub fn set_margin_right(&mut self, emu: Emu) {
        self.body.set_margin_right(emu);
    }
    /// 设置单边边距（top）。
    pub fn set_margin_top(&mut self, emu: Emu) {
        self.body.set_margin_top(emu);
    }
    /// 设置单边边距（bottom）。
    pub fn set_margin_bottom(&mut self, emu: Emu) {
        self.body.set_margin_bottom(emu);
    }

    // --------- 多列布局 ---------

    /// 文本列数（`numCol` 属性）。
    ///
    /// `None` 表示未设置（默认单列）；`Some(1)` 等同于单列。
    pub fn num_cols(&self) -> Option<u32> {
        self.body.num_cols()
    }

    /// 设置文本列数（`numCol` 属性）。
    pub fn set_num_cols(&mut self, count: u32) {
        self.body.set_num_cols(count);
    }

    /// 列间距（`spcCol` 属性，EMU）。
    pub fn col_spacing(&self) -> Option<Emu> {
        self.body.col_spacing()
    }

    /// 设置列间距（`spcCol` 属性，EMU）。
    pub fn set_col_spacing(&mut self, emu: Emu) {
        self.body.set_col_spacing(emu);
    }

    /// 一次性设置多列布局（列数 + 可选列间距）。
    ///
    /// 对标 python-pptx 中 `text_frame` 的多列设置。
    /// 当 `spacing` 为 `None` 时仅设置列数，保留已有列间距。
    pub fn set_columns(&mut self, count: u32, spacing: Option<Emu>) {
        self.body.set_num_cols(count);
        if let Some(s) = spacing {
            self.body.set_col_spacing(s);
        }
    }
}

impl<'a> From<&'a mut TextBody> for TextFrame<'a> {
    fn from(b: &'a mut TextBody) -> Self {
        TextFrame::new(b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_paragraph_simple() {
        let mut p = Paragraph::new();
        let mut r = Run::new("Hello");
        r.properties.size = Some(Pt(24.0));
        r.properties.bold = true;
        r.properties.color = RGBColor(0xFF, 0, 0).into();
        p.runs.push(r);
        let mut w = super::super::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("Hello"));
        assert!(s.contains("sz=\"2400\""));
        assert!(s.contains("b=\"1\""));
        assert!(s.contains("a:srgbClr"));
    }

    /// 验证 `TextFrame` / `ParagraphFormat` 视图能联动到底层 `TextBody`。
    #[test]
    fn textframe_view_mirrors_body() {
        let mut tb = TextBody::new();
        {
            let mut tf = TextFrame::new(&mut tb);
            // 走 view 加一段 + 走 view 设 word_wrap
            let p = tf.add_paragraph();
            p.add_run_with_text("first").set_bold(true);
            tf.add_paragraph_with_text("second\nthird");
            tf.set_word_wrap(false);
            tf.set_margins(Emu(91440), Emu(45720), Emu(91440), Emu(45720));
        }
        // 验证 view 写入已生效
        assert_eq!(tb.paragraphs.len(), 3);
        assert_eq!(tb.paragraphs[0].runs[0].text, "first");
        assert!(tb.paragraphs[0].runs[0].properties.bold);
        assert_eq!(tb.paragraphs[1].runs[0].text, "second");
        assert_eq!(tb.paragraphs[2].runs[0].text, "third");
        // 自动换行被设成 false
        assert_eq!(tb.word_wrap(), Some(false));
    }

    /// 验证 `ParagraphFormat` 的 line_spacing 互斥语义。
    #[test]
    fn paragraph_format_line_spacing_mutex() {
        let mut p = Paragraph::new();
        p.set_line_spacing(Pt(20.0));
        assert!(p.line_spacing().is_some());
        assert!(p.line_spacing_pct().is_none());
        // 改设倍数：清空固定值
        p.set_line_spacing_pct(1.5);
        assert!(p.line_spacing().is_none());
        assert_eq!(p.line_spacing_pct(), Some(1.5));
        // 走 view 验证同样互斥
        let mut p2 = Paragraph::new();
        {
            let mut pf = ParagraphFormat::new(&mut p2.properties);
            pf.set_line_spacing(Pt(15.0));
        }
        assert_eq!(p2.line_spacing(), Some(Pt(15.0)));
        {
            let mut pf = ParagraphFormat::new(&mut p2.properties);
            pf.set_line_spacing_pct(2.0);
        }
        assert!(p2.line_spacing().is_none());
        assert_eq!(p2.line_spacing_pct(), Some(2.0));
    }

    /// 验证 `Font` 的删除线/双删除线 API。
    ///
    /// 这是 TODO-017 的测试。
    #[test]
    fn font_strikethrough_api() {
        let mut r = Run::new("text");
        {
            let mut f = Font::new(&mut r.properties);
            f.set_strike(true);
        }
        assert!(r.properties.strike);
        assert!(!r.properties.strike_dbl);

        // 双删除线
        {
            let mut f = Font::new(&mut r.properties);
            f.set_double_strike(true);
        }
        assert!(r.properties.strike_dbl);

        // 验证 Font 读取
        let f = Font::new(&mut r.properties);
        assert!(f.strike());
        assert!(f.double_strike());
    }

    /// 验证 `Font` 的高亮色 API。
    ///
    /// 这是 TODO-018 的测试。
    #[test]
    fn font_highlight_api() {
        let mut r = Run::new("text");
        // 默认无高亮
        assert!(r.properties.highlight.is_none());

        // 设置高亮
        {
            let mut f = Font::new(&mut r.properties);
            f.set_highlight(Some(Color::RGB(RGBColor(0xFF, 0xFF, 0x00))));
        }
        assert!(r.properties.highlight.is_some());

        // 验证 Font 读取
        let f = Font::new(&mut r.properties);
        let hl = f.highlight().expect("应有高亮色");
        assert!(matches!(hl, Color::RGB(c) if c.0 == 0xFF && c.1 == 0xFF && c.2 == 0x00));

        // 清除高亮
        {
            let mut f = Font::new(&mut r.properties);
            f.set_highlight(None);
        }
        assert!(r.properties.highlight.is_none());

        // 用 Color::None 清除
        {
            let mut f = Font::new(&mut r.properties);
            f.set_highlight(Some(Color::RGB(RGBColor(0xFF, 0x00, 0x00))));
            f.set_highlight(Some(Color::None));
        }
        assert!(r.properties.highlight.is_none());
    }

    /// 验证 `TextBody` / `TextFrame` 的多列布局 API 与序列化。
    ///
    /// 这是 TODO-019 的测试。
    #[test]
    fn text_body_multi_column_api() {
        let mut tb = TextBody::new();
        // 默认无列数
        assert!(tb.num_cols().is_none());
        assert!(tb.col_spacing().is_none());

        // 设置 3 列 + 列间距
        tb.set_num_cols(3);
        tb.set_col_spacing(Emu(91440));
        assert_eq!(tb.num_cols(), Some(3));
        assert_eq!(tb.col_spacing(), Some(Emu(91440)));

        // 序列化验证
        let mut w = super::super::writer::XmlWriter::new();
        tb.body_properties
            .as_ref()
            .expect("应有 body_properties")
            .write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("numCol=\"3\""), "应输出 numCol=\"3\"，实际: {s}");
        assert!(
            s.contains("spcCol=\"91440\""),
            "应输出 spcCol=\"91440\"，实际: {s}"
        );
    }

    /// 验证 `TextFrame::set_columns` 一次性设置列数和列间距。
    ///
    /// 这是 TODO-019 的测试。
    #[test]
    fn text_frame_set_columns() {
        let mut tb = TextBody::new();
        {
            let mut tf = TextFrame::new(&mut tb);
            // 设置 2 列 + 列间距
            tf.set_columns(2, Some(Emu(45720)));
        }
        assert_eq!(tb.num_cols(), Some(2));
        assert_eq!(tb.col_spacing(), Some(Emu(45720)));

        // 仅设置列数，不设置列间距
        {
            let mut tf = TextFrame::new(&mut tb);
            tf.set_columns(4, None);
        }
        assert_eq!(tb.num_cols(), Some(4));
        // 列间距应保持上次的值
        assert_eq!(tb.col_spacing(), Some(Emu(45720)));

        // TextFrame 代理方法
        {
            let tf = TextFrame::new(&mut tb);
            assert_eq!(tf.num_cols(), Some(4));
            assert_eq!(tf.col_spacing(), Some(Emu(45720)));
        }
    }

    /// 验证 `ParagraphFormat` 的项目符号字符 API。
    ///
    /// 这是 TODO-014 的测试。
    #[test]
    fn paragraph_format_bullet_char() {
        let mut ppr = ParagraphProperties::default();
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.set_bullet_char("•");
        }
        assert!(ppr.bullet, "bullet 应为 true");
        assert!(pf_has_bullet(&ppr), "has_bullet 应为 true");
        // 验证序列化
        let mut w = super::super::writer::XmlWriter::new();
        ppr.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("buChar"), "应输出 buChar，实际: {s}");
        assert!(s.contains("char=\"•\""), "应包含 char=\"•\"，实际: {s}");
    }

    /// 验证 `ParagraphFormat` 的编号项目符号 API。
    ///
    /// 这是 TODO-014 的测试。
    #[test]
    fn paragraph_format_bullet_numbered() {
        let mut ppr = ParagraphProperties::default();
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.set_bullet_numbered("arabicPeriod", Some(3));
        }
        assert!(ppr.bullet);
        match &ppr.bullet_style {
            Some(BulletStyle::AutoNum {
                auto_num_type,
                start_at,
            }) => {
                assert_eq!(auto_num_type, "arabicPeriod");
                assert_eq!(*start_at, Some(3));
            }
            other => panic!("期望 AutoNum，实际: {other:?}"),
        }
        // 验证序列化
        let mut w = super::super::writer::XmlWriter::new();
        ppr.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("buAutoNum"), "应输出 buAutoNum，实际: {s}");
        assert!(
            s.contains("type=\"arabicPeriod\""),
            "应包含 type，实际: {s}"
        );
        assert!(s.contains("startAt=\"3\""), "应包含 startAt，实际: {s}");
    }

    /// 验证 `ParagraphFormat::clear_bullet` 清除项目符号。
    ///
    /// 这是 TODO-014 的测试。
    #[test]
    fn paragraph_format_clear_bullet() {
        let mut ppr = ParagraphProperties::default();
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.set_bullet_char("•");
        }
        assert!(pf_has_bullet(&ppr));
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.clear_bullet();
        }
        assert!(!ppr.bullet, "bullet 应为 false");
        assert!(!pf_has_bullet(&ppr), "has_bullet 应为 false");
        // 验证序列化输出 buNone
        let mut w = super::super::writer::XmlWriter::new();
        ppr.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("buNone"), "应输出 buNone，实际: {s}");
    }

    /// 辅助函数：检查 ParagraphProperties 是否有项目符号。
    fn pf_has_bullet(ppr: &ParagraphProperties) -> bool {
        ppr.bullet
            && ppr
                .bullet_style
                .as_ref()
                .map(|bs| !matches!(bs, BulletStyle::None))
                .unwrap_or(false)
    }

    /// 验证 `ParagraphFormat` 的制表位 API。
    ///
    /// 这是 TODO-015 的测试。
    #[test]
    fn paragraph_format_tab_stops() {
        let mut ppr = ParagraphProperties::default();
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.add_tab_stop(Emu(914400), TabAlignment::Left);
            pf.add_tab_stop(Emu(1828800), TabAlignment::Right);
            pf.add_tab_stop(Emu(2743200), TabAlignment::Center);
        }
        assert_eq!(ppr.tab_stops.len(), 3);
        assert_eq!(ppr.tab_stops[0].pos.value(), 914400);
        assert_eq!(ppr.tab_stops[0].alignment, TabAlignment::Left);
        assert_eq!(ppr.tab_stops[1].pos.value(), 1828800);
        assert_eq!(ppr.tab_stops[1].alignment, TabAlignment::Right);
        assert_eq!(ppr.tab_stops[2].pos.value(), 2743200);
        assert_eq!(ppr.tab_stops[2].alignment, TabAlignment::Center);

        // 验证序列化
        let mut w = super::super::writer::XmlWriter::new();
        ppr.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("tabLst"), "应输出 tabLst，实际: {s}");
        assert!(s.contains("pos=\"914400\""), "应包含 pos=914400，实际: {s}");
        assert!(s.contains("algn=\"l\""), "应包含 algn=l，实际: {s}");
        assert!(s.contains("algn=\"r\""), "应包含 algn=r，实际: {s}");
        assert!(s.contains("algn=\"ctr\""), "应包含 algn=ctr，实际: {s}");

        // 清除
        {
            let mut pf = ParagraphFormat::new(&mut ppr);
            pf.clear_tab_stops();
        }
        assert!(ppr.tab_stops.is_empty());
    }

    /// 验证 `Field` 的序列化和 `Paragraph::add_field` API。
    ///
    /// 这是 TODO-016 的测试。
    #[test]
    fn paragraph_field_serialization() {
        let mut p = Paragraph::new();
        p.add_field(FieldType::SlideNumber, "1");
        p.add_field(FieldType::DateTime, "1/1/2024");

        assert_eq!(p.fields.len(), 2);
        assert_eq!(p.fields[0].field_type, FieldType::SlideNumber);
        assert_eq!(p.fields[0].text, "1");
        assert_eq!(p.fields[1].field_type, FieldType::DateTime);
        assert_eq!(p.fields[1].text, "1/1/2024");

        // 验证序列化
        let mut w = super::super::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("a:fld"), "应输出 a:fld，实际: {s}");
        assert!(
            s.contains("type=\"slidenum\""),
            "应包含 type=slidenum，实际: {s}"
        );
        assert!(
            s.contains("type=\"datetime\""),
            "应包含 type=datetime，实际: {s}"
        );
        assert!(s.contains(">1<"), "应包含文本 1，实际: {s}");
        assert!(s.contains(">1/1/2024<"), "应包含文本 1/1/2024，实际: {s}");
    }

    /// 验证 `FieldType` 的转换方法。
    ///
    /// 这是 TODO-016 的测试。
    #[test]
    fn field_type_conversion() {
        // as_str
        assert_eq!(FieldType::SlideNumber.as_str(), "slidenum");
        assert_eq!(FieldType::DateTime.as_str(), "datetime");
        assert_eq!(FieldType::DateTime1.as_str(), "datetime1");
        assert_eq!(FieldType::Footer.as_str(), "footer");
        assert_eq!(FieldType::Custom("custom1".to_string()).as_str(), "custom1");

        // from_str_value
        assert_eq!(
            FieldType::from_str_value("slidenum"),
            FieldType::SlideNumber
        );
        assert_eq!(FieldType::from_str_value("datetime"), FieldType::DateTime);
        assert_eq!(FieldType::from_str_value("datetime1"), FieldType::DateTime1);
        assert_eq!(FieldType::from_str_value("footer"), FieldType::Footer);
        assert_eq!(
            FieldType::from_str_value("unknown"),
            FieldType::Custom("unknown".to_string())
        );
    }

    /// 验证 `Font` 超链接高阶 API（set_hyperlink / set_slide_jump）。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn font_hyperlink_api() {
        let mut run = Run::new("链接文本");
        // 设置 URL 超链接
        run.font().set_hyperlink("rId3", Some("点击访问"));
        let hl = run
            .properties
            .hlink_click
            .as_ref()
            .expect("hlink_click 应存在");
        assert_eq!(hl.rid.as_deref(), Some("rId3"));
        assert_eq!(hl.tooltip.as_deref(), Some("点击访问"));

        // 设置跳转幻灯片动作
        run.font().set_slide_jump();
        let hl = run
            .properties
            .hlink_click
            .as_ref()
            .expect("hlink_click 应存在");
        assert_eq!(hl.action.as_deref(), Some("ppaction://hlinksldjump"));

        // 清除
        run.font().clear_hlink_click();
        assert!(run.properties.hlink_click.is_none());
    }

    /// 验证超链接的序列化 round-trip。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn hyperlink_serialization_roundtrip() {
        let mut run = Run::new("点击这里");
        run.font().set_hyperlink("rId7", Some("提示文字"));

        // 序列化
        let mut w = super::super::writer::XmlWriter::new();
        run.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("a:hlinkClick"), "应输出 a:hlinkClick，实际: {s}");
        assert!(s.contains("r:id=\"rId7\""), "应包含 r:id=rId7，实际: {s}");
        assert!(
            s.contains("tooltip=\"提示文字\""),
            "应包含 tooltip，实际: {s}"
        );
    }

    // --------------------- TODO-047: endParaRPr 测试 ---------------------

    /// 验证 `Paragraph::set_end_para_rpr` 高阶 API 与序列化。
    ///
    /// 设置带属性 + 子元素的 endParaRPr，验证序列化输出正确。
    #[test]
    fn end_para_rpr_set_and_serialize() {
        let mut p = Paragraph::new();
        let rpr = RunProperties {
            size: Some(Pt(24.0)),
            latin_font: Some("Calibri".to_string()),
            lang: Some("en-US".to_string()),
            ..Default::default()
        };
        p.set_end_para_rpr(rpr);

        // 验证 getter
        assert!(p.end_para_rpr().is_some());
        let got = p.end_para_rpr().unwrap();
        assert_eq!(got.size, Some(Pt(24.0)));
        assert_eq!(got.latin_font.as_deref(), Some("Calibri"));
        assert_eq!(got.lang.as_deref(), Some("en-US"));

        // 序列化
        let mut w = super::super::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("a:endParaRPr"), "应输出 a:endParaRPr，实际: {s}");
        assert!(s.contains("sz=\"2400\""), "应包含 sz=2400，实际: {s}");
        assert!(s.contains("lang=\"en-US\""), "应包含 lang=en-US，实际: {s}");
        assert!(s.contains("a:latin"), "应包含 a:latin，实际: {s}");
        assert!(
            s.contains("typeface=\"Calibri\""),
            "应包含 typeface=Calibri，实际: {s}"
        );
    }

    /// 验证 `Paragraph::clear_end_para_rpr` 清除 endParaRPr。
    #[test]
    fn end_para_rpr_clear() {
        let mut p = Paragraph::new();
        let rpr = RunProperties {
            size: Some(Pt(18.0)),
            ..Default::default()
        };
        p.set_end_para_rpr(rpr);
        assert!(p.end_para_rpr().is_some());

        p.clear_end_para_rpr();
        assert!(p.end_para_rpr().is_none());

        // 序列化不应包含 endParaRPr
        let mut w = super::super::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let s = w.into_string();
        assert!(
            !s.contains("a:endParaRPr"),
            "不应输出 a:endParaRPr，实际: {s}"
        );
    }

    /// 验证 `Paragraph::end_para_rpr_mut` 可变引用修改。
    #[test]
    fn end_para_rpr_mut_modify() {
        let mut p = Paragraph::new();
        let rpr = RunProperties::default();
        p.set_end_para_rpr(rpr);

        // 通过可变引用修改
        if let Some(rpr) = p.end_para_rpr_mut() {
            rpr.size = Some(Pt(32.0));
            rpr.bold = true;
        }

        let got = p.end_para_rpr().expect("end_para_rpr 应存在");
        assert_eq!(got.size, Some(Pt(32.0)));
        assert!(got.bold);
    }

    /// 验证带子元素的 endParaRPr 的 round-trip（序列化 → 解析 → 比对）。
    ///
    /// 这是 TODO-047 的核心测试：确保带 solidFill/latin 等子元素的
    /// `<a:endParaRPr>` 能被正确解析。
    #[test]
    fn end_para_rpr_roundtrip_with_children() {
        // 1. 构造段落并设置 endParaRPr
        let mut p = Paragraph::new();
        p.add_run_with_text("hello");
        let rpr = RunProperties {
            size: Some(Pt(20.0)),
            bold: true,
            latin_font: Some("Arial".to_string()),
            lang: Some("en-US".to_string()),
            ..Default::default()
        };
        p.set_end_para_rpr(rpr);

        // 2. 序列化
        let mut w = super::super::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let xml = w.into_string();

        // 3. 解析回来
        let parsed = crate::oxml::parse_sld::parse_paragraph(&xml).expect("解析应成功");

        // 4. 验证 endParaRPr 的属性和子元素都被正确解析
        let got = parsed.end_properties.expect("end_properties 应存在");
        assert_eq!(got.size, Some(Pt(20.0)), "size 应为 20pt");
        assert!(got.bold, "bold 应为 true");
        assert_eq!(
            got.latin_font.as_deref(),
            Some("Arial"),
            "latin_font 应为 Arial"
        );
        assert_eq!(got.lang.as_deref(), Some("en-US"), "lang 应为 en-US");
    }

    /// 验证自闭合 endParaRPr（仅属性，无子元素）的解析。
    ///
    /// 这是 TODO-047 的回归测试：确保原有的 Empty 事件处理仍正常工作。
    #[test]
    fn end_para_rpr_self_closing_parse() {
        let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:r><a:t>text</a:t></a:r>
<a:endParaRPr lang="en-US" sz="1800"/>
</a:p>"#;
        let p = crate::oxml::parse_sld::parse_paragraph(xml).expect("解析应成功");
        let got = p.end_properties.expect("end_properties 应存在");
        assert_eq!(got.size, Some(Pt(18.0)));
        assert_eq!(got.lang.as_deref(), Some("en-US"));
    }

    /// 验证带 solidFill 子元素的 endParaRPr 解析。
    #[test]
    fn end_para_rpr_with_solid_fill() {
        let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:endParaRPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:endParaRPr>
</a:p>"#;
        let p = crate::oxml::parse_sld::parse_paragraph(xml).expect("解析应成功");
        let got = p.end_properties.expect("end_properties 应存在");
        assert_eq!(got.size, Some(Pt(24.0)));
        assert_eq!(got.lang.as_deref(), Some("en-US"));
        // 验证 solidFill 子元素被解析
        assert!(
            matches!(got.color, crate::oxml::color::Color::RGB(c) if c.0 == 0xFF && c.1 == 0x00 && c.2 == 0x00)
        );
    }

    // ===================== 东亚字体便捷 API 测试（TODO-005 剩余） =====================

    /// `Run::eastasia_name` / `set_eastasia_name` 往返：设置后能正确读取。
    #[test]
    fn run_eastasia_name_setter_and_getter() {
        let mut r = Run::new("你好");
        assert!(r.eastasia_name().is_none());
        r.set_eastasia_name("宋体");
        assert_eq!(r.eastasia_name(), Some("宋体"));
        assert_eq!(r.properties.eastasia_font.as_deref(), Some("宋体"));
    }

    /// `Run::complex_script_name` / `set_complex_script_name` 往返。
    #[test]
    fn run_complex_script_name_setter_and_getter() {
        let mut r = Run::new("مرحبا");
        assert!(r.complex_script_name().is_none());
        r.set_complex_script_name("Traditional Arabic");
        assert_eq!(r.complex_script_name(), Some("Traditional Arabic"));
        assert_eq!(r.properties.cs_font.as_deref(), Some("Traditional Arabic"));
    }

    /// `Run` 上 latin / ea / cs 三种字体可独立设置，互不干扰。
    #[test]
    fn run_three_fonts_independent() {
        let mut r = Run::new("Hello 你好 مرحبا");
        r.set_font_name("Arial");
        r.set_eastasia_name("Microsoft YaHei");
        r.set_complex_script_name("Tahoma");
        assert_eq!(r.font_name(), Some("Arial"));
        assert_eq!(r.eastasia_name(), Some("Microsoft YaHei"));
        assert_eq!(r.complex_script_name(), Some("Tahoma"));
        // 验证底层 properties 字段独立
        assert_eq!(r.properties.latin_font.as_deref(), Some("Arial"));
        assert_eq!(
            r.properties.eastasia_font.as_deref(),
            Some("Microsoft YaHei")
        );
        assert_eq!(r.properties.cs_font.as_deref(), Some("Tahoma"));
    }

    /// `Font::clear_eastasia_name` 清空东亚字体字段。
    #[test]
    fn font_clear_eastasia_name() {
        let mut r = Run::new("文本");
        r.set_eastasia_name("宋体");
        assert_eq!(r.eastasia_name(), Some("宋体"));
        // 通过 Font 视图清空
        {
            let mut f = Font::new(&mut r.properties);
            f.clear_eastasia_name();
        }
        assert!(r.eastasia_name().is_none());
        assert!(r.properties.eastasia_font.is_none());
    }

    /// `Font::clear_complex_script_name` 清空复杂脚本字体字段。
    #[test]
    fn font_clear_complex_script_name() {
        let mut r = Run::new("text");
        r.set_complex_script_name("Arial");
        assert_eq!(r.complex_script_name(), Some("Arial"));
        // 通过 Font 视图清空
        {
            let mut f = Font::new(&mut r.properties);
            f.clear_complex_script_name();
        }
        assert!(r.complex_script_name().is_none());
        assert!(r.properties.cs_font.is_none());
    }

    /// `Font` 视图的 eastasia/cs 访问器与 `Run` 便捷方法一致。
    #[test]
    fn font_view_eastasia_cs_consistent_with_run() {
        let mut r = Run::new("混合文本");
        r.set_eastasia_name("黑体");
        r.set_complex_script_name("Tahoma");
        // 先用 Run 便捷方法读取（避免与 Font 视图的可变借用冲突）
        let run_ea = r.eastasia_name().map(|s| s.to_string());
        let run_cs = r.complex_script_name().map(|s| s.to_string());
        // Font 视图读取应与 Run 便捷方法一致
        let f = Font::new(&mut r.properties);
        assert_eq!(f.eastasia_name(), run_ea.as_deref());
        assert_eq!(f.complex_script_name(), run_cs.as_deref());
    }

    /// 东亚字体序列化：`<a:ea typeface="..."/>` 应在 `<a:latin>` 之后写出。
    #[test]
    fn eastasia_font_serialization_order() {
        let mut r = Run::new("你好");
        r.set_font_name("Arial");
        r.set_eastasia_name("宋体");
        r.set_complex_script_name("Tahoma");
        let mut p = Paragraph::new();
        p.runs.push(r);
        let mut w = crate::oxml::writer::XmlWriter::new();
        p.write_xml(&mut w);
        let xml = w.into_string();
        // 验证三种字体元素都存在
        assert!(
            xml.contains(r#"<a:latin typeface="Arial"/>"#),
            "xml: {}",
            xml
        );
        assert!(xml.contains(r#"<a:ea typeface="宋体"/>"#), "xml: {}", xml);
        assert!(xml.contains(r#"<a:cs typeface="Tahoma"/>"#), "xml: {}", xml);
        // 验证 OOXML 顺序：latin → ea → cs
        let pos_latin = xml.find("<a:latin").expect("latin should exist");
        let pos_ea = xml.find("<a:ea").expect("ea should exist");
        let pos_cs = xml.find("<a:cs").expect("cs should exist");
        assert!(pos_latin < pos_ea, "latin must come before ea: {}", xml);
        assert!(pos_ea < pos_cs, "ea must come before cs: {}", xml);
    }
}
