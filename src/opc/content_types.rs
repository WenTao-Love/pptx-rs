//! `[Content_Types].xml` 模型。
//!
//! 该文件位于 zip 根目录，是 OPC 包的"媒体类型注册表"。
//! 它由两类元素组成：
//!
//! - `<Default Extension="..." ContentType="..."/>`：按扩展名兜底；
//! - `<Override PartName="..." ContentType="..."/>`：按 partname 精确覆盖。
//!
//! 实际解析时按"override 优先，default 次之"规则匹配。
//!
//! # 序列化
//!
//! - [`ContentTypes::to_xml`] → 序列化；
//! - 解析在 `OpcPackage::load` 流程中通过 [`super::package::parse_content_types_public`] 完成。
//!
//! # 与 OOXML 规范的对应
//!
//! - 命名空间 `http://schemas.openxmlformats.org/package/2006/content-types`；
//! - 默认值（`xml` / `rels` / `png` / `jpeg` / ...）由 [`ContentTypes::new_default`]
//!   预置，符合 PowerPoint 输出的标准文件。

use std::collections::BTreeMap;

/// 默认 Content-Type（按扩展名匹配）。
#[derive(Debug, Clone)]
pub struct DefaultExt {
    /// 扩展名（不带 `.`）。
    pub extension: String,
    /// Content-Type 字符串。
    pub content_type: String,
}

impl DefaultExt {
    /// 构造一条 Default。自动去掉前导 `.`。
    pub fn new(ext: impl Into<String>, ct: impl Into<String>) -> Self {
        DefaultExt {
            extension: ext.into().trim_start_matches('.').to_string(),
            content_type: ct.into(),
        }
    }
}

/// Override：按 part 名称覆盖 Content-Type（通常用于 XML 部件）。
#[derive(Debug, Clone)]
pub struct Override {
    /// Part 名称（绝对路径，含前导 `/`）。
    pub partname: String,
    /// Content-Type 字符串。
    pub content_type: String,
}

/// `[Content_Types].xml` 模型。
///
/// 内部用 `BTreeMap` 索引 `by_partname` 提供 O(log n) 的"是否已 override"查询；
/// `Vec<Override>` 用于序列化时保留插入顺序。
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    /// 默认项（按扩展名）。
    pub defaults: Vec<DefaultExt>,
    /// 覆盖项（按 partname）。
    pub overrides: Vec<Override>,
    by_partname: BTreeMap<String, usize>,
}

impl ContentTypes {
    /// 构造一个带默认项的 ContentTypes（xml/rels/png/jpeg/jpg/gif/bmp/svg）。
    ///
    /// 默认项与 PowerPoint 标准输出对齐，可直接被 Office 解析。
    #[allow(clippy::field_reassign_with_default)]
    pub fn new_default() -> Self {
        let mut ct = ContentTypes::default();
        ct.defaults.push(DefaultExt::new("xml", "application/xml"));
        ct.defaults.push(DefaultExt::new(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        ));
        ct.defaults.push(DefaultExt::new("png", "image/png"));
        ct.defaults.push(DefaultExt::new("jpeg", "image/jpeg"));
        ct.defaults.push(DefaultExt::new("jpg", "image/jpeg"));
        ct.defaults.push(DefaultExt::new("gif", "image/gif"));
        ct.defaults.push(DefaultExt::new("bmp", "image/bmp"));
        ct.defaults.push(DefaultExt::new("svg", "image/svg+xml"));
        ct
    }

    /// 添加或覆盖一个 partname。
    ///
    /// 已存在则**不覆盖**——首个插入的 Content-Type 生效（与 python-pptx 行为一致）。
    pub fn add_override(&mut self, partname: &str, content_type: &str) {
        if self.by_partname.contains_key(partname) {
            return;
        }
        let idx = self.overrides.len();
        self.overrides.push(Override {
            partname: partname.to_string(),
            content_type: content_type.to_string(),
        });
        self.by_partname.insert(partname.to_string(), idx);
    }

    /// 是否已 override。
    pub fn has_override(&self, partname: &str) -> bool {
        self.by_partname.contains_key(partname)
    }

    /// 转成 XML 字符串。
    pub fn to_xml(&self) -> String {
        let mut s = String::with_capacity(512);
        s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        s.push_str(
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
        );
        for d in &self.defaults {
            s.push_str(&format!(
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                super::rels::xml_escape(&d.extension),
                super::rels::xml_escape(&d.content_type),
            ));
        }
        for o in &self.overrides {
            s.push_str(&format!(
                "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                super::rels::xml_escape(&o.partname),
                super::rels::xml_escape(&o.content_type),
            ));
        }
        s.push_str("</Types>");
        s
    }
}
