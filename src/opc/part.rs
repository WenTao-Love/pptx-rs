//! `Part` 与 `PartName`：OPC 包内的逻辑条目。
//!
//! 本文件定义 OPC 容器层最细粒度的"零件"——**Part**。一个 `.pptx`
//! 文件本质上是若干个 Part 组成的扁平命名空间，由 Content-Types 描述
//! 媒体类型，由 `.rels` 描述相互链接。
//!
//! # 设计要点
//!
//! - `PartName` 是对 OPC 规范的"PartName"（绝对路径）语义的 NewType 包装；
//! - `Part` 自身**只**持字节内容（`blob`），不解析 XML；
//! - 解析与序列化交由 `oxml::parser` / `oxml::writer` 负责。
//!
//! 这与 python-pptx 中 `Part` 是"基类 + 各 part 子类"的 OOP 风格不同——
//! 本库采用 **trait-object free** 的 Rust 风格：调用方持 `Part`，在需要
//! 时按 `content_type` 自行反序列化为具体 `oxml::*` 结构体。

use std::borrow::Cow;
use std::fmt;

use super::content_types::ContentTypes;

/// Part 逻辑路径，形式如 `/ppt/slides/slide1.xml`。
///
/// 字符串以 `/` 开头，不含 `..`，以 `/` 分隔。NewType 包装确保
/// 调用方不会"误用相对路径"。
#[derive(Clone, Eq, PartialEq, Ord, PartialOrd, Hash)]
pub struct PartName(String);

impl PartName {
    /// 用裸路径字符串构造，不做合法性校验。
    ///
    /// 适用于内部已知合法的场景；外部输入请用 [`PartName::new`]。
    pub fn from_unchecked(s: impl Into<String>) -> Self {
        let s = s.into();
        PartName(if s.starts_with('/') {
            s
        } else {
            format!("/{}", s)
        })
    }

    /// 校验并构造。
    ///
    /// # 错误
    /// - [`PartNameError::NotAbsolute`]：不以 `/` 开头；
    /// - [`PartNameError::EmptySegment`]：含空段（双斜杠）；
    /// - [`PartNameError::RelSegment`]：含 `.` 或 `..` 段。
    pub fn new(s: &str) -> Result<Self, PartNameError> {
        if !s.starts_with('/') {
            return Err(PartNameError::NotAbsolute);
        }
        for seg in s.split('/').skip(1) {
            if seg.is_empty() {
                return Err(PartNameError::EmptySegment);
            }
            if seg == "." || seg == ".." {
                return Err(PartNameError::RelSegment);
            }
        }
        Ok(PartName(s.to_string()))
    }

    /// 字符串视图。
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }
    /// 取出内部字符串（消费 self）。
    #[inline]
    pub fn into_string(self) -> String {
        self.0
    }

    /// 转成 zip 内的相对路径（去掉前导 `/`）。
    #[inline]
    pub fn to_zip_path(&self) -> &str {
        &self.0[1..]
    }

    /// 取扩展名（含 `.`，小写），无扩展名时为空串。
    pub fn ext(&self) -> String {
        match self.0.rfind('.') {
            Some(i) => self.0[i..].to_ascii_lowercase(),
            None => String::new(),
        }
    }

    /// 同级目录下新文件名：取父目录，拼上 `filename`。
    pub fn sibling(&self, filename: &str) -> PartName {
        let p = self.0.rfind('/').unwrap_or(0);
        PartName(format!("{}/{}", &self.0[..p], filename))
    }

    /// 父目录 part name。已为根时返回 `None`。
    pub fn parent(&self) -> Option<PartName> {
        let p = self.0.rfind('/')?;
        if p == 0 {
            return None;
        }
        Some(PartName(self.0[..p].to_string()))
    }
}

impl fmt::Debug for PartName {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "PartName({})", self.0)
    }
}

impl fmt::Display for PartName {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::str::FromStr for PartName {
    type Err = PartNameError;
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        PartName::new(s)
    }
}

/// Part 名称错误。
#[derive(Debug, thiserror::Error)]
pub enum PartNameError {
    /// 必须以 `/` 开头。
    #[error("part name must start with '/'")]
    NotAbsolute,
    /// 包含空段（`//`）。
    #[error("part name contains empty segment")]
    EmptySegment,
    /// 包含相对段 `.` 或 `..`。
    #[error("part name contains relative segment '.' or '..'")]
    RelSegment,
}

/// 快速构造一个 PartName：未做合法性校验。
///
/// 等价于 [`PartName::from_unchecked`]。为方便函数签名紧凑，单独提供。
pub fn new_part_name(s: &str) -> PartName {
    PartName::from_unchecked(s)
}

/// 一个 OPC part。
///
/// Part 包含：
/// - 名称 [`PartName`]（绝对路径）；
/// - 内容类型 `content_type`（如
///   `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`）；
/// - `blob`（字节内容，未解析）。
#[derive(Debug, Clone)]
pub struct Part {
    /// Part 名称（绝对路径）。
    pub partname: PartName,
    /// Content-Type 字符串。
    pub content_type: String,
    /// 原始字节内容。
    pub blob: Vec<u8>,
}

impl Part {
    /// 构造一个 part。
    pub fn new<N, C, B>(partname: N, content_type: C, blob: B) -> Self
    where
        N: Into<PartName>,
        C: Into<String>,
        B: Into<Vec<u8>>,
    {
        Part {
            partname: partname.into(),
            content_type: content_type.into(),
            blob: blob.into(),
        }
    }

    /// blob 转字符串（UTF-8）。
    ///
    /// 若不是合法 UTF-8 返回 `None`，调用方应自行决定是否报错。
    pub fn blob_text(&self) -> Option<Cow<'_, str>> {
        match std::str::from_utf8(&self.blob) {
            Ok(s) => Some(Cow::Borrowed(s)),
            Err(_) => None,
        }
    }

    /// blob 长度（字节）。
    pub fn len(&self) -> usize {
        self.blob.len()
    }
    /// blob 是否为空。
    pub fn is_empty(&self) -> bool {
        self.blob.is_empty()
    }

    /// 通知 [`ContentTypes`] 自己存在：根据扩展名或 override 添加相应记录。
    ///
    /// - `.rels` 由 default ContentType 处理（`application/vnd...relationships+xml`）
    /// - `.xml` 与其它扩展都显式注册 override，PowerPoint 严格要求每个 XML part 都有显式 Content-Type。
    pub fn contribute_to(&self, ct: &mut ContentTypes) {
        let ext = self.partname.ext();
        if ext == ".rels" {
            // 使用 default 中已经注册的 relationships Content-Type
            return;
        }
        // .xml 与其它扩展都显式注册 override
        ct.add_override(self.partname.as_str(), &self.content_type);
    }
}

/// 允许 `PathBuf::from(part)` 直接把 Part 转成本地路径（zip 内相对路径）。
impl From<Part> for std::path::PathBuf {
    fn from(p: Part) -> std::path::PathBuf {
        std::path::PathBuf::from(p.partname.to_zip_path())
    }
}
