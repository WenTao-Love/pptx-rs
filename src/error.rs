//! 错误类型与 `Result` 别名。
//!
//! 本模块是整个 `pptx-rs` crate 的错误处理中枢。它做三件事：
//!
//! 1. **统一错误类型** [`enum@Error`]：使用 `thiserror` 派生，将底层 `std::io::Error` /
//!    `zip::result::ZipError` / 自定义字符串错误统一成一个枚举，方便上层 `?` 传播。
//! 2. **统一 Result 别名** [`Result<T>`]：所有公共 API（除特别声明）均返回该别名，
//!    调用方不必书写冗长的 `Result<T, pptx_rs::Error>`。
//! 3. **便捷构造器**：在 `Error` 上提供 `opc(...)` / `oxml(...)` / `not_implemented(...)`
//!    等关联函数，使错误抛出更可读，并隐含语义归类。
//!
//! # 设计原则
//!
//! - **零 `panic!`**：库路径上禁止 `unwrap` / `expect` / `panic!`。所有失败一律
//!   转化为 [`enum@Error`] 的一个变体，由调用方决定如何处理。
//! - **错误消息规范**：消息小写开头、句末无标点（与 Rust 标准库惯例一致）；
//!   需要上下文时使用 `format!` 拼接具体元素名 / 路径，例如
//!   `"relationships parse: missing Id"`。
//! - **可扩展**：新增错误类别时优先扩展 `enum Error` 变体，而非全部塞进
//!   [`Error::Other`]。后者仅用于临时过渡。
//!
//! # 与 python-pptx 的对应
//!
//! `python-pptx` 抛出 `python_pptx.PptxException` 及若干子异常（`PackageNotFoundError` /
//! `XPathOverflowError` 等）。本库以单一枚举 + 字符串消息统一表达，对调用方而言
//! 仅需 match 顶层 `Error` 即可。
//!
//! # 示例
//!
//! ```no_run
//! use pptx_rs::{Error, Result};
//!
//! fn read_slide() -> Result<()> {
//!     let p = std::fs::File::open("missing.pptx")?;  // io::Error 自动转 Error::Io
//!     Ok(())
//! }
//!
//! fn parse_attr() -> Result<()> {
//!     Err(Error::oxml("missing required <p:ph> element"))
//! }
//! ```

use std::io;
use thiserror::Error;

/// 库统一 `Result` 别名。
///
/// 简化签名书写，所有公共 API（除特别声明）均使用该别名，等价于
/// `std::result::Result<T, pptx_rs::Error>`。
///
/// # 示例
///
/// ```no_run
/// use pptx_rs::Result;
///
/// fn read_something() -> Result<String> { Ok(String::new()) }
/// ```
pub type Result<T> = std::result::Result<T, Error>;

/// 库错误。所有外部接口（除特别声明外）均返回 [`Result<T>`]。
///
/// 变体按"来源"分类，调用方可以基于 `match` 决定重试 / 跳过 / 报告策略。
/// 错误消息遵循 **小写开头、句末无标点** 的 Rust 标准库风格。
#[derive(Debug, Error)]
pub enum Error {
    /// I/O 错误：文件不存在、读取失败、写入失败、权限拒绝等。
    ///
    /// 由 `#[from] io::Error` 自动派生，调用方可用 `?` 直接把任意
    /// `std::io::Error` 提升为本变体。
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),

    /// zip 压缩包错误：CRC 校验失败、条目不存在、解压错误等。
    ///
    /// 由 `#[from] zip::result::ZipError` 派生，封装 `zip` crate 的全部错误。
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML 解析或序列化错误。
    ///
    /// 字符串内容应包含 **出错元素名 + 上下文路径**，例如
    /// `"slide layout1.xml parse: unexpected end of <p:sld> at line 42"`。
    #[error("xml error: {0}")]
    Xml(String),

    /// OPC 包结构错误。
    ///
    /// 典型场景：
    /// - 缺少必要的 part（如 `word/document.xml` 缺失）；
    /// - 关系链断裂（`r:id` 指向不存在的 target）；
    /// - `[Content_Types].xml` 中缺失 Override / Default。
    #[error("OPC error: {0}")]
    Opc(String),

    /// OOXML 模型错误。
    ///
    /// 典型场景：
    /// - 缺失必要字段（如 `slideLayout` 必须有 `cSld`）；
    /// - 命名空间不匹配；
    /// - 序列化时违反 OOXML 元素顺序约束（CT_* schema 严格顺序）。
    #[error("OOXML error: {0}")]
    Oxml(String),

    /// 元素未找到：按 Id / Name / Type 查找时未命中。
    #[error("not found: {0}")]
    NotFound(String),

    /// 索引越界：访问 `Slides` / `Shapes` 等集合时 idx >= len。
    #[error("index out of range: {0}")]
    IndexOutOfRange(usize),

    /// 不支持的功能（路线图中）。
    ///
    /// 携带 `&'static str` 描述功能名，便于编译期收集未实现项清单。
    #[error("not implemented: {0}")]
    NotImplemented(&'static str),

    /// 加密/解密错误。
    ///
    /// 典型场景：
    /// - 密码不匹配；
    /// - 加密算法不支持；
    /// - AES 密钥/IV 长度不正确；
    /// - 加密文件格式损坏。
    #[error("encryption error: {0}")]
    Encryption(String),

    /// .ppt（PowerPoint 97-2003 二进制格式）处理错误。
    ///
    /// 典型场景：
    /// - OLE2/CFB 容器结构损坏；
    /// - PPT record 树解析失败（record header / recLen 异常）；
    /// - 找不到必要 stream（PowerPoint Document / Current User）；
    /// - PersistDirectoryAtom 解析失败；
    /// - 水印注入失败（找不到 MainMaster / PPDrawing / SpgrContainer）；
    /// - persist 对象重排失败。
    ///
    /// 与 [`Error::Encryption`] 的区别：本变体专指 .ppt 二进制格式的
    /// 结构性错误；而 `Encryption` 侧重密钥/算法层面的失败。
    #[error("ppt97 error: {0}")]
    Ppt97(String),

    /// 其它错误。
    ///
    /// **仅** 用于临时过渡；正式错误请扩展 `enum Error` 的具体变体。
    #[error("{0}")]
    Other(String),
}

impl Error {
    /// 便捷构造：OPC 错误。
    ///
    /// 接受任何 `Into<String>` 的消息，避免调用方手动写 `Error::Opc(s.into())`。
    pub fn opc<S: Into<String>>(msg: S) -> Self {
        Error::Opc(msg.into())
    }

    /// 便捷构造：OOXML 错误。
    pub fn oxml<S: Into<String>>(msg: S) -> Self {
        Error::Oxml(msg.into())
    }

    /// 便捷构造：未实现。
    ///
    /// 配合 `unimplemented!` 风格的早期失败语义，用于在路线图功能
    /// 调用点快速失败。
    pub fn not_implemented(feature: &'static str) -> Self {
        Error::NotImplemented(feature)
    }

    /// 便捷构造：加密/解密错误。
    pub fn encryption<S: Into<String>>(msg: S) -> Self {
        Error::Encryption(msg.into())
    }

    /// 便捷构造：.ppt 97-2003 二进制格式处理错误。
    ///
    /// 用于 OLE2 容器、PPT record 树、persist 对象等结构性失败。
    pub fn ppt97<S: Into<String>>(msg: S) -> Self {
        Error::Ppt97(msg.into())
    }
}
