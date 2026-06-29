//! OPC（Open Packaging Convention）容器层。
//!
//! `.pptx` 本质上是一个 **zip 包**，其内部由若干个 **Part** 组成，
//! Part 之间通过 **Relationship**（关系）相互链接。本模块提供与
//! python-pptx 的 `pptx.opc` 子包对标的最小化 Rust 实现。
//!
//! # 模块构成
//!
//! - [`OpcPackage`]：包加载/保存入口，对应一个完整的 `.pptx` 文件。
//! - [`Part`]：单个 part 的抽象（名称 + Content-Type + 字节内容）。
//! - [`PartName`]：part 的逻辑路径（如 `/ppt/slides/slide1.xml`）。
//! - [`ContentTypes`]：`[Content_Types].xml` 模型（defaults + overrides）。
//! - [`Relationships`]：单个 `.rels` 文件模型。
//! - [`RelType`]：常用关系类型枚举。
//!
//! # 在三层架构中的位置
//!
//! 本模块是 [`crate`] 三个层级中的 **最底层**（容器层）。它不依赖
//! `oxml`（OOXML 模型）或 `shape` / `presentation` 等高阶 API。
//! `oxml::writer` 输出的字节最终通过 `OpcPackage::put_part` 注入本层。
//!
//! # 与 OOXML 规范的对应
//!
//! - `Part` ↔ OOXML 规范中的"部件"（part）；
//! - `PartName` ↔ "部件名称"（`PartName` 路径，根用 `/`）；
//! - `Relationships` ↔ `*.rels` XML 文件；
//! - `ContentTypes` ↔ 根目录下的 `[Content_Types].xml`。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.opc.package.OpcPackage` ←→ [`OpcPackage`]；
//! - `pptx.opc.package.Part` ←→ [`Part`]；
//! - `pptx.opc.packuri.PackURI` ←→ [`PartName`]；
//! - `pptx.opc.constants.CONTENT_TYPE` ←→ [`ct`] 模块常量；
//! - `pptx.opc.parts.*.part_class_selector` 思路未移植，由调用方决定具体 part。
//!
//! # 示例
//!
//! ```no_run
//! use pptx_rs::opc::{OpcPackage, Part, PartName};
//!
//! let mut pkg = OpcPackage::new();
//! let xml = b"<?xml version=\"1.0\"?><root/>".to_vec();
//! let p = Part::new(PartName::from_unchecked("/custom/foo.xml"),
//!                   "application/xml", xml);
//! pkg.put_part(p);
//! pkg.save("out.pptx").unwrap();
//! ```

pub mod content_types;
pub mod package;
pub mod part;
pub mod rels;

// 公共 API 重新导出：让 `pptx_rs::opc::OpcPackage` 等可直接使用，
// 不必记忆模块嵌套路径。
pub use content_types::{ContentTypes, DefaultExt, Override};
pub use package::{ct, OpcPackage};
pub use part::{new_part_name, Part, PartName, PartNameError};
pub use rels::{RelType, Relationship, Relationships, Target};
