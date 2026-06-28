//! pptx-rs —— Rust 实现的 PowerPoint `.pptx` 读写库，对标 [python-pptx](https://github.com/scanny/python-pptx)。
//!
//! # 项目定位
//!
//! `pptx-rs` 旨在以 Rust 的强类型 + 零 unsafe + 零异步的"现代化"姿态，完整复刻
//! `python-pptx` 提供的 PowerPoint 文档读写能力，并在性能与可维护性上做到更优。
//! 当前处于 **0.2.0** 阶段。
//!
//! # 顶层模块结构
//!
//! - [`Presentation`] / [`Slide`] / [`Slides`]：高阶面向用户 API。
//! - [`shape`]：形状（`AutoShape` / `Picture` / `Group` / `Connector` / `Table` / `TextBox`）。
//! - [`opc`]：OPC（Open Packaging Convention）容器层 —— zip、Part、关系。
//! - [`oxml`]：OOXML 模型层 —— PresentationML / DrawingML 的强类型 XML 模型。
//! - [`units`] / [`RGBColor`]：EMU / Pt / Inches / 颜色。
//! - [`crypto`]：.pptx 文件加密（ECMA-376 Agile Encryption：AES-256-CBC + SHA-512）。
//! - [`ppt97`]：.ppt 97-2003 二进制格式支持（水印注入 + RC4 CryptoAPI 加密）。
//! - [`Error`] / [`Result`]：错误与 Result 别名。
//!
//! # 三层架构（自下而上）
//!
//! ```text
//!   高阶 API 层    Presentation / Slide / Shapes
//!        ↑ 借用
//!   OOXML 模型层   oxml/{sp, pic, slide, master, layout, ...}
//!        ↑ 序列化
//!   OPC 容器层    opc/{Part, Relationships, ContentTypes}
//!        ↑ 读写
//!   zip crate
//! ```
//!
//! 下层绝不依赖上层。详见 [`docs::ARCHITECTURE`] 与
//! [`.trae/skills/pptx-rs-architecture`](../.trae/skills/pptx-rs-architecture/SKILL.md)。
//!
//! # 最小示例
//!
//! ```no_run
//! use pptx::Presentation;
//!
//! let mut prs = Presentation::new().unwrap();
//! let counter = prs.id_counter();
//! let slide = prs.slides_mut().add_slide(counter).unwrap();
//! let mut tb = slide.shapes_mut().add_textbox(
//!     pptx::Inches(1.0), pptx::Inches(1.0),
//!     pptx::Inches(8.0), pptx::Inches(1.0),
//! ).unwrap();
//! tb.set_text("hello");
//! prs.save("hello.pptx").unwrap();
//! ```
//!
//! # 单位与坐标系
//!
//! 全部几何计算均以 **EMU**（English Metric Unit，`i64`）为内部单位，理由：
//!
//! - OOXML 中所有几何属性（`off` / `ext` / 行高 / 列宽 / 边框）都是整数 EMU；
//! - 整数运算保证**无浮点精度漂移**；
//! - 与其他 Office 工具（python-pptx / Open XML SDK）互操作零成本。
//!
//! 外部 API 通过 [`Inches`] / [`Pt`] / [`Cm`] 三种 NewType 包装 + [`units::EmuExt`] 扩展 trait
//! 提供便捷转换；类型系统会阻止"英寸和磅混用"的潜在 bug。
//!
//! # 错误处理
//!
//! 库内**禁止** `panic!` / `unwrap()`。所有公开 API 返回 [`Result<T>`]，错误统一归入
//! [`enum Error`](Error) 的 10 个变体（`Io` / `Zip` / `Xml` / `Opc` / `Oxml` / `NotFound` /
//! `IndexOutOfRange` / `NotImplemented` / `Encryption` / `Other`）。错误消息遵循**小写 + 句末无标点**。
//!
//! # 公开 API 稳定性
//!
//! `0.1.x` 期间：
//!
//! - `pub` 方法签名（方法名 / 参数 / 返回类型）—— **不承诺稳定**，允许小调整；
//! - `pub` 字段 —— **不承诺稳定**，调整走 `CHANGELOG` 标注 `internal`；
//! - 破坏性变更 —— 需先 `#[deprecated(note = "...")]` 一段时间。
//!
//! 详见 [`docs::CHANGELOG`] 与 [`.trae/rules/project_rules.md`](../.trae/rules/project_rules.md)。
//!
//! # 路线图
//!
//! 详见 [`README.md`](../README.md) 的"路线图"段、[`docs::CHANGELOG`] 与
//! [`.trae/skills/pptx-rs-overview/SKILL.md`](../.trae/skills/pptx-rs-overview/SKILL.md)。
//!
//! [`docs::ARCHITECTURE`]: ../docs/ARCHITECTURE.md
//! [`docs::CHANGELOG`]: ../docs/CHANGELOG.md

#![deny(rust_2018_idioms)]
#![warn(missing_debug_implementations)]
//! # 编译期开关说明
//!
//! - `deny(rust_2018_idioms)`：禁止 2015 风格的 idiom（`#[macro_use]`、路径写法等）。
//! - `warn(missing_debug_implementations)`：所有 `pub` 类型应实现 `Debug`（如未实现会警告）。
//!   当前 v0.1.0 中 `Shapes` / `ShapesMut` 暂未实现，未来需补 `#[derive(Debug)]`。

/// 库统一的 `Result` 类型别名。所有公共 API（除特别声明）均返回该类型。
///
/// # 示例
///
/// ```no_run
/// use pptx::Result;
///
/// fn read_something() -> Result<String> { Ok(String::new()) }
/// ```
pub use crate::error::{Error, Result};

// 重新导出常用单位类型，保持与 python-pptx 的 `from pptx import Pt, Inches` 风格相近。
pub use crate::units::{Cm, Emu, EmuExt, EmuPoint, Inches, Pt, RGBColor};

// 重新导出常用枚举（OOXML simple types），对标 `pptx.enum.*`。
// 这些枚举在 oxml 与 shape 两侧都被用到，提到 crate 根以减少路径深度。
pub use crate::oxml::simpletypes::{
    Alignment, Cap, MsoAnchor, MsoAutoSize, MsoColorType, MsoConnectorType, MsoFillType,
    MsoLineDashStyle, MsoShapeType, MsoThemeColorIndex, PpAlign, PpPlaceholderType, PresetGeometry,
    TextDirection, TextWrapping, Underline,
};

// 重新导出颜色、变换、填充、边框等"最常用基础类型"。
pub use crate::oxml::color::{Color, ColorFormat, ColorRole, PresetColor, SchemeColor};
pub use crate::oxml::sppr::{
    Dash, EffectList, Fill, FillFormat, GlowEffect, Line, LineFormat, ReflectionEffect,
    ShadowEffect, ShapeProperties, SoftEdgeEffect, Transform,
};
// 重新导出文本体相关：TextFrame / Paragraph / Run / Font / ParagraphFormat。
pub use crate::oxml::txbody::{
    BodyProperties, Font, Indent, Inset, Paragraph, ParagraphFormat, ParagraphProperties, Run,
    RunProperties, TextBody, TextFrame,
};
// 重新导出主题样式 + 扩展列表（p:style / p:extLst）。
pub use crate::oxml::shape::{ExtensionEntry, ExtensionList, ShapeLocks, ShapeStyle, StyleRef};
// 形状锁定类型枚举（TODO-027 高阶 API）。
pub use crate::oxml::shape::LockType;
// 三维效果相关类型（TODO-050 scene3d / sp3d / backdrop）。
pub use crate::oxml::sppr::{
    Backdrop, Bevel, Camera, CameraPreset, LightRig, LightRigDirection, LightRigType,
    MaterialPreset, Point3d, Rotation3d, Scene3d, Sp3d,
};
// 图表相关类型（TODO-004 基础图表支持）。
pub use crate::oxml::chart::{Chart, ChartCategory, ChartData, ChartSeries, ChartType};
// 章节分组相关类型（TODO-039 Section 分组）。
pub use crate::oxml::section::{Section, SectionList};
// OLE 对象相关类型（TODO-043 OLE 嵌入）。
pub use crate::oxml::ole::{OleObject, OLE_GRAPHIC_DATA_URI};
// 音视频媒体相关类型（TODO-033 音视频嵌入）。
pub use crate::oxml::shape::MediaKind;
pub use crate::presentation::{AudioEntry, DiagramEntry, VideoEntry};
// SmartArt 最小保留类型（TODO-037 SmartArt 识别 + round-trip）。
pub use crate::oxml::shape::SmartArtRef;
// SmartArt 高阶句柄（TODO-037 创建 API）。
pub use crate::shape::smartartshape::SmartArtShape;
// SmartArt 结构化解析类型（TODO-037 SmartArt 数据模型）。
pub use crate::oxml::diagram::{
    ColorsDef, DataModel, DataModelConnection, DataModelPoint, LayoutCategory, LayoutDef,
    QuickStyleDef, StyleLabel,
};

// 高阶顶层 API 重新导出。
pub use crate::crypto::ModifyProtection;
pub use crate::presentation::Presentation;
pub use crate::presentation::{CoreProperties, CustomProperties, CustomPropertyValue};
pub use crate::slide::{Shapes, ShapesMut, Slide, SlideBackground, SlideId, SlideRef, Slides};
pub use crate::slide_layouts::{Placeholder, SlideLayout, SlideLayoutRef, SlideLayouts};
pub use crate::slide_masters::{SlideMaster, SlideMasterRef, SlideMasters};
// 备注母版相关类型（TODO-045 NotesMaster 访问）。
pub use crate::notes_masters::{NotesMasterRef, NotesMasters};
pub use crate::oxml::notesmaster::NotesMaster;
// 评论相关类型（TODO-036）。
pub use crate::oxml::comments::{Comment, CommentAuthor, CommentAuthorList, CommentList};
// 幻灯片过渡相关类型（TODO-020 高阶 API）。
pub use crate::oxml::slide::{
    MorphOption, SplitOrientation, Transition, TransitionDirection, TransitionSpeed, TransitionType,
};

// 按模块组织源代码。模块顺序遵循"底层→高层"：
// error → units → opc → oxml → crypto → ppt97 → 顶层。
// ppt97（.ppt 97-2003 二进制格式支持）独立于 .pptx 主线，作为可选模块放在 crypto 之后。
pub mod crypto;
pub mod error;
pub mod notes_masters;
pub mod opc;
pub mod oxml;
pub mod ppt97;
pub mod presentation;
pub mod shape;
pub mod slide;
pub mod slide_layouts;
pub mod slide_masters;
pub mod units;
