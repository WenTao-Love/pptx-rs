//! OOXML XML 模型层。
//!
//! 本模块是 [`crate`] 三个层级中的 **中间层**——OOXML 模型层。
//! 它负责把 OPC 容器中的"字节内容"反序列化为强类型 Rust 结构，
//! 同时把内存中的模型序列化为符合 OOXML 规范的 XML。
//!
//! # 子模块结构
//!
//! - [`ns`]：所有 OOXML / DrawingML / PresentationML 命名空间常量。
//! - [`simpletypes`]：枚举型简单类型（`algn` / `prstGeom` / `u` / ...）。
//! - [`parser`] / [`writer`]：XML 解析与序列化基础工具。
//! - [`color`]：`a:srgbClr` / `a:schemeClr` / `a:prstClr` 统一表达。
//! - [`sppr`]：`<p:spPr>`（变换 / 几何 / 填充 / 边框）。
//! - [`txbody`]：`<p:txBody>`（段落 / Run / 字体属性）。
//! - [`shape`]：`<p:sp>` / `<p:pic>` / `<p:grpSp>` / `<p:cxnSp>` / `<p:graphicFrame>`。
//! - [`table`]：`<a:tbl>`（嵌入在 `<p:graphicFrame>` 内）。
//! - [`presentation`]：`<p:presentation>` 根元素。
//! - [`slide`]：`<p:sld>` 单个幻灯片。
//! - [`slidemaster`] / [`slidelayout`]：母版与版式（极简版）。
//! - [`theme`]：标准 Office 主题 XML。
//!
//! # 在三层架构中的位置
//!
//! ```text
//!   高阶 API 层   presentation / slide / shape
//!        ↑ 借用
//!   OOXML 模型层  ← 你在这里
//!        ↑ 序列化
//!   OPC 容器层    opc/{OpcPackage, Part, ...}
//! ```
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.presentation.Presentation` ←→ [`PresentationRoot`]；
//! - `pptx.oxml.slide.Slide` ←→ [`Sld`]；
//! - `pptx.oxml.ns` ←→ [`ns`]；
//! - `pptx.oxml.text.*` ←→ [`txbody`]；
//! - `pptx.oxml.shape.*` ←→ [`shape`]。
//!
//! # 序列化策略
//!
//! 内部模型采用 `write_xml(&mut XmlWriter)` 显式序列化风格，**不**使用
//! `serde::Serialize` 派生。理由：OOXML 元素有严格顺序约束、命名空间混杂，
//! 显式写更可控且 byte-diff 友好。

pub mod chart;
pub mod color;
pub mod comments;
pub mod diagram;
pub mod notesmaster;
pub mod ns;
pub mod ole;
pub mod parse_sld;
pub mod parser;
pub mod presentation;
pub mod section;
pub mod shape;
pub mod simpletypes;
pub mod slide;
pub mod slidelayout;
pub mod slidemaster;
pub mod sppr;
pub mod table;
pub mod theme;
pub mod txbody;
pub mod writer;

// 公共 API 重新导出：让 `pptx_rs::oxml::Color` 等可直接使用。
pub use chart::{Chart, ChartCategory, ChartData, ChartSeries, ChartType};
pub use color::{Color, ColorFormat, ColorRole, PresetColor, SchemeColor};
pub use comments::{Comment, CommentAuthor, CommentAuthorList, CommentList};
pub use diagram::{
    ColorsDef, DataModel, DataModelConnection, DataModelPoint, LayoutCategory, LayoutDef,
    QuickStyleDef, StyleLabel,
};
pub use notesmaster::NotesMaster;
pub use ole::{OleObject, OLE_GRAPHIC_DATA_URI};
pub use presentation::{PresentationRoot, SldMasterIdEntry, SlideIdEntry};
pub use section::{Section, SectionList, SECTION_EXT_URI};
pub use shape::{
    Connector, ExtensionEntry, ExtensionList, Graphic, GraphicFrame, Group, GroupChild, MediaKind,
    Pic, ShapeLocks, ShapeStyle, SmartArtRef, Sp, StyleRef,
};
pub use simpletypes::{
    Alignment, Cap, MsoAnchor, MsoAutoSize, MsoColorType, MsoConnectorType, MsoFillType,
    MsoLineDashStyle, MsoShapeType, MsoThemeColorIndex, PpAlign, PpPlaceholderType, PresetGeometry,
    TextDirection, TextWrapping, Underline,
};
pub use slide::{notes_xml, Sld, SlideShape};
pub use slidelayout::SldLayout;
pub use slidemaster::SldMaster;
pub use sppr::{
    AdjustmentValue, ArrowHead, ArrowSize, ArrowType, Backdrop, Bevel, BlipFillMode, Camera,
    CameraPreset, CustomGeometry, Dash, EffectList, Fill, FillFormat, GeomRect, Geometry,
    GlowEffect, GradientFill, GradientPath, GradientStop, GradientType, LightRig,
    LightRigDirection, LightRigType, Line, LineFormat, LineJoin, MaterialPreset, Path, PathSegment,
    PatternFill, Point3d, ReflectionEffect, Rotation3d, Scene3d, ShadowEffect, ShapeProperties,
    SoftEdgeEffect, Sp3d, Transform,
};
pub use table::{Cell, CellBorder, Col, Row, Table, TableLook, TableStyle, VerticalAnchor};
pub use theme::{default_theme_xml, ColorScheme, FontScheme, FormatScheme, Theme, ThemeColor};
pub use txbody::{
    BodyProperties, BulletStyle, Caps, Font, Hyperlink, Indent, Inset, Paragraph, ParagraphFormat,
    ParagraphProperties, Run, RunProperties, TextBody, TextFrame,
};
