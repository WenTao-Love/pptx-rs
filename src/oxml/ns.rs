//! OOXML 命名空间常量。
//!
//! 这些常量是 OPC / DrawingML / PresentationML / VML 等的标准 URI。
//! 全部以 `pub const &str` 形式暴露，便于在序列化时直接使用而无需硬编码。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.oxml.ns.PRESENTATIONML_MAIN` ←→ [`NS_PRESENTATION_MAIN`]；
//! - `pptx.oxml.ns.DRAWINGML_MAIN` ←→ [`NS_DRAWING_MAIN`]；
//! - `pptx.opc.constants.RELATIONSHIPS` ←→ [`NS_RELATIONSHIPS`]；
//! - 等等。
//!
//! # 前缀约定（OOXML 规范推荐）
//!
//! | 常量 | 规范推荐前缀 | 用途 |
//! | ---- | ------------ | ---- |
//! | [`NS_PRESENTATION_MAIN`] | `p:` | PresentationML 主命名空间 |
//! | [`NS_DRAWING_MAIN`] | `a:` | DrawingML 主命名空间 |
//! | [`NS_DRAWING_RELS`] | `r:` | 关系 |
//! | [`NS_CP`] | `cp:` | 核心属性 |
//! | [`NS_DC`] | `dc:` | Dublin Core |

/// 关系文件命名空间。
pub const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Content-Types 命名空间。
pub const NS_CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

/// 主 PresentationML 命名空间（`p:` 前缀）。
///
/// 包含 `<p:presentation>` / `<p:sld>` / `<p:sp>` / `<p:txBody>` / `<p:sldMaster>` 等。
pub const NS_PRESENTATION_MAIN: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

/// DrawingML 主命名空间（`a:` 前缀）。
///
/// 包含 `<a:xfrm>` / `<a:prstGeom>` / `<a:solidFill>` / `<a:r>` / `<a:t>` 等。
pub const NS_DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// DrawingML PowerPoint 扩展（`p:`，与 PresentationML 同名，区分靠元素名）。
///
/// 包含 `<p:transition>` / `<p:timing>` 等 PowerPoint 特有的动画元素。
pub const NS_DRAWING_PPT: &str = "http://schemas.openxmlformats.org/drawingml/2006/ppt";

/// 关系（drawing 关系）。
pub const NS_DRAWING_RELS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// WordprocessingML（docProps 用到）。
///
/// core-properties 命名空间，含 `<dc:title>` / `<dc:creator>` / `<cp:lastModifiedBy>` 等。
pub const NS_CP: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

/// 扩展属性。
///
/// 包含 `<Application>` / `<Company>` / `<AppVersion>` 等"应用属性"。
pub const NS_EXT_PROPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

/// Dublin Core 命名空间。
pub const NS_DC: &str = "http://purl.org/dc/elements/1.1/";

/// dcterms 命名空间。
pub const NS_DCTERMS: &str = "http://purl.org/dc/terms/";

/// dcmitype 命名空间。
pub const NS_DCMITYPE: &str = "http://purl.org/dc/dcmitype/";

/// xsi 命名空间。
pub const NS_XSI: &str = "http://www.w3.org/2001/XMLSchema-instance";

/// DrawingML Diagram（SmartArt）命名空间（`dgm:` 前缀）。
///
/// 包含 `<dgm:relIds>` / `<dgm:dataModel>` / `<dgm:layoutDef>` 等 SmartArt 元素。
pub const NS_DIAGRAM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
