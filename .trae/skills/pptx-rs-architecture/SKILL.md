---
name: "pptx-rs-architecture"
description: "pptx-rs 三层架构（OPC / OOXML / 高阶 API）的模块职责、调用链、关键类型、扩展点详解。Invoke when user asks about module structure, where to add code, data flow, or how layers interact."
---

# pptx-rs 架构详解

> 对应 [docs/ARCHITECTURE.md](../../../../docs/ARCHITECTURE.md)（本文件是该文档的索引式速查）。

## 三层职责矩阵

| 层 | 路径 | 输入 | 输出 | 不可见约束 |
| --- | --- | --- | --- | --- |
| **OPC** | `src/opc/` | 文件 / `&[u8]` | `OpcPackage` | OOXML 规范 |
| **OOXML** | `src/oxml/` | `OpcPackage` / 程序构造 | XML 字符串 | OOXML schema |
| **高阶 API** | `src/presentation.rs` / `slide.rs` / `shape/` | 用户调用 | XML + zip | 业务友好性 |

## 模块依赖图

```
       presentation.rs (顶层)
            │
   ┌────────┼────────┬──────────────┐
   │        │        │              │
 slide   layouts  masters        shape/
   │        │        │              │
   └────────┴────────┴──────┬───────┘
                            │
                          oxml/
                            │
                          opc/
                            │
                         zip crate
```

**关键约束**：

- `opc` 不依赖 `oxml`（只搬运字节 + 关系）。
- `oxml` 不依赖 `shape`（模型层保持纯）。
- `shape` 不依赖 `presentation` / `slide`（高阶层依赖模型 + 借用 slide 的 spTree）。

## 模块一：OPC（Open Packaging Convention）

`src/opc/`

| 文件 | 关键类型 | 说明 |
| --- | --- | --- |
| `mod.rs` | （聚合） | 模块说明 + 公共再导出 |
| `package.rs` | `OpcPackage` | 包的"装/卸"：zip 读取、构造、序列化到 zip / 内存字节 |
| `part.rs` | `Part` / `PartName` | 单个 part 的抽象（partname + content-type + blob） |
| `rels.rs` | `Relationship` / `Relationships` / `RelType` | 一个 `.rels` 文件的模型 |
| `content_types.rs` | `ContentTypes` / `DefaultExt` / `Override` | `[Content_Types].xml` 模型 |

### 关键工具函数

- `OpcPackage::load(path)` → `Result<OpcPackage>`：从 .pptx 文件加载。
- `OpcPackage::save(path)` → `Result<()>`：保存为 .pptx。
- `OpcPackage::to_bytes()` → `Result<Vec<u8>>`：序列化到内存 zip。
- `rels_partname_for("/ppt/slides/slide1.xml")` → `"/ppt/slides/_rels/slide1.xml.rels"`。
- `derive_content_type(&content_types, partname)` → `String`：override 优先，fallback 到 defaults。

### 设计要点

- **`BTreeMap<String, Part>` 索引 partname**（有序，便于 O(N) 输出）。
- **关系与 part 分离维护**（`Relationships` 是独立结构，但作为 part 的 blob 存储）。
- **不缓存 ContentTypes 反向索引**，仅在 save 时按需 `derive_content_type`。

## 模块二：OOXML 模型层

`src/oxml/`

| 文件 | 关键类型 | 对应 OOXML 元素 |
| --- | --- | --- |
| `mod.rs` | （聚合） | 再导出 |
| `ns.rs` | `NS_*` 常量 | OPC/PresentationML/DrawingML 命名空间 |
| `simpletypes.rs` | `Alignment` / `PresetGeometry` / `Underline` / `Cap` / `TextDirection` / `MsoConnectorType` / `MsoFillType` | 枚举型属性 |
| `color.rs` | `Color` / `SchemeColor` / `PresetColor` / `ColorFormat` | `<a:srgbClr>` / `<a:schemeClr>` / `<a:prstClr>` + `<a:alpha>` |
| `sppr.rs` | `Transform` / `Fill` / `Line` / `Dash` / `ShapeProperties` | `<a:xfrm>` / `<a:solidFill>` / `<a:ln>` / `<p:spPr>` |
| `txbody.rs` | `TextBody` / `Paragraph` / `Run` / `RunProperties` / `BodyProperties` / `Inset` / `Indent` / `ParagraphProperties` | `<p:txBody>` / `<a:p>` / `<a:r>` / `<a:rPr>` / `<a:bodyPr>` |
| `shape.rs` | `Sp` / `Pic` / `Group` / `Connector` / `GraphicFrame` / `GroupChild` / `Graphic` | `<p:sp>` / `<p:pic>` / `<p:grpSp>` / `<p:cxnSp>` / `<p:graphicFrame>` |
| `table.rs` | `Table` / `Row` / `Col` / `Cell` | `<a:tbl>` / `<a:tr>` / `<a:gridCol>` / `<a:tc>` |
| `presentation.rs` | `PresentationRoot` / `SlideIdEntry` | `<p:presentation>` / `<p:sldId>` |
| `slide.rs` | `Sld` / `SlideShape` | `<p:sld>` / spTree 子元素（含 `name` 字段） |
| `slidelayout.rs` | `SldLayout` | `<p:sldLayout>` |
| `slidemaster.rs` | `SldMaster` | `<p:sldMaster>` |
| `theme.rs` | `default_theme_xml()` | `<a:theme>`（静态完整 Office Theme） |
| `parser.rs` | `escape` / `unescape` / `AttrMap` / `parse_*` / `collect_inner_text` | XML 读写底层 |
| `parse_sld.rs` | `parse_sld` / `parse_sp` / `parse_pic` / `parse_cxn` | SAX 解析 slide XML |
| `writer.rs` | `XmlWriter` | XML 写出器（链式 API） |

### 序列化约定

- **每个 `*` 模型都实现 `write_xml(&self, &mut XmlWriter)`**。
- **属性顺序遵循 OOXML 规范**（如 `<a:ln>` 必须先写属性再写子元素）。
- **复杂结构提前 `let xx_s = ...` 把 `&str` 借到 `w.close()` 之后**（避开 `Vec<&str>` 借用问题）。

### 命名空间声明

- **每个 part 根元素自带 `xmlns:a` / `xmlns:p` / `xmlns:r` 声明**。
- 不在子元素重复声明（除 `<a:blipFill xmlns:r=...>` 这种"局部分支"）。

## 模块三：高阶 API

| 路径 | 关键类型 | 一句话 |
| --- | --- | --- |
| `src/presentation.rs` | `Presentation` / `MediaEntry` / `CoreProperties` | 顶层容器 + 元数据 + 水印 + 加密占位 |
| `src/slide.rs` | `Slide` / `Slides` / `Shapes` / `ShapesMut` / `SlideBackground` | 幻灯片及其形状视图 |
| `src/slide_layouts.rs` | `SlideLayout` / `SlideLayouts` / `SlideLayoutRef` | 版式（简版） |
| `src/slide_masters.rs` | `SlideMaster` / `SlideMasters` / `SlideMasterRef` | 母版（简版） |
| `src/shape/mod.rs` | `ShapeKind` / `PlaceholderShape` / `wrap()` | 形状枚举 + oxml→高阶的桥 |
| `src/shape/base.rs` | `trait Shape` | 所有形状的抽象接口 |
| `src/shape/autoshape.rs` | `AutoShape` | 自选图形（矩形/椭圆/箭头/...） |
| `src/shape/textbox.rs` | `TextBox` | 文本框（AutoShape + txBody） |
| `src/shape/picture.rs` | `Picture` | 图片 |
| `src/shape/connector.rs` | `Connector` | 连接器（直线/折线，含 `recompute_xfrm`） |
| `src/shape/group.rs` | `Group` / `GroupChild` | 组合 |
| `src/shape/table.rs` | `TableShape` | 表格（高阶） |
| `src/shape/freeform.rs` | `FreeformBuilder` / `Freeform` / `Point` | 自由形 |

### 新增类型详解

#### CoreProperties（文档元数据）

对标 pypdf `DocumentInformation`，集中管理 11 个元数据字段：

```rust
#[derive(Debug, Clone, Default)]
pub struct CoreProperties {
    pub title: Option<String>,
    pub creator: Option<String>,
    pub subject: Option<String>,
    pub last_modified_by: Option<String>,
    pub application: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
    pub keywords: Option<String>,
    pub category: Option<String>,
    pub description: Option<String>,
    pub revision: Option<String>,
}
```

**存储位置**：`/docProps/core.xml`（Dublin Core 命名空间）。
**写入时机**：`Presentation::to_opc_package` 中序列化。
**读取时机**：`Presentation::from_opc` 中解析（v0.1.0 简版）。

#### SlideBackground（幻灯片背景）

0.1.0 占位实现，返回 `MsoFillType::Inherit`：

```rust
pub struct SlideBackground<'a> {
    slide: &'a Slide,
}
impl<'a> SlideBackground<'a> {
    pub fn fill_type(&self) -> MsoFillType { MsoFillType::Inherit }
}
```

**路线图**：v0.2+ 支持纯色/渐变/图片背景。

#### Alpha 透明度链路

水印的半透明效果通过以下链路实现：

```
Presentation::add_watermark(text, font_size, color, rotation)
    │
    ├─► 创建 Sp（水印形状）
    ├─► 设置 RunProperties.alpha = Some(40000)  // 40% 不透明
    ├─► 设置 Transform.rot = rotation * 60_000
    │
    ▼
RunProperties::write_xml
    │
    ├─► Color::write_solid_fill_with_alpha(w, color, alpha)
    │       │
    │       ├─► <a:solidFill>
    │       │     <a:srgbClr val="RRGGBB">
    │       │       <a:alpha val="40000"/>  ← 关键：0-100000
    │       │     </a:srgbClr>
    │       │   </a:solidFill>
    │
    ▼
Slide::write_xml → spTree 中包含水印 sp
```

### 关键设计：`AutoShape` 是"全能基类"

- `TextBox` / `Freeform` 内部都包了一个 `AutoShape`。
- `Shape` trait 的所有方法在每个具体类型上重复实现（不可避免，Rust 无继承）。
- **`Shape` trait 借用的是 oxml 字段**（`sp.properties.xfrm.off_x`），**不复制**。

### 形状添加到 slide 的流程

```rust
// 用户调用
let mut tb = slide.shapes_mut().add_textbox_with_text(l, t, w, h, "hi")?;

// 内部步骤
// 1) ShapesMut 创建 TextBox（包一个 AutoShape，sp.geom = Rectangle）
// 2) 设置 sp.properties.xfrm.off_x / off_y / ext_cx / ext_cy
// 3) sp.id = slide.next_shape_id()  // 共享计数器
// 4) tb.set_text("hi")             // 清空 paragraphs + push 一段一行
// 5) clone tb.shape.sp 推入 slide.inner.shapes
// 6) 返回 tb
```

**为什么 clone？** 0.1.0 采取"高阶对象与 oxml 模型双份拥有 + 各自维护"策略；后续可在 `Rc<RefCell<Sp>>` 上统一。

## 调用链：保存 .pptx

```
Presentation::save(path)
    │
    ▼
Presentation::to_opc_package(&self)
    │
    ├─► 构造 OpcPackage
    ├─► 写入 /_rels/.rels（根关系）
    ├─► 写入 /docProps/core.xml（CoreProperties 序列化）
    ├─► 写入 /docProps/app.xml
    ├─► 写入 /ppt/theme/theme1.xml（default_theme_xml）
    ├─► 写入 /ppt/slideMasters/slideMaster1.xml + .rels
    ├─► 写入 /ppt/slideLayouts/slideLayout1.xml + .rels
    ├─► 写入 /ppt/presProps.xml / viewProps.xml / tableStyles.xml
    │
    ├─► for each slide:
    │     ├─► sld.set_layout_rid("rId1")
    │     ├─► 构造 /ppt/slides/_rels/slideN.xml.rels（layout + media + notes）
    │     ├─► 写入 /ppt/slides/slideN.xml
    │     └─► 在 pres_rels 上加 <Relationship rId="..." Type="slide" Target="slides/slideN.xml">
    │
    ├─► 写入 /ppt/presentation.xml
    ├─► 写入 /ppt/_rels/presentation.xml.rels
    └─► 写入 /ppt/media/* (图片)
    │
    ▼
OpcPackage::save(path) / to_bytes()
    │
    ▼
zip::ZipWriter + xml content
```

## 调用链：打开 .pptx

```
Presentation::open(path)
    │
    ▼
Presentation::load(path)
    │
    ▼
OpcPackage::load(path)
    │   解析 [Content_Types].xml → ContentTypes
    │   遍历 zip 读所有 part
    │   关系文件以 RELATIONSHIPS Content-Type 标记
    ▼
Presentation::from_opc(pkg)
    │   ├─► 解析 core.xml → CoreProperties
    │   ├─► 解析 presentation.xml → slide 列表
    │   ├─► 解析每个 slide → Sld + shapes
    │   ├─► 解析 notes slide（如有）
    │   ⚠️ 当前仅构造空壳 + 1 个 master + 1 个 layout
    │   完整解析路线图
    ▼
Presentation
```

## 调用链：添加水印

```
Presentation::add_watermark(text, font_size, color, rotation)
    │
    ├─► for each slide:
    │     ├─► 创建 Sp（水印形状）
    │     │     ├─► name = "pptx-rs-watermark"
    │     │     ├─► geom = PresetGeometry::Rectangle
    │     │     ├─► xfrm: 全屏覆盖（0,0 → slide_width, slide_height）
    │     │     ├─► fill = None（透明填充）
    │     │     └─► line = None（无边框）
    │     │
    │     ├─► 创建 TextBody
    │     │     ├─► bodyPr: anchor="ctr"
    │     │     ├─► Paragraph: align="ctr"
    │     │     └─► Run: text + RunProperties
    │     │           ├─► size = font_size (默认 48pt)
    │     │           ├─► color = color (默认灰色 C0C0C0)
    │     │           ├─► alpha = Some(40000)  // 40% 不透明
    │     │           └─► rotation = rotation * 60_000
    │     │
    │     └─► 推入 slide.inner.shapes
    │
    ▼
保存时水印 sp 随 slide 一起序列化
```

## 共享状态

| 状态 | 类型 | 所有者 | 备注 |
| --- | --- | --- | --- |
| shape id 计数器 | `Rc<Cell<u32>>` | `Presentation` 创建 → 共享给每个 `Slide` → 每个 `add_*` 自增 | `Slide::next_shape_id()` |
| 媒体（图片） | `Vec<MediaEntry>` | `Presentation` | 保存时遍历写入 zip |
| Master / Layout | `SlideMasters` / `SlideLayouts` | `Presentation` | 0.1.0 强制 1 master + 1 layout |
| 元数据 | `CoreProperties` | `Presentation` | 11 字段，写入 `/docProps/core.xml` |
| Notes Slide 信息 | `notes_partname` / `notes_slide_rel_target` | `Slide` | 读写时维护 |

## Solution Patterns

### Pattern 1: 新增跨 slide 操作放在 Slides 集合上

```rust
// ✅ 跨 slide 操作在 Slides 上
impl Slides {
    pub fn insert_slide(&mut self, id_counter: Rc<Cell<u32>>, index: usize) -> Result<&mut Slide> { ... }
    pub fn clone_slide(&mut self, src_idx: usize, insert_at: usize) -> Result<&mut Slide> { ... }
    pub fn append_slides_from(&mut self, other: &Slides) { ... }
    fn reindex(&mut self) { ... }  // 自动重排 sld_id / rid / partname
}

// ❌ 放在 Presentation 上导致职责混乱
impl Presentation {
    fn insert_slide(&mut self, index: usize) -> Result<&mut Slide> { ... }  // 应委托给 Slides
}
```

**适用场景**：insert/clone/append/remove/reindex 等集合操作。
**不适场景**：单个 slide 的属性操作（如 `extract_text`）。

### Pattern 2: Connector 的 xfrm 必须与 begin/end 同步

```rust
// ✅ set_begin/set_end 自动调用 recompute_xfrm
impl Connector {
    pub fn set_begin(&mut self, x: Emu, y: Emu) {
        self.cxn.begin = Some((x, y));
        self.recompute_xfrm();  // 自动同步
    }

    pub fn recompute_xfrm(&mut self) {
        // 根据 begin/end 计算 off_x/off_y/ext_cx/ext_cy
    }
}

// ❌ 手动维护 xfrm（容易忘记同步）
connector.set_begin(Emu(100), Emu(200));
// 忘了更新 xfrm → PowerPoint 显示位置错误
```

**适用场景**：任何有 begin/end 端点的连接器。
**不适场景**：静态形状（xfrm 不依赖端点）。

## Workflow

### 选择代码放置层

```
1. 是否涉及 zip / part / 关系？
   → 是：OPC 层（src/opc/）
   → 否：进入第 2 步

2. 是否涉及 XML 元素 / 属性 / 命名空间？
   → 是：OOXML 层（src/oxml/）
   → 否：进入第 3 步

3. 是否是用户直接调用的便捷方法？
   → 是：高阶 API 层（src/presentation.rs / slide.rs / shape/）
   → 否：重新评估需求

4. 跨层调用规则：
   → 高阶 API 可调用 OOXML + OPC
   → OOXML 可调用 OPC（仅 PartName 等）
   → OPC 不调用 OOXML（纯容器）
```

## 扩展点

| 想加什么 | 看哪里 |
| --- | --- |
| 新形状（preset） | `oxml/simpletypes.rs::PresetGeometry` + 写出器 |
| 新 Run 属性 | `oxml/txbody.rs::RunProperties` + `write_xml` |
| 新颜色支持 | `oxml/color.rs` |
| 新图表类型 | 在 `shape/mod.rs::ShapeKind` 加 variant + `oxml/shape.rs::Graphic` |
| 新 OPC 关系类型 | `opc/rels.rs::RelType` + uri 映射 |
| 新 Content-Type | `opc/package.rs::ct::*` |
| 新元数据字段 | `presentation.rs::CoreProperties` + `to_opc_package` 序列化 |
| 新水印样式 | `presentation.rs::add_watermark` + `oxml/txbody.rs::RunProperties` |
| 新背景类型 | `slide.rs::SlideBackground` + `oxml/slide.rs::Sld` |

详见 [pptx-rs-extending](../pptx-rs-extending/SKILL.md)。

## 反模式 / 陷阱

| 陷阱 | 说明 | 应对 |
| --- | --- | --- |
| 形状 ID 冲突 | 同一 slide 内 id 必须唯一 | 共享 `Rc<Cell<u32>>` |
| Master 缺失 | PowerPoint 拒绝打开 | `Presentation::ensure_default_master_and_layout` 兜底 |
| 关系文件位置 | `/ppt/slides/_rels/slide1.xml.rels`，不是 `/ppt/_rels/...` | `rels_partname_for()` |
| `<p:defaultTextStyle>` | PowerPoint 强校验 | 必含 9 段 `lvlXpPr`（静态常量） |
| Master 的 `<p:txStyles>` | PowerPoint 强校验 | 必含 titleStyle/bodyStyle/otherStyle |
| `<p:modifyVerifier>` 位置 | WPS 对位置敏感 | 必须在 `<p:extLst>` 之前 |
| 旋转单位 | OOXML 用 1/60000 度 | `set_rotation(deg) = deg * 60_000` |
| 字号单位 | OOXML 用百分之一磅 | `Pt(24.0)` → `sz=2400` |
| Connector xfrm 不同步 | begin/end 变化后 xfrm 未更新 | `recompute_xfrm()` |
| Alpha 范围 | 0-100000（不是 0-100） | `alpha = 40000` 表示 40% 不透明 |
| Slides reindex | insert/clone/remove 后 sld_id 不连续 | `Slides::reindex()` 自动重排 |

## Review Checklist

- [ ] 新增类型在正确的层（OPC / OOXML / 高阶 API）
- [ ] 跨层依赖方向正确（高阶 → OOXML → OPC，不反向）
- [ ] 新增 Sp/Pic 等的 `write_xml` 子元素顺序正确
- [ ] 新增共享状态用 `Rc<Cell<>>` / `Rc<RefCell<>>`
- [ ] Connector 的 begin/end 变化后调用了 `recompute_xfrm`
- [ ] Slides 集合操作后调用了 `reindex`
- [ ] CoreProperties 字段在 `to_opc_package` 中正确序列化
- [ ] 水印 Sp 的 alpha 值在 0-100000 范围内

## Cross-References

- [pptx-rs-overview](../pptx-rs-overview/SKILL.md) — 项目总览
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南
- [pptx-rs-ooxml](../pptx-rs-ooxml/SKILL.md) — OOXML 速查
- [pptx-rs-testing](../pptx-rs-testing/SKILL.md) — 测试规范
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
