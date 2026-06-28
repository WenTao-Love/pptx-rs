---
name: "pptx-rs-overview"
description: "pptx-rs 项目总览，对标 python-pptx + pypdf 的 Rust .pptx 读写库。Invoke when user asks what pptx-rs is, how it relates to python-pptx or pypdf, where to start, or needs a 30-second project summary."
---

# pptx-rs 项目总览

> **版本**：v0.1.0（MVP）
> **目标**：用 Rust 1.75+ 实现一个对标 [python-pptx](https://github.com/scanny/python-pptx) 的 PowerPoint `.pptx` 读写库，同时吸收 [pypdf](https://github.com/py-pdf/pypdf) 的 API 设计模式。
> **关键事实**：零 `unsafe`、纯同步、单一 `Cargo.toml` 依赖极简（zip / quick-xml / thiserror）。

## 一句话定位

> 把 python-pptx 的"易用 API + 严格 OOXML 模型"翻译成 Rust 的"高阶包装 + 强类型 XML + Rc/RefCell 内部状态"，并借鉴 pypdf 的 Reader/Writer 分离、页面操作、加密、水印、元数据等设计模式。

## 三层架构（30 秒版）

```
┌─────────────────────────────────────────────────────────────┐
│  高阶 API 层（src/presentation.rs、src/slide.rs、src/shape/）│
│  ── Presentation / Slide / TextBox / AutoShape / Picture    │
│  ── CoreProperties / SlideBackground / extract_text         │
│  ── add_watermark / encrypt / decrypt                       │
└─────────────────────────────────────────────────────────────┘
                          │ 包装
┌─────────────────────────────────────────────────────────────┐
│  OOXML 模型层（src/oxml/）                                  │
│  ── Sp / Pic / Group / GraphicFrame / CxnSp / Sld / ...    │
│  ── SldLayout / PresentationRoot / Theme / TextBody         │
│  ── RunProperties.alpha / Color.write_solid_fill_with_alpha │
└─────────────────────────────────────────────────────────────┘
                          │ 读写
┌─────────────────────────────────────────────────────────────┐
│  OPC 容器层（src/opc/）                                     │
│  ── OpcPackage / Part / PartName / ContentTypes / RelType   │
└─────────────────────────────────────────────────────────────┘
```

详细分层职责、调用链、扩展点见 [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md)。

## 5 个关键文件

| 路径 | 一句话 |
| --- | --- |
| `src/lib.rs` | 顶层导出（`pub use`）+ 模块声明 |
| `src/presentation.rs` | `Presentation::new` / `open` / `save` + `CoreProperties` + `add_watermark` + `encrypt`/`decrypt` |
| `src/slide.rs` | `Slides::add_slide` / `insert_slide` / `clone_slide` + `Shapes` + `SlideBackground` + `extract_text` |
| `src/oxml/mod.rs` | OOXML 模型的统一再导出 |
| `src/opc/package.rs` | zip ↔ `OpcPackage` 的桥接 |

## 公共 API 与 python-pptx 对照

| python-pptx | pptx-rs | 备注 |
| --- | --- | --- |
| `Presentation()` | `Presentation::new()` | 0.1.0 不带参数 |
| `Presentation(path)` | `Presentation::open(path)` | |
| `prs.save(path)` | `prs.save(path)` | |
| `prs.slides` | `prs.slides()` / `prs.slides_mut()` | 显式 `mut` |
| `prs.slide_width` | `prs.slide_width()` → `Emu` | |
| `prs.core_properties` | `prs.core_properties()` / `prs.core_properties_mut()` | 对标 pypdf `DocumentInformation` |
| `slide.shapes` | `slide.shapes()` / `shapes_mut()` | |
| `shapes.add_textbox(l, t, w, h)` | `shapes_mut().add_textbox_with_text(l, t, w, h, "text")` | 一步式更安全 |
| `shape.text_frame.text = "x"` | `tb.set_text("x")` | 链式 |
| `run.font.size = Pt(24)` | `run.properties.size = Some(Pt(24.0))` | |
| `run.font.bold = True` | `run.properties.bold = true` | |
| `shapes.title` | `shapes.title()` | 返回 `Option<&Shape>` |
| `shapes.placeholders` | `shapes.placeholders()` | 返回占位符迭代器 |
| `slide.slide_layout` | `slide.layout()` | 0.1.0 简版 |

## pypdf API 对标

| pypdf | pptx-rs | 状态 |
| --- | --- | --- |
| `PdfReader(path)` | `Presentation::open(path)` | ✅ 已实现 |
| `PdfWriter()` | `Presentation::new()` | ✅ 已实现 |
| `writer.add_page(page)` | `slides.add_slide(counter)` | ✅ 已实现 |
| `writer.insert_page(index, page)` | `slides.insert_slide(counter, index)` | ✅ 已实现 |
| `reader.pages[i].extract_text()` | `slide.extract_text()` | ✅ 已实现 |
| `writer.clone_page_from(reader, i)` | `slides.clone_slide(src_idx, insert_at)` | ✅ 已实现 |
| `writer.append_pages_from_reader(reader)` | `slides.append_slides_from(other)` | ✅ 已实现 |
| `reader.metadata` | `prs.core_properties()` | ✅ 已实现（`CoreProperties`） |
| `writer.add_blank_page()` | `slides.add_slide(counter)` | ✅ 已实现 |
| `page.merge_page(watermark)` | `prs.add_watermark(...)` | ✅ 已实现 |
| `writer.encrypt(password)` | `prs.encrypt(password, read_only)` | 🔄 占位（返回 `NotImplemented`） |
| `reader.is_encrypted` | `prs.is_encrypted()` | 🔄 占位（始终返回 `false`） |
| `reader.decrypt(password)` | `prs.decrypt(password)` | 🔄 占位（返回 `NotImplemented`） |
| `reader.page_count` | `prs.num_slides()` | ✅ 已实现 |

## 已实现（v0.1.0）

- OPC 容器：zip 读写、`[Content_Types].xml`、Part 关系链
- Presentation / Slide / SlideLayout / SlideMaster XML 模型
- 形状树：AutoShape、Picture、Group、Connector、TextBox、TableShape、Freeform
- TextFrame / Paragraph / Run / 字体（粗/斜/下划线/字号/颜色/字距/基线/Alpha 透明度）
- 常用 DrawingML 几何（`<a:prstGeom>` 50+ preset）
- 表格（按行按列）
- 文档保护（`<p:modifyVerifier>` 注入）
- 水印（`Presentation::add_watermark`，支持 Alpha 透明度）
- 文档元数据（`CoreProperties`：title/creator/subject/keywords 等 11 字段）
- 幻灯片操作（insert/clone/append/reindex）
- 文本提取（`Slide::extract_text`）
- Notes Slide 读写
- 占位符查询（`Shapes::title` / `placeholders` / `placeholder(idx)`）

## 路线图（v0.2+）

- 完整 Master / Layout / Theme 主题切换
- Chart（基础类型）
- SmartArt
- 完整读取路径（`from_opc` 当前仅建空壳）
- 加密/解密实现（ECMA-376 Agile Encryption）
- 性能优化（流式写）
- 自定义几何（custGeom）

## Solution Patterns

### Pattern 1: 新功能应同时对标 python-pptx 和 pypdf

```rust
// 问题：添加新 API 时如何确定设计方向？

// ✅ 同时参考两个库的设计
// python-pptx: prs.slides.add_slide(layout)
// pypdf:       writer.add_blank_page()
// → pptx-rs:   slides.add_slide(counter)  // 取 python-pptx 的"指定 layout"
//                                       // 取 pypdf 的"简洁参数"

// ❌ 只参考一个库
fn add_slide() -> Slide { ... }  // 缺少 layout 选择
```

**适用场景**：添加任何新的公共 API。
**不适场景**：OOXML 底层细节（纯 OOXML 规范，pypdf 无对应）。

### Pattern 2: 元数据用结构体而非松散字段

```rust
// ✅ 集中管理
pub struct CoreProperties {
    pub title: Option<String>,
    pub creator: Option<String>,
    // ... 11 字段
}
impl Presentation {
    pub fn core_properties(&self) -> &CoreProperties { ... }
    pub fn core_properties_mut(&mut self) -> &mut CoreProperties { ... }
}

// ❌ 松散字段
impl Presentation {
    pub fn title(&self) -> Option<&str> { ... }
    pub fn set_title(&mut self, t: &str) { ... }
    pub fn creator(&self) -> Option<&str> { ... }
    // ... 22 个方法
}
```

**适用场景**：一组相关属性（元数据、样式、保护选项）。
**不适场景**：单个独立属性（如 `slide_width`）。

## Workflow

### 选择 API 设计方向

```
1. 是否在 python-pptx 中有对应？
   → 是：优先对齐 python-pptx 命名和签名
   → 否：进入第 2 步

2. 是否在 pypdf 中有类似模式？
   → 是：参考 pypdf 的 API 风格（Reader/Writer 分离、page 操作）
   → 否：按 Rust 惯用设计

3. 是否涉及 OOXML 底层？
   → 是：走 oxml 层 + write_xml
   → 否：在高阶 API 层直接实现

4. 是否需要跨 slide 操作？
   → 是：放在 Slides 集合上（如 insert_slide / clone_slide）
   → 否：放在单个 Slide 上（如 extract_text）
```

## Quick Reference（5 秒速查）

```rust
// ✅ pptx-rs 惯用
let mut prs = Presentation::new()?;
prs.core_properties_mut().title = Some("Demo".into());
prs.add_watermark("DRAFT", Some(48.0), None, Some(-30))?;
let slide = prs.slides_mut().add_slide(prs.id_counter())?;
let text = slide.extract_text();

// ❌ 反模式
let prs = Presentation::new().unwrap();  // 库代码不用 unwrap
prs.save("out.pptx").expect("save");     // 错误信息不够具体
```

## 关键术语

| 术语 | 含义 |
| --- | --- |
| **EMU** | English Metric Unit。1 inch = 914 400 EMU，1 pt = 12 700 EMU |
| **OOXML** | Office Open XML；.pptx 是其中一种 |
| **OPC** | Open Packaging Convention；zip + Content-Types + .rels |
| **DrawingML** | 图形语言，元素前缀 `a:` |
| **PresentationML** | 演示语言，元素前缀 `p:` |
| **prstGeom** | 预设几何（rect/ellipse/arrow/...） |
| **rId** | 关系 ID（`rId1`、`rId2`、...） |
| **Layout / Master** | 母版 / 版式：定义占位符与样式继承 |
| **CoreProperties** | 文档元数据（对标 pypdf `DocumentInformation`） |
| **Alpha** | DrawingML 透明度（0-100000，100000=不透明） |

## 反模式（请勿）

- ❌ 引入 `unsafe`（除 `zip` 内部）
- ❌ 引入异步运行时（tokio / async-std）
- ❌ 写 `format!("<a:sp>{}</a:sp>", ...)`（必须用 `XmlWriter`）
- ❌ 直接解析/构造 OOXML 字符串（请走 `oxml` 层）
- ❌ 0.1.0 内对 `pub` 字段做隐式破坏性变更
- ❌ 只参考 python-pptx 不参考 pypdf（两个库互补）

## Cross-References

- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 三层架构详解
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南
- [pptx-rs-ooxml](../pptx-rs-ooxml/SKILL.md) — OOXML 速查
- [pptx-rs-testing](../pptx-rs-testing/SKILL.md) — 测试规范
- [pptx-rs-development](../pptx-rs-development/SKILL.md) — 开发指南
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
- [python-pptx 文档](https://python-pptx.readthedocs.io/)
- [pypdf 文档](https://pypdf.readthedocs.io/)
