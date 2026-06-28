---
name: "pptx-rs-extending"
description: "pptx-rs 扩展指南：如何添加新形状、新属性、新 OPC 关系、新 Content-Type、水印、元数据、加密、新示例。Invoke when user asks how to add a new shape type, new attribute, new relationship type, watermark, metadata, encryption, or extend the library."
---

# pptx-rs 扩展指南

> 本指南按"加什么"组织。每节给出"先改什么 → 怎么改 → 怎么测"。
> 融入 huiali/rust-skills 的 Solution Patterns 方法论。

## Workflow：扩展前决策

```
1. 扩展属于哪一层？
   → OPC 容器：见 §添加新 OPC 关系类型 / Content-Type
   → OOXML 模型：见 §添加新自选形状 / Run 属性 / 颜色
   → 高阶 API：见 §添加新高阶形状 / 水印 / 元数据 / 加密

2. 是否需要跨层改动？
   → 是：按 OPC → OOXML → 高阶 API 顺序改
   → 否：只改对应层

3. 是否影响公共 API？
   → 是：更新 README + CHANGELOG + 加 #[non_exhaustive]
   → 否：仅内部改动
```

## 添加新自选形状（PresetGeometry）

**场景**：想加一个 `PresetGeometry::Star12`（目前已有 5/6/10 角星）。

**步骤**：

1. 在 [`src/oxml/simpletypes.rs`](../../../../src/oxml/simpletypes.rs) 加枚举变体：

```rust
pub enum PresetGeometry {
    // ...
    Star12,  // ← 新增
    // ...
}

impl PresetGeometry {
    pub fn as_str(self) -> &'static str {
        use PresetGeometry::*;
        match self {
            // ...
            Star12 => "star12",  // ← 新增映射
            // ...
        }
    }
}

impl FromStr for PresetGeometry {
    // ...
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Ok(match s {
            // ...
            "star12" => Star12,  // ← 新增反向映射
            // ...
        })
    }
}
```

2. 在 [`src/slide.rs`](../../../../src/slide.rs) 的 `add_shape` 中**无需改动**（已接受任意 `PresetGeometry`）。
3. 在 `examples/hello_pptx.rs` 加一个用例：

```rust
let mut star = slide.shapes_mut().add_shape(
    pptx::oxml::simpletypes::PresetGeometry::Star12,
    Inches(2.0), Inches(3.0),
    Inches(3.0), Inches(3.0),
)?;
```

4. 在 [`src/oxml/simpletypes.rs`](../../../../src/oxml/simpletypes.rs) 底部测试加 `assert_eq!(PresetGeometry::Star12.as_str(), "star12");`。
5. 跑 `cargo test` + `cargo run --example hello_pptx` 验证 PowerPoint 能正常打开。

## 添加新 Run 属性

**场景**：想支持 `RunProperties::shadow`（投影）。

1. 在 [`src/oxml/txbody.rs`](../../../../src/oxml/txbody.rs) 的 `RunProperties` 加字段：

```rust
pub struct RunProperties {
    // ...
    /// 投影。
    pub shadow: Option<Shadow>,
}
```

2. 定义 `Shadow`（可借用 `oxml::sppr` 的现有结构或新建）。
3. 在 `RunProperties::write_xml` 输出顺序正确位置添加序列化：

```rust
// 严格 OOXML 顺序：在 <a:ln> 之后、<a:solidFill> 之前/之后，按规范
if let Some(sh) = &self.shadow {
    sh.write_xml(w);
}
```

4. 加测试：写入 Run → 检查 XML 包含 `<a:effectLst>` 或对应标签。
5. **重要**：如果新属性影响 `<a:rPr>` 的子元素顺序，更新 [`pptx-rs-ooxml`](../pptx-rs-ooxml/SKILL.md) 文档。

## 添加新颜色

**场景**：想支持 `Color::System(SystemColor::WindowText)`（系统色）。

1. 在 [`src/oxml/color.rs`](../../../../src/oxml/color.rs) 加枚举：

```rust
pub enum Color {
    // ...
    System(SystemColor),
}

pub enum SystemColor {
    WindowText,
    Window,
    // ...
}
```

2. 实现 `Color::write_solid_fill` / `write_xml` 写出 `<a:sysClr val="windowText"/>`。
3. 文档化属性顺序约束（见 OOXML §17.18）。

## 添加新 OPC 关系类型

**场景**：想支持 `RelType::Chart`。

1. 在 [`src/opc/rels.rs`](../../../../src/oxml/rels.rs) 加枚举变体：

```rust
pub enum RelType {
    // ...
    Chart,  // 已有；若需要新类型如 Comments:
    Comments,
    Other(&'static str),
}

impl RelType {
    pub fn uri(self) -> &'static str {
        match self {
            // ...
            RelType::Chart =>
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            RelType::Comments =>
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            // ...
        }
    }
}
```

2. `Relationships::from_xml` 中加反向映射。
3. 在 `examples/` 加使用示例。
4. 文档更新：若新类型影响 Content-Types，参考 [pptx-rs-ooxml](../pptx-rs-ooxml/SKILL.md)。

## 添加新 Content-Type

**场景**：想支持 `.svg` 媒体。

1. 在 [`src/opc/package.rs`](../../../../src/opc/package.rs) 加常量：

```rust
pub mod ct {
    // ...
    pub const IMAGE_SVG: &str = "image/svg+xml";
}
```

2. 在 [`src/opc/content_types.rs`](../../../../src/opc/content_types.rs) 的 `new_default` 添加：

```rust
ct.defaults.push(DefaultExt::new("svg", "image/svg+xml"));
```

3. 在 [`src/shape/picture.rs`](../../../../src/shape/picture.rs) 的 `content_type_for` 添加：

```rust
".svg" => "image/svg+xml",
```

## 添加新高阶形状

**场景**：想加 `CalloutShape`（带引线标注）。

1. **定义 oxml 模型**（如已有 `<p:sp>` 满足，跳过）：

```rust
// src/oxml/shape.rs
pub struct Sp {
    // ...
    /// 是否为引线标注
    pub is_callout: bool,  // 新增
}
```

2. **添加 ShapeKind 变体**：

```rust
// src/shape/mod.rs
pub enum ShapeKind {
    // ...
    Callout(CalloutShape),
}
```

3. **实现 Shape trait**：

```rust
// src/shape/callout.rs（新文件）
use crate::shape::base::Shape;

#[derive(Clone, Debug, Default)]
pub struct CalloutShape {
    pub(crate) sp: crate::oxml::shape::Sp,
}

impl CalloutShape { /* ... */ }

impl Shape for CalloutShape {
    // ...
}
```

4. **在 [`src/shape/mod.rs`](../../../../src/shape/mod.rs) 的 `wrap()` 函数中加映射**。
5. **在 [`src/slide.rs`](../../../../src/slide.rs) 的 `ShapesMut` 加 `add_callout` 方法**。
6. **在 [`src/lib.rs`](../../../../src/lib.rs) 加 `pub use`**。
7. **加测试 + 例子**。

## 添加水印样式

**场景**：想支持图片水印或对角线水印。

### 文字水印变体

1. 在 [`src/presentation.rs`](../../../../src/presentation.rs) 的 `add_watermark` 增加参数：

```rust
pub fn add_watermark(
    &mut self,
    text: &str,
    font_size_pt: Option<f64>,
    color: Option<RGBColor>,
    rotation_deg: Option<i32>,
    alpha: Option<i32>,  // ← 新增：自定义透明度
    font_name: Option<&str>,  // ← 新增：自定义字体
) -> Result<()> { ... }
```

2. 更新 `RunProperties` 构造，使用传入的 `alpha` 和 `font_name`。
3. 更新 `examples/watermark_pptx.rs`。
4. 加测试：生成水印 → 检查 XML 含 `<a:alpha val="..."/>`。

### 图片水印（路线图）

1. 在 `add_watermark` 中创建 `Pic` 而非 `Sp`。
2. 设置 `blipFill` 的 `<a:alpha>` 实现透明度。
3. 需要在 `MediaEntry` 中添加水印图片。

## 添加元数据字段

**场景**：想支持 `CoreProperties` 的新字段（如 `content_status`）。

1. 在 [`src/presentation.rs`](../../../../src/presentation.rs) 的 `CoreProperties` 加字段：

```rust
pub struct CoreProperties {
    // ... 现有 11 字段
    /// 内容状态（如 "Draft" / "Final"）。
    pub content_status: Option<String>,  // ← 新增
}
```

2. 在 `to_opc_package` 的 core.xml 序列化中添加：

```rust
if let Some(cs) = &self.core_properties.content_status {
    // 写入 <cp:contentStatus>cs</cp:contentStatus>
}
```

3. 在 `from_opc` 的 core.xml 解析中添加（如已实现）。
4. 加测试：设置字段 → 保存 → 重新打开 → 验证。

## 添加加密/解密

**场景**：想实现真正的 `encrypt` / `decrypt`（v0.2+ 路线图）。

### 加密实现步骤

1. **在 [`src/error.rs`](../../../../src/error.rs) 加错误变体**：

```rust
/// 密码不匹配
#[error("password mismatch")]
PasswordMismatch,

/// 加密算法不支持
#[error("unsupported encryption algorithm: {0}")]
UnsupportedEncryption(String),
```

2. **在 [`Cargo.toml`](../../../../Cargo.toml) 加依赖**（如需要）：

```toml
[dependencies]
# ECMA-376 Agile Encryption 需要
aes = "0.8"       # AES 加密
cbc = "0.1"       # CBC 模式
hmac = "0.12"     # HMAC 验证
```

3. **实现 `encrypt`**：

```rust
pub fn encrypt(&mut self, password: &str, read_only: bool) -> Result<()> {
    // 1) 生成随机 salt（16 bytes）
    // 2) SHA-512(password + salt) spinCount=100000 次
    // 3) 构造 <p:modifyVerifier>
    // 4) 注入到 presentation.xml
    // 5) 如 read_only，设置 <p:modification access="readOnly">
    Ok(())
}
```

4. **实现 `decrypt`**：

```rust
pub fn decrypt(&mut self, password: &str) -> Result<bool> {
    // 1) 读取 <p:modifyVerifier>
    // 2) 用相同算法计算 hash
    // 3) 比对 hash
    // 4) 返回是否匹配
    Ok(true)
}
```

5. **更新 `is_encrypted`**：

```rust
pub fn is_encrypted(&self) -> bool {
    self.encrypted  // 新增字段
}
```

6. **加测试 + 更新 examples/protect_pptx.rs**。

### WPS 兼容性注意

- `<p:modifyVerifier>` 必须在 `<p:extLst>` 之前。
- `cryptProviderType` 必须为 `rsaAES`。
- `cryptAlgorithmSid` = 14（SHA-512）。
- `spinCount` = 100000。

## 添加新示例

**场景**：想加一个 `examples/dual_slide.rs`。

1. 写最小可运行代码：

```rust
//! 双 slide 演示。
//!
//! 用法：`cargo run --example dual_slide`

use pptx::{Inches, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut prs = Presentation::new()?;
    let counter = prs.id_counter();
    let s1 = prs.slides_mut().add_slide(counter)?;
    s1.shapes_mut()
        .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(8.0), Inches(1.0), "Slide 1")?;
    let counter = prs.id_counter();
    let s2 = prs.slides_mut().add_slide(counter)?;
    s2.shapes_mut()
        .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(8.0), Inches(1.0), "Slide 2")?;
    prs.save("dual.pptx")?;
    Ok(())
}
```

2. 在 [`Cargo.toml`](../../../../Cargo.toml) 中**无需配置**（`examples/` 目录下 `.rs` 文件自动被 cargo 发现）。
3. 在 [`README.md`](../../../../README.md) 的"快速开始"附近加引用。
4. 在 [`docs/CHANGELOG.md`](../../../../docs/CHANGELOG.md) 加 changelog。

## 添加新单元测试

**场景**：在 `src/oxml/simpletypes.rs` 加 `PresetGeometry` 解析测试。

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_known_geometry() {
        assert_eq!("star5".parse::<PresetGeometry>().unwrap(), PresetGeometry::Star5);
    }

    #[test]
    fn parse_unknown_geometry_falls_back_to_other() {
        assert_eq!("xxx".parse::<PresetGeometry>().unwrap(), PresetGeometry::Other);
    }
}
```

## 添加新集成测试

**场景**：在 `tests/` 加 `presentation_save.rs`。

```rust
// tests/presentation_save.rs
use pptx::*;

#[test]
fn end_to_end_save_load() {
    let mut prs = Presentation::new().unwrap();
    let counter = prs.id_counter();
    let _ = prs.slides_mut().add_slide(counter).unwrap();
    let bytes = prs.to_bytes().unwrap();
    let prs2 = Presentation::load_bytes(&bytes).unwrap();
    assert!(prs2.slide_width().value() > 0);
}
```

## 添加新公共 API

**步骤**：

1. 在 `pub` 类型上加 `#[must_use]`（若返回 `Result` / `&mut`）。
2. 写完整 doc 注释（遵循 [project_rules](../../rules/project_rules.md) §4）。
3. 加测试。
4. 更新 [`README.md`](../../../../README.md) 与 [docs/CHANGELOG.md](../../../../docs/CHANGELOG.md)。
5. 若 0.1.0 期间标记弃用，加 `#[deprecated(note = "...")]` + 给出迁移示例。

## 添加新依赖

**场景**：需要 `regex` 做属性验证。

1. 在 [`Cargo.toml`](../../../../Cargo.toml) 加：

```toml
[dependencies]
regex = "1.10"
```

2. 在 [`docs/DEVELOPMENT.md`](../../../../docs/DEVELOPMENT.md) 的"依赖策略"表登记。
3. **评估**：是否真的需要？能复用现有代码吗？能推迟到 v0.2 吗？
4. **避免**：
   - 重型依赖（机器学习栈、GUI 框架）
   - 异步运行时（与"零异步"原则冲突）
   - 重复造轮子（如 `serde-xml-rs` 与 `quick-xml` 重复）

## 添加新错误变体

**场景**：想区分"密码错误" vs "OPC 损坏"。

1. 在 [`src/error.rs`](../../../../src/error.rs) 加：

```rust
pub enum Error {
    // ...
    /// 密码不匹配
    #[error("password mismatch")]
    PasswordMismatch,
}
```

2. 加便捷构造（如 `Error::password_mismatch()`）。
3. 在 `examples/protect_pptx.rs` 中使用。
4. 文档化"何时抛"。

## 升级公共 API 时的兼容策略

```rust
// 0.1.0 → 0.2.0 重命名方法
#[deprecated(note = "use new_method instead")]
pub fn old_method(&self) {}

pub fn new_method(&self) {}
```

并在 [`docs/CHANGELOG.md`](../../../../docs/CHANGELOG.md) 标注：

```markdown
## [0.2.0] - 2026-xx-xx

### Removed
- `Presentation::old_method` → 改用 `Presentation::new_method`（#[deprecated] 自 0.1.5）
```

## Solution Patterns

### Pattern 1: 扩展按三层顺序改

```rust
// ✅ OPC → OOXML → 高阶 API 顺序
// 1) 添加 RelType::Chart（OPC 层）
// 2) 添加 Chart XML 模型（OOXML 层）
// 3) 添加 ChartShape 高阶 API（高阶 API 层）

// ❌ 反向改
// 1) 先写 ChartShape（缺底层支持）
// 2) 再补 XML 模型
// 3) 最后补关系类型
// → 中间每步都编译不过
```

**适用场景**：任何跨层扩展。
**不适场景**：单层内改动（如只加 PresetGeometry 枚举变体）。

### Pattern 2: 新增属性先加 Option 字段再实现 write_xml

```rust
// ✅ 渐进式
pub struct RunProperties {
    // ... 现有字段
    /// 投影效果。
    pub shadow: Option<Shadow>,  // Option = 向后兼容
}

impl RunProperties {
    fn write_xml(&self, w: &mut XmlWriter) {
        // ... 现有写出
        if let Some(sh) = &self.shadow {
            sh.write_xml(w);  // 有值才写
        }
    }
}

// ❌ 一步到位但破坏兼容
pub struct RunProperties {
    pub shadow: Shadow,  // 必填 → 破坏所有现有代码
}
```

**适用场景**：任何新增可选属性。
**不适场景**：OOXML 必填元素（如 `<a:bodyPr>`）。

## Review Checklist

- [ ] 代码注释完整（[`project_rules`](../../rules/project_rules.md) §4）
- [ ] 单元测试覆盖关键路径（[`pptx-rs-testing`](../pptx-rs-testing/SKILL.md)）
- [ ] `cargo fmt` + `cargo clippy` + `cargo test` 全过
- [ ] 公共 API 在 [`README.md`](../../../../README.md) / [`docs/`](../../../../docs/) 反映
- [ ] CHANGELOG 写入
- [ ] 提交信息遵循 Conventional Commits
- [ ] 新增 Option 字段而非必填字段（向后兼容）
- [ ] OOXML 子元素顺序正确（参考 [`pptx-rs-ooxml`](../pptx-rs-ooxml/SKILL.md)）
- [ ] 跨层改动按 OPC → OOXML → 高阶 API 顺序
- [ ] 水印/加密改动在 WPS 和 PowerPoint 上都验证过

## Verification Commands

```bash
# 编译检查
cargo check

# 格式化
cargo fmt --all

# Clippy
cargo clippy --all-targets -- -D warnings

# 全部测试
cargo test --all

# 运行示例验证
cargo run --example hello_pptx
cargo run --example watermark_pptx
cargo run --example protect_pptx

# 文档
cargo doc --no-deps

# Release 构建
cargo build --release
```

## Cross-References

- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 架构详解（扩展点一览）
- [pptx-rs-ooxml](../pptx-rs-ooxml/SKILL.md) — OOXML 速查（元素顺序）
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南
- [pptx-rs-testing](../pptx-rs-testing/SKILL.md) — 测试规范
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
