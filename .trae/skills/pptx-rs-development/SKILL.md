---
name: "pptx-rs-development"
description: "pptx-rs 本地开发指南：环境配置、构建命令、调试技巧、依赖升级、发布流程。Invoke when user asks how to build, run examples, debug, release, or set up dev environment."
---

# pptx-rs 开发指南

> 对应 [docs/DEVELOPMENT.md](../../../../docs/DEVELOPMENT.md)。

## 环境要求

| 工具 | 版本 | 说明 |
| --- | --- | --- |
| Rust | ≥ 1.75 | `rustup default stable` |
| Cargo | 随 Rust | — |
| Git | ≥ 2.30 | Windows 下推荐 `git for windows` |
| PowerPoint / WPS / LibreOffice | 任意 | 用于肉眼验证 .pptx 正确性 |
| 7-Zip | 任意 | 用来直接看 zip 内容（调试必备） |

可选：

| 工具 | 用途 |
| --- | --- |
| `cargo-watch` | `cargo install cargo-watch`，文件变更自动重跑 |
| `cargo-expand` | 宏展开调试 |
| `cargo-bloat` | 看二进制体积 |
| `valgrind`（Linux） | 内存检查（v0.1.0 不需要） |

## 5 分钟上手

```powershell
# 1) 克隆（如尚未克隆）
git clone <repo-url> pptx-rs && cd pptx-rs

# 2) 构建
cargo build

# 3) 跑示例（生成 hello.pptx）
cargo run --example hello_pptx

# 4) 用 PowerPoint / WPS 打开 hello.pptx 验证

# 5) 跑所有测试
cargo test --all

# 6) 生成文档
cargo doc --no-deps --open
```

## 常用命令

### 构建 / 检查

```powershell
cargo build                          # 默认 debug
cargo build --release                # release（LTO + 1 codegen-unit）
cargo check                          # 快速类型检查（不链接）
cargo clippy --all-targets -- -D warnings   # 静态检查（-D warnings 当错误）
cargo fmt --all                      # 格式化
cargo fmt --all -- --check           # 只检查不改
```

### 测试

```powershell
cargo test                                       # 全部测试
cargo test --doc                                 # 仅文档测试
cargo test units::                               # 按模块名过滤
cargo test -- --nocapture                        # 打印 println
cargo test -- --test-threads=1                   # 串行跑（资源竞争排查）
cargo test --all-features                        # 所有 feature
```

### 示例

```powershell
cargo run --example hello_pptx
cargo run --example protect_pptx
cargo run --example watermark_pptx
```

### 文档

```powershell
cargo doc --no-deps                # 生成
cargo doc --no-deps --open         # 生成 + 打开
RUSTDOCFLAGS="-D warnings" cargo doc --no-deps   # 把 warning 当错误
```

## 调试技巧

### 1. 看生成的 .pptx 内部结构

```powershell
# 用 7-Zip 直接打开 hello.pptx
# 或 PowerShell:
Expand-Archive hello.pptx -DestinationPath hello_extracted
Get-ChildItem -Recurse hello_extracted

# 看 [Content_Types].xml
Get-Content hello_extracted\[Content_Types].xml

# 看某个 slide
Get-Content hello_extracted\ppt\slides\slide1.xml
```

### 2. 与 python-pptx 的输出做 byte-diff

```powershell
# 用 python-pptx 重新生成同一份 .pptx
python -c "from pptx import Presentation; ..."
# 用 Beyond Compare / diff 比较两份 zip 的成员
```

### 3. 启用 trace 日志

`src/lib.rs` 顶部加：

```rust
#![cfg_attr(test, allow(unused_imports))]
```

或者在 `Cargo.toml` 加 feature：

```toml
[features]
trace = []
```

### 4. 排查 zip 大小 / 内容

```rust
let bytes = prs.to_bytes()?;
println!("pptx = {} bytes", bytes.len());
// 逐 part 打印
for part in pkg.iter_parts() {
    println!("{:>8}  {}  ({:>6} B)",
        part.partname.as_str(),
        part.content_type,
        part.len());
}
```

### 5. 排查 OOXML 元素顺序

最常见的问题是 PowerPoint/WPS 提示 "无法打开" / "格式不正确"。
- **第一嫌疑**：`<p:sldMaster>` 的 `clrMap` / `sldLayoutIdLst` / `txStyles` 子元素缺失。
- **第二嫌疑**：`<p:sldLayout>` 的 `type` 属性缺失。
- **第三嫌疑**：`<p:presentation>` 的 `defaultTextStyle` 段数不足。
- **第四嫌疑**：`<p:sp>` 的子元素顺序（如 `nvSpPr` → `spPr` → `txBody`）。

### 6. 排查 `quick-xml` 版本问题

`Cargo.toml` 锁在 `0.40`；`Reader::config_mut().trim_text(true)` 行为稳定。如果升级导致 `from_str` API 变化，集中在 `src/oxml/parser.rs::collect_inner_text`。

## 依赖升级

```powershell
# 1) 检查过时依赖
cargo outdated

# 2) 升级一个
cargo update -p quick-xml

# 3) 全升级（谨慎）
cargo update

# 4) 跑全部测试 + clippy
cargo test --all
cargo clippy --all-targets -- -D warnings
```

升级 `zip` / `quick-xml` 时重点关注：

| crate | 风险 | 关注点 |
| --- | --- | --- |
| `zip` | 中 | `ZipArchive::new` / `ZipWriter::start_file` 签名 |
| `quick-xml` | 高 | `Events` / `BytesStart` / `Reader` API 反复改 |
| `thiserror` | 低 | `#[from]` 行为稳定 |
| `base64` | 低 | `Engine::encode` / `STANDARD` 稳定 |
| `sha2` | 低 | `Digest::update` / `finalize` 稳定 |

## 性能分析

```powershell
# 1) 用 release profile
cargo build --release

# 2) 用 criterion（未来加入）做基准
# 暂未集成；现可通过 println! 计时
```

## 调试模式（`#[cfg(debug_assertions)]`）

允许在源码中加：

```rust
#[cfg(debug_assertions)]
eprintln!("[debug] shapes count = {}", self.shapes.len());
```

发布时自动剔除。

## IDE 集成

### VS Code

推荐扩展：

- `rust-analyzer`
- `Even Better TOML`
- `crates`（Cargo.toml 依赖提示）

`settings.json` 建议：

```json
{
  "rust-analyzer.cargo.features": "all",
  "[rust]": { "editor.formatOnSave": true }
}
```

### Trae IDE

- 启用本仓库 `.trae/rules/project_rules.md` 自动加载。
- 加载 `.trae/skills/` 下的所有 SKILL.md（按需触发）。

## 发布流程

1. 确认 `develop` 上 CI 绿、PR 全部合并。
2. `cargo update`（允许锁文件前进）。
3. 更新 `docs/CHANGELOG.md`。
4. 改 `Cargo.toml` 的 `version`。
5. `cargo test --all` + `cargo clippy --all-targets -- -D warnings`。
6. `cargo doc --no-deps`。
7. `git add -A && git commit -m "chore(release): v0.x.y"`。
8. `git tag -a v0.x.y -m "v0.x.y"`。
9. `cargo publish --dry-run`。
10. `cargo publish`。
11. 推 tag：`git push origin v0.x.y`。
12. 在 GitHub Release 写 changelog（可由 `docs/CHANGELOG.md` 截取）。

## 调试 `modifyVerifier` / 水印功能

两个 example 都基于"读 → 改 → 写"模式：
- `protect_pptx.rs`：在 `/ppt/presentation.xml` 注入 `<p:modifyVerifier>`。
- `watermark_pptx.rs`：在每个 slide 的 `</p:spTree>` 之前注入 `<p:sp>`。

跑完后 `_test_out/` 目录会生成输出文件，可用 PowerPoint 打开验证。

## 故障排查

| 现象 | 可能原因 | 解决 |
| --- | --- | --- |
| `cargo build` 报 `linker not found` | Windows 缺 MSVC | `rustup default stable-msvc` |
| 生成的 .pptx PowerPoint 拒绝 | OOXML 元素顺序错 | 用 7-Zip 打开，对照 [docs/OOXML_REFERENCE.md](../../../../docs/OOXML_REFERENCE.md) |
| 字号/位置不对 | EMU/Pt 转换错 | 重新查 `src/units.rs` 系数 |
| 中文乱码 | XML 未声明 `encoding="UTF-8"` | `XmlWriter::decl()` 强制写头 |
| 表格列宽不对 | `Table` 没指定 `Col.width` | 显式设置 |

## 代码组织建议（添加新代码时）

1. **先查 SKILL**：相关 SKILL.md 必读。
2. **先写注释**（遵循 [project_rules](../../rules/project_rules.md) §4）。
3. **再写实现**。
4. **再加测试**（[pptx-rs-testing](../pptx-rs-testing/SKILL.md)）。
5. **跑 `cargo fmt` + `cargo clippy` + `cargo test`**。
6. **更新 [docs/CHANGELOG.md](../../../../docs/CHANGELOG.md)**。
7. **提交**（遵循 Conventional Commits）。

## 常用链接

- [OOXML 标准](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/)
- [DrawingML 元素索引](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/)
- [PresentationML 元素索引](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-ctdse6b/)
- [python-pptx 文档](https://python-pptx.readthedocs.io/)
- [pypdf 文档](https://pypdf.readthedocs.io/)
- [zip crate 文档](https://docs.rs/zip/)
- [quick-xml 文档](https://docs.rs/quick-xml/)

## Solution Patterns

### Pattern 1: 开发新功能先写注释再写实现

```rust
// ✅ 先写文档注释
/// 向所有幻灯片添加文字水印。
///
/// 水印以半透明文字形式覆盖在每张幻灯片上方，
/// 支持自定义字号、颜色、旋转角度。
///
/// # 参数
/// - `text`：水印文字内容
/// - `font_size_pt`：字号（磅），默认 48pt
/// - `color`：颜色，默认灰色 C0C0C0
/// - `rotation_deg`：旋转角度，默认 -30°
///
/// # 错误
/// - `Error::Opc`：保存失败
pub fn add_watermark(&mut self, ...) -> Result<()> {
    // 然后写实现
}

// ❌ 先写实现后补注释（容易遗漏）
pub fn add_watermark(&mut self, ...) -> Result<()> {
    // 实现完了，注释？算了以后再补
}
```

**适用场景**：所有 `pub` 函数/方法。
**不适场景**：快速原型/实验代码。

### Pattern 2: 调试 OOXML 问题先拆 zip 再 diff

```bash
# ✅ 系统化排查
cargo run --example hello_pptx
Expand-Archive hello.pptx -DestinationPath debug_out
python -c "from pptx import Presentation; p=Presentation('hello.pptx'); p.save('ref.pptx')"
Expand-Archive ref.pptx -DestinationPath ref_out
Compare-Object (Get-Content debug_out\ppt\slides\slide1.xml) (Get-Content ref_out\ppt\slides\slide1.xml)

# ❌ 随意猜测
# "可能是编码问题" → 改了还是不行
```

**适用场景**：任何 PowerPoint/WPS 拒绝打开的情况。
**不适场景**：编译错误（应看编译器输出）。

## Workflow

### 开发新功能流程

```
1. 查阅 SKILL.md（相关领域知识）
2. 写文档注释（遵循 project_rules §4）
3. 写实现（三层顺序：OPC → OOXML → 高阶 API）
4. 写测试（单元 + 集成 + 端到端）
5. 跑 cargo fmt + cargo clippy + cargo test
6. 更新 CHANGELOG.md
7. 提交（Conventional Commits）
```

### 修复 Bug 流程

```
1. 复现问题（最小可复现代码）
2. 定位层（OPC / OOXML / 高阶 API）
3. 拆 zip 看 XML（OOXML 问题）
4. 修复 + 加回归测试
5. 验证（PowerPoint / WPS 打开）
6. 提交
```

## Review Checklist

- [ ] 文档注释完整（中文，含参数/返回值/错误）
- [ ] `cargo fmt --all -- --check` 通过
- [ ] `cargo clippy --all-targets -- -D warnings` 通过
- [ ] `cargo test --all` 通过
- [ ] `cargo doc --no-deps` 无 warning
- [ ] CHANGELOG 已更新
- [ ] 公共 API 在 README 反映
- [ ] 提交信息遵循 Conventional Commits

## Verification Commands

```bash
# 完整开发验证流程
cargo fmt --all
cargo clippy --all-targets -- -D warnings
cargo test --all
cargo doc --no-deps
cargo run --example hello_pptx
cargo run --example watermark_pptx
cargo run --example protect_pptx
cargo build --release
```

## Cross-References

- [pptx-rs-overview](../pptx-rs-overview/SKILL.md) — 项目总览
- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 架构详解
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南
- [pptx-rs-testing](../pptx-rs-testing/SKILL.md) — 测试规范
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
