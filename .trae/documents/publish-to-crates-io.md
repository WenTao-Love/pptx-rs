# 发布 pptx-rs 到 crates.io 的准备计划

## Context（背景）

用户希望将 `pptx-rs` 项目发布到 crates.io。经过 3 个 Explore agent 的彻底审查，发现项目对齐 python-pptx 约 90%，但存在多个发布阻断项：

- **硬阻断**：`[lib] name = "pptx"` 与 crates.io 已存在的 `pptx` crate 冲突；`cargo clippy` 有 119 个 error；README 快速开始示例无法编译；`repository` URL 是占位符
- **严重**：版本号不一致（Cargo.toml=0.2.0 / lib.rs 混用 0.1.x / project_rules=0.1.0）；缺少 LICENSE 文件；缺 `rust-version` 字段；公开类型缺 `Debug` impl
- **代码质量**：9 处 `unreachable!()`、3 处 String 拼 XML、6 处 `Error::Other` 可分类

用户决策：
1. lib 名重命名为 `pptx_rs`（删除 `[lib] name` 覆盖，用默认值）
2. 仓库 URL 用 `https://github.com/WenTao-Love/pptx-rs.git`
3. 范围：修发布阻断项（不含 round-trip 保真修复），发布为 0.3.0

## 实施步骤

### 步骤 1：Crate 重命名 `pptx` → `pptx_rs`

**目标**：删除 `[lib] name = "pptx"` 覆盖，让导入名回退为默认 `pptx_rs`（包名 `pptx-rs` 连字符转下划线）。

**改动**：
- `Cargo.toml`：删除 `[lib] name = "pptx"` 两行（保留 `path = "src/lib.rs"` 或直接删整个 `[lib]` 段，用默认）
- 全局替换 `use pptx::` → `use pptx_rs::`（src/、tests/、examples/、benches/ 所有 .rs 文件）
- 全局替换路径前缀 `pptx::` → `pptx_rs::`（如 `pptx::Result` → `pptx_rs::Result`，排除注释/字符串中的 `pptx-rs`）
- `README.md` 代码示例中的 `use pptx::` → `use pptx_rs::`
- `src/lib.rs` doctest 中的 `use pptx::` → `use pptx_rs::`

**注意**：只替换 `use pptx::` 和 `pptx::` 路径前缀，不替换项目名 `pptx-rs` 本身（注释、文档中的项目名保持不变）。

### 步骤 2：修复 Cargo.toml

**文件**：`d:\xcode\xdemo\realtime-screen-ocr-rust\pptx-rs\Cargo.toml`

- 删除 `[lib]` 段（或改为 `name = "pptx_rs"`，但删除更干净）
- `version`：`0.2.0` → `0.3.0`
- 删除第 10 行注释 `# 发布到 crates.io 时需替换为真实仓库地址`
- `repository`：`https://github.com/pptx-rs/pptx-rs` → `https://github.com/WenTao-Love/pptx-rs`
- `homepage`：同上
- 添加 `rust-version = "1.75"`（在 `edition` 行后）

### 步骤 3：修复 clippy 119 个 error

**文件**：多个 src/ 文件（参考 `_clippy_full.txt`）

按 error 类别批量修复（均为机械性改动）：

| 类别 | 数量 | 修复方式 |
|------|------|---------|
| `field_reassign_with_default` | ~40 | 测试代码中 `let mut x = T::default(); x.f = v;` → `let x = T { f: v, ..Default::default() };` |
| `derivable_impls` | ~8 | 手写 `impl Default` → `#[derive(Default)]` + `#[default]` 标注 |
| `should_implement_trait` | ~6 | `pub fn from_str` → 实现 `FromStr` trait 或改名 `parse` |
| `explicit_auto_deref` | ~8 | `&*t` → `&t` |
| `unnecessary_to_owned` | ~4 | `.to_string()` → `.as_ref()` |
| `useless_format` | ~2 | `&format!("literal")` → `"literal".to_string()` |
| `unnecessary_unwrap` | ~1 | `if x.is_some() { x.unwrap() }` → `if let Some(x) = x` |
| `collapsible_match` | ~2 | 嵌套 if → match arm |
| `vec_init_then_push` | ~1 | `Vec::with_capacity + push` → `vec![..]` |
| `type_complexity` | ~1 | 提取 type alias |
| `doc_lazy_continuation` | ~2 | 修复 doc 列表缩进 |
| `if_same_then_else` | ~1 | 合并相同分支 |
| `too_many_arguments` | ~4 | 加 `#[allow(clippy::too_many_arguments)]`（公共 API 不宜改签名） |
| `bool_assert_comparison` | ~1 | `assert_eq!(x, false)` → `assert!(!x)` |
| `approx_constant` | ~2 | `3.14` → `std::f64::consts::PI` |
| `needless_update` | ~2 | 删除多余的 `..Default::default()` |
| `non_snake_case` | ~2 | `valAx_count` → `val_ax_count` |

**策略**：从 `_clippy_full.txt` 逐条对照修复，每修一批重新跑 clippy 验证。

### 步骤 4：修复 README 快速开始示例

**文件**：`d:\xcode\xdemo\realtime-screen-ocr-rust\pptx-rs\README.md`

- 第 28 行：`use pptx::{Presentation, Pt, RGBColor};` → `use pptx_rs::{Presentation, Pt, RGBColor};`
- 第 30 行：`fn main() -> pptx::Result<()>` → `fn main() -> pptx_rs::Result<()>`
- 第 33 行：`prs.slides_mut().add_slide()?` → 需补 `id_counter` 参数（参考 lib.rs:41-42 doctest 写法）
- 第 35/36 行：`pptx::Inches(...)` → `pptx_rs::Inches(...)`
- 第 5 行版本号：`0.2.0` → `0.3.0`

### 步骤 5：添加 LICENSE 文件

**文件**：`d:\xcode\xdemo\realtime-screen-ocr-rust\pptx-rs\LICENSE`（新建）

MIT 许可证全文，版权行：`Copyright (c) 2026 pptx-rs contributors`

### 步骤 6：修复 src/lib.rs

**文件**：`d:\xcode\xdemo\realtime-screen-ocr-rust\pptx-rs\src\lib.rs`

- 第 7 行：`0.2.0` → `0.3.0`
- 第 38/44/46 行 doctest：`use pptx::` → `use pptx_rs::`、`pptx::Inches` → `pptx_rs::Inches`
- 第 65 行：`10 个变体` → `11 个变体`（补 `Ppt97`）
- 第 66 行：列出变体补 `Ppt97`
- 第 70 行：`0.1.x` → `0.3.x`
- 第 86-92 行：把 `#![deny]` / `#![warn]` 属性移到所有 `//!` 文档**之前**（Rust 惯例：inner attributes 在 inner docs 之前），或把 88-92 行的 `//!` 文档移到 86 行之前
- 第 92 行：`当前 v0.1.0 中` → `已修复`（补 Debug impl 后删除此说明）
- 添加 `#![forbid(unsafe_code)]`（在第 86 行 `#![deny]` 附近）
- 修复 broken doc links：`[`docs::ARCHITECTURE`]` / `[`docs::CHANGELOG`]` 改为普通文本或 GitHub 绝对 URL；移除指向 `.trae/` 的相对链接（被 exclude 后无效）

### 步骤 7：补 Debug impl

**文件**：`src/slide.rs`（`Shapes` / `ShapesMut` 定义处）

为 `Shapes` 和 `ShapesMut` 添加 `#[derive(Debug)]` 或手动 `impl Debug`（如果含 `Rc<RefCell<>>` 等非 Debug 字段，手动实现输出关键字段）。

### 步骤 8：同步版本号文档

- `docs/CHANGELOG.md`：把 `[Unreleased]` 段改标题为 `## [0.3.0] - 2026-06-29`，新增空 `[Unreleased]` 段
- `.trae/rules/project_rules.md` 第 24 行：`v0.1.0 MVP` → `v0.3.0`（注：此文件被 exclude，不影响发布，但保持一致）
- `src/lib.rs` 第 7 行：`0.2.0` → `0.3.0`（步骤 6 已含）

### 步骤 9：清理公开 API 文档中的 TODO 标记

**文件**：`src/slide.rs`、`src/slide_layouts.rs`、`src/slide_masters.rs` 等

移除公开 `pub` API 文档注释中的 `TODO-007` / `TODO-008` / `TODO-033` / `TODO-037` / `TODO-049` 等内部任务编号（会出现在 docs.rs 公开文档中，对外部用户是噪音）。改为描述性说明或直接删除。

### 步骤 10：处理安全隐患（_ssh_pass.bat / _ssh_run.ps1）

这两个文件已被跟踪且含明文密码 `Xwt@113806`，已在 `exclude` 列表中不会发布，但应从 git 历史中清除：
- `git rm --cached pptx-rs/_ssh_pass.bat pptx-rs/_ssh_run.ps1`（取消跟踪）
- 它们已在 `.gitignore` 的 exclude 中，取消跟踪后不会影响发布包
- **注意**：密码已暴露在 git 历史中，建议用户轮换该密码

### 步骤 11：最终验证（用户在 WSL 中执行）

```bash
cd /mnt/d/xcode/xdemo/realtime-screen-ocr-rust/pptx-rs
cargo fmt --all
cargo clippy --all-targets -- -D warnings
cargo test --all
cargo doc --no-deps
cargo publish --dry-run
```

全部通过后：
```bash
cargo publish
```

## 验证清单

- [ ] `cargo fmt --all -- --check` 通过
- [ ] `cargo clippy --all-targets -- -D warnings` 0 error
- [ ] `cargo test --all` 全绿
- [ ] `cargo doc --no-deps` 无 warning
- [ ] `cargo publish --dry-run` 成功
- [ ] README 快速开始示例可编译
- [ ] `use pptx_rs::` 全局替换无遗漏
- [ ] LICENSE 文件存在
- [ ] 版本号 0.3.0 全局一致

## 关键文件清单

| 文件 | 改动 |
|------|------|
| `Cargo.toml` | 删 [lib] 段、改 URL、加 rust-version、bump 0.3.0 |
| `src/lib.rs` | 重命名 import、版本号、forbid(unsafe_code)、doc 结构、Debug |
| `README.md` | 修示例、重命名 import、版本号 |
| `LICENSE` | 新建 MIT |
| `docs/CHANGELOG.md` | [Unreleased] → [0.3.0] |
| `src/slide.rs` | Debug impl、TODO 清理、clippy 修复 |
| `src/oxml/chart.rs` | clippy 修复（from_str、derivable_impls、useless_format 等） |
| `src/oxml/sppr.rs` | clippy 修复（derivable_impls × 3、from_str × 3） |
| `src/oxml/theme.rs` | clippy 修复（derivable_impls、unnecessary_to_owned × 4） |
| `src/oxml/txbody.rs` | clippy 修复（derivable_impls × 2） |
| `src/oxml/parse_sld.rs` | clippy 修复（collapsible_match、if_same_then_else） |
| 全部 .rs 文件 | `use pptx::` → `use pptx_rs::` |
