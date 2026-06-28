---
name: "rust-coding-standards"
description: "Rust 编码规范（融合 actionbook/rust-skills + huiali/rust-skills + pptx-rs 项目定制）。覆盖所有权、错误处理、API 设计、性能、测试、文档等。Invoke when user asks about Rust best practices, idioms, code style, or needs a coding rules reference."
---

# Rust 编码规范

> 融合 [actionbook/rust-skills](https://github.com/actionbook/rust-skills) 179 条规则 + [huiali/rust-skills](https://github.com/huiali/rust-skills) 40+ 子技能体系，按 pptx-rs 项目实际需求裁剪。
> huiali 版本新增：Solution Patterns（模式化解决方案）、Workflow（决策流程）、Review Checklist（审查清单）、Verification Commands（验证命令）。

## 1. 优先级一览

| 优先级 | 类别 | 影响 | 前缀 |
| --- | --- | --- | --- |
| 1 | 所有权与借用 | CRITICAL | `own-` |
| 2 | 错误处理 | CRITICAL | `err-` |
| 3 | 内存优化 | CRITICAL | `mem-` |
| 4 | API 设计 | HIGH | `api-` |
| 5 | 异步/await | HIGH | `async-` |
| 6 | 编译器优化 | HIGH | `opt-` |
| 7 | 命名规范 | MEDIUM | `name-` |
| 8 | 类型安全 | MEDIUM | `type-` |
| 9 | 测试 | MEDIUM | `test-` |
| 10 | 文档 | MEDIUM | `doc-` |
| 11 | 性能模式 | MEDIUM | `perf-` |
| 12 | 项目结构 | LOW | `proj-` |
| 13 | Clippy | LOW | `lint-` |
| 14 | 反模式 | REFERENCE | `anti-` |

## 2. 所有权与借用（CRITICAL）

- **`own-borrow-over-clone`**：优先 `&T` 借用，避免无谓 `.clone()`。
- **`own-slice-over-vec`**：参数用 `&[T]` / `&str`，不用 `&Vec<T>` / `&String`。
- **`own-cow-conditional`**：条件拥有时用 `Cow<'a, T>`（如 `part.blob_text()`）。
- **`own-rc-single-thread`**：单线程共享用 `Rc<T>`（本项目内部用 `Rc<Cell<u32>>` / `Rc<RefCell<T>>`）。
- **`own-refcell-interior`**：单线程内部可借用 `RefCell<T>`。
- **`own-arc-shared`**：跨线程共享用 `Arc<T>`（0.1.0 不暴露多线程）。
- **`own-copy-small`**：`Copy` 仅给小、不含引用的类型（如 `Emu` / `Pt` / `RGBColor`）。
- **`own-clone-explicit`**：显式 `Clone`，避免隐式复制。
- **`own-move-large`**：大数据 move 而非 clone。
- **`own-lifetime-elision`**：依赖 elision，避免显式 `<'a>` 滥用。

## 3. 错误处理（CRITICAL）

- **`err-thiserror-lib`**：库用 `thiserror`（本项目 `Error` 枚举已用）。
- **`err-anyhow-app`**：应用层用 `anyhow`。
- **`err-result-over-panic`**：库用 `Result`，不要 `panic!`。
- **`err-context-chain`**：用 `Error::oxml("...")` 加上下文。
- **`err-no-unwrap-prod`**：生产代码不要 `unwrap()`。
- **`err-expect-bugs-only`**：`.expect()` 仅用于编程错误（invariant 违反）。
- **`err-question-mark`**：用 `?` 传播。
- **`err-from-impl`**：`#[from]` 自动转换（`Io` / `Zip` 已用）。
- **`err-source-chain`**：`#[source]` 链接底层。
- **`err-lowercase-msg`**：错误消息小写、句末无标点。
- **`err-doc-errors`**：文档化 `# Errors` 段。
- **`err-custom-type`**：自定义错误类型，不 `Box<dyn Error>`。

## 4. 内存优化（CRITICAL）

- **`mem-with-capacity`**：`Vec::with_capacity(n)`（如 `Vec::with_capacity(entry.size() as usize)`）。
- **`mem-smallvec`**：常空小集合考虑 `SmallVec`（v0.1.0 未用）。
- **`mem-avoid-format`**：字符串字面量用 `push_str`，不要 `format!()`。
- **`mem-write-over-format`**：`write!(&mut s, ...)` 替代 `format!()`。
- **`mem-zero-copy`**：解析用 `quick_xml` SAX，不一次性 `String`。
- **`mem-smaller-integers`**：用最小够用整数类型（如 id 用 `u32`，pptx 强校验不超过 `u32::MAX`）。
- **`mem-reuse-collections`**：循环中复用 `XmlWriter` / `String`（`presentation.rs::to_opc_package` 已做）。
- **`mem-assert-type-size`**：热路径类型静态断言（v0.1.0 未用，未来可加）。

## 5. API 设计（HIGH）

- **`api-builder-pattern`**：复杂构造用 builder（本项目 `FreeformBuilder` 是雏形）。
- **`api-must-use`**：`Result` / `&mut` 返回值加 `#[must_use]`。
- **`api-newtype-ty`**：用 newtype 区分语义（如 `Emu` / `Pt` / `Inches`）。
- **`api-sealed-trait`**：sealed trait 禁止外部 impl（v0.1.0 暂未用）。
- **`api-extension-trait`**：扩展 trait 添加方法（本项目 `EmuExt`）。
- **`api-parse-dont-validate`**：边界 parse 成验证过的类型（本项目 `PartName::new`）。
- **`api-impl-into`**：接受 `impl Into<T>`（如 `name: impl Into<String>`）。
- **`api-impl-asref`**：接受 `impl AsRef<Path>`（如 `path: P: AsRef<Path>`）。
- **`api-non-exhaustive`**：`pub enum` 加 `#[non_exhaustive]` 预留扩展。
- **`api-typestate`**：类型状态机（v0.1.0 暂未用）。

## 6. 异步/await（HIGH）

- **0.1.0 不引入异步**。所有 IO 都是同步阻塞。
- 任何"看起来该 async"的需求，先用 sync 实现 + 评估是否真的需要并发。
- 未来如果引入：用 `tokio`，且仅在 `examples/`。

## 7. 编译器优化（HIGH）

- **`opt-inline`**：小函数加 `#[inline]`（如 `Emu::new` / `EmuExt::emu`）。
- **`opt-const-fn`**：纯计算加 `const fn`。
- **`opt-lto`**：release profile 开 `lto = "thin"` + `codegen-units = 1`（已配）。
- **`opt-derive-default`**：能 `#[derive(Default)]` 就 derive。
- **`opt-derive-eq`**：可比较类型 derive `Eq` / `PartialEq` / `Hash`。
- **`opt-minimize-monomorphization`**：避免过度的泛型实例化。

## 8. 命名规范（MEDIUM）

- **模块 / 类型 / 函数 / 变量**：`snake_case`。
- **类型参数**：`T` / `U` / `K` / `V`。
- **生命周期**：`'a` / `'de`。
- **常量**：`SCREAMING_SNAKE_CASE`。
- **特征（trait）**：形容词或名词（`Shape` / `EmuExt`）。
- **crate 名**：`kebab-case`（本项目 `pptx-rs`）。
- **不缩写**：`http` 不写成 `h`，`presentation` 不写成 `pres`（除专业术语）。
- **OOXML 元素名保留原样**（如 `txBody` / `spPr`），但 Rust 标识符用 `tx_body` / `sp_pr` 风格。
- **错误消息小写，句末无标点**（如 `"missing Id attribute"`，不是 `"Missing Id."`）。

## 9. 类型安全（MEDIUM）

- **`type-newtype-ty`**：`Emu(i64)` 包装避免裸 `i64` 混淆。
- **`type-validated-parse`**：`PartName::new` 校验后才返回。
- **`type-enum-repr`**：不 `#[repr(u32)]` 暴露内部数字。
- **`type-phantom-data`**：需要时用 `PhantomData<T>` 携带类型参数。
- **`type-asref-bounds`**：泛型约束 `T: AsRef<Path>` 而非具体类型。

## 10. 测试（MEDIUM）

- **覆盖率目标**：见 [project_rules](../../rules/project_rules.md) §7。
- **每个 `pub` 函数**至少有 1 个测试。
- **命名**：`<unit>_<scenario>`（如 `unit_conversion` / `relationships_round_trip`）。
- **`#[should_panic(expected = "...")]`** 优先用 `expected`。
- **`#[ignore]`** 用于重资源测试，CI 用 `cargo test -- --ignored`。
- **集成测试**放 `tests/`，端到端用 `examples/`。
- **快照测试**用 `insta`（v0.1.0 未用）。

## 11. 文档（MEDIUM）

- **每个 `pub` 项必须有 doc 注释**。
- **模块顶部** `//!` 说明。
- **crate 顶部**在 `lib.rs`。
- **`# Examples` / `# Errors` / `# Panics` 段**按 rustdoc 习惯。
- **不写"performs X" 类废话**，写"什么场景用 / 副作用 / 复杂度"。

## 12. 性能模式（MEDIUM）

- **`perf-streaming`**：解析走 SAX（`quick_xml`），不一次性 `String`。
- **`perf-reuse-buffer`**：`String::with_capacity` + 复用。
- **`perf-avoid-allocations`**：循环内不 `Vec::new()`。
- **`perf-batch-io`**：批量写文件（`zip::start_file` 一次一个 part）。
- **`perf-lazy-init`**：大结构用 `OnceCell` / `Lazy`（v0.1.0 暂未用）。

## 13. 项目结构（LOW）

- **`proj-flat-modules`**：本项目用扁平子模块（`src/oxml/...`），不用 `oxml::raw::*` 嵌套。
- **`proj-pub-use-mod`**：`lib.rs` 顶层 `pub use` 公共项。
- **`proj-feature-flags`**：v0.1.0 单一 feature；未来加 `serde` / `async`。
- **`proj-workspace`**：v0.1.0 单 crate，不 workspace。

## 14. Clippy（LOW）

- `cargo clippy --all-targets -- -D warnings`（CI 必跑）。
- 关键 lint：
  - `clippy::needless_return`
  - `clippy::redundant_clone`
  - `clippy::single_match`
  - `clippy::too_many_arguments`
  - `clippy::missing_errors_doc`
  - `clippy::missing_panics_doc`

## 15. 反模式（REFERENCE）

- **`anti-unwrap-prod`**：不要 `unwrap()` 在 lib 路径。
- **`anti-stringly-typed`**：用 enum 不用字符串。
- **`anti-mutex-soup`**：不要 `Arc<Mutex<Everything>>`。
- **`anti-unsafe-first`**：先不用 `unsafe`。
- **`anti-panic-in-lib`**：库不 panic。
- **`anti-magic-numbers`**：常量集中（如 `presentation.rs::DEFAULT_WIDTH_EMU`）。
- **`anti-deep-nesting`**：早 return / `?` / `let-else`。
- **`anti-global-state`**：避免 `static mut` / `lazy_static`（v0.1.0 不需要）。
- **`anti-dead-code`**：提交前 `cargo build --release` 检查。

## 16. pptx-rs 项目特殊约定

| 场景 | 约定 |
| --- | --- |
| 加形状 ID | 必须从 `Rc<Cell<u32>>` 拿（`Slide::next_shape_id`） |
| 序列化 XML | 必须用 `XmlWriter`，不 `format!()` |
| 写关系 | 走 `Relationships::add`，不要直接拼字符串 |
| 调试打印 | `#[cfg(debug_assertions)] eprintln!(...)` |
| 错误信息 | 必含"出错元素名" |
| OOXML 顺序 | 严格遵守（参考 [`pptx-rs-ooxml`](../pptx-rs-ooxml/SKILL.md)） |
| 公共字段 | 0.1.0 不承诺稳定；加 `#[non_exhaustive]` 留口 |
| 私有 `pub(crate)` | 必须注释（reviewer 看的） |

## 17. 速查表（5 秒）

```rust
// ✅ 推荐
pub fn add_textbox<L, T, W, H>(left: L, top: T, width: W, height: H) -> Result<TextBox>
where L: EmuExt, T: EmuExt, W: EmuExt, H: EmuExt
{ ... }

// ❌ 反模式
pub fn add_textbox(left: f64, top: f64, width: f64, height: f64) -> Result<TextBox, Box<dyn std::error::Error>>
{ ... }
```

---

## 18. Solution Patterns（模式化解决方案）

> 以下模式来自 [huiali/rust-skills](https://github.com/huiali/rust-skills) 的"Solution Patterns"方法论，
> 将常见问题归类为可复用的代码模式，每个模式包含 ✅ 正确做法 + ❌ 反模式 + 适用场景。

### Pattern 1: 错误分类决策

```rust
// 问题：如何选择 Option / Result / panic？

// "不存在"是正常情况 → Option<T>
fn find_shape_by_name(shapes: &[Sp], name: &str) -> Option<&Sp> {
    shapes.iter().find(|s| s.name == name)
}

// "失败"是预期可恢复的 → Result<T, E>
fn parse_sld(xml: &str) -> Result<Sld> {
    // ... XML 解析可能失败
}

// 不变式违反 / 编程 bug → panic!()
assert_eq!(sld.shapes.len(), expected);  // 测试中
```

### Pattern 2: 自定义错误类型（thiserror）

```rust
// ✅ 库代码用 thiserror 枚举
#[derive(Error, Debug)]
pub enum Error {
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),
    #[error("xml error: {0}")]
    Xml(String),
    #[error("not found: {0}")]
    NotFound(String),
}

// ❌ 不要用 Box<dyn Error>
fn bad() -> Result<T, Box<dyn std::error::Error>> { ... }
```

### Pattern 3: 预分配 + 复用

```rust
// ✅ 预分配已知大小
let mut vec = Vec::with_capacity(shapes.len());

// ✅ 循环中复用 XmlWriter / String
let mut w = XmlWriter::new();
for slide in &slides {
    w.clear();  // 复用内部 buffer
    slide.write_xml(&mut w);
}

// ❌ 循环内每次新建
for slide in &slides {
    let mut w = XmlWriter::new();  // 每次分配！
}
```

### Pattern 4: Newtype 区分语义

```rust
// ✅ Emu / Pt / Inches 是不同类型，编译期防混用
pub struct Emu(pub i64);
pub struct Pt(pub f64);
pub struct Inches(pub f64);

// ❌ 裸 i64 / f64 容易混淆
fn set_width(width: i64) { ... }  // 是 EMU 还是 Pt？
```

### Pattern 5: Builder 模式构造复杂对象

```rust
// ✅ 复杂构造用 builder
let mut tb = TextBox::new(name);
tb.set_text("hello");
tb.set_font_size(Pt(24.0));
tb.set_bold(true);

// ❌ 多参数构造函数
TextBox::new(name, "hello", 24.0, true, false, ...);  // 参数太多
```

---

## 19. Workflow（决策流程）

> 来自 huiali/rust-skills 的"Workflow"方法论——按步骤做决策，避免遗漏。

### 错误处理决策流程

```
1. 这是"不存在"还是"失败"？
   → 不存在（正常）: Option<T>
   → 失败（预期可恢复）: Result<T, E>
   → 不变式违反: panic!()

2. 在库代码还是应用代码？
   → 库（公共 API）: thiserror（类型化错误）
   → 应用（内部）: anyhow（灵活错误 + context）

3. 能否让调用方处理？
   → 能: 返回 Result，用 ? 传播
   → 需要加上下文: .context("why it failed")?
   → 必须在此处理: match / if let / unwrap_or
```

### 性能优化决策流程

```
1. 确认瓶颈（测量，不要猜）
   → cargo bench / Instant::now() 计时

2. 算法是否最优？（10x-1000x 影响）
   → 否 → 换算法

3. 数据结构是否合适？（2x-10x 影响）
   → 否 → 换结构

4. 是否有不必要分配？（2x-5x 影响）
   → 是 → 预分配 / 复用 / Cow

5. 缓存是否友好？（1.5x-3x 影响）
   → 否 → 调整布局

6. 能否并行化？（2x-8x 影响）
   → 能 → rayon / 多线程
```

### API 设计决策流程

```
1. 是否需要多种构造方式？
   → 是 → Builder 模式

2. 返回值是否不应被忽略？
   → 是 → #[must_use]

3. 参数是否需要多种类型？
   → 是 → impl Into<T> / impl AsRef<Path>

4. 枚举是否可能扩展？
   → 是 → #[non_exhaustive]

5. 是否需要区分语义相同的类型？
   → 是 → Newtype 模式
```

---

## 20. Review Checklist（审查清单）

> 代码审查时逐项检查，确保不遗漏关键规范。

### 错误处理审查

- [ ] 所有可失败操作返回 `Result` 或 `Option`
- [ ] 错误类型有意义（不只是 `String`）
- [ ] 错误上下文通过传播链保留
- [ ] `unwrap()` 仅在测试或有注释说明时使用
- [ ] `panic!()` 仅用于 bug 或不可恢复状态
- [ ] 库代码用类型化错误（thiserror）
- [ ] 错误消息可操作（含元素名/路径）
- [ ] 无静默吞错（`let _ = ...` 须注释原因）

### 性能审查

- [ ] 已 profile 确认瓶颈
- [ ] 算法对用例最优
- [ ] 数据结构合适
- [ ] 不必要分配已移除
- [ ] 循环内复用 buffer/collection
- [ ] 基准测试显示改进

### API 审查

- [ ] `pub` 项有文档注释
- [ ] `#[must_use]` 在 `Result` / `&mut` 返回值上
- [ ] `#[non_exhaustive]` 在可能扩展的 `enum` 上
- [ ] 参数用 `impl Into<T>` / `impl AsRef<Path>` 而非具体类型
- [ ] Newtype 区分语义相同的不同类型
- [ ] Builder 模式用于复杂构造

### pptx-rs 项目特殊审查

- [ ] 形状 ID 从 `Rc<Cell<u32>>` 拿
- [ ] XML 序列化用 `XmlWriter`，不 `format!()`
- [ ] 关系写走 `Relationships::add`
- [ ] 错误信息含"出错元素名"
- [ ] OOXML 子元素顺序正确
- [ ] `pub(crate)` 字段有注释
- [ ] 公共 API 有中文文档注释

---

## 21. Verification Commands（验证命令）

```bash
# 编译检查
cargo check

# 代码格式化
cargo fmt --all -- --check

# Clippy 检查（CI 必跑）
cargo clippy --all-targets -- -D warnings

# 检查 unwrap/expect 使用
cargo clippy -- -W clippy::unwrap_used -W clippy::expect_used

# 检查生产代码中的 panic
cargo clippy -- -W clippy::panic

# 检查未使用的 Result
cargo clippy -- -D unused_must_use

# 运行全部测试
cargo test --all

# 生成文档（无 warning）
cargo doc --no-deps

# Release 构建
cargo build --release
```

---

## 22. 外部参考

- [actionbook/rust-skills](https://github.com/actionbook/rust-skills) — 179 条规则的完整版
- [huiali/rust-skills](https://github.com/huiali/rust-skills) — 40+ 子技能体系（Solution Patterns + Workflow + Checklist）
- [Rust API Guidelines](https://rust-lang.github.io/api-guidelines/) — 官方 API 指南
- [The Rustonomicon](https://doc.rust-lang.org/nomicon/) — unsafe 细节（本项目不用）
- [Effective Rust](https://www.lurklurk.org/effective-rust/) — David Drysdale 经验集
