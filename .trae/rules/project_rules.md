# pptx-rs 项目协作开发规范（Project Rules）

> 本文件是 **pptx-rs** 项目的核心协作与开发规范。**所有** 贡献者（含 AI 助手）必须遵守。
> 规则优先级：**安全 > 兼容性 > 一致性 > 性能 > 风格**。
> 适用范围：仓库 `e:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs` 下所有 Rust 源码、文档、示例、CI 脚本。

---

## 0. 文档地图（先看这里）

| 文档 | 作用 |
| --- | --- |
| [README.md](../../../README.md) | 项目门面：快速开始 + 路线图 + 与 python-pptx 的对照 |
| [docs/ARCHITECTURE.md](../../../docs/ARCHITECTURE.md) | **架构总览**：opc / oxml / shape 三层模型、调用链 |
| [docs/DEVELOPMENT.md](../../../docs/DEVELOPMENT.md) | 开发指南：环境、构建、调试、提交 |
| [docs/TESTING.md](../../../docs/TESTING.md) | 测试规范：覆盖目标、断言策略、参考样本 |
| [docs/OOXML_REFERENCE.md](../../../docs/OOXML_REFERENCE.md) | OOXML/DrawingML 关键元素的速查表 |
| [docs/CONTRIBUTING.md](../../../docs/CONTRIBUTING.md) | 贡献流程：fork / PR / 评审 checklist |
| [docs/CHANGELOG.md](../../../docs/CHANGELOG.md) | 版本变更日志 |
| [.trae/skills/](../skills/) | AI 助手专用的领域知识包（架构/开发/测试/调试/扩展 等） |
| [rust-coding-standards](../skills/rust-coding-standards/SKILL.md) | Rust 编码规范（与 actionbook/rust-skills 对齐） |

---

## 1. 项目定位

- **目标**：用 Rust 实现一个对标 [python-pptx](https://github.com/scanny/python-pptx) 的 PowerPoint `.pptx` 读写库，**正确性优先**、**API 友好**、**零 unsafe**。
- **当前阶段**：v0.1.0 MVP（创建/打开/保存 `.pptx`；slide、形状、文本、字体、颜色、图片、表格、连接器均可用）。
- **非目标**（明确不做）：
  - 渲染 / 缩略图生成（用 LibreOffice 或 Aspose 替代）。
  - 完整的 Chart / SmartArt / VML 兼容（路线图）。
  - 任何形式的 `unsafe` 块（除 `zip` 内部）。

---

## 2. 仓库布局

```
pptx-rs/
├── Cargo.toml          # 依赖与 profile；版本字段必须与 CHANGELOG.md 同步
├── README.md
├── docs/               # 项目级文档（架构/开发/测试/OOXML/CHANGELOG）
├── examples/           # 可运行的 cargo example（hello / protect / watermark）
├── src/
│   ├── lib.rs          # 顶层导出 + crate-level doc
│   ├── error.rs        # 统一 Error + Result 别名
│   ├── units.rs        # Emu / Pt / Inches / RGBColor
│   ├── presentation.rs # Presentation 高阶 API
│   ├── slide.rs        # Slide / Slides / Shapes 视图
│   ├── slide_layouts.rs
│   ├── slide_masters.rs
│   ├── opc/            # OPC 容器层（zip + Content_Types + 关系）
│   ├── oxml/           # OOXML 模型层（强类型 XML）
│   └── shape/          # 高阶形状 API
├── _test/              # 真实 .pptx 样本（不参与 cargo test）
├── _test_out/          # 处理输出（git ignored）
├── check_*.py          # 验证脚本（保护/水印的离线核验）
├── gen_ref.py          # OOXML 参考文档生成
├── debug-pptx-output-fail.md
└── .trae/              # AI 协作上下文（rules + skills）
```

---

## 3. 语言与版本

- **Rust edition 2021**（`Cargo.toml`）。
- MSRV：**1.75+**（使用 `let-else` / `derive_default_enum` / `array::repeat` 等）。
- 公共 API 文档注释与项目内沟通语言：**中文**。
- 标识符（函数/类型/字段名）保持 **英文**，遵循 Rust 标准命名。
- 不引入拼音或缩写作为标识符（除专业术语如 `EMU`、`lnSpc`）。

---

## 4. 注释规范（强制）

> 这是本轮补全工作的核心。**所有新代码必须先有注释再写实现**。

### 4.1 必须写文档注释的位置

- 所有 `pub` 字段、函数、方法、类型 —— **强制**。
- `pub(crate)` 字段、方法 —— **强制**（便于 reviewer 理解）。
- 私有结构体字段 —— **建议**（特别是 `Rc<RefCell<...>>` / `Cell<u32>` 这类共享状态）。
- 模块顶部 `//!` 文档 —— **强制**，且必须包含：
  1. 一句话功能定位；
  2. 与 python-pptx 的对应关系（如适用）；
  3. 与 OPC/OOXML 哪些元素对应。

### 4.2 函数注释模板

```rust
/// 函数功能简述（一句话）。
///
/// 详细说明（可选）：算法/副作用/错误条件/性能。
///
/// # 参数
/// - `param_name`：说明。
///
/// # 返回值
/// - 成功：返回什么；失败：返回什么错误。
///
/// # 错误
/// - `Error::Xxx`：触发条件。
///
/// # 示例
/// ```no_run
/// // 最小可运行片段
/// ```
pub fn foo(x: i64) -> Result<Y> { ... }
```

### 4.3 内联注释

- **复杂逻辑必须解释"为什么"**，而不是"做了什么"。
- 推荐格式：`// 因为 ...，所以 ...`
- 关键 OOXML 元素顺序、命名空间等约束在源码中就地标注。
- 不写无意义的 `// 计数器` `// i++` 类注释。

### 4.4 注释语言

- 公共 API 文档注释：**中文**。
- 内联注释：**中文**（与代码上下文保持一致）。
- DOXYGEN / cargo doc 链接：本仓库内路径用相对路径，跨仓库用 `https://...`。

---

## 5. 安全红线（禁止模式）

| 模式 | 风险 | 替代方案 |
| --- | --- | --- |
| `unsafe` 块 | UB | 永远不引入；如依赖库需要，封装到 `opc::package` 一层 |
| `unwrap()` / `expect()`（在 lib 路径） | panic | 用 `?` + `Error` 显式传播；测试中允许 `unwrap` |
| 硬编码路径（`C:\\...`） | 平台不可移植 | `PathBuf` / `path.join` |
| 硬编码密码 / 密钥 | 信息泄露 | 环境变量（参考 `examples/protect_pptx.rs` 的盐值常量改为 CLI 参数） |
| `String` 拼接 XML | 转义遗漏 / 性能 | 使用 `oxml::writer::XmlWriter` |
| `serde_json::from_str` 解析 XML | 错误语义 | 走 `quick_xml` SAX |
| `_ = some_func()` 忽略错误 | 静默失败 | 显式 `let _ = ...;` 并注释原因 |

---

## 6. 代码质量

### 6.1 资源管理

- 文件/zip 句柄：在最小作用域内 `let mut` + 不必 `defer`（Rust RAII 即可），但**所有** `File::open` / `ZipArchive::new` 必须紧跟 `?` 错误传播。
- 长生命周期的 `String` / `Vec<u8>`：优先 `Cow<'_, str>` 减少 clone。

### 6.2 并发安全

- 当前 0.1.0 不暴露多线程 API；内部使用 `Rc<Cell<u32>>` / `Rc<RefCell<T>>` 是**预期设计**。
- 引入 `Send + Sync` 之前必须经过 [docs/ARCHITECTURE.md](../../../docs/ARCHITECTURE.md) 评审。
- **禁止**在 lib 路径使用 `Mutex`（仅 `Cargo.toml` 的依赖库内部允许）。

### 6.3 错误处理

- 统一使用 `crate::error::Result<T>`，**禁止**返回 `Result<T, Box<dyn Error>>`（除 `examples/main`）。
- `Error::Other` 仅用于过渡，正式错误请扩展 `enum Error`。
- 解析/序列化错误必须包含：**出错元素名 + 上下文路径**（如 `relationships parse: missing Id`）。

### 6.4 不变性

- `pub` 字段不在 0.1.0 承诺稳定；允许 0.x 调整，**但**类型签名（方法名、参数、返回）需经 deprecation 流程。
- `pub(crate)` 字段调整无需 deprecation，但要在 CHANGELOG 标注 `internal`。

---

## 7. 测试要求

| 模块 | 覆盖率目标 | 备注 |
| --- | --- | --- |
| `units` | ≥ 95% | 浮点转换是核心 |
| `opc` | ≥ 85% | zip 读写必须 round-trip |
| `oxml::*` 序列化 | ≥ 80% | 与 python-pptx 输出 byte-diff |
| `shape::*` | ≥ 75% | 高阶 API 难以全量覆盖 |
| 端到端（hello_pptx example） | 100% | 必须能跑出可被 PowerPoint 打开的 .pptx |

测试规范详见 [docs/TESTING.md](../../../docs/TESTING.md)。

---

## 8. Git 规范

### 8.1 分支策略

- `main`：稳定可发布（0.x.y 标签打在这里）。
- `develop`：日常集成分支。
- `feature/<scope>-<short-desc>`：新功能。
- `fix/<scope>-<short-desc>`：Bug 修复。
- `docs/<scope>`：纯文档。
- `chore/<scope>`：杂项（依赖/CI/工具）。

### 8.2 提交信息（Conventional Commits）

```
<type>(<scope>): <subject>

<body>

<footer>
```

**type 取值**：

| type | 用途 |
| --- | --- |
| `feat` | 新功能 |
| `fix` | Bug 修复 |
| `docs` | 仅文档 |
| `refactor` | 重构（无行为变化） |
| `perf` | 性能优化 |
| `test` | 测试相关 |
| `chore` | 构建/CI/工具 |

**scope 常用值**：`opc` / `oxml` / `shape` / `slide` / `pres` / `units` / `error` / `docs` / `examples` / `ci`。

**示例**：

```
feat(oxml): 支持 Sp 的占位符属性（ph type/idx）

新增 `ph_type` / `ph_idx` 字段，并在 write_xml 时按 OOXML 顺序
插入 <p:ph> 元素。WPS 与 PowerPoint 均能正确识别。

Refs: #42
```

### 8.3 提交前自检

- [ ] `cargo fmt --all`
- [ ] `cargo clippy --all-targets -- -D warnings`
- [ ] `cargo test --all`
- [ ] `cargo doc --no-deps`（无 warning）
- [ ] 更新 [docs/CHANGELOG.md](../../../docs/CHANGELOG.md)
- [ ] 涉及公共 API 改动时同步更新 [README.md](../../../README.md)

---

## 9. 性能 / 资源约束

- **零分配热路径**：序列化 XML 时，**不要**在循环里 `format!()`，用 `XmlWriter` 的链式 API。
- **字符串驻留**：常用 `&'static str`（如 `NS_PRESENTATION_MAIN`）。
- **缓冲复用**：批量生成 slide xml 时复用 `XmlWriter`（见 `presentation.rs::to_opc_package`）。
- **不引入 tokio / async**：保持纯同步库；任何 IO 阻塞都接受。

---

## 10. 跨平台

- 路径处理：仅使用 `std::path::Path` / `PathBuf`，**禁止** `\\` 硬编码。
- 换行符：仓库统一 LF；CI 配 `.gitattributes` 兜底。
- 行尾：源文件 LF。
- 文件权限：zip 内统一 0o644（`OpcPackage::save`）。

---

## 11. 依赖策略

| 类别 | 允许 | 备注 |
| --- | --- | --- |
| zip | `zip` crate | 默认 `deflate` 即可 |
| XML | `quick-xml` 0.40+ | 不引 `serde-xml-rs`（语义不对） |
| 错误 | `thiserror` | 库错误 |
| base64 | `base = "0.22"` | 仅 `examples/protect_pptx` 用 |
| 哈希 | `sha2` | 仅 `examples/protect_pptx` |
| 测试 | `tempfile` | dev-dep |

**禁止**：异步运行时（tokio / async-std）、重型 GUI 依赖、机器学习栈。

升级依赖必须经过 [docs/DEVELOPMENT.md](../../../docs/DEVELOPMENT.md) 中 "依赖升级" 流程。

---

## 12. 发布流程

1. 确认 `develop` 上所有 PR 已合并。
2. `cargo update -p <crate>` + `cargo test`。
3. 更新 [docs/CHANGELOG.md](../../../docs/CHANGELOG.md) 的版本号段。
4. 修改 `Cargo.toml` `version`。
5. 打 tag：`git tag -a v0.x.y -m "..."`。
6. `cargo publish --dry-run` → `cargo publish`。

---

## 13. AI 协作约定

- **AI 修改代码前必须先读 .trae/rules/ 与相关 SKILL.md**。
- **AI 写代码前必须先补全注释**，再写实现。
- **AI 提交前必须自检 §8.3**。
- 不在源码中留下 "AI generated" 之类的签名（仅在 commit message 中说明）。
- 当用户问"如何 X"时，AI 应优先引用本规则文件 + 对应 SKILL.md，而不是凭直觉回答。

---

## 14. FAQ

### Q: 公共方法没注释能不能合？
**A**: 不能。§4 是强制项。

### Q: 我从 python-pptx 抄了一段逻辑，可以省略注释吗？
**A**: 不可以。必须解释 "为什么这样写"（特别是 OOXML 顺序、命名空间等约束）。

### Q: 0.1.0 之后会破坏 API 吗？
**A**: 0.x 期间允许小调整，但需走 deprecation：标记 `#[deprecated]`，给出迁移示例，并在 CHANGELOG 写明。

### Q: 怎么判断一个改动属于 `feat` 还是 `refactor`？
**A**: 是否改变对外可观察行为。是 → `feat`；否 → `refactor`。

---

> **最后修订**：2026-06-13
> **维护者**：pptx-rs 团队
> **本规则文件变更**需提 PR 并在 `docs/CHANGELOG.md` 记录。
