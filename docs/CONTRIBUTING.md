# 贡献指南

> 欢迎向 pptx-rs 贡献代码、文档、测试、示例。
> 提交前请阅读 [project_rules.md](../.trae/rules/project_rules.md) 与 [DEVELOPMENT.md](DEVELOPMENT.md)。

## 1. 行为准则

- **尊重**：所有评论保持专业与友善。
- **就事论事**：聚焦技术问题。
- **建设性**：提出可执行的改进建议。

## 2. 贡献类型

| 类型 | 说明 | 适合 |
| --- | --- | --- |
| Bug 报告 | 提交 issue 描述复现 | 用户 |
| 功能建议 | 提交 issue 描述需求 | 用户 |
| 文档改进 | 修错别字、补示例、翻译 | 任何 |
| 代码贡献 | 修 bug、加功能 | 开发者 |
| 测试 | 加测试用例 | 开发者 |
| 性能优化 | 跑基准 + 提 PR | 高级用户 |

## 3. 报告 Bug

### 3.1 提交 issue

请包含：

- **标题**：简洁描述（如 `add_textbox 中 w=0 时 panic`）
- **环境**：OS、Rust 版本、crate 版本
- **复现**：最小代码片段
- **期望行为**
- **实际行为**（含 panic stack trace）
- **PowerPoint / WPS / LibreOffice 表现**

### 3.2 模板

```markdown
**环境**
- OS: Windows 11 / 10
- Rust: 1.78.0
- pptx-rs: 0.2.0
- 打开软件: PowerPoint 365

**复现**
```rust
let mut prs = Presentation::new()?;
prs.slides_mut().add_slide(prs.id_counter())?
   .shapes_mut()
   .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(0.0), Inches(1.0), "x")?;
prs.save("out.pptx")?;
```

**期望**：保存成功。
**实际**：`unwrap on None at line 215 in slide.rs`
```

## 4. 代码贡献流程

### 4.1 Fork & Clone

```bash
# 1) Fork 仓库
# 2) Clone
git clone https://github.com/<your-name>/pptx-rs.git
cd pptx-rs
# 3) 加 upstream
git remote add upstream https://github.com/<org>/pptx-rs.git
```

### 4.2 拉分支

```bash
git fetch upstream
git checkout -b feature/<scope>-<short-desc> upstream/develop
# 或 fix/<scope>-<short-desc>
```

### 4.3 写代码

- 阅读 [project_rules.md](../.trae/rules/project_rules.md) §4（注释规范）。
- 阅读对应 SKILL.md（[.trae/skills/](../.trae/skills/)）。
- **先写注释，再写实现**。
- 加测试（[TESTING.md](TESTING.md)）。

### 4.4 自检

```bash
cargo fmt --all
cargo clippy --all-targets -- -D warnings
cargo test --all
cargo doc --no-deps
```

### 4.5 提交

```bash
git add <files>
git commit -m "feat(oxml): 支持 Sp 的占位符属性"
# 类型 + 作用域 + 主题（英文，50 字符内）
```

### 4.6 Push & PR

```bash
git push origin feature/<branch>
# 在 GitHub 提 PR
```

PR 描述模板：

```markdown
## 改了什么
- 新增 X
- 修复 Y
- 重构 Z

## 关联 issue
Closes #42

## 测试
- 单元测试
- 集成测试
- 端到端

## 检查清单
- [ ] cargo fmt 通过
- [ ] cargo clippy 通过
- [ ] cargo test 通过
- [ ] 文档已更新
- [ ] CHANGELOG 已更新
```

## 5. 评审流程

- 1 个 reviewer LGTM 后可合入。
- 涉及公共 API 改动需 2 个 reviewer。
- 涉及安全 / 并发需 maintainer 批准。

## 6. 发布周期

- v0.x.y：每 4-6 周一个小版本。
- 大改动在 PR 阶段 review 一周。
- 安全 / 数据丢失类 bug 走 hotfix 流程。

## 7. 沟通

- GitHub Issues：bug / feature
- GitHub Discussions：设计讨论
- 邮件：<maintainer@example.com>（隐私问题）

## 8. 致谢

贡献者将列于 [README.md](../README.md) "Contributors" 段。
