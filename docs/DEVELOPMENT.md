# 开发指南

> 本文件是 pptx-rs 项目的日常开发参考手册。包含：环境、构建、调试、依赖升级、发布。
> 速查版见 [.trae/skills/pptx-rs-development/SKILL.md](../.trae/skills/pptx-rs-development/SKILL.md)。

## 1. 环境要求

| 工具 | 版本 | 说明 |
| --- | --- | --- |
| Rust | ≥ 1.75 | 推荐 `rustup default stable` |
| Cargo | 随 Rust | — |
| Git | ≥ 2.30 | Windows 下推荐 `git for windows` |
| PowerPoint / WPS / LibreOffice | 任意 | 用于肉眼验证 .pptx / .ppt 正确性 |
| 7-Zip | 任意 | 调试 .pptx 内部结构 |
| Python + msoffcrypto-tool | 任意 | 验证 .ppt 加密输出（`pip install msoffcrypto-tool`） |

可选：

- `cargo-watch`：文件变更自动重跑
- `cargo-expand`：宏展开调试
- `cargo-bloat`：看二进制体积
- `cargo-outdated`：看依赖过时
- `cargo-audit`：依赖安全审计

## 2. 5 分钟上手

```powershell
# 1) 克隆
git clone <repo-url> pptx-rs && cd pptx-rs

# 2) 构建
cargo build

# 3) 跑示例
cargo run --example hello_pptx

# 4) 验证
#    用 PowerPoint 打开 hello.pptx

# 5) 跑测试
cargo test --all

# 6) 生成文档
cargo doc --no-deps --open
```

## 3. 常用命令

### 构建

```powershell
cargo build                          # debug
cargo build --release                # release
cargo check                          # 快速类型检查
```

### 静态检查

```powershell
cargo fmt --all -- --check           # 格式检查
cargo clippy --all-targets -- -D warnings  # clippy 当错误
RUSTDOCFLAGS="-D warnings" cargo doc --no-deps  # 文档 warning 当错误
```

### 测试

```powershell
cargo test                                       # 全部
cargo test --doc                                 # 文档测试
cargo test units::                               # 按模块过滤
cargo test -- --nocapture                        # 打印 stdout
cargo test -- --test-threads=1                   # 串行
cargo test --all-features                        # 全部 feature
cargo test -- --ignored                          # 跑 @ignore 测试
```

### 示例

```powershell
# .pptx 示例
cargo run --example hello_pptx
cargo run --example protect_pptx
cargo run --example watermark_pptx
cargo run --example watermark_and_protect

# .ppt 示例（PowerPoint 97-2003 二进制格式）
# 注：业务逻辑位于 src/ppt97/ 模块，examples 仅作为薄封装的命令行入口
cargo run --example protect_ppt                 # 调用 pptx::ppt97::encrypt
cargo run --example watermark_ppt               # 调用 pptx::ppt97::add_watermark
cargo run --example watermark_and_protect_ppt   # 调用 pptx::ppt97::add_watermark_and_encrypt
```

## 4. 调试 .pptx

### 4.1 用 7-Zip 拆解

```powershell
# 1) 解压
Expand-Archive hello.pptx -DestinationPath hello_out

# 2) 看文件
Get-ChildItem -Recurse hello_out
```

### 4.2 看关键 XML

```powershell
# [Content_Types].xml
Get-Content hello_out\[Content_Types].xml

# 第 1 张 slide
Get-Content hello_out\ppt\slides\slide1.xml

# 关系文件
Get-Content hello_out\ppt\slides\_rels\slide1.xml.rels
```

### 4.3 与 python-pptx 输出对比

```python
# ref.py
from pptx import Presentation

p = Presentation()
slide = p.slides.add_slide(p.slide_layouts[0])
tb = slide.shapes.add_textbox(...)
tb.text = "Hello, rust-pptx!"

p.save("ref.pptx")
```

```powershell
# 跑参考
python ref.py
# 解压
Expand-Archive ref.pptx -DestinationPath ref_out
# diff
Compare-Object (Get-Content hello_out\ppt\presentation.xml) (Get-Content ref_out\ppt\presentation.xml)
```

### 4.4 用 LibreOffice 验证

```powershell
# 命令行转 PDF
& 'C:\Program Files\LibreOffice\program\soffice.exe' --headless --convert-to pdf hello.pptx
```

LibreOffice 容忍度高于 PowerPoint，能打开说明基本结构正确。

## 4A. 调试 .ppt（97-2003 二进制格式）

`.ppt` 文件与 `.pptx` 完全不同（OLE2/CFB 容器 + 二进制 record 树，非 ZIP+XML），
调试方法也不同。业务逻辑位于 [`src/ppt97/`](../src/ppt97/) 模块。

### 4A.1 测试样本

`.ppt` 测试样本位于 `_test/` 目录（不参与 `cargo test`）：

```powershell
# 列出测试样本
Get-ChildItem _test\*.ppt

# 输出目录（git ignored）
mkdir _test_out -Force
```

### 4A.2 用 Python 探查 OLE2 结构

```python
# probe_ppt.py
import olefile

ole = olefile.OleFileIO('_test/心理账户理论.ppt')
for stream in ole.listdir():
    print('/'.join(stream))
ole.close()
```

### 4A.3 用 msoffcrypto-tool 验证加密

```powershell
# 安装
pip install msoffcrypto-tool

# 验证（脚本会检查 is_encrypted / load_key / decrypt）
python verify_ppt_crypto.py
python verify_ppt_watermark.py
python verify_watermark_and_protect.py
```

### 4A.4 用 WPS / PowerPoint 人工验证

- 打开 `_test_out/wm_*.ppt`：确认水印可见且不可编辑
- 打开 `_test_out/protected_*.ppt`：输入密码 `pptx-rs-secret` 后正常打开
- 打开 `_test_out/wm_protected_*.ppt`：输入密码后看到水印且不可编辑

### 4A.5 常见 .ppt 错误

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| WPS 打开报"文件已损坏" | OLE2 mini FAT 多分配扇区 | 检查 `ppt97::ole::fix_mini_fat` 调用 |
| 水印可被编辑（变成文本框） | z-order 最高（在 SpgrContainer 末尾） | 改为插入到组形状本身之后（z-order 最低） |
| 加密后 WPS 密码验证通过但打不开 | 加密了 Pictures stream | 不加密 Pictures stream（WPS 严格检查） |
| `persist directory parse: entry N offset out of range` | PersistDirectoryEntry 解析错误 | 检查 `persistId(20bit) \| cPersist(12bit)` 头格式 |
| UserEditAtom recLen != 28 错误 | 文件已加密 | `encrypt` 不支持二次加密，需先解密 |

## 5. 排查常见错误

### 5.1 PowerPoint 拒绝打开

按 [.trae/skills/pptx-rs-debugging/SKILL.md](../.trae/skills/pptx-rs-debugging/SKILL.md) 流程排查。

### 5.2 中文乱码

- 检查 XML 头是否 `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`。
- 源码必须是 UTF-8（默认）。
- 字符串来源不能是 `from_utf8_lossy`（v0.1.0 不用）。

### 5.3 图片不显示

- 检查 `/ppt/media/` 下文件存在。
- 检查 rels 中 `Target` 路径正确。
- 检查 `<a:blip r:embed="rIdX">` 中 rId 与 rels 一致。

### 5.4 clippy 警告

```powershell
cargo clippy --all-targets -- -D warnings
```

按警告信息修复；常用：

| 警告 | 修复 |
| --- | --- |
| `needless_return` | 删 `return` |
| `redundant_clone` | 改借用 |
| `single_match` | `if let` / `matches!` |
| `too_many_arguments` | builder |
| `missing_errors_doc` | 加 `# Errors` |

## 6. 依赖升级

### 6.1 流程

1. 升级前看 changelog（crates.io / GitHub Releases）。
2. 在 `Cargo.toml` 改版本。
3. 跑 `cargo test --all` + `cargo clippy --all-targets`。
4. 重点测试 [`src/oxml/parser.rs`](../src/oxml/parser.rs)（`quick-xml` API 不稳）。
5. 更新 [CHANGELOG.md](CHANGELOG.md)。

### 6.2 已知风险

| crate | 风险 | 关注点 |
| --- | --- | --- |
| `zip` | 中 | `ZipArchive::new` / `ZipWriter::start_file` 签名 |
| `quick-xml` | 高 | `Events` / `BytesStart` 反复改 |
| `thiserror` | 低 | `#[from]` 稳定 |
| `base64` | 低 | `Engine::encode` 稳定 |
| `sha2` | 低 | `Digest::update` 稳定 |

## 7. 性能

### 7.1 测量

v0.1.0 暂未集成 criterion；可手动：

```rust
use std::time::Instant;

let start = Instant::now();
let bytes = prs.to_bytes()?;
println!("to_bytes = {:?}", start.elapsed());
```

### 7.2 优化

- 复用 `XmlWriter`（已做）。
- 避免 `format!` 循环（已避免）。
- 大量 `String` 用 `with_capacity`（已用）。
- 批量 IO（`zip::start_file` 一次一个 part）。

## 8. 双仓库工作流与发布流程

### 8.1 双仓库架构

本项目同时使用两个 git remote，**可见性不同，用途不同**：

| remote | URL | 可见性 | 用途 |
| --- | --- | --- | --- |
| `github` | `https://github.com/WenTao-Love/pptx-rs.git` | **公开** | crates.io 发布配套源码、社区协作 |
| `origin` | `https://gitcode.com/szlaoxu/realtime-screen-ocr.git` | **私有** | 日常开发、自用、含业务测试样本 |

> ⚠️ **关键区别**：`github` 是公开仓库，**任何提交的内容都会被全世界看到**（含爬虫/快照）；
> `origin` 是私有仓库，保密性强，可放心存放业务测试样本。

父仓库 `realtime-screen-ocr-rust` 同时管理 pptx-rs 及其他子项目。`pptx-rs-github` 分支
是专门用于推送到公开 GitHub 的独立分支，**不包含业务测试样本和敏感文件**。

#### 工作流

```text
日常开发        → 父仓库 rust 分支（推到 origin 私有仓库，含 _test/ 业务样本）
更新公开仓库    → pptx-rs-github 分支（推到 github 公开仓库，仅源码）
发布到 crates.io → cargo publish（仅打包 src/ 等必要文件）
```

### 8.2 安全红线（⚠️ 血的教训）

> **2026-06-29 事故复盘**：
> 明文密码文件 `_ssh_pass.bat` / `_ssh_run.ps1`（含 VirtualBox 密码 `Xwt@113806`）
> 和业务测试样本 `_test/心理账户理论.ppt` / `_test/文旅IP人设打造抖音短视频运营方案.pptx`
> 通过 `git subtree split --prefix=pptx-rs` 被打包到 `pptx-rs-github` 分支并推送到公开 GitHub。
> 此外 commit `7b8274c` 的 message 含非专业表述（直接把对话内容当 commit message）。
>
> **影响**：密码在公开环境暴露过（即使已清理，仍可能被爬虫/快照抓取），需更改密码。
>
> **处置**：通过 `git filter-branch --index-filter` + `--msg-filter` 重写历史并 force push；
> 父仓库执行 `git rm --cached` 移除跟踪；.gitignore 补全忽略规则。

#### 绝对禁止（⚠️ 红线）

| 禁止项 | 原因 | 正确做法 |
| --- | --- | --- |
| 提交明文密码/密钥文件到 git | 公开仓库被爬虫秒抓 | 用环境变量 / `.env`（.gitignore 忽略） |
| 提交业务测试样本到公开仓库 | 文件名含业务信息 | `_test/` 已在 .gitignore 忽略，已 `git rm --cached` |
| 用 `git subtree split` 推送到 GitHub | 父仓库历史中的敏感文件会重新暴露 | 直接在 `pptx-rs-github` 分支上开发 |
| commit message 写非专业表述 | 公开仓库会被社区看到 | 遵循 Conventional Commits（见 [project_rules.md §8.2](../.trae/rules/project_rules.md)） |
| 将 `_ssh_pass.bat` / `_ssh_run.ps1` / `_askpass.bat` 加入 git | 含明文 SSH 密码 | 已在 .gitignore 忽略，绝不入库 |

#### .gitignore 已忽略的敏感文件

见父仓库根目录 `.gitignore` 的 "pptx-rs" 段落：

```gitignore
# SSH 辅助文件（含明文密码/敏感连接信息，绝不入库）
/pptx-rs/_askpass.bat
/pptx-rs/_ssh_pass.bat
/pptx-rs/_ssh_run.ps1
```

> 注：`_test/` 目录下的业务测试样本已通过 `git rm --cached` 从父仓库 git 跟踪中移除（本地文件保留），
> 防止 `git subtree split` 时被重新打包到公开分支。如需在 gitcode 私有仓库中跟踪，可 `git add _test/`
> 重新添加（gitcode 保密性强，可接受）；但**绝不**用 `git subtree split` 推送到公开 GitHub。

#### 更新公开仓库代码的正确方式

```powershell
# ✅ 正确：直接在 pptx-rs-github 分支上开发
git checkout pptx-rs-github
# ... 修改代码 ...
git add .
git commit -m "feat(xxx): ..."
git push github pptx-rs-github:main
```

```powershell
# ❌ 错误：用 subtree split（会重新暴露父仓库历史中的敏感文件）
# git subtree split --prefix=pptx-rs
# git push github pptx-rs-github:main
```

#### 如果敏感信息已暴露到公开仓库

1. **立即更改密码**（即使后续清理了历史，密码可能已被抓取）
2. 用 `git filter-branch` 重写历史：
   ```bash
   # 同时删除文件 + 修改 commit message
   FILTER_BRANCH_SQUELCH_WARNING=1 git filter-branch --force \
     --index-filter 'git rm -r --cached --ignore-unmatch <敏感文件路径>' \
     --msg-filter 'if [ "$GIT_COMMIT" = "<commit-hash>" ]; then printf "规范的 commit message"; else cat; fi' \
     --prune-empty \
     <分支名>
   ```
3. `git push -f github <分支名>:main`（force push 覆盖远端历史）
4. 父仓库执行 `git rm --cached <敏感文件>` + 更新 .gitignore + commit
5. 通知所有 clone 过该仓库的协作者重新 clone（旧历史已失效）

> ⚠️ force push 是不可逆操作，执行前务必备份分支：`git branch -f <分支名>-backup <分支名>`

### 8.3 发布流程

1. 确认 `develop` 上 CI 绿、PR 全部合并。
2. `cargo update`（允许锁文件前进）。
3. 更新 [CHANGELOG.md](CHANGELOG.md)。
4. 改 `Cargo.toml` 的 `version`。
5. `cargo test --all` + `cargo clippy --all-targets -- -D warnings` + `cargo doc --no-deps`。
6. `git commit -m "chore(release): v0.x.y"`。
7. `git tag -a v0.x.y -m "v0.x.y"`。
8. `cargo publish --dry-run`。
9. `cargo publish`。
10. 推 tag：`git push github v0.x.y`（注意：推到 `github` 公开仓库，不是 `origin`）。
11. 写 GitHub Release。

## 9. IDE 集成

### VS Code

- `rust-analyzer` 扩展
- `settings.json`：
  ```json
  {
    "rust-analyzer.cargo.features": "all",
    "[rust]": { "editor.formatOnSave": true }
  }
  ```

### Trae IDE

- 自动加载 [`.trae/rules/project_rules.md`](../.trae/rules/project_rules.md)
- 按需加载 [`.trae/skills/`](../.trae/skills/)

## 10. 故障排查

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| `linker not found` | Windows 缺 MSVC | `rustup default stable-msvc` |
| 生成的 .pptx PowerPoint 拒绝 | OOXML 元素顺序错 | 对照 OOXML_REFERENCE.md |
| 字号/位置不对 | EMU/Pt 转换错 | 重新查 `src/units.rs` 系数 |
| 中文乱码 | XML 未声明 UTF-8 | `XmlWriter::decl()` |
| 表格列宽不对 | `Table` 没指定 `Col.width` | 显式设置 |

## 11. 工具脚本

### .pptx 验证脚本

- [`check_protect.py`](../check_protect.py) — 验证 `protect_pptx.rs` 输出
- [`check_wm.py`](../check_wm.py) — 验证 `watermark_pptx.rs` 输出
- [`gen_ref.py`](../gen_ref.py) — 生成 OOXML 参考
- [`debug-pptx-output-fail.md`](../debug-pptx-output-fail.md) — 历史失败案例

### .ppt 验证脚本

- [`verify_ppt_crypto.py`](../verify_ppt_crypto.py) — 验证 `protect_ppt.rs` 输出（msoffcrypto 加密验证）
- [`verify_ppt_watermark.py`](../verify_ppt_watermark.py) — 验证 `watermark_ppt.rs` 输出（OLE2 结构 + 水印文本）
- [`verify_watermark_and_protect.py`](../verify_watermark_and_protect.py) — 验证 `watermark_and_protect_ppt.rs` 输出（解密后检查水印）

## 12. 常用链接

- [OOXML 标准](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/)
- [DrawingML 元素](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/)
- [python-pptx 文档](https://python-pptx.readthedocs.io/)
- [zip crate](https://docs.rs/zip/)
- [quick-xml 文档](https://docs.rs/quick-xml/)
- [Rust API Guidelines](https://rust-lang.github.io/api-guidelines/)
