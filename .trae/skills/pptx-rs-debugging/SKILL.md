---
name: "pptx-rs-debugging"
description: "pptx-rs 调试指南：解析失败、PowerPoint 拒绝打开、元素顺序错乱、水印/加密问题、运行时错误、性能问题。Invoke when user reports pptx generation failure, PowerPoint/WPS error, file corruption, watermark issue, encryption problem, or test failure."
---

# pptx-rs 调试指南

> 本指南按"症状 → 排查 → 修复"组织，融入 huiali/rust-skills 的 Decision Tree 方法论。

## Decision Tree：快速定位问题

```
1. PowerPoint 能打开文件吗？
   → 不能（提示损坏）：见 §1
   → 能但显示异常：见 §2
   → 能但内容不对：见 §3

2. 是哪种异常？
   → 中文乱码：见 §3
   → 图片不显示：见 §4
   → 水印不显示/不透明：见 §10
   → 加密无效：见 §11
   → 表格错位：见 §5

3. 是运行时错误吗？
   → panic / unwrap：见 §6
   → 编译错误：见 §9
   → cargo test 失败：见 §7
   → 性能问题：见 §8
```

## 1. 症状：PowerPoint 提示"无法打开 / 文件已损坏"

### 排查清单

1. **检查 [Content_Types].xml**

```powershell
# 1) 7-Zip 打开 hello.pptx
# 2) 看 [Content_Types].xml
# 3) 关键 part 是否有 Override？
```

- ✅ 每个 XML part 都有 `<Override PartName="..." ContentType="..."/>`。
- ❌ 缺 `/ppt/presentation.xml` / `/ppt/slides/slide1.xml` 的 Override。

**修复**：检查 [`src/presentation.rs::to_opc_package`](../../../../src/presentation.rs) 的 `put_part` 调用顺序。

2. **检查 `<p:presentation>` 子元素**

- ✅ `<p:sldMasterIdLst>` / `<p:sldIdLst>` / `<p:sldSz>` / `<p:notesSz>` / `<p:defaultTextStyle>` 全有。
- ❌ 缺 `defaultTextStyle` → PowerPoint 强校验失败。

**修复**：[`src/oxml/presentation.rs::PresentationRoot::default`](../../../../src/oxml/presentation.rs) 已注入 `DEFAULT_TEXT_STYLE`。

3. **检查 `<p:sldMaster>` 子元素**

- ✅ `<p:cSld>/<p:spTree>` + `<p:clrMap>` + `<p:sldLayoutIdLst>` + `<p:txStyles>`。
- ❌ 缺 `txStyles` → 强校验失败。

**修复**：见 [`src/oxml/slidemaster.rs::SldMaster::to_xml`](../../../../src/oxml/slidemaster.rs)，已静态注入 `TX_STYLES`。

4. **检查 `<p:sldLayout>` 属性**

- ✅ 根元素有 `type="blank"` 或 `type="title"` 等。
- ❌ 缺 `type=` → 校验失败。

**修复**：[`src/oxml/slidelayout.rs::SldLayout::to_xml`](../../../../src/oxml/slidelayout.rs) 已处理空 type 默认为 `blank`。

5. **检查 `<p:sp>` 子元素顺序**

正确顺序：`p:nvSpPr` → `p:spPr` → `p:txBody` → `p:extLst`。

如果子元素颠倒，PowerPoint 多数情况会容忍，但 WPS 严格。

**修复**：对照 [`src/oxml/shape.rs::Sp::write_xml`](../../../../src/oxml/shape.rs) 调整。

### 调试命令

```powershell
# 1) 重新生成 + 用 7-Zip 拆
cargo run --example hello_pptx
Expand-Archive hello.pptx -DestinationPath debug_out

# 2) 对照 python-pptx
python -c "from pptx import Presentation; p=Presentation('hello.pptx'); p.save('ref.pptx')"
Expand-Archive ref.pptx -DestinationPath ref_out

# 3) diff 关键文件
Compare-Object (Get-Content debug_out\ppt\presentation.xml) (Get-Content ref_out\ppt\presentation.xml)
```

## 2. 症状：PowerPoint 打开后显示"格式不正确"但能修复

### 可能原因

- **属性顺序错**（OOXML schema 强校验）。
- **多余属性**（如 LibreOffice 写的扩展属性）。
- **namespace 漏声明**（如子元素用了未声明的前缀）。

### 排查

```powershell
# 用 xmllint 验证
xmllint --noout debug_out\ppt\presentation.xml
# 或 PowerShell [xml]
[xml]$xml = Get-Content debug_out\ppt\presentation.xml
# 检查所有 xmlns
$xml.DocumentElement.NamespaceURI
```

## 3. 症状：中文乱码

### 排查

1. **XML 头声明**：

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
```

✅ 正确。❌ 漏 `encoding` 或写错（如 `utf8` 不带 `-`）。

2. **字符串来源**：是否中途用了 `String::from_utf8_lossy`（v0.1.0 不用）。

3. **运行环境**：cargo build 时的源码是否是 UTF-8（默认是）。

### 修复

- 确认 [`src/oxml/writer.rs::XmlWriter::decl`](../../../../src/oxml/writer.rs) 写出正确的 XML 头。
- 检查 `rel` 关系的 `Target` 是否被转义（中文 Target 用 `&` 转义）。

## 4. 症状：图片不显示 / 显示 X

### 排查

1. **媒体文件是否写入 zip**：

```powershell
Expand-Archive hello.pptx -DestinationPath debug_out
Test-Path debug_out\ppt\media\image1.png  # 应该是 True
```

2. **关系文件 `Target` 是否正确**：

```powershell
Get-Content debug_out\ppt\slides\_rels\slide1.xml.rels
# 应有 <Relationship Id="rIdImg1" Type="...image" Target="../media/image1.png"/>
```

3. **`<a:blip r:embed="rIdImg1">`** 中 rId 与 rels 中是否一致。

### 修复

- [`src/presentation.rs::to_opc_package`](../../../../src/presentation.rs) 中图片遍历 `self.media` 的 `rid` 必须以 `rIdImg` 开头才会被加入到 slide 关系（**注意**：当前实现细节，可能漏 media）。

## 5. 症状：表格行/列错位

### 排查

- **`<a:tblGrid>` 与 `<a:tr>` 中的 `<a:tc>` 数量**必须一致。
- **`<a:tr h="...">`** 中 h 决定行高。
- **`<a:gridCol w="...">`** 中 w 决定列宽。

### 修复

- [`src/oxml/table.rs::Table::write_xml`](../../../../src/oxml/table.rs) 中 `cols.len()` 与 `rows[].cells.len()` 必须一致。

## 6. 症状：运行时 `unwrap` panic

### 排查

- `Result<_, E>` 上有没有 `unwrap()`？
- 索引用 `vec[i]` 越界？
- 字符串转数字 `s.parse::<i32>().unwrap()` 失败？

### 修复

- 库路径必须用 `?` 传播 + `Error::*` 显式表达。
- 测试路径可以用 `unwrap()` / `expect()`。

## 7. 症状：`cargo test` 偶发失败

### 排查

- 是不是用了 `lazy_static` / `static mut`？
- 是不是测试间共享 `Rc<RefCell<T>>`？

### 修复

- 加 `--test-threads=1` 临时排查。
- 重构为不共享状态。

## 8. 症状：生成的 .pptx 体积过大

### 排查

- **是否多写了 theme？** → v0.1.0 写一次，无问题。
- **是否多写了 master/layout？** → 当前每张 slide 都引用同一个 layout，但 0.1.0 写一次。
- **是否 zip 压缩开了？** → `CompressionMethod::Deflated` 是默认。

### 修复

- 检查 `OpcPackage::save` 的 `compression_method`。
- 大型媒体考虑用 `jpeg`（已支持）。

## 9. 症状：cargo clippy 报错

### 常见 clippy 警告

| 警告 | 修复 |
| --- | --- |
| `needless_return` | 删 `return` |
| `redundant_clone` | 改借用 |
| `single_match` | 用 `if let` / `matches!` |
| `too_many_arguments` | 拆 builder |
| `missing_errors_doc` | 加 `# Errors` 段 |
| `missing_panics_doc` | 加 `# Panics` 段 |
| `cast_possible_truncation` | 用 `try_into` / `i32::try_from` |
| `cast_sign_loss` | 显式 `.unsigned_abs()` |

## 10. 症状：水印不显示 / 不透明 / 位置不对

### Decision Tree

```
1. 水印完全不可见？
   → 检查 Sp 是否被推入 slide.inner.shapes
   → 检查 Sp.name 是否为 "pptx-rs-watermark"

2. 水印可见但不透明？
   → 检查 RunProperties.alpha 是否设置
   → alpha 值范围：0-100000（不是 0-100）
   → 40000 = 40% 不透明，60000 = 60% 不透明

3. 水印位置不对？
   → 检查 xfrm.off_x / off_y 是否为 0,0
   → 检查 xfrm.ext_cx / ext_cy 是否覆盖全屏

4. 水印旋转不对？
   → OOXML 旋转单位是 1/60000 度
   → -30° 应写为 -30 * 60_000 = -1_800_000
   → 检查 Transform.rot 是否正确设置

5. WPS 能看但 PowerPoint 不能？
   → 检查 <a:alpha> 是否在 <a:srgbClr> 内部
   → 检查 write_solid_fill_with_alpha 的 XML 结构
```

### 排查步骤

```powershell
# 1) 生成水印文件
cargo run --example watermark_pptx

# 2) 解压查看 slide XML
Expand-Archive _test_out\wm_*.pptx -DestinationPath wm_debug
Get-Content wm_debug\ppt\slides\slide1.xml

# 3) 检查关键元素
# 应有：
# <a:solidFill>
#   <a:srgbClr val="C0C0C0">
#     <a:alpha val="40000"/>
#   </a:srgbClr>
# </a:solidFill>
```

### 常见错误

| 错误 | 表现 | 修复 |
| --- | --- | --- |
| `alpha` 在 `<a:solidFill>` 外 | PowerPoint 忽略 alpha | 确保 `<a:alpha>` 在 `<a:srgbClr>` 内部 |
| alpha 值为 40（应为 40000） | 几乎完全透明 | alpha 范围 0-100000 |
| 旋转写 30（应为 1800000） | 旋转角度极小 | 乘以 60_000 |
| 水印 Sp 缺少 `<a:noFill/>` | 形状有白色背景 | 设置 `fill = None` |

## 11. 症状：加密无效 / WPS 不识别

### Decision Tree

```
1. encrypt() 返回 NotImplemented？
   → v0.1.0 加密是占位 API，尚未实现
   → 路线图：v0.2+ 实现 ECMA-376 Agile Encryption

2. modifyVerifier 注入后 PowerPoint 提示损坏？
   → 检查 <p:modifyVerifier> 位置
   → 必须在 <p:extLst> 之前
   → WPS 对位置更敏感

3. 密码验证失败？
   → 检查 SHA-512 + salt + spinCount 算法
   → salt 必须是 base64 编码的随机字节
   → spinCount 必须为 100000

4. 加密后文件体积异常？
   → modifyVerifier 本身很小（~200 bytes）
   → 如果体积暴增，检查是否重复注入
```

### 排查步骤

```powershell
# 1) 运行保护示例
cargo run --example protect_pptx

# 2) 解压查看 presentation.xml
Expand-Archive _test_out\protected_*.pptx -DestinationPath prot_debug
Get-Content prot_debug\ppt\presentation.xml

# 3) 检查 <p:modifyVerifier> 位置
# 应在 <p:presentation> 内部，<p:extLst> 之前
```

### 常见错误

| 错误 | 表现 | 修复 |
| --- | --- | --- |
| `modifyVerifier` 在 `extLst` 之后 | WPS 提示损坏 | 移到 `extLst` 之前 |
| `cryptAlgorithmSid` 写错 | 算法不匹配 | SHA-512 = sid 14 |
| `salt` 不是 base64 | 解码失败 | 用 `base64::engine::encode` |
| `spinCount` 不是 100000 | 验证失败 | 必须为 100000 |

## 12. 调试辅助

### 打印生成的 XML

```rust
let bytes = prs.to_bytes()?;
let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&bytes))?;
let mut s = String::new();
zip.by_name("ppt/slides/slide1.xml")?.read_to_string(&mut s)?;
println!("{}", s);
```

### 打印所有 part 名

```rust
let pkg = prs.to_opc_package()?;
for part in pkg.iter_parts() {
    println!("{:>40}  {}  {:>8} B",
        part.partname.as_str(),
        part.content_type,
        part.len());
}
```

### 与 python-pptx 输出对比

```rust
// 用 std::process::Command 调 python：
// $ python -c "from pptx import Presentation; ..."
```

## 13. 提 issue / 调试信息

报告 bug 时附：

```powershell
cargo --version
rustc --version
cargo tree -p pptx-rs
cargo build --release 2>&1 | tee build.log
cargo test --all 2>&1 | tee test.log
# 关键 pptx 文件（如可）
```

## 14. 终极排查

如果以上都不行：

1. **生成两份对比**：用 python-pptx 生成同样内容的 .pptx，逐 part diff。
2. **用 LibreOffice 打开**：LibreOffice 容忍度比 PowerPoint 高，能打开说明文件结构基本对。
3. **OOXML 官方 schema 验证**：用 [Open XML SDK 2.5](https://github.com/OfficeDev/Open-XML-SDK) 的 `DocumentFormat.OpenXml.Validation` 验证（需 .NET）。
4. **联系维护者**：附最小复现代码 + 生成的 .pptx。

## 15. 已知坑（v0.1.0）

| 坑 | 状态 | 缓解 |
| --- | --- | --- |
| `from_opc` 仅建空壳 | 路线图 | 写测试用 `to_bytes` → `load_bytes` |
| 多个 master/layout 不支持 | 路线图 | 0.1.0 强制 1 master + 1 layout |
| 自定义几何（custGeom）不支持 | 路线图 | 用 `prstGeom` |
| Chart / SmartArt 不支持 | 路线图 | v0.2+ |
| `Presentation::open` 后内容不全 | 已知 | 读路径在迭代中 |
| media rid 必须 `rIdImg` 前缀 | 实现细节 | 保持不变即可 |
| `encrypt`/`decrypt` 未实现 | 占位 | 返回 `Error::NotImplemented` |

## Solution Patterns

### Pattern 1: OOXML 问题先拆 zip 再 diff

```rust
// ✅ 系统化排查
// 1) 生成 .pptx
// 2) 解压看 XML
// 3) 与 python-pptx 输出逐 part diff
// 4) 定位差异 → 修复 write_xml

// ❌ 随意猜测
// "可能是命名空间问题" → 改了还是不行
// "可能是编码问题" → 改了还是不行
```

**适用场景**：任何 PowerPoint/WPS 拒绝打开的情况。
**不适场景**：运行时 panic（应看堆栈）。

### Pattern 2: 水印问题先看 XML 结构

```rust
// ✅ 检查 XML 层级
// <a:solidFill>
//   <a:srgbClr val="C0C0C0">
//     <a:alpha val="40000"/>    ← 必须在 srgbClr 内
//   </a:srgbClr>
// </a:solidFill>

// ❌ 只检查 alpha 值
// "alpha = 40000 应该对的啊" → 但位置错了
```

**适用场景**：水印透明度/颜色/位置问题。
**不适场景**：水印完全不可见（应先检查 Sp 是否在 shapes 中）。

## Verification Commands

```bash
# 编译检查
cargo check

# 运行所有示例
cargo run --example hello_pptx
cargo run --example protect_pptx
cargo run --example watermark_pptx

# 运行测试
cargo test --all

# Clippy 检查
cargo clippy --all-targets -- -D warnings

# 检查 unwrap 使用
cargo clippy -- -W clippy::unwrap_used

# 检查 panic
cargo clippy -- -W clippy::panic

# 生成文档
cargo doc --no-deps
```

## Cross-References

- [pptx-rs-overview](../pptx-rs-overview/SKILL.md) — 项目总览
- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 架构详解（含水印/加密调用链）
- [pptx-rs-ooxml](../pptx-rs-ooxml/SKILL.md) — OOXML 速查（元素顺序/命名空间）
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南
- [pptx-rs-testing](../pptx-rs-testing/SKILL.md) — 测试规范
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
