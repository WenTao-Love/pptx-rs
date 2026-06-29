# pptx-rs

Rust 实现的 PowerPoint `.pptx` / `.ppt` 读写库，对标 [python-pptx](https://github.com/scanny/python-pptx)。

> 当前状态：**0.3.0** —— MVP + 文件加密。可创建 / 打开 / 保存 `.pptx`；支持幻灯片增删、文本框、段落、Run（带字体/字号/粗体/颜色/下划线）、图片插入、表格/连接器等基础形状；**支持 OOXML Agile Encryption 文件加密**（AES-256-CBC + SHA512，WPS/PowerPoint 兼容）；支持水印注入。**支持 .ppt（PowerPoint 97-2003 二进制格式）文件的 RC4 CryptoAPI 加密和水印注入**。完整的 Master / Layout / Theme / Chart / SmartArt 还在路线图上。
>
> **已知限制**：打开已有 `.pptx` 再保存时，母版/版式/主题定制不会完整还原（使用默认主题）。建议用于创建新文件或修改 slide 内容，暂不建议用于完整 round-trip 保真场景。

## 路线图

- [x] OPC 包容器：zip 读写、`[Content_Types].xml`、Part 关系链
- [x] Presentation / Slide / SlideLayout / SlideMaster XML 模型
- [x] 形状树：AutoShape、Picture、Group、Connector
- [x] TextFrame / Paragraph / Run / 字体/颜色
- [x] 常用 DrawingML 几何（`<a:prstGeom>` preset shapes）
- [x] 表格
- [x] 水印（向 spTree 注入半透明文本框）
- [x] **文件加密（OOXML Agile Encryption：AES-256-CBC + SHA512）**
- [x] **水印+加密一步到位**
- [x] **.ppt 文件加密（RC4 CryptoAPI，MS-OFFCRYPTO 规范）**
- [x] **.ppt 文件水印注入（Escher OfficeArt SpContainer）**
- [x] **.ppt 水印+加密合并**
- [ ] 完整母版 / 版式 / 主题
- [ ] Chart（基础类型）
- [ ] SmartArt

## 快速开始

```rust
use pptx_rs::{Presentation, Pt, RGBColor};

fn main() -> pptx_rs::Result<()> {
    // 1) 全新创建
    let mut prs = Presentation::new()?;
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide(counter)?;
    slide.shapes_mut().add_textbox(
        pptx_rs::Inches(1.0), pptx_rs::Inches(1.0),
        pptx_rs::Inches(8.0), pptx_rs::Inches(1.0),
    )?
    .text_frame_mut()
    .set_text("Hello, rust-pptx!")
    .paragraphs_mut()
    .first_mut()
    .unwrap()
    .runs_mut()
    .first_mut()
    .unwrap()
    .font_mut()
    .set_size(Pt(40.0))
    .set_bold(true)
    .set_color(RGBColor(0x1F, 0x4E, 0x79));

    prs.save("hello.pptx")?;

    // 2) 打开已存在
    let prs = Presentation::open("hello.pptx")?;
    for (i, slide) in prs.slides().iter().enumerate() {
        println!("slide #{} has {} shape(s)", i, slide.shapes().len());
    }
    Ok(())
}
```

## .ppt 97-2003 格式支持（库 API）

`.ppt`（PowerPoint 97-2003 二进制格式）与 `.pptx`（ZIP+XML）是**完全不同**的两种格式。
本库通过 `pptx_rs::ppt97` 模块提供 `.ppt` 文件的**水印注入**与 **RC4 CryptoAPI 加密**能力，
填补 python-pptx 不支持 .ppt 二进制格式的空白。

### 模块组织

```
src/ppt97/
├── mod.rs          # 模块入口 + 公共 API
├── record.rs       # PPT record 树解析（MS-PPT 规范）
├── ole.rs          # OLE2/CFB 容器操作（基于 cfb crate）
├── watermark.rs    # 水印注入（Escher OfficeArt SpContainer）
└── crypto.rs       # RC4 CryptoAPI 加密（MS-OFFCRYPTO 规范）
```

### 公共 API

| API | 功能 |
| --- | --- |
| `pptx_rs::ppt97::add_watermark(path, &config)` | 为 .ppt 注入水印（不可编辑背景层） |
| `pptx_rs::ppt97::encrypt(path, password)` | 为 .ppt 设置 RC4 CryptoAPI 加密 |
| `pptx_rs::ppt97::add_watermark_and_encrypt(path, &config, password)` | 同时注入水印和加密（先水印后加密） |
| `pptx_rs::ppt97::WatermarkConfig` | 水印配置（文本、字号、颜色、旋转角度） |

### 库 API 示例

```rust
use pptx_rs::ppt97::{add_watermark, encrypt, add_watermark_and_encrypt, WatermarkConfig};
use std::path::Path;

fn main() -> pptx_rs::Result<()> {
    let input = Path::new("input.ppt");
    let config = WatermarkConfig {
        text: "机密".to_string(),
        ..Default::default()
    };

    // 1) 仅加水印
    let watermarked = add_watermark(input, &config)?;
    std::fs::write("watermarked.ppt", &watermarked)?;

    // 2) 仅加密
    let encrypted = encrypt(input, "my-password")?;
    std::fs::write("encrypted.ppt", &encrypted)?;

    // 3) 水印 + 加密（推荐用法）
    let both = add_watermark_and_encrypt(input, &config, "my-password")?;
    std::fs::write("both.ppt", &both)?;

    Ok(())
}
```

### 水印特性

- 注入到 **MainMaster** 的 PPDrawing（覆盖所有幻灯片，不只是单张）
- 作为 SpgrContainer 中"组形状本身"之后的**第一个子形状**（z-order 最低）
- FOPT 保护位 0x01C2 锁定，普通视图下**不可选中/不可编辑**
- 全屏覆盖（ClientAnchor 0,0 → 5760,4320）
- 默认中灰色、可配置旋转角度、可配置文本/字号/颜色

### 加密特性

- RC4 CryptoAPI 加密（MS-OFFCRYPTO 规范 2.3.5）
- 密钥派生：H₀=SHA1(salt+password_utf16le)，Hfinal=SHA1(H₀+LE32(block))
- 每个 persist 对象独立加密（block=persistId，分段加密）
- **不加密**：UserEditAtom、PersistDirectoryAtom、Pictures stream（WPS 兼容）
- 加密标记：CurrentUserAtom.headerToken=0xF3D1C4DF，UserEditAtom.recLen=28→32
- 默认密码：`pptx-rs-secret`（可在调用时自定义）

## 文件加密

### .pptx 文件加密（OOXML Agile Encryption）

使用 OOXML Agile Encryption（MS-OFFCRYPTO 规范），输出可被 WPS / PowerPoint / msoffcrypto-python 正确解密：

```bash
# 仅加密
cargo run --example protect_pptx

# 水印+加密
cargo run --example watermark_and_protect
```

加密参数：
- 算法：AES-256-CBC + SHA512
- 密钥派生：PBKDF2（100000 次迭代）
- 容器：OLE2/CFB（DataSpaces + EncryptionInfo + EncryptedPackage）
- 默认密码：`pptx-rs-secret`

### .ppt 文件加密（RC4 CryptoAPI）

支持 PowerPoint 97-2003 二进制格式（`.ppt`）的 RC4 CryptoAPI 加密，输出可被 msoffcrypto-python 正确解密。
**推荐使用 [`pptx_rs::ppt97`](#ppt-97-2003-格式支持库-api) 库 API**，examples 仅作为命令行封装：

```bash
# 仅加密（库 API: pptx_rs::ppt97::encrypt）
cargo run --example protect_ppt

# 水印+加密（库 API: pptx_rs::ppt97::add_watermark_and_encrypt）
cargo run --example watermark_and_protect_ppt
```

加密参数：
- 算法：RC4 + SHA1 密钥派生（MS-OFFCRYPTO 规范）
- 密钥位数：128 bit
- 容器：OLE2/CFB（PowerPoint Document + Current User stream）
- 加密结构：CryptSession10Container + PersistDirectoryAtom 更新
- 每个 persist 对象独立加密（block=persistId，分段加密）
- 默认密码：`pptx-rs-secret`

## 水印

### .pptx 水印

```bash
cargo run --example watermark_pptx
```

水印特性：
- 浅灰色半透明 40pt 粗体文字 "pptx-rs WATERMARK"
- 旋转 -45°，位于幻灯片中央
- 不重新声明 xmlns（WPS 兼容）

### .ppt 水印

```bash
# 库 API: pptx_rs::ppt97::add_watermark
cargo run --example watermark_ppt
```

水印特性：
- 在每个 **MainMaster** 的 PPDrawing 中注入水印 SpContainer（不可编辑背景层）
- Escher OfficeArt 二进制结构（FSP / FOPT / ClientAnchor / ClientTextbox）
- TextCharsAtom（UTF-16LE）水印文本 "pptx-rs 水印"
- 加水印后正确更新所有 persist 对象的 offset
- 作为 SpgrContainer 中"组形状本身"之后的第一个子形状（z-order 最低）
- FOPT 保护位 0x01C2 锁定，普通视图下不可选中/不可编辑

## 与 python-pptx 的对应关系

| python-pptx | pptx-rs |
| --- | --- |
| `Presentation()` | `Presentation::new()` |
| `Presentation(path)` | `Presentation::open(path)` |
| `prs.save(path)` | `prs.save(path)` |
| `prs.slides` | `prs.slides()` / `prs.slides_mut()` |
| `prs.slide_width` / `slide_height` | `Emu` 类型字段 |
| `slide.shapes` | `slide.shapes()` / `slide.shapes_mut()` |
| `shapes.add_textbox(l, t, w, h)` | `shapes_mut().add_textbox(l, t, w, h)` |
| `shapes.add_picture(path, l, t, w, h)` | `shapes_mut().add_picture(path, l, t, w, h)` |
| `shape.text_frame.text = "…"` | `shape.text_frame_mut().set_text("…")` |
| `run.font.size = Pt(24)` | `run.font_mut().set_size(Pt(24.0))` |
| `run.font.bold = True` | `run.font_mut().set_bold(true)` |

## 命名

| 类型 | 说明 |
| --- | --- |
| `Emu(i64)` | English Metric Unit，1 inch = 914400 EMU |
| `Pt(f64)` | 磅，1 pt = 12700 EMU |
| `Inches(f64)` | 英寸 |
| `RGBColor(u8, u8, u8)` | sRGB 颜色 |

## 示例

| 示例 | 说明 |
| --- | --- |
| `hello_pptx` | 创建包含文本框的演示文稿 |
| `watermark_pptx` | 为现有 .pptx 添加水印 |
| `protect_pptx` | 为现有 .pptx 设置文件加密 |
| `watermark_and_protect` | 同时添加水印和加密（.pptx） |
| `protect_ppt` | 为现有 .ppt 设置 RC4 CryptoAPI 加密 |
| `watermark_ppt` | 为现有 .ppt 添加水印 |
| `watermark_and_protect_ppt` | 同时添加水印和加密（.ppt） |
| `text_format_demo` | 文本格式化演示 |
| `shapes_demo` | 各种形状演示 |
| `connector_demo` | 连接器演示 |

## 文档

- [架构总览](docs/ARCHITECTURE.md) — 三层模型、调用链、关键类型
- [开发指南](docs/DEVELOPMENT.md) — 环境、构建、调试、发布
- [测试规范](docs/TESTING.md) — 覆盖目标、断言策略
- [OOXML 参考](docs/OOXML_REFERENCE.md) — 命名空间、元素顺序
- [贡献指南](docs/CONTRIBUTING.md) — fork / PR / 评审
- [更新日志](docs/CHANGELOG.md) — 版本变更

## 许可

MIT
