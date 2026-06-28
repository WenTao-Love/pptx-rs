# 架构总览

> 本文件描述 pptx-rs 项目的**完整架构**：模块职责、调用链、共享状态、扩展点、已知约束。
> 速查版见 [.trae/skills/pptx-rs-overview/SKILL.md](../.trae/skills/pptx-rs-overview/SKILL.md) 与 [.trae/skills/pptx-rs-architecture/SKILL.md](../.trae/skills/pptx-rs-architecture/SKILL.md)。

## 1. 设计目标

1. **正确性优先**：输出可被 PowerPoint / WPS / LibreOffice 三家打开，且 round-trip 后内容不丢。
2. **API 友好**：高层 API 与 python-pptx 对齐；低层 OOXML 模型可独立使用。
3. **零 unsafe**：库路径不出现 `unsafe`（除 `zip` crate 内部）。
4. **零异步**：纯同步实现，避免引入 tokio / async-std。
5. **依赖极简**：`zip` + `quick-xml` + `thiserror`（+ `aes` / `cbc` / `cfb` / `hmac` / `sha2` / `rand` / `base64` 用于加密功能）。

## 2. 三层架构

```
┌──────────────────────────────────────────────────────────────┐
│ Layer 3: 高阶 API（src/presentation.rs、slide.rs、shape/）   │
│ ── 用户直接调用的接口                                         │
│ ── 内部借用 oxml 模型，必要时 clone 一份推入 spTree           │
└──────────────────────────────────────────────────────────────┘
                          ▲ 调用
┌──────────────────────────────────────────────────────────────┐
│ Layer 2: OOXML 模型层（src/oxml/）                            │
│ ── 强类型 Rust 结构 + write_xml(&mut XmlWriter)              │
│ ── 不依赖 OPC、不依赖 shape                                  │
└──────────────────────────────────────────────────────────────┘
                          ▲ 序列化
┌──────────────────────────────────────────────────────────────┐
│ Layer 1: OPC 容器层（src/opc/）                               │
│ ── OpcPackage / Part / Relationships / ContentTypes          │
│ ── 加载/保存 zip；不解析任何业务 XML                          │
└──────────────────────────────────────────────────────────────┘
                          ▲ zip
                       zip crate
```

**核心原则**：
- **下层不依赖上层**（`opc` 不引 `oxml`，`oxml` 不引 `shape`）。
- **同层不互相耦合**（`opc::package` 仅用 `opc::part` / `opc::rels` / `opc::content_types`）。
- **跨层用 trait**（如未来 `EmuExt` 扩展到 OPC）。

## 3. 模块依赖图

```
                lib.rs
                   │
        ┌──────────┼──────────┬──────────┐
        │          │          │          │
   presentation  slide    slide_layouts  slide_masters
        │          │          │          │
        └──────────┼──────────┴──────────┘
                   │
                  shape/   oxml/   opc/
                   │         │       │
                   └────┬────┘       │
                        │            │
                      oxml/ ──────► opc/
                        │
                       zip
```

**禁止方向**：
- `opc → oxml`（opc 不知道有 sp/pic/...）
- `opc → shape` / `opc → presentation`
- `oxml → shape` / `oxml → presentation`
- `shape → presentation`（高阶层只借用 slide 字段）

## 4. 关键类型

| 类型 | 路径 | 作用 |
| --- | --- | --- |
| `Presentation` | `presentation.rs` | 顶层容器 |
| `Slide` / `Slides` / `Shapes` / `ShapesMut` | `slide.rs` | 幻灯片与形状视图 |
| `OpcPackage` | `opc/package.rs` | 包的装卸 |
| `Part` / `PartName` | `opc/part.rs` | 单个 part |
| `Relationship` / `Relationships` / `RelType` | `opc/rels.rs` | 关系 |
| `ContentTypes` | `opc/content_types.rs` | Content-Types 模型 |
| `Sld` / `SldMaster` / `SldLayout` / `PresentationRoot` | `oxml/{slide,slidemaster,slidelayout,presentation}.rs` | 顶层 XML 模型 |
| `Sp` / `Pic` / `Group` / `Connector` / `GraphicFrame` | `oxml/shape.rs` | 形状 XML 模型 |
| `ShapeProperties` / `Transform` / `Fill` / `Line` | `oxml/sppr.rs` | 形状属性 |
| `TextBody` / `Paragraph` / `Run` / `RunProperties` | `oxml/txbody.rs` | 文本 |
| `Table` / `Row` / `Col` / `Cell` | `oxml/table.rs` | 表格 |
| `Color` / `SchemeColor` / `PresetColor` | `oxml/color.rs` | 颜色 |
| `PresetGeometry` / `Alignment` / `Underline` / ... | `oxml/simpletypes.rs` | 枚举型属性 |
| `XmlWriter` | `oxml/writer.rs` | XML 写出器 |
| `Shape`（trait） / `ShapeKind` | `shape/{base,mod}.rs` | 形状抽象 |
| `AutoShape` / `TextBox` / `Picture` / `Group` / `Connector` / `TableShape` / `Freeform` | `shape/*.rs` | 高阶形状 |
| `ModifyProtection` | `crypto.rs` | 文档保护标记（modifyVerifier） |
| `Error::Encryption` | `error.rs` | 加密/解密错误 |
| `Error::IndexOutOfRange` | `error.rs` | 索引越界错误 |
| `Error::Ppt97` | `error.rs` | .ppt 二进制格式相关错误 |
| `ppt97::WatermarkConfig` | `ppt97/watermark.rs` | .ppt 水印配置（文本/字号/颜色/旋转） |
| `ppt97::add_watermark` / `encrypt` / `add_watermark_and_encrypt` | `ppt97/mod.rs` | .ppt 公共 API（水印/加密/合并） |
| `ppt97::record::RT_*` 常量 | `ppt97/record.rs` | MS-PPT record 类型码（MainMaster / PPDrawing / UserEditAtom 等） |
| `ppt97::ole::write_stream` / `fix_mini_fat` | `ppt97/ole.rs` | OLE2/CFB 容器操作 |
| `ppt97::crypto::encrypt_ppt_stream` | `ppt97/crypto.rs` | RC4 CryptoAPI 加密核心流程 |

## 5. 共享状态

| 状态 | 类型 | 所有者 | 备注 |
| --- | --- | --- | --- |
| 形状 ID 计数器 | `Rc<Cell<u32>>` | `Presentation` 创建 → 共享给每个 `Slide` | `Slide::next_shape_id()` |
| 媒体（图片 blob） | `Vec<MediaEntry>` | `Presentation` | 保存时遍历写入 zip |
| 母版 | `SlideMasters` | `Presentation` | 0.1.0 强制 1 个 |
| 版式 | `SlideLayouts` | `Presentation` | 0.1.0 强制 1 个 |
| 主题 | 静态字符串 | 全局常量 | `default_theme_xml()` |

**线程安全**：0.1.0 不暴露多线程 API；`Rc` / `RefCell` / `Cell` 均为单线程安全。

## 6. 写入流程（保存 .pptx）

```
Presentation::save(path)
    └─► Presentation::to_opc_package(&self)
            ├─ 构造 OpcPackage::new()
            ├─ 写入 /_rels/.rels（根关系：officeDocument + coreProps + appProps）
            ├─ 写入 /docProps/core.xml + /docProps/app.xml
            ├─ 写入 /ppt/theme/theme1.xml（default_theme_xml 静态字符串）
            ├─ 写入 /ppt/slideMasters/slideMaster1.xml + .rels
            ├─ 写入 /ppt/slideLayouts/slideLayout1.xml + .rels
            ├─ 写入 /ppt/presProps.xml + viewProps.xml + tableStyles.xml
            ├─ for each slide:
            │     ├─ sld.set_layout_rid("rId1")
            │     ├─ 构造 /ppt/slides/_rels/slideN.xml.rels（layout + media）
            │     ├─ 写入 /ppt/slides/slideN.xml
            │     └─ 在 pres_rels 加 <Relationship Type=slide Target=slides/slideN.xml>
            ├─ 写入 /ppt/presentation.xml
            ├─ 写入 /ppt/_rels/presentation.xml.rels
            └─ 写入 /ppt/media/*（图片）
    └─► OpcPackage::save(path) / to_bytes()
            └─ zip::ZipWriter::new + start_file + write_all
```

## 7. 读取流程（打开 .pptx）

```
Presentation::open(path)
    └─► Presentation::load(path)
            └─► OpcPackage::load(path)
                    ├─ 打开 zip
                    ├─ 读 [Content_Types].xml → ContentTypes
                    └─ 遍历 zip：每个 entry → Part
            └─► Presentation::from_opc(pkg)
                    ├─ ⚠️ 0.1.0 简化：仅构造 1 个 master + 1 个 layout
                    └─ 返回 Presentation（slides 为空）
```

**已知限制**：`from_opc` 不解析各 slide 内容；路线图补全。

## 8. 添加形状的流程

```rust
// 1) 用户
let mut slide = prs.slides_mut().add_slide(counter)?;
let mut tb = slide.shapes_mut()
    .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(8.0), Inches(1.0), "hi")?;

// 2) 内部
//   ShapesMut::add_textbox_with_text
//     ├─ TextBox::new("TextBox 1")  // AutoShape::new + PresetGeometry::Rectangle
//     ├─ set_left / set_top / set_width / set_height  // 写到 sp.properties.xfrm
//     ├─ set_id(slide.next_shape_id())  // 从 Rc<Cell<u32>> 自增
//     ├─ set_text("hi")  // 清空 paragraphs + push Paragraph{ runs: [Run("hi")] }
//     ├─ clone tb.shape.sp  // ⚠️ 当前设计：clone 后两份
//     └─ slide.inner.shapes.push(OxmlSlideShape::Sp(sp))
```

## 9. 序列化约定

### 9.1 命名空间声明

- 根元素自带 `xmlns:a` / `xmlns:p` / `xmlns:r` / `xmlns:p14`。
- 子元素不重复声明（除 `<a:blipFill xmlns:r=...>` 这种局部分支）。

### 9.2 属性顺序

- 严格按 OOXML schema 顺序。
- 如 `<a:rPr>` 属性顺序：`lang → altLang → sz → b → i → u → ...`。
- 如 `<p:sp>` 子元素顺序：`p:nvSpPr → p:spPr → p:txBody → p:extLst`。

### 9.3 借用延长

```rust
// ❌ 错误：Vec<&str> 在 w.close() 后借用失效
let attrs: Vec<(&str, &str)> = self.xfrm.off_x.map(|v| ...).collect();
w.open_with("...", &attrs);  // 报错

// ✅ 正确：先把所有 String 提到 w.close() 之外
let off_x_s = self.xfrm.off_x.map(|v| v.value().to_string());
let off_y_s = self.xfrm.off_y.map(|v| v.value().to_string());
w.open("a:xfrm");
w.empty_with("a:off", &[("x", &off_x_s), ("y", &off_y_s)]);
w.close("a:xfrm");
```

### 9.4 写出器用法

```rust
use crate::oxml::writer::XmlWriter;
let mut w = XmlWriter::with_decl();  // 自动写 <?xml ...?>
w.open("a:rPr");
w.empty_with("a:latin", &[("typeface", "Calibri")]);
w.close("a:rPr");
let xml = w.into_string();
```

## 10. 错误传播

```rust
use crate::error::{Error, Result};

fn parse_partname(s: &str) -> Result<PartName> {
    if !s.starts_with('/') {
        return Err(Error::opc(format!("part name must start with '/', got: {s}")));
    }
    // ...
}
```

## 11. 扩展点

| 想加什么 | 看哪里 |
| --- | --- |
| 新形状（preset） | `oxml/simpletypes.rs::PresetGeometry` + 写出器 |
| 新 Run 属性 | `oxml/txbody.rs::RunProperties` + `write_xml` |
| 新颜色 | `oxml/color.rs` |
| 新图表 | `shape/mod.rs::ShapeKind` + `oxml/shape.rs::Graphic` |
| 新关系类型 | `opc/rels.rs::RelType` + uri 映射 |
| 新 Content-Type | `opc/package.rs::ct::*` |
| 新高阶形状 | `shape/<name>.rs` + `shape/mod.rs` + `slide.rs::ShapesMut` |
| .pptx 文件加密 | `examples/protect_pptx.rs` / `examples/watermark_and_protect.rs`（OOXML Agile Encryption） |
| .ppt 文件加密 | [`src/ppt97/crypto.rs`](../src/ppt97/crypto.rs)（RC4 CryptoAPI，OLE2/CFB + 二进制 record 树）+ [`src/ppt97/mod.rs::encrypt`](../src/ppt97/mod.rs) 公共 API |
| .ppt 文件水印 | [`src/ppt97/watermark.rs`](../src/ppt97/watermark.rs)（Escher OfficeArt SpContainer 注入）+ [`src/ppt97/mod.rs::add_watermark`](../src/ppt97/mod.rs) 公共 API |
| .ppt 水印+加密 | [`src/ppt97/mod.rs::add_watermark_and_encrypt`](../src/ppt97/mod.rs)（先水印后加密，含 offset 更新） |
| .ppt record 解析扩展 | [`src/ppt97/record.rs`](../src/ppt97/record.rs)（新增 record 类型常量 / 解析函数） |
| .ppt OLE2 容器操作扩展 | [`src/ppt97/ole.rs`](../src/ppt97/ole.rs)（新增 stream 操作 / FAT 修复函数） |

详见 [.trae/skills/pptx-rs-extending/SKILL.md](../.trae/skills/pptx-rs-extending/SKILL.md)。

## 12. 反模式 / 陷阱

| 反模式 | 后果 |
| --- | --- |
| 在 `opc` 用 `oxml` 的类型 | 跨层耦合 |
| 写 `format!("<a:sp>{}</a:sp>", ...)` | 转义错误 / 性能 |
| `panic!` 库路径 | panic 抛给用户 |
| `unsafe` 块 | UB 风险 |
| 共享 `Rc<RefCell<T>>` 跨线程 | 编译期错 |
| 形状 ID 不递增 | 同 slide 内 id 冲突 |
| 关系文件路径写绝对 | 找不到 part |
| `<p:spPr>` 子元素顺序错 | 强校验失败 |

## 13. .ppt 文件格式支持（PowerPoint 97-2003）

`.ppt` 文件与 `.pptx` 完全不同，使用 OLE2/CFB（Compound File Binary）容器 + 二进制 record 树结构，而非 ZIP+XML。
本库通过 [`src/ppt97/`](../src/ppt97/) 模块提供完整支持，examples 仅作为薄封装的命令行入口。

### 13.1 模块组织

```
src/ppt97/
├── mod.rs          # 模块入口 + 3 个公共 API + OLE2 stream 读写辅助
├── record.rs       # PPT record 树解析（MS-PPT 规范）
│                   #   - record header / PersistDirectoryAtom 解析
│                   #   - find_main_masters / find_ppdrawing_in_master
├── ole.rs          # OLE2/CFB 容器操作
│                   #   - write_stream: 写入/替换 stream
│                   #   - fix_mini_fat: 修复 cfb crate 多分配的 mini FAT 扇区
├── watermark.rs    # 水印注入（Escher OfficeArt SpContainer）
│                   #   - inject_watermark: 核心注入逻辑
│                   #   - WatermarkConfig: 水印配置
└── crypto.rs       # RC4 CryptoAPI 加密（MS-OFFCRYPTO 规范）
                    #   - Rc4: RC4 流密码实现
                    #   - make_key / encrypt_persist_object: 密钥派生与加密
                    #   - build_crypt_session10_container: 加密参数容器
                    #   - reorder_persist_objects / encrypt_ppt_stream: 加密主流程
```

**模块依赖**：`mod.rs` → `record` / `ole` / `watermark` / `crypto`（四个子模块互不依赖）。

### 13.2 文件结构

```
.ppt 文件（OLE2/CFB 容器）
├── Current User stream          # 用户信息 + offsetToCurrentEdit
├── PowerPoint Document stream   # 主内容（record 树）
│   ├── MainMaster records (0x03F8)  # 母版（水印注入目标）
│   │   └── PPDrawing (0x040C)       # 绘图层（Escher OfficeArt）
│   │       └── DgContainer (0xF002)
│   │           └── SpgrContainer (0xF003)
│   │               ├── SpContainer (0xF004)  # 组形状本身
│   │               └── SpContainer (0xF004)  # 水印（z-order 最低的真正子形状）
│   │                   ├── FSP (0xF007)
│   │                   ├── FOPT (0xF008)         # 保护位 0x01C2
│   │                   ├── ClientAnchor (0xF00B) # 全屏 0,0 → 5760,4320
│   │                   └── ClientTextbox (0xF009)
│   ├── PersistDirectoryAtom (0x1772)  # persistId → offset 映射
│   └── UserEditAtom (0x0FF5)          # 编辑信息
└── Pictures stream              # 图片数据（不加密）
```

### 13.3 公共 API 调用链

#### 仅加水印 `ppt97::add_watermark(path, &config)`

```
add_watermark
    ├─► CompoundFile::open(Cursor::new(file_data))   // 打开 OLE2 容器
    ├─► read_all_streams(&mut comp)                  // 读取所有 streams
    ├─► classify_streams(streams)                    // 分类 CU / PPT / 其他
    ├─► parse_record_header(&cu_data, 0)             // 解析 CurrentUserAtom
    ├─► read_u32_le(&cu_data, 16)                    // 读取 offsetToCurrentEdit
    ├─► parse_record_header(&ppt_data, ue_offset)    // 解析 UserEditAtom
    ├─► read_u32_le(&ppt_data, ue_offset + 20)       // 读取 offsetPersistDirectory
    ├─► parse_persist_directory(&ppt_data, pd_offset) // 解析 PersistDirectoryAtom
    ├─► inject_watermark(&mut ppt_data, config, ...) // 注入水印（更新所有 offset）
    ├─► write_u32_le(&mut ppt_data, ue_offset_new + 20, pd_offset_new)  // 更新 UserEditAtom
    ├─► write_u32_le(&mut cu_data, 16, ue_offset_new)                   // 更新 CurrentUser
    └─► write_back_streams(comp, &cu_data, &ppt_data, &other_streams)   // 写回 + fix_mini_fat
```

#### 仅加密 `ppt97::encrypt(path, password)`

```
encrypt
    ├─► CompoundFile::open(...)
    ├─► read_all_streams / classify_streams
    └─► encrypt_ppt_stream(&mut ppt_data, &mut cu_data, password)
            ├─► 检查 UserEditAtom.recLen == 28（确认未加密）
            ├─► 解析 PersistDirectoryAtom
            ├─► reorder_persist_objects(...)         // 重排 persist 对象
            ├─► build_crypt_session10_container(...) // 构造加密参数容器
            ├─► 对每个 persist 对象调用 encrypt_persist_object(...)
            ├─► 更新 UserEditAtom（recLen=32, encryptSessionPersistIdRef）
            └─► 更新 CurrentUserAtom（headerToken=0xF3D1C4DF）
```

#### 水印 + 加密 `ppt97::add_watermark_and_encrypt(path, &config, password)`

```
add_watermark_and_encrypt
    ├─► 1) 执行与 add_watermark 相同的水印注入流程
    │       （更新 persist entries / UserEditAtom / CurrentUser）
    └─► 2) 调用 encrypt_ppt_stream(&mut ppt_data, &mut cu_data, password)
            （对已加水印的 ppt_data 执行加密）
```

**顺序约束**：必须**先水印后加密**。加密后所有 persist 对象被加密，无法再修改 record 结构。

### 13.4 加密机制（RC4 CryptoAPI）

`.ppt` 使用 RC4 CryptoAPI 加密（与 `.pptx` 的 AES Agile Encryption 完全不同）：

- **密钥派生**：H₀ = SHA1(salt + password_utf16le)，Hfinal = SHA1(H₀ + LE32(block))
- **加密粒度**：每个 persist 对象独立加密，block=persistId
- **不加密**：UserEditAtom、PersistDirectoryAtom、Pictures stream（WPS 兼容）
- **加密标记**：
  - `CurrentUserAtom.headerToken`：0xE391C05F（未加密）→ 0xF3D1C4DF（已加密）
  - `UserEditAtom.recLen`：28 → 32（添加 encryptSessionPersistIdRef）
  - `PersistDirectoryAtom`：添加 CryptSession10Container 条目

### 13.5 水印注入设计

水印通过在 **MainMaster** 的 PPDrawing 中注入 SpContainer（不可编辑背景层）实现：

1. 遍历所有 MainMaster record (0x03F8)（注：不是 Slide，水印要覆盖所有幻灯片）
2. 在每个 MainMaster 的 PPDrawing 中找到 SpgrContainer (0xF003)
3. 在 SpgrContainer 中"组形状本身"之后插入水印 SpContainer（**z-order 最低的真正子形状**）
4. 更新 SpgrContainer / PPDrawing / MainMaster 的 recLen
5. 更新 PersistDirectoryAtom 中所有 persist 对象的 offset（加水印导致 offset 偏移）

**z-order 设计要点**：
- 水印必须是 SpgrContainer 中"组形状本身"之后的**第一个子形状**（不是末尾）
- 末尾的子形状 z-order 最高，会被显示为可编辑文本框
- 第一个子形状 z-order 最低，作为背景层普通视图下不可选中/编辑
- FOPT 保护位 0x01C2 锁定（MS-ODRAW 规范 bit field）

### 13.6 PersistDirectoryAtom 解析

PersistDirectoryAtom (0x1772) 是 .ppt 文件"对象寻址"的核心数据结构，包含多个 PersistDirectoryEntry：

```
rgPersistDirEntry[]:
  PersistDirectoryEntry:
    - persistId (20 bits) | cPersist (12 bits) = 4 bytes
    - rgPersistOffset[cPersist] (cPersist * 4 bytes): 每个 entry 的 stream offset
```

**注意**：一个 PersistDirectoryAtom 可包含**多个** PersistDirectoryEntry，每个 entry 有自己的 persistId 起始值和 cPersist。
persistId 在 entry 内从 `persistId` 开始递增（`persistId + 0`、`persistId + 1`、...、`persistId + cPersist - 1`）。

实现见 [`src/ppt97/record.rs::parse_persist_directory`](../src/ppt97/record.rs)。

### 13.7 OLE2 容器修复

`cfb` crate 在 `create_stream` 时可能多分配 mini FAT 扇区，PowerPoint / WPS 严格检查 mini FAT 结构会拒绝打开。
[`src/ppt97/ole.rs::fix_mini_fat`](../src/ppt97/ole.rs) 在写回容器后修复此问题：

1. 沿 FAT 链遍历 mini FAT 扇区
2. 检查第二个及之后的扇区是否都是 FREESECT
3. 若是，将 FAT 中第一个 mini FAT 扇区的 next 改为 ENDOFCHAIN，其余标记为 FREESECT
4. 修改 OLE2 header 中的 `num_mini_fat_sectors` 为实际使用的扇区数

### 13.8 相关 example 与库 API 对应

| example | 库 API | 功能 |
| --- | --- | --- |
| `protect_ppt` | `ppt97::encrypt` | .ppt RC4 CryptoAPI 加密 |
| `watermark_ppt` | `ppt97::add_watermark` | .ppt 水印注入 |
| `watermark_and_protect_ppt` | `ppt97::add_watermark_and_encrypt` | .ppt 水印+加密合并 |

examples 从 1000+ 行的独立实现简化为 60~80 行的薄封装，全部业务逻辑位于 [`src/ppt97/`](../src/ppt97/)。

## 14. 未来路线图

- v0.3：完整读取路径（`from_opc` 解析各 slide 内容）
- v0.4：Chart（基础类型）
- v0.5：SmartArt
- v0.6：自定义几何（custGeom）
- v0.7：性能优化（流式写）
- v1.0：API 稳定
