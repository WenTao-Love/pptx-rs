---
name: "pptx-rs-ooxml"
description: "OOXML/DrawingML/PresentationML 关键元素的速查表：命名空间、必备元素、子元素顺序、单位换算。Invoke when user asks about OOXML structure, element ordering, namespace URIs, EMU/Pt conversion, or which elements are required."
---

# pptx-rs OOXML 速查

> 对应 [docs/OOXML_REFERENCE.md](../../../../docs/OOXML_REFERENCE.md)。

## 命名空间（URI → 前缀）

| 前缀 | URI | 用途 |
| --- | --- | --- |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML |
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | 关系引用 |
| `p14` | `http://schemas.microsoft.com/office/powerpoint/2010/main` | 2010 扩展 |
| — | `http://schemas.openxmlformats.org/package/2006/relationships` | 关系文件 |
| — | `http://schemas.openxmlformats.org/package/2006/content-types` | Content-Types |
| `cp` | `http://schemas.openxmlformats.org/package/2006/metadata/core-properties` | core.xml |
| `dc` | `http://purl.org/dc/elements/1.1/` | Dublin Core |

见 [`src/oxml/ns.rs`](../../../../src/oxml/ns.rs) 常量定义。

## 单位换算

| 物理量 | OOXML 内部单位 | 换算 |
| --- | --- | --- |
| 长度 | EMU (i64) | 1 inch = 914 400；1 cm = 360 000；1 pt = 12 700 |
| 字号 | 1/100 pt (i32) | Pt(24.0) → 2400 |
| 旋转 | 1/60000 度 (i32) | 30° → 1 800 000 |
| 缩进 / 间距 | EMU | 同长度 |
| 缩进 / 间距（百分比） | 1/1000 (i32) | 100% = 100 000；80% = 80 000 |
| 行距（百分比） | 1/1000 (i32) | 100% = 100 000；150% = 150 000 |
| 颜色 | sRGB / 主题色 / 预设色 | `<a:srgbClr val="RRGGBB"/>` 等 |

## 必备元素清单（PowerPoint 强校验）

### `<p:presentation>`（`/ppt/presentation.xml`）

```xml
<p:presentation xmlns:a="..." xmlns:p="..." xmlns:r="..." xmlns:p14="...">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>  <!-- 至少 1 个 -->
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId10"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>...</p:defaultTextStyle>  <!-- 必含 9 段 lvlXpPr -->
</p:presentation>
```

### `<p:sldMaster>`（`/ppt/slideMasters/slideMaster1.xml`）

```xml
<p:sldMaster xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>...</a:xfrm></p:grpSpPr>
      <!-- 用户 shapes -->
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" .../>  <!-- 必含 -->
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>  <!-- 至少 1 个 -->
  </p:sldLayoutIdLst>
  <p:txStyles>  <!-- 必含 titleStyle/bodyStyle/otherStyle -->
    <p:titleStyle>...</p:titleStyle>
    <p:bodyStyle>...</p:bodyStyle>
    <p:otherStyle>...</p:otherStyle>
  </p:txStyles>
</p:sldMaster>
```

### `<p:sldLayout>`（`/ppt/slideLayouts/slideLayout1.xml`）

```xml
<p:sldLayout xmlns:a="..." xmlns:p="..." xmlns:r="..."
             type="blank" preserve="1">  <!-- type 必含 -->
  <p:cSld name="blank">
    <p:spTree>
      <p:nvGrpSpPr>...</p:nvGrpSpPr>
      <p:grpSpPr>...</p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr bg1="lt1" tx1="dk1" .../>
  <p:transition/>
</p:sldLayout>
```

### `<p:sld>`（`/ppt/slides/slideN.xml`）

```xml
<p:sld xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm>...</a:xfrm></p:grpSpPr>
      <p:sp>...</p:sp>  <!-- 0..n -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
  <p:transition/>  <!-- 可选 -->
</p:sld>
```

### `<a:theme>`（`/ppt/theme/theme1.xml`）

必须含 `<a:themeElements>` → `<a:clrScheme>` + `<a:fontScheme>` + `<a:fmtScheme>`。静态完整版见 [`src/oxml/theme.rs`](../../../../src/oxml/theme.rs)。

## 形状子元素顺序（OOXML 严格规定）

### `<p:sp>` 子元素顺序

```
<p:nvSpPr>     ← 1
  <p:cNvPr/>
  <p:cNvSpPr/>
  <p:nvPr/>    ← 可选：<p:ph type=... idx=...>
</p:nvSpPr>
<p:spPr/>      ← 2
<p:txBody/>    ← 3 (或 <p:extLst/> 4)
```

### `<p:pic>` 子元素顺序

```
<p:nvPicPr>
  <p:cNvPr/>
  <p:cNvPicPr/>
  <p:nvPr/>
</p:nvPicPr>
<p:blipFill>   ← 必须在 spPr 之前
  <a:blip r:embed="rId..."/>
  <a:stretch><a:fillRect/></a:stretch>
</p:blipFill>
<p:spPr/>
```

### `<p:spPr>` 子元素顺序

```
<a:xfrm>          ← 1
<a:prstGeom>      ← 2 (或 <a:custGeom>)
<a:noFill/>       ← 3 (或 <a:solidFill>/<a:gradFill>/...)
<a:ln>            ← 4
<a:effectLst/>    ← 5
<a:scene3d>       ← 6
<a:sp3d>          ← 7
<a:extLst>        ← 8
```

### `<a:txBody>` 子元素顺序

```
<a:bodyPr/>       ← 1 (含 ln/reflection 等)
<a:lstStyle/>     ← 2
<a:p>             ← 3..n
  <a:pPr/>        ← 可选
  <a:r>           ← 1..n
    <a:rPr/>
    <a:t>...</a:t>
  </a:r>
  <a:endParaRPr/> ← 可选
</a:p>
<a:extLst/>       ← 最后
```

### `<a:rPr>` 属性顺序（OOXML 严格）

```
lang
altLang
sz
b
i
u
strike
baseline
cap
spc
normalize
kern
noProof
dirty
err
smtClean
smtId
...
```

之后是子元素（顺序）：

```
<a:ln>
<a:noFill> | <a:solidFill> | <a:gradFill> | <a:blipFill> | <a:pattFill> | <a:grpFill>
<a:effectLst> | <a:effectDag>
<a:highlight>
<a:uLnTx> | <a:uLn>
<a:uFillTx> | <a:uFill>
<a:latin>
<a:ea>
<a:cs>
<a:sym>
<a:hlinkClick>
<a:hlinkMouseOver>
<a:rtl>
<a:extLst>
```

## 关系（`.rels`）结构

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="../media/image1.png"/>
</Relationships>
```

| 文件 | 关系 | Target 相对路径基准 |
| --- | --- | --- |
| `/_rels/.rels` | 根：officeDocument + coreProps + appProps | 绝对 `/ppt/presentation.xml` |
| `/ppt/_rels/presentation.xml.rels` | presentation → slideMaster / slideLayout / theme / presProps / viewProps / tableStyles | 相对 `ppt/`，如 `slideMasters/slideMaster1.xml` |
| `/ppt/slideMasters/_rels/slideMaster1.xml.rels` | master → slideLayout / theme | 相对 `ppt/slideMasters/`，如 `../slideLayouts/slideLayout1.xml` |
| `/ppt/slideLayouts/_rels/slideLayout1.xml.rels` | layout → slideMaster | 相对 `ppt/slideLayouts/`，如 `../slideMasters/slideMaster1.xml` |
| `/ppt/slides/_rels/slideN.xml.rels` | slide → slideLayout / image | 相对 `ppt/slides/`，如 `../slideLayouts/slideLayout1.xml` |

## Content-Types 模板

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/ppt/presentation.xml"        ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml"       ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml"        ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/presProps.xml"           ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml"           ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml"         ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/docProps/core.xml"           ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
```

## 关键写值示例

### Run 属性（完整示例）

```xml
<a:r>
  <a:rPr
    lang="zh-CN"
    altLang="en-US"
    sz="2400" b="1" i="1" u="sng"
    strike="sngStrike"
    baseline="30000"
    cap="small"
    kern="1200"
    spc="-100"
    normalize="0"
    noProof="0"
    dirty="0">
    <a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
    <a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill>
    <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
    <a:latin typeface="+mn-lt"/>
    <a:ea   typeface="宋体"/>
    <a:cs   typeface="+mn-cs"/>
    <a:hlinkClick r:id="rId99"/>
  </a:rPr>
  <a:t>Hello, World</a:t>
</a:r>
```

### 形状属性（完整示例）

```xml
<p:spPr>
  <a:xfrm rot="1800000" flipH="1" flipV="0">
    <a:off x="914400" y="914400"/>
    <a:ext cx="4572000" cy="2743200"/>
  </a:xfrm>
  <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
  <a:ln w="19050" cap="flat" cmpd="sng">
    <a:solidFill><a:srgbClr val="2F5496"/></a:solidFill>
    <a:prstDash val="dash"/>
  </a:ln>
</p:spPr>
```

### 图片

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="3" name="Picture 1"/>
    <p:cNvPicPr/>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="..." r:embed="rIdImg1"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

### 表格

```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="10" name="Table 1"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="914400" y="1828800"/>
    <a:ext cx="7315200" cy="3657600"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/main">
      <a:tbl>
        <a:tblPr>
          <a:tblW w="0" type="auto"/>
          <a:tblGrid/>
          <a:tblLook val="04A0" firstRow="1" lastRow="0" firstColumn="1" lastColumn="0" noHBand="0" noVBand="1"/>
        </a:tblPr>
        <a:tblGrid>
          <a:gridCol w="2438400"/>
          <a:gridCol w="2438400"/>
        </a:tblGrid>
        <a:tr h="365760">
          <a:tc><a:txBody>...</a:txBody></a:tc>
          <a:tc><a:txBody>...</a:txBody></a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

## 文档保护

```xml
<p:modifyVerifier
  cryptProviderType="rsaAES"
  cryptAlgorithmClass="hash"
  cryptAlgorithmType="typeAny"
  cryptAlgorithmSid="14"
  cryptSpinCount="100000"
  hash="<base64>"
  salt="<base64>"/>
```

**位置**：`<p:presentation>` 内部、`<p:extLst>` 之前。
**算法**：SHA-512 + salt + spinCount=100 000。

## 主题（Theme）必含

```xml
<a:theme xmlns:a="..." name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window"      lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1>..6></a:accent1..6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <!-- 40+ 复杂脚本字体 -->
      </a:majorFont>
      <a:minorFont>...</a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>...</a:fillStyleLst>
      <a:lnStyleLst>...</a:lnStyleLst>
      <a:effectStyleLst>...</a:effectStyleLst>
      <a:bgFillStyleLst>...</a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
```

完整内容见 [`src/oxml/theme.rs`](../../../../src/oxml/theme.rs) 的 `THEME_XML` 常量。

## 常见错误

| 错误 | 表现 | 修复 |
| --- | --- | --- |
| `<p:sldMaster>` 缺 `<p:clrMap>` | PowerPoint 提示损坏 | 补全 |
| `<p:sldMaster>` 缺 `<p:txStyles>` | 强校验失败 | 补全 |
| `<p:sldLayout>` 缺 `type=` | 校验失败 | 补 `type="blank"` |
| `<p:presentation>` 缺 `<p:defaultTextStyle>` | 校验失败 | 补 9 段 |
| `<p:sp>` 子元素顺序错 | 强校验失败 | 调整 `write_xml` |
| `<a:txBody>` 缺 `<a:bodyPr>` | 校验失败 | 必含 |
| 关系文件路径写绝对 | 找不到 part | 写相对 `..` |
| 旋转 30° 写 `30` | 实际是 30/60000 度 | 写 `1 800 000` |
| 字号 24 写 `24` | 实际是 0.24 pt | 写 `2400` |
| `<a:alpha>` 在 `<a:solidFill>` 外 | PowerPoint 忽略 alpha | 移到 `<a:srgbClr>` 内部 |
| alpha 写 40（应为 40000） | 几乎完全透明 | alpha 范围 0-100000 |
| `<a:xfrm>` 属性与子元素分离 | 重复 `<a:xfrm>` 标签 | 用 `open_with` 合并属性 |
| `<p:ph>` 缺 `type` 属性 | PowerPoint 警告 | 默认 `type="body"` |

## Solution Patterns

### Pattern 1: write_xml 必须遵守子元素顺序

```rust
// ✅ 严格按 OOXML 顺序
fn write_xml(&self, w: &mut XmlWriter) {
    w.open("p:nvSpPr");  // 1
    // ...
    w.close("p:nvSpPr");
    w.open("p:spPr");    // 2
    // ...
    w.close("p:spPr");
    w.open("p:txBody");  // 3
    // ...
    w.close("p:txBody");
}

// ❌ 顺序错误
fn write_xml(&self, w: &mut XmlWriter) {
    w.open("p:txBody");  // 3 放前面了！
    w.open("p:spPr");    // 2 放后面了！
    // → WPS 可能拒绝打开
}
```

**适用场景**：任何 `write_xml` 实现。
**不适场景**：属性顺序（PowerPoint 对属性顺序不敏感）。

### Pattern 2: xfrm 属性用 open_with 合并

```rust
// ✅ 属性合并到外层标签
let mut attrs: Vec<(&str, &str)> = Vec::new();
if self.flip_h { attrs.push(("flipH", "1")); }
if let Some(s) = &rot_s { attrs.push(("rot", s.as_str())); }
w.open_with("a:xfrm", &attrs);
w.empty_with("a:off", &[("x", xs.as_str()), ("y", ys.as_str())]);
w.close("a:xfrm");

// ❌ 属性与子元素分离导致重复标签
w.open("a:xfrm");  // 无属性的空标签
w.open_with("a:xfrm", &attrs);  // 又开了一个！
// → 生成两层 <a:xfrm>，PowerPoint 报错
```

**适用场景**：`<a:xfrm>`、`<a:ln>` 等既有属性又有子元素的标签。
**不适场景**：纯属性标签（用 `empty_with`）或纯子元素标签（用 `open`/`close`）。

## Quick Reference（5 秒速查）

```
// 单位换算
1 inch = 914 400 EMU
1 pt   = 12 700 EMU
1 cm   = 360 000 EMU
字号 24pt → sz="2400"（百分之一磅）
旋转 30°  → rot="1800000"（六万分之一度）
alpha 40% → val="40000"（十万分之比）

// 子元素顺序口诀
<p:sp>  = nvSpPr → spPr → txBody
<p:pic> = nvPicPr → blipFill → spPr
<spPr>  = xfrm → prstGeom → fill → ln → effectLst
<rPr>   = 属性 → ln → fill → effectLst → latin/ea/cs
```

## Review Checklist

- [ ] `write_xml` 子元素顺序符合 OOXML 规范
- [ ] 有属性的标签用 `open_with` / `empty_with`
- [ ] 命名空间在根元素声明，不在子元素重复
- [ ] alpha 值在 0-100000 范围内
- [ ] 旋转值已乘以 60_000
- [ ] 字号值已乘以 100
- [ ] 关系 Target 用相对路径（`..` 前缀）
- [ ] 占位符 `is_placeholder=true` 时 `ph_type` 有默认值

## Cross-References

- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 架构详解（序列化约定）
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南（OOXML 问题排查）
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南（新增属性/元素）
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
