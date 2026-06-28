# OOXML 速查

> pptx-rs 涉及的 OOXML / DrawingML / PresentationML 关键元素与约束。
> 完整规范参考：<https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/>

## 1. 命名空间

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

源码常量见 [`src/oxml/ns.rs`](../src/oxml/ns.rs)。

## 2. 单位换算

| 物理量 | OOXML 单位 | 换算 |
| --- | --- | --- |
| 长度 | EMU (i64) | 1 inch = 914 400；1 cm = 360 000；1 pt = 12 700 |
| 字号 | 1/100 pt (i32) | Pt(24.0) → 2400 |
| 旋转 | 1/60000 度 (i32) | 30° → 1 800 000 |
| 缩进 | EMU | 同长度 |
| 百分比间距 | 1/1000 (i32) | 100% = 100 000；80% = 80 000 |
| 行距（百分比） | 1/1000 (i32) | 100% = 100 000；150% = 150 000 |
| 颜色 | sRGB / 主题色 / 预设色 | `<a:srgbClr val="RRGGBB"/>` 等 |

源码单位类型见 [`src/units.rs`](../src/units.rs)。

## 3. 必备元素清单

### 3.1 `<p:presentation>`

```xml
<p:presentation xmlns:a="..." xmlns:p="..." xmlns:r="..." xmlns:p14="...">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId10"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>...</p:defaultTextStyle>
</p:presentation>
```

### 3.2 `<p:sldMaster>`

必含：`<p:cSld>/<p:spTree>`、`<p:clrMap>`、`<p:sldLayoutIdLst>`、`<p:txStyles>`。

### 3.3 `<p:sldLayout>`

必含根属性 `type="blank"`（或 `title` 等）；必含 `<p:cSld>/<p:spTree>`、`<p:clrMapOvr>`。

### 3.4 `<p:sld>`

必含 `<p:cSld>/<p:spTree>`（含 `<p:nvGrpSpPr>` + `<p:grpSpPr>`）；可选 `<p:clrMapOvr>`、`<p:transition>`、`<p:timing>`。

### 3.5 `<a:theme>`

必含 `<a:themeElements>` → `<a:clrScheme>` + `<a:fontScheme>` + `<a:fmtScheme>`。完整内容见 [`src/oxml/theme.rs`](../src/oxml/theme.rs) 的 `THEME_XML`。

## 4. 形状子元素顺序

### 4.1 `<p:sp>`

```
<p:nvSpPr>     ← 1
  <p:cNvPr/>
  <p:cNvSpPr/>
  <p:nvPr/>    ← 可选：<p:ph type=... idx=...>
</p:nvSpPr>
<p:spPr/>      ← 2
<p:txBody/>    ← 3
<p:extLst/>    ← 4 (可选)
```

### 4.2 `<p:pic>`

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

### 4.3 `<p:spPr>`

```
<a:xfrm/>          ← 1
<a:prstGeom/>      ← 2 (或 <a:custGeom/>)
<a:noFill/>        ← 3 (或 <a:solidFill/>/<a:gradFill/>/...)
<a:ln/>            ← 4
<a:effectLst/>     ← 5
<a:scene3d/>       ← 6
<a:sp3d/>          ← 7
<a:extLst/>        ← 8
```

### 4.4 `<a:txBody>`

```
<a:bodyPr/>        ← 1
<a:lstStyle/>      ← 2
<a:p>              ← 3..n
  <a:pPr/>         ← 可选
  <a:r>            ← 1..n
    <a:rPr/>
    <a:t>...</a:t>
  </a:r>
  <a:endParaRPr/>  ← 可选
</a:p>
<a:extLst/>        ← 最后
```

### 4.5 `<a:rPr>` 子元素顺序

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

## 5. 关系（.rels）

| 文件 | 关系 | Target 相对路径基准 |
| --- | --- | --- |
| `/_rels/.rels` | 根：officeDocument + coreProps + appProps | 绝对 `/ppt/presentation.xml` |
| `/ppt/_rels/presentation.xml.rels` | presentation → slideMaster / slideLayout / theme / presProps / viewProps / tableStyles | 相对 `ppt/`，如 `slideMasters/slideMaster1.xml` |
| `/ppt/slideMasters/_rels/slideMaster1.xml.rels` | master → slideLayout / theme | 相对 `ppt/slideMasters/`，如 `../slideLayouts/slideLayout1.xml` |
| `/ppt/slideLayouts/_rels/slideLayout1.xml.rels` | layout → slideMaster | 相对 `ppt/slideLayouts/`，如 `../slideMasters/slideMaster1.xml` |
| `/ppt/slides/_rels/slideN.xml.rels` | slide → slideLayout / image | 相对 `ppt/slides/`，如 `../slideLayouts/slideLayout1.xml` |

关系文件结构：

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
```

## 6. Content-Types

每个 XML part 都应有 `<Override>`；`xml` / `rels` / `png` / `jpeg` 等可走 `<Default>`。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <!-- ... -->
</Types>
```

## 7. 写值示例

### 7.1 Run（完整）

```xml
<a:r>
  <a:rPr lang="zh-CN" altLang="en-US" sz="2400" b="1" i="1" u="sng"
         strike="sngStrike" baseline="30000" cap="small" kern="1200" spc="-100">
    <a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
    <a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill>
    <a:highlight><a:srgbClr val="FFFF00"/></a:highlight>
    <a:latin typeface="+mn-lt"/>
    <a:ea   typeface="宋体"/>
    <a:cs   typeface="+mn-cs"/>
  </a:rPr>
  <a:t>Hello, World</a:t>
</a:r>
```

### 7.2 形状属性

```xml
<p:spPr>
  <a:xfrm rot="1800000" flipH="1">
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

### 7.3 图片

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

### 7.4 表格

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
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

## 8. 文档保护

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

- 位置：`<p:presentation>` 内部、`<p:extLst>` 之前。
- 算法：SHA-512 + salt + spinCount=100 000。

完整实现见 [`examples/protect_pptx.rs`](../examples/protect_pptx.rs)。

## 9. 常见错误

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

## 10. 主题模板

完整 Office Theme 见 [`src/oxml/theme.rs::THEME_XML`](../src/oxml/theme.rs)。包含：

- 12 个 clrScheme 项
- 40+ 复杂脚本字体（majorFont + minorFont）
- fillStyleLst / lnStyleLst / effectStyleLst / bgFillStyleLst

## 11. 元素速查表（按 OOXML 元素）

| 元素 | 前缀 | 所在模块 | pptx-rs 类型 |
| --- | --- | --- | --- |
| `p:presentation` | p | `oxml/presentation.rs` | `PresentationRoot` |
| `p:sldMaster` | p | `oxml/slidemaster.rs` | `SldMaster` |
| `p:sldLayout` | p | `oxml/slidelayout.rs` | `SldLayout` |
| `p:sld` | p | `oxml/slide.rs` | `Sld` |
| `p:sp` | p | `oxml/shape.rs` | `Sp` |
| `p:pic` | p | `oxml/shape.rs` | `Pic` |
| `p:grpSp` | p | `oxml/shape.rs` | `Group` |
| `p:cxnSp` | p | `oxml/shape.rs` | `Connector` |
| `p:graphicFrame` | p | `oxml/shape.rs` | `GraphicFrame` |
| `p:txBody` | p | `oxml/txbody.rs` | `TextBody` |
| `p:spPr` | p | `oxml/sppr.rs` | `ShapeProperties` |
| `a:xfrm` | a | `oxml/sppr.rs` | `Transform` |
| `a:prstGeom` | a | `oxml/sppr.rs` | `PresetGeometry` |
| `a:solidFill` | a | `oxml/sppr.rs` | `Fill::Solid` |
| `a:ln` | a | `oxml/sppr.rs` | `Line` |
| `a:r` | a | `oxml/txbody.rs` | `Run` |
| `a:p` | a | `oxml/txbody.rs` | `Paragraph` |
| `a:rPr` | a | `oxml/txbody.rs` | `RunProperties` |
| `a:bodyPr` | a | `oxml/txbody.rs` | `BodyProperties` |
| `a:tbl` | a | `oxml/table.rs` | `Table` |
| `a:t` | a | `oxml/txbody.rs` | `Run::text` |
| `a:srgbClr` | a | `oxml/color.rs` | `Color::RGB` |
| `a:schemeClr` | a | `oxml/color.rs` | `Color::Scheme` |
| `a:prstClr` | a | `oxml/color.rs` | `Color::Preset` |
| `a:theme` | a | `oxml/theme.rs` | `default_theme_xml()` |
