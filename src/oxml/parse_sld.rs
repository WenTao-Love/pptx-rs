//! `<p:sld>` / `<p:sp>` / `<p:txBody>` 的反序列化（read 路径）。
//!
//! 本文件承担"读"职责——把 `slideN.xml` 的 XML 文本反序列化为强类型 [`Sld`]。
//! 由于 OOXML 元素嵌套极深、命名空间混杂，这里采用 **SAX 风格**流式解析：
//!
//! 1. 用 [`quick_xml`] 逐事件读取（`Start` / `Empty` / `End` / `Text`）；
//! 2. 在 `<p:spTree>` 阶段对每个 `<p:sp>` 调用 `parse_sp`；
//! 3. 在 sp 内部分别 parse `nvSpPr` / `spPr` / `txBody`；
//! 4. 跨过我们不识别的元素（pic / grpSp / cxnSp / graphicFrame）时**直接吞掉**子节点。
//!
//! # 设计取舍
//!
//! - **不递归到所有子元素**——例如 `a:xfrm` 内部的 off/ext 不做位置/尺寸校验，
//!   仅提取 4 个数字属性；
//! - **`a:rPr` 仅解析属性**——RunProperties 的子元素（`<a:solidFill>` 等）不递归
//!   解析为 Color，而是 **保留原始 XML 字符串** 以便原样回写；
//! - **遇到未知命名空间或属性** 静默忽略（OOXML 经常带 PowerPoint 私有扩展）。
//!
//! # 与 python-pptx 的差异
//!
//! python-pptx 中 `Slide.shapes` 持有 `SlideShapes`，每个 shape 是一个 `Shape`
//! 子类（`TextBox` / `Picture` / ...）。本库采用**强类型枚举** `SlideShape`，
//! 调用方在 `match` 时决定行为。
//!
//! # 失败模式
//!
//! 任何 XML 解析错误（标签不闭合、属性不是数字、必须属性缺失）均返回
//! [`crate::Error::Xml`]，不会 panic。`pic` / `grpSp` / `cxnSp` / `graphicFrame`
//! 的 XML 仅做"原样保留字符串"处理；它们被保存时**不会**从字符串恢复成结构体，
//! 仍可保证 zip 字节不丢。

// SAX 解析中用 Default 初始化结构体后逐字段赋值是正常模式，允许此 clippy 警告。
#![allow(clippy::field_reassign_with_default)]

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::oxml::color::{Color, PresetColor, SchemeColor};
use crate::oxml::shape::{
    Connector, GraphicFrame, Group, GroupChild, Pic, ShapeLocks, ShapeStyle, Sp, StyleRef,
};
use crate::oxml::simpletypes::{
    Alignment, MsoAnchor, PresetGeometry, TabAlignment, TextWrapping, Underline,
};
use crate::oxml::slide::Sld;
use crate::oxml::slide::{
    MorphOption, SplitOrientation, Transition, TransitionDirection, TransitionSpeed, TransitionType,
};
use crate::oxml::sppr::{
    ArrowHead, ArrowSize, ArrowType, Backdrop, Bevel, Camera, CameraPreset, CustomGeometry,
    EffectList, Fill, GeomRect, Geometry, GlowEffect, GradientFill, GradientPath, GradientStop,
    GradientType, LightRig, LightRigDirection, LightRigType, Line, LineJoin, MaterialPreset, Path,
    PathSegment, PatternFill, Point3d, ReflectionEffect, Rotation3d, Scene3d, ShadowEffect,
    ShapeProperties, SoftEdgeEffect, Sp3d,
};
use crate::oxml::txbody::{
    BodyProperties, BulletStyle, Field, FieldType, Hyperlink, Paragraph, ParagraphProperties, Run,
    RunProperties, TabStop, TextBody,
};
use crate::oxml::SlideShape as OxmlSlideShape;
use crate::units::{Emu, Pt, RGBColor};

/// 从 `slideN.xml` 文本解析出 [`Sld`]。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误或关键属性缺失。
///
/// # 注意
/// - 返回的 `Sld` 字段中，`layout_rid` **不会**自动填充——`Presentation::load`
///   会从 `slideN.xml.rels` 中解析并回填；
/// - `ext_lst` 暂不解析（保留 `None`），如需原样保留需扩展实现。
pub fn parse_sld(xml: &str) -> crate::Result<Sld> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sld = Sld::default();
    // 状态机：
    //   0: 等待 <p:sld> 入口；
    //   1: 在 <p:sld> 内部；
    //   2: 在 <p:spTree> 内部。
    let mut state: u8 = 0;
    // 使用 Vec 而非 Option，因为一张 slide 可以包含多个同类型形状
    let mut sp_bufs: Vec<String> = Vec::new();
    let mut pic_bufs: Vec<String> = Vec::new();
    let mut cxn_bufs: Vec<String> = Vec::new();
    let mut grp_bufs: Vec<String> = Vec::new();
    let mut gfx_bufs: Vec<String> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"sld" => {
                        // 读取 <p:sld> 标签上的 id 属性（虽然 0.1.0 sld.id 主要由
                        // Presentation 内部 sldIdLst 决定，但保留供 read-modify-write）。
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"id" {
                                if let Ok(v) = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse::<u32>()
                                {
                                    sld.id = v;
                                }
                            }
                        }
                        state = 1;
                    }
                    1 if local == b"cSld" => {
                        // 读取 <p:cSld name="...">（slide 用户可读名，对标 python-pptx Slide.name）。
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"name" {
                                sld.name = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                    }
                    1 if local == b"spTree" => {
                        state = 2;
                    }
                    1 if local == b"transition" => {
                        // <p:transition>（TODO-020：幻灯片过渡）
                        // 位于 <p:cSld> 之后、<p:timing> 之前
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        if let Ok(tr) = parse_transition(&inner) {
                            sld.transition = Some(tr);
                        }
                    }
                    2 if local == b"sp" => {
                        // 累积 sp 的内部 XML（含闭合标签）—— 走子解析
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp_bufs.push(inner);
                    }
                    2 if local == b"pic" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        pic_bufs.push(inner);
                    }
                    2 if local == b"cxnSp" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        cxn_bufs.push(inner);
                    }
                    2 if local == b"grpSp" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        grp_bufs.push(inner);
                    }
                    2 if local == b"graphicFrame" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        gfx_bufs.push(inner);
                    }
                    _ => {
                        // 其它子元素（nvGrpSpPr / grpSpPr 等）—— 整个吞掉
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 2 {
                    if local == b"sp" {
                        // 自闭合 sp（极少见）—— 不处理
                    } else if local == b"pic" {
                        // 自闭合 pic —— 跳过
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 1 && local == b"sld" {
                    break;
                } else if state == 2 && local == b"spTree" {
                    state = 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("sld parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 解析 sp（最常见类型）
    for s in sp_bufs {
        match parse_sp(&s) {
            Ok(sp) => sld.shapes.push(OxmlSlideShape::Sp(sp)),
            Err(e) => {
                // 容错：单个 sp 解析失败不致命，但记入 result
                return Err(e);
            }
        }
    }
    // 解析 pic（图片）
    for p in pic_bufs {
        if let Ok(pic) = parse_pic(&p) {
            sld.shapes.push(OxmlSlideShape::Pic(pic));
        }
        // 解析失败时直接跳过该 pic —— 重要：避免一次失败拖垮整张 slide
    }
    // 解析 cxnSp（连接器）
    for c in cxn_bufs {
        if let Ok(cxn) = parse_cxn_sp(&c) {
            sld.shapes.push(OxmlSlideShape::CxnSp(cxn));
        }
    }
    // 解析 grpSp（组合形状，递归）
    for g in grp_bufs {
        if let Ok(grp) = parse_grp_sp(&g) {
            sld.shapes.push(OxmlSlideShape::Group(Box::new(grp)));
        }
    }
    // 解析 graphicFrame（图形框：表格/图表）
    for f in gfx_bufs {
        if let Ok(frame) = parse_graphic_frame(&f) {
            sld.shapes.push(OxmlSlideShape::GraphicFrame(frame));
        }
    }
    Ok(sld)
}

/// 解析 `<p:transition>` 元素（TODO-020：幻灯片过渡）。
///
/// # 元素结构
///
/// ```text
/// <p:transition spd="slow|med|fast" advClick="0|1" advTm="...">
///   <p:fade thruBlk="1"/>       ← 或 push/wipe/split/cover/pull/cut/zoom/morph
/// </p:transition>
/// ```
///
/// # 参数
/// - `xml`：包含 `<p:transition>...</p:transition>` 的完整 XML 片段。
///
/// # 返回值
/// - 成功：返回 [`Transition`]；失败：返回 [`crate::Error::Xml`]。
pub fn parse_transition(xml: &str) -> crate::Result<Transition> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut tr = Transition::default();
    // 默认点击换片为 true（OOXML 默认值）
    tr.advance_click = true;
    let mut in_transition = false;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if !in_transition && local == b"transition" {
                    in_transition = true;
                    // 读取 <p:transition> 上的属性
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match key {
                            b"spd" => {
                                tr.speed = match v.as_ref() {
                                    "slow" => TransitionSpeed::Slow,
                                    "fast" => TransitionSpeed::Fast,
                                    _ => TransitionSpeed::Medium,
                                };
                            }
                            b"advClick" => {
                                tr.advance_click = v == "1" || v == "true";
                            }
                            b"advTm" => {
                                tr.advance_after_ms = v.parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                } else if in_transition {
                    // 解析过渡类型子元素
                    parse_transition_type_child(local, &e, &mut tr, false)?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if !in_transition && local == b"transition" {
                    // 自闭合 <p:transition/>：无子元素，使用默认值
                    in_transition = true;
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match key {
                            b"spd" => {
                                tr.speed = match v.as_ref() {
                                    "slow" => TransitionSpeed::Slow,
                                    "fast" => TransitionSpeed::Fast,
                                    _ => TransitionSpeed::Medium,
                                };
                            }
                            b"advClick" => {
                                tr.advance_click = v == "1" || v == "true";
                            }
                            b"advTm" => {
                                tr.advance_after_ms = v.parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                } else if in_transition {
                    // 自闭合过渡类型子元素
                    parse_transition_type_child(local, &e, &mut tr, true)?;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_transition && local == b"transition" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("transition parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(tr)
}

/// 解析过渡类型子元素（fade/push/wipe/split/cover/pull/cut/zoom/morph）。
///
/// # 参数
/// - `local`：元素本地名（不含命名空间前缀）。
/// - `e`：XML 事件（Start 或 Empty）。
/// - `tr`：待填充的 [`Transition`]。
/// - `is_empty`：是否为自闭合元素（Empty 事件）。
fn parse_transition_type_child(
    local: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    tr: &mut Transition,
    _is_empty: bool,
) -> crate::Result<()> {
    // 提取属性辅助闭包
    let attr = |key: &[u8]| -> Option<String> {
        for a in e.attributes().flatten() {
            if a.key.as_ref() == key {
                return Some(
                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .to_string(),
                );
            }
        }
        None
    };

    match local {
        b"fade" => {
            let thru_blk = attr(b"thruBlk")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            tr.transition_type = TransitionType::Fade { thru_blk };
        }
        b"push" => {
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Push { dir };
        }
        b"wipe" => {
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Wipe { dir };
        }
        b"split" => {
            let orient = match attr(b"orient").as_deref() {
                Some("vert") => SplitOrientation::Vertical,
                _ => SplitOrientation::Horizontal,
            };
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Split { orient, dir };
        }
        b"cover" => {
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Cover { dir };
        }
        b"pull" => {
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Pull { dir };
        }
        b"cut" => {
            let thru_blk = attr(b"thruBlk")
                .map(|v| v == "1" || v == "true")
                .unwrap_or(false);
            tr.transition_type = TransitionType::Cut { thru_blk };
        }
        b"zoom" => {
            let dir = parse_transition_dir(&attr(b"dir").unwrap_or_default());
            tr.transition_type = TransitionType::Zoom { dir };
        }
        b"morph" => {
            let option = match attr(b"option").as_deref() {
                Some("byWord") => MorphOption::ByWord,
                Some("byChar") => MorphOption::ByChar,
                _ => MorphOption::ByObject,
            };
            tr.transition_type = TransitionType::Morph { option };
        }
        _ => {
            // 未知过渡类型，保持默认（None）
        }
    }
    Ok(())
}

/// 解析过渡方向字符串为 [`TransitionDirection`]。
fn parse_transition_dir(s: &str) -> TransitionDirection {
    match s {
        "l" => TransitionDirection::Left,
        "r" => TransitionDirection::Right,
        "u" => TransitionDirection::Up,
        "d" => TransitionDirection::Down,
        "lu" => TransitionDirection::LeftUp,
        "ld" => TransitionDirection::LeftDown,
        "ru" => TransitionDirection::RightUp,
        "rd" => TransitionDirection::RightDown,
        _ => TransitionDirection::Right,
    }
}

/// 从 `slideN.xml.rels` 文本中查找 `rId` 对应的 `Target`。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 解析失败。
/// - `Error::Opc`：`rId` 不存在。
pub fn find_relationship_target(rels_xml: &str, rid: &str) -> crate::Result<String> {
    let rels = crate::opc::rels::Relationships::from_xml(rels_xml)?;
    rels.get(rid)
        .map(|r| r.target.as_str().to_string())
        .ok_or_else(|| crate::Error::opc(format!("relationship not found: {rid}")))
}

/// 从 `slideN.xml.rels` 中枚举所有关系（rid → target 路径）。
pub fn all_relationships(rels_xml: &str) -> crate::Result<Vec<(String, String, String)>> {
    // 返回 (rid, reltype_uri, target)
    use crate::opc::rels::Relationships;
    let rels = Relationships::from_xml(rels_xml)?;
    Ok(rels
        .iter()
        .map(|r| {
            (
                r.id.clone(),
                r.reltype.uri().to_string(),
                r.target.as_str().to_string(),
            )
        })
        .collect())
}

/// 从一段 XML 文本中解析 `p:sp` 元素（包含开始 + 子元素 + 结束）。
pub fn parse_sp(xml: &str) -> crate::Result<Sp> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sp = Sp::default();
    let mut state = 0u8; // 0: top, 1: nvSpPr, 2: spPr, 3: txBody

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"sp" => {
                        state = 1;
                    }
                    1 if local == b"nvSpPr" => {
                        // 手动遍历 nvSpPr 子元素提取 cNvPr/cNvSpPr/ph 属性，
                        // 不使用 collect_full_element_skipping（否则 Event::Empty 中的
                        // cNvPr 不会被处理）
                        let mut nv_depth = 1i32;
                        loop {
                            match rd.read_event_into(&mut buf) {
                                Ok(Event::Start(e2)) => {
                                    nv_depth += 1;
                                    // cNvSpPr 以 Start 事件出现时（含子元素如 spLocks），
                                    // 需要提取 txBox 属性
                                    let name2 = e2.name();
                                    let local2 = local_name(name2.as_ref());
                                    if local2 == b"cNvSpPr" {
                                        for a in e2.attributes().flatten() {
                                            if a.key.as_ref() == b"txBox" {
                                                let v = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default();
                                                if v == "1" || v == "true" {
                                                    sp.c_nv_sp_pr_tx_box = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(_)) => {
                                    nv_depth -= 1;
                                    if nv_depth == 0 {
                                        break;
                                    }
                                }
                                Ok(Event::Empty(e2)) => {
                                    let name2 = e2.name();
                                    let local2 = local_name(name2.as_ref());
                                    if local2 == b"cNvPr" {
                                        for a in e2.attributes().flatten() {
                                            match a.key.as_ref() {
                                                b"id" => {
                                                    if let Ok(v) = a
                                                        .normalized_value(
                                                            quick_xml::XmlVersion::Implicit1_0,
                                                        )
                                                        .unwrap_or_default()
                                                        .parse::<u32>()
                                                    {
                                                        sp.id = v;
                                                    }
                                                }
                                                b"name" => {
                                                    sp.name = a
                                                        .normalized_value(
                                                            quick_xml::XmlVersion::Implicit1_0,
                                                        )
                                                        .unwrap_or_default()
                                                        .to_string();
                                                }
                                                _ => {}
                                            }
                                        }
                                    } else if local2 == b"cNvSpPr" {
                                        for a in e2.attributes().flatten() {
                                            if a.key.as_ref() == b"txBox" {
                                                let v = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default();
                                                if v == "1" || v == "true" {
                                                    sp.c_nv_sp_pr_tx_box = true;
                                                }
                                            }
                                        }
                                    } else if local2 == b"spLocks" {
                                        // <a:spLocks noGrp="1" noSelect="1"/>
                                        sp.locks = Some(parse_sp_locks_attrs(&e2));
                                    } else if local2 == b"ph" {
                                        sp.is_placeholder = true;
                                        for a in e2.attributes().flatten() {
                                            match a.key.as_ref() {
                                                b"type" => {
                                                    sp.ph_type = Some(
                                                        String::from_utf8_lossy(a.value.as_ref())
                                                            .into_owned(),
                                                    )
                                                }
                                                b"idx" => {
                                                    if let Ok(v) = a
                                                        .normalized_value(
                                                            quick_xml::XmlVersion::Implicit1_0,
                                                        )
                                                        .unwrap_or_default()
                                                        .parse::<u32>()
                                                    {
                                                        sp.ph_idx = Some(v);
                                                    }
                                                }
                                                _ => {}
                                            }
                                        }
                                    }
                                }
                                Ok(Event::Eof) => break,
                                Err(_) => break,
                                _ => {}
                            }
                            buf.clear();
                        }
                    }
                    1 if local == b"spPr" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        // spPr 解析（独立子解析器），直接使用 inner 避免从 Option 中 unwrap
                        sp.properties = parse_sppr(&inner)?;
                        state = 2;
                    }
                    _ if state < 2 && local == b"spPr" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp.properties = parse_sppr(&inner)?;
                        state = 2;
                    }
                    2 if local == b"txBody" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        // 直接使用 inner 避免从 Option 中 unwrap
                        sp.text = parse_txbody(&inner)?;
                        state = 3;
                    }
                    _ if state >= 2 && local == b"style" => {
                        // <p:style> 在 spPr 之后、txBody 之前（或之后）出现
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp.style = Some(parse_shape_style(&inner)?);
                    }
                    _ => {
                        // 其它元素：吞掉
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 1 && local == b"ph" {
                    // 占位符（自闭合）
                    sp.is_placeholder = true;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"type" => {
                                sp.ph_type =
                                    Some(String::from_utf8_lossy(a.value.as_ref()).into_owned())
                            }
                            b"idx" => {
                                if let Ok(v) = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse::<u32>()
                                {
                                    sp.ph_idx = Some(v);
                                }
                            }
                            _ => {}
                        }
                    }
                } else if state == 1 && local == b"cNvPr" {
                    // cNvPr 自闭合（在 nvSpPr 内部）
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"id" => {
                                if let Ok(v) = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse::<u32>()
                                {
                                    sp.id = v;
                                }
                            }
                            b"name" => {
                                sp.name = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                            _ => {}
                        }
                    }
                } else if state == 1 && local == b"cNvSpPr" {
                    // cNvSpPr 自闭合
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"txBox" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if v == "1" || v == "true" {
                                sp.c_nv_sp_pr_tx_box = true;
                            }
                        }
                    }
                } else if state == 1 && local == b"spPr" {
                    // 自闭合 <p:spPr/>：无子元素，直接转到 state 2，
                    // 以便后续 <p:style> 能被正确匹配（state >= 2 守卫）
                    state = 2;
                } else if state == 2 && local == b"prstGeom" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"prst" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if let Ok(g) = v.parse::<PresetGeometry>() {
                                sp.properties.geometry = Some(Geometry::preset(g));
                            }
                        }
                    }
                } else if state == 2 && local == b"custGeom" {
                    // 自定义几何（TODO-024）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    if let Ok(cg) = parse_custom_geometry(&inner) {
                        sp.properties.geometry = Some(Geometry::Custom(cg));
                    }
                } else if state >= 2 && local == b"style" {
                    // 自闭合 <p:style/>：无子元素，创建空的 ShapeStyle
                    sp.style = Some(ShapeStyle::default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 3 && local == b"sp" {
                    break;
                }
                if state == 2 && local == b"sp" {
                    // 没有 txBody
                    break;
                }
                if state == 1 && local == b"sp" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("sp parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 兜底：若 spPr 解析失败，properties 已是默认
    Ok(sp)
}

/// 解析 `p:spPr` 元素。
pub fn parse_sppr(xml: &str) -> crate::Result<ShapeProperties> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sp = ShapeProperties::default();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"spPr" {
                    // 根元素 <p:spPr>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 xfrm / prstGeom 等无法解析）
                } else if local == b"xfrm" {
                    // 先提取 xfrm 自身的属性（rot/flipH/flipV）
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"rot" => {
                                if let Ok(v) = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse::<i32>()
                                {
                                    sp.xfrm.rot = Some(v);
                                }
                            }
                            b"flipH" => {
                                let v = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default();
                                if v == "1" {
                                    sp.xfrm.flip_h = true;
                                }
                            }
                            b"flipV" => {
                                let v = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default();
                                if v == "1" {
                                    sp.xfrm.flip_v = true;
                                }
                            }
                            _ => {}
                        }
                    }
                    // 再解析子元素（off/ext）
                    parse_xfrm_into(&mut rd, &mut sp);
                } else if local == b"prstGeom" {
                    // 先提取 prst 属性，再收集完整元素以解析 <a:avLst>
                    let mut prst_val: Option<String> = None;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"prst" {
                            prst_val = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    if let Some(v) = prst_val {
                        if let Ok(g) = v.parse::<PresetGeometry>() {
                            let adjustments = parse_av_lst(&inner);
                            sp.geometry = Some(Geometry::Preset(g, adjustments));
                        }
                    }
                } else if local == b"custGeom" {
                    // 自定义几何（TODO-024）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    if let Ok(cg) = parse_custom_geometry(&inner) {
                        sp.geometry = Some(Geometry::Custom(cg));
                    }
                } else if local == b"noFill" {
                    sp.fill = Fill::None;
                } else if local == b"solidFill" {
                    // 解析 a:solidFill 子元素 a:srgbClr / a:schemeClr
                    let color = parse_solid_fill(&mut rd, e.into_owned())?;
                    sp.fill = Fill::Solid(color);
                } else if local == b"gradFill" {
                    // 渐变填充：<a:gradFill><a:gsLst>...</a:gsLst><a:lin/>或<a:path/></a:gradFill>
                    let grad = parse_grad_fill(&mut rd, e.into_owned())?;
                    sp.fill = Fill::Gradient(grad);
                } else if local == b"pattFill" {
                    // 图案填充：<a:pattFill prst="..."><a:fgClr/>...<a:bgClr/>...</a:pattFill>
                    let patt = parse_patt_fill(&mut rd, e.into_owned())?;
                    sp.fill = Fill::Pattern(patt);
                } else if local == b"blipFill" {
                    // 图片填充（TODO-003/048）：解析 rid + mode，组装为 Fill::Blip。
                    // 注意：rid 仅是关系 id 字符串（如 "rId1"），实际 partname 解析
                    // 需要在外层 from_opc 路径中通过 slideN.xml.rels 映射——这里只保留 rid。
                    let (rid, mode) = parse_blip_fill(&mut rd, e.into_owned())?;
                    if !rid.is_empty() {
                        sp.fill = Fill::Blip { rid, mode };
                    } else {
                        // rid 缺失：保持 Inherit（避免写出无效的 Fill::Blip）
                    }
                } else if local == b"ln" {
                    let ln = parse_ln(&mut rd, e.into_owned())?;
                    sp.line = Some(ln);
                } else if local == b"effectLst" {
                    // 效果列表（TODO-011：形状效果）
                    let effects = parse_effect_lst(&mut rd, e.into_owned())?;
                    sp.effects = Some(effects);
                } else if local == b"scene3d" {
                    // 三维场景（TODO-050）
                    let scene = parse_scene_3d(&mut rd, e.into_owned())?;
                    sp.scene3d = Some(scene);
                } else if local == b"sp3d" {
                    // 形状 3D 属性（TODO-050）
                    let sp3d = parse_sp_3d(&mut rd, e.into_owned())?;
                    sp.sp3d = Some(sp3d);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"prstGeom" {
                    // 自闭合 <a:prstGeom prst="..."/>：无 avLst，调整值为空
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"prst" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if let Ok(g) = v.parse::<PresetGeometry>() {
                                sp.geometry = Some(Geometry::preset(g));
                            }
                        }
                    }
                } else if local == b"noFill" {
                    sp.fill = Fill::None;
                } else if local == b"custGeom" {
                    // 自闭合 <a:custGeom/>：罕见，用空自定义几何
                    sp.geometry = Some(Geometry::Custom(CustomGeometry::default()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("spPr parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(sp)
}

/// 解析 `a:xfrm`，就地更新 `ShapeProperties.xfrm`。
///
/// # 元素顺序（OOXML）
///
/// ```text
/// <a:xfrm rot="..." flipH="1" flipV="1">     ← 可选属性
///   <a:off x="..." y="..."/>                 ← 可选
///   <a:ext cx="..." cy="..."/>                ← 可选
/// </a:xfrm>
/// ```
/// 解析 `a:xfrm` 的子元素（off/ext），就地更新 `ShapeProperties.xfrm`。
///
/// # 调用约定
/// 调用方**必须**已消费 `<a:xfrm>` 的 Start 事件。本函数从 xfrm 的第一个子元素
/// 开始读取，直到遇到 `</a:xfrm>` End 事件为止。
///
/// xfrm 自身的属性（rot/flipH/flipV）由调用方从 Start 事件中提取，
/// 在调用本函数之前设置到 `sp.xfrm` 上。
///
/// # 元素顺序（OOXML）
///
/// ```text
/// <a:xfrm rot="..." flipH="1" flipV="1">     ← 属性由调用方处理
///   <a:off x="..." y="..."/>                 ← 可选，自闭合
///   <a:ext cx="..." cy="..."/>                ← 可选，自闭合
/// </a:xfrm>
/// ```
fn parse_xfrm_into<R: std::io::BufRead>(rd: &mut Reader<R>, sp: &mut ShapeProperties) {
    let mut buf = Vec::new();
    // 遍历 xfrm 的子元素：off / ext（均为自闭合 Empty 事件）
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"off" {
                    let (mut x, mut y) = (None, None);
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"x" => {
                                x = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            b"y" => {
                                y = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            _ => {}
                        }
                    }
                    sp.xfrm.off_x = x.map(Emu);
                    sp.xfrm.off_y = y.map(Emu);
                } else if local == b"ext" {
                    let (mut cx, mut cy) = (None, None);
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"cx" => {
                                cx = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            b"cy" => {
                                cy = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            _ => {}
                        }
                    }
                    sp.xfrm.ext_cx = cx.map(Emu);
                    sp.xfrm.ext_cy = cy.map(Emu);
                }
                // chOff / chExt 在组合形状中由 parse_grp_sppr_into 单独处理
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"xfrm" {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:prstGeom>` 内的 `<a:avLst>` 调整值列表。
///
/// 从 `collect_full_element` 产生的完整 `<a:prstGeom>` XML 中提取 `<a:gd>` 元素。
///
/// # 参数
/// - `xml`：包含 `<a:prstGeom>` 根元素的完整 XML 字符串。
///
/// # 返回值
/// 调整值列表。无 `<a:avLst>` 或 `<a:avLst/>` 为空时返回空 Vec。
fn parse_av_lst(xml: &str) -> Vec<crate::oxml::sppr::AdjustmentValue> {
    use crate::oxml::sppr::AdjustmentValue;

    let mut result = Vec::new();
    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"gd" {
                    let mut name = String::new();
                    let mut fmla = String::new();
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"name" => name = v.to_string(),
                            b"fmla" => fmla = v.to_string(),
                            _ => {}
                        }
                    }
                    // 解析 fmla="val <number>" 格式
                    if let Some(raw) = parse_fmla_val(&fmla) {
                        result.push(AdjustmentValue::new(name, raw));
                    }
                }
            }
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"gd" {
                    let mut name = String::new();
                    let mut fmla = String::new();
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"name" => name = v.to_string(),
                            b"fmla" => fmla = v.to_string(),
                            _ => {}
                        }
                    }
                    // 跳过子元素（gd 通常无子元素，但保险起见）
                    let _ = collect_full_element(&mut rd, e.into_owned());
                    if let Some(raw) = parse_fmla_val(&fmla) {
                        result.push(AdjustmentValue::new(name, raw));
                    }
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    result
}

/// 解析 `fmla="val <number>"` 格式的公式，提取数值部分。
///
/// 支持常见格式 `val 16667`（返回 `Some(16667)`）。
/// 不支持复杂公式（如 `*/ adj1 100000 50000`），返回 `None`。
fn parse_fmla_val(fmla: &str) -> Option<i64> {
    let trimmed = fmla.trim();
    if let Some(rest) = trimmed.strip_prefix("val ") {
        rest.trim().parse::<i64>().ok()
    } else if let Some(rest) = trimmed.strip_prefix("val") {
        rest.trim().parse::<i64>().ok()
    } else {
        None
    }
}

/// 解析 `<a:custGeom>...</a:custGeom>` 内部结构为 [`CustomGeometry`]。
///
/// # 元素顺序（OOXML 规范）
///
/// ```text
/// <a:custGeom>
///   <a:avLst/>          ← 调整手柄列表（暂跳过）
///   <a:fill>...</a:fill> ← 可选
///   <a:stroke>...</a:stroke> ← 可选
///   <a:rect l="..." t="..." r="..." b="..."/> ← 可选
///   <a:pathLst>
///     <a:path w="..." h="..." fill="..." stroke="...">
///       <a:moveTo>|<a:lnTo>|<a:cubicBezTo>|<a:quadBezTo>|<a:arcTo>|<a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
///
/// # 参数
/// - `xml`：包含 `<a:custGeom>` 根元素的完整 XML 字符串（由 `collect_full_element` 产生）。
///
/// # 返回值
/// - 成功：返回 [`CustomGeometry`]；失败：返回 `Error::Xml`。
fn parse_custom_geometry(xml: &str) -> crate::Result<CustomGeometry> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut geom = CustomGeometry::default();
    // 状态：0 = 顶层（寻找 custGeom 子元素），1 = 在 pathLst 内，2 = 在 path 内
    let mut state: u8 = 0;
    let mut current_path: Option<Path> = None;
    // 用于 cubicBezTo / quadBezTo 累积控制点
    let mut bez_pts: Vec<(i64, i64)> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                match state {
                    0 => match local {
                        b"fill" => {
                            // 读取文本内容
                            let text = read_element_text(&mut rd, b"fill");
                            geom.fill = Some(text);
                        }
                        b"stroke" => {
                            let text = read_element_text(&mut rd, b"stroke");
                            geom.stroke = Some(text);
                        }
                        b"rect" => {
                            // rect 是 Empty 元素，但若以 Start 形式出现也处理属性
                            geom.rect = Some(parse_geom_rect(&e));
                        }
                        b"pathLst" => {
                            state = 1;
                        }
                        b"avLst" => {
                            // 调整手柄列表，跳过（collect_full_element 已吞掉子元素，
                            // 但此处是 Start 事件，需要继续读到 End）
                            // 不做处理，让循环自然推进
                        }
                        _ => {}
                    },
                    1 => {
                        if local == b"path" {
                            // 进入 path 元素，读取属性
                            let mut p = Path::default();
                            for a in e.attributes().flatten() {
                                let key = a.key.as_ref();
                                let v = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                                match key {
                                    b"w" => p.width = v.parse().unwrap_or(0),
                                    b"h" => p.height = v.parse().unwrap_or(0),
                                    b"fill" => p.fill = Some(v),
                                    b"stroke" => p.stroke = Some(v),
                                    _ => {}
                                }
                            }
                            current_path = Some(p);
                            state = 2;
                        }
                    }
                    2 => {
                        // 在 path 内，处理路径段
                        match local {
                            b"moveTo" | b"lnTo" => {
                                // 读取单个 pt 子元素
                                let pt = read_single_pt(&mut rd, local);
                                if let Some((x, y)) = pt {
                                    let seg = if local == b"moveTo" {
                                        PathSegment::MoveTo { x, y }
                                    } else {
                                        PathSegment::LineTo { x, y }
                                    };
                                    if let Some(p) = &mut current_path {
                                        p.segments.push(seg);
                                    }
                                }
                            }
                            b"cubicBezTo" => {
                                // 读取 3 个 pt 子元素
                                bez_pts.clear();
                                read_multiple_pts(&mut rd, b"cubicBezTo", &mut bez_pts);
                                if bez_pts.len() >= 3 {
                                    let seg = PathSegment::CubicBezTo {
                                        x1: bez_pts[0].0,
                                        y1: bez_pts[0].1,
                                        x2: bez_pts[1].0,
                                        y2: bez_pts[1].1,
                                        x3: bez_pts[2].0,
                                        y3: bez_pts[2].1,
                                    };
                                    if let Some(p) = &mut current_path {
                                        p.segments.push(seg);
                                    }
                                }
                            }
                            b"quadBezTo" => {
                                // 读取 2 个 pt 子元素
                                bez_pts.clear();
                                read_multiple_pts(&mut rd, b"quadBezTo", &mut bez_pts);
                                if bez_pts.len() >= 2 {
                                    let seg = PathSegment::QuadBezTo {
                                        x1: bez_pts[0].0,
                                        y1: bez_pts[0].1,
                                        x2: bez_pts[1].0,
                                        y2: bez_pts[1].1,
                                    };
                                    if let Some(p) = &mut current_path {
                                        p.segments.push(seg);
                                    }
                                }
                            }
                            b"arcTo" => {
                                // arcTo 是自闭合元素，属性直接在标签上
                                let seg = parse_arc_to(&e);
                                if let Some(p) = &mut current_path {
                                    p.segments.push(seg);
                                }
                            }
                            _ => {}
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                match state {
                    0 => {
                        if local == b"rect" {
                            geom.rect = Some(parse_geom_rect(&e));
                        }
                        // avLst 空元素、其它空元素均忽略
                    }
                    2 => {
                        if local == b"close" {
                            if let Some(p) = &mut current_path {
                                p.segments.push(PathSegment::Close);
                            }
                        } else if local == b"arcTo" {
                            let seg = parse_arc_to(&e);
                            if let Some(p) = &mut current_path {
                                p.segments.push(seg);
                            }
                        } else if local == b"moveTo" || local == b"lnTo" {
                            // 自闭合的 moveTo/lnTo（罕见，但规范允许）含一个 pt 子元素
                            // Empty 事件没有子元素，所以无法读取 pt，跳过
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                match state {
                    2 => {
                        if local == b"path" {
                            // path 结束，推入 path_list
                            if let Some(p) = current_path.take() {
                                geom.path_list.push(p);
                            }
                            state = 1;
                        }
                    }
                    1 => {
                        if local == b"pathLst" {
                            state = 0;
                        }
                    }
                    0 if local == b"custGeom" => {
                        return Ok(geom);
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => return Ok(geom),
            Err(e) => return Err(crate::Error::Xml(format!("parse_custom_geometry: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 读取元素的文本内容（直到匹配的 End 事件）。
///
/// 用于 `<a:fill>norm</a:fill>` 这类文本叶子元素。
fn read_element_text<R: std::io::BufRead>(rd: &mut Reader<R>, expected_end: &[u8]) -> String {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                text.push_str(&t.decode().unwrap_or_default());
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == expected_end {
                    return text;
                }
            }
            Ok(Event::Eof) => return text,
            _ => {}
        }
        buf.clear();
    }
}

/// 从 `<a:pt x="..." y="..."/>` 属性中读取坐标。
fn parse_geom_rect(e: &quick_xml::events::BytesStart<'_>) -> GeomRect {
    let mut r = GeomRect::default();
    for a in e.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default()
            .to_string();
        match a.key.as_ref() {
            b"l" => r.l = v,
            b"t" => r.t = v,
            b"r" => r.r = v,
            b"b" => r.b = v,
            _ => {}
        }
    }
    r
}

/// 读取 moveTo / lnTo 内的单个 `<a:pt>` 子元素。
fn read_single_pt<R: std::io::BufRead>(rd: &mut Reader<R>, parent: &[u8]) -> Option<(i64, i64)> {
    let mut buf = Vec::new();
    let mut result = None;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let ev_name = e.name();
                if local_name(ev_name.as_ref()) == b"pt" {
                    result = parse_pt_attrs(&e);
                }
            }
            Ok(Event::Start(e)) => {
                let ev_name = e.name();
                if local_name(ev_name.as_ref()) == b"pt" {
                    // pt 可能含子元素（罕见），读取属性后跳到 End
                    result = parse_pt_attrs(&e);
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == parent {
                    return result;
                }
            }
            Ok(Event::Eof) => return result,
            _ => {}
        }
        buf.clear();
    }
}

/// 读取 cubicBezTo / quadBezTo 内的多个 `<a:pt>` 子元素。
fn read_multiple_pts<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    parent: &[u8],
    pts: &mut Vec<(i64, i64)>,
) {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let ev_name = e.name();
                if local_name(ev_name.as_ref()) == b"pt" {
                    if let Some(p) = parse_pt_attrs(&e) {
                        pts.push(p);
                    }
                }
            }
            Ok(Event::Start(e)) => {
                let ev_name = e.name();
                if local_name(ev_name.as_ref()) == b"pt" {
                    if let Some(p) = parse_pt_attrs(&e) {
                        pts.push(p);
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == parent {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            _ => {}
        }
        buf.clear();
    }
}

/// 从 `<a:pt>` 元素的属性中读取 (x, y) 坐标。
fn parse_pt_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<(i64, i64)> {
    let mut x: Option<i64> = None;
    let mut y: Option<i64> = None;
    for a in e.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match a.key.as_ref() {
            b"x" => x = v.parse().ok(),
            b"y" => y = v.parse().ok(),
            _ => {}
        }
    }
    match (x, y) {
        (Some(x), Some(y)) => Some((x, y)),
        _ => None,
    }
}

/// 从 `<a:arcTo>` 元素的属性中读取弧线参数。
fn parse_arc_to(e: &quick_xml::events::BytesStart<'_>) -> PathSegment {
    let mut w_r: i64 = 0;
    let mut h_r: i64 = 0;
    let mut st_ang: i32 = 0;
    let mut sw_ang: i32 = 0;
    for a in e.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match a.key.as_ref() {
            b"wR" => w_r = v.parse().unwrap_or(0),
            b"hR" => h_r = v.parse().unwrap_or(0),
            b"stAng" => st_ang = v.parse().unwrap_or(0),
            b"swAng" => sw_ang = v.parse().unwrap_or(0),
            _ => {}
        }
    }
    PathSegment::ArcTo {
        w_r,
        h_r,
        st_ang,
        sw_ang,
    }
}

/// 解析 `<a:solidFill>...</a:solidFill>` 内的颜色（必须包含 a:srgbClr 或 a:schemeClr）。
fn parse_solid_fill<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    _start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<Color> {
    let mut buf = Vec::new();
    // 读取第一个子元素
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"srgbClr" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if v.len() == 6 {
                                let r = u8::from_str_radix(&v[0..2], 16).unwrap_or(0);
                                let g = u8::from_str_radix(&v[2..4], 16).unwrap_or(0);
                                let b = u8::from_str_radix(&v[4..6], 16).unwrap_or(0);
                                return Ok(Color::RGB(RGBColor(r, g, b)));
                            }
                        }
                    }
                } else if local == b"schemeClr" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if let Ok(sc) = v.parse::<SchemeColor>() {
                                return Ok(Color::Scheme(sc));
                            }
                        }
                    }
                } else if local == b"prstClr" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if let Ok(p) = v.parse::<PresetColor>() {
                                return Ok(Color::Preset(p));
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"solidFill" {
                    return Ok(Color::None);
                }
            }
            Ok(Event::Eof) => return Ok(Color::None),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:gradFill>` 元素为 [`GradientFill`]。
///
/// # OOXML 结构
///
/// ```text
/// <a:gradFill flip="none" rotWithShape="1">
///   <a:gsLst>
///     <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
///     <a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs>
///   </a:gsLst>
///   <a:lin ang="5400000" scaled="1"/>     ← 线性渐变
///   <!-- 或 -->
///   <a:path path="circle">...</a:path>    ← 路径渐变
/// </a:gradFill>
/// ```
///
/// # 参数
/// - `rd`：已位于 `<a:gradFill>` Start 事件之后的 reader；
/// - `_start`：`<a:gradFill>` 的 Start 事件（用于提取 flip/rotWithShape 属性）。
fn parse_grad_fill<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<GradientFill> {
    // 先从 gradFill 自身属性提取 flip / rotWithShape
    let mut flip = None;
    let mut rot_with_shape = None;
    for a in start.attributes().flatten() {
        match a.key.as_ref() {
            b"flip" => {
                flip = Some(
                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .to_string(),
                );
            }
            b"rotWithShape" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                rot_with_shape = Some(v == "1");
            }
            _ => {}
        }
    }

    let mut stops: Vec<GradientStop> = Vec::new();
    let mut gradient_type = GradientType::Linear(0);
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"gsLst" {
                    // 解析光轨列表
                    let mut gs_depth = 1i32;
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Start(ref e2)) => {
                                if local_name(e2.name().as_ref()) == b"gs" && gs_depth == 1 {
                                    // 解析单个 <a:gs pos="...">
                                    let mut pos: u32 = 0;
                                    for a in e2.attributes().flatten() {
                                        if a.key.as_ref() == b"pos" {
                                            pos = a
                                                .normalized_value(
                                                    quick_xml::XmlVersion::Implicit1_0,
                                                )
                                                .unwrap_or_default()
                                                .parse()
                                                .unwrap_or(0);
                                        }
                                    }
                                    gs_depth += 1;
                                    // 读取 gs 内部的颜色子元素
                                    let color = parse_color_child(rd, b"gs")?;
                                    stops.push(GradientStop { pos, color });
                                } else {
                                    gs_depth += 1;
                                }
                            }
                            Ok(Event::End(ref e2)) => {
                                let n2 = e2.name();
                                let ln = local_name(n2.as_ref());
                                if ln == b"gs" {
                                    gs_depth -= 1;
                                } else if ln == b"gsLst" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"lin" {
                    // 线性渐变：<a:lin ang="5400000" scaled="1"/>
                    let mut ang: i32 = 0;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"ang" {
                            ang = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(0);
                        }
                    }
                    gradient_type = GradientType::Linear(ang);
                    // lin 可能是 Start 或 Empty；如果是 Start 需要读到 End
                    if e.is_empty() {
                        // Empty 事件不会触发 End，无需处理
                    } else {
                        // 读到 </a:lin>
                        let _ = collect_full_element(rd, e.into_owned())?;
                    }
                } else if local == b"path" {
                    // 路径渐变：<a:path path="circle|rect|shape">...</a:path>
                    let mut path_type = GradientPath::Circle;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"path" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            path_type = match v.as_ref() {
                                "circle" => GradientPath::Circle,
                                "rect" => GradientPath::Rect,
                                "shape" => GradientPath::Shape,
                                _ => GradientPath::Circle,
                            };
                        }
                    }
                    gradient_type = GradientType::Path(path_type);
                    // 吞掉 path 的子元素（fillToRect 等）
                    let _ = collect_full_element(rd, e.into_owned())?;
                } else {
                    let _ = collect_full_element(rd, e.into_owned())?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"lin" {
                    // 自闭合的 <a:lin ang="..." scaled="..."/>
                    let mut ang: i32 = 0;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"ang" {
                            ang = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(0);
                        }
                    }
                    gradient_type = GradientType::Linear(ang);
                } else if local == b"path" {
                    let mut path_type = GradientPath::Circle;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"path" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            path_type = match v.as_ref() {
                                "circle" => GradientPath::Circle,
                                "rect" => GradientPath::Rect,
                                "shape" => GradientPath::Shape,
                                _ => GradientPath::Circle,
                            };
                        }
                    }
                    gradient_type = GradientType::Path(path_type);
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"gradFill" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("gradFill parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(GradientFill {
        stops,
        gradient_type,
        flip,
        rot_with_shape,
    })
}

/// 解析 `<a:pattFill>` 元素为 [`PatternFill`]。
///
/// # OOXML 结构
///
/// ```text
/// <a:pattFill prst="pct5">
///   <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
///   <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
/// </a:pattFill>
/// ```
fn parse_patt_fill<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<PatternFill> {
    let mut prst = String::new();
    for a in start.attributes().flatten() {
        if a.key.as_ref() == b"prst" {
            prst = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default()
                .to_string();
        }
    }

    let mut fg_color = Color::None;
    let mut bg_color = Color::None;
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"fgClr" {
                    fg_color = parse_color_child(rd, b"fgClr")?;
                } else if local == b"bgClr" {
                    bg_color = parse_color_child(rd, b"bgClr")?;
                } else {
                    let _ = collect_full_element(rd, e.into_owned())?;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"pattFill" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("pattFill parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(PatternFill {
        prst,
        fg_color,
        bg_color,
    })
}

/// 从 `<a:blipFill>` 事件中解析图片填充（TODO-003/048）。
///
/// 解析 `<a:blip r:embed="rIdN"/>` 提取 rid，以及可选的 `<a:stretch>` / `<a:tile>` 填充模式。
/// `<a:srcRect>` 暂不解析（SpPr 级 blipFill 较少用 srcRect，主要在 `<p:pic>` 内使用）。
///
/// # 参数
/// - `rd`：reader，已消费 `<a:blipFill>` 的 Start 事件；
/// - `start`：`<a:blipFill>` 的 Start 事件（保留属性，目前不用）。
///
/// # 返回值
/// 返回 `(rid, mode)` 二元组，由调用方组装为 `Fill::Blip { rid, mode }`。
///
/// # OOXML 结构
///
/// ```text
/// <a:blipFill>
///   <a:blip r:embed="rId1"/>          ← 必填，引用图片 part
///   <a:srcRect l="..." t="..." .../>  ← 可选，裁剪（暂不解析）
///   <a:stretch><a:fillRect/></a:stretch>  ← 拉伸（默认）
///   或
///   <a:tile tx="..." ty="..." .../>   ← 平铺
/// </a:blipFill>
/// ```
fn parse_blip_fill<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    _start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<(String, crate::oxml::sppr::BlipFillMode)> {
    let mut rid = String::new();
    let mut mode = crate::oxml::sppr::BlipFillMode::default(); // 默认 Stretch
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"blip" {
                    // `<a:blip r:embed="rIdN">...</a:blip>`：可能有子元素（如 alphaModFix）
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        if key == b"r:embed" || key.ends_with(b":embed") {
                            rid = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                    // 吞掉 blip 的子元素（如 alphaModFix），直到 blip End
                    let _ = collect_full_element(rd, e.into_owned())?;
                } else if local == b"stretch" {
                    // `<a:stretch><a:fillRect/></a:stretch>`
                    let _ = collect_full_element(rd, e.into_owned())?;
                    mode = crate::oxml::sppr::BlipFillMode::Stretch;
                } else if local == b"tile" {
                    // `<a:tile tx="..." ty="..." sx="..." sy="..." flip="..." algn="..."/>`
                    mode = parse_tile_attrs(&e);
                    let _ = collect_full_element(rd, e.into_owned())?;
                } else if local == b"srcRect" {
                    // 暂不解析 srcRect（SpPr 级 blipFill 较少用）
                    let _ = collect_full_element(rd, e.into_owned())?;
                } else {
                    let _ = collect_full_element(rd, e.into_owned())?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"blip" {
                    // 自闭合 `<a:blip r:embed="rIdN"/>`
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        if key == b"r:embed" || key.ends_with(b":embed") {
                            rid = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if local == b"stretch" {
                    mode = crate::oxml::sppr::BlipFillMode::Stretch;
                } else if local == b"tile" {
                    mode = parse_tile_attrs(&e);
                }
                // srcRect / 其它自闭合元素忽略
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"blipFill" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("blipFill parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok((rid, mode))
}

/// 从 reader 中解析一个颜色子元素（`<a:srgbClr>` / `<a:schemeClr>` / `<a:prstClr>`）。
///
/// 调用方必须已消费父元素的 Start 事件。本函数读取父元素的子元素直到 End。
///
/// # 参数
/// - `rd`：reader；
/// - `parent`：父元素名（用于匹配 End 事件，如 `b"gs"` / `b"fgClr"` / `b"bgClr"`）。
fn parse_color_child<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    parent: &[u8],
) -> crate::Result<Color> {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if let Some(c) = parse_color_from_event(&e) {
                    return Ok(c);
                }
                // 不是颜色元素则忽略
                let _ = local;
            }
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"srgbClr" || local == b"schemeClr" || local == b"prstClr" {
                    // 非自闭合颜色元素：读属性后跳到 End
                    if let Some(c) = parse_color_from_event(&e) {
                        // 吞掉子元素（如 alpha）直到 End
                        let _ = collect_full_element(rd, e.into_owned())?;
                        return Ok(c);
                    }
                } else {
                    let _ = collect_full_element(rd, e.into_owned())?;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == parent {
                    return Ok(Color::None);
                }
            }
            Ok(Event::Eof) => return Ok(Color::None),
            Err(_) => return Ok(Color::None),
            _ => {}
        }
        buf.clear();
    }
}

/// 从一个 XML 事件（Start 或 Empty）中提取颜色。
///
/// 支持 `<a:srgbClr val="RRGGBB"/>` / `<a:schemeClr val="..."/>` / `<a:prstClr val="..."/>`。
fn parse_color_from_event(e: &quick_xml::events::BytesStart<'_>) -> Option<Color> {
    let name = e.name();
    let local = local_name(name.as_ref());
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"val" {
            let v = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default();
            if local == b"srgbClr" && v.len() == 6 {
                let r = u8::from_str_radix(&v[0..2], 16).unwrap_or(0);
                let g = u8::from_str_radix(&v[2..4], 16).unwrap_or(0);
                let b = u8::from_str_radix(&v[4..6], 16).unwrap_or(0);
                return Some(Color::RGB(RGBColor(r, g, b)));
            } else if local == b"schemeClr" {
                if let Ok(sc) = v.parse::<SchemeColor>() {
                    return Some(Color::Scheme(sc));
                }
            } else if local == b"prstClr" {
                if let Ok(p) = v.parse::<PresetColor>() {
                    return Some(Color::Preset(p));
                }
            }
        }
    }
    None
}

/// 解析 `<a:ln>...</a:ln>` 元素。
fn parse_ln<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<Line> {
    let mut ln = Line::default();
    // 解析 w / cap / cmpd 等属性
    for a in start.attributes().flatten() {
        match a.key.as_ref() {
            b"w" => {
                if let Ok(v) = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .parse::<i64>()
                {
                    ln.width = Some(Emu(v));
                }
            }
            b"cap" => {
                ln.cap = Some(
                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .to_string(),
                );
            }
            b"cmpd" => {
                ln.compound = Some(
                    a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .to_string(),
                );
            }
            _ => {}
        }
    }
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"solidFill" {
                    ln.color = parse_solid_fill(rd, e.into_owned())?;
                } else if local == b"noFill" {
                    ln.no_fill = true;
                    let _ = collect_full_element(rd, e.into_owned());
                } else if local == b"prstDash" {
                    // 寻找子元素 a:prstDash val="..."
                    parse_prstdash_into(rd, &mut ln);
                } else if local == b"gradFill" {
                    // 线条渐变填充
                    let grad = parse_grad_fill(rd, e.into_owned())?;
                    ln.fill = Fill::Gradient(grad);
                } else if local == b"pattFill" {
                    // 线条图案填充
                    let patt = parse_patt_fill(rd, e.into_owned())?;
                    ln.fill = Fill::Pattern(patt);
                } else if local == b"headEnd" {
                    ln.head_end = Some(parse_arrow_head(&e));
                    let _ = collect_full_element(rd, e.into_owned());
                } else if local == b"tailEnd" {
                    ln.tail_end = Some(parse_arrow_head(&e));
                    let _ = collect_full_element(rd, e.into_owned());
                } else if local == b"round" {
                    ln.join = Some(LineJoin::Round);
                    let _ = collect_full_element(rd, e.into_owned());
                } else if local == b"miter" {
                    let mut lim = 800000i32;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"lim" {
                            lim = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(800000);
                        }
                    }
                    ln.join = Some(LineJoin::Miter(lim));
                    let _ = collect_full_element(rd, e.into_owned());
                } else if local == b"bevel" {
                    ln.join = Some(LineJoin::Bevel);
                    let _ = collect_full_element(rd, e.into_owned());
                } else {
                    let _ = collect_full_element(rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"noFill" {
                    ln.no_fill = true;
                } else if local == b"headEnd" {
                    ln.head_end = Some(parse_arrow_head(&e));
                } else if local == b"tailEnd" {
                    ln.tail_end = Some(parse_arrow_head(&e));
                } else if local == b"round" {
                    ln.join = Some(LineJoin::Round);
                } else if local == b"miter" {
                    let mut lim = 800000i32;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"lim" {
                            lim = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(800000);
                        }
                    }
                    ln.join = Some(LineJoin::Miter(lim));
                } else if local == b"bevel" {
                    ln.join = Some(LineJoin::Bevel);
                }
                // solidFill 自闭合罕见，兜底忽略
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"ln" {
                    return Ok(ln);
                }
            }
            Ok(Event::Eof) => return Ok(ln),
            _ => {}
        }
        buf.clear();
    }
}

/// 从 `<a:headEnd>` / `<a:tailEnd>` 事件中提取箭头属性（type/w/len）。
///
/// 该函数仅读取属性，**不**消费子元素——调用方负责处理后续事件。
fn parse_arrow_head(e: &quick_xml::events::BytesStart<'_>) -> ArrowHead {
    let mut arrow_type = ArrowType::None;
    let mut width = ArrowSize::Medium;
    let mut length = ArrowSize::Medium;
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"type" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                arrow_type = match v.as_ref() {
                    "none" => ArrowType::None,
                    "triangle" => ArrowType::Triangle,
                    "stealth" => ArrowType::Stealth,
                    "diamond" => ArrowType::Diamond,
                    "oval" => ArrowType::Oval,
                    "arrow" => ArrowType::Arrow,
                    _ => ArrowType::None,
                };
            }
            b"w" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                width = match v.as_ref() {
                    "sm" => ArrowSize::Small,
                    "lg" => ArrowSize::Large,
                    _ => ArrowSize::Medium,
                };
            }
            b"len" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                length = match v.as_ref() {
                    "sm" => ArrowSize::Small,
                    "lg" => ArrowSize::Large,
                    _ => ArrowSize::Medium,
                };
            }
            _ => {}
        }
    }
    ArrowHead {
        arrow_type,
        width,
        length,
    }
}

fn parse_prstdash_into<R: std::io::BufRead>(rd: &mut Reader<R>, ln: &mut Line) {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"prst" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            if let Ok(d) = v.parse::<crate::oxml::sppr::Dash>() {
                                ln.dash = Some(d);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"prstDash" {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `p:txBody` 元素。
pub fn parse_txbody(xml: &str) -> crate::Result<TextBody> {
    let mut tb = TextBody::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut p_bufs: Vec<String> = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"txBody" {
                    // 根元素 <p:txBody>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 <a:p> 等无法解析）
                } else if local == b"p" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    p_bufs.push(inner);
                } else if local == b"bodyPr" {
                    // 解析 <a:bodyPr> 的属性（numCol / spcCol / anchor / wrap 等）
                    // 并保留其子元素（如 <a:spAutoFit/>）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    let bp = parse_body_pr(&inner)?;
                    tb.body_properties = Some(bp);
                } else {
                    // lstStyle 等跳过
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"bodyPr" {
                    // 自闭合 <a:bodyPr/>：仅解析属性
                    let bp = parse_body_pr_attrs(&e);
                    tb.body_properties = Some(bp);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("txBody parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    for p_xml in p_bufs {
        if let Ok(p) = parse_paragraph(&p_xml) {
            tb.paragraphs.push(p);
        }
    }
    Ok(tb)
}

/// 解析 `<a:effectLst>` 元素（TODO-011：形状效果）。
///
/// # 支持的子元素
/// - `<a:outerShdw>`：外阴影（dir/dist/blurRad/color/rotWithShape）
/// - `<a:innerShdw>`：内阴影（dir/dist/blurRad/color）
/// - `<a:glow>`：发光（rad/color）
/// - `<a:softEdge>`：柔化边缘（rad）
/// - `<a:reflection>`：反射（blurRad/stA/stPos/endA/endPos/dist/dir/rotWithShape）
fn parse_effect_lst<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<EffectList> {
    let mut effects = EffectList::default();
    let mut buf = Vec::new();
    // 消费 start 事件的属性（effectLst 本身无属性，但需进入子元素循环）
    let _ = start;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"outerShdw" => {
                        let shdw = parse_shadow_attrs(&e, false);
                        // 解析子元素（srgbClr/schemeClr）
                        let color = parse_color_child(rd, b"outerShdw")?;
                        let mut shdw = shdw;
                        shdw.color = color;
                        effects.outer_shadow = Some(shdw);
                    }
                    b"innerShdw" => {
                        let shdw = parse_shadow_attrs(&e, true);
                        let color = parse_color_child(rd, b"innerShdw")?;
                        let mut shdw = shdw;
                        shdw.color = color;
                        effects.inner_shadow = Some(shdw);
                    }
                    b"glow" => {
                        let mut glow = GlowEffect::default();
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"rad" {
                                glow.rad = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .unwrap_or(0);
                            }
                        }
                        glow.color = parse_color_child(rd, b"glow")?;
                        effects.glow = Some(glow);
                    }
                    b"softEdge" => {
                        let mut se = SoftEdgeEffect::default();
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"rad" {
                                se.rad = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .unwrap_or(0);
                            }
                        }
                        // softEdge 无子元素，消费到 End
                        let _ = collect_full_element(rd, e.into_owned());
                        effects.soft_edge = Some(se);
                    }
                    b"reflection" => {
                        let mut refl = ReflectionEffect::default();
                        for a in e.attributes().flatten() {
                            let key = a.key.as_ref();
                            let v = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default();
                            match key {
                                b"blurRad" => refl.blur_rad = v.parse().ok(),
                                b"stA" => refl.st_a = v.parse().ok(),
                                b"stPos" => refl.st_pos = v.parse().ok(),
                                b"endA" => refl.end_a = v.parse().ok(),
                                b"endPos" => refl.end_pos = v.parse().ok(),
                                b"dist" => refl.dist = v.parse().ok(),
                                b"dir" => refl.dir = v.parse().ok(),
                                b"rotWithShape" => {
                                    refl.rot_with_shape = Some(v == "1" || v == "true")
                                }
                                _ => {}
                            }
                        }
                        // reflection 无子元素，消费到 End
                        let _ = collect_full_element(rd, e.into_owned());
                        effects.reflection = Some(refl);
                    }
                    _ => {
                        // 未识别的效果子元素：吞掉
                        let _ = collect_full_element(rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"outerShdw" {
                    let mut shdw = parse_shadow_attrs(&e, false);
                    shdw.color = Color::None;
                    effects.outer_shadow = Some(shdw);
                } else if local == b"innerShdw" {
                    let mut shdw = parse_shadow_attrs(&e, true);
                    shdw.color = Color::None;
                    effects.inner_shadow = Some(shdw);
                } else if local == b"glow" {
                    let mut glow = GlowEffect::default();
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"rad" {
                            glow.rad = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(0);
                        }
                    }
                    effects.glow = Some(glow);
                } else if local == b"softEdge" {
                    let mut se = SoftEdgeEffect::default();
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"rad" {
                            se.rad = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse()
                                .unwrap_or(0);
                        }
                    }
                    effects.soft_edge = Some(se);
                } else if local == b"reflection" {
                    let mut refl = ReflectionEffect::default();
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match key {
                            b"blurRad" => refl.blur_rad = v.parse().ok(),
                            b"stA" => refl.st_a = v.parse().ok(),
                            b"stPos" => refl.st_pos = v.parse().ok(),
                            b"endA" => refl.end_a = v.parse().ok(),
                            b"endPos" => refl.end_pos = v.parse().ok(),
                            b"dist" => refl.dist = v.parse().ok(),
                            b"dir" => refl.dir = v.parse().ok(),
                            b"rotWithShape" => refl.rot_with_shape = Some(v == "1" || v == "true"),
                            _ => {}
                        }
                    }
                    effects.reflection = Some(refl);
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"effectLst" {
                    return Ok(effects);
                }
            }
            Ok(Event::Eof) => return Ok(effects),
            Err(e) => return Err(crate::Error::Xml(format!("effectLst parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:rot lat="..." lon="..." rev="..."/>` 的三个属性为 [`Rotation3d`]。
///
/// 用于 `camera` / `lightRig` 内的可选旋转元素。
fn parse_rotation_3d(e: &quick_xml::events::BytesStart<'_>) -> Rotation3d {
    let mut rot = Rotation3d::default();
    for a in e.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match a.key.as_ref() {
            b"lat" => rot.lat = v.parse().unwrap_or(0),
            b"lon" => rot.lon = v.parse().unwrap_or(0),
            b"rev" => rot.rev = v.parse().unwrap_or(0),
            _ => {}
        }
    }
    rot
}

/// 解析 `<a:scene3d>` 元素（TODO-050）。
///
/// # 元素顺序（OOXML）
/// ```text
/// <a:scene3d>
///   <a:camera prst="..." fov="..." zoom="...">    ← 必填
///     <a:rot lat="..." lon="..." rev="..."/>      ← 可选
///   </a:camera>
///   <a:lightRig rig="..." dir="...">              ← 必填
///     <a:rot lat="..." lon="..." rev="..."/>      ← 可选
///   </a:lightRig>
/// </a:scene3d>
/// ```
fn parse_scene_3d<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<Scene3d> {
    let mut scene = Scene3d::default();
    let mut buf = Vec::new();
    let _ = start;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"camera" {
                    let mut camera = Camera::default();
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            // v 是 Cow<'_, str>，from_str 期望 &str，需要 &v
                            b"prst" => camera.preset = CameraPreset::parse(&v),
                            b"fov" => camera.fov = v.parse().unwrap_or(0),
                            b"zoom" => camera.zoom = v.parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    // 解析可选 <a:rot>
                    let mut buf2 = Vec::new();
                    loop {
                        match rd.read_event_into(&mut buf2) {
                            Ok(Event::Start(e2)) => {
                                let n2 = e2.name();
                                if local_name(n2.as_ref()) == b"rot" {
                                    camera.rotation = Some(parse_rotation_3d(&e2));
                                    let _ = collect_full_element(rd, e2.into_owned());
                                }
                            }
                            Ok(Event::End(e2)) => {
                                let n2 = e2.name();
                                if local_name(n2.as_ref()) == b"camera" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                        buf2.clear();
                    }
                    scene.camera = camera;
                } else if local == b"lightRig" {
                    let mut rig = LightRig::default();
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            // v 是 Cow<'_, str>，from_str 期望 &str，需要 &v
                            b"rig" => rig.rig = LightRigType::parse(&v),
                            b"dir" => rig.dir = LightRigDirection::parse(&v),
                            _ => {}
                        }
                    }
                    let mut buf2 = Vec::new();
                    loop {
                        match rd.read_event_into(&mut buf2) {
                            Ok(Event::Start(e2)) => {
                                let n2 = e2.name();
                                if local_name(n2.as_ref()) == b"rot" {
                                    rig.rotation = Some(parse_rotation_3d(&e2));
                                    let _ = collect_full_element(rd, e2.into_owned());
                                }
                            }
                            Ok(Event::End(e2)) => {
                                let n2 = e2.name();
                                if local_name(n2.as_ref()) == b"lightRig" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                        buf2.clear();
                    }
                    scene.light_rig = rig;
                } else if local == b"backdrop" {
                    // 解析可选 <a:backdrop>（TODO-050）
                    let backdrop = parse_backdrop(rd, e.into_owned())?;
                    scene.backdrop = Some(backdrop);
                } else {
                    let _ = collect_full_element(rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"scene3d" {
                    return Ok(scene);
                }
            }
            Ok(Event::Eof) => return Ok(scene),
            Err(e) => return Err(crate::Error::Xml(format!("scene3d parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:backdrop>` 元素（TODO-050）。
///
/// # 元素顺序（OOXML CT_Backdrop）
/// ```text
/// <a:backdrop>
///   <a:anchor x="..." y="..." z="..."/>   ← 可选
///   <a:floor/>                            ← 可选
///   <a:wall/>                             ← 可选
///   <a:l/>                                ← 可选
///   <a:r/>                                ← 可选
///   <a:t/>                                ← 可选
///   <a:b/>                                ← 可选
/// </a:backdrop>
/// ```
fn parse_backdrop<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<Backdrop> {
    let mut bd = Backdrop::default();
    let mut buf = Vec::new();
    let _ = start;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"anchor" {
                    let mut pt = Point3d::default();
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"x" => pt.x = v.parse().unwrap_or(0),
                            b"y" => pt.y = v.parse().unwrap_or(0),
                            b"z" => pt.z = v.parse().unwrap_or(0),
                            _ => {}
                        }
                    }
                    bd.anchor = Some(pt);
                } else if local == b"floor" {
                    bd.floor = true;
                } else if local == b"wall" {
                    bd.wall = true;
                } else if local == b"l" {
                    bd.left = true;
                } else if local == b"r" {
                    bd.right = true;
                } else if local == b"t" {
                    bd.top = true;
                } else if local == b"b" {
                    bd.bottom = true;
                }
            }
            Ok(Event::Start(e)) => {
                // backdrop 的子元素都是自闭合的（Empty），无 Start 分支需要处理；
                // 但兼容罕见的 open-close 形式（如 <a:floor></a:floor>），用 collect_full_element 吞掉。
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"floor" {
                    bd.floor = true;
                } else if local == b"wall" {
                    bd.wall = true;
                } else if local == b"l" {
                    bd.left = true;
                } else if local == b"r" {
                    bd.right = true;
                } else if local == b"t" {
                    bd.top = true;
                } else if local == b"b" {
                    bd.bottom = true;
                }
                let _ = collect_full_element(rd, e.into_owned());
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"backdrop" {
                    return Ok(bd);
                }
            }
            Ok(Event::Eof) => return Ok(bd),
            Err(e) => return Err(crate::Error::Xml(format!("backdrop parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:sp3d>` 元素（TODO-050）。
///
/// # 元素顺序（OOXML）
/// ```text
/// <a:sp3d extrusionH="..." contourW="..." prstMaterial="...">
///   <a:bevelT w="..." h="..."/>     ← 可选
///   <a:bevelB w="..." h="..."/>     ← 可选
///   <a:extrusionClr>...</a:extrusionClr>   ← 可选
///   <a:contourClr>...</a:contourClr>       ← 可选
/// </a:sp3d>
/// ```
fn parse_sp_3d<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<Sp3d> {
    let mut sp3d = Sp3d::default();
    // 解析 sp3d 自身属性
    for a in start.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match a.key.as_ref() {
            b"extrusionH" => sp3d.extrusion_h = v.parse().unwrap_or(0),
            b"contourW" => sp3d.contour_w = v.parse().unwrap_or(0),
            // v 是 Cow<'_, str>，from_str 期望 &str，需要 &v
            b"prstMaterial" => sp3d.prst_material = MaterialPreset::parse(&v),
            _ => {}
        }
    }
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                // 自闭合 <a:bevelT w="..." h="..."/> 等
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"bevelT" => {
                        sp3d.bevel_top = Some(parse_bevel_attrs(&e));
                    }
                    b"bevelB" => {
                        sp3d.bevel_bottom = Some(parse_bevel_attrs(&e));
                    }
                    _ => {}
                }
            }
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"bevelT" => {
                        let bvl = parse_bevel_attrs(&e);
                        let _ = collect_full_element(rd, e.into_owned());
                        sp3d.bevel_top = Some(bvl);
                    }
                    b"bevelB" => {
                        let bvl = parse_bevel_attrs(&e);
                        let _ = collect_full_element(rd, e.into_owned());
                        sp3d.bevel_bottom = Some(bvl);
                    }
                    b"extrusionClr" => {
                        // 解析子元素 srgbClr / schemeClr
                        let color = parse_color_child(rd, b"extrusionClr")?;
                        sp3d.extrusion_color = Some(color);
                    }
                    b"contourClr" => {
                        let color = parse_color_child(rd, b"contourClr")?;
                        sp3d.contour_color = Some(color);
                    }
                    _ => {
                        let _ = collect_full_element(rd, e.into_owned());
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"sp3d" {
                    return Ok(sp3d);
                }
            }
            Ok(Event::Eof) => return Ok(sp3d),
            Err(e) => return Err(crate::Error::Xml(format!("sp3d parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 从 `<a:bevelT>` / `<a:bevelB>` 的 Start/Empty 事件提取 w/h 属性。
fn parse_bevel_attrs(e: &quick_xml::events::BytesStart<'_>) -> Bevel {
    let mut bvl = Bevel::default();
    for a in e.attributes().flatten() {
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match a.key.as_ref() {
            b"w" => bvl.w = v.parse().unwrap_or(0),
            b"h" => bvl.h = v.parse().unwrap_or(0),
            _ => {}
        }
    }
    bvl
}

/// 从 `<a:outerShdw>` / `<a:innerShdw>` 的 Start/Empty 事件提取属性。
///
/// `is_inner` 为 true 时忽略 `rotWithShape`（仅 outerShdw 有此属性）。
fn parse_shadow_attrs(e: &quick_xml::events::BytesStart<'_>, is_inner: bool) -> ShadowEffect {
    let mut shdw = ShadowEffect::default();
    for a in e.attributes().flatten() {
        let key = a.key.as_ref();
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        match key {
            b"dir" => shdw.dir = v.parse().unwrap_or(0),
            b"dist" => shdw.dist = v.parse().unwrap_or(0),
            b"blurRad" => shdw.blur_rad = v.parse().unwrap_or(0),
            b"rotWithShape" if !is_inner => {
                shdw.rot_with_shape = Some(v == "1" || v == "true");
            }
            _ => {}
        }
    }
    shdw
}

/// 从 `<a:bodyPr ...>...</a:bodyPr>` 完整 XML 字符串解析 [`BodyProperties`]。
///
/// 提取属性：`numCol` / `spcCol` / `anchor` / `wrap` / `lIns`/`tIns`/`rIns`/`bIns` /
/// `vert` / `rot`，以及子元素 `a:spAutoFit` / `a:normAutofit` 标志位。
fn parse_body_pr(xml: &str) -> crate::Result<BodyProperties> {
    let mut bp = BodyProperties::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"bodyPr" {
                    // 根元素：解析属性
                    apply_body_pr_attrs(&mut bp, &e);
                } else if local == b"spAutoFit" {
                    bp.sp_auto_fit = true;
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else if local == b"normAutofit" {
                    bp.norm_autofit = true;
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"bodyPr" {
                    apply_body_pr_attrs(&mut bp, &e);
                } else if local == b"spAutoFit" {
                    bp.sp_auto_fit = true;
                } else if local == b"normAutofit" {
                    bp.norm_autofit = true;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("bodyPr parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(bp)
}

/// 从自闭合 `<a:bodyPr .../>` 事件解析 [`BodyProperties`]。
fn parse_body_pr_attrs(e: &quick_xml::events::BytesStart<'_>) -> BodyProperties {
    let mut bp = BodyProperties::default();
    apply_body_pr_attrs(&mut bp, e);
    bp
}

/// 把 `<a:bodyPr>` 元素的属性应用到 [`BodyProperties`]。
///
/// 支持的属性：`numCol` / `spcCol` / `anchor` / `wrap` / `vert` / `rot` /
/// `lIns` / `tIns` / `rIns` / `bIns`。
fn apply_body_pr_attrs(bp: &mut BodyProperties, e: &quick_xml::events::BytesStart<'_>) {
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"numCol" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                bp.num_cols = v.parse::<u32>().ok();
            }
            b"spcCol" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(n) = v.parse::<i64>() {
                    bp.col_spacing = Some(Emu(n));
                }
            }
            b"anchor" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                bp.anchor = match v.as_ref() {
                    "t" => Some(MsoAnchor::Top),
                    "ctr" => Some(MsoAnchor::Middle),
                    "b" => Some(MsoAnchor::Bottom),
                    _ => None,
                };
            }
            b"wrap" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                bp.wrap = match v.as_ref() {
                    "square" => Some(TextWrapping::Square),
                    "none" => Some(TextWrapping::None),
                    _ => None,
                };
            }
            b"vert" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string();
                bp.vertical = Some(v);
            }
            b"rot" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                bp.rotation = v.parse::<i32>().ok();
            }
            b"lIns" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(n) = v.parse::<i64>() {
                    bp.insets.get_or_insert_with(Default::default).left = Emu(n);
                }
            }
            b"tIns" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(n) = v.parse::<i64>() {
                    bp.insets.get_or_insert_with(Default::default).top = Emu(n);
                }
            }
            b"rIns" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(n) = v.parse::<i64>() {
                    bp.insets.get_or_insert_with(Default::default).right = Emu(n);
                }
            }
            b"bIns" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(n) = v.parse::<i64>() {
                    bp.insets.get_or_insert_with(Default::default).bottom = Emu(n);
                }
            }
            _ => {}
        }
    }
}

/// 解析 `a:p` 元素（段落）。
pub fn parse_paragraph(xml: &str) -> crate::Result<Paragraph> {
    let mut p = Paragraph::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_ppr = false;
    let mut ppr_done = false;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"p" {
                    // 根元素 <a:p>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 pPr / r 等无法解析）
                } else if local == b"pPr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    p.properties = parse_paragraph_properties(&inner)?;
                    ppr_done = true;
                } else if local == b"r" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    if let Ok(r) = parse_run(&inner) {
                        p.runs.push(r);
                    }
                } else if local == b"fld" {
                    // 字段元素 <a:fld id="..." type="...">
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    if let Ok(f) = parse_field(&inner) {
                        p.fields.push(f);
                    }
                } else if local == b"br" {
                    // 段落内的换行
                    p.runs.push(Run::line_break());
                } else if local == b"endParaRPr" {
                    // 带子元素的 endParaRPr（如 <a:endParaRPr lang="en-US"><a:latin typeface="..."/></a:endParaRPr>）
                    // 收集完整元素后用 parse_run_properties 解析属性 + 子元素
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    p.end_properties = Some(parse_run_properties(&inner)?);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
                in_ppr = false;
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"pPr" {
                    // 自闭合 pPr
                    p.properties = parse_paragraph_properties_attrs(&e);
                    ppr_done = true;
                } else if local == b"endParaRPr" {
                    // 自闭合 endParaRPr
                    p.end_properties = Some(RunProperties::from_bytes_start(&e));
                } else if local == b"fld" {
                    // 自闭合 <a:fld/>（无文本，罕见但合法）
                    if let Some(f) = parse_field_attrs(&e) {
                        p.fields.push(f);
                    }
                }
            }
            Ok(Event::End(_)) => {
                if !ppr_done {
                    in_ppr = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("paragraph parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    let _ = in_ppr;
    Ok(p)
}

fn parse_paragraph_properties(xml: &str) -> crate::Result<ParagraphProperties> {
    let mut ppr = ParagraphProperties::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"pPr" {
                    // 根元素 <a:pPr>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 lnSpc / spcBef 等无法解析）
                } else if local == b"lnSpc" {
                    parse_ln_spc_into(&mut rd, &mut ppr, e.into_owned())?;
                } else if local == b"spcBef" {
                    parse_spc_bef_into(&mut rd, &mut ppr, e.into_owned())?;
                } else if local == b"spcAft" {
                    parse_spc_aft_into(&mut rd, &mut ppr, e.into_owned())?;
                } else if local == b"buNone" {
                    // 无项目符号：保留详细信息到 bullet_style
                    ppr.bullet = false;
                    ppr.bullet_style = Some(BulletStyle::None);
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else if local == b"buChar" {
                    // 自定义字符项目符号：保留 char 属性
                    let bs = parse_bu_char(&e);
                    ppr.bullet = true;
                    ppr.bullet_style = Some(bs);
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else if local == b"buAutoNum" {
                    // 自动编号项目符号：保留 type/startAt 属性
                    let bs = parse_bu_auto_num(&e);
                    ppr.bullet = true;
                    ppr.bullet_style = Some(bs);
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else if local == b"tabLst" {
                    // 制表位列表：解析所有 <a:tab> 子元素
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    ppr.tab_stops = parse_tab_lst(&inner)?;
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"buNone" {
                    ppr.bullet = false;
                    ppr.bullet_style = Some(BulletStyle::None);
                } else if local == b"buChar" {
                    ppr.bullet = true;
                    ppr.bullet_style = Some(parse_bu_char(&e));
                } else if local == b"buAutoNum" {
                    ppr.bullet = true;
                    ppr.bullet_style = Some(parse_bu_auto_num(&e));
                } else if local == b"tab" {
                    // 自闭合 <a:tab/>（在 tabLst 外罕见，但兼容）
                    if let Some(tab) = parse_tab(&e) {
                        ppr.tab_stops.push(tab);
                    }
                } else {
                    apply_ppr_attrs(&e, &mut ppr);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("pPr parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(ppr)
}

fn parse_paragraph_properties_attrs(e: &quick_xml::events::BytesStart<'_>) -> ParagraphProperties {
    let mut ppr = ParagraphProperties::default();
    apply_ppr_attrs(e, &mut ppr);
    ppr
}

/// 从 `<a:buChar char="..."/>` 事件解析 [`BulletStyle::Char`]。
///
/// 提取 `char` 属性（项目符号字符，如 `"•"` / `"▪"` / `"→"`）。
fn parse_bu_char(e: &quick_xml::events::BytesStart<'_>) -> BulletStyle {
    let mut ch = String::new();
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"char" {
            ch = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default()
                .to_string();
        }
    }
    BulletStyle::Char { char: ch }
}

/// 从 `<a:buAutoNum type="..." startAt="..."/>` 事件解析 [`BulletStyle::AutoNum`]。
///
/// 提取 `type` 属性（编号类型，如 `"arabicPeriod"` / `"alphaLcParenR"` / `"romanLcParenBoth"`）
/// 和可选的 `startAt` 属性（起始编号）。
fn parse_bu_auto_num(e: &quick_xml::events::BytesStart<'_>) -> BulletStyle {
    let mut auto_num_type = String::new();
    let mut start_at: Option<u32> = None;
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"type" => {
                auto_num_type = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string();
            }
            b"startAt" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                start_at = v.parse::<u32>().ok();
            }
            _ => {}
        }
    }
    BulletStyle::AutoNum {
        auto_num_type,
        start_at,
    }
}

/// 从 `<a:tabLst>...</a:tabLst>` XML 字符串解析制表位列表。
///
/// 提取所有 `<a:tab pos="..." algn="..."/>` 子元素。
fn parse_tab_lst(xml: &str) -> crate::Result<Vec<TabStop>> {
    let mut tabs = Vec::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tabLst" {
                    // 根元素 <a:tabLst>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有 <a:tab> 子元素）
                } else if local == b"tab" {
                    if let Some(tab) = parse_tab(&e) {
                        tabs.push(tab);
                    }
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tab" {
                    if let Some(tab) = parse_tab(&e) {
                        tabs.push(tab);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("tabLst parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(tabs)
}

/// 从 `<a:tab pos="..." algn="..."/>` 事件解析单个 [`TabStop`]。
///
/// 返回 `None` 表示缺少 `pos` 属性或解析失败。
fn parse_tab(e: &quick_xml::events::BytesStart<'_>) -> Option<TabStop> {
    let mut pos: Option<i64> = None;
    let mut alignment = TabAlignment::Left;
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"pos" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                pos = v.parse::<i64>().ok();
            }
            b"algn" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(al) = v.parse::<TabAlignment>() {
                    alignment = al;
                }
            }
            _ => {}
        }
    }
    pos.map(|p| TabStop {
        pos: Emu(p),
        alignment,
    })
}

/// 从 `<a:fld id="..." type="...">...</a:fld>` XML 字符串解析 [`Field`]。
///
/// 提取 `id` / `type` 属性，以及 `<a:rPr>` 和 `<a:t>` 子元素。
fn parse_field(xml: &str) -> crate::Result<Field> {
    let mut field = Field::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"fld" {
                    // 根元素：提取 id / type 属性
                    apply_field_attrs(&mut field, &e);
                } else if local == b"rPr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    field.properties = parse_run_properties(&inner)?;
                } else if local == b"t" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    field.text = extract_text_content(&inner);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"fld" {
                    apply_field_attrs(&mut field, &e);
                } else if local == b"rPr" {
                    field.properties = RunProperties::from_bytes_start(&e);
                }
                // <a:t/> 自闭合（空文本）——text 保持默认空字符串
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("fld parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(field)
}

/// 从自闭合 `<a:fld id="..." type="..."/>` 事件解析 [`Field`]。
///
/// 仅提取属性，text 为空。
fn parse_field_attrs(e: &quick_xml::events::BytesStart<'_>) -> Option<Field> {
    let mut field = Field::default();
    let mut has_type = false;
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"id" => {
                field.id = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string();
            }
            b"type" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                field.field_type = FieldType::from_str_value(&v);
                has_type = true;
            }
            _ => {}
        }
    }
    if has_type {
        Some(field)
    } else {
        None
    }
}

/// 把 `<a:fld>` 元素的属性应用到 [`Field`]。
fn apply_field_attrs(field: &mut Field, e: &quick_xml::events::BytesStart<'_>) {
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"id" => {
                field.id = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string();
            }
            b"type" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                field.field_type = FieldType::from_str_value(&v);
            }
            _ => {}
        }
    }
}

/// 从 `<a:t>...</a:t>` XML 字符串提取纯文本内容。
fn extract_text_content(xml: &str) -> String {
    let mut text = String::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                text.push_str(&t.decode().unwrap_or_default());
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    text
}

fn apply_ppr_attrs(e: &quick_xml::events::BytesStart<'_>, ppr: &mut ParagraphProperties) {
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"algn" => {
                let v = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default();
                if let Ok(al) = v.parse::<Alignment>() {
                    ppr.alignment = Some(al);
                }
            }
            b"lvl" => {
                if let Ok(v) = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .parse::<u8>()
                {
                    ppr.level = v;
                }
            }
            b"indent" | b"marL" => {
                if let Ok(v) = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .parse::<i64>()
                {
                    ppr.indent.left = Some(Emu(v));
                }
            }
            #[allow(unreachable_patterns)]
            b"indent" | b"r" => {
                if let Ok(v) = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .parse::<i64>()
                {
                    ppr.indent.right = Some(Emu(v));
                }
            }
            #[allow(unreachable_patterns)]
            b"indent" | b"firstLine" => {
                if let Ok(v) = a
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .parse::<i64>()
                {
                    ppr.indent.first_line = Some(Emu(v));
                }
            }
            _ => {}
        }
    }
}

fn parse_ln_spc_into<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    ppr: &mut ParagraphProperties,
    _start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<()> {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"spcPct" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<i32>()
                            {
                                ppr.line_spacing_pct = Some(v);
                            }
                        }
                    }
                } else if local == b"spcPts" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<i32>()
                            {
                                ppr.line_spacing = Some(v);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"lnSpc" {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

fn parse_spc_bef_into<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    ppr: &mut ParagraphProperties,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<()> {
    let _ = start;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"spcPts" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<i32>()
                            {
                                ppr.space_before = Some(Emu(v as i64 * 100)); // 1pt = 100 EMU
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"spcBef" {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

fn parse_spc_aft_into<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    ppr: &mut ParagraphProperties,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<()> {
    let _ = start;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"spcPts" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<i32>()
                            {
                                ppr.space_after = Some(Emu(v as i64 * 100));
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if local_name(name.as_ref()) == b"spcAft" {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `a:r` 元素（Run）。
pub fn parse_run(xml: &str) -> crate::Result<Run> {
    let mut run = Run::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rpr_props: Option<RunProperties> = None;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"r" {
                    // 根元素 <a:r>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 rPr / t 等无法解析）
                } else if local == b"rPr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    rpr_props = Some(parse_run_properties(&inner)?);
                } else if local == b"t" {
                    // 文本内容
                    let txt = collect_inner_text_only(&mut rd)?;
                    run.text = txt;
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"rPr" {
                    // 自闭合 rPr
                    rpr_props = Some(RunProperties::from_bytes_start(&e));
                } else if local == b"t" {
                    // 自闭合：空文本
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("run parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    if let Some(p) = rpr_props {
        run.properties = p;
    }
    Ok(run)
}

/// 收集当前 Start 之后的**所有**事件（包括子元素）直到匹配的 End 出现。
/// 返回完整 XML 字符串。
pub fn collect_full_element<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<String> {
    let mut out = String::new();
    // 写开始标签
    out.push('<');
    // 绑定到本地变量以延长临时值的生命周期（quick-xml 0.40 的 `name()` 返回
    // 借用的 `QName<'_>`，直接 .as_ref() 链式写法会被新版 borrow checker 拒绝）。
    let name = start.name();
    out.push_str(&String::from_utf8_lossy(name.as_ref()));
    for a in start.attributes().flatten() {
        out.push(' ');
        out.push_str(&String::from_utf8_lossy(a.key.as_ref()));
        out.push_str("=\"");
        out.push_str(
            &a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default(),
        );
        out.push('"');
    }
    out.push('>');
    // 递归
    let mut depth = 1i32;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                depth += 1;
                write_event_start(&mut out, &e);
            }
            Ok(Event::End(e)) => {
                depth -= 1;
                write_event_end(&mut out, &e);
                if depth == 0 {
                    return Ok(out);
                }
            }
            Ok(Event::Empty(e)) => {
                write_event_empty(&mut out, &e);
            }
            Ok(Event::Text(t)) => {
                out.push_str(&t.decode().unwrap_or_default());
            }
            Ok(Event::CData(c)) => {
                // 直接拼字符串，不走 `write!`（String 不实现 `std::io::Write`）。
                out.push_str("<![CDATA[");
                let cd = String::from_utf8_lossy(c.as_ref());
                out.push_str(&cd);
                out.push_str("]]>");
            }
            Ok(Event::Eof) => return Ok(out),
            Err(e) => return Err(crate::Error::Xml(format!("collect: {e}"))),
            _ => {}
        }
        buf.clear();
    }
}

/// 与 [`collect_full_element`] 类似，但**不**保留原始 XML，只跳过事件。
/// 用于"我们知道这个元素不需要解析、但仍要消费完以保持 reader 状态正确"的场景。
pub fn collect_full_element_skipping<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    _start: quick_xml::events::BytesStart<'static>,
) {
    let mut depth = 1i32;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            _ => {}
        }
        buf.clear();
    }
}

fn write_event_start(out: &mut String, e: &quick_xml::events::BytesStart<'_>) {
    out.push('<');
    // 同 `collect_full_element`：绑定到本地变量延长 `QName<'_>` 借用的生命周期。
    let name = e.name();
    out.push_str(&String::from_utf8_lossy(name.as_ref()));
    for a in e.attributes().flatten() {
        out.push(' ');
        out.push_str(&String::from_utf8_lossy(a.key.as_ref()));
        out.push_str("=\"");
        out.push_str(
            &a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default(),
        );
        out.push('"');
    }
    out.push('>');
}

fn write_event_empty(out: &mut String, e: &quick_xml::events::BytesStart<'_>) {
    out.push('<');
    let name = e.name();
    out.push_str(&String::from_utf8_lossy(name.as_ref()));
    for a in e.attributes().flatten() {
        out.push(' ');
        out.push_str(&String::from_utf8_lossy(a.key.as_ref()));
        out.push_str("=\"");
        out.push_str(
            &a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default(),
        );
        out.push('"');
    }
    out.push_str("/>");
}

fn write_event_end(out: &mut String, e: &quick_xml::events::BytesEnd<'_>) {
    out.push_str("</");
    let name = e.name();
    out.push_str(&String::from_utf8_lossy(name.as_ref()));
    out.push('>');
}

/// 解析 `a:rPr` 元素（含子元素如 `<a:solidFill>` 等）。
pub fn parse_run_properties(xml: &str) -> crate::Result<RunProperties> {
    let mut rp = RunProperties::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"rPr" || local == b"endParaRPr" {
                    // 根元素 <a:rPr> 或 <a:endParaRPr>：
                    // 1. 解析起始标签上的属性（sz/b/i/u/strike/lang/...）
                    // 2. 仅标记进入，不调用 collect_full_element
                    //    （否则会吞掉所有子元素，导致 solidFill / latin 等无法解析）
                    // 先把起始标签属性读入 rp，后续子元素解析会补充 color/font 等
                    rp = RunProperties::from_bytes_start(&e);
                } else if local == b"solidFill" {
                    rp.color = parse_solid_fill(&mut rd, e.into_owned())?;
                } else if local == b"highlight" {
                    let color = parse_solid_fill(&mut rd, e.into_owned())?;
                    rp.highlight = Some(color);
                } else if local == b"latin" {
                    rp.latin_font = Some(collect_attr_value(&mut rd, "typeface")?);
                } else if local == b"ea" {
                    rp.eastasia_font = Some(collect_attr_value(&mut rd, "typeface")?);
                } else if local == b"cs" {
                    rp.cs_font = Some(collect_attr_value(&mut rd, "typeface")?);
                } else if local == b"hlinkClick" {
                    // 超链接（点击）：提取 r:id / tooltip / action 属性，然后跳过子元素
                    let hl = parse_hyperlink_attrs(&e);
                    collect_full_element_skipping(&mut rd, e.into_owned());
                    rp.hlink_click = Some(hl);
                } else if local == b"hlinkHover" {
                    // 超链接（悬停）：提取 r:id / tooltip / action 属性，然后跳过子元素
                    let hl = parse_hyperlink_attrs(&e);
                    collect_full_element_skipping(&mut rd, e.into_owned());
                    rp.hlink_hover = Some(hl);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"latin" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"typeface" {
                            rp.latin_font = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                } else if local == b"ea" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"typeface" {
                            rp.eastasia_font = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                } else if local == b"cs" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"typeface" {
                            rp.cs_font = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                } else if local == b"hlinkClick" {
                    // 自闭合的超链接（点击）：<a:hlinkClick r:id="..." tooltip="..."/>
                    rp.hlink_click = Some(parse_hyperlink_attrs(&e));
                } else if local == b"hlinkHover" {
                    // 自闭合的超链接（悬停）：<a:hlinkHover r:id="..." tooltip="..."/>
                    rp.hlink_hover = Some(parse_hyperlink_attrs(&e));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("rPr parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(rp)
}

/// 从 `<a:hlinkClick>` / `<a:hlinkHover>` 事件中提取属性，构造 [`Hyperlink`]。
///
/// 提取的属性：
/// - `r:id`（或任意 `:id` 后缀）：关系 ID，指向 `.rels` 中的目标 URL；
/// - `tooltip`：鼠标悬停提示；
/// - `action`：动作类型（如 `"ppaction://hlinksldjump"` 跳转幻灯片）。
///
/// 若三个属性均缺失，返回的 `Hyperlink` 的 `invalid` 标记为 `true`（用于继承场景）。
fn parse_hyperlink_attrs(e: &quick_xml::events::BytesStart<'_>) -> Hyperlink {
    let mut hl = Hyperlink::default();
    let mut has_any = false;
    for a in e.attributes().flatten() {
        let key = a.key.as_ref();
        let val = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default()
            .to_string();
        // r:id 可能以不同前缀出现（r:id / ns0:id 等），统一用后缀匹配
        if key == b"r:id" || key.ends_with(b":id") {
            hl.rid = Some(val);
            has_any = true;
        } else if key == b"tooltip" {
            hl.tooltip = Some(val);
            has_any = true;
        } else if key == b"action" {
            hl.action = Some(val);
            has_any = true;
        }
    }
    if !has_any {
        hl.invalid = true;
    }
    hl
}

/// 从 `<a:tile>` 事件中提取属性，构造 [`BlipFillMode::Tile`]。
///
/// 提取的属性：
/// - `tx` / `ty`：水平/垂直偏移（EMU）；
/// - `sx` / `sy`：水平/垂直缩放（千分比，100000 = 100%）；
/// - `flip`：翻转模式（`"none"` / `"x"` / `"y"` / `"xy"`）；
/// - `algn`：对齐方式（`"tl"` / `"ctr"` / `"br"` 等）。
fn parse_tile_attrs(e: &quick_xml::events::BytesStart<'_>) -> crate::oxml::sppr::BlipFillMode {
    let mut tx = None;
    let mut ty = None;
    let mut sx = None;
    let mut sy = None;
    let mut flip = None;
    let mut algn = None;
    for a in e.attributes().flatten() {
        let key = a.key.as_ref();
        let val = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default()
            .to_string();
        match key {
            b"tx" => tx = val.parse().ok(),
            b"ty" => ty = val.parse().ok(),
            b"sx" => sx = val.parse().ok(),
            b"sy" => sy = val.parse().ok(),
            b"flip" => flip = Some(val),
            b"algn" => algn = Some(val),
            _ => {}
        }
    }
    crate::oxml::sppr::BlipFillMode::Tile {
        tx,
        ty,
        sx,
        sy,
        flip,
        algn,
    }
}

/// 从 `<a:spLocks>` 事件中提取锁定属性，构造 [`ShapeLocks`]。
///
/// 所有属性均为布尔值，值为 `"1"` 或 `"true"` 时设为 `true`。
fn parse_sp_locks_attrs(e: &quick_xml::events::BytesStart<'_>) -> ShapeLocks {
    let mut locks = ShapeLocks::default();
    for a in e.attributes().flatten() {
        let key = a.key.as_ref();
        let v = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .unwrap_or_default();
        let is_true = v == "1" || v == "true";
        match key {
            b"noGrp" => locks.no_grp = is_true,
            b"noDrilldown" => locks.no_drilldown = is_true,
            b"noSelect" => locks.no_select = is_true,
            b"noChangeAspect" => locks.no_change_aspect = is_true,
            b"noMove" => locks.no_move = is_true,
            b"noResize" => locks.no_resize = is_true,
            b"noRot" => locks.no_rot = is_true,
            b"noEditPoints" => locks.no_edit_points = is_true,
            b"noAdjustHandles" => locks.no_adjust_handles = is_true,
            b"noChangeArrowheads" => locks.no_change_arrowheads = is_true,
            b"noChangeShapeType" => locks.no_change_shape_type = is_true,
            b"noCrop" => locks.no_crop = is_true,
            _ => {}
        }
    }
    locks
}

/// 解析 `<p:style>` XML 为 [`ShapeStyle`]。
///
/// `<p:style>` 包含 4 个子元素（按 OOXML 顺序）：
/// - `<a:lnRef idx="..." schemeClr="..."/>` 线条样式引用
/// - `<a:fillRef idx="..." schemeClr="..."/>` 填充样式引用
/// - `<a:effectRef idx="..." schemeClr="..."/>` 效果样式引用
/// - `<a:fontRef idx="minor|major" schemeClr="..."/>` 字体样式引用
pub fn parse_shape_style(xml: &str) -> crate::Result<ShapeStyle> {
    let mut style = ShapeStyle::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"lnRef" => {
                        style.line_ref = Some(parse_style_ref(&mut rd, &e)?);
                    }
                    b"fillRef" => {
                        style.fill_ref = Some(parse_style_ref(&mut rd, &e)?);
                    }
                    b"effectRef" => {
                        style.effect_ref = Some(parse_style_ref(&mut rd, &e)?);
                    }
                    b"fontRef" => {
                        style.font_ref = Some(parse_style_ref(&mut rd, &e)?);
                    }
                    b"style" => {
                        // 根元素 <p:style>：仅标记进入，不调用 collect_full_element
                    }
                    _ => {
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"lnRef" => {
                        style.line_ref = Some(parse_style_ref_empty(&e));
                    }
                    b"fillRef" => {
                        style.fill_ref = Some(parse_style_ref_empty(&e));
                    }
                    b"effectRef" => {
                        style.effect_ref = Some(parse_style_ref_empty(&e));
                    }
                    b"fontRef" => {
                        style.font_ref = Some(parse_style_ref_empty(&e));
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("style parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(style)
}

/// 解析 `<a:lnRef>` / `<a:fillRef>` / `<a:effectRef>` / `<a:fontRef>` 的 Start 事件。
///
/// 提取 `idx` 属性和子元素 `<a:schemeClr val="..."/>`。
fn parse_style_ref<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    e: &quick_xml::events::BytesStart<'_>,
) -> crate::Result<StyleRef> {
    let mut sr = parse_style_ref_empty(e);
    // 读取子元素直到匹配的 End
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e2)) => {
                let name2 = e2.name();
                let local = local_name(name2.as_ref());
                if local == b"schemeClr" {
                    for a in e2.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            sr.scheme_color = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                }
            }
            Ok(Event::End(_)) => return Ok(sr),
            Ok(Event::Eof) => return Ok(sr),
            _ => {}
        }
        buf.clear();
    }
}

/// 从 Empty 事件提取 StyleRef 属性（idx）。
fn parse_style_ref_empty(e: &quick_xml::events::BytesStart<'_>) -> StyleRef {
    let mut sr = StyleRef::default();
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"idx" {
            sr.idx = Some(
                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .unwrap_or_default()
                    .to_string(),
            );
        }
    }
    sr
}

fn collect_attr_value<R: std::io::BufRead>(rd: &mut Reader<R>, key: &str) -> crate::Result<String> {
    // 期望的语义：在当前层找到一个自闭合元素 `<... key="value"/>` 并返回 `value`。
    // 常见用法：`<a:latin typeface="Calibri"/>` / `<a:ea typeface="宋体"/>` / `<a:cs typeface="Times"/>`。
    //
    // 早期实现把"元素名 == key"作为匹配条件，导致**永远匹配不上**（元素名是 latin/ea/cs，
    // key 是 typeface）。修正后：忽略元素名，只看属性 key 是否命中。
    let mut buf = Vec::new();
    let mut val = String::new();
    let key_b = key.as_bytes();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                for a in e.attributes().flatten() {
                    if a.key.as_ref() == key_b {
                        val = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .to_string();
                    }
                }
            }
            Ok(Event::End(_)) => return Ok(val),
            Ok(Event::Eof) => return Ok(val),
            _ => {}
        }
        buf.clear();
    }
}

/// 取 quick-xml 元素名的**本地部分**（去前缀）。
///
/// 历史上存在的 `local_name_check` noop 包装已被删除（它没有做任何"去前缀"工作，
/// 只是把字节返回），保留仅是历史包袱。
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&c| c == b':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

/// 仅收集 Text 内容（用于 `<a:t>...</a:t>`）。
fn collect_inner_text_only<R: std::io::BufRead>(rd: &mut Reader<R>) -> crate::Result<String> {
    let mut out = String::new();
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => {
                out.push_str(&t.decode().unwrap_or_default());
            }
            Ok(Event::CData(c)) => {
                out.push_str(&String::from_utf8_lossy(c.as_ref()));
            }
            Ok(Event::End(_)) => return Ok(out),
            Ok(Event::Eof) => return Ok(out),
            _ => {}
        }
        buf.clear();
    }
}

impl RunProperties {
    /// 从 `a:rPr` 的自闭合属性中解析（不递归子元素）。
    pub fn from_bytes_start(e: &quick_xml::events::BytesStart<'_>) -> Self {
        let mut rp = RunProperties::default();
        for a in e.attributes().flatten() {
            let v = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default();
            match a.key.as_ref() {
                b"sz" => {
                    if let Ok(p) = v.parse::<i32>() {
                        rp.size = Some(Pt(p as f64 / 100.0));
                    }
                }
                b"b" => {
                    if v == "1" || v == "true" {
                        rp.bold = true;
                    }
                }
                b"i" => {
                    if v == "1" || v == "true" {
                        rp.italic = true;
                    }
                }
                b"u" => {
                    if let Ok(u) = v.parse::<Underline>() {
                        rp.underline = Some(u);
                    }
                }
                b"strike" => {
                    if v == "sngStrike" {
                        rp.strike = true;
                    }
                    if v == "dblStrike" {
                        rp.strike_dbl = true;
                    }
                }
                b"baseline" => {
                    if let Ok(v) = v.parse::<i32>() {
                        rp.baseline = Some(v);
                    }
                }
                b"kern" => {
                    if let Ok(v) = v.parse::<i32>() {
                        rp.kerning = Some(v);
                    }
                }
                b"spc" => {
                    if let Ok(v) = v.parse::<i32>() {
                        rp.spc = Some(v);
                    }
                }
                b"lang" => {
                    rp.lang = Some(v.to_string());
                }
                _ => {}
            }
        }
        rp
    }
}

/// 解析 `p:pic` 元素（仅最小字段；其它信息丢失）。
///
/// 走 **正规 SAX**：在 `<p:blipFill>` 内逐事件读 `<a:blip r:embed="..."/>`
/// 等自闭合元素；不再使用"字符串 find r:embed"——后者在缩进/属性顺序变化时会漏取。
pub fn parse_pic(xml: &str) -> crate::Result<Pic> {
    let mut pic = Pic::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_blipfill = false;
    let mut blipfill_depth: i32 = 0;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"pic" {
                    // 根元素 <p:pic>：仅标记进入，不调用 collect_full_element
                    // （否则会吞掉所有子元素，导致 nvPicPr / blipFill 等无法解析）
                } else if local == b"nvPicPr" {
                    // 不使用 collect_full_element 吞掉 nvPicPr，
                    // 而是手动遍历子元素提取 cNvPr 的 id/name 属性
                    let mut nv_depth = 1i32;
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Start(_)) => nv_depth += 1,
                            Ok(Event::End(_)) => {
                                nv_depth -= 1;
                                if nv_depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Empty(e2)) => {
                                let name2 = e2.name();
                                let local2 = local_name(name2.as_ref());
                                if local2 == b"cNvPr" {
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                if let Ok(v) = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse::<u32>()
                                                {
                                                    pic.id = v;
                                                }
                                            }
                                            b"name" => {
                                                pic.name = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .to_string();
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"blipFill" {
                    in_blipfill = true;
                    blipfill_depth = 1;
                    // 不立即吞掉，改为手写 SAX 解析 blipFill 的子元素
                    // 让 Event::Start/Empty 单独触发下方的 blip 处理
                } else if local == b"spPr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    pic.properties = parse_sppr(&inner)?;
                } else if in_blipfill && local == b"blip" {
                    // `<a:blip r:embed="rIdN"/>` 可能是 Start+End 也可能是 Empty。
                    // 直接从当前事件读取属性，**不再**走 collect_full_element。
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        // 兼容 `r:embed` 与 `embed` 两种命名空间形态
                        if key == b"r:embed" || key.ends_with(b":embed") {
                            pic.rid = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if in_blipfill && local == b"srcRect" {
                    // `<a:srcRect l="..." t="..." r="..." b="..."/>`
                    let (mut l, mut t, mut r, mut b) = (None, None, None, None);
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"l" => l = v.parse().ok(),
                            b"t" => t = v.parse().ok(),
                            b"r" => r = v.parse().ok(),
                            b"b" => b = v.parse().ok(),
                            _ => {}
                        }
                    }
                    if let (Some(ll), Some(tt), Some(rr), Some(bb)) = (l, t, r, b) {
                        pic.src_rect = Some((ll, tt, rr, bb));
                    }
                } else if in_blipfill && local == b"stretch" {
                    // `<a:stretch><a:fillRect/></a:stretch>`：拉伸填充模式
                    let _ = collect_full_element(&mut rd, e.into_owned())?;
                    pic.fill_mode = crate::oxml::sppr::BlipFillMode::Stretch;
                } else if in_blipfill && local == b"tile" {
                    // `<a:tile tx="..." ty="..." sx="..." sy="..." flip="..." algn="..."/>`
                    // 先解析属性，再吞掉子元素（tile 很少有子元素，但保险起见）
                    let mode = parse_tile_attrs(&e);
                    let _ = collect_full_element(&mut rd, e.into_owned())?;
                    pic.fill_mode = mode;
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_blipfill && local == b"blip" {
                    // 自闭合形态的 `<a:blip r:embed="rIdN"/>`
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        if key == b"r:embed" || key.ends_with(b":embed") {
                            pic.rid = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if in_blipfill && local == b"srcRect" {
                    let (mut l, mut t, mut r, mut b) = (None, None, None, None);
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"l" => l = v.parse().ok(),
                            b"t" => t = v.parse().ok(),
                            b"r" => r = v.parse().ok(),
                            b"b" => b = v.parse().ok(),
                            _ => {}
                        }
                    }
                    if let (Some(ll), Some(tt), Some(rr), Some(bb)) = (l, t, r, b) {
                        pic.src_rect = Some((ll, tt, rr, bb));
                    }
                } else if in_blipfill && local == b"stretch" {
                    // 自闭合 `<a:stretch/>`（罕见，通常带 fillRect 子元素）
                    pic.fill_mode = crate::oxml::sppr::BlipFillMode::Stretch;
                } else if in_blipfill && local == b"tile" {
                    // 自闭合 `<a:tile .../>`
                    pic.fill_mode = parse_tile_attrs(&e);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // blipFill 可能是 `<a:blipFill><a:blip/>...</a:blipFill>` 这种
                // 形式：进入 blipFill 时 depth=1，遇到结束事件 depth-1。
                // depth 归 0 即离开 blipFill 范围。
                if in_blipfill && local == b"blipFill" {
                    blipfill_depth -= 1;
                    if blipfill_depth == 0 {
                        in_blipfill = false;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("pic parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(pic)
}

/// 解析 `<p:cxnSp>` 元素（连接器）。
///
/// # 元素结构（OOXML）
///
/// ```text
/// <p:cxnSp>
///   <p:nvCxnSpPr>
///     <p:cNvPr id="..." name="..."/>
///     <p:cNvCxnSpPr/>
///     <p:nvPr/>
///     <p:stCxn id="..." idx="..."/>   ← 可选：起点挂接
///     <p:endCxn id="..." idx="..."/>  ← 可选：终点挂接
///   </p:nvCxnSpPr>
///   <p:spPr>...</p:spPr>
///   <p:style>...</p:style>           ← 可选
///   <p:extLst>...</p:extLst>         ← 可选
/// </p:cxnSp>
/// ```
///
/// # 行为
/// - 提取 `id` / `name` 属性；
/// - 解析 `stCxn` / `endCxn` 挂接信息（`Some((shape_id, idx))`）；
/// - 解析 `spPr`（变换 / 几何 / 填充 / 边框）；
/// - `style` / `extLst` 暂不解析（保留 `None`）。
pub fn parse_cxn_sp(xml: &str) -> crate::Result<Connector> {
    let mut cxn = Connector::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"cxnSp" {
                    // 根元素 <p:cxnSp>：仅标记进入
                } else if local == b"nvCxnSpPr" {
                    // 手动遍历 nvCxnSpPr 子元素提取 cNvPr / stCxn / endCxn
                    let mut nv_depth = 1i32;
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Start(_)) => nv_depth += 1,
                            Ok(Event::End(_)) => {
                                nv_depth -= 1;
                                if nv_depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Empty(e2)) => {
                                let name2 = e2.name();
                                let local2 = local_name(name2.as_ref());
                                if local2 == b"cNvPr" {
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                if let Ok(v) = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse::<u32>()
                                                {
                                                    cxn.id = v;
                                                }
                                            }
                                            b"name" => {
                                                cxn.name = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .to_string();
                                            }
                                            _ => {}
                                        }
                                    }
                                } else if local2 == b"stCxn" {
                                    // 起点挂接：<p:stCxn id="..." idx="..."/>
                                    let mut sid: Option<u32> = None;
                                    let mut idx: Option<u32> = None;
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                sid = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse()
                                                    .ok();
                                            }
                                            b"idx" => {
                                                idx = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse()
                                                    .ok();
                                            }
                                            _ => {}
                                        }
                                    }
                                    if let (Some(s), Some(i)) = (sid, idx) {
                                        cxn.st_cxn = Some((s, i));
                                    }
                                } else if local2 == b"endCxn" {
                                    // 终点挂接：<p:endCxn id="..." idx="..."/>
                                    let mut sid: Option<u32> = None;
                                    let mut idx: Option<u32> = None;
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                sid = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse()
                                                    .ok();
                                            }
                                            b"idx" => {
                                                idx = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse()
                                                    .ok();
                                            }
                                            _ => {}
                                        }
                                    }
                                    if let (Some(s), Some(i)) = (sid, idx) {
                                        cxn.end_cxn = Some((s, i));
                                    }
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"spPr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    cxn.properties = parse_sppr(&inner)?;
                } else if local == b"style" {
                    // <p:style> 主题样式引用（TODO-006 一致性补全）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    cxn.style = Some(parse_shape_style(&inner)?);
                } else {
                    // extLst 等暂不解析，吞掉
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"cxnSp" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("cxnSp parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(cxn)
}

/// 解析 `<p:grpSp>` 元素（组合形状，递归）。
///
/// # 元素结构（OOXML）
///
/// ```text
/// <p:grpSp>
///   <p:nvGrpSpPr>
///     <p:cNvPr id="..." name="..."/>
///     <p:cNvGrpSpPr/>
///     <p:nvPr/>
///   </p:nvGrpSpPr>
///   <p:grpSpPr>
///     <a:xfrm>
///       <a:off x="..." y="..."/>
///       <a:ext cx="..." cy="..."/>
///       <a:chOff x="..." y="..."/>
///       <a:chExt cx="..." cy="..."/>
///     </a:xfrm>
///   </p:grpSpPr>
///   子形状 (sp/pic/cxnSp/grpSp/graphicFrame)   ← 递归
///   <p:extLst>...</p:extLst>                  ← 可选
/// </p:grpSp>
/// ```
///
/// # 行为
/// - 提取 `id` / `name`；
/// - 解析 `grpSpPr` 内的 `a:xfrm`（off/ext/chOff/chExt）；
/// - 递归解析子形状（sp/pic/cxnSp/grpSp/graphicFrame）。
pub fn parse_grp_sp(xml: &str) -> crate::Result<Group> {
    let mut grp = Group::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    // 收集子形状的原始 XML，稍后递归解析
    let mut child_sp_bufs: Vec<String> = Vec::new();
    let mut child_pic_bufs: Vec<String> = Vec::new();
    let mut child_cxn_bufs: Vec<String> = Vec::new();
    let mut child_grp_bufs: Vec<String> = Vec::new();
    let mut child_gfx_bufs: Vec<String> = Vec::new();
    // 标记是否已遇到根 <p:grpSp>（用于区分根元素与嵌套 grpSp）
    let mut entered_root = false;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"grpSp" {
                    if !entered_root {
                        // 根元素 <p:grpSp>：仅标记进入
                        entered_root = true;
                    } else {
                        // 嵌套组合：递归
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        child_grp_bufs.push(inner);
                    }
                } else if local == b"nvGrpSpPr" {
                    // 手动遍历 nvGrpSpPr 子元素提取 cNvPr
                    let mut nv_depth = 1i32;
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Start(_)) => nv_depth += 1,
                            Ok(Event::End(_)) => {
                                nv_depth -= 1;
                                if nv_depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Empty(e2)) => {
                                let name2 = e2.name();
                                let local2 = local_name(name2.as_ref());
                                if local2 == b"cNvPr" {
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                if let Ok(v) = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse::<u32>()
                                                {
                                                    grp.id = v;
                                                }
                                            }
                                            b"name" => {
                                                grp.name = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .to_string();
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"grpSpPr" {
                    // 解析 grpSpPr 内的 a:xfrm（off/ext/chOff/chExt）
                    parse_grp_sppr_into(&mut rd, &mut grp, e.into_owned())?;
                } else if local == b"sp" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    child_sp_bufs.push(inner);
                } else if local == b"pic" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    child_pic_bufs.push(inner);
                } else if local == b"cxnSp" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    child_cxn_bufs.push(inner);
                } else if local == b"graphicFrame" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    child_gfx_bufs.push(inner);
                } else if local == b"style" {
                    // <p:style> 主题样式引用（TODO-006 一致性补全）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    grp.style = Some(parse_shape_style(&inner)?);
                } else {
                    // extLst 等暂不解析，吞掉
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"grpSp" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("grpSp parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 递归解析子形状
    for s in child_sp_bufs {
        if let Ok(sp) = parse_sp(&s) {
            grp.children.push(GroupChild::Sp(sp));
        }
    }
    for p in child_pic_bufs {
        if let Ok(pic) = parse_pic(&p) {
            grp.children.push(GroupChild::Pic(pic));
        }
    }
    for c in child_cxn_bufs {
        if let Ok(cxn) = parse_cxn_sp(&c) {
            grp.children.push(GroupChild::CxnSp(cxn));
        }
    }
    for g in child_grp_bufs {
        if let Ok(sub) = parse_grp_sp(&g) {
            grp.children.push(GroupChild::Group(Box::new(sub)));
        }
    }
    for f in child_gfx_bufs {
        if let Ok(frame) = parse_graphic_frame(&f) {
            grp.children.push(GroupChild::GraphicFrame(frame));
        }
    }
    Ok(grp)
}

/// 解析 `<p:grpSpPr>` 内的 `<a:xfrm>`，提取 off/ext/chOff/chExt。
fn parse_grp_sppr_into<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    grp: &mut Group,
    _start: quick_xml::events::BytesStart<'static>,
) -> crate::Result<()> {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"off" {
                    let (mut x, mut y) = (None, None);
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"x" => {
                                x = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            b"y" => {
                                y = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            _ => {}
                        }
                    }
                    if let (Some(x), Some(y)) = (x, y) {
                        grp.off = (Emu(x), Emu(y));
                    }
                } else if local == b"ext" {
                    let (mut cx, mut cy) = (None, None);
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"cx" => {
                                cx = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            b"cy" => {
                                cy = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            _ => {}
                        }
                    }
                    if let (Some(cx), Some(cy)) = (cx, cy) {
                        grp.ext = (Emu(cx), Emu(cy));
                    }
                }
                // chOff / chExt 暂不解析（组合内部坐标系，0.1.0 写出时固定为 0,0）
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"grpSpPr" {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<p:graphicFrame>` 元素（图形框：表格/图表）。
///
/// # 元素结构（OOXML）
///
/// ```text
/// <p:graphicFrame>
///   <p:nvGraphicFramePr>
///     <p:cNvPr id="..." name="..."/>
///     <p:cNvGraphicFramePr/>
///     <p:nvPr/>
///   </p:nvGraphicFramePr>
///   <p:xfrm>
///     <a:off x="..." y="..."/>
///     <a:ext cx="..." cy="..."/>
///   </p:xfrm>
///   <a:graphic>
///     <a:graphicData uri="...">
///       <a:tbl>...</a:tbl>     ← 或 chart / smartArt
///     </a:graphicData>
///   </a:graphic>
///   <p:extLst>...</p:extLst>   ← 可选
/// </p:graphicFrame>
/// ```
///
/// # 行为
/// - 提取 `id` / `name`；
/// - 解析 `<p:xfrm>` 内的 off/ext（写入 `properties.xfrm`）；
/// - 解析 `<a:graphicData uri="...">` 内的 `<a:tbl>`（仅表格，其它类型暂留空）。
pub fn parse_graphic_frame(xml: &str) -> crate::Result<GraphicFrame> {
    let mut frame = GraphicFrame::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"graphicFrame" {
                    // 根元素 <p:graphicFrame>：仅标记进入
                } else if local == b"nvGraphicFramePr" {
                    // 手动遍历 nvGraphicFramePr 子元素提取 cNvPr
                    let mut nv_depth = 1i32;
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Start(_)) => nv_depth += 1,
                            Ok(Event::End(_)) => {
                                nv_depth -= 1;
                                if nv_depth == 0 {
                                    break;
                                }
                            }
                            Ok(Event::Empty(e2)) => {
                                let name2 = e2.name();
                                let local2 = local_name(name2.as_ref());
                                if local2 == b"cNvPr" {
                                    for a in e2.attributes().flatten() {
                                        match a.key.as_ref() {
                                            b"id" => {
                                                if let Ok(v) = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .parse::<u32>()
                                                {
                                                    frame.id = v;
                                                }
                                            }
                                            b"name" => {
                                                frame.name = a
                                                    .normalized_value(
                                                        quick_xml::XmlVersion::Implicit1_0,
                                                    )
                                                    .unwrap_or_default()
                                                    .to_string();
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                            Ok(Event::Eof) => break,
                            Err(_) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"xfrm" {
                    // 先提取 xfrm 自身的属性（rot/flipH/flipV）
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"rot" => {
                                if let Ok(v) = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse::<i32>()
                                {
                                    frame.properties.xfrm.rot = Some(v);
                                }
                            }
                            b"flipH" => {
                                let v = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default();
                                if v == "1" {
                                    frame.properties.xfrm.flip_h = true;
                                }
                            }
                            b"flipV" => {
                                let v = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default();
                                if v == "1" {
                                    frame.properties.xfrm.flip_v = true;
                                }
                            }
                            _ => {}
                        }
                    }
                    // 再解析子元素（off/ext）
                    parse_xfrm_into(&mut rd, &mut frame.properties);
                } else if local == b"graphic" {
                    // 解析 <a:graphic> 内的 <a:graphicData uri="...">
                    parse_graphic_into(&mut rd, &mut frame)?;
                } else {
                    // extLst 等暂不解析，吞掉
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"graphicFrame" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("graphicFrame parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(frame)
}

/// 解析 `<a:graphic>` 内的 `<a:graphicData uri="...">`，根据 uri 分发到表格解析器。
fn parse_graphic_into<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    frame: &mut GraphicFrame,
) -> crate::Result<()> {
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"graphicData" {
                    // 读取 uri 属性判断图形类型
                    let mut uri = String::new();
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"uri" {
                            uri = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                    // 根据 uri 分发：表格 / 图表 / smartArt
                    if uri == "http://schemas.openxmlformats.org/drawingml/2006/table" {
                        // 收集 <a:tbl> 子元素
                        loop {
                            match rd.read_event_into(&mut buf) {
                                Ok(Event::Start(e2)) => {
                                    if local_name(e2.name().as_ref()) == b"tbl" {
                                        let tbl_xml = collect_full_element(rd, e2.into_owned())?;
                                        if let Ok(tbl) = parse_table(&tbl_xml) {
                                            frame.graphic = crate::oxml::shape::Graphic::Table(tbl);
                                        }
                                    } else {
                                        let _ = collect_full_element(rd, e2.into_owned());
                                    }
                                }
                                Ok(Event::End(e2)) => {
                                    if local_name(e2.name().as_ref()) == b"graphicData" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                            buf.clear();
                        }
                    } else if uri == "http://schemas.openxmlformats.org/drawingml/2006/diagram" {
                        // SmartArt 最小保留（TODO-037）：保留完整 <a:graphicData> 元素 XML。
                        // collect_full_element 会从当前 Start 事件（graphicData）开始，
                        // 收集到对应 End 事件，返回完整 XML（含外壳）。
                        let raw_xml = collect_full_element(rd, e.into_owned())?;
                        // 从 raw_xml 中提取 dgm:relIds 的 4 个关系 id（r:dm / r:lo / r:qs / r:cs）。
                        // 这些 id 仅供调用方查询，序列化时不单独使用（直接走 raw_xml）。
                        let (dm_rid, lo_rid, qs_rid, cs_rid) = parse_smartart_rel_ids(&raw_xml);
                        frame.graphic = crate::oxml::shape::Graphic::SmartArt(
                            crate::oxml::shape::SmartArtRef {
                                raw_xml,
                                dm_rid,
                                lo_rid,
                                qs_rid,
                                cs_rid,
                            },
                        );
                    } else if uri == "http://schemas.openxmlformats.org/drawingml/2006/chart" {
                        // 图表读路径（TODO-004）：slide 的 graphicFrame 仅含 `<c:chart r:id="rIdX"/>`
                        // 引用，真正的图表内容在独立的 `/ppt/charts/chartN.xml` part 中。
                        // 这里仅提取 r:id 构造 Chart 占位模型，由 `Presentation::from_opc`
                        // 读取 chartN.xml 后调用 `Chart::parse_from_xml` 填充真实类型与数据。
                        let raw_xml = collect_full_element(rd, e.into_owned())?;
                        let rid = parse_chart_rid(&raw_xml);
                        frame.graphic =
                            crate::oxml::shape::Graphic::Chart(crate::oxml::chart::Chart::new(
                                crate::oxml::chart::ChartType::Column,
                                crate::oxml::chart::ChartData::default(),
                            ));
                        // 把 rid 写入 Chart.rid，from_opc 阶段根据 rid 查 chart_rel_map 读 chartN.xml。
                        if let crate::oxml::shape::Graphic::Chart(c) = &mut frame.graphic {
                            c.rid = rid;
                        }
                    } else {
                        // 其它类型（chart 等）暂不解析，吞掉
                        let _ = collect_full_element(rd, e.into_owned());
                    }
                } else {
                    let _ = collect_full_element(rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"graphic" {
                    return Ok(());
                }
            }
            Ok(Event::Eof) => return Ok(()),
            _ => {}
        }
        buf.clear();
    }
}

/// 从 SmartArt 的 `<a:graphicData>` 完整 XML 中提取 `<dgm:relIds>` 的 4 个关系 id（TODO-037）。
///
/// # 元素结构（OOXML）
///
/// ```text
/// <a:graphicData uri=".../diagram">
///   <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
/// </a:graphicData>
/// ```
///
/// # 行为
///
/// - 用简单的字符串查找提取 `r:dm` / `r:lo` / `r:qs` / `r:cs` 属性值；
/// - 不做完整 XML 解析（避免 quick-xml 命名空间处理的复杂性，且 raw_xml 已 byte-exact 保留）；
/// - 找不到时对应字段返回 `None`。
///
/// # 返回值
///
/// `(dm_rid, lo_rid, qs_rid, cs_rid)` 四元组，每个字段为 `Option<String>`。
fn parse_smartart_rel_ids(
    raw_xml: &str,
) -> (
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
) {
    /// 在 XML 字符串中查找 `attr_name="value"` 模式的 value。
    ///
    /// 仅做简单字符串查找，不处理转义（OOXML 关系 id 不含特殊字符，安全）。
    fn find_attr(xml: &str, attr_name: &str) -> Option<String> {
        // 查找 `attr_name="` 模式（注意 r:dm 等带冒号的属性名）
        let pattern = format!("{}=\"", attr_name);
        let pos = xml.find(&pattern)?;
        let value_start = pos + pattern.len();
        let value_end = xml[value_start..].find('"')?;
        Some(xml[value_start..value_start + value_end].to_string())
    }

    let dm = find_attr(raw_xml, "r:dm");
    let lo = find_attr(raw_xml, "r:lo");
    let qs = find_attr(raw_xml, "r:qs");
    let cs = find_attr(raw_xml, "r:cs");
    (dm, lo, qs, cs)
}

/// 从 chart 的 `<a:graphicData>` 完整 XML 中提取 `<c:chart r:id="..."/>` 的关系 id（TODO-004 读路径）。
///
/// # 元素结构（OOXML）
///
/// ```text
/// <a:graphicData uri=".../chart">
///   <c:chart xmlns:c="..." xmlns:r="..." r:id="rId1"/>
/// </a:graphicData>
/// ```
///
/// # 行为
///
/// - 用简单的字符串查找提取 `r:id` 属性值；
/// - 兼容 `r:id="..."` 与 `r:id = "..."`（带空格，罕见）两种写法；
/// - 找不到时返回空字符串（`from_opc` 阶段会跳过该 chart 的 part 读取）。
///
/// # 返回值
///
/// 关系 id 字符串（如 `"rId1"`）。找不到时返回空字符串。
fn parse_chart_rid(raw_xml: &str) -> String {
    // 优先匹配 `r:id="value"` 模式（OOXML 关系 id 不含特殊字符，安全）
    let pattern = "r:id=\"";
    if let Some(pos) = raw_xml.find(pattern) {
        let value_start = pos + pattern.len();
        if let Some(end) = raw_xml[value_start..].find('"') {
            return raw_xml[value_start..value_start + end].to_string();
        }
    }
    String::new()
}

/// 解析 `<a:tbl>` 元素为 [`crate::oxml::table::Table`]。
///
/// # 元素结构
///
/// ```text
/// <a:tbl>
///   <a:tblPr>...</a:tblPr>
///   <a:tblGrid>
///     <a:gridCol w="..."/>
///     ...
///   </a:tblGrid>
///   <a:tr h="...">
///     <a:tc>
///       <a:txBody>...</a:txBody>
///       <a:tcPr anchor="..." marT="..." .../>
///     </a:tc>
///     ...
///   </a:tr>
///   ...
/// </a:tbl>
/// ```
fn parse_table(xml: &str) -> crate::Result<crate::oxml::table::Table> {
    use crate::oxml::table::{Col, Table};

    let mut tbl = Table::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    // 收集行和列的原始 XML
    let mut tr_bufs: Vec<String> = Vec::new();
    let mut grid_col_widths: Vec<Emu> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tbl" {
                    // 根元素
                } else if local == b"tblPr" {
                    // 解析 tblPr 内的 tblLook 属性和 tableStyleId
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    parse_tbl_pr_into(&inner, &mut tbl);
                } else if local == b"tblGrid" {
                    // 解析 tblGrid 内的 gridCol
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Empty(e2)) => {
                                if local_name(e2.name().as_ref()) == b"gridCol" {
                                    let mut w: Option<i64> = None;
                                    for a in e2.attributes().flatten() {
                                        if a.key.as_ref() == b"w" {
                                            w = a
                                                .normalized_value(
                                                    quick_xml::XmlVersion::Implicit1_0,
                                                )
                                                .unwrap_or_default()
                                                .parse()
                                                .ok();
                                        }
                                    }
                                    grid_col_widths.push(Emu(w.unwrap_or(0)));
                                }
                            }
                            Ok(Event::End(e2)) => {
                                if local_name(e2.name().as_ref()) == b"tblGrid" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                } else if local == b"tr" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    tr_bufs.push(inner);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("tbl parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    // 填充列
    tbl.cols = grid_col_widths.iter().map(|w| Col { width: *w }).collect();

    // 解析行
    for tr_xml in tr_bufs {
        if let Ok(row) = parse_table_row(&tr_xml) {
            tbl.rows.push(row);
        }
    }

    Ok(tbl)
}

/// 解析 `<a:tblPr>` 内的 `<a:tblLook>` 属性和 `<a:tableStyleId>` 元素。
///
/// 将解析结果写入 `tbl.tbl_look` 和 `tbl.table_style`。
fn parse_tbl_pr_into(xml: &str, tbl: &mut crate::oxml::table::Table) {
    use crate::oxml::table::TableStyle;

    let look = &mut tbl.tbl_look;
    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                if local_name(e.name().as_ref()) == b"tblLook" {
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"val" => look.val = v.to_string(),
                            b"firstRow" => look.first_row = v == "1",
                            b"lastRow" => look.last_row = v == "1",
                            b"firstColumn" => look.first_column = v == "1",
                            b"lastColumn" => look.last_column = v == "1",
                            b"noHBand" => look.no_h_band = v == "1",
                            b"noVBand" => look.no_v_band = v == "1",
                            _ => {}
                        }
                    }
                }
                // tableStyleId 自闭合形式（罕见但合法：<a:tableStyleId/>）
                if local_name(e.name().as_ref()) == b"tableStyleId" {
                    // 自闭合无文本内容，跳过
                }
            }
            Ok(Event::Start(e)) => {
                if local_name(e.name().as_ref()) == b"tableStyleId" {
                    // 读取文本内容（GUID）
                    let mut guid = String::new();
                    loop {
                        match rd.read_event_into(&mut buf) {
                            Ok(Event::Text(t)) => {
                                guid.push_str(&t.decode().unwrap_or_default());
                            }
                            Ok(Event::End(e2)) => {
                                if local_name(e2.name().as_ref()) == b"tableStyleId" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                        buf.clear();
                    }
                    let trimmed = guid.trim();
                    if !trimmed.is_empty() {
                        tbl.table_style = Some(TableStyle::new(trimmed));
                    }
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:tr>` 元素为 [`Row`]。
fn parse_table_row(xml: &str) -> crate::Result<crate::oxml::table::Row> {
    use crate::oxml::table::Row;

    let mut row = Row::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    // 行高
    // 从根元素 <a:tr h="..."> 读取属性
    let mut tc_bufs: Vec<String> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tr" {
                    // 读取行高属性
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"h" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<i64>()
                            {
                                row.height = Emu(v);
                            }
                        }
                    }
                } else if local == b"tc" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    tc_bufs.push(inner);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(_e)) => {
                // 自闭合 tr（罕见）
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("tr parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    // 解析单元格
    for tc_xml in tc_bufs {
        if let Ok(cell) = parse_table_cell(&tc_xml) {
            row.cells.push(cell);
        }
    }

    Ok(row)
}

/// 解析 `<a:tc>` 元素为 [`Cell`]。
fn parse_table_cell(xml: &str) -> crate::Result<crate::oxml::table::Cell> {
    let mut cell = crate::oxml::table::Cell::default();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tc" {
                    // 读取 gridSpan / rowSpan / hMerge / vMerge 属性
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"gridSpan" => {
                                if let Ok(g) = v.parse::<u32>() {
                                    cell.grid_span = g;
                                }
                            }
                            b"rowSpan" => {
                                if let Ok(r) = v.parse::<u32>() {
                                    cell.row_span = r;
                                }
                            }
                            b"hMerge" if v == "1" => {
                                cell.h_merge = true;
                            }
                            b"vMerge" if v == "1" => {
                                cell.v_merge = true;
                            }
                            _ => {}
                        }
                    }
                } else if local == b"txBody" {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    cell.text = parse_txbody(&inner)?;
                } else if local == b"tcPr" {
                    // 解析 tcPr 属性（marT/marL/marB/marR/anchor）和子元素（lnL/lnR/lnT/lnB/solidFill）
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    parse_tc_pr_into(&inner, &mut cell);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tc" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("tc parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(cell)
}

/// 解析 `<a:tcPr>` 的属性和子元素到 [`Cell`]。
fn parse_tc_pr_into(xml: &str, cell: &mut crate::oxml::table::Cell) {
    use crate::oxml::table::VerticalAnchor;

    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // tcPr 自身的属性（marT/marL/marB/marR/anchor）
                // 注意：根元素是 tcPr，其属性在 Start 事件上
                if local == b"tcPr" {
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"marT" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.0 = Some(Emu(m));
                                }
                            }
                            b"marL" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.1 = Some(Emu(m));
                                }
                            }
                            b"marB" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.2 = Some(Emu(m));
                                }
                            }
                            b"marR" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.3 = Some(Emu(m));
                                }
                            }
                            b"anchor" => {
                                // 使用 &*v 解引用 Cow<str> 为 &str，避免不稳定的 str::as_str()
                                cell.anchor = match &*v {
                                    "t" => VerticalAnchor::Top,
                                    "ctr" => VerticalAnchor::Middle,
                                    "b" => VerticalAnchor::Bottom,
                                    _ => VerticalAnchor::Middle,
                                };
                            }
                            _ => {}
                        }
                    }
                } else if local == b"solidFill" {
                    // 单元格填充色
                    if let Ok(color) = parse_solid_fill(&mut rd, e.into_owned()) {
                        cell.fill = color;
                    }
                } else if local == b"lnL" {
                    // 单元格左边框：分支独立，避免 e.into_owned() 后 local 引用失效
                    let border = parse_cell_border(&mut rd, e.into_owned());
                    cell.border_left = Some(border);
                } else if local == b"lnR" {
                    let border = parse_cell_border(&mut rd, e.into_owned());
                    cell.border_right = Some(border);
                } else if local == b"lnT" {
                    let border = parse_cell_border(&mut rd, e.into_owned());
                    cell.border_top = Some(border);
                } else if local == b"lnB" {
                    let border = parse_cell_border(&mut rd, e.into_owned());
                    cell.border_bottom = Some(border);
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"tcPr" {
                    // 自闭合 tcPr：仅属性
                    for a in e.attributes().flatten() {
                        let v = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match a.key.as_ref() {
                            b"marT" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.0 = Some(Emu(m));
                                }
                            }
                            b"marL" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.1 = Some(Emu(m));
                                }
                            }
                            b"marB" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.2 = Some(Emu(m));
                                }
                            }
                            b"marR" => {
                                if let Ok(m) = v.parse::<i64>() {
                                    cell.margin.3 = Some(Emu(m));
                                }
                            }
                            b"anchor" => {
                                // 使用 &*v 解引用 Cow<str> 为 &str，避免不稳定的 str::as_str()
                                cell.anchor = match &*v {
                                    "t" => VerticalAnchor::Top,
                                    "ctr" => VerticalAnchor::Middle,
                                    "b" => VerticalAnchor::Bottom,
                                    _ => VerticalAnchor::Middle,
                                };
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"tcPr" {
                    return;
                }
            }
            Ok(Event::Eof) => return,
            _ => {}
        }
        buf.clear();
    }
}

/// 解析单元格边框 `<a:lnL>` / `<a:lnR>` / `<a:lnT>` / `<a:lnB>` 为 [`CellBorder`]。
fn parse_cell_border<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    start: quick_xml::events::BytesStart<'static>,
) -> crate::oxml::table::CellBorder {
    use crate::oxml::table::CellBorder;
    let mut border = CellBorder::default();
    // 读取 w 属性
    for a in start.attributes().flatten() {
        if a.key.as_ref() == b"w" {
            if let Ok(v) = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default()
                .parse::<i64>()
            {
                border.width = Emu(v);
            }
        }
    }
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"noFill" {
                    border.no_fill = true;
                }
                // solidFill 自闭合（<a:solidFill/>）无颜色子元素，跳过
            }
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"solidFill" {
                    if let Ok(color) = parse_solid_fill(rd, e.into_owned()) {
                        border.color = color;
                    }
                } else {
                    let _ = collect_full_element(rd, e.into_owned());
                }
            }
            Ok(Event::End(e)) => {
                // 返回到调用者：边框元素结束
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"lnL" || local == b"lnR" || local == b"lnT" || local == b"lnB" {
                    return border;
                }
            }
            Ok(Event::Eof) => return border,
            _ => {}
        }
        buf.clear();
    }
}

/// 从 `<p:notes>` / `<p:notesSlide>` 元素解析出 `TextBody`。
///
/// 元素结构（OOXML 严格顺序）：
///
/// ```text
/// <p:notes>
///   <p:cSld>
///     <p:spTree>
///       <p:nvGrpSpPr/>
///       <p:grpSpPr/>
///       <p:sp>
///         <p:nvSpPr>...</p:nvSpPr>
///         <p:spPr>...</p:spPr>
///         <p:txBody>...</p:txBody>   ← 备注文本
///       </p:sp>
///     </p:spTree>
///   </p:cSld>
/// </p:notes>
/// ```
///
/// # 行为
/// - 仅取第一个 `p:sp` 内的 `p:txBody`，**忽略** `p:sp` 之外的其它内容；
/// - 若没有 `p:sp` 或没有 `p:txBody`，返回 `Ok(TextBody::new())`（空备注）。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_notes(xml: &str) -> crate::Result<TextBody> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut tx_buf: Option<String> = None;
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"txBody" && tx_buf.is_none() {
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    tx_buf = Some(inner);
                }
                // 其它 Start 事件（notes / cSld / spTree / sp / nvSpPr / spPr 等）
                // 不调用 collect_full_element——否则会把 txBody 后代一起吞掉。
                // 让子元素在后续迭代中自然展开，直到遇到 txBody 为止。
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("notes parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    if let Some(buf) = tx_buf {
        parse_txbody(&buf)
    } else {
        Ok(TextBody::new())
    }
}

/// 从 `commentN.xml` 解析评论列表（`<p:cmLst>`）。
///
/// # 行为
/// - 遍历所有 `<p:cm>` 元素，提取 `authorId` / `dt` / `idx` 属性；
/// - 对每个 `<p:cm>` 内部的 `<p:pos x="..." y="..."/>` 和 `<p:text>...</p:text>` 分别解析；
/// - 忽略无法识别的子元素。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_comments(xml: &str) -> crate::Result<crate::oxml::comments::CommentList> {
    use crate::oxml::comments::{Comment, CommentList};

    let mut lst = CommentList::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    // 当前正在解析的 <p:cm> 的属性
    let mut cur_author_id: Option<u32> = None;
    let mut cur_dt: String = String::new();
    let mut cur_idx: Option<u32> = None;
    // 当前是否在 <p:cm> 内部
    let mut in_cm = false;
    // 当前是否在 <p:text> 内部（用于收集文本）
    let mut in_text = false;
    // 当前 <p:cm> 的 pos
    let mut cur_pos_x: i64 = 0;
    let mut cur_pos_y: i64 = 0;
    let mut cur_text = String::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"cm" {
                    // 进入新的 <p:cm>：重置状态
                    in_cm = true;
                    cur_author_id = None;
                    cur_dt.clear();
                    cur_idx = None;
                    cur_pos_x = 0;
                    cur_pos_y = 0;
                    cur_text.clear();
                    // 提取属性
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let val = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .to_string();
                        match key {
                            b"authorId" => {
                                cur_author_id = val.parse::<u32>().ok();
                            }
                            b"dt" => {
                                cur_dt = val;
                            }
                            b"idx" => {
                                cur_idx = val.parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                } else if in_cm && local == b"text" {
                    in_text = true;
                    cur_text.clear();
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_cm && local == b"pos" {
                    // 自闭合 <p:pos x="..." y="..."/>
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let val = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default();
                        match key {
                            b"x" => {
                                cur_pos_x = val.parse::<i64>().unwrap_or(0);
                            }
                            b"y" => {
                                cur_pos_y = val.parse::<i64>().unwrap_or(0);
                            }
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if in_text {
                    cur_text.push_str(&t.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"cm" && in_cm {
                    // 结束当前 <p:cm>：组装并推入列表
                    let c = Comment {
                        author_id: cur_author_id.unwrap_or(0),
                        date_time: std::mem::take(&mut cur_dt),
                        idx: cur_idx.unwrap_or(0),
                        pos_x: cur_pos_x,
                        pos_y: cur_pos_y,
                        text: std::mem::take(&mut cur_text),
                    };
                    lst.push(c);
                    in_cm = false;
                    in_text = false;
                } else if local == b"text" && in_text {
                    in_text = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("comments parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(lst)
}

/// 从 `commentAuthors.xml` 解析评论作者列表（`<p:cmAuthorLst>`）。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_comment_authors(xml: &str) -> crate::Result<crate::oxml::comments::CommentAuthorList> {
    use crate::oxml::comments::{CommentAuthor, CommentAuthorList};

    let mut lst = CommentAuthorList::new();
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"cmAuthor" {
                    let mut id: u32 = 0;
                    let mut name_val = String::new();
                    let mut initials = String::new();
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let val = a
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .unwrap_or_default()
                            .to_string();
                        match key {
                            b"id" => {
                                id = val.parse::<u32>().unwrap_or(0);
                            }
                            b"name" => {
                                name_val = val;
                            }
                            b"initials" => {
                                initials = val;
                            }
                            _ => {}
                        }
                    }
                    lst.push(CommentAuthor::new(id, name_val, initials));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("commentAuthors parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(lst)
}

/// 从 `presentation.xml` 解析出关键字段（`sldIdLst` + `sldSz` + `sectionLst`）。
///
/// # 返回值
/// 元组 `(slide_ids, slide_width, slide_height, sld_master_ids, sections)`：
/// - `slide_ids`：按文档顺序排列的 `(id, rid)` 列表；
/// - `slide_width` / `slide_height`：EMU 尺寸；若未指定则回落到 `None`；
/// - `sld_master_ids`：母版 `(id, rid)` 列表（TODO-001）；
/// - `sections`：章节分组列表（TODO-039），若未指定则为空。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误或关键属性缺失。
#[allow(clippy::type_complexity)]
pub fn parse_pres_root(
    xml: &str,
) -> crate::Result<(
    Vec<(u32, String)>,
    Option<Emu>,
    Option<Emu>,
    Vec<(u32, String)>,
    crate::oxml::section::SectionList,
)> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sld_ids: Vec<(u32, String)> = Vec::new();
    let mut sld_sz: (Option<Emu>, Option<Emu>) = (None, None);
    // TODO-001：解析 sldMasterIdLst
    let mut sld_master_ids: Vec<(u32, String)> = Vec::new();
    // TODO-039：解析 sectionLst（位于 extLst 内的 p14 扩展）
    let mut sections = crate::oxml::section::SectionList::default();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"sldId" {
                    let mut id: Option<u32> = None;
                    let mut rid: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        // 无前缀 `id` 与带前缀 `r:id` 都接受。
                        if k == b"id" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<u32>()
                            {
                                id = Some(v);
                            }
                        } else if k == b"r:id" || k.ends_with(b":id") {
                            rid = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                    if let (Some(i), Some(r)) = (id, rid) {
                        sld_ids.push((i, r));
                    }
                } else if local == b"sldMasterId" {
                    // TODO-001：解析 sldMasterId 元素
                    let mut id: Option<u32> = None;
                    let mut rid: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        if k == b"id" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<u32>()
                            {
                                id = Some(v);
                            }
                        } else if k == b"r:id" || k.ends_with(b":id") {
                            rid = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                    if let (Some(i), Some(r)) = (id, rid) {
                        sld_master_ids.push((i, r));
                    }
                } else if local == b"sldSz" {
                    let mut cx: Option<i64> = None;
                    let mut cy: Option<i64> = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"cx" => {
                                cx = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            b"cy" => {
                                cy = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .parse()
                                    .ok()
                            }
                            _ => {}
                        }
                    }
                    sld_sz = (cx.map(Emu), cy.map(Emu));
                }
            }
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"sldMasterId" {
                    // Start 形式的 sldMasterId（含子元素，罕见但规范允许）
                    let mut id: Option<u32> = None;
                    let mut rid: Option<String> = None;
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        if k == b"id" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<u32>()
                            {
                                id = Some(v);
                            }
                        } else if k == b"r:id" || k.ends_with(b":id") {
                            rid = Some(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                    if let (Some(i), Some(r)) = (id, rid) {
                        sld_master_ids.push((i, r));
                    }
                    // 跳过子元素直到 End
                    let _ = collect_full_element(&mut rd, e.into_owned());
                } else if local == b"extLst" {
                    // TODO-039：解析 extLst 内的 sectionLst 扩展（p14 命名空间）。
                    // 收集完整 extLst XML 后单独解析，避免主状态机嵌套过深。
                    let ext_lst_xml = collect_full_element(&mut rd, e.into_owned())?;
                    let parsed = parse_sections_from_ext_lst(&ext_lst_xml)?;
                    if !parsed.is_empty() {
                        sections = parsed;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("pres root parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok((sld_ids, sld_sz.0, sld_sz.1, sld_master_ids, sections))
}

/// 从 `<p:extLst>` 的 XML 片段中解析出 [`crate::oxml::section::SectionList`]。
///
/// 仅识别 `uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E64}"` 的 `<p:ext>`，
/// 内部 `<p14:sectionLst>` 中的每个 `<p14:section>` 会被还原为
/// [`crate::oxml::section::Section`]（name + slide_ids）。
///
/// # 返回值
/// - 解析到的章节列表（可能为空，表示 extLst 中无 section 扩展）。
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
///
/// # 示例输入
/// ```xml
/// <p:extLst>
///   <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E64}">
///     <p14:sectionLst>
///       <p14:section name="章节一">
///         <p14:sldIdLst>
///           <p14:sldId id="256"/>
///         </p14:sldIdLst>
///       </p14:section>
///     </p14:sectionLst>
///   </p:ext>
/// </p:extLst>
/// ```
fn parse_sections_from_ext_lst(xml: &str) -> crate::Result<crate::oxml::section::SectionList> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut result = crate::oxml::section::SectionList::default();
    // 状态机：
    //   0 = 在 extLst 内寻找 ext
    //   1 = 在匹配的 ext 内寻找 sectionLst
    //   2 = 在 sectionLst 内解析 section
    //   3 = 在 section 内解析 sldIdLst
    let mut state: u8 = 0;
    // 当前正在构建的 section（在状态 2/3 中累积）
    let mut cur_name = String::new();
    let mut cur_ids: Vec<u32> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"ext" => {
                        // 检查 uri 属性是否匹配 section 扩展
                        let mut uri_val = String::new();
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"uri" {
                                uri_val = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                        if uri_val == crate::oxml::section::SECTION_EXT_URI {
                            state = 1;
                        } else {
                            // 非 section 扩展——跳过整个 ext
                            let _ = collect_full_element(&mut rd, e.into_owned());
                        }
                    }
                    1 if local == b"sectionLst" => {
                        state = 2;
                    }
                    2 if local == b"section" => {
                        // 读取 name 属性
                        cur_name.clear();
                        cur_ids.clear();
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"name" {
                                cur_name = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                        state = 3;
                    }
                    3 if local == b"sldIdLst" => {
                        // 进入 sldIdLst，继续读取其中的 sldId（Empty 事件）
                    }
                    _ => {
                        // 其它未识别的子元素——整体跳过
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                // 在 sldIdLst 内遇到自闭合的 sldId
                if state == 3 && local == b"sldId" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"id" {
                            if let Ok(v) = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .parse::<u32>()
                            {
                                cur_ids.push(v);
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                // 把 e.name() 绑定到 let，避免临时 QName 在 as_ref() 后 drop
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    3 if local == b"section" => {
                        // section 结束——提交
                        result.push(crate::oxml::section::Section {
                            name: std::mem::take(&mut cur_name),
                            slide_ids: std::mem::take(&mut cur_ids),
                        });
                        state = 2;
                    }
                    2 if local == b"sectionLst" => {
                        state = 1;
                    }
                    1 if local == b"ext" => {
                        state = 0;
                    }
                    0 if local == b"extLst" => {
                        break;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("sectionLst parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(result)
}

/// 从 `slideMasterN.xml` 文本解析出 [`SldMaster`]。
///
/// # 解析内容
/// - `<p:spTree>` 内的所有 `<p:sp>`（其它子类型如 pic/grpSp 暂不解析，保留位置）
/// - `<p:sldLayoutIdLst>` 内的所有 `r:id`（layout 关系 id 列表）
///
/// # 忽略内容（后续扩展）
/// - `<p:clrMap>` 颜色映射
/// - `<p:txStyles>` 文本样式
/// - `<p:extLst>` 扩展列表
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_sld_master(xml: &str) -> crate::Result<crate::oxml::slidemaster::SldMaster> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut master = crate::oxml::slidemaster::SldMaster::default();
    // 状态机：0=等待 sldMaster，1=在 sldMaster 内，2=在 spTree 内，3=在 sldLayoutIdLst 内
    let mut state: u8 = 0;
    let mut sp_bufs: Vec<String> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"sldMaster" => {
                        state = 1;
                    }
                    // cSld 是容器元素，不能跳过（内含 spTree）
                    1 if local == b"cSld" => {}
                    1 if local == b"spTree" => {
                        state = 2;
                    }
                    1 if local == b"sldLayoutIdLst" => {
                        state = 3;
                    }
                    2 if local == b"sp" => {
                        // 累积 sp 的内部 XML，走子解析
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp_bufs.push(inner);
                    }
                    _ => {
                        // 其它子元素（clrMap / txStyles 等）—— 整个吞掉
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 3 && local == b"sldLayoutId" {
                    // 提取 r:id 属性
                    for a in e.attributes().flatten() {
                        let k = a.key.as_ref();
                        if k == b"r:id" || k.ends_with(b":id") {
                            let rid = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                            master.layout_rids.push(rid);
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // state==2（spTree 结束）或 state==3（sldLayoutIdLst 结束）都回到 state=1
                if (state == 2 && local == b"spTree") || (state == 3 && local == b"sldLayoutIdLst")
                {
                    state = 1;
                } else if state == 1 && local == b"sldMaster" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("sldMaster parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 解析所有累积的 sp
    for s in sp_bufs {
        if let Ok(sp) = parse_sp(&s) {
            master.shapes.push(sp);
        }
    }
    Ok(master)
}

/// 从 `notesMasterN.xml` 文本解析出 [`crate::oxml::notesmaster::NotesMaster`]（TODO-045）。
///
/// 与 [`parse_sld_master`] 结构类似，但根元素是 `<p:notesMaster>`，
/// 且不含 `<p:sldLayoutIdLst>`（备注母版不挂版式）。
///
/// # 解析内容
/// - `<p:spTree>` 内的所有 `<p:sp>`（其它子类型如 pic/grpSp 暂不解析，保留位置）
///
/// # 忽略内容（后续扩展）
/// - `<p:clrMap>` 颜色映射
/// - `<p:notesStyle>` 备注文本样式
/// - `<p:extLst>` 扩展列表
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_notes_master(xml: &str) -> crate::Result<crate::oxml::notesmaster::NotesMaster> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut master = crate::oxml::notesmaster::NotesMaster::default();
    // 状态机：0=等待 notesMaster，1=在 notesMaster 内，2=在 spTree 内
    let mut state: u8 = 0;
    let mut sp_bufs: Vec<String> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"notesMaster" => {
                        state = 1;
                    }
                    // cSld 是容器元素，不能跳过（内含 spTree）
                    1 if local == b"cSld" => {}
                    1 if local == b"spTree" => {
                        state = 2;
                    }
                    2 if local == b"sp" => {
                        // 累积 sp 的内部 XML，走子解析
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp_bufs.push(inner);
                    }
                    _ => {
                        // 其它子元素（clrMap / notesStyle 等）—— 整个吞掉
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 2 && local == b"spTree" {
                    state = 1;
                } else if state == 1 && local == b"notesMaster" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("notesMaster parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 解析所有累积的 sp
    for s in sp_bufs {
        if let Ok(sp) = parse_sp(&s) {
            master.shapes.push(sp);
        }
    }
    Ok(master)
}

/// 从 `slideLayoutN.xml` 文本解析出 [`SldLayout`]。
///
/// # 解析内容
/// - `<p:sldLayout type="...">` 根元素的 `type` 属性
/// - `<p:cSld name="...">` 的 `name` 属性
/// - `<p:spTree>` 内的所有 `<p:sp>`
///
/// # 忽略内容（后续扩展）
/// - `<p:clrMapOvr>` 颜色映射覆盖
/// - `<p:transition>` 过渡
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_sld_layout(xml: &str) -> crate::Result<crate::oxml::slidelayout::SldLayout> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut layout = crate::oxml::slidelayout::SldLayout::default();
    // 状态机：0=等待 sldLayout，1=在 sldLayout 内，2=在 spTree 内
    let mut state: u8 = 0;
    let mut sp_bufs: Vec<String> = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match state {
                    0 if local == b"sldLayout" => {
                        // 提取 type 属性
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"type" {
                                layout.type_ = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                        state = 1;
                    }
                    1 if local == b"cSld" => {
                        // 提取 name 属性
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"name" {
                                layout.name = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                    }
                    1 if local == b"spTree" => {
                        state = 2;
                    }
                    2 if local == b"sp" => {
                        let inner = collect_full_element(&mut rd, e.into_owned())?;
                        sp_bufs.push(inner);
                    }
                    _ => {
                        let _ = collect_full_element(&mut rd, e.into_owned());
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if state == 2 && local == b"spTree" {
                    state = 1;
                } else if state == 1 && local == b"sldLayout" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("sldLayout parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    // 解析所有累积的 sp
    for s in sp_bufs {
        if let Ok(sp) = parse_sp(&s) {
            layout.shapes.push(sp);
        }
    }
    Ok(layout)
}

/// 从 `themeN.xml` 文本解析出 [`Theme`]。
///
/// # 解析内容
/// - `<a:theme name="...">` 根元素的 `name` 属性
/// - `<a:clrScheme>` 内的 12 个颜色（dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink）
/// - `<a:fontScheme>` 内的 majorFont/minorFont 的 latin typeface
///
/// # 忽略内容（后续扩展）
/// - `<a:fmtScheme>` 格式方案
/// - `<a:objectDefaults>` 对象默认值
/// - `<a:extraClrSchemeLst>` 额外颜色方案
///
/// # 错误
/// - [`crate::Error::Xml`]：XML 语法错误。
pub fn parse_theme(xml: &str) -> crate::Result<crate::oxml::theme::Theme> {
    let mut rd = Reader::from_str(xml);
    rd.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut theme = crate::oxml::theme::Theme::default();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"theme" {
                    // 提取 name 属性
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"name" {
                            theme.name = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if local == b"clrScheme" {
                    // 解析颜色方案
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    theme.color_scheme = parse_clr_scheme(&inner);
                } else if local == b"fontScheme" {
                    // 解析字体方案
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    theme.font_scheme = parse_font_scheme(&inner);
                } else if local == b"fmtScheme" {
                    // 解析格式方案（TODO-005：FormatScheme 结构化解析）
                    // 保留原始 XML 支持 round-trip，同时填充 4 个结构化字段
                    let inner = collect_full_element(&mut rd, e.into_owned())?;
                    let (fmt_name, raw) = parse_fmt_scheme(&inner);
                    theme.format_scheme.name = fmt_name;
                    theme.format_scheme.raw_xml = raw;
                    // 结构化解析：把 raw_xml 拆分为 fill_styles / line_styles / effect_styles / bg_fill_styles
                    theme.format_scheme.parse_from_raw_xml();
                } else if local == b"themeElements" {
                    // themeElements 是容器元素，不跳过（内含 clrScheme / fontScheme / fmtScheme）
                } else {
                    let _ = collect_full_element(&mut rd, e.into_owned());
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::Error::Xml(format!("theme parse: {e}"))),
            _ => {}
        }
        buf.clear();
    }
    Ok(theme)
}

/// 解析 `<a:fmtScheme>` 的 name 属性和内部 raw XML。
///
/// 返回 `(name, raw_xml)`：
/// - `name`：`<a:fmtScheme name="...">` 的 name 属性值；
/// - `raw_xml`：`<a:fmtScheme>` 内部所有子元素的原始 XML 字符串（不含 fmtScheme 标签本身）。
fn parse_fmt_scheme(xml: &str) -> (String, String) {
    let mut fmt_name = String::new();
    let mut raw_xml = String::new();
    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut in_fmt_scheme = false;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                if !in_fmt_scheme && local == b"fmtScheme" {
                    in_fmt_scheme = true;
                    // 提取 name 属性
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"name" {
                            fmt_name = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if in_fmt_scheme {
                    // 收集子元素到 raw_xml
                    let inner = collect_full_element(&mut rd, e.into_owned()).unwrap_or_default();
                    raw_xml.push_str(&inner);
                }
            }
            Ok(Event::Empty(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                if in_fmt_scheme {
                    // 自闭合子元素，写入 raw_xml
                    let mut s = String::new();
                    s.push('<');
                    s.push_str(std::str::from_utf8(local).unwrap_or(""));
                    for a in e.attributes().flatten() {
                        s.push(' ');
                        s.push_str(std::str::from_utf8(a.key.as_ref()).unwrap_or(""));
                        s.push_str("=\"");
                        s.push_str(
                            &a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default(),
                        );
                        s.push('"');
                    }
                    s.push_str("/>");
                    raw_xml.push_str(&s);
                }
            }
            Ok(Event::End(e)) => {
                let ev_name = e.name();
                let local = local_name(ev_name.as_ref());
                if in_fmt_scheme && local == b"fmtScheme" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    (fmt_name, raw_xml)
}

/// 解析 `<a:clrScheme>` 内的 12 个颜色。
fn parse_clr_scheme(xml: &str) -> crate::oxml::theme::ColorScheme {
    use crate::oxml::theme::{ColorScheme, ThemeColor};

    let mut scheme = ColorScheme::default();
    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // 根据颜色槽位名分发，把颜色值写入对应字段
                let slot: &mut Option<ThemeColor> = match local {
                    b"dk1" => &mut scheme.dk1,
                    b"lt1" => &mut scheme.lt1,
                    b"dk2" => &mut scheme.dk2,
                    b"lt2" => &mut scheme.lt2,
                    b"accent1" => &mut scheme.accent1,
                    b"accent2" => &mut scheme.accent2,
                    b"accent3" => &mut scheme.accent3,
                    b"accent4" => &mut scheme.accent4,
                    b"accent5" => &mut scheme.accent5,
                    b"accent6" => &mut scheme.accent6,
                    b"hlink" => &mut scheme.hlink,
                    b"folHlink" => &mut scheme.fol_hlink,
                    // clrScheme 根元素自身 —— 提取 name 属性，不跳过子树
                    b"clrScheme" => {
                        // 提取 name 属性（TODO-005）
                        for a in e.attributes().flatten() {
                            if a.key.as_ref() == b"name" {
                                scheme.name = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                        }
                        continue;
                    }
                    _ => {
                        // 非颜色槽位未知元素 —— 跳过子树
                        let _ = collect_full_element(&mut rd, e.into_owned());
                        continue;
                    }
                };
                *slot = Some(read_color_child(&mut rd, local));
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"clrScheme" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    scheme
}

/// 读取颜色槽位（如 `<a:dk1>`）内的子元素 `<a:srgbClr>` 或 `<a:sysClr>`。
///
/// `parent_local` 是颜色槽位的 local name（如 `b"dk1"`），用于判断何时返回。
fn read_color_child<R: std::io::BufRead>(
    rd: &mut Reader<R>,
    parent_local: &[u8],
) -> crate::oxml::theme::ThemeColor {
    use crate::oxml::theme::ThemeColor;
    let mut buf = Vec::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"srgbClr" {
                    let mut val = String::new();
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            val = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                    // 消费 srgbClr 的 End 事件
                    let _ = collect_full_element(rd, e.into_owned());
                    return ThemeColor::Srgb(val);
                } else if local == b"sysClr" {
                    let mut val = String::new();
                    let mut last_clr = String::new();
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"val" => {
                                val = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                            b"lastClr" => {
                                last_clr = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                            _ => {}
                        }
                    }
                    let _ = collect_full_element(rd, e.into_owned());
                    return ThemeColor::Sys(val, last_clr);
                } else {
                    let _ = collect_full_element(rd, e.into_owned());
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"srgbClr" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"val" {
                            return ThemeColor::Srgb(
                                a.normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string(),
                            );
                        }
                    }
                } else if local == b"sysClr" {
                    let mut val = String::new();
                    let mut last_clr = String::new();
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"val" => {
                                val = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                            b"lastClr" => {
                                last_clr = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .unwrap_or_default()
                                    .to_string();
                            }
                            _ => {}
                        }
                    }
                    return ThemeColor::Sys(val, last_clr);
                }
            }
            Ok(Event::End(e)) => {
                // 回到颜色槽位的 End —— 没有找到颜色子元素
                if local_name(e.name().as_ref()) == parent_local {
                    return ThemeColor::None;
                }
            }
            Ok(Event::Eof) => return ThemeColor::None,
            _ => {}
        }
        buf.clear();
    }
}

/// 解析 `<a:fontScheme>` 内的 majorFont/minorFont。
fn parse_font_scheme(xml: &str) -> crate::oxml::theme::FontScheme {
    let mut scheme = crate::oxml::theme::FontScheme::default();
    let mut rd = Reader::from_str(xml);
    let mut buf = Vec::new();
    // 0=等待 fontScheme，1=在 fontScheme 内
    let mut state: u8 = 0;

    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"fontScheme" {
                    state = 1;
                    // 提取 name 属性（TODO-005）
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"name" {
                            scheme.name = a
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default()
                                .to_string();
                        }
                    }
                } else if state == 1 && local == b"majorFont" {
                    // 读取 latin / ea / cs typeface（TODO-005：扩展 FontScheme）
                    let (latin, ea, cs) = read_font_typefaces(&mut rd);
                    scheme.major_latin = latin;
                    scheme.major_ea = ea;
                    scheme.major_cs = cs;
                } else if state == 1 && local == b"minorFont" {
                    let (latin, ea, cs) = read_font_typefaces(&mut rd);
                    scheme.minor_latin = latin;
                    scheme.minor_ea = ea;
                    scheme.minor_cs = cs;
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == b"fontScheme" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }
    scheme
}

/// 从 majorFont/minorFont 内读取 `<a:latin>` / `<a:ea>` / `<a:cs>` 的 typeface。
///
/// 返回 `(latin, ea, cs)` 三个 typeface 字符串。
/// 在 majorFont/minorFont 的 End 事件时返回（空值表示未找到）。
fn read_font_typefaces<R: std::io::BufRead>(rd: &mut Reader<R>) -> (String, String, String) {
    let mut buf = Vec::new();
    let mut latin = String::new();
    let mut ea = String::new();
    let mut cs = String::new();
    loop {
        match rd.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                // 提取 typeface 属性
                let typeface = extract_typeface(&e);
                match local {
                    b"latin" => latin = typeface,
                    b"ea" => ea = typeface,
                    b"cs" => cs = typeface,
                    _ => {}
                }
            }
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                let typeface = extract_typeface(&e);
                match local {
                    b"latin" => latin = typeface,
                    b"ea" => ea = typeface,
                    b"cs" => cs = typeface,
                    _ => {}
                }
                // 消费 Start 元素的子树（含 End）
                let _ = collect_full_element(rd, e.into_owned());
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"majorFont" || local == b"minorFont" {
                    return (latin, ea, cs);
                }
            }
            Ok(Event::Eof) => return (latin, ea, cs),
            _ => {}
        }
        buf.clear();
    }
}

/// 从元素的属性中提取 `typeface` 值。
fn extract_typeface(e: &quick_xml::events::BytesStart<'_>) -> String {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"typeface" {
            return a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .unwrap_or_default()
                .to_string();
        }
    }
    String::new()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::oxml::txbody::TextBody;

    // ===================== TODO-050 backdrop 解析测试 =====================

    /// 验证 `parse_backdrop` 解析完整 backdrop（含 anchor + 全平面）。
    #[test]
    fn parse_backdrop_full() {
        let xml = r#"<?xml version="1.0"?>
<a:backdrop xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:anchor x="100000" y="200000" z="300000"/>
  <a:floor/>
  <a:wall/>
  <a:l/>
  <a:r/>
  <a:t/>
  <a:b/>
</a:backdrop>"#;
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"backdrop" => {
                    break e.into_owned();
                }
                Ok(Event::Eof) => panic!("backdrop not found"),
                _ => {}
            }
            buf.clear();
        };
        let bd = parse_backdrop(&mut rd, start).expect("parse backdrop");

        assert_eq!(
            bd.anchor,
            Some(Point3d {
                x: 100000,
                y: 200000,
                z: 300000
            })
        );
        assert!(bd.floor);
        assert!(bd.wall);
        assert!(bd.left);
        assert!(bd.right);
        assert!(bd.top);
        assert!(bd.bottom);
    }

    /// 验证 `parse_backdrop` 仅启用部分平面。
    #[test]
    fn parse_backdrop_partial() {
        let xml = r#"<?xml version="1.0"?>
<a:backdrop xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:floor/>
  <a:wall/>
</a:backdrop>"#;
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"backdrop" => {
                    break e.into_owned();
                }
                Ok(Event::Eof) => panic!("backdrop not found"),
                _ => {}
            }
            buf.clear();
        };
        let bd = parse_backdrop(&mut rd, start).expect("parse backdrop");

        assert!(bd.anchor.is_none());
        assert!(bd.floor);
        assert!(bd.wall);
        assert!(!bd.left);
        assert!(!bd.right);
        assert!(!bd.top);
        assert!(!bd.bottom);
    }

    /// 验证 `parse_scene_3d` 正确解析 backdrop 子元素（round-trip 验证）。
    #[test]
    fn parse_scene_3d_with_backdrop() {
        let xml = r#"<?xml version="1.0"?>
<a:scene3d xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:camera prst="orthographicFront"/>
  <a:lightRig rig="balanced" dir="t"/>
  <a:backdrop>
    <a:anchor x="0" y="0" z="0"/>
    <a:floor/>
    <a:wall/>
  </a:backdrop>
</a:scene3d>"#;
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"scene3d" => {
                    break e.into_owned();
                }
                Ok(Event::Eof) => panic!("scene3d not found"),
                _ => {}
            }
            buf.clear();
        };
        let scene = parse_scene_3d(&mut rd, start).expect("parse scene3d");

        // 验证 backdrop 被解析
        let bd = scene.backdrop.expect("backdrop should be parsed");
        assert!(bd.floor);
        assert!(bd.wall);
        assert!(!bd.left);
        assert_eq!(bd.anchor, Some(Point3d { x: 0, y: 0, z: 0 }));
    }

    /// 验证 `parse_scene_3d` 无 backdrop 时 backdrop 字段为 None（向后兼容）。
    #[test]
    fn parse_scene_3d_without_backdrop() {
        let xml = r#"<?xml version="1.0"?>
<a:scene3d xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:camera prst="orthographicFront"/>
  <a:lightRig rig="balanced" dir="t"/>
</a:scene3d>"#;
        let mut rd = Reader::from_str(xml);
        rd.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let start = loop {
            match rd.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"scene3d" => {
                    break e.into_owned();
                }
                Ok(Event::Eof) => panic!("scene3d not found"),
                _ => {}
            }
            buf.clear();
        };
        let scene = parse_scene_3d(&mut rd, start).expect("parse scene3d");
        assert!(scene.backdrop.is_none());
    }

    #[test]
    fn parse_minimal_sld() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 1"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="914400"/>
            <a:ext cx="3657600" cy="457200"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr>
          <a:p>
            <a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        assert_eq!(sld.shapes.len(), 1);
        if let OxmlSlideShape::Sp(sp) = &sld.shapes[0] {
            assert_eq!(sp.id, 2);
            assert_eq!(sp.name, "TextBox 1");
            assert!(sp.c_nv_sp_pr_tx_box);
            assert_eq!(sp.text.paragraphs.len(), 1);
            assert_eq!(sp.text.paragraphs[0].runs.len(), 1);
            assert_eq!(sp.text.paragraphs[0].runs[0].text, "Hello");
        } else {
            panic!("expected Sp");
        }
    }

    #[test]
    fn parse_empty_txbody_keeps_one_paragraph() {
        let xml = r#"<p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:bodyPr/>
        </p:txBody>"#;
        let tb = parse_txbody(xml).unwrap();
        assert!(tb.paragraphs.is_empty());
        // TextBody 解析出来确实没有段落——这与 PowerPoint 期望"至少一个段落"略不同，
        // 但符合"无段落就不生成"的最小语义。
        let tb2 = TextBody::new();
        assert!(tb2.paragraphs.is_empty());
    }

    #[test]
    fn parse_pres_root_extracts_sld_id_lst_and_size() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#;
        let (ids, w, h, master_ids, _sections) = parse_pres_root(xml).expect("parse ok");
        assert_eq!(ids.len(), 2);
        assert_eq!(ids[0], (256, "rId2".to_string()));
        assert_eq!(ids[1], (257, "rId3".to_string()));
        assert_eq!(w.map(|v| v.0), Some(9_144_000));
        assert_eq!(h.map(|v| v.0), Some(6_858_000));
        // TODO-001：验证 sldMasterIdLst 解析
        assert_eq!(master_ids.len(), 1);
        assert_eq!(master_ids[0], (2147483648, "rId1".to_string()));
    }

    #[test]
    fn parse_notes_returns_text_body() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Notes Placeholder"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>hello notes</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>"#;
        let tb = parse_notes(xml).expect("parse notes ok");
        assert_eq!(tb.paragraphs.len(), 1);
        assert_eq!(tb.paragraphs[0].runs.len(), 1);
        assert_eq!(tb.paragraphs[0].runs[0].text, "hello notes");
    }

    /// `parse_pic` 必须能正确解析 r:embed 与 srcRect。
    ///
    /// 早期版本用 `find("r:embed=")` 字符串扫描；本测试覆盖**自闭合 `<a:blip/>`** 与
    /// **非自闭合 `<a:blip>...</a:blip>`** 两种形态 + 缩进变化，确保 SAX 解析稳定。
    #[test]
    fn parse_pic_extracts_rid_and_src_rect() {
        // 自闭合 blip + 有 srcRect
        let xml1 = r#"<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr><p:cNvPr id="42" name="MyPic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rIdImg5"/>
    <a:srcRect l="1000" t="2000" r="3000" b="4000"/>
  </p:blipFill>
  <p:spPr/>
</p:pic>"#;
        let p1 = parse_pic(xml1).expect("parse pic 1");
        assert_eq!(p1.id, 42);
        assert_eq!(p1.name, "MyPic");
        assert_eq!(p1.rid, "rIdImg5");
        assert_eq!(p1.src_rect, Some((1000, 2000, 3000, 4000)));

        // 非自闭合 blip
        let xml2 = r#"<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:nvPicPr><p:cNvPr id="7" name="p2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
<p:blipFill><a:blip r:embed="rIdImg99"></a:blip></p:blipFill>
<p:spPr/>
</p:pic>"#;
        let p2 = parse_pic(xml2).expect("parse pic 2");
        assert_eq!(p2.id, 7);
        assert_eq!(p2.rid, "rIdImg99");
    }

    /// `collect_attr_value` 之前是 noop + 比较 noop，导致 typeface 永远取不到。
    /// 本测试确保 `<a:latin typeface="..."/>` 能正确解析。
    #[test]
    fn parse_run_with_typeface_works() {
        let xml = r#"<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:rPr>
    <a:latin typeface="Calibri"/>
    <a:ea typeface="宋体"/>
    <a:cs typeface="Times New Roman"/>
  </a:rPr>
  <a:t>hello</a:t>
</a:r>"#;
        let r = parse_run(xml).expect("parse run");
        assert_eq!(r.text, "hello");
        assert_eq!(r.properties.latin_font.as_deref(), Some("Calibri"));
        assert_eq!(r.properties.eastasia_font.as_deref(), Some("宋体"));
        assert_eq!(r.properties.cs_font.as_deref(), Some("Times New Roman"));
    }

    /// `Sp::write_xml` 在 `is_placeholder=true` 但 `ph_type=None` 时，**必须**默认写出 `type="body"`。
    /// 早期版本会写出无属性的 `<p:ph/>`，PowerPoint 弹警告。
    #[test]
    fn write_sp_placeholder_defaults_to_body() {
        use crate::oxml::shape::Sp;
        use crate::oxml::writer::XmlWriter;
        let sp = Sp {
            id: 3,
            name: "Ph".into(),
            is_placeholder: true,
            ph_idx: Some(0),
            ph_type: None, // 关键：None 时应默认 body
            ..Default::default()
        };
        let mut w = XmlWriter::new();
        sp.write_xml(&mut w);
        let s = w.into_string();
        assert!(
            s.contains(r#"<p:ph type="body" idx="0"/>"#),
            "应写出 type=\"body\"，实际：{s}"
        );
    }

    /// `Connector::set_begin/set_end` 必须**自动**同步 xfrm 边界盒，
    /// 否则序列化后 begin/end 位置信息丢失。
    #[test]
    fn connector_set_begin_end_recomputes_xfrm() {
        use crate::oxml::simpletypes::MsoConnectorType;
        use crate::shape::connector::Connector;
        use crate::units::EmuPoint;
        let mut c = Connector::new_with_type("c1", MsoConnectorType::Straight);
        c.set_begin(EmuPoint::new(1000, 2000));
        c.set_end(EmuPoint::new(5000, 6000));
        // bounding box: off=(1000,2000), ext=(4000,4000)
        assert_eq!(c.properties().xfrm.off_x.unwrap().value(), 1000);
        assert_eq!(c.properties().xfrm.off_y.unwrap().value(), 2000);
        assert_eq!(c.properties().xfrm.ext_cx.unwrap().value(), 4000);
        assert_eq!(c.properties().xfrm.ext_cy.unwrap().value(), 4000);
    }

    /// 验证 `parse_sld_master` 能从最小 slideMaster XML 中提取 shapes 和 layout_rids。
    #[test]
    fn parse_minimal_sld_master() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Placeholder 1"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr><p:ph type="title" idx="0"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>"#;
        let master = parse_sld_master(xml).unwrap();
        assert_eq!(master.shapes.len(), 1, "应解析出 1 个 shape");
        assert_eq!(master.shapes[0].id, 2);
        assert_eq!(master.shapes[0].name, "Title Placeholder 1");
        assert!(master.shapes[0].is_placeholder);
        assert_eq!(master.shapes[0].ph_type.as_deref(), Some("title"));
        assert_eq!(master.layout_rids.len(), 1, "应解析出 1 个 layout rid");
        assert_eq!(master.layout_rids[0], "rId1");
    }

    /// 验证 `parse_sld_layout` 能从最小 slideLayout XML 中提取 type/name/shapes。
    #[test]
    fn parse_minimal_sld_layout() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr><p:ph type="title" idx="0"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="2130425"/><a:ext cx="7772400" cy="1470025"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/>
</p:sldLayout>"#;
        let layout = parse_sld_layout(xml).unwrap();
        assert_eq!(layout.type_, "title");
        assert_eq!(layout.name, "Title Slide");
        assert_eq!(layout.shapes.len(), 1, "应解析出 1 个 shape");
        assert_eq!(layout.shapes[0].name, "Title 1");
        assert!(layout.shapes[0].is_placeholder);
    }

    /// 验证 `parse_theme` 能从 theme XML 中提取 name/color_scheme/font_scheme/fmt_scheme。
    #[test]
    fn parse_theme_extracts_name_and_colors() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst/>
      <a:effectStyleLst/>
      <a:bgFillStyleLst/>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"#;
        let theme = parse_theme(xml).unwrap();
        assert_eq!(theme.name, "Office Theme");
        // 验证颜色方案名（TODO-005）
        assert_eq!(theme.color_scheme.name, "Office");
        // 验证颜色方案（用 assert_eq! + Debug 格式以便定位问题）
        assert_eq!(
            format!("{:?}", theme.color_scheme.dk1),
            r#"Some(Sys("windowText", "000000"))"#,
            "dk1 mismatch"
        );
        assert_eq!(
            format!("{:?}", theme.color_scheme.lt1),
            r#"Some(Sys("window", "FFFFFF"))"#,
            "lt1 mismatch"
        );
        assert_eq!(
            format!("{:?}", theme.color_scheme.accent1),
            r#"Some(Srgb("4F81BD"))"#,
            "accent1 mismatch"
        );
        assert_eq!(
            format!("{:?}", theme.color_scheme.hlink),
            r#"Some(Srgb("0000FF"))"#,
            "hlink mismatch"
        );
        // 验证字体方案名（TODO-005）
        assert_eq!(theme.font_scheme.name, "Office");
        // 验证字体方案
        assert_eq!(theme.font_scheme.major_latin, "Calibri");
        assert_eq!(theme.font_scheme.minor_latin, "Calibri");
        // 验证格式方案（TODO-005）
        assert_eq!(theme.format_scheme.name, "Office");
        assert!(
            !theme.format_scheme.raw_xml.is_empty(),
            "fmtScheme raw_xml should not be empty"
        );
        assert!(
            theme.format_scheme.raw_xml.contains("<a:fillStyleLst>"),
            "fmtScheme raw_xml should contain fillStyleLst"
        );
    }

    /// 验证 `parse_theme` 能解析 fontScheme 的 ea/cs typeface（TODO-005）。
    #[test]
    fn parse_theme_font_scheme_ea_cs() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
  <a:themeElements>
    <a:clrScheme name="Test"/>
    <a:fontScheme name="Test">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface="宋体"/>
        <a:cs typeface="Times New Roman"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="黑体"/>
        <a:cs typeface="Arial"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Test"/>
  </a:themeElements>
</a:theme>"#;
        let theme = parse_theme(xml).unwrap();
        assert_eq!(theme.font_scheme.major_latin, "Calibri Light");
        assert_eq!(theme.font_scheme.major_ea, "宋体");
        assert_eq!(theme.font_scheme.major_cs, "Times New Roman");
        assert_eq!(theme.font_scheme.minor_latin, "Calibri");
        assert_eq!(theme.font_scheme.minor_ea, "黑体");
        assert_eq!(theme.font_scheme.minor_cs, "Arial");
    }

    /// 验证 theme 的 round-trip：parse → to_xml → parse 一致性（TODO-005）。
    #[test]
    fn theme_round_trip() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="RoundTrip">
  <a:themeElements>
    <a:clrScheme name="RT">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="RT">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>"#;
        // 第一次解析
        let theme1 = parse_theme(xml).unwrap();
        assert_eq!(theme1.name, "RoundTrip");
        assert_eq!(theme1.color_scheme.name, "RT");
        assert_eq!(theme1.font_scheme.name, "RT");
        // 序列化回 XML
        let serialized = theme1.to_xml();
        // 第二次解析
        let theme2 = parse_theme(&serialized).unwrap();
        // 验证关键字段一致
        assert_eq!(theme2.name, theme1.name);
        assert_eq!(theme2.color_scheme.name, theme1.color_scheme.name);
        assert_eq!(theme2.color_scheme.dk1, theme1.color_scheme.dk1);
        assert_eq!(theme2.color_scheme.accent1, theme1.color_scheme.accent1);
        assert_eq!(theme2.font_scheme.name, theme1.font_scheme.name);
        assert_eq!(
            theme2.font_scheme.major_latin,
            theme1.font_scheme.major_latin
        );
        assert_eq!(
            theme2.font_scheme.minor_latin,
            theme1.font_scheme.minor_latin
        );
    }

    /// 验证 `parse_cxn_sp` 能正确解析连接器的 id/name/begin/end/stCxn/endCxn。
    ///
    /// 连接器是 OOXML 中特殊的形状（`<p:cxnSp>`），用于在形状间画线。
    /// `stCxn`/`endCxn` 是"挂接"信息——表示连接线粘附到哪个形状的哪个连接点。
    #[test]
    fn parse_cxn_sp_extracts_geometry_and_connections() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvCxnSpPr>
    <p:cNvPr id="10" name="Connector1"/>
    <p:cNvCxnSpPr/>
    <p:nvPr>
      <p:stCxn id="2" idx="3"/>
      <p:endCxn id="5" idx="1"/>
    </p:nvPr>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="1000" y="2000"/>
      <a:ext cx="5000" cy="4000"/>
    </a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:cxnSp>"#;
        let cxn = parse_cxn_sp(xml).expect("parse cxnSp ok");
        assert_eq!(cxn.id, 10);
        assert_eq!(cxn.name, "Connector1");
        // 挂接信息
        assert_eq!(cxn.st_cxn, Some((2, 3)), "stCxn 应解析出 (id=2, idx=3)");
        assert_eq!(cxn.end_cxn, Some((5, 1)), "endCxn 应解析出 (id=5, idx=1)");
        // 几何坐标（xfrm 内的 off/ext）
        assert_eq!(cxn.properties.xfrm.off_x.unwrap().value(), 1000);
        assert_eq!(cxn.properties.xfrm.off_y.unwrap().value(), 2000);
        assert_eq!(cxn.properties.xfrm.ext_cx.unwrap().value(), 5000);
        assert_eq!(cxn.properties.xfrm.ext_cy.unwrap().value(), 4000);
    }

    /// 验证 `parse_cxn_sp` 对**无挂接**的连接器也能正常解析（纯几何连接线）。
    #[test]
    fn parse_cxn_sp_without_connections() {
        let xml = r#"<p:cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvCxnSpPr><p:cNvPr id="20" name="FreeLine"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
</p:cxnSp>"#;
        let cxn = parse_cxn_sp(xml).unwrap();
        assert_eq!(cxn.id, 20);
        assert_eq!(cxn.name, "FreeLine");
        assert!(cxn.st_cxn.is_none(), "无 stCxn 时应为 None");
        assert!(cxn.end_cxn.is_none(), "无 endCxn 时应为 None");
    }

    /// 验证 `parse_cxn_sp` 能解析 `<p:style>` 主题样式引用（TODO-002 round-trip 补全）。
    #[test]
    fn parse_cxn_sp_with_style() {
        let xml = r#"<p:cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvCxnSpPr><p:cNvPr id="30" name="StyledConn"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent2"/></a:fillRef>
  </p:style>
</p:cxnSp>"#;
        let cxn = parse_cxn_sp(xml).expect("parse cxnSp ok");
        let style = cxn.style.as_ref().expect("style 应存在");
        let ln = style.line_ref.as_ref().expect("line_ref 应存在");
        assert_eq!(ln.idx.as_deref(), Some("2"));
        assert_eq!(ln.scheme_color.as_deref(), Some("accent1"));
        let fill = style.fill_ref.as_ref().expect("fill_ref 应存在");
        assert_eq!(fill.idx.as_deref(), Some("0"));
        assert_eq!(fill.scheme_color.as_deref(), Some("accent2"));
    }

    /// 验证 `parse_grp_sp` 能解析组合自身的 `<p:style>` 主题样式引用。
    #[test]
    fn parse_grp_sp_with_style() {
        let xml = r#"<p:grpSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvGrpSpPr><p:cNvPr id="40" name="StyledGroup"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="tx1"/></a:lnRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
</p:grpSp>"#;
        let grp = parse_grp_sp(xml).expect("parse grpSp ok");
        assert_eq!(grp.id, 40);
        let style = grp.style.as_ref().expect("style 应存在");
        let ln = style.line_ref.as_ref().expect("line_ref 应存在");
        assert_eq!(ln.idx.as_deref(), Some("1"));
        assert_eq!(ln.scheme_color.as_deref(), Some("tx1"));
        let font = style.font_ref.as_ref().expect("font_ref 应存在");
        assert_eq!(font.idx.as_deref(), Some("minor"));
        assert_eq!(font.scheme_color.as_deref(), Some("dk1"));
    }

    /// 验证 `parse_grp_sp` 能解析组合形状及其子形状。
    ///
    /// 组合（`<p:grpSp>`）是 OOXML 中把多个形状打包为一个整体的容器。
    /// 子形状可以是 sp/pic/cxnSp/grpSp(嵌套)/graphicFrame。
    #[test]
    fn parse_grp_sp_extracts_children() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:grpSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvGrpSpPr>
    <p:cNvPr id="100" name="Group1"/>
    <p:cNvGrpSpPr/>
    <p:nvPr/>
  </p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="8000" cy="6000"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="8000" cy="6000"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="101" name="Child1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="2000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:r><a:t>child text</a:t></a:r></a:p></p:txBody>
  </p:sp>
  <p:cxnSp>
    <p:nvCxnSpPr><p:cNvPr id="102" name="Link"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000" cy="3000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
  </p:cxnSp>
</p:grpSp>"#;
        let grp = parse_grp_sp(xml).expect("parse grpSp ok");
        assert_eq!(grp.id, 100);
        assert_eq!(grp.name, "Group1");
        assert_eq!(grp.children.len(), 2, "应解析出 2 个子形状");

        // 第一个子形状：Sp
        match &grp.children[0] {
            GroupChild::Sp(sp) => {
                assert_eq!(sp.id, 101);
                assert_eq!(sp.name, "Child1");
                assert_eq!(sp.text.paragraphs[0].runs[0].text, "child text");
            }
            other => panic!("第一个子形状应为 Sp，实际：{other:?}"),
        }
        // 第二个子形状：CxnSp
        match &grp.children[1] {
            GroupChild::CxnSp(cxn) => {
                assert_eq!(cxn.id, 102);
                assert_eq!(cxn.name, "Link");
            }
            other => panic!("第二个子形状应为 CxnSp，实际：{other:?}"),
        }
    }

    /// 验证 `parse_graphic_frame` 能解析承载表格的图形框。
    ///
    /// `<p:graphicFrame>` 是 OOXML 中承载表格/图表等复合图形的容器。
    /// 内部 `<a:graphicData uri="...">` 的 uri 决定具体类型：
    /// - `http://schemas.openxmlformats.org/drawingml/2006/table` → 表格
    #[test]
    fn parse_graphic_frame_with_table() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvGraphicFramePr>
    <p:cNvPr id="200" name="TableFrame"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="457200" y="1828800"/>
    <a:ext cx="8229600" cy="1143000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1"/>
        <a:tblGrid>
          <a:gridCol w="4114800"/>
          <a:gridCol w="4114800"/>
        </a:tblGrid>
        <a:tr h="370840" h="0">
          <a:tc>
            <a:txBody><a:bodyPr/><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody><a:bodyPr/><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
        <a:tr h="370840" h="0">
          <a:tc>
            <a:txBody><a:bodyPr/><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody>
            <a:tcPr/>
          </a:tc>
          <a:tc>
            <a:txBody><a:bodyPr/><a:p><a:r><a:t>B2</a:t></a:r></a:p></a:txBody>
            <a:tcPr/>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#;
        let frame = parse_graphic_frame(xml).expect("parse graphicFrame ok");
        assert_eq!(frame.id, 200);
        assert_eq!(frame.name, "TableFrame");
        // xfrm
        assert_eq!(frame.properties.xfrm.off_x.unwrap().value(), 457200);
        assert_eq!(frame.properties.xfrm.off_y.unwrap().value(), 1828800);
        assert_eq!(frame.properties.xfrm.ext_cx.unwrap().value(), 8229600);
        assert_eq!(frame.properties.xfrm.ext_cy.unwrap().value(), 1143000);
        // 表格内容
        match &frame.graphic {
            crate::oxml::shape::Graphic::Table(tbl) => {
                assert_eq!(tbl.cols.len(), 2, "应有 2 列");
                assert_eq!(tbl.rows.len(), 2, "应有 2 行");
                assert_eq!(tbl.cols[0].width.value(), 4114800);
                // 验证单元格文本
                let cell_a1 = &tbl.rows[0].cells[0];
                let text_a1: String = cell_a1
                    .text
                    .paragraphs
                    .iter()
                    .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
                    .collect();
                assert_eq!(text_a1, "A1");
                let cell_b2 = &tbl.rows[1].cells[1];
                let text_b2: String = cell_b2
                    .text
                    .paragraphs
                    .iter()
                    .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
                    .collect();
                assert_eq!(text_b2, "B2");
            }
            _ => panic!("期望 Table 变体"),
        }
    }

    /// 验证 `parse_sld` 能解析包含**多种形状类型**的 slide（Sp + Pic + CxnSp + Group）。
    ///
    /// 这是 round-trip 测试的基础：如果 parse 阶段就丢失形状，save 后内容必然缺失。
    #[test]
    fn parse_sld_with_mixed_shape_types() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="2000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>text</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Pic1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdImg1"/></p:blipFill>
        <p:spPr><a:xfrm><a:off x="5000" y="5000"/><a:ext cx="3000" cy="2000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="4" name="Conn1"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="8000" cy="6000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
      </p:cxnSp>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="5" name="Group1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="6" name="Inner"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>inner</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse sld ok");
        assert_eq!(sld.shapes.len(), 4, "应解析出 4 个形状");

        // 验证每个形状的类型和关键字段
        assert!(
            matches!(&sld.shapes[0], OxmlSlideShape::Sp(s) if s.id == 2 && s.name == "TextBox")
        );
        assert!(
            matches!(&sld.shapes[1], OxmlSlideShape::Pic(p) if p.id == 3 && p.rid == "rIdImg1")
        );
        assert!(
            matches!(&sld.shapes[2], OxmlSlideShape::CxnSp(c) if c.id == 4 && c.name == "Conn1")
        );
        match &sld.shapes[3] {
            OxmlSlideShape::Group(grp) => {
                assert_eq!(grp.id, 5);
                assert_eq!(grp.name, "Group1");
                assert_eq!(grp.children.len(), 1, "组合内应有 1 个子形状");
                assert!(
                    matches!(&grp.children[0], GroupChild::Sp(s) if s.id == 6 && s.name == "Inner")
                );
            }
            other => panic!("第 4 个形状应为 Group，实际：{other:?}"),
        }
    }

    /// **Round-trip 保留测试**：parse → write → parse，验证关键字段不丢失。
    ///
    /// 这是 TODO-041 的核心测试：打开已有 PPTX 后修改再保存，
    /// 形状的 id/name/文本/几何坐标/挂接信息必须全部保留。
    #[test]
    fn round_trip_sld_preserves_shapes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TB1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="3657600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>round-trip</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="3" name="C1"/><p:cNvCxnSpPr/><p:nvPr><p:stCxn id="2" idx="0"/><p:endCxn id="5" idx="1"/></p:nvPr></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="4000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;

        // 第一次解析
        let sld1 = parse_sld(xml).expect("first parse ok");
        assert_eq!(sld1.shapes.len(), 2);

        // 序列化回 XML
        let rewritten = sld1.to_xml();

        // 第二次解析（round-trip）
        let sld2 = parse_sld(&rewritten).expect("second parse ok");
        assert_eq!(sld2.shapes.len(), 2, "round-trip 后形状数应不变");

        // 验证 Sp
        match (&sld1.shapes[0], &sld2.shapes[0]) {
            (OxmlSlideShape::Sp(a), OxmlSlideShape::Sp(b)) => {
                assert_eq!(a.id, b.id, "Sp id 应保留");
                assert_eq!(a.name, b.name, "Sp name 应保留");
                assert_eq!(
                    a.text.paragraphs[0].runs[0].text, b.text.paragraphs[0].runs[0].text,
                    "Sp 文本应保留"
                );
                // 几何坐标
                assert_eq!(
                    a.properties.xfrm.off_x, b.properties.xfrm.off_x,
                    "Sp off_x 应保留"
                );
                assert_eq!(
                    a.properties.xfrm.ext_cx, b.properties.xfrm.ext_cx,
                    "Sp ext_cx 应保留"
                );
            }
            _ => panic!("形状类型不匹配"),
        }

        // 验证 CxnSp（含挂接信息）
        match (&sld1.shapes[1], &sld2.shapes[1]) {
            (OxmlSlideShape::CxnSp(a), OxmlSlideShape::CxnSp(b)) => {
                assert_eq!(a.id, b.id, "CxnSp id 应保留");
                assert_eq!(a.name, b.name, "CxnSp name 应保留");
                assert_eq!(a.st_cxn, b.st_cxn, "stCxn 挂接应保留");
                assert_eq!(a.end_cxn, b.end_cxn, "endCxn 挂接应保留");
            }
            _ => panic!("形状类型不匹配"),
        }
    }

    /// **Round-trip 表格保留测试**：parse → write → parse，验证表格结构/单元格文本不丢失。
    #[test]
    fn round_trip_table_preserves_cells() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="10" name="Tbl1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="8000" cy="4000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr/>
              <a:tblGrid><a:gridCol w="4000"/><a:gridCol w="4000"/></a:tblGrid>
              <a:tr h="2000">
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>R1C1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>R1C2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;

        let sld1 = parse_sld(xml).expect("first parse ok");
        let rewritten = sld1.to_xml();
        let sld2 = parse_sld(&rewritten).expect("second parse ok");

        assert_eq!(sld1.shapes.len(), sld2.shapes.len(), "形状数应一致");
        match (&sld1.shapes[0], &sld2.shapes[0]) {
            (OxmlSlideShape::GraphicFrame(f1), OxmlSlideShape::GraphicFrame(f2)) => {
                assert_eq!(f1.id, f2.id, "GraphicFrame id 应保留");
                assert_eq!(f1.name, f2.name, "GraphicFrame name 应保留");
                match (&f1.graphic, &f2.graphic) {
                    (
                        crate::oxml::shape::Graphic::Table(t1),
                        crate::oxml::shape::Graphic::Table(t2),
                    ) => {
                        assert_eq!(t1.cols.len(), t2.cols.len(), "列数应保留");
                        assert_eq!(t1.rows.len(), t2.rows.len(), "行数应保留");
                        // 验证单元格文本
                        let get_text = |t: &crate::oxml::table::Table| -> String {
                            t.rows[0].cells[0]
                                .text
                                .paragraphs
                                .iter()
                                .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
                                .collect()
                        };
                        assert_eq!(get_text(t1), get_text(t2), "单元格文本应保留");
                        assert_eq!(get_text(t1), "R1C1");
                    }
                    _ => panic!("期望两侧均为 Table 变体"),
                }
            }
            _ => panic!("形状类型不匹配"),
        }
    }

    /// 验证 `parse_sld` 对**空 spTree**（无形状）的容错。
    #[test]
    fn parse_sld_empty_sp_tree() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        assert_eq!(sld.shapes.len(), 0, "空 spTree 应解析出 0 个形状");
    }

    /// 验证 `parse_grp_sp` 能正确解析**嵌套组合**（grpSp 内有 grpSp）。
    ///
    /// 早期版本因 `if local == b"grpSp"` 分支重复，嵌套组合被根元素分支吞掉，
    /// 导致嵌套 grpSp 的子形状全部丢失。本测试回归此 bug。
    #[test]
    fn parse_grp_sp_nested_group() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:grpSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
         xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvGrpSpPr><p:cNvPr id="1" name="Outer"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000" cy="10000"/><a:chOff x="0" y="0"/><a:chExt cx="10000" cy="10000"/></a:xfrm></p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="OuterChild"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:r><a:t>outer</a:t></a:r></a:p></p:txBody>
  </p:sp>
  <p:grpSp>
    <p:nvGrpSpPr><p:cNvPr id="3" name="Inner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="2000" y="2000"/><a:ext cx="5000" cy="5000"/><a:chOff x="0" y="0"/><a:chExt cx="5000" cy="5000"/></a:xfrm></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="4" name="InnerChild"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="50" y="50"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>inner</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:grpSp>
</p:grpSp>"#;
        let grp = parse_grp_sp(xml).expect("parse ok");
        assert_eq!(grp.id, 1);
        assert_eq!(grp.name, "Outer");
        assert_eq!(
            grp.children.len(),
            2,
            "外层组合应有 2 个子形状（sp + 嵌套 grpSp）"
        );

        // 第一个子形状：外层 Sp
        assert!(
            matches!(&grp.children[0], GroupChild::Sp(s) if s.id == 2 && s.name == "OuterChild")
        );

        // 第二个子形状：嵌套组合
        match &grp.children[1] {
            GroupChild::Group(inner) => {
                assert_eq!(inner.id, 3);
                assert_eq!(inner.name, "Inner");
                assert_eq!(inner.children.len(), 1, "内层组合应有 1 个子形状");
                assert!(
                    matches!(&inner.children[0], GroupChild::Sp(s) if s.id == 4 && s.name == "InnerChild")
                );
            }
            other => panic!("第二个子形状应为 Group，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**线性渐变填充**（`<a:gradFill><a:gsLst>...<a:lin ang="..."/></a:gradFill>`）。
    ///
    /// 这是 TODO-003 的核心测试：确保渐变光轨、角度、颜色全部正确提取。
    #[test]
    fn parse_sppr_grad_fill_linear() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:gradFill flip="none" rotWithShape="1">
    <a:gsLst>
      <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
      <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
    </a:gsLst>
    <a:lin ang="5400000" scaled="1"/>
  </a:gradFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Gradient(grad) => {
                assert_eq!(grad.stops.len(), 2, "应有 2 个渐变光轨");
                assert_eq!(grad.stops[0].pos, 0);
                assert_eq!(grad.stops[1].pos, 100000);
                // 验证颜色
                assert!(
                    matches!(grad.stops[0].color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0x00 && c.2 == 0x00)
                );
                assert!(
                    matches!(grad.stops[1].color, Color::RGB(c) if c.0 == 0x00 && c.1 == 0x00 && c.2 == 0xFF)
                );
                // 验证线性角度（5400000 = 90° = 向下）
                assert_eq!(grad.gradient_type, GradientType::Linear(5400000));
                // 验证属性
                assert_eq!(grad.flip.as_deref(), Some("none"));
                assert_eq!(grad.rot_with_shape, Some(true));
            }
            other => panic!("fill 应为 Gradient，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**路径渐变填充**（`<a:gradFill><a:path path="circle"/></a:gradFill>`）。
    #[test]
    fn parse_sppr_grad_fill_path() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:gradFill>
    <a:gsLst>
      <a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs>
      <a:gs pos="50000"><a:srgbClr val="FFFF00"/></a:gs>
      <a:gs pos="100000"><a:srgbClr val="FF00FF"/></a:gs>
    </a:gsLst>
    <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
  </a:gradFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Gradient(grad) => {
                assert_eq!(grad.stops.len(), 3, "应有 3 个渐变光轨");
                assert_eq!(grad.stops[1].pos, 50000);
                // 验证路径渐变类型
                assert_eq!(grad.gradient_type, GradientType::Path(GradientPath::Circle));
            }
            other => panic!("fill 应为 Gradient，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**图案填充**（`<a:pattFill prst="..."><a:fgClr/>...<a:bgClr/>...</a:pattFill>`）。
    #[test]
    fn parse_sppr_patt_fill() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:pattFill prst="pct5">
    <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
    <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
  </a:pattFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Pattern(patt) => {
                assert_eq!(patt.prst, "pct5", "预置图案应为 pct5");
                // 前景色：红色
                assert!(
                    matches!(patt.fg_color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0x00 && c.2 == 0x00)
                );
                // 背景色：白色
                assert!(
                    matches!(patt.bg_color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0xFF && c.2 == 0xFF)
                );
            }
            other => panic!("fill 应为 Pattern，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**图片填充（blipFill）**的 stretch 模式（TODO-003/048）。
    ///
    /// 这是 TODO-003/048 的核心测试：确保 `<a:blip r:embed="rIdN"/>` + `<a:stretch>`
    /// 被正确解析为 `Fill::Blip { rid, mode: Stretch }`。
    #[test]
    fn parse_sppr_blip_fill_stretch() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:blipFill>
    <a:blip r:embed="rId1"/>
    <a:stretch><a:fillRect/></a:stretch>
  </a:blipFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Blip { rid, mode } => {
                assert_eq!(rid, "rId1", "rid 应为 rId1");
                assert!(
                    matches!(mode, crate::oxml::sppr::BlipFillMode::Stretch),
                    "mode 应为 Stretch，实际：{:?}",
                    mode
                );
            }
            other => panic!("fill 应为 Blip，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**图片填充（blipFill）**的 tile 模式 + 自闭合 blip。
    #[test]
    fn parse_sppr_blip_fill_tile() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:blipFill>
    <a:blip r:embed="rIdImg7"/>
    <a:tile tx="100" ty="200" sx="50000" sy="50000" flip="x" algn="ctr"/>
  </a:blipFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Blip { rid, mode } => {
                assert_eq!(rid, "rIdImg7", "rid 应为 rIdImg7");
                match mode {
                    crate::oxml::sppr::BlipFillMode::Tile {
                        tx,
                        ty,
                        sx,
                        sy,
                        flip,
                        algn,
                    } => {
                        assert_eq!(*tx, Some(100), "tx 应为 100");
                        assert_eq!(*ty, Some(200), "ty 应为 200");
                        assert_eq!(*sx, Some(50000), "sx 应为 50000");
                        assert_eq!(*sy, Some(50000), "sy 应为 50000");
                        assert_eq!(flip.as_deref(), Some("x"), "flip 应为 x");
                        assert_eq!(algn.as_deref(), Some("ctr"), "algn 应为 ctr");
                    }
                    other => panic!("mode 应为 Tile，实际：{other:?}"),
                }
            }
            other => panic!("fill 应为 Blip，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**自闭合 blipFill 子元素**（`<a:blip/>` + `<a:stretch/>` 均为 Empty）。
    #[test]
    fn parse_sppr_blip_fill_self_closing() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch/>
  </a:blipFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Blip { rid, mode } => {
                assert_eq!(rid, "rId2");
                assert!(matches!(mode, crate::oxml::sppr::BlipFillMode::Stretch));
            }
            other => panic!("fill 应为 Blip，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析**自闭合的渐变填充**（`<a:lin/>` 为 Empty 事件）。
    ///
    /// 某些 PPTX 生成器会把 `<a:lin ang="..." scaled="..."/>` 写成自闭合标签，
    /// 本测试确保 Empty 分支也能正确提取角度。
    #[test]
    fn parse_sppr_grad_fill_self_closing_lin() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:gradFill>
    <a:gsLst>
      <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
      <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
    </a:gsLst>
    <a:lin ang="2700000" scaled="0"/>
  </a:gradFill>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        match &sp.fill {
            Fill::Gradient(grad) => {
                assert_eq!(grad.gradient_type, GradientType::Linear(2700000));
                assert_eq!(grad.stops.len(), 2);
            }
            other => panic!("fill 应为 Gradient，实际：{other:?}"),
        }
    }

    /// 验证 `parse_ln` 能解析**箭头端点**（`<a:headEnd>` / `<a:tailEnd>`）。
    ///
    /// 这是 TODO-012 的核心测试：确保线条起止箭头的 type/w/len 属性全部正确提取。
    #[test]
    fn parse_sppr_ln_with_arrow_heads() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln w="12700">
    <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
    <a:headEnd type="triangle" w="lg" len="med"/>
    <a:tailEnd type="stealth" w="sm" len="lg"/>
  </a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        // headEnd
        let head = ln.head_end.expect("应有 headEnd");
        assert_eq!(head.arrow_type, ArrowType::Triangle);
        assert_eq!(head.width, ArrowSize::Large);
        assert_eq!(head.length, ArrowSize::Medium);
        // tailEnd
        let tail = ln.tail_end.expect("应有 tailEnd");
        assert_eq!(tail.arrow_type, ArrowType::Stealth);
        assert_eq!(tail.width, ArrowSize::Small);
        assert_eq!(tail.length, ArrowSize::Large);
    }

    /// 验证 `parse_ln` 能解析**连接类型**（`<a:round>` / `<a:miter>` / `<a:bevel>`）。
    #[test]
    fn parse_sppr_ln_with_join() {
        // round 连接
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/></a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        assert_eq!(ln.join, Some(LineJoin::Round));

        // miter 连接（带 lim 属性）
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="600000"/></a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        assert_eq!(ln.join, Some(LineJoin::Miter(600000)));

        // bevel 连接
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:bevel/></a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        assert_eq!(ln.join, Some(LineJoin::Bevel));
    }

    /// 验证 `parse_ln` 能解析**线条渐变填充**（`<a:gradFill>` 在 `<a:ln>` 内）。
    #[test]
    fn parse_sppr_ln_with_grad_fill() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln w="12700">
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs>
      </a:gsLst>
      <a:lin ang="0" scaled="1"/>
    </a:gradFill>
  </a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        match &ln.fill {
            Fill::Gradient(grad) => {
                assert_eq!(grad.stops.len(), 2);
                assert_eq!(grad.gradient_type, GradientType::Linear(0));
            }
            other => panic!("line fill 应为 Gradient，实际：{other:?}"),
        }
    }

    /// 验证 `parse_ln` 能解析**线条图案填充**（`<a:pattFill>` 在 `<a:ln>` 内）。
    #[test]
    fn parse_sppr_ln_with_patt_fill() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:ln w="12700">
    <a:pattFill prst="cross">
      <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
      <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>
    </a:pattFill>
  </a:ln>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let ln = sp.line.expect("应有 line");
        match &ln.fill {
            Fill::Pattern(patt) => {
                assert_eq!(patt.prst, "cross");
                assert!(matches!(patt.fg_color, Color::RGB(c) if c.0 == 0xFF));
                assert!(
                    matches!(patt.bg_color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0xFF && c.2 == 0xFF)
                );
            }
            other => panic!("line fill 应为 Pattern，实际：{other:?}"),
        }
    }

    /// 验证 `parse_sppr` 能解析 `<a:effectLst>` 中的外阴影（TODO-011）。
    #[test]
    fn parse_sppr_effect_lst_outer_shadow() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:effectLst>
    <a:outerShdw blurRad="40000" dist="38100" dir="2700000" rotWithShape="0">
      <a:srgbClr val="000000"/>
    </a:outerShdw>
  </a:effectLst>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let effects = sp.effects.expect("应有 effects");
        let shdw = effects.outer_shadow.expect("应有 outer_shadow");
        assert_eq!(shdw.dir, 2_700_000);
        assert_eq!(shdw.dist, 38100);
        assert_eq!(shdw.blur_rad, 40000);
        assert_eq!(shdw.rot_with_shape, Some(false));
        assert!(matches!(shdw.color, Color::RGB(c) if c.0 == 0x00 && c.1 == 0x00 && c.2 == 0x00));
    }

    /// 验证 `parse_sppr` 能解析发光和柔化边缘效果。
    #[test]
    fn parse_sppr_effect_lst_glow_and_soft_edge() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:effectLst>
    <a:glow rad="50000">
      <a:srgbClr val="FF00FF"/>
    </a:glow>
    <a:softEdge rad="25000"/>
  </a:effectLst>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let effects = sp.effects.expect("应有 effects");
        let glow = effects.glow.expect("应有 glow");
        assert_eq!(glow.rad, 50000);
        assert!(matches!(glow.color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0x00 && c.2 == 0xFF));
        let se = effects.soft_edge.expect("应有 soft_edge");
        assert_eq!(se.rad, 25000);
    }

    /// 验证 `parse_sppr` 能解析反射效果（仅属性）。
    #[test]
    fn parse_sppr_effect_lst_reflection() {
        let xml = r#"<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:effectLst>
    <a:reflection blurRad="50000" stA="52000" stPos="0" endA="30000" endPos="50000" dist="38100" dir="5400000"/>
  </a:effectLst>
</p:spPr>"#;
        let sp = parse_sppr(xml).expect("parse ok");
        let effects = sp.effects.expect("应有 effects");
        let refl = effects.reflection.expect("应有 reflection");
        assert_eq!(refl.blur_rad, Some(50000));
        assert_eq!(refl.st_a, Some(52000));
        assert_eq!(refl.st_pos, Some(0));
        assert_eq!(refl.end_a, Some(30000));
        assert_eq!(refl.end_pos, Some(50000));
        assert_eq!(refl.dist, Some(38100));
        assert_eq!(refl.dir, Some(5_400_000));
    }

    /// 验证 `Cell` 的 `gridSpan`/`rowSpan`/`hMerge`/`vMerge` 属性在 XML 往返中保持不变。
    ///
    /// 这是 TODO-029 的核心测试：确保合并单元格信息在序列化 → 解析后不丢失。
    #[test]
    fn round_trip_cell_merge_attributes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="MergedTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr>
            <a:tblGrid>
              <a:gridCol w="2000000"/>
              <a:gridCol w="2000000"/>
              <a:gridCol w="2000000"/>
            </a:tblGrid>
            <a:tr h="2000000">
              <a:tc gridSpan="3" rowSpan="2">
                <a:txBody><a:bodyPr/><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
            <a:tr h="2000000">
              <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
              <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
              <a:tc vMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        // 找到 graphicFrame
        let frame = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::GraphicFrame(g) = s {
                    Some(g)
                } else {
                    None
                }
            })
            .expect("应有 graphicFrame");
        match &frame.graphic {
            crate::oxml::shape::Graphic::Table(tbl) => {
                assert_eq!(tbl.rows.len(), 2, "应有 2 行");
                // 验证 tableStyleId 解析（TODO-030）
                assert_eq!(
                    tbl.table_style
                        .as_ref()
                        .expect("应有 table_style")
                        .style_id(),
                    "{5940675A-B579-460E-94D1-54222C63F5DA}"
                );
                // 第一行：合并源单元格
                let origin = &tbl.rows[0].cells[0];
                assert_eq!(origin.grid_span, 3, "gridSpan 应为 3");
                assert_eq!(origin.row_span, 2, "rowSpan 应为 2");
                assert!(!origin.h_merge, "合并源不应有 hMerge");
                assert!(!origin.v_merge, "合并源不应有 vMerge");
                // 第二行：被合并方
                let h1 = &tbl.rows[1].cells[0];
                assert!(h1.h_merge, "应有 hMerge");
                let v1 = &tbl.rows[1].cells[2];
                assert!(v1.v_merge, "应有 vMerge");
            }
            _ => panic!("期望 Table 变体"),
        }
    }

    /// 验证表格样式（`<a:tableStyleId>`）在序列化 → 解析后保持不变。
    ///
    /// 这是 TODO-030 的核心 round-trip 测试。
    #[test]
    fn round_trip_table_style_id() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="StyledTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
            <a:tr h="2000000">
              <a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
              <a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let frame = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::GraphicFrame(g) = s {
                    Some(g)
                } else {
                    None
                }
            })
            .expect("应有 graphicFrame");
        match &frame.graphic {
            crate::oxml::shape::Graphic::Table(tbl) => {
                // 验证 tableStyleId 解析
                let style = tbl.table_style.as_ref().expect("应有 table_style");
                assert_eq!(style.style_id(), "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
                // 验证序列化后仍包含 tableStyleId
                let mut w = crate::oxml::writer::XmlWriter::new();
                tbl.write_xml(&mut w);
                let out = &w.buf;
                assert!(out.contains(
                    "<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>"
                ));
            }
            _ => panic!("期望 Table 变体"),
        }
    }

    /// 验证表格无 `<a:tableStyleId>` 时 `table_style` 为 `None`。
    #[test]
    fn parse_table_without_style_id() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="PlainTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr/>
            <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
            <a:tr h="2000000">
              <a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
              <a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody><a:tcPr/></a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let frame = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::GraphicFrame(g) = s {
                    Some(g)
                } else {
                    None
                }
            })
            .expect("应有 graphicFrame");
        match &frame.graphic {
            crate::oxml::shape::Graphic::Table(tbl) => {
                assert!(
                    tbl.table_style.is_none(),
                    "无 tableStyleId 时 table_style 应为 None"
                );
            }
            _ => panic!("期望 Table 变体"),
        }
    }

    /// 验证 `Cell` 的边框（lnL/lnR/lnT/lnB）、边距（marT/marL/marB/marR）、垂直对齐（anchor）
    /// 在 XML 往返中保持不变。
    ///
    /// 这是 TODO-013 的核心测试：确保单元格格式化信息在序列化 → 解析后不丢失。
    #[test]
    fn round_trip_cell_borders_margins_anchor() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="2" name="FormattedTable"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl>
            <a:tblPr><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr>
            <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
            <a:tr h="2000000">
              <a:tc>
                <a:txBody><a:bodyPr/><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody>
                <a:tcPr marT="50000" marL="60000" marB="70000" marR="80000" anchor="ctr">
                  <a:lnL w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>
                  <a:lnR w="9525"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnR>
                  <a:lnT w="9525"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:lnT>
                  <a:lnB w="9525"><a:noFill/></a:lnB>
                </a:tcPr>
              </a:tc>
              <a:tc>
                <a:txBody><a:bodyPr/><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody>
                <a:tcPr/>
              </a:tc>
            </a:tr>
          </a:tbl>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let frame = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::GraphicFrame(g) = s {
                    Some(g)
                } else {
                    None
                }
            })
            .expect("应有 graphicFrame");
        let tbl = match &frame.graphic {
            crate::oxml::shape::Graphic::Table(t) => t,
            _ => panic!("期望 Table 变体"),
        };
        let cell = &tbl.rows[0].cells[0];
        // 边距
        assert_eq!(cell.margin.0.map(|m| m.value()), Some(50000), "marT");
        assert_eq!(cell.margin.1.map(|m| m.value()), Some(60000), "marL");
        assert_eq!(cell.margin.2.map(|m| m.value()), Some(70000), "marB");
        assert_eq!(cell.margin.3.map(|m| m.value()), Some(80000), "marR");
        // 垂直对齐
        assert_eq!(
            cell.anchor,
            crate::oxml::table::VerticalAnchor::Middle,
            "anchor=ctr"
        );
        // 边框
        let bl = cell.border_left.as_ref().expect("应有 lnL");
        assert_eq!(bl.width.value(), 9525);
        assert!(matches!(bl.color, Color::RGB(c) if c.0 == 0xFF && c.1 == 0x00 && c.2 == 0x00));
        let br = cell.border_right.as_ref().expect("应有 lnR");
        assert!(matches!(br.color, Color::RGB(c) if c.0 == 0x00 && c.1 == 0xFF && c.2 == 0x00));
        let bt = cell.border_top.as_ref().expect("应有 lnT");
        assert!(matches!(bt.color, Color::RGB(c) if c.0 == 0x00 && c.1 == 0x00 && c.2 == 0xFF));
        let bb = cell.border_bottom.as_ref().expect("应有 lnB");
        assert!(bb.no_fill, "lnB 应为 noFill");
    }

    /// 验证 `<a:bodyPr>` 的 `numCol` / `spcCol` / `anchor` / `wrap` 属性解析。
    ///
    /// 这是 TODO-019 的 round-trip 测试。
    #[test]
    fn parse_txbody_multi_column() {
        let xml = r#"<p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:bodyPr numCol="3" spcCol="91440" anchor="ctr" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
  <a:p><a:r><a:t>multi-col text</a:t></a:r></a:p>
</p:txBody>"#;
        let tb = parse_txbody(xml).expect("parse ok");
        let bp = tb.body_properties.as_ref().expect("应有 body_properties");
        assert_eq!(bp.num_cols, Some(3), "numCol");
        assert_eq!(bp.col_spacing.map(|v| v.value()), Some(91440), "spcCol");
        assert_eq!(bp.anchor, Some(MsoAnchor::Middle), "anchor=ctr");
        assert_eq!(bp.wrap, Some(TextWrapping::Square), "wrap=square");
        let insets = bp.insets.as_ref().expect("应有 insets");
        assert_eq!(insets.left.value(), 91440, "lIns");
        assert_eq!(insets.top.value(), 45720, "tIns");
        assert_eq!(insets.right.value(), 91440, "rIns");
        assert_eq!(insets.bottom.value(), 45720, "bIns");
        // 段落仍应正确解析
        assert_eq!(tb.paragraphs.len(), 1, "应有一个段落");
        assert_eq!(tb.paragraphs[0].runs.len(), 1, "应有一个 run");
        assert_eq!(tb.paragraphs[0].runs[0].text, "multi-col text");
    }

    /// 验证自闭合 `<a:bodyPr/>` 的属性解析。
    ///
    /// 这是 TODO-019 的测试。
    #[test]
    fn parse_txbody_self_closing_body_pr() {
        let xml = r#"<p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:bodyPr numCol="2" spcCol="45720"/>
  <a:p/>
</p:txBody>"#;
        let tb = parse_txbody(xml).expect("parse ok");
        let bp = tb.body_properties.as_ref().expect("应有 body_properties");
        assert_eq!(bp.num_cols, Some(2), "numCol");
        assert_eq!(bp.col_spacing.map(|v| v.value()), Some(45720), "spcCol");
    }

    /// 验证 `<a:bodyPr>` 含 `<a:spAutoFit/>` 子元素的解析。
    ///
    /// 这是 TODO-019 的测试（顺带覆盖 autoSize 子元素解析）。
    #[test]
    fn parse_txbody_with_sp_auto_fit() {
        let xml = r#"<p:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <a:bodyPr numCol="1"><a:spAutoFit/></a:bodyPr>
  <a:p/>
</p:txBody>"#;
        let tb = parse_txbody(xml).expect("parse ok");
        let bp = tb.body_properties.as_ref().expect("应有 body_properties");
        assert_eq!(bp.num_cols, Some(1), "numCol=1");
        assert!(bp.sp_auto_fit, "应启用 spAutoFit");
        assert!(!bp.norm_autofit, "不应启用 normAutofit");
    }

    /// 验证 `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` 的详细属性解析。
    ///
    /// 这是 TODO-014 的 round-trip 测试。
    #[test]
    fn parse_paragraph_bullet_styles() {
        // buChar
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buChar char="•"/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        assert!(ppr.bullet, "bullet 应为 true");
        match &ppr.bullet_style {
            Some(BulletStyle::Char { char }) => assert_eq!(char, "•"),
            other => panic!("期望 Char，实际: {other:?}"),
        }

        // buAutoNum with startAt
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buAutoNum type="arabicPeriod" startAt="3"/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        assert!(ppr.bullet);
        match &ppr.bullet_style {
            Some(BulletStyle::AutoNum {
                auto_num_type,
                start_at,
            }) => {
                assert_eq!(auto_num_type, "arabicPeriod");
                assert_eq!(*start_at, Some(3));
            }
            other => panic!("期望 AutoNum，实际: {other:?}"),
        }

        // buAutoNum without startAt
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buAutoNum type="alphaLcParenR"/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        match &ppr.bullet_style {
            Some(BulletStyle::AutoNum {
                auto_num_type,
                start_at,
            }) => {
                assert_eq!(auto_num_type, "alphaLcParenR");
                assert_eq!(*start_at, None);
            }
            other => panic!("期望 AutoNum，实际: {other:?}"),
        }

        // buNone
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buNone/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        assert!(!ppr.bullet, "bullet 应为 false");
        assert!(matches!(ppr.bullet_style, Some(BulletStyle::None)));
    }

    /// 验证自闭合的 `<a:buChar/>` / `<a:buAutoNum/>` / `<a:buNone/>` 解析。
    ///
    /// 这是 TODO-014 的测试。
    #[test]
    fn parse_paragraph_bullet_styles_self_closing() {
        // 自闭合 buChar（无 char 属性）
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buChar char="▪"/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        match &ppr.bullet_style {
            Some(BulletStyle::Char { char }) => assert_eq!(char, "▪"),
            other => panic!("期望 Char，实际: {other:?}"),
        }

        // 自闭合 buAutoNum
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:buAutoNum type="romanLcParenBoth" startAt="5"/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        match &ppr.bullet_style {
            Some(BulletStyle::AutoNum {
                auto_num_type,
                start_at,
            }) => {
                assert_eq!(auto_num_type, "romanLcParenBoth");
                assert_eq!(*start_at, Some(5));
            }
            other => panic!("期望 AutoNum，实际: {other:?}"),
        }
    }

    /// 验证 `<a:tabLst>` / `<a:tab>` 的解析。
    ///
    /// 这是 TODO-015 的 round-trip 测试。
    #[test]
    fn parse_paragraph_tab_stops() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:tabLst>
    <a:tab pos="914400" algn="l"/>
    <a:tab pos="1828800" algn="r"/>
    <a:tab pos="2743200" algn="ctr"/>
    <a:tab pos="3657600" algn="dec"/>
  </a:tabLst>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        assert_eq!(ppr.tab_stops.len(), 4, "应有 4 个制表位");
        assert_eq!(ppr.tab_stops[0].pos.value(), 914400);
        assert_eq!(ppr.tab_stops[0].alignment, TabAlignment::Left);
        assert_eq!(ppr.tab_stops[1].pos.value(), 1828800);
        assert_eq!(ppr.tab_stops[1].alignment, TabAlignment::Right);
        assert_eq!(ppr.tab_stops[2].pos.value(), 2743200);
        assert_eq!(ppr.tab_stops[2].alignment, TabAlignment::Center);
        assert_eq!(ppr.tab_stops[3].pos.value(), 3657600);
        assert_eq!(ppr.tab_stops[3].alignment, TabAlignment::Decimal);
    }

    /// 验证空 `<a:tabLst/>` 的解析。
    ///
    /// 这是 TODO-015 的测试。
    #[test]
    fn parse_paragraph_empty_tab_lst() {
        let xml = r#"<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:tabLst/>
</a:pPr>"#;
        let ppr = parse_paragraph_properties(xml).expect("parse ok");
        assert!(ppr.tab_stops.is_empty(), "空 tabLst 应无制表位");
    }

    /// 验证 `<a:fld>` 字段元素的解析。
    ///
    /// 这是 TODO-016 的 round-trip 测试。
    #[test]
    fn parse_paragraph_field() {
        let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:fld id="{12345678-ABCD-EF01-2345-678901234567}" type="slidenum">
    <a:rPr lang="en-US"/>
    <a:t>1</a:t>
  </a:fld>
</a:p>"#;
        let p = parse_paragraph(xml).expect("parse ok");
        assert_eq!(p.fields.len(), 1, "应有 1 个字段");
        let f = &p.fields[0];
        assert_eq!(f.id, "{12345678-ABCD-EF01-2345-678901234567}");
        assert_eq!(f.field_type, FieldType::SlideNumber);
        assert_eq!(f.text, "1");
    }

    /// 验证含多个字段和 Run 的段落解析。
    ///
    /// 这是 TODO-016 的测试。
    #[test]
    fn parse_paragraph_mixed_runs_and_fields() {
        let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:r><a:t>Page </a:t></a:r>
  <a:fld id="{AAA}" type="slidenum"><a:t>1</a:t></a:fld>
  <a:fld id="{BBB}" type="datetime"><a:t>1/1/2024</a:t></a:fld>
</a:p>"#;
        let p = parse_paragraph(xml).expect("parse ok");
        assert_eq!(p.runs.len(), 1, "应有 1 个 run");
        // 注意：trim_text(true) 会去掉尾部空格
        assert_eq!(p.runs[0].text, "Page");
        assert_eq!(p.fields.len(), 2, "应有 2 个字段");
        assert_eq!(p.fields[0].field_type, FieldType::SlideNumber);
        assert_eq!(p.fields[0].text, "1");
        assert_eq!(p.fields[1].field_type, FieldType::DateTime);
        assert_eq!(p.fields[1].text, "1/1/2024");
    }

    /// 验证自定义字段类型的解析。
    ///
    /// 这是 TODO-016 的测试。
    #[test]
    fn parse_paragraph_custom_field_type() {
        let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:fld id="{CCC}" type="customField"><a:t>custom</a:t></a:fld>
</a:p>"#;
        let p = parse_paragraph(xml).expect("parse ok");
        assert_eq!(p.fields.len(), 1);
        assert_eq!(
            p.fields[0].field_type,
            FieldType::Custom("customField".to_string())
        );
        assert_eq!(p.fields[0].text, "custom");
    }

    /// 验证 `<a:hlinkClick>` 自闭合超链接的解析（r:id + tooltip）。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn parse_run_properties_hlink_click_self_closing() {
        let xml = r#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" lang="en-US">
  <a:hlinkClick r:id="rId3" tooltip="点击访问"/>
</a:rPr>"#;
        let rp = parse_run_properties(xml).expect("parse ok");
        let hl = rp.hlink_click.expect("hlink_click 应存在");
        assert_eq!(hl.rid.as_deref(), Some("rId3"));
        assert_eq!(hl.tooltip.as_deref(), Some("点击访问"));
        assert!(hl.action.is_none());
        assert!(!hl.invalid);
    }

    /// 验证 `<a:hlinkClick>` 带动作（跳转幻灯片）的解析。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn parse_run_properties_hlink_click_action() {
        let xml = r#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:hlinkClick action="ppaction://hlinksldjump"/>
</a:rPr>"#;
        let rp = parse_run_properties(xml).expect("parse ok");
        let hl = rp.hlink_click.expect("hlink_click 应存在");
        assert_eq!(hl.action.as_deref(), Some("ppaction://hlinksldjump"));
        assert!(hl.rid.is_none());
    }

    /// 验证 `<a:hlinkHover>` 的解析。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn parse_run_properties_hlink_hover() {
        let xml = r#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:hlinkHover r:id="rId5" tooltip="悬停提示"/>
</a:rPr>"#;
        let rp = parse_run_properties(xml).expect("parse ok");
        let hl = rp.hlink_hover.expect("hlink_hover 应存在");
        assert_eq!(hl.rid.as_deref(), Some("rId5"));
        assert_eq!(hl.tooltip.as_deref(), Some("悬停提示"));
    }

    /// 验证同时存在 hlinkClick 和 hlinkHover 的解析。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn parse_run_properties_both_hlinks() {
        let xml = r#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:hlinkClick r:id="rId1" tooltip="click"/>
  <a:hlinkHover r:id="rId2" tooltip="hover"/>
</a:rPr>"#;
        let rp = parse_run_properties(xml).expect("parse ok");
        assert_eq!(
            rp.hlink_click.as_ref().unwrap().rid.as_deref(),
            Some("rId1")
        );
        assert_eq!(
            rp.hlink_hover.as_ref().unwrap().rid.as_deref(),
            Some("rId2")
        );
    }

    /// 验证空属性 `<a:hlinkClick/>` 标记为 invalid（继承场景）。
    ///
    /// 这是 TODO-026 的测试。
    #[test]
    fn parse_run_properties_hlink_click_empty() {
        let xml = r#"<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:hlinkClick/>
</a:rPr>"#;
        let rp = parse_run_properties(xml).expect("parse ok");
        let hl = rp.hlink_click.expect("hlink_click 应存在");
        assert!(hl.invalid);
        assert!(hl.rid.is_none());
    }

    /// 验证 `<a:spLocks>` 自闭合形状锁定的解析。
    ///
    /// 这是 TODO-027 的测试。
    #[test]
    fn parse_sp_with_locks() {
        let xml = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="2" name="Locked"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1" noSelect="1" noResize="1"/>
    </p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody><a:bodyPr/><a:p/></p:txBody>
</p:sp>"#;
        let sp = parse_sp(xml).expect("parse ok");
        let locks = sp.locks.as_ref().expect("locks 应存在");
        assert!(locks.no_grp);
        assert!(locks.no_select);
        assert!(locks.no_resize);
        assert!(!locks.no_move);
        assert!(!locks.no_rot);
    }

    /// 验证 `<a:spLocks>` 与 `txBox` 同时存在的解析。
    ///
    /// 这是 TODO-027 的测试。
    #[test]
    fn parse_sp_locks_with_txbox() {
        let xml = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="3" name="TB"/>
    <p:cNvSpPr txBox="1">
      <a:spLocks noChangeAspect="1"/>
    </p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody><a:bodyPr/><a:p/></p:txBody>
</p:sp>"#;
        let sp = parse_sp(xml).expect("parse ok");
        assert!(sp.c_nv_sp_pr_tx_box, "txBox 应为 true");
        let locks = sp.locks.as_ref().expect("locks 应存在");
        assert!(locks.no_change_aspect);
        assert!(!locks.no_grp);
    }

    /// 验证 `<p:style>` 主题样式引用的解析。
    ///
    /// 这是 TODO-006 的测试。
    #[test]
    fn parse_sp_with_style() {
        let xml = r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="4" name="Styled"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr/>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent3"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>
  </p:style>
  <p:txBody><a:bodyPr/><a:p/></p:txBody>
</p:sp>"#;
        let sp = parse_sp(xml).expect("parse ok");
        let style = sp.style.as_ref().expect("style 应存在");
        let ln = style.line_ref.as_ref().expect("line_ref 应存在");
        assert_eq!(ln.idx.as_deref(), Some("1"));
        assert_eq!(ln.scheme_color.as_deref(), Some("accent1"));
        let fill = style.fill_ref.as_ref().expect("fill_ref 应存在");
        assert_eq!(fill.idx.as_deref(), Some("2"));
        assert_eq!(fill.scheme_color.as_deref(), Some("accent2"));
        let eff = style.effect_ref.as_ref().expect("effect_ref 应存在");
        assert_eq!(eff.idx.as_deref(), Some("0"));
        let font = style.font_ref.as_ref().expect("font_ref 应存在");
        assert_eq!(font.idx.as_deref(), Some("minor"));
        assert_eq!(font.scheme_color.as_deref(), Some("tx1"));
    }

    /// 验证 `<p:style>` 自闭合子元素的解析。
    ///
    /// 这是 TODO-006 的测试。
    #[test]
    fn parse_sp_style_self_closing_refs() {
        let xml = r#"<p:style xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:lnRef idx="3"/>
  <a:fillRef idx="1"/>
</p:style>"#;
        let style = parse_shape_style(xml).expect("parse ok");
        assert_eq!(style.line_ref.as_ref().unwrap().idx.as_deref(), Some("3"));
        assert!(style.line_ref.as_ref().unwrap().scheme_color.is_none());
        assert_eq!(style.fill_ref.as_ref().unwrap().idx.as_deref(), Some("1"));
        assert!(style.effect_ref.is_none());
        assert!(style.font_ref.is_none());
    }

    // ===== TODO-020：幻灯片过渡测试 =====

    /// 验证 `<p:transition>` 的 fade 类型解析。
    #[test]
    fn parse_transition_fade() {
        let xml = r#"<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow" advClick="0" advTm="5000">
  <p:fade thruBlk="1"/>
</p:transition>"#;
        let tr = parse_transition(xml).expect("parse ok");
        assert_eq!(tr.speed, TransitionSpeed::Slow);
        assert!(!tr.advance_click);
        assert_eq!(tr.advance_after_ms, Some(5000));
        match &tr.transition_type {
            TransitionType::Fade { thru_blk } => assert!(*thru_blk),
            other => panic!("expected Fade, got {:?}", other),
        }
    }

    /// 验证 `<p:transition>` 的 push 类型解析（带方向）。
    #[test]
    fn parse_transition_push() {
        let xml = r#"<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="fast">
  <p:push dir="l"/>
</p:transition>"#;
        let tr = parse_transition(xml).expect("parse ok");
        assert_eq!(tr.speed, TransitionSpeed::Fast);
        // 默认 advClick=true
        assert!(tr.advance_click);
        assert!(tr.advance_after_ms.is_none());
        match &tr.transition_type {
            TransitionType::Push { dir } => assert_eq!(*dir, TransitionDirection::Left),
            other => panic!("expected Push, got {:?}", other),
        }
    }

    /// 验证 `<p:transition>` 的 morph 类型解析。
    #[test]
    fn parse_transition_morph() {
        let xml = r#"<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:morph option="byChar"/>
</p:transition>"#;
        let tr = parse_transition(xml).expect("parse ok");
        match &tr.transition_type {
            TransitionType::Morph { option } => assert_eq!(*option, MorphOption::ByChar),
            other => panic!("expected Morph, got {:?}", other),
        }
    }

    /// 验证 `<p:transition>` 的 split 类型解析（带方向和方向）。
    #[test]
    fn parse_transition_split() {
        let xml = r#"<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:split orient="vert" dir="d"/>
</p:transition>"#;
        let tr = parse_transition(xml).expect("parse ok");
        match &tr.transition_type {
            TransitionType::Split { orient, dir } => {
                assert_eq!(*orient, SplitOrientation::Vertical);
                assert_eq!(*dir, TransitionDirection::Down);
            }
            other => panic!("expected Split, got {:?}", other),
        }
    }

    /// 验证 `<p:transition>` 自闭合（无子元素）解析。
    #[test]
    fn parse_transition_self_closing() {
        let xml = r#"<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med" advClick="1"/>"#;
        let tr = parse_transition(xml).expect("parse ok");
        assert_eq!(tr.speed, TransitionSpeed::Medium);
        assert!(tr.advance_click);
        // 无子元素时，transition_type 为 None
        assert_eq!(tr.transition_type, TransitionType::None);
    }

    /// 验证 `parse_sld` 能正确解析包含 `<p:transition>` 的 slide。
    #[test]
    fn parse_sld_with_transition() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
  <p:transition spd="slow" advTm="3000">
    <p:fade/>
  </p:transition>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let tr = sld.transition.expect("transition should exist");
        assert_eq!(tr.speed, TransitionSpeed::Slow);
        assert_eq!(tr.advance_after_ms, Some(3000));
        match &tr.transition_type {
            TransitionType::Fade { thru_blk } => assert!(!*thru_blk),
            other => panic!("expected Fade, got {:?}", other),
        }
    }

    /// 验证 `Transition::write_xml` 序列化 + `parse_transition` 反序列化的 round-trip。
    #[test]
    fn transition_round_trip() {
        use crate::oxml::slide::Transition;
        use crate::oxml::writer::XmlWriter;

        let original = Transition {
            speed: TransitionSpeed::Fast,
            advance_click: false,
            advance_after_ms: Some(8000),
            transition_type: TransitionType::Push {
                dir: TransitionDirection::Up,
            },
        };
        // 序列化
        let mut w = XmlWriter::new();
        original.write_xml(&mut w);
        let xml = w.buf.clone();
        // 反序列化
        let parsed = parse_transition(&xml).expect("parse ok");
        assert_eq!(parsed.speed, original.speed);
        assert_eq!(parsed.advance_click, original.advance_click);
        assert_eq!(parsed.advance_after_ms, original.advance_after_ms);
        assert_eq!(parsed.transition_type, original.transition_type);
    }

    // ===== TODO-024：自定义几何（custGeom）测试 =====

    /// 验证 `parse_custom_geometry` 能解析基本的 moveTo/lnTo/close 路径。
    #[test]
    fn parse_custom_geometry_basic() {
        let xml = r#"<a:custGeom>
<a:avLst/>
<a:rect l="0" t="0" r="100" b="100"/>
<a:pathLst>
<a:path w="100" h="100" fill="norm" stroke="norm">
<a:moveTo><a:pt x="0" y="0"/></a:moveTo>
<a:lnTo><a:pt x="100" y="0"/></a:lnTo>
<a:lnTo><a:pt x="50" y="100"/></a:lnTo>
<a:close/>
</a:path>
</a:pathLst>
</a:custGeom>"#;
        let geom = parse_custom_geometry(xml).expect("parse ok");
        assert_eq!(
            geom.rect,
            Some(GeomRect {
                l: "0".to_string(),
                t: "0".to_string(),
                r: "100".to_string(),
                b: "100".to_string(),
            })
        );
        assert_eq!(geom.path_list.len(), 1);
        let p = &geom.path_list[0];
        assert_eq!(p.width, 100);
        assert_eq!(p.height, 100);
        assert_eq!(p.fill.as_deref(), Some("norm"));
        assert_eq!(p.stroke.as_deref(), Some("norm"));
        assert_eq!(p.segments.len(), 4);
        // 验证各段
        match &p.segments[0] {
            PathSegment::MoveTo { x, y } => {
                assert_eq!(*x, 0);
                assert_eq!(*y, 0);
            }
            other => panic!("expected MoveTo, got {:?}", other),
        }
        match &p.segments[1] {
            PathSegment::LineTo { x, y } => {
                assert_eq!(*x, 100);
                assert_eq!(*y, 0);
            }
            other => panic!("expected LineTo, got {:?}", other),
        }
        match &p.segments[2] {
            PathSegment::LineTo { x, y } => {
                assert_eq!(*x, 50);
                assert_eq!(*y, 100);
            }
            other => panic!("expected LineTo, got {:?}", other),
        }
        match &p.segments[3] {
            PathSegment::Close => {}
            other => panic!("expected Close, got {:?}", other),
        }
    }

    /// 验证 `parse_custom_geometry` 能解析贝塞尔曲线段。
    #[test]
    fn parse_custom_geometry_with_bez() {
        let xml = r#"<a:custGeom>
<a:avLst/>
<a:pathLst>
<a:path w="200" h="200">
<a:moveTo><a:pt x="10" y="10"/></a:moveTo>
<a:cubicBezTo>
<a:pt x="50" y="0"/>
<a:pt x="150" y="0"/>
<a:pt x="190" y="10"/>
</a:cubicBezTo>
<a:quadBezTo>
<a:pt x="190" y="100"/>
<a:pt x="100" y="190"/>
</a:quadBezTo>
</a:path>
</a:pathLst>
</a:custGeom>"#;
        let geom = parse_custom_geometry(xml).expect("parse ok");
        assert_eq!(geom.path_list.len(), 1);
        let p = &geom.path_list[0];
        assert_eq!(p.segments.len(), 3);
        match &p.segments[1] {
            PathSegment::CubicBezTo {
                x1,
                y1,
                x2,
                y2,
                x3,
                y3,
            } => {
                assert_eq!(*x1, 50);
                assert_eq!(*y1, 0);
                assert_eq!(*x2, 150);
                assert_eq!(*y2, 0);
                assert_eq!(*x3, 190);
                assert_eq!(*y3, 10);
            }
            other => panic!("expected CubicBezTo, got {:?}", other),
        }
        match &p.segments[2] {
            PathSegment::QuadBezTo { x1, y1, x2, y2 } => {
                assert_eq!(*x1, 190);
                assert_eq!(*y1, 100);
                assert_eq!(*x2, 100);
                assert_eq!(*y2, 190);
            }
            other => panic!("expected QuadBezTo, got {:?}", other),
        }
    }

    /// 验证 `parse_custom_geometry` 能解析 arcTo 段。
    #[test]
    fn parse_custom_geometry_with_arc() {
        let xml = r#"<a:custGeom>
<a:avLst/>
<a:pathLst>
<a:path w="100" h="100">
<a:moveTo><a:pt x="0" y="50"/></a:moveTo>
<a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
<a:close/>
</a:path>
</a:pathLst>
</a:custGeom>"#;
        let geom = parse_custom_geometry(xml).expect("parse ok");
        let p = &geom.path_list[0];
        assert_eq!(p.segments.len(), 3);
        match &p.segments[1] {
            PathSegment::ArcTo {
                w_r,
                h_r,
                st_ang,
                sw_ang,
            } => {
                assert_eq!(*w_r, 50);
                assert_eq!(*h_r, 50);
                assert_eq!(*st_ang, 0);
                assert_eq!(*sw_ang, 5400000);
            }
            other => panic!("expected ArcTo, got {:?}", other),
        }
    }

    /// 验证 `parse_custom_geometry` 能解析 fill/stroke 文本元素。
    #[test]
    fn parse_custom_geometry_fill_stroke() {
        let xml = r#"<a:custGeom>
<a:avLst/>
<a:fill>norm</a:fill>
<a:stroke>none</a:stroke>
<a:pathLst>
<a:path w="10" h="10">
<a:moveTo><a:pt x="0" y="0"/></a:moveTo>
</a:path>
</a:pathLst>
</a:custGeom>"#;
        let geom = parse_custom_geometry(xml).expect("parse ok");
        assert_eq!(geom.fill.as_deref(), Some("norm"));
        assert_eq!(geom.stroke.as_deref(), Some("none"));
    }

    /// 验证 `CustomGeometry::write_xml` + `parse_custom_geometry` 的 round-trip。
    #[test]
    fn custom_geometry_round_trip() {
        use crate::oxml::writer::XmlWriter;

        let original = CustomGeometry {
            fill: Some("norm".to_string()),
            stroke: None,
            rect: Some(GeomRect {
                l: "0".to_string(),
                t: "0".to_string(),
                r: "100".to_string(),
                b: "100".to_string(),
            }),
            path_list: vec![Path {
                width: 100,
                height: 100,
                fill: Some("norm".to_string()),
                stroke: None,
                segments: vec![
                    PathSegment::MoveTo { x: 0, y: 0 },
                    PathSegment::LineTo { x: 100, y: 0 },
                    PathSegment::LineTo { x: 100, y: 100 },
                    PathSegment::LineTo { x: 0, y: 100 },
                    PathSegment::Close,
                ],
            }],
        };
        // 序列化
        let mut w = XmlWriter::new();
        original.write_xml(&mut w);
        let xml = w.buf.clone();
        // 反序列化
        let parsed = parse_custom_geometry(&xml).expect("parse ok");
        assert_eq!(parsed.fill, original.fill);
        assert_eq!(parsed.stroke, original.stroke);
        assert_eq!(parsed.rect, original.rect);
        assert_eq!(parsed.path_list.len(), 1);
        let p0 = &parsed.path_list[0];
        assert_eq!(p0.width, 100);
        assert_eq!(p0.height, 100);
        assert_eq!(p0.fill.as_deref(), Some("norm"));
        assert_eq!(p0.segments.len(), 5);
        // 验证第一段
        match &p0.segments[0] {
            PathSegment::MoveTo { x, y } => {
                assert_eq!(*x, 0);
                assert_eq!(*y, 0);
            }
            other => panic!("expected MoveTo, got {:?}", other),
        }
        // 验证最后一段
        match &p0.segments[4] {
            PathSegment::Close => {}
            other => panic!("expected Close, got {:?}", other),
        }
    }

    /// 验证 `parse_sld` 能从完整 slide XML 中解析带 custGeom 的形状。
    #[test]
    fn parse_sp_with_custgeom() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="10" name="Freeform 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="100" y="100"/>
            <a:ext cx="200" cy="200"/>
          </a:xfrm>
          <a:custGeom>
            <a:avLst/>
            <a:pathLst>
              <a:path w="200" h="200">
                <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
                <a:lnTo><a:pt x="200" y="0"/></a:lnTo>
                <a:lnTo><a:pt x="100" y="200"/></a:lnTo>
                <a:close/>
              </a:path>
            </a:pathLst>
          </a:custGeom>
        </p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr/>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        assert_eq!(sld.shapes.len(), 1);
        match &sld.shapes[0] {
            OxmlSlideShape::Sp(sp) => {
                // 验证几何是 Custom
                let geom = sp
                    .properties
                    .geometry
                    .as_ref()
                    .expect("geometry should exist");
                match geom {
                    Geometry::Custom(cg) => {
                        assert_eq!(cg.path_list.len(), 1);
                        let p = &cg.path_list[0];
                        assert_eq!(p.width, 200);
                        assert_eq!(p.height, 200);
                        assert_eq!(p.segments.len(), 4);
                        match &p.segments[0] {
                            PathSegment::MoveTo { x, y } => {
                                assert_eq!(*x, 0);
                                assert_eq!(*y, 0);
                            }
                            other => panic!("expected MoveTo, got {:?}", other),
                        }
                    }
                    other => panic!("expected Custom geometry, got {:?}", other),
                }
            }
            other => panic!("expected Sp, got {:?}", other),
        }
    }

    /// 验证 `Geometry::write_xml` 对 Preset 变体的序列化（确保 prstGeom 路径未被破坏）。
    #[test]
    fn geometry_preset_write_xml() {
        use crate::oxml::simpletypes::PresetGeometry;
        use crate::oxml::writer::XmlWriter;

        let geom = Geometry::preset(PresetGeometry::Rectangle);
        let mut w = XmlWriter::new();
        geom.write_xml(&mut w);
        let xml = w.buf.clone();
        assert!(xml.contains("<a:prstGeom prst=\"rect\">"), "xml: {}", xml);
        assert!(xml.contains("<a:avLst/>"), "xml: {}", xml);
        assert!(xml.contains("</a:prstGeom>"), "xml: {}", xml);
    }

    /// 验证 `Geometry::write_xml` 对 Custom 变体的序列化。
    #[test]
    fn geometry_custom_write_xml() {
        use crate::oxml::writer::XmlWriter;

        let geom = Geometry::Custom(CustomGeometry {
            fill: None,
            stroke: None,
            rect: None,
            path_list: vec![Path {
                width: 50,
                height: 50,
                fill: None,
                stroke: None,
                segments: vec![
                    PathSegment::MoveTo { x: 0, y: 0 },
                    PathSegment::LineTo { x: 50, y: 50 },
                    PathSegment::Close,
                ],
            }],
        });
        let mut w = XmlWriter::new();
        geom.write_xml(&mut w);
        let xml = w.buf.clone();
        assert!(xml.contains("<a:custGeom>"), "xml: {}", xml);
        assert!(xml.contains("<a:avLst/>"), "xml: {}", xml);
        assert!(xml.contains("<a:pathLst>"), "xml: {}", xml);
        assert!(xml.contains("<a:path w=\"50\" h=\"50\">"), "xml: {}", xml);
        assert!(xml.contains("<a:moveTo>"), "xml: {}", xml);
        assert!(xml.contains("<a:lnTo>"), "xml: {}", xml);
        assert!(xml.contains("<a:close/>"), "xml: {}", xml);
    }

    /// 验证带调整值的 `<a:prstGeom>` 在 XML 往返中保持不变（TODO-038）。
    ///
    /// 测试场景：圆角矩形（roundRect）带单个调整值 `adj=16667`（16.667%）。
    /// 期望：解析后 `Geometry::Preset` 携带一个 `AdjustmentValue`，
    /// 序列化后 XML 仍包含 `<a:gd name="adj" fmla="val 16667"/>`。
    #[test]
    fn round_trip_prst_geom_with_adjustments() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="RoundRect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="100000" y="100000"/><a:ext cx="500000" cy="300000"/></a:xfrm>
        <a:prstGeom prst="roundRect">
          <a:avLst>
            <a:gd name="adj" fmla="val 16667"/>
          </a:avLst>
        </a:prstGeom>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let sp = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::Sp(sp) = s {
                    Some(sp)
                } else {
                    None
                }
            })
            .expect("应有 Sp");
        // 验证几何为 Preset(roundRect) 且携带一个调整值
        match &sp.properties.geometry {
            Some(Geometry::Preset(prst, adjustments)) => {
                assert_eq!(*prst, PresetGeometry::RoundRectangle);
                assert_eq!(adjustments.len(), 1, "应有 1 个调整值");
                let adj = &adjustments[0];
                assert_eq!(adj.name, "adj");
                assert_eq!(adj.raw_value, 16667);
                assert!((adj.effective_value() - 0.16667).abs() < 1e-6);
            }
            other => panic!("期望 Preset 几何，得到 {:?}", other),
        }
        // 验证序列化后 XML 仍包含调整值
        let mut w = crate::oxml::writer::XmlWriter::new();
        sp.properties.geometry.as_ref().unwrap().write_xml(&mut w);
        let out = &w.buf;
        assert!(
            out.contains("<a:prstGeom prst=\"roundRect\">"),
            "xml: {}",
            out
        );
        assert!(out.contains("<a:avLst>"), "xml: {}", out);
        assert!(
            out.contains("<a:gd name=\"adj\" fmla=\"val 16667\"/>"),
            "xml: {}",
            out
        );
        assert!(out.contains("</a:prstGeom>"), "xml: {}", out);
    }

    /// 验证带多个调整值的 `<a:prstGeom>` 在 XML 往返中保持不变（TODO-038）。
    ///
    /// 测试场景：`rect` 形状带两个调整值 `adj1=50000` 和 `adj2=25000`。
    /// 注意：rect 通常无调整值，但本测试验证解析器能正确处理多个 `<a:gd>`。
    #[test]
    fn round_trip_prst_geom_with_multiple_adjustments() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="MultiAdj"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst>
            <a:gd name="adj1" fmla="val 50000"/>
            <a:gd name="adj2" fmla="val 25000"/>
          </a:avLst>
        </a:prstGeom>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let sp = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::Sp(sp) = s {
                    Some(sp)
                } else {
                    None
                }
            })
            .expect("应有 Sp");
        match &sp.properties.geometry {
            Some(Geometry::Preset(_, adjustments)) => {
                assert_eq!(adjustments.len(), 2, "应有 2 个调整值");
                assert_eq!(adjustments[0].name, "adj1");
                assert_eq!(adjustments[0].raw_value, 50000);
                assert_eq!(adjustments[1].name, "adj2");
                assert_eq!(adjustments[1].raw_value, 25000);
            }
            other => panic!("期望 Preset 几何，得到 {:?}", other),
        }
        // 验证序列化
        let mut w = crate::oxml::writer::XmlWriter::new();
        sp.properties.geometry.as_ref().unwrap().write_xml(&mut w);
        let out = &w.buf;
        assert!(
            out.contains("<a:gd name=\"adj1\" fmla=\"val 50000\"/>"),
            "xml: {}",
            out
        );
        assert!(
            out.contains("<a:gd name=\"adj2\" fmla=\"val 25000\"/>"),
            "xml: {}",
            out
        );
    }

    /// 验证空 `<a:avLst/>` 解析为空调整值列表（TODO-038）。
    #[test]
    fn parse_prst_geom_with_empty_av_lst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="4" name="EmptyAdj"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let sp = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::Sp(sp) = s {
                    Some(sp)
                } else {
                    None
                }
            })
            .expect("应有 Sp");
        match &sp.properties.geometry {
            Some(Geometry::Preset(prst, adjustments)) => {
                assert_eq!(*prst, PresetGeometry::Rectangle);
                assert!(adjustments.is_empty(), "空 avLst 应解析为空列表");
            }
            other => panic!("期望 Preset 几何，得到 {:?}", other),
        }
    }

    /// 验证无 `<a:avLst>` 的 `<a:prstGeom>` 解析为空调整值列表（TODO-038）。
    #[test]
    fn parse_prst_geom_without_av_lst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm/></p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="5" name="NoAdj"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>
        <a:prstGeom prst="rect"/>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>"#;
        let sld = parse_sld(xml).expect("parse ok");
        let sp = sld
            .shapes
            .iter()
            .find_map(|s| {
                if let OxmlSlideShape::Sp(sp) = s {
                    Some(sp)
                } else {
                    None
                }
            })
            .expect("应有 Sp");
        match &sp.properties.geometry {
            Some(Geometry::Preset(prst, adjustments)) => {
                assert_eq!(*prst, PresetGeometry::Rectangle);
                assert!(adjustments.is_empty(), "无 avLst 应解析为空列表");
            }
            other => panic!("期望 Preset 几何，得到 {:?}", other),
        }
    }

    /// 验证 `parse_fmla_val` 对各种公式格式的解析（TODO-038）。
    #[test]
    fn parse_fmla_val_formats() {
        // 标准格式 "val 16667"
        assert_eq!(parse_fmla_val("val 16667"), Some(16667));
        // 无空格 "val16667"（罕见但合法）
        assert_eq!(parse_fmla_val("val16667"), Some(16667));
        // 带前后空白
        assert_eq!(parse_fmla_val("  val 16667  "), Some(16667));
        // 负数
        assert_eq!(parse_fmla_val("val -5000"), Some(-5000));
        // 非数字
        assert_eq!(parse_fmla_val("val abc"), None);
        // 非 val 公式（如乘法公式）
        assert_eq!(parse_fmla_val("*/ adj1 100000 50000"), None);
        // 空字符串
        assert_eq!(parse_fmla_val(""), None);
    }

    /// 验证 `parse_graphic_frame` 能识别 SmartArt 并保留 raw_xml（TODO-037 最小保留）。
    ///
    /// 关键断言：
    /// - `<a:graphicData uri=".../diagram">` 应被识别为 `Graphic::SmartArt`；
    /// - `raw_xml` 应包含完整的 `<a:graphicData>` 元素（含外壳）；
    /// - 4 个关系 id（r:dm / r:lo / r:qs / r:cs）应被正确提取。
    #[test]
    fn parse_graphic_frame_with_smartart() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <p:nvGraphicFramePr>
    <p:cNvPr id="300" name="SmartArt 1"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="457200" y="1828800"/>
    <a:ext cx="8229600" cy="1143000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
      <dgm:relIds r:dm="rIdDm1" r:lo="rIdLo1" r:qs="rIdQs1" r:cs="rIdCs1"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#;
        let frame = parse_graphic_frame(xml).expect("parse ok");
        assert_eq!(frame.id, 300);
        assert_eq!(frame.name, "SmartArt 1");
        match &frame.graphic {
            crate::oxml::shape::Graphic::SmartArt(s) => {
                // raw_xml 应包含完整 graphicData 元素
                assert!(
                    s.raw_xml.contains("<a:graphicData"),
                    "raw_xml 应包含 graphicData 外壳，实际: {}",
                    s.raw_xml
                );
                assert!(
                    s.raw_xml.contains(
                        "uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\""
                    ),
                    "raw_xml 应包含 diagram uri，实际: {}",
                    s.raw_xml
                );
                // 4 个关系 id 应被正确提取
                assert_eq!(s.dm_rid.as_deref(), Some("rIdDm1"));
                assert_eq!(s.lo_rid.as_deref(), Some("rIdLo1"));
                assert_eq!(s.qs_rid.as_deref(), Some("rIdQs1"));
                assert_eq!(s.cs_rid.as_deref(), Some("rIdCs1"));
            }
            other => panic!("期望 SmartArt，得到 {:?}", other),
        }
    }

    /// 验证 `parse_smartart_rel_ids` 对各种 raw_xml 格式的解析（TODO-037）。
    #[test]
    fn parse_smartart_rel_ids_formats() {
        // 标准格式：4 个关系 id 都存在
        let raw = r#"<a:graphicData uri=".../diagram"><dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/></a:graphicData>"#;
        let (dm, lo, qs, cs) = parse_smartart_rel_ids(raw);
        assert_eq!(dm.as_deref(), Some("rId1"));
        assert_eq!(lo.as_deref(), Some("rId2"));
        assert_eq!(qs.as_deref(), Some("rId3"));
        assert_eq!(cs.as_deref(), Some("rId4"));

        // 缺失部分属性
        let raw_partial =
            r#"<a:graphicData uri=".../diagram"><dgm:relIds r:dm="rId1"/></a:graphicData>"#;
        let (dm, lo, qs, cs) = parse_smartart_rel_ids(raw_partial);
        assert_eq!(dm.as_deref(), Some("rId1"));
        assert!(lo.is_none(), "lo 应为 None");
        assert!(qs.is_none(), "qs 应为 None");
        assert!(cs.is_none(), "cs 应为 None");

        // 完全缺失 dgm:relIds
        let raw_empty = r#"<a:graphicData uri=".../diagram"></a:graphicData>"#;
        let (dm, lo, qs, cs) = parse_smartart_rel_ids(raw_empty);
        assert!(dm.is_none() && lo.is_none() && qs.is_none() && cs.is_none());
    }
}
