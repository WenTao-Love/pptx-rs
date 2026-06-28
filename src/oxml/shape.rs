//! 形状 XML 模型：`p:sp` / `p:pic` / `p:grpSp` / `p:cxnSp` / `p:graphicFrame`。
//!
//! 包含 OOXML 中所有"可视对象"的强类型表达。每个类型都自带 `write_xml` 方法，
//! 按规范要求的元素顺序输出。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.shapetree.*` ←→ [`Sp`] / [`Pic`] / [`Group`] / [`Connector`]；
//! - `pptx.oxml.shape.*` ←→ 同名 oxml 模型；
//! - `pptx.shapes.graphfrm.GraphicFrame` ←→ [`GraphicFrame`]（承载表格/图表）。
//!
//! # 序列化约定
//!
//! - **非视觉属性隐藏**（`<p:nvSpPr>` 中的 cNvPr）必须最先输出；
//! - **spPr**（几何/变换/填充/线）紧跟其后；
//! - **txBody**（文本）最后输出；
//! - **顺序错误**会导致 PowerPoint 报错 "Invalid OOXML"。

use crate::oxml::sppr::ShapeProperties;
use crate::oxml::txbody::TextBody;
use crate::units::Emu;

/// 形状锁定属性（`<a:spLocks>`）。
///
/// 对应 OOXML 中 `<p:cNvSpPr>` 内的 `<a:spLocks>` 元素。
/// 所有属性均为布尔值，`true` 表示禁止对应操作。
///
/// 对标 python-pptx 中 `shape.locks` 的底层支撑（python-pptx 尚未暴露高阶 API，
/// 但读写时保留这些属性）。
#[derive(Clone, Debug, Default)]
pub struct ShapeLocks {
    /// 禁止组合（`noGrp="1"`）。
    pub no_grp: bool,
    /// 禁止进入组合（`noDrilldown="1"`）。
    pub no_drilldown: bool,
    /// 禁止选择（`noSelect="1"`）。
    pub no_select: bool,
    /// 禁止改变宽高比（`noChangeAspect="1"`）。
    pub no_change_aspect: bool,
    /// 禁止移动（`noMove="1"`）。
    pub no_move: bool,
    /// 禁止缩放（`noResize="1"`）。
    pub no_resize: bool,
    /// 禁止旋转（`noRot="1"`）。
    pub no_rot: bool,
    /// 禁止编辑顶点（`noEditPoints="1"`）。
    pub no_edit_points: bool,
    /// 禁止调整手柄（`noAdjustHandles="1"`）。
    pub no_adjust_handles: bool,
    /// 禁止修改箭头（`noChangeArrowheads="1"`）。
    pub no_change_arrowheads: bool,
    /// 禁止修改形状类型（`noChangeShapeType="1"`）。
    pub no_change_shape_type: bool,
    /// 禁止裁剪（`noCrop="1"`）。
    pub no_crop: bool,
}

impl ShapeLocks {
    /// 是否所有属性都为 false（即无任何锁定）。
    pub fn is_empty(&self) -> bool {
        !self.no_grp
            && !self.no_drilldown
            && !self.no_select
            && !self.no_change_aspect
            && !self.no_move
            && !self.no_resize
            && !self.no_rot
            && !self.no_edit_points
            && !self.no_adjust_handles
            && !self.no_change_arrowheads
            && !self.no_change_shape_type
            && !self.no_crop
    }

    /// 写出 `<a:spLocks .../>` 自闭合标签。
    ///
    /// 若所有属性均为 false，则**不**写出（调用方应先调用 [`Self::is_empty`] 判断）。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        let mut attrs: Vec<(&str, &str)> = Vec::new();
        if self.no_grp {
            attrs.push(("noGrp", "1"));
        }
        if self.no_drilldown {
            attrs.push(("noDrilldown", "1"));
        }
        if self.no_select {
            attrs.push(("noSelect", "1"));
        }
        if self.no_change_aspect {
            attrs.push(("noChangeAspect", "1"));
        }
        if self.no_move {
            attrs.push(("noMove", "1"));
        }
        if self.no_resize {
            attrs.push(("noResize", "1"));
        }
        if self.no_rot {
            attrs.push(("noRot", "1"));
        }
        if self.no_edit_points {
            attrs.push(("noEditPoints", "1"));
        }
        if self.no_adjust_handles {
            attrs.push(("noAdjustHandles", "1"));
        }
        if self.no_change_arrowheads {
            attrs.push(("noChangeArrowheads", "1"));
        }
        if self.no_change_shape_type {
            attrs.push(("noChangeShapeType", "1"));
        }
        if self.no_crop {
            attrs.push(("noCrop", "1"));
        }
        w.empty_with("a:spLocks", &attrs);
    }

    /// 统一设置指定类型的锁定（TODO-027 高阶 API）。
    ///
    /// 对标 python-pptx 风格 `shape.set_lock(LockType::Select, true)`。
    /// 此方法是所有 `no_*` 字段写入器的统一入口，避免调用方记忆 12 个字段名。
    ///
    /// # 参数
    /// - `lock_type`：锁定类型（见 [`LockType`] 枚举）；
    /// - `locked`：`true` 启用该锁定，`false` 解除该锁定。
    ///
    /// # 示例
    ///
    /// ```no_run
    /// # use pptx::oxml::shape::{ShapeLocks, LockType};
    /// let mut l = ShapeLocks::default();
    /// l.set_lock(LockType::Select, true);
    /// l.set_lock(LockType::Rotate, true);
    /// assert!(l.no_select);
    /// assert!(l.no_rot);
    /// ```
    pub fn set_lock(&mut self, lock_type: LockType, locked: bool) {
        match lock_type {
            LockType::Grouping => self.no_grp = locked,
            LockType::Drilldown => self.no_drilldown = locked,
            LockType::Select => self.no_select = locked,
            LockType::ChangeAspect => self.no_change_aspect = locked,
            LockType::Move => self.no_move = locked,
            LockType::Resize => self.no_resize = locked,
            LockType::Rotate => self.no_rot = locked,
            LockType::EditPoints => self.no_edit_points = locked,
            LockType::AdjustHandles => self.no_adjust_handles = locked,
            LockType::ChangeArrowheads => self.no_change_arrowheads = locked,
            LockType::ChangeShapeType => self.no_change_shape_type = locked,
            LockType::Crop => self.no_crop = locked,
        }
    }

    /// 读取指定类型的锁定状态。详见 [`ShapeLocks::set_lock`]。
    pub fn get_lock(&self, lock_type: LockType) -> bool {
        match lock_type {
            LockType::Grouping => self.no_grp,
            LockType::Drilldown => self.no_drilldown,
            LockType::Select => self.no_select,
            LockType::ChangeAspect => self.no_change_aspect,
            LockType::Move => self.no_move,
            LockType::Resize => self.no_resize,
            LockType::Rotate => self.no_rot,
            LockType::EditPoints => self.no_edit_points,
            LockType::AdjustHandles => self.no_adjust_handles,
            LockType::ChangeArrowheads => self.no_change_arrowheads,
            LockType::ChangeShapeType => self.no_change_shape_type,
            LockType::Crop => self.no_crop,
        }
    }
}

/// 形状锁定类型枚举（对应 `<a:spLocks>` 的 12 个属性）。
///
/// 用于 [`ShapeLocks::set_lock`] / [`ShapeLocks::get_lock`] 高阶 API，
/// 避免调用方直接操作 12 个 `no_*` 字段。对标 python-pptx 的
/// `MSO_SHAPE_LOCK_TYPE` 枚举。
#[derive(Copy, Clone, Debug, Eq, PartialEq, Hash)]
pub enum LockType {
    /// 禁止组合（`noGrp`）。
    Grouping,
    /// 禁止进入组合（`noDrilldown`）。
    Drilldown,
    /// 禁止选择（`noSelect`）。
    Select,
    /// 禁止改变宽高比（`noChangeAspect`）。
    ChangeAspect,
    /// 禁止移动（`noMove`）。
    Move,
    /// 禁止缩放（`noResize`）。
    Resize,
    /// 禁止旋转（`noRot`）。
    Rotate,
    /// 禁止编辑顶点（`noEditPoints`）。
    EditPoints,
    /// 禁止调整手柄（`noAdjustHandles`）。
    AdjustHandles,
    /// 禁止修改箭头（`noChangeArrowheads`）。
    ChangeArrowheads,
    /// 禁止修改形状类型（`noChangeShapeType`）。
    ChangeShapeType,
    /// 禁止裁剪（`noCrop`）。
    Crop,
}

/// `<p:sp>` 普通形状（如文本框、矩形、椭圆等）。
#[derive(Clone, Debug, Default)]
pub struct Sp {
    /// 形状 ID（同一 slide 内唯一）。
    pub id: u32,
    /// 形状名（仅显示用）。
    pub name: String,
    /// 是否为占位符。
    pub is_placeholder: bool,
    /// 占位符 idx（placeholder idx，仅当 `is_placeholder`）。
    pub ph_idx: Option<u32>,
    /// 占位符类型（`title` / `body` / ...）。
    ///
    /// 完整取值见 `PP_PLACEHOLDER_TYPE` 枚举：
    /// - `title` / `ctrTitle` / `subTitle`
    /// - `body` / `ftr` / `dt` / `sldNum` / `hdr`
    /// - `tbl` / `chart` / `pic` / `sldImg` / `media`
    /// - `obj` / `vertAlign` 等等
    ///
    /// 当 `is_placeholder=true` 但 `ph_type=None` 时，写 XML 时**默认**按 `body` 写出。
    pub ph_type: Option<String>,
    /// 形状属性（xfrm / prstGeom / 填充 / 边框）。
    pub properties: ShapeProperties,
    /// 主题样式引用（`p:style`，可选）。
    pub style: Option<ShapeStyle>,
    /// 文本体（`p:txBody`）。
    pub text: TextBody,
    /// 扩展列表（`p:extLst`），用于承载 PowerPoint 私有扩展。
    pub ext_lst: Option<ExtensionList>,
    /// 标记该 sp 为**纯文本框**（对应 `<p:cNvSpPr txBox="1"/>`）。
    ///
    /// python-pptx 中 `shapes.add_textbox(...)` 创建的就是带这个标志的 `p:sp`。
    /// 该标志让 PowerPoint 把它识别为"自由文本框"而非"自选图形 + 文本"，从而
    /// 走**非自动布局**（不套用母版占位符约束）。
    pub c_nv_sp_pr_tx_box: bool,
    /// 形状锁定属性（`<a:spLocks>`，可选）。
    ///
    /// `None` 表示不写出 `<a:spLocks>` 元素。
    /// `Some(ShapeLocks { .. })` 且 `is_empty() == false` 时写出。
    pub locks: Option<ShapeLocks>,
    // 非视觉属性（隐藏、跳转等）——目前不实现。
}

impl Sp {
    /// 写 XML。
    ///
    /// 按 OOXML `CT_Shape` 的子元素顺序：
    ///
    /// ```text
    /// <p:sp>
    ///   <p:nvSpPr>...</p:nvSpPr>
    ///   <p:spPr>...</p:spPr>            ← spPr 必填
    ///   <p:style>...</p:style>          ← 可选：主题样式引用
    ///   <p:txBody>...</p:txBody>        ← 可选：文本体
    ///   <p:extLst>...</p:extLst>        ← 可选：扩展
    /// </p:sp>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:sp");
        // nvSpPr
        w.open("p:nvSpPr");
        let id_s = self.id.to_string();
        w.empty_with(
            "p:cNvPr",
            &[("id", id_s.as_str()), ("name", self.name.as_str())],
        );
        // 提前取出 idx 字符串，扩展到 if 块外
        let idx_s = self.ph_idx.map(|v| v.to_string());
        if self.is_placeholder {
            let mut pattrs: Vec<(&str, &str)> = Vec::new();
            // ph_type 缺省回落到 `body`——这是 OOXML 规范的默认 placeholder 类型，
            // 也是 PowerPoint 在用户新增占位符时实际写入的值。
            // 显式 None 会让 PowerPoint 弹"无法识别的占位符"警告。
            let type_str: &str = self.ph_type.as_deref().unwrap_or("body");
            pattrs.push(("type", type_str));
            if let Some(s) = &idx_s {
                pattrs.push(("idx", s.as_str()));
            }
            w.empty_with("p:ph", &pattrs);
        }
        // 纯文本框：写 `txBox="1"` 让 PowerPoint 识别为 textbox（不应用母版占位符）。
        // 若有形状锁定（spLocks），则 cNvSpPr 需要写成 open-close 形式以容纳子元素。
        let has_locks = self.locks.as_ref().map(|l| !l.is_empty()).unwrap_or(false);
        if has_locks {
            // 有 spLocks 子元素：写成 open-close
            if self.c_nv_sp_pr_tx_box {
                w.open_with("p:cNvSpPr", &[("txBox", "1")]);
            } else {
                w.open("p:cNvSpPr");
            }
            self.locks.as_ref().unwrap().write_xml(w);
            w.close("p:cNvSpPr");
        } else if self.c_nv_sp_pr_tx_box {
            w.empty_with("p:cNvSpPr", &[("txBox", "1")]);
        } else {
            w.empty_with("p:cNvSpPr", &[]);
        }
        w.empty("p:nvPr");
        w.close("p:nvSpPr");
        // spPr
        self.properties.write_xml(w, "p:spPr");
        // p:style（暂未填充具体内容；保留位置供后续主题样式引用）
        if let Some(style) = &self.style {
            style.write_xml(w);
        }
        // txBody
        self.text.write_xml(w);
        // extLst
        write_extlst(w, self.ext_lst.as_ref());
        w.close("p:sp");
    }
}

/// `<p:pic>` 图片。
#[derive(Clone, Debug, Default)]
pub struct Pic {
    pub id: u32,
    pub name: String,
    pub rid: String, // r:embed 关系 id
    /// 是否为占位符（TODO-007）。
    ///
    /// `true` 时 `write_xml` 会在 `<p:nvPicPr>/<p:nvPr>` 内写出 `<p:ph .../>`。
    /// 图片占位符对应版式中的 `<p:ph type="pic"/>`，用户点击后弹出"插入图片"对话框。
    pub is_placeholder: bool,
    /// 占位符 idx（`<p:ph idx="..."/>`，仅当 `is_placeholder` 为 true 时有效）。
    pub ph_idx: Option<u32>,
    /// 占位符类型（`<p:ph type="..."/>`）。
    ///
    /// 图片占位符通常为 `"pic"`；`None` 时 `write_xml` 不写出 `type` 属性。
    pub ph_type: Option<String>,
    pub properties: ShapeProperties,
    /// 图片裁剪。
    pub src_rect: Option<(i32, i32, i32, i32)>, // l%, t%, r%, b%  (100000=100%)
    /// 填充模式（拉伸/平铺/无）。
    pub fill_mode: crate::oxml::sppr::BlipFillMode,
    /// 图片透明度（0-100000，30000 = 30% 不透明 / 70% 透明）。
    ///
    /// 对应 `<a:blip>` 内的 `<a:alphaModFix amt="..."/>` 元素。
    /// 常用于图片水印场景：`alpha = Some(30_000)` 表示 30% 不透明。
    pub alpha: Option<i32>,
    /// 主题样式引用（`p:style`，可选）。
    pub style: Option<ShapeStyle>,
    /// 扩展列表。
    pub ext_lst: Option<ExtensionList>,
    /// 媒体引用（TODO-033 音视频嵌入）。
    ///
    /// `None` 表示普通图片；`Some(MediaKind::Video { rid })` / `Some(MediaKind::Audio { rid })`
    /// 表示该 `<p:pic>` 实际是视频/音频形状，`write_xml` 会在 `<p:nvPr>` 内写出
    /// `<a:videoFile r:link="..."/>` 或 `<a:audioFile r:link="..."/>`。
    /// `rid` 指向媒体 part 的关系 id（与 `self.rid` 指向海报帧图片不同）。
    pub media: Option<MediaKind>,
}

/// 媒体类型（TODO-033 音视频嵌入）。
///
/// OOXML 中视频/音频形状实际上是带 `<a:videoFile>` / `<a:audioFile>` 的 `<p:pic>`，
/// 用海报帧图片（`<a:blip r:embed="..."/>`）作为视觉占位。
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MediaKind {
    /// 视频（`<a:videoFile r:link="..."/>`）。
    Video {
        /// 媒体 part 的关系 id（指向 `/ppt/media/mediaN.mp4`）。
        rid: String,
    },
    /// 音频（`<a:audioFile r:link="..."/>`）。
    Audio {
        /// 媒体 part 的关系 id（指向 `/ppt/media/mediaN.mp3`）。
        rid: String,
    },
}

impl Default for MediaKind {
    fn default() -> Self {
        // 默认视频空 rid（避免 Option 双层包装）
        MediaKind::Video { rid: String::new() }
    }
}

impl Pic {
    /// 写 XML。
    ///
    /// 按 OOXML `CT_Picture` 的子元素顺序：
    ///
    /// ```text
    /// <p:pic>
    ///   <p:nvPicPr>...</p:nvPicPr>
    ///   <p:blipFill>...</p:blipFill>
    ///   <p:spPr>...</p:spPr>
    ///   <p:style>...</p:style>          ← 可选
    ///   <p:extLst>...</p:extLst>        ← 可选
    /// </p:pic>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:pic");
        w.open("p:nvPicPr");
        let id_s = self.id.to_string();
        w.empty_with("p:cNvPr", &[("id", &id_s), ("name", &self.name)]);
        w.empty("p:cNvPicPr");
        // nvPr：若 is_placeholder，写出 <p:ph type="..." idx="..."/>（TODO-007）。
        // 若有 media（TODO-033），写出 <a:videoFile r:link="..."/> 或 <a:audioFile r:link="..."/>。
        // OOXML 顺序：<p:cNvPr> → <p:cNvPicPr> → <p:nvPr>（含 <p:ph> / <a:videoFile> 等）。
        if self.is_placeholder || self.media.is_some() {
            w.open_with("p:nvPr", &[("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS)]);
            if self.is_placeholder {
                // idx_s 提到 pattrs 之前，保证生命周期覆盖 empty_with 调用
                let idx_s = self.ph_idx.map(|i| i.to_string());
                let mut pattrs: Vec<(&str, &str)> = Vec::new();
                if let Some(t) = self.ph_type.as_deref() {
                    pattrs.push(("type", t));
                }
                if let Some(s) = idx_s.as_deref() {
                    pattrs.push(("idx", s));
                }
                w.empty_with("p:ph", &pattrs);
            }
            // 媒体引用（TODO-033）
            if let Some(media) = &self.media {
                match media {
                    MediaKind::Video { rid } => {
                        w.empty_with("a:videoFile", &[("r:link", rid.as_str())]);
                    }
                    MediaKind::Audio { rid } => {
                        w.empty_with("a:audioFile", &[("r:link", rid.as_str())]);
                    }
                }
            }
            w.close("p:nvPr");
        } else {
            w.empty("p:nvPr");
        }
        w.close("p:nvPicPr");
        // blipFill：开始标签 + 子元素 + 结束标签
        // 子元素是 `<a:blip r:embed="rIdN"/>`（自闭合）和可选的 `<a:stretch>` / `<a:srcRect>`。
        // `xmlns:r` 放在最外层 blipFill 上以声明 r: 前缀。
        w.open_with(
            "p:blipFill",
            &[("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS)],
        );
        // a:blip：若有 alpha 透明度，需要开标签 + 子元素 + 关标签
        if let Some(alpha_val) = self.alpha {
            w.open_with("a:blip", &[("r:embed", self.rid.as_str())]);
            let amt_s = alpha_val.to_string();
            w.empty_with("a:alphaModFix", &[("amt", amt_s.as_str())]);
            w.close("a:blip");
        } else {
            w.empty_with("a:blip", &[("r:embed", self.rid.as_str())]);
        }
        // srcRect 用于裁剪：l/t/r/b 均为 0..=100000 的千分比（100000=100%）。
        if let Some((l, t, r, b)) = self.src_rect {
            let l_s = l.to_string();
            let t_s = t.to_string();
            let r_s = r.to_string();
            let b_s = b.to_string();
            w.empty_with(
                "a:srcRect",
                &[("l", &l_s), ("t", &t_s), ("r", &r_s), ("b", &b_s)],
            );
        }
        // 写出填充模式（拉伸/平铺/无）
        self.fill_mode.write_xml(w);
        w.close("p:blipFill");
        // spPr
        self.properties.write_xml(w, "p:spPr");
        // p:style
        if let Some(style) = &self.style {
            style.write_xml(w);
        }
        // extLst
        write_extlst(w, self.ext_lst.as_ref());
        w.close("p:pic");
    }
}

/// `<p:grpSp>` 组合形状。
#[derive(Clone, Debug, Default)]
pub struct Group {
    pub id: u32,
    pub name: String,
    pub properties: ShapeProperties,
    /// 子形状（递归）。
    pub children: Vec<GroupChild>,
    /// 组合自己的大小（写在 a:xfrm 上）。
    pub ext: (Emu, Emu),
    pub off: (Emu, Emu),
    /// 主题样式引用（`p:style`，可选）。
    pub style: Option<ShapeStyle>,
    /// 扩展列表。
    pub ext_lst: Option<ExtensionList>,
}

#[derive(Clone, Debug)]
pub enum GroupChild {
    Sp(Sp),
    Pic(Pic),
    CxnSp(Connector),
    Group(Box<Group>),
    /// 图形框（承载表格/图表等）。
    GraphicFrame(GraphicFrame),
}

impl Group {
    /// 写 XML。
    ///
    /// 按 OOXML `CT_GroupShape` 的子元素顺序：
    ///
    /// ```text
    /// <p:grpSp>
    ///   <p:nvGrpSpPr>...</p:nvGrpSpPr>
    ///   <p:grpSpPr>...</p:grpSpPr>
    ///   子形状 (sp/pic/cxnSp/grpSp/graphicFrame/contentPart)
    ///   <p:extLst>...</p:extLst>        ← 可选
    /// </p:grpSp>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:grpSp");
        w.open("p:nvGrpSpPr");
        let id_s = self.id.to_string();
        w.empty_with("p:cNvPr", &[("id", &id_s), ("name", &self.name)]);
        w.empty("p:cNvGrpSpPr");
        w.empty("p:nvPr");
        w.close("p:nvGrpSpPr");
        // 组合本身 spPr 中有 chExt / chOff
        w.open("p:grpSpPr");
        w.open("a:xfrm");
        let off_x_s = self.off.0.value().to_string();
        let off_y_s = self.off.1.value().to_string();
        let ext_cx_s = self.ext.0.value().to_string();
        let ext_cy_s = self.ext.1.value().to_string();
        w.empty_with("a:off", &[("x", &off_x_s), ("y", &off_y_s)]);
        w.empty_with("a:ext", &[("cx", &ext_cx_s), ("cy", &ext_cy_s)]);
        w.empty_with("a:chOff", &[("x", "0"), ("y", "0")]);
        w.empty_with("a:chExt", &[("cx", &ext_cx_s), ("cy", &ext_cy_s)]);
        w.close("a:xfrm");
        w.close("p:grpSpPr");
        for c in &self.children {
            match c {
                GroupChild::Sp(s) => s.write_xml(w),
                GroupChild::Pic(p) => p.write_xml(w),
                GroupChild::CxnSp(c) => c.write_xml(w),
                GroupChild::Group(g) => g.write_xml(w),
                GroupChild::GraphicFrame(g) => g.write_xml(w),
            }
        }
        // p:style
        if let Some(style) = &self.style {
            style.write_xml(w);
        }
        // extLst（组合末尾）
        write_extlst(w, self.ext_lst.as_ref());
        w.close("p:grpSp");
    }
}

/// `<p:cxnSp>` 连接器。
#[derive(Clone, Debug, Default)]
pub struct Connector {
    pub id: u32,
    pub name: String,
    pub properties: ShapeProperties,
    /// 起点几何坐标（EMU，绝对值，相对于 slide 原点）。
    /// 仅在 `begin_type == None` 时由 OOXML 渲染。
    pub begin: Option<(Emu, Emu)>,
    /// 终点几何坐标（EMU）。
    pub end: Option<(Emu, Emu)>,
    /// 起点挂接（`stCxn`）——`Some((shape_id, idx))` 表示挂接到某 shape 的连接点。
    pub st_cxn: Option<(u32, u32)>,
    /// 终点挂接（`endCxn`）。
    pub end_cxn: Option<(u32, u32)>,
    /// 显式几何类型（直线 / 折线 / 曲线）。`None` 则用 `properties.xfrm.prst_geom`。
    pub connector_type: Option<crate::oxml::simpletypes::MsoConnectorType>,
    /// 主题样式引用（`p:style`，可选）。
    pub style: Option<ShapeStyle>,
    /// 扩展列表。
    pub ext_lst: Option<ExtensionList>,
}

impl Connector {
    /// 写 XML。
    ///
    /// 按 OOXML `CT_Connector` 的子元素顺序：
    ///
    /// ```text
    /// <p:cxnSp>
    ///   <p:nvCxnSpPr>...</p:nvCxnSpPr>
    ///   <p:spPr>...</p:spPr>
    ///   <p:style>...</p:style>          ← 可选
    ///   <p:extLst>...</p:extLst>        ← 可选
    /// </p:cxnSp>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:cxnSp");
        w.open("p:nvCxnSpPr");
        let id_s = self.id.to_string();
        w.empty_with("p:cNvPr", &[("id", &id_s), ("name", &self.name)]);
        w.empty("p:cNvCxnSpPr");
        w.empty("p:nvPr");
        if let Some((shape_id, idx)) = &self.st_cxn {
            let sid = shape_id.to_string();
            let idxs = idx.to_string();
            w.empty_with("p:stCxn", &[("id", &sid), ("idx", &idxs)]);
        }
        if let Some((shape_id, idx)) = &self.end_cxn {
            let sid = shape_id.to_string();
            let idxs = idx.to_string();
            w.empty_with("p:endCxn", &[("id", &sid), ("idx", &idxs)]);
        }
        w.close("p:nvCxnSpPr");
        self.properties.write_xml(w, "p:spPr");
        // p:style
        if let Some(style) = &self.style {
            style.write_xml(w);
        }
        // extLst
        write_extlst(w, self.ext_lst.as_ref());
        w.close("p:cxnSp");
    }
}

/// `<p:graphicFrame>` 容器：表格、图表等使用。
#[derive(Clone, Debug, Default)]
pub struct GraphicFrame {
    /// 图形框 ID（同一 slide 内唯一）。
    pub id: u32,
    /// 图形框名（仅显示用）。
    pub name: String,
    /// 形状属性（仅 xfrm 有效，不含 prstGeom/填充/边框）。
    pub properties: ShapeProperties,
    /// 内部图形。
    pub graphic: Graphic,
    /// 扩展列表。
    pub ext_lst: Option<ExtensionList>,
    /// **是否**为占位符（TODO-007 图表/表格占位符填充）。
    ///
    /// 为 `true` 时 `write_xml` 会在 `<p:nvPr>` 内写出 `<p:ph .../>`。
    /// 对应 OOXML 中 `<p:nvGraphicFramePr>/<p:nvPr>/<p:ph type="..." idx="..."/>`。
    pub is_placeholder: bool,
    /// 占位符索引（对应 `<p:ph idx="..."/>`）。
    pub ph_idx: Option<u32>,
    /// 占位符类型（对应 `<p:ph type="..."/>`，如 `"chart"` / `"tbl"` / `"obj"`）。
    pub ph_type: Option<String>,
}

#[derive(Clone, Debug)]
pub enum Graphic {
    Table(super::table::Table),
    /// 图表（`<c:chart>` 引用 chart part）。TODO-004 基础图表支持。
    ///
    /// 与 Table 不同，Chart 数据**不**直接嵌在 `<a:graphicData>` 中，
    /// 而是通过 `r:id="..."` 引用独立的 chartN.xml part。
    /// 序列化时只写出 `<c:chart xmlns:c=... xmlns:r=... r:id="..."/>` 自闭合元素。
    Chart(super::chart::Chart),
    /// OLE 对象嵌入（`<p:oleObj>` 引用 oleObjectN.bin）。TODO-043。
    ///
    /// 与 Chart 类似，OLE 数据**不**直接嵌在 `<a:graphicData>` 中，
    /// 而是通过 `r:id="..."` 引用独立的 `/ppt/embeddings/oleObjectN.bin` part。
    /// 序列化时写出 `<p:oleObj ...>` 完整元素（含 `<p:embed/>` 与可选 `<p:pic>` 图标）。
    OleObject(super::ole::OleObject),
    /// SmartArt 图形（`<dgm:relIds>` 引用 4 个 diagram parts）。TODO-037 最小保留。
    ///
    /// SmartArt 是 PowerPoint 的"智能图形"（流程图/层次结构/循环/矩阵等），
    /// 由 4 个独立 part 组成：
    /// - `diagramDataN.xml`（数据模型，r:dm）；
    /// - `diagramLayoutN.xml`（布局定义，r:lo）；
    /// - `diagramQuickStyleN.xml`（样式，r:qs）；
    /// - `diagramColorsN.xml`（颜色，r:cs）。
    ///
    /// **当前实现（最小保留）**：
    /// - 读路径：识别 `<a:graphicData uri=".../diagram">`，保留完整 XML（含 graphicData 外壳）；
    /// - 写路径：按原样输出 `raw_xml`，保证 slide XML 中的 SmartArt 引用 byte-exact 保留；
    /// - **限制**：不解析也不保留 4 个 diagram parts，read→save 后 SmartArt 会因为
    ///   diagram parts 丢失而无法渲染。完整 round-trip 需要 OPC 层保留未识别关系，
    ///   计划在 0.2.x 实现。
    SmartArt(SmartArtRef),
}

/// SmartArt 引用（TODO-037 最小保留）。
///
/// 持有完整的 `<a:graphicData uri=".../diagram">...</a:graphicData>` 元素 XML，
/// 用于 read→save 时 byte-exact 保留 SmartArt 在 slide XML 中的引用。
///
/// # 字段说明
///
/// - `raw_xml`：完整的 `<a:graphicData>` 元素 XML（**含** `<a:graphicData>` 外壳），
///   通常形如 `<a:graphicData uri=".../diagram"><dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/></a:graphicData>`。
///   序列化时由 `GraphicFrame::write_xml` 直接 raw 输出，**跳过** open_with/close 流程。
/// - `dm_rid` / `lo_rid` / `qs_rid` / `cs_rid`：从 raw_xml 中提取的 4 个关系 id，
///   仅供调用方查询使用，序列化时**不**单独输出（直接走 raw_xml）。
#[derive(Clone, Debug, Default)]
pub struct SmartArtRef {
    /// 完整的 `<a:graphicData>` 元素 XML（byte-exact 保留，含外壳）。
    pub raw_xml: String,
    /// 数据模型关系 id（`r:dm`，指向 `/ppt/diagrams/diagramDataN.xml`）。
    pub dm_rid: Option<String>,
    /// 布局定义关系 id（`r:lo`，指向 `/ppt/diagrams/diagramLayoutN.xml`）。
    pub lo_rid: Option<String>,
    /// 样式关系 id（`r:qs`，指向 `/ppt/diagrams/diagramQuickStyleN.xml`）。
    pub qs_rid: Option<String>,
    /// 颜色关系 id（`r:cs`，指向 `/ppt/diagrams/diagramColorsN.xml`）。
    pub cs_rid: Option<String>,
}

impl SmartArtRef {
    /// 从 4 个关系 id 构造 SmartArtRef（用于创建路径，TODO-037 创建 API）。
    ///
    /// 用 [`XmlWriter`](crate::oxml::writer::XmlWriter) 链式 API 构造 `raw_xml`，
    /// **不**用 `format!` 字符串拼接（遵守 `.trae/rules/project_rules.md` §5 安全红线）。
    ///
    /// # 参数
    /// - `dm_rid` / `lo_rid` / `qs_rid` / `cs_rid`：4 个关系 id，分别指向
    ///   data / layout / quickStyle / colors 四个 diagram part。
    ///
    /// # 生成内容
    ///
    /// ```xml
    /// <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
    ///   <dgm:relIds xmlns:dgm="..." xmlns:r="..."
    ///               r:dm="..." r:lo="..." r:qs="..." r:cs="..."/>
    /// </a:graphicData>
    /// ```
    ///
    /// # 示例
    ///
    /// ```no_run
    /// use pptx::oxml::shape::SmartArtRef;
    /// let r = SmartArtRef::from_rids("rId1", "rId2", "rId3", "rId4");
    /// assert!(r.raw_xml.contains("r:dm=\"rId1\""));
    /// ```
    pub fn from_rids(dm_rid: &str, lo_rid: &str, qs_rid: &str, cs_rid: &str) -> Self {
        let mut w = super::writer::XmlWriter::new();
        w.open_with(
            "a:graphicData",
            &[(
                "uri",
                "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            )],
        );
        w.empty_with(
            "dgm:relIds",
            &[
                ("xmlns:dgm", crate::oxml::ns::NS_DIAGRAM),
                ("xmlns:r", crate::oxml::ns::NS_DRAWING_RELS),
                ("r:dm", dm_rid),
                ("r:lo", lo_rid),
                ("r:qs", qs_rid),
                ("r:cs", cs_rid),
            ],
        );
        w.close("a:graphicData");
        Self {
            raw_xml: w.into_string(),
            dm_rid: Some(dm_rid.to_string()),
            lo_rid: Some(lo_rid.to_string()),
            qs_rid: Some(qs_rid.to_string()),
            cs_rid: Some(cs_rid.to_string()),
        }
    }
}

// 由于 `#[default]` 仅支持单元变体，Graphic 需手动实现 Default。
impl Default for Graphic {
    fn default() -> Self {
        Graphic::Table(super::table::Table::default())
    }
}

impl GraphicFrame {
    /// 写 XML。
    ///
    /// 按 OOXML `CT_GraphicalObjectFrame` 的子元素顺序：
    ///
    /// ```text
    /// <p:graphicFrame>
    ///   <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
    ///   <p:xfrm>...</p:xfrm>            ← 用 ShapeProperties::write_xfrm_only
    ///   <a:graphic>
    ///     <a:graphicData uri="...">     ← uri 必填
    ///       <a:tbl>...</a:tbl>          ← 或 chart / smartArt
    ///     </a:graphicData>
    ///   </a:graphic>
    ///   <p:extLst>...</p:extLst>        ← 可选
    /// </p:graphicFrame>
    /// ```
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        w.open("p:graphicFrame");
        // nvGraphicFramePr
        w.open("p:nvGraphicFramePr");
        let id_s = self.id.to_string();
        w.empty_with("p:cNvPr", &[("id", &id_s), ("name", &self.name)]);
        w.empty("p:cNvGraphicFramePr");
        // nvPr：若 is_placeholder=true，写出 <p:ph type="..." idx="..."/>（TODO-007 图表/表格占位符填充）
        if self.is_placeholder {
            w.open("p:nvPr");
            // idx_s 提到 pattrs 之前，保证生命周期覆盖 empty_with 调用
            let idx_s = self.ph_idx.map(|i| i.to_string());
            let mut pattrs: Vec<(&str, &str)> = Vec::new();
            if let Some(t) = self.ph_type.as_deref() {
                pattrs.push(("type", t));
            }
            if let Some(s) = idx_s.as_deref() {
                pattrs.push(("idx", s));
            }
            w.empty_with("p:ph", &pattrs);
            w.close("p:nvPr");
        } else {
            w.empty("p:nvPr");
        }
        w.close("p:nvGraphicFramePr");
        // xfrm（用专用方法，不输出 prstGeom/填充/边框）
        self.properties.write_xfrm_only(w);
        // graphic
        w.open("a:graphic");
        // SmartArt 走完整 raw_xml 输出（含 graphicData 外壳），其它类型走 open_with + 子元素 + close。
        // 这样设计是因为 SmartArt 的 raw_xml 在读路径上 byte-exact 保留了完整的 <a:graphicData> 元素，
        // 重新拆解为 open_with + 子元素会丢失原始格式（如命名空间声明位置、属性顺序等）。
        if let Graphic::SmartArt(s) = &self.graphic {
            // SmartArt 最小保留：直接输出读路径保存的 raw_xml（byte-exact）。
            // raw_xml 是完整的 <a:graphicData uri=".../diagram">...</a:graphicData> 元素。
            w.raw(&s.raw_xml);
        } else {
            // a:graphicData 的 uri 标识内部图形类型（表/图/OLE/...），
            // 必须在 `<a:graphicData uri="...">` 元素上指定，不能放到子元素上。
            let uri = match &self.graphic {
                Graphic::Table(_) => "http://schemas.openxmlformats.org/drawingml/2006/table",
                Graphic::Chart(_) => "http://schemas.openxmlformats.org/drawingml/2006/chart",
                Graphic::OleObject(_) => super::ole::OLE_GRAPHIC_DATA_URI,
                Graphic::SmartArt(_) => unreachable!(), // 已在 if 分支处理
            };
            w.open_with("a:graphicData", &[("uri", uri)]);
            match &self.graphic {
                Graphic::Table(t) => {
                    // table 自带 <a:tbl> 顶层元素
                    t.write_xml(w);
                }
                Graphic::Chart(c) => {
                    // chart 仅写出引用元素：<c:chart xmlns:c=... xmlns:r=... r:id="rid"/>
                    // 实际 chart 数据在独立的 chartN.xml part 中。
                    let rid = if c.rid.is_empty() {
                        "rId1"
                    } else {
                        c.rid.as_str()
                    };
                    w.empty_with(
                        "c:chart",
                        &[
                            ("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                            ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                            ("r:id", rid),
                        ],
                    );
                }
                Graphic::OleObject(o) => {
                    // oleObj 写出完整 <p:oleObj ...> 元素（含 <p:embed/> 与可选 <p:pic>）。
                    // 实际 OLE 二进制数据在独立的 oleObjectN.bin part 中。
                    // 注意：oleObj 元素需要 r 命名空间，但已在 <p:graphicFrame> 外层默认声明，
                    // 这里不再重复声明 xmlns:r。
                    o.write_xml(w);
                }
                Graphic::SmartArt(_) => unreachable!(),
            }
            w.close("a:graphicData");
        }
        w.close("a:graphic");
        // extLst
        write_extlst(w, self.ext_lst.as_ref());
        w.close("p:graphicFrame");
    }
}

// ============== 形状样式（p:style） ==============

/// 形状主题样式引用（`<p:style>`）。
///
/// OOXML 规范定义 `CT_ShapeStyle` 由 4 个引用组成：
/// `<a:lnRef>`、`<a:fillRef>`、`<a:effectRef>`、`<a:fontRef>`。
/// 在 PowerPoint 中，把"主题样式"应用到形状时，**就是写入这 4 个引用**；
/// 一旦 `<p:style>` 存在，spPr 中的 prstGeom/fill/ln 会被其覆盖。
///
/// # 与 python-pptx 的对应
///
/// - python-pptx 同样在 `pptx.oxml.shape.CT_Shape` 上提供 `style` 属性；
/// - python-pptx 通过 `shape.style = "..."` 或读取主题样式来设置。
/// - 本库采用**强类型**字段，便于序列化与回读。
#[derive(Clone, Debug, Default)]
pub struct ShapeStyle {
    /// `<a:lnRef>` 线条样式引用（`idx` + `schemeClr`）。
    pub line_ref: Option<StyleRef>,
    /// `<a:fillRef>` 填充样式引用。
    pub fill_ref: Option<StyleRef>,
    /// `<a:effectRef>` 效果样式引用。
    pub effect_ref: Option<StyleRef>,
    /// `<a:fontRef>` 字体样式引用（minor/major）。
    pub font_ref: Option<StyleRef>,
}

impl ShapeStyle {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        // 即便 4 个引用都是 None，也按 OOXML 规范至少输出 `<p:style>` 外壳以占位。
        w.open("p:style");
        if let Some(r) = &self.line_ref {
            write_style_ref(w, "a:lnRef", r);
        }
        if let Some(r) = &self.fill_ref {
            write_style_ref(w, "a:fillRef", r);
        }
        if let Some(r) = &self.effect_ref {
            write_style_ref(w, "a:effectRef", r);
        }
        if let Some(r) = &self.font_ref {
            // fontRef 特殊：`idx` 取值是 `"minor"` / `"major"` 而非数字
            let idx_attr = r.idx.as_deref().unwrap_or("minor");
            w.open_with("a:fontRef", &[("idx", idx_attr)]);
            if let Some(scheme) = &r.scheme_color {
                w.empty_with("a:schemeClr", &[("val", scheme)]);
            }
            w.close("a:fontRef");
        }
        w.close("p:style");
    }
}

/// 4 个样式引用的通用结构。
#[derive(Clone, Debug, Default)]
pub struct StyleRef {
    /// 索引（数字索引或 `"minor"` / `"major"` 字面量）。
    pub idx: Option<String>,
    /// 关联的主题色（`schemeClr val`）。
    pub scheme_color: Option<String>,
}

fn write_style_ref(w: &mut super::writer::XmlWriter, tag: &str, r: &StyleRef) {
    let idx_s = r.idx.clone().unwrap_or_else(|| "1".to_string());
    w.open_with(tag, &[("idx", &idx_s)]);
    if let Some(scheme) = &r.scheme_color {
        w.empty_with("a:schemeClr", &[("val", scheme)]);
    }
    w.close(tag);
}

// ============== 扩展列表（p:extLst） ==============

/// 扩展列表（`<p:extLst>`）。
///
/// OOXML 中"扩展"是 PowerPoint 用来存放私有属性 / future 兼容字段的地方，
/// 大多数"插入 → 背景效果" 等 UI 操作都会在 extLst 写一条 Microsoft 私有
/// 扩展。**只要形状被 UI 触碰过**，extLst 就很可能非空。
///
/// # 与 python-pptx 的对应
///
/// python-pptx 中 `CT_Shape` / `CT_GroupShape` / `CT_GraphicalObjectFrame`
/// 都有 `ext_lst` 属性；本库同等暴露。
#[derive(Clone, Debug, Default)]
pub struct ExtensionList {
    /// 扩展条目。每个条目包含 `uri`（必填）和原始 XML 字符串。
    pub entries: Vec<ExtensionEntry>,
}

#[derive(Clone, Debug)]
pub struct ExtensionEntry {
    /// 扩展的 URI（命名空间 GUID）。
    pub uri: String,
    /// 扩展 XML 内容（**已转义**的合法 XML 子串）。
    /// 通常为 `<a:...>...</a:...>` 或 `<p14:...>...</p14:...>`。
    pub xml: String,
}

impl ExtensionList {
    /// 写 XML。
    pub fn write_xml(&self, w: &mut super::writer::XmlWriter) {
        if self.entries.is_empty() {
            return;
        }
        w.open("p:extLst");
        for e in &self.entries {
            w.open_with("p:ext", &[("uri", e.uri.as_str())]);
            w.raw(&e.xml);
            w.close("p:ext");
        }
        w.close("p:extLst");
    }
}

/// extLst 便捷写出：`Some` 时写出，`None` 时跳过。
fn write_extlst(w: &mut super::writer::XmlWriter, ext: Option<&ExtensionList>) {
    if let Some(e) = ext {
        e.write_xml(w);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 验证 `ShapeLocks` 的序列化。
    ///
    /// 这是 TODO-027 的测试。
    #[test]
    fn shape_locks_serialization() {
        let locks = ShapeLocks {
            no_grp: true,
            no_select: true,
            no_resize: true,
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        locks.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("noGrp=\"1\""), "应包含 noGrp=1，实际: {s}");
        assert!(s.contains("noSelect=\"1\""), "应包含 noSelect=1，实际: {s}");
        assert!(s.contains("noResize=\"1\""), "应包含 noResize=1，实际: {s}");
        assert!(!s.contains("noMove"), "不应包含 noMove，实际: {s}");
        assert!(!s.contains("noRot"), "不应包含 noRot，实际: {s}");
    }

    /// 验证空 `ShapeLocks` 的 `is_empty` 判断。
    ///
    /// 这是 TODO-027 的测试。
    #[test]
    fn shape_locks_is_empty() {
        let empty = ShapeLocks::default();
        assert!(empty.is_empty());

        let non_empty = ShapeLocks {
            no_grp: true,
            ..Default::default()
        };
        assert!(!non_empty.is_empty());
    }

    /// 验证 `Sp` 带 locks 的完整序列化。
    ///
    /// 这是 TODO-027 的测试。
    #[test]
    fn sp_with_locks_serialization() {
        let sp = Sp {
            id: 1,
            name: "Test".into(),
            locks: Some(ShapeLocks {
                no_select: true,
                no_move: true,
                ..Default::default()
            }),
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        sp.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("a:spLocks"), "应包含 a:spLocks，实际: {s}");
        assert!(s.contains("noSelect=\"1\""), "应包含 noSelect=1，实际: {s}");
        assert!(s.contains("noMove=\"1\""), "应包含 noMove=1，实际: {s}");
        // cNvSpPr 应为 open-close 形式（因为有子元素）
        assert!(
            s.contains("<p:cNvSpPr>") || s.contains("<p:cNvSpPr "),
            "cNvSpPr 应为 open 标签，实际: {s}"
        );
        assert!(s.contains("</p:cNvSpPr>"), "应包含 </p:cNvSpPr>，实际: {s}");
    }

    /// 验证 `Pic` 带 `MediaKind::Video` 的序列化（TODO-033）。
    ///
    /// 关键断言：
    /// - `<a:videoFile r:link="rIdVideo1"/>` 必须出现在 `<p:nvPr>` 内；
    /// - `r:link` 属性（不是 `r:embed`）；
    /// - `xmlns:r` 命名空间声明必须出现在 `<p:nvPr>` 上。
    #[test]
    fn pic_with_video_media_serialize() {
        let pic = Pic {
            id: 1,
            name: "Video 1".into(),
            rid: "rIdImg1".into(),
            media: Some(MediaKind::Video {
                rid: "rIdVideo1".into(),
            }),
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        pic.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("<a:videoFile"), "应包含 a:videoFile，实际: {s}");
        assert!(
            s.contains("r:link=\"rIdVideo1\""),
            "应包含 r:link=\"rIdVideo1\"，实际: {s}"
        );
        assert!(
            !s.contains("r:embed=\"rIdVideo1\""),
            "视频不应使用 r:embed，实际: {s}"
        );
        assert!(
            s.contains("xmlns:r="),
            "nvPr 应声明 xmlns:r 命名空间，实际: {s}"
        );
    }

    /// 验证 `Pic` 带 `MediaKind::Audio` 的序列化（TODO-033）。
    #[test]
    fn pic_with_audio_media_serialize() {
        let pic = Pic {
            id: 2,
            name: "Audio 1".into(),
            rid: "rIdImg2".into(),
            media: Some(MediaKind::Audio {
                rid: "rIdAudio1".into(),
            }),
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        pic.write_xml(&mut w);
        let s = w.into_string();
        assert!(s.contains("<a:audioFile"), "应包含 a:audioFile，实际: {s}");
        assert!(
            s.contains("r:link=\"rIdAudio1\""),
            "应包含 r:link=\"rIdAudio1\"，实际: {s}"
        );
    }

    /// 验证普通 `Pic`（无 media）不写出 videoFile/audioFile（TODO-033）。
    #[test]
    fn pic_without_media_no_video_file() {
        let pic = Pic {
            id: 3,
            name: "Plain Picture".into(),
            rid: "rIdImg3".into(),
            // media = None
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        pic.write_xml(&mut w);
        let s = w.into_string();
        assert!(
            !s.contains("videoFile"),
            "普通图片不应包含 videoFile，实际: {s}"
        );
        assert!(
            !s.contains("audioFile"),
            "普通图片不应包含 audioFile，实际: {s}"
        );
    }

    /// 验证 `MediaKind` 的 `PartialEq` 与 `Default`（TODO-033）。
    #[test]
    fn media_kind_eq_and_default() {
        // Default 应为 Video { rid: "" }
        let d = MediaKind::default();
        assert!(matches!(d, MediaKind::Video { ref rid } if rid.is_empty()));

        // PartialEq
        let v1 = MediaKind::Video { rid: "rId1".into() };
        let v2 = MediaKind::Video { rid: "rId1".into() };
        let v3 = MediaKind::Video { rid: "rId2".into() };
        let a1 = MediaKind::Audio { rid: "rId1".into() };
        assert_eq!(v1, v2, "相同 rid 的 Video 应相等");
        assert_ne!(v1, v3, "不同 rid 的 Video 应不等");
        assert_ne!(v1, a1, "Video 与 Audio 即使 rid 相同也不等");
    }

    /// 验证 `GraphicFrame` 带 `Graphic::SmartArt` 的序列化（TODO-037 最小保留）。
    ///
    /// 关键断言：
    /// - `<a:graphicData uri=".../diagram">` 必须出现（uri 来自 raw_xml，不再由 write_xml 生成）；
    /// - `<dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>` 必须 byte-exact 保留；
    /// - 不应出现重复的 `<a:graphicData>` 外壳（write_xml 跳过 open_with/close）。
    #[test]
    fn graphic_frame_with_smartart_serialize() {
        // 模拟读路径保存的 raw_xml（完整 graphicData 元素，含外壳）
        let raw_xml =
            "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">\
                       <dgm:relIds r:dm=\"rId1\" r:lo=\"rId2\" r:qs=\"rId3\" r:cs=\"rId4\"/>\
                       </a:graphicData>";
        let frame = GraphicFrame {
            id: 100,
            name: "SmartArt 1".into(),
            graphic: Graphic::SmartArt(SmartArtRef {
                raw_xml: raw_xml.into(),
                dm_rid: Some("rId1".into()),
                lo_rid: Some("rId2".into()),
                qs_rid: Some("rId3".into()),
                cs_rid: Some("rId4".into()),
            }),
            ..Default::default()
        };
        let mut w = super::super::writer::XmlWriter::new();
        frame.write_xml(&mut w);
        let s = w.into_string();
        // 应包含 uri
        assert!(
            s.contains("uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\""),
            "应包含 diagram uri，实际: {s}"
        );
        // 应包含 dgm:relIds（byte-exact）
        assert!(
            s.contains("<dgm:relIds r:dm=\"rId1\" r:lo=\"rId2\" r:qs=\"rId3\" r:cs=\"rId4\"/>"),
            "应包含原始 dgm:relIds，实际: {s}"
        );
        // 不应出现重复的 graphicData 外壳（write_xml 跳过 open_with/close）
        // 统计 <a:graphicData 出现次数：应为 1
        let count = s.matches("<a:graphicData").count();
        assert_eq!(
            count, 1,
            "SmartArt 不应重复写出 graphicData 外壳，实际出现 {} 次: {s}",
            count
        );
    }

    /// 验证 `SmartArtRef` 的 Default 与字段访问（TODO-037）。
    #[test]
    fn smartart_ref_default() {
        let r = SmartArtRef::default();
        assert!(r.raw_xml.is_empty(), "默认 raw_xml 应为空");
        assert!(r.dm_rid.is_none(), "默认 dm_rid 应为 None");
        assert!(r.lo_rid.is_none(), "默认 lo_rid 应为 None");
        assert!(r.qs_rid.is_none(), "默认 qs_rid 应为 None");
        assert!(r.cs_rid.is_none(), "默认 cs_rid 应为 None");
    }
}
