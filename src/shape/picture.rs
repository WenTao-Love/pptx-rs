//! `Picture`：图片（`<p:pic>`）。
//!
//! 图片是高阶对象中**唯一**自带二进制 blob 的类型——其它形状都是纯 XML。
//! 保存时 [`Picture::blob`] 会被 `presentation::to_opc_package` 注入为
//! `/ppt/media/imageN.<ext>` part，并自动建立 `r:id` 关系。
//!
//! # 与 python-pptx 的对应
//!
//! - `pptx.shapes.picture.Picture` ←→ [`Picture`]；
//! - `Slide.shapes.add_picture(path, left, top, width, height)` 返回 [`Picture`]。
//!
//! # 内容类型推导
//!
//! [`content_type_for`] 按扩展名查表；未知扩展名回退到 `application/octet-stream`。
//! 该 Content-Type 会同步写入 `[Content_Types].xml`。
//!
//! # Base64 诊断
//!
//! [`Picture::base64`] 把当前图片 base64 编码——主要服务于在线图片预览 / 调试，
//! 不参与持久化。

use std::path::Path;

use base64::Engine;

use crate::oxml::shape::Pic as OxmlPic;
use crate::oxml::sppr::ShapeProperties;
use crate::shape::base::Shape;
use crate::units::Emu;

/// 一张图片。
#[derive(Clone, Debug, Default)]
pub struct Picture {
    /// 内部 oxml 句柄。
    pub(crate) pic: OxmlPic,
    /// 图片字节。保存时写入 zip。
    pub blob: Option<Vec<u8>>,
    /// 扩展名（含 `.`，如 `.png`），用于推导 Content-Type。
    pub ext: String,
}

impl Picture {
    /// 从本地文件创建。
    ///
    /// # 错误
    /// - [`crate::Error::Io`]：文件读取失败。
    pub fn from_path(path: impl AsRef<Path>) -> crate::Result<Self> {
        let path = path.as_ref();
        let blob = std::fs::read(path)?;
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| format!(".{}", e.to_lowercase()))
            .unwrap_or_else(|| ".png".to_string());
        Ok(Picture {
            pic: OxmlPic::default(),
            blob: Some(blob),
            ext,
        })
    }

    /// 从内存创建。
    pub fn from_bytes(bytes: impl Into<Vec<u8>>, ext: impl Into<String>) -> Self {
        let ext_s = ext.into();
        let ext_norm = if ext_s.starts_with('.') {
            ext_s
        } else {
            format!(".{}", ext_s)
        };
        Picture {
            pic: OxmlPic::default(),
            blob: Some(bytes.into()),
            ext: ext_norm,
        }
    }

    /// 从 oxml [`OxmlPic`] 构造（不携带 blob，用于读取路径）。
    pub fn from_pic(pic: OxmlPic) -> Self {
        Picture {
            pic,
            blob: None,
            ext: ".png".to_string(),
        }
    }

    /// 取出 oxml 引用。
    pub fn pic(&self) -> &OxmlPic {
        &self.pic
    }
    /// 取出 oxml 可变引用。
    pub fn pic_mut(&mut self) -> &mut OxmlPic {
        &mut self.pic
    }

    /// 形状属性不可变引用。
    pub fn properties(&self) -> &ShapeProperties {
        &self.pic.properties
    }
    /// 形状属性可变引用。
    pub fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.pic.properties
    }

    /// 设置图片填充模式（拉伸/平铺/无）。
    ///
    /// 对应 OOXML `<a:stretch>` / `<a:tile>` 元素。
    /// 默认为 `BlipFillMode::Stretch`（拉伸铺满）。
    pub fn set_fill_mode(&mut self, mode: crate::oxml::sppr::BlipFillMode) {
        self.pic.fill_mode = mode;
    }

    /// 设置为拉伸填充模式（便捷方法）。
    pub fn set_stretch(&mut self) {
        self.pic.fill_mode = crate::oxml::sppr::BlipFillMode::Stretch;
    }

    /// 设置为平铺填充模式。
    ///
    /// # 参数
    /// - `tx` / `ty`：水平/垂直偏移（EMU），`None` 表示不设置；
    /// - `sx` / `sy`：水平/垂直缩放（千分比，100000 = 100%），`None` 表示不设置；
    /// - `flip`：翻转模式（`"none"` / `"x"` / `"y"` / `"xy"`），`None` 表示不设置；
    /// - `algn`：对齐方式（`"tl"` / `"ctr"` / `"br"` 等），`None` 表示不设置。
    pub fn set_tile(
        &mut self,
        tx: Option<i64>,
        ty: Option<i64>,
        sx: Option<i32>,
        sy: Option<i32>,
        flip: Option<&str>,
        algn: Option<&str>,
    ) {
        self.pic.fill_mode = crate::oxml::sppr::BlipFillMode::Tile {
            tx,
            ty,
            sx,
            sy,
            flip: flip.map(|s| s.to_string()),
            algn: algn.map(|s| s.to_string()),
        };
    }

    /// 取当前填充模式。
    pub fn fill_mode(&self) -> &crate::oxml::sppr::BlipFillMode {
        &self.pic.fill_mode
    }

    /// 裁剪图片（`picture.crop_left/top/right/bottom` 的复合设置）。
    ///
    /// # 参数
    /// - `left` / `top` / `right` / `bottom`：**千分比**（取值 `0..=100000`），
    ///   表示从原图四边裁掉的占比。如 `left=25000` 即裁掉左 25%。
    ///   对应 python-pptx `picture.crop_left = 0.25`（库内做一次 ×100000 转换）。
    ///
    /// # 示例
    /// ```ignore
    /// // 裁掉左 10%、右 10%，保持上下
    /// pic.crop(10_000, 0, 10_000, 0);
    /// ```
    pub fn crop(&mut self, left: i32, top: i32, right: i32, bottom: i32) {
        self.pic.src_rect = Some((left, top, right, bottom));
    }

    /// 设置裁剪（TODO-044 高阶 API，`crop` 的 python-pptx 风格别名）。
    ///
    /// 与 [`crop`](Self::crop) 完全等价，仅方法名对齐 python-pptx
    /// `picture.set_crop(left, top, right, bottom)` 风格。
    pub fn set_crop(&mut self, left: i32, top: i32, right: i32, bottom: i32) {
        self.crop(left, top, right, bottom);
    }

    /// 清除裁剪（恢复原图）。
    pub fn clear_crop(&mut self) {
        self.pic.src_rect = None;
    }

    /// 取当前裁剪矩形（千分比），未裁剪返回 `None`。
    pub fn crop_rect(&self) -> Option<(i32, i32, i32, i32)> {
        self.pic.src_rect
    }

    /// 取左侧裁剪量（千分比，TODO-044 高阶 API）。
    ///
    /// 未裁剪时返回 0。
    pub fn crop_left(&self) -> i32 {
        self.pic.src_rect.map(|(l, _, _, _)| l).unwrap_or(0)
    }
    /// 取顶部裁剪量（千分比，TODO-044 高阶 API）。
    pub fn crop_top(&self) -> i32 {
        self.pic.src_rect.map(|(_, t, _, _)| t).unwrap_or(0)
    }
    /// 取右侧裁剪量（千分比，TODO-044 高阶 API）。
    pub fn crop_right(&self) -> i32 {
        self.pic.src_rect.map(|(_, _, r, _)| r).unwrap_or(0)
    }
    /// 取底部裁剪量（千分比，TODO-044 高阶 API）。
    pub fn crop_bottom(&self) -> i32 {
        self.pic.src_rect.map(|(_, _, _, b)| b).unwrap_or(0)
    }

    /// 设置左侧裁剪量（千分比，TODO-044 高阶 API），保留其它三边。
    pub fn set_crop_left(&mut self, left: i32) {
        let mut r = self.pic.src_rect.unwrap_or((0, 0, 0, 0));
        r.0 = left;
        self.pic.src_rect = Some(r);
    }
    /// 设置顶部裁剪量（千分比，TODO-044 高阶 API），保留其它三边。
    pub fn set_crop_top(&mut self, top: i32) {
        let mut r = self.pic.src_rect.unwrap_or((0, 0, 0, 0));
        r.1 = top;
        self.pic.src_rect = Some(r);
    }
    /// 设置右侧裁剪量（千分比，TODO-044 高阶 API），保留其它三边。
    pub fn set_crop_right(&mut self, right: i32) {
        let mut r = self.pic.src_rect.unwrap_or((0, 0, 0, 0));
        r.2 = right;
        self.pic.src_rect = Some(r);
    }
    /// 设置底部裁剪量（千分比，TODO-044 高阶 API），保留其它三边。
    pub fn set_crop_bottom(&mut self, bottom: i32) {
        let mut r = self.pic.src_rect.unwrap_or((0, 0, 0, 0));
        r.3 = bottom;
        self.pic.src_rect = Some(r);
    }

    // --------------------- 占位符 API（TODO-007 高阶） ---------------------
    //
    // 对标 python-pptx 中"图片占位符"——版式中的 `<p:ph type="pic"/>`。
    // 调用方通常通过 `slide.add_picture_to_placeholder(idx, path)` 创建带占位符
    // 标记的图片，PowerPoint 会按版式中的占位符位置/尺寸自动布局。

    /// 标记本图片为占位符填充（`<p:ph type="pic" idx="..."/>`）。
    ///
    /// # 参数
    /// - `ph_idx`：占位符 idx（对应版式中 `<p:ph idx="N"/>` 的 N）。
    /// - `ph_type`：占位符类型字符串（通常 `"pic"`，传 `None` 则不写出 `type` 属性）。
    ///
    /// # 示例
    /// ```no_run
    /// use pptx_rs::Presentation;
    /// use pptx_rs::units::Inches;
    ///
    /// let mut p = Presentation::new().unwrap();
    /// let counter = p.id_counter();
    /// let s = p.slides_mut().add_slide(counter).unwrap();
    /// // 假设版式有 idx=10 的图片占位符
    /// let mut pic = s.shapes_mut().add_picture("logo.png", Inches(1.0), Inches(1.0), Inches(4.0), Inches(3.0)).unwrap();
    /// pic.set_placeholder(10, Some("pic"));
    /// ```
    pub fn set_placeholder(&mut self, ph_idx: u32, ph_type: Option<&str>) {
        self.pic.is_placeholder = true;
        self.pic.ph_idx = Some(ph_idx);
        self.pic.ph_type = ph_type.map(|s| s.to_string());
    }

    /// 清除占位符标记（让本图片回到"自由图片"状态）。
    pub fn clear_placeholder(&mut self) {
        self.pic.is_placeholder = false;
        self.pic.ph_idx = None;
        self.pic.ph_type = None;
    }

    /// 是否为占位符填充。
    pub fn is_placeholder(&self) -> bool {
        self.pic.is_placeholder
    }

    /// 占位符 idx（仅当 [`is_placeholder`](Self::is_placeholder) 为 true 时有意义）。
    pub fn ph_idx(&self) -> Option<u32> {
        self.pic.ph_idx
    }

    /// 占位符类型字符串（如 `"pic"`）。
    pub fn ph_type(&self) -> Option<&str> {
        self.pic.ph_type.as_deref()
    }

    /// base64 输出当前图片（便于诊断 / 浏览器内嵌）。
    pub fn base64(&self) -> Option<String> {
        self.blob
            .as_ref()
            .map(|b| base64::engine::general_purpose::STANDARD.encode(b))
    }

    // --------------------- 媒体 API（TODO-033 高阶） ---------------------
    //
    // 对标 python-pptx 中"媒体形状"——OOXML 中视频/音频形状实际上是带
    // `<a:videoFile r:link="..."/>` / `<a:audioFile r:link="..."/>` 的 `<p:pic>`，
    // 用海报帧图片（`<a:blip r:embed="..."/>`）作为视觉占位。
    // 调用方通常通过 `slide.shapes_mut().add_video(...)` / `add_audio(...)` 创建，
    // 这里提供低层 setter 供高级用户手动构造媒体形状。

    /// 把本图片标记为**视频**形状（`<a:videoFile r:link="..."/>`）。
    ///
    /// 调用后 `Pic::write_xml` 会在 `<p:nvPr>` 内写出 `<a:videoFile r:link="..."/>`，
    /// PowerPoint 渲染时会把该 `<p:pic>` 当作视频形状处理（双击播放）。
    ///
    /// # 参数
    /// - `rid`：媒体 part 的关系 id（指向 `/ppt/media/mediaN.mp4`）。
    ///   该 rid **必须**与 `slideN.xml.rels` 中的 `<Relationship Type=".../video"/>` 一致，
    ///   由 [`crate::slide::ShapesMut::add_video`] 自动分配，或由调用方手动维护。
    ///
    /// # 与海报帧 rid 的区别
    /// - 海报帧图片的 `r:embed` 关系通过 `pic_mut().rid` 设置（指向 imageN.png）；
    /// - `set_video` 设置的是**视频文件**的 `r:link` 关系（指向 mediaN.mp4）。
    ///   二者独立，互不影响。
    ///
    /// # 示例
    /// ```no_run
    /// # use pptx_rs::Presentation;
    /// # use pptx_rs::units::Inches;
    /// # let mut p = Presentation::new().unwrap();
    /// # let counter = p.id_counter();
    /// # let s = p.slides_mut().add_slide(counter).unwrap();
    /// let mut pic = s.shapes_mut().add_picture("poster.png",
    ///     Inches(1.0), Inches(1.0), Inches(4.0), Inches(3.0)).unwrap();
    /// pic.set_video("rIdVideo1");
    /// ```
    pub fn set_video(&mut self, rid: impl Into<String>) {
        self.pic.media = Some(crate::oxml::shape::MediaKind::Video { rid: rid.into() });
    }

    /// 把本图片标记为**音频**形状（`<a:audioFile r:link="..."/>`）。
    ///
    /// 与 [`Self::set_video`] 对称，仅媒体类型不同。
    ///
    /// # 参数
    /// - `rid`：媒体 part 的关系 id（指向 `/ppt/media/mediaN.mp3`）。
    pub fn set_audio(&mut self, rid: impl Into<String>) {
        self.pic.media = Some(crate::oxml::shape::MediaKind::Audio { rid: rid.into() });
    }

    /// 取当前媒体类型（视频/音频/None）。
    ///
    /// 返回 `None` 表示普通图片；`Some(MediaKind::Video { rid })` / `Some(MediaKind::Audio { rid })`
    /// 表示已通过 [`Self::set_video`] / [`Self::set_audio`] 标记为媒体形状。
    pub fn media_kind(&self) -> Option<&crate::oxml::shape::MediaKind> {
        self.pic.media.as_ref()
    }

    /// 清除媒体标记（让本图片回到普通图片状态）。
    pub fn clear_media(&mut self) {
        self.pic.media = None;
    }
}

impl Shape for Picture {
    fn id(&self) -> u32 {
        self.pic.id
    }
    fn set_id(&mut self, id: u32) {
        self.pic.id = id;
    }
    fn name(&self) -> &str {
        &self.pic.name
    }
    fn set_name(&mut self, name: String) {
        self.pic.name = name;
    }
    fn shape_type(&self) -> &'static str {
        "picture"
    }

    fn left(&self) -> Emu {
        self.pic.properties.xfrm.off_x.unwrap_or_default()
    }
    fn set_left(&mut self, emu: Emu) {
        self.pic.properties.xfrm.off_x = Some(emu);
    }
    fn top(&self) -> Emu {
        self.pic.properties.xfrm.off_y.unwrap_or_default()
    }
    fn set_top(&mut self, emu: Emu) {
        self.pic.properties.xfrm.off_y = Some(emu);
    }
    fn width(&self) -> Emu {
        self.pic.properties.xfrm.ext_cx.unwrap_or_default()
    }
    fn set_width(&mut self, emu: Emu) {
        self.pic.properties.xfrm.ext_cx = Some(emu);
    }
    fn height(&self) -> Emu {
        self.pic.properties.xfrm.ext_cy.unwrap_or_default()
    }
    fn set_height(&mut self, emu: Emu) {
        self.pic.properties.xfrm.ext_cy = Some(emu);
    }

    fn rotation(&self) -> f64 {
        self.pic.properties.rot_deg.unwrap_or(0.0)
    }
    fn set_rotation(&mut self, deg: f64) {
        self.pic.properties.rot_deg = Some(deg);
        let rot = (deg * 60_000.0) as i32;
        self.pic.properties.xfrm.rot = Some(rot);
    }
}

/// 推导 Content-Type（按扩展名）。
///
/// 未知扩展名回退到 `application/octet-stream`。
pub fn content_type_for(ext: &str) -> &'static str {
    match ext.to_ascii_lowercase().as_str() {
        ".png" => "image/png",
        ".jpg" | ".jpeg" => "image/jpeg",
        ".gif" => "image/gif",
        ".bmp" => "image/bmp",
        ".svg" => "image/svg+xml",
        ".tif" | ".tiff" => "image/tiff",
        _ => "application/octet-stream",
    }
}
