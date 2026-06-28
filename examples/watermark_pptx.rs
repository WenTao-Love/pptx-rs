//! 读取现有 .pptx + 给每张幻灯片加水印 + 保存。
//!
//! 使用 OpcPackage 底层操作：直接在 slide XML 中注入水印 shape。
//!
//! 关键：水印 XML 不能重新声明 xmlns:p 和 xmlns:a 命名空间，
//! 因为父元素 <p:sld> 已经声明了。WPS 对重复命名空间声明非常敏感，
//! 会导致 shape 不渲染。
//!
//! 用法：在 pptx-rs 目录执行
//!   cargo run --example watermark_pptx

use std::path::Path;

use pptx::opc::OpcPackage;

/// 水印 shape XML 模板。
///
/// **不重新声明 xmlns:p 和 xmlns:a**（已在父元素 <p:sld> 中声明）。
/// 格式与 python-pptx 生成的 textbox 完全一致，确保 WPS 兼容。
///
/// - 浅灰色半透明 40pt 粗体文字 "pptx-rs WATERMARK"
/// - 旋转 -45°（OOXML 单位：-45 × 60000 = -2700000）
/// - 位于幻灯片中央
/// - `WATERMARK_ID` 会被替换为唯一 id
const WATERMARK_SP: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="WATERMARK_ID" name="pptx-rs Watermark"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="-2700000"><a:off x="457200" y="2743200"/><a:ext cx="8229600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="4000" b="1"><a:solidFill><a:srgbClr val="BFBFBF"><a:alpha val="40000"/></a:srgbClr></a:solidFill><a:latin typeface="Calibri"/><a:ea typeface="宋体"/></a:rPr><a:t>pptx-rs WATERMARK</a:t></a:r></a:p></p:txBody></p:sp>"#;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let in_path = Path::new("_test/文旅IP人设打造抖音短视频运营方案.pptx");
    println!("正在读取 {}", in_path.display());

    // 1) 用 OpcPackage 加载（保留所有原始 part 的 blob）
    let mut pkg = OpcPackage::load(in_path)?;
    println!("包含 {} 个 part", pkg.part_count());

    // 2) 找到所有 slide*.xml 的 partname
    let mut slide_names: Vec<String> = Vec::new();
    for part in pkg.iter_parts() {
        let n = part.partname.as_str();
        if n.starts_with("/ppt/slides/slide") && n.ends_with(".xml") {
            slide_names.push(n.to_string());
        }
    }
    slide_names.sort();
    println!("共发现 {} 张幻灯片", slide_names.len());

    // 3) 对每个 slide XML 注入水印 sp
    let mut max_id: u32 = 9000;
    for name in &slide_names {
        max_id += 1;
        if let Some(part) = pkg.get_part_mut(name) {
            let new_xml = inject_watermark(&part.blob, max_id);
            part.blob = new_xml.into_bytes();
        }
    }

    // 4) 输出
    std::fs::create_dir_all("_test_out")?;
    let out_path = "_test_out/wm_文旅IP人设打造抖音短视频运营方案.pptx";
    pkg.save(out_path)?;
    println!("已生成 {}", out_path);
    Ok(())
}

/// 在 slide XML 的 `</p:spTree>` 之前注入水印 shape。
///
/// 水印放在 spTree 末尾（z-order 最顶层），配合半透明效果，
/// 既能看到水印又不严重影响阅读。
fn inject_watermark(blob: &[u8], id: u32) -> String {
    let s = std::str::from_utf8(blob).unwrap_or("").to_string();
    let sp_xml = WATERMARK_SP.replace("WATERMARK_ID", &id.to_string());

    // 检查是否已经存在水印
    if s.contains("pptx-rs Watermark") || s.contains("pptx-rs WATERMARK") {
        return s;
    }

    // 在 </p:spTree> 之前插入水印
    if let Some(pos) = s.find("</p:spTree>") {
        let mut out = String::with_capacity(s.len() + sp_xml.len());
        out.push_str(&s[..pos]);
        out.push_str(&sp_xml);
        out.push_str(&s[pos..]);
        out
    } else {
        s
    }
}
