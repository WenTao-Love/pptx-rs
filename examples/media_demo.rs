//! # 端到端示例：在幻灯片中嵌入视频和音频（TODO-033）
//!
//! 演示 `pptx-rs` 的音视频嵌入 API：
//!
//! 1. 新建演示文稿；
//! 2. 添加一张幻灯片；
//! 3. 生成示例视频/音频文件（简单的二进制 blob）+ 海报帧 PNG；
//! 4. 通过 `add_video` 把视频嵌入到幻灯片；
//! 5. 通过 `add_audio` 把音频嵌入到幻灯片；
//! 6. 保存为 `media_demo.pptx`。
//!
//! # 运行方式
//!
//! ```bash
//! cargo run --example media_demo
//! ```
//!
//! # 预期输出
//!
//! - `media_demo.pptx` 包含 1 张幻灯片；
//! - 幻灯片上有 1 个视频形状 + 1 个音频形状；
//! - 视频通过 `<a:videoFile r:link="rIdVideo1"/>` 引用独立的 `/ppt/media/media1.mp4` part；
//! - 音频通过 `<a:audioFile r:link="rIdAudio1"/>` 引用独立的 `/ppt/media/media2.mp3` part；
//! - 海报帧图片通过 `<a:blip r:embed="rIdImg1"/>` / `rIdImg2` 引用 `/ppt/media/image1.png` / `image2.png`；
//! - slide1.xml.rels 中含 `<Relationship Type=".../video"/>` 与 `<Relationship Type=".../audio"/>`。
//!
//! # 注意
//!
//! 本示例为了演示 API 流程，使用简单的二进制 blob 作为视频/音频数据。
//! 实际使用时，应传入真实的 `.mp4` / `.mp3` 文件，PowerPoint 才能正确解码播放。
//! 海报帧图片建议传入代表视频首帧的 `.png` / `.jpg` 文件。

use pptx::{Inches, Presentation};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // ---------- 1) 新建演示文稿 ----------
    let mut prs = Presentation::new()?;

    // ---------- 2) 添加幻灯片 ----------
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide(counter)?;

    // ---------- 3) 生成示例视频/音频/海报文件 ----------
    // 实际使用时，应替换为真实的媒体文件路径。
    // 这里用简单的二进制文件作为演示数据（PowerPoint 无法解码，但文件结构完整）。
    let video_path = "media_sample_video.mp4";
    let audio_path = "media_sample_audio.mp3";
    let poster_path = "media_sample_poster.png";

    // 写入示例视频数据（伪 MP4 头 + 数据）
    let video_data: Vec<u8> = {
        let mut v = Vec::new();
        // 伪 MP4 ftyp box 头（仅用于演示，PowerPoint 实际无法解码）
        v.extend_from_slice(&[0x00, 0x00, 0x00, 0x18, b'f', b't', b'y', b'p']);
        v.extend_from_slice(b"mp42");
        v.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]);
        v.extend_from_slice(b"mp42isom");
        // 后续填充伪数据
        v.extend_from_slice(b"\x00\x00\x00\x01VIDEO_BLOB_FOR_PPTX_RS_DEMO");
        v
    };
    std::fs::write(video_path, &video_data)?;

    // 写入示例音频数据（伪 MP3 头 + 数据）
    let audio_data: Vec<u8> = {
        let mut v = Vec::new();
        // 伪 MP3 ID3 头（仅用于演示）
        v.extend_from_slice(b"ID3\x03\x00\x00\x00\x00\x00\x00");
        v.extend_from_slice(b"AUDIO_BLOB_FOR_PPTX_RS_DEMO");
        v
    };
    std::fs::write(audio_path, &audio_data)?;

    // 写入示例海报帧 PNG（1x1 蓝色 PNG，48 字节）
    let poster_data: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xD4,
        0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00,
        0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    std::fs::write(poster_path, poster_data)?;

    // ---------- 4) 嵌入视频 ----------
    // add_video 内部流程：
    //   1. 读取 video_path 字节 → 写入 /ppt/media/media1.mp4
    //   2. 读取 poster_path 字节 → 写入 /ppt/media/image1.png
    //   3. 分配 rIdImg1（海报帧 r:embed） + rIdVideo1（视频 r:link）
    //   4. 生成 <p:pic>，nvPr 内含 <a:videoFile r:link="rIdVideo1"/>
    let _video = slide.shapes_mut().add_video(
        video_path,
        Some(poster_path),
        Inches(0.5),
        Inches(1.0),
        Inches(6.0),
        Inches(4.0),
    )?;
    println!(
        "已嵌入视频：rid={}，media_kind={:?}",
        _video.pic().rid,
        _video.media_kind()
    );

    // ---------- 5) 嵌入音频 ----------
    // add_audio 与 add_video 对称：
    //   - 视频写入 /ppt/media/media2.mp4 → 此处音频写入 /ppt/media/media2.mp3
    //   - 关系类型 .../audio（区别于 .../video）
    //   - <a:audioFile r:link="rIdAudio1"/>
    // 此处 poster_path 传 None，演示内置 1x1 透明 PNG 占位能力。
    let _audio = slide.shapes_mut().add_audio(
        audio_path,
        None, // 用内置 1x1 透明 PNG 占位（PowerPoint 会显示空白）
        Inches(7.0),
        Inches(5.5),
        Inches(2.0),
        Inches(2.0),
    )?;
    println!(
        "已嵌入音频：rid={}，media_kind={:?}",
        _audio.pic().rid,
        _audio.media_kind()
    );

    // ---------- 6) 保存 ----------
    prs.save("media_demo.pptx")?;

    // 清理临时文件
    let _ = std::fs::remove_file(video_path);
    let _ = std::fs::remove_file(audio_path);
    let _ = std::fs::remove_file(poster_path);

    println!(
        "\n已生成 media_demo.pptx（{} 张幻灯片，1 个视频 + 1 个音频）",
        prs.slides().len()
    );
    println!("视频二进制写入 /ppt/media/media1.mp4");
    println!("音频二进制写入 /ppt/media/media2.mp3");
    println!("海报帧图片写入 /ppt/media/image1.png");
    println!("slide1.xml.rels 中含 video / audio / image 三种关系");
    println!("slide1.xml 中含 <a:videoFile r:link=\"rIdVideo1\"/> 与 <a:audioFile r:link=\"rIdAudio1\"/>");

    Ok(())
}
