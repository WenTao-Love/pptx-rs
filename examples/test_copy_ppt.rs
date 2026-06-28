//! 测试：直接复制原始 .ppt 的 stream 到新 OLE2 容器（不做任何修改）。
//!
//! 如果输出文件无法被 PowerPoint 打开，说明 cfb crate 创建的 OLE2 容器本身不兼容。

use std::io::{Read, Write};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = std::path::Path::new("_test/心理账户理论.ppt");
    let file_data = std::fs::read(input)?;
    let cursor = std::io::Cursor::new(file_data);
    let mut comp = cfb::CompoundFile::open(cursor)?;

    // 先收集所有路径
    let mut stream_paths: Vec<String> = Vec::new();
    let mut storage_paths: Vec<String> = Vec::new();

    for entry in comp.walk() {
        let path = entry.path().to_string_lossy().to_string();
        if entry.is_root() {
            continue;
        }
        if entry.is_storage() {
            storage_paths.push(path);
        } else if entry.is_stream() {
            stream_paths.push(path);
        }
    }

    // 再读取所有 streams
    let mut streams: Vec<(String, Vec<u8>)> = Vec::new();
    for path in &stream_paths {
        let mut stream = comp.open_stream(path)?;
        let mut data = Vec::new();
        stream.read_to_end(&mut data)?;
        streams.push((path.clone(), data));
    }

    // 用 V3 格式创建新 OLE2 容器，设置 CLSID
    let ppt_clsid = uuid::Uuid::parse_str("64818d10-4f9b-11cf-86ea-00aa00b929e8")?;
    let mut buf = std::io::Cursor::new(Vec::new());
    {
        let mut new_comp = cfb::CompoundFile::create_with_version(cfb::Version::V3, &mut buf)?;
        new_comp.set_storage_clsid("/", ppt_clsid)?;

        for storage_path in &storage_paths {
            new_comp.create_storage_all(storage_path)?;
        }

        for (path, data) in &streams {
            {
                let mut stream = new_comp.create_stream(path)?;
                stream.write_all(data)?;
                stream.flush()?;
            }
        }

        new_comp.flush()?;
    }

    std::fs::create_dir_all("_test_out")?;
    std::fs::write("_test_out/test_copy.ppt", buf.into_inner())?;
    println!("已创建测试文件：_test_out/test_copy.ppt");
    println!("请用 WPS/PowerPoint 打开，验证是否能正常打开");

    Ok(())
}
