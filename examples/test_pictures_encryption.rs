//! 测试 Pictures stream 加密的 round-trip 正确性。
//!
//! 读取原始文件的 Pictures stream，先解析所有 record 边界，
//! 再按 record 字段用与 protect_ppt.rs 相同的逻辑加密/解密，
//! 比较是否与原始数据完全一致。

use sha1::{Digest, Sha1};
use std::io::Read;

const PASSWORD: &str = "pptx-rs-secret";
const KEY_SIZE_BITS: u32 = 128;
const SALT_SIZE: usize = 16;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input_path = std::path::Path::new("_test/心理账户理论.ppt");
    let file_data = std::fs::read(input_path)?;
    let cursor = std::io::Cursor::new(file_data);
    let mut comp = cfb::CompoundFile::open(cursor)?;

    let mut pictures = Vec::new();
    {
        let mut stream = comp.open_stream("Pictures")?;
        stream.read_to_end(&mut pictures)?;
    }

    let original = pictures.clone();

    // 生成 salt
    let salt: Vec<u8> = (0..SALT_SIZE).map(|i| i as u8).collect();

    // 解析 record 边界
    let records = parse_picture_records(&pictures)?;
    println!("解析到 {} 个 picture records", records.len());

    // 加密
    encrypt_pictures_stream(PASSWORD, &salt, KEY_SIZE_BITS, &mut pictures, &records)?;

    // 解密（与加密对称，使用相同的 record 边界）
    encrypt_pictures_stream(PASSWORD, &salt, KEY_SIZE_BITS, &mut pictures, &records)?;

    if pictures == original {
        println!("Pictures stream round-trip 验证通过！");
        Ok(())
    } else {
        for (i, (a, b)) in original.iter().zip(pictures.iter()).enumerate() {
            if a != b {
                eprintln!(
                    "第一个不同字节 @ {}: 原始=0x{:02X}, 解密后=0x{:02X}",
                    i, a, b
                );
                break;
            }
        }
        eprintln!(
            "Pictures stream round-trip 验证失败！大小: 原始={}, 解密后={}",
            original.len(),
            pictures.len()
        );
        Err("Pictures stream 加密/解密不对称".into())
    }
}

#[derive(Debug, Clone)]
struct PicRecord {
    offset: usize,
    len: usize, // 包含 header 的总长度
    rec_type: u16,
    rec_inst: u16,
}

fn parse_picture_records(data: &[u8]) -> Result<Vec<PicRecord>, Box<dyn std::error::Error>> {
    let mut records = Vec::new();
    let mut offset = 0;
    while offset + 8 <= data.len() {
        let ver_inst = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let rec_type = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
        let rlen = u32::from_le_bytes([
            data[offset + 4],
            data[offset + 5],
            data[offset + 6],
            data[offset + 7],
        ]) as usize;
        let rec_inst = (ver_inst >> 4) & 0x0FFF;
        let total_len = 8 + rlen;
        if offset + total_len > data.len() {
            return Err(format!("record @ {} 越界: len={}", offset, total_len).into());
        }
        records.push(PicRecord {
            offset,
            len: total_len,
            rec_type,
            rec_inst,
        });
        offset += total_len;
    }
    Ok(records)
}

struct Rc4 {
    s: [u8; 256],
    i: usize,
    j: usize,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        let mut s = [0u8; 256];
        for (i, slot) in s.iter_mut().enumerate() {
            *slot = i as u8;
        }
        let mut j = 0;
        for i in 0..256 {
            j = (j + s[i] as usize + key[i % key.len()] as usize) % 256;
            s.swap(i, j);
        }
        Self { s, i: 0, j: 0 }
    }

    fn process(&mut self, data: &mut [u8]) {
        for b in data.iter_mut() {
            self.i = (self.i + 1) % 256;
            self.j = (self.j + self.s[self.i] as usize) % 256;
            self.s.swap(self.i, self.j);
            let k = self.s[(self.s[self.i] as usize + self.s[self.j] as usize) % 256];
            *b ^= k;
        }
    }
}

fn make_key(password: &str, salt: &[u8], key_bits: u32, block: u32) -> Vec<u8> {
    let password_utf16le: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&password_utf16le);
    let h0 = hasher.finalize();

    let mut hasher = Sha1::new();
    hasher.update(h0);
    hasher.update(block.to_le_bytes());
    let hfinal = hasher.finalize();

    let key_bytes = (key_bits / 8) as usize;
    if key_bits == 40 {
        let mut key = vec![0u8; 16];
        key[..5].copy_from_slice(&hfinal[..5]);
        key
    } else {
        hfinal[..key_bytes].to_vec()
    }
}

fn encrypt_pictures_stream(
    password: &str,
    salt: &[u8],
    key_bits: u32,
    data: &mut [u8],
    records: &[PicRecord],
) -> Result<(), Box<dyn std::error::Error>> {
    for rec in records {
        let offset = rec.offset;
        let rlen = rec.len - 8;
        let rec_type = rec.rec_type;
        let rec_inst = rec.rec_inst;

        encrypt_pic_field(password, salt, key_bits, data, offset, 8);

        let mut pos = offset + 8;
        let end_offset = pos + rlen;

        if rec_type == 0xF007 {
            let cb_name = u16::from_le_bytes([data[pos + 33], data[pos + 34]]) as usize;
            let parts = [1usize, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1];
            for part in &parts {
                encrypt_pic_field(password, salt, key_bits, data, pos, *part);
                pos += part;
            }
            if cb_name > 0 {
                encrypt_pic_field(password, salt, key_bits, data, pos, cb_name);
                pos += cb_name;
            }
            if pos >= end_offset {
                continue;
            }
            let ver_inst2 = u16::from_le_bytes([data[pos], data[pos + 1]]);
            let rec_type2 = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
            let rec_inst2 = (ver_inst2 >> 4) & 0x0FFF;
            encrypt_pic_field(password, salt, key_bits, data, pos, 8);
            pos += 8;
            encrypt_blip_fields(
                password, salt, key_bits, data, &mut pos, end_offset, rec_type2, rec_inst2,
            );
        } else {
            encrypt_blip_fields(
                password, salt, key_bits, data, &mut pos, end_offset, rec_type, rec_inst,
            );
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn encrypt_blip_fields(
    password: &str,
    salt: &[u8],
    key_bits: u32,
    data: &mut [u8],
    pos: &mut usize,
    end_offset: usize,
    rec_type: u16,
    rec_inst: u16,
) {
    let rgb_uid_cnt = if rec_inst == 0x217
        || rec_inst == 0x3D5
        || rec_inst == 0x46B
        || rec_inst == 0x543
        || rec_inst == 0x6E1
        || rec_inst == 0x6E3
        || rec_inst == 0x6E5
        || rec_inst == 0x7A9
    {
        2
    } else {
        1
    };

    for _ in 0..rgb_uid_cnt {
        encrypt_pic_field(password, salt, key_bits, data, *pos, 16);
        *pos += 16;
    }

    let next_bytes = if rec_type == 0xF01A || rec_type == 0xF01B || rec_type == 0xF01C {
        34
    } else {
        1
    };
    encrypt_pic_field(password, salt, key_bits, data, *pos, next_bytes);
    *pos += next_bytes;

    let blip_len = end_offset - *pos;
    if blip_len > 0 {
        encrypt_pic_field(password, salt, key_bits, data, *pos, blip_len);
        *pos += blip_len;
    }
}

fn encrypt_pic_field(
    password: &str,
    salt: &[u8],
    key_bits: u32,
    data: &mut [u8],
    offset: usize,
    len: usize,
) {
    if len == 0 {
        return;
    }
    let key = make_key(password, salt, key_bits, 0);
    let mut rc4 = Rc4::new(&key);
    rc4.process(&mut data[offset..offset + len]);
}
