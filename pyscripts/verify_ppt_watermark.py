# -*- coding: utf-8 -*-
"""验证加水印后的 .ppt 文件结构完整性。"""
import sys
import io
import struct
import olefile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    is_container = (ver == 0xF)
    return (ver, inst, rec_type, rec_len, is_container)

def find_watermark(data, offset, end, depth=0, max_depth=8):
    """递归查找水印 SpContainer。"""
    results = []
    pos = offset
    while pos + 8 <= end:
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len, is_container = hdr
        total_len = 8 + rec_len

        if is_container and depth < max_depth:
            child_results = find_watermark(data, pos + 8, pos + 8 + rec_len, depth + 1, max_depth)
            results.extend(child_results)

        # 查找水印文本
        if rec_type == 0x0FA0:  # TextCharsAtom
            if rec_len > 0:
                try:
                    text = data[pos+8:pos+8+rec_len].decode('utf-16-le', errors='replace').rstrip('\x00')
                    if '水印' in text:
                        results.append(('WatermarkText', pos, rec_len, text))
                except:
                    pass

        pos += total_len
        if not is_container and rec_len == 0:
            break
    return results

def main():
    wm_path = "_test_out/wm_心理账户理论.ppt"

    # 1. 验证 OLE2 容器完整性
    try:
        ole = olefile.OleFileIO(wm_path)
        print("[OK] OLE2 容器解析成功")
        streams = ['/'.join(e) for e in ole.listdir()]
        print(f"  Streams: {streams}")
    except Exception as e:
        print(f"[FAIL] OLE2 容器解析失败: {e}")
        sys.exit(1)

    # 2. 验证 PowerPoint Document stream
    try:
        ppt_data = ole.openstream("PowerPoint Document").read()
        print(f"[OK] PowerPoint Document stream 读取成功，大小: {len(ppt_data)}")
    except Exception as e:
        print(f"[FAIL] PowerPoint Document stream 读取失败: {e}")
        sys.exit(1)

    # 3. 查找水印文本
    watermarks = find_watermark(ppt_data, 0, len(ppt_data))
    if watermarks:
        print(f"[OK] 找到 {len(watermarks)} 个水印文本")
        for name, pos, rec_len, text in watermarks[:5]:
            print(f"  {name} at offset={pos}, len={rec_len}, text={text!r}")
    else:
        print("[WARN] 未找到水印文本")

    # 4. 验证 Current User stream
    try:
        cu_data = ole.openstream("Current User").read()
        print(f"[OK] Current User stream 读取成功，大小: {len(cu_data)}")
        # 检查 headerToken（offset 12）
        header_token = struct.unpack_from('<I', cu_data, 12)[0]
        if header_token == 0xE391C05F:
            print("  headerToken = 0xE391C05F (未加密)")
        elif header_token == 0xF3D1C4DF:
            print("  headerToken = 0xF3D1C4DF (已加密)")
        else:
            print(f"  headerToken = 0x{header_token:08X} (未知)")
    except Exception as e:
        print(f"[FAIL] Current User stream 读取失败: {e}")

    ole.close()
    print("[SUCCESS] 文件结构验证完成")

if __name__ == "__main__":
    main()
