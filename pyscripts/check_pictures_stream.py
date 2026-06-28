"""检查 msoffcrypto 测试文件 rc4cryptoapi_password.ppt 的 Pictures stream 加密方式。

目标：
1. 确认 rc4cryptoapi_password.ppt 是否有 Pictures stream
2. 如果有，检查其结构（record type, recLen 等）
3. 对比我们生成的加密文件的 Pictures stream 结构
"""
import olefile
import sys
import os

def check_ppt_structure(ppt_path, label):
    print(f"\n{'='*60}")
    print(f"文件: {ppt_path}")
    print(f"标签: {label}")
    print(f"{'='*60}")

    if not os.path.exists(ppt_path):
        print(f"文件不存在！")
        return None

    ole = olefile.OleFileIO(ppt_path)
    print(f"\nOLE2 streams:")
    for stream in ole.listdir():
        path = '/'.join(stream)
        size = ole.get_size(path)
        print(f"  {path} ({size} bytes)")

    # 检查 Pictures stream
    if ole.exists('Pictures'):
        pics_data = ole.openstream('Pictures').read()
        print(f"\nPictures stream: {len(pics_data)} bytes")

        # 解析 record 结构
        offset = 0
        record_count = 0
        while offset + 8 <= len(pics_data):
            ver_inst = int.from_bytes(pics_data[offset:offset+2], 'little')
            rec_type = int.from_bytes(pics_data[offset+2:offset+4], 'little')
            rec_len = int.from_bytes(pics_data[offset+4:offset+8], 'little')
            ver = ver_inst & 0x0F
            inst = (ver_inst >> 4) & 0x0FFF

            print(f"  Record #{record_count} at offset {offset}:")
            print(f"    ver={ver:#x}, inst={inst:#x}({inst}), type={rec_type:#06x}, len={rec_len}")

            # 如果是加密的，数据应该是随机字节，无法解析
            # 如果是未加密的，可以尝试解析 FBSE 结构
            if rec_type == 0xF007:  # FBSE
                # BLIB_STORE_ENTRY_PARTS = [1,1,16,2,4,4,4,1,1,1,1] = 36 bytes
                if offset + 8 + 36 <= len(pics_data):
                    bt_win_blip_type = pics_data[offset+8]
                    bt_mac_blip_type = pics_data[offset+9]
                    rgb_uid = pics_data[offset+10:offset+26]
                    tag = int.from_bytes(pics_data[offset+26:offset+28], 'little')
                    size = int.from_bytes(pics_data[offset+28:offset+32], 'little')
                    ref_count = int.from_bytes(pics_data[offset+32:offset+36], 'little')
                    delay_stream_offset = int.from_bytes(pics_data[offset+36:offset+40], 'little')
                    usage = pics_data[offset+40]
                    cb_name = int.from_bytes(pics_data[offset+41:offset+43], 'little')
                    print(f"    FBSE: btWinBlipType={bt_win_blip_type:#x}, rgbUid={rgb_uid.hex()}")
                    print(f"    FBSE: tag={tag}, size={size}, refCount={ref_count}, delayOffset={delay_stream_offset}")
                    print(f"    FBSE: usage={usage}, cbName={cb_name}")

            record_count += 1
            offset += 8 + rec_len
            if record_count > 20:  # 限制输出
                print(f"  ... (更多记录省略)")
                break

        return pics_data
    else:
        print(f"\n无 Pictures stream")
        return None


# 检查 msoffcrypto 测试文件
check_ppt_structure(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\rc4cryptoapi_password.ppt",
    "msoffcrypto 测试文件（已加密）"
)

# 检查我们生成的加密文件
check_ppt_structure(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt",
    "我们生成的加密文件"
)

# 检查原始未加密文件
check_ppt_structure(
    r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt",
    "原始未加密文件"
)
