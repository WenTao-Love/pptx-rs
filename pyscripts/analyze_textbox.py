# -*- coding: utf-8 -*-
"""查看 ClientTextbox 的原始数据。"""
import struct
import sys
import io
import olefile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def main():
    ppt_path = "_test/心理账户理论.ppt"
    ole = olefile.OleFileIO(ppt_path)
    ppt_data = ole.openstream("PowerPoint Document").read()

    # ClientTextbox at offset=10211, len=80
    print("=== ClientTextbox at offset=10211, len=80 ===")
    ct_start = 10211
    # 读取 record header
    ver_inst = struct.unpack_from('<H', ppt_data, ct_start)[0]
    rec_type = struct.unpack_from('<H', ppt_data, ct_start + 2)[0]
    rec_len = struct.unpack_from('<I', ppt_data, ct_start + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    print(f"ClientTextbox: ver={ver}, inst={inst}, type=0x{rec_type:04X}, len={rec_len}")

    # 遍历子 record
    pos = ct_start + 8
    end = ct_start + 8 + rec_len
    while pos + 8 <= end:
        vi = struct.unpack_from('<H', ppt_data, pos)[0]
        rt = struct.unpack_from('<H', ppt_data, pos + 2)[0]
        rl = struct.unpack_from('<I', ppt_data, pos + 4)[0]
        v = vi & 0x0F
        i = (vi >> 4) & 0x0FFF
        print(f"  子 record at offset={pos}: ver={v}, inst={i}, type=0x{rt:04X}, len={rl}")
        if rl > 0:
            raw = ppt_data[pos+8:pos+8+min(60, rl)]
            print(f"    原始数据: {raw.hex()}")
            # 尝试解码为 UTF-16LE
            if rl >= 2:
                try:
                    text = raw.decode('utf-16-le', errors='replace').rstrip('\x00')
                    print(f"    UTF-16LE 文本: {text!r}")
                except:
                    pass
            # 尝试解码为 ASCII
            try:
                text = raw.decode('ascii', errors='replace').rstrip('\x00')
                print(f"    ASCII 文本: {text!r}")
            except:
                pass
        pos += 8 + rl

    ole.close()

if __name__ == "__main__":
    main()
