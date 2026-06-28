#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析加密文件解密后的结构问题：
1. 检查 Pictures stream 是否被加密
2. 检查 persist 对象的 offset 是否正确
3. 对比原始文件和解密后文件的结构
"""
import io
import os
import struct
import olefile
import msoffcrypto

PASSWORD = "pptx-rs-secret"


def list_streams(filepath):
    """列出 OLE 文件中的所有 stream。"""
    ole = olefile.OleFileIO(filepath)
    try:
        streams = []
        for parts in ole.listdir():
            name = "/".join(parts)
            size = ole.get_size(name)
            streams.append((name, size))
        return streams
    finally:
        ole.close()


def parse_record_header(data, pos):
    if pos + 8 > len(data):
        return None
    ver_inst, rec_type, rec_len = struct.unpack_from("<HHI", data, pos)
    ver = ver_inst & 0xF
    inst = (ver_inst >> 4) & 0xFFF
    return ver, inst, rec_type, rec_len


def analyze_ppt_stream(data, label=""):
    """分析 PowerPoint Document stream 的顶层 record 结构。"""
    print(f"\n  --- {label} PowerPoint Document 顶层 records ---", flush=True)
    pos = 0
    records = []
    while pos + 8 <= len(data):
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len = hdr
        is_container = ver == 0xF
        total_len = 8 + rec_len
        rt_names = {
            0x03E8: "Document", 0x03EE: "Slide", 0x03F0: "Notes",
            0x03F8: "MainMaster", 0x0FF0: "SlideListWithText",
            0x0FF5: "UserEditAtom", 0x0FF6: "CurrentUserAtom",
            0x1772: "PersistDirectoryAtom", 0x2F14: "CryptSession10",
        }
        name = rt_names.get(rec_type, f"0x{rec_type:04X}")
        records.append((pos, rec_type, rec_len, name, is_container))
        if len(records) <= 5 or rec_type in (0x0FF5, 0x1772, 0x2F14):
            print(f"    offset={pos:>8} type=0x{rec_type:04X}({name:>20}) len={rec_len:>8} container={'Y' if is_container else 'N'}", flush=True)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    print(f"    ... 共 {len(records)} 个顶层 record, stream 总长 {len(data)}", flush=True)
    # 打印最后几个 record
    for r in records[-5:]:
        print(f"    offset={r[0]:>8} type=0x{r[1]:04X}({r[3]:>20}) len={r[2]:>8}", flush=True)


def check_pictures_stream(filepath):
    """检查 Pictures stream 是否存在及其内容。"""
    ole = olefile.OleFileIO(filepath)
    try:
        stream_names = ["/".join(parts) for parts in ole.listdir()]
        print(f"  streams: {stream_names}", flush=True)
        if "Pictures" in stream_names:
            with ole.openstream("Pictures") as f:
                pic_data = f.read()
            print(f"  Pictures stream 大小: {len(pic_data)} 字节", flush=True)
            # 检查第一个字节是否是有效的 record header
            if len(pic_data) >= 8:
                ver, inst, rec_type, rec_len = parse_record_header(pic_data, 0)
                print(f"  Pictures 第一个 record: ver={ver}, inst={inst}, type=0x{rec_type:04X}, len={rec_len}", flush=True)
                # OfficeArtBStoreRecord (0xF018) 是 Pictures stream 的顶层 record
                if rec_type == 0xF018:
                    print(f"  [OK] Pictures stream 顶层是 OfficeArtBStoreRecord (0xF018)", flush=True)
                else:
                    print(f"  [WARN] Pictures stream 顶层不是 0xF018", flush=True)
        else:
            print(f"  [INFO] 无 Pictures stream", flush=True)
    finally:
        ole.close()


def compare_original_and_decrypted(original_path, encrypted_path):
    """对比原始文件和解密后文件的结构。"""
    print(f"\n=== 对比: {os.path.basename(original_path)} ===", flush=True)

    print("\n[原始文件 streams]:", flush=True)
    list_streams(original_path)
    for name, size in list_streams(original_path):
        print(f"  {name}: {size} 字节", flush=True)

    print("\n[加密文件 streams]:", flush=True)
    for name, size in list_streams(encrypted_path):
        print(f"  {name}: {size} 字节", flush=True)

    # 解密加密文件
    print("\n[解密加密文件...]", flush=True)
    with open(encrypted_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=PASSWORD)
        out = io.BytesIO()
        office_file.decrypt(out)
        decrypted_data = out.getvalue()

    tmp_path = encrypted_path + ".decrypted.ppt"
    with open(tmp_path, "wb") as f:
        f.write(decrypted_data)

    try:
        print("\n[解密后文件 streams]:", flush=True)
        for name, size in list_streams(tmp_path):
            print(f"  {name}: {size} 字节", flush=True)

        # 对比 PowerPoint Document stream
        ole_orig = olefile.OleFileIO(original_path)
        ole_dec = olefile.OleFileIO(tmp_path)
        try:
            with ole_orig.openstream("PowerPoint Document") as f:
                orig_ppt = f.read()
            with ole_dec.openstream("PowerPoint Document") as f:
                dec_ppt = f.read()

            print(f"\n  原始 PPT stream 大小: {len(orig_ppt)}", flush=True)
            print(f"  解密后 PPT stream 大小: {len(dec_ppt)}", flush=True)

            analyze_ppt_stream(orig_ppt, "原始")
            analyze_ppt_stream(dec_ppt, "解密后")

            # 检查 Pictures stream
            print("\n  [原始文件 Pictures]:", flush=True)
            check_pictures_stream(original_path)
            print("\n  [解密后文件 Pictures]:", flush=True)
            check_pictures_stream(tmp_path)
        finally:
            ole_orig.close()
            ole_dec.close()
    finally:
        os.remove(tmp_path)


if __name__ == "__main__":
    test_dir = "_test"
    out_dir = "_test_out"
    # 找到原始 .ppt 文件
    for fname in sorted(os.listdir(test_dir)):
        if fname.endswith(".ppt"):
            orig = os.path.join(test_dir, fname)
            prot = os.path.join(out_dir, "protected_" + fname)
            if os.path.exists(prot):
                compare_original_and_decrypted(orig, prot)
                break
