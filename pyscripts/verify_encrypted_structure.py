#!/usr/bin/env python3
"""验证加密后的 .ppt 文件结构，检查密码通过后报错的根因。

用 msoffcrypto 解密文件，检查解密后的结构是否完整。
"""

import sys
import struct
from pathlib import Path

try:
    import olefile
except ImportError:
    print("需要 olefile: pip install olefile")
    sys.exit(1)

try:
    from msoffcrypto import OfficeFile
except ImportError:
    print("需要 msoffcrypto: pip install msoffcrypto-tool")
    sys.exit(1)


def parse_header(data, offset):
    """解析 8 字节 record header。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


RECORD_TYPE_NAMES = {
    0x03E8: "Document",
    0x03EE: "Slide",
    0x03F8: "MainMaster",
    0x040C: "PPDrawing",
    0x0FF5: "UserEditAtom",
    0x0FF6: "CurrentUserAtom",
    0x1772: "PersistDirectoryAtom",
    0x2F14: "CryptSession10Container",
}


def type_name(t):
    return RECORD_TYPE_NAMES.get(t, f"0x{t:04X}")


def dump_top_records(data, label, max_records=20):
    """打印 stream 的顶层 record。"""
    print(f"\n--- {label} 顶层 records ---")
    pos = 0
    count = 0
    while pos + 8 <= len(data) and count < max_records:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        name = type_name(rec_type)
        print(
            f"  offset={pos:>8} type=0x{rec_type:04X}({name:>25}) len={rec_len:>10} container={'Y' if is_container else 'N'}"
        )
        pos += total_len
        count += 1
        if not is_container and rec_len == 0:
            break
    print(f"  ... 共 {count} 个顶层 record, stream 总长 {len(data)}")


def verify_encryption(path, password):
    """验证加密文件能否被 msoffcrypto 解密。"""
    print(f"\n{'='*70}")
    print(f"验证: {path}")
    print(f"{'='*70}")

    if not Path(path).exists():
        print(f"  文件不存在")
        return False

    # 1. 检查 OLE2 容器
    try:
        ole = olefile.OleFileIO(path)
    except Exception as e:
        print(f"  OLE2 容器打开失败: {e}")
        return False

    print(f"  OLE2 streams: {ole.listdir()}")
    ole.close()

    # 2. 用 msoffcrypto 验证密码并解密
    try:
        with open(path, "rb") as f:
            office_file = OfficeFile(f)
            office_file.load_key(password=password)
            print(f"  密码验证通过 ✓")

            # 解密到临时文件
            out_path = str(Path(path).with_suffix(".decrypted.ppt"))
            with open(out_path, "wb") as out:
                office_file.decrypt(out)
            print(f"  解密成功: {out_path}")

            # 检查解密后的结构
            ole2 = olefile.OleFileIO(out_path)
            if ole2.exists("PowerPoint Document"):
                ppt_data = ole2.openstream("PowerPoint Document").read()
                dump_top_records(ppt_data, "解密后 PowerPoint Document")
            else:
                print(f"  解密后文件中没有 PowerPoint Document stream")
                print(f"  streams: {ole2.listdir()}")
            ole2.close()

            return True
    except Exception as e:
        print(f"  解密失败: {e}")
        import traceback

        traceback.print_exc()
        return False


def check_persist_directory(path):
    """检查加密文件的 PersistDirectoryAtom 和 UserEditAtom 结构。"""
    print(f"\n{'='*70}")
    print(f"检查加密结构: {path}")
    print(f"{'='*70}")

    ole = olefile.OleFileIO(path)

    # 检查 Current User
    if not ole.exists("Current User"):
        print("  找不到 Current User stream")
        ole.close()
        return

    cu_data = ole.openstream("Current User").read()
    h = parse_header(cu_data, 0)
    if h is None or h[2] != 0x0FF6:
        print(f"  CurrentUserAtom 类型错误: 0x{h[2]:04X}" if h else "  解析失败")
        ole.close()
        return

    header_token = struct.unpack_from("<I", cu_data, 12)[0]
    offset_to_current_edit = struct.unpack_from("<I", cu_data, 16)[0]
    print(f"  CurrentUserAtom:")
    print(f"    headerToken = 0x{header_token:08X} ({'已加密' if header_token == 0xF3D1C4DF else '未加密'})")
    print(f"    offsetToCurrentEdit = {offset_to_current_edit}")

    # 检查 PowerPoint Document
    if not ole.exists("PowerPoint Document"):
        print("  找不到 PowerPoint Document stream")
        ole.close()
        return

    ppt_data = ole.openstream("PowerPoint Document").read()
    print(f"  PowerPoint Document stream 长度: {len(ppt_data)}")

    # 检查 UserEditAtom
    ue_offset = offset_to_current_edit
    if ue_offset + 8 > len(ppt_data):
        print(f"  UserEditAtom offset 超出范围: {ue_offset}")
        ole.close()
        return

    h = parse_header(ppt_data, ue_offset)
    if h is None or h[2] != 0x0FF5:
        print(f"  UserEditAtom 类型错误: 0x{h[2]:04X}" if h else "  解析失败")
        ole.close()
        return

    ver, inst, rec_type, rec_len = h
    print(f"  UserEditAtom (offset={ue_offset}):")
    print(f"    recLen = {rec_len} ({'已加密' if rec_len == 32 else '未加密'})")

    offset_persist_dir = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
    persist_id_seed = struct.unpack_from("<I", ppt_data, ue_offset + 28)[0]
    print(f"    offsetPersistDirectory = {offset_persist_dir}")
    print(f"    persistIdSeed = {persist_id_seed}")

    if rec_len == 32:
        encrypt_session_pid = struct.unpack_from("<I", ppt_data, ue_offset + 8 + 28)[0]
        print(f"    encryptSessionPersistIdRef = {encrypt_session_pid}")

    # 检查 PersistDirectoryAtom
    pd_offset = offset_persist_dir
    if pd_offset + 8 > len(ppt_data):
        print(f"  PersistDirectoryAtom offset 超出范围: {pd_offset}")
        ole.close()
        return

    h = parse_header(ppt_data, pd_offset)
    if h is None or h[2] != 0x1772:
        print(f"  PersistDirectoryAtom 类型错误: 0x{h[2]:04X}" if h else "  解析失败")
        ole.close()
        return

    ver, inst, rec_type, rec_len = h
    print(f"  PersistDirectoryAtom (offset={pd_offset}):")
    print(f"    recLen = {rec_len}")

    # 解析 persist entries
    pd_data = ppt_data[pd_offset + 8 : pd_offset + 8 + rec_len]
    pos = 0
    total_persist = 0
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        print(f"    entry: persistId={persist_id}, cPersist={c_persist}")
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                poff = struct.unpack_from("<I", pd_data, pos)[0]
                pid = persist_id + j
                # 检查 offset 是否指向有效的 record
                if poff + 8 <= len(ppt_data):
                    rh = parse_header(ppt_data, poff)
                    if rh:
                        rtype = rh[2]
                        rlen = rh[3]
                        print(f"      persistId={pid}, offset={poff}, type=0x{rtype:04X}({type_name(rtype)}), len={rlen}")
                    else:
                        print(f"      persistId={pid}, offset={poff}, 解析失败")
                else:
                    print(f"      persistId={pid}, offset={poff}, 超出范围")
                pos += 4
                total_persist += 1

    print(f"    总共 {total_persist} 个 persist 对象")

    # 检查 Pictures stream
    if ole.exists("Pictures"):
        pic_data = ole.openstream("Pictures").read()
        print(f"\n  Pictures stream 长度: {len(pic_data)}")
        # 检查第一个 record header（加密后应该是乱码）
        if len(pic_data) >= 8:
            h = parse_header(pic_data, 0)
            if h:
                ver, inst, rec_type, rec_len = h
                print(f"    第一个 record: type=0x{rec_type:04X}, len={rec_len}")
                if rec_len > 1000000:
                    print(f"    ⚠️ record len 异常大，可能未正确加密或解密")
    else:
        print(f"\n  Pictures stream 不存在")

    ole.close()


def main():
    out_dir = Path("_test_out")

    # 检查加密文件结构
    for ppt in out_dir.glob("protected_*.ppt"):
        check_persist_directory(ppt)

    for ppt in out_dir.glob("wm_protected_*.ppt"):
        check_persist_directory(ppt)

    # 用 msoffcrypto 验证解密
    for ppt in out_dir.glob("protected_*.ppt"):
        verify_encryption(ppt, "pptx-rs-secret")

    for ppt in out_dir.glob("wm_protected_*.ppt"):
        verify_encryption(ppt, "pptx-rs-secret")


if __name__ == "__main__":
    main()
