"""检查 msoffcrypto 解密后的 PowerPoint Document stream 结构。

从 Current User stream 获取 offsetToCurrentEdit，然后分析 persist 对象。
"""
import sys
import io
import os
import struct

try:
    import msoffcrypto
except ImportError:
    print("请先安装 msoffcrypto-tool: pip install msoffcrypto-tool")
    sys.exit(1)

try:
    import olefile
except ImportError:
    print("请先安装 olefile: pip install olefile")
    sys.exit(1)


def parse_record_header(data, offset):
    """解析 8 字节 record header。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def parse_persist_directory(data, offset):
    """解析 PersistDirectoryAtom。"""
    rh = parse_record_header(data, offset)
    if rh is None or rh[2] != 0x1772:
        return None

    rec_len = rh[3]
    pd_data = data[offset + 8 : offset + 8 + rec_len]
    entries = []
    pos = 0

    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4

        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + j, persist_offset))
                pos += 4

    return entries


def get_offset_to_current_edit(cu_data):
    """从 Current User stream 获取 offsetToCurrentEdit。"""
    # CurrentUserAtom 结构：
    # RecordHeader (8 bytes)
    # size (4 bytes)
    # headerToken (4 bytes) - offset 12
    # offsetToCurrentEdit (4 bytes) - offset 16
    return struct.unpack_from("<I", cu_data, 16)[0]


def analyze_ppt_stream(ppt_data, cu_data, label=""):
    """分析 PowerPoint Document stream 的结构。"""
    print(f"\n  --- {label} ---")
    print(f"  PPT stream 大小: {len(ppt_data)} bytes")
    print(f"  CU stream 大小: {len(cu_data)} bytes")

    # 从 Current User stream 获取 offsetToCurrentEdit
    offset_to_current_edit = get_offset_to_current_edit(cu_data)
    print(f"  offsetToCurrentEdit: {offset_to_current_edit}")

    # 解析 UserEditAtom
    ue_offset = offset_to_current_edit
    rh = parse_record_header(ppt_data, ue_offset)
    if rh is None or rh[2] != 0x0FF5:
        print(f"  ✗ 在 offset {ue_offset} 找不到 UserEditAtom")
        print(f"    实际找到: {rh}")
        return None

    ue_len = rh[3]
    print(f"  UserEditAtom: offset={ue_offset}, recLen={ue_len}")

    offset_persist_dir = struct.unpack_from("<I", ppt_data, ue_offset + 20)[0]
    persist_id_seed = struct.unpack_from("<I", ppt_data, ue_offset + 28)[0]

    if ue_len == 32:
        encrypt_session_pid = struct.unpack_from("<I", ppt_data, ue_offset + 36)[0]
        print(f"  encryptSessionPersistIdRef: {encrypt_session_pid}")

    print(f"  offsetPersistDirectory: {offset_persist_dir}")
    print(f"  persistIdSeed: {persist_id_seed}")

    # 解析 PersistDirectoryAtom
    persist_entries = parse_persist_directory(ppt_data, offset_persist_dir)
    if persist_entries is None:
        print("  ✗ 找不到 PersistDirectoryAtom")
        return None

    print(f"\n  Persist 目录 ({len(persist_entries)} 个 entry):")
    print(f"  {'pid':>5} {'offset':>10} {'type':>10} {'recLen':>10} {'ver':>4} {'状态':>10}")
    print(f"  {'-'*5} {'-'*10} {'-'*10} {'-'*10} {'-'*4} {'-'*10}")

    for pid, poff in persist_entries:
        poff = int(poff)
        if poff + 8 > len(ppt_data):
            print(f"  {pid:>5} {poff:>10} {'N/A':>10} {'N/A':>10} {'N/A':>4} {'越界':>10}")
            continue

        rh = parse_record_header(ppt_data, poff)
        if rh is None:
            print(f"  {pid:>5} {poff:>10} {'N/A':>10} {'N/A':>10} {'N/A':>4} {'解析失败':>10}")
            continue

        ver, inst, rec_type, rec_len = rh

        # 检查 record header 是否有效
        status = "OK"
        if ver not in [0, 0xF, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xA, 0xB, 0xC, 0xD, 0xE]:
            status = "ver异常"
        if rec_len > 1000000:
            status = "len异常"

        print(f"  {pid:>5} {poff:>10} 0x{rec_type:04X}    {rec_len:>10} 0x{ver:X}   {status:>10}")

    return persist_entries


def main():
    encrypted_file = "_test_out/protected_心理账户理论.ppt"
    original_file = "_test/心理账户理论.ppt"
    password = "pptx-rs-secret"

    if not os.path.exists(original_file):
        for f in os.listdir("_test"):
            if f.endswith(".ppt"):
                original_file = f"_test/{f}"
                break

    print(f"加密文件: {encrypted_file}")
    print(f"原始文件: {original_file}")

    # 解密
    with open(encrypted_file, "rb") as f:
        officefile = msoffcrypto.OfficeFile(f)
        officefile.load_key(password=password)
        out = io.BytesIO()
        officefile.decrypt(out)
        decrypted_data = out.getvalue()

    # 提取 streams
    dec_ole = olefile.OleFileIO(io.BytesIO(decrypted_data))
    dec_ppt = dec_ole.openstream("PowerPoint Document").read()
    dec_cu = dec_ole.openstream("Current User").read()

    with open(original_file, "rb") as f:
        original_data = f.read()
    orig_ole = olefile.OleFileIO(io.BytesIO(original_data))
    orig_ppt = orig_ole.openstream("PowerPoint Document").read()
    orig_cu = orig_ole.openstream("Current User").read()

    # 分析两个 stream
    orig_entries = analyze_ppt_stream(orig_ppt, orig_cu, "原始文件")
    dec_entries = analyze_ppt_stream(dec_ppt, dec_cu, "解密后文件")

    # 比较 persist 对象
    if orig_entries and dec_entries:
        print(f"\n  --- 比较 persist 对象 ---")
        print(f"  {'pid':>5} {'orig offset':>12} {'dec offset':>12} {'orig type':>10} {'dec type':>10} {'状态':>10}")
        print(f"  {'-'*5} {'-'*12} {'-'*12} {'-'*10} {'-'*10} {'-'*10}")

        # 按 persistId 比较
        orig_dict = {pid: poff for pid, poff in orig_entries}
        dec_dict = {pid: poff for pid, poff in dec_entries}

        all_pids = sorted(set(orig_dict.keys()) | set(dec_dict.keys()))
        for pid in all_pids:
            orig_poff = orig_dict.get(pid)
            dec_poff = dec_dict.get(pid)

            orig_type = "N/A"
            dec_type = "N/A"

            if orig_poff is not None:
                orig_poff = int(orig_poff)
                rh = parse_record_header(orig_ppt, orig_poff)
                if rh:
                    orig_type = f"0x{rh[2]:04X}"

            if dec_poff is not None:
                dec_poff = int(dec_poff)
                rh = parse_record_header(dec_ppt, dec_poff)
                if rh:
                    dec_type = f"0x{rh[2]:04X}"

            status = "✓ 相同" if orig_type == dec_type else "✗ 不同"
            if orig_poff is None:
                status = "解密新增"
            elif dec_poff is None:
                status = "解密缺失"

            print(f"  {pid:>5} {str(orig_poff):>12} {str(dec_poff):>12} {orig_type:>10} {dec_type:>10} {status:>10}")

    dec_ole.close()
    orig_ole.close()


if __name__ == "__main__":
    main()
