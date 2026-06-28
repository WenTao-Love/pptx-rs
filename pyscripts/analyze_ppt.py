"""分析加密 .ppt 文件的 PowerPoint Document stream 内部结构。"""
import olefile
import struct

def parse_record_header(data, offset):
    """解析 8 字节 record header。"""
    if offset + 8 > len(data):
        raise ValueError(f"offset {offset} 超出范围 {len(data)}")
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len

def analyze_ppt_stream(path):
    """分析 PowerPoint Document stream。"""
    print(f"\n=== 分析 {path} ===")
    ole = olefile.OleFileIO(path)

    # 读取 Current User
    cu = ole.openstream("Current User").read()
    header_token = struct.unpack_from("<I", cu, 12)[0]
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    print(f"  headerToken: 0x{header_token:08X} ({'ENCRYPTED' if header_token == 0xF3D1C4D0 else 'NOT ENCRYPTED'})")
    print(f"  offsetToCurrentEdit: {offset_to_current_edit}")

    # 读取 PowerPoint Document
    ppt = ole.openstream("PowerPoint Document").read()
    print(f"  PowerPoint Document size: {len(ppt)} bytes")

    # 解析 UserEditAtom
    ue_offset = offset_to_current_edit
    print(f"\n  UserEditAtom at offset {ue_offset}:")
    if ue_offset + 8 > len(ppt):
        print(f"    ERROR: offset out of range")
        ole.close()
        return
    ver, inst, rec_type, rec_len = parse_record_header(ppt, ue_offset)
    print(f"    ver={ver}, inst={inst}, rec_type=0x{rec_type:04X}, rec_len={rec_len}")
    if rec_type != 0x0FF5:
        print(f"    ERROR: expected UserEditAtom (0x0FF5), got 0x{rec_type:04X}")
        # 尝试解密后查看
        ole.close()
        return

    if rec_len >= 28:
        last_slide_id = struct.unpack_from("<I", ppt, ue_offset + 8)[0]
        version = struct.unpack_from("<I", ppt, ue_offset + 12)[0]
        minor_version = struct.unpack_from("<I", ppt, ue_offset + 16)[0]
        offset_persist_dir = struct.unpack_from("<I", ppt, ue_offset + 20)[0]
        persist_id_seed = struct.unpack_from("<I", ppt, ue_offset + 24)[0]
        max_persist_written = struct.unpack_from("<I", ppt, ue_offset + 28)[0]
        print(f"    lastSlideIdRef: {last_slide_id}")
        print(f"    version: {version}")
        print(f"    minorVersion: {minor_version}")
        print(f"    offsetPersistDirectory: {offset_persist_dir}")
        print(f"    persistIdSeed: {persist_id_seed}")
        print(f"    maxPersistWritten: {max_persist_written}")
        if rec_len >= 32:
            encrypt_session_pid = struct.unpack_from("<I", ppt, ue_offset + 32)[0]
            print(f"    encryptSessionPersistIdRef: {encrypt_session_pid}")
        else:
            print(f"    encryptSessionPersistIdRef: (not present, rec_len={rec_len})")

    # 解析 PersistDirectoryAtom
    pd_offset = offset_persist_dir
    print(f"\n  PersistDirectoryAtom at offset {pd_offset}:")
    if pd_offset + 8 > len(ppt):
        print(f"    ERROR: offset out of range")
        ole.close()
        return
    ver, inst, rec_type, rec_len = parse_record_header(ppt, pd_offset)
    print(f"    ver={ver}, inst={inst}, rec_type=0x{rec_type:04X}, rec_len={rec_len}")
    if rec_type != 0x1772:
        print(f"    ERROR: expected PersistDirectoryAtom (0x1772), got 0x{rec_type:04X}")
        ole.close()
        return

    # 解析 persist entries
    pd_data = ppt[pd_offset + 8:pd_offset + 8 + rec_len]
    pos = 0
    entries = []
    while pos + 4 <= len(pd_data):
        entry = struct.unpack_from("<I", pd_data, pos)[0]
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        pos += 4
        print(f"    PersistDirectoryEntry: persistId={persist_id}, cPersist={c_persist}")
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                poff = struct.unpack_from("<I", pd_data, pos)[0]
                entries.append((persist_id + j, poff))
                print(f"      [{persist_id + j}] offset={poff}")
                pos += 4

    # 检查 CryptSession10Container（最后一个 persist 对象）
    if entries:
        last_pid, last_offset = entries[-1]
        print(f"\n  Last persist object (pid={last_pid}) at offset {last_offset}:")
        if last_offset + 8 <= len(ppt):
            ver, inst, rec_type, rec_len = parse_record_header(ppt, last_offset)
            print(f"    ver={ver}, inst={inst}, rec_type=0x{rec_type:04X}, rec_len={rec_len}")
            if rec_type == 0x2F14:
                print(f"    -> CryptSession10Container (correct)")
                # 解析 EncryptionVersionInfo
                data_offset = last_offset + 8
                if data_offset + 4 <= len(ppt):
                    v_major = struct.unpack_from("<H", ppt, data_offset)[0]
                    v_minor = struct.unpack_from("<H", ppt, data_offset + 2)[0]
                    print(f"    EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}")
                if data_offset + 8 <= len(ppt):
                    outer_flags = struct.unpack_from("<I", ppt, data_offset + 4)[0]
                    print(f"    outer flags: 0x{outer_flags:08X}")
                if data_offset + 12 <= len(ppt):
                    header_size = struct.unpack_from("<I", ppt, data_offset + 8)[0]
                    print(f"    headerSize: {header_size}")
                    # EncryptionHeader
                    eh_offset = data_offset + 12
                    if eh_offset + 24 <= len(ppt):
                        eh_flags = struct.unpack_from("<I", ppt, eh_offset)[0]
                        alg_id = struct.unpack_from("<I", ppt, eh_offset + 8)[0]
                        alg_id_hash = struct.unpack_from("<I", ppt, eh_offset + 12)[0]
                        key_size = struct.unpack_from("<I", ppt, eh_offset + 16)[0]
                        print(f"    EncryptionHeader: flags=0x{eh_flags:08X}, algId=0x{alg_id:08X}, algIdHash=0x{alg_id_hash:08X}, keySize={key_size}")
            else:
                print(f"    -> NOT CryptSession10Container (expected 0x2F14)")

    ole.close()


# 分析加密文件
files = [
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        analyze_ppt_stream(f)
    except Exception as e:
        print(f"  分析失败: {e}")
        import traceback
        traceback.print_exc()
