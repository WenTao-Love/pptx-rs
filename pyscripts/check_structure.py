"""全面检查 .ppt 文件的结构，用于诊断加密问题。"""
import struct
import sys
import os
import io
import olefile

RT_CURRENT_USER_ATOM = 0x0FF6
RT_USER_EDIT_ATOM = 0x0FF5
RT_PERSIST_DIRECTORY_ATOM = 0x1772
RT_CRYPT_SESSION10_CONTAINER = 0x2F14

def read_u32_le(data, offset):
    return struct.unpack_from('<I', data, offset)[0]

def read_u16_le(data, offset):
    return struct.unpack_from('<H', data, offset)[0]

def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = read_u16_le(data, offset)
    rec_type = read_u16_le(data, offset + 2)
    rec_len = read_u32_le(data, offset + 4)
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)

def parse_persist_directory(data, offset):
    hdr = parse_record_header(data, offset)
    if not hdr or hdr[2] != RT_PERSIST_DIRECTORY_ATOM:
        print(f'  错误：不是 PersistDirectoryAtom (0x{hdr[2]:04X})')
        return []

    pd_data = data[offset + 8:offset + 8 + hdr[3]]
    entries = []
    pos = 0
    entry_idx = 0
    while pos + 4 <= len(pd_data):
        entry = read_u32_le(pd_data, pos)
        persist_id = entry & 0xFFFFF
        c_persist = (entry >> 20) & 0xFFF
        print(f'  PersistDirectoryEntry[{entry_idx}]: persistId={persist_id}, cPersist={c_persist}')
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = read_u32_le(pd_data, pos)
                entries.append((persist_id + j, persist_offset))
                pos += 4
        entry_idx += 1
    return entries

def check_file(fpath, label):
    print(f'=== {label}: {fpath} ===')
    if not os.path.exists(fpath):
        print(f'  文件不存在')
        return None
    with open(fpath, 'rb') as f:
        data = f.read()
    print(f'  文件大小: {len(data)} bytes')

    # 用 olefile 打开
    ole = olefile.OleFileIO(io.BytesIO(data))

    # 列出所有 streams
    print(f'  OLE2 streams:')
    for entry in ole.listdir():
        path = '/'.join(entry)
        size = ole.get_size(path)
        print(f'    {path} ({size} bytes)')

    # 找到 PowerPoint Document stream
    ppt_data = None
    cu_data = None
    for entry in ole.listdir():
        name = '/'.join(entry)
        if name == 'PowerPoint Document':
            ppt_data = ole.openstream(name).read()
        elif name == 'Current User':
            cu_data = ole.openstream(name).read()

    if not ppt_data:
        print(f'  找不到 PowerPoint Document stream')
        return data
    if not cu_data:
        print(f'  找不到 Current User stream')
        return data

    # 解析 CurrentUserAtom
    cu_hdr = parse_record_header(cu_data, 0)
    if cu_hdr and cu_hdr[2] == RT_CURRENT_USER_ATOM:
        header_token = read_u32_le(cu_data, 12)
        offset_to_current_edit = read_u32_le(cu_data, 16)
        print(f'  CurrentUserAtom: headerToken=0x{header_token:08X}, offsetToCurrentEdit={offset_to_current_edit}')
        if header_token == 0xE391C05F:
            print(f'    状态：未加密')
        elif header_token == 0xF3D1C4DF:
            print(f'    状态：已加密')
        else:
            print(f'    状态：未知 headerToken')

    # 解析 UserEditAtom
    ue_offset = offset_to_current_edit
    ue_hdr = parse_record_header(ppt_data, ue_offset)
    if not ue_hdr or ue_hdr[2] != RT_USER_EDIT_ATOM:
        print(f'  错误：offsetToCurrentEdit 指向的不是 UserEditAtom')
        return data

    print(f'  UserEditAtom (offset={ue_offset}):')
    print(f'    recLen={ue_hdr[3]}', end='')
    if ue_hdr[3] == 28:
        print(' (未加密)')
    elif ue_hdr[3] == 32:
        print(' (已加密)')
    else:
        print(f' (异常)')

    last_slide_id_ref = read_u32_le(ppt_data, ue_offset + 8)
    version = read_u16_le(ppt_data, ue_offset + 12)
    minor_ver = ppt_data[ue_offset + 14]
    major_ver = ppt_data[ue_offset + 15]
    offset_last_edit = read_u32_le(ppt_data, ue_offset + 16)
    offset_persist_dir = read_u32_le(ppt_data, ue_offset + 20)
    doc_persist_id_ref = read_u32_le(ppt_data, ue_offset + 24)
    max_persist_written = read_u32_le(ppt_data, ue_offset + 28)
    print(f'    lastSlideIdRef={last_slide_id_ref}, version={version}, minor={minor_ver}, major={major_ver}')
    print(f'    offsetLastEdit={offset_last_edit}, offsetPersistDirectory={offset_persist_dir}')
    print(f'    docPersistIdRef={doc_persist_id_ref}, maxPersistWritten={max_persist_written}')
    encrypt_session_pid = None
    if ue_hdr[3] == 32:
        encrypt_session_pid = read_u32_le(ppt_data, ue_offset + 36)
        print(f'    encryptSessionPersistIdRef={encrypt_session_pid}')

    # 解析 PersistDirectoryAtom
    print(f'  PersistDirectoryAtom (offset={offset_persist_dir}):')
    pd_hdr = parse_record_header(ppt_data, offset_persist_dir)
    if pd_hdr:
        print(f'    recLen={pd_hdr[3]}')
    entries = parse_persist_directory(ppt_data, offset_persist_dir)
    print(f'    共 {len(entries)} 个 persist entries')

    # 检查每个 persist entry 的 record type
    print(f'  persist entries 详情:')
    for pid, poff in entries:
        if poff >= len(ppt_data):
            print(f'    pid={pid}: offset={poff} 超出范围!')
            continue
        hdr = parse_record_header(ppt_data, poff)
        if hdr:
            print(f'    pid={pid}: offset={poff}, recType=0x{hdr[2]:04X}, recLen={hdr[3]}')
        else:
            print(f'    pid={pid}: offset={poff}, 无法解析 record header')

    # 检查 CryptSession10Container
    if encrypt_session_pid is not None:
        cs_found = False
        for pid, poff in entries:
            if pid == encrypt_session_pid:
                cs_found = True
                cs_hdr = parse_record_header(ppt_data, poff)
                if cs_hdr and cs_hdr[2] == RT_CRYPT_SESSION10_CONTAINER:
                    print(f'  CryptSession10Container (pid={pid}, offset={poff}):')
                    print(f'    recLen={cs_hdr[3]}')
                    # 解析 EncryptionVersionInfo
                    v_major = read_u16_le(ppt_data, poff + 8)
                    v_minor = read_u16_le(ppt_data, poff + 10)
                    outer_flags = read_u32_le(ppt_data, poff + 12)
                    header_size = read_u32_le(ppt_data, poff + 16)
                    print(f'    vMajor={v_major}, vMinor={v_minor}, flags=0x{outer_flags:08X}, headerSize={header_size}')
                else:
                    print(f'  错误：encryptSessionPersistIdRef 指向的不是 CryptSession10Container (recType=0x{hdr[2]:04X})')
                break
        if not cs_found:
            print(f'  错误：找不到 encryptSessionPersistIdRef={encrypt_session_pid} 对应的 persist entry')

    return data

if __name__ == '__main__':
    files = [
        ('_test/心理账户理论.ppt', '原始文件'),
        ('_test_out/protected_心理账户理论.ppt', '纯加密'),
        ('_test_out/wm_protected_心理账户理论.ppt', '水印+加密'),
        ('_test_out/wm_心理账户理论.ppt', '纯水印'),
    ]
    for fpath, label in files:
        check_file(fpath, label)
        print()
