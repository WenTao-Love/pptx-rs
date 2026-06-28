"""检查 UserEditAtom 的所有字段，特别是 encryptSessionPersistIdRef。"""
import olefile
import struct

def check_user_edit_atom(path):
    """检查 UserEditAtom 的所有字段。"""
    print(f"\n=== {path} ===")
    ole = olefile.OleFileIO(path)

    cu = ole.openstream("Current User").read()
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    print(f"  offsetToCurrentEdit: {offset_to_current_edit}")

    ppt = ole.openstream("PowerPoint Document").read()
    ue_offset = offset_to_current_edit

    # UserEditAtom header
    ver_inst = struct.unpack_from("<H", ppt, ue_offset)[0]
    rec_type = struct.unpack_from("<H", ppt, ue_offset + 2)[0]
    rec_len = struct.unpack_from("<I", ppt, ue_offset + 4)[0]
    print(f"  UserEditAtom: ver={ver_inst & 0x0F}, inst={ver_inst >> 4}, type=0x{rec_type:04X}, rec_len={rec_len}")

    # UserEditAtom fields
    print(f"    lastSlideIdRef: {struct.unpack_from('<I', ppt, ue_offset + 8)[0]}")
    print(f"    version: {struct.unpack_from('<I', ppt, ue_offset + 12)[0]}")
    print(f"    minorVersion: {struct.unpack_from('<I', ppt, ue_offset + 16)[0]}")
    print(f"    offsetPersistDirectory: {struct.unpack_from('<I', ppt, ue_offset + 20)[0]}")
    print(f"    persistIdSeed: {struct.unpack_from('<I', ppt, ue_offset + 24)[0]}")
    print(f"    maxPersistWritten: {struct.unpack_from('<I', ppt, ue_offset + 28)[0]}")
    print(f"    cLastSavePoint: {struct.unpack_from('<I', ppt, ue_offset + 32)[0]}")
    if rec_len >= 32:
        print(f"    encryptSessionPersistIdRef: {struct.unpack_from('<I', ppt, ue_offset + 36)[0]}")

    # 也检查原始文件的 UserEditAtom
    ole.close()

# 检查所有文件
files = [
    "_test/心理账户理论.ppt",
    "_test_out/protected_心理账户理论.ppt",
    "_test_out/wm_protected_心理账户理论.ppt",
]

for f in files:
    try:
        check_user_edit_atom(f)
    except Exception as e:
        print(f"  分析失败: {e}")
