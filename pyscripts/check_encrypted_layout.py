#!/usr/bin/env python3
"""检查加密后文件中 persist 对象的布局，确认重排是否生效。"""
import struct
import olefile

def check_encrypted_layout(filepath):
    print(f"\n{'='*70}")
    print(f"检查: {filepath}")
    print(f"{'='*70}")

    ole = olefile.OleFileIO(filepath)
    ppt = ole.openstream("PowerPoint Document").read()
    cu = ole.openstream("Current User").read()

    # CurrentUserAtom
    header_token = struct.unpack_from("<I", cu, 12)[0]
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    print(f"headerToken = 0x{header_token:08X} (encrypted={header_token == 0xF3D1C4DF})")
    print(f"offsetToCurrentEdit = {offset_to_current_edit}")

    # UserEditAtom
    ue_off = offset_to_current_edit
    ue_rec_type = struct.unpack_from("<H", ppt, ue_off + 2)[0]
    ue_rec_len = struct.unpack_from("<I", ppt, ue_off + 4)[0]
    offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
    encrypt_session_pid = struct.unpack_from("<I", ppt, ue_off + 32)[0] if ue_rec_len >= 32 else None
    print(f"UserEditAtom: recType=0x{ue_rec_type:04X}, recLen={ue_rec_len}, offsetPersistDirectory={offset_persist_dir}")
    print(f"  encryptSessionPersistIdRef = {encrypt_session_pid}")

    # PersistDirectoryAtom
    pd_off = offset_persist_dir
    pd_rec_type = struct.unpack_from("<H", ppt, pd_off + 2)[0]
    pd_rec_len = struct.unpack_from("<I", ppt, pd_off + 4)[0]
    print(f"PersistDirectoryAtom: recType=0x{pd_rec_type:04X}, recLen={pd_rec_len}")

    # PDA entries
    entry_val = struct.unpack_from("<I", ppt, pd_off + 8)[0]
    persist_id = entry_val & 0xFFFFF
    c_persist = (entry_val >> 20) & 0xFFF
    print(f"PDA entry: persistId={persist_id}, cPersist={c_persist}")

    # Read all offsets
    offsets = []
    for i in range(c_persist):
        off = struct.unpack_from("<I", ppt, pd_off + 12 + i * 4)[0]
        offsets.append(off)

    print(f"\nPersist object offsets (persistId order), first 10:")
    for i in range(min(10, len(offsets))):
        print(f"  pid {persist_id + i}: offset {offsets[i]}")

    print(f"  ...")
    print(f"  pid {persist_id + len(offsets) - 1}: offset {offsets[-1]}")

    # Check if sorted by offset
    is_sorted = all(offsets[i] < offsets[i + 1] for i in range(len(offsets) - 1))
    print(f"\nOffsets in persistId order are ascending? {is_sorted}")

    if not is_sorted:
        # Find first out-of-order pair
        for i in range(len(offsets) - 1):
            if offsets[i] >= offsets[i + 1]:
                print(f"  First out-of-order: pid {persist_id + i} (offset {offsets[i]}) >= pid {persist_id + i + 1} (offset {offsets[i + 1]})")
                break

    # Check stream size
    print(f"\nStream length = {len(ppt)}")
    print(f"Last offset = {offsets[-1]}")

    # Check what's at the last offset (should be CryptSession10Container)
    last_off = offsets[-1]
    if last_off + 8 <= len(ppt):
        last_rec_type = struct.unpack_from("<H", ppt, last_off + 2)[0]
        last_rec_len = struct.unpack_from("<I", ppt, last_off + 4)[0]
        print(f"Last persist object: recType=0x{last_rec_type:04X}, recLen={last_rec_len}")
        if last_rec_type == 0x2F14:
            print(f"  ✓ CryptSession10Container found at last offset")

check_encrypted_layout("_test_out/protected_心理账户理论.ppt")
check_encrypted_layout("_test_out/wm_protected_心理账户理论.ppt")
