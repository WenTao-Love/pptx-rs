#!/usr/bin/env python3
"""检查加密文件的实际结构（不解密）。"""
import struct
import olefile

def inspect(filepath):
    print(f"\n检查: {filepath}")
    ole = olefile.OleFileIO(filepath)
    cu = ole.openstream("Current User").read()
    ppt = ole.openstream("PowerPoint Document").read()

    header_token = struct.unpack_from("<I", cu, 12)[0]
    offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
    print(f"  Current User: headerToken=0x{header_token:08X}, offsetToCurrentEdit={offset_to_current_edit}")

    ue_off = offset_to_current_edit
    if ue_off + 8 <= len(ppt):
        ver_inst = struct.unpack_from("<H", ppt, ue_off)[0]
        rec_type = struct.unpack_from("<H", ppt, ue_off + 2)[0]
        rec_len = struct.unpack_from("<I", ppt, ue_off + 4)[0]
        print(f"  UserEditAtom @ {ue_off}: ver_inst=0x{ver_inst:04X}, recType=0x{rec_type:04X}, recLen={rec_len}")
        if rec_len >= 28:
            last_slide = struct.unpack_from("<I", ppt, ue_off + 8)[0]
            version = struct.unpack_from("<H", ppt, ue_off + 12)[0]
            minor_major = struct.unpack_from("<BB", ppt, ue_off + 14)
            offset_last_edit = struct.unpack_from("<I", ppt, ue_off + 16)[0]
            offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
            doc_persist_id = struct.unpack_from("<I", ppt, ue_off + 24)[0]
            persist_id_seed = struct.unpack_from("<I", ppt, ue_off + 28)[0]
            print(f"    lastSlideIdRef={last_slide}, version={version}, minor/major={minor_major}")
            print(f"    offsetLastEdit={offset_last_edit}, offsetPersistDirectory={offset_persist_dir}")
            print(f"    docPersistIdRef={doc_persist_id}, persistIdSeed={persist_id_seed}")
            if rec_len == 32:
                encrypt_session = struct.unpack_from("<I", ppt, ue_off + 36)[0]
                print(f"    encryptSessionPersistIdRef={encrypt_session}")

    # 找到 PersistDirectoryAtom
    pd_off = offset_persist_dir
    if pd_off + 8 <= len(ppt):
        pd_ver_inst = struct.unpack_from("<H", ppt, pd_off)[0]
        pd_rec_type = struct.unpack_from("<H", ppt, pd_off + 2)[0]
        pd_rec_len = struct.unpack_from("<I", ppt, pd_off + 4)[0]
        print(f"  PersistDirectoryAtom @ {pd_off}: ver_inst=0x{pd_ver_inst:04X}, recType=0x{pd_rec_type:04X}, recLen={pd_rec_len}")
        entry_val = struct.unpack_from("<I", ppt, pd_off + 8)[0]
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        print(f"    entry: persistId={persist_id}, cPersist={c_persist}")
        for i in range(min(c_persist, 5)):
            off = struct.unpack_from("<I", ppt, pd_off + 12 + i * 4)[0]
            print(f"    pid {persist_id + i}: offset={off}")

    # 检查 stream 末尾附近是否有 CryptSession10Container
    # 从 offsetPersistDirectory 向后搜索 0x2F14
    pos = offset_persist_dir
    while pos + 8 <= len(ppt):
        rec_type = struct.unpack_from("<H", ppt, pos + 2)[0]
        if rec_type == 0x2F14:
            rec_len = struct.unpack_from("<I", ppt, pos + 4)[0]
            print(f"  CryptSession10Container @ {pos}: recLen={rec_len}")
            break
        rec_len = struct.unpack_from("<I", ppt, pos + 4)[0]
        pos += 8 + rec_len

    ole.close()

inspect("_test_out/protected_心理账户理论.ppt")
inspect("_test_out/wm_protected_心理账户理论.ppt")
inspect("_test_out/rc4cryptoapi_password.ppt")
