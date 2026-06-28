#!/usr/bin/env python3
"""检查原始未加密文件的结构。"""
import struct
import olefile

ole = olefile.OleFileIO("_test/心理账户理论.ppt")
cu = ole.openstream("Current User").read()
ppt = ole.openstream("PowerPoint Document").read()

offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
print(f"offsetToCurrentEdit = {offset_to_current_edit}")

ue_off = offset_to_current_edit
ver_inst = struct.unpack_from("<H", ppt, ue_off)[0]
rec_type = struct.unpack_from("<H", ppt, ue_off + 2)[0]
rec_len = struct.unpack_from("<I", ppt, ue_off + 4)[0]
print(f"UserEditAtom: ver_inst=0x{ver_inst:04X}, recType=0x{rec_type:04X}, recLen={rec_len}")

last_slide = struct.unpack_from("<I", ppt, ue_off + 8)[0]
version = struct.unpack_from("<H", ppt, ue_off + 12)[0]
minor = ppt[ue_off + 14]
major = ppt[ue_off + 15]
offset_last_edit = struct.unpack_from("<I", ppt, ue_off + 16)[0]
offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
doc_persist_id = struct.unpack_from("<I", ppt, ue_off + 24)[0]
persist_id_seed = struct.unpack_from("<I", ppt, ue_off + 28)[0]
print(f"  lastSlideIdRef={last_slide}, version={version} (0x{version:04X}), minor={minor}, major={major}")
print(f"  offsetLastEdit={offset_last_edit}, offsetPersistDirectory={offset_persist_dir}")
print(f"  docPersistIdRef={doc_persist_id}, persistIdSeed={persist_id_seed}")

pd_off = offset_persist_dir
pd_rec_len = struct.unpack_from("<I", ppt, pd_off + 4)[0]
entry_val = struct.unpack_from("<I", ppt, pd_off + 8)[0]
persist_id = entry_val & 0xFFFFF
c_persist = (entry_val >> 20) & 0xFFF
print(f"PersistDirectoryAtom: recLen={pd_rec_len}, persistId={persist_id}, cPersist={c_persist}")

ole.close()
