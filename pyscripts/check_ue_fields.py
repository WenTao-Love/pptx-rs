#!/usr/bin/env python3
"""检查 UserEditAtom 的所有字段。"""
import struct
import olefile

ole = olefile.OleFileIO("_test/心理账户理论.ppt")
ppt = ole.openstream("PowerPoint Document").read()
cu = ole.openstream("Current User").read()
ue_off = struct.unpack_from("<I", cu, 16)[0]

# UserEditAtom fields (after 8-byte header)
last_slide_id = struct.unpack_from("<I", ppt, ue_off + 8)[0]
version = struct.unpack_from("<H", ppt, ue_off + 12)[0]
minor_major = struct.unpack_from("<BB", ppt, ue_off + 14)
offset_last_edit = struct.unpack_from("<I", ppt, ue_off + 16)[0]
offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
doc_persist_id = struct.unpack_from("<I", ppt, ue_off + 24)[0]
persist_id_seed = struct.unpack_from("<I", ppt, ue_off + 28)[0]

print(f"ue_off={ue_off}")
print(f"lastSlideIdRef={last_slide_id}")
print(f"version={version}")
print(f"minorVersion={minor_major[0]}, majorVersion={minor_major[1]}")
print(f"offsetLastEdit={offset_last_edit} (hex={hex(offset_last_edit)})")
print(f"offsetPersistDirectory={offset_persist_dir}")
print(f"docPersistIdRef={doc_persist_id}")
print(f"persistIdSeed={persist_id_seed}")
