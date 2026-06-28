#!/usr/bin/env python3
"""诊断 persist 对象布局，确认 UserEditAtom/PDA 是否在所有 persist 对象之后。"""
import struct
import olefile

ole = olefile.OleFileIO("_test/心理账户理论.ppt")
ppt = ole.openstream("PowerPoint Document").read()
cu = ole.openstream("Current User").read()

# CurrentUserAtom
offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
print(f"offsetToCurrentEdit (UserEditAtom offset) = {offset_to_current_edit}")

# UserEditAtom
ue_off = offset_to_current_edit
ue_rec_len = struct.unpack_from("<I", ppt, ue_off + 4)[0]
offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
offset_last_edit = struct.unpack_from("<I", ppt, ue_off + 12)[0]
print(f"UserEditAtom at {ue_off}, recLen={ue_rec_len}, offsetPersistDirectory={offset_persist_dir}, offsetLastEdit={offset_last_edit}")

# PersistDirectoryAtom
pd_off = offset_persist_dir
pd_rec_len = struct.unpack_from("<I", ppt, pd_off + 4)[0]
print(f"PersistDirectoryAtom at {pd_off}, recLen={pd_rec_len}, ends at {pd_off + 8 + pd_rec_len}")

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
print(f"Persist object offsets (persistId order), first 5: {offsets[:5]}")
print(f"Persist object offsets (persistId order), last 3: {offsets[-3:]}")

# Check last persist object end
sorted_offsets = sorted(offsets)
last_off = sorted_offsets[-1]
last_rec_len = struct.unpack_from("<I", ppt, last_off + 4)[0]
last_end = last_off + 8 + last_rec_len
print(f"Last persist object by offset: at {last_off}, recLen={last_rec_len}, ends at {last_end}")

# Check what is at the end of stream
print(f"Stream length = {len(ppt)}")
print(f"UserEditAtom at {ue_off}, PDA at {pd_off}")
print(f"Are UE/PDA after all persist objects? {ue_off > last_end and pd_off > last_end}")

# Check if persist objects are contiguous (no gaps)
sorted_with_len = []
for off in sorted_offsets:
    rl = struct.unpack_from("<I", ppt, off + 4)[0]
    sorted_with_len.append((off, off + 8 + rl))

has_gaps = False
for i in range(len(sorted_with_len) - 1):
    gap = sorted_with_len[i + 1][0] - sorted_with_len[i][1]
    if gap != 0:
        print(f"GAP between {sorted_with_len[i]} and {sorted_with_len[i+1]}: {gap} bytes")
        has_gaps = True
if not has_gaps:
    print("No gaps between persist objects (contiguous)")

# Check persistId order vs offset order
pid_order = list(range(persist_id, persist_id + c_persist))
offset_order_by_pid = offsets
print(f"\nFirst 10 (persistId, offset) pairs:")
for i in range(min(10, len(offsets))):
    print(f"  pid {pid_order[i]}: offset {offsets[i]}")

# Check if sorted by offset
is_sorted = all(offsets[i] < offsets[i + 1] for i in range(len(offsets) - 1))
print(f"\nOffsets in persistId order are ascending? {is_sorted}")
