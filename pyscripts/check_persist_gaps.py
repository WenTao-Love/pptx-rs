"""
检查 persist 对象是否按 offset 顺序连续排列，没有间隙。
PowerPoint 可能用"下一个 offset - 当前 offset"来确定 persist 对象的加密长度。
"""
import olefile
import struct

def read_u32_le(data, off):
    return struct.unpack_from('<I', data, off)[0]

def read_u16_le(data, off):
    return struct.unpack_from('<H', data, off)[0]

def parse_record_header(data, off):
    ver_inst = read_u16_le(data, off)
    rec_type = read_u16_le(data, off + 2)
    rec_len = read_u32_le(data, off + 4)
    return ver_inst & 0xF, (ver_inst >> 4) & 0xFFF, rec_type, rec_len

def parse_persist_directory(data, off):
    _, _, rec_type, rec_len = parse_record_header(data, off)
    assert rec_type == 0x1772
    entries = []
    pos = off + 8
    end = pos + rec_len
    while pos < end:
        entry_val = read_u32_le(data, pos)
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        offsets = []
        for i in range(c_persist):
            offsets.append(read_u32_le(data, pos + 4 + i * 4))
        entries.append((persist_id, c_persist, offsets))
        pos += 4 + c_persist * 4
    return entries

# 读取原始文件
o = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt')
orig_ppt = o.openstream('PowerPoint Document').read()
orig_cu = o.openstream('Current User').read()
o.close()

# 获取 UserEditAtom 和 PersistDirectoryAtom 的 offset
ue_off = read_u32_le(orig_cu, 16)
pd_off = read_u32_le(orig_ppt, ue_off + 20)

print(f"UserEditAtom offset: {ue_off}")
print(f"PersistDirectoryAtom offset: {pd_off}")

# 解析 PersistDirectoryAtom
entries = parse_persist_directory(orig_ppt, pd_off)
persist_dir = {}
for pid, cpersist, offsets in entries:
    for i, off in enumerate(offsets):
        persist_dir[pid + i] = off

# 按 offset 排序
sorted_persists = sorted(persist_dir.items(), key=lambda x: x[1])

print(f"\n=== Persist objects sorted by offset ===")
print(f"Total: {len(sorted_persists)} persist objects")

gaps = []
overlaps = []
for i, (pid, offset) in enumerate(sorted_persists):
    _, _, rt, rl = parse_record_header(orig_ppt, offset)
    total_len = 8 + rl
    end = offset + total_len

    if i + 1 < len(sorted_persists):
        next_pid, next_offset = sorted_persists[i + 1]
        gap = next_offset - end
        if gap > 0:
            gaps.append((pid, offset, end, next_offset, gap))
        elif gap < 0:
            overlaps.append((pid, offset, end, next_offset, gap))

    if i < 10 or i >= len(sorted_persists) - 5:
        print(f"  persistId={pid:3d} offset={offset:8d} recType={hex(rt):8s} recLen={rl:8d} end={end:8d}")

print(f"\n=== Gaps between persist objects ===")
print(f"Total gaps: {len(gaps)}")
for pid, offset, end, next_offset, gap in gaps[:10]:
    print(f"  persistId={pid}: offset={offset}, end={end}, next_offset={next_offset}, gap={gap}")

print(f"\n=== Overlaps between persist objects ===")
print(f"Total overlaps: {len(overlaps)}")
for pid, offset, end, next_offset, gap in overlaps[:10]:
    print(f"  persistId={pid}: offset={offset}, end={end}, next_offset={next_offset}, overlap={-gap}")

# 检查最后一个 persist 对象和 PersistDirectoryAtom 之间的间隙
last_pid, last_offset = sorted_persists[-1]
_, _, last_rt, last_rl = parse_record_header(orig_ppt, last_offset)
last_end = last_offset + 8 + last_rl
print(f"\n=== Gap between last persist object and PersistDirectoryAtom ===")
print(f"  last persistId={last_pid}, offset={last_offset}, end={last_end}")
print(f"  PersistDirectoryAtom offset={pd_off}")
print(f"  gap={pd_off - last_end}")

# 检查 PersistDirectoryAtom 和 UserEditAtom 之间的间隙
_, _, pd_rt, pd_rl = parse_record_header(orig_ppt, pd_off)
pd_end = pd_off + 8 + pd_rl
print(f"\n=== Gap between PersistDirectoryAtom and UserEditAtom ===")
print(f"  PersistDirectoryAtom: offset={pd_off}, end={pd_end}")
print(f"  UserEditAtom: offset={ue_off}")
print(f"  gap={ue_off - pd_end}")

# 检查 UserEditAtom 之后的数据
_, _, ue_rt, ue_rl = parse_record_header(orig_ppt, ue_off)
ue_end = ue_off + 8 + ue_rl
print(f"\n=== After UserEditAtom ===")
print(f"  UserEditAtom: offset={ue_off}, end={ue_end}")
print(f"  stream length={len(orig_ppt)}")
print(f"  remaining={len(orig_ppt) - ue_end}")
