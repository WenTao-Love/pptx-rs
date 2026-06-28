# -*- coding: utf-8 -*-
"""详细解析 .ppt 文件中的文本框 SpContainer 结构。"""
import struct
import sys
import io
import olefile

# 强制 UTF-8 输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    is_container = (ver == 0xF)
    return (ver, inst, rec_type, rec_len, is_container)

OFFICE_ART_NAMES = {
    0xF000: 'DggContainer', 0xF001: 'DgContainer', 0xF003: 'SpgrContainer',
    0xF004: 'SpContainer', 0xF005: 'SolverContainer', 0xF006: 'FDG',
    0xF007: 'FSPGR', 0xF008: 'FSP', 0xF009: 'FOPT', 0xF00A: 'FOT',
    0xF00B: 'FOPT', 0xF00D: 'ClientTextbox', 0xF00E: 'ClientData',
    0xF00F: 'ChildAnchor', 0xF010: 'ClientAnchor',
}

# PPT Text record types
PPT_TEXT_TYPES = {
    0x0FA0: 'TextHeaderAtom', 0x0FA1: 'TextCharsAtom', 0x0FA2: 'StyleTextPropAtom',
    0x0FA3: 'TextMasterStyleAtom', 0x0FA4: 'TextOther', 0x0FA5: 'TextClickInfoAtom',
    0x0FA6: 'TextCharsAtom', 0x0FA7: 'TextSpecialInfoAtom', 0x0FA8: 'TextHeaderAtom',
    0x0FD9: 'TextCharsAtom', 0x0FDA: 'TextCharsAtom',
}

def dump_container(data, offset, indent=0, max_depth=6):
    """递归解析 container 的内容。"""
    hdr = parse_record_header(data, offset)
    if hdr is None:
        return
    ver, inst, rec_type, rec_len, is_container = hdr
    name = OFFICE_ART_NAMES.get(rec_type, PPT_TEXT_TYPES.get(rec_type, f'Type0x{rec_type:04X}'))
    prefix = "  " * indent

    if rec_type == 0xF00B:  # FOPT
        num_props = inst
        print(f"{prefix}{name} (offset={offset}, len={rec_len}, {num_props} props)")
        pos = offset + 8
        for i in range(min(num_props, 20)):
            if pos + 6 > offset + 8 + rec_len:
                break
            opid = struct.unpack_from('<H', data, pos)[0]
            opid_val = opid & 0x3FFF
            opid_is_complex = (opid & 0x8000) != 0
            op = struct.unpack_from('<I', data, pos + 2)[0]
            print(f"{prefix}  [{i}] optId=0x{opid_val:04X} complex={opid_is_complex} val=0x{op:08X}")
            pos += 6
    elif rec_type == 0xF00A:  # FSP (实际是 0xF00A = FOT)
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len >= 8:
            spid = struct.unpack_from('<I', data, offset + 8)[0]
            flags = struct.unpack_from('<I', data, offset + 12)[0]
            print(f"{prefix}  shapeId={spid}, flags=0x{flags:08X}")
    elif rec_type == 0xF00F:  # ChildAnchor
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len >= 16:
            left, top, right, bottom = struct.unpack_from('<iiii', data, offset + 8)
            print(f"{prefix}  left={left}, top={top}, right={right}, bottom={bottom}")
    elif rec_type in PPT_TEXT_TYPES:
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len > 0:
            raw = data[offset+8:offset+8+min(60, rec_len)]
            try:
                text = raw.decode('utf-16-le', errors='replace').rstrip('\x00')
                print(f"{prefix}  text={text!r}")
            except:
                print(f"{prefix}  data={raw.hex()}")
    elif is_container:
        print(f"{prefix}{name} (offset={offset}, len={rec_len}, container)")
        if indent < max_depth:
            pos = offset + 8
            end = offset + 8 + rec_len
            while pos + 8 <= end:
                child_hdr = parse_record_header(data, pos)
                if child_hdr is None:
                    break
                dump_container(data, pos, indent + 1, max_depth)
                pos += 8 + child_hdr[3]
    else:
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")

def main():
    ppt_path = "_test/心理账户理论.ppt"
    ole = olefile.OleFileIO(ppt_path)
    ppt_data = ole.openstream("PowerPoint Document").read()

    # 找到第一个 PPDrawing 中的第二个 SpContainer（offset=9016, len=1275）
    # 这是一个包含文本的形状
    print("=== SpContainer at offset=9016, len=1275 ===")
    dump_container(ppt_data, 9016)

    print("\n=== SpContainer at offset=23663, len=72 (简单形状) ===")
    dump_container(ppt_data, 23663)

    ole.close()

if __name__ == "__main__":
    main()
