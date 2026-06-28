"""精确解析水印 SpContainer (0xF004) 的完整结构。"""
import olefile
import struct

def parse_record_header(data, offset):
    """解析 8 字节 record header。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0xFFF
    return (ver, inst, rec_type, rec_len)

def parse_fopt_props(data, offset, num_props):
    """解析 FOPT 属性。"""
    prop_names = {
        0x00BD: 'rotation',
        0x0180: 'fillType',
        0x0181: 'fillColor',
        0x0183: 'fillBackColor',
        0x018D: 'shadowColor',
        0x0191: 'shadowColor',
        0x01BF: 'FillStyleBool',
        0x01C1: 'LineStyleBool',
    }
    for i in range(num_props):
        if offset + 6 > len(data):
            break
        prop_id = struct.unpack_from('<H', data, offset)[0]
        prop_val = struct.unpack_from('<I', data, offset + 2)[0]
        name = prop_names.get(prop_id, f'prop_0x{prop_id:04X}')
        print(f'        [{i}] {name} (0x{prop_id:04X}) = 0x{prop_val:08X}')
        offset += 6
    return offset

def parse_container(data, start_offset, depth=0):
    """递归解析 container。"""
    indent = '  ' * depth
    hdr = parse_record_header(data, start_offset)
    if hdr is None:
        return start_offset + 8
    ver, inst, rec_type, rec_len = hdr

    rec_names = {
        0xF004: 'SpContainer',
        0xF003: 'SpgrContainer',
        0xF005: 'BStoreContainer',
        0xF006: 'SolverContainer',
        0xF00A: 'FSP',
        0xF00B: 'FOPT',
        0xF00D: 'ClientTextbox',
        0xF010: 'ClientAnchor',
        0xF011: 'ClientData',
        0x0F9F: 'TextHeaderAtom',
        0x0FA0: 'TextCharsAtom',
        0x0FA1: 'StyleTextPropAtom',
    }
    name = rec_names.get(rec_type, f'Unknown_0x{rec_type:04X}')
    end = start_offset + 8 + rec_len

    print(f'{indent}offset={start_offset}: {name} (type=0x{rec_type:04X}, ver={ver}, inst={inst}, len={rec_len})')

    if rec_type == 0xF00A:  # FSP
        if rec_len >= 8:
            shape_id = struct.unpack_from('<I', data, start_offset + 8)[0]
            flags = struct.unpack_from('<I', data, start_offset + 12)[0]
            print(f'{indent}  shapeId={shape_id}, flags=0x{flags:08X}, inst(shapeType)=0x{inst:04X}')

    elif rec_type == 0xF00B:  # FOPT
        print(f'{indent}  num_props={inst}')
        parse_fopt_props(data, start_offset + 8, inst)

    elif rec_type == 0xF010:  # ClientAnchor
        if rec_len >= 8:
            top, left, right, bottom = struct.unpack_from('<hhhh', data, start_offset + 8)
            print(f'{indent}  top={top}, left={left}, right={right}, bottom={bottom}')

    elif rec_type == 0x0FA0:  # TextCharsAtom
        text = data[start_offset+8:start_offset+8+rec_len].decode('utf-16-le', errors='replace')
        print(f'{indent}  text="{text}"')

    elif rec_type == 0x0FA1:  # StyleTextPropAtom
        style_data = data[start_offset+8:start_offset+8+rec_len]
        if len(style_data) >= 20:
            lfo = struct.unpack_from('<I', style_data, 0)[0]
            count1 = struct.unpack_from('<I', style_data, 4)[0]
            indent_level = struct.unpack_from('<h', style_data, 8)[0]
            pf_flags = struct.unpack_from('<I', style_data, 10)[0]
            count2 = struct.unpack_from('<I', style_data, 14)[0]
            cf_flags = struct.unpack_from('<I', style_data, 18)[0]
            print(f'{indent}  lfo={lfo}, count1={count1}, indentLevel={indent_level}, pfFlags=0x{pf_flags:08X}')
            print(f'{indent}  count2={count2}, cfFlags=0x{cf_flags:08X}')
            if cf_flags & 0x40 and len(style_data) >= 24:
                sz = struct.unpack_from('<H', style_data, 22)[0]
                print(f'{indent}  sz={sz} (={sz/100}pt)')
            if cf_flags & 0x80 and len(style_data) >= 28:
                r, g, b, idx = style_data[24], style_data[25], style_data[26], style_data[27]
                print(f'{indent}  color=({r},{g},{b},idx={idx})')

    # 如果是 container，递归解析子 record
    if ver == 0xF:
        offset = start_offset + 8
        while offset < end:
            child_hdr = parse_record_header(data, offset)
            if child_hdr is None:
                break
            child_ver, child_inst, child_type, child_len = child_hdr
            if child_len > rec_len or offset + 8 + child_len > end:
                break
            parse_container(data, offset, depth + 1)
            offset += 8 + child_len

    return end

fname = '_test_out/wm_心理账户理论.ppt'
ole = olefile.OleFileIO(fname)
data = ole.openstream('PowerPoint Document').read()

# 查找水印文本
watermark_text = "pptx-rs 水印".encode('utf-16-le')
text_pos = data.find(watermark_text)
print(f'水印文本位置: {text_pos}')

# 从水印文本向前搜索 SpContainer (0xF004)
# SpContainer 的 ver=0xF, type=0xF004
# 搜索 pattern: 0F 00 04 F0
pattern = b'\x0F\x00\x04\xF0'
search_start = max(0, text_pos - 10000)
pos = data.rfind(pattern, search_start, text_pos)
if pos >= 0:
    print(f'\n找到 SpContainer: offset={pos}')
    print(f'\n=== SpContainer 完整结构 ===')
    parse_container(data, pos)
else:
    print('未找到 SpContainer')

    # 打印水印文本前 100 字节的 hex dump
    print(f'\n=== 水印文本前 100 字节 hex dump ===')
    dump_start = max(0, text_pos - 100)
    for i in range(0, 100, 16):
        offset = dump_start + i
        if offset + 16 > len(data):
            break
        hex_str = ' '.join(f'{b:02X}' for b in data[offset:offset+16])
        print(f'  {offset:06X}: {hex_str}')

ole.close()
