"""全面诊断：OLE2 结构 + CurrentUser + 水印 shape 搜索。"""
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

def check_ole_structure(fname):
    """检查 OLE2 容器结构。"""
    print(f'\n=== {fname} - OLE2 结构 ===')
    ole = olefile.OleFileIO(fname)
    print(f'  Root entry CLSID: {ole.root.clsid}')
    print(f'  Sector size: {ole.sectorsize}')
    print(f'  Mini sector size: {ole.minisectorcutoff}')
    print(f'  Streams:')
    for stream in ole.listdir():
        name = '/'.join(stream)
        size = ole.get_size(name)
        print(f'    {name}: {size} bytes')
    ole.close()

def check_current_user(fname):
    """检查 Current User stream。"""
    print(f'\n=== {fname} - Current User ===')
    ole = olefile.OleFileIO(fname)
    try:
        data = ole.openstream('Current User').read()
        print(f'  大小: {len(data)} bytes')
        if len(data) >= 20:
            # CurrentUserAtom 结构
            ver, inst, rec_type, rec_len = parse_record_header(data, 0)
            print(f'  RecordHeader: ver={ver}, inst={inst}, type=0x{rec_type:04X}, len={rec_len}')
            if rec_type == 0x0FF6 and rec_len >= 20:
                size = struct.unpack_from('<I', data, 8)[0]
                header_token = struct.unpack_from('<I', data, 12)[0]
                offset_to_current_edit = struct.unpack_from('<I', data, 16)[0]
                print(f'  size: {size}')
                print(f'  headerToken: 0x{header_token:08X} (期望 0xF3D1C4D0)')
                print(f'  offsetToCurrentEdit: {offset_to_current_edit}')
                if header_token == 0xF3D1C4D0:
                    print(f'  ✓ headerToken 正确（加密）')
                else:
                    print(f'  ✗ headerToken 错误')
    finally:
        ole.close()

def find_watermark_shapes(fname):
    """查找所有水印 shape。"""
    print(f'\n=== {fname} - 查找水印 shape ===')
    ole = olefile.OleFileIO(fname)
    try:
        data = ole.openstream('PowerPoint Document').read()
        print(f'  PowerPoint Document 大小: {len(data)} bytes')

        # 查找水印文本
        watermark_text = "pptx-rs 水印".encode('utf-16-le')
        text_positions = []
        pos = 0
        while True:
            pos = data.find(watermark_text, pos)
            if pos < 0:
                break
            text_positions.append(pos)
            pos += 1

        print(f'  找到 {len(text_positions)} 处水印文本')
        for text_pos in text_positions:
            print(f'    offset={text_pos}')

            # 向前查找 SpContainer (0xF009)
            # 搜索范围：前 5000 字节
            search_start = max(0, text_pos - 5000)
            found_sp = False
            for offset in range(text_pos, search_start, -1):
                if offset + 8 > len(data):
                    continue
                ver_inst = struct.unpack_from('<H', data, offset)[0]
                rec_type = struct.unpack_from('<H', data, offset + 2)[0]
                rec_len = struct.unpack_from('<I', data, offset + 4)[0]

                if rec_type == 0xF009 and (ver_inst & 0x0F) == 0xF:
                    container_end = offset + 8 + rec_len
                    if container_end > text_pos and container_end < len(data) + 1:
                        # 找到包含水印文本的 SpContainer
                        print(f'    找到 SpContainer: offset={offset}, len={rec_len}')
                        parse_spcontainer(data, offset)
                        found_sp = True
                        break

            if not found_sp:
                print(f'    ✗ 未找到包含水印的 SpContainer')

    finally:
        ole.close()

def parse_spcontainer(data, start_offset):
    """解析 SpContainer。"""
    ver, inst, rec_type, rec_len = parse_record_header(data, start_offset)
    print(f'      SpContainer: ver={ver}, inst={inst}, type=0x{rec_type:04X}, len={rec_len}')

    offset = start_offset + 8
    end = start_offset + 8 + rec_len
    while offset < end:
        hdr = parse_record_header(data, offset)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len = hdr

        if rec_type == 0xF00A:  # FSP
            shape_id = struct.unpack_from('<I', data, offset + 8)[0]
            flags = struct.unpack_from('<I', data, offset + 12)[0]
            print(f'      FSP: inst=0x{inst:04X}, shapeId={shape_id}, flags=0x{flags:08X}')

        elif rec_type == 0xF00B:  # FOPT
            num_props = inst
            print(f'      FOPT: num_props={num_props}')
            prop_offset = offset + 8
            for i in range(num_props):
                if prop_offset + 6 > len(data):
                    break
                prop_id = struct.unpack_from('<H', data, prop_offset)[0]
                prop_val = struct.unpack_from('<I', data, prop_offset + 2)[0]
                prop_names = {
                    0x00BD: 'rotation',
                    0x01BF: 'FillStyleBool',
                    0x01C1: 'LineStyleBool',
                    0x0180: 'fillColor',
                    0x0181: 'lineColor',
                }
                name = prop_names.get(prop_id, f'prop_0x{prop_id:04X}')
                print(f'        [{i}] {name} (0x{prop_id:04X}) = 0x{prop_val:08X}')
                prop_offset += 6

        elif rec_type == 0xF010:  # ClientAnchor
            if rec_len >= 8:
                top, left, right, bottom = struct.unpack_from('<hhhh', data, offset + 8)
                print(f'      ClientAnchor: top={top}, left={left}, right={right}, bottom={bottom}')

        elif rec_type == 0xF00D:  # ClientTextbox (container)
            print(f'      ClientTextbox: len={rec_len}')

        offset += 8 + rec_len

# 检查水印文件
check_ole_structure('_test_out/wm_心理账户理论.ppt')
find_watermark_shapes('_test_out/wm_心理账户理论.ppt')

# 检查加密文件
check_ole_structure('_test_out/protected_心理账户理论.ppt')
check_current_user('_test_out/protected_心理账户理论.ppt')
check_current_user('_test_out/wm_protected_心理账户理论.ppt')
