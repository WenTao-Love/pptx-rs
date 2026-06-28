"""详细解析 UserEditAtom 的所有字段，确认偏移量。"""
import io
import os
import struct
import olefile


def main():
    orig_path = None
    for f in os.listdir('_test'):
        if f.endswith('.ppt'):
            orig_path = '_test/' + f
            break

    with open(orig_path, 'rb') as f:
        orig_data = f.read()

    ole = olefile.OleFileIO(io.BytesIO(orig_data))
    ppt = ole.openstream('PowerPoint Document').read()
    cu = ole.openstream('Current User').read()

    offset_to_current_edit = struct.unpack_from('<I', cu, 16)[0]
    ue_offset = offset_to_current_edit

    print(f'UserEditAtom offset: {ue_offset}')
    print(f'原始数据 (36 bytes):')
    for i in range(0, 36, 4):
        val = struct.unpack_from('<I', ppt, ue_offset + i)[0]
        print(f'  offset {ue_offset}+{i:2d} = {val:10d} (0x{val:08X})')

    print(f'\n按 MS-PPT 2.3.3 规范解析:')
    print(f'  header.ver_inst = 0x{struct.unpack_from("<H", ppt, ue_offset)[0]:04X}')
    print(f'  header.recType  = 0x{struct.unpack_from("<H", ppt, ue_offset+2)[0]:04X}')
    print(f'  header.recLen   = {struct.unpack_from("<I", ppt, ue_offset+4)[0]}')
    print(f'  lastSlideIdRef  = {struct.unpack_from("<I", ppt, ue_offset+8)[0]}')
    print(f'  version         = 0x{struct.unpack_from("<H", ppt, ue_offset+12)[0]:04X}')
    print(f'  minorVersion    = {ppt[ue_offset+14]}')
    print(f'  majorVersion    = {ppt[ue_offset+15]}')
    print(f'  offsetLastEdit  = {struct.unpack_from("<I", ppt, ue_offset+16)[0]}')
    print(f'  offsetPersistDir= {struct.unpack_from("<I", ppt, ue_offset+20)[0]}')
    print(f'  persistIdSeed   = {struct.unpack_from("<I", ppt, ue_offset+24)[0]}')
    print(f'  docPersistRef   = {struct.unpack_from("<I", ppt, ue_offset+28)[0]}')
    print(f'  maxPersistWritten= {struct.unpack_from("<I", ppt, ue_offset+32)[0]}')

    # 验证 offsetPersistDirectory 指向 PersistDirectoryAtom
    pd_offset = struct.unpack_from('<I', ppt, ue_offset+20)[0]
    if pd_offset + 8 <= len(ppt):
        pd_type = struct.unpack_from('<H', ppt, pd_offset+2)[0]
        pd_len = struct.unpack_from('<I', ppt, pd_offset+4)[0]
        print(f'\n  PersistDirectoryAtom at offset {pd_offset}: type=0x{pd_type:04X}, recLen={pd_len}')

    ole.close()


if __name__ == '__main__':
    main()
