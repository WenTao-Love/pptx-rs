"""调试 msoffcrypto 解密后的文件结构"""
import olefile
import io
from struct import unpack


def parse_record_header(data, offset):
    ver_inst = int.from_bytes(data[offset:offset+2], 'little')
    rec_type = int.from_bytes(data[offset+2:offset+4], 'little')
    rec_len = int.from_bytes(data[offset+4:offset+8], 'little')
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return ver, inst, rec_type, rec_len


def main():
    decrypted_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_protected_心理账户理论.ppt.decrypted2.ppt"
    ole = olefile.OleFileIO(decrypted_path)

    ppt = ole.openstream('PowerPoint Document').read()
    cu = ole.openstream('Current User').read()

    print(f"PowerPoint Document: {len(ppt)} bytes")
    print(f"Current User: {len(cu)} bytes")

    # 解析 Current User
    ver, inst, rec_type, rec_len = parse_record_header(cu, 0)
    print(f"\nCurrentUserAtom: type={rec_type:#06x}, len={rec_len}")
    size = int.from_bytes(cu[8:12], 'little')
    header_token = int.from_bytes(cu[12:16], 'little')
    offset_to_current_edit = int.from_bytes(cu[16:20], 'little')
    print(f"  size={size:#x}, headerToken={header_token:#010x}, offsetToCurrentEdit={offset_to_current_edit}")

    # 解析 UserEditAtom
    print(f"\nUserEditAtom at offset {offset_to_current_edit}:")
    ver, inst, rec_type, rec_len = parse_record_header(ppt, offset_to_current_edit)
    print(f"  type={rec_type:#06x}, len={rec_len}")
    print(f"  (期望: type=0x0FF5, len=28 未加密)")

    if rec_type == 0x0FF5:
        ue_offset = offset_to_current_edit
        last_slide_id_ref = int.from_bytes(ppt[ue_offset+8:ue_offset+12], 'little')
        version = int.from_bytes(ppt[ue_offset+12:ue_offset+14], 'little')
        minor_major = ppt[ue_offset+14:ue_offset+16]
        offset_last_edit = int.from_bytes(ppt[ue_offset+16:ue_offset+20], 'little')
        offset_persist_dir = int.from_bytes(ppt[ue_offset+20:ue_offset+24], 'little')
        doc_persist_id_ref = int.from_bytes(ppt[ue_offset+24:ue_offset+28], 'little')
        max_persist_written = int.from_bytes(ppt[ue_offset+28:ue_offset+32], 'little')
        print(f"  lastSlideIdRef={last_slide_id_ref:#x}")
        print(f"  version={version:#x}")
        print(f"  offsetLastEdit={offset_last_edit:#x}")
        print(f"  offsetPersistDirectory={offset_persist_dir} ({offset_persist_dir:#x})")
        print(f"  docPersistIdRef={doc_persist_id_ref}")
        print(f"  maxPersistWritten={max_persist_written}")

        # 解析 PersistDirectoryAtom
        print(f"\nPersistDirectoryAtom at offset {offset_persist_dir}:")
        if offset_persist_dir + 8 <= len(ppt):
            ver, inst, rec_type, rec_len = parse_record_header(ppt, offset_persist_dir)
            print(f"  type={rec_type:#06x}, len={rec_len}")
            print(f"  (期望: type=0x1772)")

            if rec_type == 0x1772:
                pd_data = ppt[offset_persist_dir+8:offset_persist_dir+8+rec_len]
                pos = 0
                while pos + 4 <= len(pd_data):
                    entry_val = int.from_bytes(pd_data[pos:pos+4], 'little')
                    persist_id = entry_val & 0xFFFFF
                    c_persist = (entry_val >> 20) & 0xFFF
                    pos += 4
                    print(f"  Entry: persistId={persist_id}, cPersist={c_persist}")
                    for j in range(c_persist):
                        if pos + 4 <= len(pd_data):
                            persist_offset = int.from_bytes(pd_data[pos:pos+4], 'little')
                            print(f"    [{persist_id+j}] offset={persist_offset} ({persist_offset:#x})")
                            pos += 4
        else:
            print(f"  offset 超出范围！")

    ole.close()


if __name__ == "__main__":
    main()
