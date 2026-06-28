"""比较原始文件和加水印文件的 persist directory，确认 offset 是否变化。"""
import olefile
from struct import unpack


def parse_record_header(data, offset):
    ver_inst = int.from_bytes(data[offset:offset+2], 'little')
    rec_type = int.from_bytes(data[offset+2:offset+4], 'little')
    rec_len = int.from_bytes(data[offset+4:offset+8], 'little')
    return ver_inst & 0x0F, (ver_inst >> 4) & 0x0FFF, rec_type, rec_len


def parse_persist_directory(data, offset):
    ver, inst, rec_type, rec_len = parse_record_header(data, offset)
    if rec_type != 0x1772:
        return []
    pd_data = data[offset+8:offset+8+rec_len]
    entries = []
    pos = 0
    while pos + 4 <= len(pd_data):
        entry_val = int.from_bytes(pd_data[pos:pos+4], 'little')
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        pos += 4
        for j in range(c_persist):
            if pos + 4 <= len(pd_data):
                persist_offset = int.from_bytes(pd_data[pos:pos+4], 'little')
                entries.append((persist_id + j, persist_offset))
                pos += 4
    return entries


def main():
    original_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt"
    wm_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_心理账户理论.ppt"

    ole_orig = olefile.OleFileIO(original_path)
    ole_wm = olefile.OleFileIO(wm_path)

    ppt_orig = ole_orig.openstream('PowerPoint Document').read()
    ppt_wm = ole_wm.openstream('PowerPoint Document').read()

    print(f"原始 PowerPoint Document: {len(ppt_orig)} bytes")
    print(f"加水印 PowerPoint Document: {len(ppt_wm)} bytes")
    print(f"差异: {len(ppt_wm) - len(ppt_orig)} bytes")

    cu_orig = ole_orig.openstream('Current User').read()
    offset_to_current_edit_orig = int.from_bytes(cu_orig[16:20], 'little')
    offset_persist_dir_orig = int.from_bytes(
        ppt_orig[offset_to_current_edit_orig+20:offset_to_current_edit_orig+24], 'little')
    persist_entries_orig = parse_persist_directory(ppt_orig, offset_persist_dir_orig)

    cu_wm = ole_wm.openstream('Current User').read()
    offset_to_current_edit_wm = int.from_bytes(cu_wm[16:20], 'little')
    offset_persist_dir_wm = int.from_bytes(
        ppt_wm[offset_to_current_edit_wm+20:offset_to_current_edit_wm+24], 'little')
    persist_entries_wm = parse_persist_directory(ppt_wm, offset_persist_dir_wm)

    print(f"\n原始文件 persist entries: {len(persist_entries_orig)}")
    print(f"加水印文件 persist entries: {len(persist_entries_wm)}")

    orig_dict = dict(persist_entries_orig)
    wm_dict = dict(persist_entries_wm)

    print(f"\n{'pid':>4} | {'orig off':>10} {'orig type':>10} {'orig len':>8} | {'wm off':>10} {'wm type':>10} {'wm len':>8} | {'off diff':>10}")
    print("-" * 100)
    for pid in sorted(orig_dict.keys()):
        orig_offset = orig_dict[pid]
        ver, inst, rec_type, rec_len = parse_record_header(ppt_orig, orig_offset)
        if pid in wm_dict:
            wm_offset = wm_dict[pid]
            ver2, inst2, rec_type2, rec_len2 = parse_record_header(ppt_wm, wm_offset)
            diff = wm_offset - orig_offset
            print(f"{pid:>4} | {orig_offset:>10} {hex(rec_type):>10} {rec_len:>8} | {wm_offset:>10} {hex(rec_type2):>10} {rec_len2:>8} | {diff:>10}")
        else:
            print(f"{pid:>4} | {orig_offset:>10} {hex(rec_type):>10} {rec_len:>8} | {'N/A':>10}")

    ole_orig.close()
    ole_wm.close()


if __name__ == "__main__":
    main()
