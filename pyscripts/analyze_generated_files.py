# -*- coding: utf-8 -*-
"""深入分析生成的水印文件和加密文件结构，找出 PowerPoint 拒绝的根因。"""
import os
import struct
import olefile

BASE = os.path.dirname(__file__)
WM_FILE = os.path.join(BASE, "_test_out", "wm_心理账户理论.ppt")
ENC_FILE = os.path.join(BASE, "_test_out", "protected_心理账户理论.ppt")
ORIG_FILE = os.path.join(BASE, "_test", "心理账户理论.ppt")


def read_u32_le(data, offset):
    return struct.unpack_from("<I", data, offset)[0]


def read_u16_le(data, offset):
    return struct.unpack_from("<H", data, offset)[0]


def parse_record_header(data, offset):
    """返回 (ver, inst, recType, recLen, is_container)。"""
    if offset + 8 > len(data):
        return None
    ver_inst = read_u16_le(data, offset)
    rec_type = read_u16_le(data, offset + 2)
    rec_len = read_u32_le(data, offset + 4)
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    is_container = (ver == 0xF)
    return (ver, inst, rec_type, rec_len, is_container)


RT_NAMES = {
    0x03F8: "Slide",
    0x040C: "PPDrawing",
    0xF002: "SpgrContainer",
    0xF003: "SpContainer",
    0xF004: "UnknownContainer_F004",
    0xF006: "FSPGR",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "ClientTextbox(atom)",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF00D: "ClientTextbox(container)",
    0xF010: "ChildAnchor",
    0xF011: "ChildTextbox",
    0xF122: "Unknown_F122",
    0x0F9F: "TextHeaderAtom",
    0x0FA0: "TextCharsAtom",
    0x0FF5: "UserEditAtom",
    0x0FF6: "CurrentUserAtom",
    0x1772: "PersistDirectoryAtom",
    0x2F14: "CryptSession10Container",
}


def rt_name(rt):
    return RT_NAMES.get(rt, f"0x{rt:04X}")


def dump_record(data, offset, indent=0, max_depth=10):
    """递归打印 record 结构。"""
    if max_depth < 0:
        return 0
    hdr = parse_record_header(data, offset)
    if hdr is None:
        return 0
    ver, inst, rec_type, rec_len, is_container = hdr
    prefix = "  " * indent
    name = rt_name(rec_type)
    print(f"{prefix}{name}(type=0x{rec_type:04X}, ver={ver}, inst={inst}, len={rec_len}) @ {offset}")

    next_offset = offset + 8 + rec_len
    if is_container:
        pos = offset + 8
        end = offset + 8 + rec_len
        while pos + 8 <= end:
            pos = dump_record(data, pos, indent + 1, max_depth - 1)
            if pos == 0:
                break
    else:
        # 打印关键 atom 的数据
        data_start = offset + 8
        data_end = data_start + rec_len
        if rec_type == 0xF007:  # FSP
            if rec_len >= 8:
                shape_id = read_u32_le(data, data_start)
                flags = read_u32_le(data, data_start + 4)
                print(f"{prefix}  shapeId={shape_id}, flags=0x{flags:08X}")
        elif rec_type == 0xF00A:  # ClientAnchor
            anchor_data = data[data_start:data_end]
            print(f"{prefix}  raw ({len(anchor_data)}B): {anchor_data.hex()}")
        elif rec_type == 0x0F9F:  # TextHeaderAtom
            if rec_len >= 4:
                print(f"{prefix}  txType={read_u32_le(data, data_start)}")
        elif rec_type == 0x0FA0:  # TextCharsAtom
            try:
                text = data[data_start:data_end].decode('utf-16-le', errors='replace').rstrip('\x00')
                print(f"{prefix}  text={text!r}")
            except:
                print(f"{prefix}  raw: {data[data_start:data_end].hex()}")

    return next_offset


def analyze_watermark_file():
    """分析水印文件中的 SpContainer 结构。"""
    print("=" * 70)
    print("分析水印文件中的 SpContainer 结构")
    print("=" * 70)

    ole = olefile.OleFileIO(WM_FILE)
    ppt = ole.openstream("PowerPoint Document").read()

    # 找到第一个 Slide
    pos = 0
    slide_count = 0
    while pos + 8 <= len(ppt) and slide_count < 1:
        hdr = parse_record_header(ppt, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len, is_container = hdr
        total_len = 8 + rec_len

        if is_container and rec_type == 0x03F8:  # Slide
            slide_count += 1
            print(f"\n=== Slide at offset={pos}, len={rec_len} ===")

            # 在 Slide 中找 PPDrawing
            slide_end = pos + 8 + rec_len
            sp = pos + 8
            while sp + 8 <= slide_end:
                s_hdr = parse_record_header(ppt, sp)
                if s_hdr is None:
                    break
                s_ver, s_inst, s_type, s_len, s_container = s_hdr
                s_total = 8 + s_len

                if s_container and s_type == 0x040C:  # PPDrawing
                    print(f"\n  PPDrawing at offset={sp}, len={s_len}")
                    # 打印 PPDrawing 的完整结构
                    ppd_end = sp + 8 + s_len
                    pp = sp + 8
                    while pp + 8 <= ppd_end:
                        pp_hdr = parse_record_header(ppt, pp)
                        if pp_hdr is None:
                            break
                        pp_ver, pp_inst, pp_type, pp_len, pp_container = pp_hdr
                        print(f"\n    {rt_name(pp_type)} at offset={pp}, len={pp_len}")
                        if pp_container and pp_type == 0xF002:  # SpgrContainer
                            # 打印 SpgrContainer 中的所有子元素
                            spgr_end = pp + 8 + pp_len
                            sc = pp + 8
                            sp_idx = 0
                            while sc + 8 <= spgr_end:
                                sc_hdr = parse_record_header(ppt, sc)
                                if sc_hdr is None:
                                    break
                                sc_ver, sc_inst, sc_type, sc_len, sc_container = sc_hdr
                                sp_idx += 1
                                print(f"\n      === 子形状 #{sp_idx}: {rt_name(sc_type)} at offset={sc}, len={sc_len} ===")
                                sc = dump_record(ppt, sc, 7, max_depth=8)
                                if sc == 0:
                                    break
                        pp += 8 + pp_len
                        if not pp_container and pp_len == 0:
                            break
                    break

                sp += s_total
                if not s_container and s_len == 0:
                    break

        pos += total_len
        if not is_container and rec_len == 0:
            break

    ole.close()


def analyze_encrypt_file():
    """分析加密文件的关键字段。"""
    print("\n" + "=" * 70)
    print("分析加密文件的关键字段")
    print("=" * 70)

    ole = olefile.OleFileIO(ENC_FILE)
    ppt = ole.openstream("PowerPoint Document").read()
    cu = ole.openstream("Current User").read()

    print(f"\nPowerPoint Document 大小: {len(ppt)}")
    print(f"Current User 大小: {len(cu)}")

    # 解析 CurrentUserAtom
    print("\n--- CurrentUserAtom ---")
    cu_hdr = parse_record_header(cu, 0)
    print(f"  RecordHeader: ver={cu_hdr[0]}, inst={cu_hdr[1]}, type=0x{cu_hdr[2]:04X}, len={cu_hdr[3]}")
    print(f"  期望: type=0x0FF6 (CurrentUserAtom)")

    # CurrentUserAtom 字段
    # offset 0-3: header (已解析)
    # offset 8-11: size (4 bytes)
    # offset 12-15: headerToken (4 bytes) - 0xE391C05F=未加密, 0xF3D1C4DF=已加密
    # offset 16-19: offsetToCurrentEdit (4 bytes)
    cu_size = read_u32_le(cu, 8)
    cu_token = read_u32_le(cu, 12)
    cu_offset_to_edit = read_u32_le(cu, 16)
    print(f"  size={cu_size}")
    print(f"  headerToken=0x{cu_token:08X} ({'已加密' if cu_token == 0xF3D1C4DF else '未加密' if cu_token == 0xE391C05F else '未知!'})")
    print(f"  offsetToCurrentEdit={cu_offset_to_edit}")

    # 解析 UserEditAtom
    print("\n--- UserEditAtom ---")
    ue_offset = cu_offset_to_edit
    if ue_offset + 40 > len(ppt):
        print(f"  [ERROR] UserEditAtom offset {ue_offset} 超出范围 (len={len(ppt)})")
        ole.close()
        return

    ue_hdr = parse_record_header(ppt, ue_offset)
    print(f"  RecordHeader: ver={ue_hdr[0]}, inst={ue_hdr[1]}, type=0x{ue_hdr[2]:04X}, len={ue_hdr[3]}")
    print(f"  期望: type=0x0FF5 (UserEditAtom), len=32 (已加密) 或 28 (未加密)")

    # UserEditAtom 字段
    # offset 0-7: header
    # offset 8-11: lastSlideIdRef (4 bytes)
    # offset 12-15: version (2 bytes) + minorVersion (2 bytes)
    # offset 16-19: offsetLastEdit (4 bytes)
    # offset 20-23: offsetPersistDirectory (4 bytes)
    # offset 24-27: documentRef (4 bytes)
    # offset 28-31: maxPersistWritten (4 bytes) - persistIdSeed
    # offset 32-35: lastViewType (2 bytes) + unused (2 bytes)
    # offset 36-39: encryptSessionPersistIdRef (4 bytes) - 仅当 len=32 时存在
    ue_last_slide = read_u32_le(ppt, ue_offset + 8)
    ue_version = read_u16_le(ppt, ue_offset + 12)
    ue_minor = read_u16_le(ppt, ue_offset + 14)
    ue_offset_last_edit = read_u32_le(ppt, ue_offset + 16)
    ue_offset_pd = read_u32_le(ppt, ue_offset + 20)
    ue_doc_ref = read_u32_le(ppt, ue_offset + 24)
    ue_persist_seed = read_u32_le(ppt, ue_offset + 28)
    print(f"  lastSlideIdRef={ue_last_slide}")
    print(f"  version={ue_version}, minorVersion={ue_minor}")
    print(f"  offsetLastEdit={ue_offset_last_edit}")
    print(f"  offsetPersistDirectory={ue_offset_pd}")
    print(f"  documentRef={ue_doc_ref}")
    print(f"  persistIdSeed={ue_persist_seed}")

    if ue_hdr[3] >= 32:
        ue_encrypt_pid = read_u32_le(ppt, ue_offset + 36)
        print(f"  encryptSessionPersistIdRef={ue_encrypt_pid}")
        print(f"  persistIdSeed 应该 > encryptSessionPersistIdRef: {ue_persist_seed} > {ue_encrypt_pid} = {ue_persist_seed > ue_encrypt_pid}")

    # 解析 PersistDirectoryAtom
    print("\n--- PersistDirectoryAtom ---")
    pd_offset = ue_offset_pd
    if pd_offset + 8 > len(ppt):
        print(f"  [ERROR] PersistDirectoryAtom offset {pd_offset} 超出范围")
        ole.close()
        return

    pd_hdr = parse_record_header(ppt, pd_offset)
    print(f"  RecordHeader: ver={pd_hdr[0]}, inst={pd_hdr[1]}, type=0x{pd_hdr[2]:04X}, len={pd_hdr[3]}")
    print(f"  期望: type=0x1772 (PersistDirectoryAtom)")

    # 解析 persist entries
    pd_data_start = pd_offset + 8
    pd_data_end = pd_data_start + pd_hdr[3]
    pos = pd_data_start
    entry_count = 0
    total_persist = 0
    crypt_session_offset = None
    encrypt_pid = None

    if ue_hdr[3] >= 32:
        encrypt_pid = read_u32_le(ppt, ue_offset + 36)

    while pos + 4 <= pd_data_end:
        entry_val = read_u32_le(ppt, pos)
        persist_id = entry_val & 0xFFFFF
        c_persist = (entry_val >> 20) & 0xFFF
        pos += 4
        entry_count += 1

        for j in range(c_persist):
            if pos + 4 <= pd_data_end:
                poff = read_u32_le(ppt, pos)
                pid = persist_id + j
                if encrypt_pid is not None and pid == encrypt_pid:
                    crypt_session_offset = poff
                    print(f"  CryptSession10Container: persistId={pid}, offset={poff}")
                pos += 4
                total_persist += 1

    print(f"  总共 {entry_count} 个 entry, {total_persist} 个 persist 对象")

    # 解析 CryptSession10Container
    if crypt_session_offset is not None:
        print("\n--- CryptSession10Container ---")
        cs_offset = crypt_session_offset
        if cs_offset + 8 > len(ppt):
            print(f"  [ERROR] CryptSession10Container offset {cs_offset} 超出范围")
            ole.close()
            return

        cs_hdr = parse_record_header(ppt, cs_offset)
        print(f"  RecordHeader: ver={cs_hdr[0]}, inst={cs_hdr[1]}, type=0x{cs_hdr[2]:04X}, len={cs_hdr[3]}")
        print(f"  期望: type=0x2F14 (CryptSession10Container)")

        # 打印 CryptSession10Container 的原始数据（前 100 字节）
        cs_data_start = cs_offset + 8
        cs_data_end = cs_data_start + cs_hdr[3]
        cs_data = ppt[cs_data_start:cs_data_end]

        # EncryptionVersionInfo
        v_major = read_u16_le(ppt, cs_data_start)
        v_minor = read_u16_le(ppt, cs_data_start + 2)
        print(f"  EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}")

        # flags
        flags = read_u32_le(ppt, cs_data_start + 4)
        print(f"  flags: 0x{flags:08X}")

        # headerSize
        header_size = read_u32_le(ppt, cs_data_start + 8)
        print(f"  headerSize: {header_size}")

        # EncryptionHeader
        eh_start = cs_data_start + 12
        eh_flags = read_u32_le(ppt, eh_start)
        eh_size_extra = read_u32_le(ppt, eh_start + 4)
        eh_alg_id = read_u32_le(ppt, eh_start + 8)
        eh_alg_id_hash = read_u32_le(ppt, eh_start + 12)
        eh_key_size = read_u32_le(ppt, eh_start + 16)
        eh_provider_type = read_u32_le(ppt, eh_start + 20)
        print(f"  EncryptionHeader.flags: 0x{eh_flags:08X}")
        print(f"  EncryptionHeader.sizeExtra: {eh_size_extra}")
        print(f"  EncryptionHeader.algId: 0x{eh_alg_id:08X} (期望 0x00006801 = RC4)")
        print(f"  EncryptionHeader.algIdHash: 0x{eh_alg_id_hash:08X} (期望 0x00008004 = SHA1)")
        print(f"  EncryptionHeader.keySize: {eh_key_size} (期望 128)")
        print(f"  EncryptionHeader.providerType: {eh_provider_type}")

        # EncryptionVerifier
        ev_start = eh_start + header_size
        ev_salt_size = read_u32_le(ppt, ev_start)
        ev_salt = ppt[ev_start + 4:ev_start + 20]
        ev_enc_verifier = ppt[ev_start + 20:ev_start + 36]
        ev_hash_size = read_u32_le(ppt, ev_start + 36)
        ev_enc_hash = ppt[ev_start + 40:ev_start + 60]
        print(f"  EncryptionVerifier.saltSize: {ev_salt_size}")
        print(f"  EncryptionVerifier.salt: {ev_salt.hex()}")
        print(f"  EncryptionVerifier.encryptedVerifier: {ev_enc_verifier.hex()}")
        print(f"  EncryptionVerifier.verifierHashSize: {ev_hash_size}")
        print(f"  EncryptionVerifier.encryptedVerifierHash: {ev_enc_hash.hex()}")

    ole.close()


def compare_with_original():
    """对比原始文件和水印文件中的 SpContainer 结构。"""
    print("\n" + "=" * 70)
    print("对比原始文件和水印文件中的第一个 SpContainer 结构")
    print("=" * 70)

    for label, path in [("原始文件", ORIG_FILE), ("水印文件", WM_FILE)]:
        print(f"\n--- {label}: {os.path.basename(path)} ---")
        ole = olefile.OleFileIO(path)
        ppt = ole.openstream("PowerPoint Document").read()

        # 找到第一个 Slide 的 PPDrawing
        pos = 0
        found = False
        while pos + 8 <= len(ppt) and not found:
            hdr = parse_record_header(ppt, pos)
            if hdr is None:
                break
            ver, inst, rec_type, rec_len, is_container = hdr
            total_len = 8 + rec_len

            if is_container and rec_type == 0x03F8:  # Slide
                slide_end = pos + 8 + rec_len
                sp = pos + 8
                while sp + 8 <= slide_end:
                    s_hdr = parse_record_header(ppt, sp)
                    if s_hdr is None:
                        break
                    s_ver, s_inst, s_type, s_len, s_container = s_hdr

                    if s_container and s_type == 0x040C:  # PPDrawing
                        # 找到 SpgrContainer
                        ppd_end = sp + 8 + s_len
                        pp = sp + 8
                        while pp + 8 <= ppd_end:
                            pp_hdr = parse_record_header(ppt, pp)
                            if pp_hdr is None:
                                break
                            pp_ver, pp_inst, pp_type, pp_len, pp_container = pp_hdr

                            if pp_container and pp_type == 0xF002:  # SpgrContainer
                                # 打印 SpgrContainer 中的前 2 个子元素
                                spgr_end = pp + 8 + pp_len
                                sc = pp + 8
                                sp_idx = 0
                                while sc + 8 <= spgr_end and sp_idx < 2:
                                    sc_hdr = parse_record_header(ppt, sc)
                                    if sc_hdr is None:
                                        break
                                    sc_ver, sc_inst, sc_type, sc_len, sc_container = sc_hdr
                                    sp_idx += 1
                                    print(f"\n  子形状 #{sp_idx}: {rt_name(sc_type)} at offset={sc}")
                                    dump_record(ppt, sc, 2, max_depth=5)
                                    sc += 8 + sc_len
                                found = True
                                break

                            pp += 8 + pp_len
                            if not pp_container and pp_len == 0:
                                break
                        break

                    sp += 8 + s_len
                    if not s_container and s_len == 0:
                        break

            pos += total_len
            if not is_container and rec_len == 0:
                break

        ole.close()


def main():
    analyze_watermark_file()
    analyze_encrypt_file()
    compare_with_original()


if __name__ == "__main__":
    main()
