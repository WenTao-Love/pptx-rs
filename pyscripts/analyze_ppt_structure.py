# -*- coding: utf-8 -*-
"""分析原始 .ppt 文件中的 SpContainer 结构，以及加密文件的 CryptSession10Container 结构。"""
import sys
import os
import struct
import olefile

ORIG_FILE = os.path.join(os.path.dirname(__file__), "_test", "心理账户理论.ppt")
ENC_FILE = os.path.join(os.path.dirname(__file__), "_test_out", "protected_心理账户理论.ppt")


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


# OfficeArt record types
RT_SPGR_CONTAINER = 0xF002
RT_SP_CONTAINER = 0xF003
RT_FSPGR = 0xF006
RT_FSP = 0xF007
RT_FOPT = 0xF008
RT_CLIENT_TEXTBOX = 0xF009
RT_CLIENT_ANCHOR = 0xF00A
RT_CLIENT_DATA = 0xF00B

RT_NAMES = {
    0xF002: "SpgrContainer",
    0xF003: "SpContainer",
    0xF006: "FSPGR",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "ClientTextbox",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF010: "ChildAnchor",
    0x0F9F: "TextHeaderAtom",
    0x0FA0: "TextCharsAtom",
    0x0FA8: "TextBytesAtom",
}


def rt_name(rt):
    return RT_NAMES.get(rt, f"0x{rt:04X}")


def dump_spcontainer(data, offset, indent=0):
    """递归打印 SpContainer 的结构。"""
    hdr = parse_record_header(data, offset)
    if hdr is None:
        return
    ver, inst, rec_type, rec_len, is_container = hdr
    prefix = "  " * indent
    name = rt_name(rec_type)
    print(f"{prefix}{name}(type=0x{rec_type:04X}, ver={ver}, inst={inst}, len={rec_len})")

    if is_container:
        pos = offset + 8
        end = offset + 8 + rec_len
        while pos + 8 <= end:
            dump_spcontainer(data, pos, indent + 1)
            child_hdr = parse_record_header(data, pos)
            if child_hdr is None:
                break
            _, _, _, child_len, child_container = child_hdr
            pos += 8 + child_len
            if not child_container and child_len == 0:
                break
    else:
        # 打印非 container 的数据
        data_start = offset + 8
        data_end = data_start + rec_len
        if rec_type == RT_FSP:
            # FSP: shapeId(4) + flags(4)
            if rec_len >= 8:
                shape_id = read_u32_le(data, data_start)
                flags = read_u32_le(data, data_start + 4)
                print(f"{prefix}  shapeId={shape_id}, flags=0x{flags:08X}")
        elif rec_type == RT_FOPT:
            # FOPT: inst=num_props, 每个属性 6 字节
            print(f"{prefix}  num_props={inst}")
            for i in range(inst):
                p_off = data_start + i * 6
                if p_off + 6 <= data_end:
                    pid = read_u16_le(data, p_off)
                    val = read_u32_le(data, p_off + 2)
                    print(f"{prefix}  prop[{i}]: id=0x{pid:04X}, val=0x{val:08X}")
        elif rec_type == RT_CLIENT_ANCHOR:
            # ClientAnchor: 打印原始数据
            anchor_data = data[data_start:data_end]
            print(f"{prefix}  raw data ({len(anchor_data)} bytes): {anchor_data.hex()}")
            if len(anchor_data) >= 16:
                top = struct.unpack_from("<i", data, data_start)[0]
                left = struct.unpack_from("<i", data, data_start + 4)[0]
                right = struct.unpack_from("<i", data, data_start + 8)[0]
                bottom = struct.unpack_from("<i", data, data_start + 12)[0]
                print(f"{prefix}  top={top}, left={left}, right={right}, bottom={bottom}")
            elif len(anchor_data) >= 8:
                t = struct.unpack_from("<H", data, data_start)[0]
                l = struct.unpack_from("<H", data, data_start + 2)[0]
                r = struct.unpack_from("<H", data, data_start + 4)[0]
                b = struct.unpack_from("<H", data, data_start + 6)[0]
                print(f"{prefix}  (u16) top={t}, left={l}, right={r}, bottom={b}")
        elif rec_type == 0x0F9F:  # TextHeaderAtom
            if rec_len >= 4:
                tx_type = read_u32_le(data, data_start)
                print(f"{prefix}  txType={tx_type}")
        elif rec_type == 0x0FA0:  # TextCharsAtom
            try:
                text = data[data_start:data_end].decode('utf-16-le', errors='replace').rstrip('\x00')
                print(f"{prefix}  text={text!r}")
            except:
                print(f"{prefix}  raw: {data[data_start:data_end].hex()}")


def find_first_spcontainer(data):
    """找到第一个 SpContainer 并打印其结构。"""
    pos = 0
    while pos + 8 <= len(data):
        hdr = parse_record_header(data, pos)
        if hdr is None:
            break
        ver, inst, rec_type, rec_len, is_container = hdr
        total_len = 8 + rec_len

        if is_container and rec_type == RT_SP_CONTAINER:
            print(f"\n找到第一个 SpContainer at offset={pos}:")
            dump_spcontainer(data, pos)
            return

        pos += total_len
        if not is_container and rec_len == 0:
            break


def analyze_crypt_session(path):
    """分析 CryptSession10Container 的结构。"""
    print(f"\n=== 分析 CryptSession10Container: {os.path.basename(path)} ===")
    ole = olefile.OleFileIO(path)
    ppt = ole.openstream("PowerPoint Document").read()

    # 读取 CurrentUser 获取 offsetToCurrentEdit
    cu = ole.openstream("Current User").read()
    offset_to_current_edit = read_u32_le(cu, 16)

    # 读取 UserEditAtom
    ue_offset = offset_to_current_edit
    offset_persist_dir = read_u32_le(ppt, ue_offset + 20)
    encrypt_session_pid_ref = read_u32_le(ppt, ue_offset + 36)

    # 解析 PersistDirectoryAtom，找到 CryptSession10Container 的 offset
    pd_data_start = offset_persist_dir + 8
    entry_val = read_u32_le(ppt, pd_data_start)
    entry_pid = entry_val & 0xFFFFF
    entry_cpersist = (entry_val >> 20) & 0xFFF

    crypt_session_offset = None
    for i in range(entry_cpersist):
        poff = read_u32_le(ppt, pd_data_start + 4 + i * 4)
        pid = entry_pid + i
        if pid == encrypt_session_pid_ref:
            crypt_session_offset = poff
            break

    if crypt_session_offset is None:
        print("  [FAIL] 找不到 CryptSession10Container")
        ole.close()
        return

    print(f"  CryptSession10Container offset: {crypt_session_offset}")

    # 解析 CryptSession10Container
    cs_start = crypt_session_offset
    ver, inst, cs_type, cs_len, _ = parse_record_header(ppt, cs_start)
    print(f"  RecordHeader: ver={ver}, inst={inst}, type=0x{cs_type:04X}, len={cs_len}")

    data_start = cs_start + 8
    data_end = data_start + cs_len
    pos = data_start

    # EncryptionVersionInfo
    v_major = read_u16_le(ppt, pos)
    v_minor = read_u16_le(ppt, pos + 2)
    print(f"  EncryptionVersionInfo: vMajor={v_major}, vMinor={v_minor}")
    pos += 4

    # flags
    flags = read_u32_le(ppt, pos)
    print(f"  flags: 0x{flags:08X}")
    pos += 4

    # headerSize
    header_size = read_u32_le(ppt, pos)
    print(f"  headerSize: {header_size}")
    pos += 4

    # EncryptionHeader
    header_start = pos
    eh_flags = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.flags: 0x{eh_flags:08X}")
    pos += 4

    size_extra = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.sizeExtra: {size_extra}")
    pos += 4

    alg_id = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.algId: 0x{alg_id:08X}")
    pos += 4

    alg_id_hash = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.algIdHash: 0x{alg_id_hash:08X}")
    pos += 4

    key_size = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.keySize: {key_size}")
    pos += 4

    provider_type = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.providerType: {provider_type}")
    pos += 4

    reserved1 = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.reserved1: {reserved1}")
    pos += 4

    reserved2 = read_u32_le(ppt, pos)
    print(f"  EncryptionHeader.reserved2: {reserved2}")
    pos += 4

    # cspName (剩余的 header 数据)
    csp_data = ppt[pos:header_start + header_size]
    try:
        csp_name = csp_data.decode('utf-16-le', errors='replace').rstrip('\x00')
        print(f"  EncryptionHeader.cspName: {csp_name!r} ({len(csp_data)} bytes)")
    except:
        print(f"  EncryptionHeader.cspName: {csp_data.hex()}")
    pos = header_start + header_size

    # EncryptionVerifier
    salt_size = read_u32_le(ppt, pos)
    print(f"  EncryptionVerifier.saltSize: {salt_size}")
    pos += 4

    salt = ppt[pos:pos + 16]
    print(f"  EncryptionVerifier.salt: {salt.hex()}")
    pos += 16

    encrypted_verifier = ppt[pos:pos + 16]
    print(f"  EncryptionVerifier.encryptedVerifier: {encrypted_verifier.hex()}")
    pos += 16

    verifier_hash_size = read_u32_le(ppt, pos)
    print(f"  EncryptionVerifier.verifierHashSize: {verifier_hash_size}")
    pos += 4

    encrypted_verifier_hash = ppt[pos:pos + 20]
    print(f"  EncryptionVerifier.encryptedVerifierHash: {encrypted_verifier_hash.hex()}")
    pos += 20

    print(f"  总字节数: {pos - cs_start} (header 8 + data {cs_len})")

    ole.close()


def main():
    # 分析原始文件中的 SpContainer 结构
    print("=" * 60)
    print("分析原始 .ppt 文件中的 SpContainer 结构")
    print("=" * 60)

    ole = olefile.OleFileIO(ORIG_FILE)
    ppt = ole.openstream("PowerPoint Document").read()

    # 找到所有 Slide record
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
            print(f"\n找到 Slide at offset={pos}")

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
                    print(f"  PPDrawing at offset={sp}")

                    # 在 PPDrawing 中找 SpgrContainer
                    ppd_end = sp + 8 + s_len
                    pp = sp + 8
                    while pp + 8 <= ppd_end:
                        p_hdr = parse_record_header(ppt, pp)
                        if p_hdr is None:
                            break
                        p_ver, p_inst, p_type, p_len, p_container = p_hdr
                        p_total = 8 + p_len

                        if p_container and p_type == RT_SPGR_CONTAINER:
                            print(f"    SpgrContainer at offset={pp}")

                            # 打印 SpgrContainer 中的所有 SpContainer
                            spgr_end = pp + 8 + p_len
                            sc = pp + 8
                            sp_count = 0
                            while sc + 8 <= spgr_end and sp_count < 3:
                                sc_hdr = parse_record_header(ppt, sc)
                                if sc_hdr is None:
                                    break
                                sc_ver, sc_inst, sc_type, sc_len, sc_container = sc_hdr
                                sc_total = 8 + sc_len

                                if sc_container and sc_type == RT_SP_CONTAINER:
                                    sp_count += 1
                                    print(f"\n      SpContainer #{sp_count} at offset={sc}:")
                                    dump_spcontainer(ppt, sc, 3)

                                sc += sc_total
                                if not sc_container and sc_len == 0:
                                    break
                            break

                        pp += p_total
                        if not p_container and p_len == 0:
                            break
                    break

                sp += s_total
                if not s_container and s_len == 0:
                    break

        pos += total_len
        if not is_container and rec_len == 0:
            break

    ole.close()

    # 分析加密文件的 CryptSession10Container
    analyze_crypt_session(ENC_FILE)


if __name__ == "__main__":
    main()
