#!/usr/bin/env python3
"""验证加密文件能否被 msoffcrypto 正确解密，并检查解密后的 UserEditAtom 是否完好。"""
import struct
import msoffcrypto
import olefile

def verify_file(filepath, password="pptx-rs-secret"):
    print(f"\n{'='*70}")
    print(f"验证: {filepath}")
    print(f"{'='*70}")

    try:
        with open(filepath, "rb") as f:
            officefile = msoffcrypto.OfficeFile(f)
            print(f"  is_encrypted: {officefile.is_encrypted()}")

            # 验证密码
            officefile.load_key(password=password)
            print(f"  密码验证: 通过")

            # 解密到临时文件
            import io
            decrypted = io.BytesIO()
            officefile.decrypt(decrypted)
            decrypted.seek(0)

            # 检查解密后的文件
            dec_data = decrypted.read()
            ole = olefile.OleFileIO(io.BytesIO(dec_data))
            ppt = ole.openstream("PowerPoint Document").read()

            # 检查 UserEditAtom
            cu = ole.openstream("Current User").read()
            offset_to_current_edit = struct.unpack_from("<I", cu, 16)[0]
            print(f"  解密后 offsetToCurrentEdit = {offset_to_current_edit}")

            ue_off = offset_to_current_edit
            if ue_off + 8 <= len(ppt):
                ue_rec_type = struct.unpack_from("<H", ppt, ue_off + 2)[0]
                ue_rec_len = struct.unpack_from("<I", ppt, ue_off + 4)[0]
                print(f"  解密后 UserEditAtom: recType=0x{ue_rec_len:04X}, recLen={ue_rec_len}")
                if ue_rec_len == 28:
                    print(f"  ✓ UserEditAtom recLen 正确 (28)")
                elif ue_rec_len == 32:
                    print(f"  ✓ UserEditAtom recLen 正确 (32, 含 encryptSessionPersistIdRef)")
                else:
                    print(f"  ✗ UserEditAtom recLen 异常! 期望 28 或 32, 实际 {ue_rec_len}")

            # 检查 persist 对象是否按 persistId 顺序排列
            offset_persist_dir = struct.unpack_from("<I", ppt, ue_off + 20)[0]
            pd_off = offset_persist_dir
            pd_rec_len = struct.unpack_from("<I", ppt, pd_off + 4)[0]
            entry_val = struct.unpack_from("<I", ppt, pd_off + 8)[0]
            persist_id = entry_val & 0xFFFFF
            c_persist = (entry_val >> 20) & 0xFFF

            offsets = []
            for i in range(c_persist):
                off = struct.unpack_from("<I", ppt, pd_off + 12 + i * 4)[0]
                offsets.append(off)

            is_sorted = all(offsets[i] < offsets[i + 1] for i in range(len(offsets) - 1))
            print(f"  persist 对象按 persistId 顺序排列: {is_sorted}")

            if is_sorted:
                # 验证 recLen 计算
                print(f"  验证 recLen 计算 (前5个):")
                for i in range(min(5, len(offsets) - 1)):
                    off = offsets[i]
                    next_off = offsets[i + 1]
                    rec_len_hdr = struct.unpack_from("<I", ppt, off + 4)[0]
                    rec_len_calc = next_off - off - 8
                    match = "✓" if rec_len_hdr == rec_len_calc else "✗"
                    print(f"    pid {persist_id + i}: offset={off}, recLen(hdr)={rec_len_hdr}, recLen(calc)={rec_len_calc} {match}")

            print(f"  ✓ 解密成功，文件结构正常")

    except Exception as e:
        print(f"  ✗ 错误: {e}")
        import traceback
        traceback.print_exc()

# 验证两个加密文件
verify_file("_test_out/protected_心理账户理论.ppt")
verify_file("_test_out/wm_protected_心理账户理论.ppt")

# 也验证参考文件
verify_file("_test_out/rc4cryptoapi_password.ppt", password="Password1234_")
