"""验证加密的 .ppt 文件能否被 msoffcrypto 正确解密。"""
import msoffcrypto
import os
import sys

files = [
    '_test_out/protected_心理账户理论.ppt',
    '_test_out/wm_protected_心理账户理论.ppt',
]
password = 'pptx-rs-secret'

all_ok = True
for fpath in files:
    print(f'=== 验证: {fpath} ===')
    if not os.path.exists(fpath):
        print(f'  文件不存在')
        all_ok = False
        continue
    try:
        with open(fpath, 'rb') as f:
            office = msoffcrypto.OfficeFile(f)
            print(f'  is_encrypted: {office.is_encrypted()}')
            office.load_key(password=password)
            out_path = fpath + '.verified.ppt'
            with open(out_path, 'wb') as out:
                office.decrypt(out)
            print(f'  解密成功: {out_path} ({os.path.getsize(out_path)} bytes)')
    except Exception as e:
        print(f'  解密失败: {type(e).__name__}: {e}')
        all_ok = False
    print()

sys.exit(0 if all_ok else 1)
