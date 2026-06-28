"""检查解密后的 .ppt 文件中是否包含水印文本。"""
import os
import sys

files = [
    '_test_out/protected_心理账户理论.ppt.verified.ppt',
    '_test_out/wm_protected_心理账户理论.ppt.verified.ppt',
    '_test_out/wm_心理账户理论.ppt',
]

watermark_text = 'pptx-rs 水印'
watermark_bytes_utf8 = watermark_text.encode('utf-8')
watermark_bytes_utf16le = watermark_text.encode('utf-16-le')

for fpath in files:
    print(f'=== 检查: {fpath} ===')
    if not os.path.exists(fpath):
        print(f'  文件不存在')
        continue
    with open(fpath, 'rb') as f:
        data = f.read()
    print(f'  文件大小: {len(data)} bytes')
    # 水印文本在 PPT 中以 UTF-16LE 编码存储（TextCharsAtom）
    found_utf16le = watermark_bytes_utf16le in data
    found_utf8 = watermark_bytes_utf8 in data
    print(f'  水印文本 (UTF-16LE): {"找到" if found_utf16le else "未找到"}')
    print(f'  水印文本 (UTF-8): {"找到" if found_utf8 else "未找到"}')

    # 检查 FOPT 属性值 0x01C2=0x0000000D
    # FOPT 属性格式: property_id (u16) + property_value (u32)
    target = b'\xC2\x01' + (0x0000000D).to_bytes(4, 'little')
    found_lock = target in data
    print(f'  0x01C2=0x0000000D (锁定选择+编辑): {"找到" if found_lock else "未找到"}')

    # 检查旧的错误值 0x01C2=0x00000045
    target_old = b'\xC2\x01' + (0x00000045).to_bytes(4, 'little')
    found_lock_old = target_old in data
    print(f'  0x01C2=0x00000045 (旧错误值): {"找到" if found_lock_old else "未找到"}')
    print()

sys.exit(0)
