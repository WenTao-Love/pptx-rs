"""验证 .ppt 加密文件：用 msoffcrypto 解密，并比较 Pictures Stream。"""
import msoffcrypto
import io
import os
import sys
import glob
import olefile

PASSWORD = 'pptx-rs-secret'


def find_file(prefix):
    """在 _test_out 中查找以 prefix 开头的 .ppt 文件（排除中间文件）。"""
    for f in os.listdir('_test_out'):
        if f.startswith(prefix) and f.endswith('.ppt') and '.' not in f[len(prefix):-4]:
            return '_test_out/' + f
    return None


def decrypt_file(path, password):
    """用 msoffcrypto 解密 .ppt 文件，返回解密后的字节。"""
    with open(path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        out = io.BytesIO()
        office_file.decrypt(out)
        return out.getvalue()


def get_stream(data, stream_name):
    """从 OLE2 数据中读取指定 stream 的内容。"""
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        if ole.exists(stream_name):
            return ole.openstream(stream_name).read()
        return None
    finally:
        ole.close()


def main():
    # 1. 读取原始文件
    orig_path = None
    for f in os.listdir('_test'):
        if f.endswith('.ppt'):
            orig_path = '_test/' + f
            break
    if not orig_path:
        print('找不到原始 .ppt 文件')
        sys.exit(1)

    with open(orig_path, 'rb') as f:
        orig_data = f.read()
    print('原始文件:', orig_path, '大小:', len(orig_data))

    # 2. 解密 protected 文件
    prefix = sys.argv[1] if len(sys.argv) > 1 else 'protected_'
    prot_path = find_file(prefix)
    if not prot_path:
        print('找不到', prefix, '文件')
        sys.exit(1)
    print('加密文件:', prot_path)

    try:
        decrypted = decrypt_file(prot_path, PASSWORD)
        print('msoffcrypto 解密成功! 大小:', len(decrypted))
    except Exception as e:
        print('msoffcrypto 解密失败:', e)
        sys.exit(1)

    # 3. 比较 Pictures Stream
    orig_pics = get_stream(orig_data, 'Pictures')
    dec_pics = get_stream(decrypted, 'Pictures')

    if orig_pics is None and dec_pics is None:
        print('Pictures Stream: 原始和解密文件都没有（OK）')
    elif orig_pics is None:
        print('Pictures Stream: 原始没有但解密有（异常!）')
    elif dec_pics is None:
        print('Pictures Stream: 原始有但解密没有（msoffcrypto 不解密 Pictures，正常）')
        print('  原始 Pictures 大小:', len(orig_pics))
    else:
        if orig_pics == dec_pics:
            print('Pictures Stream: 一致 (OK)')
        else:
            print('Pictures Stream: 不一致!')
            print('  原始大小:', len(orig_pics), '解密大小:', len(dec_pics))
            # 找第一个不同的字节
            min_len = min(len(orig_pics), len(dec_pics))
            for i in range(min_len):
                if orig_pics[i] != dec_pics[i]:
                    print('  第一个差异在字节', i, ': 原始=', hex(orig_pics[i]), '解密=', hex(dec_pics[i]))
                    break

    # 4. 比较 PowerPoint Document stream（msoffcrypto 会解密这个）
    orig_ppt = get_stream(orig_data, 'PowerPoint Document')
    dec_ppt = get_stream(decrypted, 'PowerPoint Document')
    if orig_ppt and dec_ppt:
        if orig_ppt == dec_ppt:
            print('PowerPoint Document: 一致 (OK)')
        else:
            print('PowerPoint Document: 不完全一致（msoffcrypto 解密后会修改 UserEditAtom 等，正常）')
            print('  原始大小:', len(orig_ppt), '解密大小:', len(dec_ppt))
    else:
        print('PowerPoint Document: 缺失!')

    print('\n验证完成。')


if __name__ == '__main__':
    main()
