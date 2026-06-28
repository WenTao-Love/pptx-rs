"""
比较加密文件和原始文件的所有 OLE2 stream，找出差异。
"""
import olefile

# 读取加密文件
enc_ole = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\protected_心理账户理论.ppt')
# 读取原始文件
orig_ole = olefile.OleFileIO(r'd:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt')

print("=== OLE2 streams comparison ===")
print(f"\nenc streams: {enc_ole.listdir()}")
print(f"orig streams: {orig_ole.listdir()}")

print(f"\n=== Stream sizes ===")
enc_streams = {tuple(p): s for p, s in [(s, enc_ole.get_size(s)) for s in enc_ole.listdir()]}
orig_streams = {tuple(p): s for p, s in [(s, orig_ole.get_size(s)) for s in orig_ole.listdir()]}

for stream in sorted(set(list(enc_streams.keys()) + list(orig_streams.keys()))):
    enc_size = enc_streams.get(stream, 'MISSING')
    orig_size = orig_streams.get(stream, 'MISSING')
    name = '/'.join(stream)
    diff = ''
    if isinstance(enc_size, int) and isinstance(orig_size, int):
        diff = f' (diff: {enc_size - orig_size})' if enc_size != orig_size else ''
    print(f"  {name}: enc={enc_size}, orig={orig_size}{diff}")

# 比较未加密的 stream 内容
print(f"\n=== Compare unencrypted streams ===")
for stream in sorted(set(list(enc_streams.keys()) + list(orig_streams.keys()))):
    name = '/'.join(stream)
    if name in ['PowerPoint Document', 'Pictures']:
        continue  # 这些 stream 被加密了，跳过
    enc_data = enc_ole.openstream(stream).read() if stream in enc_streams else None
    orig_data = orig_ole.openstream(stream).read() if stream in orig_streams else None
    if enc_data is not None and orig_data is not None:
        if enc_data == orig_data:
            print(f"  {name}: IDENTICAL")
        else:
            print(f"  {name}: DIFFERENT (enc={len(enc_data)}, orig={len(orig_data)})")
            # 找到第一个不同的字节
            for i in range(min(len(enc_data), len(orig_data))):
                if enc_data[i] != orig_data[i]:
                    print(f"    first diff at byte {i}: enc={hex(enc_data[i])} orig={hex(orig_data[i])}")
                    break
    elif enc_data is not None:
        print(f"  {name}: only in enc")
    else:
        print(f"  {name}: only in orig")

# 检查 Current User stream 的差异
print(f"\n=== Current User stream diff ===")
enc_cu = enc_ole.openstream('Current User').read()
orig_cu = orig_ole.openstream('Current User').read()
print(f"  enc CU len: {len(enc_cu)}, orig CU len: {len(orig_cu)}")
for i in range(min(len(enc_cu), len(orig_cu))):
    if enc_cu[i] != orig_cu[i]:
        print(f"  diff at byte {i}: enc={hex(enc_cu[i])} orig={hex(orig_cu[i])}")

enc_ole.close()
orig_ole.close()
