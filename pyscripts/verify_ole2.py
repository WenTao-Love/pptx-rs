import olefile
ole = olefile.OleFileIO('_test_out/py_agile_encrypted.pptx')
print('Streams:', ole.listdir())
for stream in ole.listdir():
    name = '/'.join(stream)
    data = ole.openstream(stream).read()
    print(f'  {name}: {len(data)} bytes')
    if name == 'EncryptionInfo' and len(data) > 0:
        text = data[:300].decode('utf-8', errors='replace')
        print(f'  Content: {text}')
ole.close()
