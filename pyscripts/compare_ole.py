#!/usr/bin/env python3
"""比较原始文件和加密文件的 OLE2 结构差异。"""
import olefile
import struct

def ole_info(path, label):
    print(f"\n{label}: {path}")
    ole = olefile.OleFileIO(path)
    print(f"  streams: {ole.listdir()}")
    print(f"  total_size: {len(ole.openstream('PowerPoint Document').read())}")
    print(f"  mini_fat_sectors: {ole.num_mini_fat_sectors}")
    print(f"  num_fat_sectors: {ole.num_fat_sectors}")
    print(f"  first_mini_fat_sector: {ole.first_mini_fat_sector}")
    print(f"  sector_size: {ole.sector_size}")
    ole.close()

ole_info("_test/心理账户理论.ppt", "原始")
ole_info("_test_out/protected_心理账户理论.ppt", "加密")
ole_info("_test_out/wm_protected_心理账户理论.ppt", "水印+加密")
ole_info("_test_out/rc4cryptoapi_password.ppt", "参考")

print("\n\n原始文件 Current User 字节:")
with open("_test/心理账户理论.ppt", "rb") as f:
    ole = olefile.OleFileIO(f)
    cu = ole.openstream("Current User").read()
    print(cu[:64].hex())
    ole.close()

print("\n加密文件 Current User 字节:")
with open("_test_out/protected_心理账户理论.ppt", "rb") as f:
    ole = olefile.OleFileIO(f)
    cu = ole.openstream("Current User").read()
    print(cu[:64].hex())
    ole.close()

print("\n参考文件 Current User 字节:")
with open("_test_out/rc4cryptoapi_password.ppt", "rb") as f:
    ole = olefile.OleFileIO(f)
    cu = ole.openstream("Current User").read()
    print(cu[:64].hex())
    ole.close()
