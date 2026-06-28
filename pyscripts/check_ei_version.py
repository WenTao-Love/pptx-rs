"""检查 Rust 生成的 EncryptionInfo stream 前 8 字节。"""
import olefile
import struct

ole = olefile.OleFileIO("_test_out/protected_文旅IP人设打造抖音短视频运营方案.pptx")
ei = ole.openstream("EncryptionInfo").read()
print(f"First 16 bytes hex: {ei[:16].hex()}")
vmaj, vmin = struct.unpack_from("<HH", ei, 0)
print(f"MajorVersion: {vmaj}, MinorVersion: {vmin}")
# 如果 version 是 4.0，说明 cfb crate 没有覆盖，而是我们写入的数据本身有问题
# 检查我们写入的前 8 字节
print(f"Bytes 0-3: {ei[0:4].hex()} (should be 04000000)")
print(f"Bytes 4-7: {ei[4:8].hex()} (should be 04000000)")
ole.close()
