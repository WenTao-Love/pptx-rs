"""精确解析 msoffcrypto 生成的 DataSpaces 二进制格式。"""
import olefile
import struct

path = "_test_out/msoffcrypto_encrypted.pptx"
ole = olefile.OleFileIO(path)

# Version
ver = ole.openstream("\x06DataSpaces/Version").read()
print("=== Version ===")
print(f"Total: {len(ver)} bytes")
print(f"Hex: {ver.hex()}")
# Parse: string_len(4) + string + version fields
offset = 0
str_len = struct.unpack_from("<I", ver, offset)[0]
offset += 4
name = ver[offset:offset+str_len].decode("utf-16-le").rstrip("\x00")
offset += str_len
print(f"  String len: {str_len}, Name: {name}")
# Remaining fields
while offset < len(ver):
    val = struct.unpack_from("<I", ver, offset)[0]
    print(f"  Offset {offset}: {val} (0x{val:08X})")
    offset += 4

# DataSpaceMap
dsm = ole.openstream("\x06DataSpaces/DataSpaceMap").read()
print("\n=== DataSpaceMap ===")
print(f"Total: {len(dsm)} bytes")
print(f"Hex: {dsm.hex()}")
offset = 0
entry_len = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
num_entries = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
print(f"  entryLength: {entry_len}, numEntries: {num_entries}")
for i in range(num_entries):
    ds_entry_len = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
    num_comp = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
    print(f"  Entry {i}: length={ds_entry_len}, components={num_comp}")
    for j in range(num_comp):
        comp_len = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
        comp_name = dsm[offset:offset+comp_len].decode("utf-16-le").rstrip("\x00")
        offset += comp_len
        print(f"    Component {j}: len={comp_len}, name={comp_name}")
    ds_name_len = struct.unpack_from("<I", dsm, offset)[0]; offset += 4
    ds_name = dsm[offset:offset+ds_name_len].decode("utf-16-le").rstrip("\x00")
    offset += ds_name_len
    print(f"    DataSpace: len={ds_name_len}, name={ds_name}")

# DataSpaceInfo
dsi = ole.openstream("\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace").read()
print("\n=== DataSpaceInfo ===")
print(f"Total: {len(dsi)} bytes")
print(f"Hex: {dsi.hex()}")
offset = 0
entry_len = struct.unpack_from("<I", dsi, offset)[0]; offset += 4
num_entries = struct.unpack_from("<I", dsi, offset)[0]; offset += 4
print(f"  entryLength: {entry_len}, numEntries: {num_entries}")
for i in range(num_entries):
    name_len = struct.unpack_from("<I", dsi, offset)[0]; offset += 4
    name = dsi[offset:offset+name_len].decode("utf-16-le").rstrip("\x00")
    offset += name_len
    print(f"  Entry {i}: len={name_len}, name={name}")

# TransformInfo
dst = ole.openstream("\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary").read()
print("\n=== TransformInfo ===")
print(f"Total: {len(dst)} bytes")
print(f"Hex: {dst.hex()}")
offset = 0
entry_len = struct.unpack_from("<I", dst, offset)[0]; offset += 4
num_entries = struct.unpack_from("<I", dst, offset)[0]; offset += 4
print(f"  entryLength: {entry_len}, numEntries: {num_entries}")
clsid_len = struct.unpack_from("<I", dst, offset)[0]; offset += 4
clsid = dst[offset:offset+clsid_len].decode("utf-16-le").rstrip("\x00")
offset += clsid_len
print(f"  CLSID: len={clsid_len}, name={clsid}")
transform_len = struct.unpack_from("<I", dst, offset)[0]; offset += 4
transform = dst[offset:offset+transform_len].decode("utf-16-le").rstrip("\x00")
offset += transform_len
print(f"  Transform: len={transform_len}, name={transform}")
while offset < len(dst):
    val = struct.unpack_from("<I", dst, offset)[0]
    print(f"  Offset {offset}: {val} (0x{val:08X})")
    offset += 4

ole.close()
