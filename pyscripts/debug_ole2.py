"""调试 OLE2 文件的目录条目"""
import struct

with open('_test_out/py_agile_encrypted.pptx', 'rb') as f:
    data = f.read()

# 读取 header
signature = data[0:8]
print(f'Signature: {signature.hex()}')
fat_sectors = struct.unpack_from('<I', data, 44)[0]
dir_sector_id = struct.unpack_from('<I', data, 48)[0]
print(f'FAT sectors: {fat_sectors}')
print(f'Dir sector: {dir_sector_id}')

# 读取目录 sector
dir_offset = 512 + dir_sector_id * 512
print(f'Dir offset: {dir_offset}')

for i in range(4):
    entry_offset = dir_offset + i * 128
    if entry_offset + 128 > len(data):
        break
    name_size = struct.unpack_from('<H', data, entry_offset + 64)[0]
    obj_type = data[entry_offset + 66]
    left_did = struct.unpack_from('<I', data, entry_offset + 68)[0]
    right_did = struct.unpack_from('<I', data, entry_offset + 72)[0]
    child_did = struct.unpack_from('<I', data, entry_offset + 76)[0]
    start_sector = struct.unpack_from('<I', data, entry_offset + 116)[0]
    size = struct.unpack_from('<I', data, entry_offset + 120)[0]

    name_bytes = data[entry_offset:entry_offset+min(name_size, 64)]
    try:
        name = name_bytes.decode('utf-16-le').rstrip('\x00')
    except:
        name = repr(name_bytes)

    print(f'Entry {i}: name="{name}", type={obj_type}, left={left_did}, right={right_did}, '
          f'child={child_did}, start_sector={start_sector}, size={size}')

# 检查 FAT 中 EncryptionInfo 的链
# 先找 enc_info_start
# 从目录条目1的 start_sector 读取
enc_info_sector = struct.unpack_from('<I', data, dir_offset + 128 + 116)[0]
print(f'\nEncryptionInfo start sector: {enc_info_sector}')

# 读取 FAT 中该 sector 的值
fat_offset = 512  # FAT sector 0 starts at offset 512
fat_entry = struct.unpack_from('<I', data, fat_offset + enc_info_sector * 4)[0]
print(f'FAT[{enc_info_sector}] = {fat_entry} (0x{fat_entry:08X})')

# 读取 EncryptionInfo 数据
enc_info_offset = 512 + enc_info_sector * 512
enc_info_data = data[enc_info_offset:enc_info_offset+100]
print(f'EncryptionInfo data (first 100 bytes): {enc_info_data[:50]}')
