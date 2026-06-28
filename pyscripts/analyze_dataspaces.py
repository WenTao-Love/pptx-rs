"""详细分析 msoffcrypto 生成的加密文件中 DataSpaces 的内容。"""
import olefile
import struct

path = "_test_out/msoffcrypto_encrypted.pptx"
ole = olefile.OleFileIO(path)

for stream_path in ole.listdir():
    name = "/".join(stream_path)
    data = ole.openstream(stream_path).read()
    print(f"Stream: {name}")
    print(f"  Size: {len(data)} bytes")

    if name == "\x06DataSpaces/Version":
        print(f"  Hex: {data.hex()}")
        # Version stream format: 4 bytes reserved + 4 bytes version
        if len(data) >= 8:
            reserved = struct.unpack_from("<I", data, 0)[0]
            version = struct.unpack_from("<I", data, 4)[0]
            print(f"  Reserved: {reserved}, Version: {version}")

    elif name == "\x06DataSpaces/DataSpaceMap":
        # DataSpaceMap format
        if len(data) >= 12:
            entry_length = struct.unpack_from("<I", data, 0)[0]
            num_entries = struct.unpack_from("<I", data, 4)[0]
            print(f"  Entry length: {entry_length}, Num entries: {num_entries}")
            offset = 8
            for i in range(num_entries):
                if offset + 8 > len(data):
                    break
                ds_entry_len = struct.unpack_from("<I", data, offset)[0]
                num_components = struct.unpack_from("<I", data, offset + 4)[0]
                print(f"  Entry {i}: length={ds_entry_len}, components={num_components}")
                # Read component names
                comp_offset = offset + 8
                for j in range(num_components):
                    if comp_offset + 4 > len(data):
                        break
                    name_len = struct.unpack_from("<I", data, comp_offset)[0]
                    comp_offset += 4
                    name_bytes = data[comp_offset:comp_offset + name_len]
                    # UTF-16LE, null-terminated
                    try:
                        comp_name = name_bytes.decode("utf-16-le").rstrip("\x00")
                    except:
                        comp_name = repr(name_bytes)
                    print(f"    Component {j}: {comp_name}")
                    comp_offset += name_len
                # Read DataSpace name
                if comp_offset + 4 <= len(data):
                    ds_name_len = struct.unpack_from("<I", data, comp_offset)[0]
                    comp_offset += 4
                    ds_name_bytes = data[comp_offset:comp_offset + ds_name_len]
                    try:
                        ds_name = ds_name_bytes.decode("utf-16-le").rstrip("\x00")
                    except:
                        ds_name = repr(ds_name_bytes)
                    print(f"    DataSpace: {ds_name}")

    elif "DataSpaceInfo" in name:
        print(f"  Hex: {data[:100].hex()}")
        # DataSpaceInfo: 4 bytes num entries + entries
        if len(data) >= 4:
            num_entries = struct.unpack_from("<I", data, 0)[0]
            print(f"  Num entries: {num_entries}")
            offset = 4
            for i in range(num_entries):
                if offset + 4 > len(data):
                    break
                entry_len = struct.unpack_from("<I", data, offset)[0]
                entry_data = data[offset + 4:offset + 4 + entry_len]
                try:
                    entry_text = entry_data.decode("utf-16-le").rstrip("\x00")
                except:
                    entry_text = repr(entry_data)
                print(f"  Entry {i}: {entry_text}")
                offset += 4 + entry_len

    elif "TransformInfo" in name:
        print(f"  Hex: {data[:200].hex()}")
        # TransformInfo format
        if len(data) >= 12:
            # Read as much as we can
            print(f"  First 100 bytes as text: {data[:100]}")

    elif name == "EncryptionInfo":
        # Skip first 8 bytes (version), then XML
        ei_xml = data[8:].decode("utf-8", errors="replace")
        print(f"  XML:\n{ei_xml}")

    print()

ole.close()
