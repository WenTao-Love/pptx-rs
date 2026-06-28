#!/usr/bin/env python3
"""检查 OLE2 容器中所有 stream 是否能正常读取。"""
import olefile

for path in ["_test/心理账户理论.ppt", "_test_out/protected_心理账户理论.ppt", "_test_out/wm_protected_心理账户理论.ppt"]:
    print(f"\n{path}:")
    try:
        ole = olefile.OleFileIO(path)
        for stream_path in ole.listdir():
            name = "/".join(stream_path)
            try:
                data = ole.openstream(stream_path).read()
                print(f"  OK: {name!r} ({len(data)} bytes)")
            except Exception as e:
                print(f"  FAIL: {name!r}: {e}")
        ole.close()
    except Exception as e:
        print(f"  OPEN FAIL: {e}")
