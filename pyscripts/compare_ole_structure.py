#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对比原始文件和加密文件的 OLE2 容器结构，找出容器层面的差异。
"""

import olefile


def dump_ole_structure(path, label):
    """Dump OLE2 container structure."""
    print(f"\n{'='*60}")
    print(f"{label}: {path}")
    print(f"{'='*60}")

    ole = olefile.OleFileIO(path)

    # 基本信息
    print(f"\n基本信息:")
    print(f"  Sector size: {ole.sectorsize}")
    print(f"  Mini sector cutoff: {ole.minisectorcutoff}")
    print(f"  Num sectors: {ole.nb_sect}")
    print(f"  Num FAT sectors: {ole.num_fat_sectors}")
    print(f"  Num mini FAT sectors: {ole.num_mini_fat_sectors}")
    print(f"  Num DIFAT sectors: {ole.num_difat_sectors}")

    # 所有 streams
    print(f"\nStreams:")
    for entry in ole.listdir(streams=True):
        name = "/".join(entry)
        try:
            data = ole.openstream(entry).read()
            print(f"  {name}: {len(data)} bytes")
        except Exception as e:
            print(f"  {name}: ERROR - {e}")

    # 目录条目
    print(f"\n目录条目:")
    for sid in range(len(ole.direntries)):
        e = ole.direntries[sid]
        if e is None:
            continue
        name = e.name
        if isinstance(name, bytes):
            name = name.decode('utf-8', errors='replace')
        etype = {0: 'unknown', 1: 'storage', 2: 'stream', 5: 'root'}.get(e.entry_type, str(e.entry_type))
        sid_start = getattr(e, 'start', getattr(e, 'sid_start', '?'))
        print(f"  SID={sid:>2} name={name:<40} type={etype:<8} size={e.size:>10}")

    ole.close()


def main():
    orig_path = "_test/心理账户理论.ppt"
    enc_path = "_test_out/protected_心理账户理论.ppt"

    dump_ole_structure(orig_path, "原始文件")
    dump_ole_structure(enc_path, "加密文件")

    # 对比 stream 内容
    print(f"\n{'='*60}")
    print(f"Stream 内容对比:")
    print(f"{'='*60}")

    ole_o = olefile.OleFileIO(orig_path)
    ole_e = olefile.OleFileIO(enc_path)

    orig_streams = set()
    for entry in ole_o.listdir(streams=True):
        orig_streams.add("/".join(entry))
    enc_streams = set()
    for entry in ole_e.listdir(streams=True):
        enc_streams.add("/".join(entry))

    print(f"\n原始文件 streams: {sorted(orig_streams)}")
    print(f"加密文件 streams: {sorted(enc_streams)}")
    print(f"仅在原始文件: {orig_streams - enc_streams}")
    print(f"仅在加密文件: {enc_streams - orig_streams}")

    for name in sorted(orig_streams & enc_streams):
        o_data = ole_o.openstream(name).read()
        e_data = ole_e.openstream(name).read()
        if len(o_data) == len(e_data):
            if o_data == e_data:
                status = "完全相同"
            else:
                status = f"大小相同但内容不同 ({len(o_data)} bytes)"
        else:
            status = f"大小不同 (原始={len(o_data)} 加密={len(e_data)})"
        print(f"  {name}: {status}")

    ole_o.close()
    ole_e.close()


if __name__ == "__main__":
    main()
