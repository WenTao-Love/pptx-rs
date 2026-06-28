# -*- coding: utf-8 -*-
"""详细解析 .ppt 文件中的 SpContainer 结构，了解如何构造水印。"""
import struct
import olefile

def parse_record_header(data, offset):
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from('<H', data, offset)[0]
    rec_type = struct.unpack_from('<H', data, offset + 2)[0]
    rec_len = struct.unpack_from('<I', data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    is_container = (ver == 0xF)
    return (ver, inst, rec_type, rec_len, is_container)

# OfficeArt record type 名称映射
OFFICE_ART_NAMES = {
    0xF000: 'DggContainer',
    0xF001: 'DgContainer',
    0xF003: 'SpgrContainer',
    0xF004: 'SpContainer',
    0xF005: 'SolverContainer',
    0xF006: 'FDG',
    0xF007: 'FSPGR',
    0xF008: 'FSP',
    0xF009: 'FOPT',
    0xF00A: 'FOT',
    0xF00B: 'FOPT',
    0xF00D: 'ClientTextbox',
    0xF00E: 'ClientData',
    0xF00F: 'ChildAnchor',
    0xF010: 'ClientAnchor',
    0xF011: 'ClientRule',
    0xF012: 'FOPTE',
    0xF013: 'FOPTEComplex',
}

# OfficeArt 属性 ID 名称映射（部分）
OPT_NAMES = {
    0x0040: 'transform',
    0x0080: 'wzName',
    0x0081: 'wzDescription',
    0x0083: 'wzProgId',
    0x00BF: 'txflTextDir',
    0x0140: 'rotation',
    0x0141: 'cx',
    0x0142: 'cy',
    0x0143: 'shapePath',
    0x0150: 'fillColor',
    0x0180: 'fillBackColor',
    0x0181: 'fillCrMod',
    0x0183: 'fillBackCrMod',
    0x0184: 'blipMod',
    0x0185: 'blip',
    0x0186: 'fillType',
    0x0187: 'fillWidth',
    0x0188: 'fillHeight',
    0x0189: 'fillDztype',
    0x018A: 'fillRectLeft',
    0x018B: 'fillRectTop',
    0x018C: 'fillRectRight',
    0x018D: 'fillRectBottom',
    0x018F: 'fNoFillHitTest',
    0x0190: 'lineColor',
    0x0191: 'lineWidth',
    0x0192: 'lineStyle',
    0x0193: 'lineDashing',
    0x0194: 'lineDashStyle',
    0x0195: 'lineStartArrowhead',
    0x0196: 'lineEndArrowhead',
    0x0197: 'lineStartArrowWidth',
    0x0198: 'lineStartArrowLength',
    0x0199: 'lineEndArrowWidth',
    0x019A: 'lineEndArrowLength',
    0x019B: 'lineJoinStyle',
    0x019C: 'lineEndCapStyle',
    0x019D: 'fNoLineDrawDash',
    0x019F: 'fLine',
    0x01A0: 'shadowType',
    0x01A1: 'shadowColor',
    0x01A2: 'shadowHighlight',
    0x01A3: 'shadowCrMod',
    0x01A4: 'shadowHighlightCrMod',
    0x01A5: 'shadowOpacity',
    0x01A6: 'shadowOffsetX',
    0x01A7: 'shadowOffsetY',
    0x01A8: 'shadowSecondOffsetX',
    0x01A9: 'shadowSecondOffsetY',
    0x01AA: 'shadowScaleXToX',
    0x01AB: 'shadowScaleYToY',
    0x01AC: 'shadowScaleXToX',
    0x01AD: 'shadowScaleYToY',
    0x01AE: 'shadowPerspectiveX',
    0x01AF: 'shadowPerspectiveY',
    0x01B0: 'shadowOriginX',
    0x01B1: 'shadowOriginY',
    0x01B2: 'fShadow',
    0x01B3: 'fshadowObscured',
    0x0200: 'puzzle',
    0x0201: 'groupShadow',
    0x0204: 'groupSelect',
    0x0205: 'groupHidden',
    0x0206: 'groupShapeName',
    0x0208: 'groupDescription',
    0x020B: 'groupShapeId',
    0x020C: 'groupFlags',
    0x0300: 'shapeId',
    0x0301: 'shapeName',
    0x0302: 'shapeDescription',
    0x0303: 'shapeProgId',
    0x0304: 'shapeFlags',
    0x0380: 'lTxid',
    0x0381: 'dxTextLeft',
    0x0382: 'dyTextTop',
    0x0383: 'dxTextRight',
    0x0384: 'dyTextBottom',
    0x0385: 'wrapText',
    0x0386: 'anchorText',
    0x0387: 'textDirection',
    0x0388: 'cxTextMargin',
    0x0389: 'cyTextMargin',
    0x038A: 'fAutoTextMargin',
    0x038B: 'txSpacingBefore',
    0x038C: 'txSpacingAfter',
    0x038D: 'txSpacingLine',
    0x038E: 'txSpacingLineRule',
    0x038F: 'txAutospaceMINDoubleByte',
    0x0390: 'txAutospaceMINLanguage',
    0x0391: 'txAutospaceMINNumeric',
    0x0392: 'txAutospaceDEndDoubleByte',
    0x0393: 'txAutospaceDEndLanguage',
    0x0394: 'txAutospaceDEndNumeric',
    0x0395: 'txScaleText',
    0x0396: 'txRelativeSize',
    0x0397: 'txIndent',
    0x0398: 'txBorderLeft',
    0x0399: 'txBorderTop',
    0x039A: 'txBorderRight',
    0x039B: 'txBorderBottom',
    0x039C: 'txVerticalTextAlignment',
    0x039D: 'txHorizontalTextAlignment',
    0x039E: 'txTextOrientation',
    0x039F: 'txFlow',
    0x03A0: 'txDirection',
    0x03A1: 'txRotation',
    0x03A2: 'txOrientation',
    0x03A3: 'txShadow',
    0x03A4: 'txFill',
    0x03A5: 'txLine',
    0x03A6: 'txEffect',
    0x03A7: 'txEffectExtent',
    0x03A8: 'txEffectExtentLeft',
    0x03A9: 'txEffectExtentTop',
    0x03AA: 'txEffectExtentRight',
    0x03AB: 'txEffectExtentBottom',
    0x03AC: 'txEffectExtentDepth',
    0x03AD: 'txEffectExtentDepthDirection',
    0x03AE: 'txEffectExtentDepthAlignment',
    0x03AF: 'txEffectExtentDepthOffset',
    0x03B0: 'txEffectExtentDepthSpacing',
    0x03B1: 'txEffectExtentDepthSpacingRule',
    0x03B2: 'txEffectExtentDepthSpacingFont',
    0x03B3: 'txEffectExtentDepthSpacingFontRule',
    0x03B4: 'txEffectExtentDepthSpacingFontScale',
    0x03B5: 'txEffectExtentDepthSpacingFontScaleRule',
    0x03B6: 'txEffectExtentDepthSpacingFontOrigin',
    0x03B7: 'txEffectExtentDepthSpacingFontOriginRule',
    0x03B8: 'txEffectExtentDepthSpacingFontOriginOffset',
    0x03B9: 'txEffectExtentDepthSpacingFontOriginOffsetRule',
    0x03BA: 'txEffectExtentDepthSpacingFontOriginOffsetScale',
    0x03BB: 'txEffectExtentDepthSpacingFontOriginOffsetScaleRule',
    0x03BC: 'txEffectExtentDepthSpacingFontOriginOffsetScaleOrigin',
    0x03BD: 'txEffectExtentDepthSpacingFontOriginOffsetScaleOriginRule',
    0x03BE: 'txEffectExtentDepthSpacingFontOriginOffsetScaleOriginOffset',
    0x03BF: 'txEffectExtentDepthSpacingFontOriginOffsetScaleOriginOffsetRule',
    0x0380: 'lTxid',
    0x0381: 'dxTextLeft',
    0x0382: 'dyTextTop',
    0x0383: 'dxTextRight',
    0x0384: 'dyTextBottom',
    0x0385: 'wrapText',
    0x0386: 'anchorText',
    0x0387: 'textDirection',
    0x0388: 'cxTextMargin',
    0x0389: 'cyTextMargin',
    0x038A: 'fAutoTextMargin',
    0x038B: 'txSpacingBefore',
    0x038C: 'txSpacingAfter',
    0x038D: 'txSpacingLine',
    0x038E: 'txSpacingLineRule',
    0x038F: 'txAutospaceMINDoubleByte',
    0x0390: 'txAutospaceMINLanguage',
    0x0391: 'txAutospaceMINNumeric',
    0x0392: 'txAutospaceDEndDoubleByte',
    0x0393: 'txAutospaceDEndLanguage',
    0x0394: 'txAutospaceDEndNumeric',
    0x0395: 'txScaleText',
    0x0396: 'txRelativeSize',
    0x0397: 'txIndent',
    0x0380: 'lTxid',
    0x0381: 'dxTextLeft',
    0x0382: 'dyTextTop',
    0x0383: 'dxTextRight',
    0x0384: 'dyTextBottom',
    0x0385: 'wrapText',
    0x0386: 'anchorText',
    0x0387: 'textDirection',
    0x0388: 'cxTextMargin',
    0x0389: 'cyTextMargin',
    0x038A: 'fAutoTextMargin',
    0x038B: 'txSpacingBefore',
    0x038C: 'txSpacingAfter',
    0x038D: 'txSpacingLine',
    0x038E: 'txSpacingLineRule',
    0x038F: 'txAutospaceMINDoubleByte',
    0x0390: 'txAutospaceMINLanguage',
    0x0391: 'txAutospaceMINNumeric',
    0x0392: 'txAutospaceDEndDoubleByte',
    0x0393: 'txAutospaceDEndLanguage',
    0x0394: 'txAutospaceDEndNumeric',
    0x0395: 'txScaleText',
    0x0396: 'txRelativeSize',
    0x0397: 'txIndent',
    0x0380: 'lTxid',
    0x0381: 'dxTextLeft',
    0x0382: 'dyTextTop',
    0x0383: 'dxTextRight',
    0x0384: 'dyTextBottom',
    0x0385: 'wrapText',
    0x0386: 'anchorText',
    0x0387: 'textDirection',
    0x0388: 'cxTextMargin',
    0x0389: 'cyTextMargin',
    0x038A: 'fAutoTextMargin',
    0x038B: 'txSpacingBefore',
    0x038C: 'txSpacingAfter',
    0x038D: 'txSpacingLine',
    0x038E: 'txSpacingLineRule',
    0x038F: 'txAutospaceMINDoubleByte',
    0x0390: 'txAutospaceMINLanguage',
    0x0391: 'txAutospaceMINNumeric',
    0x0392: 'txAutospaceDEndDoubleByte',
    0x0393: 'txAutospaceDEndLanguage',
    0x0394: 'txAutospaceDEndNumeric',
    0x0395: 'txScaleText',
    0x0396: 'txRelativeSize',
    0x0397: 'txIndent',
    0x0380: 'lTxid',
    0x0381: 'dxTextLeft',
    0x0382: 'dyTextTop',
    0x0383: 'dxTextRight',
    0x0384: 'dyTextBottom',
    0x0385: 'wrapText',
    0x0386: 'anchorText',
    0x0387: 'textDirection',
    0x0388: 'cxTextMargin',
    0x0389: 'cyTextMargin',
    0x038A: 'fAutoTextMargin',
    0x038B: 'txSpacingBefore',
    0x038C: 'txSpacingAfter',
    0x038D: 'txSpacingLine',
    0x038E: 'txSpacingLineRule',
    0x038F: 'txAutospaceMINDoubleByte',
    0x0390: 'txAutospaceMINLanguage',
    0x0391: 'txAutospaceMINNumeric',
    0x0392: 'txAutospaceDEndDoubleByte',
    0x0393: 'txAutospaceDEndLanguage',
    0x0394: 'txAutospaceDEndNumeric',
    0x0395: 'txScaleText',
    0x0396: 'txRelativeSize',
    0x0397: 'txIndent',
}

def dump_spcontainer(data, offset, indent=0):
    """详细解析 SpContainer 的内容。"""
    hdr = parse_record_header(data, offset)
    if hdr is None:
        return
    ver, inst, rec_type, rec_len, is_container = hdr
    name = OFFICE_ART_NAMES.get(rec_type, f'Unknown(0x{rec_type:04X})')
    prefix = "  " * indent

    if rec_type == 0xF00B:  # FOPT
        # 解析 FOPT 属性
        num_props = inst  # inst 字段包含属性数量
        print(f"{prefix}{name} (offset={offset}, len={rec_len}, {num_props} props)")
        pos = offset + 8
        complex_data_start = pos + num_props * 6
        for i in range(num_props):
            if pos + 6 > offset + 8 + rec_len:
                break
            opid = struct.unpack_from('<H', data, pos)[0]
            opid_val = opid & 0x3FFF
            opid_is_complex = (opid & 0x8000) != 0
            op = struct.unpack_from('<I', data, pos + 2)[0]
            opt_name = OPT_NAMES.get(opid_val, f'Unknown(0x{opid_val:04X})')
            if opid_is_complex:
                # complex 属性，op 是 complex 数据的偏移
                complex_data = data[complex_data_start + op:complex_data_start + op + 20]
                try:
                    complex_str = complex_data.decode('utf-16-le', errors='replace').rstrip('\x00')
                except:
                    complex_str = complex_data.hex()
                print(f"{prefix}  [{i}] {opt_name} (0x{opid_val:04X}) = complex, offset={op}, data={complex_str!r}")
            else:
                print(f"{prefix}  [{i}] {opt_name} (0x{opid_val:04X}) = 0x{op:08X} ({op})")
            pos += 6
    elif rec_type == 0xF00A:  # FSP
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len >= 8:
            spid = struct.unpack_from('<I', data, offset + 8)[0]
            flags = struct.unpack_from('<I', data, offset + 12)[0]
            print(f"{prefix}  shapeId={spid}, flags=0x{flags:08X}")
    elif rec_type == 0xF00F:  # ChildAnchor
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len >= 16:
            left, top, right, bottom = struct.unpack_from('<iiii', data, offset + 8)
            print(f"{prefix}  left={left}, top={top}, right={right}, bottom={bottom}")
    elif rec_type == 0xF00D:  # ClientTextbox
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        # ClientTextbox 包含文本内容，通常是 TextHeaderAtom + TextCharsAtom 等
        if rec_len > 0:
            print(f"{prefix}  data (first 100 bytes): {data[offset+8:offset+8+min(100, rec_len)].hex()}")
    elif is_container:
        print(f"{prefix}{name} (offset={offset}, len={rec_len}, container)")
        pos = offset + 8
        end = offset + 8 + rec_len
        while pos + 8 <= end:
            child_hdr = parse_record_header(data, pos)
            if child_hdr is None:
                break
            dump_spcontainer(data, pos, indent + 1)
            pos += 8 + child_hdr[3]
    else:
        print(f"{prefix}{name} (offset={offset}, len={rec_len})")
        if rec_len > 0 and rec_len <= 32:
            print(f"{prefix}  data: {data[offset+8:offset+8+rec_len].hex()}")

def main():
    ppt_path = "_test/心理账户理论.ppt"
    ole = olefile.OleFileIO(ppt_path)
    ppt_data = ole.openstream("PowerPoint Document").read()

    # 解析第一个 PPDrawing 中的最后一个 SpContainer（offset=23663, len=72）
    print("=== 解析 SpContainer at offset=23663, len=72 ===")
    dump_spcontainer(ppt_data, 23663)

    print("\n=== 解析第一个 PPDrawing 中的第一个 SpContainer（offset=8968, len=40）===")
    dump_spcontainer(ppt_data, 8968)

    print("\n=== 解析第一个 PPDrawing 中的第二个 SpContainer（offset=9016, len=1275）===")
    dump_spcontainer(ppt_data, 9016)

    ole.close()

if __name__ == "__main__":
    main()
