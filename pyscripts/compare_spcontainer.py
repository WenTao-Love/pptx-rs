#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对比分析原始 .ppt 文件与注入水印后的 .ppt 文件中 SpContainer 的结构差异。

目的：
  在 .ppt 文件中注入水印 SpContainer 后，水印不显示。
  本脚本通过对比原始文件中 PowerPoint 原生生成的 SpContainer 结构
  与我们注入的水印 SpContainer 结构，找出导致水印不显示的差异。

流程：
  1. 用 olefile 打开 .ppt，读取 "PowerPoint Document" stream
  2. 找到第一个 Slide container (recType=0x03EE)
     注意：0x03F8 是 RT_MainMaster（主母版），不是 Slide
  3. 在 Slide 中找到 PPDrawing (recType=0x040C)
  4. 在 PPDrawing 中找到 SpgrContainer (0xF002)
  5. 遍历 SpgrContainer 中的所有 SpContainer (0xF003)
  6. 对每个 SpContainer 解析其子 record（FSP/FOPT/ClientAnchor/ClientTextbox 等）
  7. 对比原始文件与水印文件的 SpContainer 结构差异
"""

import sys
import io
import struct
import olefile

# 全局输出缓冲：同时输出到 stdout 和 UTF-8 文件
_OUTPUT_LINES = []

def out(msg=""):
    """统一输出函数：同时输出到 stdout 和全局缓冲（用于写 UTF-8 文件）。

    使用 flush=True 避免输出缓冲。
    """
    sys.stdout.write(str(msg) + "\n")
    sys.stdout.flush()
    _OUTPUT_LINES.append(str(msg))

def save_output(filepath):
    """把全局缓冲的内容写到 UTF-8 文件。"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(_OUTPUT_LINES))
        f.write("\n")


# ============================================================================
# Record type 常量（MS-PPT / MS-ODRAW 规范）
# ============================================================================

# MS-PPT 顶层 record types
RT_SLIDE = 0x03EE              # RT_Slide（幻灯片，注意不是 0x03F8）
RT_MAIN_MASTER = 0x03F8        # RT_MainMaster（主母版，仅用于参考）
RT_PPDRAWING = 0x040C          # RT_PPDrawing（PPDrawing container）

# MS-ODRAW OfficeArt record types
RT_SPGR_CONTAINER = 0xF002     # OfficeArtSpgrContainer
RT_SP_CONTAINER = 0xF003       # OfficeArtSpContainer
RT_F004 = 0xF004               # OfficeArtSpContainer 变体（PowerPoint 97-2003 常用）
RT_FSP = 0xF007                # OfficeArtFSP（形状属性 atom）
RT_FOPT = 0xF008               # OfficeArtFOPT（形状选项 atom）
RT_FSPGR = 0xF009              # OfficeArtFSPGR（形状组属性 atom）
RT_CLIENT_ANCHOR = 0xF00A      # OfficeArtClientAnchor
RT_CLIENT_DATA = 0xF00B        # OfficeArtClientData
RT_CLIENT_TEXTBOX_ATOM = 0xF009  # 注意：0xF009 既是 FSPGR 又是 ClientTextbox(atom)，按 ver 区分
RT_CLIENT_TEXTBOX = 0xF00D     # OfficeArtClientTextbox（container 版本，ver=0xF）

# PPT Text record types（位于 ClientTextbox 内部）
RT_TEXT_HEADER_ATOM = 0x0F9F   # TextHeaderAtom
RT_TEXT_CHARS_ATOM = 0x0FA0    # TextCharsAtom（UTF-16LE 文本）
RT_TEXT_BYTES_ATOM = 0x0FA8    # TextBytesAtom（Latin-1 文本）
RT_STYLE_ATOM = 0x0FA1         # TextRulerAtom / StyleAtom
RT_TEXT_SPEC_INFO = 0x0FA5     # TextSpecInfoAtom

# OfficeArt record type 名称映射
OFFICEART_TYPES = {
    0xF000: "DggContainer",
    0xF001: "BStoreContainer",
    0xF002: "SpgrContainer",
    0xF003: "SpContainer",
    0xF004: "SpContainer(F004变体)",
    0xF005: "SolverContainer",
    0xF006: "FDGG",
    0xF007: "FSP",
    0xF008: "FOPT",
    0xF009: "FSPGR",
    0xF00A: "ClientAnchor",
    0xF00B: "ClientData",
    0xF00C: "FRITContainer",
    0xF00D: "ClientTextbox(container)",
    0xF00E: "TertiaryFOPT",
    0xF00F: "ChildAnchor",
    0xF010: "FSPGR",
    0xF011: "FConnector",
    0xF012: "FDGGBlock",
    0xF122: "UnknownF122",
}

# MS-PPT 顶层 record type 名称映射（非 OfficeArt）
PPT_TYPES = {
    0x03EE: "Slide",
    0x03F8: "MainMaster",
    0x040C: "PPDrawing",
    0x03F9: "SlideList",
    0x03FF: "UserEditAtom",
    0x0FF5: "UserEditAtom",
    0x0FF6: "CurrentUserAtom",
    0x1772: "PersistDirectoryAtom",
    0x03F3: "SlideAtom",
    0x03F4: "SlideShowSlideInfoAtom",
    0x03FD: "SlideSchemeColorSchemeAtom",
    0x03FE: "SlideName",
    0x0400: "SlideProgTagsContainer",
    0x0401: "SlideProgBinaryTagContainer",
    0x1388: "SlideProgBinaryTagAtom",
    0x00AF: "SlideFlagsAtom(0xAF)",
    0x0141: "SlideFlagsAtom(0x141)",
    0x0080: "OfficeColorRGBAtom",
    0x0081: "OfficeColorCyanAtom",
    0x0082: "OfficeColorMagentaAtom",
    0x0083: "OfficeColorYellowAtom",
    0x0084: "OfficeColorBlackAtom",
}

# FOPT 属性 ID 到名称的映射（MS-ODRAW 规范，仅保留实际常见的属性）
FOPT_PROPS = {
    0x0080: "transform.rotation",
    0x0081: "transform.lockRotation",
    0x00BF: "ProtectionBooleanProperties",
    0x00C0: "shapePath",
    0x0104: "geoLeft",
    0x0105: "geoTop",
    0x0106: "geoRight",
    0x0107: "geoBottom",
    0x0140: "cxform",
    0x0141: "fillColor",
    0x0142: "fillBackColor",
    0x0143: "fillCrMod",
    0x0144: "lineColor",
    0x0145: "lineBackColor",
    0x0146: "lineCrMod",
    0x0147: "fillStyle",
    0x0148: "lineStyle",
    0x0149: "lineWidth",
    0x014A: "lineDashing",
    0x014B: "lineDashStyle",
    0x014C: "fillStyleBooleanProperties",
    0x014D: "shadowColor",
    0x014E: "shadowBackColor",
    0x014F: "shadowCrMod",
    0x0150: "shadowOpacity",
    0x0151: "shadowOffsetX",
    0x0152: "shadowOffsetY",
    0x0153: "shadowStyle",
    0x0154: "shadowStyleBooleanProperties",
    0x0155: "fillBlip",
    0x0156: "fillBlipName",
    0x0157: "fillBlipFlags",
    0x0158: "fillWidth",
    0x0159: "fillHeight",
    0x015A: "fillDztype",
    0x015B: "fillRectLeft",
    0x015C: "fillRectTop",
    0x015D: "fillRectRight",
    0x015E: "fillRectBottom",
    0x015F: "fillAngle",
    0x0160: "fillFocus",
    0x0161: "fillToLeft",
    0x0162: "fillToTop",
    0x0163: "fillToRight",
    0x0164: "fillToBottom",
    0x0165: "rectLeft",
    0x0166: "rectTop",
    0x0167: "rectRight",
    0x0168: "rectBottom",
    0x0169: "pWrapPolygonVertices",
    0x016A: "wrapDistLeft",
    0x016B: "wrapDistTop",
    0x016C: "wrapDistRight",
    0x016D: "wrapDistBottom",
    0x016E: "wrapText",
    0x016F: "dxWrapDistLeft",
    0x0170: "dyWrapDistTop",
    0x0171: "dxWrapDistRight",
    0x0172: "dyWrapDistBottom",
    0x0179: "hspNext",
    0x017A: "hspPrev",
    0x017B: "relativeHorizontalPosition",
    0x017C: "relativeVerticalPosition",
    0x017D: "relativeHorizontalSize",
    0x017E: "relativeVerticalSize",
    0x017F: "metaFileBlip",
    0x0180: "lineBlip",
    0x0181: "lineBlipName",
    0x0182: "lineBlipFlags",
    0x0183: "lineOpacity",
    0x0184: "lineWidth",
    0x0185: "lineStyle",
    0x0186: "fillType",
    0x0187: "fillBlip",
    0x0188: "fillBlipName",
    0x0189: "fillBlipFlags",
    0x018A: "shadowColor",
    0x018B: "shadowBackColor",
    0x018C: "shadowCrMod",
    0x018D: "shadowOpacity",
    0x018E: "shadowOffsetX",
    0x018F: "shadowOffsetY",
    0x0190: "lineColor",
    0x0191: "lineBackColor",
    0x0192: "lineCrMod",
    0x0193: "lineWidth",
    0x0194: "lineStyle",
    0x0195: "lineDashing",
    0x0196: "lineDashStyle",
    0x0197: "lineStyleBooleanProperties",
    0x0198: "lineStartArrowhead",
    0x0199: "lineEndArrowhead",
    0x019A: "lineStartArrowWidth",
    0x019B: "lineStartArrowLength",
    0x019C: "lineEndArrowWidth",
    0x019D: "lineEndArrowLength",
    0x019E: "lineJoinStyle",
    0x019F: "lineEndCapStyle",
    0x01BF: "fNoFillHitTest",
    0x01C0: "fNoLineDrawDash",
    0x01FF: "fNoLineDrawDash",
    0x0200: "lineColor",
    0x0201: "lineBackColor",
    0x0202: "lineCrMod",
    0x0203: "lineWidth",
    0x0204: "lineStyle",
    0x0205: "lineDashing",
    0x0206: "lineDashStyle",
    0x0207: "lineStyleBooleanProperties",
    0x0208: "lineStartArrowhead",
    0x0209: "lineEndArrowhead",
    0x020A: "lineStartArrowWidth",
    0x020B: "lineStartArrowLength",
    0x020C: "lineEndArrowWidth",
    0x020D: "lineEndArrowLength",
    0x020E: "lineJoinStyle",
    0x020F: "lineEndCapStyle",
    0x0304: "shapeFlags",
    0x0305: "shapePath",
    0x033F: "pImageMapProperties",
    0x0380: "Text ID",
    0x0381: "wzName",
    0x0382: "wzDescription",
    0x0383: "pWrapPolygonVertices",
    0x0384: "wzName",
    0x0385: "wzDescription",
    0x0386: "pWrapPolygonVertices",
    0x0387: "wzName",
    0x0388: "wzDescription",
    0x0389: "pWrapPolygonVertices",
    0x038A: "wzName",
    0x038B: "wzDescription",
    0x038C: "pWrapPolygonVertices",
    0x038D: "wzName",
    0x038E: "wzDescription",
    0x038F: "pWrapPolygonVertices",
    0x0390: "wzName",
    0x0391: "wzDescription",
    0x0392: "pWrapPolygonVertices",
    0x0393: "wzName",
    0x0394: "wzDescription",
    0x0395: "pWrapPolygonVertices",
    0x0396: "wzName",
    0x0397: "wzDescription",
    0x0398: "pWrapPolygonVertices",
    0x0399: "pWrapPolygonVertices",
    0x039A: "wzName",
    0x039B: "wzDescription",
    0x039C: "pWrapPolygonVertices",
    0x039D: "wzName",
    0x039E: "wzDescription",
    0x039F: "pWrapPolygonVertices",
    0x03A0: "wzName",
    0x03A1: "wzDescription",
    0x03A2: "pWrapPolygonVertices",
    0x03A3: "wzName",
    0x03A4: "wzDescription",
    0x03A5: "pWrapPolygonVertices",
    0x03A6: "wzName",
    0x03A7: "wzDescription",
    0x03A8: "pWrapPolygonVertices",
    0x03A9: "wzName",
    0x03AA: "wzDescription",
    0x03AB: "pWrapPolygonVertices",
    0x03AC: "wzName",
    0x03AD: "wzDescription",
    0x03AE: "pWrapPolygonVertices",
    0x03AF: "wzName",
    0x03B0: "wzDescription",
    0x03B1: "pWrapPolygonVertices",
    0x03B2: "wzName",
    0x03B3: "wzDescription",
    0x03B4: "pWrapPolygonVertices",
    0x03B5: "wzName",
    0x03B6: "wzDescription",
    0x03B7: "pWrapPolygonVertices",
    0x03B8: "pWrapPolygonVertices",
    0x03B9: "pWrapPolygonVertices",
    0x03BA: "pWrapPolygonVertices",
    0x03BB: "pWrapPolygonVertices",
    0x03BC: "pWrapPolygonVertices",
    0x03BD: "pWrapPolygonVertices",
    0x03BE: "pWrapPolygonVertices",
    0x03BF: "pWrapPolygonVertices",
    0x03C0: "pWrapPolygonVertices",
    0x03C1: "pWrapPolygonVertices",
    0x03C2: "pWrapPolygonVertices",
    0x03C3: "pWrapPolygonVertices",
    0x03C4: "pWrapPolygonVertices",
    0x03C5: "pWrapPolygonVertices",
    0x03C6: "pWrapPolygonVertices",
    0x03C7: "pWrapPolygonVertices",
    0x03C8: "pWrapPolygonVertices",
    0x03C9: "pWrapPolygonVertices",
    0x03CA: "pWrapPolygonVertices",
    0x03CB: "pWrapPolygonVertices",
    0x03CC: "pWrapPolygonVertices",
    0x03CD: "pWrapPolygonVertices",
    0x03CE: "pWrapPolygonVertices",
    0x03CF: "pWrapPolygonVertices",
    0x03D0: "pWrapPolygonVertices",
    0x03D1: "pWrapPolygonVertices",
    0x03D2: "pWrapPolygonVertices",
    0x03D3: "pWrapPolygonVertices",
    0x03D4: "pWrapPolygonVertices",
    0x03D5: "pWrapPolygonVertices",
    0x03D6: "pWrapPolygonVertices",
    0x03D7: "pWrapPolygonVertices",
    0x03D8: "pWrapPolygonVertices",
    0x03D9: "pWrapPolygonVertices",
    0x03DA: "pWrapPolygonVertices",
    0x03DB: "pWrapPolygonVertices",
    0x03DC: "pWrapPolygonVertices",
    0x03DD: "pWrapPolygonVertices",
    0x03DE: "pWrapPolygonVertices",
    0x03DF: "pWrapPolygonVertices",
    0x03E0: "pWrapPolygonVertices",
    0x03E1: "pWrapPolygonVertices",
    0x03E2: "pWrapPolygonVertices",
    0x03E3: "pWrapPolygonVertices",
    0x03E4: "pWrapPolygonVertices",
    0x03E5: "pWrapPolygonVertices",
    0x03E6: "pWrapPolygonVertices",
    0x03E7: "pWrapPolygonVertices",
    0x03E8: "pWrapPolygonVertices",
    0x03E9: "pWrapPolygonVertices",
    0x03EA: "pWrapPolygonVertices",
    0x03EB: "pWrapPolygonVertices",
    0x03EC: "pWrapPolygonVertices",
    0x03ED: "pWrapPolygonVertices",
    0x03EE: "pWrapPolygonVertices",
    0x03EF: "pWrapPolygonVertices",
    0x03F0: "pWrapPolygonVertices",
    0x03F1: "pWrapPolygonVertices",
    0x03F2: "pWrapPolygonVertices",
    0x03F3: "pWrapPolygonVertices",
    0x03F4: "pWrapPolygonVertices",
    0x03F5: "pWrapPolygonVertices",
    0x03F6: "pWrapPolygonVertices",
    0x03F7: "pWrapPolygonVertices",
    0x03F8: "pWrapPolygonVertices",
    0x03F9: "pWrapPolygonVertices",
    0x03FA: "pWrapPolygonVertices",
    0x03FB: "pWrapPolygonVertices",
    0x03FC: "pWrapPolygonVertices",
    0x03FD: "pWrapPolygonVertices",
    0x03FE: "pWrapPolygonVertices",
    0x03FF: "pWrapPolygonVertices",
    0x043F: "tableProperties",
    0x0440: "tableRowProperties",
}

# MSO shape type（FSP inst 字段）名称映射（部分常见类型）
MSO_SHAPE_TYPES = {
    0x00: "NotPrimitive",
    0x01: "Rectangle",
    0x02: "RoundRectangle",
    0x03: "Ellipse",
    0x04: "Diamond",
    0x05: "IsocelesTriangle",
    0x06: "RightTriangle",
    0x07: "Parallelogram",
    0x08: "Trapezoid",
    0x09: "Hexagon",
    0x0A: "Octagon",
    0x0B: "Plus",
    0x0C: "Star",
    0x0D: "Arrow",
    0x14: "Plaque",
    0x1A: "Can",
    0x1B: "Cube",
    0x23: "SmileyFace",
    0x4B: "TextBox",          # 75 - 文本框
    0x14: "Plaque",
    0x4A: "TextPlainText",    # 74
    0x4C: "TextChevron",      # 76
    0x4D: "TextChevronInverted",
    0x4E: "TextRingInside",
    0x4F: "TextRingOutside",
    0x50: "TextArchUp",
    0x51: "TextArchDown",
    0x52: "TextCircle",
    0x53: "TextButton",
    0x54: "TextArchUpPour",
    0x55: "TextArchDownPour",
    0x56: "TextCirclePour",
    0x57: "TextButtonPour",
    0x58: "TextCurveUp",
    0x59: "TextCurveDown",
    0x5A: "TextCascadeUp",
    0x5B: "TextCascadeDown",
    0x5C: "TextWave1",
    0x5D: "TextWave2",
    0x5E: "TextWave3",
    0x5F: "TextWave4",
    0x60: "TextInflate",
    0x61: "TextDeflate",
    0x62: "TextInflateBottom",
    0x63: "TextDeflateBottom",
    0x64: "TextInflateTop",
    0x65: "TextDeflateTop",
    0x66: "TextDeflateInflate",
    0x67: "TextDeflateInflateDeflate",
    0x68: "TextFadeRight",
    0x69: "TextFadeLeft",
    0x6A: "TextFadeUp",
    0x6B: "TextFadeDown",
    0x6C: "TextSlantUp",
    0x6D: "TextSlantDown",
    0x6E: "TextCanUp",
    0x6F: "TextCanDown",
    0x70: "FlowChartProcess",
    0x71: "FlowChartDecision",
    0x72: "FlowChartInputOutput",
    0x73: "FlowChartPredefinedProcess",
    0x74: "FlowChartInternalStorage",
    0x75: "FlowChartDocument",
    0x76: "FlowChartMultidocument",
    0x77: "FlowChartTerminator",
    0x78: "FlowChartPreparation",
    0x79: "FlowChartManualInput",
    0x7A: "FlowChartManualOperation",
    0x7B: "FlowChartConnector",
    0x7C: "FlowChartPunchedCard",
    0x7D: "FlowChartPunchedTape",
    0x7E: "FlowChartSummingJunction",
    0x7F: "FlowChartOr",
    0x80: "FlowChartCollate",
    0x81: "FlowChartSort",
    0x82: "FlowChartExtract",
    0x83: "FlowChartMerge",
    0x84: "FlowChartOfflineStorage",
    0x85: "FlowChartOnlineStorage",
    0x86: "FlowChartMagneticTape",
    0x87: "FlowChartMagneticDisk",
    0x88: "FlowChartMagneticDrum",
    0x89: "FlowChartDisplay",
    0x8A: "FlowChartDelay",
    0x8B: "TextPlainText",
    0x8C: "TextStop",
    0x8D: "TextTriangle",
    0x8E: "TextTriangleInverted",
    0xC8: "PictureFrame",     # 200
    0xCB: "HostControl",      # 203
    0x8F: "TextPlain",
}


# ============================================================================
# 通用解析函数
# ============================================================================

def parse_header(data, offset):
    """解析 8 字节 record header。

    OfficeArt record header 结构（MS-ODRAW 规范 2.3.1）：
      - ver (4 bits): 版本号；container 用 0xF
      - inst (12 bits): 实例号；FOPT 中表示属性数量
      - recType (16 bits): record 类型
      - recLen (32 bytes): 数据长度（不含 header）

    返回 (ver, inst, recType, recLen)，offset 越界返回 None。
    """
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def type_name(rec_type):
    """返回 record type 的可读名称，优先查 OfficeArt，再查 PPT 类型表。"""
    if rec_type in OFFICEART_TYPES:
        return OFFICEART_TYPES[rec_type]
    if rec_type in PPT_TYPES:
        return PPT_TYPES[rec_type]
    return f"Unknown(0x{rec_type:04X})"


def mso_shape_name(inst):
    """返回 FSP inst 字段对应的 MSO 形状类型名称。"""
    return MSO_SHAPE_TYPES.get(inst, f"Unknown(0x{inst:X})")


def find_all_slides(data):
    """在 PowerPoint Document stream 中找到所有 Slide record 的 offset。

    遍历顶层 record，返回所有 type=0x03EE (RT_Slide) 的 container record 位置。
    注意：0x03F8 是 RT_MainMaster（主母版），不是 Slide。
    """
    slides = []
    pos = 0
    while pos + 8 <= len(data):
        h = parse_header(data, pos)
        if h is None:
            break
        ver, _, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SLIDE:
            slides.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return slides


def find_record(data, start, end, target_type, container_only=True):
    """在 [start, end) 范围内找到第一个指定类型的 record。

    参数：
      data: 完整的 PowerPoint Document stream
      start/end: 搜索范围
      target_type: 目标 record type
      container_only: True 时只匹配 container（ver=0xF）
    """
    pos = start
    while pos + 8 <= end:
        h = parse_header(data, pos)
        if h is None:
            return None
        ver, _, rec_type, rec_len = h
        is_container = ver == 0xF
        if rec_type == target_type:
            if not container_only or is_container:
                return pos
        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break
    return None


def find_ppdrawing_in_slide(data, slide_offset):
    """在 Slide container 中找到 PPDrawing 的 offset。

    Slide (container) 的子 record 包括 SlideAtom、PPDrawing 等。
    返回 PPDrawing 的 offset，找不到返回 None。
    """
    h = parse_header(data, slide_offset)
    if h is None:
        return None
    _, _, _, slide_len = h
    slide_end = slide_offset + 8 + slide_len
    return find_record(data, slide_offset + 8, slide_end, RT_PPDRAWING, container_only=True)


def find_spgr_in_ppdrawing(data, ppd_offset):
    """在 PPDrawing container 中找到 SpgrContainer 的 offset。"""
    h = parse_header(data, ppd_offset)
    if h is None:
        return None
    _, _, _, ppd_len = h
    ppd_end = ppd_offset + 8 + ppd_len
    return find_record(data, ppd_offset + 8, ppd_end, RT_SPGR_CONTAINER, container_only=True)


def find_spcontainers_in_spgr(data, spgr_offset):
    """在 SpgrContainer 中找到所有 SpContainer 的 offset 列表。

    SpgrContainer 的第一个子 record 通常是 FSPGR（形状组属性），
    后续是多个 SpContainer（每个形状一个）。
    """
    spcontainers = []
    h = parse_header(data, spgr_offset)
    if h is None:
        return spcontainers
    _, _, _, spgr_len = h
    spgr_end = spgr_offset + 8 + spgr_len
    pos = spgr_offset + 8
    while pos + 8 <= spgr_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, _, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SP_CONTAINER:
            spcontainers.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return spcontainers


def parse_spcontainer_children(data, sp_offset):
    """解析 SpContainer 的所有子 record，返回 [(pos, ver, inst, recType, recLen)]。"""
    h = parse_header(data, sp_offset)
    if h is None:
        return []
    _, _, _, sp_len = h
    sp_end = sp_offset + 8 + sp_len
    children = []
    pos = sp_offset + 8
    while pos + 8 <= sp_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        children.append((pos, ver, inst, rec_type, rec_len))
        pos += 8 + rec_len
        if ver != 0xF and rec_len == 0:
            break
    return children


def parse_fsp(data, offset):
    """解析 FSP record。

    FSP 结构（MS-ODRAW 2.3.6.1）：
      - header (8 bytes): ver=2, inst=shapeType, type=0xF007
      - shapeId (4 bytes): 形状 ID
      - flags (4 bytes): 形状标志位

    flags 关键位：
      bit 0  (0x00000001): fBackground
      bit 9  (0x00000200): fHaveAnchor — 形状有 ClientAnchor
      bit 10 (0x00000400): fInGroup
      bit 11 (0x00000800): fPropagatedDelete — 形状被删除（不可用！）
      bit 16 (0x00010000): fFlipH
      bit 17 (0x00020000): fFlipV
    """
    h = parse_header(data, offset)
    if h is None:
        return None
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_FSP:
        return None
    if rec_len < 8:
        return None
    shape_id = struct.unpack_from("<I", data, offset + 8)[0]
    flags = struct.unpack_from("<I", data, offset + 12)[0]
    return {
        "ver": ver,
        "inst": inst,            # 形状类型（MSO shape type）
        "shape_type_name": mso_shape_name(inst),
        "shape_id": shape_id,
        "flags": flags,
    }


def parse_fopt(data, offset):
    """解析 FOPT record，返回属性列表 [(propId, propName, propValue, complexDataLen)]。

    FOPT 结构（MS-ODRAW 2.3.8.1）：
      - header (8 bytes): ver=3, inst=属性数量, type=0xF008
      - rgFOPTEntry: 每个属性 6 bytes
        - opid (16 bits): propId(14) + fComplex(1) + fBlip(1)
        - op (32 bits): 属性值（或 complex data 长度，若 fComplex=1）

    若属性是 complex/blip 类型，op 是 complex data 的长度，
    complex data 紧跟在所有固定属性之后。
    """
    h = parse_header(data, offset)
    if h is None:
        return []
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_FOPT:
        return []
    num_props = inst
    props = []
    pos = offset + 8
    end = offset + 8 + rec_len

    # 第一遍：解析所有固定属性
    for _ in range(num_props):
        if pos + 6 > end:
            break
        opid = struct.unpack_from("<H", data, pos)[0]
        op = struct.unpack_from("<I", data, pos + 2)[0]
        prop_id = opid & 0x3FFF       # 低 14 位是 propId
        f_complex = (opid >> 14) & 0x1
        f_blip = (opid >> 15) & 0x1
        prop_name = FOPT_PROPS.get(prop_id, f"unknown_0x{prop_id:04X}")
        props.append({
            "prop_id": prop_id,
            "prop_name": prop_name,
            "op": op,
            "f_complex": f_complex,
            "f_blip": f_blip,
            "complex_data": None,
        })
        pos += 6

    # 第二遍：解析 complex 属性的数据
    for p in props:
        if p["f_complex"] or p["f_blip"]:
            complex_len = p["op"]
            if pos + complex_len <= end:
                p["complex_data"] = data[pos:pos + complex_len]
                pos += complex_len

    return props


def parse_client_anchor(data, offset):
    """解析 ClientAnchor record。

    ClientAnchor 结构（MS-PPT 2.5.5）：
      - header (8 bytes): ver=2, inst=0, type=0xF00A
      - top (int16): 顶部位置（单位 1/635 cm，即 EMU/635）
      - left (int16): 左侧位置
      - right (int16): 右侧位置
      - bottom (int16): 底部位置

    注意：某些版本的 ClientAnchor 可能包含额外字段。
    """
    h = parse_header(data, offset)
    if h is None:
        return None
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_CLIENT_ANCHOR:
        return None
    anchor_data = data[offset + 8:offset + 8 + rec_len]
    # 解析为 int16 列表
    values = []
    for i in range(0, len(anchor_data), 2):
        if i + 2 <= len(anchor_data):
            v = struct.unpack_from("<h", anchor_data, i)[0]  # signed int16
            values.append(v)
    return {
        "ver": ver,
        "inst": inst,
        "len": rec_len,
        "raw_hex": anchor_data.hex(),
        "values": values,
    }


def parse_client_textbox(data, offset):
    """解析 ClientTextbox container，返回子 record 列表。

    ClientTextbox (0xF00D, ver=0xF) 是 container，包含：
      - TextHeaderAtom (0x0F9F): 文本类型
      - TextCharsAtom (0x0FA0): UTF-16LE 文本
      - TextBytesAtom (0x0FA8): Latin-1 文本（旧格式）
      - StyleAtom (0x0FA1): 文本样式
      - TextSpecInfoAtom (0x0FA5): 文本特殊信息
    """
    h = parse_header(data, offset)
    if h is None:
        return []
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_CLIENT_TEXTBOX:
        return []
    tb_end = offset + 8 + rec_len
    children = []
    pos = offset + 8
    while pos + 8 <= tb_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver2, inst2, rec_type2, rec_len2 = h
        child = {
            "pos": pos,
            "ver": ver2,
            "inst": inst2,
            "rec_type": rec_type2,
            "rec_len": rec_len2,
            "type_name": "",
            "text": None,
            "tx_type": None,
            "raw_hex": None,
        }
        if rec_type2 == RT_TEXT_HEADER_ATOM:
            child["type_name"] = "TextHeaderAtom"
            if rec_len2 >= 4:
                child["tx_type"] = struct.unpack_from("<I", data, pos + 8)[0]
        elif rec_type2 == RT_TEXT_CHARS_ATOM:
            child["type_name"] = "TextCharsAtom"
            if rec_len2 > 0:
                text_data = data[pos + 8:pos + 8 + rec_len2]
                try:
                    child["text"] = text_data.decode("utf-16-le")
                except Exception:
                    child["raw_hex"] = text_data.hex()
        elif rec_type2 == RT_TEXT_BYTES_ATOM:
            child["type_name"] = "TextBytesAtom"
            if rec_len2 > 0:
                text_data = data[pos + 8:pos + 8 + rec_len2]
                try:
                    child["text"] = text_data.decode("latin-1")
                except Exception:
                    child["raw_hex"] = text_data.hex()
        elif rec_type2 == RT_STYLE_ATOM:
            child["type_name"] = "StyleAtom/TextRulerAtom"
            if rec_len2 > 0:
                child["raw_hex"] = data[pos + 8:pos + 8 + rec_len2].hex()
        elif rec_type2 == RT_TEXT_SPEC_INFO:
            child["type_name"] = "TextSpecInfoAtom"
            if rec_len2 > 0:
                child["raw_hex"] = data[pos + 8:pos + 8 + rec_len2].hex()
        else:
            child["type_name"] = f"Unknown(0x{rec_type2:04X})"
            if rec_len2 > 0:
                child["raw_hex"] = data[pos + 8:pos + 8 + min(rec_len2, 64)].hex()
        children.append(child)
        pos += 8 + rec_len2
        if ver2 != 0xF and rec_len2 == 0:
            break
    return children


def parse_fspgr(data, offset):
    """解析 FSPGR record（形状组属性）。

    FSPGR 结构（MS-ODRAW 2.3.5）：
      - header (8 bytes): ver=1, inst=0, type=0xF009
      - rcgBoxLeft (int32): 组边界左侧
      - rcgBoxTop (int32): 组边界顶部
      - rcgBoxRight (int32): 组边界右侧
      - rcgBoxBottom (int32): 组边界底部
    """
    h = parse_header(data, offset)
    if h is None:
        return None
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_FSPGR:
        return None
    if rec_len < 16:
        return None
    left = struct.unpack_from("<i", data, offset + 8)[0]
    top = struct.unpack_from("<i", data, offset + 12)[0]
    right = struct.unpack_from("<i", data, offset + 16)[0]
    bottom = struct.unpack_from("<i", data, offset + 20)[0]
    return {"ver": ver, "inst": inst, "left": left, "top": top, "right": right, "bottom": bottom}


# ============================================================================
# SpContainer 详细 dump
# ============================================================================

def dump_spcontainer(data, sp_offset, indent=0, title=None):
    """详细 dump 一个 SpContainer 的完整结构。

    参数：
      data: 完整的 PowerPoint Document stream
      sp_offset: SpContainer 的 offset
      indent: 缩进级别
      title: 可选的标题
    """
    prefix = "  " * indent
    h = parse_header(data, sp_offset)
    if h is None:
        out(f"{prefix}[错误] 无法解析 SpContainer header at offset=0x{sp_offset:X}")
        return
    ver, inst, rec_type, rec_len = h
    if title:
        out(f"\n{prefix}{'='*70}")
        out(f"{prefix}{title}")
        out(f"{prefix}{'='*70}")
    out(f"{prefix}SpContainer (offset=0x{sp_offset:X}, ver=0x{ver:X}, inst=0x{inst:03X}, "
          f"type=0x{rec_type:04X}, len={rec_len})")
    out(f"{prefix}总长度（含 header）: {8 + rec_len} 字节")

    children = parse_spcontainer_children(data, sp_offset)
    out(f"{prefix}子 record 数量: {len(children)}")
    out(f"{prefix}子 record 顺序:")
    for idx, (cpos, cver, cinst, ctype, clen) in enumerate(children):
        tname = type_name(ctype)
        out(f"{prefix}  [{idx}] offset=0x{cpos:X} ver=0x{cver:X} inst=0x{cinst:03X} "
              f"type=0x{ctype:04X} ({tname}) len={clen}")

    out(f"\n{prefix}--- 各子 record 详细解析 ---")
    for idx, (cpos, cver, cinst, ctype, clen) in enumerate(children):
        out(f"\n{prefix}[{idx}] {type_name(ctype)} (offset=0x{cpos:X}, "
              f"ver=0x{cver:X}, inst=0x{cinst:03X}, len={clen})")

        if ctype == RT_FSP:
            fsp = parse_fsp(data, cpos)
            if fsp:
                out(f"{prefix}  FSP 详细:")
                out(f"{prefix}    ver: 0x{fsp['ver']:X} (期望 0x2)")
                out(f"{prefix}    inst (形状类型): 0x{fsp['inst']:X} ({fsp['inst']}) "
                      f"-> {fsp['shape_type_name']}")
                out(f"{prefix}    shapeId: {fsp['shape_id']} (0x{fsp['shape_id']:08X})")
                out(f"{prefix}    flags: 0x{fsp['flags']:08X}")
                # 解析 flags 各位
                flags = fsp["flags"]
                flag_bits = []
                if flags & 0x00000001:
                    flag_bits.append("bit0:fBackground")
                if flags & 0x00000200:
                    flag_bits.append("bit9:fHaveAnchor")
                if flags & 0x00000400:
                    flag_bits.append("bit10:fInGroup")
                if flags & 0x00000800:
                    flag_bits.append("bit11:fPropagatedDelete(已删除!)")
                if flags & 0x00010000:
                    flag_bits.append("bit16:fFlipH")
                if flags & 0x00020000:
                    flag_bits.append("bit17:fFlipV")
                out(f"{prefix}    flags 位解析: {' | '.join(flag_bits) if flag_bits else '(无)'}")

        elif ctype == RT_FOPT:
            props = parse_fopt(data, cpos)
            out(f"{prefix}  FOPT 详细 (属性数: {len(props)}):")
            for p in props:
                complex_flag = ""
                if p["f_complex"]:
                    complex_flag = " [complex]"
                if p["f_blip"]:
                    complex_flag = " [blip]"
                out(f"{prefix}    propId=0x{p['prop_id']:04X} ({p['prop_name']}): "
                      f"op=0x{p['op']:08X}{complex_flag}")
                if p["complex_data"] is not None:
                    cd = p["complex_data"]
                    out(f"{prefix}      complex data ({len(cd)} 字节): "
                          f"{cd[:64].hex()}{'...' if len(cd) > 64 else ''}")

        elif ctype == RT_CLIENT_ANCHOR:
            anchor = parse_client_anchor(data, cpos)
            if anchor:
                out(f"{prefix}  ClientAnchor 详细:")
                out(f"{prefix}    ver: {anchor['ver']} (期望 2)")
                out(f"{prefix}    inst: {anchor['inst']}")
                out(f"{prefix}    len: {anchor['len']} 字节")
                out(f"{prefix}    raw hex: {anchor['raw_hex']}")
                out(f"{prefix}    values (int16): {anchor['values']}")
                if len(anchor["values"]) >= 4:
                    out(f"{prefix}    解析:")
                    out(f"{prefix}      top:    {anchor['values'][0]} (0x{anchor['values'][0] & 0xFFFF:04X})")
                    out(f"{prefix}      left:   {anchor['values'][1]} (0x{anchor['values'][1] & 0xFFFF:04X})")
                    out(f"{prefix}      right:  {anchor['values'][2]} (0x{anchor['values'][2] & 0xFFFF:04X})")
                    out(f"{prefix}      bottom: {anchor['values'][3]} (0x{anchor['values'][3] & 0xFFFF:04X})")
                    out(f"{prefix}    (单位: 1/635 cm = EMU/635)")

        elif ctype == RT_CLIENT_TEXTBOX:
            tb_children = parse_client_textbox(data, cpos)
            out(f"{prefix}  ClientTextbox 详细 (子 record 数: {len(tb_children)}):")
            for tc in tb_children:
                out(f"{prefix}    [{tc['pos']:6d}] ver=0x{tc['ver']:X} inst=0x{tc['inst']:03X} "
                      f"type=0x{tc['rec_type']:04X} ({tc['type_name']}) len={tc['rec_len']}")
                if tc["tx_type"] is not None:
                    out(f"{prefix}      txType: {tc['tx_type']} "
                          f"(0x{tc['tx_type']:08X})")
                if tc["text"] is not None:
                    out(f"{prefix}      文本: {tc['text']!r}")
                if tc["raw_hex"]:
                    out(f"{prefix}      raw hex: {tc['raw_hex']}")

        elif ctype == RT_F004:
            # 0xF004 是 PowerPoint 97-2003 中实际使用的 SpContainer 变体，
            # 内部结构与标准 SpContainer (0xF003) 相同，包含 FSP/FOPT/ClientAnchor/ClientTextbox。
            # 递归解析其内部 record。
            out(f"{prefix}  F004 Container 详细（递归解析内部 record）:")
            f004_children = parse_spcontainer_children(data, cpos)
            out(f"{prefix}    内部子 record 数量: {len(f004_children)}")
            for fidx, (fpos, fver, finst, ftype, flen) in enumerate(f004_children):
                ftname = type_name(ftype)
                out(f"{prefix}    [{fidx}] offset=0x{fpos:X} ver=0x{fver:X} inst=0x{finst:03X} "
                      f"type=0x{ftype:04X} ({ftname}) len={flen}")
                # 解析 F004 内部的 FSP
                if ftype == RT_FSP:
                    fsp = parse_fsp(data, fpos)
                    if fsp:
                        out(f"{prefix}      FSP: inst=0x{fsp['inst']:X} ({fsp['shape_type_name']}) "
                              f"shapeId={fsp['shape_id']} flags=0x{fsp['flags']:08X}")
                        flags = fsp["flags"]
                        flag_bits = []
                        if flags & 0x00000001:
                            flag_bits.append("bit0:fBackground")
                        if flags & 0x00000200:
                            flag_bits.append("bit9:fHaveAnchor")
                        if flags & 0x00000400:
                            flag_bits.append("bit10:fInGroup")
                        if flags & 0x00000800:
                            flag_bits.append("bit11:fPropagatedDelete(已删除!)")
                        if flags & 0x00010000:
                            flag_bits.append("bit16:fFlipH")
                        if flags & 0x00020000:
                            flag_bits.append("bit17:fFlipV")
                        out(f"{prefix}      flags 位: {' | '.join(flag_bits) if flag_bits else '(无)'}")
                # 解析 F004 内部的 FOPT
                elif ftype == RT_FOPT:
                    props = parse_fopt(data, fpos)
                    out(f"{prefix}      FOPT 属性数: {len(props)}")
                    for p in props:
                        complex_flag = ""
                        if p["f_complex"]:
                            complex_flag = " [complex]"
                        if p["f_blip"]:
                            complex_flag = " [blip]"
                        out(f"{prefix}        0x{p['prop_id']:04X} ({p['prop_name']}): "
                              f"0x{p['op']:08X}{complex_flag}")
                # 解析 F004 内部的 ClientAnchor
                elif ftype == RT_CLIENT_ANCHOR:
                    anchor = parse_client_anchor(data, fpos)
                    if anchor:
                        out(f"{prefix}      ClientAnchor: ver={anchor['ver']} len={anchor['len']} "
                              f"values={anchor['values']}")
                # 解析 F004 内部的 ClientTextbox
                elif ftype == RT_CLIENT_TEXTBOX:
                    tb_children = parse_client_textbox(data, fpos)
                    out(f"{prefix}      ClientTextbox 子 record 数: {len(tb_children)}")
                    for tc in tb_children:
                        out(f"{prefix}        {tc['type_name']} (len={tc['rec_len']})")
                        if tc["text"]:
                            out(f"{prefix}          文本: {tc['text']!r}")
                        if tc["tx_type"] is not None:
                            out(f"{prefix}          txType: {tc['tx_type']}")
                # 解析 F004 内部的 FSPGR
                elif ftype == RT_FSPGR:
                    fspgr = parse_fspgr(data, fpos)
                    if fspgr:
                        out(f"{prefix}      FSPGR: left={fspgr['left']} top={fspgr['top']} "
                              f"right={fspgr['right']} bottom={fspgr['bottom']}")
                # 其他类型打印 raw hex
                else:
                    if flen > 0:
                        raw = data[fpos + 8:fpos + 8 + min(flen, 64)]
                        out(f"{prefix}      raw hex: {raw.hex()}")

        elif ctype == RT_CLIENT_DATA:
            out(f"{prefix}  ClientData (存在，{clen} 字节)")
            if clen > 0:
                raw = data[cpos + 8:cpos + 8 + min(clen, 64)]
                out(f"{prefix}    raw hex: {raw.hex()}")

        elif ctype == RT_FSPGR:
            fspgr = parse_fspgr(data, cpos)
            if fspgr:
                out(f"{prefix}  FSPGR 详细 (形状组属性):")
                out(f"{prefix}    left:   {fspgr['left']}")
                out(f"{prefix}    top:    {fspgr['top']}")
                out(f"{prefix}    right:  {fspgr['right']}")
                out(f"{prefix}    bottom: {fspgr['bottom']}")

        else:
            out(f"{prefix}  其他 record 类型，raw hex:")
            if clen > 0:
                raw = data[cpos + 8:cpos + 8 + min(clen, 64)]
                out(f"{prefix}    {raw.hex()}")


# ============================================================================
# 文件分析主流程
# ============================================================================

def collect_f004_summary(data, f004_offset):
    """收集 0xF004 container 内部的结构摘要。

    0xF004 是 PowerPoint 97-2003 中实际使用的 SpContainer 变体，
    内部结构与标准 SpContainer (0xF003) 相同，包含 FSP/FOPT/ClientAnchor/ClientTextbox。
    但原始文件中的 0xF004 可能不包含 FSP/FOPT，而是用 FConnector (0xF011) 代替。

    返回包含内部 record 信息的字典。
    """
    children = parse_spcontainer_children(data, f004_offset)
    info = {
        "offset": f004_offset,
        "child_types": [(c[3], c[2]) for c in children],  # (recType, inst)
        "fsp": None,
        "fopt_props": [],
        "anchor": None,
        "textbox_children": [],
        "client_data": None,
        "fconnector_raw": None,
        "fspgr": None,
    }
    for (cpos, cver, cinst, ctype, clen) in children:
        if ctype == RT_FSP:
            info["fsp"] = parse_fsp(data, cpos)
        elif ctype == RT_FOPT:
            info["fopt_props"] = parse_fopt(data, cpos)
        elif ctype == RT_CLIENT_ANCHOR:
            info["anchor"] = parse_client_anchor(data, cpos)
        elif ctype == RT_CLIENT_TEXTBOX:
            info["textbox_children"] = parse_client_textbox(data, cpos)
        elif ctype == RT_CLIENT_DATA:
            info["client_data"] = {"len": clen, "raw_hex": data[cpos+8:cpos+8+min(clen, 64)].hex()}
        elif ctype == 0xF011:  # FConnector
            info["fconnector_raw"] = data[cpos+8:cpos+8+min(clen, 64)].hex()
        elif ctype == 0xF010 or ctype == RT_FSPGR:  # FSPGR
            info["fspgr"] = parse_fspgr(data, cpos) if ctype == RT_FSPGR else {"raw_hex": data[cpos+8:cpos+8+min(clen, 64)].hex()}
    return info


def load_ppt_stream(filepath):
    """用 olefile 打开 .ppt 文件，返回 PowerPoint Document stream 的字节数据。

    返回 (ppt_data, stream_names)，失败返回 (None, None)。
    """
    ole = olefile.OleFileIO(filepath)
    stream_names = ["/".join(p) for p in ole.listdir()]
    if "PowerPoint Document" not in stream_names:
        ole.close()
        return None, None
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()
    return ppt_data, stream_names


def analyze_file(filepath, label, max_slides=1):
    """分析一个 .ppt 文件中 Slide 的 SpContainer 结构。

    参数：
      filepath: .ppt 文件路径
      label: 文件标签（"原始" 或 "水印"）
      max_slides: 最多分析的 Slide 数量
    """
    out(f"\n{'#'*80}")
    out(f"# 文件标签: {label}")
    out(f"# 文件路径: {filepath}")
    out(f"{'#'*80}")

    ppt_data, stream_names = load_ppt_stream(filepath)
    if ppt_data is None:
        out(f"[错误] 找不到 PowerPoint Document stream")
        out(f"  可用 streams: {stream_names}")
        return None

    out(f"PowerPoint Document stream 大小: {len(ppt_data)} 字节")

    slides = find_all_slides(ppt_data)
    out(f"找到 {len(slides)} 个 Slide (recType=0x03EE)")

    if not slides:
        out(f"[警告] 未找到 Slide，可能 RT_SLIDE 常量需要调整")
        return None

    results = []
    for i, slide_offset in enumerate(slides[:max_slides]):
        out(f"\n{'='*70}")
        out(f"--- Slide {i+1} (offset=0x{slide_offset:X}) ---")
        out(f"{'='*70}")

        h = parse_header(ppt_data, slide_offset)
        out(f"Slide header: ver=0x{h[0]:X} inst=0x{h[1]:03X} type=0x{h[2]:04X} len={h[3]}")

        ppd_offset = find_ppdrawing_in_slide(ppt_data, slide_offset)
        if ppd_offset is None:
            out(f"  [错误] 找不到 PPDrawing")
            continue
        h = parse_header(ppt_data, ppd_offset)
        out(f"PPDrawing offset: 0x{ppd_offset:X} (ver=0x{h[0]:X} inst=0x{h[1]:03X} "
              f"type=0x{h[2]:04X} len={h[3]})")

        spgr_offset = find_spgr_in_ppdrawing(ppt_data, ppd_offset)
        if spgr_offset is None:
            out(f"  [错误] 找不到 SpgrContainer")
            continue
        h = parse_header(ppt_data, spgr_offset)
        out(f"SpgrContainer offset: 0x{spgr_offset:X} (ver=0x{h[0]:X} inst=0x{h[1]:03X} "
              f"type=0x{h[2]:04X} len={h[3]})")

        spcontainers = find_spcontainers_in_spgr(ppt_data, spgr_offset)
        out(f"SpContainer 数量: {len(spcontainers)}")

        slide_result = {
            "slide_index": i + 1,
            "slide_offset": slide_offset,
            "ppd_offset": ppd_offset,
            "spgr_offset": spgr_offset,
            "spcontainers": [],
        }

        for j, sp_offset in enumerate(spcontainers):
            title = f"SpContainer {j+1}/{len(spcontainers)} (offset=0x{sp_offset:X})"
            dump_spcontainer(ppt_data, sp_offset, indent=1, title=title)

            # 收集结构摘要用于对比
            children = parse_spcontainer_children(ppt_data, sp_offset)
            summary = {
                "index": j + 1,
                "offset": sp_offset,
                "child_types": [(c[3], c[2]) for c in children],  # (recType, inst)
                "fsp": None,
                "fopt_props": [],
                "anchor": None,
                "textbox_children": [],
                "f004_children": [],  # 0xF004 内部的子形状信息
            }
            for (cpos, cver, cinst, ctype, clen) in children:
                if ctype == RT_FSP:
                    summary["fsp"] = parse_fsp(ppt_data, cpos)
                elif ctype == RT_FOPT:
                    summary["fopt_props"] = parse_fopt(ppt_data, cpos)
                elif ctype == RT_CLIENT_ANCHOR:
                    summary["anchor"] = parse_client_anchor(ppt_data, cpos)
                elif ctype == RT_CLIENT_TEXTBOX:
                    summary["textbox_children"] = parse_client_textbox(ppt_data, cpos)
                elif ctype == RT_F004:
                    # 0xF004 是 SpContainer 变体，递归收集其内部信息
                    f004_info = collect_f004_summary(ppt_data, cpos)
                    summary["f004_children"].append(f004_info)
            slide_result["spcontainers"].append(summary)

        results.append(slide_result)

    return results


# ============================================================================
# 对比分析
# ============================================================================

def compare_results(orig_results, wm_results):
    """对比原始文件与水印文件的 SpContainer 结构差异。

    参数：
      orig_results: 原始文件的分析结果
      wm_results: 水印文件的分析结果
    """
    out(f"\n\n{'#'*80}")
    out(f"# 结构对比分析")
    out(f"{'#'*80}")

    if not orig_results or not wm_results:
        out(f"[错误] 缺少分析结果，无法对比")
        return

    # 取第一个 Slide 进行对比
    orig_slide = orig_results[0]
    wm_slide = wm_results[0]

    out(f"\n原始文件 Slide 1: {len(orig_slide['spcontainers'])} 个 SpContainer")
    out(f"水印文件 Slide 1: {len(wm_slide['spcontainers'])} 个 SpContainer")

    # 找出水印文件中多出来的 SpContainer（应该是水印）
    orig_count = len(orig_slide["spcontainers"])
    wm_count = len(wm_slide["spcontainers"])
    if wm_count > orig_count:
        out(f"\n水印文件比原始文件多 {wm_count - orig_count} 个 SpContainer（应为水印）")

    # ============================================================
    # 原始文件 SpContainer 结构概览
    # ============================================================
    out(f"\n{'='*70}")
    out(f"--- 原始文件 SpContainer 结构概览 ---")
    out(f"{'='*70}")
    if orig_slide["spcontainers"]:
        orig_sp = orig_slide["spcontainers"][0]  # 原始文件只有 1 个 SpContainer
        out(f"SpContainer #{orig_sp['index']} (offset=0x{orig_sp['offset']:X})")
        out(f"顶层子 record 顺序: {[(type_name(t), f'inst=0x{i:03X}') for t, i in orig_sp['child_types']]}")

        # 检查是否包含 0xF004
        if orig_sp["f004_children"]:
            out(f"\n原始文件 SpContainer 内部包含 {len(orig_sp['f004_children'])} 个 0xF004 container")
            out(f"这是 PowerPoint 97-2003 的组形状结构，每个 0xF004 是一个子形状")
            for idx, f004 in enumerate(orig_sp["f004_children"]):
                out(f"\n  [0xF004 #{idx+1}] (offset=0x{f004['offset']:X})")
                out(f"    内部 record 顺序: {[(type_name(t), f'inst=0x{i:03X}') for t, i in f004['child_types']]}")
                if f004["fsp"]:
                    out(f"    FSP: inst=0x{f004['fsp']['inst']:X} ({f004['fsp']['shape_type_name']}) "
                          f"shapeId={f004['fsp']['shape_id']} flags=0x{f004['fsp']['flags']:08X}")
                else:
                    out(f"    FSP: 无（原始文件 0xF004 内部不使用 FSP）")
                if f004["anchor"]:
                    out(f"    ClientAnchor: ver={f004['anchor']['ver']} len={f004['anchor']['len']} "
                          f"values={f004['anchor']['values']}")
                if f004["textbox_children"]:
                    out(f"    ClientTextbox 子 record:")
                    for tc in f004["textbox_children"]:
                        out(f"      {tc['type_name']} (len={tc['rec_len']})")
                        if tc["text"]:
                            out(f"        文本: {tc['text']!r}")
                        if tc["tx_type"] is not None:
                            out(f"        txType: {tc['tx_type']}")
                if f004["client_data"]:
                    out(f"    ClientData: len={f004['client_data']['len']}")
                if f004["fconnector_raw"]:
                    out(f"    FConnector raw: {f004['fconnector_raw']}")

    # ============================================================
    # 水印 SpContainer 结构概览
    # ============================================================
    out(f"\n{'='*70}")
    out(f"--- 水印文件最后一个 SpContainer（注入的水印）---")
    out(f"{'='*70}")
    if wm_slide["spcontainers"]:
        wm_sp = wm_slide["spcontainers"][-1]  # 最后一个是水印
        out(f"SpContainer #{wm_sp['index']} (offset=0x{wm_sp['offset']:X})")
        out(f"顶层子 record 顺序: {[(type_name(t), f'inst=0x{i:03X}') for t, i in wm_sp['child_types']]}")
        if wm_sp["fsp"]:
            out(f"FSP: inst=0x{wm_sp['fsp']['inst']:X} ({wm_sp['fsp']['shape_type_name']}) "
                  f"shapeId={wm_sp['fsp']['shape_id']} flags=0x{wm_sp['fsp']['flags']:08X}")
        if wm_sp["fopt_props"]:
            out(f"FOPT 属性:")
            for p in wm_sp["fopt_props"]:
                out(f"  0x{p['prop_id']:04X} ({p['prop_name']}): 0x{p['op']:08X}")
        if wm_sp["anchor"]:
            out(f"ClientAnchor: ver={wm_sp['anchor']['ver']} len={wm_sp['anchor']['len']} "
                  f"values={wm_sp['anchor']['values']}")
        if wm_sp["textbox_children"]:
            out(f"ClientTextbox 子 record:")
            for tc in wm_sp["textbox_children"]:
                out(f"  {tc['type_name']} (len={tc['rec_len']})")
                if tc["text"]:
                    out(f"    文本: {tc['text']!r}")
                if tc["tx_type"] is not None:
                    out(f"    txType: {tc['tx_type']}")

    # ============================================================
    # 关键差异点分析
    # ============================================================
    out(f"\n{'='*70}")
    out(f"--- 关键差异点分析 ---")
    out(f"{'='*70}")

    if orig_slide["spcontainers"] and wm_slide["spcontainers"]:
        orig_sp = orig_slide["spcontainers"][0]
        wm_sp = wm_slide["spcontainers"][-1]

        # 找一个原始文件中带文本的 0xF004 作为参考
        orig_text_f004 = None
        for f004 in orig_sp["f004_children"]:
            if f004["textbox_children"]:
                orig_text_f004 = f004
                break

        out(f"\n[1] Container 类型差异:")
        out(f"    原始文件: SpContainer(0xF003) 内部用 0xF004 container 表示子形状")
        out(f"    水印文件: SpContainer(0xF003) 直接包含 FSP/FOPT/ClientAnchor/ClientTextbox")
        out(f"    说明: 0xF004 是 PowerPoint 97-2003 的组形状子容器，")
        out(f"          而水印直接用标准 SpContainer 结构（符合 MS-ODRAW 规范）")

        out(f"\n[2] FSP (0xF007) 差异:")
        if orig_text_f004 and orig_text_f004["fsp"]:
            out(f"    原始文本框: 有 FSP, inst=0x{orig_text_f004['fsp']['inst']:X} "
                  f"({orig_text_f004['fsp']['shape_type_name']})")
        else:
            out(f"    原始文本框: 无 FSP (0xF004 内部不使用 FSP，用 FConnector 0xF011 代替)")
        if wm_sp["fsp"]:
            out(f"    水印:       有 FSP, inst=0x{wm_sp['fsp']['inst']:X} "
                  f"({wm_sp['fsp']['shape_type_name']}) "
                  f"shapeId={wm_sp['fsp']['shape_id']} flags=0x{wm_sp['fsp']['flags']:08X}")
        out(f"    说明: 原始文件用 FConnector(0xF011) 表示形状属性，水印用标准 FSP(0xF007)")

        out(f"\n[3] FOPT (0xF008) 差异:")
        if orig_text_f004 and orig_text_f004["fopt_props"]:
            out(f"    原始文本框: 有 FOPT, 属性数={len(orig_text_f004['fopt_props'])}")
            for p in orig_text_f004["fopt_props"]:
                out(f"      0x{p['prop_id']:04X} ({p['prop_name']}): 0x{p['op']:08X}")
        else:
            out(f"    原始文本框: 无 FOPT (0xF004 内部不使用 FOPT)")
        if wm_sp["fopt_props"]:
            out(f"    水印:       有 FOPT, 属性数={len(wm_sp['fopt_props'])}")
            for p in wm_sp["fopt_props"]:
                out(f"      0x{p['prop_id']:04X} ({p['prop_name']}): 0x{p['op']:08X}")
        out(f"    说明: 原始文件不用 FOPT 设置填充/线条，水印用 FOPT 设置无填充无边框")

        out(f"\n[4] ClientAnchor (0xF00A) 差异:")
        if orig_text_f004 and orig_text_f004["anchor"]:
            a = orig_text_f004["anchor"]
            out(f"    原始文本框: ver={a['ver']} inst={a['inst']} len={a['len']} "
                  f"values={a['values']}")
        if wm_sp["anchor"]:
            a = wm_sp["anchor"]
            out(f"    水印:       ver={a['ver']} inst={a['inst']} len={a['len']} "
                  f"values={a['values']}")
        out(f"    说明: ClientAnchor 格式相同（ver=2, len=8, 4个int16），但坐标值不同")

        out(f"\n[5] ClientTextbox (0xF00D) 子 record 差异:")
        if orig_text_f004 and orig_text_f004["textbox_children"]:
            orig_tb_types = [tc["type_name"] for tc in orig_text_f004["textbox_children"]]
            out(f"    原始文本框: {orig_tb_types}")
            for tc in orig_text_f004["textbox_children"]:
                if tc["text"]:
                    out(f"      文本: {tc['text']!r}")
                if tc["tx_type"] is not None:
                    out(f"      txType: {tc['tx_type']}")
        if wm_sp["textbox_children"]:
            wm_tb_types = [tc["type_name"] for tc in wm_sp["textbox_children"]]
            out(f"    水印:       {wm_tb_types}")
            for tc in wm_sp["textbox_children"]:
                if tc["text"]:
                    out(f"      文本: {tc['text']!r}")
                if tc["tx_type"] is not None:
                    out(f"      txType: {tc['tx_type']}")
        out(f"    说明: 原始文本框包含 StyleAtom/TextRulerAtom 等样式 record，")
        out(f"          水印仅包含 TextHeaderAtom + TextCharsAtom（缺少样式）")

        out(f"\n[6] ClientData (0xF00B) 差异:")
        if orig_text_f004 and orig_text_f004["client_data"]:
            out(f"    原始文本框: 有 ClientData, len={orig_text_f004['client_data']['len']}")
            out(f"      raw hex: {orig_text_f004['client_data']['raw_hex']}")
        else:
            out(f"    原始文本框: 无 ClientData")
        out(f"    水印:       无 ClientData")
        out(f"    说明: ClientData 包含形状的额外属性（如占位符类型、形状几何等），")
        out(f"          水印缺少 ClientData 可能导致 PowerPoint 无法正确渲染形状")

        out(f"\n[7] FConnector (0xF011) vs FSP (0xF007) 差异:")
        if orig_text_f004 and orig_text_f004["fconnector_raw"]:
            out(f"    原始文本框: 有 FConnector, raw={orig_text_f004['fconnector_raw']}")
        out(f"    水印:       无 FConnector，用 FSP 代替")
        out(f"    说明: FConnector 和 FSP 都包含 shapeId/flags，但 type code 不同")

        # ============================================================
        # 总结
        # ============================================================
        out(f"\n{'='*70}")
        out(f"--- 总结：水印不显示的可能原因 ---")
        out(f"{'='*70}")
        out(f"1. 结构差异: 原始文件用 0xF004+FConnector 结构，水印用标准 FSP+FOPT 结构")
        out(f"   - 水印结构符合 MS-ODRAW 规范，但可能与 PowerPoint 97-2003 兼容性不佳")
        out(f"2. 缺少 ClientData (0xF00B): 原始形状都有 ClientData，水印没有")
        out(f"   - ClientData 包含形状几何、占位符等关键信息")
        out(f"3. 缺少样式 record: 原始文本框有 StyleAtom/TextRulerAtom，水印没有")
        out(f"   - 可能导致文本无法正确渲染")
        out(f"4. FOPT 属性: 水印设置了 fillType=0(无填充) 和 fNoFillHitTest/fNoLineDrawDash")
        out(f"   - 这些属性可能影响形状可见性")
        out(f"5. ClientAnchor 坐标: 水印用 [2000,2000,17000,12000]，需确认是否在 slide 范围内")
        out(f"6. TextHeaderAtom txType: 水印用 txType=4，原始文件也用 4（一致）")


# ============================================================================
# 主入口
# ============================================================================

def main():
    """主入口：分析原始文件与水印文件，对比 SpContainer 结构。"""
    orig_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test\心理账户理论.ppt"
    wm_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\_test_out\wm_心理账户理论.ppt"
    output_path = r"d:\xwtwork\xdemo\realtime-screen-ocr-rust\pptx-rs\compare_spcontainer_output.txt"

    out("=" * 80)
    out("SpContainer 结构对比分析")
    out("=" * 80)
    out(f"原始文件: {orig_path}")
    out(f"水印文件: {wm_path}")
    out(f"说明: 使用 RT_SLIDE=0x03EE (Slide)，而非 0x03F8 (MainMaster)")

    # 分析原始文件
    orig_results = analyze_file(orig_path, label="原始文件", max_slides=1)

    # 分析水印文件
    wm_results = analyze_file(wm_path, label="水印文件", max_slides=1)

    # 对比
    if orig_results and wm_results:
        compare_results(orig_results, wm_results)

    out(f"\n{'#'*80}")
    out(f"# 分析完成")
    out(f"{'#'*80}")

    # 保存输出到 UTF-8 文件
    save_output(output_path)
    out(f"\n输出已保存到: {output_path}")


if __name__ == "__main__":
    main()
