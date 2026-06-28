#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析原始 .ppt 文件中 SpContainer 的详细结构。

目的：找出 PowerPoint 原生生成的形状（特别是文本框）的完整结构，
作为水印 SpContainer 的参考模板。
"""

import sys
import struct
import olefile

# Record type 常量
RT_SLIDE = 0x03F8
RT_PPDRAWING = 0x040C
RT_SPGR_CONTAINER = 0xF002
RT_SP_CONTAINER = 0xF003
RT_FSP = 0xF007
RT_FOPT = 0xF008
RT_CLIENT_ANCHOR = 0xF00A
RT_CLIENT_DATA = 0xF00B
RT_CLIENT_TEXTBOX = 0xF00D
RT_TEXT_HEADER_ATOM = 0x0F9F
RT_TEXT_CHARS_ATOM = 0x0FA0
RT_TEXT_BYTES_ATOM = 0x0FA8

# FOPT 属性 ID 到名称的映射（部分常见属性）
FOPT_PROPS = {
    0x0080: "transform.rotation",
    0x0081: "transform.lockRotation",
    0x00BF: " ProtectionBooleanProperties",
    0x00C0: "shapePath",
    0x00C2: "pVertices",
    0x00C3: "pSegmentInfo",
    0x00C4: "pAdjustHandles",
    0x00C5: "pGuides",
    0x00C6: "pInscribe",
    0x00C7: "pFragments",
    0x00C8: "pFragmentHeaders",
    0x00C9: "pFragmentRecords",
    0x00CA: "pFragmentRules",
    0x00CB: "pFragmentText",
    0x00CC: "pFragmentImageMap",
    0x00CD: "pFragmentProperties",
    0x00CE: "pFragmentVariables",
    0x00CF: "pFragmentMorphs",
    0x00D0: "pFragmentWraps",
    0x00D1: "pFragmentPaths",
    0x00D2: "pFragmentTextProperties",
    0x00D3: "pFragmentImageMapProperties",
    0x00D4: "pFragmentVariablesProperties",
    0x00D5: "pFragmentMorphsProperties",
    0x00D6: "pFragmentWrapsProperties",
    0x00D7: "pFragmentPathsProperties",
    0x00D8: "pFragmentTextPropertiesProperties",
    0x00D9: "pFragmentImageMapPropertiesProperties",
    0x00DA: "pFragmentVariablesPropertiesProperties",
    0x00DB: "pFragmentMorphsPropertiesProperties",
    0x00DC: "pFragmentWrapsPropertiesProperties",
    0x00DD: "pFragmentPathsPropertiesProperties",
    0x00DE: "pFragmentTextPropertiesPropertiesProperties",
    0x00DF: "pFragmentImageMapPropertiesPropertiesProperties",
    0x00E0: "pFragmentVariablesPropertiesPropertiesProperties",
    0x00E1: "pFragmentMorphsPropertiesPropertiesProperties",
    0x00E2: "pFragmentWrapsPropertiesPropertiesProperties",
    0x00E3: "pFragmentPathsPropertiesPropertiesProperties",
    0x00E4: "pFragmentTextPropertiesPropertiesPropertiesProperties",
    0x00E5: "pFragmentImageMapPropertiesPropertiesPropertiesProperties",
    0x00E6: "pFragmentVariablesPropertiesPropertiesPropertiesProperties",
    0x00E7: "pFragmentMorphsPropertiesPropertiesPropertiesProperties",
    0x00E8: "pFragmentWrapsPropertiesPropertiesPropertiesProperties",
    0x00E9: "pFragmentPathsPropertiesPropertiesPropertiesProperties",
    0x00EA: "pFragmentTextPropertiesPropertiesPropertiesPropertiesProperties",
    0x00EB: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesProperties",
    0x00EC: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00ED: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesProperties",
    0x00EE: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesProperties",
    0x00EF: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F0: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F1: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F2: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F3: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F4: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F5: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F6: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F7: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F8: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00F9: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FA: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FB: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FC: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FD: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FE: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x00FF: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0100: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0101: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0104: "geoLeft",
    0x0105: "geoTop",
    0x0106: "geoRight",
    0x0107: "geoBottom",
    0x0108: "shapePath",
    0x0109: "pAdjustHandles",
    0x010A: "pGuides",
    0x010B: "pInscribe",
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
    0x0173: "pWrapPolygonVertices",
    0x0174: "wrapText",
    0x0175: "dxWrapDistLeft",
    0x0176: "dyWrapDistTop",
    0x0177: "dxWrapDistRight",
    0x0178: "dyWrapDistBottom",
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
    0x01A0: "lineFillBlip",
    0x01A1: "lineFillBlipName",
    0x01A2: "lineFillBlipFlags",
    0x01A3: "lineFillColor",
    0x01A4: "lineBackColor",
    0x01A5: "lineCrMod",
    0x01A6: "lineOpacity",
    0x01A7: "lineWidth",
    0x01A8: "lineStyle",
    0x01A9: "lineDashing",
    0x01AA: "lineDashStyle",
    0x01AB: "lineStyleBooleanProperties",
    0x01AC: "lineStartArrowhead",
    0x01AD: "lineEndArrowhead",
    0x01AE: "lineStartArrowWidth",
    0x01AF: "lineStartArrowLength",
    0x01B0: "lineEndArrowWidth",
    0x01B1: "lineEndArrowLength",
    0x01B2: "lineJoinStyle",
    0x01B3: "lineEndCapStyle",
    0x01B4: "fillColor",
    0x01B5: "fillBackColor",
    0x01B6: "fillCrMod",
    0x01B7: "fillStyle",
    0x01B8: "fillStyleBooleanProperties",
    0x01B9: "fillBlip",
    0x01BA: "fillBlipName",
    0x01BB: "fillBlipFlags",
    0x01BC: "fillWidth",
    0x01BD: "fillHeight",
    0x01BE: "fillDztype",
    0x01BF: "fNoFillHitTest",
    0x01C0: "fNoLineDrawDash",
    0x01C1: "fNoFillHitTest",
    0x01C2: "fNoLineDrawDash",
    0x01C3: "fillColor",
    0x01C4: "fillBackColor",
    0x01C5: "fillCrMod",
    0x01C6: "fillStyle",
    0x01C7: "fillStyleBooleanProperties",
    0x01C8: "fillBlip",
    0x01C9: "fillBlipName",
    0x01CA: "fillBlipFlags",
    0x01CB: "fillWidth",
    0x01CC: "fillHeight",
    0x01CD: "fillDztype",
    0x01CE: "fillRectLeft",
    0x01CF: "fillRectTop",
    0x01D0: "fillRectRight",
    0x01D1: "fillRectBottom",
    0x01D2: "fillAngle",
    0x01D3: "fillFocus",
    0x01D4: "fillToLeft",
    0x01D5: "fillToTop",
    0x01D6: "fillToRight",
    0x01D7: "fillToBottom",
    0x01D8: "fillType",
    0x01D9: "fillBlip",
    0x01DA: "fillBlipName",
    0x01DB: "fillBlipFlags",
    0x01DC: "fillWidth",
    0x01DD: "fillHeight",
    0x01DE: "fillDztype",
    0x01DF: "fillRectLeft",
    0x01E0: "fillRectTop",
    0x01E1: "fillRectRight",
    0x01E2: "fillRectBottom",
    0x01E3: "fillAngle",
    0x01E4: "fillFocus",
    0x01E5: "fillToLeft",
    0x01E6: "fillToTop",
    0x01E7: "fillToRight",
    0x01E8: "fillToBottom",
    0x01E9: "fillType",
    0x01EA: "fillBlip",
    0x01EB: "fillBlipName",
    0x01EC: "fillBlipFlags",
    0x01ED: "fillWidth",
    0x01EE: "fillHeight",
    0x01EF: "fillDztype",
    0x01F0: "fillRectLeft",
    0x01F1: "fillRectTop",
    0x01F2: "fillRectRight",
    0x01F3: "fillRectBottom",
    0x01F4: "fillAngle",
    0x01F5: "fillFocus",
    0x01F6: "fillToLeft",
    0x01F7: "fillToTop",
    0x01F8: "fillToRight",
    0x01F9: "fillToBottom",
    0x01FA: "fillType",
    0x01FB: "fillBlip",
    0x01FC: "fillBlipName",
    0x01FD: "fillBlipFlags",
    0x01FE: "fillWidth",
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
    0x0210: "lineFillBlip",
    0x0211: "lineFillBlipName",
    0x0212: "lineFillBlipFlags",
    0x0213: "lineFillColor",
    0x0214: "lineBackColor",
    0x0215: "lineCrMod",
    0x0216: "lineOpacity",
    0x0217: "lineWidth",
    0x0218: "lineStyle",
    0x0219: "lineDashing",
    0x021A: "lineDashStyle",
    0x021B: "lineStyleBooleanProperties",
    0x021C: "lineStartArrowhead",
    0x021D: "lineEndArrowhead",
    0x021E: "lineStartArrowWidth",
    0x021F: "lineStartArrowLength",
    0x0220: "lineEndArrowWidth",
    0x0221: "lineEndArrowLength",
    0x0222: "lineJoinStyle",
    0x0223: "lineEndCapStyle",
    0x0224: "fillColor",
    0x0225: "fillBackColor",
    0x0226: "fillCrMod",
    0x0227: "fillStyle",
    0x0228: "fillStyleBooleanProperties",
    0x0229: "fillBlip",
    0x022A: "fillBlipName",
    0x022B: "fillBlipFlags",
    0x022C: "fillWidth",
    0x022D: "fillHeight",
    0x022E: "fillDztype",
    0x022F: "fillRectLeft",
    0x0230: "fillRectTop",
    0x0231: "fillRectRight",
    0x0232: "fillRectBottom",
    0x0233: "fillAngle",
    0x0234: "fillFocus",
    0x0235: "fillToLeft",
    0x0236: "fillToTop",
    0x0237: "fillToRight",
    0x0238: "fillToBottom",
    0x0239: "fillType",
    0x023A: "fillBlip",
    0x023B: "fillBlipName",
    0x023C: "fillBlipFlags",
    0x023D: "fillWidth",
    0x023E: "fillHeight",
    0x023F: "fillDztype",
    0x0240: "fillRectLeft",
    0x0241: "fillRectTop",
    0x0242: "fillRectRight",
    0x0243: "fillRectBottom",
    0x0244: "fillAngle",
    0x0245: "fillFocus",
    0x0246: "fillToLeft",
    0x0247: "fillToTop",
    0x0248: "fillToRight",
    0x0249: "fillToBottom",
    0x024A: "fillType",
    0x024B: "fillBlip",
    0x024C: "fillBlipName",
    0x024D: "fillBlipFlags",
    0x024E: "fillWidth",
    0x024F: "fillHeight",
    0x0250: "fillDztype",
    0x0251: "fillRectLeft",
    0x0252: "fillRectTop",
    0x0253: "fillRectRight",
    0x0254: "fillRectBottom",
    0x0255: "fillAngle",
    0x0256: "fillFocus",
    0x0257: "fillToLeft",
    0x0258: "fillToTop",
    0x0259: "fillToRight",
    0x025A: "fillToBottom",
    0x025B: "fillType",
    0x025C: "fillBlip",
    0x025D: "fillBlipName",
    0x025E: "fillBlipFlags",
    0x025F: "fillWidth",
    0x0260: "fillHeight",
    0x0261: "fillDztype",
    0x0262: "fillRectLeft",
    0x0263: "fillRectTop",
    0x0264: "fillRectRight",
    0x0265: "fillRectBottom",
    0x0266: "fillAngle",
    0x0267: "fillFocus",
    0x0268: "fillToLeft",
    0x0269: "fillToTop",
    0x026A: "fillToRight",
    0x026B: "fillToBottom",
    0x026C: "fillType",
    0x026D: "fillBlip",
    0x026E: "fillBlipName",
    0x026F: "fillBlipFlags",
    0x0270: "fillWidth",
    0x0271: "fillHeight",
    0x0272: "fillDztype",
    0x0273: "fillRectLeft",
    0x0274: "fillRectTop",
    0x0275: "fillRectRight",
    0x0276: "fillRectBottom",
    0x0277: "fillAngle",
    0x0278: "fillFocus",
    0x0279: "fillToLeft",
    0x027A: "fillToTop",
    0x027B: "fillToRight",
    0x027C: "fillToBottom",
    0x027D: "fillType",
    0x027E: "fillBlip",
    0x027F: "fillBlipName",
    0x0280: "fillBlipFlags",
    0x0281: "fillWidth",
    0x0282: "fillHeight",
    0x0283: "fillDztype",
    0x0284: "fillRectLeft",
    0x0285: "fillRectTop",
    0x0286: "fillRectRight",
    0x0287: "fillRectBottom",
    0x0288: "fillAngle",
    0x0289: "fillFocus",
    0x028A: "fillToLeft",
    0x028B: "fillToTop",
    0x028C: "fillToRight",
    0x028D: "fillToBottom",
    0x028E: "fillType",
    0x028F: "fillBlip",
    0x0290: "fillBlipName",
    0x0291: "fillBlipFlags",
    0x0292: "fillWidth",
    0x0293: "fillHeight",
    0x0294: "fillDztype",
    0x0295: "fillfillRectLeft",
    0x0296: "fillRectTop",
    0x0297: "fillRectRight",
    0x0298: "fillRectBottom",
    0x0299: "fillAngle",
    0x029A: "fillFocus",
    0x029B: "fillToLeft",
    0x029C: "fillToTop",
    0x029D: "fillToRight",
    0x029E: "fillToBottom",
    0x029F: "fillType",
    0x02A0: "fillBlip",
    0x02A1: "fillBlipName",
    0x02A2: "fillBlipFlags",
    0x02A3: "fillWidth",
    0x02A4: "fillHeight",
    0x02A5: "fillDztype",
    0x02A6: "fillRectLeft",
    0x02A7: "fillRectTop",
    0x02A8: "fillRectRight",
    0x02A9: "fillRectBottom",
    0x02AA: "fillAngle",
    0x02AB: "fillFocus",
    0x02AC: "fillToLeft",
    0x02AD: "fillToTop",
    0x02AE: "fillToRight",
    0x02AF: "fillToBottom",
    0x02B0: "fillType",
    0x02B1: "fillBlip",
    0x02B2: "fillBlipName",
    0x02B3: "fillBlipFlags",
    0x02B4: "fillWidth",
    0x02B5: "fillHeight",
    0x02B6: "fillDztype",
    0x02B7: "fillRectLeft",
    0x02B8: "fillRectTop",
    0x02B9: "fillRectRight",
    0x02BA: "fillRectBottom",
    0x02BB: "fillAngle",
    0x02BC: "fillFocus",
    0x02BD: "fillToLeft",
    0x02BE: "fillToTop",
    0x02BF: "fillToRight",
    0x02C0: "fillToBottom",
    0x02C1: "fillType",
    0x02C2: "fillBlip",
    0x02C3: "fillBlipName",
    0x02C4: "fillBlipFlags",
    0x02C5: "fillWidth",
    0x02C6: "fillHeight",
    0x02C7: "fillDztype",
    0x02C8: "fillRectLeft",
    0x02C9: "fillRectTop",
    0x02CA: "fillRectRight",
    0x02CB: "fillRectBottom",
    0x02CC: "fillAngle",
    0x02CD: "fillFocus",
    0x02CE: "fillToLeft",
    0x02CF: "fillToTop",
    0x02D0: "fillToRight",
    0x02D1: "fillToBottom",
    0x02D2: "fillType",
    0x02D3: "fillBlip",
    0x02D4: "fillBlipName",
    0x02D5: "fillBlipFlags",
    0x02D6: "fillWidth",
    0x02D7: "fillHeight",
    0x02D8: "fillDztype",
    0x02D9: "fillRectLeft",
    0x02DA: "fillRectTop",
    0x02DB: "fillRectRight",
    0x02DC: "fillRectBottom",
    0x02DD: "fillAngle",
    0x02DE: "fillFocus",
    0x02DF: "fillToLeft",
    0x02E0: "fillToTop",
    0x02E1: "fillToRight",
    0x02E2: "fillToBottom",
    0x02E3: "fillType",
    0x02E4: "fillBlip",
    0x02E5: "fillBlipName",
    0x02E6: "fillBlipFlags",
    0x02E7: "fillWidth",
    0x02E8: "fillHeight",
    0x02E9: "fillDztype",
    0x02EA: "fillRectLeft",
    0x02EB: "fillRectTop",
    0x02EC: "fillRectRight",
    0x02ED: "fillRectBottom",
    0x02EE: "fillAngle",
    0x02EF: "fillFocus",
    0x02F0: "fillToLeft",
    0x02F1: "fillToTop",
    0x02F2: "fillToRight",
    0x02F3: "fillToBottom",
    0x02F4: "fillType",
    0x02F5: "fillBlip",
    0x02F6: "fillBlipName",
    0x02F7: "fillBlipFlags",
    0x02F8: "fillWidth",
    0x02F9: "fillHeight",
    0x02FA: "fillDztype",
    0x02FB: "fillRectLeft",
    0x02FC: "fillRectTop",
    0x02FD: "fillRectRight",
    0x02FE: "fillRectBottom",
    0x02FF: "fillAngle",
    0x0300: "fillFocus",
    0x0301: "fillToLeft",
    0x0302: "fillToTop",
    0x0303: "fillToRight",
    0x0304: "shapeFlags",
    0x0305: "shapePath",
    0x0306: "pAdjustHandles",
    0x0307: "pGuides",
    0x0308: "pInscribe",
    0x0309: "pFragments",
    0x030A: "pFragmentHeaders",
    0x030B: "pFragmentRecords",
    0x030C: "pFragmentRules",
    0x030D: "pFragmentText",
    0x030E: "pFragmentImageMap",
    0x030F: "pFragmentProperties",
    0x0310: "pFragmentVariables",
    0x0311: "pFragmentMorphs",
    0x0312: "pFragmentWraps",
    0x0313: "pFragmentPaths",
    0x0314: "pFragmentTextProperties",
    0x0315: "pFragmentImageMapProperties",
    0x0316: "pFragmentVariablesProperties",
    0x0317: "pFragmentMorphsProperties",
    0x0318: "pFragmentWrapsProperties",
    0x0319: "pFragmentPathsProperties",
    0x031A: "pFragmentTextPropertiesProperties",
    0x031B: "pFragmentImageMapPropertiesProperties",
    0x031C: "pFragmentVariablesPropertiesProperties",
    0x031D: "pFragmentMorphsPropertiesProperties",
    0x031E: "pFragmentWrapsPropertiesProperties",
    0x031F: "pFragmentPathsPropertiesProperties",
    0x0320: "pFragmentTextPropertiesPropertiesProperties",
    0x0321: "pFragmentImageMapPropertiesPropertiesProperties",
    0x0322: "pFragmentVariablesPropertiesPropertiesProperties",
    0x0323: "pFragmentMorphsPropertiesPropertiesProperties",
    0x0324: "pFragmentWrapsPropertiesPropertiesProperties",
    0x0325: "pFragmentPathsPropertiesPropertiesProperties",
    0x0326: "pFragmentTextPropertiesPropertiesPropertiesProperties",
    0x0327: "pFragmentImageMapPropertiesPropertiesPropertiesProperties",
    0x0328: "pFragmentVariablesPropertiesPropertiesPropertiesProperties",
    0x0329: "pFragmentMorphsPropertiesPropertiesPropertiesProperties",
    0x032A: "pFragmentWrapsPropertiesPropertiesPropertiesProperties",
    0x032B: "pFragmentPathsPropertiesPropertiesPropertiesProperties",
    0x032C: "pFragmentTextPropertiesPropertiesPropertiesPropertiesProperties",
    0x032D: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesProperties",
    0x032E: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesProperties",
    0x032F: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0330: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0331: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0332: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0333: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0334: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0335: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0336: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0337: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0338: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0339: "pFragmentImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033A: "pFragmentVariablesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033B: "pFragmentMorphsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033C: "pFragmentWrapsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033D: "pFragmentPathsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033E: "pFragmentTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033F: "pImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
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
    0x0399: "wpWrapPolygonVertices",
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
    0x0380: "Text ID",
    0x0381: "wzName",
    0x0382: "wzDescription",
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
    0x033F: "pImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0304: "shapeFlags",
    0x0305: "shapePath",
    0x0306: "pAdjustHandles",
    0x0307: "pGuides",
    0x0308: "pInscribe",
    0x0309: "pFragments",
    0x030A: "pFragmentHeaders",
    0x030B: "pFragmentRecords",
    0x030C: "pFragmentRules",
    0x030D: "pFragmentText",
    0x030E: "pFragmentImageMap",
    0x030F: "pFragmentProperties",
    0x0310: "pFragmentVariables",
    0x0311: "pFragmentMorphs",
    0x0312: "pFragmentWraps",
    0x0313: "pFragmentPaths",
    0x0314: "pFragmentTextProperties",
    0x0315: "pFragmentImageMapProperties",
    0x0316: "pFragmentVariablesProperties",
    0x0317: "pFragmentMorphsProperties",
    0x0318: "pFragmentWrapsProperties",
    0x0319: "pFragmentPathsProperties",
    0x031A: "pFragmentTextPropertiesProperties",
    0x031B: "pFragmentImageMapPropertiesProperties",
    0x031C: "pFragmentVariablesPropertiesProperties",
    0x031D: "pFragmentMorphsPropertiesProperties",
    0x031E: "pFragmentWrapsPropertiesProperties",
    0x031F: "pFragmentPathsPropertiesProperties",
    0x0320: "pFragmentTextPropertiesPropertiesProperties",
    0x0321: "pImageMapPropertiesPropertiesProperties",
    0x0322: "pVariablesPropertiesPropertiesProperties",
    0x0323: "pMorphsPropertiesPropertiesProperties",
    0x0324: "pWrapsPropertiesPropertiesProperties",
    0x0325: "pPathsPropertiesPropertiesProperties",
    0x0326: "pTextPropertiesPropertiesPropertiesProperties",
    0x0327: "pImageMapPropertiesPropertiesPropertiesProperties",
    0x0328: "pVariablesPropertiesPropertiesPropertiesProperties",
    0x0329: "pMorphsPropertiesPropertiesPropertiesProperties",
    0x032A: "pWrapsPropertiesPropertiesPropertiesProperties",
    0x032B: "pPathsPropertiesPropertiesPropertiesProperties",
    0x032C: "pTextPropertiesPropertiesPropertiesPropertiesProperties",
    0x032D: "pImageMapPropertiesPropertiesPropertiesPropertiesProperties",
    0x032E: "pVariablesPropertiesPropertiesPropertiesPropertiesProperties",
    0x032F: "pMorphsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0330: "pWrapsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0331: "pPathsPropertiesPropertiesPropertiesPropertiesProperties",
    0x0332: "pTextPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0333: "pImageMapPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0334: "pVariablesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0335: "pMorphsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0336: "pWrapsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0337: "pPathsPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0338: "pTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0339: "pImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033A: "pVariablesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033B: "pMorphsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033C: "pWrapsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033D: "pPathsPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033E: "pTextPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x033F: "pImageMapPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesPropertiesProperties",
    0x0380: "Text ID",
    0x0381: "wzName",
    0x0382: "wzDescription",
    0x0384: "wzName",
    0x0385: "wzDescription",
    0x0386: "pWrapPolygonVertices",
    0x0387: "zName",
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
    0x039C: "pWrapPolygonPolygonVertices",
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
    0x03BD: "pWrapPolygon (Vertices",
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
    0x03E8: "pzWrapPolygonVertices",
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
}


def parse_header(data, offset):
    """解析 8 字节 record header，返回 (ver, inst, rec_type, rec_len)。"""
    if offset + 8 > len(data):
        return None
    ver_inst = struct.unpack_from("<H", data, offset)[0]
    rec_type = struct.unpack_from("<H", data, offset + 2)[0]
    rec_len = struct.unpack_from("<I", data, offset + 4)[0]
    ver = ver_inst & 0x0F
    inst = (ver_inst >> 4) & 0x0FFF
    return (ver, inst, rec_type, rec_len)


def find_all_slides(data):
    """找到所有 Slide record 的 offset。"""
    slides = []
    pos = 0
    while pos + 8 <= len(data):
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SLIDE:
            slides.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return slides


def find_ppdrawing(data, slide_offset):
    """在 Slide 中找到 PPDrawing 的 offset。"""
    h = parse_header(data, slide_offset)
    if h is None:
        return None
    _, _, _, slide_len = h
    slide_end = slide_offset + 8 + slide_len
    pos = slide_offset + 8
    while pos + 8 <= slide_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_PPDRAWING:
            return pos
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return None


def find_spgr_container(data, ppd_offset):
    """在 PPDrawing 中找到 SpgrContainer 的 offset。"""
    h = parse_header(data, ppd_offset)
    if h is None:
        return None
    _, _, _, ppd_len = h
    ppd_end = ppd_offset + 8 + ppd_len
    pos = ppd_offset + 8
    while pos + 8 <= ppd_end:
        h = parse_header(data, pos)
        if h is None:
            break
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SPGR_CONTAINER:
            return pos
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return None


def find_spcontainers_in_spgr(data, spgr_offset):
    """在 SpgrContainer 中找到所有 SpContainer 的 offset。"""
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
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        total_len = 8 + rec_len
        if is_container and rec_type == RT_SP_CONTAINER:
            spcontainers.append(pos)
        pos += total_len
        if not is_container and rec_len == 0:
            break
    return spcontainers


def parse_fopt(data, offset):
    """解析 FOPT record，返回属性列表。"""
    h = parse_header(data, offset)
    if h is None:
        return []
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_FOPT:
        return []
    num_props = inst
    props = []
    pos = offset + 8
    for i in range(num_props):
        if pos + 6 > len(data):
            break
        prop_id = struct.unpack_from("<H", data, pos)[0]
        prop_val = struct.unpack_from("<I", data, pos + 2)[0]
        prop_name = FOPT_PROPS.get(prop_id, f"unknown_0x{prop_id:04X}")
        props.append((prop_id, prop_name, prop_val))
        pos += 6
    return props


def parse_fsp(data, offset):
    """解析 FSP record，返回 (shape_id, flags)。"""
    h = parse_header(data, offset)
    if h is None:
        return None
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_FSP:
        return None
    shape_id = struct.unpack_from("<I", data, offset + 8)[0]
    flags = struct.unpack_from("<I", data, offset + 12)[0]
    return (inst, shape_id, flags)


def parse_client_anchor(data, offset):
    """解析 ClientAnchor record。"""
    h = parse_header(data, offset)
    if h is None:
        return None
    ver, inst, rec_type, rec_len = h
    if rec_type != RT_CLIENT_ANCHOR:
        return None
    anchor_data = data[offset + 8:offset + 8 + rec_len]
    # 解析为 u16 列表
    values = []
    for i in range(0, len(anchor_data), 2):
        if i + 2 <= len(anchor_data):
            values.append(struct.unpack_from("<H", anchor_data, i)[0])
    return (ver, inst, rec_len, values)


def parse_client_textbox(data, offset):
    """解析 ClientTextbox container，返回子 record 列表。"""
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
        ver, inst, rec_type, rec_len = h
        is_container = ver == 0xF
        children.append((pos, ver, inst, rec_type, rec_len))
        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break
    return children


def parse_spcontainer_children(data, sp_offset):
    """解析 SpContainer 的子 record。"""
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
        is_container = ver == 0xF
        pos += 8 + rec_len
        if not is_container and rec_len == 0:
            break
    return children


def analyze_file(filepath):
    """分析一个 .ppt 文件。"""
    print(f"\n{'='*80}")
    print(f"分析文件: {filepath}")
    print(f"{'='*80}")

    ole = olefile.OleFileIO(filepath)
    # olefile.listdir() 返回 [[name1], [name2], ...] 形式
    stream_names = []
    for parts in ole.listdir():
        stream_names.append("/".join(parts))
    print(f"OLE2 streams: {stream_names}")
    if "PowerPoint Document" not in stream_names:
        print("找不到 PowerPoint Document stream")
        ole.close()
        return
    ppt_data = ole.openstream("PowerPoint Document").read()
    ole.close()

    print(f"PowerPoint Document stream 大小: {len(ppt_data)} 字节")

    slides = find_all_slides(ppt_data)
    print(f"找到 {len(slides)} 个 Slide")

    # 只分析前 2 个 Slide
    for i, slide_offset in enumerate(slides[:2]):
        print(f"\n--- Slide {i+1} (offset=0x{slide_offset:X}) ---")
        ppd_offset = find_ppdrawing(ppt_data, slide_offset)
        if ppd_offset is None:
            print("  找不到 PPDrawing")
            continue
        print(f"  PPDrawing offset: 0x{ppd_offset:X}")

        spgr_offset = find_spgr_container(ppt_data, ppd_offset)
        if spgr_offset is None:
            print("  找不到 SpgrContainer")
            continue
        print(f"  SpgrContainer offset: 0x{spgr_offset:X}")

        spcontainers = find_spcontainers_in_spgr(ppt_data, spgr_offset)
        print(f"  SpContainer 数量: {len(spcontainers)}")

        for j, sp_offset in enumerate(spcontainers):
            print(f"\n  === SpContainer {j+1} (offset=0x{sp_offset:X}) ===")
            children = parse_spcontainer_children(ppt_data, sp_offset)
            print(f"  子 record 数量: {len(children)}")

            for (child_pos, ver, inst, rec_type, rec_len) in children:
                type_name = {
                    RT_FSP: "FSP",
                    RT_FOPT: "FOPT",
                    RT_CLIENT_ANCHOR: "ClientAnchor",
                    RT_CLIENT_DATA: "ClientData",
                    RT_CLIENT_TEXTBOX: "ClientTextbox",
                }.get(rec_type, f"Unknown(0x{rec_type:04X})")
                print(f"    [{child_pos:6d}] ver=0x{ver:X} inst=0x{inst:03X} type=0x{rec_type:04X} ({type_name}) len={rec_len}")

                if rec_type == RT_FSP:
                    fsp = parse_fsp(ppt_data, child_pos)
                    if fsp:
                        print(f"      FSP: inst=0x{fsp[0]:X} shapeId={fsp[1]} flags=0x{fsp[2]:08X}")
                elif rec_type == RT_FOPT:
                    props = parse_fopt(ppt_data, child_pos)
                    print(f"      FOPT 属性数: {len(props)}")
                    for (pid, pname, pval) in props:
                        print(f"        0x{pid:04X} ({pname}): 0x{pval:08X}")
                elif rec_type == RT_CLIENT_ANCHOR:
                    anchor = parse_client_anchor(ppt_data, child_pos)
                    if anchor:
                        print(f"      ClientAnchor: ver={anchor[0]} inst={anchor[1]} len={anchor[2]} values={anchor[3]}")
                elif rec_type == RT_CLIENT_TEXTBOX:
                    tb_children = parse_client_textbox(ppt_data, child_pos)
                    print(f"      ClientTextbox 子 record 数: {len(tb_children)}")
                    for (tb_pos, tb_ver, tb_inst, tb_type, tb_len) in tb_children:
                        tb_type_name = {
                            RT_TEXT_HEADER_ATOM: "TextHeaderAtom",
                            RT_TEXT_CHARS_ATOM: "TextCharsAtom",
                            RT_TEXT_BYTES_ATOM: "TextBytesAtom",
                        }.get(tb_type, f"Unknown(0x{tb_type:04X})")
                        print(f"        [{tb_pos:6d}] ver=0x{tb_ver:X} inst=0x{tb_inst:03X} type=0x{tb_type:04X} ({tb_type_name}) len={tb_len}")
                        if tb_type == RT_TEXT_CHARS_ATOM:
                            text_data = ppt_data[tb_pos + 8:tb_pos + 8 + tb_len]
                            try:
                                text = text_data.decode("utf-16-le")
                                print(f"          文本: {text!r}")
                            except Exception:
                                print(f"          文本(原始): {text_data!r}")
                        elif tb_type == RT_TEXT_BYTES_ATOM:
                            text_data = ppt_data[tb_pos + 8:tb_pos + 8 + tb_len]
                            try:
                                text = text_data.decode("latin-1")
                                print(f"          文本: {text!r}")
                            except Exception:
                                print(f"          文本(原始): {text_data!r}")
                        elif tb_type == RT_TEXT_HEADER_ATOM:
                            tx_type = struct.unpack_from("<I", ppt_data, tb_pos + 8)[0]
                            print(f"          txType: {tx_type}")


def main():
    import os
    test_dir = "_test"
    if not os.path.isdir(test_dir):
        print(f"找不到 {test_dir} 目录")
        return
    for fname in sorted(os.listdir(test_dir)):
        if fname.lower().endswith(".ppt"):
            analyze_file(os.path.join(test_dir, fname))


if __name__ == "__main__":
    main()
