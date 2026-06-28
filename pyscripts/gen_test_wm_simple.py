"""用 Python 生成各种水印变体，找出 WPS 能渲染的方案。"""
import zipfile

# 变体1：最简单的水印 - 无旋转、无透明、红色大字、居中
WM_SIMPLE = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
             '<p:nvSpPr><p:cNvPr id="9999" name="Watermark"/>'
             '<p:cNvSpPr/><p:nvSpPr/></p:nvSpPr>'
             '<p:spPr>'
             '<a:xfrm><a:off x="1143000" y="4572000"/>'
             '<a:ext cx="6858000" cy="914400"/></a:xfrm>'
             '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
             '<a:noFill/><a:ln w="0"><a:noFill/></a:ln>'
             '</p:spPr>'
             '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
             '<a:r><a:rPr lang="zh-CN" sz="6000" b="1">'
             '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
             '<a:latin typeface="Calibri"/>'
             '</a:rPr><a:t>WATERMARK TEST</a:t></a:r>'
             '</a:p></p:txBody></p:sp>')

wm_id = 9001
with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_wm_simple.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                xml_str = data.decode('utf-8')
                if 'WATERMARK TEST' not in xml_str and 'pptx-rs WATERMARK' not in xml_str:
                    wm = WM_SIMPLE.replace('9999', str(wm_id))
                    # 插在 </p:spTree> 前（z-order 最顶层）
                    xml_str = xml_str.replace('</p:spTree>', wm + '</p:spTree>')
                    data = xml_str.encode('utf-8')
                    wm_id += 1
            zout.writestr(item, data)

print(f'py_wm_simple.pptx created ({wm_id - 9001} slides)')
