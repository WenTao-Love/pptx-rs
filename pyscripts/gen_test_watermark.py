"""用 Python 生成水印 PPTX，用于对照测试 WPS 水印可见性。"""
import zipfile

# 水印 shape XML：浅灰色半透明 40pt -45° 旋转，位于 spTree 末尾（z-order 最顶层）
WATERMARK = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
             '<p:nvSpPr><p:cNvPr id="9999" name="Watermark"/>'
             '<p:cNvSpPr txBox="1"/><p:nvSpPr/></p:nvSpPr>'
             '<p:spPr>'
             '<a:xfrm rot="-2700000"><a:off x="457200" y="2743200"/>'
             '<a:ext cx="8229600" cy="1828800"/></a:xfrm>'
             '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
             '<a:noFill/><a:ln><a:noFill/></a:ln>'
             '</p:spPr>'
             '<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>'
             '<a:lstStyle/><a:p><a:pPr algn="ctr"/>'
             '<a:r><a:rPr lang="zh-CN" sz="4000" b="1">'
             '<a:solidFill><a:srgbClr val="BFBFBF"><a:alpha val="30000"/></a:srgbClr></a:solidFill>'
             '<a:latin typeface="Calibri"/><a:ea typeface="宋体"/>'
             '</a:rPr><a:t>pptx-rs WATERMARK</a:t></a:r>'
             '</a:p></p:txBody></p:sp>')

wm_id = 9001
with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_watermark.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                xml_str = data.decode('utf-8')
                if 'pptx-rs WATERMARK' not in xml_str:
                    wm = WATERMARK.replace('9999', str(wm_id))
                    xml_str = xml_str.replace('</p:spTree>', wm + '</p:spTree>')
                    data = xml_str.encode('utf-8')
                    wm_id += 1
            zout.writestr(item, data)

print(f'py_watermark.pptx created ({wm_id - 9001} slides watermarked)')
