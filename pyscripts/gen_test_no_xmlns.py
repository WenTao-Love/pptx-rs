"""用 Python 生成无 xmlns 重复声明的水印和加密文件，与 Rust 版对照。"""
import zipfile

# === 水印：无 xmlns 重复声明 ===
WM = ('<p:sp><p:nvSpPr><p:cNvPr id="9999" name="pptx-rs Watermark"/>'
      '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
      '<p:spPr><a:xfrm rot="-2700000"><a:off x="457200" y="2743200"/>'
      '<a:ext cx="8229600" cy="1828800"/></a:xfrm>'
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
      '<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>'
      '<a:lstStyle/><a:p><a:pPr algn="ctr"/>'
      '<a:r><a:rPr lang="zh-CN" sz="4000" b="1">'
      '<a:solidFill><a:srgbClr val="BFBFBF"><a:alpha val="40000"/></a:srgbClr></a:solidFill>'
      '<a:latin typeface="Calibri"/><a:ea typeface="宋体"/>'
      '</a:rPr><a:t>pptx-rs WATERMARK</a:t></a:r>'
      '</a:p></p:txBody></p:sp>')

wm_id = 9001
with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_wm_no_xmlns.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                xml_str = data.decode('utf-8')
                if 'pptx-rs WATERMARK' not in xml_str and 'pptx-rs Watermark' not in xml_str:
                    wm = WM.replace('9999', str(wm_id))
                    xml_str = xml_str.replace('</p:spTree>', wm + '</p:spTree>')
                    data = xml_str.encode('utf-8')
                    wm_id += 1
            zout.writestr(item, data)

print(f'py_wm_no_xmlns.pptx created ({wm_id - 9001} slides)')

# === 加密：无 xmlns 重复声明 ===
MV = ('<p:modifyVerifier cryptProviderType="rsaAES" '
      'cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" '
      'cryptAlgorithmSid="14" cryptSpinCount="100000" '
      'hash="U2lhgW/8YJml2ca8JlrUCWw33ioHu5IfT8edg95+5GkjExhvLOTdFQWveOwuLxZOaQ36ZyUHV+pIdvCGPnx1Sw==" '
      'salt="TTlvTgyim6fB9RqIcyk/zQ=="/>')

with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_enc_no_xmlns.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'ppt/presentation.xml':
                xml_str = data.decode('utf-8')
                if 'modifyVerifier' not in xml_str:
                    xml_str = xml_str.replace('</p:presentation>', MV + '</p:presentation>')
                    data = xml_str.encode('utf-8')
            zout.writestr(item, data)

print('py_enc_no_xmlns.pptx created')
