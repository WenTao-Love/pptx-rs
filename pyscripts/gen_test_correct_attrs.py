"""用 Python 生成加密 PPTX，使用正确的属性名 hashData/saltData。"""
import zipfile

# 使用正确的属性名：hashData/saltData（不是 hash/salt）
MV = ('<p:modifyVerifier cryptProviderType="rsaAES" '
      'cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" '
      'cryptAlgorithmSid="14" cryptSpinCount="100000" '
      'hashData="U2lhgW/8YJml2ca8JlrUCWw33ioHu5IfT8edg95+5GkjExhvLOTdFQWveOwuLxZOaQ36ZyUHV+pIdvCGPnx1Sw==" '
      'saltData="TTlvTgyim6fB9RqIcyk/zQ=="/>')

with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_enc_correct_attrs.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'ppt/presentation.xml':
                xml_str = data.decode('utf-8')
                if 'modifyVerifier' not in xml_str:
                    xml_str = xml_str.replace('</p:presentation>', MV + '</p:presentation>')
                    data = xml_str.encode('utf-8')
            zout.writestr(item, data)

print('py_enc_correct_attrs.pptx created')

# 验证
with zipfile.ZipFile('_test_out/py_enc_correct_attrs.pptx') as z:
    pres = z.read('ppt/presentation.xml').decode('utf-8')
    idx = pres.find('modifyVerifier')
    if idx > 0:
        end = pres.find('/>', idx) + 2
        print('modifyVerifier:')
        print(pres[idx:end])
