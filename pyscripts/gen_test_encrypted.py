"""用 Python 生成加密 PPTX，用于对照测试 WPS 是否认 modifyVerifier。"""
import zipfile
import shutil

# 1) 纯拷贝
shutil.copy2('_test/文旅IP人设打造抖音短视频运营方案.pptx', '_test_out/copy_test.pptx')
print('copy_test.pptx created')

# 2) Python 直接修改 XML 插入 modifyVerifier
#    不重复声明 xmlns:p（已在父元素 <p:presentation> 声明）
mv = ('<p:modifyVerifier cryptProviderType="rsaAES" '
      'cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" '
      'cryptAlgorithmSid="14" cryptSpinCount="100000" '
      'hash="U2lhgW/8YJml2ca8JlrUCWw33ioHu5IfT8edg95+5GkjExhvLOTdFQWveOwuLxZOaQ36ZyUHV+pIdvCGPnx1Sw==" '
      'salt="TTlvTgyim6fB9RqIcyk/zQ=="/>')

with zipfile.ZipFile('_test/文旅IP人设打造抖音短视频运营方案.pptx', 'r') as zin:
    with zipfile.ZipFile('_test_out/py_encrypted.pptx', 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'ppt/presentation.xml':
                xml_str = data.decode('utf-8')
                xml_str = xml_str.replace('</p:presentation>', mv + '</p:presentation>')
                data = xml_str.encode('utf-8')
            zout.writestr(item, data)

print('py_encrypted.pptx created')

# 验证
with zipfile.ZipFile('_test_out/py_encrypted.pptx') as z:
    pres = z.read('ppt/presentation.xml').decode('utf-8')
    idx = pres.find('modifyVerifier')
    if idx > 0:
        end = pres.find('/>', idx) + 2
        print('modifyVerifier:')
        print(pres[idx:end])
