"""验证 protected_*.pptx 是否真的包含文档保护设置（直接看 XML 原文）。"""
import sys
import zipfile

path = sys.argv[1]
print(f"检查: {path}")

with zipfile.ZipFile(path) as z:
    names = z.namelist()
    print(f"包含 {len(names)} 个 part")
    pres = z.read("ppt/presentation.xml").decode("utf-8")

    if "modifyVerifier" in pres:
        i = pres.find("modifyVerifier")
        # 输出 modifyVerifier 元素的内容
        end = pres.find("/>", i) + 2
        snippet = pres[i - 5:end + 5]
        print("FOUND <p:modifyVerifier>:")
        print(f"  {snippet}")
        print()
        print("VERIFIED: 文档保护已设置！WPS 打开时会要求输入修改密码")
    else:
        print("ERROR: <p:modifyVerifier> not found in presentation.xml!")

    # 顺便看看 presentation.xml 的尾部
    print()
    print("presentation.xml 尾部 500 字符:")
    print(pres[-500:])
