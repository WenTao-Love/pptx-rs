# debug-pptx-output-fail

## 症状
1. `hello.pptx` 用 WPS 打不开
2. `_test_out/wm_*.pptx` 用 WPS 打开看不到水印
3. `_test_out/protected_*.pptx` 用 WPS 打开并未加密（可能未触发"输入修改密码"）

## 状态
[OPEN] 验证中

## 假设（H1 ~ H5）

### H1: hello.pptx 的 [Content_Types].xml 仍然缺 Override 或缺 namespace
可能 `contribute_to` 中只有部分 part 注册了 override，
或 `<Override PartName=...>` 的 PartName 写法不对（PowerPoint 严格要求以 `/` 开头）。

### H2: hello.pptx 缺关键 part（theme/presProps/viewProps/tableStyles/app）或者内容不规范
- theme1.xml 命名空间或元素顺序不对，PowerPoint 拒绝
- presProps.xml/viewProps.xml/tableStyles.xml 缺根元素
- app.xml 不是 valid ExtendedProperties

### H3: hello.pptx 内部命名空间/前缀错误
- presentation.xml 中 `r:id` 应该绑定 `xmlns:r="http://...relationships"`，但当前 `r:id` 在 sldId 元素中可能没有正确映射
- 类似 `p14:` 这种 Microsoft 扩展 namespace 在不需要时强制声明会让 PowerPoint 拒绝

### H4: watermark 注入的 sp XML 不规范
- id 与已有 sp id 冲突
- 插入位置错误（必须 `<p:sp>` 且在 `<p:spTree>` 内）
- `<p:txBody>` 中必须按顺序有 `<a:bodyPr><a:lstStyle><a:p>...`
- 命名空间未声明（应在 `<p:sp>` 所在根元素 `<p:sld>` 上声明）

### H5: protect 注入的 `<p:modifyVerifier>` 元素顺序/属性不对
- modifyVerifier 必须在 `<p:sldIdLst>` 等之前，还是在 `</p:presentation>` 之前即可
- 哈希算法 sid=14 对应 SHA-512，但属性顺序/可选项可能漏
- 修改密码保护在 PowerPoint 打开时可能只弹"以只读打开"提示，不一定强要求密码

## 验证步骤
1. 解析 hello.pptx 的所有 XML，看每个是否符合 OOXML schema
2. 用 WPS 实际打开 hello.pptx，看错误信息
3. 对比 python-pptx 生成的 hello.pptx 结构
4. 修复代码后重新生成，再次用 WPS 验证
