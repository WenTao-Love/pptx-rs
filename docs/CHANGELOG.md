# 更新日志

> 所有重要变更记录在此。格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。
> 版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### 新增

- 无（待下一版本规划）。

## [0.3.0] - 2026-06-29

### 新增

- **.ppt 97-2003 二进制格式支持**：水印注入 + RC4 CryptoAPI 加密
  - `pptx_rs::ppt97` 模块（基于 `cfb` crate）：`add_watermark` / `encrypt` / `add_watermark_and_encrypt`
  - 填补 python-pptx 不支持 .ppt 二进制格式的空白
  - 水印特性：注入到 MainMaster 的 PPDrawing（覆盖所有幻灯片），FOPT 0x01C2 锁定不可编辑
  - 加密特性：MS-OFFCRYPTO 规范 2.3.5 RC4 CryptoAPI，每个 persist 对象独立加密
- `Error::Ppt97` 错误变体（错误枚举从 10 → 11 变体）
- `#![forbid(unsafe_code)]` crate 级属性（与项目规则 §5 安全红线一致）

### 变更

- **crate 重命名**：`pptx` → `pptx_rs`（lib name 与 crates.io 已占用 crate 解冲突）
  - 所有 `use pptx::...` 改为 `use pptx_rs::...`（41 文件）
  - crates.io 上 crate 名为 `pptx-rs2`（原名 `pptx-rs` 已被他人占坑，只有 3 行代码的空 crate）
  - 通过 `[lib] name = "pptx_rs"` 解耦 crate 名与 lib 名，代码中 `use pptx_rs::...` 保持不变
- `Cargo.toml` 元数据完善：
  - 添加 `rust-version = "1.75"`（MSRV）
  - 添加 `repository` / `homepage` / `documentation` URL
  - 完善 `keywords` / `categories` / `exclude` 列表
- `Presentation::slides_mut().add_slide()` 新增 `id_counter: Rc<Cell<u32>>` 参数（API 破坏性变更）
- 公开 API 稳定性窗口调整为 `0.3.x` 期间

### 修复

- 修复 `src/lib.rs` 文档/属性结构错乱（`//!` 文档块出现在 `#![deny]` 之后）
- 修复 broken doc links：`[docs::ARCHITECTURE]` / `[docs::CHANGELOG]` 等改为 GitHub 绝对 URL
- 补全 `Shapes` / `ShapesMut` 的 `Debug` 实现
- 修复多项 clippy lint（`-D warnings` 全绿）：
  - `derivable_impls`：8 处手写 `impl Default` 改为 `#[derive(Default)]` + `#[default]`
  - `should_implement_trait`：6 处 `from_str` 方法改名为 `parse`
  - `explicit_auto_deref`：8 处 `&*t` → `&t`
  - `unnecessary_to_owned`：4 处 `.to_string()` → `.as_ref()`
  - `field_reassign_with_default`：多处改为 struct literal 语法
  - `collapsible_match` / `if_same_then_else` / `vec_init_then_push` / `replace_box` / `useless_format` / `unnecessary_unwrap` / `type_complexity` / `bool_assert_comparison` / `approx_constant` / `needless_update` / `non_snake_case` / `doc_lazy_continuation`
  - `too_many_arguments`：4 处生产 API 加 `#[allow]` 属性（0.4 路线图再重构为 builder 模式）

### 文档

- `README.md` 更新至 0.3.0，添加已知限制说明（round-trip 保真缺口）
- `LICENSE` 文件添加（MIT，版权 `2026 pptx-rs contributors`）

### 持续开发积累（v6.6 - v6.7，自 0.2.0 后陆续合入，统一归入 0.3.0 发布）

#### 新增

- **测试覆盖率大幅提升（v6.7，TODO-041 持续推进）**：新增 25 个 P0 级集成测试，覆盖文本格式化、图表、备注/批注/自定义属性
  - `tests/text_format_integration.rs`（8 个测试）：Run 字体大小/加粗/斜体/下划线/删除线/RGB 颜色/主题色/段落对齐行距/文本框锚定自适应边距/多段落多 Run 混合/AutoShape 填充线条格式
  - `tests/chart_integration.rs`（8 个测试）：6 种图表类型 round-trip（柱/条/线/饼/散点/面积）+ 多图表混合单 slide + 多 slide 不同图表
  - `tests/notes_comments_integration.rs`（9 个测试）：备注设置/清除/多 slide 混合 + 批注添加/多条/清除 + 自定义属性 5 种类型/删除 + 核心属性 round-trip，含 zip part 内容验证辅助函数
  - 测试总数：436 → 461（lib 358 + 集成 42 + doctest 61）
- **图表数据标签 / 次坐标轴支持（TODO-004 剩余小项，v6.6）**：补齐图表 dLbls 解析与系列级次轴绑定
  - `oxml::chart::ChartData` 新增 `data_labels: Option<DataLabels>` 字段（图表级 `<c:dLbls>`）
  - `Chart::parse_from_xml` SAX 状态机新增 dLbls 解析上下文跟踪（`chart_dl` / `series_dl` / `dl_target` / `has_secondary_axis`）+ 新增 `DlTarget` 枚举区分 Chart/Series 归属
  - Start 事件识别 `dLbls` 元素（根据 `cur_ser.is_some()` 判断归属）+ 识别 `<c:crosses val="max"/>` 标记次坐标轴
  - Empty 事件解析 9 个 dLbls 子元素（showVal / showCatName / showSerName / showLegendKey / showPercent / showBubbleSize / dLblPos / separator / numFmt）
  - End 事件处理 `c:ser` 关闭时写入 `series_dl` 到 `cur_ser.data_labels`；末尾遍历系列设置 `secondary_axis=true`（非散点/非饼图）
  - 新增 `parse_bool_val` / `attr_val` 辅助函数（兼容 `0/1/true/false` + 命名空间前缀后缀匹配）
  - `shape::chartshape::ChartShape` 新增 7 个便捷方法：`data_labels()` / `set_data_labels()` / `series_data_labels()` / `set_series_data_labels()` / `is_series_secondary()` / `set_series_secondary()` / `push_series()`，全部延续零 panic 设计
  - 19 个单元测试（15 个 chart parse + 4 个 chartshape）
- **SmartArt 文本节点编辑 API（TODO-037 剩余小项，v6.6）**：暴露 DataModelPoint 文本编辑能力
  - `oxml::diagram::DataModelPoint` 新增 3 个方法：`set_text(new_text)`（双字段同步：更新 `text` 字段 + 替换 `raw_xml` 中的 `<a:t>...</a:t>` 内容，自动 XML 转义）/ `clear_text()`（同步清空 text 字段和 raw_xml 中 `<a:t>` 内容）/ `is_type(type_str)`（便捷类型查询）
  - 新增 `escape_xml_text(s)` 辅助函数（XML 文本转义：`&` / `<` / `>` / `'` / `"` → 实体引用）
  - `DataModel::to_xml()` 新增结构化重建分支：当 `raw_xml` 为空但 `text` 非空时，构造完整 `<dgm:pt><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>文本</a:t></a:r></a:p></dgm:t></dgm:pt>` 子树（对齐 OOXML 规范的 SmartArt 节点完整结构）
  - `DataModel` 新增 3 个便捷方法：`point(model_id)` / `point_mut(model_id)` / `set_point_text(model_id, new_text)`
  - `presentation::DiagramEntry` 新增 2 个写回方法：`set_data_model(data_model)`（批量编辑后一次性写回 `data_xml`）+ `set_point_text(model_id, new_text) -> Result<bool>`（单点编辑：解析→修改→序列化回 `data_xml`）
  - 8 个单元测试覆盖 set_text 双字段同步 / clear_text / 结构化重建 / point 查询 / set_point_text 写回
- **Font 东亚字体便捷 API（TODO-005 剩余小项，v6.6）**：暴露东亚字体 / 复杂脚本字体高阶访问
  - `oxml::txbody::Run` 新增 4 个便捷方法（紧接 `set_font_name` 之后）：`eastasia_name() -> Option<&str>` / `set_eastasia_name(name)` / `complex_script_name() -> Option<&str>` / `set_complex_script_name(name)`，底层访问 `RunProperties.eastasia_font` / `cs_font` 字段（v5 已存在但未暴露高阶 API）
  - `Font<'a>` 新增 2 个 clear 方法：`clear_eastasia_name()` / `clear_complex_script_name()`，设置字段为 `None`（走主题继承），对标现有 `clear_name()` 模式
  - 7 个单元测试覆盖 Run getter/setter 双向 / Run 三字体独立性 / Font clear 方法 / Font view 与 Run 一致性 / 序列化顺序验证
- **.ppt（PowerPoint 97-2003 二进制格式）文件加密**：完整的 RC4 CryptoAPI 加密实现
  - OLE2/CFB 容器读写（基于 `cfb` crate）
  - RC4 流密码 + SHA1 密钥派生（对齐 MS-OFFCRYPTO 规范）
  - CryptSession10Container 构造、PersistDirectoryAtom 更新、CurrentUserAtom 加密标记
  - 每个 persist 对象独立加密（block=persistId，分段加密）
  - 示例：`protect_ppt`（仅加密）
  - msoffcrypto-python 验证通过（is_encrypted / load_key / decrypt 全部成功）
- **.ppt 文件水印注入**：在每个 MainMaster 的 PPDrawing 中注入水印 SpContainer（不可编辑背景层）
  - Escher OfficeArt 二进制结构构造（FSP / FOPT / ClientAnchor / ClientTextbox）
  - TextCharsAtom（UTF-16LE）水印文本
  - 加水印后正确更新 PersistDirectoryAtom 中所有 persist 对象的 offset
  - 水印作为 SpgrContainer 中"组形状本身"之后的第一个子形状（z-order 最低，普通视图不可选中/编辑）
  - FOPT 保护位 0x01C2 锁定，全屏覆盖 ClientAnchor
  - 示例：`watermark_ppt`（仅水印）
- **.ppt 水印+加密合并**：`watermark_and_protect_ppt` 示例，先加水印再加密
  - 解决加水印后 offset 偏移问题：更新 UserEditAtom / PersistDirectoryAtom / persist entries
  - msoffcrypto-python 验证通过（解密后正确找到水印文本）
- **`ppt97` 库模块**：将 examples 中的水印/加密逻辑提炼为公共库 API
  - 新增模块路径 `src/ppt97/`，含 4 个子模块：`record` / `ole` / `watermark` / `crypto`
  - 公共 API：`pptx::ppt97::add_watermark`、`pptx::ppt97::encrypt`、`pptx::ppt97::add_watermark_and_encrypt`
  - 配置类型：`pptx::ppt97::WatermarkConfig`（文本、字号、颜色、旋转角度）
  - 错误类型扩展：`Error::Ppt97(String)` 变体 + `Error::ppt97()` 便捷构造器
  - examples 从 1000+ 行简化为 60~80 行薄封装（仅参数解析 + 库 API 调用）
  - 修复 `parse_persist_directory` 的实现 bug：PersistDirectoryEntry 由 `persistId(20bit) | cPersist(12bit)` 头 + offsets 组成，原实现错误假设为 `cPersist(4byte) + offsets`
- Python 验证脚本：`verify_ppt_crypto.py`（加密验证）、`verify_ppt_watermark.py`（水印验证）、`verify_watermark_and_protect.py`（合并验证）
- **基础图表支持（TODO-004）**：完成 P1 级核心缺口，对标 python-pptx `add_chart`
  - 新增 `oxml::chart` 模块：`Chart` / `ChartType`（Column/Bar/Line/Pie）/ `ChartData` / `ChartCategory` / `ChartSeries` 完整模型
  - `Chart::to_xml()` 生成完整 `<c:chartSpace>` XML（含 numCache/strCache 内嵌数据，避免依赖嵌入式 Excel）
  - `Graphic::Chart` 变体 + `GraphicFrame::write_xml` Chart 分支处理
  - 高阶 `ChartShape` 类型（包装 `OxmlFrame`，零 panic 设计：`chart()`/`chart_mut()` 返回 `Option`）
  - `ShapeKind::Chart(ChartShape)` 变体 + `wrap()` 工厂函数 Chart 分支
  - `ShapesMut::add_chart(chart_type, data, left, top, width, height)` 高阶 API
  - `Slide` 新增 `chart_entries` / `chart_index_counter` / `chart_rid_counter` 跟踪机制（`Rc<Cell<u32>>` 模式）
  - `ChartEntry` 类型 + `to_opc_package` 写出 `/ppt/charts/chartN.xml` 独立 part + slide rels
  - `ct::CHART` Content-Type 常量
  - 示例：`chart_demo`（柱状图 + 折线图 + 饼图端到端演示）
- **SpPr 内 blipFill 解析（TODO-003/048）**：补齐 shape 级图片填充解析路径
  - 新增 `parse_blip_fill` 函数：`<a:blipFill>` → `(rid, BlipFillMode)`
  - 支持 `<a:blip r:embed="..."/>` 提取 rid（兼容 `r:embed` / `*:embed` 两种命名空间写法）
  - 支持 `<a:stretch>` / `<a:tile>` 填充模式解析（tile 全属性：tx/ty/sx/sy/flip/algn）
  - 兼容自闭合 `<a:blip/>` / `<a:stretch/>` / `<a:tile/>` 形态
  - `parse_sppr` 中 blipFill 分支从"collect 跳过"改为调用 `parse_blip_fill`，写入 `Fill::Blip { rid, mode }`
  - 3 个单元测试：stretch 模式 / tile 全属性 / 自闭合标签
- **形状锁定高阶 API（TODO-027）**：暴露 `LockType` 枚举 + `set_lock` 便捷方法
  - `LockType` 枚举（12 个变体：Grouping/Rotation/Selection/Resize/...）
  - `AutoShape::set_lock` / `TextBox::set_lock` 高阶方法
- **形状效果高阶 API（TODO-011）**：暴露 shadow/glow/soft_edge/reflection 便捷方法
  - `AutoShape` / `TextBox` 新增：`set_outer_shadow` / `set_inner_shadow` / `set_glow` / `set_soft_edge` / `set_reflection`
- **ShapeStyle 高阶 API（TODO-006）**：暴露 `style()` / `set_style()` 便捷方法
  - `AutoShape` / `TextBox` 新增 `style()` 访问器 + `set_style(ShapeStyle)` 设置器
- **FreeformBuilder 输出 custGeom（TODO-024）**：打通"用户 API 路径"
  - `FreeformBuilder::build()` 输出 `Geometry::Custom(CustomGeometry)` 而非退化为 `prstGeom=rect`
  - 测试断言：`<a:custGeom>` 存在且 `<a:prstGeom>` 不存在
- **表格单元格拆分（TODO-029）**：`TableShape::split_cell(row, col)` 高阶 API
- **SlideMaster 高阶编辑（TODO-049）**：`SlideMasterRef` 完整编辑能力
  - `shapes()` / `shapes_mut()` 母版形状视图
  - `placeholders()` 占位符列表
  - `background()` / `set_background()` / `set_background_solid()` / `clear_background()` 背景编辑
  - `add_shape()` / `remove_shape()` 形状增删
- **组合内子形状编辑（TODO-032）**：`Group` 子形状增删 API
  - `add_autoshape` / `add_picture` / `add_connector` / `add_table` / `add_group` 类型安全追加
  - `remove_child(id)` 递归匹配移除
- **超链接高阶 API（TODO-026）**：`Run` 便捷方法补齐
  - `hlink_click()` / `set_hlink_click()` / `clear_hlink_click()` 点击超链接访问器
  - `hlink_hover()` / `set_hlink_hover()` / `clear_hlink_hover()` 悬停超链接访问器
  - `set_hyperlink(rid, tooltip)` 外部 URL 超链接便捷方法（对标 python-pptx `run.hyperlink.address`）
  - `set_slide_jump()` 跳转幻灯片动作便捷方法
- **删除线高阶 API（TODO-017）**：`Run::double_strike()` / `set_double_strike()` 便捷方法
- **高亮高阶 API（TODO-018）**：`Run::highlight()` / `set_highlight()` / `clear_highlight()` 便捷方法
- **图片裁剪高阶 API（TODO-044）**：`Picture` 单边裁剪访问器
  - `set_crop(left, top, right, bottom)` 作为 `crop` 的 python-pptx 风格别名
  - `crop_left()` / `crop_top()` / `crop_right()` / `crop_bottom()` 单边 getter
  - `set_crop_left()` / `set_crop_top()` / `set_crop_right()` / `set_crop_bottom()` 单边 setter（保留其它三边）
- **幻灯片过渡 setter（TODO-020）**：确认 `Slide::transition()` / `set_transition()` / `clear_transition()` 已完整存在
- **章节分组（TODO-039）**：新增 `Section` / `SectionList` oxml 模型 + `Presentation::sections()` / `sections_mut()` 高阶 API
  - 新增 `src/oxml/section.rs`：`Section`（name + slide_ids）+ `SectionList` + `write_xml` 输出 `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E64}"><p14:sectionLst>`
  - `PresentationRoot` 加 `sections` 字段，`to_xml` 在 `<p:defaultTextStyle>` 之后输出 sectionLst 扩展
  - `parse_pres_root` 扩展为 5 元组返回（新增 `SectionList`），新增 `parse_sections_from_ext_lst` 辅助函数解析 `<p:extLst>` 内的 sectionLst
  - `Presentation` 加 `sections` 字段 + `sections()` / `sections_mut()` 访问器；`from_opc` / `to_opc_package` 透传 sections
  - `lib.rs` 导出 `Section` / `SectionList`
- **备注母版访问（TODO-045）**：新增 `NotesMaster` / `NotesMasterRef` / `NotesMasters` 只读访问模型
  - OPC 层：`RelType::NotesMaster` + `ct::NOTES_MASTER` Content-Type 常量
  - oxml 层：新增 `src/oxml/notesmaster.rs`（`NotesMaster` 极简只读模型，承载 shapes + background）+ `parse_notes_master` 函数
  - 高阶 API 层：新增 `src/notes_masters.rs`（`NotesMasterRef` + `NotesMasters`，参考 `SlideMasterRef` / `SlideMasters`）
  - `Presentation` 加 `notes_masters` 字段 + `notes_masters()` / `notes_masters_mut()` / `notes_master()` 访问器
  - `from_opc` 解析 `presentation.xml.rels` 中的 `NotesMaster` 关系并还原 `NotesMasterRef`
  - `lib.rs` 导出 `NotesMaster` / `NotesMasterRef` / `NotesMasters`
- **图片占位符类型化填充（TODO-007）**：补齐 `Pic` 占位符字段 + `add_picture_to_placeholder` 高阶 API
  - `PpPlaceholderType` 新增 `Picture`（`pic`）变体 + `from_str` 解析方法（修复 `pic` 错误回落为 `body` 的 bug）
  - `Pic` 结构体新增 `is_placeholder` / `ph_idx` / `ph_type` 字段，`write_xml` 在 `<p:nvPr>` 内写出 `<p:ph type="..." idx="..."/>`
  - `Picture` 高阶 API 新增 `set_placeholder` / `clear_placeholder` / `is_placeholder` / `ph_idx` / `ph_type` 方法
  - `ShapesMut::add_picture_to_placeholder(ph_idx, path, layout)` 高阶 API：自动从版式占位符继承位置/尺寸
  - `placeholders_inherited` / `placeholder_inherited` 扩展识别 `Pic` 占位符（返回 `ShapeKind::Picture`）
- **图表/表格占位符类型化填充（TODO-007 剩余）**：补齐 `GraphicFrame` 占位符字段 + `add_chart_to_placeholder` / `add_table_to_placeholder` 高阶 API
  - `GraphicFrame` 结构体新增 `is_placeholder` / `ph_idx` / `ph_type` 字段，`write_xml` 在 `<p:nvGraphicFramePr>/<p:nvPr>` 内写出 `<p:ph type="..." idx="..."/>`
  - `ChartShape` 高阶 API 新增 `set_placeholder` / `clear_placeholder` / `is_placeholder` / `ph_idx` / `ph_type` 方法
  - `TableShape` 高阶 API 新增 `set_placeholder` / `clear_placeholder` / `is_placeholder` / `ph_idx` / `ph_type` 方法
  - `ShapesMut::add_chart_to_placeholder(ph_idx, chart_type, data, layout)` 高阶 API：自动从版式占位符继承位置/尺寸，标记 `type="chart"`
  - `ShapesMut::add_table_to_placeholder(ph_idx, rows, cols, layout)` 高阶 API：自动从版式占位符继承位置/尺寸，标记 `type="tbl"`
  - `placeholders_inherited` / `placeholder_inherited` 扩展识别 `GraphicFrame` 占位符（返回 `ShapeKind::Chart` / `ShapeKind::Table`）
- **进阶图表类型（TODO-004 进阶）**：新增散点图 / 面积图支持
  - `ChartType` 新增 `Scatter` / `Area` 变体
  - `ChartSeries` 新增 `x_values: Option<Vec<f64>>` 字段（散点图 X 坐标）+ `new_scatter(name, x_values, y_values)` 构造器
  - `Chart::to_xml` 处理散点图特殊语义：`<c:scatterStyle val="lineMarker"/>` + `<c:xVal>` / `<c:yVal>`（替代 `c:cat` / `c:val`）+ 两个 `<c:valAx>`（替代 `catAx + valAx`）
  - `Chart::to_xml` 处理面积图：`<c:areaChart>` + `<c:grouping val="standard"/>` + `catAx + valAx`
  - 2 个单元测试：散点图（验证 `c:xVal` / `c:yVal` / 双 `valAx` / 无 `catAx` / 无 `c:cat`）+ 面积图（验证 `c:areaChart` / `grouping=standard` / `catAx + valAx` / `c:cat`）
  - `chart_demo.rs` 示例扩展为 2 张幻灯片 5 个图表（柱/线/饼/散点/面积）
- **OLE 对象嵌入（TODO-043）**：新增 `add_ole_object` 高阶 API + 完整 OPC/oxml/shape 三层支持
  - OPC 层：新增 `RelType::OleObject` 变体 + URI 映射 + `ct::OLE_OBJECT` Content-Type 常量
  - oxml 层：新建 `src/oxml/ole.rs`（`OleObject` 结构体 + `write_xml` + `OLE_GRAPHIC_DATA_URI` 常量），支持 `<p:oleObj spid="..." name="..." r:id="..." imgW="..." imgH="..." progId="..." showAsIcon="1">` + `<p:embed/>` + 可选 `<p:pic>` 图标
  - `Graphic` 枚举新增 `OleObject(OleObject)` 变体，`GraphicFrame::write_xml` 添加 OleObject 分支（`<a:graphicData uri=".../ole">` 包裹）
  - 高阶层：新建 `src/shape/oleshape.rs`（`OleObjectShape` + `Shape` trait 实现 + 零 panic 设计），提供 `rid` / `set_rid` / `image_rid` / `set_image_rid` / `prog_id` / `ole_name` / `show_as_icon` / `set_icon_size` / `set_pic_id_name` / `set_placeholder` 等方法
  - `ShapeKind` 新增 `OleObject(OleObjectShape)` 变体，`wrap()` 函数添加 Graphic::OleObject 分支
  - `lib.rs` 导出 `OleObject` / `OLE_GRAPHIC_DATA_URI`
  - Presentation 层：新增 `OleEntry` 结构体（partname + blob + rid），`Slide` 新增 `ole_entries` / `ole_rid_counter` / `ole_index_counter` 字段 + `allocate_ole_rid` / `next_ole_index` / `register_ole` 方法
  - `ShapesMut::add_ole_object(path, prog_id, name, left, top, width, height)` 高阶 API：读取文件 → 创建 OleObjectShape → 分配 rid → 注册 OleEntry → 推入 spTree
  - `to_opc_package` 写出 `/ppt/embeddings/oleObjectN.bin` part + `slideN.xml.rels` oleObject 关系（全局索引避免多 slide 冲突）
  - 8 个单元测试（ole.rs 4 个 + oleshape.rs 4 个）+ `examples/ole_demo.rs` 端到端示例
- **三维效果（TODO-050）**：新增 `scene3d` / `sp3d` 完整 oxml 模型 + 解析层
  - oxml 层：新增 10 个结构体/枚举（`Rotation3d` / `CameraPreset` / `Camera` / `LightRigType` / `LightRigDirection` / `LightRig` / `Scene3d` / `Bevel` / `MaterialPreset` / `Sp3d`），每个都带 `write_xml` / `as_str` / `from_str` 方法
  - `ShapeProperties` 新增 `scene3d: Option<Scene3d>` / `sp3d: Option<Sp3d>` 字段，`write_xml` 在 `effectLst` 之后、`close(tag)` 之前按 OOXML 顺序输出
  - 解析层：`parse_sppr` 新增 `scene3d` / `sp3d` 分支，新增 `parse_rotation_3d` / `parse_scene_3d` / `parse_sp_3d` / `parse_bevel_attrs` 函数（Start/Empty 事件分离处理，兼容自闭合 `<a:bevelT w="..." h="..."/>` 与开闭形式）
  - `Sp3d::write_xml` 智能省略默认值（`prstMaterial` 仅在非 `WarmMatte` 时输出，避免覆盖默认值）
  - `lib.rs` 导出全部 10 个 3D 类型
  - 8 个单元测试覆盖序列化与 `from_str` 解析
- **音视频嵌入（TODO-033）**：新增 `add_video` / `add_audio` 高阶 API + 完整 OPC/oxml/shape 三层支持
  - OPC 层：`RelType` 新增 `Video` / `Audio` / `Media` 变体 + URI 映射；`ct` 新增 `VIDEO_MP4` / `AUDIO_MP3` Content-Type 常量
  - oxml 层：`Pic` 结构体新增 `media: Option<MediaKind>` 字段；新增 `MediaKind` 枚举（`Video { rid }` / `Audio { rid }`）；`Pic::write_xml` 在 `<p:nvPr>` 内输出 `<a:videoFile r:link="..."/>` / `<a:audioFile r:link="..."/>`（`r:link` 而非 `r:embed`，区别于海报帧图片）
  - 高阶层：`Picture` 新增 `set_video(rid)` / `set_audio(rid)` / `media_kind()` / `clear_media()` 便捷方法
  - Presentation 层：新增 `VideoEntry` / `AudioEntry` 结构体（partname + blob + rid）；`Slide` 新增 `video_entries` / `audio_entries` 字段 + `allocate_video_rid` / `next_video_index` / `register_video` / `allocate_audio_rid` / `next_audio_index` / `register_audio` 方法
  - `ShapesMut::add_video(video_path, poster_path, left, top, width, height)` 高阶 API：读取视频文件 → 读取海报帧图片（None 时用内置 1x1 透明 PNG 占位）→ 注册海报帧 `MediaEntry`（`rIdImgN` + `imageN.png`）→ 分配 `rIdVideoN` → 调用 `pic.set_video` → 注册 `VideoEntry` → 推入 spTree
  - `ShapesMut::add_audio(audio_path, poster_path, left, top, width, height)` 高阶 API：与 `add_video` 对称，仅媒体类型与 Content-Type 不同
  - `to_opc_package` 写出 `/ppt/media/mediaN.mp4` / `mediaN.mp3` part + `slideN.xml.rels` Video / Audio 关系（全局索引避免多 slide 冲突）
  - `lib.rs` 导出 `MediaKind` / `VideoEntry` / `AudioEntry`
  - 4 个单元测试（Pic 序列化 video/audio/无 media + MediaKind PartialEq）+ `examples/media_demo.rs` 端到端示例
- **SmartArt 最小保留（TODO-037）**：识别 + XML round-trip（不解析也不保留 diagram parts）
  - `Graphic` 枚举新增 `SmartArt(SmartArtRef)` 变体；新增 `SmartArtRef` 结构体（`raw_xml` + `dm_rid` / `lo_rid` / `qs_rid` / `cs_rid` 4 个关系 id）
  - 解析层：`parse_graphic_into` 新增 `diagram` uri 分支，调用 `collect_full_element` 保留完整 `<a:graphicData>` 元素 XML（byte-exact，含外壳）；新增 `parse_smartart_rel_ids` 函数从 raw_xml 提取 `<dgm:relIds>` 的 4 个关系 id（简单字符串查找，避免 quick-xml 命名空间处理复杂性）
  - 序列化层：`GraphicFrame::write_xml` 检测 `Graphic::SmartArt` 时跳过 `open_with("a:graphicData")` + `close("a:graphicData")` 流程，直接 `w.raw(&s.raw_xml)` 输出完整元素（避免重新拆解丢失原始格式）
  - `lib.rs` 导出 `SmartArtRef`
  - 4 个单元测试（序列化 byte-exact / Default / 解析 graphicFrame / parse_smartart_rel_ids 多种格式）
- **主题 fmtScheme 结构化解析（TODO-005）**：保留 raw_xml 用于 round-trip 的同时，把 4 个 style 列表拆分为可查询/修改的结构化字段
  - `FormatScheme` 新增 4 个结构化字段：`fill_styles` / `line_styles` / `effect_styles` / `bg_fill_styles`（每个元素是对应子元素的原始 XML 字符串，如 `<a:gradFill>...</a:gradFill>`）
  - 新增 `parse_from_raw_xml()` 方法：使用 quick-xml 状态机（Seeking → InContainer）从 raw_xml 拆分 4 个 `<a:xxxStyleLst>` 容器的直接子元素
  - `write_xml` 重构为三级优先：**结构化字段 > raw_xml > 默认 Office 格式方案**（DEFAULT_FMT_SCHEME）
  - 新增 4 个查询方法：`fill_style_count()` / `line_style_count()` / `effect_style_count()` / `bg_fill_style_count()`
  - 新增 3 个辅助函数：`collect_style_lst_children`（状态机收集容器子元素）/ `local_name_quick`（去命名空间前缀）/ `collect_full_element_str`（收集完整 XML 字符串）
  - `parse_theme` 集成：解析 fmtScheme 后调用 `parse_from_raw_xml()` 填充结构化字段
  - 11 个单元测试覆盖：parse 拆分 / 子元素内容 / 结构化 write_xml 顺序 / count 查询 / 空 raw_xml noop / raw_xml 回退 / 默认方案 / **默认 Office 主题 round-trip（3/3/3/3）** / 字段可变性 / 自闭合子元素 / local_name / 容器不存在
- **3D 高阶 API（TODO-050 高阶）**：在 AutoShape / TextBox 上添加 3D 效果便捷方法
  - `AutoShape` 新增 9 个 3D 方法：`set_3d_rotation(lat, lon, rev)` / `set_3d_extrusion(height, color)` / `set_3d_bevel(top_w, top_h, bottom_w, bottom_h)` / `set_3d_material(preset)` / `clear_3d()` / `scene_3d()` / `scene_3d_mut()` / `sp_3d()` / `sp_3d_mut()`
  - `TextBox` 委托全部 9 个方法（与 `set_outer_shadow` 等效果 API 一致的设计模式）
  - 角度参数统一使用**度**（用户直觉），内部转换为 1/60000 度（OOXML ST_Angle）
  - 零 panic 设计：所有方法在不变量被破坏时返回 Option/默认值
- **进阶图表：雷达图 / 气泡图（TODO-004 进阶）**：扩展 ChartType 枚举与 to_xml 实现
  - `ChartType` 新增 `Radar` / `Bubble` 变体
  - 新增 `is_xy_chart()`（覆盖 Scatter + Bubble）/ `is_bubble()` 辅助方法
  - `ChartSeries` 新增 `bubble_sizes: Option<Vec<f64>>` 字段 + `new_bubble(name, x, y, sizes)` 构造器
  - `Chart::to_xml` 新增分支：雷达图（`<c:radarChart>` + `<c:radarStyle val="marker"/>` + catAx + valAx）/ 气泡图（`<c:bubbleChart>` + `<c:bubbleScale val="100"/>` + c:xVal/c:yVal/c:bubbleSize + 两个 valAx）
  - 系列处理与轴处理统一用 `is_xy_chart()` 覆盖散点图与气泡图
  - 2 个单元测试覆盖雷达/气泡 XML 结构
- **SmartArt 完整 round-trip（TODO-037 完整）**：保留 4 个 diagram parts，read→save 后 SmartArt 可正确渲染
  - OPC 层：`RelType` 新增 `DiagramData` / `DiagramLayout` / `DiagramQuickStyle` / `DiagramColors` 4 个变体 + URI 映射；`from_xml` 识别这 4 类关系；`ct` 新增 `DIAGRAM_DATA` / `DIAGRAM_LAYOUT` / `DIAGRAM_QUICK_STYLE` / `DIAGRAM_COLORS` 4 个 Content-Type 常量
  - Presentation 层：新增 `DiagramEntry` 结构体（4 个 partname + 4 份原始 XML 字符串 + 4 个 rId），采用**完整 round-trip**策略——读路径保留 4 个 part 的原始 XML，写路径直接写入 zip（不重新序列化），保证任何 SmartArt 模板/布局/颜色变体都能正确保留
  - `Slide` 新增 `diagram_entries` / `diagram_index_counter` / `diagram_rid_counter` 字段 + `next_diagram_index()` / `allocate_diagram_rids()` / `register_diagram()` 方法
  - `to_opc_package` 写出 4 个 `/ppt/diagrams/{data,layout,quickStyles,colors}N.xml` part + `slideN.xml.rels` 中 4 个 diagram 关系（全局索引避免多 slide 冲突）
  - 读路径 `from_opc` 增强：收集 slide rels 中的 4 类 diagram 关系到 `diagram_rel_map`；遍历 `sld.shapes` 中的 `GraphicFrame.Graphic::SmartArt`，根据其 4 个 rid 配对查 rels 找 target，读取 4 个 part 内容，构造 `DiagramEntry` 注入 `slide.diagram_entries`
  - `lib.rs` 导出 `DiagramEntry`
  - 7 个单元测试覆盖：rid 分配递增 / index 递增 / register_diagram 存储 / 端到端 to_bytes 写出 4 parts + 4 rels / 多 slide 全局索引不冲突 / RelType URI 正确 / from_xml 识别 4 类关系
- **图表读路径（TODO-004 读路径）**：解析已有 chart graphicFrame 的 `<c:chart>` 内容，支持读取 PowerPoint 创建的图表
  - `Chart::parse_from_xml(xml)` 实现：基于 quick-xml SAX 事件流的状态机解析器，跟踪 `in_chart_elem` / `in_title_text` / `in_cache` / `ser_field` 等上下文状态
  - 覆盖 8 种图表类型（Column/Bar/Line/Pie/Scatter/Area/Radar/Bubble）的 `<c:xxxChart>` 元素 + `<c:ser>` 系列 + `<c:numCache>` / `<c:strCache>` 缓存 + `<c:title>` 标题
  - `presentation.rs::from_opc` 两阶段策略：`parse_sld` 阶段从 slide 的 graphicFrame 提取 `<c:chart r:id="rIdX"/>` 的 rid，构造占位 `Chart`；`from_opc` 阶段读取对应 `chartN.xml` part 内容，调用 `Chart::parse_from_xml` 还原真实模型，用解析结果替换占位 Chart
  - 借用冲突两阶段解决：`for shape in &slide.inner.shapes` 不可变借用期间无法调用 `slide.inner.shapes.iter_mut()`，先收集 `(rid, partname)` 对到 Vec，结束不可变借用后再可变借用替换 `Graphic::Chart` 内容
  - 11 个单元测试覆盖：柱状图 / 条形图 / 折线图 / 饼图 / 散点图 / 气泡图 / 雷达图 round-trip + 多系列 + 无标题 + 空 chartSpace + 错误 XML 容错
- **SmartArt 数据模型结构化解析（TODO-037 数据模型）**：新建 `src/oxml/diagram.rs` 模块，4 个 part 的结构化模型
  - `DataModel`（**完全结构化**，对应 `<dgm:dataModel>`）：`points: Vec<DataModelPoint>` + `connections: Vec<DataModelConnection>`
    - `DataModelPoint`：`model_id` / `pt_type`（doc/par/ch/sib/prev/next）/ `text`（从 `<dgm:t>/<a:p>/<a:r>/<a:t>` 提取的首个文本）/ `properties`（ang / cx / cy 等节点属性）
    - `DataModelConnection`：`cxn_type`（parChld/sibTx 等）/ `src_id` / `dest_id` / `par_trans` / `sib_trans`
  - `LayoutDef`（**半结构化**，对应 `<dgm:layoutDef>`）：`unique_id` / `name` / `verb` / `style_lbl` / `category` 列表 + `layout_node_xml`（保留 layoutNode 整段子树原始 XML，不展开为强类型树）
  - `QuickStyleDef`（**半结构化**，对应 `<dgm:styleData>`）：`unique_id` / `name` / `style_lbl` 列表 + `raw_xml`
  - `ColorsDef`（**半结构化**，对应 `<dgm:colorsDef>`）：`unique_id` / `name` / `category` 列表 + `style_clr_lbl` 列表 + `raw_xml`
  - **按需解析（lazy parsing）**：`DiagramEntry` 仍以 `String` blob 持有原始 XML 保证 byte-exact round-trip；新增 `data_model()` / `layout_def()` / `quick_style_def()` / `colors_def()` 4 个方法按需触发解析，返回强类型模型
  - **零 panic 设计**：解析失败返回 `Error::Xml`，不阻塞 round-trip
  - `lib.rs` 导出 8 个类型：`DataModel` / `DataModelPoint` / `DataModelConnection` / `LayoutDef` / `LayoutCategory` / `QuickStyleDef` / `ColorsDef` / `StyleLabel`
- **3D 场景背景 backdrop（TODO-050 backdrop）**：`<a:backdrop>` 元素支持
  - 新增 `Backdrop` 结构体（OOXML CT_Backdrop）：定义 3D 场景中的 6 个背景平面
    - `anchor: Option<Point3d>`：锚点位置（`<a:anchor x="..." y="..." z="..."/>`）
    - `floor: bool` / `wall: bool` / `left: bool` / `right: bool` / `top: bool` / `bottom: bool`：6 个独立启用/禁用的平面（`<a:floor/>` / `<a:wall/>` / `<a:l/>` / `<a:r/>` / `<a:t/>` / `<a:b/>`）
  - `Backdrop::write_xml` 按 OOXML 元素顺序输出：anchor → floor → wall → l → r → t → b（仅写出启用平面）
  - `Scene3d.backdrop: Option<Backdrop>` 字段挂载，`Scene3d::write_xml` 在 camera + lightRig 之后输出 backdrop
  - `lib.rs` 导出 `Backdrop`
  - 2 个单元测试覆盖：默认序列化（无平面启用，仅空 backdrop 容器）+ 含 anchor 与全部平面启用的序列化（验证元素顺序与属性）
- **性能基准测试建立（TODO-040）**：新增 `benches/` 目录 + criterion 基准测试基线
  - `Cargo.toml` 新增 `criterion = "0.5"` dev-dependency + 两个 `[[bench]]` 配置（`save_pptx` / `large_pptx`，`harness = false`）
  - 新增 `[profile.bench]`：`opt-level = 3` + `debug = true`（保留调试信息以便 perf 分析）
  - `benches/save_pptx.rs`：5 个基础场景基准（new_presentation / save_empty_to_bytes / save_hello_pptx / save_with_shapes / round_trip_save_load_save）
  - `benches/large_pptx.rs`：3 个大型 PPTX 场景基准（save_large_pptx 100/500 slides / serialize_only 100/500 slides / round_trip_large 100 slides）
  - 运行方式：`cargo bench --bench save_pptx` 或 `cargo bench --bench large_pptx`
- **集成测试覆盖（TODO-041）**：新增 `tests/` 目录 + 17 个集成测试
  - `tests/presentation_save.rs`：6 个 Presentation 全流程测试（空保存 / slide 数量保持 / 临时文件 round-trip / load-modify-resave / 多次 round-trip 稳定性）
  - `tests/shape_integration.rs`：6 个形状端到端测试（文本框 / 多自选形状 / 表格 / 连接器 / 混合形状 / 多幻灯片不同形状）
  - `tests/large_pptx.rs`：5 个大型 PPTX 测试（50/100 slides round-trip / 100 slides 保存到文件 / 多次 round-trip / load-modify-resave 30→50 slides）
  - 覆盖 TODO-041 关键差距：复杂 PPTX 读写测试 + 大型 PPTX（100+ slides）读写测试
- **crates.io 发布就绪（TODO-042）**：完善 Cargo.toml 元数据 + CI 配置
  - `Cargo.toml` 新增 `repository` / `homepage` / `documentation` / `authors` 字段（发布时需替换为真实仓库地址）
  - `Cargo.toml` 新增 `exclude` 字段：排除 `_test/` / `_test_out/` / `pyscripts/` / `.trae/` / `docs/` / `examples/` 等非必要资源，减小发布包体积
  - 新增 `.github/workflows/ci.yml`：4 个 CI job（lint: fmt+clippy+doc / test: 跨平台 ubuntu+windows / e2e: hello_pptx 示例 + 产物上传 / publish-dry-run: cargo publish --dry-run）
  - CI 触发条件：push 到 main/develop + pull request
- **图表读路径（TODO-004 读路径）**：实现 `chartN.xml` → 强类型 `Chart` 模型的反向解析
  - `oxml::chart::Chart::parse_from_xml` 新增 ~270 行 SAX 状态机：解析 `<c:chartSpace>` 提取 chart_type / title / categories / series（含 x_values / bubble_sizes）
  - 状态变量跟踪：`in_chart_elem` / `in_title_text` / `in_cache` / `ser_field` 上下文管理
  - `chart_type` 推断：根据 `<c:plotArea>` 下的图表元素名（barChart/lineChart/pieChart/scatterChart/areaChart/radarChart/bubbleChart）+ barDir 属性区分 Column/Bar
  - `parse_sld::parse_graphic_into` 新增 chart uri 分支：提取 `<c:chart r:id="..."/>` 的 rid，构造占位 Chart
  - `Presentation::from_opc` 新增 chart 处理段落（两阶段策略避免借用冲突）：阶段一收集 (rid, partname) 对，阶段二读取 chartN.xml → parse_from_xml → 替换 graphic + 注册 ChartEntry
  - 11 个单元测试覆盖：6 种图表类型 round-trip + barDir 区分 Bar/Column + 多系列 + 无标题 + 空 chartSpace 零 panic + 畸形 XML 错误返回
- **SmartArt 数据模型（TODO-037 数据模型）**：4 个 diagram parts 的结构化解析能力
  - 新增 `oxml::diagram` 模块（~1145 行）：4 个强类型结构体 + parse_from_xml + to_xml
  - `DataModel`（完全结构化）：节点列表（`DataModelPoint`：model_id / pt_type / text / raw_xml）+ 连接列表（`DataModelConnection`：src_id / dest_id / cxn_type / raw_xml）
  - `LayoutDef`（半结构化）：uniqueId / title / desc / categories + layoutNode 子树原始 XML（保留 byte-exact）
  - `QuickStyleDef`（半结构化）：styleLbl 列表（name + raw_xml）
  - `ColorsDef`（半结构化）：uniqueId / title / desc + styleClrLbl 列表
  - **按需解析（lazy parsing）策略**：`DiagramEntry` 仍以 String blob 持有原始 XML，保证 byte-exact round-trip；新增 4 个方法 `data_model()` / `layout_def()` / `quick_style_def()` / `colors_def()` 按需触发解析
  - 零 panic 设计：解析失败返回 `Error::Xml`，不阻塞 round-trip
  - 12 个单元测试覆盖：4 个 part 各自的 parse / raw_xml 保留 / round-trip / 空 XML 零 panic / 畸形 XML 错误返回
- **backdrop 背景元素（TODO-050 backdrop）**：3D 场景背景平面支持
  - 新增 `Backdrop` 结构体：anchor (Option<Point3d>) + 6 个 plane bool 字段（floor/wall/left/right/top/bottom）
  - 新增 `Point3d` 结构体（CT_Point3D）：x/y/z 三个 EMU 坐标字段
  - `Scene3d` 新增 `backdrop: Option<Backdrop>` 字段，`write_xml` 按 OOXML 顺序输出（camera → lightRig → backdrop）
  - `parse_sld::parse_scene_3d` 新增 backdrop 分支 + 新增 `parse_backdrop` 函数（兼容 Empty 和 open-close 形式）
  - 7 个单元测试覆盖：Backdrop 默认/完整/部分序列化 + Scene3d 带/不带 backdrop + Point3d 序列化 + parse_backdrop 端到端解析
- **图表 Excel 嵌入（TODO-004 Excel 嵌入）**：支持在图表中嵌入 Excel 工作簿，PowerPoint "编辑数据" 时启动 Excel
  - **OPC 层**：`opc/package.rs` 新增 `ct::SPREADSHEET_XLSX` 常量（`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`）；`opc/rels.rs` 新增 `RelType::Package` 变体 + URI 映射 + `from_xml` 识别分支
  - **OOXML 层**：`Chart` 结构体新增 `external_data_rid: Option<String>` 字段
    - `Chart::to_xml` 在 `</c:chart>` 之后、`</c:chartSpace>` 之前写出 `<c:externalData r:id="..."><c:autoUpdate val="0"/></c:externalData>`
    - `Chart::parse_from_xml` SAX 循环提取 `externalData` 的 `r:id`（兼容 `r:id` 与 `:id` 两种写法，部分工具不严格加 r: 前缀）
  - **Presentation 层**：`ChartEntry` 新增 `xlsx_blob: Option<Vec<u8>>` 字段
    - `to_opc_package` 在写出 chart part 前检查 `xlsx_blob`，若非空则写出 `/ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` part + 独立的 `/ppt/charts/_rels/chartN.xml.rels` 关系文件
    - 关系 id 用 `rIdXlsxN` 命名空间避免与 slide 的 `rIdChartN` 冲突；全局 `chart_xlsx_global_index` 避免多 slide 冲突
    - 在 chart 模型上设置 `external_data_rid` 后重新 `to_xml`
  - **Slide 高阶 API**：`ShapesMut::add_chart_with_excel(chart_type, data, xlsx_blob, left, top, width, height)` 与 `add_chart` 对称实现，唯一差异是 `ChartEntry.xlsx_blob = Some(xlsx_blob)`
  - 4 个单元测试：`to_xml` 写出 externalData / 省略 externalData（None 时）/ `parse_from_xml` round-trip / 裸 `id` 形式兼容
- **SmartArt 创建 API（TODO-037 创建 API）**：从结构化模型或原始 XML 创建 SmartArt 图形
  - **OOXML 层**：`SmartArtRef::from_rids(dm_rid, lo_rid, qs_rid, cs_rid)` 工厂方法
    - 使用 `XmlWriter` 链式 API 构造 `<a:graphicData uri=".../diagram"><dgm:relIds r:dm=".." r:lo=".." r:qs=".." r:cs=".."/></a:graphicData>` 完整元素 XML
    - 遵守 §5 安全红线（禁止 `format!` 拼接 XML）
    - `oxml/ns.rs` 新增 `NS_DIAGRAM` 命名空间常量
  - **高阶层**：新建 `shape/smartartshape.rs`（~280 行）—— `SmartArtShape` 结构体持有 `OxmlFrame`（`Graphic::SmartArt`）
    - 三个构造器：`new()` / `from_rids()` / `from_frame()`
    - 4 个 getter：`dm_rid()` / `lo_rid()` / `qs_rid()` / `cs_rid()`
    - 4 个单 rid setter：`set_dm_rid()` / `set_lo_rid()` / `set_qs_rid()` / `set_cs_rid()`（任一变更触发 `raw_xml` 整体重建，委托 `SmartArtRef::from_rids` 保证写路径一致）
    - `set_all_rids()` 一次性更新 4 个 rid（仅重建一次）
    - 占位符方法：`set_placeholder` / `clear_placeholder` / `is_placeholder` / `ph_idx` / `ph_type`
    - `Shape` trait 完整实现（旋转始终为 0，OOXML 规范约束）
    - 零 panic 设计：所有方法在不变量被破坏时返回 Option/默认值
  - **ShapeKind 扩展**：`shape/mod.rs` 新增 `ShapeKind::SmartArt(SmartArtShape)` 变体，5 个 match 表达式（shape_type/name/id/left/top/width/height）+ `wrap()` 工厂同步添加分支
  - **Slide 双入口**：
    - `add_smartart_from_xml(data_xml, layout_xml, quick_style_xml, colors_xml, left, top, width, height)` 逃生舱入口（从 4 份原始 XML 创建，round-trip 友好）
    - `add_smartart(data_model, layout_def, quick_style_def, colors_def, left, top, width, height)` 高阶友好入口（从结构化模型创建，调用各模型 `to_xml()` 转 XML 后委托 `add_smartart_from_xml`）
  - `lib.rs` 导出 `SmartArtShape`
  - 7 个单元测试：new 基础 / from_rids / 单 setter 触发重建 / set_all_rids / set_placeholder / Shape trait 几何 / 旋转忽略
- **页脚/日期/幻灯片编号占位符（TODO-007 页脚占位符）**：高阶 setter/getter API
  - `ShapesMut` 新增 6 个高阶 API：
    - `set_footer_text(text) -> bool` / `footer_text() -> Option<String>`
    - `set_date_text(text) -> bool` / `date_text() -> Option<String>`
    - `set_slide_number_text(text) -> bool` / `slide_number_text() -> Option<String>`
  - **查找策略**：仅按 `ph_type` 字符串匹配（`"ftr"` / `"dt"` / `"sldNum"`），不按 `ph_idx` 回退（这三类占位符的 idx 在不同版式中取值不一）
  - 遍历 slide 的 `inner.shapes`，匹配 `OxmlSlideShape::Sp` 且 `is_placeholder && ph_type == target` 的形状，构造新 `TextBody` 调用 `set_text(text)` 替换
  - 零 panic 设计：找不到占位符时 setter 返回 `false`、getter 返回 `None`

#### 修复

- **编译错误与 clippy lint（v6.7）**：修复全部编译错误和 72+ 个 clippy lint
  - 修复 `examples/test_copy_ppt.rs` 未使用的 `Seek` 导入
  - 修复 `src/slide.rs` 4 个 doctest 借用错误（`prs.slides_mut()` 与 `prs.slide_layouts()` 可变借用冲突，改为先获取 layout 再 add_slide）
  - 修复 `src/slide.rs` 1 个 doctest 缺失 `Shape` trait 导入（`add_smartart_from_xml` 示例中 `sa.shape_type()` 需要 `use pptx::shape::Shape;`）
  - 修复 `tests/shape_integration.rs` 5 处 `Inches(x).value()` 返回 f64 与 i64 类型不匹配（改为 `Inches(x).emu().value()`）
  - 修复 `benches/save_pptx.rs` `add_table` 参数顺序错误
  - 修复 `examples/chart_demo.rs` 5 处 `field_reassign_with_default` clippy lint（改为结构体字面量初始化）
  - 修复 `examples/test_pictures_encryption.rs` `needless_range_loop` 和 `ptr_arg` clippy lint
  - 修复 lib 代码 47 个 clippy lint（`from_str`→`parse`、`#[derive(Default)]`+`#[default]`、结构体字面量初始化、`.as_ref()` 替代 `.to_string()` 等）
- **11 个失败单元测试（v6.7）**：修复 diagram 和 chart 模块的 SAX 解析测试
  - 修复 7 个 diagram 测试：quick-xml 0.40 `BytesStart::as_ref()` 不含 `<` 前缀，改为手动累积 raw_xml（`push('<') + push(e.as_ref()) + push('>')`）
  - 修复 4 个 chart 测试：`externalData` 兼容 bare `id` 属性、引入 `in_plot_area` 标志修复次坐标轴检测、修复饼图 categories 断言、修复 malformed XML 测试用未闭合注释触发真正解析错误
- **Fill::Blip 双重 `<a:blip>` 标签（BUG-001）**：修复 sppr.rs 中 `Fill::Blip.write_xml` 生成嵌套无效 XML 的问题
  - 原实现 `w.open("a:blip")` 后又调用 `w.empty_with("a:blip", ...)`，输出 `<a:blip><a:blip .../></a:blip>`
  - 修复为单次 `empty_with("a:blip", ...)`，生成正确的自闭合 `<a:blip r:embed="..."/>`
- **OLE2 容器兼容性**：修复 PowerPoint/WPS 无法打开输出文件的问题
  - 使用 `cfb::Version::V3`（512 字节扇区）替代默认的 V4（4096 字节扇区）
  - 设置根目录 CLSID 为 `64818D10-4F9B-11CF-86EA-00AA00B929E8`（PowerPoint 文件标识）
- **水印 z-order**：修复水印可被编辑的"文本框"问题
  - 水印从"插入到 SpgrContainer 末尾（z-order 最高）"改为"插入到组形状本身之后（z-order 最低的真正子形状）"
  - 注入目标从 Slide 的 PPDrawing 改为 MainMaster 的 PPDrawing（覆盖所有幻灯片）
- **加密 Pictures stream**：修复加密后 WPS 无法打开的问题
  - 不再加密 Pictures stream（WPS 严格检查 Pictures stream 完整性）

#### 依赖

- 新增 `sha1 = "0.10"`（.ppt RC4 CryptoAPI 密钥派生）
- 新增 `cfb = "0.10"`（OLE2/CFB 容器读写）
- 新增 `rand = "0.8"`（加密 salt / verifier 随机生成）

## [0.2.0] - 2026-06-16

### 新增

- **OOXML Agile Encryption 文件加密**：完整的 AES-256-CBC + SHA512 密码派生，输出可被 WPS / PowerPoint / msoffcrypto-python 正确解密
  - 对齐 MS-OFFCRYPTO 规范：PBKDF2 迭代、blockKey 子密钥派生、分段加密（4096 字节/段）
  - OLE2/CFB 容器：DataSpaces 目录结构 + EncryptionInfo + EncryptedPackage
  - 示例：`protect_pptx`（仅加密）、`watermark_and_protect`（水印+加密）
- **水印+加密合并**：`watermark_and_protect` 示例，先加水印再加密，一步到位
- `crypto` 模块：`ModifyProtection` 类型（文档保护标记）
- `OpcPackage::to_bytes()`：序列化到内存（用于加密前的中间步骤）
- 完整的项目文档体系（[`docs/`](.) + [`.trae/`](../.trae/)）
- 8 个 AI 协作 skill（overview / architecture / development / testing / ooxml / extending / coding-standards / debugging）
- 项目级协作开发规范 [`.trae/rules/project_rules.md`](../.trae/rules/project_rules.md)
- 全源码中文注释补全（opc / oxml / shape / presentation / slide / slide_layouts / slide_masters / crypto / examples）
- [`docs/ARCHITECTURE.md`](ARCHITECTURE.md)、[`docs/DEVELOPMENT.md`](DEVELOPMENT.md)、[`docs/TESTING.md`](TESTING.md)、[`docs/OOXML_REFERENCE.md`](OOXML_REFERENCE.md)、[`docs/CONTRIBUTING.md`](CONTRIBUTING.md)

### 改进

- 公共 API 文档注释与项目内沟通统一使用中文
- 错误消息与日志统一小写、句末无标点
- README 中加入对 `docs/` 与 `.trae/skills/` 的引用
- `Relationships::from_xml` 保留内部关系的原始相对路径（`Target::InternalStr`），修复 notes 幻灯片加载失败
- `from_opc` 使用 `resolve_relative_partname` 统一路径解析，正确处理 notes 和 slide 的相对路径
- 修复 clippy 警告：`needless_borrows_for_generic_args`、`manual_div_ceil`、`manual_is_multiple_of`、`too_many_arguments`
- 修复全部 46 个 `cargo doc` 链接警告（未解析链接→反引号、私有项链接→反引号、歧义链接→`enum@Error`）
- 消除 lib 路径中 14 处 `expect()`/`unwrap()` 违规：改用 `ok_or(Error)?` 或索引直接访问
- 新增 `Error::Encryption` 变体，统一加密/解密错误
- 新增 `Error::IndexOutOfRange` 变体，替代 `expect("slide exists")`

### 修复

- 修复 notes 幻灯片 round-trip 测试失败（4 个测试）
- 修复 `Picture::crop` doctest 编译错误

### 依赖

- 新增 `aes = "0.8"`、`cbc = { version = "0.1", features = ["alloc"] }`、`cfb = "0.10"`、`hmac = "0.12"`、`rand = "0.8"`、`sha2 = "0.10"`、`base64 = "0.22"`（加密功能）

## [0.1.0] - 2026-06-13

### 新增

- OPC 容器：zip 读写、`[Content_Types].xml`、Part 关系链
- Presentation / Slide / SlideLayout / SlideMaster XML 模型
- 形状树：AutoShape、Picture、Group、Connector、TextBox、TableShape、Freeform
- TextFrame / Paragraph / Run / 字体（粗/斜/下划线/字号/颜色/字距/基线）
- 常用 DrawingML 几何（50+ preset）
- 表格（按行按列）
- 文档保护（`<p:modifyVerifier>` 注入）
- 水印（向 spTree 注入 sp）
- 示例：hello_pptx、protect_pptx、watermark_pptx
- Python 验证脚本：check_protect.py、check_wm.py、gen_ref.py

### 已知限制

- `Presentation::open` 后内容不全（读路径在建空壳）
- 单 master + 单 layout 强制
- 不支持 Chart / SmartArt
- 不支持 custGeom（自定义几何）
- 媒体 rId 必须 `rIdImg` 前缀（实现细节）
- Performance 未优化
