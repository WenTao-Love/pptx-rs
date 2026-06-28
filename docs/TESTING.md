# 测试规范

> pptx-rs 项目的测试规范。覆盖目标、断言策略、组织、CI 集成。
> 速查版见 [.trae/skills/pptx-rs-testing/SKILL.md](../.trae/skills/pptx-rs-testing/SKILL.md)。

## 1. 覆盖率目标

| 模块 | 目标 | 关键场景 |
| --- | --- | --- |
| `units` | ≥ 95% | EMU ↔ Pt ↔ Inches ↔ Cm 转换、RGB 构造 |
| `opc::package` | ≥ 85% | zip 加载/保存、Content-Types 解析 |
| `opc::part` | ≥ 90% | `PartName` 校验、`sibling` / `parent` |
| `opc::rels` | ≥ 90% | rels XML 解析与序列化 round-trip |
| `opc::content_types` | ≥ 85% | override / default 行为 |
| `oxml::*` 序列化 | ≥ 80% | 关键元素属性 + 子元素顺序 |
| `oxml::txbody` | ≥ 85% | Paragraph / Run / RunProperties 全字段 |
| `oxml::sppr` | ≥ 85% | xfrm / fill / line / dash 全枚举 |
| `oxml::shape` | ≥ 80% | Sp / Pic / Group / CxnSp / GraphicFrame |
| `oxml::presentation/slide/master/layout` | ≥ 80% | 必含子元素齐备 |
| `shape::*` | ≥ 75% | 抽象 trait 实现 |
| `ppt97::record` | ≥ 85% | record header 解析、PersistDirectoryAtom 解析、MainMaster 定位 |
| `ppt97::ole` | ≥ 80% | write_stream round-trip、fix_mini_fat 修复 |
| `ppt97::watermark` | ≥ 75% | WatermarkConfig 默认值、SpContainer 注入、recLen 更新 |
| `ppt97::crypto` | ≥ 80% | RC4 加密、密钥派生、CryptSession10Container 构造 |
| 端到端（hello_pptx） | 100% | 必须生成可被 PowerPoint 打开的文件 |
| 端到端（ppt97 examples） | 100% | 必须生成可被 WPS / PowerPoint 打开的 .ppt 文件 |

## 2. 测试组织

### 2.1 单元测试

每个文件底部：

```rust
#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn unit_conversion() {
        assert_eq!(Pt(72.0).emu().value(), 72 * 12_700);
    }
}
```

### 2.2 集成测试

需要多文件协作的放 `tests/`：

```
tests/
├── opc_round_trip.rs        # OpcPackage 加载→保存
├── presentation_save.rs     # Presentation 全流程
└── shape_integration.rs     # 形状→slide→presentation
```

### 2.3 端到端

放 `examples/`，作为"可运行的活文档"。

### 2.4 真实样本

`_test/` 目录提供真实 .pptx 样本（来自 PPT 生态），用于回归测试。

## 3. 测试命名

- **格式**：`<unit>_<scenario>` 或 `<unit>_<action>_<expected>`。
- **示例**：
  - `unit_conversion` ✓
  - `relationships_round_trip` ✓
  - `parse_unknown_geometry_falls_back_to_other` ✓
  - `test1` ✗
  - `test_works` ✗

## 4. 断言策略

| 场景 | 断言 |
| --- | --- |
| XML 含某元素 | `assert!(s.contains("<a:sp>"))` |
| XML 含某属性 | `assert!(s.contains("sz=\"2400\""))` |
| 数量 | `assert_eq!(prs.slides().len(), 2)` |
| 必含子元素 | `assert!(xml.contains("<p:txStyles>"))` |
| 错误传播 | `assert!(matches!(res, Err(Error::Opc(_))))` |
| 浮点 | `assert!((a - b).abs() < 1e-9)` |
| 整数 | `assert_eq!(a, b)` |
| 失败时打印 | `assert!(cond, "actual = {:?}", x)` |

## 5. 常用测试模式

### 5.1 单元

```rust
#[test]
#[should_panic(expected = "must start with '/'")]
fn part_name_rejects_relative() {
    let _ = PartName::new("ppt/slides/slide1.xml").unwrap();
}
```

### 5.2 round-trip

```rust
#[test]
fn relationships_round_trip() {
    let mut r = Relationships::new();
    r.add(Relationship::internal("rId1", RelType::Slide, new_part_name("/ppt/slides/slide1.xml"))).unwrap();
    let xml = r.to_xml();
    let r2 = Relationships::from_xml(&xml).unwrap();
    assert_eq!(r2.len(), 1);
}
```

### 5.3 端到端

```rust
// tests/presentation_save.rs
use pptx::*;

#[test]
fn end_to_end_save() {
    let mut prs = Presentation::new().unwrap();
    let counter = prs.id_counter();
    let _ = prs.slides_mut().add_slide(counter).unwrap();
    let bytes = prs.to_bytes().unwrap();
    assert!(!bytes.is_empty());
}
```

### 5.4 byte-diff（与 python-pptx 对比）

```rust
#[test]
#[ignore] // 默认跳过
fn matches_python_pptx() {
    // ...
}
```

### 5.5 临时文件

```rust
use tempfile::TempDir;

#[test]
fn save_to_temp() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("out.pptx");
    let prs = Presentation::new().unwrap();
    prs.save(&path).unwrap();
    assert!(path.exists());
}
```

## 6. 真实文件验证

### 6.1 .pptx 文件验证

```powershell
# 1) 跑示例
cargo run --example hello_pptx
cargo run --example protect_pptx
cargo run --example watermark_pptx
cargo run --example watermark_and_protect

# 2) 人工验证
#    - PowerPoint 打开 hello.pptx
#    - 打开 _test_out/protected_*.pptx，输入密码 pptx-rs-secret
#    - 打开 _test_out/wm_*.pptx，确认水印
#    - 打开 _test_out/wm_protect_*.pptx，确认水印+密码

# 3) 自动化验证（可选）
python check_protect.py
python check_wm.py
```

### 6.2 .ppt 文件验证

```powershell
# 1) 跑示例（需要 _test/ 目录下有 .ppt 文件）
cargo run --example protect_ppt
cargo run --example watermark_ppt
cargo run --example watermark_and_protect_ppt

# 2) 自动化验证（需要 pip install msoffcrypto-tool）
python verify_ppt_crypto.py              # 验证加密
python verify_ppt_watermark.py           # 验证水印结构
python verify_watermark_and_protect.py   # 验证水印+加密合并

# 3) 人工验证
#    - WPS/PowerPoint 打开 _test_out/protected_*.ppt，输入密码 pptx-rs-secret
#    - WPS/PowerPoint 打开 _test_out/wm_*.ppt，确认水印
#    - WPS/PowerPoint 打开 _test_out/wm_protected_*.ppt，确认水印+密码
```

### 6.3 .ppt 验证脚本说明

| 脚本 | 验证内容 |
| --- | --- |
| `verify_ppt_crypto.py` | msoffcrypto is_encrypted / load_key / decrypt |
| `verify_ppt_watermark.py` | OLE2 结构完整性 + 水印文本数量 |
| `verify_watermark_and_protect.py` | 解密后检查水印文本是否存在 |

## 7. CI 集成

`.github/workflows/ci.yml`：

```yaml
name: ci
on: [push, pull_request]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - uses: Swatinem/rust-cache@v2
      - run: cargo fmt --all -- --check
      - run: cargo clippy --all-targets -- -D warnings
      - run: cargo test --all
      - run: cargo doc --no-deps
      - run: cargo run --example hello_pptx
      - uses: actions/upload-artifact@v4
        with:
          name: hello
          path: hello.pptx
```

## 7A. ppt97 模块测试规范

`.ppt`（PowerPoint 97-2003 二进制格式）测试需要真实样本与外部工具辅助验证。

### 7A.1 单元测试要点

| 子模块 | 关键测试场景 | 断言策略 |
| --- | --- | --- |
| `ppt97::record` | `parse_record_header` 边界（offset+8 越界）、`parse_persist_directory` 多 entry 解析 | `assert!(matches!(res, Err(Error::Ppt97(_))))` |
| `ppt97::ole` | `write_stream` round-trip（写入后读取一致）、`fix_mini_fat` 修复后 num_mini_fat_sectors 正确 | byte-diff |
| `ppt97::watermark` | `WatermarkConfig::default()` 字段值、`inject_watermark` 后 recLen 正确更新 | `assert_eq!(rec_len_new, rec_len_old + inserted_size)` |
| `ppt97::crypto` | `Rc4::new` + `encrypt` round-trip（加密后解密一致）、`make_key` 输出长度 = key_bits/8 | `assert_eq!(key.len(), 16)` |

### 7A.2 集成测试要点

```rust
// tests/ppt97_integration.rs（计划中）
use pptx::ppt97::{add_watermark, encrypt, add_watermark_and_encrypt, WatermarkConfig};

#[test]
#[ignore] // 需要真实 .ppt 样本
fn ppt97_watermark_round_trip() {
    let input = std::path::Path::new("_test/心理账户理论.ppt");
    if !input.exists() {
        return; // 跳过：无样本
    }
    let config = WatermarkConfig::default();
    let output = add_watermark(input, &config).expect("watermark failed");
    assert!(!output.is_empty());
    // 进一步断言：解析输出 OLE2 容器，确认水印 SpContainer 存在
}

#[test]
#[ignore]
fn ppt97_encrypt_round_trip() {
    let input = std::path::Path::new("_test/心理账户理论.ppt");
    if !input.exists() {
        return;
    }
    let output = encrypt(input, "test-password").expect("encrypt failed");
    assert!(!output.is_empty());
    // 进一步断言：CurrentUserAtom.headerToken == 0xF3D1C4DF
}
```

### 7A.3 端到端测试（examples 验证）

```powershell
# 1) 跑全部 .ppt examples（需要 _test/ 目录下有 .ppt 文件）
cargo run --example watermark_ppt
cargo run --example protect_ppt
cargo run --example watermark_and_protect_ppt

# 2) 自动化验证（msoffcrypto-tool + olefile）
python verify_ppt_crypto.py              # 验证加密
python verify_ppt_watermark.py           # 验证水印结构
python verify_watermark_and_protect.py   # 验证水印+加密合并

# 3) 人工验证（WPS / PowerPoint）
#    - 打开 _test_out/wm_*.ppt：水印可见且不可编辑
#    - 打开 _test_out/protected_*.ppt：输入密码 pptx-rs-secret 后正常打开
#    - 打开 _test_out/wm_protected_*.ppt：输入密码后看到水印且不可编辑
```

### 7A.4 .ppt 测试样本管理

- 样本目录：`_test/`（git tracked，**不参与** `cargo test`）
- 输出目录：`_test_out/`（git ignored）
- 样本来源：用户提供真实 .ppt 文件
- 样本命名：建议中文名（如 `心理账户理论.ppt`），便于人工识别

### 7A.5 .ppt 测试 checklist

提交 .ppt 相关代码前：

- [ ] `cargo build --examples` 编译通过
- [ ] `cargo clippy --examples -- -D warnings` 无 ppt97 相关警告
- [ ] `cargo doc --no-deps --lib` 无 ppt97 相关 warning
- [ ] `cargo run --example watermark_ppt` 成功处理 1 个文件
- [ ] `cargo run --example protect_ppt` 成功处理 1 个文件
- [ ] `cargo run --example watermark_and_protect_ppt` 成功处理 1 个文件
- [ ] `python verify_ppt_crypto.py` 全部通过
- [ ] WPS 人工验证：水印可见不可编辑 / 密码打开正常

## 8. 调试单个测试

```powershell
cargo test test_name -- --nocapture
cargo test test_name -- --nocapture --test-threads=1
```

VS Code：在测试函数左侧打断点 → `cargo test test_name` → F5。

## 9. 性能测试（v0.2+ 计划）

```toml
[dev-dependencies]
criterion = "0.5"

[[bench]]
name = "save_pptx"
harness = false
```

```rust
// benches/save_pptx.rs
use criterion::*;
use pptx::*;

fn bench_save(c: &mut Criterion) {
    c.bench_function("save hello", |b| {
        b.iter(|| {
            let prs = Presentation::new().unwrap();
            prs.to_bytes().unwrap()
        });
    });
}

criterion_group!(benches, bench_save);
criterion_main!(benches);
```

## 10. 反模式

- ❌ `assert!(some_fn().is_ok())` —— 应明确消息
- ❌ `unimplemented!()` 当占位
- ❌ `panic!` 不带消息
- ❌ 测试名 `test1` / `test_works`
- ❌ `static mut` / `lazy_static` 共享状态
- ❌ 跳过 `cargo fmt` 后提交

## 11. 测试 checklist

提交前：

- [ ] `cargo fmt --all`
- [ ] `cargo clippy --all-targets -- -D warnings`
- [ ] `cargo test --all` 全过
- [ ] `cargo test -- --ignored` 跑过重资源测试
- [ ] `cargo doc --no-deps` 无 warning
- [ ] 至少跑过 `cargo run --example hello_pptx` + 肉眼打开

## 12. 故障排查

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| `assertion failed: xml.contains(...)` | 属性顺序变化 | 调整 `write_xml` |
| `zip: InvalidArchive` | Content-Type 错 | 检查 `derive_content_type` |
| `should_panic` 没触发 | panic 信息改了 | 用 `#[should_panic = "regex"]` |
| 偶发失败 | 共享可变状态 | `--test-threads=1` |
| 浮点 assert 失败 | 精度 | `(a - b).abs() < 1e-9` |
