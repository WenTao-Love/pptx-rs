---
name: "pptx-rs-testing"
description: "pptx-rs 测试规范与实践：单元/集成/端到端测试、覆盖率目标、断言策略、PowerPoint 验证。Invoke when user asks how to test, write tests, debug failing tests, or interpret test failures."
---

# pptx-rs 测试指南

> 对应 [docs/TESTING.md](../../../../docs/TESTING.md)。

## 覆盖率目标

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
| 端到端（hello_pptx） | 100% | 必须生成可被 PowerPoint 打开的文件 |

## 测试组织

```rust
// 每个文件底部放：
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xxx() { ... }
}
```

集成测试（需要多文件协作）放 `tests/`：

```
tests/
├── opc_round_trip.rs        # OpcPackage 加载→保存
├── presentation_save.rs     # Presentation 全流程
└── shape_integration.rs     # 形状→slide→presentation
```

## 单元测试模板

```rust
#[test]
fn unit_conversion() {
    // 入参与期望
    assert_eq!(Pt(72.0).emu().value(), 72 * 12_700);
    // 浮点
    assert!((Emu(914_400).inches() - 1.0).abs() < 1e-9);
}

#[test]
#[should_panic(expected = "must start with '/'")]
fn part_name_rejects_relative() {
    let _ = PartName::new("ppt/slides/slide1.xml").unwrap();
}

#[test]
fn xml_escape_unicode() {
    let s = "中文 & <tag>";
    assert_eq!(escape(s), "中文 &amp; &lt;tag&gt;");
}
```

## 端到端测试

```rust
// tests/presentation_save.rs
use pptx::*;

#[test]
fn end_to_end_hello() {
    let mut prs = Presentation::new().unwrap();
    let counter = prs.id_counter();
    let slide = prs.slides_mut().add_slide(counter).unwrap();
    slide.shapes_mut()
        .add_textbox_with_text(Inches(1.0), Inches(1.0), Inches(8.0), Inches(1.0), "hi")
        .unwrap();

    let bytes = prs.to_bytes().unwrap();
    assert!(!bytes.is_empty());
    assert!(bytes.len() > 1000); // 必有默认 master/layout/theme

    // round-trip：再打开
    let prs2 = Presentation::load_bytes(&bytes).unwrap();
    assert!(prs2.slide_width().value() > 0);
}
```

## round-trip 测试

**目标**：序列化与解析（若支持）应互逆。

```rust
#[test]
fn relationships_round_trip() {
    let mut r = Relationships::new();
    r.add(Relationship::internal("rId1", RelType::Slide, new_part_name("/ppt/slides/slide1.xml"))).unwrap();
    r.add(Relationship::internal("rId2", RelType::SlideLayout, new_part_name("/ppt/slideLayouts/slideLayout1.xml"))).unwrap();
    let xml = r.to_xml();
    let r2 = Relationships::from_xml(&xml).unwrap();
    assert_eq!(r2.len(), 2);
    assert!(r2.get("rId1").is_some());
    assert!(r2.get("rId2").is_some());
}
```

## byte-diff 测试（与 python-pptx 对比）

```rust
#[test]
#[ignore] // 默认跳过，CI 可启用
fn matches_python_pptx_template() {
    // 1) 调用本库生成
    let bytes1 = generate_with_pptx_rs();

    // 2) 调用 python-pptx 生成
    let bytes2 = generate_with_python_pptx();

    // 3) 解析 zip 比对关键 part
    let z1 = parse_zip(&bytes1);
    let z2 = parse_zip(&bytes2);
    for name in &["[Content_Types].xml", "ppt/presentation.xml", "ppt/slideMasters/slideMaster1.xml"] {
        assert_eq!(z1[name], z2[name], "diff in {name}");
    }
}
```

## 真实文件验证（半自动）

```powershell
# 1) 跑示例
cargo run --example hello_pptx
cargo run --example protect_pptx
cargo run --example watermark_pptx

# 2) 用 PowerPoint 打开 hello.pptx
# 3) 用 PowerPoint 打开 _test_out/protected_*.pptx 并尝试编辑（应要求输入密码）
# 4) 用 PowerPoint 打开 _test_out/wm_*.pptx，确认每页有水印
```

`_test/` 目录提供两份真实样本（来自 PPT 生态，非本项目生成）作为回归测试输入。

## 断言策略

| 场景 | 断言 |
| --- | --- |
| XML 含某元素 | `assert!(s.contains("<a:sp>"))` |
| XML 含某属性 | `assert!(s.contains("sz=\"2400\""))` |
| 数量 | `assert_eq!(prs.slides().len(), 2)` |
| 必含子元素 | `assert!(xml.contains("<p:txStyles>"))` |
| 错误传播 | `assert!(matches!(res, Err(Error::Opc(_))))` |
| 浮点 | `assert!((a - b).abs() < 1e-9)` |
| 整数 | `assert_eq!(a, b)` |

## 模拟 IO

```rust
use std::io::Cursor;

fn read_zip_part(bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(cursor).unwrap();
    let mut s = String::new();
    zip.by_name(name).unwrap().read_to_string(&mut s).unwrap();
    s
}
```

## 临时文件

dev-dep 已声明 `tempfile`：

```rust
use tempfile::TempDir;

#[test]
fn save_to_tempfile() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("out.pptx");
    let prs = Presentation::new().unwrap();
    prs.save(&path).unwrap();
    assert!(path.exists());
}
```

## CI 集成建议

`.github/workflows/ci.yml`（参考）：

```yaml
name: ci
on: [push, pull_request]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - run: cargo fmt --all -- --check
      - run: cargo clippy --all-targets -- -D warnings
      - run: cargo test --all
      - run: cargo doc --no-deps
      - run: cargo run --example hello_pptx
```

## 常见失败模式

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| `assertion failed: xml.contains("...")` | 属性顺序变化 | 调整 `write_xml` 中的属性 push 顺序 |
| `zip: InvalidArchive` | Content-Type 写错 | 检查 `derive_content_type` |
| `Presentation::load_bytes` 报 `not implemented` | 当前读路径只建空壳 | 0.1.0 已知；写测试用 `to_bytes` → `load_bytes` 即可 |
| `should_panic` 没触发 | panic 信息改了 | 用 `#[should_panic = "regex"]` |
| `cargo test` 偶发失败 | `Rc<Cell<u32>>` 并发 | 加 `--test-threads=1` |

## 调试特定测试

```powershell
# 只跑一个测试
cargo test test_name -- --nocapture

# 进入调试（VS Code）
# 在测试函数左侧打断点 → `cargo test test_name` → F5

# 用 println 调试
#[test]
fn foo() {
    let s = ...;
    println!("s = {:?}", s);
    assert!(false); // 故意失败看输出
}
```

## 性能测试（占位）

0.1.0 未集成 criterion；如需：

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

## 反模式

- ❌ `assert!(some_fn().is_ok())` —— 应明确 `assert!(some_fn().is_ok(), "{:?}", some_fn())`
- ❌ `unimplemented!()` 当占位 —— 改用 `Error::not_implemented("...")` 返回
- ❌ 在测试里 `panic!` 而不写消息
- ❌ 测试名只叫 `test1` / `test_works`
- ❌ 共享可变状态（`static mut` / `lazy_static`）—— 用 thread-local 或单线程 `--test-threads=1`
- ❌ 跳过 `cargo fmt` 后提交

## Solution Patterns

### Pattern 1: OOXML 序列化测试用 XML 包含断言

```rust
// ✅ 检查 XML 包含关键元素
#[test]
fn sp_write_xml_includes_placeholder() {
    let sp = Sp { is_placeholder: true, ph_type: None, ..Default::default() };
    let mut w = XmlWriter::new();
    sp.write_xml(&mut w);
    let xml = w.into_string();
    assert!(xml.contains(r#"<p:ph type="body"/>"#), "placeholder should have default type");
}

// ❌ 只检查方法不报错
#[test]
fn sp_write_xml_works() {
    let sp = Sp::default();
    let mut w = XmlWriter::new();
    sp.write_xml(&mut w);  // 没验证输出内容
}
```

**适用场景**：任何 `write_xml` 实现的测试。
**不适场景**：round-trip 测试（应比较完整 XML）。

### Pattern 2: round-trip 测试验证序列化/解析互逆

```rust
// ✅ 完整 round-trip
#[test]
fn relationships_round_trip() {
    let mut r = Relationships::new();
    r.add(Relationship::internal("rId1", RelType::Slide, ...)).unwrap();
    let xml = r.to_xml();
    let r2 = Relationships::from_xml(&xml).unwrap();
    assert_eq!(r2.len(), 1);
    assert!(r2.get("rId1").is_some());
}

// ❌ 只测序列化不测解析
#[test]
fn relationships_to_xml() {
    let r = Relationships::new();
    let xml = r.to_xml();
    assert!(!xml.is_empty());  // 没验证解析回来
}
```

**适用场景**：有 `to_xml` + `from_xml` 的类型。
**不适场景**：只写不读的类型（如 `PresentationRoot`）。

## Workflow

### 选择测试类型

```
1. 测试的是什么？
   → 单个函数/方法：单元测试（文件底部 #[cfg(test)]）
   → 多文件协作：集成测试（tests/）
   → 完整 .pptx 生成：端到端测试（examples/ + 手动验证）

2. 需要验证什么？
   → XML 输出正确：XML 包含断言
   → 序列化/解析互逆：round-trip 测试
   → 与 python-pptx 一致：byte-diff 测试
   → PowerPoint 能打开：手动验证 + CI 示例

3. 测试是否涉及 IO？
   → 是：用 tempfile / Cursor<Vec<u8>>
   → 否：纯内存测试
```

## Review Checklist

- [ ] 每个 `pub` 函数至少 1 个测试
- [ ] 测试命名描述场景（非 `test1`）
- [ ] `write_xml` 测试检查输出 XML 内容
- [ ] round-trip 测试覆盖序列化/解析
- [ ] 端到端测试生成可被 PowerPoint 打开的文件
- [ ] 无 `unimplemented!()` 占位（用 `Error::not_implemented`）
- [ ] `cargo test --all` 通过
- [ ] `cargo clippy --all-targets -- -D warnings` 通过

## Verification Commands

```bash
# 运行全部测试
cargo test --all

# 运行单个测试（带输出）
cargo test test_name -- --nocapture

# 运行特定模块测试
cargo test oxml::

# 串行运行（排查并发问题）
cargo test --all -- --test-threads=1

# 运行被忽略的测试
cargo test --all -- --ignored

# 运行文档测试
cargo test --doc

# Clippy 检查
cargo clippy --all-targets -- -D warnings

# 运行示例验证
cargo run --example hello_pptx
cargo run --example watermark_pptx
cargo run --example protect_pptx
```

## Cross-References

- [pptx-rs-architecture](../pptx-rs-architecture/SKILL.md) — 架构详解
- [pptx-rs-debugging](../pptx-rs-debugging/SKILL.md) — 调试指南（测试失败排查）
- [pptx-rs-extending](../pptx-rs-extending/SKILL.md) — 扩展指南（添加新测试）
- [rust-coding-standards](../rust-coding-standards/SKILL.md) — Rust 编码规范
