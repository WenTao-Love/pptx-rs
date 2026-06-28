"""验证 msoffcrypto 是否能解密我们的加密文件。

如果 msoffcrypto 能成功解密，说明加密逻辑基本正确。
如果 msoffcrypto 解密失败，说明加密逻辑有问题。
"""
import sys
import io
import traceback

try:
    import msoffcrypto
except ImportError:
    print("请先安装 msoffcrypto-tool: pip install msoffcrypto-tool")
    sys.exit(1)


def verify_file(filepath, password="pptx-rs-secret"):
    """验证单个文件。"""
    print(f"\n{'='*60}")
    print(f"验证文件: {filepath}")
    print(f"{'='*60}")

    try:
        with open(filepath, "rb") as f:
            officefile = msoffcrypto.OfficeFile(f)

            # 检查是否支持加密
            if not officefile.is_encrypted():
                print("  文件未加密")
                return False

            print(f"  文件类型: {type(officefile).__name__}")
            print(f"  文件已加密: True")

            # 尝试验证密码
            try:
                officefile.load_key(password=password)
                print(f"  密码验证: ✓ 通过")
            except Exception as e:
                print(f"  密码验证: ✗ 失败 - {e}")
                return False

            # 尝试解密
            try:
                out = io.BytesIO()
                officefile.decrypt(out)
                decrypted_data = out.getvalue()
                print(f"  解密: ✓ 成功")
                print(f"  解密后大小: {len(decrypted_data)} bytes")

                # 检查解密后的数据是否有效
                # 解密后的数据应该是有效的 OLE2 文件
                out.seek(0)
                try:
                    import olefile
                    ole = olefile.OleFileIO(out)
                    print(f"  OLE2 结构: ✓ 有效")
                    print(f"  Streams: {ole.listdir()}")
                    ole.close()
                except Exception as e:
                    print(f"  OLE2 结构: ✗ 无效 - {e}")

                return True
            except Exception as e:
                print(f"  解密: ✗ 失败")
                print(f"  错误: {e}")
                traceback.print_exc()
                return False

    except Exception as e:
        print(f"  文件读取失败: {e}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    files = [
        "_test_out/protected_心理账户理论.ppt",
        "_test_out/wm_protected_心理账户理论.ppt",
    ]

    results = []
    for f in files:
        try:
            result = verify_file(f)
            results.append((f, result))
        except Exception as e:
            print(f"  异常: {e}")
            results.append((f, False))

    print(f"\n{'='*60}")
    print("总结:")
    print(f"{'='*60}")
    for f, r in results:
        status = "✓" if r else "✗"
        print(f"  {status} {f}")
