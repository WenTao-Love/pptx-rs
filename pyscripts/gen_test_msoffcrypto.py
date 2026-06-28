"""用 msoffcrypto-python 库加密 PPTX 文件。"""
import msoffcrypto
import io

input_path = "_test/文旅IP人设打造抖音短视频运营方案.pptx"
output_path = "_test_out/msoffcrypto_encrypted.pptx"
password = "pptx-rs-secret"

# 加密：不需要 load_key，直接 encrypt
with open(input_path, "rb") as f_in:
    office_file = msoffcrypto.OfficeFile(f_in)
    with open(output_path, "wb") as f_out:
        office_file.encrypt(password, f_out)

print(f"Created: {output_path}")

# 验证：尝试解密
with open(output_path, "rb") as f_enc:
    enc_file = msoffcrypto.OfficeFile(f_enc)
    enc_file.load_key(password=password)
    decrypted = io.BytesIO()
    enc_file.decrypt(decrypted)
    print(f"Decryption test: OK ({decrypted.tell()} bytes)")
