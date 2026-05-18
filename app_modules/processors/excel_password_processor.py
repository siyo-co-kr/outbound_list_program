import msoffcrypto
from io import BytesIO

# 엑셀 파일 암호화 확인 및 암호 복호화 로직
def decrypt_excel(file_path, password):
    with open(file_path, "rb") as f:
        data = BytesIO(f.read())
    office_file = msoffcrypto.OfficeFile(data)
    if office_file.is_encrypted():
        if not password: raise ValueError("비밀번호가 필요합니다.")
        office_file.load_key(password=password)
        decrypt_data = BytesIO()
        office_file.decrypt(decrypt_data)
        return decrypt_data
    return data