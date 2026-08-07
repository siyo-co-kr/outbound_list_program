"""암호가 걸린 엑셀 파일을 확인하고 복호화한다."""

from io import BytesIO

import msoffcrypto


def decrypt_excel(file_path, password):
    """읽을 대상을 돌려준다. 암호가 걸려 있으면 복호화한 스트림, 아니면 원본 경로.

    암호가 없는 파일까지 통째로 메모리에 올릴 이유가 없으므로, 암호화 여부만
    확인하고 평문 파일은 pandas가 직접 열도록 경로를 그대로 넘긴다.
    """
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)

        if not office_file.is_encrypted():
            return file_path

        if not password:
            raise ValueError("비밀번호가 필요합니다.")

        office_file.load_key(password=password)
        decrypted = BytesIO()
        office_file.decrypt(decrypted)

    decrypted.seek(0)
    return decrypted
