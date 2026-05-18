import re

# 컬럼 매핑 사전
COLUMN_MAP = {
    '차트번호' : ['차트번호', '환자번호'],
    '이름' : ['이름', '성명', '환자명', '고객명', '환자', '고객'],
    '연락처' : ['연락처', '휴대폰번호', '전화번호', '핸드폰번호', '핸드폰', '전화', '휴대폰'],
    '마지막 내원일자' : ['마지막내원일자', '내원일', '최종내원일'],
    '생년월일' : ['생년월일', '생일'],
    '아웃바운드 제한 설정': ['아웃바운드제한설정']
}

def normalize_text(text):
    """문자열에서 모든 공백 및 특수문자를 제거하고 소문자로 변환"""
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', str(text)).lower()

def find_actual_column(df_columns, target_key):
    """정규식 정규화된 값을 기준으로 컬럼명을 찾음"""
    target_candidates = [normalize_text(c) for c in COLUMN_MAP[target_key]]

    for col in df_columns:
        if normalize_text(col) in target_candidates:
            return col
    return None