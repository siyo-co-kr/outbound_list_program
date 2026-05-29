from .excel_password_processor import decrypt_excel
import pandas as pd
from .column_utils import find_actual_column

def outbound_limit(file_path, password):
    # 1. 복호화 실행
    data_to_read = decrypt_excel(file_path, password)

    # 2. 데이터 로드
    df = pd.read_excel(data_to_read, sheet_name=0)
    df.columns = df.columns.str.strip()

    # 3. 컬럼 매핑 적용
    mapping_result = {}
    required_keys = ['차트번호', '이름', '생년월일', '연락처', '마지막 내원일자']

    for key in required_keys:
        actual_col = find_actual_column(df.columns.tolist(), key)
        if not actual_col:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {key}")
        mapping_result[key] = actual_col

    # 표준 컬럼명으로 변경
    df = df.rename(columns={v: k for k, v in mapping_result.items()})

    # 4. 숫자 정제 (특수문자 및 소수점 제거)
    for col in ['생년월일', '연락처', '마지막 내원일자']:
        df[col] = df[col].astype(str).str.replace(r'[^0-9]', '', regex=True)

    # 5. 연락처 보정 및 유효성 검사
    mask = (df['연락처'].str.len() == 10) & (df['연락처'].str.startswith('10'))
    df.loc[mask, '연락처'] = '010' + df['연락처'].str[2:]
    df = df[df['연락처'].str.match(r'^010\d{8}$')]

    # 6. '아웃바운드 제한 설정' 컬럼이 없으면 생성하고 'O' 입력
    if '아웃바운드 제한 설정' not in df.columns:
        df['아웃바운드 제한 설정'] = 'O'
    else:
        df['아웃바운드 제한 설정'] = 'O'

    # 7. 최종 출력할 컬럼 정의
    final_keys = required_keys + ['아웃바운드 제한 설정']
    df_final = df[final_keys].copy()

    # 결측치(NaN) 행 제거
    df_final = df_final.dropna(subset=final_keys)

    # 문자열 공백("", " ", "nan") 필터링
    for col in final_keys:
        df_final = df_final[df_final[col].astype(str).str.strip() != '']
        df_final = df_final[df_final[col].astype(str).str.lower() != 'nan']

    return df_final.reset_index(drop=True)
