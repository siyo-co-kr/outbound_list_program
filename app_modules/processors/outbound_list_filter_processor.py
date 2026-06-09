from .excel_password_processor import decrypt_excel
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from .column_utils import find_actual_column


def outbound_list_filter(file_path, password, period_type, period_value, use_birth, start_date, end_date):
    """
    엑셀 파일을 복호화하고 내원일/생년월일 필터링 및 연락처 정제를 수행합니다.
    """
    # 1. 복호화 실행
    data_to_read = decrypt_excel(file_path, password)

    # 2. 데이터 로드 (시트가 2개 이상이면 두 번째 시트, 1개면 첫 번째 시트 사용)
    excel_file = pd.ExcelFile(data_to_read)
    sheet_idx = 1 if len(excel_file.sheet_names) > 1 else 0
    df = pd.read_excel(excel_file, sheet_name=sheet_idx)
    df.columns = df.columns.str.strip()

    # 3. 컬럼 매핑 적용
    mapping_result = {}
    required_keys = ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자']
    if use_birth:
        required_keys.append('생년월일')

    for key in required_keys:
        actual_col = find_actual_column(df.columns.tolist(), key)
        if not actual_col:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {key}")
        mapping_result[key] = actual_col

    # 표준 컬럼명으로 변경
    df = df.rename(columns={v: k for k, v in mapping_result.items()})

    # 4. 필터링 (내원일 기준)
    if period_type != "전체":
        df['마지막 내원일자'] = pd.to_datetime(df['마지막 내원일자'], errors='coerce')

    if period_type == "개월":
        threshold = datetime.now() - relativedelta(months=int(period_value))
        df = df[df['마지막 내원일자'] >= threshold]
    elif period_type == "년":
        threshold = datetime.now() - relativedelta(years=int(period_value))
        df = df[df['마지막 내원일자'] >= threshold]

    # 5. 필터링 (생년월일 기준)
    if use_birth:
        df['생년월일'] = pd.to_datetime(df['생년월일'], errors='coerce')
        s_date = datetime.combine(start_date, datetime.min.time())
        e_date = datetime.combine(end_date, datetime.max.time())
        df = df[(df['생년월일'] >= s_date) & (df['생년월일'] <= e_date)]

    # 6. 휴대폰번호 정제 및 중복 제거
    df['휴대폰번호'] = df['휴대폰번호'].astype(str).str.replace(r'[^0-9]', '', regex=True)
    mask = (df['휴대폰번호'].str.len() == 10) & (df['휴대폰번호'].str.startswith('10'))
    df.loc[mask, '휴대폰번호'] = '010' + df['휴대폰번호'].str[2:]

    df_final = df.drop_duplicates(subset=['휴대폰번호'])
    df_final = df_final[df_final['휴대폰번호'].str.match(r'^010\d{8}$')].copy()

    # 7. 최종 데이터 반환 (차트번호 원본 유지)
    return df_final[['차트번호', '환자 이름', '휴대폰번호']]
