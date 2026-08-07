from datetime import datetime
from typing import NamedTuple

import pandas as pd
from dateutil.relativedelta import relativedelta

from .column_utils import resolve_columns
from .excel_password_processor import decrypt_excel
# 결과 컬럼은 템플릿 머리글과 반드시 같아야 하므로 한 곳에서만 정의한다
from .template_writer import TEMPLATE_COLUMNS as OUTPUT_COLUMNS

# 기간 필터 단위 -> relativedelta 인자명
PERIOD_UNITS = {'개월': 'months', '년': 'years'}

# 머리글이 첫 행이 아닌 파일을 위해 위에서부터 훑어볼 행 수
_MAX_HEADER_SCAN = 8

# 엑셀 날짜 일련번호로 인정할 범위.
# 내원일자는 작은 숫자를 날짜로 오인하지 않도록 하한을 높게 두지만, 생년월일에
# 같은 하한을 쓰면 1954년 이전 출생자가 통째로 누락되므로 따로 잡는다.
_SERIAL_MIN, _SERIAL_MAX = 20000, 60000   # 1954-10-03 ~ 2064-04-08
_BIRTH_SERIAL_MIN = 2                     # 1900-01-01
_EXCEL_EPOCH = '1899-12-30'


class FilterResult(NamedTuple):
    """필터 결과와, 사용자가 결과를 검증할 수 있도록 하는 처리 내역"""
    df: pd.DataFrame
    report: list      # 단계별 건수
    warnings: list    # 주의가 필요한 사항 (컬럼 선택, 해석 실패 등)


def _threshold(period_type, period_value):
    """기준 시점 계산. period_type이 '전체'면 None (필터 미적용)"""
    unit = PERIOD_UNITS.get(period_type)
    if unit is None:
        return None
    return datetime.now() - relativedelta(**{unit: int(period_value)})


def _dedupe_headers(names):
    """같은 이름의 컬럼이 여러 개면 뒤쪽에 번호를 붙여 고유하게 만든다.

    빈 셀은 'nan'/'None' 같은 엉뚱한 컬럼명이 되지 않도록 빈 이름으로 통일한다.
    """
    seen = {}
    result = []
    for name in names:
        blank = name is None or name != name          # None 또는 NaN
        text = '' if blank else str(name).strip()
        if text in seen:
            seen[text] += 1
            text = f"{text}.{seen[text]}"
        else:
            seen[text] = 0
        result.append(text)
    return result


def _load_sheet(data_to_read, required_keys):
    """필수 컬럼이 모두 인식되는 (시트, 머리글 행)을 찾아 데이터프레임으로 로드한다.

    기존 동작(시트가 여러 개면 두 번째 시트, 머리글은 첫 행)을 최우선으로 시도하고,
    거기서 컬럼을 찾지 못할 때만 다른 시트·머리글 행으로 탐색 범위를 넓힌다.

    모든 값을 문자열로 읽는다. 숫자로 추론되면 '00123' -> 123 처럼 앞자리 0이
    사라지고, 빈칸이 섞인 숫자 컬럼은 1012345678.0 처럼 소수점이 붙어버린다.

    반환: (df, mapping, notes)
      notes - 어느 시트·머리글 행을 썼는지 등, 엉뚱한 위치를 읽었을 때
              사용자가 바로 알아챌 수 있게 하는 안내 문구
    """
    with pd.ExcelFile(data_to_read) as excel_file:
        names = excel_file.sheet_names
        preferred = 1 if len(names) > 1 else 0
        sheet_order = [preferred] + [i for i in range(len(names)) if i != preferred]

        first_error = None
        for sheet in sheet_order:
            raw = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str)
            for header_row in range(min(_MAX_HEADER_SCAN, len(raw))):
                columns = _dedupe_headers(raw.iloc[header_row].tolist())
                try:
                    mapping, notes = resolve_columns(columns, required_keys)
                except ValueError as exc:
                    first_error = first_error or exc
                    continue

                df = raw.iloc[header_row + 1:].copy()
                df.columns = columns
                df = df.reset_index(drop=True)

                if sheet != preferred:
                    notes.append(f"'{names[preferred]}' 시트에 필수 컬럼이 없어 "
                                 f"'{names[sheet]}' 시트를 사용했습니다.")
                if header_row > 0:
                    notes.append(f"머리글을 '{names[sheet]}' 시트 {header_row + 1}행에서 찾았습니다.")
                return df, mapping, notes

    # 시트가 모두 비어 있으면 위 루프가 한 번도 돌지 않아 first_error가 None으로 남는다.
    # 어디까지 찾아봤는지 함께 알려야, 머리글이 스캔 범위보다 아래 있는 파일에서
    # "컬럼이 없다"는 메시지만 보고 원인을 못 찾는 일이 없다.
    detail = first_error or ValueError("읽을 수 있는 데이터가 없습니다.")
    raise ValueError(f"{detail} "
                     f"(시트 {len(names)}개의 상위 {_MAX_HEADER_SCAN}행까지 머리글을 찾아봤습니다.)")


def _parse_dates(series, serial_min=_SERIAL_MIN):
    """엑셀 날짜 셀·문자열·YYYYMMDD 숫자·엑셀 일련번호를 모두 datetime으로 변환한다.

    모든 값을 문자열로 읽기 때문에 형식이 뒤섞여 있다. pandas의 자동 추론은
    첫 값의 형식을 나머지에 그대로 적용해 버려 '20260702' 같은 값이 통째로
    NaT가 되므로, 형식별로 나누어 해석한다.

    serial_min - 엑셀 일련번호로 인정할 하한. 생년월일은 고령 환자를 잃지 않도록
    내원일자보다 낮은 값을 쓴다.
    """
    text = series.fillna('').astype(str).str.strip()
    digits = text.str.replace(r'\D', '', regex=True)

    # 1) '2026-07-01', '2026-07-01 00:00:00', '2026/07/01' 등 일반 날짜 문자열
    result = pd.to_datetime(text, errors='coerce', format='mixed')

    # 2) YYYYMMDD 8자리 숫자
    is_ymd = digits.str.fullmatch(r'\d{8}')
    if is_ymd.any():
        result = result.fillna(
            pd.to_datetime(digits.where(is_ymd), format='%Y%m%d', errors='coerce')
        )

    # 3) 엑셀 날짜 일련번호 (서식이 '일반'으로 저장된 날짜 셀)
    serial = pd.to_numeric(text.where(~is_ymd), errors='coerce')
    in_range = serial.between(serial_min, _SERIAL_MAX)
    if in_range.any():
        result = result.fillna(
            pd.to_datetime(serial.where(in_range), unit='D', origin=_EXCEL_EPOCH, errors='coerce')
        )

    return result


def _clean_phone(series):
    """연락처에서 숫자만 남기고 국가번호·앞자리 0 누락을 보정한다"""
    text = series.fillna('').astype(str).str.strip()

    # 엑셀이 숫자로 저장한 셀은 '1012345678.0'처럼 소수점이 붙어 읽힐 수 있다
    # (.xls는 xlrd가 모든 숫자를 float으로 돌려준다). 먼저 떼어내지 않으면
    # 숫자만 남길 때 뒤에 0이 하나 더 붙어 멀쩡한 번호가 통째로 탈락한다.
    text = text.str.replace(r'\.0+$', '', regex=True)

    digits = text.str.replace(r'\D', '', regex=True)

    # +82 국제 표기: 821012345678 / 8201012345678 -> 국가번호 제거
    intl = digits.str.fullmatch(r'82\d{9,11}')
    digits = digits.mask(intl, digits.str[2:])

    # 엑셀이 앞자리 0을 지운 경우: 1012345678 -> 01012345678
    dropped_zero = digits.str.fullmatch(r'10\d{8}')
    digits = digits.mask(dropped_zero, '0' + digits)

    return digits


def outbound_list_filter(file_path, password, period_type, period_type_old, period_value,
                         period_value_old, use_birth, start_date, end_date):
    """엑셀 파일을 복호화하고 내원일/생년월일 필터링 및 연락처 정제를 수행한다.

    반환: FilterResult(df, report, warnings)
    """
    report = []
    warnings = []

    # 1. 복호화 및 로드 (전부 문자열로 읽어 원본 표기를 보존)
    data_to_read = decrypt_excel(file_path, password)

    required_keys = ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자']
    if use_birth:
        required_keys.append('생년월일')

    df, mapping, notes = _load_sheet(data_to_read, required_keys)
    warnings.extend(notes)

    # 2. 필요한 컬럼만 골라 표준 컬럼명으로 변경
    #    (필요한 컬럼만 남기므로 rename 과정에서 이름이 겹쳐 값이 섞일 수 없다)
    df = df[list(mapping.values())].copy()
    df.columns = list(mapping.keys())

    # 완전히 빈 행 제거 (엑셀 하단 여백)
    df = df.dropna(how='all').reset_index(drop=True)
    total = len(df)
    report.append(f"원본 {total}건")

    # 3. 내원일 기준 필터
    visited_after = _threshold(period_type, period_value)           # 이 시점 이후에 내원한 대상만
    visited_before = _threshold(period_type_old, period_value_old)  # 이 시점 이전이 마지막인 대상만

    if visited_after is not None or visited_before is not None:
        df['마지막 내원일자'] = _parse_dates(df['마지막 내원일자'])

        unparsed = int(df['마지막 내원일자'].isna().sum())
        if unparsed:
            warnings.append(f"마지막 내원일자를 해석하지 못한 {unparsed}건은 제외했습니다.")
            # 제외 건수도 내역에 남겨야 단계별 숫자가 앞뒤로 맞아떨어진다
            before = len(df)
            df = df[df['마지막 내원일자'].notna()]
            report.append(f"마지막 내원일자 해석 가능: {len(df)}건 (-{before - len(df)})")

        if visited_after is not None:
            before = len(df)
            df = df[df['마지막 내원일자'] >= visited_after]
            report.append(
                f"최근 {period_value}{period_type} 이내 내원 "
                f"({visited_after:%Y-%m-%d} 이후): {len(df)}건 (-{before - len(df)})"
            )
        if visited_before is not None:
            before = len(df)
            df = df[df['마지막 내원일자'] <= visited_before]
            report.append(
                f"{period_value_old}{period_type_old} 이상 미내원 "
                f"({visited_before:%Y-%m-%d} 이전): {len(df)}건 (-{before - len(df)})"
            )

    # 4. 생년월일 기준 필터
    if use_birth:
        df['생년월일'] = _parse_dates(df['생년월일'], serial_min=_BIRTH_SERIAL_MIN)

        unparsed = int(df['생년월일'].isna().sum())
        if unparsed:
            warnings.append(f"생년월일을 해석하지 못한 {unparsed}건은 제외했습니다.")
            before = len(df)
            df = df[df['생년월일'].notna()]
            report.append(f"생년월일 해석 가능: {len(df)}건 (-{before - len(df)})")

        s_date = datetime.combine(start_date, datetime.min.time())
        e_date = datetime.combine(end_date, datetime.max.time())
        before = len(df)
        df = df[(df['생년월일'] >= s_date) & (df['생년월일'] <= e_date)]
        report.append(
            f"생년월일 {start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}: {len(df)}건 (-{before - len(df)})"
        )

    # 5. 휴대폰번호 정제 -> 형식 검증 -> 중복 제거 (검증을 먼저 해야 유효한 번호끼리 비교된다)
    df = df.copy()
    df['휴대폰번호'] = _clean_phone(df['휴대폰번호'])

    before = len(df)
    df = df[df['휴대폰번호'].str.fullmatch(r'010\d{8}')]
    if before != len(df):
        report.append(f"휴대폰번호 형식 유효: {len(df)}건 (-{before - len(df)})")

    before = len(df)
    df = df.drop_duplicates(subset=['휴대폰번호'], keep='first')
    if before != len(df):
        report.append(f"중복 연락처 제거: {len(df)}건 (-{before - len(df)})")

    # 6. 최종 데이터 (차트번호·이름은 원본 표기 그대로)
    result = df[OUTPUT_COLUMNS].reset_index(drop=True)
    report.append(f"최종 {len(result)}건")

    return FilterResult(df=result, report=report, warnings=warnings)
