"""원본 환자 정보 엑셀을 읽어 내원일/생년월일로 필터링하고 연락처를 정제한다.

값은 전부 문자열로 다룬다. 숫자로 해석하면 '00123'의 앞자리 0이 사라지고,
연락처에 소수점이 붙어('1012345678.0') 번호가 통째로 망가진다.
"""

import math
from datetime import datetime, timedelta
from typing import NamedTuple

import openpyxl
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
_EXCEL_EPOCH = datetime(1899, 12, 30)

# 날짜 문자열로 인정할 형식. 위에서부터 순서대로 시도한다.
_DATE_FORMATS = (
    '%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M',
    '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m', '%m/%d/%Y',
)

_OLE2_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'   # 구형 .xls 파일 머리표식


class FilterResult(NamedTuple):
    """필터 결과와, 사용자가 결과를 검증할 수 있도록 하는 처리 내역"""
    rows: list        # {표준 컬럼명: 값} 목록
    report: list      # 단계별 건수
    warnings: list    # 주의가 필요한 사항 (컬럼 선택, 해석 실패 등)


def _threshold(period_type, period_value):
    """기준 시점 계산. period_type이 '전체'면 None (필터 미적용)"""
    unit = PERIOD_UNITS.get(period_type)
    if unit is None:
        return None
    return datetime.now() - relativedelta(**{unit: int(period_value)})


# --------------------------------------------------------------------------- 읽기

def _cell_text(value):
    """셀 값을 원본 표기 그대로의 문자열로. 빈 셀은 None."""
    if value is None:
        return None
    if isinstance(value, str):
        return value.strip() or None
    # xlrd는 모든 수를 float으로 준다. 정수로 떨어지면 소수점을 붙이지 않는다.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _peek(source, size=8):
    """파일 형식 판별용 머리 바이트. 경로와 스트림 양쪽을 받는다."""
    if hasattr(source, 'read'):
        position = source.tell()
        head = source.read(size)
        source.seek(position)
        return head
    with open(source, 'rb') as f:
        return f.read(size)


def _read_xlsx(source):
    workbook = openpyxl.load_workbook(source, read_only=True, data_only=True)
    try:
        return [(sheet.title,
                 [[_cell_text(cell.value) for cell in row] for row in sheet.iter_rows()])
                for sheet in workbook.worksheets]
    finally:
        workbook.close()


def _read_xls(source):
    # .xls를 읽을 때만 부담하도록 지연 임포트
    import xlrd

    if hasattr(source, 'read'):
        source.seek(0)
        book = xlrd.open_workbook(file_contents=source.read())
    else:
        book = xlrd.open_workbook(source)

    sheets = []
    for sheet in book.sheets():
        grid = []
        for index in range(sheet.nrows):
            row = []
            for cell in sheet.row(index):
                if cell.ctype == xlrd.XL_CELL_DATE:
                    row.append(str(xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)))
                else:
                    row.append(_cell_text(cell.value))
            grid.append(row)
        sheets.append((sheet.name, grid))
    return sheets


def _read_grid(source):
    """엑셀을 [(시트명, 2차원 격자)]로 읽는다. 확장자가 아닌 내용으로 형식을 판별한다."""
    return _read_xls(source) if _peek(source) == _OLE2_MAGIC else _read_xlsx(source)


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


def _load_sheet(source, required_keys):
    """필수 컬럼이 모두 인식되는 (시트, 머리글 행)을 찾아 행 목록으로 읽는다.

    기존 동작(시트가 여러 개면 두 번째 시트, 머리글은 첫 행)을 최우선으로 시도하고,
    거기서 컬럼을 찾지 못할 때만 다른 시트·머리글 행으로 탐색 범위를 넓힌다.

    반환: (rows, notes)
      rows  - {표준 컬럼명: 값} 목록
      notes - 어느 시트·머리글 행을 썼는지 등, 엉뚱한 위치를 읽었을 때
              사용자가 바로 알아챌 수 있게 하는 안내 문구
    """
    sheets = _read_grid(source)
    preferred = 1 if len(sheets) > 1 else 0
    sheet_order = [preferred] + [i for i in range(len(sheets)) if i != preferred]

    first_error = None
    for index in sheet_order:
        name, grid = sheets[index]
        for header_row in range(min(_MAX_HEADER_SCAN, len(grid))):
            columns = _dedupe_headers(grid[header_row])
            try:
                mapping, notes = resolve_columns(columns, required_keys)
            except ValueError as exc:
                first_error = first_error or exc
                continue

            position = {key: columns.index(column) for key, column in mapping.items()}
            rows = [{key: row[at] if at < len(row) else None for key, at in position.items()}
                    for row in grid[header_row + 1:]]

            if index != preferred:
                notes.append(f"'{sheets[preferred][0]}' 시트에 필수 컬럼이 없어 "
                             f"'{name}' 시트를 사용했습니다.")
            if header_row > 0:
                notes.append(f"머리글을 '{name}' 시트 {header_row + 1}행에서 찾았습니다.")
            return rows, notes

    # 시트가 모두 비어 있으면 위 루프가 한 번도 돌지 않아 first_error가 None으로 남는다.
    # 어디까지 찾아봤는지 함께 알려야, 머리글이 스캔 범위보다 아래 있는 파일에서
    # "컬럼이 없다"는 메시지만 보고 원인을 못 찾는 일이 없다.
    detail = first_error or ValueError("읽을 수 있는 데이터가 없습니다.")
    raise ValueError(f"{detail} "
                     f"(시트 {len(sheets)}개의 상위 {_MAX_HEADER_SCAN}행까지 머리글을 찾아봤습니다.)")


# --------------------------------------------------------------------------- 정제

def _strptime_or_none(text, fmt):
    try:
        return datetime.strptime(text, fmt)
    except ValueError:
        return None


def _sane(parsed):
    """상식 밖 연도는 해석 실패로 본다 ('26.07.01'을 서기 26년으로 읽는 식의 사고 방지)"""
    if parsed is None or not 1900 <= parsed.year <= 2100:
        return None
    return parsed


def _parse_date(value, serial_min=_SERIAL_MIN):
    """엑셀 날짜 셀·문자열·YYYYMMDD/YYMMDD 숫자·엑셀 일련번호를 datetime으로 변환한다.

    모든 값을 문자열로 읽기 때문에 한 컬럼 안에서도 형식이 뒤섞여 있다.
    형식을 명시적으로 나열해 순서대로 시도하므로, 값마다 다른 형식이어도 해석되고
    어떤 입력이 어떻게 읽힐지가 코드에 그대로 드러난다.

    serial_min - 엑셀 일련번호로 인정할 하한. 생년월일은 고령 환자를 잃지 않도록
    내원일자보다 낮은 값을 쓴다.

    해석할 수 없으면 None.
    """
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    for fmt in _DATE_FORMATS:
        parsed = _strptime_or_none(text, fmt)
        if parsed is not None:
            return _sane(parsed)

    if text.isdigit():
        # YYYYMMDD / YYMMDD / YYYY. 자릿수로 구분해야 일련번호와 섞이지 않는다.
        if len(text) == 8:
            return _sane(_strptime_or_none(text, '%Y%m%d'))
        if len(text) == 6:
            # 한국식 생년월일 표기(900101). 두 자리 연도는 미래가 되지 않도록
            # 올해까지는 2000년대, 그보다 크면 1900년대로 읽는다.
            century = 2000 if int(text[:2]) <= datetime.now().year % 100 else 1900
            return _sane(_strptime_or_none(f"{century + int(text[:2]):04d}{text[2:]}", '%Y%m%d'))
        if len(text) == 4:
            return _sane(_strptime_or_none(text, '%Y'))

    # 엑셀 날짜 일련번호 (서식이 '일반'으로 저장된 날짜 셀)
    try:
        serial = float(text)
    except ValueError:
        return None
    if not math.isfinite(serial) or not serial_min <= serial <= _SERIAL_MAX:
        return None
    return _EXCEL_EPOCH + timedelta(days=serial)


def _clean_phone(value):
    """연락처에서 숫자만 남기고 국가번호·앞자리 0 누락을 보정한다"""
    text = str(value).strip() if value is not None else ''

    # 엑셀이 숫자로 저장한 셀은 '1012345678.0'처럼 소수점이 붙어 읽힐 수 있다.
    # 먼저 떼어내지 않으면 숫자만 남길 때 뒤에 0이 하나 더 붙어 멀쩡한 번호가 탈락한다.
    if text.endswith('.0'):
        text = text.rstrip('0').rstrip('.')

    digits = ''.join(c for c in text if c.isdigit())

    # +82 국제 표기: 821012345678 / 8201012345678 -> 국가번호 제거
    if digits.startswith('82') and 11 <= len(digits) <= 13:
        digits = digits[2:]

    # 엑셀이 앞자리 0을 지운 경우: 1012345678 -> 01012345678
    if len(digits) == 10 and digits.startswith('10'):
        digits = '0' + digits

    return digits


def _is_mobile(digits):
    return len(digits) == 11 and digits.startswith('010') and digits.isdigit()


# --------------------------------------------------------------------------- 필터

def _step(rows, keep, report, label):
    """조건을 적용하고 몇 건이 걸러졌는지 내역에 남긴다"""
    before = len(rows)
    rows = [row for row in rows if keep(row)]
    report.append(f"{label}: {len(rows)}건 (-{before - len(rows)})")
    return rows


def _resolve_column_dates(rows, key, serial_min, report, warnings):
    """날짜 컬럼을 해석하고, 실패한 행은 건수를 내역에 남기며 제외한다"""
    for row in rows:
        row[key] = _parse_date(row[key], serial_min=serial_min)

    unparsed = sum(1 for row in rows if row[key] is None)
    if not unparsed:
        return rows

    warnings.append(f"{key}를 해석하지 못한 {unparsed}건은 제외했습니다.")
    # 제외 건수도 내역에 남겨야 단계별 숫자가 앞뒤로 맞아떨어진다
    return _step(rows, lambda r: r[key] is not None, report, f"{key} 해석 가능")


def outbound_list_filter(file_path, password, period_type, period_type_old, period_value,
                         period_value_old, use_birth, start_date, end_date):
    """엑셀 파일을 복호화하고 내원일/생년월일 필터링 및 연락처 정제를 수행한다.

    반환: FilterResult(rows, report, warnings)
    """
    report = []
    warnings = []

    # 1. 복호화 및 로드 (전부 문자열로 읽어 원본 표기를 보존)
    source = decrypt_excel(file_path, password)

    required_keys = ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자']
    if use_birth:
        required_keys.append('생년월일')

    rows, notes = _load_sheet(source, required_keys)
    warnings.extend(notes)

    # 완전히 빈 행 제거 (엑셀 하단 여백)
    rows = [row for row in rows if any(value is not None for value in row.values())]
    report.append(f"원본 {len(rows)}건")

    # 2. 내원일 기준 필터
    visited_after = _threshold(period_type, period_value)           # 이 시점 이후에 내원한 대상만
    visited_before = _threshold(period_type_old, period_value_old)  # 이 시점 이전이 마지막인 대상만

    if visited_after is not None or visited_before is not None:
        rows = _resolve_column_dates(rows, '마지막 내원일자', _SERIAL_MIN, report, warnings)

        if visited_after is not None:
            rows = _step(rows, lambda r: r['마지막 내원일자'] >= visited_after, report,
                         f"최근 {period_value}{period_type} 이내 내원 "
                         f"({visited_after:%Y-%m-%d} 이후)")
        if visited_before is not None:
            rows = _step(rows, lambda r: r['마지막 내원일자'] <= visited_before, report,
                         f"{period_value_old}{period_type_old} 이상 미내원 "
                         f"({visited_before:%Y-%m-%d} 이전)")

    # 3. 생년월일 기준 필터
    if use_birth:
        rows = _resolve_column_dates(rows, '생년월일', _BIRTH_SERIAL_MIN, report, warnings)

        since = datetime.combine(start_date, datetime.min.time())
        until = datetime.combine(end_date, datetime.max.time())
        rows = _step(rows, lambda r: since <= r['생년월일'] <= until, report,
                     f"생년월일 {start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}")

    # 4. 휴대폰번호 정제 -> 형식 검증 -> 중복 제거
    #    (검증을 먼저 해야 유효한 번호끼리 비교된다)
    for row in rows:
        row['휴대폰번호'] = _clean_phone(row['휴대폰번호'])

    before = len(rows)
    rows = [row for row in rows if _is_mobile(row['휴대폰번호'])]
    if before != len(rows):
        report.append(f"휴대폰번호 형식 유효: {len(rows)}건 (-{before - len(rows)})")

    before = len(rows)
    rows = _unique_by(rows, '휴대폰번호')
    if before != len(rows):
        report.append(f"중복 연락처 제거: {len(rows)}건 (-{before - len(rows)})")

    # 5. 최종 데이터 (차트번호·이름은 원본 표기 그대로)
    result = [{key: row[key] for key in OUTPUT_COLUMNS} for row in rows]
    report.append(f"최종 {len(result)}건")

    return FilterResult(rows=result, report=report, warnings=warnings)


def _unique_by(rows, key):
    """먼저 등장한 행만 남긴다"""
    seen = set()
    unique = []
    for row in rows:
        if row[key] not in seen:
            seen.add(row[key])
            unique.append(row)
    return unique
