"""resources/outbound_template.xlsx 의 양식 그대로 결과 엑셀을 생성한다.

서식을 코드로 재현하지 않고 템플릿 파일 자체를 열어서 값만 채우므로,
글꼴·테두리·머리글 색·열 너비·작성예시 블록이 원본과 100% 동일하게 유지된다.
"""

import os
import sys
from copy import copy

import openpyxl

TEMPLATE_NAME = "outbound_template.xlsx"

# 템플릿 A1:C1 머리글과 동일한 순서
TEMPLATE_COLUMNS = ["차트번호", "환자 이름", "휴대폰번호"]

_HEADER_ROW = 1
_FIRST_DATA_ROW = 2
_EXAMPLE_BLOCK = ("F1:F3", 6, 9, 3)      # 병합범위, 시작열(F), 끝열(I), 끝행

# 제한 명단 시트의 열 너비 (템플릿 본문 열 너비와 같은 계열)
_LIMIT_COLUMN_WIDTHS = {
    "차트번호": 13.5,
    "이름": 13.5,
    "생년월일": 15.0,
    "연락처": 18.8,
    "마지막 내원일자": 18.0,
    "아웃바운드 제한 설정": 22.0,
}


def resource_path(*parts):
    """PyInstaller 번들 실행과 소스 실행 양쪽에서 통하는 리소스 절대 경로"""
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    return os.path.join(base, *parts)


def _load_template():
    path = resource_path("resources", TEMPLATE_NAME)
    if not os.path.exists(path):
        raise FileNotFoundError(f"양식 파일을 찾을 수 없습니다: {path}")
    return openpyxl.load_workbook(path)


def _to_text(value):
    """텍스트 서식(@) 셀에 넣을 문자열로 변환. 12345.0 -> '12345'"""
    if value is None or value != value:                  # None 또는 NaN
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.lower() == "nan":
        return None
    return text or None


def _copy_style(ws, src_row, src_col, dst_row, dst_col):
    ws.cell(row=dst_row, column=dst_col)._style = copy(
        ws.cell(row=src_row, column=src_col)._style
    )


def _ensure_styled_rows(ws, column_count, needed_rows):
    """데이터 행이 템플릿에 미리 서식이 잡힌 행 수를 넘으면 서식을 이어서 복제"""
    last_row = _FIRST_DATA_ROW + needed_rows - 1
    for row in range(ws.max_row + 1, last_row + 1):
        for col in range(1, column_count + 1):
            _copy_style(ws, _FIRST_DATA_ROW, col, row, col)


def _write_rows(ws, rows):
    for offset, record in enumerate(rows):
        row = _FIRST_DATA_ROW + offset
        for col, value in enumerate(record, start=1):
            ws.cell(row=row, column=col).value = _to_text(value)
    return len(rows)


def _clear_example_block(ws):
    """머리글이 3열을 넘는 시트를 만들 때, 템플릿 우측 '작성예시' 블록을 제거"""
    merged, first_col, last_col, last_row = _EXAMPLE_BLOCK
    if merged in [str(r) for r in ws.merged_cells.ranges]:
        ws.unmerge_cells(merged)

    for row in range(1, last_row + 1):
        for col in range(first_col, last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell._style = copy(ws.cell(row=1, column=4)._style)   # D1 = 서식 없는 셀

    for col in range(first_col, last_col + 1):
        letter = openpyxl.utils.get_column_letter(col)
        if letter in ws.column_dimensions:
            del ws.column_dimensions[letter]


def save_outbound_list(df, save_path):
    """아웃바운드 리스트 필터 결과를 템플릿 양식 그대로 저장한다."""
    missing = [c for c in TEMPLATE_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"양식에 필요한 컬럼이 없습니다: {', '.join(missing)}")

    wb = _load_template()
    ws = wb.worksheets[0]

    rows = list(df[TEMPLATE_COLUMNS].itertuples(index=False, name=None))
    _ensure_styled_rows(ws, len(TEMPLATE_COLUMNS), len(rows))
    count = _write_rows(ws, rows)

    wb.save(save_path)
    return count


def save_limit_list(df, save_path):
    """제한 명단 결과를 템플릿과 같은 서식(글꼴·테두리·머리글 색)으로 저장한다.

    컬럼 구성이 템플릿(3열)과 달라 우측 '작성예시' 블록은 제거하고
    머리글/본문 셀 서식만 템플릿에서 그대로 가져와 사용한다.
    """
    columns = list(df.columns)

    wb = _load_template()
    ws = wb.worksheets[0]
    _clear_example_block(ws)

    # 머리글: 템플릿 A1 서식을 그대로 복제해 컬럼 수만큼 확장
    for col, name in enumerate(columns, start=1):
        _copy_style(ws, _HEADER_ROW, 1, _HEADER_ROW, col)
        ws.cell(row=_HEADER_ROW, column=col).value = name

    rows = list(df.itertuples(index=False, name=None))

    # 본문: 템플릿 A2 서식을 데이터 영역 전체(가로·세로)로 복제
    template_body_rows = ws.max_row - _FIRST_DATA_ROW + 1
    for offset in range(max(len(rows), template_body_rows)):
        for col in range(1, len(columns) + 1):
            _copy_style(ws, _FIRST_DATA_ROW, 1, _FIRST_DATA_ROW + offset, col)

    count = _write_rows(ws, rows)

    for col, name in enumerate(columns, start=1):
        letter = openpyxl.utils.get_column_letter(col)
        ws.column_dimensions[letter].width = _LIMIT_COLUMN_WIDTHS.get(name, 15.0)

    wb.save(save_path)
    return count
