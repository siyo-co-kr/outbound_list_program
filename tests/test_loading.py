"""시트·머리글 자동 탐색 (_load_sheet)."""

import openpyxl
import pytest

from app_modules.processors.outbound_list_filter_processor import _load_sheet
from conftest import HEADER

KEYS = ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자']
ROW = ['00123', '홍길동', '01012345678', '2026-07-01']


def test_prefers_second_sheet_when_several_exist(make_xlsx):
    path = make_xlsx([HEADER, ROW], sheets=('요약', '데이터'))
    df, mapping, notes = _load_sheet(path, KEYS)
    assert len(df) == 1
    assert notes == []                     # 예상대로 읽었으면 알릴 것이 없다


def test_uses_only_sheet_when_single(make_xlsx):
    df, _, notes = _load_sheet(make_xlsx([HEADER, ROW]), KEYS)
    assert len(df) == 1 and notes == []


def test_finds_header_below_the_first_row(make_xlsx):
    path = make_xlsx([HEADER, ROW], skip_rows=3)
    df, _, notes = _load_sheet(path, KEYS)
    assert len(df) == 1
    assert any("4행" in n for n in notes)


def test_falls_back_to_another_sheet_and_says_so(make_xlsx):
    """엉뚱한 시트를 읽었을 때 사용자가 알아챌 수 있어야 한다."""
    path = make_xlsx([HEADER, ROW], sheets=('표지', '메모', '데이터'))
    df, _, notes = _load_sheet(path, KEYS)
    assert len(df) == 1
    assert any("'메모'" in n and "'데이터'" in n for n in notes)


def test_header_deeper_than_scan_window_explains_where_it_looked(make_xlsx):
    path = make_xlsx([HEADER, ROW], skip_rows=9)
    with pytest.raises(ValueError, match="8행"):
        _load_sheet(path, KEYS)


def test_empty_workbook_raises_a_readable_error(tmp_path):
    """빈 시트만 있으면 내부 루프가 한 번도 돌지 않는다 (raise None 회귀)."""
    wb = openpyxl.Workbook()
    wb.active.title = "s1"
    wb.create_sheet("s2")
    path = tmp_path / "empty.xlsx"
    wb.save(path)

    with pytest.raises(ValueError, match="읽을 수 있는 데이터가 없습니다"):
        _load_sheet(str(path), KEYS)


def test_missing_column_names_the_key(make_xlsx):
    path = make_xlsx([['차트번호', '이름'], ['00123', '홍길동']])
    with pytest.raises(ValueError, match="휴대폰번호"):
        _load_sheet(path, KEYS)


def test_values_are_read_as_text(make_xlsx):
    """앞자리 0이 사라지거나 소수점이 붙으면 안 된다."""
    df, mapping, _ = _load_sheet(make_xlsx([HEADER, ROW]), KEYS)
    assert df[mapping['차트번호']][0] == '00123'
    assert df[mapping['휴대폰번호']][0] == '01012345678'
