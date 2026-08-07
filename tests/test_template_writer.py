"""템플릿 양식 저장 (template_writer)."""

import pandas as pd
import pytest

from app_modules.processors.template_writer import (TEMPLATE_COLUMNS, _to_text,
                                                    save_outbound_list)


@pytest.mark.parametrize("value,expected", [
    ('00123', '00123'),        # 앞자리 0 유지
    ('123', '123'),
    (123.0, '123'),
    ('123.0', '123'),          # .xls 숫자 셀이 문자열로 읽힌 경우
    ('1012345678.0', '1012345678'),
    ('3.05', '3.05'),          # 정수+.0 형태가 아니면 건드리지 않음
    (None, None),
    (float('nan'), None),
    ('nan', None),
    ('   ', None),
])
def test_to_text(value, expected):
    assert _to_text(value) == expected


def test_save_round_trips_values_without_losing_leading_zeros(tmp_path):
    df = pd.DataFrame({
        '차트번호': ['00123', '00124'],
        '환자 이름': ['홍길동', '김철수'],
        '휴대폰번호': ['01011112222', '01033334444'],
    })
    out = tmp_path / "result.xlsx"

    assert save_outbound_list(df, str(out)) == 2

    back = pd.read_excel(out, dtype=str)
    assert list(back.columns[:3]) == TEMPLATE_COLUMNS
    assert list(back['차트번호'])[:2] == ['00123', '00124']
    assert list(back['휴대폰번호'])[:2] == ['01011112222', '01033334444']


def test_save_rejects_a_frame_missing_template_columns(tmp_path):
    df = pd.DataFrame({'차트번호': ['1'], '환자 이름': ['홍길동']})
    with pytest.raises(ValueError, match="휴대폰번호"):
        save_outbound_list(df, str(tmp_path / "x.xlsx"))
