"""날짜·연락처 파싱 (_parse_dates / _clean_phone / _dedupe_headers)."""

import pandas as pd
import pytest

from app_modules.processors.outbound_list_filter_processor import (
    _BIRTH_SERIAL_MIN, _clean_phone, _dedupe_headers, _parse_dates)


def parse_one(value, **kw):
    return _parse_dates(pd.Series([value]), **kw)[0]


@pytest.mark.parametrize("value,expected", [
    ('2026-07-01', '2026-07-01'),
    ('2026-07-01 00:00:00', '2026-07-01'),
    ('2026/07/01', '2026-07-01'),
    ('2026.07.01', '2026-07-01'),
    ('20260701', '2026-07-01'),      # YYYYMMDD 숫자
    ('45000', '2023-03-15'),         # 엑셀 일련번호
])
def test_parses_every_supported_date_shape(value, expected):
    assert parse_one(value) == pd.Timestamp(expected)


@pytest.mark.parametrize("value", ['', None, '미상', '1990-13-45', '해당없음'])
def test_unparseable_values_become_nat(value):
    assert pd.isna(parse_one(value))


def test_mixed_formats_in_one_column_all_parse():
    """pandas 자동 추론은 첫 값의 형식을 나머지에 적용해 버린다."""
    got = _parse_dates(pd.Series(['2026-07-01', '20260702', '45000']))
    assert got.notna().all()


def test_birth_serial_floor_keeps_elderly_patients():
    """일반 서식으로 저장된 1954년 이전 생년월일이 사라지면 안 된다."""
    born_1945 = '16603'
    assert pd.isna(parse_one(born_1945))                                  # 내원일자 기준
    assert parse_one(born_1945, serial_min=_BIRTH_SERIAL_MIN).year == 1945


def test_visit_serial_floor_rejects_small_numbers():
    """내원일자는 작은 숫자를 날짜로 오인하지 않아야 한다."""
    assert pd.isna(parse_one('12951'))


@pytest.mark.parametrize("raw,expected", [
    ('010-1234-5678', '01012345678'),
    ('01012345678', '01012345678'),
    ('1012345678', '01012345678'),        # 엑셀이 앞자리 0을 지운 경우
    ('+82 10-1234-5678', '01012345678'),  # 국제 표기
    ('+82 010-1234-5678', '01012345678'),
    ('821012345678', '01012345678'),
    ('1012345678.0', '01012345678'),      # .xls 숫자 셀이 float으로 읽힌 경우
    ('0212345678', '0212345678'),         # 일반전화는 보정하지 않음
    ('', ''),
    (None, ''),
])
def test_clean_phone(raw, expected):
    assert _clean_phone(pd.Series([raw]))[0] == expected


def test_float_string_phone_does_not_gain_a_digit():
    """'.0'을 먼저 떼지 않으면 뒤에 0이 붙어 멀쩡한 번호가 탈락한다."""
    assert _clean_phone(pd.Series(['1012345678.0']))[0] != '10123456780'


def test_dedupe_headers_numbers_duplicates():
    assert _dedupe_headers(['연락처', '연락처', '연락처']) == ['연락처', '연락처.1', '연락처.2']


def test_dedupe_headers_strips_and_blanks_out_empty_cells():
    assert _dedupe_headers([' 차트번호 ', None, float('nan')]) == ['차트번호', '', '.1']
