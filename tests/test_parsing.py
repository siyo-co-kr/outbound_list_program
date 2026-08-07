"""날짜·연락처 파싱 (_parse_date / _clean_phone / _dedupe_headers)."""

from datetime import datetime

import pytest

from app_modules.processors.outbound_list_filter_processor import (
    _BIRTH_SERIAL_MIN, _clean_phone, _dedupe_headers, _is_mobile, _parse_date)


@pytest.mark.parametrize("value,expected", [
    ('2026-07-01', '2026-07-01'),
    ('2026-07-01 00:00:00', '2026-07-01'),
    ('2026-07-01T00:00:00', '2026-07-01'),
    ('2026/07/01', '2026-07-01'),
    ('2026.07.01', '2026-07-01'),
    ('2026-7-1', '2026-07-01'),
    (' 2026-07-01 ', '2026-07-01'),
    ('20260701', '2026-07-01'),      # YYYYMMDD 숫자
    ('45000', '2023-03-15'),         # 엑셀 일련번호
    ('45000.0', '2023-03-15'),       # .xls가 float으로 준 일련번호
])
def test_parses_every_supported_date_shape(value, expected):
    assert _parse_date(value) == datetime.fromisoformat(expected)


@pytest.mark.parametrize("value", [
    '', ' ', None, '미상', '해당없음', 'N/A', 'nan', 'abc',
    '1990-13-45', '20261301', '20260732', '901301',
    '010-1234-5678', '01012345678',
])
def test_unparseable_values_return_none(value):
    assert _parse_date(value) is None


def test_mixed_formats_in_one_column_all_parse():
    """한 컬럼 안에 형식이 섞여 있어도 값마다 해석된다."""
    assert all(_parse_date(v) for v in ['2026-07-01', '20260702', '45000'])


@pytest.mark.parametrize("value,year", [('900101', 1990), ('691231', 1969), ('000101', 2000)])
def test_six_digit_korean_birth_notation(value, year):
    """두 자리 연도는 미래가 되지 않게 읽는다 (000101은 2000년생)."""
    assert _parse_date(value, serial_min=_BIRTH_SERIAL_MIN).year == year


def test_absurd_year_is_rejected():
    """'26.07.01'을 서기 26년으로 읽어 통과시키면 안 된다."""
    assert _parse_date('26.07.01') is None


def test_birth_serial_floor_keeps_elderly_patients():
    """일반 서식으로 저장된 1954년 이전 생년월일이 사라지면 안 된다."""
    born_1945 = '16603'
    assert _parse_date(born_1945) is None                                  # 내원일자 기준
    assert _parse_date(born_1945, serial_min=_BIRTH_SERIAL_MIN).year == 1945


def test_visit_serial_floor_rejects_small_numbers():
    """내원일자는 작은 숫자를 날짜로 오인하지 않아야 한다."""
    assert _parse_date('12951') is None


def test_fractional_serial_keeps_the_time():
    assert _parse_date('45000.5') == datetime(2023, 3, 15, 12, 0)


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
    assert _clean_phone(raw) == expected


def test_float_string_phone_does_not_gain_a_digit():
    """'.0'을 먼저 떼지 않으면 뒤에 0이 붙어 멀쩡한 번호가 탈락한다."""
    assert _clean_phone('1012345678.0') != '10123456780'


@pytest.mark.parametrize("digits,ok", [
    ('01012345678', True), ('0212345678', False), ('10123456780', False),
    ('0101234567', False), ('', False),
])
def test_is_mobile(digits, ok):
    assert _is_mobile(digits) is ok


def test_dedupe_headers_numbers_duplicates():
    assert _dedupe_headers(['연락처', '연락처', '연락처']) == ['연락처', '연락처.1', '연락처.2']


def test_dedupe_headers_strips_and_blanks_out_empty_cells():
    assert _dedupe_headers([' 차트번호 ', None, float('nan')]) == ['차트번호', '', '.1']
