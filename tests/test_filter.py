"""필터 전체 흐름 (outbound_list_filter)."""

import re
from datetime import date, timedelta

import pytest

from app_modules.processors.outbound_list_filter_processor import outbound_list_filter
from conftest import HEADER, HEADER_WITH_BIRTH

RECENT = (date.today() - timedelta(days=10)).strftime('%Y-%m-%d')
LONG_AGO = (date.today() - timedelta(days=365 * 5)).strftime('%Y-%m-%d')


def run(path, period_type='전체', period_value='6', period_type_old='전체',
        period_value_old='6', use_birth=False, start_date=None, end_date=None):
    return outbound_list_filter(
        file_path=path, password='',
        period_type=period_type, period_type_old=period_type_old,
        period_value=period_value, period_value_old=period_value_old,
        use_birth=use_birth, start_date=start_date, end_date=end_date)


def counts(report):
    """각 내역 줄에서 (건수, 제외건수)를 뽑는다."""
    out = []
    for line in report:
        m = re.search(r'(\d+)건(?:\s*\(-(\d+)\))?', line)
        out.append((line, int(m.group(1)), int(m.group(2)) if m.group(2) else None))
    return out


def assert_tally_is_consistent(report):
    """단계별 숫자가 앞뒤로 맞아떨어져야 한다 (내역이 검증 수단이므로)."""
    rows = counts(report)
    for (prev_line, prev_n, _), (line, n, delta) in zip(rows, rows[1:]):
        if delta is None:
            assert n == prev_n, f"'{line}'가 '{prev_line}'과 이어지지 않음"
        else:
            assert prev_n - delta == n, f"'{line}'의 숫자가 맞지 않음 (앞: {prev_n})"


def test_report_accounts_for_rows_dropped_as_unparseable(make_xlsx):
    rows = [HEADER]
    rows += [[f'A{i}', f'환자{i}', f'0101234000{i}', RECENT] for i in range(6)]
    rows += [[f'B{i}', f'환자B{i}', f'0101234100{i}', '미상'] for i in range(4)]

    result = run(make_xlsx(rows, sheets=('요약', '데이터')), period_type='개월')

    assert len(result.df) == 6
    assert_tally_is_consistent(result.report)
    assert any('해석하지 못한 4건' in w for w in result.warnings)


def test_keeps_original_notation_and_drops_invalid_contacts(make_xlsx):
    rows = [
        HEADER,
        ['00123', '김철수', '010-1111-2222', RECENT],
        ['00124', '이영희', '+82 10-3333-4444', RECENT],
        ['00125', '중복자', '010-1111-2222', RECENT],    # 중복 연락처
        ['00126', '탈락자', '02-123-4567', RECENT],      # 일반전화
    ]
    result = run(make_xlsx(rows), period_type='개월')

    assert list(result.df['차트번호']) == ['00123', '00124']
    assert list(result.df['휴대폰번호']) == ['01011112222', '01033334444']
    assert_tally_is_consistent(result.report)


def test_duplicate_contacts_keep_the_first_row(make_xlsx):
    rows = [HEADER,
            ['A1', '먼저', '01012345678', RECENT],
            ['A2', '나중', '010-1234-5678', RECENT]]
    result = run(make_xlsx(rows))
    assert list(result.df['환자 이름']) == ['먼저']


def test_recent_and_long_absent_filters(make_xlsx):
    rows = [HEADER,
            ['A1', '최근내원', '01011112222', RECENT],
            ['A2', '장기미내원', '01033334444', LONG_AGO]]

    recent_only = run(make_xlsx(rows), period_type='개월', period_value='6')
    assert list(recent_only.df['환자 이름']) == ['최근내원']

    long_absent = run(make_xlsx(rows), period_type_old='년', period_value_old='3')
    assert list(long_absent.df['환자 이름']) == ['장기미내원']


def test_birth_filter_keeps_patients_born_before_1954(make_xlsx):
    """생년월일이 일반 서식(일련번호)이어도 고령 환자가 사라지면 안 된다."""
    rows = [HEADER_WITH_BIRTH,
            ['A1', '고령환자', '01011112222', RECENT, '16603'],   # 1945-06-15
            ['A2', '젊은환자', '01033334444', RECENT, '1990-05-05']]

    result = run(make_xlsx(rows), use_birth=True,
                 start_date=date(1940, 1, 1), end_date=date(1950, 1, 1))

    assert list(result.df['환자 이름']) == ['고령환자']
    assert_tally_is_consistent(result.report)


def test_empty_result_still_returns_a_report(make_xlsx):
    rows = [HEADER, ['A1', '홍길동', '01012345678', LONG_AGO]]
    result = run(make_xlsx(rows), period_type='개월', period_value='1')

    assert result.df.empty
    assert result.report[0].startswith('원본')
    assert result.report[-1] == '최종 0건'


def test_output_columns_match_the_template(make_xlsx):
    from app_modules.processors.template_writer import TEMPLATE_COLUMNS

    rows = [HEADER, ['A1', '홍길동', '01012345678', RECENT]]
    result = run(make_xlsx(rows))
    assert list(result.df.columns) == TEMPLATE_COLUMNS


def test_missing_column_is_reported_by_name(make_xlsx):
    rows = [['차트번호', '이름', '연락처'], ['A1', '홍길동', '01012345678']]
    with pytest.raises(ValueError, match="마지막 내원일자"):
        run(make_xlsx(rows))
