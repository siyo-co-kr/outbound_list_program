"""구형 .xls 읽기 경로 (xlrd).

.xls는 모든 숫자를 float으로 돌려주므로, 연락처에 소수점이 붙어 번호가 통째로
탈락하는 사고가 나기 쉽다. 형식 판별도 확장자가 아닌 내용으로 해야 한다.
"""

from datetime import date, timedelta

import pytest

xlwt = pytest.importorskip("xlwt", reason=".xls 픽스처 생성에 xlwt 필요")

from app_modules.processors.outbound_list_filter_processor import (  # noqa: E402
    _read_grid, outbound_list_filter)

HEADER = ['차트번호', '이름', '연락처', '마지막 내원일자']
RECENT = (date.today() - timedelta(days=10)).strftime('%Y-%m-%d')


@pytest.fixture
def make_xls(tmp_path):
    counter = {'n': 0}

    def _make(rows, sheets=('데이터',)):
        book = xlwt.Workbook()
        written = [book.add_sheet(name) for name in sheets]
        sheet = written[-1]
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                sheet.write(r, c, value)
        counter['n'] += 1
        path = tmp_path / f"book{counter['n']}.xls"
        book.save(str(path))
        return str(path)

    return _make


def test_detects_xls_by_content_not_extension(make_xls, tmp_path):
    """확장자가 틀려도 내용으로 .xls를 알아본다."""
    src = make_xls([HEADER, ['A1', '홍길동', '01012345678', RECENT]])
    renamed = tmp_path / "mislabelled.xlsx"
    renamed.write_bytes(open(src, 'rb').read())

    sheets = _read_grid(str(renamed))
    assert sheets[0][1][0] == HEADER


def test_numeric_cells_do_not_gain_a_decimal_point(make_xls):
    """xlrd는 숫자를 float으로 준다. '1012345678.0'이 되면 번호가 망가진다."""
    path = make_xls([HEADER, ['A1', '홍길동', 1012345678, 20260701]])
    grid = _read_grid(path)[0][1]

    assert grid[1][2] == '1012345678'
    assert grid[1][3] == '20260701'


def test_end_to_end_recovers_number_stripped_contacts(make_xls):
    """연락처가 숫자 셀이면 앞자리 0이 지워진 채 들어온다. 보정되어야 한다."""
    rows = [HEADER,
            ['00123', '김철수', 1012345678, RECENT],       # 숫자 셀
            ['00124', '이영희', '010-3333-4444', RECENT]]  # 문자열 셀
    result = outbound_list_filter(make_xls(rows), '', '개월', '전체', '6', '6',
                                  False, None, None)

    assert [r['휴대폰번호'] for r in result.rows] == ['01012345678', '01033334444']
    assert [r['차트번호'] for r in result.rows] == ['00123', '00124']


def test_real_date_cells_are_read(make_xls):
    """서식이 날짜인 셀은 xlrd가 별도 타입으로 돌려준다."""
    style = xlwt.XFStyle()
    style.num_format_str = 'YYYY-MM-DD'

    book = xlwt.Workbook()
    sheet = book.add_sheet('데이터')
    for c, name in enumerate(HEADER):
        sheet.write(0, c, name)
    sheet.write(1, 0, 'A1')
    sheet.write(1, 1, '홍길동')
    sheet.write(1, 2, '01012345678')
    sheet.write(1, 3, date.today() - timedelta(days=10), style)

    import tempfile, os
    path = os.path.join(tempfile.mkdtemp(), "dated.xls")
    book.save(path)

    result = outbound_list_filter(path, '', '개월', '전체', '6', '6', False, None, None)
    assert len(result.rows) == 1


def test_second_sheet_preference_applies_to_xls(make_xls):
    path = make_xls([HEADER, ['A1', '홍길동', '01012345678', RECENT]],
                    sheets=('요약', '데이터'))
    result = outbound_list_filter(path, '', '전체', '전체', '6', '6', False, None, None)
    assert len(result.rows) == 1
    assert result.warnings == []
