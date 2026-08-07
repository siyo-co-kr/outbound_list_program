"""테스트 공용 설정 및 픽스처."""

import sys
from pathlib import Path

import openpyxl
import pytest

# 프로젝트 루트를 임포트 경로에 추가 (설치 없이 소스 그대로 테스트)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# 원본 파일에서 흔히 쓰이는 머리글
HEADER = ['차트번호', '이름', '연락처', '마지막 내원일자']
HEADER_WITH_BIRTH = HEADER + ['생년월일']


@pytest.fixture
def make_xlsx(tmp_path):
    """시트/머리글 위치를 지정해 엑셀 파일을 만들어 경로를 돌려주는 팩토리.

    sheets      - 시트 이름 목록. 데이터는 마지막 시트에 쓴다.
    skip_rows   - 머리글 앞에 끼워 넣을 안내문 행 수.
    """
    counter = {'n': 0}

    def _make(rows, sheets=('데이터',), skip_rows=0):
        wb = openpyxl.Workbook()
        wb.active.title = sheets[0]
        for extra in sheets[1:]:
            wb.create_sheet(extra)

        ws = wb[sheets[-1]]
        for _ in range(skip_rows):
            ws.append(["안내문"])
        for row in rows:
            ws.append(row)

        counter['n'] += 1
        path = tmp_path / f"book{counter['n']}.xlsx"
        wb.save(path)
        return str(path)

    return _make
