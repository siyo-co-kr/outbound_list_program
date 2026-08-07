"""컬럼명 유연 매칭 (column_utils)."""

import pytest

from app_modules.processors.column_utils import (COLUMN_MAP, normalize_text,
                                                 resolve_columns)

ALL_KEYS = ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자', '생년월일']


def test_normalize_text_strips_spaces_and_symbols():
    assert normalize_text(" 마지막 내원일자 ") == "마지막내원일자"
    assert normalize_text("휴대폰-번호(2)") == "휴대폰번호2"


def test_resolves_common_header():
    mapping, notes = resolve_columns(
        ['차트번호', '이름', '연락처', '마지막내원일자'],
        ['차트번호', '환자 이름', '휴대폰번호', '마지막 내원일자'])
    assert mapping == {'차트번호': '차트번호', '환자 이름': '이름',
                       '휴대폰번호': '연락처', '마지막 내원일자': '마지막내원일자'}
    assert notes == []


def test_more_specific_candidate_wins_over_generic():
    """'휴대폰번호'와 '연락처'가 함께 있으면 구체적인 쪽을 쓴다."""
    mapping, notes = resolve_columns(['연락처', '휴대폰번호'], ['휴대폰번호'])
    assert mapping['휴대폰번호'] == '휴대폰번호'
    assert len(notes) == 1 and "'연락처'" in notes[0]


def test_result_is_independent_of_column_order():
    """파일의 컬럼 순서가 바뀌어도 같은 결과가 나와야 한다."""
    a, _ = resolve_columns(['연락처', '휴대폰번호'], ['휴대폰번호'])
    b, _ = resolve_columns(['휴대폰번호', '연락처'], ['휴대폰번호'])
    assert a == b


def test_missing_required_column_names_the_key():
    with pytest.raises(ValueError, match="생년월일"):
        resolve_columns(['차트번호', '이름', '연락처'], ['차트번호', '생년월일'])


def test_one_column_is_never_assigned_to_two_keys():
    columns = ['차트번호', '환자이름', '휴대폰번호', '마지막내원일자', '생년월일']
    mapping, _ = resolve_columns(columns, ALL_KEYS)
    assert len(set(mapping.values())) == len(mapping)


def test_candidate_lists_do_not_overlap_between_keys():
    """후보가 두 키에 겹치면 어느 항목이 컬럼을 가져갈지 순서에 좌우된다."""
    seen = {}
    for key, candidates in COLUMN_MAP.items():
        for candidate in candidates:
            norm = normalize_text(candidate)
            assert norm not in seen, f"'{candidate}' 후보가 {seen.get(norm)}와 {key}에 중복"
            seen[norm] = key
