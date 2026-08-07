import re

# 컬럼 매핑 사전
# key: 프로세서에서 사용하는 표준 컬럼명, value: 실제 파일에서 인식할 후보 컬럼명 목록
#
# 후보는 "구체적인 이름일수록 앞"에 둔다. 한 파일에 여러 후보가 동시에 존재할 때
# 이 순서대로 우선 선택하므로, 파일의 컬럼 순서와 무관하게 결과가 항상 같다.
COLUMN_MAP = {
    '차트번호'      : ['차트번호', '환자번호'],
    '환자 이름'     : ['환자이름', '환자명', '이름', '성명', '고객명', '환자', '고객'],
    '휴대폰번호'    : ['휴대폰번호', '핸드폰번호', '휴대폰', '핸드폰', '연락처', '전화번호', '전화'],
    '마지막 내원일자': ['마지막내원일자', '최종내원일', '내원일', '방문일자', '방문일', '예약일자', '예약일'],
    '생년월일'      : ['생년월일', '생일'],
}


def normalize_text(text):
    """문자열에서 모든 공백 및 특수문자를 제거하고 소문자로 변환"""
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', str(text)).lower()


def find_matching_columns(df_columns, target_key, exclude=()):
    """표준 키에 매칭되는 컬럼을 후보 우선순위 순서로 모두 반환"""
    # 컬럼명 정규화는 컬럼당 한 번만 한다. 후보 루프 안에서 다시 계산하면
    # 머리글 행마다 (후보 수 x 컬럼 수)만큼 정규식이 돌아간다.
    normalized = [(col, normalize_text(col)) for col in df_columns]

    matches = []
    seen = set()
    for candidate in COLUMN_MAP[target_key]:
        target = normalize_text(candidate)
        for col, col_norm in normalized:
            if col_norm == target and col not in exclude and col not in seen:
                seen.add(col)
                matches.append(col)
    return matches


def resolve_columns(df_columns, required_keys):
    """표준 키 -> 실제 컬럼명 매핑을 결정적으로 해석한다.

    - 후보 목록 순서로 매칭하므로 파일의 컬럼 순서에 결과가 좌우되지 않는다.
    - 이미 다른 항목에 배정된 컬럼은 재사용하지 않아, 서로 다른 항목이
      같은 컬럼을 가리켜 값이 뒤섞이는 일이 없다.

    반환: (mapping, notes)
      mapping - {표준 키: 실제 컬럼명}
      notes   - 후보가 둘 이상이라 하나를 골라야 했던 경우의 안내 문구
    반환값이 없으면(필수 컬럼 누락) ValueError를 발생시킨다.
    """
    mapping = {}
    notes = []
    taken = set()

    for key in required_keys:
        matches = find_matching_columns(df_columns, key, exclude=taken)
        if not matches:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {key}")

        chosen = matches[0]
        mapping[key] = chosen
        taken.add(chosen)

        if len(matches) > 1:
            ignored = ", ".join(f"'{c}'" for c in matches[1:])
            notes.append(
                f"'{key}' 후보가 여러 개입니다. '{chosen}' 컬럼을 사용하고 {ignored}은(는) 무시했습니다."
            )

    return mapping, notes
