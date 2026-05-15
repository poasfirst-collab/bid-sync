from rapidfuzz import fuzz
from config import RESULT_UPDATE_COLS


def normalize_key(bid_no: str) -> str:
    """-000 suffix 제거 후 대문자 정규화"""
    s = str(bid_no).strip().upper()
    return s[:-4] if s.endswith('-000') else s


def find_match(record: dict, pool: list[dict], threshold: int = 80) -> int:
    """
    pool에서 record 매칭 인덱스 반환. 없으면 -1.
    1순위: 공고번호 정규화 일치
    2순위: 사업명 유사도 threshold% 이상
    """
    key  = normalize_key(record.get('공고번호', ''))
    name = str(record.get('사업명', '')).strip()

    if key:
        for i, r in enumerate(pool):
            if normalize_key(r.get('공고번호', '')) == key:
                return i

    if name:
        for i, r in enumerate(pool):
            existing = str(r.get('사업명', '')).strip()
            if existing and fuzz.ratio(name, existing) >= threshold:
                return i

    return -1


def merge_into(target: dict, source: dict, cols: list) -> bool:
    """source의 cols를 target에 덮어씀. 변경 발생 시 True."""
    changed = False
    for col in cols:
        val = source.get(col, '')
        if val is not None and str(val).strip() not in ('', '0', 'nan'):
            target[col] = val
            changed = True
    return changed


def merge_sales_into(target: dict, source: dict) -> None:
    """영업현황 → 기존 레코드에 빈 필드만 채움 (덮어쓰지 않음)"""
    for col, val in source.items():
        if col.startswith('_'):
            continue
        if not target.get(col) and val and str(val).strip() not in ('', '0', 'nan'):
            target[col] = val


def sort_by_date(records: list[dict]) -> list[dict]:
    """문의 / 공고일 기준 오름차순 정렬. 빈 값은 맨 뒤."""
    return sorted(records, key=lambda r: r.get('문의 / 공고일', '') or '9999-99-99')
