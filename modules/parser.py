import pandas as pd
import re
from datetime import datetime
from modules.classifier import classify

CURRENT_YEAR = datetime.now().year


def _str(val) -> str:
    if val is None or (isinstance(val, float) and val != val):
        return ''
    s = str(val).strip()
    return '' if s.lower() == 'nan' else s


def _normalize_bid_no(val) -> str:
    return _str(val).upper()


def _normalize_date(val) -> str:
    if not val or (isinstance(val, float) and val != val):
        return ''
    s = str(val).strip()
    if s.lower() == 'nan':
        return ''
    if re.match(r'\d{4}-\d{2}-\d{2}', s):
        return s[:16]
    m = re.match(r'^(\d{2})-(\d{2})\s+(\d{2}:\d{2})', s)
    if m:
        return f'{CURRENT_YEAR}-{m.group(1)}-{m.group(2)} {m.group(3)}'
    return s


def _to_int(val) -> int:
    try:
        v = str(val).replace(',', '').strip()
        return int(float(v)) if v and v.lower() != 'nan' else 0
    except Exception:
        return 0


def _parse_kind(raw: str, override: str = '') -> str:
    if override and override.strip():
        return override.strip()
    if not raw:
        return '본공고'
    if re.match(r'^R\d{2}BD', raw, re.I):
        return '사전공고'
    m = re.match(r'^.+-(\d{3})$', raw)
    if m:
        return '본공고' if m.group(1) == '000' else '재공고'
    return '본공고'


def _parse_bid_no(raw: str) -> str:
    """공고번호 정규화 (-000 그대로 유지, 공백 제거)"""
    return _str(raw).upper()


def parse_bid_notice(path: str) -> list[dict]:
    """입찰공고_YYYYMMDD.xlsx 파싱"""
    df = pd.read_excel(path)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    records = []
    for _, row in df.iterrows():
        bid_no = _parse_bid_no(row.get('공고번호', ''))
        if not bid_no:
            continue
        records.append({
            '영업속성':                    '비드메이트',
            '문의 / 공고일':               _normalize_date(row.get('입력일', '')),
            '사전/본공고':                 _parse_kind(bid_no),
            '공고번호':                    bid_no,
            '사업명':                      _str(row.get('공고명', '')),
            'RFP':                         '',
            '수요처':                      _str(row.get('발주기관', '')),
            '사업종류':                    classify(_str(row.get('공고명', ''))),
            '사업예산 (VAT포함)':          _to_int(row.get('추정가격', 0)),
            '입찰 마감일':                 _normalize_date(row.get('참가마감', '')),
            '사업 예상 시기':              '',
            '대응여부':                    '',
            '사업 제안 지원부서 / 담당자': '',
            '비고':                        '',
            '낙찰사':                      '',
            '낙찰율':                      '',
            '참여업체수':                  '',
            '내순위':                      '',
            '개찰일시':                    _normalize_date(row.get('개찰일시', '')),
            '단계':                        '',
            '원본출처':                    '비드메이트_입찰공고',
            '최종수정일':                  now,
        })
    return records


def parse_bid_result(path: str) -> list[dict]:
    """낙찰정보_YYYYMMDD.xlsx 파싱 — 업데이트용"""
    df = pd.read_excel(path)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    records = []
    for _, row in df.iterrows():
        bid_no = _parse_bid_no(row.get('공고번호', ''))
        if not bid_no:
            continue
        records.append({
            '공고번호':    bid_no,
            '사업명':      _str(row.get('공고명', '')),
            '낙찰사':      _str(row.get('1순위 업체', '')),
            '낙찰율':      _str(row.get('1순위 사정율', '')),
            '참여업체수':  str(_to_int(row.get('참여업체수', 0))),
            '내순위':      _str(row.get('내순위', '')),
            '개찰일시':    _normalize_date(row.get('개찰일시', '')),
            '단계':        '개찰완료',
            '원본출처':    '비드메이트_낙찰정보',
            '최종수정일':  now,
        })
    return records


def parse_sales(path: str, sheet_index: int = 1) -> list[dict]:
    """영업현황 시트2 파싱"""
    df = pd.read_excel(path, sheet_name=sheet_index)
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    records = []
    for _, row in df.iterrows():
        bid_no = _parse_bid_no(row.get('공고번호', ''))
        kind_raw = _str(row.get('사전/본공고', ''))
        records.append({
            '영업속성':                    _str(row.get('영업속성', '조달청')) or '조달청',
            '문의 / 공고일':               _normalize_date(row.get('문의 / 공고일', '')),
            '사전/본공고':                 _parse_kind(bid_no, kind_raw),
            '공고번호':                    bid_no,
            '사업명':                      _str(row.get('사업명', '')),
            'RFP':                         _str(row.get('RFP', '')),
            '수요처':                      _str(row.get('수요처', '')),
            '사업종류':                    classify(_str(row.get('사업종류', '')) or _str(row.get('사업명', ''))),
            '사업예산 (VAT포함)':          _to_int(row.get('사업예산 (VAT포함)', 0)),
            '입찰 마감일':                 _normalize_date(row.get('입찰 마감일', '')),
            '사업 예상 시기':              _normalize_date(row.get('사업 예상 시기', '')),
            '대응여부':                    _str(row.get('대응여부', '')),
            '사업 제안 지원부서 / 담당자': _str(row.get(' 사업 제안 지원부서 / 담당자', '')),
            '비고':                        _str(row.get('비고', '')),
            '낙찰사':                      '',
            '낙찰율':                      '',
            '참여업체수':                  '',
            '내순위':                      '',
            '개찰일시':                    '',
            '단계':                        '',
            '원본출처':                    '영업현황',
            '최종수정일':                  now,
        })
    return records
