import re
import time
import requests
from datetime import datetime

BID_BASE     = 'http://apis.data.go.kr/1230000/ad/BidPublicInfoService'
STD_BASE     = 'http://apis.data.go.kr/1230000/ao/PubDataOpnStdService'
PRESPEC_BASE = 'http://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService'

BID_OPS = [
    'getBidPblancListInfoServc',
    'getBidPblancListInfoThng',
    'getBidPblancListInfoCnstwk',
    'getBidPblancListInfoFrgcpt',
    'getBidPblancListInfoEtc',
]
PRESPEC_OPS = [
    'getPublicPrcureThngInfoServc',
    'getPublicPrcureThngInfoThng',
    'getPublicPrcureThngInfoCnstwk',
    'getPublicPrcureThngInfoFrgcpt',
]


def _call(url: str, params: dict) -> list:
    for attempt in range(3):
        try:
            res = requests.get(url, params=params, timeout=15)
            body = res.json()
            if res.status_code in (401, 403):
                raise PermissionError('API 인증 실패 — data.go.kr 서비스 키 확인 필요')
            raw = body.get('response', {}).get('body', {}).get('items') or []
            items = raw.get('item', raw) if isinstance(raw, dict) else raw
            return items if isinstance(items, list) else [items] if items else []
        except PermissionError:
            raise
        except Exception:
            if attempt == 2:
                return []
            time.sleep(1)
    return []


def _build_date_range(bid_no: str) -> dict:
    year = datetime.now().year
    month = None
    if re.match(r'^R\d{2}[A-Z]', bid_no, re.I):
        year = 2000 + int(bid_no[1:3])
    elif re.match(r'^\d{13}', bid_no):
        year = int(bid_no[:4])
        month = int(bid_no[4:6])

    now = datetime.now()
    m = month or now.month
    import calendar
    last_day = calendar.monthrange(year, m)[1]
    pad = lambda n: str(n).zfill(2)
    return {
        'bidNtceBgnDt': f'{year}{pad(m)}010000',
        'bidNtceEndDt': f'{year}{pad(m)}{last_day}2359',
    }


def _parse_item(it: dict, source: str) -> dict:
    attachments = []
    for i in range(1, 11):
        url = it.get(f'ntceSpecDocUrl{i}')
        name = it.get(f'ntceSpecFileNm{i}', f'attachment_{i}')
        if url:
            attachments.append({'url': url, 'name': name})

    budget_raw = it.get('bdgtAmt') or it.get('asignBdgtAmt') or it.get('presmptPrce') or 0
    try:
        budget = int(float(budget_raw))
    except Exception:
        budget = 0

    return {
        '공고번호':   (it.get('bidNtceNo') or '').strip(),
        '공고구분순번': it.get('bidNtceOrd', '000'),
        '공고명':    it.get('bidNtceNm', ''),
        '발주기관':  it.get('dminsttNm') or it.get('dmndInsttNm', ''),
        '공고기관':  it.get('ntceInsttNm', ''),
        '예산':      budget,
        '공고일':    it.get('bidNtceDt') or it.get('bidNtceDate', ''),
        '참가마감':  it.get('bidClseDt') or it.get('bidClseDate', ''),
        '개찰일시':  it.get('opengDt') or it.get('opengDate', ''),
        '계약방법':  it.get('cntrctCnclsMthdNm', ''),
        '낙찰방법':  it.get('bidwinMthdNm', ''),
        '담당자명':  it.get('ntceInsttOfclNm', ''),
        '담당자연락처': it.get('ntceInsttOfclTelNo') or it.get('ntceInsttOfclTel', ''),
        '나라장터URL': it.get('bidNtceDtlUrl', ''),
        'attachments': attachments,
        '원본출처':  '나라장터_API',
    }


def fetch_bid(api_key: str, bid_no: str) -> dict | None:
    """입찰공고 공고번호로 조회. R26BK... 또는 R26BK...-001 형식"""
    bid_no = bid_no.strip().upper()
    m = re.match(r'^(.+)-(\d{3})$', bid_no)
    base_no = m.group(1) if m else bid_no
    seq     = m.group(2) if m else '000'

    params = {
        'ServiceKey': api_key,
        'numOfRows': 10,
        'pageNo': 1,
        'type': 'json',
        'inqryDiv': 2,
        'bidNtceNo': base_no,
        **({'bidNtceOrd': seq} if seq != '000' else {}),
    }

    for op in BID_OPS:
        items = _call(f'{BID_BASE}/{op}', params)
        found = next((
            it for it in items
            if (it.get('bidNtceNo') or '').strip().upper() == base_no
            and (seq == '000' or (it.get('bidNtceOrd') or '000').strip() == seq)
        ), None)
        if found:
            return _parse_item(found, op)

    # 보조: 날짜 범위 검색
    dr = _build_date_range(base_no)
    std_params = {
        'ServiceKey': api_key,
        'numOfRows': 100,
        'pageNo': 1,
        'type': 'json',
        **dr,
    }
    items = _call(f'{STD_BASE}/getDataSetOpnStdBidPblancInfo', std_params)
    found = next((
        it for it in items
        if (it.get('bidNtceNo') or '').strip().upper() == base_no
    ), None)
    return _parse_item(found, 'PubDataOpnStdService') if found else None


def fetch_prespec(api_key: str, spec_no: str) -> dict | None:
    """사전규격 번호(R26BD...) 조회"""
    spec_no = spec_no.strip().upper()
    params = {
        'ServiceKey': api_key,
        'numOfRows': 10,
        'pageNo': 1,
        'type': 'json',
        'inqryDiv': 2,
        'bfSpecRgstNo': spec_no,
    }
    for op in PRESPEC_OPS:
        items = _call(f'{PRESPEC_BASE}/{op}', params)
        found = next((
            it for it in items
            if (it.get('bfSpecRgstNo') or '').strip().upper() == spec_no
        ), None)
        if found:
            budget_raw = found.get('asignBdgtAmt') or found.get('bdgtAmt') or 0
            try:
                budget = int(float(budget_raw))
            except Exception:
                budget = 0
            return {
                '공고번호':  found.get('bfSpecRgstNo', ''),
                '공고명':   found.get('prdctClsfcNoNm', ''),
                '발주기관': found.get('orderInsttNm', ''),
                '예산':     budget,
                '공고일':   found.get('rcptDt', ''),
                '참가마감': found.get('opninRgstClseDt', ''),
                '담당자명': found.get('ofclNm', ''),
                '담당자연락처': found.get('ofclTelNo', ''),
                '나라장터URL': f'https://www.g2b.go.kr/pn/pnz/pnza/spBfSpec/retrieveBfSpecDtlInfo.do?bfSpecRgstNo={spec_no}',
                '원본출처': '나라장터_사전규격',
            }
    return None
