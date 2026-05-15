"""
나라장터 공고번호 직접 검색 -> 구글 시트 추가
usage: python g2b_search.py
"""
import sys, os, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import G2B_API_KEY, SPREADSHEET_ID, COLUMNS, SERVICE_ACCOUNT_FILE
from modules.g2b_api import fetch_bid, fetch_prespec
from modules.classifier import classify
from modules.dedup import find_match
from modules.sheets import read_main_sheet
from modules.logger import log
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]
DIVIDER = '-' * 60


def is_prespec(bid_no):
    return bool(re.match(r'^R\d{2}BD', bid_no, re.I))


def display_result(data):
    print()
    print(DIVIDER)
    print(f'  공고번호  : {data.get("공고번호", "")}')
    print(f'  공고명    : {data.get("공고명", "")}')
    print(f'  발주기관  : {data.get("발주기관", "") or data.get("공고기관", "")}')
    budget = data.get('예산', 0)
    if budget:
        print(f'  예산      : {int(budget):,}원')
    print(f'  마감일    : {data.get("참가마감", "")}')
    if data.get('나라장터URL'):
        print(f'  나라장터  : {data["나라장터URL"]}')
    print(DIVIDER)


def to_sheet_record(data):
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    bid_no = str(data.get('공고번호', '')).strip().upper()
    seq = str(data.get('공고구분순번', '000')).strip()
    if seq and seq != '000':
        bid_no_full = f'{bid_no}-{seq}'
        kind = f'재공고({int(seq)}차)'
    else:
        bid_no_full = bid_no
        kind = '본공고'
    if re.match(r'^R\d{2}BD', bid_no, re.I):
        kind = '사전공고'

    return {
        '영업속성':                    '조달청',
        '문의 / 공고일':               now[:10],
        '사전/본공고':                 kind,
        '공고번호':                    bid_no_full,
        '사업명':                      data.get('공고명', ''),
        'RFP':                         '',
        '수요처':                      data.get('발주기관') or data.get('공고기관', ''),
        '사업종류':                    classify(data.get('공고명', '')),
        '사업예산 (VAT포함)':          str(data.get('예산', '')),
        '입찰 마감일':                 data.get('참가마감', ''),
        '사업 예상 시기':              '',
        '대응여부':                    '',
        '사업 제안 지원부서 / 담당자': '',
        '비고':                        '',
        '낙찰사':                      '',
        '낙찰율':                      '',
        '참여업체수':                  '',
        '내순위':                      '',
        '개찰일시':                    data.get('개찰일시') or data.get('개찰일', ''),
        '단계':                        '',
        '원본출처':                    '나라장터_API',
        '최종수정일':                  now,
    }


def run():
    print('=' * 60)
    print('  나라장터 공고번호 직접 검색')
    print('  종료: q')
    print('=' * 60)

    if not SPREADSHEET_ID:
        print('\n[오류] config.py의 SPREADSHEET_ID가 비어 있습니다.')
        return

    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)

    while True:
        print()
        bid_no = input('  공고번호 입력 > ').strip().upper()
        if bid_no in ('Q', ''):
            print('  종료합니다.')
            break

        print('  [검색 중...]', end='', flush=True)

        try:
            if is_prespec(bid_no):
                data = fetch_prespec(G2B_API_KEY, bid_no)
                label = '사전규격'
            else:
                data = fetch_bid(G2B_API_KEY, bid_no)
                label = '입찰공고'
        except PermissionError as e:
            print(f'\n[오류] {e}')
            log.error(str(e))
            break
        except Exception as e:
            print(f'\n[오류] API: {e}')
            log.error(f'API 오류: {e}')
            continue

        if not data:
            print(f'\n  찾을 수 없습니다: {bid_no}')
            log.info(f'미발견: {bid_no}')
            continue

        print(f' 발견 ({label})')
        display_result(data)

        ans = input('  구글 시트에 추가하시겠습니까? [Y/N] > ').strip().upper()
        if ans != 'Y':
            print('  추가하지 않았습니다.')
            continue

        sheet_rows = read_main_sheet()
        # 헤더 제외, 공고번호=인덱스3, 사업명=인덱스4
        pool = [
            {'공고번호': r[3], '사업명': r[4]}
            for r in sheet_rows[1:] if len(r) > 4
        ]
        check = {'공고번호': bid_no, '사업명': data.get('공고명', '')}
        if find_match(check, pool, threshold=80) >= 0:
            print('  [중복] 결과1/2에 이미 존재합니다. 추가하지 않습니다.')
            log.info(f'중복 스킵: {bid_no}')
            continue

        rec = to_sheet_record(data)
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet('통합공고관리')
        ws.append_row(
            [str(rec.get(col, '')) for col in COLUMNS],
            value_input_option='RAW'
        )
        print('  구글 시트에 추가 완료 (1건)')
        log.info(f'추가: {bid_no} - {data.get("공고명", "")}')


if __name__ == '__main__':
    run()
