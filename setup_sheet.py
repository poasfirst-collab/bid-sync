"""
구글 스프레드시트 초기 구조 설정 스크립트
- 기존 시트에 헤더·탭 구성
usage: python setup_sheet.py
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gspread
from google.oauth2.service_account import Credentials
from config import SERVICE_ACCOUNT_FILE, SPREADSHEET_ID, COLUMNS
from config import SHEET_MAIN, SHEET_LOG, SHEET_RULE
from modules.classifier import RULES

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

def main():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SPREADSHEET_ID)

    existing = [ws.title for ws in ss.worksheets()]

    # ── 메인 시트 ──────────────────────────────────────
    if SHEET_MAIN not in existing:
        ws_main = ss.add_worksheet(title=SHEET_MAIN, rows=5000, cols=len(COLUMNS))
    else:
        ws_main = ss.worksheet(SHEET_MAIN)

    # 헤더가 없을 때만 입력
    first = ws_main.row_values(1)
    if not first or first[0] != '공고번호':
        ws_main.update([COLUMNS], 'A1')
        # 헤더 행 굵게 (배경색 설정)
        ws_main.format('A1:U1', {
            'textFormat': {'bold': True},
            'backgroundColor': {'red': 0.2, 'green': 0.4, 'blue': 0.7},
        })
        time.sleep(1)
    print(f'  OK [{SHEET_MAIN}] 헤더 설정 완료')

    # 기본 Sheet1 이름이 남아 있으면 제거
    if 'Sheet1' in existing and len(ss.worksheets()) > 1:
        ss.del_worksheet(ss.worksheet('Sheet1'))
        time.sleep(0.5)

    # ── 로그 시트 ──────────────────────────────────────
    if SHEET_LOG not in existing:
        ws_log = ss.add_worksheet(title=SHEET_LOG, rows=1000, cols=10)
        ws_log.update([['실행일시', '파일명', '신규추가', '업데이트', '스킵(중복)', '비고']], 'A1')
        ws_log.format('A1:F1', {'textFormat': {'bold': True}})
        time.sleep(1)
    print(f'  OK [{SHEET_LOG}] 설정 완료')

    # ── 분류 규칙 시트 ─────────────────────────────────
    if SHEET_RULE not in existing:
        ws_rule = ss.add_worksheet(title=SHEET_RULE, rows=50, cols=2)
        rule_data = [['키워드', '분류']] + [[', '.join(kws), label] for kws, label in RULES]
        ws_rule.update(rule_data, 'A1')
        ws_rule.format('A1:B1', {'textFormat': {'bold': True}})
        time.sleep(1)
    print(f'  OK [{SHEET_RULE}] 설정 완료')

    print()
    print(f'OK 시트 초기화 완료')
    print(f'   URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}')

if __name__ == '__main__':
    main()
