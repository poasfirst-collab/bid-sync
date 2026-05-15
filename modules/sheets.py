import time
import gspread
from google.oauth2.service_account import Credentials
from config import SERVICE_ACCOUNT_FILE, SPREADSHEET_ID, SHEET_MAIN, SHEET_LOG, COLUMNS

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

_client = None


def _get_client():
    global _client
    if _client is None:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        _client = gspread.authorize(creds)
    return _client


def _retry(fn, retries=3, delay=2):
    for i in range(retries):
        try:
            return fn()
        except gspread.exceptions.APIError:
            if i == retries - 1:
                raise
            time.sleep(delay)


def get_sheet():
    return _get_client().open_by_key(SPREADSHEET_ID)


def read_main_sheet() -> list[list]:
    ws = get_sheet().worksheet(SHEET_MAIN)
    return _retry(lambda: ws.get_all_values())


def write_all(records: list[dict]):
    """서식 없이 텍스트만 전체 재작성"""
    ss = get_sheet()
    ws = ss.worksheet(SHEET_MAIN)

    rows = [COLUMNS]
    for rec in records:
        rows.append([str(rec.get(col, '')) for col in COLUMNS])

    _retry(lambda: ws.clear())
    time.sleep(0.5)
    # USER_ENTERED: 날짜/숫자 자동 변환 없이 텍스트 그대로
    _retry(lambda: ws.update(rows, 'A1', value_input_option='RAW'))


def append_to_sheet(record: dict):
    """단건 추가 (g2b 검색 결과)"""
    ss = get_sheet()
    ws = ss.worksheet(SHEET_MAIN)
    row = [str(record.get(col, '')) for col in COLUMNS]
    _retry(lambda: ws.append_row(row, value_input_option='RAW'))


def append_log(filename: str, inserts: int, updates: int, skips: int, note: str = ''):
    from datetime import datetime
    ws  = get_sheet().worksheet(SHEET_LOG)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    _retry(lambda: ws.append_row([now, filename, inserts, updates, skips, note]))
