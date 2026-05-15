import os, sys

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS                          # 번들 리소스 (읽기전용)
    APP_DIR  = os.path.dirname(sys.executable)       # exe 옆 폴더 (쓰기)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_DIR  = BASE_DIR

SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'service_account.json')
RAW_DIR  = os.path.join(APP_DIR, 'raw')
LOG_DIR  = os.path.join(APP_DIR, 'logs')

SPREADSHEET_ID   = '1gI6CJbGdXMtWJk-e2lDXyimn9cxXxgnG6p7SAt2ZaXI'
SHEET_MAIN       = '통합공고관리'
SHEET_LOG        = 'SYNC_LOG'
SHEET_RULE       = 'CATEGORY_RULE'

G2B_API_KEY      = 'd9d6cf7019685b20326957a7e6828dbbca2e72cab82b8284ef5039988f5f28a2'
SALES_SHEET_INDEX   = 1
SIMILARITY_THRESHOLD = 80

# 낙찰정보 업데이트 대상 컬럼
RESULT_UPDATE_COLS = [
    '낙찰사', '낙찰율', '참여업체수', '내순위',
    '개찰일시', '단계', '원본출처', '최종수정일',
]

# 구글 시트 컬럼 (서식 없음, 텍스트만)
COLUMNS = [
    '영업속성',
    '문의 / 공고일',
    '사전/본공고',
    '공고번호',
    '사업명',
    'RFP',
    '수요처',
    '사업종류',
    '사업예산 (VAT포함)',
    '입찰 마감일',
    '사업 예상 시기',
    '대응여부',
    '사업 제안 지원부서 / 담당자',
    '비고',
    '낙찰사',
    '낙찰율',
    '참여업체수',
    '내순위',
    '개찰일시',
    '단계',
    '원본출처',
    '최종수정일',
]
