"""
핀시큐리티 공공입찰 통합관리 - 동기화
usage: python bid_sync.py [--dir 폴더경로]
"""
import os, sys, glob, argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import RAW_DIR, SPREADSHEET_ID, SALES_SHEET_INDEX, SIMILARITY_THRESHOLD, RESULT_UPDATE_COLS
from modules.parser import parse_bid_notice, parse_bid_result, parse_sales
from modules.dedup import find_match, merge_into, merge_sales_into, sort_by_date, normalize_key
from modules.sheets import write_all, append_log
from modules.logger import log

DIV = '=' * 60


def find_file(directory, pattern):
    m = glob.glob(os.path.join(directory, pattern))
    return m[0] if m else None


def confirm_missing(missing):
    print()
    print('[경고] 아래 파일을 찾을 수 없습니다:')
    for f in missing:
        print(f'   - {f}')
    print('   누락 파일 데이터는 업데이트되지 않습니다.')
    while True:
        ans = input('   계속 진행하시겠습니까? [Y/N] > ').strip().upper()
        if ans == 'Y': return True
        if ans == 'N': return False


def sync(directory):
    print(DIV)
    print('  핀시큐리티 공공입찰 통합관리 - 동기화')
    print(f'  {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print(DIV)

    f_notice = find_file(directory, '입찰공고_*.xlsx')
    f_result = find_file(directory, '낙찰정보_*.xlsx')
    f_sales  = find_file(directory, '영업현황_*.xlsx')

    print('\n[파일 확인]')
    files = {'입찰공고_*.xlsx': f_notice, '낙찰정보_*.xlsx': f_result, '영업현황_*.xlsx': f_sales}
    for pat, path in files.items():
        print(f'  {"[OK]" if path else "[없음]"} {pat:<25} {os.path.basename(path) if path else "-"}')

    if not any(files.values()):
        print('\n  파일 없음. 종료합니다.')
        return

    missing = [p for p, f in files.items() if not f]
    if missing and not confirm_missing(missing):
        print('  취소.')
        return

    if not SPREADSHEET_ID:
        print('\n[오류] SPREADSHEET_ID 미설정.')
        return

    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    # ── Step1: 영업현황 + 입찰공고 병합 (중복 제거) ────────
    print('\n[Step1] 영업현황 + 입찰공고 병합')
    pool: list[dict] = []

    # 영업현황 먼저 (우선순위 높음)
    if f_sales:
        for rec in parse_sales(f_sales, SALES_SHEET_INDEX):
            pool.append(rec)

    # 입찰공고: 공고번호 중복이면 빈 필드만 채움, 없으면 신규 추가
    notice_new = notice_merge = 0
    if f_notice:
        for rec in parse_bid_notice(f_notice):
            idx = find_match(rec, pool, SIMILARITY_THRESHOLD)
            if idx >= 0:
                merge_sales_into(pool[idx], rec)   # 빈 필드 보완
                notice_merge += 1
            else:
                pool.append(rec)
                notice_new += 1

    print(f'  영업현황: {len([r for r in pool if r["원본출처"]=="영업현황"])}건')
    print(f'  입찰공고: 신규 {notice_new}건 / 기존 보완 {notice_merge}건')
    print(f'  병합 결과: {len(pool)}건')

    # ── Step2: 낙찰정보로 업데이트 ─────────────────────────
    result_upd = result_skip = 0
    if f_result:
        print(f'\n[Step2] 낙찰정보 업데이트: {os.path.basename(f_result)}')
        for rec in parse_bid_result(f_result):
            idx = find_match(rec, pool, SIMILARITY_THRESHOLD)
            if idx >= 0:
                rec['최종수정일'] = now
                merge_into(pool[idx], rec, RESULT_UPDATE_COLS)
                result_upd += 1
            else:
                result_skip += 1  # pool에 없으면 무시
        print(f'  업데이트 {result_upd}건 / 무시(미매칭) {result_skip}건')

    # ── Step3: 문의/공고일 기준 정렬 ───────────────────────
    pool = sort_by_date(pool)

    # ── Step4: 구글 시트 업로드 ─────────────────────────────
    print(f'\n[Step3] 구글 시트 업로드 중... ({len(pool)}건)')
    write_all(pool)
    append_log('동기화', len(pool), result_upd, 0, f'총{len(pool)}건')

    print()
    print(DIV)
    print(f'  완료: 총 {len(pool)}건 업로드')
    print(f'  시트: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}')
    print(DIV)
    log.info(f'완료 - 총:{len(pool)}건')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', default=RAW_DIR)
    args = parser.parse_args()
    sync(args.dir)
