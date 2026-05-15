"""
핀시큐리티 공공입찰 통합관리 - 웹 포털 (Streamlit)
"""
import sys, os, io, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
from datetime import datetime

from config import (
    SPREADSHEET_ID, SALES_SHEET_INDEX, SIMILARITY_THRESHOLD,
    G2B_API_KEY, RESULT_UPDATE_COLS, COLUMNS
)

st.set_page_config(
    page_title='공공입찰 통합관리 — NAS Portal',
    page_icon='📋',
    layout='wide',
)

st.markdown("""
<style>
/* ── 포털 디자인 시스템 변수 ── */
:root {
  --bg:      #f1f5f9;
  --surface: #ffffff;
  --text:    #0f172a;
  --muted:   #64748b;
  --border:  #e2e8f0;
  --accent:  #3b82f6;
  --accent-hv: #2563eb;
  --radius:  8px;
  --radius-lg: 14px;
  --shadow:  0 1px 3px rgba(0,0,0,.08);
  --shadow-lg: 0 10px 24px rgba(0,0,0,.10);
}

/* 상단 헤더 숨기기 / 여백 */
[data-testid="stHeader"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
.block-container { padding-top: 0 !important; padding-bottom: 2rem !important; max-width: 1280px; }

/* 포털 네비게이션 바 */
.portal-nav {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: 0 24px;
  height: 56px;
  display: flex;
  align-items: center;
  gap: 12px;
  margin: -1rem -1rem 1.5rem -1rem;
  box-shadow: 0 1px 3px rgba(0,0,0,.06);
}
.portal-brand-icon {
  width: 32px; height: 32px;
  border-radius: 9px;
  background: linear-gradient(135deg, #3b82f6, #8b5cf6);
  display: inline-flex; align-items: center; justify-content: center;
  color: #fff; font-size: 15px;
}
.portal-brand-text {
  font-size: 16px; font-weight: 700; color: var(--text);
}
.portal-back {
  margin-left: auto;
  font-size: 12px; color: var(--muted);
  text-decoration: none;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 4px 10px;
}
.portal-back:hover { color: var(--accent); border-color: var(--accent); }

/* 카드 */
.p-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow);
  padding: 20px;
  margin-bottom: 1rem;
}
.p-card-title {
  font-size: 13px; font-weight: 700;
  color: var(--muted); text-transform: uppercase;
  letter-spacing: .06em; margin-bottom: 14px;
}

/* 버튼 */
.stButton > button {
  background: var(--accent) !important;
  color: #fff !important;
  border: none !important;
  border-radius: var(--radius) !important;
  font-weight: 600 !important;
  font-size: 14px !important;
  padding: 8px 20px !important;
  transition: background .18s ease !important;
}
.stButton > button:hover {
  background: var(--accent-hv) !important;
}

/* 탭 */
[data-testid="stTabs"] [role="tab"] {
  font-size: 14px !important;
  font-weight: 600 !important;
  color: var(--muted) !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
  color: var(--accent) !important;
}
[data-testid="stTabs"] [role="tablist"] {
  border-bottom: 2px solid var(--border) !important;
}

/* 파일 업로더 */
[data-testid="stFileUploader"] {
  border: 1px solid var(--border) !important;
  border-radius: var(--radius) !important;
}

/* 입력 필드 */
[data-testid="stTextInput"] input {
  border-radius: var(--radius) !important;
  border-color: var(--border) !important;
  font-size: 14px !important;
}
[data-testid="stTextInput"] input:focus {
  border-color: var(--accent) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important;
}

/* 로그박스 */
.log-box {
  background: #0f172a; color: #cbd5e1;
  font-family: "SF Mono", Consolas, monospace;
  font-size: 12px; line-height: 1.6;
  padding: 16px; border-radius: var(--radius);
  max-height: 380px; overflow-y: auto;
  white-space: pre-wrap;
  border: 1px solid #1e293b;
}

/* 성공/경고/에러 */
.stSuccess { border-radius: var(--radius) !important; }
.stWarning { border-radius: var(--radius) !important; }
.stError   { border-radius: var(--radius) !important; }

/* 섹션 구분선 */
hr { border-color: var(--border) !important; margin: 1.25rem 0 !important; }

/* 메트릭 */
[data-testid="metric-container"] {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 14px 16px;
}
</style>

<div class="portal-nav">
  <span class="portal-brand-icon">📋</span>
  <span class="portal-brand-text">공공입찰 통합관리</span>
  <a class="portal-back" href="http://100.95.127.38:8080/dashboard">← 대시보드</a>
</div>
""", unsafe_allow_html=True)

st.markdown(f"<p style='color:#64748b;font-size:13px;margin-bottom:1.5rem;'>Sheet: <code>{SPREADSHEET_ID[:24]}…</code></p>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📂 파일 동기화", "🔍 나라장터 직접 검색"])


# ══════════════════════════════════════════════════════════
# 탭1: 파일 동기화
# ══════════════════════════════════════════════════════════
with tab1:
    st.subheader("엑셀 파일 업로드")

    col1, col2, col3 = st.columns(3)
    with col1:
        f_notice = st.file_uploader("입찰공고_*.xlsx", type="xlsx", key="notice")
    with col2:
        f_result = st.file_uploader("낙찰정보_*.xlsx", type="xlsx", key="result")
    with col3:
        f_sales = st.file_uploader("영업현황_*.xlsx", type="xlsx", key="sales")

    missing = []
    if not f_notice: missing.append("입찰공고")
    if not f_result: missing.append("낙찰정보")
    if not f_sales:  missing.append("영업현황")

    if missing and (f_notice or f_result or f_sales):
        st.warning(f"누락 파일 (업로드 없이 진행): {', '.join(missing)}")

    if not (f_notice or f_result or f_sales):
        st.info("파일을 하나 이상 업로드하면 동기화 버튼이 활성화됩니다.")

    run_btn = st.button("▶ 구글 시트 동기화 실행", disabled=not (f_notice or f_result or f_sales))

    log_placeholder = st.empty()

    if run_btn:
        log_lines = []

        def log(msg):
            log_lines.append(msg)
            log_placeholder.markdown(
                f'<div class="log-box">' + "\n".join(log_lines) + '</div>',
                unsafe_allow_html=True
            )

        try:
            import pandas as pd
            from modules.parser import parse_bid_notice, parse_bid_result, parse_sales
            from modules.dedup import find_match, merge_into, merge_sales_into, sort_by_date
            from modules.sheets import write_all, append_log

            now = datetime.now().strftime('%Y-%m-%d %H:%M')
            log("=" * 52)
            log(f"  동기화 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            log("=" * 52)

            pool = []

            # Step1: 영업현황 + 입찰공고 병합
            log("\n[Step1] 영업현황 + 입찰공고 병합")

            if f_sales:
                sales_bytes = io.BytesIO(f_sales.read())
                for rec in parse_sales(sales_bytes, SALES_SHEET_INDEX):
                    pool.append(rec)

            notice_new = notice_merge = 0
            if f_notice:
                notice_bytes = io.BytesIO(f_notice.read())
                for rec in parse_bid_notice(notice_bytes):
                    idx = find_match(rec, pool, SIMILARITY_THRESHOLD)
                    if idx >= 0:
                        merge_sales_into(pool[idx], rec)
                        notice_merge += 1
                    else:
                        pool.append(rec)
                        notice_new += 1

            log(f"  영업현황: {len([r for r in pool if r.get('원본출처')=='영업현황'])}건")
            log(f"  입찰공고: 신규 {notice_new}건 / 기존 보완 {notice_merge}건")
            log(f"  병합 결과: {len(pool)}건")

            # Step2: 낙찰정보 업데이트
            result_upd = result_skip = 0
            if f_result:
                log(f"\n[Step2] 낙찰정보 업데이트")
                result_bytes = io.BytesIO(f_result.read())
                for rec in parse_bid_result(result_bytes):
                    idx = find_match(rec, pool, SIMILARITY_THRESHOLD)
                    if idx >= 0:
                        rec['최종수정일'] = now
                        merge_into(pool[idx], rec, RESULT_UPDATE_COLS)
                        result_upd += 1
                    else:
                        result_skip += 1
                log(f"  업데이트 {result_upd}건 / 무시(미매칭) {result_skip}건")

            # Step3: 정렬
            pool = sort_by_date(pool)

            # Step4: 업로드
            log(f"\n[Step3] 구글 시트 업로드 중... ({len(pool)}건)")
            write_all(pool)
            append_log('웹동기화', len(pool), result_upd, 0, f'총{len(pool)}건')

            log("\n" + "=" * 52)
            log(f"  완료: 총 {len(pool)}건 업로드")
            log("=" * 52)

            st.success(f"동기화 완료 — 총 {len(pool)}건 업로드")
            st.markdown(
                f"[구글 시트 열기](https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID})",
                unsafe_allow_html=False
            )

        except Exception as e:
            import traceback
            log(f"\n[오류] {e}")
            log(traceback.format_exc())
            st.error(f"오류 발생: {e}")


# ══════════════════════════════════════════════════════════
# 탭2: 나라장터 직접 검색
# ══════════════════════════════════════════════════════════
with tab2:
    st.subheader("공고번호 직접 조회")
    st.caption("예: R26BK01516539 (입찰공고) 또는 R26BD00219883 (사전규격)")

    col_inp, col_btn = st.columns([4, 1])
    with col_inp:
        bid_input = st.text_input("공고번호", placeholder="R26BK...", label_visibility="collapsed")
    with col_btn:
        search_btn = st.button("검색", use_container_width=True)

    if "g2b_result" not in st.session_state:
        st.session_state.g2b_result = None

    if search_btn and bid_input.strip():
        bid_no = bid_input.strip().upper()
        with st.spinner(f"검색 중: {bid_no}"):
            try:
                from modules.g2b_api import fetch_bid, fetch_prespec
                if re.match(r'^R\d{2}BD', bid_no, re.I):
                    data = fetch_prespec(G2B_API_KEY, bid_no)
                    label = "사전규격"
                else:
                    data = fetch_bid(G2B_API_KEY, bid_no)
                    label = "입찰공고"

                if data:
                    data['_label'] = label
                    data['_bid_no_input'] = bid_no
                    st.session_state.g2b_result = data
                else:
                    st.warning(f"찾을 수 없습니다: {bid_no}")
                    st.session_state.g2b_result = None
            except Exception as e:
                st.error(f"API 오류: {e}")
                st.session_state.g2b_result = None

    data = st.session_state.g2b_result
    if data:
        label = data.get('_label', '')
        budget = data.get('예산', 0)
        budget_str = f"{int(budget):,}원" if budget else "-"

        st.success(f"[{label}] 발견")

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**공고번호** {data.get('공고번호', '-')}")
            st.markdown(f"**공고명** {data.get('공고명', '-')}")
            st.markdown(f"**발주기관** {data.get('발주기관', '-') or data.get('공고기관', '-')}")
            st.markdown(f"**예산** {budget_str}")
        with col_b:
            st.markdown(f"**입찰마감** {data.get('참가마감', '-')}")
            st.markdown(f"**개찰일시** {data.get('개찰일시', '-') or data.get('개찰일', '-')}")
            url = data.get('나라장터URL', '')
            if url:
                st.markdown(f"**나라장터** [{url}]({url})")

        st.divider()

        # 중복 체크 후 추가 버튼
        if st.button("구글 시트 중복 확인 후 추가"):
            bid_no = data['_bid_no_input']
            with st.spinner("중복 확인 중..."):
                try:
                    from modules.sheets import read_main_sheet, append_to_sheet
                    from modules.dedup import find_match
                    from modules.classifier import classify

                    rows = read_main_sheet()
                    pool = [
                        {'공고번호': r[3], '사업명': r[4]}
                        for r in rows[1:] if len(r) > 4
                    ]
                    check = {'공고번호': bid_no, '사업명': data.get('공고명', '')}

                    if find_match(check, pool, 80) >= 0:
                        st.warning("이미 시트에 존재합니다. 추가하지 않았습니다.")
                    else:
                        seq = str(data.get('공고구분순번', '000')).strip()
                        bid_no_full = f"{bid_no}-{seq}" if seq != '000' else bid_no
                        if re.match(r'^R\d{2}BD', bid_no, re.I):
                            kind = '사전공고'
                        elif seq != '000':
                            kind = f'재공고({int(seq)}차)'
                        else:
                            kind = '본공고'

                        now = datetime.now().strftime('%Y-%m-%d %H:%M')
                        rec = {
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
                            '개찰일시':                    data.get('개찰일시', '') or data.get('개찰일', ''),
                            '단계':                        '',
                            '원본출처':                    '나라장터_API',
                            '최종수정일':                  now,
                        }
                        append_to_sheet(rec)
                        st.success("구글 시트에 추가 완료")
                        st.session_state.g2b_result = None
                        st.rerun()

                except Exception as e:
                    st.error(f"오류: {e}")
