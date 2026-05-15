"""
핀시큐리티 공공입찰 통합관리 - GUI
usage: python app.py
"""
import sys, os, threading, io, glob
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config import SPREADSHEET_ID, SALES_SHEET_INDEX, SIMILARITY_THRESHOLD, G2B_API_KEY
from modules.logger import log

# ── 색상 상수 ───────────────────────────────────────────────
C_HEADER  = '#1B5E20'   # 헤더 진녹색
C_BTN     = '#2E7D32'   # 버튼 녹색
C_BTN_H   = '#1B5E20'   # 버튼 hover
C_BTN2    = '#1565C0'   # 검색 버튼 파란
C_BG      = '#F5F5F5'   # 배경
C_WHITE   = '#FFFFFF'
C_RED     = '#C62828'
C_TEXT    = '#212121'
C_BORDER  = '#BDBDBD'
C_LOG_BG  = '#1E1E1E'
C_LOG_FG  = '#D4D4D4'
C_OK      = '#43A047'
C_WARN    = '#FB8C00'


class LogRedirect(io.StringIO):
    """print 출력을 Text 위젯으로 리다이렉트"""
    def __init__(self, widget: tk.Text):
        super().__init__()
        self.widget = widget

    def write(self, text):
        self.widget.after(0, self._insert, text)

    def _insert(self, text):
        self.widget.config(state='normal')
        self.widget.insert('end', text)
        self.widget.see('end')
        self.widget.config(state='disabled')

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('핀시큐리티 공공입찰 통합관리')
        self.geometry('900x680')
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = 900, 680
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f'{w}x{h}+{x}+{y}')

    def _build_ui(self):
        # ── 헤더 바 ───────────────────────────────────────
        header = tk.Frame(self, bg=C_HEADER, height=52)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='핀시큐리티  공공입찰 통합관리',
                 bg=C_HEADER, fg=C_WHITE,
                 font=('맑은 고딕', 14, 'bold')).pack(side='left', padx=20, pady=12)
        tk.Label(header, text=f'Sheet ID: {SPREADSHEET_ID[:20]}...',
                 bg=C_HEADER, fg='#A5D6A7',
                 font=('맑은 고딕', 9)).pack(side='right', padx=20)

        # ── 탭 ────────────────────────────────────────────
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook',        background=C_BG, borderwidth=0)
        style.configure('TNotebook.Tab',    background='#E0E0E0', foreground=C_TEXT,
                         font=('맑은 고딕', 10), padding=[14, 6])
        style.map('TNotebook.Tab',
                  background=[('selected', C_WHITE)],
                  foreground=[('selected', C_HEADER)])

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=12, pady=(8, 4))

        tab1 = tk.Frame(nb, bg=C_WHITE)
        tab2 = tk.Frame(nb, bg=C_WHITE)
        nb.add(tab1, text='  파일 동기화  ')
        nb.add(tab2, text='  나라장터 직접 검색  ')

        self._build_sync_tab(tab1)
        self._build_search_tab(tab2)

        # ── 상태바 ────────────────────────────────────────
        self.status_var = tk.StringVar(value='준비')
        status_bar = tk.Frame(self, bg='#E0E0E0', height=26)
        status_bar.pack(fill='x', side='bottom')
        status_bar.pack_propagate(False)
        tk.Label(status_bar, textvariable=self.status_var,
                 bg='#E0E0E0', fg='#555', font=('맑은 고딕', 9)).pack(side='left', padx=10)

    # ═══════════════════════════════════════════════════════
    # 탭1: 파일 동기화
    # ═══════════════════════════════════════════════════════
    def _build_sync_tab(self, parent):
        pad = {'padx': 20, 'pady': 6}

        # 파일 선택 섹션
        section = tk.LabelFrame(parent, text='  엑셀 파일 선택  ',
                                 bg=C_WHITE, fg=C_HEADER,
                                 font=('맑은 고딕', 10, 'bold'),
                                 relief='groove', bd=1)
        section.pack(fill='x', **pad)

        self.file_vars = {}
        file_defs = [
            ('입찰공고', '입찰공고_*.xlsx', '비드메이트 입찰공고 파일'),
            ('낙찰정보', '낙찰정보_*.xlsx', '비드메이트 낙찰정보 파일'),
            ('영업현황', '영업현황_*.xlsx', '내부 영업현황 파일 (시트2)'),
        ]
        for label, pattern, hint in file_defs:
            self._file_row(section, label, pattern, hint)

        # 폴더 일괄 선택 버튼
        folder_row = tk.Frame(section, bg=C_WHITE)
        folder_row.pack(fill='x', padx=12, pady=(4, 10))
        tk.Button(folder_row, text='폴더에서 자동 탐색',
                  bg='#ECEFF1', fg=C_TEXT, relief='flat',
                  font=('맑은 고딕', 9), cursor='hand2',
                  command=self._pick_folder).pack(side='left')
        tk.Label(folder_row,
                 text='파일명 패턴: 입찰공고_*.xlsx / 낙찰정보_*.xlsx / 영업현황_*.xlsx',
                 bg=C_WHITE, fg='#9E9E9E', font=('맑은 고딕', 8)).pack(side='left', padx=10)

        # 실행 버튼
        btn_row = tk.Frame(parent, bg=C_WHITE)
        btn_row.pack(fill='x', padx=20, pady=(4, 10))
        self.sync_btn = tk.Button(
            btn_row, text='  구글 시트 동기화 실행  ',
            bg=C_BTN, fg=C_WHITE, activebackground=C_BTN_H, activeforeground=C_WHITE,
            font=('맑은 고딕', 11, 'bold'), relief='flat', cursor='hand2',
            padx=20, pady=8, command=self._run_sync)
        self.sync_btn.pack(side='left')

        self.progress = ttk.Progressbar(btn_row, mode='indeterminate', length=200)
        self.progress.pack(side='left', padx=16)

        # 로그 박스
        log_frame = tk.LabelFrame(parent, text='  실행 로그  ',
                                   bg=C_WHITE, fg=C_HEADER,
                                   font=('맑은 고딕', 10, 'bold'),
                                   relief='groove', bd=1)
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 10))

        self.log_box = tk.Text(log_frame, bg=C_LOG_BG, fg=C_LOG_FG,
                                font=('Consolas', 9), relief='flat',
                                state='disabled', wrap='word')
        sb = ttk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.log_box.pack(fill='both', expand=True, padx=4, pady=4)

    def _file_row(self, parent, label, pattern, hint):
        row = tk.Frame(parent, bg=C_WHITE)
        row.pack(fill='x', padx=12, pady=3)

        tk.Label(row, text=f'{label}', width=8, anchor='w',
                 bg=C_WHITE, fg=C_TEXT, font=('맑은 고딕', 10)).pack(side='left')

        var = tk.StringVar(value='')
        self.file_vars[label] = var

        entry = tk.Entry(row, textvariable=var, width=52,
                         font=('맑은 고딕', 9), relief='solid', bd=1,
                         bg='#FAFAFA', fg=C_TEXT)
        entry.pack(side='left', padx=(4, 6))

        tk.Button(row, text='찾아보기',
                  bg='#E8F5E9', fg=C_BTN, relief='flat',
                  font=('맑은 고딕', 9), cursor='hand2',
                  command=lambda l=label: self._pick_file(l)).pack(side='left')

        tk.Label(row, text=hint, bg=C_WHITE, fg='#9E9E9E',
                 font=('맑은 고딕', 8)).pack(side='left', padx=8)

    def _pick_file(self, label):
        path = filedialog.askopenfilename(
            title=f'{label} 파일 선택',
            filetypes=[('Excel 파일', '*.xlsx'), ('모든 파일', '*.*')])
        if path:
            self.file_vars[label].set(path)

    def _pick_folder(self):
        folder = filedialog.askdirectory(title='엑셀 파일이 있는 폴더 선택')
        if not folder:
            return
        patterns = {
            '입찰공고': '입찰공고_*.xlsx',
            '낙찰정보': '낙찰정보_*.xlsx',
            '영업현황': '영업현황_*.xlsx',
        }
        found = 0
        for label, pat in patterns.items():
            matches = glob.glob(os.path.join(folder, pat))
            if matches:
                self.file_vars[label].set(matches[0])
                found += 1
        self.status_var.set(f'폴더 탐색 완료 — {found}개 파일 발견')

    def _run_sync(self):
        paths = {k: v.get() for k, v in self.file_vars.items()}

        # 없는 파일 경고
        missing = [k for k, v in paths.items() if not v]
        if missing:
            msg = f'아래 파일이 선택되지 않았습니다:\n  - {chr(10).join(missing)}\n\n계속 진행하시겠습니까?'
            if not messagebox.askyesno('파일 없음 경고', msg):
                return

        if all(not v for v in paths.values()):
            messagebox.showwarning('파일 없음', '업로드할 파일을 하나 이상 선택해주세요.')
            return

        self.sync_btn.config(state='disabled')
        self.progress.start(12)
        self.status_var.set('동기화 진행 중...')

        # print → 로그박스 리다이렉트
        redir = LogRedirect(self.log_box)
        old_stdout = sys.stdout
        sys.stdout = redir

        def worker():
            try:
                from datetime import datetime
                from config import RESULT_UPDATE_COLS, SIMILARITY_THRESHOLD as THR
                from modules.parser import parse_bid_notice, parse_bid_result, parse_sales
                from modules.dedup import find_match, merge_into, merge_sales_into, sort_by_date
                from modules.sheets import write_all, append_log

                now = datetime.now().strftime('%Y-%m-%d %H:%M')
                print('=' * 54)
                print('  동기화 시작:', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                print('=' * 54)

                f_notice = paths.get('입찰공고', '')
                f_result = paths.get('낙찰정보', '')
                f_sales  = paths.get('영업현황', '')
                pool = []

                # Step1: 영업현황 + 입찰공고 병합
                print('\n[Step1] 영업현황 + 입찰공고 병합')
                if f_sales and os.path.exists(f_sales):
                    for rec in parse_sales(f_sales, SALES_SHEET_INDEX):
                        pool.append(rec)

                notice_new = notice_merge = 0
                if f_notice and os.path.exists(f_notice):
                    for rec in parse_bid_notice(f_notice):
                        idx = find_match(rec, pool, THR)
                        if idx >= 0:
                            merge_sales_into(pool[idx], rec)
                            notice_merge += 1
                        else:
                            pool.append(rec)
                            notice_new += 1

                print(f'  영업현황: {len([r for r in pool if r.get("원본출처")=="영업현황"])}건')
                print(f'  입찰공고: 신규 {notice_new}건 / 기존 보완 {notice_merge}건')
                print(f'  병합 결과: {len(pool)}건')

                # Step2: 낙찰정보 업데이트
                result_upd = result_skip = 0
                if f_result and os.path.exists(f_result):
                    print(f'\n[Step2] 낙찰정보 업데이트')
                    for rec in parse_bid_result(f_result):
                        idx = find_match(rec, pool, THR)
                        if idx >= 0:
                            rec['최종수정일'] = now
                            merge_into(pool[idx], rec, RESULT_UPDATE_COLS)
                            result_upd += 1
                        else:
                            result_skip += 1  # pool에 없으면 무시
                    print(f'  업데이트 {result_upd}건 / 무시(미매칭) {result_skip}건')

                pool = sort_by_date(pool)

                print(f'\n[Step3] 구글 시트 업로드 중... ({len(pool)}건)')
                write_all(pool)
                append_log('GUI동기화', len(pool), result_upd, 0, f'총{len(pool)}건')

                print('\n' + '=' * 54)
                print(f'  완료: 총 {len(pool)}건 업로드')
                print('=' * 54)

                self.after(0, lambda: self.status_var.set(f'완료 — 총 {len(pool)}건 업로드'))

            except Exception as e:
                import traceback
                print(f'\n[오류] {e}')
                print(traceback.format_exc())
                log.error(str(e))
                self.after(0, lambda: self.status_var.set(f'오류 발생: {e}'))
            finally:
                sys.stdout = old_stdout
                self.after(0, self._sync_done)

        threading.Thread(target=worker, daemon=True).start()

    def _sync_done(self):
        self.sync_btn.config(state='normal')
        self.progress.stop()

    # ═══════════════════════════════════════════════════════
    # 탭2: 나라장터 직접 검색
    # ═══════════════════════════════════════════════════════
    def _build_search_tab(self, parent):
        pad = {'padx': 20, 'pady': 8}

        # 검색 입력
        section = tk.LabelFrame(parent, text='  공고번호 직접 조회  ',
                                  bg=C_WHITE, fg=C_BTN2,
                                  font=('맑은 고딕', 10, 'bold'),
                                  relief='groove', bd=1)
        section.pack(fill='x', **pad)

        inp_row = tk.Frame(section, bg=C_WHITE)
        inp_row.pack(fill='x', padx=12, pady=10)

        tk.Label(inp_row, text='공고번호', bg=C_WHITE, fg=C_TEXT,
                 font=('맑은 고딕', 10)).pack(side='left')

        self.search_var = tk.StringVar()
        entry = tk.Entry(inp_row, textvariable=self.search_var, width=30,
                         font=('맑은 고딕', 11), relief='solid', bd=1)
        entry.pack(side='left', padx=(8, 10))
        entry.bind('<Return>', lambda e: self._run_search())

        tk.Label(inp_row, text='예: R26BK01516539  또는  R26BD00219883',
                 bg=C_WHITE, fg='#9E9E9E', font=('맑은 고딕', 9)).pack(side='left')

        self.search_btn = tk.Button(
            inp_row, text='  검색  ',
            bg=C_BTN2, fg=C_WHITE, activebackground='#0D47A1',
            font=('맑은 고딕', 10, 'bold'), relief='flat', cursor='hand2',
            padx=14, pady=4, command=self._run_search)
        self.search_btn.pack(side='right', padx=4)

        # 결과 표시
        result_frame = tk.LabelFrame(parent, text='  검색 결과  ',
                                      bg=C_WHITE, fg=C_BTN2,
                                      font=('맑은 고딕', 10, 'bold'),
                                      relief='groove', bd=1)
        result_frame.pack(fill='x', padx=20, pady=(0, 8))

        self.result_fields = {}
        fields = [
            ('공고번호', ''), ('공고명', ''), ('발주기관', ''),
            ('예산', ''), ('입찰마감', ''), ('개찰일', ''), ('나라장터URL', ''),
        ]
        for i, (fname, _) in enumerate(fields):
            row = tk.Frame(result_frame, bg=C_WHITE)
            row.pack(fill='x', padx=12, pady=2)
            tk.Label(row, text=f'{fname}', width=10, anchor='w',
                     bg=C_WHITE, fg='#616161',
                     font=('맑은 고딕', 9)).pack(side='left')
            var = tk.StringVar(value='-')
            self.result_fields[fname] = var
            lbl = tk.Label(row, textvariable=var, anchor='w',
                           bg=C_WHITE, fg=C_TEXT, font=('맑은 고딕', 10),
                           wraplength=680, justify='left')
            lbl.pack(side='left', padx=4)

        # 추가 버튼 + 상태
        add_row = tk.Frame(result_frame, bg=C_WHITE)
        add_row.pack(fill='x', padx=12, pady=(4, 10))

        self.add_btn = tk.Button(
            add_row, text='  구글 시트에 추가  ',
            bg=C_BTN, fg=C_WHITE, activebackground=C_BTN_H,
            font=('맑은 고딕', 10, 'bold'), relief='flat', cursor='hand2',
            padx=14, pady=5, state='disabled', command=self._add_to_sheet)
        self.add_btn.pack(side='left')

        self.add_status = tk.Label(add_row, text='', bg=C_WHITE,
                                    font=('맑은 고딕', 9))
        self.add_status.pack(side='left', padx=12)

        # 검색 로그
        log_frame = tk.LabelFrame(parent, text='  검색 로그  ',
                                   bg=C_WHITE, fg=C_BTN2,
                                   font=('맑은 고딕', 10, 'bold'),
                                   relief='groove', bd=1)
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 10))
        self.search_log = tk.Text(log_frame, bg=C_LOG_BG, fg=C_LOG_FG,
                                   font=('Consolas', 9), relief='flat',
                                   state='disabled', height=8, wrap='word')
        sb2 = ttk.Scrollbar(log_frame, command=self.search_log.yview)
        self.search_log.configure(yscrollcommand=sb2.set)
        sb2.pack(side='right', fill='y')
        self.search_log.pack(fill='both', expand=True, padx=4, pady=4)

        self._current_result = None

    def _slog(self, msg):
        self.search_log.config(state='normal')
        self.search_log.insert('end', msg + '\n')
        self.search_log.see('end')
        self.search_log.config(state='disabled')

    def _run_search(self):
        bid_no = self.search_var.get().strip().upper()
        if not bid_no:
            return

        self.search_btn.config(state='disabled')
        self.add_btn.config(state='disabled')
        self.add_status.config(text='')
        self.status_var.set(f'검색 중: {bid_no}')
        self._slog(f'\n검색: {bid_no}')

        def worker():
            try:
                import re
                from modules.g2b_api import fetch_bid, fetch_prespec

                if re.match(r'^R\d{2}BD', bid_no, re.I):
                    data = fetch_prespec(G2B_API_KEY, bid_no)
                    label = '사전규격'
                else:
                    data = fetch_bid(G2B_API_KEY, bid_no)
                    label = '입찰공고'

                if not data:
                    self.after(0, lambda: self._search_result(None, bid_no))
                else:
                    self.after(0, lambda d=data, lb=label: self._search_result(d, bid_no, lb))
            except Exception as e:
                self.after(0, lambda: self._search_error(str(e)))
            finally:
                self.after(0, lambda: self.search_btn.config(state='normal'))

        threading.Thread(target=worker, daemon=True).start()

    def _search_result(self, data, bid_no, label=''):
        if not data:
            self._slog(f'  찾을 수 없습니다: {bid_no}')
            self.status_var.set('검색 결과 없음')
            for var in self.result_fields.values():
                var.set('-')
            self._current_result = None
            return

        self._current_result = data
        budget = data.get('예산', 0)
        budget_str = f'{int(budget):,}원' if budget else '-'

        self.result_fields['공고번호'].set(data.get('공고번호', '-'))
        self.result_fields['공고명'].set(data.get('공고명', '-'))
        self.result_fields['발주기관'].set(data.get('발주기관', '-') or data.get('공고기관', '-'))
        self.result_fields['예산'].set(budget_str)
        self.result_fields['입찰마감'].set(data.get('참가마감', '-'))
        self.result_fields['개찰일'].set(data.get('개찰일시', '-') or data.get('개찰일', '-'))
        self.result_fields['나라장터URL'].set(data.get('나라장터URL', '-'))

        self._slog(f'  [{label}] 발견: {data.get("공고명", "")}')

        # 중복 체크
        def check_dup():
            try:
                from modules.sheets import read_main_sheet
                from modules.dedup import find_match
                rows = read_main_sheet()
                pool = [{'공고번호': r[3], '사업명': r[4]}
                        for r in rows[1:] if len(r) > 4]
                check = {'공고번호': bid_no, '사업명': data.get('공고명', '')}
                if find_match(check, pool, 80) >= 0:
                    self.after(0, lambda: self.add_status.config(
                        text='이미 시트에 존재합니다', fg=C_WARN))
                    self.after(0, lambda: self.add_btn.config(state='disabled'))
                    self.after(0, lambda: self._slog('  [중복] 이미 존재하는 공고'))
                else:
                    self.after(0, lambda: self.add_btn.config(state='normal'))
                    self.after(0, lambda: self.add_status.config(
                        text='추가 가능', fg=C_OK))
            except Exception as e:
                self.after(0, lambda: self.add_btn.config(state='normal'))

        threading.Thread(target=check_dup, daemon=True).start()
        self.status_var.set(f'검색 완료: {data.get("공고명", "")}')

    def _search_error(self, msg):
        self._slog(f'  [오류] {msg}')
        self.status_var.set(f'오류: {msg}')

    def _add_to_sheet(self):
        if not self._current_result:
            return
        self.add_btn.config(state='disabled')

        def worker():
            try:
                import gspread, re
                from google.oauth2.service_account import Credentials
                from config import SERVICE_ACCOUNT_FILE, SPREADSHEET_ID, COLUMNS
                from modules.classifier import classify
                from datetime import datetime

                data = self._current_result
                bid_no = str(data.get('공고번호', '')).strip().upper()
                seq = str(data.get('공고구분순번', '000')).strip()
                bid_no_full = f'{bid_no}-{seq}' if seq != '000' else bid_no
                kind = f'재공고({int(seq)}차)' if seq != '000' else '본공고'
                if re.match(r'^R\d{2}BD', bid_no, re.I):
                    kind = '사전공고'

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
                creds = Credentials.from_service_account_file(
                    SERVICE_ACCOUNT_FILE,
                    scopes=['https://www.googleapis.com/auth/spreadsheets',
                            'https://www.googleapis.com/auth/drive'])
                gc = gspread.authorize(creds)
                ws = gc.open_by_key(SPREADSHEET_ID).worksheet('통합공고관리')
                ws.append_row([str(rec.get(col, '')) for col in COLUMNS],
                              value_input_option='RAW')

                self.after(0, lambda: self.add_status.config(
                    text='추가 완료', fg=C_OK))
                self.after(0, lambda: self._slog(
                    f'  추가 완료: {data.get("공고명", "")}'))
                self.after(0, lambda: self.status_var.set('구글 시트 추가 완료'))
                log.info(f'추가: {bid_no_full}')
            except Exception as e:
                self.after(0, lambda: self.add_status.config(
                    text=f'오류: {e}', fg=C_RED))
                self.after(0, lambda: self._slog(f'  [오류] {e}'))
                self.after(0, lambda: self.add_btn.config(state='normal'))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()
