"""
run_gui.py  ─  KT 정산 자동화 GUI 실행기
===========================================
더블클릭 또는 [정산실행.bat] 으로 실행합니다.
run_settlement.py 와 같은 폴더에 있어야 합니다.

필요 패키지: pandas, openpyxl  (최초 1회 자동 설치)
"""

# ── 0. 의존성 확인 (설치 안 된 경우에만 pip 실행) ────────────────
import importlib, subprocess, sys

def _try_install(pkg):
    """이미 import 가능하면 pip 호출 없이 바로 반환.
    사내 pip 미러 타임아웃 시 공식 PyPI로 재시도."""
    if importlib.util.find_spec(pkg) is not None:
        return  # 이미 설치됨
    print(f"[설치 중] {pkg} ...")

    def _pip(extra_args=()):
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pkg, "-q", *extra_args],
            timeout=30,
        )

    # 1차: 현재 pip 설정(사내 미러) 그대로 시도
    try:
        _pip()
        return
    except Exception as e1:
        print(f"[경고] 기본 미러 설치 실패: {e1}")

    # 2차: 공식 PyPI 로 재시도 (사내 미러 접속 불가 시)
    try:
        print(f"[재시도] 공식 PyPI 에서 {pkg} 설치 중...")
        _pip(["--index-url", "https://pypi.org/simple/",
              "--trusted-host", "pypi.org"])
        return
    except Exception as e2:
        print(f"[경고] 공식 PyPI 설치도 실패: {e2}\n"
              f"       수동으로 설치해 주세요: pip install {pkg}")

for _pkg in ["pandas", "openpyxl"]:
    _try_install(_pkg)

# ── 1. 표준 라이브러리 ───────────────────────────────────────────
import os, json, datetime, calendar, threading, logging, traceback
from pathlib import Path

# ── 2. Tkinter (Python 기본 내장) ────────────────────────────────
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ── 3. run_settlement 임포트 ─────────────────────────────────────
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

try:
    from run_settlement import (
        SettlementRunner, MappingMaster,
        build_col_map, detect_col,
        _KT_CANDIDATES, _PL_CANDIDATES, _PL_OPT,
        ReturnMappingNeeded,
    )
except ImportError as _ie:
    # Tkinter 없이도 오류 메시지를 보여주기 위한 fallback
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror(
            "임포트 오류",
            f"run_settlement.py 를 찾을 수 없습니다.\n"
            f"run_gui.py 와 같은 폴더에 있는지 확인하세요.\n\n{_ie}"
        )
    except Exception:
        print(f"[오류] run_settlement.py 임포트 실패: {_ie}")
    sys.exit(1)


# ════════════════════════════════════════════════════════
# 로그 → Tkinter Text 위젯으로 실시간 스트리밍
# ════════════════════════════════════════════════════════
class _TkLogHandler(logging.Handler):
    _TAG = {
        "DEBUG":    "gray",
        "INFO":     "info",
        "WARNING":  "warn",
        "ERROR":    "err",
        "CRITICAL": "err",
    }

    def __init__(self, text_widget: scrolledtext.ScrolledText):
        super().__init__()
        self._w = text_widget

    def emit(self, record: logging.LogRecord):
        msg = self.format(record)
        tag = self._TAG.get(record.levelname, "info")
        def _append():
            self._w.configure(state="normal")
            self._w.insert("end", msg + "\n", tag)
            self._w.configure(state="disabled")
            self._w.see("end")
        # GUI 스레드에서 안전하게 실행
        self._w.after(0, _append)


# ════════════════════════════════════════════════════════
# input() 없이 동작하는 GUI 전용 SettlementRunner
# ════════════════════════════════════════════════════════
class _GUIRunner(SettlementRunner):
    """파일 경로·정산 기간을 생성자로 받아 사용자 입력 없이 전체 실행"""

    def __init__(self, kt: str, pl: str, mm: str, rm: str,
                 start: datetime.date, end: datetime.date,
                 round_override: str = None):
        # base_dir = KT 파일이 있는 폴더 → 결과 파일도 같은 위치에 저장
        super().__init__(base_dir=str(Path(kt).parent))
        self._kt, self._pl = kt, pl
        self._mm, self._rm = mm, rm
        self._start, self._end = start, end
        self._round_override = round_override  # None이면 자동 판단

    # ── 파일 탐지: input() 대신 생성자 파라미터 사용 ──────────────
    def detect_files(self):
        self.log.info("=" * 60)
        self.log.info("  KT 정산 자동화 시스템  v2.0")
        self.log.info("=" * 60)

        self.kt_path = Path(self._kt) if self._kt else None
        self.pl_path = Path(self._pl) if self._pl else None
        self.rm_path = Path(self._rm) if self._rm else None
        mm_path      = Path(self._mm) if self._mm else None

        # 중견기업목록.xlsx: KT 파일과 같은 폴더에서 탐색
        midcorp_path = None
        if self.base.exists():
            _mc = self.base / "중견기업목록.xlsx"
            if _mc.exists():
                midcorp_path = _mc

        for label, p in [
            ("KT 파일        ", self.kt_path),
            ("플랫폼 파일    ", self.pl_path),
            ("소싱그룹 매핑  ", mm_path),
            ("반품 매핑      ", self.rm_path),
            ("중견기업목록   ", midcorp_path),
        ]:
            self.log.info(f"  {label}: {p or '없음'}")

        self.master = MappingMaster(mm_path, midcorp_path)
        self.master.load(self.log)

    # ── 컬럼 탐지: 사용자 확인 단계 생략 ─────────────────────────
    def step0_init(self):
        from run_settlement import _read_excel_smart
        self.log.info("\n[STEP 0] 파일 컬럼 구조 분석")
        df_kt = _read_excel_smart(self.kt_path, search_cols=["구매문서번호"], nrows=3, dtype=str)
        df_pl = _read_excel_smart(self.pl_path, search_cols=["주문번호", "일정산번호"], nrows=3, dtype=str)

        self.kt_cols = build_col_map(df_kt, _KT_CANDIDATES)
        self.pl_cols = build_col_map(df_pl, _PL_CANDIDATES)
        self.pl_cols["요청번호"] = detect_col(df_pl, _PL_OPT["요청번호"], required=False)

        self.log.info("  KT 컬럼:")
        for k, v in self.kt_cols.items():
            self.log.info(f"    [{'O' if v else '-'}] {k:<16}: {v or '(미탐지)'}")
        self.log.info("  플랫폼 컬럼:")
        for k, v in self.pl_cols.items():
            self.log.info(f"    [{'O' if v else '-'}] {k:<16}: {v or '(없음)'}")

    # ── 정산 기간: 생성자 파라미터 그대로 사용 ────────────────────
    def get_period(self):
        self.log.info(f"  정산 기간: {self._start} ~ {self._end}")
        return self._start, self._end


# ════════════════════════════════════════════════════════
# GUI 애플리케이션
# ════════════════════════════════════════════════════════
_SETTINGS = str(_HERE / "gui_settings.json")


class SettlementApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("KT 정산 자동화 시스템  v2.0")
        self.root.geometry("760x660")
        self.root.resizable(True, True)
        self._result_path: str | None = None

        self._apply_style()
        self._build_ui()
        self._load_settings()
        self._auto_date()

    # ── 스타일 ───────────────────────────────────────────────────
    def _apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")   # Windows 기본 테마
        except Exception:
            style.theme_use("clam")
        style.configure("TLabelframe.Label", font=("맑은 고딕", 9, "bold"))
        style.configure("Run.TButton",       font=("맑은 고딕", 10, "bold"))

    # ── UI 구성 ──────────────────────────────────────────────────
    def _build_ui(self):
        wrap = ttk.Frame(self.root, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        # 제목
        ttk.Label(
            wrap, text="KT 정산 자동화 시스템",
            font=("맑은 고딕", 16, "bold"), foreground="#1F4E79",
        ).pack(anchor="w", pady=(0, 10))

        # ── 파일 선택 ──────────────────────────────────────────
        ff = ttk.LabelFrame(wrap, text=" 파일 선택 ", padding=(10, 6))
        ff.pack(fill=tk.X, pady=(0, 8))
        ff.columnconfigure(1, weight=1)

        self.v_kt = tk.StringVar()
        self.v_pl = tk.StringVar()
        self.v_mm = tk.StringVar()
        self.v_rm = tk.StringVar()

        file_rows = [
            ("KT 정산파일  ★",       self.v_kt, "KT 정산 raw 파일 (필수)"),
            ("통합플랫폼파일  ★",    self.v_pl, "플랫폼 입고내역 파일 (필수)"),
            ("소싱그룹 매핑파일",    self.v_mm, "서비스카테고리→관리회계/담당자 매핑"),
            ("반품 매핑파일",        self.v_rm, "return_mapping.xlsx (반품 경로 B 시)"),
        ]
        for i, (lbl, var, tip) in enumerate(file_rows):
            ttk.Label(ff, text=lbl, width=20, anchor="e").grid(
                row=i, column=0, padx=(0, 8), pady=4, sticky="e")
            ent = ttk.Entry(ff, textvariable=var)
            ent.grid(row=i, column=1, padx=(0, 6), pady=4, sticky="ew")
            ent.bind("<FocusOut>", lambda e: self._save_settings())
            ttk.Button(
                ff, text="찾아보기", width=9,
                command=lambda v=var: self._browse(v),
            ).grid(row=i, column=2, pady=4)
            # 툴팁 (레이블 오른쪽에 회색 힌트)
            ttk.Label(ff, text=f"  ← {tip}", foreground="#888888",
                      font=("맑은 고딕", 8)).grid(
                row=i, column=3, sticky="w", padx=(0, 4))

        # ── 정산 기간 ──────────────────────────────────────────
        df = ttk.LabelFrame(wrap, text=" 정산 기간 ", padding=(10, 6))
        df.pack(fill=tk.X, pady=(0, 8))
        di = ttk.Frame(df); di.pack(fill=tk.X)

        ttk.Label(di, text="시작일").pack(side=tk.LEFT)
        self.v_s = tk.StringVar()
        ttk.Entry(di, textvariable=self.v_s, width=13,
                  font=("Consolas", 10)).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(di, text="  ~  종료일").pack(side=tk.LEFT)
        self.v_e = tk.StringVar()
        ttk.Entry(di, textvariable=self.v_e, width=13,
                  font=("Consolas", 10)).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(di, text="자동 판단",
                   command=self._auto_date).pack(side=tk.LEFT, padx=(14, 0))
        ttk.Label(di, text="  YYYY-MM-DD",
                  foreground="#888888", font=("맑은 고딕", 8)).pack(side=tk.LEFT)

        # 차수 선택
        di2 = ttk.Frame(df); di2.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(di2, text="차수 선택").pack(side=tk.LEFT)
        self.v_round = tk.StringVar(value="자동")
        ttk.Combobox(
            di2, textvariable=self.v_round,
            values=["자동", "1차", "2차"],
            state="readonly", width=6,
            font=("맑은 고딕", 10),
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(di2, text="  ※ 자동: 시작일 기준 판단 (1일~15일→1차, 16일~→2차)",
                  foreground="#888888", font=("맑은 고딕", 8)).pack(side=tk.LEFT)

        # ── 실행 버튼 / 상태 ───────────────────────────────────
        bf = ttk.Frame(wrap); bf.pack(fill=tk.X, pady=(4, 10))

        self.btn_run = ttk.Button(
            bf, text="  정산 실행  ", style="Run.TButton",
            command=self._on_run,
        )
        self.btn_run.pack(side=tk.LEFT, ipadx=20, ipady=5)

        self.btn_open = ttk.Button(
            bf, text="결과 파일 열기",
            command=self._open_result, state="disabled",
        )
        self.btn_open.pack(side=tk.LEFT, padx=(12, 0), ipadx=10, ipady=5)

        self.v_status = tk.StringVar(value="준비")
        ttk.Label(
            bf, textvariable=self.v_status,
            foreground="#888888", font=("맑은 고딕", 9),
        ).pack(side=tk.RIGHT, padx=6)

        # ── 로그 창 ────────────────────────────────────────────
        lf = ttk.LabelFrame(wrap, text=" 실행 로그 ", padding=4)
        lf.pack(fill=tk.BOTH, expand=True)

        self.log_w = scrolledtext.ScrolledText(
            lf, state="disabled", height=16,
            font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white", relief="flat",
            selectbackground="#264f78",
        )
        self.log_w.pack(fill=tk.BOTH, expand=True)

        # 로그 색상 태그
        self.log_w.tag_config("info", foreground="#9cdcfe")
        self.log_w.tag_config("warn", foreground="#ffd700")
        self.log_w.tag_config("err",  foreground="#f44747")
        self.log_w.tag_config("gray", foreground="#666666")
        self.log_w.tag_config("ok",   foreground="#6ab187")

    # ── 파일 찾아보기 ─────────────────────────────────────────────
    def _browse(self, var: tk.StringVar):
        init = str(Path(var.get()).parent) if var.get() else str(_HERE)
        p = filedialog.askopenfilename(
            initialdir=init,
            filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if p:
            var.set(p)
            self._save_settings()

    # ── 날짜 자동 판단 ────────────────────────────────────────────
    def _auto_date(self):
        t = datetime.date.today()
        d, y, m = t.day, t.year, t.month
        if d == 16:
            s = datetime.date(y, m, 1)
            e = datetime.date(y, m, 15)
        elif d == 1:
            py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
            s = datetime.date(py, pm, 16)
            e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
        elif d <= 15:
            # 상반월 중: 전월 하반월 기간
            py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
            s = datetime.date(py, pm, 16)
            e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
        else:
            # 하반월 중: 당월 상반월 기간
            s = datetime.date(y, m, 1)
            e = datetime.date(y, m, 15)
        self.v_s.set(s.isoformat())
        self.v_e.set(e.isoformat())

    # ── 설정 저장 / 로드 ──────────────────────────────────────────
    def _save_settings(self):
        try:
            with open(_SETTINGS, "w", encoding="utf-8") as f:
                json.dump({
                    "kt": self.v_kt.get(),
                    "pl": self.v_pl.get(),
                    "mm": self.v_mm.get(),
                    "rm": self.v_rm.get(),
                }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            if os.path.exists(_SETTINGS):
                with open(_SETTINGS, encoding="utf-8") as f:
                    s = json.load(f)
                self.v_kt.set(s.get("kt", ""))
                self.v_pl.set(s.get("pl", ""))
                self.v_mm.set(s.get("mm", ""))
                self.v_rm.set(s.get("rm", ""))
        except Exception:
            pass

    # ── 로그 출력 ─────────────────────────────────────────────────
    def _print_log(self, msg: str, tag: str = "info"):
        def _a():
            self.log_w.configure(state="normal")
            self.log_w.insert("end", msg + "\n", tag)
            self.log_w.configure(state="disabled")
            self.log_w.see("end")
        self.root.after(0, _a)

    # ── 실행 ─────────────────────────────────────────────────────
    def _on_run(self):
        kt = self.v_kt.get().strip()
        pl = self.v_pl.get().strip()

        if not kt or not os.path.exists(kt):
            messagebox.showerror("파일 오류", "KT 정산파일을 선택해주세요.")
            return
        if not pl or not os.path.exists(pl):
            messagebox.showerror("파일 오류", "통합플랫폼파일을 선택해주세요.")
            return

        try:
            s_date = datetime.date.fromisoformat(self.v_s.get().strip())
            e_date = datetime.date.fromisoformat(self.v_e.get().strip())
        except ValueError:
            messagebox.showerror("날짜 오류", "날짜를 YYYY-MM-DD 형식으로 입력해주세요.\n예) 2025-05-01")
            return
        if s_date > e_date:
            messagebox.showerror("날짜 오류", "시작일이 종료일보다 늦습니다.")
            return

        mm = self.v_mm.get().strip() or ""
        rm = self.v_rm.get().strip() or ""

        # UI 초기화
        self.btn_run.config(state="disabled")
        self.btn_open.config(state="disabled")
        self._result_path = None
        self._save_settings()
        self.root.after(0, lambda: self.v_status.set("실행 중..."))

        self.log_w.configure(state="normal")
        self.log_w.delete("1.0", "end")
        self.log_w.configure(state="disabled")

        def _worker():
            try:
                _round_sel = self.v_round.get()
                _round_ov  = None if _round_sel == "자동" else _round_sel
                runner = _GUIRunner(kt, pl, mm, rm, s_date, e_date,
                                    round_override=_round_ov)

                # 기존 TkLogHandler 제거 (재실행 시 중복 방지)
                for h in list(runner.log.handlers):
                    if isinstance(h, _TkLogHandler):
                        runner.log.removeHandler(h)

                # GUI 로그 핸들러 연결
                hdl = _TkLogHandler(self.log_w)
                hdl.setFormatter(logging.Formatter(
                    "%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S",
                ))
                runner.log.addHandler(hdl)

                runner.run()

                # 결과 파일 경로: runner.result_path 에 저장된 실제 경로 사용
                candidate = getattr(runner, "result_path", None)
                self._result_path = candidate if candidate and os.path.exists(candidate) else None

                self.root.after(0, self._on_done)

            except ReturnMappingNeeded as exc:
                # 반품 매핑 템플릿 생성 완료 — 오류가 아닌 안내 메시지로 표시
                self._print_log(f"\n📋 {exc}", "warn")
                self.root.after(0, lambda: self.v_status.set(
                    "반품 매핑 파일 작성 후 재실행 필요"
                ))
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

            except Exception as exc:
                self._print_log(f"\n오류 발생: {exc}", "err")
                self._print_log(traceback.format_exc(), "err")
                self.root.after(0, lambda: self.v_status.set("오류 발생 — 로그 확인"))
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self):
        self.btn_run.config(state="normal")
        if self._result_path and os.path.exists(self._result_path):
            self.v_status.set(f"완료  →  {Path(self._result_path).name}")
            self.btn_open.config(state="normal")
            self._print_log(f"\n완료: {self._result_path}", "ok")
        else:
            self.v_status.set("완료 (결과 파일 경로 확인)")
            self._print_log("실행 완료 — 결과 파일을 직접 확인하세요.", "warn")

    def _open_result(self):
        if self._result_path and os.path.exists(self._result_path):
            os.startfile(self._result_path)


# ════════════════════════════════════════════════════════
# 진입점
# ════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = SettlementApp(root)
    root.mainloop()
