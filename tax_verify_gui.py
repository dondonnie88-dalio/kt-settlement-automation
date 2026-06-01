"""
tax_verify_gui.py  ─  과세구분 검증 GUI 실행기
================================================
더블클릭 또는 [과세구분검증.bat] 으로 실행합니다.
run_tax_verify.py 와 같은 폴더에 있어야 합니다.
"""

import importlib, importlib.util, subprocess, sys

def _try_install(pkg):
    if importlib.util.find_spec(pkg) is not None:
        return
    print(f"[설치 중] {pkg} ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"], timeout=30)
    except Exception as e:
        print(f"[경고] 설치 실패: {e}")

for _pkg in ["pandas", "openpyxl"]:
    _try_install(_pkg)

import os, json, threading, logging, traceback, datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

try:
    from run_tax_verify import validate_tax, write_tax_sheet, load_svccat, load_whitelist, auto_period
except ImportError as _ie:
    try:
        root = tk.Tk(); root.withdraw()
        from tkinter import messagebox
        messagebox.showerror(
            "임포트 오류",
            f"run_tax_verify.py 를 찾을 수 없습니다.\n"
            f"같은 폴더에 있는지 확인하세요.\n\n{_ie}"
        )
    except Exception:
        print(f"[오류] run_tax_verify.py 임포트 실패: {_ie}")
    sys.exit(1)

import openpyxl


# ── 로그 → Text 위젯 실시간 스트리밍 ──────────────────────────────
class _TkLogHandler(logging.Handler):
    _TAG = {"DEBUG": "gray", "INFO": "info", "WARNING": "warn",
            "ERROR": "err", "CRITICAL": "err"}

    def __init__(self, widget):
        super().__init__()
        self._w = widget

    def emit(self, record):
        msg = self.format(record)
        tag = self._TAG.get(record.levelname, "info")
        def _a():
            self._w.configure(state="normal")
            self._w.insert("end", msg + "\n", tag)
            self._w.configure(state="disabled")
            self._w.see("end")
        self._w.after(0, _a)


# ── 설정 파일 경로 ────────────────────────────────────────────────
_SETTINGS = str(_HERE / "tax_verify_settings.json")


# ── 메인 앱 ──────────────────────────────────────────────────────
class TaxVerifyApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("과세구분 검증 도구")
        self.root.geometry("680x620")
        self.root.resizable(True, True)
        self._result_path: str | None = None

        self._apply_style()
        self._build_ui()
        self._load_settings()
        self._auto_date()

    def _apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            style.theme_use("clam")
        style.configure("TLabelframe.Label", font=("맑은 고딕", 9, "bold"))
        style.configure("Run.TButton",       font=("맑은 고딕", 10, "bold"))

    def _build_ui(self):
        wrap = ttk.Frame(self.root, padding=14)
        wrap.pack(fill=tk.BOTH, expand=True)

        # 제목
        ttk.Label(
            wrap, text="과세구분 검증 도구",
            font=("맑은 고딕", 16, "bold"), foreground="#1F4E79",
        ).pack(anchor="w", pady=(0, 10))

        # ── 파일 선택 ──────────────────────────────────────────────
        ff = ttk.LabelFrame(wrap, text=" 파일 선택 ", padding=(10, 8))
        ff.pack(fill=tk.X, pady=(0, 8))
        ff.columnconfigure(1, weight=1)

        self.v_pl = tk.StringVar()
        self.v_mm = tk.StringVar()
        self.v_wl = tk.StringVar()

        # 플랫폼 파일 변경 시 같은 폴더의 과세확인목록.xlsx 자동 탐지
        self.v_pl.trace_add("write", self._auto_detect_whitelist)

        rows = [
            ("통합플랫폼 파일  ★", self.v_pl, "과세구분 검증할 플랫폼 Excel (필수)"),
            ("매핑마스터 파일",    self.v_mm, "서비스카테고리→담당자 조회용 (선택)"),
            ("과세확인목록",       self.v_wl, "확인 완료 상품코드 목록 (선택 · 자동탐지)"),
        ]
        for i, (lbl, var, tip) in enumerate(rows):
            ttk.Label(ff, text=lbl, width=20, anchor="e").grid(
                row=i, column=0, padx=(0, 8), pady=5, sticky="e")
            ent = ttk.Entry(ff, textvariable=var)
            ent.grid(row=i, column=1, padx=(0, 6), pady=5, sticky="ew")
            ent.bind("<FocusOut>", lambda e: self._save_settings())
            ttk.Button(
                ff, text="찾아보기", width=9,
                command=lambda v=var: self._browse(v),
            ).grid(row=i, column=2, pady=5)
            ttk.Label(ff, text=f"  ← {tip}", foreground="#888888",
                      font=("맑은 고딕", 8)).grid(
                row=i, column=3, sticky="w", padx=(0, 4))

        # ── 검증 기간 ──────────────────────────────────────────────
        df = ttk.LabelFrame(wrap, text=" 검증 기간 (승인일 기준) ", padding=(10, 6))
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
        ttk.Label(di, text="  YYYY-MM-DD  (비우면 전체 기간)",
                  foreground="#888888", font=("맑은 고딕", 8)).pack(side=tk.LEFT)

        # ── 실행 버튼 / 상태 ───────────────────────────────────────
        bf = ttk.Frame(wrap)
        bf.pack(fill=tk.X, pady=(4, 10))

        self.btn_run = ttk.Button(
            bf, text="  검증 실행  ", style="Run.TButton",
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

        # ── 로그 창 ────────────────────────────────────────────────
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
        self.log_w.tag_config("info", foreground="#9cdcfe")
        self.log_w.tag_config("warn", foreground="#ffd700")
        self.log_w.tag_config("err",  foreground="#f44747")
        self.log_w.tag_config("gray", foreground="#666666")
        self.log_w.tag_config("ok",   foreground="#6ab187")

    # ── 과세확인목록 자동 탐지 ───────────────────────────────────
    def _auto_detect_whitelist(self, *_):
        pl = self.v_pl.get().strip()
        if not pl:
            return
        candidate = Path(pl).parent / "과세확인목록.xlsx"
        if candidate.exists() and not self.v_wl.get().strip():
            self.v_wl.set(str(candidate))

    # ── 날짜 자동 판단 ────────────────────────────────────────────
    def _auto_date(self):
        start, end = auto_period()
        self.v_s.set(start.isoformat())
        self.v_e.set(end.isoformat())

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

    # ── 설정 저장/로드 ────────────────────────────────────────────
    def _save_settings(self):
        try:
            with open(_SETTINGS, "w", encoding="utf-8") as f:
                json.dump({"pl": self.v_pl.get(),
                           "mm": self.v_mm.get(),
                           "wl": self.v_wl.get()},
                          f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            if os.path.exists(_SETTINGS):
                with open(_SETTINGS, encoding="utf-8") as f:
                    s = json.load(f)
                self.v_pl.set(s.get("pl", ""))
                self.v_mm.set(s.get("mm", ""))
                self.v_wl.set(s.get("wl", ""))
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
        pl = self.v_pl.get().strip()
        if not pl or not os.path.exists(pl):
            from tkinter import messagebox
            messagebox.showerror("파일 오류", "통합플랫폼 파일을 선택해 주세요.")
            return

        # 날짜 파싱 (비어 있으면 None → 필터 없음)
        def _parse_date(s):
            s = s.strip()
            if not s:
                return None
            try:
                return datetime.date.fromisoformat(s)
            except ValueError:
                from tkinter import messagebox
                messagebox.showerror("날짜 오류", f"날짜 형식이 올바르지 않습니다: {s!r}\nYYYY-MM-DD 형식으로 입력하세요.")
                return "error"

        start_date = _parse_date(self.v_s.get())
        end_date   = _parse_date(self.v_e.get())
        if start_date == "error" or end_date == "error":
            return

        mm = self.v_mm.get().strip()
        wl = self.v_wl.get().strip()

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
                # 로거 설정
                log = logging.getLogger("tax_verify")
                log.setLevel(logging.INFO)
                # 기존 TkLogHandler 제거 (재실행 중복 방지)
                for h in list(log.handlers):
                    if isinstance(h, _TkLogHandler):
                        log.removeHandler(h)
                hdl = _TkLogHandler(self.log_w)
                hdl.setFormatter(logging.Formatter(
                    "%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S",
                ))
                log.addHandler(hdl)

                # 담당자 마스터 로드 (선택)
                svccat = {}
                if mm and os.path.exists(mm):
                    svccat = load_svccat(mm, log)

                # 과세확인목록 로드 (선택 · 자동탐지 포함)
                whitelist = set()
                wl_path = wl if wl and os.path.exists(wl) else None
                if not wl_path:
                    _auto = Path(pl).parent / "과세확인목록.xlsx"
                    if _auto.exists():
                        wl_path = str(_auto)
                if wl_path:
                    whitelist = load_whitelist(wl_path, log)

                # 과세구분 검증
                tax_issues = validate_tax(pl, svccat, log,
                                          start_date=start_date,
                                          end_date=end_date,
                                          whitelist=whitelist)

                # 결과 파일 저장
                pl_stem  = Path(pl).stem
                out_path = Path(pl).parent / f"과세구분검증_{pl_stem}.xlsx"
                wb = openpyxl.Workbook()
                write_tax_sheet(wb, tax_issues,
                                start_date=start_date, end_date=end_date)
                wb.save(str(out_path))

                self._result_path = str(out_path)
                self.root.after(0, self._on_done)

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
            self.v_status.set("완료")

    def _open_result(self):
        if self._result_path and os.path.exists(self._result_path):
            os.startfile(self._result_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = TaxVerifyApp(root)
    root.mainloop()
