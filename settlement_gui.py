"""
settlement_gui.py  ─  KT 정산 자동화 GUI
==========================================
더블클릭 또는 [정산자동화.bat] 으로 실행합니다.
run_settlement.py 와 같은 폴더에 있어야 합니다.
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

import os, json, threading, logging, traceback, datetime, calendar
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

_HERE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
sys.path.insert(0, str(_HERE))

try:
    from run_settlement import SettlementRunner
except ImportError as _ie:
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror(
            "임포트 오류",
            f"run_settlement.py 를 찾을 수 없습니다.\n"
            f"같은 폴더에 있는지 확인하세요.\n\n{_ie}"
        )
    except Exception:
        print(f"[오류] run_settlement.py 임포트 실패: {_ie}")
    sys.exit(1)


# ── 로그 → Text 위젯 실시간 스트리밍 ─────────────────────────────
class _TkLogHandler(logging.Handler):
    _TAG = {"DEBUG": "gray", "INFO": "info", "WARNING": "warn",
            "ERROR": "err",  "CRITICAL": "err"}

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


_SETTINGS = str(_HERE / "settlement_gui_settings.json")


def _auto_period():
    """오늘 날짜 기준 정산 기간 자동 판단"""
    today = datetime.date.today()
    d, y, m = today.day, today.year, today.month
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    if d == 16:
        # 정산 마감일: 당월 상반월(1~15)
        s = datetime.date(y, m, 1)
        e = datetime.date(y, m, 15)
    elif d == 1:
        # 정산 마감일: 전월 하반월(16~말일)
        s = datetime.date(py, pm, 16)
        e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
    elif d <= 15:
        # 상반월 진행 중 → 전월 하반월이 가장 최근 정산
        s = datetime.date(py, pm, 16)
        e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
    else:
        # 하반월 진행 중 → 당월 상반월이 가장 최근 정산
        s = datetime.date(y, m, 1)
        e = datetime.date(y, m, 15)
    return s, e


# ── 메인 앱 ──────────────────────────────────────────────────────
class SettlementApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("KT 정산 자동화")
        self.root.geometry("760x700")
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

        ttk.Label(
            wrap, text="KT 정산 자동화",
            font=("맑은 고딕", 16, "bold"), foreground="#1F4E79",
        ).pack(anchor="w", pady=(0, 10))

        # ── 작업 폴더 ──────────────────────────────────────────────
        ff = ttk.LabelFrame(wrap, text=" 작업 폴더 ", padding=(10, 8))
        ff.pack(fill=tk.X, pady=(0, 8))
        ff.columnconfigure(1, weight=1)

        self.v_base = tk.StringVar(value=str(_HERE))
        ttk.Label(ff, text="정산 파일 폴더  ★", width=20, anchor="e").grid(
            row=0, column=0, padx=(0, 8), pady=5, sticky="e")
        ent = ttk.Entry(ff, textvariable=self.v_base)
        ent.grid(row=0, column=1, padx=(0, 6), pady=5, sticky="ew")
        ent.bind("<FocusOut>", lambda e: self._save_settings())
        ttk.Button(ff, text="찾아보기", width=9,
                   command=self._browse_folder).grid(row=0, column=2, pady=5)
        ttk.Label(ff, text="  ← KT파일·플랫폼·매핑마스터 등이 있는 폴더",
                  foreground="#888888", font=("맑은 고딕", 8)).grid(
            row=0, column=3, sticky="w", padx=(0, 4))

        # ── 정산 기간 ──────────────────────────────────────────────
        df = ttk.LabelFrame(wrap, text=" 정산 기간 (입고승인일 기준) ", padding=(10, 6))
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

        # ── 실행 버튼 ──────────────────────────────────────────────
        bf = ttk.Frame(wrap)
        bf.pack(fill=tk.X, pady=(4, 10))

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

        # ── 로그 창 ────────────────────────────────────────────────
        lf = ttk.LabelFrame(wrap, text=" 실행 로그 ", padding=4)
        lf.pack(fill=tk.BOTH, expand=True)

        self.log_w = scrolledtext.ScrolledText(
            lf, state="disabled", height=20,
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

    # ── 날짜 자동 판단 ────────────────────────────────────────────
    def _auto_date(self):
        s, e = _auto_period()
        self.v_s.set(s.isoformat())
        self.v_e.set(e.isoformat())

    # ── 폴더 찾아보기 ─────────────────────────────────────────────
    def _browse_folder(self):
        init = self.v_base.get() or str(_HERE)
        p = filedialog.askdirectory(initialdir=init)
        if p:
            self.v_base.set(p)
            self._save_settings()

    # ── 설정 저장/로드 ────────────────────────────────────────────
    def _save_settings(self):
        try:
            with open(_SETTINGS, "w", encoding="utf-8") as f:
                json.dump({"base": self.v_base.get()}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_settings(self):
        try:
            if os.path.exists(_SETTINGS):
                with open(_SETTINGS, encoding="utf-8") as f:
                    s = json.load(f)
                if s.get("base") and os.path.isdir(s["base"]):
                    self.v_base.set(s["base"])
        except Exception:
            pass

    # ── 실행 ─────────────────────────────────────────────────────
    def _on_run(self):
        base = self.v_base.get().strip()
        if not base or not os.path.isdir(base):
            messagebox.showerror("폴더 오류", "작업 폴더를 선택해 주세요.")
            return

        def _parse_date(s, label):
            s = s.strip()
            try:
                return datetime.date.fromisoformat(s)
            except ValueError:
                messagebox.showerror("날짜 오류",
                    f"{label} 형식이 올바르지 않습니다: {s!r}\nYYYY-MM-DD 형식으로 입력하세요.")
                return None

        start_date = _parse_date(self.v_s.get(), "시작일")
        end_date   = _parse_date(self.v_e.get(), "종료일")
        if start_date is None or end_date is None:
            return

        self.btn_run.config(state="disabled")
        self.btn_open.config(state="disabled")
        self._result_path = None
        self._save_settings()
        self.v_status.set("실행 중...")

        self.log_w.configure(state="normal")
        self.log_w.delete("1.0", "end")
        self.log_w.configure(state="disabled")

        def _worker():
            try:
                runner = SettlementRunner(
                    base_dir=base,
                    start_date=start_date,
                    end_date=end_date,
                    gui_mode=True,
                )
                # 로거에 TkLogHandler 연결
                log = runner.log
                log.setLevel(logging.DEBUG)
                for h in list(log.handlers):
                    if isinstance(h, _TkLogHandler):
                        log.removeHandler(h)
                hdl = _TkLogHandler(self.log_w)
                hdl.setFormatter(logging.Formatter(
                    "%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S",
                ))
                log.addHandler(hdl)

                runner.run()
                self._result_path = getattr(runner, "result_path", None)
                self.root.after(0, self._on_done)

            except Exception as exc:
                def _err():
                    self.log_w.configure(state="normal")
                    self.log_w.insert("end", f"\n오류 발생: {exc}\n", "err")
                    self.log_w.insert("end", traceback.format_exc(), "err")
                    self.log_w.configure(state="disabled")
                    self.log_w.see("end")
                    self.v_status.set("오류 발생 — 로그 확인")
                self.root.after(0, _err)
            finally:
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self):
        self.btn_run.config(state="normal")
        if self._result_path and os.path.exists(self._result_path):
            self.v_status.set(f"완료  →  {Path(self._result_path).name}")
            self.btn_open.config(state="normal")
            self.log_w.configure(state="normal")
            self.log_w.insert("end", f"\n완료: {self._result_path}\n", "ok")
            self.log_w.configure(state="disabled")
            self.log_w.see("end")
        else:
            self.v_status.set("완료")

    def _open_result(self):
        if self._result_path and os.path.exists(self._result_path):
            os.startfile(self._result_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = SettlementApp(root)
    root.mainloop()
