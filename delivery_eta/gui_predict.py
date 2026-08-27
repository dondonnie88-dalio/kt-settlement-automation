# -*- coding: utf-8 -*-
"""
배송 예측 조회 GUI - 상품코드 + 주문일자만 입력하면 도착확률을 바로 보여준다.
serve.py의 DeliveryPredictor를 그대로 사용(중복 로직 없음). tkinter는 파이썬 표준
라이브러리라 별도 설치 없이 바로 실행 가능(사내망 pip 설치 제약과 무관).

실행: python gui_predict.py (또는 gui_predict.bat 더블클릭)
"""
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from serve import DeliveryPredictor

HERE = Path(__file__).resolve().parent

# 주문 마감시간 가정 (2026-07-07): 이 시각 이후 조회하면 "지금 주문해도 사실상 내일 접수"로
# 보고 주문일을 다음날로 간주해서 계산한다. 확률 모델 자체는 건드리지 않고(원본 데이터에
# 시:분 정보가 아예 없어서 마감 효과를 검증할 방법이 없음 - known_risks.md 참고) 주문일
# 해석만 바꾸는 방식. 15시는 네이버 "15:00까지 결제 시 오늘 발송" 패턴을 참고한 가정값이며
# 실제 KT커머스/통합플랫폼 주문 처리 마감시간이 확인되면 이 값을 바꿀 것(API_SPEC.md 6번).
# 발주 접수는 평일만 함(2026-07-15 사용자 확정: 토요일 신규 발주 접수 안 함. 예외
# 협력사가 있어도 개별 확인이 현실적으로 불가능해 "접수 안 함"을 기본 가정으로 유지
# - known_risks.md 참고). 토요일도 배송완료 자체는 9.8%로 실제 발생하지만(일요일
# 0.43%와 다름) 이건 "금요일까지 접수된 건이 토요일에 도착만 한 것"이라 주말(토/일)
# 전부 접수 안 함으로 처리하는 게 맞음 - 확정 사항, 재확인 불필요.
CUTOFF_HOUR = 15

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]
LEVEL_LABEL = {
    "sku": "상품코드 자체 데이터",
    "vendor_mid": "협력사+상품중분류 평균",
    "vendor": "협력사 전체 평균",
    "sla_bucket": "표준납기일 기반 평균 (추정치 사용)",
    "cat": "상품대분류 전체 평균",
    "global": "전체 평균 (신규/미등록 상품)",
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("배송 예측 조회")
        self.geometry("800x560")
        self.resizable(False, False)

        self.predictor = None
        self._load_predictor()

        self._build_input_row()
        self._build_result_area()

    def _load_predictor(self):
        try:
            self.predictor = DeliveryPredictor()
        except Exception as e:
            messagebox.showerror("초기화 실패", f"segment_distributions.csv 등 파일을 불러오지 못했습니다.\n\n{e}")
            self.destroy()
            return
        # 표준납기일 자동 조회용 (2026-07-23 추가): 매번 손으로 입력할 수 없으니, 학습 데이터에서
        # 뽑아둔 상품코드별/협력사별 표준납기_lt 최빈값을 자동으로 붙여서 sla_bucket 레벨까지
        # 활용한다 - 어디까지나 추정치이며 통합플랫폼의 실제 값과는 다를 수 있음(화면에 명시).
        try:
            vdf = pd.read_csv(HERE / "vendor_sla_lookup.csv")
            self.vendor_sla_lookup = dict(zip(vdf["협력사명"], vdf["표준납기_lt_추정"]))
        except FileNotFoundError:
            self.vendor_sla_lookup = {}

    def _estimate_sla_lt(self, code: str):
        """상품코드 -> 표준납기_lt 추정치(일수). 상품코드 자체 값 우선, 없으면 협력사 최빈값,
        둘 다 없으면 None(표준납기일 없이 조회 - 기존 cat/global 폴백과 동일)."""
        pl = self.predictor.product_lookup
        if code not in pl.index:
            return None
        row = pl.loc[code]
        val = row.get("표준납기_lt_추정")
        if pd.notna(val):
            return int(val)
        vendor_name = row.get("협력사명")
        val = self.vendor_sla_lookup.get(vendor_name)
        return int(val) if pd.notna(val) else None

    def _build_input_row(self):
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="x")

        ttk.Label(frame, text="상품코드", font=("맑은 고딕", 11)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.code_entry = ttk.Entry(frame, font=("맑은 고딕", 11), width=18)
        self.code_entry.grid(row=0, column=1, padx=(0, 20))
        self.code_entry.bind("<Return>", lambda e: self.on_search())

        ttk.Label(frame, text="주문일자 (YYYY-MM-DD)", font=("맑은 고딕", 11)).grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.date_entry = ttk.Entry(frame, font=("맑은 고딕", 11), width=14)
        self.date_entry.insert(0, date.today().isoformat())
        self.date_entry.grid(row=0, column=3, padx=(0, 20))
        self.date_entry.bind("<Return>", lambda e: self.on_search())

        self.search_btn = ttk.Button(frame, text="조회", command=self.on_search)
        self.search_btn.grid(row=0, column=4)

        self.status_label = ttk.Label(self, text="상품코드를 입력하고 조회를 누르세요.",
                                       font=("맑은 고딕", 9), foreground="#6E7B8B", padding=(12, 0))
        self.status_label.pack(fill="x")

    def _build_result_area(self):
        info_frame = ttk.LabelFrame(self, text="상품 정보", padding=10)
        info_frame.pack(fill="x", padx=12, pady=(6, 6))

        self.info_var = tk.StringVar(value="-")
        ttk.Label(info_frame, textvariable=self.info_var, font=("맑은 고딕", 10), justify="left",
                  wraplength=760).pack(anchor="w")

        headline_frame = ttk.LabelFrame(self, text="핵심 결과 (90% 확신 도착일)", padding=10)
        headline_frame.pack(fill="x", padx=12, pady=(0, 6))

        self.headline_var = tk.StringVar(value="-")
        ttk.Label(headline_frame, textvariable=self.headline_var, font=("맑은 고딕", 15, "bold"),
                  foreground="#1E2761", wraplength=740, justify="left").pack(anchor="w")

        table_frame = ttk.LabelFrame(self, text="날짜별 누적 도착확률", padding=10)
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        cols = ("day", "date", "weekday", "prob")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=9)
        headers = {"day": "경과일", "date": "날짜", "weekday": "요일", "prob": "누적 도착확률"}
        widths = {"day": 70, "date": 110, "weekday": 60, "prob": 140}
        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths[c], anchor="center")
        # 배송 느린 상품은 45/60/90일 지점이 추가로 붙어 9행을 넘을 수 있어(serve.py 참고)
        # 스크롤바를 달아둔다 - 평소엔 안 보이다가 필요할 때만 스크롤하면 됨.
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, side="left")

    def on_search(self):
        code = self.code_entry.get().strip()
        date_str = self.date_entry.get().strip()
        if not code:
            messagebox.showwarning("입력 필요", "상품코드를 입력하세요.")
            return
        try:
            order_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showwarning("날짜 형식 오류", "주문일자는 YYYY-MM-DD 형식으로 입력하세요. 예: 2026-07-08")
            return

        # 주문일자가 오늘(기본값)일 때만 마감시간 로직 적용 - 다른 날짜를 직접 입력한
        # 경우는 시뮬레이션 목적이므로 입력값 그대로 계산한다.
        cutoff_note = ""
        if order_date == date.today():
            reason = None
            if datetime.now().hour >= CUTOFF_HOUR:
                order_date = order_date + timedelta(days=1)
                reason = f"{CUTOFF_HOUR}시 마감 경과"
            # 발주 접수는 평일만 한다고 가정(2026-07-08 사용자 확인) - 마감시간 계산 결과가
            # 토/일이면 다음 월요일로 건너뛴다. 오늘 자체가 토/일인 경우(마감 전이라도)도
            # 동일하게 적용 - 주말엔 아예 신규 접수를 안 한다는 가정이므로.
            if order_date.weekday() >= 5:
                order_date = order_date + timedelta(days=7 - order_date.weekday())
                reason = (reason + " + 주말 제외" if reason else "주말 제외")
            if reason:
                cutoff_note = f"{reason} - {order_date.month}/{order_date.day}({WEEKDAY_KO[order_date.weekday()]}) 주문 접수 기준"
            else:
                cutoff_note = f"오늘 {CUTOFF_HOUR}시 이전 주문 기준"

        sla_lt = self._estimate_sla_lt(code)
        표준납기일 = order_date + timedelta(days=sla_lt) if sla_lt is not None else None

        try:
            result = self.predictor.predict(order_date=order_date, product_code=code, 표준납기일=표준납기일)
        except Exception as e:
            messagebox.showerror("조회 실패", str(e))
            return

        self._show_result(code, result, cutoff_note, sla_estimate=(sla_lt, 표준납기일))

    def _show_result(self, code, result, cutoff_note="", sla_estimate=(None, None)):
        if not result["product_recognized"]:
            # 상품코드 자체가 학습 데이터 어디에도 없는 경우(오타 등) - 신규 상품인지
            # 아예 존재하지 않는 코드인지 구분할 방법이 없어서(품목 마스터가 아니라
            # 주문 이력 기반), 확신에 찬 확률 대신 확인 안내로 대체한다(2026-07-06 사용자 지적).
            self.info_var.set(
                f"상품코드: {code}      상품명: (확인 불가)\n"
                f"⚠ 이 상품코드를 찾을 수 없습니다. 오타이거나, 아직 주문 이력이 없는\n"
                f"  신규 상품일 수 있습니다. 상품코드를 다시 확인해 주세요."
            )
            self.headline_var.set("확인 필요 - 상품코드를 다시 입력해 주세요")
            for row in self.tree.get_children():
                self.tree.delete(row)
            self.status_label.config(text=f"조회 완료 - {code} (상품코드 미확인)")
            return

        name = result["product_name"] or "(상품명 미등록)"
        vendor = result["resolved_vendor"] or "-"
        category = result["resolved_category"] or "-"
        level = result["level"]
        level_label = LEVEL_LABEL.get(level, level)

        sla_lt, sla_date = sla_estimate
        sla_line = ""
        if sla_lt is not None:
            sla_line = (f"\n표준납기일(추정): {sla_date.month}/{sla_date.day} "
                        f"({sla_lt}일, 학습 이력 기준 추정치 - 실제 값과 다를 수 있음)")

        self.info_var.set(
            f"상품코드: {code}      상품명: {name}\n"
            f"협력사: {vendor}      카테고리: {category}\n"
            f"예측 근거: {level_label}  (학습 건수 {result['sample_size']:,}건, 중앙값 {result['median_days']}일)"
            + ("      ※ 공휴일 연휴 직전이라 단기 확률이 보정됨" if result["holiday_adjusted"] else "")
            + sla_line
        )

        h = result["headline"]
        target = datetime.strptime(h["date"], "%Y-%m-%d").date()
        prefix = f"[{cutoff_note}]  " if cutoff_note else ""
        if h["threshold_met"]:
            self.headline_var.set(f"{prefix}{target.month}/{target.day}({h['weekday']})까지 도착 확률 {h['prob']*100:.1f}%  "
                                   f"({h['day']}일 이내)")
        else:
            self.headline_var.set(f"{prefix}{h['day']}일이 지나도 {int(h['threshold']*100)}%를 넘지 못함 "
                                   f"(그때까지 확률 {h['prob']*100:.1f}%) - 배송이 불안정한 상품일 수 있음")

        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in result["by_day"]:
            self.tree.insert("", "end", values=(f"{row['day']}일", row["date"], row["weekday"] + "요일",
                                                 f"{row['prob']*100:.1f}%"))

        self.status_label.config(text=f"조회 완료 - {code}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
