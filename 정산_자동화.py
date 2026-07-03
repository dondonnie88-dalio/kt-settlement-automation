# -*- coding: utf-8 -*-
"""
KT 정산 자동화 스크립트
입력: 매출.xlsx  (시트: 2024년, 2025년)
출력: 정산보고서.xlsx

정산 단위: 협력사 × 정산확정일 (예: 오렌지스펙트럼(주) / 1월 1차)
청구액 = 주문금액  /  정산액 = 정산금액
"""

import sys, os
from pathlib import Path

for pkg in ["pandas", "openpyxl"]:
    try:
        __import__(pkg)
    except ImportError:
        print(f"[설치 필요] pip install {pkg}")
        sys.exit(1)

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

BASE_DIR    = Path(__file__).parent
INPUT_FILE  = BASE_DIR / "매출.xlsx"
OUTPUT_FILE = BASE_DIR / "정산보고서.xlsx"

# ── 오차 허용 범위 (원 단위, 반올림 차이 허용) ────────────────────────────
TOLERANCE = 0

# ════════════════════════════════════════════════════════════════════════════
# 스타일 상수
# ════════════════════════════════════════════════════════════════════════════
FONT = "맑은 고딕"
# 배경색
BG_HEADER   = "1F4E79"   # 진파랑 (헤더)
BG_TITLE    = "2E75B6"   # 중간파랑 (타이틀)
BG_COMPLETE = "C6EFCE"   # 연초록 (정산완료)
BG_UNPAID   = "FFCCCC"   # 연빨강 (미수금)
BG_OVER     = "FFE699"   # 연노랑 (초과입금)
BG_QTY_ERR  = "FCE4D6"   # 연주황 (수량오차)
BG_TOTAL    = "FFF2CC"   # 연노랑 (합계행)
BG_ALT      = "DEEAF1"   # 교번 행

FG_WHITE    = "FFFFFF"
FG_RED      = "C00000"   # 오차 금액 글자색
FG_BLUE     = "1F4E79"   # 합계 글자색

NUM     = "#,##0"
NUM_D   = "#,##0;[Red]-#,##0;\"-\""   # 음수 빨강
PCT     = "0.0%"
DATE_FMT = "YYYY-MM-DD"

def _s(): return Side(style="thin", color="BFBFBF")
def BD(): return Border(left=_s(), right=_s(), top=_s(), bottom=_s())

def hcell(cell, text, bg=BG_HEADER, fg=FG_WHITE, bold=True, wrap=False, size=10):
    cell.value = text
    cell.font  = Font(name=FONT, bold=bold, color=fg, size=size)
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = BD()

def dcell(cell, value, fmt=None, bold=False, bg=None, align="right", color=None):
    cell.value = value
    c = color or "000000"
    cell.font  = Font(name=FONT, bold=bold, color=c, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = BD()
    if fmt:  cell.number_format = fmt
    if bg:   cell.fill = PatternFill("solid", start_color=bg)

def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

def sheet_title(ws, t1, t2=""):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    c = ws.cell(1, 1, t1)
    c.font = Font(name=FONT, bold=True, size=14, color=FG_WHITE)
    c.fill = PatternFill("solid", start_color=BG_TITLE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    if t2:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=20)
        c2 = ws.cell(2, 1, t2)
        c2.font = Font(name=FONT, size=10, color="595959")
        c2.fill = PatternFill("solid", start_color="D6E4F0")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 18

def row_bg(상태):
    return {
        "정산완료":  BG_COMPLETE,
        "미수금":    BG_UNPAID,
        "초과입금":  BG_OVER,
        "수량오차":  BG_QTY_ERR,
    }.get(상태, None)


# ════════════════════════════════════════════════════════════════════════════
# 1. 데이터 로드
# ════════════════════════════════════════════════════════════════════════════
def load_data():
    if not INPUT_FILE.exists():
        sys.exit(f"[오류] 파일 없음: {INPUT_FILE}")

    xl = pd.ExcelFile(INPUT_FILE)
    print(f"[읽기] 시트: {xl.sheet_names}")

    frames = []
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df["시트"] = sheet
        frames.append(df)
        print(f"  [{sheet}] {len(df):,}행")

    data = pd.concat(frames, ignore_index=True)

    # 필수 컬럼 확인
    must = ["협력사명", "정산확정일", "주문금액", "정산금액", "주문수량", "정산수량"]
    missing = [c for c in must if c not in data.columns]
    if missing:
        sys.exit(f"[오류] 필수 컬럼 없음: {missing}")

    # 숫자 변환
    for col in ["주문금액", "정산금액", "주문수량", "정산수량",
                "매출총이익", "매입금액", "주문단가", "정산단가"]:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    # 날짜 변환
    if "일일정산월" in data.columns:
        data["일일정산월"] = pd.to_datetime(data["일일정산월"], errors="coerce")
        data["년월"] = data["일일정산월"].dt.to_period("M").astype(str)
    else:
        data["년월"] = ""

    print(f"\n[전체] {len(data):,}행 | 협력사 {data['협력사명'].nunique()}개 "
          f"| 기간 {data.get('년월', pd.Series()).min()} ~ {data.get('년월', pd.Series()).max()}")
    return data


# ════════════════════════════════════════════════════════════════════════════
# 2. 정산 분석
# ════════════════════════════════════════════════════════════════════════════
def analyze(data):
    df = data.copy()

    # ── 라인별 오차 계산 ────────────────────────────────────────────────────
    df["금액오차"] = df["주문금액"] - df["정산금액"]        # + = 미수금, - = 초과입금
    df["수량오차"] = df["주문수량"] - df["정산수량"]
    df["단가오차"] = df["주문단가"] - df["정산단가"]

    def classify(row):
        if abs(row["금액오차"]) <= TOLERANCE and row["수량오차"] == 0:
            return "정산완료"
        elif row["수량오차"] != 0 and abs(row["금액오차"]) <= TOLERANCE:
            return "수량오차"
        elif row["금액오차"] > TOLERANCE:
            return "미수금"
        else:
            return "초과입금"

    df["정산상태"] = df.apply(classify, axis=1)
    df["오차여부"] = df["정산상태"] != "정산완료"

    # ── 협력사 × 정산확정일 집계 ────────────────────────────────────────────
    grp_key = ["협력사명", "협력사코드", "정산확정일"]
    agg = df.groupby(grp_key, sort=False).agg(
        총주문금액=("주문금액",  "sum"),
        총정산금액=("정산금액",  "sum"),
        총매입금액=("매입금액",  "sum"),
        총이익=    ("매출총이익","sum"),
        총건수=    ("주문금액",  "count"),
        정산완료=  ("오차여부",  lambda x: (~x).sum()),
        오차건수=  ("오차여부",  "sum"),
    ).reset_index()

    agg["총오차금액"] = agg["총주문금액"] - agg["총정산금액"]
    agg["이익률"]    = (agg["총이익"] / agg["총정산금액"].replace(0, pd.NA) * 100).round(2)

    def classify_agg(row):
        if abs(row["총오차금액"]) <= TOLERANCE and row["오차건수"] == 0:
            return "정산완료"
        elif row["총오차금액"] > TOLERANCE:
            return "미수금"
        elif row["총오차금액"] < -TOLERANCE:
            return "초과입금"
        else:
            return "일부오차"

    agg["정산상태"] = agg.apply(classify_agg, axis=1)
    agg.sort_values(["정산확정일", "협력사명"], inplace=True)

    # ── 월별 집계 ────────────────────────────────────────────────────────────
    if "년월" in df.columns and df["년월"].ne("").any():
        monthly_agg = df.groupby("년월").agg(
            총주문금액=("주문금액",  "sum"),
            총정산금액=("정산금액",  "sum"),
            총오차금액=("금액오차",  "sum"),
            총건수=    ("주문금액",  "count"),
            오차건수=  ("오차여부",  "sum"),
        ).reset_index()
    else:
        monthly_agg = pd.DataFrame()

    # ── 협력사별 요약 ────────────────────────────────────────────────────────
    coop_summary = df.groupby("협력사명").agg(
        총주문금액=("주문금액",  "sum"),
        총정산금액=("정산금액",  "sum"),
        총이익=    ("매출총이익","sum"),
        총건수=    ("주문금액",  "count"),
        오차건수=  ("오차여부",  "sum"),
    ).reset_index()
    coop_summary["총오차금액"] = coop_summary["총주문금액"] - coop_summary["총정산금액"]
    coop_summary["이익률"] = (coop_summary["총이익"] / coop_summary["총정산금액"].replace(0, pd.NA) * 100).round(2)
    coop_summary.sort_values("총주문금액", ascending=False, inplace=True)

    # ── 상태 분류별 필터 ────────────────────────────────────────────────────
    unpaid = agg[agg["정산상태"] == "미수금"].copy()
    over   = agg[agg["정산상태"] == "초과입금"].copy()
    done   = agg[agg["정산상태"] == "정산완료"].copy()

    print(f"\n[분석 결과]")
    print(f"  정산완료  : {len(done):>4}건  (합계 {done['총정산금액'].sum():>20,.0f}원)")
    print(f"  미수금    : {len(unpaid):>4}건  (합계 {unpaid['총오차금액'].sum():>20,.0f}원)")
    print(f"  초과입금  : {len(over):>4}건  (합계 {abs(over['총오차금액']).sum():>20,.0f}원)")
    print(f"  전체건수  : {len(agg):>4}건")

    return df, agg, monthly_agg, coop_summary, unpaid, over, done


# ════════════════════════════════════════════════════════════════════════════
# 3. 시트별 작성 함수
# ════════════════════════════════════════════════════════════════════════════

# ── 대시보드 ────────────────────────────────────────────────────────────────
def write_dashboard(wb, agg, coop_summary, unpaid, over, done):
    ws = wb.create_sheet("대시보드", 0)
    sheet_title(ws, "KT 정산 현황 대시보드",
                f"작성일: {datetime.today().strftime('%Y년 %m월 %d일')}")
    ws.sheet_view.showGridLines = False

    total_billed   = agg["총주문금액"].sum()
    total_settled  = agg["총정산금액"].sum()
    total_unpaid   = unpaid["총오차금액"].sum() if len(unpaid) else 0
    total_over     = abs(over["총오차금액"]).sum() if len(over) else 0
    total_profit   = agg["총이익"].sum()
    profit_rate    = total_profit / total_settled * 100 if total_settled else 0
    done_rate      = len(done) / len(agg) * 100 if len(agg) else 0

    # KPI 카드 (2행 × 3열)
    kpis = [
        ("총 청구금액(주문금액)",  total_billed,          NUM,   "1F4E79", "DEEAF1"),
        ("총 정산금액",            total_settled,         NUM,   "375623", "E2EFDA"),
        ("미수금 합계",            total_unpaid,          NUM,   "C00000", "FFCCCC"),
        ("초과입금 합계",          total_over,            NUM,   "7F6000", "FFE699"),
        ("매출총이익",             total_profit,          NUM,   "1F4E79", "DEEAF1"),
        ("정산완료율",             done_rate / 100,       PCT,   "375623", "E2EFDA"),
    ]
    for i, (label, val, fmt, fg, bg_v) in enumerate(kpis):
        col = (i % 3) * 4 + 1
        row_label = 4 + (i // 3) * 3
        row_val   = row_label + 1

        ws.merge_cells(start_row=row_label, start_column=col,
                       end_row=row_label, end_column=col + 2)
        ws.merge_cells(start_row=row_val, start_column=col,
                       end_row=row_val, end_column=col + 2)

        lc = ws.cell(row_label, col, label)
        lc.font  = Font(name=FONT, bold=True, color=FG_WHITE, size=9)
        lc.fill  = PatternFill("solid", start_color=fg)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border = BD()
        for dc in range(col+1, col+3):
            ws.cell(row_label, dc).fill  = PatternFill("solid", start_color=fg)
            ws.cell(row_label, dc).border = BD()

        vc = ws.cell(row_val, col, val)
        vc.number_format = fmt
        vc.font  = Font(name=FONT, bold=True, color=fg, size=15)
        vc.fill  = PatternFill("solid", start_color=bg_v)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = BD()
        for dc in range(col+1, col+3):
            ws.cell(row_val, dc).fill  = PatternFill("solid", start_color=bg_v)
            ws.cell(row_val, dc).border = BD()

        ws.row_dimensions[row_label].height = 18
        ws.row_dimensions[row_val].height   = 38

    # 정산 상태 요약 테이블
    sr = 11
    ws.cell(sr, 1, "정산 상태 요약").font = Font(name=FONT, bold=True, size=11, color=BG_TITLE)
    ws.row_dimensions[sr].height = 22
    for c, h in enumerate(["상태", "건수", "청구금액", "정산금액", "오차금액"], 1):
        hcell(ws.cell(sr+1, c), h)
    status_rows = [
        ("정산완료",  len(done),   done["총주문금액"].sum(),
         done["총정산금액"].sum(),   0),
        ("미수금",    len(unpaid),  unpaid["총주문금액"].sum(),
         unpaid["총정산금액"].sum(), unpaid["총오차금액"].sum()),
        ("초과입금",  len(over),    over["총주문금액"].sum(),
         over["총정산금액"].sum(),   over["총오차금액"].sum()),
    ]
    bg_map = {"정산완료": BG_COMPLETE, "미수금": BG_UNPAID, "초과입금": BG_OVER}
    for r, (status, cnt, billed, settled, diff) in enumerate(status_rows, sr+2):
        bg = bg_map[status]
        dcell(ws.cell(r, 1), status, bold=True, bg=bg, align="center")
        dcell(ws.cell(r, 2), cnt, NUM, bg=bg)
        dcell(ws.cell(r, 3), billed,  NUM, bg=bg)
        dcell(ws.cell(r, 4), settled, NUM, bg=bg)
        color = FG_RED if diff > 0 else ("7F6000" if diff < 0 else None)
        dcell(ws.cell(r, 5), diff, NUM_D, bg=bg, color=color)

    # 협력사별 요약 (상위 10개)
    cr = sr
    col_off = 7
    ws.cell(cr, col_off, "협력사별 정산 현황 (상위 10)").font = Font(
        name=FONT, bold=True, size=11, color=BG_TITLE)
    for c, h in enumerate(["협력사명", "청구금액", "정산금액", "오차금액", "이익률"], col_off):
        hcell(ws.cell(cr+1, c), h, wrap=True)
    for r, row in enumerate(coop_summary.head(10).itertuples(), cr+2):
        bg = BG_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, col_off),   row.협력사명,    align="left", bg=bg)
        dcell(ws.cell(r, col_off+1), row.총주문금액,  NUM, bg=bg)
        dcell(ws.cell(r, col_off+2), row.총정산금액,  NUM, bg=bg)
        diff = row.총오차금액
        dbg = BG_UNPAID if diff > 0 else (BG_OVER if diff < 0 else bg)
        dcell(ws.cell(r, col_off+3), diff, NUM_D, bg=dbg,
              color=FG_RED if diff > 0 else None)
        ir = row.이익률
        dcell(ws.cell(r, col_off+4),
              (ir/100) if ir and not pd.isna(ir) else None, PCT, bg=bg)

    for c in range(1, 14): cw(ws, c, 16)
    cw(ws, col_off, 20)


# ── 청구 vs 정산 대조표 ─────────────────────────────────────────────────────
def write_reconciliation(wb, agg):
    ws = wb.create_sheet("청구 vs 정산 대조표")
    sheet_title(ws, "협력사별 청구액 vs 정산액 대조표",
                "※ 오차 항목(미수금/초과입금)은 색상으로 구분됩니다")

    headers = [
        "정산확정일", "협력사코드", "협력사명", "총 청구금액(원)", "총 정산금액(원)",
        "오차금액(원)", "오차율(%)", "건수", "완료건", "오차건", "정산상태"
    ]
    SR = 4
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, wrap=True)
    ws.row_dimensions[SR].height = 32

    for r, row in enumerate(agg.itertuples(), SR+1):
        bg = row_bg(row.정산상태)

        dcell(ws.cell(r, 1),  row.정산확정일,   align="center", bg=bg)
        dcell(ws.cell(r, 2),  row.협력사코드,   align="center", bg=bg)
        dcell(ws.cell(r, 3),  row.협력사명,     align="left",   bg=bg)
        dcell(ws.cell(r, 4),  row.총주문금액,   NUM,             bg=bg)
        dcell(ws.cell(r, 5),  row.총정산금액,   NUM,             bg=bg)

        diff = row.총오차금액
        diff_color = FG_RED if diff > 0 else ("7F6000" if diff < 0 else None)
        diff_bold  = abs(diff) > 0
        dcell(ws.cell(r, 6),  diff, NUM_D, bold=diff_bold, bg=bg, color=diff_color)

        err_rate = (abs(diff) / row.총주문금액) if row.총주문금액 != 0 else 0
        dcell(ws.cell(r, 7),  err_rate, "0.00%", bg=bg,
              color=FG_RED if err_rate > 0 else None)

        dcell(ws.cell(r, 8),  row.총건수,     NUM, bg=bg)
        dcell(ws.cell(r, 9),  row.정산완료,   NUM, bg=bg)
        dcell(ws.cell(r, 10), row.오차건수,   NUM, bg=bg,
              color=FG_RED if row.오차건수 > 0 else None)

        status_color = {"정산완료": "375623", "미수금": "C00000",
                        "초과입금": "7F6000", "일부오차": "833C00"}.get(row.정산상태, "000000")
        dcell(ws.cell(r, 11), row.정산상태, bold=True, bg=bg,
              align="center", color=status_color)

    # 합계 행
    last = SR + len(agg)
    tot  = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=BG_TOTAL, align="center")
    dcell(ws.cell(tot, 3), f"총 {len(agg)}건", bold=True, bg=BG_TOTAL, align="center")
    for ci, col in [(4, "총주문금액"), (5, "총정산금액"), (6, "총오차금액"),
                    (8, "총건수"), (9, "정산완료"), (10, "오차건수")]:
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = NUM
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=BG_TOTAL)
        ws.cell(tot, ci).border = BD()
        ws.cell(tot, ci).alignment = Alignment(horizontal="right", vertical="center")

    for c, w in [(1,14),(2,12),(3,22),(4,20),(5,20),(6,20),(7,10),
                 (8,8),(9,8),(10,8),(11,12)]:
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"

    # 범례
    lr = tot + 2
    ws.cell(lr, 1, "범례").font = Font(name=FONT, bold=True, size=9)
    for i, (label, bg) in enumerate([
        ("정산완료", BG_COMPLETE), ("미수금", BG_UNPAID),
        ("초과입금", BG_OVER), ("수량오차", BG_QTY_ERR)
    ], 0):
        r_leg = lr + 1 + i
        ws.cell(r_leg, 1).fill   = PatternFill("solid", start_color=bg)
        ws.cell(r_leg, 1).border = BD()
        ws.cell(r_leg, 2, label).font = Font(name=FONT, size=9)


# ── 미수금 현황 ─────────────────────────────────────────────────────────────
def write_unpaid(wb, unpaid, detail_df):
    ws = wb.create_sheet("미수금 현황")
    sheet_title(ws, "미수금 현황",
                f"총 {len(unpaid)}건 | 미수금 합계: {unpaid['총오차금액'].sum():,.0f}원")

    # 집계 테이블
    SR = 4
    headers = ["정산확정일", "협력사명", "청구금액", "정산금액", "미수금액", "미수율(%)", "건수"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, bg="C00000")
    ws.row_dimensions[SR].height = 28

    for r, row in enumerate(unpaid.itertuples(), SR+1):
        bg = BG_UNPAID if r % 2 == 0 else "FFE0E0"
        dcell(ws.cell(r, 1), row.정산확정일,  align="center", bg=bg)
        dcell(ws.cell(r, 2), row.협력사명,    align="left",   bg=bg)
        dcell(ws.cell(r, 3), row.총주문금액,  NUM,             bg=bg)
        dcell(ws.cell(r, 4), row.총정산금액,  NUM,             bg=bg)
        dcell(ws.cell(r, 5), row.총오차금액,  NUM, bold=True,
              bg=bg, color=FG_RED)
        rate = row.총오차금액 / row.총주문금액 if row.총주문금액 != 0 else 0
        dcell(ws.cell(r, 6), rate, "0.00%", bg=bg, color=FG_RED)
        dcell(ws.cell(r, 7), row.총건수, NUM, bg=bg)

    last = SR + len(unpaid); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=BG_TOTAL, align="center")
    for ci in [3, 4, 5, 7]:
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = NUM
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=BG_TOTAL)
        ws.cell(tot, ci).border = BD()
        ws.cell(tot, ci).alignment = Alignment(horizontal="right", vertical="center")

    # 미수금 상세 내역 (라인별)
    detail_row = tot + 3
    unpaid_keys = set(zip(unpaid["협력사명"], unpaid["정산확정일"]))
    detail = detail_df[
        detail_df.apply(lambda r: (r["협력사명"], r["정산확정일"]) in unpaid_keys, axis=1)
    ].copy()

    ws.cell(detail_row, 1, "상세 내역 (라인별 오차)").font = Font(
        name=FONT, bold=True, size=11, color="C00000")
    detail_row += 1
    d_headers = ["정산확정일", "협력사명", "주문번호", "상품명",
                 "주문수량", "정산수량", "주문금액", "정산금액", "금액오차", "정산상태"]
    for c, h in enumerate(d_headers, 1):
        hcell(ws.cell(detail_row, c), h, bg="C00000")

    for r_idx, row in detail.iterrows():
        r = detail_row + 1 + list(detail.index).index(r_idx)
        bg = row_bg(row["정산상태"])
        vals = [
            (row["정산확정일"], "center"),
            (row["협력사명"],   "left"),
            (str(row.get("주문번호", "")), "center"),
            (str(row.get("상품명", ""))[:30], "left"),
            (row["주문수량"], NUM),
            (row["정산수량"], NUM),
            (row["주문금액"], NUM),
            (row["정산금액"], NUM),
        ]
        for ci, (val, fmt_or_align) in enumerate(vals, 1):
            if fmt_or_align in ("left", "center", "right"):
                dcell(ws.cell(r, ci), val, align=fmt_or_align, bg=bg)
            else:
                dcell(ws.cell(r, ci), val, fmt_or_align, bg=bg)
        diff = row["금액오차"]
        dcell(ws.cell(r, 9), diff, NUM_D, bold=True, bg=bg,
              color=FG_RED if diff > 0 else None)
        dcell(ws.cell(r, 10), row["정산상태"], bold=True, bg=bg,
              align="center", color=FG_RED)

    for c, w in [(1,14),(2,22),(3,14),(4,30),(5,10),(6,10),
                 (7,16),(8,16),(9,16),(10,12)]:
        cw(ws, c, w)


# ── 초과입금 현황 ────────────────────────────────────────────────────────────
def write_overpaid(wb, over, detail_df):
    ws = wb.create_sheet("초과입금 현황")
    if len(over) == 0:
        sheet_title(ws, "초과입금 현황")
        ws.cell(4, 1, "초과입금 항목이 없습니다.").font = Font(
            name=FONT, size=12, color="375623")
        return

    sheet_title(ws, "초과입금 현황",
                f"총 {len(over)}건 | 초과입금 합계: {abs(over['총오차금액']).sum():,.0f}원")
    SR = 4
    headers = ["정산확정일", "협력사명", "청구금액", "정산금액", "초과금액", "초과율(%)", "건수"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, bg="7F6000")

    for r, row in enumerate(over.itertuples(), SR+1):
        bg = BG_OVER if r % 2 == 0 else "FFF0B3"
        dcell(ws.cell(r, 1), row.정산확정일, align="center", bg=bg)
        dcell(ws.cell(r, 2), row.협력사명,   align="left",   bg=bg)
        dcell(ws.cell(r, 3), row.총주문금액, NUM, bg=bg)
        dcell(ws.cell(r, 4), row.총정산금액, NUM, bg=bg)
        over_amt = abs(row.총오차금액)
        dcell(ws.cell(r, 5), over_amt, NUM, bold=True, bg=bg, color="7F6000")
        rate = over_amt / row.총주문금액 if row.총주문금액 != 0 else 0
        dcell(ws.cell(r, 6), rate, "0.00%", bg=bg, color="7F6000")
        dcell(ws.cell(r, 7), row.총건수, NUM, bg=bg)

    last = SR + len(over); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=BG_TOTAL, align="center")
    for ci in [3, 4, 5, 7]:
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = NUM
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=BG_TOTAL)
        ws.cell(tot, ci).border = BD()
        ws.cell(tot, ci).alignment = Alignment(horizontal="right", vertical="center")

    for c, w in [(1,14),(2,22),(3,18),(4,18),(5,18),(6,12),(7,8)]:
        cw(ws, c, w)


# ── 정산완료 목록 ────────────────────────────────────────────────────────────
def write_done(wb, done):
    ws = wb.create_sheet("정산완료 목록")
    sheet_title(ws, "정산완료 목록",
                f"총 {len(done)}건 | 정산 완료 금액: {done['총정산금액'].sum():,.0f}원")
    SR = 4
    headers = ["정산확정일", "협력사명", "청구금액", "정산금액", "매출총이익", "이익률(%)", "건수"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, bg="375623")

    for r, row in enumerate(done.itertuples(), SR+1):
        bg = BG_COMPLETE if r % 2 == 0 else "E8F5E9"
        dcell(ws.cell(r, 1), row.정산확정일, align="center", bg=bg)
        dcell(ws.cell(r, 2), row.협력사명,   align="left",   bg=bg)
        dcell(ws.cell(r, 3), row.총주문금액, NUM, bg=bg)
        dcell(ws.cell(r, 4), row.총정산금액, NUM, bg=bg)
        dcell(ws.cell(r, 5), row.총이익,     NUM, bg=bg)
        ir = row.이익률
        dcell(ws.cell(r, 6), (ir/100) if ir and not pd.isna(ir) else None,
              PCT, bg=bg)
        dcell(ws.cell(r, 7), row.총건수, NUM, bg=bg)

    last = SR + len(done); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=BG_TOTAL, align="center")
    for ci in [3, 4, 5, 7]:
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = NUM
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=BG_TOTAL)
        ws.cell(tot, ci).border = BD()
        ws.cell(tot, ci).alignment = Alignment(horizontal="right", vertical="center")

    for c, w in [(1,14),(2,22),(3,18),(4,18),(5,16),(6,10),(7,8)]:
        cw(ws, c, w)


# ── 전체 정산 내역 (원본 + 상태 컬럼) ─────────────────────────────────────
def write_detail(wb, df):
    ws = wb.create_sheet("전체 정산 내역")
    sheet_title(ws, "전체 정산 라인 내역 (오차 항목 색상 표시)")

    show_cols = [
        "정산확정일", "협력사명", "사업장", "부서명", "주문번호",
        "상품명", "주문수량", "정산수량", "수량오차",
        "주문단가", "주문금액", "정산금액", "금액오차",
        "매입금액", "매출총이익", "정산상태"
    ]
    available = [c for c in show_cols if c in df.columns]
    SR = 4

    for c, h in enumerate(available, 1):
        hcell(ws.cell(SR, c), h, wrap=True)
    ws.row_dimensions[SR].height = 30

    num_cols  = {"주문수량","정산수량","수량오차","주문단가",
                 "주문금액","정산금액","금액오차","매입금액","매출총이익"}
    right_col = num_cols

    for r_idx, (_, row) in enumerate(df[available].iterrows()):
        r = SR + 1 + r_idx
        bg = row_bg(row.get("정산상태", ""))

        for c, col in enumerate(available, 1):
            val = row[col]
            if pd.isna(val):
                val = None

            if col in right_col:
                fmt = NUM_D if col in ("금액오차", "수량오차") else NUM
                color = None
                if col == "금액오차" and isinstance(val, (int, float)) and val != 0:
                    color = FG_RED if val > 0 else "7F6000"
                dcell(ws.cell(r, c), val, fmt, bg=bg, color=color)
            elif col == "정산상태":
                sc = {"정산완료": "375623", "미수금": "C00000",
                      "초과입금": "7F6000", "수량오차": "833C00"}.get(str(val), "000000")
                dcell(ws.cell(r, c), val, bold=True, bg=bg,
                      align="center", color=sc)
            else:
                dcell(ws.cell(r, c), str(val) if val is not None else "",
                      align="left", bg=bg)

    col_widths = {
        "정산확정일": 14, "협력사명": 22, "사업장": 16, "부서명": 14,
        "주문번호": 14, "상품명": 30, "주문수량": 10, "정산수량": 10,
        "수량오차": 10, "주문단가": 12, "주문금액": 14, "정산금액": 14,
        "금액오차": 14, "매입금액": 14, "매출총이익": 14, "정산상태": 12
    }
    for c, col in enumerate(available, 1):
        cw(ws, c, col_widths.get(col, 14))

    ws.freeze_panes = f"A{SR+1}"


# ════════════════════════════════════════════════════════════════════════════
# 4. 메인
# ════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  KT 정산 자동화 스크립트")
    print("=" * 60)

    data = load_data()
    df, agg, monthly_agg, coop_summary, unpaid, over, done = analyze(data)

    wb = openpyxl.Workbook()
    del wb["Sheet"]

    write_dashboard(wb, agg, coop_summary, unpaid, over, done)
    write_reconciliation(wb, agg)
    write_unpaid(wb, unpaid, df)
    write_overpaid(wb, over, df)
    write_done(wb, done)
    write_detail(wb, df)

    wb.save(OUTPUT_FILE)

    print(f"\n[저장 완료] {OUTPUT_FILE}")
    print("-" * 60)
    print(f"  청구 vs 정산 대조 : {len(agg)}건")
    print(f"  정산완료          : {len(done)}건")
    print(f"  미수금            : {len(unpaid)}건  "
          f"({unpaid['총오차금액'].sum():,.0f}원)")
    print(f"  초과입금          : {len(over)}건  "
          f"({abs(over['총오차금액']).sum():,.0f}원)")
    print("=" * 60)

    if sys.platform == "win32":
        try:
            os.startfile(OUTPUT_FILE)
        except OSError:
            print(f"  파일 경로: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
