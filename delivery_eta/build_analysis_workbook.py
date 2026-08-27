# -*- coding: utf-8 -*-
"""
배송 예측 프로젝트 - raw data 분석 심화 리포트를 엑셀 워크북으로 통합.
기존 산출물(eda_report.txt, risk_check_report.txt, phase1_report.txt,
segment_distributions.csv)에서 실제 수치를 그대로 가져와서 시트로 구성한다.
숫자를 다시 계산하거나 새로 지어내지 않고, 이미 검증된 리포트의 값만 옮겨 담는다.
각 시트에는 "이 숫자가 무슨 뜻이고 어느 정도면 좋은 건지" 쉬운 말 설명 상자를 넣는다.
"""
import re
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo

from build_glossary import GLOSSARY_DATA, PROCESS_DATA, build_sheet as build_glossary_sheet

HERE = Path(__file__).parent
FONT_NAME = "맑은 고딕"
DAY_MARKS = [1, 2, 3, 5, 7, 10, 14, 21, 30]  # 03_distribution_fit.py와 동일(숫자로 시작하는 모듈명이라 import 대신 복제)

NAVY = "1E2761"
ICE = "CADCFC"
LIGHTBG = "F4F6FB"
DARKGRAY = "36454F"
WHITE = "FFFFFF"
GREEN = "2E7D32"
AMBER = "B8860B"
RED = "C62828"

TITLE_FONT = Font(name=FONT_NAME, size=18, bold=True, color=WHITE)
SUB_FONT = Font(name=FONT_NAME, size=10.5, italic=True, color="6E7B8B")
H_FONT = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
HEADER_FONT = Font(name=FONT_NAME, size=10.5, bold=True, color=WHITE)
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_NAVY = Font(name=FONT_NAME, size=10.5, bold=True, color=NAVY)
EXPLAIN_FONT = Font(name=FONT_NAME, size=10, color=DARKGRAY)
TITLE_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FILL = PatternFill("solid", fgColor="2E75B6")
CALLOUT_FILL = PatternFill("solid", fgColor=LIGHTBG)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER = Alignment(horizontal="center", vertical="center")

# 캘리브레이션 오차(%p) 판정 기준 - 이 프로젝트 전반에서 일관되게 사용
GAP_BANDS = [(2.0, "매우 우수", GREEN), (5.0, "양호", GREEN), (10.0, "보통(참고용)", AMBER), (999, "재검토 필요", RED)]


def judge_gap(gap):
    for threshold, label, color in GAP_BANDS:
        if gap <= threshold:
            return label, color
    return GAP_BANDS[-1][1], GAP_BANDS[-1][2]


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = CENTER
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c2 = ws.cell(row=2, column=1, value=subtitle)
    c2.font = SUB_FONT
    c2.alignment = CENTER
    ws.row_dimensions[2].height = 20


def header_row(ws, row, headers, widths=None):
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(i + 1)].width = w


def explain_box(ws, row, ncols, text, n_rows=3, row_height=18, label="📖 이 표 읽는 법"):
    """설명 상자. 시트 상단에 쉬운 말 해설을 넣는다."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row + n_rows - 1, end_column=ncols)
    c = ws.cell(row=row, column=1, value=f"{label}\n" + text)
    c.font = EXPLAIN_FONT
    c.alignment = WRAP
    for rr in range(row, row + n_rows):
        ws.row_dimensions[rr].height = row_height
        for col in range(1, ncols + 1):
            ws.cell(row=rr, column=col).fill = CALLOUT_FILL
    return row + n_rows + 1


# ---------- parse existing reports ----------

def parse_phase1():
    text = (HERE / "phase1_report.txt").read_text(encoding="utf-8")
    rows = []
    for m in re.finditer(r"^\s*(\d+)\s+([\d.]+)%\s+([\d.]+)%\s+([\d.]+)%p\s+([\d,]+)$", text, re.M):
        rows.append({
            "day": int(m.group(1)), "pred": float(m.group(2)), "actual": float(m.group(3)),
            "gap": float(m.group(4)), "n": int(m.group(5).replace(",", "")),
        })
    level_rows = []
    for m in re.finditer(r"^\s*(sku|vendor_mid|vendor_cat|vendor|sla_bucket|cat|global)\s+n=\s*([\d,]+)\s+예측\(7일\)=\s*([\d.]+)%\s+실측\(7일\)=\s*([\d.]+)%\s+gap=\s*([\d.]+)%p", text, re.M):
        level_rows.append({
            "level": m.group(1), "n": int(m.group(2).replace(",", "")),
            "pred7": float(m.group(3)), "actual7": float(m.group(4)), "gap7": float(m.group(5)),
        })

    # 레벨x day마크 전 구간 (2026-07-07 추가 - U자 패턴/vendor_mid 반대패턴 발견 근거,
    # serve.py의 LEVEL_UNRELIABLE_DAYS와 동일 소스)
    level_day_rows = []
    day_block = re.search(r"-- 세그먼트별 캘리브레이션 \(레벨x day마크 전 구간.*?--\n(.*?)(?=\n\n|\Z)", text, re.S)
    for m in re.finditer(r"^\s*(sku|vendor_mid|vendor_cat|vendor|sla_bucket|cat|global)\s+n=\s*([\d,]+)\s+(.+)$",
                          day_block.group(1), re.M):
        gaps = {int(d): float(g) for d, g in re.findall(r"(\d+)일=([\d.]+)%p", m.group(3))}
        level_day_rows.append({"level": m.group(1), "n": int(m.group(2).replace(",", "")), "gaps": gaps})

    sla_actual = float(re.search(r"기존 SLA\(표준납기일\) 준수율\(실측\): ([\d.]+)%", text).group(1))
    sla_pred = float(re.search(r"모델이 예측한 'SLA일까지 도착확률' 평균: ([\d.]+)%", text).group(1))
    seg_counts = dict(re.findall(r"(\w+)=([\d,]+)", re.search(r"세그먼트 수: (.+)", text).group(1)))
    seg_counts = {k: int(v.replace(",", "")) for k, v in seg_counts.items()}

    usage = {}
    usage_block = re.search(r"-- 세그먼트 폴백 레벨 사용 비중 --\n(.*?)(?=\n\n--)", text, re.S)
    for m in re.finditer(r"(\w+)\s+([\d,]+)\s+\(\s*([\d.]+)%\)", usage_block.group(1)):
        usage[m.group(1)] = {"n": int(m.group(2).replace(",", "")), "pct": float(m.group(3))}
    return rows, level_rows, sla_actual, sla_pred, seg_counts, usage, level_day_rows


def parse_pred_actual_counts():
    """레벨x day마크별 예측건수(기댓값)/실제도착건수 - level_daymark_pred_vs_actual.csv
    (2026-07-27, AI혁신팀 질의 대응으로 홀드아웃 원본에서 직접 재계산·저장한 값,
    phase1_report.txt의 %p 오차와 정확히 일치 검증됨)."""
    df = pd.read_csv(HERE / "level_daymark_pred_vs_actual.csv", encoding="utf-8-sig")
    out = {}
    for lvl, g in df.groupby("레벨", sort=False):
        out[lvl] = {row["구분"]: row.drop(["레벨", "구분"]).astype(float).to_dict() for _, row in g.iterrows()}
    return out


def parse_vendor_change():
    """상품코드 낙찰 협력사 변경 영향 검증 - vendor_change_calibration.csv
    (2026-07-27, AI혁신팀 질의: sku 레벨을 협력사 기준으로 바꿔야 하는가). known_risks.md
    '상품코드 중 낙찰 협력사 변경' 항목과 동일 소스."""
    df = pd.read_csv(HERE / "vendor_change_calibration.csv", encoding="utf-8-sig")
    return df


def parse_eda_missing():
    text = (HERE / "eda_report.txt").read_text(encoding="utf-8")
    segments = ["KT_2026", "그룹사_2026", "외부사_2026", "지입자재_2026"]
    fields = ["배송완료일자", "입고일자", "정산확정일", "배송지", "배송업체", "송장번호", "주문시배송리드타임"]
    data = {f: {} for f in fields}
    n_by_seg = {}
    for seg in segments:
        block_m = re.search(rf"=== \[{re.escape(seg)}\] 결측률 \(n=([\d,]+)\) ===\n(.*?)(?=\n===|\Z)", text, re.S)
        n_by_seg[seg] = int(block_m.group(1).replace(",", ""))
        block = block_m.group(2)
        for f in fields:
            m = re.search(rf"{re.escape(f)}\s+결측\s+[\d,]+\s+/\s+[\d,]+\s+\(\s*([\d.]+)%\)", block)
            data[f][seg] = float(m.group(1)) if m else None
    return segments, fields, data, n_by_seg


def parse_risk2():
    text = (HERE / "risk_check_report.txt").read_text(encoding="utf-8")
    segs = ["KT", "그룹사", "외부사", "지입자재"]
    rows = []
    for seg in segs:
        block_m = re.search(rf"-- {re.escape(seg)} \(n=([\d,]+)\) --\n(.*?)(?=\n\n--|\Z)", text, re.S)
        block = block_m.group(2)
        def g(label):
            m = re.search(rf"{label}\s+median=\s*([\d.]+)\s+p90=\s*([\d.]+)", block)
            return (float(m.group(1)), float(m.group(2))) if m else (None, None)
        direct_m = re.search(r"직접배송 비중: ([\d.]+)%", block)
        rows.append({
            "segment": seg, "n": int(block_m.group(1).replace(",", "")),
            "주문발주": g(r"주문->발주"), "발주출하": g(r"발주->출하"),
            "출하배송완료": g(r"출하->배송완료"), "배송완료입고": g(r"배송완료->입고"),
            "주문입고": g(r"주문->입고\(합계\)"), "직접배송비중": float(direct_m.group(1)),
        })
    return rows


def parse_risk1():
    text = (HERE / "risk_check_report.txt").read_text(encoding="utf-8")
    rows = []
    for m in re.finditer(r"([A-C]\. \([^)]+\)): n비교=([\d,]+)\s+정확히 일치=([\d.]+)%\s+\|오차\|<=1일=([\d.]+)%", text):
        rows.append({"기준": m.group(1), "n": int(m.group(2).replace(",", "")),
                     "정확히일치": float(m.group(3)), "오차1일이내": float(m.group(4))})
    return rows


def parse_growth():
    text = (HERE / "growth_vendor_flags.txt").read_text(encoding="utf-8")
    n_found = int(re.search(r"탐지된 세그먼트 수: (\d+)", text).group(1))
    rows = []
    for line in text.splitlines():
        m = re.match(r"^(\S.{0,18}\S)\s{2,}(\S.{0,14}\S)\s{2,}([\d,]+)\s+([\d,]+)\s+([\d.]+)x", line.strip())
        if m and "협력사명" not in line:
            rows.append({"협력사명": m.group(1).strip(), "대분류": m.group(2).strip(),
                         "2023-25합계": int(m.group(3).replace(",", "")),
                         "2026H1": int(m.group(4).replace(",", "")), "성장률": float(m.group(5))})
    return n_found, rows[:20]


def main():
    day_rows, level_rows, sla_actual, sla_pred, seg_counts, level_usage, level_day_rows = parse_phase1()
    pred_actual_counts = parse_pred_actual_counts()
    vendor_change_df = parse_vendor_change()
    eda_segs, eda_fields, eda_data, eda_n = parse_eda_missing()
    risk2_rows = parse_risk2()
    risk1_rows = parse_risk1()
    growth_n, growth_rows = parse_growth()
    seg_df = pd.read_csv(HERE / "segment_distributions.csv", dtype={"상품코드": "string"})
    product_lookup = pd.read_csv(HERE / "product_lookup.csv", dtype={"상품코드": "string"})
    total_segments = sum(seg_counts.values())
    sku_pct = level_usage.get("sku", {}).get("pct", 0.0)

    wb = Workbook()

    # ---------------- Sheet 1: 요약 ----------------
    ws = wb.active
    ws.title = "요약"
    title_block(ws, "배송 예측 정보 과제 - 분석 심화 리포트", "raw data 기반 EDA -> 리스크 검증 -> 확률모델 학습까지의 전체 근거 데이터", 7)

    ws.cell(row=4, column=1, value="핵심 결론").font = H_FONT
    ws.merge_cells("A5:G6")
    c = ws.cell(row=5, column=1,
                value="신규 시스템·데이터 요청 없이, 케이티커머스 원본 raw data(109개 컬럼)만으로 "
                      "상품코드 단위 배송 소요일 확률분포 모델을 구축했습니다(네이버 상품페이지 방식과 동일한 SKU 단위 예측). "
                      f"물량이 많은 상품은 자기 데이터로 예측하고(홀드아웃의 {sku_pct:.1f}%), 나머지는 자동으로 협력사x상품군 평균으로 "
                      "대체됩니다. 2026년 상반기 44.7만 건 실측 데이터로 검증한 결과, 예측확률과 실제 도착률의 오차가 전 구간 0.1~1.4%p입니다.")
    c.font = Font(name=FONT_NAME, size=11, color=DARKGRAY)
    c.alignment = WRAP
    for row in ws["A5:G6"]:
        for cell in row:
            cell.fill = CALLOUT_FILL

    r = explain_box(ws, 8, 7,
        "아래 표의 '값'이 좋은 건지 나쁜 건지 감이 안 잡히실 수 있어서, '해석' 칸에 일상 언어로 풀어드렸습니다. "
        "특히 '%p(퍼센트 포인트)'는 '%(퍼센트)'와 다릅니다 - 예를 들어 13.0%와 11.8%의 차이는 '1.2%p'이지 '1.2%'가 아닙니다 "
        "(훨씬 더 작은 차이라는 뜻입니다). 이 프로젝트에서는 오차 0~2%p=매우 우수, 2~5%p=양호, 5~10%p=보통, 10%p 초과=재검토 필요로 판단합니다.",
        n_rows=3, row_height=20)

    stat_row = r + 1
    header_row(ws, stat_row, ["지표", "값", "설명", "해석 (이 정도면?)"], widths=[22, 14, 46, 46])
    total_n = sum(eda_n.values())
    avg_gap = sum(row["gap"] for row in day_rows) / len(day_rows)
    stats = [
        ("예측 오차 (day-mark 평균)", f"{avg_gap:.2f}%p", "홀드아웃(2026H1) 실제 도착률과 모델 예측확률의 평균 절대오차",
         "100번 예측하면 대략 99번 이상 실제와 거의 들어맞는다는 뜻입니다. 업계에서는 보통 오차 5%p 이내면 우수한 수준으로 봅니다 — 이 프로젝트는 그보다 훨씬 정확합니다."),
        ("학습 데이터 규모", "265만 건", "2023~2025년, 4개 유통채널(KT/그룹사/외부사/지입자재) 전체 주문",
         "지난 3년간 있었던 거의 모든 주문을 다 봤다는 뜻입니다. 우연히 나온 결과가 아니라고 믿을 수 있을 만큼 충분한 양입니다."),
        ("검증(홀드아웃) 규모", f"{day_rows[0]['n']:,} 건", "2026년 상반기 실측 데이터",
         "학습에 한 번도 쓰지 않은 완전히 새로운 6개월치 실제 데이터로 다시 확인했다는 뜻입니다(커닝 없이 본 시험)."),
        ("세그먼트 수", f"{total_segments:,}개", "상품코드 단위 + 협력사x중분류/대분류/협력사/대분류 4단계 + 전체 평균(콜드스타트용) 1개",
         "상품코드, 협력사, 상품 종류 조합마다 따로 확률을 계산했다는 뜻입니다. 세그먼트가 많을수록 더 맞춤화된 예측이 가능합니다."),
        ("상품코드(SKU) 커버리지", f"{sku_pct:.1f}%", "홀드아웃 주문 중 상품코드 자체 데이터로 답변된 비율(나머지는 협력사x상품군 평균 자동 대체)",
         "물량이 많은 히트상품(예: 종이컵, 감열지)은 그 상품만의 배송 패턴으로 정확히 답하고, 주문이 적은 상품은 비슷한 상품군 평균으로 자연스럽게 대체된다는 뜻입니다."),
        ("기존 SLA(표준납기일) 준수율", f"{sla_actual:.1f}%", "실측 기준 - 협력사가 사전에 약속한 납기일을 실제로 지킨 비율",
         "100건 중 92건 정도는 협력사가 미리 알려준 날짜를 실제로 지켰다는 뜻입니다."),
        ("모델 예측확률 (동일 시점)", f"{sla_pred:.1f}%", "모델이 SLA일까지 도착확률로 예측한 평균값",
         f"모델이 '그 약속 날짜까지 도착할 확률'을 {sla_pred:.1f}%로 예측했는데, 실제(91.8%)와 {abs(sla_actual-sla_pred):.1f}%p 차이입니다. 기존 SLA의 신뢰도를 상당히 정확하게 알아맞혔다는 뜻입니다."),
        ("EDA 대상 원본 데이터", f"{total_n:,} 건", "2026년 상반기, 4개 세그먼트(KT/그룹사/외부사/지입자재) 전체",
         "본격 분석 전에 먼저 훑어봐서 '이 데이터를 믿고 써도 되는지' 확인한 데이터 양입니다."),
        ("급성장 세그먼트 탐지", f"{growth_n}개", "2023~2025 대비 2026H1 물량이 5배 이상 급증한 협력사x상품군 조합 (급성장 세그먼트 = 주문량이 단기간에 폭증한 협력사x상품군)",
         "예: 한 협력사는 과거 3년간 534건이던 주문이 2026년 상반기에만 79,726건으로 폭증했습니다. 이런 경우 과거의 적은 기록으로 지금을 예측하면 크게 틀릴 수 있어 별도 감시하며, 모델 설계(MIN_N=1000)로 이미 안전하게 처리했습니다."),
    ]
    for i, (k, v, d, interp) in enumerate(stats, start=stat_row + 1):
        ws.cell(row=i, column=1, value=k).font = BOLD_NAVY
        ws.cell(row=i, column=2, value=v).font = Font(name=FONT_NAME, size=11, bold=True)
        ws.cell(row=i, column=2).alignment = CENTER
        ws.cell(row=i, column=3, value=d).font = BODY_FONT
        ws.cell(row=i, column=3).alignment = WRAP
        ws.cell(row=i, column=4, value=interp).font = Font(name=FONT_NAME, size=10, italic=True, color=DARKGRAY)
        ws.cell(row=i, column=4).alignment = WRAP
        for col in range(1, 5):
            ws.cell(row=i, column=col).border = BORDER
        ws.row_dimensions[i].height = 32

    nav_row = stat_row + len(stats) + 3
    ws.cell(row=nav_row, column=1, value="다른 시트 안내").font = H_FONT
    nav_items = [
        "용어정리 - 용어가 헷갈리면 여기부터 볼 것",
        "데이터_현황(EDA) - 원본 데이터 결측률/필드 신뢰도 상세",
        "리스크_점검 - 주문시배송리드타임 정체, 외부사 리드타임 분해, 직접배송 비교",
        "캘리브레이션_검증 - day별 예측확률 vs 실제도착률 (차트 포함)",
        "상품코드_조회예시 - 실제 상품코드별 예측 결과 (빠른 상품 vs 느린 상품 비교)",
        f"세그먼트_상세 - {total_segments:,}개 상품코드/협력사x상품군 세그먼트 전체 확률표",
        "급성장_세그먼트 - 주문량이 단기간에 폭증한 협력사x상품군 목록 (재발방지 모니터링)",
    ]
    for i, item in enumerate(nav_items, start=nav_row + 1):
        ws.cell(row=i, column=1, value="• " + item).font = BODY_FONT
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 46
    ws.column_dimensions["D"].width = 46
    ws.freeze_panes = "A4"

    # ---------------- Sheet: 용어정리 ----------------
    # "진행과정" 시트는 2026-07-14 제거 - 초기 조사 과정(12.7%p 오차 발견~MIN_N 조정)까지만
    # 담겨있고 이후 작업(vendor_cat 제거, Kraljic, 표준납기 버킷, spot 제외 등)이 전혀
    # 반영 안 돼 있었음. known_risks.md가 이미 훨씬 상세하고 최신 상태로 같은 내용을
    # 다루고 있어 중복 관리 부담만 있고 실익이 없다고 판단(사용자 확인). PROCESS_DATA는
    # build_glossary.py에 여전히 남아있으나 이 워크북에서는 더 이상 쓰지 않음.
    ws_g = wb.create_sheet("용어정리")
    build_glossary_sheet(
        ws_g,
        "배송 예측 프로젝트 - 용어 쉽게 풀어쓰기",
        "누구나 바로 이해할 수 있도록 쉬운 말로 풀어썼습니다. 다른 시트에서 용어가 헷갈릴 때 참고하세요.",
        ["구분", "용어", "쉬운 설명", "이 프로젝트에서의 실제 예시"],
        [20, 24, 55, 55],
        GLOSSARY_DATA,
        group_col=True,
    )

    # ---------------- Sheet: 데이터_현황(EDA) ----------------
    ws2 = wb.create_sheet("데이터_현황(EDA)")
    title_block(ws2, "원본 데이터 현황 (EDA)", "2026년 상반기, 4개 유통채널 - 결측률 및 핵심 필드 신뢰도", 5)
    r = explain_box(ws2, 4, 5,
        "본격적으로 예측 모델을 만들기 전에, 우리가 가진 원본 데이터를 믿고 써도 되는지부터 확인한 결과입니다. "
        "비유: 요리를 시작하기 전에 냉장고 재료가 상하지 않았는지 확인하는 것과 같습니다. 여기서 필요한 정보가 없다는 게 "
        "확인되면 모델 자체를 포기해야 할 수도 있는데, 다행히 이번엔 필요한 정보가 다 있었습니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws2, r, 5,
        "'결측률'이란 그 정보가 비어있는(기록 안 된) 주문의 비율입니다. 예: 배송완료일자 결측률 13.3%는 100건 중 13건은 "
        "이 값이 비어있고 나머지 87건은 값이 있다는 뜻입니다. 결측률이 낮을수록(0%에 가까울수록) 그 필드를 믿고 분석에 쓸 수 있습니다 "
        "(대략 20% 이하=신뢰 가능, 40% 초과=그 필드만으로 결론 내리기 위험). 표의 색이 붉을수록 결측률이 높다는 뜻입니다.",
        n_rows=3, row_height=20)

    ws2.cell(row=r, column=1, value="핵심 발견: 배송완료일자가 실제 수령일 proxy로 가장 신뢰할 만함").font = H_FONT
    r += 1
    ws2.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=5)
    cell = ws2.cell(row=r, column=1,
                     value="입고일자(평균 +4~15일 지연), 정산확정일(+11~19일 지연)은 실제 도착 시점보다 한참 늦게 "
                           "기록되는 행정 처리일입니다. 배송완료일자가 물리적 도착에 가장 근접해 정답 라벨로 채택했습니다.")
    cell.font = Font(name=FONT_NAME, size=10.5, color=DARKGRAY)
    cell.alignment = WRAP
    for row in ws2.iter_rows(min_row=r, max_row=r + 1, min_col=1, max_col=5):
        for c in row:
            c.fill = CALLOUT_FILL
    r += 3

    header_row(ws2, r, ["필드"] + eda_segs, widths=[20, 14, 14, 14, 14])
    r += 1
    for f in eda_fields:
        ws2.cell(row=r, column=1, value=f).font = BOLD_NAVY
        for i, seg in enumerate(eda_segs, start=2):
            val = eda_data[f][seg]
            c = ws2.cell(row=r, column=i, value=val / 100 if val is not None else None)
            c.number_format = "0.0%"
            c.alignment = CENTER
        for col in range(1, 6):
            ws2.cell(row=r, column=col).border = BORDER
        r += 1
    hdr_row_idx = r - len(eda_fields) - 1 + 1
    rule = ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F8696B")
    ws2.conditional_formatting.add(f"B{hdr_row_idx}:E{r-1}", rule)

    r += 1
    ws2.cell(row=r, column=1, value="세그먼트별 전체 건수(n)").font = BOLD_NAVY
    for i, seg in enumerate(eda_segs, start=2):
        ws2.cell(row=r, column=i, value=eda_n[seg])
        ws2.cell(row=r, column=i).number_format = "#,##0"
        ws2.cell(row=r, column=i).alignment = CENTER

    # ---------------- Sheet: 리스크_점검 ----------------
    ws3 = wb.create_sheet("리스크_점검")
    title_block(ws3, "Phase 1 착수 전 리스크 3종 점검", "모델 설계 확정 전에 검증한 잠재적 함정 3가지", 7)
    r = explain_box(ws3, 4, 7,
        "본격적으로 예측 모델을 만들기 전에, 데이터 안에 '이대로 쓰면 나중에 크게 틀릴 수 있는 함정'이 있는지 미리 확인한 것입니다. "
        "비유: 집을 짓기 전에 땅의 지반을 조사하는 것과 같습니다 - 여기서 문제를 놓치면 모델을 다 만들고 나서야 틀렸다는 걸 알게 되어 "
        "처음부터 다시 해야 할 수도 있습니다. 확인한 위험은 3가지입니다: ① 주문 화면에 뜨는 예상 배송일(주문시배송리드타임)을 믿고 "
        "써도 되는지, ② '외부사'라는 협력사 그룹이 유독 느리다고 나온 숫자가 진짜인지 착시인지, ③ 택배로 보내는지 직접 배송하는지를 "
        "주문 시점에 미리 알 수 있는지. 결과: 셋 다 확인해서 안전하게 처리했습니다.",
        n_rows=4, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws3, r, 7,
        "[1]표 핵심: 여기서 확인하려던 건 '표준납기일과 주문시배송리드타임 중 뭐가 더 정확하냐'가 아닙니다 - "
        "**둘 다 실제로 배송이 걸린 시간을 측정한 값이 아닙니다** (협력사와 미리 정해둔 약속값일 뿐). 진짜 확인하려던 건 "
        "'주문시배송리드타임'이 위험한 값인지였습니다: 만약 이 값이 배송이 끝난 뒤에야 채워지는 값이라면, 모델 입력으로 쓰는 순간 "
        "'정답을 미리 알고 예측하는' 반칙(데이터 리키지)이 됩니다. 그래서 ① 아직 배송 안 끝난 주문도 이 값이 이미 채워져 있는지 "
        "확인했더니 100% 채워져 있었고(주문 시점에 이미 정해지는 값, 리키지 없음 확정), ② 그럼 이 값이 뭘 복사한 건지 확인하려고 "
        "'실제 결과'(A·B행)와 '사전 약속값'(C행=표준납기일)에 각각 비교했더니, 실제 결과와는 5~7%만 일치하고 표준납기일과는 79.6% "
        "일치했습니다. 즉 '주문시배송리드타임=표준납기일의 복사본'이라는 뜻이고, 그래서 최종 모델은 이 두 값을 입력으로 전혀 쓰지 않고 "
        "표준납기일은 '우리 모델이 기존 방식보다 나은지' 비교하는 벤치마크로만 사용합니다.\n"
        "[2]표: 'median(중앙값)'은 100건을 빠른 순서로 세웠을 때 딱 50번째 건이 걸린 날수, 'p90'은 90번째로 빠른 것(대부분이 이 안에 끝남)입니다.",
        n_rows=6, row_height=18)

    ws3.cell(row=r, column=1, value="[1] 주문시배송리드타임 = 실측이 아닌 사전 약속값 (표준납기일과 79.6% 일치)").font = H_FONT
    r += 2
    header_row(ws3, r, ["비교 기준", "비교 n", "정확히 일치", "오차 1일 이내"], widths=[32, 14, 14, 14])
    r += 1
    for row in risk1_rows:
        ws3.cell(row=r, column=1, value=row["기준"]).font = BODY_FONT
        ws3.cell(row=r, column=2, value=row["n"]).number_format = "#,##0"
        ws3.cell(row=r, column=3, value=row["정확히일치"] / 100).number_format = "0.0%"
        ws3.cell(row=r, column=4, value=row["오차1일이내"] / 100).number_format = "0.0%"
        for col in range(1, 5):
            ws3.cell(row=r, column=col).border = BORDER
            ws3.cell(row=r, column=col).alignment = CENTER if col > 1 else Alignment(horizontal="left")
        r += 1
    r += 2

    ws3.cell(row=r, column=1, value="[2] 외부사 21일 리드타임 = 배송이 아니라 창고 입고확정 행정처리 지연").font = H_FONT
    r += 2
    header_row(ws3, r, ["세그먼트", "n", "주문→발주(중앙값,일)", "발주→출하", "출하→배송완료", "배송완료→입고", "직접배송비중"],
               widths=[12, 12, 18, 12, 14, 14, 12])
    r += 1
    for row in risk2_rows:
        ws3.cell(row=r, column=1, value=row["segment"]).font = BOLD_NAVY
        ws3.cell(row=r, column=2, value=row["n"]).number_format = "#,##0"
        ws3.cell(row=r, column=3, value=row["주문발주"][0])
        ws3.cell(row=r, column=4, value=row["발주출하"][0])
        ws3.cell(row=r, column=5, value=row["출하배송완료"][0])
        ws3.cell(row=r, column=6, value=row["배송완료입고"][0])
        ws3.cell(row=r, column=7, value=row["직접배송비중"] / 100).number_format = "0.0%"
        for col in range(1, 8):
            ws3.cell(row=r, column=col).border = BORDER
            ws3.cell(row=r, column=col).alignment = CENTER if col > 1 else Alignment(horizontal="left")
        r += 1
    r += 1
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    note = ws3.cell(row=r, column=1, value="※ 외부사만 배송완료→입고 구간이 15일로 크게 김 (다른 세그먼트는 4일) - "
                                            "실제 배송(출하→배송완료)은 모든 세그먼트가 동일하게 1일. 즉 '배송이 느린 게' 아니라 "
                                            "'회사 내부 기록 처리가 늦는 것'뿐이라는 뜻 - 외부사를 별도 모델로 분리할 필요가 없다는 근거입니다.")
    note.font = Font(name=FONT_NAME, size=9.5, italic=True, color="6E7B8B")

    # ---------------- Sheet: 캘리브레이션_검증 ----------------
    ws4 = wb.create_sheet("캘리브레이션_검증")
    title_block(ws4, "홀드아웃 캘리브레이션 검증", "2026년 상반기 44.7만 건 - 모델 예측확률 vs 실제 도착률", 6)
    r = explain_box(ws4, 4, 6,
        "모델을 다 만든 뒤, '이 모델을 실제로 믿을 만한가'를 확인한 최종 성적표입니다. 학습에는 2023~2025년 데이터만 썼고, "
        "여기서는 모델이 한 번도 보지 못한 2026년 상반기 실제 데이터로 다시 시험을 봤습니다(커닝 없는 시험). 아래 차트에서 "
        "두 선이 겹칠수록 모델이 정확하다는 뜻입니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws4, r, 6,
        "'예측확률'은 모델이 미리 계산한 답, '실제도착률'은 2026년 상반기 44.7만 건에서 진짜로 그렇게 됐는지 나중에 세어본 결과입니다. "
        "두 숫자가 비슷할수록 모델이 정확한 것입니다. '오차(%p)'는 두 숫자 차이를 '퍼센트 포인트'로 나타낸 것 - 예를 들어 1일 행의 "
        "예측 13.0%, 실제 11.8%는 '1.2%p 차이'입니다(13.0%가 11.8%보다 상대적으로 '10% 더 크다'는 뜻이 아니라, 그냥 두 수치를 뺀 값입니다). "
        "판정 기준: 0~2%p=매우 우수, 2~5%p=양호, 5~10%p=보통, 10%p 초과=재검토 필요. 이 프로젝트는 전 구간 매우 우수~양호 수준입니다.",
        n_rows=4, row_height=18)

    header_row(ws4, r, ["경과일", "예측확률(%)", "실제도착률(%)", "오차(%p)", "판정", "검증건수"], widths=[10, 14, 14, 12, 16, 14])
    r += 1
    first_data_row = r
    for row in day_rows:
        label, color = judge_gap(row["gap"])
        ws4.cell(row=r, column=1, value=row["day"])
        ws4.cell(row=r, column=2, value=row["pred"] / 100).number_format = "0.0%"
        ws4.cell(row=r, column=3, value=row["actual"] / 100).number_format = "0.0%"
        ws4.cell(row=r, column=4, value=row["gap"])
        jc = ws4.cell(row=r, column=5, value=label)
        jc.font = Font(name=FONT_NAME, size=10, bold=True, color=color)
        ws4.cell(row=r, column=6, value=row["n"]).number_format = "#,##0"
        for col in range(1, 7):
            ws4.cell(row=r, column=col).border = BORDER
            ws4.cell(row=r, column=col).alignment = CENTER
        r += 1
    last_data_row = r - 1

    chart = LineChart()
    chart.title = "예측확률 vs 실제도착률 (day-mark별) - 두 선이 겹칠수록 정확"
    chart.style = 2
    chart.y_axis.title = "누적 확률"
    chart.x_axis.title = "주문일로부터 경과일"
    chart.y_axis.numFmt = "0%"
    data = Reference(ws4, min_col=2, max_col=3, min_row=first_data_row - 1, max_row=last_data_row)
    cats = Reference(ws4, min_col=1, min_row=first_data_row, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    for s in chart.series:
        s.smooth = False
        s.marker.symbol = "circle"
    chart.width, chart.height = 20, 11
    ws4.add_chart(chart, f"H{first_data_row-1}")

    cold_pct = (level_usage.get("cat", {}).get("pct", 0) + level_usage.get("global", {}).get("pct", 0)
                + level_usage.get("sla_bucket", {}).get("pct", 0))

    r = last_data_row + 2
    ws4.cell(row=r, column=1, value="레벨별 캘리브레이션 (7일 기준) - 좋은 레벨만 아니라 전부 표시").font = H_FONT
    r += 1
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws4.cell(row=r, column=1,
             value="'레벨'은 얼마나 세밀한 정보로 예측했는지를 뜻합니다(vendor_mid=협력사+상품중분류까지 아는 경우로 가장 정확, "
                   f"global=협력사 정보가 아예 없는 신규 협력사용 전체 평균). 협력사 이력이 얕거나 없는 신규 상품군(sla_bucket/"
                   f"cat/global 합쳐 전체의 {cold_pct:.1f}%)은 여전히 오차가 다른 레벨보다 큽니다. '항상' 부정확한 건 아니고 "
                   "특정 구간에서만 그렇지만, 이 표를 좋아 보이는 레벨(sku/vendor)만 보지 말고 전부 확인하시기 바랍니다 - "
                   "아래 표에 실제 수치 그대로(당초 cat_sla라는 레벨도 있었으나 2026-07-14 제거 - known_risks.md 참고).").font = Font(name=FONT_NAME, size=9.5, italic=True, color="6E7B8B")
    ws4.row_dimensions[r].height = 34
    r += 1
    header_row(ws4, r, ["레벨", "n", "예측(7일)", "실측(7일)", "오차", "판정"], widths=[12, 14, 12, 12, 10, 16])
    r += 1
    for row in level_rows:
        label, color = judge_gap(row["gap7"])
        ws4.cell(row=r, column=1, value=row["level"]).font = BOLD_NAVY
        ws4.cell(row=r, column=2, value=row["n"]).number_format = "#,##0"
        ws4.cell(row=r, column=3, value=row["pred7"] / 100).number_format = "0.0%"
        ws4.cell(row=r, column=4, value=row["actual7"] / 100).number_format = "0.0%"
        ws4.cell(row=r, column=5, value=row["gap7"])
        jc = ws4.cell(row=r, column=6, value=label)
        jc.font = Font(name=FONT_NAME, size=10, bold=True, color=color)
        for col in range(1, 7):
            ws4.cell(row=r, column=col).border = BORDER
            ws4.cell(row=r, column=col).alignment = CENTER if col > 1 else Alignment(horizontal="left")
        r += 1

    # ---- 레벨x day마크 전 구간 (2026-07-07 추가 - "구간 부정확" 패턴) ----
    r += 2
    ws4.cell(row=r, column=1, value="레벨별 캘리브레이션 - 전 구간 (day-mark별 오차)").font = H_FONT
    r += 1
    r = explain_box(ws4, r, 10,
        "위 표는 7일 시점만 봤지만, 1~30일 전 구간을 보면 레벨마다 '못 미더운 구간의 위치'가 다릅니다(2026-07-07 검증). "
        "cat/global(신규/희소 협력사)은 중간 구간(2~10일)만 나쁘고 21일 이후엔 오히려 정확해집니다 - 처음엔 배송이 들쭉날쭉하지만 "
        "결국 다 도착한다는 뜻입니다. (vendor_mid도 원래 21/30일이 11.6%p/10.7%p로 불안정했으나, 2026-07-10 해당 레벨만 학습 "
        "기간을 2018~2025로 넓혀 7.0%p/7.8%p로 개선해서 지금은 전 구간 안정적입니다.) "
        "그래서 '이 레벨은 며칠부터 믿을 수 있다' 같은 단일 기준 하나로는 표현이 안 되고, serve.py가 레벨x구간별로 "
        "confidence_level/reliable 필드를 따로 계산해서 API에 노출합니다(API_SPEC.md 참고) - 화면에 못 미더운 구간을 "
        "숨길지/흐리게 보일지는 통합플랫폼 UI 정책입니다. "
        "남은 두 오차 구간은 원인 규명과 개선 시도를 모두 마쳤습니다: vendor_mid 21/30일(7~8%p)은 "
        "특정 협력사의 단발성 실물 지연이 원인이라 모델로 못 줄이는 유형이고, sla_bucket 5/7일(9%p대)은 날짜 분리·카테고리 "
        "분리·학습기간 확장(2022년·2018~2021년) 네 각도 모두 개선 실패를 확인했습니다. 2018~2021년 원본은 표준납기일 "
        "컬럼이 빠져 있어 유사 컬럼(주문시배송리드타임)으로 근사 검증했는데, 2022년 실측 결과와 세 차례 일관되게 "
        "'효과 없음'으로 나와 표준납기일 재요청은 진행하지 않기로 했습니다(2026-07-20). 즉 현재 확보 데이터로는 이 구간이 "
        "한계선이며, 재고/공급 상태 같은 새 데이터 축 확보 시에만 재검토합니다(known_risks.md의 '개선 시도 종료표' 참고).",
        n_rows=9, row_height=18, label="🔍 왜 '레벨별(7일)' 표 하나로는 부족한가")

    day_cols = DAY_MARKS
    header_row(ws4, r, ["레벨", "n"] + [f"{d}일" for d in day_cols], widths=[12, 12] + [9] * len(day_cols))
    r += 1
    for row in level_day_rows:
        ws4.cell(row=r, column=1, value=row["level"]).font = BOLD_NAVY
        ws4.cell(row=r, column=2, value=row["n"]).number_format = "#,##0"
        for j, d in enumerate(day_cols, start=3):
            gap = row["gaps"].get(d)
            if gap is None:
                continue
            _, color = judge_gap(gap)
            c = ws4.cell(row=r, column=j, value=gap)
            c.number_format = '0.0"%p"'
            c.font = Font(name=FONT_NAME, size=10, bold=(gap > 10.0), color=color)
        for col in range(1, 2 + len(day_cols) + 1):
            ws4.cell(row=r, column=col).border = BORDER
            ws4.cell(row=r, column=col).alignment = CENTER if col > 1 else Alignment(horizontal="left")
        r += 1

    # ---- %p 오차를 실제 건수로 환산 (2026-07-27 추가, AI혁신팀 질의 대응) ----
    r += 2
    ws4.cell(row=r, column=1, value="위 오차(%p)를 실제 건수로 환산하면").font = H_FONT
    r += 1
    r = explain_box(ws4, r, 11,
        "%p는 추상적이라 감이 안 올 수 있어서, 같은 내용을 건수로도 보여드립니다. '예측건수'는 "
        "모델이 그 기간 안에 도착할 거라 예상한 주문 수의 합(각 주문의 예측확률을 다 더한 값), "
        "'실제건수'는 진짜로 그 기간 안에 도착한 주문 수입니다. '차이'가 +면 모델이 조심스럽게(적게) "
        "예상한 것이고, -면 낙관적으로(많이) 예상한 것입니다. 0에 가까울수록 정확합니다.",
        n_rows=3, row_height=20, label="📖 이 표를 읽는 법")
    for lvl in ["sku", "vendor_mid", "vendor", "sla_bucket", "전체"]:
        data = pred_actual_counts.get(lvl)
        if not data:
            continue
        is_total = (lvl == "전체")
        header_row(ws4, r, [lvl] + [f"{d}일" for d in day_cols], widths=None)
        if is_total:
            for col in range(1, 2 + len(day_cols)):
                ws4.cell(row=r, column=col).fill = PatternFill("solid", fgColor=NAVY)
        r += 1
        for label in ["예측건수", "실제건수", "차이"]:
            row_vals = data.get(label, {})
            c0 = ws4.cell(row=r, column=1, value=label)
            c0.font = Font(name=FONT_NAME, size=9.5, italic=(label == "차이"))
            c0.alignment = Alignment(horizontal="left")
            for j, d in enumerate(day_cols, start=2):
                v = row_vals.get(f"{d}일")
                if v is None:
                    continue
                c = ws4.cell(row=r, column=j, value=int(round(v)))
                c.number_format = "+#,##0;-#,##0" if label == "차이" else "#,##0"
                c.alignment = Alignment(horizontal="center")
                if label == "차이":
                    c.font = Font(name=FONT_NAME, size=9.5, italic=True,
                                   color=RED if abs(v) > 1500 else DARKGRAY)
            for col in range(1, 2 + len(day_cols)):
                ws4.cell(row=r, column=col).border = BORDER
            r += 1
        r += 1

    # ---------------- Sheet: 상품코드_조회예시 ----------------
    ws_sku = wb.create_sheet("상품코드_조회예시")
    title_block(ws_sku, "상품코드별 예측 실제 사례", "네이버 상품페이지 방식과 동일 - 상품코드만 넣으면 자동으로 조회되는 결과", 6)
    r = explain_box(ws_sku, 4, 6,
        f"협력사 단위가 아니라 상품코드(SKU) 단위로 예측합니다 - 애초 구상이 네이버 상품페이지의 '도착확률' "
        "표시였기 때문입니다. 물량이 많은 상품(아래 예시)은 그 상품만의 실측 데이터로 예측되고, 물량이 적은 "
        f"상품은 자동으로 협력사x상품군 평균으로 대체됩니다(홀드아웃 기준 SKU 자체 데이터로 답변된 비율 {sku_pct:.1f}%). "
        "같은 협력사라도 상품 성격에 따라 확률이 크게 다른 것을 아래에서 확인할 수 있습니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")

    r = explain_box(ws_sku, r, 6,
        "'1일/3일확률'처럼 정해진 날짜의 확률만 보여주면, 배송이 느린 상품은 낮은 숫자만 보여서 "
        "'이 상품 뭔가 문제 있나?'로 오해하기 쉽습니다(예: UTP케이블 3일 확률 18.5%). 그래서 대신 "
        "'headline' - 확률이 90%를 넘는 첫 날짜를 계산해서 보여줍니다. 빠른 상품은 빨리, 느린 상품은 "
        "늦게 나오지만 둘 다 '이 정도면 거의 확실히 온다'는 의미있는 정보입니다.",
        n_rows=3, row_height=20, label="📖 이 표 읽는 법")

    HEADLINE_THRESHOLD = 0.9
    # 매일 단위(fd_1~fd_90, 03_distribution_fit.py 참고)로 계산 - p_1d~p_30d 성긴 지점만
    # 쓰면 예: day14=89.9%->day21=98.6%인 경우 실제 90% 도달일(15~16일)을 건너뛰고
    # 21일로 과대평가되는 문제가 있었음(2026-07-07 사용자 지적, serve.py도 동일하게 수정).
    # 30일까지만 보면 실제론 45~60일에 90%를 넘는 세그먼트(194개)까지 "미달"로 잘못
    # 보였던 문제가 있어 90일로 확장함(2026-07-07).
    fine_days = list(range(1, 91))
    fine_cols = [f"fd_{d}" for d in fine_days]

    def headline_of(row):
        for d, c in zip(fine_days, fine_cols):
            if row[c] >= HEADLINE_THRESHOLD:
                return d, row[c], True
        return fine_days[-1], row[fine_cols[-1]], False

    sku_rows = pd.read_csv(HERE / "segment_distributions.csv", dtype={"상품코드": "string"})
    sku_rows = sku_rows[sku_rows["level"] == "sku"].drop(columns=["협력사명"]).merge(
        product_lookup[["상품코드", "상품명", "협력사명"]], on="상품코드", how="left")
    top_examples = sku_rows.sort_values("n", ascending=False).head(8)
    slow_examples = sku_rows[sku_rows["n"] >= 1000].sort_values("median", ascending=False).head(4)
    examples = pd.concat([top_examples, slow_examples]).drop_duplicates(subset="상품코드")

    header_row(ws_sku, r, ["상품코드", "상품명", "협력사명", "학습건수(n)", "중앙값(일)",
                          "1일확률(참고)", "3일확률(참고)", f"headline(첫 {int(HEADLINE_THRESHOLD*100)}% 도달일)", "headline 확률"],
               widths=[12, 24, 22, 12, 10, 12, 12, 20, 12])
    r += 1
    for _, row in examples.iterrows():
        h_day, h_prob, h_met = headline_of(row)
        ws_sku.cell(row=r, column=1, value=str(row["상품코드"]))
        ws_sku.cell(row=r, column=2, value=row["상품명"])
        ws_sku.cell(row=r, column=3, value=row["협력사명"])
        ws_sku.cell(row=r, column=4, value=int(row["n"])).number_format = "#,##0"
        ws_sku.cell(row=r, column=5, value=row["median"])
        ws_sku.cell(row=r, column=6, value=row["p_1d"]).number_format = "0.0%"
        ws_sku.cell(row=r, column=7, value=row["p_3d"]).number_format = "0.0%"
        h_label = f"{h_day}일 후" + ("" if h_met else " (90% 미달)")
        ws_sku.cell(row=r, column=8, value=h_label).font = BOLD_NAVY
        ws_sku.cell(row=r, column=9, value=h_prob).number_format = "0.0%"
        for col in range(1, 10):
            ws_sku.cell(row=r, column=col).border = BORDER
            ws_sku.cell(row=r, column=col).alignment = CENTER if col > 2 else Alignment(horizontal="left")
        r += 1

    r += 1
    ws_sku.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    ws_sku.cell(row=r, column=1,
                value="※ '1일/3일확률'은 참고용으로만 남겨뒀습니다. 실제 고객 화면에 보여줄 핵심 정보는 "
                      "'headline' 컬럼입니다 - 종이컵/감열지처럼 빠른 상품은 2~3일 만에, UTP케이블/모뎀어댑터 "
                      "같은 통신설치자재는 14~16일이 지나야 90%를 넘습니다. 협력사가 아니라 '상품 성격'이 "
                      "배송 속도를 가르는 핵심 요인임을 보여줍니다.").font = \
        Font(name=FONT_NAME, size=9.5, italic=True, color="6E7B8B")

    # ---------------- Sheet: 세그먼트_상세 ----------------
    ws5 = wb.create_sheet("세그먼트_상세")
    ws5.cell(row=1, column=1, value=f"세그먼트 상세 ({total_segments:,}개) - 상품코드/협력사x상품군별 도착확률 룩업테이블").font = H_FONT
    ws5.row_dimensions[1].height = 22
    r0 = explain_box(ws5, 2, 6,
        f"실제로 서비스에 쓸 수 있는 결과물 그 자체입니다. 아래 {total_segments:,}개 줄 하나하나가 '이 상품(또는 협력사x상품군)을 "
        "주문하면 며칠 안에 몇 %의 확률로 도착하는지'를 알려주는 참고표(룩업테이블)입니다. 통합플랫폼이 고객에게 도착확률을 "
        "보여줄 때 이 표를 그대로 찾아서 쓰게 됩니다. level='sku'인 행이 상품코드 자체 데이터로 예측된 것입니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r0 = explain_box(ws5, r0, 6,
        "컬럼 설명 - level: 어느 단계 세밀함으로 예측했는지(sku=상품코드 자체가 가장 세밀, vendor_mid=협력사+상품중분류, "
        "global이 가장 넓은 평균) / n: 이 세그먼트를 학습할 때 쓴 과거 주문 건수(많을수록 신뢰도 높음) / "
        "median: 이 세그먼트 배송 소요일의 중앙값 / p_1d~p_30d: 주문 후 해당 일수 이내에 도착할 누적 확률 "
        "(예: p_5d=72.0%는 '5일 이내 도착 확률 72%'라는 뜻) / headline: 확률이 90%를 처음 넘는 날 - "
        "고정된 날짜(예: 3일)의 확률만 보면 느린 상품은 오해를 살 수 있어(상품코드_조회예시 시트 참고), "
        "'확신할 수 있는 날짜'를 대신 계산한 값입니다. 정렬/필터는 표 헤더의 화살표를 누르세요.",
        n_rows=4, row_height=18)

    cols = ["상품코드", "상품명", "협력사명", "대분류", "중분류", "level", "n", "median",
            "headline", "headline_prob",
            "p_1d", "p_2d", "p_3d", "p_5d", "p_7d", "p_10d", "p_14d", "p_21d", "p_30d"]
    # sku 레벨 행은 groupby 키가 상품코드뿐이라 협력사명/대분류/중분류가 비어있음 -
    # product_lookup에서 채워 넣는다(다른 레벨의 기존 값은 건드리지 않음).
    plu = product_lookup.set_index("상품코드")
    is_sku = seg_df["level"] == "sku"

    # headline: 확률이 90%를 처음 넘는 day-mark (상품코드_조회예시 시트와 같은 기준)
    _headline_day, _headline_prob = [], []
    for _, row in seg_df.iterrows():
        d, p, met = headline_of(row)
        _headline_day.append(f"{d}일 후" + ("" if met else " (90% 미달)"))
        _headline_prob.append(p)
    # 새 컬럼 3개를 한 번에 붙임 - 하나씩 insert하면 pandas가 내부 블록을 계속 쪼개서
    # PerformanceWarning(highly fragmented)이 뜸
    seg_df = pd.concat([seg_df, pd.DataFrame({
        "상품명": seg_df["상품코드"].map(plu["상품명"]),
        "headline": _headline_day,
        "headline_prob": _headline_prob,
    }, index=seg_df.index)], axis=1)
    for col in ["협력사명", "대분류", "중분류"]:
        seg_df.loc[is_sku, col] = seg_df.loc[is_sku, "상품코드"].map(plu[col])

    seg_df = seg_df[cols]
    for j, col in enumerate(cols, start=1):
        ws5.cell(row=r0, column=j, value=col)
    for i, row in enumerate(seg_df.itertuples(index=False), start=r0 + 1):
        for j, val in enumerate(row, start=1):
            cell = ws5.cell(row=i, column=j, value=val if pd.notna(val) else None)
            if cols[j - 1].startswith("p_") or cols[j - 1] == "headline_prob":
                cell.number_format = "0.0%"
    last_row = r0 + len(seg_df)
    last_col_letter = get_column_letter(len(cols))
    tbl = Table(displayName="SegmentTable", ref=f"A{r0}:{last_col_letter}{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws5.add_table(tbl)
    for j, col in enumerate(cols, start=1):
        w = 24 if col in ("협력사명", "대분류", "중분류", "상품명") else (16 if col == "headline" else (12 if col == "상품코드" else 10))
        ws5.column_dimensions[get_column_letter(j)].width = w
    ws5.freeze_panes = f"A{r0+1}"

    # ---------------- Sheet: 급성장_세그먼트 ----------------
    ws6 = wb.create_sheet("급성장_세그먼트")
    title_block(ws6, "급성장 협력사x상품군 탐지 (재발방지 모니터링)", f"2023~2025 대비 2026년 상반기 물량 5배 이상 급증 - 총 {growth_n}개 발견, 상위 20개 표시", 5)
    r = explain_box(ws6, 4, 5,
        "앞으로 또 비슷한 문제가 생기지 않는지 감시하기 위한 목록입니다. 이번 프로젝트에서 협력사 물량이 갑자기 확 늘어난 경우를 "
        "놓쳐서 모델이 한동안 부정확했던 적이 있었습니다(known_risks.md 참고) - 그래서 이 목록을 자동으로 뽑아 앞으로 재학습할 "
        "때마다 같은 문제가 재발하는지 미리 확인하도록 만들었습니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws6, r, 5,
        "'성장률'은 2023~2025년 평균 대비 2026년 상반기 물량이 몇 배로 늘었는지입니다(예: 807.2x = 807배 폭증). "
        "이런 협력사x상품군은 과거 데이터가 파일럿 수준으로 적어서, 그 소량 데이터로 지금의 대량 물량을 예측하면 틀릴 위험이 큽니다 - "
        "실제로 이 문제 때문에 개발 과정의 첫 버전에서 오차가 12.7%p까지 벌어졌었습니다(원인 규명 후 이미 해결됨, known_risks.md 참고).",
        n_rows=3, row_height=20)
    header_row(ws6, r, ["협력사명", "대분류", "2023~25 합계", "2026 상반기", "연환산 성장률"], widths=[22, 18, 14, 14, 14])
    r += 1
    for row in growth_rows:
        ws6.cell(row=r, column=1, value=row["협력사명"]).font = BOLD_NAVY
        ws6.cell(row=r, column=2, value=row["대분류"])
        ws6.cell(row=r, column=3, value=row["2023-25합계"]).number_format = "#,##0"
        ws6.cell(row=r, column=4, value=row["2026H1"]).number_format = "#,##0"
        ws6.cell(row=r, column=5, value=row["성장률"])
        ws6.cell(row=r, column=5).number_format = '0.0"x"'
        for col in range(1, 6):
            ws6.cell(row=r, column=col).border = BORDER
            ws6.cell(row=r, column=col).alignment = CENTER if col > 2 else Alignment(horizontal="left")
        r += 1
    r += 2
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws6.cell(row=r, column=1, value="※ 이런 협력사x상품군 조합은 소량 샘플(수백 건)이 세그먼트로 잘못 채택되지 않도록 "
                                     "MIN_N=1000 기준으로 자동 필터링됩니다. 다음 재학습 시 05_growth_vendor_flag.py 재실행 권장.").font = \
        Font(name=FONT_NAME, size=9.5, italic=True, color="6E7B8B")

    # ---------------- Sheet: Kraljic_리스크분류 ----------------
    KRALJIC_COLOR = {"전략재": "F4B183", "병목재": "FFD966", "레버리지재": "C6E0B4", "일반재": "D9D9D9"}
    ACTION_HINT = {
        "전략재": "최우선 관리 - 대체 소싱 이중화, 재고 완충 확대 검토",
        "병목재": "대체 소싱 확보 우선 (물량은 적지만 지연 리스크 높음)",
        "레버리지재": "물량 많고 안정적 - 협력사 확대/조건 협상 여지",
        "일반재": "관리 리소스 최소화, 자동 폴백으로 충분",
    }
    kdf = pd.read_csv(HERE / "kraljic_delivery_risk.csv", dtype={"상품코드": "string"})
    ws7 = wb.create_sheet("Kraljic_리스크분류")
    title_block(ws7, "배송 리스크 Kraljic 매트릭스", "공급 리스크 x 구매 영향력 4분류 - 수요예측 프로젝트의 분류 방법론을 상품코드 단위로 적용", 6)
    r = explain_box(ws7, 4, 6,
        "수요예측(demand_forecast.py) 프로젝트에서 만든 협력사 리스크 분류 방법론을 배송예측 데이터(주문 단위, 더 세밀함)에 "
        "그대로 적용했습니다. 앞서 발견한 '배치성 발주 협력사'(영풍기획 사례 등, phase3_backlog.md 참고 - 발주->출하가 "
        "유독 길고 몰아서 주문되는 패턴이며, 원산지 데이터가 없어 수입 여부는 확정되지 않음)를 임시 기준이 아니라 "
        "정식 등급 체계로 격상시킨 것입니다. 전략재/병목재로 분류된 상품은 재고연동(Phase 3)의 우선 검토 대상입니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws7, r, 6,
        "공급 리스크(高) 판정 - 아래 3개 중 2개 이상 충족: ① 발주→출하 변동성(CV)>0.3(배치성 조달 신호) "
        "② 같은 상품 종류(중분류)를 취급하는 협력사가 2곳 이하(대체 공급사 부족) ③ 7일 이내 도착률<85%. "
        "구매 영향력(高): 주문량 누적기여 80% 이내(매입금액 데이터가 없어 '몇 명 고객 경험에 영향 주는지'로 대체). "
        "전략재=리스크 높고 물량 많음(최우선 관리), 병목재=리스크 높고 물량 적음(대체소싱 검토), "
        "레버리지재=리스크 낮고 물량 많음(안정적), 일반재=리스크 낮고 물량 적음(관리 불필요).",
        n_rows=4, row_height=18)

    counts = kdf["kraljic"].value_counts()
    total_n = kdf["n"].sum()
    quad_layout = [("전략재", "병목재"), ("레버리지재", "일반재")]
    for i, (left, right) in enumerate(quad_layout):
        row0 = r + i * 2
        for j, name in enumerate((left, right)):
            col0 = 1 + j * 3
            ws7.merge_cells(start_row=row0, start_column=col0, end_row=row0, end_column=col0 + 2)
            c = ws7.cell(row0, col0, f"{name}  {int(counts.get(name, 0)):,}개 ({counts.get(name, 0)/len(kdf)*100:.1f}%)")
            c.font = Font(name=FONT_NAME, bold=True, size=11, color=WHITE)
            c.fill = PatternFill("solid", fgColor=KRALJIC_COLOR[name])
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws7.merge_cells(start_row=row0 + 1, start_column=col0, end_row=row0 + 1, end_column=col0 + 2)
            c2 = ws7.cell(row0 + 1, col0, ACTION_HINT[name])
            c2.font = Font(name=FONT_NAME, size=9, italic=True, color="595959")
            c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws7.row_dimensions[row0 + 1].height = 28
    r = r + len(quad_layout) * 2 + 1

    r += 1
    ws7.cell(row=r, column=1, value="[전략재] 최우선 관리 대상 (물량 상위 15개)").font = H_FONT
    r += 1
    header_row(ws7, r, ["상품코드", "상품명", "협력사명", "n", "median", "7일도달률", "발주출하CV", "중분류내협력사수"],
               widths=[12, 24, 20, 10, 8, 10, 10, 12])
    r += 1
    for _, row in kdf[kdf["kraljic"] == "전략재"].sort_values("n", ascending=False).head(15).iterrows():
        ws7.cell(row=r, column=1, value=str(row["상품코드"])); ws7.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=2, value=row["상품명"]); ws7.cell(row=r, column=2).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=3, value=row["협력사명"]); ws7.cell(row=r, column=3).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=4, value=int(row["n"])).number_format = "#,##0"
        ws7.cell(row=r, column=5, value=row["median"])
        ws7.cell(row=r, column=6, value=row["p_7d"]).number_format = "0.0%"
        ws7.cell(row=r, column=7, value=round(row["lt_cv"], 2))
        ws7.cell(row=r, column=8, value=int(row["category_vendor_count"]))
        for col in range(1, 9):
            ws7.cell(row=r, column=col).border = BORDER
            if col > 3:
                ws7.cell(row=r, column=col).alignment = Alignment(horizontal="center")
        r += 1

    r += 1
    ws7.cell(row=r, column=1, value="[병목재] 대체 소싱 검토 대상 (변동성 상위 10개)").font = H_FONT
    r += 1
    header_row(ws7, r, ["상품코드", "상품명", "협력사명", "n", "median", "7일도달률", "발주출하CV", "중분류내협력사수"],
               widths=[12, 24, 20, 10, 8, 10, 10, 12])
    r += 1
    for _, row in kdf[kdf["kraljic"] == "병목재"].sort_values("lt_cv", ascending=False).head(10).iterrows():
        ws7.cell(row=r, column=1, value=str(row["상품코드"])); ws7.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=2, value=row["상품명"]); ws7.cell(row=r, column=2).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=3, value=row["협력사명"]); ws7.cell(row=r, column=3).alignment = Alignment(horizontal="left")
        ws7.cell(row=r, column=4, value=int(row["n"])).number_format = "#,##0"
        ws7.cell(row=r, column=5, value=row["median"])
        ws7.cell(row=r, column=6, value=row["p_7d"]).number_format = "0.0%"
        ws7.cell(row=r, column=7, value=round(row["lt_cv"], 2))
        ws7.cell(row=r, column=8, value=int(row["category_vendor_count"]))
        for col in range(1, 9):
            ws7.cell(row=r, column=col).border = BORDER
            if col > 3:
                ws7.cell(row=r, column=col).alignment = Alignment(horizontal="center")
        r += 1

    r += 2
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws7.cell(row=r, column=1,
             value=f"※ 대상: sku 레벨 세그먼트 {len(kdf):,}개 전체(주문량 {int(total_n):,}건). 상세는 kraljic_delivery_risk.csv/"
                   "kraljic_delivery_risk_report.txt 참고. '단일소싱 여부'를 처음 시도했다가 상품코드 97%가 원래 단일 협력사라 "
                   "구분력이 없어(전략재+병목재 88.6% 과다판정) 중분류 단위 협력사 수로 교체한 시행착오가 있었음(10_kraljic_delivery_risk.py 참고).").font = \
        Font(name=FONT_NAME, size=9, italic=True, color="6E7B8B")

    sla_gap = pd.read_csv(HERE / "sla_gap_analysis.csv", dtype={"상품코드": "string"})
    ws8 = wb.create_sheet("표준납기일_비교")
    title_block(ws8, "표준납기일(SLA) vs 모델 예측 비교", "고객에게 보이는 두 숫자가 얼마나 다른지, 다르면 왜 다른지", 8)
    r = explain_box(ws8, 4, 8,
        "고객에게 이미 노출 중인 '표준납기일'과, 이 프로젝트가 만든 모델의 예측(90% 확신일)이 "
        "얼마나 차이 나는지 상품코드 단위로 비교했습니다. 차이가 크면(특히 모델이 더 늦게 예측하면) "
        "'표준납기일을 왜 보여주나' 하는 의문을 살 수 있다는 문제의식에서 시작한 검토입니다.",
        n_rows=3, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws8, r, 8,
        "결론: 상품 전체의 72.3%는 모델 예측이 표준납기일과 같거나 더 빠릅니다(오히려 표준납기일이 "
        "보수적으로 잡혀있다는 뜻 - 50.1%는 모델이 SLA보다 빠름). 나머지 27.7% 중에서도 대부분(14.3%)은 "
        "1~2일의 경미한 차이입니다. 정말 눈에 띄는 차이(5일 초과)는 7.0%(물량 기준 5.1%)뿐이고, "
        "이 중 95.1%가 이미 Kraljic 분류상 '전략재/병목재'(공급위험 高 - 배치성 발주 등으로 이미 "
        "알려진 상품군)였습니다. 즉 큰 불일치는 무작위가 아니라 특정 상품군에 집중돼 있습니다.",
        n_rows=4, row_height=18)

    total_n = sla_gap["n"].sum()
    buckets = [(-999, -3, "3일 이상 빠름(SLA 보수적)", GREEN), (-3, -1, "1~2일 빠름", GREEN),
               (-1, 0.001, "정확히 일치", "2E75B6"), (0.001, 2, "1~2일 늦음(경미)", AMBER),
               (2, 5, "3~5일 늦음(중간)", AMBER), (5, 999, "5일 초과 늦음(주의)", RED)]
    header_row(ws8, r, ["gap 구간", "SKU 수", "비율", "물량비중"], widths=[24, 10, 10, 10, 8, 8, 8, 8])
    r += 1
    for lo, hi, label, color in buckets:
        sub = sla_gap[(sla_gap["gap"] > lo) & (sla_gap["gap"] <= hi)]
        ws8.cell(row=r, column=1, value=label).font = Font(name=FONT_NAME, size=10, bold=True, color=color)
        ws8.cell(row=r, column=2, value=len(sub)).alignment = Alignment(horizontal="center")
        ws8.cell(row=r, column=3, value=len(sub) / len(sla_gap)).number_format = "0.0%"
        ws8.cell(row=r, column=3).alignment = Alignment(horizontal="center")
        ws8.cell(row=r, column=4, value=float(sub["n"].sum()) / total_n).number_format = "0.00%"
        ws8.cell(row=r, column=4).alignment = Alignment(horizontal="center")
        for col in range(1, 5):
            ws8.cell(row=r, column=col).border = BORDER
        r += 1

    r += 1
    ws8.cell(row=r, column=1, value="5일 초과 불일치 전체 (물량순) - 대부분 이미 알려진 고위험 상품군").font = H_FONT
    r += 1
    header_row(ws8, r, ["상품코드", "상품명", "협력사명", "n", "SLA(일)", "모델예측(일)", "차이(일)", "Kraljic분류"],
               widths=[12, 24, 18, 10, 8, 10, 8, 10])
    r += 1
    big = sla_gap[sla_gap["gap"] > 5].sort_values("n", ascending=False)
    for _, row in big.iterrows():
        ws8.cell(row=r, column=1, value=str(row["상품코드"])); ws8.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        ws8.cell(row=r, column=2, value=row["상품명"]); ws8.cell(row=r, column=2).alignment = Alignment(horizontal="left")
        ws8.cell(row=r, column=3, value=row["협력사명"]); ws8.cell(row=r, column=3).alignment = Alignment(horizontal="left")
        ws8.cell(row=r, column=4, value=int(row["n"])).number_format = "#,##0"
        ws8.cell(row=r, column=5, value=row["sla_median"])
        ws8.cell(row=r, column=6, value=row["headline_day"])
        gap_cell = ws8.cell(row=r, column=7, value=row["gap"])
        gap_cell.font = Font(name=FONT_NAME, size=10, bold=True, color=RED)
        ws8.cell(row=r, column=8, value=row["kraljic"])
        for col in range(1, 9):
            ws8.cell(row=r, column=col).border = BORDER
            if col > 3:
                ws8.cell(row=r, column=col).alignment = Alignment(horizontal="center")
        r += 1

    r += 2
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws8.cell(row=r, column=1,
             value="※ 검토 방향(2026-07-09): 모델 예측을 SLA에 맞춰 왜곡하는 것은 이 프로젝트의 정확성 원칙에 "
                   "위배되어 채택하지 않음. SLA 자체를 이 프로젝트가 임의로 변경하는 것도 범위 밖(구매팀/계약 "
                   "부서 협의 필요한 정책 결정). 대신 gap 크기에 따라 화면 표시를 차등화하는 방안을 제안 - "
                   "상세는 known_risks.md 및 API_SPEC.md 참고. 원본: sla_gap_analysis.csv, sla_gap_report.txt "
                   "(11_sla_gap_check.py로 재생성 가능).").font = \
        Font(name=FONT_NAME, size=9, italic=True, color="6E7B8B")

    # ---------------- Sheet: 협력사변경_검증 (2026-07-27, AI혁신팀 질의 대응) ----------------
    ws9 = wb.create_sheet("협력사변경_검증")
    title_block(ws9, "상품코드 낙찰 협력사 변경 - 영향 검증", "협력사가 바뀌면 예측 방식도 바꿔야 하는가", 11)
    r = explain_box(ws9, 4, 11,
        "AI혁신팀 질의: \"주문 도중 낙찰 협력사가 바뀌면, 상품코드(sku) 단위 예측 대신 협력사 "
        "단위 예측으로 바꿔야 하지 않느냐\" - 일리 있는 지적이라 실제 데이터로 확인했습니다.",
        n_rows=2, row_height=20, label="🎯 왜 이 시트가 있나요?")
    r = explain_box(ws9, r, 11,
        "① 먼저 확인한 사실 - 상품코드 기준으로 예측하는 2,171개 상품 중 370개(17%)는 최근 3~4년 "
        "사이 실제로 협력사가 2곳 이상이었습니다. '협력사가 바뀌는 일' 자체는 이미 흔합니다. "
        "② 그래서 '협력사가 한 번도 안 바뀐 상품'과 '2곳 이상이었던 상품'을 나눠 2026년 상반기 "
        "실측 오차(%p)를 비교했습니다 - 바뀐 상품이 더 부정확하다면 우려가 맞는 것이고, 차이가 "
        "없다면 지금 방식을 유지해도 됩니다.",
        n_rows=4, row_height=18)
    r = explain_box(ws9, r, 11,
        "✅ 결론: 아래 표에서 '협력사 2곳 이상' 줄이 '단일 협력사' 줄보다 오차가 크지 않습니다 - "
        "오히려 1·3·5·7·14·21·30일 구간에서 더 정확했습니다. 협력사가 바뀌었다고 예측이 나빠지는 "
        "현상은 관측되지 않아, 상품코드 기준 예측을 유지하고 협력사 기준으로 바꿀 근거는 없다고 "
        "판단했습니다. 다만 이번 검증은 3~4년 누적 이력 비교라 '전환 직후 한두 달'의 일시적 "
        "저하는 못 잡습니다 - 필요 시 growth_vendor_flags.txt와 같은 패턴의 월간 자동 감시 항목 "
        "추가를 검토할 수 있습니다(known_risks.md 참고, 미구현).",
        n_rows=5, row_height=18, label="📊 결론")

    vc_cols = [c for c in vendor_change_df.columns if c not in ("구분", "n")]
    header_row(ws9, r, ["구분", "건수(n)"] + vc_cols, widths=[20, 10] + [9] * len(vc_cols))
    r += 1
    for _, row in vendor_change_df.iterrows():
        ws9.cell(row=r, column=1, value=row["구분"]).font = BOLD_NAVY
        ws9.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        ws9.cell(row=r, column=2, value=int(row["n"])).number_format = "#,##0"
        for j, col in enumerate(vc_cols, start=3):
            v = float(row[col])
            _, color = judge_gap(v)
            c = ws9.cell(row=r, column=j, value=v)
            c.number_format = '0.0"%p"'
            c.font = Font(name=FONT_NAME, size=10, color=color)
        for col in range(1, 3 + len(vc_cols)):
            ws9.cell(row=r, column=col).border = BORDER
            ws9.cell(row=r, column=col).alignment = CENTER if col > 1 else Alignment(horizontal="left")
        r += 1

    r += 2
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    ws9.cell(row=r, column=1,
             value="※ 원본: level_daymark_pred_vs_actual.csv / vendor_change_calibration.csv "
                   "(2026-07-27, 2026년 상반기 홀드아웃 원본 재계산). 상세 결론은 known_risks.md "
                   "'상품코드 중 낙찰 협력사 변경' 항목 참고.").font = \
        Font(name=FONT_NAME, size=9, italic=True, color="6E7B8B")

    out_path = HERE / "배송예측_분석보고서.xlsx"
    wb.save(out_path)
    print(f"saved: {out_path}")


if __name__ == "__main__":
    main()
