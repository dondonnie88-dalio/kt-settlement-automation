
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

# ─────────────────────────────────────────
# 공통 설정
# ─────────────────────────────────────────
PARTNERS = ["삼성전자", "LG전자", "SK하이닉스", "현대모비스", "롯데케미칼"]
TAX_TYPES = ["과세", "면세"]
FONT_NAME = "Arial"

HEADER_FILL = PatternFill("solid", start_color="1F4E79")   # 짙은 파랑
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
CELL_FONT   = Font(name=FONT_NAME, size=10)
EVEN_FILL   = PatternFill("solid", start_color="EBF3FB")
ODD_FILL    = PatternFill("solid", start_color="FFFFFF")

thin = Side(style="thin", color="B0B0B0")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

COLS = ["주문번호", "요청번호", "입고일", "정산금액", "협력사명", "매출과세구분", "주문유형"]
COL_WIDTHS = [18, 18, 14, 16, 14, 14, 10]

def fmt_sheet(ws):
    """헤더 스타일 + 열 너비 적용"""
    for col_idx, (col_name, width) in enumerate(zip(COLS, COL_WIDTHS), 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20

def write_row(ws, row_num, values):
    fill = EVEN_FILL if row_num % 2 == 0 else ODD_FILL
    for col_idx, value in enumerate(values, 1):
        cell = ws.cell(row=row_num, column=col_idx, value=value)
        cell.font = CELL_FONT
        cell.fill = fill
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col_idx == 4:  # 정산금액
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")

def date_str(day):
    return datetime.date(2025, 5, day)

# ─────────────────────────────────────────
# 데이터 정의
# ─────────────────────────────────────────

# [A] 일반 주문 15건 (KT ↔ 플랫폼 완전 일치)
normal_orders = [
    # (주문번호, 요청번호, 입고일day, 정산금액, 협력사, 과세구분)
    ("ORD-20250501-001", "REQ-001", 1,  1_200_000, "삼성전자",   "과세"),
    ("ORD-20250501-002", "REQ-002", 1,    850_000, "LG전자",     "과세"),
    ("ORD-20250502-003", "REQ-003", 2,  2_340_000, "SK하이닉스", "과세"),
    ("ORD-20250502-004", "REQ-004", 2,    670_000, "현대모비스", "면세"),
    ("ORD-20250503-005", "REQ-005", 3,  1_890_000, "롯데케미칼", "과세"),
    ("ORD-20250505-006", "REQ-006", 5,    530_000, "삼성전자",   "면세"),
    ("ORD-20250506-007", "REQ-007", 6,  3_100_000, "LG전자",     "과세"),
    ("ORD-20250507-008", "REQ-008", 7,    720_000, "SK하이닉스", "과세"),
    ("ORD-20250508-009", "REQ-009", 8,  1_450_000, "현대모비스", "과세"),
    ("ORD-20250509-010", "REQ-010", 9,    980_000, "롯데케미칼", "면세"),
    ("ORD-20250510-011", "REQ-011", 10, 2_650_000, "삼성전자",   "과세"),
    ("ORD-20250512-012", "REQ-012", 12,   440_000, "LG전자",     "면세"),
    ("ORD-20250513-013", "REQ-013", 13, 1_770_000, "SK하이닉스", "과세"),
    ("ORD-20250514-014", "REQ-014", 14,   610_000, "현대모비스", "과세"),
    ("ORD-20250515-015", "REQ-015", 15, 3_280_000, "롯데케미칼", "과세"),
]

# [B] KT에만 있는 주문 3건 (통합플랫폼 누락 ← 핵심 테스트 케이스)
kt_only_orders = [
    ("ORD-20250504-016", "REQ-016", 4,    990_000, "삼성전자",   "과세"),
    ("ORD-20250508-017", "REQ-017", 8,  1_580_000, "LG전자",     "과세"),
    ("ORD-20250511-018", "REQ-018", 11,   730_000, "SK하이닉스", "면세"),
]

# [C] 반품 2건
#   KT  → 반품 주문번호(RTN-...), 요청번호는 원 주문의 REQ와 동일
#   플랫폼 → 원 주문번호(ORD-...), 동일 요청번호
returns_kt = [
    # (반품주문번호, 요청번호, 입고일day, 정산금액(음수), 협력사, 과세구분)
    ("RTN-20250506-019", "REQ-019", 6,  -450_000, "현대모비스", "과세"),
    ("RTN-20250513-020", "REQ-020", 13, -870_000, "롯데케미칼", "과세"),
]
returns_platform = [
    # (원주문번호, 요청번호, 입고일day, 정산금액(음수), 협력사, 과세구분)
    ("ORD-20250504-019", "REQ-019", 4,  -450_000, "현대모비스", "과세"),
    ("ORD-20250511-020", "REQ-020", 11, -870_000, "롯데케미칼", "과세"),
]

# ─────────────────────────────────────────
# kt_raw.xlsx 생성  (20건)
# ─────────────────────────────────────────
wb_kt = openpyxl.Workbook()
ws_kt = wb_kt.active
ws_kt.title = "KT_입고내역"
ws_kt.freeze_panes = "A2"

fmt_sheet(ws_kt)

all_kt_rows = []
for o in normal_orders:
    all_kt_rows.append((o[0], o[1], date_str(o[2]), o[3], o[4], o[5], "일반"))
for o in kt_only_orders:
    all_kt_rows.append((o[0], o[1], date_str(o[2]), o[3], o[4], o[5], "일반"))
for o in returns_kt:
    all_kt_rows.append((o[0], o[1], date_str(o[2]), o[3], o[4], o[5], "반품"))

for i, row_vals in enumerate(all_kt_rows, 2):
    write_row(ws_kt, i, row_vals)

wb_kt.save("kt_raw.xlsx")
print(f"[OK] kt_raw.xlsx 생성 완료 ({len(all_kt_rows)}건)")

# ─────────────────────────────────────────
# platform.xlsx 생성  (17건 = 15 + 반품2)
# ─────────────────────────────────────────
wb_pl = openpyxl.Workbook()
ws_pl = wb_pl.active
ws_pl.title = "플랫폼_입고내역"
ws_pl.freeze_panes = "A2"

fmt_sheet(ws_pl)

all_pl_rows = []
for o in normal_orders:
    all_pl_rows.append((o[0], o[1], date_str(o[2]), o[3], o[4], o[5], "일반"))
for o in returns_platform:
    all_pl_rows.append((o[0], o[1], date_str(o[2]), o[3], o[4], o[5], "반품"))

for i, row_vals in enumerate(all_pl_rows, 2):
    write_row(ws_pl, i, row_vals)

wb_pl.save("platform.xlsx")
print(f"[OK] platform.xlsx 생성 완료 ({len(all_pl_rows)}건)")
print("\n[데이터 요약]")
print(f"  KT 전체   : {len(all_kt_rows)}건  (일반15 + KT누락3 + 반품2)")
print(f"  플랫폼 전체: {len(all_pl_rows)}건  (일반15 + 반품2)")
print(f"  KT전용(누락예상): {[o[0] for o in kt_only_orders]}")
print(f"  반품 요청번호: {[o[1] for o in returns_kt]}")
