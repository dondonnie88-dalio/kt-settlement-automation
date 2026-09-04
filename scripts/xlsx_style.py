"""
공통 Excel 서식 모듈 (요청서 12장 기준)
모든 output/*.xlsx 생성 스크립트에서 공유해서 사용한다.
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---- 색상 정의 ----
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")   # 진한 파란색
HEADER_FONT = Font(color="FFFFFF", bold=True)

ORIGIN_FONT = Font(color="008000")       # 원본값: 녹색 글씨
INPUT_FONT = Font(color="0000FF")        # 사용자 입력값: 파란색 글씨

REVIEW_FILL = PatternFill("solid", fgColor="FFE699")    # 검토 필요: 주황
ERROR_FILL = PatternFill("solid", fgColor="FFC7CE")     # 오류/불일치: 연한 빨강
CONFIRMED_FILL = PatternFill("solid", fgColor="C6EFCE")  # 확정 후보: 연한 녹색

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

WRAP_TOP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

# ---- 숫자 서식 ----
FMT_AMOUNT = '#,##0;[Red](#,##0);"-"'          # 천단위 구분, 음수 빨강 괄호, 0은 '-'
FMT_QTY = '#,##0;[Red](#,##0);"-"'
FMT_PERCENT = '0.0%;[Red](0.0%);"-"'
FMT_DATE = 'yyyy-mm-dd'


def style_header(ws: Worksheet, header_row: int, n_cols: int, start_col: int = 1):
    for c in range(start_col, start_col + n_cols):
        cell = ws.cell(row=header_row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP_CENTER
        cell.border = THIN_BORDER


def freeze_and_filter(ws: Worksheet, header_row: int, n_cols: int, n_rows: int, start_col: int = 1):
    first_col_letter = get_column_letter(start_col)
    last_col_letter = get_column_letter(start_col + n_cols - 1)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=start_col).coordinate
    last_row = max(header_row + 1, n_rows)
    ws.auto_filter.ref = f"{first_col_letter}{header_row}:{last_col_letter}{last_row}"


def autosize_columns(ws: Worksheet, n_cols: int, start_col: int = 1, max_width: int = 60, min_width: int = 8):
    for c in range(start_col, start_col + n_cols):
        col_letter = get_column_letter(c)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.value is not None:
                l = len(str(cell.value).split("\n")[0])
                max_len = max(max_len, l)
        width = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = width


def apply_wrap_all(ws: Worksheet, header_row: int, n_cols: int, n_rows: int, start_col: int = 1):
    for r in range(header_row + 1, n_rows + 1):
        for c in range(start_col, start_col + n_cols):
            ws.cell(row=r, column=c).alignment = WRAP_TOP


def finalize_sheet(ws: Worksheet, header_row: int, n_cols: int, n_rows: int, start_col: int = 1):
    """헤더 스타일 + 틀고정 + 자동필터 + 줄바꿈 + 열너비 자동조정 일괄 적용."""
    style_header(ws, header_row, n_cols, start_col)
    apply_wrap_all(ws, header_row, n_cols, n_rows, start_col)
    freeze_and_filter(ws, header_row, n_cols, n_rows, start_col)
    autosize_columns(ws, n_cols, start_col)
    if n_rows > header_row:
        ws.row_dimensions[header_row].height = 32


def mark_fill(ws: Worksheet, row: int, col: int, fill: PatternFill):
    ws.cell(row=row, column=col).fill = fill


def mark_font(ws: Worksheet, row: int, col: int, font: Font):
    ws.cell(row=row, column=col).font = font
