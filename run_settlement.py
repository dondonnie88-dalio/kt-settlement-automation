"""
run_settlement.py  ─  KT 정산 자동화 시스템 v1.0
=======================================================
실행: python run_settlement.py
로그: settlement_log.txt
출력: 정산결과_YYYYMMDD.xlsx  /  대사검토_YYYYMMDD.xlsx
"""

# ══════════════════════════════════════════════════════════════════
# 0.  라이브러리 자동 설치
# ══════════════════════════════════════════════════════════════════
import subprocess, sys

for _pkg in ["pandas", "openpyxl", "xlsxwriter"]:
    try:
        __import__(_pkg)
    except ImportError:
        print(f"  [설치중] {_pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", _pkg, "-q"])

import os, re, logging, datetime, calendar
from pathlib import Path
from typing import Optional, Tuple, Dict

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════
# 1.  인보이스 컬럼 정의 (명세 기준 전체 순서)
# ══════════════════════════════════════════════════════════════════

INVOICE_COLS = [
    "일정산번호", "정산확정일", "일일정산월", "주문번호", "품목번호", "주문&품목",
    "인보이스", "사업장코드", "사업장", "사업자번호",
    "상위부서코드", "상위부서명", "부서명", "부서코드", "상품ID", "카테고리ID",
    "마스터카테고리(CMS)", "서비스카테고리",
    "관리회계(IP만변경적용)",   # ← 서비스카테고리 XLOOKUP
    "담당자",                   # ← 서비스카테고리 XLOOKUP
    "일반/통신구분",            # ← 담당자 기준 분류
    "매출구분", "상품코드", "상품명", "대표규격", "상태", "선정산여부", "단위",
    "주문수량", "주문단가", "주문금액",
    "정산수량", "정산단가", "정산금액",
    "매입단가", "매입금액", "매출총이익",
    "계정ID", "계정명",
    "주문자", "주문자ID", "주문형태", "코스트센터코드", "코스트센터명",
    "WBS코드", "WBS명", "세금코드", "SA_ID", "입고자ID", "입고자명",
    "협력사코드", "협력사명", "상품타입", "매출과세구분", "매입과세구분",
    "주문유형", "주문일시", "발주일", "출하지시일", "입고일", "승인일",
    "배송완료일", "일일정산일", "정산시점",
    "계약번호", "공사명", "발주처", "공사시작일", "공사종료일",
    "법인", "그룹", "사이트", "VAT포함여부",
    "담당자부서", "배송형태", "VMI배송구분", "픽업배송구분", "IP/DIP",
    "주문자이메일", "고객사상품코드", "상품관리자",
    "배송지", "배송지상세주소", "송장번호", "배송메모",
    "공정명", "지급상태", "발행요청상태", "발행요청일", "결제유형",
    # ── 대사 결과 추가 컬럼 (맨 끝) ──
    "플랫폼대사결과", "플랫폼정산금액", "금액차이",
]

# 금액 서식이 필요한 컬럼 세트
MONEY_COLS = {
    "주문금액", "정산금액", "매입금액", "매출총이익",
    "주문단가", "정산단가", "매입단가",
    "플랫폼정산금액", "금액차이",
}


# ══════════════════════════════════════════════════════════════════
# 2.  컬럼 매핑 테이블  (KT 파일 컬럼명 → 인보이스 컬럼명)
#     ※ 실제 KT 파일 컬럼명이 다르면 아래 딕셔너리만 수정하세요.
# ══════════════════════════════════════════════════════════════════

KT_COL_MAP: Dict[str, str] = {
    # KT 컬럼명          : 인보이스 컬럼명
    "주문번호"           : "주문번호",
    "요청번호"           : "__요청번호__",   # 내부키 (인보이스 미포함)
    "입고일"             : "입고일",
    "정산금액"           : "정산금액",
    "협력사명"           : "협력사명",
    "매출과세구분"       : "매출과세구분",
    "주문유형"           : "주문유형",
    # ── 실제 KT 파일에 추가 컬럼이 있다면 아래에 계속 추가 ──
    # "KT컬럼명"         : "인보이스컬럼명",
}


# ══════════════════════════════════════════════════════════════════
# 3.  XLOOKUP 매핑 테이블
#     ※ 실제 값으로 교체하거나 existing_report.xlsx 에서 읽어옵니다.
# ══════════════════════════════════════════════════════════════════

# 서비스카테고리 → (관리회계, 담당자)
SVCCAT_MAP: Dict[str, Tuple[str, str]] = {
    "IT서비스"   : ("IT관리회계",   "홍길동"),
    "통신장비"   : ("통신회계",     "김철수"),
    "사무용품"   : ("총무회계",     "이영희"),
    # 추가...
}

# 담당자 → 일반/통신 구분
PERSON_TYPE_MAP: Dict[str, str] = {
    "홍길동" : "일반",
    "김철수" : "통신",
    "이영희" : "일반",
    # 추가...
}

# 협력사명 → 거래처 규모  (어음 발행 판단용)
# 값:  "대기업" / "중견기업" / "중소기업"
COMPANY_SIZE_MAP: Dict[str, str] = {
    # "삼성전자" : "대기업",
    # 실제 협력사 목록으로 교체하세요.
}


# ══════════════════════════════════════════════════════════════════
# 4.  색상 팔레트
# ══════════════════════════════════════════════════════════════════

C = {
    "hdr_blue"   : "1F4E79",
    "hdr_red"    : "C00000",
    "hdr_green"  : "375623",
    "hdr_orange" : "C55A11",
    "hdr_purple" : "44336A",
    "row_red"    : "FCE4D6",
    "row_yellow" : "FFFF99",
    "row_green"  : "E2EFDA",
    "row_blue"   : "DDEEFF",
    "row_even"   : "F2F2F2",
    "ok_green"   : "70AD47",
    "ng_red"     : "FF0000",
}
FONT_NAME = "Arial"


# ══════════════════════════════════════════════════════════════════
# 5.  스타일 헬퍼 함수
# ══════════════════════════════════════════════════════════════════

def _border():
    s = Side(style="thin", color="B0B0B0")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_cell(cell, color):
    cell.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", start_color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _border()

def _data_cell(cell, fill_color=None, money=False):
    cell.font      = Font(name=FONT_NAME, size=10)
    cell.border    = _border()
    cell.alignment = Alignment(horizontal="right" if money else "center",
                               vertical="center", wrap_text=False)
    if fill_color:
        cell.fill  = PatternFill("solid", start_color=fill_color)
    if money:
        cell.number_format = '#,##0'

def write_df_to_sheet(ws, df, hdr_color,
                      money_cols=None, highlight_fn=None,
                      freeze="A2"):
    """DataFrame → 워크시트 (헤더 + 데이터 + 서식)"""
    cols = list(df.columns)
    money_cols = money_cols or set()

    # 헤더 행
    for c, name in enumerate(cols, 1):
        _hdr_cell(ws.cell(1, c, name), hdr_color)
        ws.column_dimensions[get_column_letter(c)].width = max(len(str(name)) + 2, 13)
    ws.row_dimensions[1].height = 20

    # 데이터 행
    for r, (_, row) in enumerate(df.iterrows(), 2):
        row_color = highlight_fn(row) if highlight_fn else (
            C["row_even"] if r % 2 == 0 else None
        )
        for c, col in enumerate(cols, 1):
            val  = row[col]
            is_m = col in money_cols
            cell = ws.cell(r, c, val)
            _data_cell(cell, row_color, is_m)

    if freeze:
        ws.freeze_panes = freeze

def section_header(ws, row, col, text, color):
    cell = ws.cell(row, col, text)
    cell.font = Font(name=FONT_NAME, bold=True, size=11, color=color)
    return cell

def table_header(ws, row, headers, color):
    for c, h in enumerate(headers, 1):
        _hdr_cell(ws.cell(row, c, h), color)

def money_cell(ws, row, col, val):
    c = ws.cell(row, col, val)
    c.font          = Font(name=FONT_NAME, size=10)
    c.number_format = '#,##0'
    c.alignment     = Alignment(horizontal="right", vertical="center")
    c.border        = _border()
    return c

def bold_cell(ws, row, col, val, color=None):
    c = ws.cell(row, col, val)
    c.font   = Font(name=FONT_NAME, bold=True, size=11,
                    color=color or "000000")
    c.border = _border()
    c.alignment = Alignment(horizontal="center", vertical="center")
    return c

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ══════════════════════════════════════════════════════════════════
# 6.  SettlementRunner  클래스
# ══════════════════════════════════════════════════════════════════

class SettlementRunner:

    def __init__(self, base_dir: str = "."):
        self.base_dir     = Path(base_dir)
        self.today        = datetime.date.today()
        self.log          = self._setup_logger()
        # 파일 경로
        self.kt_path: Optional[Path] = None
        self.pl_path: Optional[Path] = None
        self.er_path: Optional[Path] = None
        # 정산 기간
        self.period_start: Optional[datetime.date] = None
        self.period_end:   Optional[datetime.date] = None
        # 처리 결과 DataFrames
        self.df_kt_all       = None   # KT 원본 전체
        self.df_pl_all       = None   # 플랫폼 원본 전체
        self.df_kt_filtered  = None   # 기간 필터 후 KT
        self.df_pl_filtered  = None   # 기간 필터 후 플랫폼
        self.df_normal_kt    = None   # KT 일반 주문
        self.df_return_kt    = None   # KT 반품 주문
        self.df_matched      = None   # 양쪽 일치 (+ 금액비교)
        self.df_kt_missing   = None   # KT에만 존재 (플랫폼 누락)
        self.df_pl_only      = None   # 플랫폼에만 존재
        self.df_amount_diff  = None   # 금액 불일치
        self.df_ret_matched  = None   # 반품 정상처리
        self.df_ret_unmatch  = None   # 미확인 반품
        self.df_invoice      = None   # 인보이스 최종 DataFrame

    # ── 로거 ─────────────────────────────────────────────────────
    def _setup_logger(self):
        log = logging.getLogger("settlement")
        log.setLevel(logging.DEBUG)
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        if not log.handlers:
            ch = logging.StreamHandler(sys.stdout)
            ch.setFormatter(fmt)
            log.addHandler(ch)
            fh = logging.FileHandler("settlement_log.txt", encoding="utf-8")
            fh.setFormatter(fmt)
            log.addHandler(fh)
        return log

    # ── 파일 자동 탐지 ────────────────────────────────────────────
    def detect_files(self):
        self.log.info("=" * 65)
        self.log.info("  KT 정산 자동화 시스템  v1.0")
        self.log.info("=" * 65)

        def _find(candidates):
            for name in candidates:
                p = self.base_dir / name
                if p.exists():
                    return p
            return None

        self.kt_path = _find(["kt_raw.xlsx", "KT_raw.xlsx", "kt_raw.xlsm"])
        self.pl_path = _find(["platform.xlsx", "플랫폼.xlsx", "통합플랫폼.xlsx"])
        self.er_path = _find(["existing_report.xlsx", "기존리포트.xlsx", "매출현황.xlsx"])

        self.log.info(f"  KT 파일     : {self.kt_path or '미발견'}")
        self.log.info(f"  플랫폼 파일 : {self.pl_path or '미발견'}")
        self.log.info(f"  기존 리포트 : {self.er_path or '없음 → 표준 양식 사용'}")

        if not self.kt_path:
            fname = input("\nKT 파일명을 입력하세요 (예: kt_raw.xlsx): ").strip()
            self.kt_path = self.base_dir / fname
            if not self.kt_path.exists():
                raise FileNotFoundError(f"파일 없음: {self.kt_path}")

        if not self.pl_path:
            fname = input("플랫폼 파일명을 입력하세요 (예: platform.xlsx): ").strip()
            self.pl_path = self.base_dir / fname
            if not self.pl_path.exists():
                raise FileNotFoundError(f"파일 없음: {self.pl_path}")

    # ── STEP 0: 초기 컬럼 분석 + 매핑 확인 ──────────────────────
    def step0_init(self):
        self.log.info("\n[STEP 0] 파일 구조 분석 및 컬럼 매핑 확인")

        df_kt_peek = pd.read_excel(self.kt_path, nrows=3, dtype=str, engine="openpyxl")
        df_pl_peek = pd.read_excel(self.pl_path, nrows=3, dtype=str, engine="openpyxl")

        LINE = "─" * 60
        print(f"\n{LINE}")
        print("▶ KT 파일 컬럼 구조 및 인보이스 매핑:")
        print(f"  {'No':>3}  {'KT 컬럼명':<22}  {'→ 인보이스 컬럼'}")
        print(f"  {'─'*3}  {'─'*22}  {'─'*22}")
        for i, col in enumerate(df_kt_peek.columns, 1):
            mapped = KT_COL_MAP.get(col, f"(인보이스 미매핑 → 그대로 사용)")
            if mapped.startswith("__"):
                mapped = "(내부 키, 출력 미포함)"
            print(f"  {i:>3}  {col:<22}  → {mapped}")

        print(f"\n▶ 플랫폼 파일 컬럼 구조:")
        print(f"  {', '.join(df_pl_peek.columns)}")

        if self.er_path:
            wb_er = openpyxl.load_workbook(self.er_path, read_only=True)
            sheets = wb_er.sheetnames
            wb_er.close()
            print(f"\n▶ 기존 리포트 시트: {sheets}")
            self._load_mappings_from_er()
        else:
            print("\n▶ 기존 리포트 없음 → 시트2~4 표준 양식으로 생성")
            print("▶ XLOOKUP 매핑은 스크립트 상단 딕셔너리(SVCCAT_MAP 등)에서 수정하세요.")

        print(f"\n▶ 현재 XLOOKUP 매핑 테이블 상태:")
        print(f"  서비스카테고리 매핑 : {len(SVCCAT_MAP)}건  "
              f"{'(비어 있음 → 스크립트 상단에 추가 필요)' if not SVCCAT_MAP else ''}")
        print(f"  담당자→일반/통신   : {len(PERSON_TYPE_MAP)}건")
        print(f"  협력사 규모 목록   : {len(COMPANY_SIZE_MAP)}건  "
              f"{'(비어 있음 → 어음 발행 대상 확인 불가)' if not COMPANY_SIZE_MAP else ''}")

        print(f"\n{LINE}")
        ans = input("위 설정으로 실행하시겠습니까? (Y/n): ").strip().lower()
        if ans == "n":
            print("→ 스크립트 상단의 KT_COL_MAP / SVCCAT_MAP 등을 수정 후 재실행하세요.")
            sys.exit(0)

    def _load_mappings_from_er(self):
        """existing_report.xlsx 에서 매핑 테이블 읽기 시도"""
        global SVCCAT_MAP, PERSON_TYPE_MAP, COMPANY_SIZE_MAP
        try:
            wb = openpyxl.load_workbook(self.er_path, read_only=True)
            for sheet_name in wb.sheetnames:
                name_lower = sheet_name.lower().replace(" ", "")
                if "매핑" in sheet_name or "mapping" in name_lower:
                    self.log.info(f"  매핑 시트 발견: '{sheet_name}' → 읽기 시도")
                    df_map = pd.read_excel(self.er_path, sheet_name=sheet_name,
                                          dtype=str, engine="openpyxl")
                    cols = [c.strip() for c in df_map.columns]
                    # 서비스카테고리 매핑 탐지
                    if "서비스카테고리" in cols and "관리회계" in cols:
                        for _, r in df_map.iterrows():
                            k = str(r.get("서비스카테고리", "")).strip()
                            v1 = str(r.get("관리회계", "")).strip()
                            v2 = str(r.get("담당자", "")).strip()
                            if k:
                                SVCCAT_MAP[k] = (v1, v2)
                        self.log.info(f"    서비스카테고리 매핑 {len(SVCCAT_MAP)}건 로드")
                    # 담당자→일반/통신 탐지
                    if "담당자" in cols and ("일반/통신구분" in cols or "구분" in cols):
                        type_col = "일반/통신구분" if "일반/통신구분" in cols else "구분"
                        for _, r in df_map.iterrows():
                            k = str(r.get("담당자", "")).strip()
                            v = str(r.get(type_col, "")).strip()
                            if k:
                                PERSON_TYPE_MAP[k] = v
                        self.log.info(f"    담당자 구분 매핑 {len(PERSON_TYPE_MAP)}건 로드")
                    # 협력사 규모 탐지
                    if "협력사명" in cols and "규모" in " ".join(cols):
                        size_col = next((c for c in cols if "규모" in c), None)
                        if size_col:
                            for _, r in df_map.iterrows():
                                k = str(r.get("협력사명", "")).strip()
                                v = str(r.get(size_col, "")).strip()
                                if k:
                                    COMPANY_SIZE_MAP[k] = v
                            self.log.info(f"    협력사 규모 매핑 {len(COMPANY_SIZE_MAP)}건 로드")
            wb.close()
        except Exception as e:
            self.log.warning(f"  기존 리포트 매핑 로드 실패 ({e}) → 기본값 사용")

    # ── 정산 기간 자동 판단 ───────────────────────────────────────
    def get_period(self) -> Tuple[datetime.date, datetime.date]:
        d = self.today.day
        y, m = self.today.year, self.today.month

        if d == 16:
            start = datetime.date(y, m, 1)
            end   = datetime.date(y, m, 15)
            self.log.info(f"  자동 판단 (16일 실행): {start} ~ {end}")
        elif d == 1:
            py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
            last   = calendar.monthrange(py, pm)[1]
            start  = datetime.date(py, pm, 16)
            end    = datetime.date(py, pm, last)
            self.log.info(f"  자동 판단 (1일 실행): {start} ~ {end}")
        else:
            self.log.warning(f"  오늘({self.today})은 자동 판단 기준일이 아닙니다.")
            s = input("  정산 시작일 (YYYY-MM-DD): ").strip()
            e = input("  정산 종료일 (YYYY-MM-DD): ").strip()
            start = datetime.date.fromisoformat(s)
            end   = datetime.date.fromisoformat(e)
            self.log.info(f"  수동 입력: {start} ~ {end}")

        return start, end

    # ── STEP 1: 기간 필터링 ──────────────────────────────────────
    def step1_filter(self):
        self.log.info("\n[STEP 1] 정산 기간 필터링")
        self.period_start, self.period_end = self.get_period()
        ps, pe = self.period_start, self.period_end

        def _load(path):
            return pd.read_excel(path,
                                 dtype={"주문번호": str, "요청번호": str,
                                        "__요청번호__": str},
                                 engine="openpyxl")

        df_kt = _load(self.kt_path)
        df_pl = _load(self.pl_path)
        self.df_kt_all = df_kt.copy()
        self.df_pl_all = df_pl.copy()

        def _find_date_col(df, label):
            for c in df.columns:
                if "입고일" in c:
                    return c
            raise ValueError(f"{label} 에서 '입고일' 컬럼을 찾을 수 없습니다.")

        dc_kt = _find_date_col(df_kt, "KT")
        dc_pl = _find_date_col(df_pl, "플랫폼")

        df_kt[dc_kt] = pd.to_datetime(df_kt[dc_kt], errors="coerce").dt.date
        df_pl[dc_pl] = pd.to_datetime(df_pl[dc_pl], errors="coerce").dt.date

        self.df_kt_filtered = df_kt[
            (df_kt[dc_kt] >= ps) & (df_kt[dc_kt] <= pe)
        ].copy().reset_index(drop=True)

        self.df_pl_filtered = df_pl[
            (df_pl[dc_pl] >= ps) & (df_pl[dc_pl] <= pe)
        ].copy().reset_index(drop=True)

        self.log.info(f"  KT 전체 {len(df_kt)}건  →  기간 내 {len(self.df_kt_filtered)}건")
        self.log.info(f"  플랫폼 전체 {len(df_pl)}건  →  기간 내 {len(self.df_pl_filtered)}건")

    # ── STEP 2: 반품 선분리 ──────────────────────────────────────
    def step2_separate_returns(self):
        self.log.info("\n[STEP 2] 반품 선분리 (KT 기준)")

        df = self.df_kt_filtered
        return_col = next(
            (c for c in df.columns if c in {"반품대상", "반품여부", "주문유형"}),
            None
        )

        if not return_col:
            self.log.warning("  반품 판별 컬럼 미발견 → 전체를 일반으로 처리")
            self.df_normal_kt = df.copy()
            self.df_return_kt = df.iloc[0:0].copy()  # 빈 DataFrame
            return

        self.log.info(f"  반품 판별 컬럼: '{return_col}'")

        # 반품 판별값 자동 탐지
        vals = df[return_col].dropna().unique().tolist()
        self.log.info(f"  '{return_col}' 고유값: {vals}")

        return_vals = {"반품대상", "반품", "Y", "반품주문"}
        mask = df[return_col].isin(return_vals)

        self.df_return_kt = df[mask].copy().reset_index(drop=True)
        self.df_normal_kt = df[~mask].copy().reset_index(drop=True)

        self.log.info(f"  일반 주문 : {len(self.df_normal_kt)}건")
        self.log.info(f"  반품 주문 : {len(self.df_return_kt)}건")

    # ── STEP 3: 일반 주문 대사 ───────────────────────────────────
    def step3_reconcile_normal(self):
        self.log.info("\n[STEP 3] 일반 주문 대사 (주문번호 기준)")

        df_kt_n = self.df_normal_kt
        df_pl   = self.df_pl_filtered

        # 플랫폼에서 반품 제외
        if "주문유형" in df_pl.columns:
            pl_normal = df_pl[~df_pl["주문유형"].isin({"반품", "반품주문"})].copy()
        else:
            pl_normal = df_pl.copy()

        kt_set = set(df_kt_n["주문번호"])
        pl_set = set(pl_normal["주문번호"])

        matched_nos = kt_set & pl_set         # A: 양쪽 일치
        kt_only_nos = kt_set - pl_set         # B: KT에만 (플랫폼 누락)
        pl_only_nos = pl_set - kt_set         # C: 플랫폼에만

        # A: 일치 + 금액 비교
        df_matched = df_kt_n[df_kt_n["주문번호"].isin(matched_nos)].copy()
        pl_amt = (
            pl_normal[pl_normal["주문번호"].isin(matched_nos)]
            [["주문번호", "정산금액"]]
            .rename(columns={"정산금액": "플랫폼정산금액"})
        )
        df_matched = df_matched.merge(pl_amt, on="주문번호", how="left")
        df_matched["정산금액"]    = pd.to_numeric(df_matched["정산금액"],    errors="coerce")
        df_matched["플랫폼정산금액"] = pd.to_numeric(df_matched["플랫폼정산금액"], errors="coerce")
        df_matched["금액차이"]    = df_matched["정산금액"] - df_matched["플랫폼정산금액"]
        df_matched["플랫폼대사결과"] = df_matched["금액차이"].apply(
            lambda x: "정상" if pd.notna(x) and abs(x) <= 1 else "금액불일치"
        )
        self.df_matched     = df_matched
        self.df_amount_diff = df_matched[df_matched["플랫폼대사결과"] == "금액불일치"].copy()

        # B: 플랫폼 누락
        self.df_kt_missing = df_kt_n[df_kt_n["주문번호"].isin(kt_only_nos)].copy()

        # C: KT 미포함
        self.df_pl_only = pl_normal[pl_normal["주문번호"].isin(pl_only_nos)].copy()

        self.log.info(f"  [A] 일치          : {len(matched_nos)}건")
        self.log.info(f"  [B] 플랫폼 누락   : {len(kt_only_nos)}건  ← KT에만 존재")
        self.log.info(f"  [C] KT 미포함     : {len(pl_only_nos)}건  ← 플랫폼에만 존재")
        self.log.info(f"  금액 불일치       : {len(self.df_amount_diff)}건")

        if len(self.df_kt_missing) > 0:
            self.log.info(f"  누락 주문번호: {list(self.df_kt_missing['주문번호'])}")

    # ── STEP 4: 반품 대사 ────────────────────────────────────────
    def step4_reconcile_returns(self):
        self.log.info("\n[STEP 4] 반품 처리 (요청번호 기준)")

        df_ret = self.df_return_kt
        if len(df_ret) == 0:
            self.log.info("  반품 건 없음")
            self.df_ret_matched = df_ret.copy()
            self.df_ret_unmatch = df_ret.copy()
            return

        req_col = "요청번호"
        if req_col not in df_ret.columns:
            self.log.warning("  KT 반품에 '요청번호' 컬럼 없음 → 전체 미확인 처리")
            self.df_ret_matched = df_ret.iloc[0:0].copy()
            self.df_ret_unmatch = df_ret.copy()
            return

        pl_reqs = (
            set(self.df_pl_filtered[req_col].dropna())
            if req_col in self.df_pl_filtered.columns else set()
        )

        mask = df_ret[req_col].isin(pl_reqs)
        self.df_ret_matched = df_ret[mask].copy().reset_index(drop=True)
        self.df_ret_unmatch = df_ret[~mask].copy().reset_index(drop=True)

        self.log.info(f"  반품 정상 처리 : {len(self.df_ret_matched)}건")
        self.log.info(f"  미확인 반품    : {len(self.df_ret_unmatch)}건")
        if len(self.df_ret_matched) > 0:
            for _, r in self.df_ret_matched.iterrows():
                self.log.info(f"    {r['주문번호']}  (요청번호: {r[req_col]})")

    # ─────────────────────────────────────────────────────────────
    # 인보이스 DataFrame 조립
    # ─────────────────────────────────────────────────────────────
    def _build_invoice_df(self) -> pd.DataFrame:
        """KT 기간 내 전체 데이터 → 대사 결과 병합 → 인보이스 컬럼 순서로 정렬"""
        df = self.df_kt_filtered.copy()

        # 내부 요청번호 컬럼 임시 보존
        req_tmp = None
        if "요청번호" in df.columns:
            req_tmp = df["요청번호"].copy()

        # ── KT 컬럼 → 인보이스 컬럼명 변환 ──
        rename_map = {k: v for k, v in KT_COL_MAP.items()
                      if not v.startswith("__") and k in df.columns}
        df = df.rename(columns=rename_map)

        # ── 대사 결과 컬럼 초기화 ──
        df["플랫폼대사결과"] = "미확인"
        df["플랫폼정산금액"] = None
        df["금액차이"]       = None

        # 일치 건 반영
        if self.df_matched is not None and len(self.df_matched) > 0:
            for _, r in self.df_matched.iterrows():
                mask = df["주문번호"] == r["주문번호"]
                df.loc[mask, "플랫폼대사결과"] = r["플랫폼대사결과"]
                df.loc[mask, "플랫폼정산금액"] = r.get("플랫폼정산금액")
                df.loc[mask, "금액차이"]       = r.get("금액차이")

        # 플랫폼 누락 반영
        if self.df_kt_missing is not None and len(self.df_kt_missing) > 0:
            miss_nos = set(self.df_kt_missing["주문번호"])
            df.loc[df["주문번호"].isin(miss_nos), "플랫폼대사결과"] = "플랫폼누락"

        # 반품 정상 반영
        if self.df_ret_matched is not None and len(self.df_ret_matched) > 0:
            ret_nos = set(self.df_ret_matched["주문번호"])
            df.loc[df["주문번호"].isin(ret_nos), "플랫폼대사결과"] = "반품"

        # ── XLOOKUP: 서비스카테고리 → 관리회계 / 담당자 ──
        if "서비스카테고리" in df.columns and SVCCAT_MAP:
            df["관리회계(IP만변경적용)"] = df["서비스카테고리"].map(
                lambda x: SVCCAT_MAP.get(str(x), (None, None))[0]
            )
            df["담당자"] = df["서비스카테고리"].map(
                lambda x: SVCCAT_MAP.get(str(x), (None, None))[1]
            )
        else:
            if "관리회계(IP만변경적용)" not in df.columns:
                df["관리회계(IP만변경적용)"] = None
            if "담당자" not in df.columns:
                df["담당자"] = None

        # ── XLOOKUP: 담당자 → 일반/통신 구분 ──
        if "담당자" in df.columns and PERSON_TYPE_MAP:
            df["일반/통신구분"] = df["담당자"].map(PERSON_TYPE_MAP)
        else:
            if "일반/통신구분" not in df.columns:
                df["일반/통신구분"] = None

        # ── 인보이스 컬럼 순서로 재정렬 (없는 컬럼 → None) ──
        for col in INVOICE_COLS:
            if col not in df.columns:
                df[col] = None

        # 내부키 제거 후 출력 컬럼 선택
        out_cols = [c for c in INVOICE_COLS if c in df.columns]
        return df[out_cols].reset_index(drop=True)

    # ─────────────────────────────────────────────────────────────
    # 시트1: 인보이스
    # ─────────────────────────────────────────────────────────────
    def build_invoice_sheet(self, wb):
        self.log.info("  [시트1] 인보이스 작성...")
        ws = wb.active
        ws.title = "①인보이스"

        self.df_invoice = self._build_invoice_df()
        df = self.df_invoice

        def hl(row):
            result = row.get("플랫폼대사결과", "")
            if result == "플랫폼누락":
                return C["row_red"]
            diff = row.get("금액차이", None)
            if pd.notna(diff):
                try:
                    if abs(float(diff)) > 1:
                        return C["row_yellow"]
                except (ValueError, TypeError):
                    pass
            return C["row_even"] if True else None  # 교차색은 write_df_to_sheet 내부서 처리

        # 헤더 쓰기
        cols = list(df.columns)
        for c, name in enumerate(cols, 1):
            _hdr_cell(ws.cell(1, c, name), C["hdr_blue"])
            ws.column_dimensions[get_column_letter(c)].width = max(len(str(name)) + 2, 12)
        ws.row_dimensions[1].height = 20

        # 데이터 쓰기
        for r, (_, row) in enumerate(df.iterrows(), 2):
            result = row.get("플랫폼대사결과", "")
            diff   = row.get("금액차이", None)
            if result == "플랫폼누락":
                row_color = C["row_red"]
            elif pd.notna(diff):
                try:
                    row_color = C["row_yellow"] if abs(float(diff)) > 1 else (
                        C["row_even"] if r % 2 == 0 else None
                    )
                except (ValueError, TypeError):
                    row_color = C["row_even"] if r % 2 == 0 else None
            else:
                row_color = C["row_even"] if r % 2 == 0 else None

            for c, col in enumerate(cols, 1):
                is_m = col in MONEY_COLS
                cell = ws.cell(r, c, row[col])
                _data_cell(cell, row_color, is_m)

        ws.freeze_panes = "D2"  # 주문번호까지 고정
        self.log.info(f"    → {len(df)}행 완료")

    # ─────────────────────────────────────────────────────────────
    # 시트2: 매출현황
    # ─────────────────────────────────────────────────────────────
    def build_sales_sheet(self, wb):
        self.log.info("  [시트2] 매출현황 작성...")
        ws = wb.create_sheet("②매출현황")
        df = self.df_invoice.copy() if self.df_invoice is not None else pd.DataFrame()

        if df.empty or "정산금액" not in df.columns:
            ws["A1"] = "데이터 없음"
            return

        df["정산금액"] = pd.to_numeric(df["정산금액"], errors="coerce").fillna(0)
        amt = "정산금액"

        def _block(ws, start_row, title, grp_col, grp_df, color):
            """집계 블록 한 개 그리기"""
            section_header(ws, start_row, 1, title, color)
            r = start_row + 1
            hdrs = [grp_col, "건수", "정산금액 합계"]
            table_header(ws, r, hdrs, color)
            r += 1
            for _, row in grp_df.iterrows():
                ws.cell(r, 1, row[grp_col]).font = Font(name=FONT_NAME, size=10)
                ws.cell(r, 2, int(row["건수"])).font = Font(name=FONT_NAME, size=10)
                money_cell(ws, r, 3, row["합계"])
                r += 1
            # 소계
            ws.cell(r, 1, "소계").font = Font(name=FONT_NAME, bold=True)
            ws.cell(r, 2, int(grp_df["건수"].sum())).font = Font(name=FONT_NAME, bold=True)
            c3 = money_cell(ws, r, 3, grp_df["합계"].sum())
            c3.font = Font(name=FONT_NAME, bold=True)
            r += 2
            return r

        row = 1
        # ── 집계1: 매출과세구분별 ──
        if "매출과세구분" in df.columns:
            g = df.groupby("매출과세구분")[amt].agg(
                건수="count", 합계="sum"
            ).reset_index()
            row = _block(ws, row, "▶ 매출과세구분별 집계", "매출과세구분", g, C["hdr_blue"])

        # ── 집계2: 일반/통신 구분별 ──
        if "일반/통신구분" in df.columns and df["일반/통신구분"].notna().any():
            g2 = df.groupby("일반/통신구분")[amt].agg(
                건수="count", 합계="sum"
            ).reset_index()
            row = _block(ws, row, "▶ 일반/통신 구분별 집계", "일반/통신구분", g2, C["hdr_green"])
        else:
            ws.cell(row, 1, "▶ 일반/통신 구분별 — 매핑 테이블 입력 필요 (PERSON_TYPE_MAP)")
            ws.cell(row, 1).font = Font(name=FONT_NAME, bold=True, color="FF6600")
            row += 2

        # ── 집계3: 담당자별 ──
        if "담당자" in df.columns and df["담당자"].notna().any():
            g3 = df.groupby("담당자")[amt].agg(
                건수="count", 합계="sum"
            ).reset_index().sort_values("합계", ascending=False)
            row = _block(ws, row, "▶ 담당자별 집계", "담당자", g3, C["hdr_orange"])
        else:
            ws.cell(row, 1, "▶ 담당자별 — SVCCAT_MAP 매핑 후 재실행 필요")
            ws.cell(row, 1).font = Font(name=FONT_NAME, bold=True, color="FF6600")
            row += 2

        # ── 집계4: 협력사(거래처)별 ──
        if "협력사명" in df.columns:
            g4 = df.groupby("협력사명")[amt].agg(
                건수="count", 합계="sum"
            ).reset_index().sort_values("합계", ascending=False)
            row = _block(ws, row, "▶ 협력사별 매출 집계", "협력사명", g4, C["hdr_purple"])

        set_col_widths(ws, [26, 10, 22])

    # ─────────────────────────────────────────────────────────────
    # 시트3: 어음확인_최종
    # ─────────────────────────────────────────────────────────────
    def build_bill_sheet(self, wb):
        self.log.info("  [시트3] 어음확인_최종 작성...")
        ws = wb.create_sheet("③어음확인_최종")
        df = self.df_invoice.copy() if self.df_invoice is not None else pd.DataFrame()

        if df.empty:
            ws["A1"] = "데이터 없음"
            return

        df["정산금액"] = pd.to_numeric(df.get("정산금액", pd.Series(dtype=float)),
                                       errors="coerce").fillna(0)

        # ── 비과세 주문 추출 ──
        tax_col = "매출과세구분"
        if tax_col in df.columns:
            df_nontax = df[df[tax_col] == "비과세"].copy()
        else:
            df_nontax = pd.DataFrame()
            ws["A1"] = "매출과세구분 컬럼 없음 — 인보이스 매핑 확인 필요"
            ws["A1"].font = Font(name=FONT_NAME, color="FF0000", bold=True)

        DISPLAY = ["주문번호", "협력사명", "합산금액", "거래처규모",
                   "플랫폼대사결과", "발행조건_충족", "발행여부"]

        if not df_nontax.empty and "협력사명" in df_nontax.columns:
            grp = df_nontax.groupby(["주문번호", "협력사명"]).agg(
                합산금액   = ("정산금액", "sum"),
                대사결과   = ("플랫폼대사결과", "first"),
            ).reset_index()
            grp = grp.rename(columns={"대사결과": "플랫폼대사결과"})
            grp["거래처규모"]    = grp["협력사명"].map(COMPANY_SIZE_MAP).fillna("확인필요")
            grp["조건_금액"]     = grp["합산금액"] >= 200_000_000
            grp["조건_규모"]     = grp["거래처규모"].isin(["중견기업", "대기업"])
            grp["조건_대사"]     = grp["플랫폼대사결과"] == "정상"
            grp["발행조건_충족"] = grp[["조건_금액","조건_규모","조건_대사"]].apply(
                lambda r: f"{'O' if r['조건_금액'] else 'X'} {'O' if r['조건_규모'] else 'X'} {'O' if r['조건_대사'] else 'X'}",
                axis=1
            )
            grp["발행여부"] = grp.apply(
                lambda r: "발행대상"  if (r["조건_금액"] and r["조건_규모"] and r["조건_대사"])
                     else ("확인필요" if (r["조건_금액"] and not r["조건_대사"])
                     else "미해당"),
                axis=1
            )
        else:
            grp = pd.DataFrame(columns=DISPLAY)

        df_issue = grp[grp["발행여부"] == "발행대상"]
        df_check = grp[grp["발행여부"] == "확인필요"]
        df_no    = grp[grp["발행여부"] == "미해당"]

        row = 1
        blocks = [
            ("▶ 어음 발행 대상 (3조건 모두 충족)",   df_issue, C["hdr_green"]),
            ("▶ 확인필요 목록 (금액충족 + 대사오류)", df_check, C["hdr_red"]),
            ("▶ 비과세 — 미해당 (조건 미충족)",       df_no,    C["hdr_blue"]),
        ]
        for title, sub, color in blocks:
            ws.cell(row, 1, title).font = Font(name=FONT_NAME, bold=True,
                                               size=11, color=color)
            row += 1
            # 범례 행
            legend = ["주문번호", "협력사명", "합산금액(원)",
                      "거래처규모", "플랫폼대사결과",
                      "조건(금액/규모/대사)", "발행여부"]
            for c, h in enumerate(legend, 1):
                _hdr_cell(ws.cell(row, c, h), color)
            row += 1

            if sub.empty:
                ws.cell(row, 1, "해당 없음").font = Font(name=FONT_NAME,
                                                         italic=True, color="808080")
                row += 1
            else:
                for _, r in sub.iterrows():
                    ws.cell(row, 1, r["주문번호"]).font  = Font(name=FONT_NAME, size=10)
                    ws.cell(row, 2, r["협력사명"]).font  = Font(name=FONT_NAME, size=10)
                    money_cell(ws, row, 3, r["합산금액"])
                    ws.cell(row, 4, r["거래처규모"]).font = Font(name=FONT_NAME, size=10)
                    ws.cell(row, 5, r["플랫폼대사결과"]).font = Font(name=FONT_NAME, size=10)
                    ws.cell(row, 6, r["발행조건_충족"]).font = Font(name=FONT_NAME, size=10)
                    result_color = (C["ok_green"] if r["발행여부"] == "발행대상"
                                    else C["ng_red"] if r["발행여부"] == "확인필요"
                                    else "808080")
                    ws.cell(row, 7, r["발행여부"]).font = Font(
                        name=FONT_NAME, bold=True, color=result_color
                    )
                    row += 1
            row += 1

        # 조건 범례 주석
        ws.cell(row + 1, 1, "※ 발행 조건: ①비과세합산 ≥ 2억  ②거래처규모=중견/대기업  ③플랫폼대사='정상'")
        ws.cell(row + 1, 1).font = Font(name=FONT_NAME, italic=True, color="595959", size=9)

        set_col_widths(ws, [22, 16, 20, 14, 16, 18, 12])

    # ─────────────────────────────────────────────────────────────
    # 시트4: 재무회계팀 자금예측보고
    # ─────────────────────────────────────────────────────────────
    def build_forecast_sheet(self, wb):
        self.log.info("  [시트4] 재무회계팀 자금예측보고 작성...")
        ws = wb.create_sheet("④자금예측보고")
        df = self.df_invoice.copy() if self.df_invoice is not None else pd.DataFrame()

        ps = self.period_start.strftime("%Y.%m.%d") if self.period_start else "-"
        pe = self.period_end.strftime("%Y.%m.%d")   if self.period_end   else "-"

        # 타이틀
        ws.merge_cells("A1:H1")
        tc = ws["A1"]
        tc.value     = f"재무회계팀  자금예측보고    정산기간: {ps} ~ {pe}"
        tc.font      = Font(name=FONT_NAME, bold=True, size=14, color="1F4E79")
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        if df.empty or "정산금액" not in df.columns:
            ws["A3"] = "데이터 없음"
            return

        df["정산금액"] = pd.to_numeric(df["정산금액"], errors="coerce").fillna(0)
        tax_col = "매출과세구분"
        if tax_col not in df.columns:
            df[tax_col] = "미분류"

        # 예측 수금일: 정산기간 종료일 기준
        if self.period_end:
            d_cash = (self.period_end + datetime.timedelta(days=30)).strftime("%Y.%m.%d")
            d_bill = (self.period_end + datetime.timedelta(days=90)).strftime("%Y.%m.%d")
        else:
            d_cash = d_bill = "-"

        # ── 과세구분별 집계 ──
        grp = df.groupby(tax_col)["정산금액"].agg(
            건수="count", 합계="sum"
        ).reset_index()

        # ── 어음 발행 대상 금액 (비과세 중 정상대사 2억↑) ──
        bill_mask = (
            (df[tax_col] == "비과세") &
            (df["정산금액"] >= 200_000_000) &
            (df.get("플랫폼대사결과", pd.Series("미확인", index=df.index)) == "정상")
        )
        bill_total = df[bill_mask]["정산금액"].sum()
        cash_total = df["정산금액"].sum() - bill_total

        HDR = ["구분", "과세구분", "건수", "금액합계(원)", "수금방식",
               "예측수금일", "비고"]
        row = 3
        for c, h in enumerate(HDR, 1):
            _hdr_cell(ws.cell(row, c, h), C["hdr_blue"])
        ws.row_dimensions[row].height = 18
        row += 1

        for _, r in grp.iterrows():
            tax = r[tax_col]
            is_bill = (tax == "비과세" and bill_total > 0)
            method  = "어음"  if is_bill else "현금"
            exp_d   = d_bill  if is_bill else d_cash
            gubun   = "어음수금" if is_bill else "현금수금"
            note    = f"발행대상 {bill_total:,.0f}원" if is_bill else ""

            ws.cell(row, 1, gubun).font  = Font(name=FONT_NAME, size=10, bold=True)
            ws.cell(row, 2, tax).font    = Font(name=FONT_NAME, size=10)
            ws.cell(row, 3, int(r["건수"])).font = Font(name=FONT_NAME, size=10)
            money_cell(ws, row, 4, r["합계"])
            ws.cell(row, 5, method).font = Font(name=FONT_NAME, size=10)
            ws.cell(row, 6, exp_d).font  = Font(name=FONT_NAME, size=10)
            ws.cell(row, 7, note).font   = Font(name=FONT_NAME, size=10, italic=True,
                                                 color="595959")
            fill = PatternFill("solid", start_color=C["row_blue"] if is_bill
                               else (C["row_even"] if row % 2 == 0 else "FFFFFF"))
            for c in range(1, 8):
                ws.cell(row, c).fill = fill
            row += 1

        # 합계 행
        ws.cell(row, 1, "합계").font     = Font(name=FONT_NAME, bold=True, size=11)
        ws.cell(row, 3, int(grp["건수"].sum())).font = Font(name=FONT_NAME, bold=True)
        c4 = money_cell(ws, row, 4, grp["합계"].sum())
        c4.font = Font(name=FONT_NAME, bold=True)

        row += 2
        # 현금/어음 분리 요약
        ws.cell(row, 1, "현금 수금 예정").font = Font(name=FONT_NAME, bold=True, color="1F4E79")
        money_cell(ws, row, 2, cash_total)
        ws.cell(row, 3, d_cash).font = Font(name=FONT_NAME, size=10)
        row += 1
        ws.cell(row, 1, "어음 수금 예정").font = Font(name=FONT_NAME, bold=True, color="375623")
        money_cell(ws, row, 2, bill_total)
        ws.cell(row, 3, d_bill).font = Font(name=FONT_NAME, size=10)

        set_col_widths(ws, [14, 14, 8, 22, 10, 14, 26])

    # ─────────────────────────────────────────────────────────────
    # 대사검토_YYYYMMDD.xlsx  (시트A~E)
    # ─────────────────────────────────────────────────────────────
    def build_review_file(self, path: str):
        self.log.info(f"\n[출력] 대사검토 파일: {path}")
        wb = openpyxl.Workbook()
        MONEY = MONEY_COLS

        # ── 시트A: 플랫폼 누락 ──
        ws_a = wb.active; ws_a.title = "A_플랫폼누락"
        if self.df_kt_missing is not None and len(self.df_kt_missing) > 0:
            write_df_to_sheet(ws_a, self.df_kt_missing, C["hdr_red"],
                              money_cols=MONEY)
        else:
            ws_a["A1"] = "플랫폼 누락 주문 없음"
            ws_a["A1"].font = Font(name=FONT_NAME, color=C["ok_green"], bold=True)

        # ── 시트B: 반품 처리 결과 ──
        ws_b = wb.create_sheet("B_반품처리결과")
        if self.df_ret_matched is not None and len(self.df_ret_matched) > 0:
            write_df_to_sheet(ws_b, self.df_ret_matched, C["hdr_green"],
                              money_cols=MONEY)
        else:
            ws_b["A1"] = "반품 정상처리 건 없음"
            ws_b["A1"].font = Font(name=FONT_NAME, italic=True)

        # ── 시트C: 금액 불일치 ──
        ws_c = wb.create_sheet("C_금액불일치")
        if self.df_amount_diff is not None and len(self.df_amount_diff) > 0:
            write_df_to_sheet(ws_c, self.df_amount_diff, C["hdr_orange"],
                              money_cols=MONEY)
        else:
            ws_c["A1"] = "금액 불일치 없음"
            ws_c["A1"].font = Font(name=FONT_NAME, color=C["ok_green"], bold=True)

        # ── 시트D: 미확인 반품 ──
        ws_d = wb.create_sheet("D_미확인반품")
        if self.df_ret_unmatch is not None and len(self.df_ret_unmatch) > 0:
            write_df_to_sheet(ws_d, self.df_ret_unmatch, C["hdr_purple"],
                              money_cols=MONEY)
        else:
            ws_d["A1"] = "미확인 반품 없음"
            ws_d["A1"].font = Font(name=FONT_NAME, color=C["ok_green"], bold=True)

        # ── 시트E: 전체 요약 ──
        ws_e = wb.create_sheet("E_전체요약")
        ps = self.period_start.strftime("%Y.%m.%d") if self.period_start else "-"
        pe = self.period_end.strftime("%Y.%m.%d")   if self.period_end   else "-"

        ws_e.merge_cells("A1:C1")
        ws_e["A1"].value = f"정산 대사 요약  |  {ps} ~ {pe} 입고분"
        ws_e["A1"].font  = Font(name=FONT_NAME, bold=True, size=14, color="1F4E79")
        ws_e["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_e.row_dimensions[1].height = 32

        summary = [
            ("KT 총 건수",             len(self.df_kt_filtered)  if self.df_kt_filtered  is not None else 0, False),
            ("플랫폼 총 건수",         len(self.df_pl_filtered)  if self.df_pl_filtered  is not None else 0, False),
            ("일반 주문 일치",         len(self.df_matched)      if self.df_matched       is not None else 0, False),
            ("플랫폼 누락 (B)",        len(self.df_kt_missing)   if self.df_kt_missing    is not None else 0, True),
            ("KT 미포함 (C)",          len(self.df_pl_only)      if self.df_pl_only       is not None else 0, True),
            ("금액 불일치",            len(self.df_amount_diff)  if self.df_amount_diff   is not None else 0, True),
            ("반품 정상처리",          len(self.df_ret_matched)  if self.df_ret_matched   is not None else 0, False),
            ("미확인 반품",            len(self.df_ret_unmatch)  if self.df_ret_unmatch   is not None else 0, True),
        ]

        for r, (label, cnt, is_bad) in enumerate(summary, 3):
            ws_e.cell(r, 1, label).font = Font(name=FONT_NAME, bold=True, size=11)
            ws_e.row_dimensions[r].height = 24

            cnt_cell = ws_e.cell(r, 2, cnt)
            cnt_cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_bad:
                if cnt > 0:
                    cnt_cell.font = Font(name=FONT_NAME, bold=True, size=16, color=C["ng_red"])
                    cnt_cell.fill = PatternFill("solid", start_color=C["row_red"])
                    ws_e.cell(r, 3, "▲ 확인 필요").font = Font(
                        name=FONT_NAME, bold=True, color=C["ng_red"]
                    )
                else:
                    cnt_cell.font = Font(name=FONT_NAME, bold=True, size=16, color=C["ok_green"])
                    cnt_cell.fill = PatternFill("solid", start_color=C["row_green"])
                    ws_e.cell(r, 3, "정상").font = Font(
                        name=FONT_NAME, color=C["ok_green"]
                    )
            else:
                cnt_cell.font = Font(name=FONT_NAME, bold=True, size=16, color="1F4E79")

        set_col_widths(ws_e, [26, 12, 16])

        wb.save(path)
        self.log.info("  → 저장 완료")

    # ─────────────────────────────────────────────────────────────
    # 메인 실행
    # ─────────────────────────────────────────────────────────────
    def run(self):
        try:
            self.detect_files()
            self.step0_init()
            self.step1_filter()
            self.step2_separate_returns()
            self.step3_reconcile_normal()
            self.step4_reconcile_returns()

            today_str    = self.today.strftime("%Y%m%d")
            result_path  = f"정산결과_{today_str}.xlsx"
            review_path  = f"대사검토_{today_str}.xlsx"

            # ─── 정산결과 파일 ───
            self.log.info(f"\n[출력] 정산결과 파일: {result_path}")
            wb = openpyxl.Workbook()
            self.build_invoice_sheet(wb)
            self.build_sales_sheet(wb)
            self.build_bill_sheet(wb)
            self.build_forecast_sheet(wb)
            wb.save(result_path)
            self.log.info("  → 저장 완료")

            # ─── 대사검토 파일 ───
            self.build_review_file(review_path)

            # ─── 최종 요약 ───
            n_miss = len(self.df_kt_missing)   if self.df_kt_missing   is not None else 0
            n_diff = len(self.df_amount_diff)  if self.df_amount_diff  is not None else 0
            n_unmt = len(self.df_ret_unmatch)  if self.df_ret_unmatch  is not None else 0

            self.log.info("")
            self.log.info("=" * 65)
            self.log.info("  실행 완료")
            self.log.info(f"  정산결과  : {result_path}")
            self.log.info(f"  대사검토  : {review_path}")
            self.log.info(f"  로그파일  : settlement_log.txt")
            self.log.info("-" * 65)
            status = "이상 없음" if (n_miss + n_diff + n_unmt) == 0 else "확인 필요"
            self.log.info(f"  플랫폼 누락   : {n_miss}건")
            self.log.info(f"  금액 불일치   : {n_diff}건")
            self.log.info(f"  미확인 반품   : {n_unmt}건")
            self.log.info(f"  → 종합 상태  : {status}")
            self.log.info("=" * 65)

        except Exception as e:
            self.log.error(f"오류 발생: {e}", exc_info=True)
            raise


# ══════════════════════════════════════════════════════════════════
# 7.  Entry Point
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    SettlementRunner(base_dir=".").run()
