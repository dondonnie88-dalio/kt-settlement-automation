"""
settlement_web.py  ─  웹 포털용 비대화형 정산 실행기
SettlementRunner 의 interactive 프롬프트를 제거하고
웹에서 파라미터를 받아 실행 후 결과 딕셔너리 반환.
"""

import os, logging, datetime, calendar
from pathlib import Path
from typing import Optional, Tuple, Dict

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 인보이스 컬럼 (run_settlement.py 와 동일) ──────────────────────
INVOICE_COLS = [
    "일정산번호", "정산확정일", "일일정산월", "주문번호", "품목번호", "주문&품목",
    "인보이스", "사업장코드", "사업장", "사업자번호",
    "상위부서코드", "상위부서명", "부서명", "부서코드", "상품ID", "카테고리ID",
    "마스터카테고리(CMS)", "서비스카테고리",
    "관리회계(IP만변경적용)", "담당자", "일반/통신구분",
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
    "플랫폼대사결과", "플랫폼정산금액", "금액차이",
]

MONEY_COLS = {
    "주문금액", "정산금액", "매입금액", "매출총이익",
    "주문단가", "정산단가", "매입단가",
    "플랫폼정산금액", "금액차이",
}

KT_COL_MAP: Dict[str, str] = {
    "주문번호"     : "주문번호",
    "요청번호"     : "__요청번호__",
    "입고일"       : "입고일",
    "정산금액"     : "정산금액",
    "협력사명"     : "협력사명",
    "매출과세구분" : "매출과세구분",
    "주문유형"     : "주문유형",
}

SVCCAT_MAP: Dict[str, Tuple[str, str]] = {
    "IT서비스" : ("IT관리회계", "홍길동"),
    "통신장비"  : ("통신회계",  "김철수"),
    "사무용품"  : ("총무회계",  "이영희"),
}

PERSON_TYPE_MAP: Dict[str, str] = {
    "홍길동": "일반",
    "김철수": "통신",
    "이영희": "일반",
}

COMPANY_SIZE_MAP: Dict[str, str] = {}

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


# ── 스타일 헬퍼 ───────────────────────────────────────────────────
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

def write_df_to_sheet(ws, df, hdr_color, money_cols=None, freeze="A2"):
    cols = list(df.columns)
    money_cols = money_cols or set()
    for c, name in enumerate(cols, 1):
        _hdr_cell(ws.cell(1, c, name), hdr_color)
        ws.column_dimensions[get_column_letter(c)].width = max(len(str(name)) + 2, 13)
    ws.row_dimensions[1].height = 20
    for r, (_, row) in enumerate(df.iterrows(), 2):
        row_color = C["row_even"] if r % 2 == 0 else None
        for c, col in enumerate(cols, 1):
            cell = ws.cell(r, c, row[col])
            _data_cell(cell, row_color, col in money_cols)
    if freeze:
        ws.freeze_panes = freeze

def money_cell(ws, row, col, val):
    c = ws.cell(row, col, val)
    c.font          = Font(name=FONT_NAME, size=10)
    c.number_format = '#,##0'
    c.alignment     = Alignment(horizontal="right", vertical="center")
    c.border        = _border()
    return c

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── WebSettlementRunner ────────────────────────────────────────────
class WebSettlementRunner:
    """
    웹 포털 전용 비대화형 정산 실행기.
    모든 파라미터를 생성자에서 받고 stdin 호출 없음.
    """

    def __init__(self,
                 kt_path: str,
                 pl_path: str,
                 period_start: datetime.date,
                 period_end: datetime.date,
                 output_dir: str):
        self.kt_path      = Path(kt_path)
        self.pl_path      = Path(pl_path)
        self.period_start = period_start
        self.period_end   = period_end
        self.output_dir   = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.today        = datetime.date.today()

        # 로거 (파일만; stdout은 Flask가 처리)
        self.log = logging.getLogger(f"settlement.{output_dir}")
        self.log.setLevel(logging.DEBUG)
        if not self.log.handlers:
            fh = logging.FileHandler(self.output_dir / "settlement_log.txt", encoding="utf-8")
            fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                              datefmt="%Y-%m-%d %H:%M:%S"))
            self.log.addHandler(fh)

        # 결과 DataFrames
        self.df_kt_filtered = None
        self.df_pl_filtered = None
        self.df_normal_kt   = None
        self.df_return_kt   = None
        self.df_matched     = None
        self.df_kt_missing  = None
        self.df_pl_only     = None
        self.df_amount_diff = None
        self.df_ret_matched = None
        self.df_ret_unmatch = None
        self.df_invoice     = None

    # ── STEP 1 ────────────────────────────────────────────────────
    def step1_filter(self):
        def _load(path):
            return pd.read_excel(path,
                                 dtype={"주문번호": str, "요청번호": str},
                                 engine="openpyxl")

        df_kt = _load(self.kt_path)
        df_pl = _load(self.pl_path)

        def _find_date_col(df, label):
            for c in df.columns:
                if "입고일" in c:
                    return c
            raise ValueError(f"{label} 파일에서 '입고일' 컬럼을 찾을 수 없습니다.")

        dc_kt = _find_date_col(df_kt, "KT")
        dc_pl = _find_date_col(df_pl, "플랫폼")

        df_kt[dc_kt] = pd.to_datetime(df_kt[dc_kt], errors="coerce").dt.date
        df_pl[dc_pl] = pd.to_datetime(df_pl[dc_pl], errors="coerce").dt.date

        ps, pe = self.period_start, self.period_end
        self.df_kt_filtered = df_kt[
            (df_kt[dc_kt] >= ps) & (df_kt[dc_kt] <= pe)
        ].copy().reset_index(drop=True)

        self.df_pl_filtered = df_pl[
            (df_pl[dc_pl] >= ps) & (df_pl[dc_pl] <= pe)
        ].copy().reset_index(drop=True)

    # ── STEP 2 ────────────────────────────────────────────────────
    def step2_separate_returns(self):
        df = self.df_kt_filtered
        return_col = next(
            (c for c in df.columns if c in {"반품대상", "반품여부", "주문유형"}), None
        )
        if not return_col:
            self.df_normal_kt = df.copy()
            self.df_return_kt = df.iloc[0:0].copy()
            return

        return_vals = {"반품대상", "반품", "Y", "반품주문"}
        mask = df[return_col].isin(return_vals)
        self.df_return_kt = df[mask].copy().reset_index(drop=True)
        self.df_normal_kt = df[~mask].copy().reset_index(drop=True)

    # ── STEP 3 ────────────────────────────────────────────────────
    def step3_reconcile_normal(self):
        df_kt_n = self.df_normal_kt
        df_pl   = self.df_pl_filtered

        if "주문유형" in df_pl.columns:
            pl_normal = df_pl[~df_pl["주문유형"].isin({"반품", "반품주문"})].copy()
        else:
            pl_normal = df_pl.copy()

        kt_set = set(df_kt_n["주문번호"])
        pl_set = set(pl_normal["주문번호"])

        matched_nos = kt_set & pl_set
        kt_only_nos = kt_set - pl_set
        pl_only_nos = pl_set - kt_set

        df_matched = df_kt_n[df_kt_n["주문번호"].isin(matched_nos)].copy()
        pl_amt = (
            pl_normal[pl_normal["주문번호"].isin(matched_nos)]
            [["주문번호", "정산금액"]]
            .rename(columns={"정산금액": "플랫폼정산금액"})
        )
        df_matched = df_matched.merge(pl_amt, on="주문번호", how="left")
        df_matched["정산금액"]       = pd.to_numeric(df_matched["정산금액"],       errors="coerce")
        df_matched["플랫폼정산금액"] = pd.to_numeric(df_matched["플랫폼정산금액"], errors="coerce")
        df_matched["금액차이"]       = df_matched["정산금액"] - df_matched["플랫폼정산금액"]
        df_matched["플랫폼대사결과"] = df_matched["금액차이"].apply(
            lambda x: "정상" if pd.notna(x) and abs(x) <= 1 else "금액불일치"
        )

        self.df_matched     = df_matched
        self.df_amount_diff = df_matched[df_matched["플랫폼대사결과"] == "금액불일치"].copy()
        self.df_kt_missing  = df_kt_n[df_kt_n["주문번호"].isin(kt_only_nos)].copy()
        self.df_pl_only     = pl_normal[pl_normal["주문번호"].isin(pl_only_nos)].copy()

    # ── STEP 4 ────────────────────────────────────────────────────
    def step4_reconcile_returns(self):
        df_ret = self.df_return_kt
        if len(df_ret) == 0:
            self.df_ret_matched = df_ret.copy()
            self.df_ret_unmatch = df_ret.copy()
            return

        req_col = "요청번호"
        if req_col not in df_ret.columns:
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

    # ── 인보이스 조립 ─────────────────────────────────────────────
    def _build_invoice_df(self) -> pd.DataFrame:
        df = self.df_kt_filtered.copy()

        rename_map = {k: v for k, v in KT_COL_MAP.items()
                      if not v.startswith("__") and k in df.columns}
        df = df.rename(columns=rename_map)

        df["플랫폼대사결과"] = "미확인"
        df["플랫폼정산금액"] = None
        df["금액차이"]       = None

        if self.df_matched is not None and len(self.df_matched) > 0:
            for _, r in self.df_matched.iterrows():
                mask = df["주문번호"] == r["주문번호"]
                df.loc[mask, "플랫폼대사결과"] = r["플랫폼대사결과"]
                df.loc[mask, "플랫폼정산금액"] = r.get("플랫폼정산금액")
                df.loc[mask, "금액차이"]       = r.get("금액차이")

        if self.df_kt_missing is not None and len(self.df_kt_missing) > 0:
            miss_nos = set(self.df_kt_missing["주문번호"])
            df.loc[df["주문번호"].isin(miss_nos), "플랫폼대사결과"] = "플랫폼누락"

        if self.df_ret_matched is not None and len(self.df_ret_matched) > 0:
            ret_nos = set(self.df_ret_matched["주문번호"])
            df.loc[df["주문번호"].isin(ret_nos), "플랫폼대사결과"] = "반품"

        if "서비스카테고리" in df.columns and SVCCAT_MAP:
            df["관리회계(IP만변경적용)"] = df["서비스카테고리"].map(
                lambda x: SVCCAT_MAP.get(str(x), (None, None))[0])
            df["담당자"] = df["서비스카테고리"].map(
                lambda x: SVCCAT_MAP.get(str(x), (None, None))[1])
        else:
            if "관리회계(IP만변경적용)" not in df.columns:
                df["관리회계(IP만변경적용)"] = None
            if "담당자" not in df.columns:
                df["담당자"] = None

        if "담당자" in df.columns and PERSON_TYPE_MAP:
            df["일반/통신구분"] = df["담당자"].map(PERSON_TYPE_MAP)
        else:
            if "일반/통신구분" not in df.columns:
                df["일반/통신구분"] = None

        for col in INVOICE_COLS:
            if col not in df.columns:
                df[col] = None

        out_cols = [c for c in INVOICE_COLS if c in df.columns]
        return df[out_cols].reset_index(drop=True)

    # ── Excel 출력 ────────────────────────────────────────────────
    def _build_result_xlsx(self, path: str):
        wb = openpyxl.Workbook()
        self.df_invoice = self._build_invoice_df()

        # 시트1: 인보이스
        ws = wb.active
        ws.title = "①인보이스"
        df = self.df_invoice
        cols = list(df.columns)
        for c, name in enumerate(cols, 1):
            _hdr_cell(ws.cell(1, c, name), C["hdr_blue"])
            ws.column_dimensions[get_column_letter(c)].width = max(len(str(name)) + 2, 12)
        ws.row_dimensions[1].height = 20
        for r, (_, row) in enumerate(df.iterrows(), 2):
            result = row.get("플랫폼대사결과", "")
            diff   = row.get("금액차이", None)
            if result == "플랫폼누락":
                row_color = C["row_red"]
            elif pd.notna(diff):
                try:
                    row_color = C["row_yellow"] if abs(float(diff)) > 1 else (
                        C["row_even"] if r % 2 == 0 else None)
                except (ValueError, TypeError):
                    row_color = C["row_even"] if r % 2 == 0 else None
            else:
                row_color = C["row_even"] if r % 2 == 0 else None
            for c, col in enumerate(cols, 1):
                cell = ws.cell(r, c, row[col])
                _data_cell(cell, row_color, col in MONEY_COLS)
        ws.freeze_panes = "D2"

        # 시트2: 매출현황
        ws2 = wb.create_sheet("②매출현황")
        df2 = self.df_invoice.copy()
        if "정산금액" in df2.columns:
            df2["정산금액"] = pd.to_numeric(df2["정산금액"], errors="coerce").fillna(0)
            amt = "정산금액"
            row = 1
            def _block(ws, start_row, title, grp_col, grp_df, color):
                ws.cell(start_row, 1, title).font = Font(name=FONT_NAME, bold=True,
                                                          size=11, color=color)
                r = start_row + 1
                for c, h in enumerate([grp_col, "건수", "정산금액 합계"], 1):
                    _hdr_cell(ws.cell(r, c, h), color)
                r += 1
                for _, row in grp_df.iterrows():
                    ws.cell(r, 1, row[grp_col]).font = Font(name=FONT_NAME, size=10)
                    ws.cell(r, 2, int(row["건수"])).font = Font(name=FONT_NAME, size=10)
                    money_cell(ws, r, 3, row["합계"])
                    r += 1
                ws.cell(r, 1, "소계").font = Font(name=FONT_NAME, bold=True)
                ws.cell(r, 2, int(grp_df["건수"].sum())).font = Font(name=FONT_NAME, bold=True)
                money_cell(ws, r, 3, grp_df["합계"].sum()).font = Font(name=FONT_NAME, bold=True)
                return r + 2

            if "매출과세구분" in df2.columns:
                g = df2.groupby("매출과세구분")[amt].agg(건수="count", 합계="sum").reset_index()
                row = _block(ws2, row, "▶ 매출과세구분별 집계", "매출과세구분", g, C["hdr_blue"])
            if "협력사명" in df2.columns:
                g4 = df2.groupby("협력사명")[amt].agg(
                    건수="count", 합계="sum").reset_index().sort_values("합계", ascending=False)
                row = _block(ws2, row, "▶ 협력사별 매출 집계", "협력사명", g4, C["hdr_purple"])
            set_col_widths(ws2, [26, 10, 22])

        # 시트A~E: 대사검토 (정산결과 파일에 포함)
        ws_a = wb.create_sheet("③플랫폼누락")
        if self.df_kt_missing is not None and len(self.df_kt_missing) > 0:
            write_df_to_sheet(ws_a, self.df_kt_missing, C["hdr_red"], money_cols=MONEY_COLS)
        else:
            ws_a["A1"] = "플랫폼 누락 주문 없음 ✓"
            ws_a["A1"].font = Font(name=FONT_NAME, color=C["ok_green"], bold=True)

        ws_b = wb.create_sheet("④금액불일치")
        if self.df_amount_diff is not None and len(self.df_amount_diff) > 0:
            write_df_to_sheet(ws_b, self.df_amount_diff, C["hdr_orange"], money_cols=MONEY_COLS)
        else:
            ws_b["A1"] = "금액 불일치 없음 ✓"
            ws_b["A1"].font = Font(name=FONT_NAME, color=C["ok_green"], bold=True)

        ws_c = wb.create_sheet("⑤반품처리")
        if self.df_ret_matched is not None and len(self.df_ret_matched) > 0:
            write_df_to_sheet(ws_c, self.df_ret_matched, C["hdr_green"], money_cols=MONEY_COLS)
        else:
            ws_c["A1"] = "반품 정상처리 건 없음"
            ws_c["A1"].font = Font(name=FONT_NAME, italic=True)

        wb.save(path)

    # ── 메인 실행 ─────────────────────────────────────────────────
    def run(self) -> dict:
        """
        정산 실행 후 결과 요약 딕셔너리 반환.
        {
          "period_start": ..., "period_end": ...,
          "kt_total": int, "pl_total": int,
          "matched": int, "missing": int, "pl_only": int,
          "amount_diff": int,
          "ret_matched": int, "ret_unmatch": int,
          "status": "정상" | "확인필요",
          "result_path": str,
          "missing_rows": [...],   # 플랫폼 누락 주요 컬럼
          "diff_rows": [...],      # 금액불일치 주요 컬럼
        }
        """
        self.step1_filter()
        self.step2_separate_returns()
        self.step3_reconcile_normal()
        self.step4_reconcile_returns()

        today_str   = self.today.strftime("%Y%m%d")
        result_path = str(self.output_dir / f"정산결과_{today_str}.xlsx")
        self._build_result_xlsx(result_path)

        n_miss = len(self.df_kt_missing)  if self.df_kt_missing  is not None else 0
        n_diff = len(self.df_amount_diff) if self.df_amount_diff is not None else 0
        n_unmt = len(self.df_ret_unmatch) if self.df_ret_unmatch is not None else 0

        # 미리보기용 행 데이터
        def _to_rows(df, cols):
            if df is None or df.empty:
                return []
            available = [c for c in cols if c in df.columns]
            return df[available].fillna("").head(20).to_dict("records")

        miss_cols = ["주문번호", "입고일", "정산금액", "협력사명", "주문유형"]
        diff_cols = ["주문번호", "정산금액", "플랫폼정산금액", "금액차이", "협력사명"]

        # 정산금액 합계
        total_amt = 0
        if self.df_invoice is not None and "정산금액" in self.df_invoice.columns:
            total_amt = int(pd.to_numeric(
                self.df_invoice["정산금액"], errors="coerce").fillna(0).sum())

        return {
            "period_start" : self.period_start.strftime("%Y-%m-%d"),
            "period_end"   : self.period_end.strftime("%Y-%m-%d"),
            "kt_total"     : len(self.df_kt_filtered)  if self.df_kt_filtered  is not None else 0,
            "pl_total"     : len(self.df_pl_filtered)  if self.df_pl_filtered  is not None else 0,
            "matched"      : len(self.df_matched)      if self.df_matched      is not None else 0,
            "missing"      : n_miss,
            "pl_only"      : len(self.df_pl_only)      if self.df_pl_only      is not None else 0,
            "amount_diff"  : n_diff,
            "ret_matched"  : len(self.df_ret_matched)  if self.df_ret_matched  is not None else 0,
            "ret_unmatch"  : n_unmt,
            "total_amount" : total_amt,
            "status"       : "정상" if (n_miss + n_diff + n_unmt) == 0 else "확인필요",
            "result_path"  : result_path,
            "missing_rows" : _to_rows(self.df_kt_missing, miss_cols),
            "diff_rows"    : _to_rows(self.df_amount_diff, diff_cols),
        }
