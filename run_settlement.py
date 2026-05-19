"""
run_settlement.py  ─  KT 정산 자동화 시스템 v2.0
=====================================================
실행: python run_settlement.py
로그: settlement_log.txt
출력: 정산결과_YYYYMMDD.xlsx
"""

# ── 0. 라이브러리 자동 설치 ───────────────────────────────────────
import subprocess, sys
for _p in ["pandas", "openpyxl"]:
    try:
        __import__(_p)
    except ImportError:
        print(f"[설치중] {_p}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", _p, "-q"])

import os, logging, datetime, calendar
from pathlib import Path
from typing import Optional, Dict, Tuple, List

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 1. 상수 ──────────────────────────────────────────────────────

FONT = "Arial"

# 행 하이라이트 색상
ROW_MISSING  = "FFCCCC"   # 플랫폼 누락  → 빨강
ROW_RETURN   = "DDEEFF"   # 반품          → 파랑
ROW_AMOUNT   = "FFFF99"   # 금액 불일치   → 노랑
ROW_EVEN     = "F2F2F2"   # 짝수 행 기본

# 헤더 색상
HDR_BLUE     = "1F4E79"
HDR_GREEN    = "375623"
HDR_RED      = "C00000"
HDR_GRAY     = "808080"
ROW_TOTAL    = "D9D9D9"   # 합계 행 배경

# KT 핵심 컬럼 후보 목록  (실제 KT raw 파일 기준)
_KT_CANDIDATES = {
    "주문번호"    : ["구매문서번호"],                          # KT 구매문서번호
    "구매품목"    : ["구매품목"],                              # 라인 아이템 번호 (복합 키 후반부)
    "요청번호"    : ["인수증"],                                # KT 인수증 (플랫폼 요청번호 대응)
    "입고일"      : ["납품일자"],                              # KT 납품일자
    "정산금액"    : ["공급가액"],                              # KT 공급가액 (세전)
    "협력사명"    : ["공급자명", "협력사명"],                  # KT 공급자명
    "매출과세구분": ["세금코드명", "매출과세구분"],            # KT 세금코드명 → 정규화
    "이동유형"    : ["이동유형명", "이동유형"],                # 이동유형명 우선
    "관리회계"    : ["관리회계", "관리회계(IP만변경적용)"],
    "담당자"      : ["담당자"],
    "일반통신구분": ["일반/통신구분", "일반통신구분"],
}
# 플랫폼 핵심 컬럼 후보 목록  (플랫폼이 인보이스 기준)
_PL_CANDIDATES = {
    "주문번호"      : ["주문번호"],
    "품목번호"      : ["품목번호"],                            # 라인 아이템 번호 (복합 키 후반부)
    "입고일"        : ["입고일"],
    "정산금액"      : ["정산금액"],
    "협력사명"      : ["협력사명"],
    "매출과세구분"  : ["매출과세구분"],
    "서비스카테고리": ["서비스카테고리"],   # XLOOKUP 소스 (플랫폼에 있음)
}
_PL_OPT = {"요청번호": ["요청번호"]}  # optional — 없으면 경로 B

MONEY_FMT = "#,##0"

# ── 2. 스타일 헬퍼 ───────────────────────────────────────────────

def _bdr():
    s = Side(style="thin", color="C0C0C0")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_cell(cell, color=HDR_BLUE):
    cell.font      = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", start_color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _bdr()

def data_cell(cell, fill=None, money=False, bold=False):
    cell.font      = Font(name=FONT, size=10, bold=bold)
    cell.border    = _bdr()
    cell.alignment = Alignment(horizontal="right" if money else "center",
                               vertical="center")
    if fill:
        cell.fill  = PatternFill("solid", start_color=fill)
    if money:
        cell.number_format = MONEY_FMT

def total_cell(cell, money=False):
    data_cell(cell, fill=ROW_TOTAL, money=money, bold=True)

def col_letter(ws, col_name):
    """워크시트에서 헤더명으로 열 문자 반환"""
    for c in ws.iter_cols(1, ws.max_column, 1, 1):
        if c[0].value == col_name:
            return get_column_letter(c[0].column)
    return None

def auto_col_width(ws, min_w=10, max_w=30):
    for col in ws.columns:
        length = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(length + 2, min_w), max_w)

# ── 3. 컬럼 탐지 헬퍼 ───────────────────────────────────────────

def detect_col(df: pd.DataFrame, candidates: List[str],
               required=True) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    for col in df.columns:
        for cand in candidates:
            if cand in col:
                return col
    if required:
        raise ValueError(f"필수 컬럼 미발견: {candidates}")
    return None

def build_col_map(df: pd.DataFrame, spec: dict) -> Dict[str, Optional[str]]:
    result = {}
    _optional = {"관리회계", "담당자", "일반통신구분", "서비스카테고리", "요청번호",
                 "구매품목", "품목번호"}
    for key, cands in spec.items():
        result[key] = detect_col(df, cands, required=(key not in _optional))
    return result

# ── 4. MappingMaster ─────────────────────────────────────────────

class MappingMaster:
    def __init__(self, path: Optional[Path]):
        self.path      = path
        self.svccat    : Dict[str, Tuple[str, str]] = {}  # {카테고리: (관리회계, 담당자)}
        self.person_type: Dict[str, str]             = {}  # {담당자: 일반/통신}
        self.corp_size  : Dict[str, str]             = {}  # {협력사명: 규모}
        self.loaded     = False

    def load(self, log):
        if not self.path or not self.path.exists():
            log.warning("  mapping_master.xlsx 없음 → 관리회계·담당자·일반통신구분·기업규모 공란 처리")
            return
        try:
            # 시트1: 서비스카테고리
            df1 = pd.read_excel(self.path, sheet_name=0, dtype=str, engine="openpyxl").fillna("")
            c1 = df1.columns.tolist()
            # 컬럼 위치 유연하게 탐지
            cat_c = next((c for c in c1 if "서비스카테고리" in c or "카테고리" in c), c1[0])
            acc_c = next((c for c in c1 if "관리회계" in c), c1[1] if len(c1) > 1 else None)
            mgr_c = next((c for c in c1 if "담당자" in c), c1[2] if len(c1) > 2 else None)
            for _, r in df1.iterrows():
                k = str(r[cat_c]).strip()
                if k:
                    self.svccat[k] = (
                        str(r[acc_c]).strip() if acc_c else "",
                        str(r[mgr_c]).strip() if mgr_c else "",
                    )

            # 시트2: 일반통신구분
            df2 = pd.read_excel(self.path, sheet_name=1, dtype=str, engine="openpyxl").fillna("")
            c2 = df2.columns.tolist()
            per_c  = next((c for c in c2 if "담당자" in c), c2[0])
            type_c = next((c for c in c2 if "구분" in c or "통신" in c), c2[1] if len(c2) > 1 else None)
            for _, r in df2.iterrows():
                k = str(r[per_c]).strip()
                if k and type_c:
                    self.person_type[k] = str(r[type_c]).strip()

            # 시트3: 기업규모
            df3 = pd.read_excel(self.path, sheet_name=2, dtype=str, engine="openpyxl").fillna("")
            c3 = df3.columns.tolist()
            corp_c = next((c for c in c3 if "협력사" in c or "기업명" in c), c3[0])
            size_c = next((c for c in c3 if "규모" in c or "기업규모" in c), c3[1] if len(c3) > 1 else None)
            for _, r in df3.iterrows():
                k = str(r[corp_c]).strip()
                if k and size_c:
                    self.corp_size[k] = str(r[size_c]).strip()

            self.loaded = True
            log.info(f"  mapping_master 로드: 서비스카테고리 {len(self.svccat)}건 / "
                     f"일반통신구분 {len(self.person_type)}건 / 기업규모 {len(self.corp_size)}건")
        except Exception as e:
            log.warning(f"  mapping_master 로드 실패: {e}")

    def apply(self, df: pd.DataFrame, kt_cols: Dict[str, Optional[str]], log) -> pd.DataFrame:
        """서비스카테고리 → 관리회계/담당자/일반통신구분 채우기 (기존값 우선, 벡터 연산)"""
        svc_col  = kt_cols.get("서비스카테고리")
        acc_col  = kt_cols.get("관리회계")  or "관리회계"
        mgr_col  = kt_cols.get("담당자")    or "담당자"
        type_col = kt_cols.get("일반통신구분") or "일반/통신구분"

        # 컬럼 생성 + 문자열 dtype 강제 (빈 셀이 float64로 읽힐 수 있음)
        for col in [acc_col, mgr_col, type_col]:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].fillna("").astype(str)

        def _blank(series):
            """공백/None인 셀 마스크"""
            return series.astype(str).str.strip().isin(["", "None", "nan"])

        fail_cats = set()
        stats = {"관리회계": [0, 0], "담당자": [0, 0], "일반통신구분": [0, 0]}

        # ── 서비스카테고리 → 관리회계 / 담당자 (벡터) ──
        if svc_col and svc_col in df.columns:
            df["_svc"] = df[svc_col].astype(str).str.strip()
            df["_acc_mapped"] = df["_svc"].map(
                lambda x: self.svccat.get(x, (None, None))[0]
            )
            df["_mgr_mapped"] = df["_svc"].map(
                lambda x: self.svccat.get(x, (None, None))[1]
            )
            # 기존값 없으면 매핑값으로 채움
            acc_blank = _blank(df[acc_col])
            df.loc[acc_blank, acc_col] = df.loc[acc_blank, "_acc_mapped"]
            mgr_blank = _blank(df[mgr_col])
            df.loc[mgr_blank, mgr_col] = df.loc[mgr_blank, "_mgr_mapped"]

            # 통계
            mapped   = df["_acc_mapped"].notna() & (df["_acc_mapped"] != "")
            unmapped = ~mapped
            stats["관리회계"] = [int(mapped.sum()), int(unmapped.sum())]
            stats["담당자"]   = [int(mapped.sum()), int(unmapped.sum())]
            fail_cats = set(df.loc[unmapped & (df["_svc"] != "") & (df["_svc"] != "nan"),
                                   "_svc"].tolist())
            df.drop(columns=["_svc", "_acc_mapped", "_mgr_mapped"], inplace=True)

        # ── 담당자 → 일반/통신구분 (담당자 컬럼이 채워진 뒤에 적용) ──
        df["_mgr_clean"] = df[mgr_col].astype(str).str.strip()
        df["_type_mapped"] = df["_mgr_clean"].map(self.person_type)
        type_blank = _blank(df[type_col])
        df.loc[type_blank, type_col] = df.loc[type_blank, "_type_mapped"]

        ok_type = df["_type_mapped"].notna()
        stats["일반통신구분"] = [int(ok_type.sum()), int((~ok_type).sum())]
        df.drop(columns=["_mgr_clean", "_type_mapped"], inplace=True)

        log.info(f"  관리회계 매핑  : {stats['관리회계'][0]}건 성공 / {stats['관리회계'][1]}건 실패")
        log.info(f"  담당자 매핑    : {stats['담당자'][0]}건 성공 / {stats['담당자'][1]}건 실패")
        log.info(f"  일반통신구분   : {stats['일반통신구분'][0]}건 성공 / {stats['일반통신구분'][1]}건 실패")
        if fail_cats:
            log.warning(f"  매핑 실패 서비스카테고리: {sorted(fail_cats)}")
        return df, stats

# ── 5. SettlementRunner ──────────────────────────────────────────

class SettlementRunner:

    def __init__(self, base_dir="."):
        self.base       = Path(base_dir)
        self.today      = datetime.date.today()
        self.log        = self._setup_logger()
        self.kt_path    : Optional[Path] = None
        self.pl_path    : Optional[Path] = None
        self.rm_path    : Optional[Path] = None   # return_mapping.xlsx
        self.master     : Optional[MappingMaster] = None
        # 정산 기간
        self.ps         : Optional[datetime.date] = None
        self.pe         : Optional[datetime.date] = None
        # 컬럼 매핑
        self.kt_cols    : Dict[str, Optional[str]] = {}
        self.pl_cols    : Dict[str, Optional[str]] = {}
        # 처리 결과
        self.df_kt      = None   # KT 기간 필터 전체
        self.df_pl      = None
        self.df_normal  = None   # KT 일반
        self.df_ret_kt  = None   # KT 반품
        self.df_matched = None
        self.df_missing = None   # 플랫폼 누락
        self.df_pl_only = None
        self.df_amtdiff = None
        self.df_ret_ok  = None   # 반품 매칭 성공
        self.df_ret_ng  = None   # 반품 미확인
        self.ret_route  = None   # 'A' or 'B'
        self.df_invoice    = None
        self.map_stats     = {}
        self.composite_key = False   # True = 주문번호+품목번호 복합 키 사용

    def _setup_logger(self):
        log = logging.getLogger("settlement")
        log.setLevel(logging.DEBUG)
        if not log.handlers:
            fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                    datefmt="%H:%M:%S")
            sh = logging.StreamHandler(sys.stdout); sh.setFormatter(fmt)
            fh = logging.FileHandler("settlement_log.txt", encoding="utf-8")
            fh.setFormatter(fmt)
            log.addHandler(sh); log.addHandler(fh)
        return log

    # ── 파일 탐지 ───────────────────────────────────────────────
    def detect_files(self):
        self.log.info("=" * 60)
        self.log.info("  KT 정산 자동화 시스템  v2.0")
        self.log.info("=" * 60)

        def _find(*names):
            for n in names:
                p = self.base / n
                if p.exists():
                    return p
            return None

        self.kt_path = _find("kt_raw.xlsx", "KT_raw.xlsx", "kt_raw.xlsm")
        self.pl_path = _find("platform.xlsx", "플랫폼.xlsx", "통합플랫폼.xlsx")
        self.rm_path = _find("return_mapping.xlsx")
        mm_path      = _find("mapping_master.xlsx", "매핑마스터.xlsx")

        for label, path, req in [
            ("KT 파일        ", self.kt_path, True),
            ("플랫폼 파일    ", self.pl_path, True),
            ("mapping_master ", mm_path,      False),
            ("return_mapping ", self.rm_path, False),
        ]:
            self.log.info(f"  {label}: {path or '없음'}")
            if req and not path:
                fname = input(f"\n  '{label.strip()}' 파일명 입력: ").strip()
                if label.strip().startswith("KT"):
                    self.kt_path = self.base / fname
                else:
                    self.pl_path = self.base / fname

        self.master = MappingMaster(mm_path)
        self.master.load(self.log)

    # ── STEP 0: 컬럼 분석 + 확인 ──────────────────────────────
    def step0_init(self):
        self.log.info("\n[STEP 0] 파일 컬럼 구조 분석")
        df_kt = pd.read_excel(self.kt_path, nrows=3, dtype=str, engine="openpyxl")
        df_pl = pd.read_excel(self.pl_path, nrows=3, dtype=str, engine="openpyxl")

        self.kt_cols = build_col_map(df_kt, _KT_CANDIDATES)
        self.pl_cols = build_col_map(df_pl, _PL_CANDIDATES)
        # 플랫폼 요청번호(optional) 별도 탐지
        self.pl_cols["요청번호"] = detect_col(df_pl, _PL_OPT["요청번호"], required=False)

        LINE = "─" * 55
        print(f"\n{LINE}")
        print("▶ KT 파일 컬럼 탐지 결과:")
        for key, found in self.kt_cols.items():
            mark = "✓" if found else "─"
            print(f"  {mark} {key:<14} → {found or '(미탐지)'}")

        print(f"\n▶ 플랫폼 파일 컬럼 탐지 결과:")
        for key, found in self.pl_cols.items():
            mark = "✓" if found else "─"
            print(f"  {mark} {key:<14} → {found or '(없음)'}")

        req_col = self.pl_cols.get("요청번호")
        kt_item = self.kt_cols.get("구매품목")
        pl_item = self.pl_cols.get("품목번호")
        key_mode = (f"복합 키  ({self.kt_cols['주문번호']}+{kt_item}  ↔  "
                    f"{self.pl_cols['주문번호']}+{pl_item})"
                    if kt_item and pl_item else
                    f"단순 키  ({self.kt_cols['주문번호']} ↔ {self.pl_cols['주문번호']})")
        print(f"\n▶ 매칭 키 모드  : {key_mode}")
        print(f"▶ 반품 처리 경로: {'A (요청번호 직접 매칭)' if req_col else 'B (return_mapping.xlsx 사용)'}")
        print(LINE)

        ans = input("위 설정으로 진행하시겠습니까? (Y/n): ").strip().lower()
        if ans == "n":
            sys.exit(0)

    # ── 정산 기간 ──────────────────────────────────────────────
    def get_period(self):
        d, y, m = self.today.day, self.today.year, self.today.month
        if d == 16:
            s, e = datetime.date(y, m, 1), datetime.date(y, m, 15)
        elif d == 1:
            py, pm = (y-1, 12) if m == 1 else (y, m-1)
            s = datetime.date(py, pm, 16)
            e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
        else:
            self.log.warning(f"  오늘({self.today})은 자동 판단 기준일이 아닙니다.")
            s = datetime.date.fromisoformat(input("  시작일 (YYYY-MM-DD): ").strip())
            e = datetime.date.fromisoformat(input("  종료일 (YYYY-MM-DD): ").strip())
        self.log.info(f"  정산 기간: {s} ~ {e}")
        return s, e

    # ── KT 전처리 ──────────────────────────────────────────────
    def _preprocess_kt(self, df: pd.DataFrame) -> pd.DataFrame:
        """① 합계 행 제거(구매문서번호 공란)  ② 세금코드명 → 과세/면세/비과세 정규화"""
        ord_col = self.kt_cols.get("주문번호")  # 구매문서번호
        if ord_col and ord_col in df.columns:
            before = len(df)
            df = df[df[ord_col].fillna("").astype(str).str.strip() != ""].copy()
            removed = before - len(df)
            if removed > 0:
                self.log.info(f"  합계/빈값 행 {removed}건 제거 ({ord_col} 공란)")

        tax_col = self.kt_cols.get("매출과세구분")  # 세금코드명
        if tax_col and tax_col in df.columns:
            def _norm(v):
                v = str(v).strip()
                if "비과세" in v: return "비과세"
                if "면세"   in v: return "면세"
                if "과세"   in v: return "과세"
                return v
            df[tax_col] = df[tax_col].map(_norm)
            self.log.info(f"  세금코드명 정규화 완료 (→ 과세/면세/비과세)")
        return df

    # ── 복합 키 생성 ────────────────────────────────────────────
    @staticmethod
    def _item_str(series: pd.Series) -> pd.Series:
        """품목번호를 정수 문자열로 정규화: 10.0 → '10', '010' → '10', '' → ''"""
        def _conv(v):
            s = str(v).strip()
            if s in ("", "nan", "None"): return ""
            try:
                return str(int(float(s)))
            except (ValueError, TypeError):
                return s
        return series.apply(_conv)

    def _assign_match_keys(self):
        """KT와 플랫폼 DataFrame에 _match_key 컬럼 부여.
        구매문서번호+구매품목 ↔ 주문번호+품목번호 (구분자 없이 연결)
        품목번호 컬럼이 없으면 주문번호만 사용(단순 키 폴백).
        """
        kt_ord  = self.kt_cols["주문번호"]
        pl_ord  = self.pl_cols["주문번호"]
        kt_item = self.kt_cols.get("구매품목")
        pl_item = self.pl_cols.get("품목번호")

        if kt_item and pl_item and kt_item in self.df_kt.columns and pl_item in self.df_pl.columns:
            self.composite_key = True
            self.df_kt["_match_key"] = (
                self.df_kt[kt_ord].fillna("").astype(str)
                + self._item_str(self.df_kt[kt_item])
            )
            self.df_pl["_match_key"] = (
                self.df_pl[pl_ord].fillna("").astype(str)
                + self._item_str(self.df_pl[pl_item])
            )
            self.log.info(f"  복합 키 생성: {kt_ord}+{kt_item} ↔ {pl_ord}+{pl_item}")
        else:
            self.composite_key = False
            self.df_kt["_match_key"] = self.df_kt[kt_ord].fillna("").astype(str)
            self.df_pl["_match_key"] = self.df_pl[pl_ord].fillna("").astype(str)
            self.log.info(f"  단순 키 사용: {kt_ord} ↔ {pl_ord}")

    # ── STEP 1: 기간 필터 ──────────────────────────────────────
    def step1_filter(self):
        self.log.info("\n[STEP 1] 정산 기간 필터링")
        self.ps, self.pe = self.get_period()

        def _load(path, str_cols):
            return pd.read_excel(path, dtype={c: str for c in str_cols
                                              if c}, engine="openpyxl")

        str_kt = [self.kt_cols.get("주문번호"), self.kt_cols.get("요청번호")]
        str_pl = [self.pl_cols.get("주문번호"), self.pl_cols.get("요청번호")]

        df_kt = _load(self.kt_path, str_kt)
        df_pl = _load(self.pl_path, str_pl)

        # KT 전처리: 합계 행 제거 + 세금코드명 정규화
        df_kt = self._preprocess_kt(df_kt)

        dc_kt = self.kt_cols["입고일"]
        dc_pl = self.pl_cols["입고일"]
        df_kt[dc_kt] = pd.to_datetime(df_kt[dc_kt], errors="coerce").dt.date
        df_pl[dc_pl] = pd.to_datetime(df_pl[dc_pl], errors="coerce").dt.date

        self.df_kt = df_kt[(df_kt[dc_kt] >= self.ps) & (df_kt[dc_kt] <= self.pe)].copy().reset_index(drop=True)
        self.df_pl = df_pl[(df_pl[dc_pl] >= self.ps) & (df_pl[dc_pl] <= self.pe)].copy().reset_index(drop=True)
        self.log.info(f"  KT 기간 내: {len(self.df_kt)}건 / 플랫폼 기간 내: {len(self.df_pl)}건")
        # 복합 키 생성 (구매문서번호+구매품목 ↔ 주문번호+품목번호)
        self._assign_match_keys()

    # ── STEP 2: 반품 선분리 ────────────────────────────────────
    def step2_separate_returns(self):
        self.log.info("\n[STEP 2] 반품 선분리 (이동유형 기준)")
        mv_col = self.kt_cols.get("이동유형")
        if not mv_col:
            self.log.warning("  이동유형 컬럼 미발견 → 전체 일반 처리")
            self.df_normal = self.df_kt.copy()
            self.df_ret_kt = self.df_kt.iloc[0:0].copy()
            return

        self.log.info(f"  이동유형 컬럼: '{mv_col}'  고유값: {self.df_kt[mv_col].dropna().unique().tolist()}")
        mask = self.df_kt[mv_col].astype(str).str.contains("반품", na=False)
        self.df_ret_kt = self.df_kt[mask].copy().reset_index(drop=True)
        self.df_normal = self.df_kt[~mask].copy().reset_index(drop=True)
        self.log.info(f"  일반: {len(self.df_normal)}건 / 반품: {len(self.df_ret_kt)}건")

    # ── STEP 3: 일반 주문 대사 ─────────────────────────────────
    def step3_reconcile_normal(self):
        key_label = "복합 키" if self.composite_key else "주문번호"
        self.log.info(f"\n[STEP 3] 일반 주문 대사 ({key_label} 기준)")
        amt_kt = self.kt_cols["정산금액"]
        amt_pl = self.pl_cols["정산금액"]

        # _match_key 기반 집합 비교
        kt_set = set(self.df_normal["_match_key"])
        pl_set = set(self.df_pl["_match_key"])

        matched = kt_set & pl_set
        kt_only = kt_set - pl_set
        pl_only = pl_set - kt_set

        # A: 일치 + 금액 비교
        df_m = self.df_normal[self.df_normal["_match_key"].isin(matched)].copy()
        pl_amt_map = (self.df_pl.set_index("_match_key")[amt_pl]
                      .apply(pd.to_numeric, errors="coerce")
                      .to_dict())
        df_m["_amt_kt"] = pd.to_numeric(df_m[amt_kt], errors="coerce")
        df_m["플랫폼정산금액"] = df_m["_match_key"].map(pl_amt_map)
        df_m["금액차이"] = df_m["_amt_kt"] - df_m["플랫폼정산금액"]
        df_m["플랫폼대사결과"] = "정상"
        df_m.drop(columns=["_amt_kt"], inplace=True)

        self.df_matched  = df_m
        self.df_amtdiff  = df_m[df_m["금액차이"].abs() > 1].copy()
        self.df_missing  = self.df_normal[self.df_normal["_match_key"].isin(kt_only)].copy()
        self.df_pl_only  = self.df_pl[self.df_pl["_match_key"].isin(pl_only)].copy()

        self.log.info(f"  [A] 일치        : {len(matched)}건")
        self.log.info(f"  [B] 플랫폼 누락 : {len(kt_only)}건  ← KT에만 존재")
        self.log.info(f"  [C] KT 미포함   : {len(pl_only)}건  ← 플랫폼에만 존재")
        self.log.info(f"  금액 불일치     : {len(self.df_amtdiff)}건  (±1원 초과)")
        if kt_only:
            self.log.info(f"  누락 복합키     : {sorted(kt_only)}")

    # ── STEP 4: 반품 대사 ──────────────────────────────────────
    def step4_reconcile_returns(self):
        self.log.info("\n[STEP 4] 반품 처리")
        if len(self.df_ret_kt) == 0:
            self.log.info("  반품 건 없음")
            self.df_ret_ok = self.df_ret_kt.copy()
            self.df_ret_ng = self.df_ret_kt.copy()
            self.ret_route = "없음"
            return

        pl_req_col = self.pl_cols.get("요청번호")
        if pl_req_col:
            self._route_a(pl_req_col)
        else:
            self._route_b()

    def _route_a(self, pl_req_col):
        self.ret_route = "A"
        self.log.info("  경로 A: 요청번호 직접 매칭")
        req_col = self.kt_cols.get("요청번호")
        if not req_col:
            self.log.warning("  KT 요청번호 컬럼 없음 → 전체 미확인")
            self.df_ret_ok = self.df_ret_kt.iloc[0:0].copy()
            self.df_ret_ng = self.df_ret_kt.copy()
            return

        pl_reqs = set(self.df_pl[pl_req_col].dropna().astype(str))
        mask = self.df_ret_kt[req_col].astype(str).isin(pl_reqs)

        ok = self.df_ret_kt[mask].copy()
        ok["플랫폼대사결과"] = "반품"
        # 플랫폼 금액 매칭
        pl_amt_map = self.df_pl.set_index(pl_req_col)[self.pl_cols["정산금액"]].apply(
            pd.to_numeric, errors="coerce"
        ).to_dict()
        ok["플랫폼정산금액"] = ok[req_col].astype(str).map(pl_amt_map)
        ok["금액차이"] = pd.to_numeric(ok[self.kt_cols["정산금액"]], errors="coerce") - ok["플랫폼정산금액"]

        self.df_ret_ok = ok
        ng = self.df_ret_kt[~mask].copy()
        ng["플랫폼대사결과"] = "미확인"
        self.df_ret_ng = ng
        self.log.info(f"  반품 매칭 성공: {len(ok)}건 / 미확인: {len(ng)}건")

    def _route_b(self):
        self.ret_route = "B"
        self.log.info("  경로 B: return_mapping.xlsx 사용")
        req_col  = self.kt_cols.get("요청번호")
        ord_col  = self.kt_cols["주문번호"]
        corp_col = self.kt_cols.get("협력사명")
        amt_col  = self.kt_cols["정산금액"]

        if not self.rm_path or not self.rm_path.exists():
            # ① 파일 없음 → 생성 후 종료
            kt_item_col = self.kt_cols.get("구매품목")
            rows = []
            for _, r in self.df_ret_kt.iterrows():
                rows.append({
                    "KT반품복합키"   : r.get("_match_key", r.get(ord_col, "")),
                    "KT반품주문번호" : r.get(ord_col, ""),
                    "KT구매품목"     : (r.get(kt_item_col, "") if kt_item_col else ""),
                    "KT요청번호"     : (r.get(req_col, "") if req_col else ""),
                    "플랫폼원복합키" : "",  # 사용자 입력: 플랫폼 주문번호+품목번호 연결
                    "협력사명"       : (r.get(corp_col, "") if corp_col else ""),
                    "정산금액"       : r.get(amt_col, ""),
                    "처리상태"       : "미입력",
                })
            rm_df = pd.DataFrame(rows)
            rm_df.to_excel("return_mapping.xlsx", index=False)
            n = len(rows)
            key_hint = ("복합 키(주문번호+품목번호)" if self.composite_key
                        else "주문번호")
            print(f"""
{"="*60}
  반품 {n}건의 매핑 파일이 생성됐습니다: return_mapping.xlsx

  1. return_mapping.xlsx 열기
  2. 각 행의 KT요청번호로 통합플랫폼에서 대응 건 검색
  3. 플랫폼 '{key_hint}'를 E열(플랫폼원복합키)에 입력
     예) 주문번호 'ORD-001' + 품목번호 '10' → 'ORD-00110'
  4. python run_settlement.py 재실행

  ※ 전산팀에 플랫폼 요청번호 컬럼 추가 요청 시
    완전 자동화(경로 A) 가능
{"="*60}""")
            sys.exit(0)

        # ② 파일 있음 → 매칭 실행
        rm = pd.read_excel(self.rm_path, dtype=str, engine="openpyxl").fillna("")
        pl_amt_col = self.pl_cols["정산금액"]
        # 플랫폼 금액 룩업: _match_key 기반
        pl_amt_map = (self.df_pl.set_index("_match_key")[pl_amt_col]
                      .apply(pd.to_numeric, errors="coerce")
                      .to_dict())

        ok_list, ng_list = [], []
        for _, ret_row in self.df_ret_kt.iterrows():
            kt_key = str(ret_row.get("_match_key", "")).strip()
            # return_mapping에서 이 반품 건 찾기 (KT반품복합키 컬럼 우선, 없으면 KT반품주문번호)
            if "KT반품복합키" in rm.columns:
                rm_row = rm[rm["KT반품복합키"].str.strip() == kt_key]
            else:
                rm_row = rm[rm["KT반품주문번호"].str.strip() == str(ret_row.get(ord_col, "")).strip()]
            if rm_row.empty:
                ng_list.append({**ret_row, "플랫폼대사결과": "미확인",
                                 "플랫폼정산금액": None, "금액차이": None})
                continue

            # 플랫폼원복합키 (신규) 또는 플랫폼원주문번호 (구버전) 읽기
            if "플랫폼원복합키" in rm_row.columns:
                pl_key = str(rm_row.iloc[0].get("플랫폼원복합키", "")).strip()
            else:
                pl_key = str(rm_row.iloc[0].get("플랫폼원주문번호", "")).strip()

            if not pl_key:
                ng_list.append({**ret_row, "플랫폼대사결과": "미확인",
                                 "플랫폼정산금액": None, "금액차이": None})
                continue

            pl_amt = pl_amt_map.get(pl_key)
            kt_amt = pd.to_numeric(ret_row.get(amt_col), errors="coerce")
            ok_list.append({**ret_row,
                             "플랫폼대사결과" : "반품",
                             "플랫폼정산금액" : pl_amt,
                             "금액차이"       : (kt_amt - pl_amt) if (pd.notna(kt_amt) and pl_amt is not None) else None})

        self.df_ret_ok = pd.DataFrame(ok_list) if ok_list else self.df_ret_kt.iloc[0:0].copy()
        self.df_ret_ng = pd.DataFrame(ng_list) if ng_list else self.df_ret_kt.iloc[0:0].copy()
        self.log.info(f"  반품 매칭 성공: {len(ok_list)}건 / 미확인: {len(ng_list)}건")

    # ── 인보이스 DataFrame 조립 (플랫폼 기반) ─────────────────
    def build_invoice_df(self) -> pd.DataFrame:
        pl_ord_col = self.pl_cols["주문번호"]
        pl_amt_col = self.pl_cols["정산금액"]
        kt_ord_col = self.kt_cols["주문번호"]   # 구매문서번호
        kt_amt_col = self.kt_cols["정산금액"]   # 공급가액

        # ── 플랫폼 데이터를 인보이스 기준으로 ──
        df = self.df_pl.copy()
        df["대사결과"]   = "정상"
        df["KT공급가액"] = pd.NA
        df["금액차이"]   = pd.NA

        # KT 금액 룩업 (_match_key → 공급가액)
        kt_amt_map = (
            self.df_kt[self.df_kt["_match_key"] != ""]
            .set_index("_match_key")[kt_amt_col]
            .apply(pd.to_numeric, errors="coerce")
            .to_dict()
        )
        df["KT공급가액"] = df["_match_key"].map(kt_amt_map)
        df[pl_amt_col]   = pd.to_numeric(df[pl_amt_col], errors="coerce")
        df["금액차이"]   = df[pl_amt_col] - df["KT공급가액"]

        # KT 미포함 (플랫폼에만 있는 주문)
        if self.df_pl_only is not None and len(self.df_pl_only) > 0:
            pl_only_set = set(self.df_pl_only["_match_key"])
            m = df["_match_key"].isin(pl_only_set)
            df.loc[m, "대사결과"]   = "KT미포함"
            df.loc[m, ["KT공급가액", "금액차이"]] = pd.NA

        # 금액 불일치 (KT 금액 있고 ±1원 초과)
        amt_diff_mask = df["KT공급가액"].notna() & (df["금액차이"].abs() > 1)
        df.loc[amt_diff_mask & (df["대사결과"] == "정상"), "대사결과"] = "금액불일치"

        # 반품 표시 — 경로 B (return_mapping.xlsx)
        if self.ret_route == "B" and self.rm_path and self.rm_path.exists():
            rm = pd.read_excel(self.rm_path, dtype=str, engine="openpyxl").fillna("")
            # 신규 형식(플랫폼원복합키) 또는 구버전(플랫폼원주문번호) 대응
            pl_key_col = ("플랫폼원복합키" if "플랫폼원복합키" in rm.columns
                          else "플랫폼원주문번호")
            pl_from_rm = set(rm[pl_key_col].str.strip()) - {""}
            df.loc[df["_match_key"].isin(pl_from_rm), "대사결과"] = "반품"

        # 반품 표시 — 경로 A (요청번호 매칭)
        pl_req_col = self.pl_cols.get("요청번호")
        if pl_req_col and self.df_ret_ok is not None and len(self.df_ret_ok) > 0:
            kt_req_col = self.kt_cols.get("요청번호")
            if kt_req_col and kt_req_col in self.df_ret_ok.columns:
                ok_reqs = set(self.df_ret_ok[kt_req_col].astype(str))
                df.loc[df[pl_req_col].astype(str).isin(ok_reqs), "대사결과"] = "반품"

        # 플랫폼 누락 (KT에만 있는 주문) → 하단 추가, 빨간 하이라이트
        if self.df_missing is not None and len(self.df_missing) > 0:
            miss = self.df_missing.copy()
            # KT 컬럼명 → 플랫폼 컬럼명으로 리매핑
            col_remap = {}
            for key in ["주문번호", "입고일", "정산금액", "협력사명", "매출과세구분"]:
                kt_c = self.kt_cols.get(key)
                pl_c = self.pl_cols.get(key)
                if kt_c and pl_c and kt_c != pl_c and kt_c in miss.columns:
                    col_remap[kt_c] = pl_c
            miss = miss.rename(columns=col_remap)
            miss["대사결과"]   = "플랫폼누락"
            miss["KT공급가액"] = pd.to_numeric(miss.get(pl_amt_col), errors="coerce")
            miss["금액차이"]   = pd.NA
            df = pd.concat([df, miss], ignore_index=True)

        # XLOOKUP 매핑 (플랫폼 서비스카테고리 → 관리회계/담당자/일반통신구분)
        if self.master and self.master.loaded:
            # pl_cols 에 서비스카테고리, kt_cols 에 타겟 컬럼명
            merged_cols = {**self.pl_cols, **self.kt_cols}
            df, self.map_stats = self.master.apply(df, merged_cols, self.log)

        # 내부 작업 컬럼 제거 (Excel 출력 제외)
        drop_cols = [c for c in df.columns if str(c).startswith("_")]
        df = df.drop(columns=drop_cols, errors="ignore")

        self.df_invoice = df
        return df

    # ── 시트1: 인보이스 ────────────────────────────────────────
    def write_invoice_sheet(self, wb):
        self.log.info("  [시트1] 인보이스 작성...")
        ws = wb.active; ws.title = "①인보이스"
        df = self.df_invoice
        cols = list(df.columns)
        pl_amt_col = self.pl_cols["정산금액"]
        money_set = {pl_amt_col, "KT공급가액", "금액차이"}

        # 헤더
        for c, name in enumerate(cols, 1):
            hdr_cell(ws.cell(1, c, name))
        ws.row_dimensions[1].height = 20

        # 데이터
        for r, (_, row) in enumerate(df.iterrows(), 2):
            res  = row.get("대사결과", "")
            diff = row.get("금액차이")
            if res == "플랫폼누락":
                fill = ROW_MISSING
            elif res == "반품":
                fill = ROW_RETURN
            elif res == "금액불일치":
                fill = ROW_AMOUNT
            else:
                fill = ROW_EVEN if r % 2 == 0 else None

            for c, col in enumerate(cols, 1):
                val = row[col]
                if pd.isna(val):
                    val = None   # pd.NA / NaN → openpyxl이 빈 셀로 처리
                cell = ws.cell(r, c, val)
                data_cell(cell, fill=fill, money=(col in money_set))

        ws.freeze_panes = "A2"
        auto_col_width(ws)
        self.log.info(f"    → {len(df)}행")

    # ── 시트2: 매출현황 ────────────────────────────────────────
    def write_sales_sheet(self, wb):
        self.log.info("  [시트2] 매출현황 작성...")
        ws = wb.create_sheet("②매출현황")
        df = self.df_invoice.copy()
        amt_col  = self.pl_cols["정산금액"]
        tax_col  = self.pl_cols.get("매출과세구분")
        type_col = self.kt_cols.get("일반통신구분") or "일반/통신구분"

        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)

        def _write_block(ws, start_row, title, grp_col, grp_df, col_label):
            """집계 블록 작성 → SUM 수식 합계행 포함. 다음 시작 행 반환"""
            ws.cell(start_row, 1, title).font = Font(name=FONT, bold=True,
                                                      size=12, color=HDR_BLUE)
            r = start_row + 1
            for c, h in enumerate([col_label, "건수", "정산금액합계"], 1):
                hdr_cell(ws.cell(r, c, h))
            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 10
            ws.column_dimensions["C"].width = 22
            data_start = r + 1
            r += 1
            rows_written = 0
            for _, row in grp_df.iterrows():
                fill = ROW_EVEN if r % 2 == 0 else None
                ws.cell(r, 1, row[grp_col]).font = Font(name=FONT, size=10)
                if fill:
                    ws.cell(r, 1).fill = PatternFill("solid", start_color=fill)
                data_cell(ws.cell(r, 2, int(row["건수"])))
                data_cell(ws.cell(r, 3, row["합계"]), money=True)
                r += 1; rows_written += 1
            # 합계 행 (SUM 수식)
            data_end = r - 1
            total_cell(ws.cell(r, 1, "합계"))
            total_cell(ws.cell(r, 2, f"=SUM(B{data_start}:B{data_end})"))
            ws.cell(r, 2).number_format = "#,##0"
            tc = ws.cell(r, 3, f"=SUM(C{data_start}:C{data_end})")
            total_cell(tc, money=True)
            return r + 2

        row = 1
        # ─ 과세구분별 ─
        if tax_col and tax_col in df.columns:
            grp1 = df.groupby(tax_col)[amt_col].agg(건수="count", 합계="sum").reset_index()
            row = _write_block(ws, row, "▶ 과세구분별 집계", tax_col, grp1, "과세구분")
        else:
            ws.cell(row, 1, "매출과세구분 컬럼 없음").font = Font(color="FF0000", bold=True, name=FONT)
            row += 2

        # ─ 일반/통신 구분별 ─
        if type_col in df.columns and df[type_col].notna().any():
            grp2 = df.groupby(type_col)[amt_col].agg(건수="count", 합계="sum").reset_index()
            _write_block(ws, row, "▶ 일반/통신 구분별 집계", type_col, grp2, "구분")
        else:
            ws.cell(row, 1, "일반/통신구분 매핑 확인 필요 (mapping_master.xlsx)").font = \
                Font(color="FF6600", bold=True, name=FONT, italic=True)

    # ── 시트3: 어음확인_최종 ───────────────────────────────────
    def write_bill_sheet(self, wb):
        self.log.info("  [시트3] 어음확인_최종 작성...")
        ws = wb.create_sheet("③어음확인_최종")
        df = self.df_invoice.copy()
        amt_col  = self.pl_cols["정산금액"]
        tax_col  = self.pl_cols.get("매출과세구분")
        corp_col = self.pl_cols.get("협력사명")
        ord_col  = self.pl_cols["주문번호"]

        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)

        if not tax_col or not corp_col:
            ws["A1"] = "매출과세구분 또는 협력사명 컬럼 없음"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT)
            return

        if not self.master or not self.master.loaded:
            ws["A1"] = "⚠ mapping_master.xlsx 없음 — 기업규모 판별 불가"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT, size=11)

        # 비과세만 추출
        df_nt = df[df[tax_col] == "비과세"].copy()
        if df_nt.empty:
            ws["A1"] = "비과세 주문 없음"
            return

        # (주문번호, 협력사명) 기준 합산
        grp = df_nt.groupby([ord_col, corp_col]).agg(
            합산청구액 = (amt_col, "sum"),
            대사결과   = ("대사결과", "first"),
        ).reset_index()
        grp.rename(columns={ord_col: "주문번호", corp_col: "협력사명"}, inplace=True)

        # 기업규모 매핑
        grp["기업규모"] = grp["협력사명"].map(
            self.master.corp_size if self.master else {}
        ).fillna("확인필요")

        # 3조건 판별
        grp["cond1"] = grp["합산청구액"] >= 200_000_000
        grp["cond2"] = grp["기업규모"].isin(["중견기업", "대기업"])
        grp["cond3"] = grp["대사결과"] == "정상"
        grp["어음대상"] = grp.apply(
            lambda r: "발행대상"  if (r.cond1 and r.cond2 and r.cond3)
                 else ("확인필요" if (r.cond1 and (not r.cond3 or r["기업규모"] == "확인필요"))
                 else "미해당"),
            axis=1
        )
        grp["비고"] = grp.apply(
            lambda r: "" if r["어음대상"] == "발행대상"
                      else ("대사 미확인" if r["대사결과"] not in ("정상",)
                            else ("기업규모 미확인" if r["기업규모"] == "확인필요"
                                  else ("금액미달" if not r.cond1
                                        else "규모조건 미충족"))),
            axis=1
        )
        grp.drop(columns=["cond1","cond2","cond3"], inplace=True)
        grp.sort_values("합산청구액", ascending=False, inplace=True)

        # 발행 대상 건수 상단 표시
        n_issue = (grp["어음대상"] == "발행대상").sum()
        ws.merge_cells("A1:F1")
        ws["A1"] = f"어음 발행 대상: {n_issue}건   ※ 비과세 + 2억 이상 + 중견/대기업 + 대사정상"
        ws["A1"].font      = Font(name=FONT, bold=True, size=12,
                                  color=("375623" if n_issue > 0 else "808080"))
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        COLS = ["주문번호", "협력사명", "합산청구액", "기업규모", "어음대상", "비고"]
        for c, h in enumerate(COLS, 1):
            hdr_cell(ws.cell(2, c, h))

        fill_map = {"발행대상": "CCFFCC", "확인필요": "FFFF99"}
        for r, (_, row) in enumerate(grp[COLS].iterrows(), 3):
            f = fill_map.get(row["어음대상"])
            for c, col in enumerate(COLS, 1):
                cell = ws.cell(r, c, row[col])
                data_cell(cell, fill=f, money=(col == "합산청구액"))

        set_col_widths = [22, 18, 20, 12, 12, 18]
        for i, w in enumerate(set_col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A3"

    # ── 시트4: 재무회계팀 자금예측보고 ────────────────────────
    def write_forecast_sheet(self, wb):
        self.log.info("  [시트4] 재무회계팀 자금예측보고 작성...")
        ws = wb.create_sheet("④자금예측보고")
        df = self.df_invoice.copy()
        amt_col  = self.pl_cols["정산금액"]
        tax_col  = self.pl_cols.get("매출과세구분")
        corp_col = self.pl_cols.get("협력사명")
        ps_s = self.ps.strftime("%Y.%m.%d") if self.ps else "-"
        pe_s = self.pe.strftime("%Y.%m.%d") if self.pe else "-"

        # 제목
        ws.merge_cells("A1:G1")
        ws["A1"] = f"재무회계팀  자금예측보고    정산기간: {ps_s} ~ {pe_s}"
        ws["A1"].font      = Font(name=FONT, bold=True, size=14, color="1F4E79")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        if not tax_col:
            ws["A3"] = "매출과세구분 컬럼 없음"
            return

        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        if tax_col not in df.columns:
            df[tax_col] = "미분류"

        # 어음 판별: 비과세 + 기업규모(중견/대기업)
        size_map = self.master.corp_size if (self.master and self.master.loaded) else {}
        def is_bill_row(row):
            if str(row.get(tax_col, "")) != "비과세":
                return False
            corp = str(row.get(corp_col, "")) if corp_col else ""
            return size_map.get(corp, "") in ("중견기업", "대기업")

        df["_수금방법"] = df.apply(lambda r: "어음" if is_bill_row(r) else "현금", axis=1)

        tax_vals = sorted(df[tax_col].dropna().unique().tolist())

        # ─ 구성 A: 피벗 집계 ─
        ws.cell(3, 1, "▶ 수금방법 × 과세구분 집계").font = Font(name=FONT, bold=True,
                                                                  size=11, color="1F4E79")
        # 헤더 행
        hdr_row = 4
        hdr_cell(ws.cell(hdr_row, 1, "수금방법"))
        for ci, tv in enumerate(tax_vals, 2):
            hdr_cell(ws.cell(hdr_row, ci, tv))
        hdr_cell(ws.cell(hdr_row, len(tax_vals) + 2, "합계"))
        ws.row_dimensions[hdr_row].height = 18

        data_rows_start = hdr_row + 1
        pivot_rows = []
        for ri, method in enumerate(["현금", "어음"], data_rows_start):
            pivot_rows.append(ri)
            sub = df[df["_수금방법"] == method]
            ws.cell(ri, 1, method).font = Font(name=FONT, size=10, bold=True)
            for ci, tv in enumerate(tax_vals, 2):
                val = sub[sub[tax_col] == tv][amt_col].sum()
                data_cell(ws.cell(ri, ci, val), money=True)
            # 합계 열: SUM 수식
            sum_range = f"{get_column_letter(2)}{ri}:{get_column_letter(len(tax_vals)+1)}{ri}"
            tc = ws.cell(ri, len(tax_vals) + 2, f"=SUM({sum_range})")
            total_cell(tc, money=True)

        # 합계 행 (각 열 SUM)
        total_row = data_rows_start + 2
        total_cell(ws.cell(total_row, 1, "합계"))
        for ci in range(2, len(tax_vals) + 3):
            col_l = get_column_letter(ci)
            tc = ws.cell(total_row, ci,
                         f"=SUM({col_l}{data_rows_start}:{col_l}{data_rows_start+1})")
            total_cell(tc, money=True)

        # ─ 구성 B: 수금 예측 안내 ─
        ann_row = total_row + 2
        ws.cell(ann_row, 1, "▶ 수금 예측 안내").font = Font(name=FONT, bold=True,
                                                              size=11, color="1F4E79")
        for i, txt in enumerate([
            "  현금: 정산확정 후 익월 말일 수금 예정",
            "  어음: 어음 발행일 기준 만기일 확인 필요",
        ], ann_row + 1):
            ws.cell(i, 1, txt).font = Font(name=FONT, size=10, italic=True, color="595959")

        auto_col_width(ws)
        ws.freeze_panes = "A5"

    # ── 메인 실행 ────────────────────────────────────────────────
    def run(self):
        try:
            self.detect_files()
            self.step0_init()
            self.step1_filter()
            self.step2_separate_returns()
            self.step3_reconcile_normal()
            self.step4_reconcile_returns()

            self.log.info("\n[XLOOKUP] 매핑 적용 중...")
            self.build_invoice_df()

            today_str   = self.today.strftime("%Y%m%d")
            result_path = f"정산결과_{today_str}.xlsx"

            self.log.info(f"\n[출력] {result_path}")
            wb = openpyxl.Workbook()
            self.write_invoice_sheet(wb)
            self.write_sales_sheet(wb)
            self.write_bill_sheet(wb)
            self.write_forecast_sheet(wb)
            wb.save(result_path)

            # ─ 콘솔 최종 요약 ─
            n_miss = len(self.df_missing)  if self.df_missing  is not None else 0
            n_ok   = len(self.df_matched)  if self.df_matched   is not None else 0
            n_diff = len(self.df_amtdiff)  if self.df_amtdiff  is not None else 0
            n_rok  = len(self.df_ret_ok)   if self.df_ret_ok   is not None else 0
            n_rng  = len(self.df_ret_ng)   if self.df_ret_ng   is not None else 0

            miss_flag = "  ← 확인 필요!" if n_miss > 0 else "  (없음)"
            self.log.info("")
            self.log.info("=" * 60)
            self.log.info("  실행 완료")
            n_pl = len(self.df_pl) if self.df_pl is not None else 0
            self.log.info(f"  정산 기간    : {self.ps} ~ {self.pe}")
            self.log.info(f"  KT 입고 건수 : {len(self.df_kt)}건  (구매문서번호 기준)")
            self.log.info(f"  플랫폼 건수  : {n_pl}건  (인보이스 기준)")
            self.log.info(f"  정상 매칭    : {n_ok}건")
            self.log.info(f"  플랫폼 누락  : {n_miss}건{miss_flag}")
            self.log.info(f"  반품         : 매칭 {n_rok}건 / 미확인 {n_rng}건  [경로 {self.ret_route}]")
            self.log.info(f"  금액 불일치  : {n_diff}건")
            if self.map_stats:
                for k, (ok, ng) in self.map_stats.items():
                    self.log.info(f"  {k:<12}: {ok}건 성공 / {ng}건 실패")
            self.log.info(f"  결과 파일    : {result_path}")
            self.log.info("=" * 60)

        except Exception as e:
            self.log.error(f"오류: {e}", exc_info=True)
            raise

# ── main ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    SettlementRunner(base_dir=".").run()
