"""
CPSM 수요예측 자동화 시스템
demand_forecast.py
"""

# ── 라이브러리 자동 설치 ──────────────────────────────────────────────────────
import subprocess, sys
from pathlib import Path   # _install 에서 wheels_dir 경로에 필요

# pip 패키지명 → 실제 import 모듈명 매핑
REQUIRED_PKGS = {
    "pandas":       "pandas",
    "numpy":        "numpy",
    "statsmodels":  "statsmodels",
    "scikit-learn": "sklearn",      # ← pip명과 모듈명이 다름
    "openpyxl":     "openpyxl",
    "xlsxwriter":   "xlsxwriter",
}

def _detect_win_proxy():
    """Windows 레지스트리에서 프록시 설정 읽기"""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        )
        enabled = winreg.QueryValueEx(key, "ProxyEnable")[0]
        if enabled:
            server = winreg.QueryValueEx(key, "ProxyServer")[0]
            if server and "=" not in server:          # 단일 프록시
                if not server.startswith("http"):
                    server = "http://" + server
                return server
            elif server:                              # 프로토콜별 분리 형식
                for part in server.split(";"):
                    if part.startswith("https="):
                        return "http://" + part.split("=",1)[1]
                    if part.startswith("http="):
                        return "http://" + part.split("=",1)[1]
    except Exception:
        pass
    # 환경변수 fallback
    import os
    return os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY") or None

def _install(pkg):
    proxy = _detect_win_proxy()
    wheels_dir = Path(__file__).parent / "wheels"

    # 시도 순서 구성
    attempts = []

    # 0순위: 로컬 wheels 폴더 (오프라인 우선)
    if wheels_dir.exists() and any(wheels_dir.iterdir()):
        attempts.append({
            "label": f"로컬 wheels 폴더 ({wheels_dir})",
            "extra": ["--no-index", "--find-links", str(wheels_dir)],
            "timeout": 30,
        })

    # 1순위: 프록시 경유 공식 PyPI
    if proxy:
        attempts.append({
            "label": f"공식 PyPI (프록시: {proxy})",
            "extra": ["--index-url", "https://pypi.org/simple",
                      "--trusted-host", "pypi.org",
                      "--trusted-host", "files.pythonhosted.org",
                      "--proxy", proxy],
            "timeout": 45,
        })

    # 2순위: 공식 PyPI 직접
    attempts.append({
        "label": "공식 PyPI (직접)",
        "extra": ["--index-url", "https://pypi.org/simple",
                  "--trusted-host", "pypi.org",
                  "--trusted-host", "files.pythonhosted.org"],
        "timeout": 45,
    })

    # 3순위: 사내 Nexus
    attempts.append({"label": "사내 Nexus", "extra": [], "timeout": 45})

    for att in attempts:
        try:
            cmd = ([sys.executable, "-m", "pip", "install", pkg,
                    "--timeout", "15", "-q"] + att["extra"])
            print(f"  → 시도: {att['label']}")
            subprocess.check_call(cmd, timeout=att["timeout"])
            print(f"  → [{pkg}] 설치 완료")
            return True
        except Exception as e:
            short = str(e)[:100]
            print(f"  → 실패: {short}")
            continue
    return False

# ── 패키지 확인 및 설치 ───────────────────────────────────────────────────────
_missing = []
for _pkg, _mod in REQUIRED_PKGS.items():
    try:
        __import__(_mod)
    except ImportError:
        _missing.append(_pkg)

if _missing:
    print(f"\n[설치 필요] {', '.join(_missing)}")
    _failed = []
    for _pkg in _missing:
        print(f"\n[설치] {_pkg} ...")
        ok = _install(_pkg)
        if not ok:
            _failed.append(_pkg)

    if _failed:
        print("\n" + "="*60)
        print("[수동 설치 안내]")
        print("아래 명령어를 인터넷이 되는 PC에서 실행해 .whl 파일을 받고,")
        print("이 폴더에 복사한 뒤 아래 명령어로 오프라인 설치하세요.\n")
        print("  # 인터넷 PC에서 다운로드:")
        dl_cmd = f"pip download {' '.join(_failed)} -d ./wheels --platform win_amd64 --python-version 3 --only-binary=:all:"
        print(f"  {dl_cmd}\n")
        print("  # 이 PC에서 오프라인 설치:")
        print(f"  pip install --no-index --find-links ./wheels {' '.join(_failed)}")
        print("="*60 + "\n")
        input("설치 완료 후 Enter 를 눌러 계속하거나, Ctrl+C 로 종료하세요: ")
    else:
        print("\n[완료] 모든 패키지 설치 성공\n")

# ── 임포트 ────────────────────────────────────────────────────────────────────
import os, re, warnings, logging
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

warnings.filterwarnings("ignore")

# ── 로그 설정 ─────────────────────────────────────────────────────────────────
LOG_FILE = Path(__file__).parent / "forecast_log.txt"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "forecast_config.json"

# ════════════════════════════════════════════════════════════════════════════
# 0-A. 설정 저장/로드
# ════════════════════════════════════════════════════════════════════════════
import json

def save_config(filepaths: list, lead_time: int, lt_file_path, col_map: dict):
    """분석 설정을 JSON 파일에 저장 (다음 실행 시 재사용)"""
    cfg = {
        "filepaths":   [str(p) for p in filepaths],
        "lead_time":   lead_time,
        "lt_file":     str(lt_file_path) if lt_file_path else None,
        "col_map":     col_map,
        "saved_at":    datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        log.info(f"설정 저장: {CONFIG_FILE.name}")
    except Exception as e:
        log.warning(f"설정 저장 실패: {e}")


def load_config() -> dict | None:
    """저장된 설정 로드. 파일 없거나 파싱 실패 시 None 반환."""
    if not CONFIG_FILE.exists():
        return None
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def prompt_use_saved_config(cfg: dict) -> bool:
    """저장된 설정 요약 출력 후 사용 여부 확인"""
    print("\n" + "=" * 65)
    print("  [저장된 설정 발견]")
    print(f"  저장일시: {cfg.get('saved_at', '알 수 없음')}")
    print()

    fps = cfg.get("filepaths", [])
    print(f"  파일 ({len(fps)}개):")
    for fp in fps:
        exists_mark = "✓" if Path(fp).exists() else "✗ (없음)"
        print(f"    {exists_mark}  {Path(fp).name}")

    lt_file = cfg.get("lt_file")
    if lt_file:
        lt_exists = "✓" if Path(lt_file).exists() else "✗ (없음)"
        print(f"  리드타임 파일: {lt_exists}  {Path(lt_file).name}")
    else:
        print(f"  리드타임: 기본값 {cfg.get('lead_time', 30)}일")

    cmap = cfg.get("col_map", {})
    print(f"  컬럼 매핑: date→{cmap.get('date','?')}, "
          f"amount→{cmap.get('amount','?')}, "
          f"qty→{cmap.get('qty','?')}, "
          f"vendor→{cmap.get('vendor','?')}")
    print("=" * 65)

    ans = input("\n이 설정으로 바로 실행하시겠습니까? (Y/n): ").strip().lower()
    return ans != "n"


# ════════════════════════════════════════════════════════════════════════════
# 0. 상수 및 스타일
# ════════════════════════════════════════════════════════════════════════════
Z_95 = 1.65   # 서비스수준 95%
Z_CI = 1.96   # 95% 신뢰구간

FONT_NAME = "맑은 고딕"
C_HEADER  = "1F4E79"; C_HFG = "FFFFFF"
C_ACTUAL  = "BDD7EE"   # 파란 계열 (실제값)
C_FCST    = "E2EFDA"   # 연두 (예측값)
C_A       = "FF9999"   # A등급
C_B       = "FFFF99"   # B등급
C_WARN    = "FFB347"   # MAPE>30%
C_TITLE   = "2E75B6"
C_ALT     = "DEEAF1"
C_TOTAL   = "FFF2CC"
NUM_FMT   = '#,##0'
PCT_FMT   = '0.0%'

def _side():
    return Side(style="thin", color="BFBFBF")

def tborder():
    return Border(left=_side(), right=_side(), top=_side(), bottom=_side())

def hcell(cell, text, bg=C_HEADER, fg=C_HFG, bold=True, wrap=True, size=10):
    cell.value = text
    cell.font  = Font(name=FONT_NAME, bold=bold, color=fg, size=size)
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = tborder()

def dcell(cell, value, fmt=None, bold=False, bg=None, align="right", size=10):
    cell.value = value
    cell.font  = Font(name=FONT_NAME, bold=bold, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = tborder()
    if fmt:  cell.number_format = fmt
    if bg:   cell.fill = PatternFill("solid", start_color=bg)

def cw(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w

def sheet_title(ws, title, sub=""):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    c = ws.cell(1, 1, title)
    c.font      = Font(name=FONT_NAME, bold=True, size=14, color="FFFFFF")
    c.fill      = PatternFill("solid", start_color=C_TITLE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=20)
        c2 = ws.cell(2, 1, sub)
        c2.font      = Font(name=FONT_NAME, size=10, color="595959")
        c2.fill      = PatternFill("solid", start_color="D6E4F0")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 16


# ════════════════════════════════════════════════════════════════════════════
# 1. 파일 탐지 & 컬럼 매핑
# ════════════════════════════════════════════════════════════════════════════
SEARCH_PATTERNS = [
    "platform*.xlsx", "platform*.csv",
    "매출*.xlsx", "정산*.xlsx",
    "*.xlsx", "*.csv",
]

# 결과·스크립트·임시 파일 제외 키워드
EXCLUDE_KEYWORDS = ["수요예측", "분석결과", "forecast", "sample", "verify_packages",
                    "리드타임_입력템플릿"]

def _is_data_file(p: Path) -> bool:
    # ~$ 접두어 = Excel 임시 잠금 파일
    if p.name.startswith("~$"):
        return False
    name_lower = p.name.lower()
    return not any(kw in name_lower for kw in EXCLUDE_KEYWORDS) \
           and p.name != "demand_forecast.py"

def find_input_files() -> list[Path]:
    """현재 폴더에서 분석 대상 데이터 파일 전체 탐지, 연도순 정렬"""
    seen = set()
    hits = []
    for pat in SEARCH_PATTERNS:
        for p in BASE_DIR.glob(pat):
            if p not in seen and _is_data_file(p):
                seen.add(p)
                hits.append(p)
    # 파일명 오름차순(연도 순서 유지)
    hits.sort(key=lambda p: p.name)
    return hits

COLUMN_ALIASES = {
    # ── 필수 ──────────────────────────────────────────────────────────────
    "date":      ["정산일자", "입고일자", "입고일", "주문일자", "주문일시", "주문일",
                  "정산확정일", "일일정산월", "일일정산일", "날짜", "일자", "date"],
    "amount":    ["판매금액", "정산금액", "주문금액", "매출액", "금액", "revenue", "amount"],
    "qty":       ["정산수량", "입고수량", "수량", "주문수량", "qty", "quantity"],
    "vendor":    ["협력사명", "협력사", "공급업체", "vendor", "supplier"],
    "prod_id":   ["상품코드", "상품ID", "품목코드", "product_id", "prod_id"],
    # ── 권장 ──────────────────────────────────────────────────────────────
    "prod_name": ["상품명", "품목명", "product_name"],
    "category":  ["서비스카테고리", "마스터카테고리(CMS)", "카테고리", "category"],
    "cost":      ["매입금액", "매입금액(원)", "purchase_amount"],       # 마진 분석
    "lt_order":  ["주문시배송리드타임", "주문리드타임", "order_lead_time"],  # 리드타임 실측
    "lt_std":    ["상품배송리드타임", "표준납기일", "standard_lead_time"],  # 표준 리드타임
    "cancel_qty":["취소수량", "cancel_qty"],                            # 실수요 보정
    "return_qty":["반품수량", "return_qty"],                            # 실수요 보정
    "order_date":["주문일자", "주문일시", "order_date"],                 # LT 실측 계산용
    "recv_date": ["입고일자", "입고일", "recv_date"],                    # LT 실측 계산용
    "dept":      ["부서명", "부서", "department"],                       # 부서별 분석
    "bizplace":  ["사업장", "site", "business_place"],                   # 사업장별 분석
    "order_type":["주문형태", "IP/DIP", "order_type"],                   # 유형별 분석
    "sale_type": ["매출구분", "sale_type"],                              # 매출 유형
    "regular":   ["정기주문여부", "is_regular", "regular_order"],        # 정기/비정기
    "tax":       ["과세상태", "부가세포함여부", "tax_status"],            # 세금 구분
}

def map_columns(df: pd.DataFrame) -> dict:
    found = {}
    cols_lower = {c.strip(): c for c in df.columns}
    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in cols_lower:
                found[key] = cols_lower[alias]
                break
    return found


# ════════════════════════════════════════════════════════════════════════════
# 2. 데이터 로드 & 전처리
# ════════════════════════════════════════════════════════════════════════════
def _detect_excel_engine(filepath: Path) -> str:
    """파일 헤더 매직바이트로 실제 Excel 포맷 판별 → 엔진 반환"""
    try:
        with open(filepath, "rb") as f:
            header = f.read(8)
        # OOXML (xlsx/xlsm): PK zip 시그니처
        if header[:2] == b"PK":
            return "openpyxl"
        # OLE2 compound (xls/xlsb): D0 CF 11 E0
        if header[:4] == b"\xd0\xcf\x11\xe0":
            return "xlrd"
    except Exception:
        pass
    # 확장자 기반 fallback
    ext = filepath.suffix.lower()
    return "openpyxl" if ext in (".xlsx", ".xlsm") else "xlrd"


def _read_excel_safe(filepath: Path, **kwargs) -> dict:
    """엔진 자동 감지 후 pd.read_excel 호출. xlrd 없으면 openpyxl 재시도."""
    engine = _detect_excel_engine(filepath)
    try:
        return pd.read_excel(filepath, engine=engine, **kwargs)
    except Exception as e1:
        # xlrd 미설치 등으로 실패 시 openpyxl 재시도
        if engine != "openpyxl":
            log.warning(f"  엔진({engine}) 실패, openpyxl 재시도: {e1}")
            return pd.read_excel(filepath, engine="openpyxl", **kwargs)
        raise


def _read_single_file(filepath: Path) -> pd.DataFrame:
    """단일 파일(xlsx/xls/csv) → DataFrame. 멀티시트는 전부 합친다."""
    if filepath.suffix.lower() == ".csv":
        # 인코딩 자동 감지 (utf-8-sig → cp949 순으로 시도)
        for enc in ("utf-8-sig", "cp949", "utf-8"):
            try:
                return pd.read_csv(filepath, encoding=enc)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(filepath, encoding="cp949", errors="replace")

    sheets = _read_excel_safe(filepath, sheet_name=None)
    frames = []
    for sdf in sheets.values():
        sdf.columns = sdf.columns.str.strip()
        frames.append(sdf)
    return pd.concat(frames, ignore_index=True)


def load_and_preprocess(filepaths: list[Path], col_map: dict, lead_time: int):
    """복수 파일을 읽어 하나의 DataFrame으로 병합 후 전처리"""
    all_frames = []
    for fp in filepaths:
        log.info(f"파일 로드: {fp.name}")
        try:
            raw = _read_single_file(fp)
            raw.columns = raw.columns.str.strip()
            raw["_source_file"] = fp.name   # 출처 파일 추적용
            all_frames.append(raw)
            log.info(f"  → {len(raw):,}행 읽음")
        except Exception as e:
            log.warning(f"  → 파일 읽기 실패 ({fp.name}): {e}")

    if not all_frames:
        print("[오류] 읽을 수 있는 파일이 없습니다.")
        sys.exit(1)

    df = pd.concat(all_frames, ignore_index=True)
    df.columns = df.columns.str.strip()
    log.info(f"  전체 병합 행수: {len(df):,}")

    # 날짜
    date_col = col_map["date"]
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    if df["_date"].isna().all():
        def parse_kor(s):
            try:
                m = int(str(s).replace("월","").split()[0])
                return pd.Timestamp(f"2024-{m:02d}-01")
            except Exception:
                return pd.NaT
        df["_date"] = df[date_col].apply(parse_kor)
    df.dropna(subset=["_date"], inplace=True)
    df["_ym"] = df["_date"].dt.to_period("M")

    # 중복 제거 (동일 날짜+금액+수량 행이 여러 파일에 걸쳐 중복될 경우)
    key_cols = [date_col, col_map["amount"]]
    if col_map.get("vendor"):  key_cols.append(col_map["vendor"])
    before = len(df)
    df.drop_duplicates(subset=key_cols, inplace=True)
    dropped = before - len(df)
    if dropped:
        log.info(f"  중복 제거: {dropped:,}행 삭제")

    # 금액 / 수량
    df["_amount"] = pd.to_numeric(df[col_map["amount"]], errors="coerce").fillna(0)
    df["_qty"]    = pd.to_numeric(df[col_map["qty"]], errors="coerce").fillna(0) \
                    if col_map.get("qty") else 0

    # 식별자 컬럼 구성
    if col_map.get("prod_id"):
        df["_prod_key"] = df[col_map["prod_id"]].astype(str).str.strip()
    else:
        cat = df[col_map["category"]].astype(str).str.strip() if col_map.get("category") else "N/A"
        pn  = df[col_map["prod_name"]].astype(str).str.strip() if col_map.get("prod_name") else "N/A"
        df["_prod_key"] = cat.str[:20] + "|" + pn.str[:20]

    df["_vendor"] = df[col_map["vendor"]].astype(str).str.strip() if col_map.get("vendor") else "미분류"

    # ── 실수요 보정: 취소·반품 차감 ──────────────────────────────────────
    if col_map.get("cancel_qty"):
        cq = pd.to_numeric(df[col_map["cancel_qty"]], errors="coerce").fillna(0)
        df["_qty"] = (df["_qty"] - cq).clip(lower=0)
    if col_map.get("return_qty"):
        rq = pd.to_numeric(df[col_map["return_qty"]], errors="coerce").fillna(0)
        df["_qty"]    = (df["_qty"]    - rq).clip(lower=0)
        df["_amount"] = (df["_amount"] * (df["_qty"] /
                          (df["_qty"] + rq).replace(0, np.nan)).fillna(1)).fillna(0)

    # ── 리드타임 실측값 추출 ─────────────────────────────────────────────
    # 우선순위: ① 주문시배송리드타임 ② 상품배송리드타임 ③ 입고일-주문일 계산
    if col_map.get("lt_order"):
        df["_lt"] = pd.to_numeric(df[col_map["lt_order"]], errors="coerce")
    elif col_map.get("lt_std"):
        df["_lt"] = pd.to_numeric(df[col_map["lt_std"]], errors="coerce")
    elif col_map.get("order_date") and col_map.get("recv_date"):
        odt = pd.to_datetime(df[col_map["order_date"]], errors="coerce")
        rdt = pd.to_datetime(df[col_map["recv_date"]],  errors="coerce")
        df["_lt"] = (rdt - odt).dt.days
    else:
        df["_lt"] = np.nan

    # 이상값 제거: 0 이하 또는 365일 초과
    df.loc[df["_lt"] <= 0,   "_lt"] = np.nan
    df.loc[df["_lt"] > 365,  "_lt"] = np.nan

    # ── 매입금액 (마진 분석용) ────────────────────────────────────────────
    if col_map.get("cost"):
        df["_cost"] = pd.to_numeric(df[col_map["cost"]], errors="coerce").fillna(0)

    log.info(f"  최종 행수: {len(df):,}  |  기간: {df['_ym'].min()} ~ {df['_ym'].max()}")
    lt_valid = df["_lt"].notna().sum()
    if lt_valid > 0:
        log.info(f"  리드타임 실측값: {lt_valid:,}건 (평균 {df['_lt'].mean():.1f}일, "
                 f"σ={df['_lt'].std():.1f}일)")
    return df


def build_monthly_series(df: pd.DataFrame, group_cols=None) -> pd.DataFrame:
    """group_cols=None → 전체 집계, list → 그룹별"""
    if group_cols:
        grp = df.groupby(group_cols + ["_ym"]).agg(
            amount=("_amount", "sum"),
            qty=("_qty", "sum"),
        ).reset_index()
    else:
        grp = df.groupby("_ym").agg(
            amount=("_amount", "sum"),
            qty=("_qty", "sum"),
        ).reset_index()
    return grp


def fill_missing_months(series: pd.Series, min_period, max_period) -> pd.Series:
    """결측 연월 0으로 보간"""
    full_idx = pd.period_range(min_period, max_period, freq="M")
    return series.reindex(full_idx, fill_value=0)


def trim_to_36months(df_ym: pd.DataFrame, ym_col="_ym") -> pd.DataFrame:
    max_ym = df_ym[ym_col].max()
    cutoff  = max_ym - 35
    return df_ym[df_ym[ym_col] >= cutoff].copy()


def winsorize_series(arr: np.ndarray, pct: float = 0.05) -> np.ndarray:
    """상·하위 pct% 이상치를 경계값으로 클리핑 (이상치가 MAPE를 왜곡하는 것 방지)"""
    if len(arr) < 6:
        return arr
    lo = np.percentile(arr, pct * 100)
    hi = np.percentile(arr, (1 - pct) * 100)
    return np.clip(arr, lo, hi)


def check_consecutive_zeros(series: pd.Series, threshold=3) -> bool:
    """연속 0이 threshold개 이상이면 True(제외 대상)"""
    count = 0
    for v in series:
        if v == 0:
            count += 1
            if count >= threshold:
                return True
        else:
            count = 0
    return False


# ════════════════════════════════════════════════════════════════════════════
# 3. ABC × CV 분류
# ════════════════════════════════════════════════════════════════════════════
def classify_abc_cv(group_df: pd.DataFrame) -> pd.DataFrame:
    """
    group_df: vendor × prod_key 기준 월별 집계 롱포맷
    반환: 분류 결과 DataFrame
    """
    # 그룹별 총 금액
    summary = (
        group_df.groupby(["_vendor", "_prod_key"])
        .agg(total_amount=("amount", "sum"),
             mean_qty=("qty", "mean"),
             std_qty=("qty", "std"),
             n_months=("amount", "count"))
        .reset_index()
    )
    summary["std_qty"]  = summary["std_qty"].fillna(0)
    summary["cv"]       = np.where(
        summary["mean_qty"] > 0,
        summary["std_qty"] / summary["mean_qty"],
        0
    )

    # ABC
    summary.sort_values("total_amount", ascending=False, inplace=True)
    total = summary["total_amount"].sum()
    summary["cum_pct"] = summary["total_amount"].cumsum() / total * 100
    summary["ABC"] = summary["cum_pct"].apply(
        lambda x: "A" if x <= 80 else ("B" if x <= 95 else "C")
    )

    # CV 구분
    summary["CV_class"] = summary["cv"].apply(
        lambda x: "안정" if x < 0.2 else ("보통" if x <= 0.5 else "불안정")
    )

    # 예측 전략
    def strategy(row):
        abc, cv_c = row["ABC"], row["CV_class"]
        if abc == "A" and cv_c in ("안정", "보통"):
            return "Holt-Winters"
        elif abc == "A" and cv_c == "불안정":
            return "SARIMA"
        elif abc == "B":
            return "선형회귀"
        else:
            return "MA3"
    summary["strategy"] = summary.apply(strategy, axis=1)

    return summary


# ════════════════════════════════════════════════════════════════════════════
# 4. 예측 모델
# ════════════════════════════════════════════════════════════════════════════
def _mape(actual, pred):
    actual, pred = np.array(actual), np.array(pred)
    mask = actual != 0
    if mask.sum() == 0:
        return np.nan
    return np.mean(np.abs((actual[mask] - pred[mask]) / actual[mask])) * 100

def _mae(actual, pred):
    return np.mean(np.abs(np.array(actual) - np.array(pred)))

def _rmse(actual, pred):
    return np.sqrt(np.mean((np.array(actual) - np.array(pred))**2))


def model_holt_winters(train: np.ndarray, h: int):
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    n = len(train)
    sp = 12 if n >= 24 else (6 if n >= 12 else None)
    if sp is None or n < sp * 2:
        raise ValueError("HW: 데이터 부족")
    try:
        m = ExponentialSmoothing(
            train, trend="add", seasonal="add", seasonal_periods=sp,
            initialization_method="estimated",
        ).fit(optimized=True, use_brute=False)
        fc = m.forecast(h)
        resid_std = np.std(m.resid)
        return fc, resid_std
    except Exception as e:
        raise ValueError(f"HW 실패: {e}")


def model_regression(train: np.ndarray, h: int):
    from sklearn.linear_model import LinearRegression
    n = len(train)
    t = np.arange(n)
    months = np.arange(n) % 12  # 0~11

    def make_X(t_arr):
        X = np.column_stack([t_arr, t_arr**2] +
                            [(months_arr == i).astype(int) for i in range(11)
                             for months_arr in [t_arr % 12]])
        return X

    # 더미 생성 올바르게
    def make_features(t_arr):
        feats = [t_arr, t_arr**2]
        for i in range(11):
            feats.append((t_arr % 12 == i).astype(int))
        return np.column_stack(feats)

    X_train = make_features(t)
    reg = LinearRegression().fit(X_train, train)
    t_fc = np.arange(n, n + h)
    X_fc = make_features(t_fc)
    fc = reg.predict(X_fc)
    resid_std = np.std(train - reg.predict(X_train))
    return fc, resid_std


def model_sarima(train: np.ndarray, h: int):
    from statsmodels.tsa.statespace.sarimax import SARIMAX
    if len(train) < 24:
        raise ValueError("SARIMA: 24개월 미만")
    try:
        m = SARIMAX(train, order=(1,1,1), seasonal_order=(1,1,0,12),
                    enforce_stationarity=False, enforce_invertibility=False).fit(
            disp=False, maxiter=200)
        fc_obj = m.get_forecast(steps=h)
        fc = fc_obj.predicted_mean
        ci = fc_obj.conf_int(alpha=0.05)
        resid_std = np.std(m.resid)
        return fc, resid_std, ci
    except Exception as e:
        raise ValueError(f"SARIMA 실패: {e}")


def model_ma3(train: np.ndarray, h: int):
    val = float(np.mean(train[-3:])) if len(train) >= 3 else float(np.mean(train))
    fc = np.full(h, val)
    resid_std = float(np.std(train[-6:])) if len(train) >= 6 else float(np.std(train))
    return fc, resid_std


# ════════════════════════════════════════════════════════════════════════════
# 5. Walk-forward 검증 & 최적 모델 선택
# ════════════════════════════════════════════════════════════════════════════
def walk_forward_eval(series: np.ndarray, strategy: str, n_test=6):
    if len(series) <= n_test + 6:
        n_test = max(1, len(series) // 4)

    # 이상치 윈소라이징 (상하위 5% 클리핑) — 일시적 급등락이 MAPE 왜곡 방지
    series_w = winsorize_series(series.copy(), pct=0.05)

    train, test = series_w[:-n_test], series_w[-n_test:]
    h = len(test)

    results = {}

    # Holt-Winters
    try:
        fc_hw, _ = model_holt_winters(train, h)
        results["Holt-Winters"] = {
            "mape": _mape(test, fc_hw),
            "mae":  _mae(test, fc_hw),
            "rmse": _rmse(test, fc_hw),
        }
    except Exception as e:
        log.debug(f"HW 검증 실패: {e}")
        results["Holt-Winters"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

    # 선형회귀
    try:
        fc_reg, _ = model_regression(train, h)
        results["선형회귀"] = {
            "mape": _mape(test, fc_reg),
            "mae":  _mae(test, fc_reg),
            "rmse": _rmse(test, fc_reg),
        }
    except Exception as e:
        log.debug(f"회귀 검증 실패: {e}")
        results["선형회귀"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

    # SARIMA
    try:
        fc_sa, _, _ = model_sarima(train, h)
        results["SARIMA"] = {
            "mape": _mape(test, fc_sa),
            "mae":  _mae(test, fc_sa),
            "rmse": _rmse(test, fc_sa),
        }
    except Exception as e:
        log.debug(f"SARIMA 검증 실패: {e}")
        results["SARIMA"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

    # MA3
    try:
        fc_ma, _ = model_ma3(train, h)
        results["MA3"] = {
            "mape": _mape(test, fc_ma),
            "mae":  _mae(test, fc_ma),
            "rmse": _rmse(test, fc_ma),
        }
    except Exception:
        results["MA3"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

    # 최적 모델 선택 (MAPE 기준, NaN 제외)
    valid = {k: v for k, v in results.items() if not np.isnan(v["mape"])}
    if not valid:
        best = strategy  # 검증 불가 → 전략 기본값
    else:
        # 전략 우선 + MAPE 최소
        if strategy in valid:
            candidates = {strategy: valid[strategy]}
        else:
            candidates = valid
        best = min(candidates, key=lambda k: candidates[k]["mape"])
        # 전략 외 다른 모델이 30% 이상 더 낫지 않으면 전략 유지
        if strategy in valid:
            strat_mape = valid[strategy]["mape"]
            global_best = min(valid, key=lambda k: valid[k]["mape"])
            if valid[global_best]["mape"] < strat_mape * 0.7:
                best = global_best

    return results, best


def forecast_series(series: np.ndarray, best_model: str, h=6):
    """최적 모델로 h개월 예측 + 신뢰구간"""
    ci_lo, ci_hi = None, None
    try:
        if best_model == "Holt-Winters":
            fc, rs = model_holt_winters(series, h)
        elif best_model == "선형회귀":
            fc, rs = model_regression(series, h)
        elif best_model == "SARIMA":
            res = model_sarima(series, h)
            fc, rs = res[0], res[1]
            ci = res[2]
            ci_lo = ci.iloc[:, 0].values
            ci_hi = ci.iloc[:, 1].values
        else:
            fc, rs = model_ma3(series, h)
    except Exception as e:
        log.warning(f"예측 모델 폴백(MA3): {e}")
        fc, rs = model_ma3(series, h)

    fc = np.maximum(fc, 0)
    if ci_lo is None:
        ci_lo = np.maximum(fc - Z_CI * rs, 0)
        ci_hi = fc + Z_CI * rs
    return fc, ci_lo, ci_hi


# ════════════════════════════════════════════════════════════════════════════
# 6. CPSM 지표
# ════════════════════════════════════════════════════════════════════════════
DEFAULT_LEAD_TIME = 30   # 리드타임 파일에 없는 품목의 기본값 (일)

def build_lt_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    원본 데이터에서 협력사×상품코드별 리드타임 통계 산출.
    반환: vendor, prod_key, lt_mean(일), lt_std(일), lt_n(건수)
    """
    if "_lt" not in df.columns or df["_lt"].notna().sum() == 0:
        return pd.DataFrame(columns=["vendor", "prod_key", "lt_mean", "lt_std", "lt_n"])

    lt_grp = (
        df[df["_lt"].notna()]
        .groupby(["_vendor", "_prod_key"])["_lt"]
        .agg(lt_mean="mean", lt_std="std", lt_n="count")
        .reset_index()
        .rename(columns={"_vendor": "vendor", "_prod_key": "prod_key"})
    )
    lt_grp["lt_std"] = lt_grp["lt_std"].fillna(0)
    log.info(f"리드타임 통계: {len(lt_grp)}개 품목 산출 "
             f"(전체 평균 {lt_grp['lt_mean'].mean():.1f}일)")
    return lt_grp


def get_lt_stats(lt_stats: pd.DataFrame, lt_df: pd.DataFrame,
                 vendor: str, prod_key: str) -> tuple[float, float]:
    """
    품목별 리드타임 (평균, 표준편차) 반환.
    우선순위: ① 실측 통계 ② 리드타임 파일 ③ 전체 평균 ④ 기본값
    """
    # ① 실측 통계
    if not lt_stats.empty:
        row = lt_stats[(lt_stats["vendor"] == vendor) &
                       (lt_stats["prod_key"] == prod_key)]
        if not row.empty:
            return float(row.iloc[0]["lt_mean"]), float(row.iloc[0]["lt_std"])
        # 협력사 평균
        row_v = lt_stats[lt_stats["vendor"] == vendor]
        if not row_v.empty:
            return float(row_v["lt_mean"].mean()), float(row_v["lt_std"].mean())
        # 전체 평균
        return float(lt_stats["lt_mean"].mean()), float(lt_stats["lt_std"].mean())

    # ② 리드타임 파일 (std=0 가정)
    lt_file_val = get_lead_time(lt_df, vendor, prod_key)
    return float(lt_file_val), 0.0

def load_lead_time_file(filepath: Path) -> pd.DataFrame:
    """
    리드타임 파일 로드.
    필수 컬럼: 협력사명(또는 vendor), 상품코드(또는 prod_id), 리드타임
    선택 컬럼: 상품명, 비고
    """
    try:
        raw = _read_excel_safe(filepath, sheet_name=0) \
              if filepath.suffix.lower() != ".csv" \
              else pd.read_csv(filepath, encoding="utf-8-sig")
        raw.columns = raw.columns.str.strip()

        # 컬럼 자동 매핑
        col_map = {}
        for key, aliases in {
            "vendor":    ["협력사명", "협력사", "공급업체", "vendor"],
            "prod_key":  ["상품코드", "상품ID", "품목코드", "product_id", "prod_id"],
            "lead_time": ["리드타임", "리드타임(일)", "lead_time", "납기일수", "납기", "LT"],
        }.items():
            for a in aliases:
                if a in raw.columns:
                    col_map[key] = a
                    break

        missing = [k for k in ["vendor", "prod_key", "lead_time"] if k not in col_map]
        if missing:
            log.warning(f"리드타임 파일 컬럼 미발견: {missing} → 기본값 {DEFAULT_LEAD_TIME}일 사용")
            return pd.DataFrame()

        lt_df = raw[[col_map["vendor"], col_map["prod_key"], col_map["lead_time"]]].copy()
        lt_df.columns = ["vendor", "prod_key", "lead_time"]
        lt_df["vendor"]    = lt_df["vendor"].astype(str).str.strip()
        lt_df["prod_key"]  = lt_df["prod_key"].astype(str).str.strip()
        lt_df["lead_time"] = pd.to_numeric(lt_df["lead_time"], errors="coerce").fillna(DEFAULT_LEAD_TIME)
        lt_df = lt_df[lt_df["lead_time"] > 0]
        log.info(f"리드타임 파일 로드: {len(lt_df)}개 품목")
        return lt_df

    except Exception as e:
        log.warning(f"리드타임 파일 읽기 실패: {e} → 기본값 {DEFAULT_LEAD_TIME}일 사용")
        return pd.DataFrame()


def get_lead_time(lt_df: pd.DataFrame, vendor: str, prod_key: str) -> int:
    """협력사×상품코드로 리드타임 조회. 없으면 협력사만, 그것도 없으면 기본값."""
    if lt_df.empty:
        return DEFAULT_LEAD_TIME
    # 정확 매칭
    row = lt_df[(lt_df["vendor"] == vendor) & (lt_df["prod_key"] == prod_key)]
    if not row.empty:
        return int(row.iloc[0]["lead_time"])
    # 협력사만 매칭 (평균)
    row_v = lt_df[lt_df["vendor"] == vendor]
    if not row_v.empty:
        return int(row_v["lead_time"].mean())
    return DEFAULT_LEAD_TIME


def make_lead_time_template(abc_df: pd.DataFrame, out_path: Path):
    """분석 결과 기반 리드타임 입력 템플릿 Excel 생성"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "리드타임_입력"

    headers = ["협력사명", "상품코드", "ABC등급", "리드타임(일)", "비고"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(1, c), h)
    ws.row_dimensions[1].height = 24

    # iterrows() 사용 — _로 시작하는 컬럼명도 안전하게 접근
    for r, (_, row) in enumerate(abc_df.iterrows(), 2):
        bg = C_A if row["ABC"] == "A" else (C_B if row["ABC"] == "B" else None)
        dcell(ws.cell(r, 1), str(row["_vendor"]),   align="left", bg=bg)
        dcell(ws.cell(r, 2), str(row["_prod_key"]), align="left", bg=bg)
        dcell(ws.cell(r, 3), row["ABC"], align="center", bold=True, bg=bg)
        ws.cell(r, 4).value     = DEFAULT_LEAD_TIME
        ws.cell(r, 4).fill      = PatternFill("solid", start_color="FFFACD")
        ws.cell(r, 4).border    = tborder()
        ws.cell(r, 4).font      = Font(name=FONT_NAME, size=10)
        ws.cell(r, 4).alignment = Alignment(horizontal="center", vertical="center")
        dcell(ws.cell(r, 5), "", align="left", bg=bg)

    for c, w in [(1, 25), (2, 30), (3, 10), (4, 14), (5, 30)]:
        cw(ws, c, w)

    note_r = len(abc_df) + 3
    for i, txt in enumerate([
        "※ 노란색 셀(리드타임)을 품목별로 수정 후 저장하세요.",
        "※ 파일명을 '리드타임_입력템플릿.xlsx' 로 유지하면 다음 실행 시 자동 인식됩니다.",
        f"※ 목록에 없는 품목은 기본값 {DEFAULT_LEAD_TIME}일이 적용됩니다.",
    ]):
        ws.cell(note_r + i, 1, txt).font = Font(name=FONT_NAME, size=9,
                                                italic=True, color="595959")
    wb.save(out_path)
    log.info(f"리드타임 템플릿 생성: {out_path.name}")


def calc_safety_stock(qty_series: np.ndarray,
                      lt_mean: float, lt_std: float = 0.0) -> tuple[float, str]:
    """
    리드타임 변동성을 반영한 완전한 안전재고 공식 (단위: 수량/월 기준).

    수요 변동만 있을 때 (기존):
        SS = Z × σ_D × √LT_months

    리드타임 변동성도 있을 때 (개선):
        SS = Z × √(LT_months × σ_D² + D_avg² × σ_LT_months²)

    - D_avg       : 최근 12개월 월평균 수요
    - σ_D         : 최근 12개월 수요 표준편차 (월)
    - LT_months   : 평균 리드타임 (일 → 월 환산: /30)
    - σ_LT_months : 리드타임 표준편차 (일 → 월 환산: /30)
    """
    window = qty_series[-12:] if len(qty_series) >= 12 else qty_series
    d_avg  = float(np.mean(window))
    sigma_d = float(np.std(window))

    lt_m     = lt_mean / 30          # 일 → 월
    sigma_lt = lt_std  / 30

    if sigma_lt > 0:
        ss = Z_95 * np.sqrt(lt_m * sigma_d**2 + d_avg**2 * sigma_lt**2)
        formula = "완전공식(수요+LT변동)"
    else:
        ss = Z_95 * sigma_d * np.sqrt(lt_m)
        formula = "기본공식(수요변동만)"

    return max(ss, 0.0), formula


def calc_rop(qty_series: np.ndarray, lt_mean: float, ss: float) -> float:
    """ROP = 평균일수요 × 리드타임(일) + 안전재고"""
    window = qty_series[-12:] if len(qty_series) >= 12 else qty_series
    avg_daily = float(np.mean(window)) / 30
    return avg_daily * lt_mean + ss


# ════════════════════════════════════════════════════════════════════════════
# 7. Excel 출력
# ════════════════════════════════════════════════════════════════════════════
def write_sheet1_overall(wb, overall_actual: pd.DataFrame, fc_vals, ci_lo, ci_hi,
                          best_model: str, model_mape: float, future_periods):
    ws = wb.create_sheet("전체_월별예측")
    sheet_title(ws, "전체 월별 매출 예측",
                f"사용모델: {best_model}  |  MAPE: {model_mape:.1f}%")
    SR = 4
    headers = ["연월", "실제정산금액(원)", "실제주문수량",
               "예측정산금액(원)", "예측하한(95%CI)", "예측상한(95%CI)",
               "사용모델", "MAPE(%)"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 30

    # 실제값 행
    r = SR + 1
    for _, row in overall_actual.iterrows():
        bg = C_ACTUAL
        dcell(ws.cell(r, 1), str(row["_ym"]), align="center", bg=bg)
        dcell(ws.cell(r, 2), row["amount"], NUM_FMT, bg=bg)
        dcell(ws.cell(r, 3), row["qty"],    NUM_FMT, bg=bg)
        dcell(ws.cell(r, 4), None, bg=bg)
        dcell(ws.cell(r, 5), None, bg=bg)
        dcell(ws.cell(r, 6), None, bg=bg)
        dcell(ws.cell(r, 7), best_model, align="center", bg=bg)
        dcell(ws.cell(r, 8), None, bg=bg)
        r += 1

    # 예측값 행
    for i, period in enumerate(future_periods):
        bg = C_FCST
        warn_bg = C_WARN if model_mape > 30 else C_FCST
        dcell(ws.cell(r, 1), str(period),   align="center", bg=bg)
        dcell(ws.cell(r, 2), None, bg=bg)
        dcell(ws.cell(r, 3), None, bg=bg)
        dcell(ws.cell(r, 4), float(fc_vals[i]),  NUM_FMT, bg=bg)
        dcell(ws.cell(r, 5), float(ci_lo[i]),    NUM_FMT, bg=bg)
        dcell(ws.cell(r, 6), float(ci_hi[i]),    NUM_FMT, bg=bg)
        dcell(ws.cell(r, 7), best_model, align="center", bg=bg)
        mape_disp = f"{model_mape:.1f}%" if not np.isnan(model_mape) else "N/A"
        dcell(ws.cell(r, 8), mape_disp, align="center", bg=warn_bg)
        r += 1

    for c, w in enumerate([14,20,16,20,20,20,16,12], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"

    # 차트
    n_total = len(overall_actual) + len(future_periods)
    chart = LineChart()
    chart.title = "월별 정산금액 실제 vs 예측"
    chart.style = 10; chart.width = 28; chart.height = 14
    chart.add_data(Reference(ws, min_col=2, min_row=SR, max_row=SR+n_total), titles_from_data=True)
    chart.add_data(Reference(ws, min_col=4, min_row=SR, max_row=SR+n_total), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=SR+1, max_row=SR+n_total))
    ws.add_chart(chart, f"J{SR}")


def write_sheet2_abc(wb, abc_df: pd.DataFrame, ss_rop: dict):
    ws = wb.create_sheet("ABC_CV_분류표")
    sheet_title(ws, "ABC × CV 분류표 (CPSM 공급망 분석)")
    SR = 4
    headers = ["협력사명", "상품코드", "ABC등급", "CV값", "CV구분",
               "연간정산금액(원)", "누적기여율(%)", "예측전략",
               "LT평균(일)", "LT표준편차(일)", "안전재고공식", "안전재고(수량)", "ROP(수량)"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 30

    for r, row in enumerate(abc_df.itertuples(), SR+1):
        abc = row.ABC
        bg = C_A if abc == "A" else (C_B if abc == "B" else None)
        key = (row._vendor, row._prod_key)
        d       = ss_rop.get(key, {})
        ss_val  = d.get("ss", 0)
        rop_val = d.get("rop", 0)
        lt_val  = d.get("lead_time", DEFAULT_LEAD_TIME)
        lt_s    = d.get("lt_std", 0)
        formula = d.get("formula", "-")

        dcell(ws.cell(r, 1), row._vendor,    align="left", bg=bg)
        dcell(ws.cell(r, 2), row._prod_key,  align="left", bg=bg)
        dcell(ws.cell(r, 3), abc, align="center", bg=bg, bold=True)
        dcell(ws.cell(r, 4), round(row.cv, 3), "0.000", bg=bg)
        dcell(ws.cell(r, 5), row.CV_class, align="center", bg=bg)
        dcell(ws.cell(r, 6), row.total_amount, NUM_FMT, bg=bg)
        dcell(ws.cell(r, 7), round(row.cum_pct, 2), "0.00%", bg=bg)
        dcell(ws.cell(r, 8), row.strategy, align="center", bg=bg)
        # LT평균 — 기본값이면 연한 노란 배경
        lt_bg = "FFF9C4" if abs(lt_val - DEFAULT_LEAD_TIME) < 0.1 and lt_s == 0 else bg
        dcell(ws.cell(r, 9),  round(lt_val, 1), "0.0", bg=lt_bg)
        dcell(ws.cell(r, 10), round(lt_s, 1),   "0.0",
              bg="E8F5E9" if lt_s > 0 else bg)   # 연두 = 실측 변동성 반영
        dcell(ws.cell(r, 11), formula, align="center",
              bg="E8F5E9" if "완전" in formula else bg)
        dcell(ws.cell(r, 12), round(ss_val, 1),  "0.0", bg=bg)
        dcell(ws.cell(r, 13), round(rop_val, 1), "0.0", bg=bg)

    for c, w in enumerate([22, 28, 10, 10, 12, 20, 14, 14, 12, 14, 18, 14, 14], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"


def write_sheet3_detail(wb, detail_rows: list):
    ws = wb.create_sheet("협력사별_상품코드_예측")
    sheet_title(ws, "협력사 × 상품코드별 향후 6개월 예측 (A등급 전체 + B등급 상위 20개)")
    SR = 4
    headers = ["협력사명", "상품코드", "ABC등급",
               "최근12개월평균(원)",
               "예측+1M(원)", "예측+2M(원)", "예측+3M(원)",
               "예측+4M(원)", "예측+5M(원)", "예측+6M(원)",
               "사용모델", "MAPE(%)", "신뢰도"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 30

    for r, dr in enumerate(detail_rows, SR+1):
        abc = dr["ABC"]
        mape_val = dr.get("mape", np.nan)
        bg_base = C_A if abc == "A" else (C_B if abc == "B" else None)
        bg_warn = C_WARN if (not np.isnan(mape_val) and mape_val > 30) else bg_base

        dcell(ws.cell(r, 1), dr["vendor"],   align="left",   bg=bg_base)
        dcell(ws.cell(r, 2), dr["prod_key"], align="left",   bg=bg_base)
        dcell(ws.cell(r, 3), abc, align="center", bold=True, bg=bg_base)
        dcell(ws.cell(r, 4), dr["avg12"],    NUM_FMT,        bg=bg_base)
        for i, fc_v in enumerate(dr["fc6"], 1):
            dcell(ws.cell(r, 4+i), float(fc_v), NUM_FMT, bg=C_FCST)
        dcell(ws.cell(r, 11), dr["model"], align="center", bg=bg_base)
        mape_str = f"{mape_val:.1f}%" if not np.isnan(mape_val) else "N/A"
        dcell(ws.cell(r, 12), mape_str, align="center", bg=bg_warn)
        flag = "⚠ 낮음" if (not np.isnan(mape_val) and mape_val > 30) else "✓ 양호"
        dcell(ws.cell(r, 13), flag, align="center", bg=bg_warn)

    for c, w in enumerate([22,28,10,18,16,16,16,16,16,16,14,12,12], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"


def write_sheet4_accuracy(wb, acc_rows: list):
    ws = wb.create_sheet("모델_정확도_비교")
    sheet_title(ws, "모델별 예측 정확도 비교 (Walk-forward 검증)")
    SR = 4
    headers = ["시계열명", "HW MAPE(%)", "회귀 MAPE(%)", "SARIMA MAPE(%)",
               "MA3 MAPE(%)", "최적모델", "비고"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 28

    for r, row in enumerate(acc_rows, SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), row["name"], align="left", bg=bg)
        for ci, key in enumerate(["Holt-Winters","선형회귀","SARIMA","MA3"], 2):
            v = row.get(key, np.nan)
            dcell(ws.cell(r, ci), round(v,1) if not np.isnan(v) else "N/A",
                  "0.0" if not isinstance(v, str) else None,
                  bg=C_WARN if (not isinstance(v,str) and not np.isnan(v) and v>30) else bg)
        dcell(ws.cell(r, 6), row["best"], align="center", bg=bg, bold=True)
        dcell(ws.cell(r, 7), row.get("note",""), align="left", bg=bg)

    for c, w in enumerate([35,14,14,14,14,14,20], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"


def write_sheet5_dashboard(wb, summary: dict):
    ws = wb.create_sheet("CPSM_요약대시보드", 0)
    sheet_title(ws, "CPSM 수요예측 요약 대시보드",
                f"실행일시: {summary['run_dt']}  |  데이터 기간: {summary['data_range']}")

    kpis = [
        ("분석기간",         summary["data_range"],         "1F4E79"),
        ("전체 시계열 수",    summary["n_series"],           "2E75B6"),
        ("A등급 품목 수",     summary["n_A"],                "375623"),
        ("향후6개월 예측총액", f"{summary['total_fc6']:,.0f}원", "843C0C"),
        ("평균 예측정확도",   f"MAPE {summary['avg_mape']:.1f}%", "7B2D8B"),
        ("주의 시계열 수",    summary["n_warn"],             "C00000"),
    ]
    for i, (label, val, bg) in enumerate(kpis):
        col = (i % 3) * 4 + 1
        row_base = 4 + (i // 3) * 3
        ws.merge_cells(start_row=row_base,   start_column=col, end_row=row_base,   end_column=col+2)
        ws.merge_cells(start_row=row_base+1, start_column=col, end_row=row_base+1, end_column=col+2)
        lc = ws.cell(row_base, col, label)
        lc.font      = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
        lc.fill      = PatternFill("solid", start_color=bg)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border    = tborder()
        for adj in range(1, 3):
            ws.cell(row_base, col+adj).fill   = PatternFill("solid", start_color=bg)
            ws.cell(row_base, col+adj).border = tborder()
            ws.cell(row_base+1, col+adj).fill   = PatternFill("solid", start_color="F2F2F2")
            ws.cell(row_base+1, col+adj).border = tborder()
        vc = ws.cell(row_base+1, col, val)
        vc.font      = Font(name=FONT_NAME, bold=True, color=bg, size=13)
        vc.fill      = PatternFill("solid", start_color="F2F2F2")
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = tborder()

    # 주의 시계열 목록
    warn_start = 12
    ws.cell(warn_start, 1, "▶ 주의 필요 시계열 (MAPE > 30%)").font = Font(
        name=FONT_NAME, bold=True, size=11, color=C_TITLE)
    for c, h in enumerate(["시계열명", "MAPE(%)", "최적모델", "비고"], 1):
        hcell(ws.cell(warn_start+1, c), h)
    for r, w in enumerate(summary.get("warn_list", []), warn_start+2):
        dcell(ws.cell(r, 1), w["name"],  align="left",   bg=C_WARN)
        dcell(ws.cell(r, 2), f"{w['mape']:.1f}%", align="center", bg=C_WARN)
        dcell(ws.cell(r, 3), w["model"], align="center", bg=C_WARN)
        dcell(ws.cell(r, 4), "예측 신뢰도 낮음", align="left", bg=C_WARN)

    for c in range(1, 14):
        cw(ws, c, 16)
    ws.sheet_view.showGridLines = False


# ════════════════════════════════════════════════════════════════════════════
# 8. 메인 파이프라인
# ════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 65)
    print("   CPSM 수요예측 자동화 시스템")
    print("=" * 65)

    # ── 저장된 설정 확인 ──────────────────────────────────────────────────
    saved_cfg = load_config()
    use_saved = False
    filepaths = []
    lead_time = DEFAULT_LEAD_TIME
    lt_df     = pd.DataFrame()
    lt_file_used = None
    col_map   = {}

    if saved_cfg:
        # 저장된 파일 경로가 실제 존재하는지 확인
        saved_fps = [Path(p) for p in saved_cfg.get("filepaths", []) if Path(p).exists()]
        if saved_fps and saved_cfg.get("col_map"):
            use_saved = prompt_use_saved_config(saved_cfg)
        else:
            print("\n[안내] 저장된 설정의 파일이 없어 새로 설정합니다.")

    if use_saved:
        # ── 저장된 설정 복원 ──────────────────────────────────────────────
        filepaths    = [Path(p) for p in saved_cfg["filepaths"] if Path(p).exists()]
        lead_time    = saved_cfg.get("lead_time", DEFAULT_LEAD_TIME)
        col_map      = saved_cfg.get("col_map", {})
        lt_file_path = saved_cfg.get("lt_file")
        if lt_file_path and Path(lt_file_path).exists():
            lt_df        = load_lead_time_file(Path(lt_file_path))
            lt_file_used = Path(lt_file_path)
        else:
            lt_df        = pd.DataFrame()
            lt_file_used = None
        print(f"\n[분석 대상] {len(filepaths)}개 파일:")
        for f in filepaths:
            print(f"  - {f.name}")
        if lt_df.empty:
            print(f"  리드타임: 기본값 {lead_time}일")
        else:
            print(f"  리드타임: {len(lt_df)}개 품목 파일 적용")
    else:
        # ── 파일 탐지 ─────────────────────────────────────────────────────
        auto_files = find_input_files()

        if auto_files:
            print(f"\n[파일 탐지] {len(auto_files)}개 발견:")
            for i, f in enumerate(auto_files, 1):
                print(f"  {i}. {f.name}")
            print()
            print("  선택 방법:")
            print("  - 전체 사용  : Enter (또는 Y)")
            print("  - 일부 선택  : 번호를 쉼표로 입력  예) 1,3,5")
            print("  - 직접 입력  : 파일 경로를 입력")
            sel = input("\n  선택 > ").strip()

            if sel == "" or sel.lower() == "y":
                filepaths = auto_files
            elif sel.replace(",","").replace(" ","").isdigit() and "," in sel:
                idxs = [int(x.strip())-1 for x in sel.split(",") if x.strip().isdigit()]
                filepaths = [auto_files[i] for i in idxs if 0 <= i < len(auto_files)]
            elif sel.isdigit():
                idx = int(sel) - 1
                filepaths = [auto_files[idx]] if 0 <= idx < len(auto_files) else auto_files
            else:
                # 직접 경로 입력 (여러 개는 세미콜론으로 구분)
                filepaths = [Path(p.strip().strip('"')) for p in sel.split(";") if p.strip()]
        else:
            print("\n[안내] 현재 폴더에서 데이터 파일을 찾지 못했습니다.")
            print("  파일 경로를 입력하세요. (여러 파일은 세미콜론으로 구분)")
            print("  예) C:\\data\\2023.xlsx;C:\\data\\2024.xlsx")
            raw = input("  경로 > ").strip()
            filepaths = [Path(p.strip().strip('"')) for p in raw.split(";") if p.strip()]

        # 존재 확인
        filepaths = [f for f in filepaths if f.exists()]
        if not filepaths:
            print("[오류] 유효한 파일이 없습니다.")
            sys.exit(1)

        print(f"\n[분석 대상] {len(filepaths)}개 파일:")
        for f in filepaths:
            print(f"  - {f.name}")

        # ── 리드타임 설정 ──────────────────────────────────────────────────
        lt_file_candidates = sorted(
            [f for f in BASE_DIR.glob("리드타임*.xlsx")] +
            [f for f in BASE_DIR.glob("lead_time*.xlsx")] +
            [f for f in BASE_DIR.glob("리드타임*.csv")],
            key=lambda p: p.stat().st_mtime, reverse=True
        )

        print("\n[리드타임 설정]")
        print("  품목별 리드타임 파일을 사용하면 안전재고/ROP가 더 정확하게 산출됩니다.")

        lt_file_used = None
        if lt_file_candidates:
            print(f"  → 탐지된 리드타임 파일: {lt_file_candidates[0].name}")
            lt_ans = input("  이 파일을 사용하시겠습니까? (Y/n/파일경로): ").strip()
            if lt_ans.lower() == "n":
                lt_df = pd.DataFrame()
                lead_time = DEFAULT_LEAD_TIME
                print(f"  → 전체 기본값 {DEFAULT_LEAD_TIME}일 적용")
            elif lt_ans == "" or lt_ans.lower() == "y":
                lt_df = load_lead_time_file(lt_file_candidates[0])
                lt_file_used = lt_file_candidates[0]
                lead_time = DEFAULT_LEAD_TIME
            else:
                lt_path = Path(lt_ans.strip('"'))
                lt_df = load_lead_time_file(lt_path)
                lt_file_used = lt_path
                lead_time = DEFAULT_LEAD_TIME
        else:
            print("  → 리드타임 파일을 찾지 못했습니다.")
            print("     파일 경로 입력 / 숫자 입력(전체 단일값) / Enter(기본 30일)")
            lt_ans = input("  선택 > ").strip()
            if lt_ans == "":
                lt_df = pd.DataFrame()
                lead_time = DEFAULT_LEAD_TIME
                print(f"  → 기본값 {DEFAULT_LEAD_TIME}일 적용")
            elif lt_ans.isdigit():
                lt_df = pd.DataFrame()
                lead_time = int(lt_ans)
                print(f"  → 전체 단일값 {lead_time}일 적용")
            else:
                lt_path = Path(lt_ans.strip('"'))
                lt_df = load_lead_time_file(lt_path)
                lt_file_used = lt_path
                lead_time = DEFAULT_LEAD_TIME

        if lt_df.empty:
            print(f"  → 품목별 리드타임 미적용 (기본값 {lead_time}일)")
        else:
            print(f"  → 품목별 리드타임 {len(lt_df)}개 품목 적용 (미등록 품목: {lead_time}일)")

        # ── 컬럼 탐지 & 사용자 확인 (첫 번째 파일 기준) ────────────────────
        ref_file = filepaths[0]
        if ref_file.suffix.lower() == ".csv":
            for enc in ("utf-8-sig", "cp949", "utf-8"):
                try:
                    sample_df = pd.read_csv(ref_file, nrows=3, encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
        else:
            sheets = _read_excel_safe(ref_file, sheet_name=None, nrows=3)
            sample_df = list(sheets.values())[0]
        sample_df.columns = sample_df.columns.str.strip()

        print(f"\n[컬럼 목록] {ref_file.name} 기준 ({len(sample_df.columns)}개)")
        for i, c in enumerate(sample_df.columns, 1):
            print(f"  {i:3}. {c}")

        col_map = map_columns(sample_df)
        print("\n[자동 매핑 결과]")
        for k, v in col_map.items():
            print(f"  {k:12} → {v}")
        for req in ["date", "amount"]:
            if req not in col_map:
                print(f"  [경고] '{req}' 컬럼 미발견!")

        confirm = input("\n위 매핑으로 분석을 시작하시겠습니까? (Y/n): ").strip().lower()
        if confirm == "n":
            print("  컬럼명을 수동으로 입력하세요.")
            for key in ["date", "amount", "qty", "vendor", "prod_id"]:
                val = input(f"  {key} 컬럼명 (현재: {col_map.get(key,'없음')}): ").strip()
                if val:
                    col_map[key] = val

        if "date" not in col_map or "amount" not in col_map:
            print("[오류] 날짜/금액 컬럼이 필수입니다.")
            sys.exit(1)

        # ── 설정 저장 (다음 실행 시 재사용) ──────────────────────────────
        save_config(filepaths, lead_time, lt_file_used, col_map)
        print(f"\n[설정 저장] 다음 실행 시 자동으로 로드됩니다. ({CONFIG_FILE.name})")

    # ── 데이터 로드 (복수 파일 병합) ─────────────────────────────────────
    df = load_and_preprocess(filepaths, col_map, lead_time)

    # ── 전체 월별 집계 ────────────────────────────────────────────────────
    overall_grp = build_monthly_series(df)
    overall_grp = trim_to_36months(overall_grp, "_ym")
    overall_grp.sort_values("_ym", inplace=True)

    min_ym = overall_grp["_ym"].min()
    max_ym = overall_grp["_ym"].max()
    full_idx = pd.period_range(min_ym, max_ym, freq="M")
    overall_grp = overall_grp.set_index("_ym").reindex(full_idx, fill_value=0).reset_index()
    overall_grp.rename(columns={"index": "_ym"}, inplace=True)

    log.info(f"전체 분석기간: {min_ym} ~ {max_ym} ({len(overall_grp)}개월)")

    # ── 전체 시계열 모델 평가 & 예측 ─────────────────────────────────────
    print("\n[STEP 3-4] 전체 시계열 모델 평가 중...")
    overall_arr = overall_grp["amount"].values.astype(float)

    # 이상치 통계 출력
    arr_w = winsorize_series(overall_arr, pct=0.05)
    if not np.allclose(overall_arr, arr_w):
        clipped = np.sum(overall_arr != arr_w)
        log.info(f"  이상치 윈소라이징: {clipped}개 월 조정 "
                 f"(원본 범위 {overall_arr.min():,.0f}~{overall_arr.max():,.0f} → "
                 f"조정 {arr_w.min():,.0f}~{arr_w.max():,.0f})")

    acc_results, best_overall = walk_forward_eval(overall_arr, "Holt-Winters")
    best_mape = acc_results.get(best_overall, {}).get("mape", np.nan)

    if not np.isnan(best_mape) and best_mape > 50:
        log.warning(f"전체 시계열 MAPE={best_mape:.1f}% → 구조적 변화 감지. "
                    "데이터 기간 단축(최근 24개월) 후 재평가.")
        overall_arr_short = overall_arr[-24:] if len(overall_arr) >= 24 else overall_arr
        acc_short, best_short = walk_forward_eval(overall_arr_short, "Holt-Winters")
        mape_short = acc_short.get(best_short, {}).get("mape", np.nan)
        if not np.isnan(mape_short) and mape_short < best_mape:
            log.info(f"  → 최근 24개월 사용: {best_short} MAPE={mape_short:.1f}% (개선)")
            best_overall = best_short
            best_mape    = mape_short
            acc_results  = acc_short
            overall_arr  = overall_arr_short          # 예측도 단축 데이터로

    log.info(f"전체 최적모델: {best_overall} (MAPE={best_mape:.1f}%)")

    future_periods = pd.period_range(max_ym + 1, periods=6, freq="M")
    fc_vals, ci_lo, ci_hi = forecast_series(overall_arr, best_overall, h=6)

    overall_acc_row = {
        "name": "전체 매출",
        "best": best_overall,
        "note": f"MAPE {best_mape:.1f}%" if not np.isnan(best_mape) else "",
    }
    for k, v in acc_results.items():
        overall_acc_row[k] = v["mape"]

    # ── 협력사×상품코드 집계 ──────────────────────────────────────────────
    print("[STEP 1] 협력사×상품코드 월별 집계 중...")
    group_grp = build_monthly_series(df, group_cols=["_vendor", "_prod_key"])
    group_grp = trim_to_36months(group_grp, "_ym")

    # ── ABC × CV ─────────────────────────────────────────────────────────
    print("[STEP 2] ABC × CV 분류 중...")
    abc_df = classify_abc_cv(group_grp)
    log.info(f"ABC 분류: A={len(abc_df[abc_df.ABC=='A'])} B={len(abc_df[abc_df.ABC=='B'])} C={len(abc_df[abc_df.ABC=='C'])}")

    # ── 리드타임 통계 (원본 데이터에서 실측) ────────────────────────────
    lt_stats = build_lt_stats(df)

    has_lt_data = not lt_stats.empty
    if has_lt_data:
        print(f"  → 실측 리드타임 데이터 활용: {len(lt_stats)}개 품목 "
              f"(전체 평균 {lt_stats['lt_mean'].mean():.1f}일, "
              f"σ={lt_stats['lt_std'].mean():.1f}일)")
        print("  → 안전재고 공식: SS = Z × √(LT × σ_D² + D² × σ_LT²)  [완전공식]")
    else:
        print(f"  → 리드타임 실측 데이터 없음 → 파일/기본값 사용")
        print("  → 안전재고 공식: SS = Z × σ_D × √LT  [기본공식]")

    # ── 리드타임 템플릿 생성 (실측·파일 모두 없을 때) ────────────────────
    tpl_path = BASE_DIR / "리드타임_입력템플릿.xlsx"
    if not has_lt_data and lt_df.empty and not tpl_path.exists():
        make_lead_time_template(abc_df, tpl_path)
        print(f"\n[안내] 리드타임 입력 템플릿 생성: {tpl_path.name}")
        print("       노란색 셀에 품목별 리드타임을 입력 후 저장하면")
        print("       다음 실행 시 자동 적용됩니다.\n")

    # ── 안전재고 / ROP ────────────────────────────────────────────────────
    ss_rop = {}
    for _, row in abc_df.iterrows():
        key = (row["_vendor"], row["_prod_key"])
        lt_m, lt_s = get_lt_stats(lt_stats, lt_df, row["_vendor"], row["_prod_key"])

        sub = group_grp[(group_grp["_vendor"]==row["_vendor"]) &
                        (group_grp["_prod_key"]==row["_prod_key"])].sort_values("_ym")
        qty_arr = sub["qty"].values.astype(float)
        if len(qty_arr) >= 2:
            ss, formula = calc_safety_stock(qty_arr, lt_m, lt_s)
            rop = calc_rop(qty_arr, lt_m, ss)
        else:
            ss, rop, formula = 0.0, 0.0, "데이터부족"
        ss_rop[key] = {"ss": ss, "rop": rop, "lead_time": lt_m,
                       "lt_std": lt_s, "formula": formula}

    # ── 개별 시계열 예측 ──────────────────────────────────────────────────
    print("[STEP 3-4] 개별 시계열 예측 중...")
    a_items  = abc_df[abc_df["ABC"] == "A"]
    b_items  = abc_df[abc_df["ABC"] == "B"].head(20)
    target_items = pd.concat([a_items, b_items], ignore_index=True)

    n_series = len(target_items)
    large_mode = n_series > 100
    if large_mode:
        log.info(f"시계열 {n_series}개 > 100 → A등급만 전체 비교, 나머지 HW 단일 적용")

    detail_rows = []
    acc_rows    = [overall_acc_row]
    warn_list   = []
    if not np.isnan(best_mape) and best_mape > 30:
        warn_list.append({"name": "전체 매출", "mape": best_mape, "model": best_overall})

    for idx, (_, row) in enumerate(target_items.iterrows(), 1):
        vendor   = row["_vendor"]
        prod_key = row["_prod_key"]
        abc      = row["ABC"]
        strategy = row["strategy"]

        sub = group_grp[(group_grp["_vendor"]==vendor) &
                        (group_grp["_prod_key"]==prod_key)].sort_values("_ym")
        if len(sub) < 3:
            log.debug(f"건너뜀(데이터 부족): {vendor} / {prod_key}")
            continue

        # 결측 월 보간
        sub_idx = pd.period_range(sub["_ym"].min(), sub["_ym"].max(), freq="M")
        sub = sub.set_index("_ym").reindex(sub_idx, fill_value=0).reset_index()
        sub.rename(columns={"index": "_ym"}, inplace=True)

        if check_consecutive_zeros(sub["amount"].values, threshold=3):
            log.info(f"제외(연속0): {vendor} / {prod_key}")
            continue

        arr = sub["amount"].values.astype(float)

        try:
            if large_mode and abc != "A":
                # 단일 모델 (HW 또는 전략 모델)
                _strat = strategy if strategy != "SARIMA" else "Holt-Winters"
                sub_acc = {_strat: {"mape": np.nan, "mae": np.nan, "rmse": np.nan}}
                best_item = _strat
            else:
                sub_acc, best_item = walk_forward_eval(arr, strategy)

            fc6, fc_lo, fc_hi = forecast_series(arr, best_item, h=6)
            item_mape = sub_acc.get(best_item, {}).get("mape", np.nan)

            avg12 = float(np.mean(arr[-12:])) if len(arr) >= 12 else float(np.mean(arr))
            detail_rows.append({
                "vendor":   vendor,
                "prod_key": prod_key,
                "ABC":      abc,
                "avg12":    avg12,
                "fc6":      fc6,
                "model":    best_item,
                "mape":     item_mape,
            })

            acc_row = {"name": f"{vendor} / {prod_key}", "best": best_item,
                       "note": abc}
            for k, v in sub_acc.items():
                acc_row[k] = v["mape"]
            acc_rows.append(acc_row)

            if not np.isnan(item_mape) and item_mape > 30:
                warn_list.append({"name": f"{vendor}/{prod_key}",
                                  "mape": item_mape, "model": best_item})

            if idx % 10 == 0 or idx == n_series:
                print(f"  진행: {idx}/{n_series}  ({vendor[:15]} / {prod_key[:15]})")

        except Exception as e:
            log.warning(f"시계열 실패 건너뜀: {vendor} / {prod_key} → {e}")
            continue

    # ── 요약 지표 ─────────────────────────────────────────────────────────
    all_mapes = [r.get(r["best"], np.nan) for r in acc_rows
                 if r["best"] in r and not np.isnan(r.get(r["best"], np.nan))]
    avg_mape = float(np.nanmean(all_mapes)) if all_mapes else np.nan

    total_fc6 = float(np.sum(fc_vals))

    summary = {
        "run_dt":     datetime.now().strftime("%Y-%m-%d %H:%M"),
        "data_range": f"{min_ym} ~ {max_ym}",
        "n_series":   n_series,
        "n_A":        int((abc_df["ABC"]=="A").sum()),
        "total_fc6":  total_fc6,
        "avg_mape":   avg_mape if not np.isnan(avg_mape) else 0.0,
        "n_warn":     len(warn_list),
        "warn_list":  warn_list,
    }

    # ── Excel 출력 ────────────────────────────────────────────────────────
    out_name = f"수요예측_{datetime.now().strftime('%Y%m')}.xlsx"
    out_path = BASE_DIR / out_name
    print(f"\n[출력] {out_name} 작성 중...")

    import openpyxl
    wb = openpyxl.Workbook()
    del wb["Sheet"]

    write_sheet5_dashboard(wb, summary)
    write_sheet1_overall(wb, overall_grp, fc_vals, ci_lo, ci_hi,
                          best_overall, best_mape, future_periods)
    write_sheet2_abc(wb, abc_df, ss_rop)
    write_sheet3_detail(wb, detail_rows)
    write_sheet4_accuracy(wb, acc_rows)

    wb.save(out_path)
    log.info(f"저장 완료: {out_path}")

    # ── 최종 요약 출력 ────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print(f"  분석 기간          : {summary['data_range']}")
    print(f"  전체 최적 모델     : {best_overall}  (MAPE {best_mape:.1f}%)")
    print(f"  향후 6개월 예측총액: {total_fc6:,.0f} 원")
    print(f"  A등급 품목 수      : {summary['n_A']}개")
    print(f"  주의 시계열 수     : {len(warn_list)}개 (MAPE > 30%)")
    print(f"  출력 파일          : {out_path.name}")
    print("=" * 65)

    if sys.platform == "win32":
        try:
            os.startfile(out_path)
        except OSError:
            pass


if __name__ == "__main__":
    main()
