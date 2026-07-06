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
    "pyarrow":      "pyarrow",      # ← parquet 저장/로드용
    "prophet":      "prophet",      # ← Meta Prophet 시계열 모델
    "lightgbm":     "lightgbm",     # ← LightGBM 래그 피처 모델
    "requests":     "requests",     # ← 한국은행 ECOS API 호출
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
CONFIG_FILE  = BASE_DIR / "forecast_config.json"
PARQUET_FILE = BASE_DIR / "processed_data.parquet"


# ════════════════════════════════════════════════════════════════════════════
# 0-A. 설정 저장/로드
# ════════════════════════════════════════════════════════════════════════════
import json

def save_config(filepaths: list, lead_time: int, lt_file_path, col_map: dict,
                processed_files: list = None):
    """분석 설정을 JSON 파일에 저장 (다음 실행 시 재사용)"""
    cfg = {
        "filepaths":        [str(p) for p in filepaths],
        "lead_time":        lead_time,
        "lt_file":          str(lt_file_path) if lt_file_path else None,
        "col_map":          col_map,
        "saved_at":         datetime.now().strftime("%Y-%m-%d %H:%M"),
        "processed_files":  processed_files or [],
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
# 0-B. 누적 데이터 저장/로드 (RAW 파일 삭제 후에도 분석 가능)
# ════════════════════════════════════════════════════════════════════════════
def save_processed_data(df: pd.DataFrame, processed_files: list):
    """전처리된 DataFrame을 parquet으로 저장. 이후 RAW 파일 없이도 분석 가능."""
    if df is None or len(df) == 0:
        log.warning("저장할 데이터가 없음 (0행) → parquet 저장 건너뜀")
        return
    try:
        # Period/혼합타입 컬럼 → 문자열 변환 후 저장
        df_save = df.copy()
        if "_ym" in df_save.columns:
            df_save["_ym"] = df_save["_ym"].astype(str)
        if "_date" in df_save.columns:
            df_save["_date"] = df_save["_date"].astype(str)
        # object 타입 컬럼(혼합 int/str 등) → str 통일
        for col in df_save.columns:
            if df_save[col].dtype == object:
                df_save[col] = df_save[col].astype(str)
        df_save.to_parquet(PARQUET_FILE, index=False, engine="pyarrow")
        size_mb = PARQUET_FILE.stat().st_size / 1024 / 1024
        log.info(f"누적 데이터 저장: {PARQUET_FILE.name} "
                 f"({len(df_save):,}행, {size_mb:.1f}MB, "
                 f"파일 {len(processed_files)}개 누적)")
    except Exception as e:
        log.warning(f"누적 데이터 저장 실패: {e}")


def load_processed_data() -> pd.DataFrame | None:
    """저장된 parquet 로드. 없거나 실패 시 None 반환."""
    if not PARQUET_FILE.exists():
        return None
    try:
        df = pd.read_parquet(PARQUET_FILE, engine="pyarrow")
        # 빈 parquet(이전 오류 실행 잔재) → None 반환해 재처리 유도
        if len(df) == 0:
            log.warning("누적 데이터 0행 감지 (이전 오류 잔재) → RAW 파일에서 재처리합니다.")
            PARQUET_FILE.unlink(missing_ok=True)
            return None
        # 문자열 → 원래 타입 복원
        if "_ym" in df.columns:
            df["_ym"] = df["_ym"].apply(lambda x: pd.Period(x, freq="M")
                                         if pd.notna(x) and x else pd.NaT)
        if "_date" in df.columns:
            df["_date"] = pd.to_datetime(df["_date"], errors="coerce")
        # _channel 컬럼 없으면 _source_file에서 복원
        if "_channel" not in df.columns and "_source_file" in df.columns:
            df["_channel"] = df["_source_file"].apply(_extract_channel)
            log.info("  _channel 컬럼 복원 완료 (파일명 기반)")
        # _mgmt_acct 없으면 원본 컬럼(관리회계)에서 복원
        if "_mgmt_acct" not in df.columns:
            _raw_mgmt = next((c for c in df.columns if c in ("관리회계", "management_account", "mgmt_acct")), None)
            if _raw_mgmt:
                df["_mgmt_acct"] = df[_raw_mgmt].astype(str).str.strip().replace({"nan": "—", "None": "—", "": "—"})
                log.info(f"  _mgmt_acct 컬럼 복원 완료 ({_raw_mgmt} 기반, {df['_mgmt_acct'].nunique()}개 고유값)")
            else:
                log.warning("  _mgmt_acct 복원 불가 — parquet 삭제 후 재실행하면 원본에서 로드됩니다")
        size_mb = PARQUET_FILE.stat().st_size / 1024 / 1024
        log.info(f"누적 데이터 로드: {PARQUET_FILE.name} ({len(df):,}행, {size_mb:.1f}MB)")
        return df
    except Exception as e:
        log.warning(f"누적 데이터 로드 실패: {e} → RAW 파일에서 재처리합니다.")
        return None


def _apply_settle_priority(df: pd.DataFrame) -> pd.DataFrame:
    """
    정산현황 파일 우선순위 적용.
    정산현황(확정 정산) 데이터가 주문현황과 중복되지 않도록 주문현황의 _settle_ym을 무효화.
    - 발주일 기준(_ym)은 주문현황 그대로 유지 → 수량·발주 예측에 영향 없음
    - 정산일 기준(_settle_ym)만 정산현황이 담당 → 채널별_분석·매출예측 정확도 확보

    중복 제거 방식 (자동 선택):
      ① 주문번호(_order_id) 있음 → 주문 단위 정확 매칭 (권장)
      ② 주문번호 없음           → 채널×정산월 단위 블록 대체 (폴백)
    """
    if "_settle_ym" not in df.columns or "_source_file" not in df.columns:
        return df

    _settle_mask = df["_source_file"].str.contains("정산현황", na=False)
    if not _settle_mask.any():
        return df

    df = df.copy()
    _order_mask = ~_settle_mask

    # ── ① 주문 라인 키 단위 정확 중복 제거 (주문번호+품목번호) ─────────────
    _key_col = "_line_key" if "_line_key" in df.columns else (
               "_order_id" if "_order_id" in df.columns else None)
    if _key_col and df.loc[_settle_mask, _key_col].notna().any():
        _settle_keys = set(
            df.loc[_settle_mask, _key_col].dropna().astype(str).unique()
        )
        _nullify = _order_mask & df[_key_col].astype(str).isin(_settle_keys)
        df.loc[_nullify, "_settle_ym"] = pd.NaT
        _basis = "주문번호+품목번호" if _key_col == "_line_key" else "주문번호"
        log.info(f"정산현황 우선 적용 ({_basis} 기준): {len(_settle_keys):,}건 — "
                 f"주문현황 중복 {_nullify.sum():,}행 _settle_ym 무효화")
        return df

    # ── ② 주문번호 없으면 블록 무효화 하지 않음 ─────────────────────────────
    # 채널×정산월 단위 블록 무효화는 정산현황 행수가 주문현황보다 훨씬 적을 때
    # 수만 건의 주문현황 데이터를 통째로 날리는 부작용이 있어 비활성화.
    # 주문번호+품목번호가 포함된 데이터를 받으면 ① 경로로 정확하게 처리됨.
    log.info("정산현황 우선 적용: 주문번호 없음 → 블록 무효화 생략 (중복 허용)")
    return df


def load_incremental(filepaths: list, col_map: dict, lead_time: int,
                     processed_files: list) -> tuple:
    """
    증분 로드: 이미 처리된 파일은 건너뛰고 새 파일만 처리 후 parquet에 병합.
    반환: (최종 DataFrame, 업데이트된 processed_files 목록)
    """
    existing_df = load_processed_data()

    # parquet이 없으면 processed_files 무시하고 전체 재처리
    if existing_df is None:
        if processed_files:
            log.warning("누적 데이터(parquet) 없음 → processed_files 초기화 후 전체 재처리")
        existing_names = set()
        new_files = list(filepaths)
    else:
        existing_names = set(processed_files)
        # 새 파일 = 목록에 없는 것
        new_files = [f for f in filepaths if f.name not in existing_names]
        # 삭제된 RAW 파일 (목록엔 있지만 실제 없는 것) → parquet으로 대체
        missing_files = [n for n in existing_names
                         if not any(f.name == n for f in filepaths)]
        if missing_files:
            log.info(f"RAW 파일 없음 (parquet으로 대체): {len(missing_files)}개 — "
                     + ", ".join(missing_files[:3]) + ("..." if len(missing_files) > 3 else ""))

    if not new_files and existing_df is not None and len(existing_df) > 0:
        log.info("새 RAW 파일 없음 → 누적 데이터 그대로 사용")
        return existing_df, processed_files

    if new_files:
        print(f"\n[증분 로드] 새 파일 {len(new_files)}개 처리 중...")
        for f in new_files:
            print(f"  + {f.name}")
        new_df = load_and_preprocess(new_files, col_map, lead_time)
        new_names = [f.name for f in new_files]
    else:
        new_df = None
        new_names = []

    # 기존 + 신규 병합
    if existing_df is not None and new_df is not None:
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        log.info(f"병합: 기존 {len(existing_df):,}행 + 신규 {len(new_df):,}행 = {len(combined):,}행")
        # ── 내용 기반 중복 제거 ───────────────────────────────────────────────
        # 신규 파일에 _order_id(주문번호)가 있으면 → 주문번호+품목번호 키로 정확 제거
        # 없으면 → 날짜+공급사+상품코드+수량+금액 조합으로 근사 제거
        _before_dedup = len(combined)
        if "_order_id" in combined.columns and combined["_order_id"].notna().any():
            _key_col = "_line_key" if "_line_key" in combined.columns else "_order_id"
            # 정산현황 행 우선 보존 (is_settle=True를 sort 기준으로 사용)
            combined["_is_settle"] = combined["_source_file"].str.contains("정산현황", na=False)
            combined = (combined
                        .sort_values("_is_settle", ascending=False)
                        .drop_duplicates(subset=[_key_col], keep="first")
                        .drop(columns=["_is_settle"])
                        .reset_index(drop=True))
            log.info(f"내용 기반 중복 제거 ({_key_col} 키): {_before_dedup - len(combined):,}행 제거 → {len(combined):,}행")
        else:
            _fuzzy_keys = [c for c in ["_date", "_vendor", "_prod_key", "_qty", "_amount"]
                           if c in combined.columns]
            if _fuzzy_keys:
                combined["_is_settle"] = combined["_source_file"].str.contains("정산현황", na=False)
                combined = (combined
                            .sort_values("_is_settle", ascending=False)
                            .drop_duplicates(subset=_fuzzy_keys, keep="first")
                            .drop(columns=["_is_settle"])
                            .reset_index(drop=True))
                log.info(f"내용 기반 중복 제거 (근사 키): {_before_dedup - len(combined):,}행 제거 → {len(combined):,}행")
    elif existing_df is not None:
        combined = existing_df
    else:
        combined = new_df

    # 정산현황 파일 우선순위 적용 (concat 이후 전체 데이터 기준으로 중복 제거)
    combined = _apply_settle_priority(combined)

    updated_files = sorted(set(list(existing_names) + new_names))

    # parquet 저장 (pyarrow 없으면 건너뜀)
    try:
        import pyarrow  # noqa
        save_processed_data(combined, updated_files)
    except ImportError:
        log.warning("pyarrow 미설치 → parquet 저장 생략 (pip install pyarrow 로 설치 가능)")

    return combined, updated_files


def _build_rightmost_map(prod_type_map: dict) -> dict:
    """
    기존 매핑에서 '>' 기준 마지막 키워드 → 통신/일반 딕셔너리 생성.
    동일 키워드가 통신/일반 양쪽에 있으면 충돌로 제외.
    예) 'CS전용 > GiGAeyes > 허브/공유기' → {'허브/공유기': '통신'}
    """
    rightmost_map: dict = {}
    conflicts: set = set()
    for cat, ptype in prod_type_map.items():
        rightmost = cat.split(">")[-1].strip()
        if not rightmost:
            continue
        if rightmost in conflicts:
            continue
        if rightmost in rightmost_map:
            if rightmost_map[rightmost] != ptype:
                conflicts.add(rightmost)
                del rightmost_map[rightmost]
        else:
            rightmost_map[rightmost] = ptype
    log.debug(f"  rightmost 매핑: {len(rightmost_map)}개 (충돌 제외: {len(conflicts)}개)")
    return rightmost_map


def ensure_prod_type(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    parquet 재사용 시 _prod_type 컬럼이 없을 수 있으므로 사후 보완.
    이미 있으면 아무것도 하지 않음.

    매칭 순서:
      1단계) 서비스카테고리 전체 문자열 → 정확 매핑
      2단계) 미분류 잔여 → '>' 마지막 키워드로 추론 매핑
    """
    if "_prod_type" in df.columns:
        return df

    prod_type_map = load_prod_type_map()
    cat_col_name  = col_map.get("category")

    if not prod_type_map or not cat_col_name or cat_col_name not in df.columns:
        df = df.copy()
        df["_prod_type"] = "미분류"
        if not prod_type_map:
            log.info("  통신/일반 구분: 매핑 파일 없음 → 전체 미분류")
        elif not cat_col_name:
            log.info("  통신/일반 구분: 서비스카테고리 컬럼 없음 → 전체 미분류")
        return df

    df = df.copy()
    cat_series = df[cat_col_name].astype(str).str.strip()

    # ── 1단계: 정확 매핑 ────────────────────────────────────────────────────
    df["_prod_type"] = cat_series.map(prod_type_map).fillna("미분류")
    n_miss = (df["_prod_type"] == "미분류").sum()

    # ── 2단계: 미분류 → 마지막 키워드로 추론 ────────────────────────────────
    if n_miss > 0:
        rightmost_map = _build_rightmost_map(prod_type_map)
        miss_mask = df["_prod_type"] == "미분류"
        rightmost_series = cat_series[miss_mask].str.split(">").str[-1].str.strip()
        inferred = rightmost_series.map(rightmost_map)
        df.loc[miss_mask, "_prod_type"] = inferred.fillna("미분류")
        n_inferred = inferred.notna().sum()
        log.info(f"  2단계 추론(마지막 키워드): {n_inferred:,}행 추가 분류")

    n_t = (df["_prod_type"] == "통신").sum()
    n_g = (df["_prod_type"] == "일반").sum()
    n_u = (df["_prod_type"] == "미분류").sum()
    log.info(f"  통신/일반 구분 완료: 통신 {n_t:,}행 / 일반 {n_g:,}행 / 미분류 {n_u:,}행")
    return df


# ════════════════════════════════════════════════════════════════════════════
# 0. 상수 및 스타일
# ════════════════════════════════════════════════════════════════════════════
Z_SS  = 2.05  # 서비스수준 98% (MRO 가용성 기준 — KT SCM 이관 품목 기본값)
Z_CI  = 1.96  # 95% 신뢰구간
Z_CUST = 1.65  # 고객사 자체 관리 시 서비스수준 95% 가정 (QBR 절감액 산출 기준)

# ── QBR 안전재고 절감액 산출 설정 ──────────────────────────────────────────
# 연간 재고 보유비용률: 자본비용 + 창고비 + 진부화 손실 합산 업종 표준 20~30%
HOLDING_COST_RATE = 0.25  # 25% — 고객사 QBR 절감액 추정 시 사용

# ── MRO 이관 설정 ──────────────────────────────────────────────────────────
# KT SCM 이관 시점. None 이면 전체 기간 학습. 설정 시 이관 후 데이터만 사용.
# 예) TRANSFER_DATE = "2023-04-01"
TRANSFER_DATE: str | None = None

# 콜드스타트 비활성화 — 이관 품목은 이미 이력이 존재하므로 불필요
ENABLE_COLD_START = False

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
                    "리드타임_입력템플릿", "mapping_master",
                    "카테고리목록", "미분류",       # 카테고리 참조 파일
                    "보완 자료",                    # 배송완료일 보완용 원시 파일 (이미 병합됨)
                    # "정산현황" 은 제외하지 않음 — 발주일자 없을 시 정산일자 fallback 처리
                    ]

# 카테고리-담당자-통신구분 파일 자동 탐지
# 파일명에 아래 키워드가 모두 포함된 xlsx를 우선 사용, 없으면 mapping_master.xlsx fallback
_CATEGORY_FILE_KEYWORDS = ["서비스", "카테고리", "담당자"]

def _find_category_mapping_file() -> Path | None:
    """BASE_DIR에서 카테고리-담당자 매핑 파일을 자동으로 찾아 반환.
    ~$ 로 시작하는 Excel 임시 잠금 파일은 제외."""
    candidates = list(BASE_DIR.glob("*.xlsx")) + list(BASE_DIR.glob("*.xlsm"))
    for p in candidates:
        if p.name.startswith("~$"):          # Excel 잠금 파일 제외
            continue
        if all(kw in p.name for kw in _CATEGORY_FILE_KEYWORDS):
            return p
    # fallback: mapping_master.xlsx
    fallback = BASE_DIR / "mapping_master.xlsx"
    return fallback if fallback.exists() else None

def _dept_to_prod_type(dept: str) -> str:
    """
    담당부서(안) 값 → 통신/일반 변환.
    네트워크 관련 부서 → '통신', 그 외 → '일반'
    부서명 추가 필요 시 TELECOM_DEPT_KEYWORDS 목록에 추가하면 됨.
    """
    TELECOM_DEPT_KEYWORDS = ["네트워크", "network", "통신", "회선", "NW"]
    dept_str = str(dept).strip()
    if not dept_str or dept_str in ("nan", "None", ""):
        return "미분류"
    if any(kw in dept_str for kw in TELECOM_DEPT_KEYWORDS):
        return "통신"
    return "일반"


def _read_cat_sheet(mapping_file: Path, sheet_name: str) -> dict:
    """
    단일 시트에서 { 카테고리명: "통신"|"일반" } 반환.
    지원 형식:
      - 신형식: 카테고리명 + 담당부서(안)  [카테고리 담당자(251229) 이후]
      - 구형식: 카테고리 분류 + 담당부서(안)  [카테고리 담당자 (개편전)]
    """
    df = pd.read_excel(mapping_file, sheet_name=sheet_name,
                       engine="openpyxl", header=0)
    df.columns = df.columns.str.strip()

    # 담당부서 컬럼 탐지
    dept_col = next((c for c in df.columns if "담당부서" in c), None)
    if dept_col is None:
        dept_col = next((c for c in df.columns if "부서" in c), None)
    if dept_col is None:
        log.debug(f"  [{sheet_name}] 담당부서 컬럼 미발견 → 건너뜀 "
                  f"(보유 컬럼: {list(df.columns)})")
        return {}

    # 카테고리 키 컬럼 우선순위:
    #   1) '서비스 카테고리'  → 실제 데이터의 서비스카테고리 컬럼 값과 일치 (신형식)
    #   2) '카테고리 분류'    → 구형식
    #   3) 그 외 카테고리 관련 컬럼
    # 우선순위대로 모두 매핑하되, 앞 컬럼이 없는 카테고리만 뒤 컬럼으로 보완
    KEY_COLS_PRIORITY = ["서비스 카테고리", "카테고리 분류", "카테고리명"]
    key_cols = [c for c in KEY_COLS_PRIORITY if c in df.columns]
    if not key_cols:
        key_cols = [c for c in df.columns if "카테고리" in c and "UID" not in c]
    if not key_cols:
        log.debug(f"  [{sheet_name}] 카테고리 컬럼 미발견 → 건너뜀")
        return {}

    result = {}
    for key_col in reversed(key_cols):   # 낮은 우선순위부터 채운 뒤 높은 우선순위가 덮어씀
        for _, row in df.iterrows():
            cat  = str(row[key_col]).strip()
            dept = str(row[dept_col]).strip()
            if cat and cat not in ("nan", "None"):
                result[cat] = _dept_to_prod_type(dept)

    log.debug(f"  [{sheet_name}] {len(result)}개 카테고리 로드 "
              f"(key_cols={key_cols}, dept_col='{dept_col}')")
    return result


def load_prod_type_map() -> dict:
    """
    카테고리-담당자 매핑 파일에서 카테고리 → 통신/일반 구분 딕셔너리 반환.

    파일 구조:
      - 신형식 시트 (카테고리 담당자(251229) 이후):
          카테고리 UID | 카테고리명 | 서비스 카테고리 | 관리회계 카테고리(안) |
          정산용 중분류 | 담당부서(안) | 상품담당자(26.05)
      - 구형식 시트 (개편전):
          카테고리 분류 | 관리회계 | 담당부서(안) | 변경담당자(26년 5월)

    담당부서(안)에서 통신/일반 자동 판별:
      - "네트워크", "통신", "회선", "NW" 포함 → 통신
      - 그 외 → 일반

    신형식 시트 우선, 누락 카테고리는 구형식으로 보완.
    반환: { 카테고리명(str): "통신" | "일반" }
    """
    mapping_file = _find_category_mapping_file()
    if mapping_file is None:
        log.warning("카테고리 매핑 파일을 찾을 수 없음 → 통신/일반 구분 불가 "
                    f"(파일명에 {_CATEGORY_FILE_KEYWORDS} 포함 필요)")
        return {}
    log.info(f"카테고리 매핑 파일: {mapping_file.name}")

    try:
        import openpyxl as _oxl
        _wb = _oxl.load_workbook(mapping_file, read_only=True, data_only=True)
        sheet_names = _wb.sheetnames
        _wb.close()
        log.info(f"  시트 목록: {sheet_names}")

        # ── 신형식 시트: "개편전"이 아닌 카테고리 담당자 시트 (여러 개일 수 있음) ──
        new_sheets = [s for s in sheet_names
                      if "카테고리" in s and "담당자" in s and "개편전" not in s]
        # ── 구형식 시트: "개편전" 포함 ──────────────────────────────────────
        old_sheets = [s for s in sheet_names if "개편전" in s]

        if not new_sheets and not old_sheets:
            # fallback: 카테고리가 포함된 모든 시트 시도
            new_sheets = [s for s in sheet_names if "카테고리" in s]

        cat_to_type: dict = {}

        # 신형식 우선 로드
        for sh in new_sheets:
            partial = _read_cat_sheet(mapping_file, sh)
            cat_to_type.update(partial)
            log.info(f"  신형식 시트 [{sh}]: {len(partial)}개 로드")

        # 구형식은 신형식에 없는 카테고리만 보완
        for sh in old_sheets:
            partial = _read_cat_sheet(mapping_file, sh)
            added = {k: v for k, v in partial.items() if k not in cat_to_type}
            cat_to_type.update(added)
            log.info(f"  구형식 시트 [{sh}]: {len(partial)}개 중 {len(added)}개 보완")

        if not cat_to_type:
            log.warning(f"{mapping_file.name}: 카테고리 데이터를 읽지 못함 "
                        f"(시트: {sheet_names})")
            return {}

        n_t = sum(1 for v in cat_to_type.values() if v == "통신")
        n_g = sum(1 for v in cat_to_type.values() if v == "일반")
        n_u = sum(1 for v in cat_to_type.values() if v == "미분류")
        log.info(f"통신/일반 매핑 완료: 총 {len(cat_to_type)}개 카테고리 "
                 f"(통신 {n_t}개 / 일반 {n_g}개 / 미분류 {n_u}개)")
        return cat_to_type

    except Exception as e:
        log.warning(f"{mapping_file.name} 로드 실패 → 통신/일반 구분 건너뜀: {e}")
        return {}

def _is_data_file(p: Path) -> bool:
    # ~$ 접두어 = Excel 임시 잠금 파일
    if p.name.startswith("~$"):
        return False
    name_lower = p.name.lower()
    # EXCLUDE_KEYWORDS 포함 파일 제외
    if any(kw in name_lower for kw in EXCLUDE_KEYWORDS):
        return False
    # 카테고리-담당자 매핑 파일 제외 (키워드 전부 포함 시)
    if all(kw in p.name for kw in _CATEGORY_FILE_KEYWORDS):
        return False
    return p.name != "demand_forecast.py"

def find_input_files() -> list[Path]:
    """현재 폴더 + _배송완료일자추가 하위 폴더에서 분석 대상 데이터 파일 전체 탐지"""
    seen = set()
    hits = []

    # 스캔 대상 디렉토리: BASE_DIR 만 (하위 폴더 자동 포함 없음)
    scan_dirs = [BASE_DIR]

    for scan_dir in scan_dirs:
        for pat in SEARCH_PATTERNS:
            for p in scan_dir.glob(pat):
                if p not in seen and _is_data_file(p):
                    seen.add(p)
                    hits.append(p)

    # 파일명 오름차순(연도 순서 유지)
    hits.sort(key=lambda p: p.name)

    # _배송완료일자추가 버전이 있으면 원본 파일 제거 (중복 적재 방지)
    # 같은 파일명(폴더 불문)이면 enhanced 우선
    enhanced = {p for p in hits if "_배송완료일자추가" in p.name}
    if enhanced:
        orig_names = {p.name.replace("_배송완료일자추가", "") for p in enhanced}
        hits = [p for p in hits if p.name not in orig_names]
        log.info(f"배송완료일자추가 버전 {len(enhanced)}개 감지 → 원본 {len(orig_names)}개 자동 제외")

    return hits

COLUMN_ALIASES = {
    # ── 필수 ──────────────────────────────────────────────────────────────
    "date":      ["발주일자", "주문일자", "주문일시", "주문일",   # 발주일자 최우선
                  "정산일자", "입고일자", "입고일",
                  "정산확정일", "일일정산월", "일일정산일", "날짜", "일자", "date"],
    "amount":    ["정산금액", "판매금액", "주문금액", "매출액", "금액", "revenue", "amount"],
    "qty":       ["정산수량", "입고수량", "수량", "주문수량", "qty", "quantity"],
    "vendor":    ["협력사명", "협력사", "공급업체", "vendor", "supplier"],
    "prod_id":   ["상품코드", "상품ID", "품목코드", "product_id", "prod_id"],
    # ── 권장 ──────────────────────────────────────────────────────────────
    "prod_name": ["상품명", "품목명", "product_name"],
    "category":  ["서비스카테고리", "마스터카테고리(CMS)", "카테고리", "category"],
    "cost":      ["매입금액", "매입금액(원)", "purchase_amount"],       # 마진 분석
    "lt_order":  ["주문시배송리드타임", "주문리드타임", "order_lead_time"],  # 리드타임 실측
    "lt_std":    ["상품배송리드타임", "표준납기일", "standard_lead_time"],  # 표준 리드타임
    "settle_date":["일일정산월", "정산확정일", "정산일자", "settlement_date"],  # 매출 기준일 (정산확정일 우선)
    "cancel_qty":["취소수량", "cancel_qty"],                            # 실수요 보정
    "return_qty":["반품수량", "return_qty"],                            # 실수요 보정
    "order_date":["주문일자", "주문일시", "order_date"],                 # LT 실측 계산용
    "recv_date": ["입고일자", "입고일", "배송완료일자", "배송완료일", "recv_date"],  # LT 실측 계산용
    "dept":      ["부서명", "부서", "department"],                       # 부서별 분석
    "bizplace":  ["사업장", "site", "business_place"],                   # 사업장별 분석
    "order_id":  ["주문번호", "주문ID", "order_id", "order_no"],          # 주문 고유번호
    "line_id":   ["품목번호", "품목ID", "line_id", "item_no", "품번"],    # 주문 라인 번호 (주문번호+품목번호 = 라인 키)
    "order_status":["주문상태", "order_status", "status"],               # 주문 상태 (취소/반품/완료 등)
    "order_type":["주문형태", "IP/DIP", "order_type"],                   # 유형별 분석
    "sale_type": ["매출구분", "sale_type"],                              # 매출 유형
    "regular":   ["정기주문여부", "is_regular", "regular_order"],        # 정기/비정기
    "tax":       ["과세상태", "부가세포함여부", "tax_status"],            # 세금 구분
    "spec":      ["규격", "대표규격", "상품규격", "사양", "spec", "specification"],  # 상품 규격
    "recv_approve_date": ["입고승인일자", "recv_approve_date"],          # KT 귀속 기준일
    "mgmt_acct": ["관리회계", "mgmt_acct", "management_account"],       # 유가증권/배송료/프로모션 분류
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


CHANNELS = ["KT", "그룹사", "외부사", "지입자재"]


def _parse_date_col(series: pd.Series) -> pd.Series:
    """날짜 컬럼 스마트 파싱.
    YYYYMMDD 정수(20220106.0)와 일반 datetime 문자열/객체 모두 처리.
    """
    if series.empty or series.dropna().empty:
        return pd.Series(pd.NaT, index=series.index)
    _sample = series.dropna().iloc[0]
    # YYYYMMDD 정수 감지: 19000101 ~ 20991231 범위 숫자
    if isinstance(_sample, (int, float)) and 19000101 <= _sample <= 20991231:
        _s = pd.to_numeric(series, errors="coerce")
        _valid = _s.between(19000101, 20991231)
        result = pd.Series(pd.NaT, index=series.index)
        result[_valid] = pd.to_datetime(
            _s[_valid].astype(int).astype(str), format="%Y%m%d", errors="coerce"
        )
        return result
    return pd.to_datetime(series, errors="coerce")


def _extract_channel(filename: str) -> str:
    """파일명에서 채널 추출 (KT/그룹사/외부사/지입자재)"""
    for ch in CHANNELS:
        if ch in filename:
            return ch
    return "기타"


def _looks_like_report_header(df: pd.DataFrame) -> bool:
    """컬럼 대부분이 Unnamed: 이면 리포트형 파일로 판단."""
    unnamed = sum(1 for c in df.columns if str(c).startswith("Unnamed:"))
    return len(df.columns) > 5 and unnamed / len(df.columns) > 0.7


def _find_real_header_row(filepath: Path, sheet_name) -> int | None:
    """헤더 행 자동 탐색: header=1~8 중 Unnamed: 비율이 30% 미만인 첫 번째 행 번호 반환."""
    for h in range(1, 9):
        try:
            probe = pd.read_excel(filepath, sheet_name=sheet_name, header=h, nrows=2,
                                  engine=_detect_excel_engine(filepath))
            unnamed = sum(1 for c in probe.columns if str(c).startswith("Unnamed:"))
            if len(probe.columns) > 3 and unnamed / len(probe.columns) < 0.3:
                return h
        except Exception:
            continue
    return None


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
    for sheet_name, sdf in sheets.items():
        sdf.columns = sdf.columns.str.strip()
        # 리포트형 파일 감지: Unnamed: 컬럼이 70% 이상이면 헤더 행 재탐색
        if _looks_like_report_header(sdf):
            real_h = _find_real_header_row(filepath, sheet_name)
            if real_h is not None:
                log.info(f"  리포트형 파일 감지 → header={real_h} 로 재로드 ({filepath.name})")
                sdf = _read_excel_safe(filepath, sheet_name=sheet_name, header=real_h)
                sdf.columns = sdf.columns.str.strip()
        frames.append(sdf)
    return pd.concat(frames, ignore_index=True)


def load_and_preprocess(filepaths: list[Path], col_map: dict, lead_time: int):
    """복수 파일을 읽어 하나의 DataFrame으로 병합 후 전처리"""

    # ── 파일 레벨 skip 키워드 (트랜잭션 데이터 아닌 참조/매핑 파일) ──────
    _SKIP_KEYWORDS = ["카테고리목록", "미분류", "담당자", "mapping_master"]

    all_frames = []
    for fp in filepaths:
        # 비데이터 파일 skip
        if any(kw in fp.name for kw in _SKIP_KEYWORDS):
            log.info(f"파일 skip (비데이터): {fp.name}")
            continue

        log.info(f"파일 로드: {fp.name}")
        try:
            raw = _read_single_file(fp)
            raw.columns = raw.columns.str.strip()

            # ── 컬럼 목록 로그 기록 ─────────────────────────────────────
            log.info(f"  컬럼 목록 ({len(raw.columns)}개): {list(raw.columns)}")

            # ── 정산현황 파일 호환 처리 ───────────────────────────────────
            # 발주일자가 없으면 정산현황 파일로 간주 → 정산일자를 _date 대용으로 주입
            date_col = col_map.get("date", "발주일자")
            if date_col not in raw.columns:
                # 발주일 계열 → 정산일자 계열 순으로 탐색
                _settle_candidates = ["발주일", "주문일자", "주문일",
                                      "정산일자", "정산확정일", "일일정산일", "일일정산월"]
                _fallback_date = next(
                    (c for c in _settle_candidates if c in raw.columns), None
                )
                if _fallback_date:
                    log.info(f"  '{date_col}' 없음 → '{_fallback_date}'를 발주일자 대용으로 사용 "
                             f"({fp.name})")
                    raw[date_col] = raw[_fallback_date]
                else:
                    log.warning(f"  날짜 컬럼({date_col}) 및 정산일자 모두 없음 → skip ({fp.name})")
                    log.warning(f"  사용 가능한 컬럼: {list(raw.columns)}")
                    continue

            # ── 금액 컬럼 없으면 skip ────────────────────────────────────
            amount_col = col_map.get("amount", "판매금액")
            if amount_col not in raw.columns:
                _amount_candidates = ["정산금액", "매출액", "금액", "판매금액"]
                _fallback_amt = next(
                    (c for c in _amount_candidates if c in raw.columns), None
                )
                if _fallback_amt:
                    log.info(f"  '{amount_col}' 없음 → '{_fallback_amt}'를 금액 대용으로 사용 "
                             f"({fp.name})")
                    raw[amount_col] = raw[_fallback_amt]
                else:
                    log.warning(f"  금액 컬럼({amount_col}) 없음 → skip ({fp.name})")
                    continue

            raw["_source_file"] = fp.name   # 출처 파일 추적용

            # ── 필요 컬럼만 선택 (concat 시 불필요한 94컬럼 폭발 방지) ──────
            _needed = set(col_map.values()) | {"_source_file"}
            # 정산현황 파일에 정산확정일/일일정산월이 별도로 있으면 추가 보존
            _SETTLE_EXTRA = {"정산확정일", "일일정산월"}
            _keep = [c for c in raw.columns if c in _needed or c in _SETTLE_EXTRA]
            raw = raw[_keep]

            all_frames.append(raw)
            log.info(f"  → {len(raw):,}행 읽음 ({len(raw.columns)}개 컬럼 선택)")
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

    # ── 숫자형 날짜 감지 & 재변환 ────────────────────────────────────────
    # pd.to_datetime(숫자) 는 기본적으로 나노초로 해석 → 1970-01-01 오파싱
    # 파싱 후 날짜가 1970년 근방이면 숫자형 날짜로 판단해 형식 자동 감지
    valid_dates = df["_date"].dropna()
    if len(valid_dates) > 0 and valid_dates.dt.year.between(1969, 1972).mean() > 0.5:
        numeric = pd.to_numeric(df[date_col], errors="coerce")
        num_valid = numeric.dropna()

        if len(num_valid) > 0:
            med = float(num_valid.median())

            if 20_000_101 <= med <= 20_991_231:
                # YYYYMMDD 정수 예: 20221001 → 2022-10-01
                log.warning(f"날짜 컬럼({date_col}) YYYYMMDD 정수 감지(중앙값={med:.0f}) → 재변환")
                df["_date"] = pd.to_datetime(
                    numeric.astype("Int64").astype(str).str.zfill(8),
                    format="%Y%m%d", errors="coerce"
                )

            elif 30_000 <= med <= 80_000:
                # Excel 시리얼 (일 단위, 1899-12-30 기준) 예: 43831 → 2020-01-06
                log.warning(f"날짜 컬럼({date_col}) Excel 시리얼(일) 감지(중앙값={med:.0f}) → 재변환")
                df["_date"] = pd.to_datetime(numeric, unit="D",
                                              origin="1899-12-30", errors="coerce")

            elif 900_000_000_000 <= med <= 2_000_000_000_000:
                # Unix 밀리초 타임스탬프 예: 1580947200000 → 2020-02-06
                log.warning(f"날짜 컬럼({date_col}) Unix ms 타임스탬프 감지(중앙값={med:.0f}) → 재변환")
                df["_date"] = pd.to_datetime(numeric, unit="ms", errors="coerce")

            elif 900_000_000 <= med <= 2_000_000_000:
                # Unix 초 타임스탬프 예: 1580947200 → 2020-02-06
                log.warning(f"날짜 컬럼({date_col}) Unix 초 타임스탬프 감지(중앙값={med:.0f}) → 재변환")
                df["_date"] = pd.to_datetime(numeric, unit="s", errors="coerce")

            else:
                # 마지막 수단: 문자열로 변환 후 여러 포맷 시도
                log.warning(f"날짜 컬럼({date_col}) 형식 미감지(중앙값={med:.0f}) → 문자열 포맷 시도")
                s = df[date_col].astype(str).str.strip()
                for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y"):
                    parsed = pd.to_datetime(s, format=fmt, errors="coerce")
                    if parsed.notna().mean() > 0.5:
                        df["_date"] = parsed
                        break
                else:
                    df["_date"] = pd.to_datetime(s, errors="coerce")

            converted = df["_date"].notna().sum()
            sample_yr  = df["_date"].dropna().dt.year.value_counts().head(3).to_dict()
            log.info(f"  → 날짜 재변환: {converted:,}건 성공, 연도 분포: {sample_yr}")

    if df["_date"].isna().all():
        def parse_kor(s):
            try:
                m = int(str(s).replace("월","").split()[0])
                return pd.Timestamp(f"2024-{m:02d}-01")
            except Exception:
                return pd.NaT
        df["_date"] = df[date_col].apply(parse_kor)

    df = df.dropna(subset=["_date"])
    df["_ym"] = df["_date"].dt.to_period("M")

    # ── 당월(불완전) 데이터 제외 ─────────────────────────────────────────
    # 오늘 기준 현재 월은 데이터가 불완전 → 훈련 제외, 예측 대상으로 전환
    current_ym = pd.Period(datetime.now(), freq="M")
    before_cut = len(df)
    df = df.loc[df["_ym"] < current_ym].reset_index(drop=True)
    cut_rows = before_cut - len(df)
    if cut_rows > 0:
        log.info(f"  당월({current_ym}) 불완전 데이터 제외: {cut_rows:,}행 "
                 f"(예측 대상으로 전환)")

    # ── 정산확정일 → _settle_ym (매출 귀속 기준) ─────────────────────────
    # KT 귀속 규칙: 정산확정일 1~15일 → 전월 귀속, 16일~ → 당월 귀속
    if col_map.get("settle_date") and col_map["settle_date"] != date_col:
        raw_settle = df[col_map["settle_date"]]
        settle_dt  = pd.to_datetime(raw_settle, errors="coerce")
        # YYYYMMDD 정수 감지
        valid_s = settle_dt.dropna()
        if len(valid_s) > 0 and valid_s.dt.year.between(1969, 1972).mean() > 0.5:
            num_s = pd.to_numeric(raw_settle, errors="coerce")
            med_s = float(num_s.dropna().median()) if len(num_s.dropna()) > 0 else 0
            if 20_000_101 <= med_s <= 20_991_231:
                settle_dt = pd.to_datetime(
                    num_s.astype("Int64").astype(str).str.zfill(8),
                    format="%Y%m%d", errors="coerce")
        df["_settle_ym"] = settle_dt.dt.to_period("M")
        # NaT → _ym fallback (정산일 없는 행은 발주월 기준 처리)
        null_settle = df["_settle_ym"].isna()
        if null_settle.any():
            df.loc[null_settle, "_settle_ym"] = df.loc[null_settle, "_ym"]
            log.info(f"  _settle_ym NaT {null_settle.sum():,}행 → _ym으로 대체")
        # 일일정산월 컬럼이 있으면 _settle_ym 보정
        # KT 정산현황: 5월 정산이 6월 초 확정 → 정산확정일=2026-06 이지만 일일정산월=2026-05
        if "일일정산월" in df.columns:
            _ym_parsed = pd.to_datetime(
                df["일일정산월"].astype(str).str.strip(), format="%Y-%m", errors="coerce"
            ).dt.to_period("M")
            _has_ym = _ym_parsed.notna()
            if _has_ym.any():
                df.loc[_has_ym, "_settle_ym"] = _ym_parsed[_has_ym]
                log.info(f"  일일정산월 → _settle_ym 보정: {_has_ym.sum():,}행")
        # 당월 이후 정산 데이터는 미래 정산 → 유지 (예측 목적)
        # 단, 과거 발주 중 당월 이후 정산 예정인 것도 포함
        log.info(f"  정산일 기준 기간: "
                 f"{df['_settle_ym'].dropna().min()} ~ "
                 f"{df['_settle_ym'].dropna().max()}")
    else:
        df["_settle_ym"] = df["_ym"]   # fallback: 발주일과 동일

    # ── _recv_ym: 채널별_분석 실적 귀속 기준 월 ────────────────────────────
    # 규칙 (행 단위 적용):
    #   ① 입고일자 있음 → 입고일자 기준 월  (KT 전체, 그룹사 전량입고)
    #   ② 입고일자 없음 + 주문상태=배송완료 + 정산일자 있음
    #      → 정산일자 기준 월               (그룹사 배송완료 정산)
    #   ③ 위 조건 모두 불충족 → _settle_ym fallback
    df["_recv_ym"] = df["_settle_ym"]   # 기본값: fallback

    # ① 입고일자/배송완료일자 → _recv_ym
    # 후보 컬럼: col_map["recv_date"] 우선 + 하드코딩 alias 목록 (df에 있는 것만)
    _recv_candidates = []
    if col_map.get("recv_date") and col_map["recv_date"] in df.columns:
        _recv_candidates.append(col_map["recv_date"])
    for _alias in ["입고일자", "입고일", "배송완료일자", "배송완료일", "recv_date"]:
        if _alias in df.columns and _alias not in _recv_candidates:
            _recv_candidates.append(_alias)
    log.info(f"  _recv_ym 후보 컬럼: {_recv_candidates}")
    for _dc in _recv_candidates:
        _nn = int(df[_dc].notna().sum())
        _sample = df[_dc].dropna().head(3).tolist()
        log.info(f"    [{_dc}] non-null: {_nn:,}행, sample: {_sample}")

    _n_recv_total = 0
    for _recv_col_try in _recv_candidates:
        # YYYYMMDD 정수 포함 스마트 파싱
        rdt = _parse_date_col(df[_recv_col_try])
        recv_ym = rdt.dt.to_period("M")
        mask_recv = recv_ym.notna() & (rdt.dt.year >= 2015) & (rdt.dt.year <= 2035)
        _n_valid = int(mask_recv.sum())
        _n_bad = int(recv_ym.notna().sum()) - _n_valid
        log.info(f"  _recv_ym 시도 [{_recv_col_try}]: non-null={int(recv_ym.notna().sum()):,}, "
                 f"유효={_n_valid:,}, 제외={_n_bad:,}")
        if _n_bad > 0:
            log.warning(f"  _recv_ym [{_recv_col_try}] 유효범위 외 {_n_bad:,}행 무시")
        if _n_valid > 0:
            _not_yet_set = df["_recv_ym"] == df["_settle_ym"]
            _apply_mask = mask_recv & _not_yet_set
            df.loc[_apply_mask, "_recv_ym"] = recv_ym[_apply_mask]
            _n_recv_total += int(_apply_mask.sum())
            log.info(f"  _recv_ym ① [{_recv_col_try}] 적용: {int(_apply_mask.sum()):,}행")
    log.info(f"  _recv_ym ① 입고/배송완료일 적용 합계: {_n_recv_total:,}행")

    # ② 배송완료 + 입고일자 없음 + 정산일자 있음 → 정산일자 기준
    #    그룹사 중 배송완료 정산 사업장: 정산일자가 실적 귀속 월
    _has_status = "_order_status" in df.columns
    _has_settle = "_settle_ym" in df.columns
    if _has_status and _has_settle:
        _no_recv = df["_recv_ym"] == df["_settle_ym"]   # 입고일자 미적용 행
        _delivery_done = df["_order_status"].isin(ORDER_STATUS_COMPLETE - {"전량입고"})
        # "배송완료"만: 전량입고는 이미 입고일 적용됨
        _delivery_only = df["_order_status"] == "배송완료"
        mask_deliv = _no_recv & _delivery_only & _has_settle
        # 이미 _settle_ym을 기본값으로 넣었으므로 추가 작업 불필요
        # 단, 어느 행이 이 규칙을 따르는지 로그만 남김
        _n_deliv = (_no_recv & _delivery_only).sum()
        if _n_deliv > 0:
            log.info(f"  _recv_ym ② 배송완료+입고일 없음 → 정산일 기준: {_n_deliv:,}행")

    log.info(f"  _recv_ym 최종 기간: "
             f"{df['_recv_ym'].dropna().min()} ~ "
             f"{df['_recv_ym'].dropna().max()}")

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

    # ── 통신/일반 구분 (mapping_master.xlsx 기반) ────────────────────────
    prod_type_map = load_prod_type_map()
    if prod_type_map and col_map.get("category"):
        cat_series = df[col_map["category"]].astype(str).str.strip()
        # 1단계: 정확 매핑
        df["_prod_type"] = cat_series.map(prod_type_map).fillna("미분류")
        # 2단계: 미분류 → 마지막 키워드 추론
        n_miss = (df["_prod_type"] == "미분류").sum()
        if n_miss > 0:
            rightmost_map = _build_rightmost_map(prod_type_map)
            miss_mask = df["_prod_type"] == "미분류"
            rightmost_s = cat_series[miss_mask].str.split(">").str[-1].str.strip()
            df.loc[miss_mask, "_prod_type"] = rightmost_s.map(rightmost_map).fillna("미분류")
        n_t = (df["_prod_type"] == "통신").sum()
        n_g = (df["_prod_type"] == "일반").sum()
        n_u = (df["_prod_type"] == "미분류").sum()
        log.info(f"  통신/일반 구분: 통신 {n_t:,}행 / 일반 {n_g:,}행 / 미분류 {n_u:,}행")
    else:
        df["_prod_type"] = "미분류"
        if not prod_type_map:
            log.info("  통신/일반 구분: 매핑 파일 없음 → 전체 미분류")

    # ── 채널 추출 (파일명 기반: KT/그룹사/외부사/지입자재) ───────────────
    df["_channel"] = df["_source_file"].apply(_extract_channel)

    # ── 주문번호 + 품목번호 → 주문 라인 키 ──────────────────────────────────
    # 하나의 주문번호에 N개 품목번호가 매칭되어 라인을 형성
    # 중복 제거 및 OLFR 계산의 기본 단위는 (주문번호, 품목번호) 조합
    if col_map.get("order_id") and col_map["order_id"] in df.columns:
        df["_order_id"] = df[col_map["order_id"]].astype(str).str.strip()
    if col_map.get("line_id") and col_map["line_id"] in df.columns:
        df["_line_id"] = df[col_map["line_id"]].astype(str).str.strip()
    # 라인 고유 키: 주문번호-품목번호 (둘 다 있을 때만)
    if "_order_id" in df.columns and "_line_id" in df.columns:
        df["_line_key"] = df["_order_id"] + "-" + df["_line_id"]
    elif "_order_id" in df.columns:
        df["_line_key"] = df["_order_id"]

    # ── 주문상태 기반 수요 보정 ──────────────────────────────────────────
    if col_map.get("order_status") and col_map["order_status"] in df.columns:
        status = df[col_map["order_status"]].astype(str).str.strip()
        df["_order_status"] = status

        # 취소/반려 → 수량·금액 0 처리 (수요에서 완전 제외)
        cancel_mask = status.isin(ORDER_STATUS_CANCEL)
        df.loc[cancel_mask, "_qty"]    = 0
        df.loc[cancel_mask, "_amount"] = 0

        # 반품완료 → 수량·금액 음수 처리 (실수요 차감)
        return_mask = status.isin(ORDER_STATUS_RETURN)
        df.loc[return_mask, "_qty"]    = -df.loc[return_mask, "_qty"].abs()
        df.loc[return_mask, "_amount"] = -df.loc[return_mask, "_amount"].abs()

        n_cancel = cancel_mask.sum()
        n_return = return_mask.sum()
        if n_cancel > 0 or n_return > 0:
            log.info(f"  주문상태 보정: 취소/반려 {n_cancel:,}건(=0), 반품완료 {n_return:,}건(음수처리)")
    else:
        # 주문상태 컬럼 없을 때 기존 취소수량·반품수량 방식으로 폴백
        if col_map.get("cancel_qty"):
            cq = pd.to_numeric(df[col_map["cancel_qty"]], errors="coerce").fillna(0)
            df["_qty"] = (df["_qty"] - cq).clip(lower=0)
        if col_map.get("return_qty"):
            rq = pd.to_numeric(df[col_map["return_qty"]], errors="coerce").fillna(0)
            df["_qty"]    = (df["_qty"]    - rq).clip(lower=0)
            df["_amount"] = (df["_amount"] * (df["_qty"] /
                              (df["_qty"] + rq).replace(0, np.nan)).fillna(1)).fillna(0)

    # ── 리드타임 실측값 추출 ─────────────────────────────────────────────
    # 기준: 배송완료일(recv_date) - 발주일자(date) 실측값 우선
    # 주문시배송리드타임은 사전 예상값으로 실질과 다르므로 fallback으로만 사용
    # 우선순위: ① 배송완료일-발주일자 실측 ② 입고일-주문일 계산 ③ 주문시배송리드타임(표준값)
    _lt_computed = False
    if col_map.get("recv_date") and col_map["recv_date"] in df.columns:
        rdt = _parse_date_col(df[col_map["recv_date"]])
        rdt = rdt.where(rdt.dt.year >= 2015, other=pd.NaT)
        odt_base = _parse_date_col(df[date_col])   # 발주일자
        _lt_real = (rdt - odt_base).dt.days
        valid = _lt_real.between(1, 365)
        if valid.sum() > 0:
            df["_lt"] = np.where(valid, _lt_real, np.nan)
            log.info(f"  리드타임: 배송완료일-발주일자 실측 {valid.sum():,}건 산정")
            _lt_computed = True

    if not _lt_computed:
        # fallback ①: 입고일-주문일
        if col_map.get("order_date") and col_map["order_date"] in df.columns \
                and col_map.get("recv_date") and col_map["recv_date"] in df.columns:
            odt = _parse_date_col(df[col_map["order_date"]])
            rdt = _parse_date_col(df[col_map["recv_date"]])
            _lt_calc = (rdt - odt).dt.days
            valid = _lt_calc.between(1, 365)
            if valid.sum() > 0:
                df["_lt"] = np.where(valid, _lt_calc, np.nan)
                log.info(f"  리드타임: 입고일-주문일 계산 {valid.sum():,}건 산정 (fallback)")
                _lt_computed = True

    # _lt_order: 주문시배송리드타임 표준값 보존 (품목별 fallback용)
    if col_map.get("lt_order") and col_map["lt_order"] in df.columns:
        df["_lt_order"] = pd.to_numeric(df[col_map["lt_order"]], errors="coerce")
        df["_lt_order"] = df["_lt_order"].where(df["_lt_order"].between(1, 365), other=np.nan)

    if not _lt_computed:
        # fallback ②: 주문시배송리드타임 표준값
        if col_map.get("lt_order") and col_map["lt_order"] in df.columns:
            df["_lt"] = pd.to_numeric(df[col_map["lt_order"]], errors="coerce")
            log.info("  리드타임: 주문시배송리드타임(표준값) 사용 (fallback)")
        elif col_map.get("lt_std") and col_map["lt_std"] in df.columns:
            df["_lt"] = pd.to_numeric(df[col_map["lt_std"]], errors="coerce")
            log.info("  리드타임: 상품배송리드타임(표준값) 사용 (fallback)")
        else:
            df["_lt"] = np.nan
            log.info("  리드타임: 산정 불가 (관련 컬럼 없음)")

    # 이상값 제거: 0 이하 또는 365일 초과
    df.loc[df["_lt"] <= 0,   "_lt"] = np.nan
    df.loc[df["_lt"] > 365,  "_lt"] = np.nan

    # ── 매입금액 (마진 분석용) ────────────────────────────────────────────
    if col_map.get("cost"):
        df["_cost"] = pd.to_numeric(df[col_map["cost"]], errors="coerce").fillna(0)

    # ── 관리회계 (유가증권/배송료/프로모션 분류) ──────────────────────────
    if col_map.get("mgmt_acct") and col_map["mgmt_acct"] in df.columns:
        df["_mgmt_acct"] = df[col_map["mgmt_acct"]].astype(str).str.strip()
        df["_mgmt_acct"] = df["_mgmt_acct"].replace({"nan": "—", "None": "—", "": "—"})
        log.info(f"  관리회계: {df['_mgmt_acct'].nunique()}개 고유값 로드")

    log.info(f"  최종 행수: {len(df):,}  |  기간: {df['_ym'].min()} ~ {df['_ym'].max()}")
    lt_valid = df["_lt"].notna().sum()
    if lt_valid > 0:
        log.info(f"  리드타임 실측값: {lt_valid:,}건 (평균 {df['_lt'].mean():.1f}일, "
                 f"σ={df['_lt'].std():.1f}일)")
    return df


def build_monthly_series(df: pd.DataFrame, group_cols=None,
                         ym_col: str = "_ym") -> pd.DataFrame:
    """
    group_cols=None → 전체 집계, list → 그룹별
    ym_col: 집계 기준 연월 컬럼 (_ym=발주일 기준, _settle_ym=정산일 기준)
    """
    valid = df[df[ym_col].notna()].copy()
    if group_cols:
        grp = valid.groupby(group_cols + [ym_col]).agg(
            amount=("_amount", "sum"),
            qty=("_qty", "sum"),
        ).reset_index()
        grp.rename(columns={ym_col: "_ym"}, inplace=True)
    else:
        grp = valid.groupby(ym_col).agg(
            amount=("_amount", "sum"),
            qty=("_qty", "sum"),
        ).reset_index()
        grp.rename(columns={ym_col: "_ym"}, inplace=True)
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
    # 그룹별 총 금액 (채널 정보 포함)
    grp_cols = ["_vendor", "_prod_key"]
    if "_channel" in group_df.columns:
        grp_cols = ["_channel", "_vendor", "_prod_key"]
    summary = (
        group_df.groupby(grp_cols)
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
# 3-B. 제품 수명주기 감지 (MRO 이관 품목용)
# ════════════════════════════════════════════════════════════════════════════

def detect_lifecycle(qty_series: np.ndarray,
                     base_ym: "pd.Period | None" = None
                     ) -> dict:
    """
    월별 수량 시계열에서 수명주기 단계와 EOL 예상 시점을 산출.

    반환 dict:
        stage    : "성장기" | "성숙기" | "쇠퇴기" | "단종임박"
        eol_ym   : pd.Period | None  — 수요 0 예상 월 (쇠퇴기/단종임박만)
        slope    : float             — 최근 12개월 월간 수량 변화율 (개/월)
        decline_pct : float          — 최근 6개월 평균 / 피크 6개월 평균 비율 (%)
        note     : str               — 판단 근거 메시지

    판정 기준:
        - 쇠퇴기   : 최근 6개월 평균 < 피크 6개월 평균의 60%, 하락 추세 확인
        - 단종임박  : 최근 3개월 평균 < 피크 6개월 평균의 30%
        - 성장기   : 최근 6개월 평균 > 피크 6개월 평균의 110%
        - 성숙기   : 나머지
    """
    from sklearn.linear_model import LinearRegression

    n = len(qty_series)
    result = {"stage": "성숙기", "eol_ym": None, "slope": 0.0,
              "decline_pct": 100.0, "note": ""}

    if n < 6:
        result["note"] = f"데이터 {n}개월 — 판단 불가"
        return result

    peak_window = min(6, n)
    peak_avg = float(np.max([
        np.mean(qty_series[i:i + peak_window])
        for i in range(n - peak_window + 1)
    ]))
    recent6 = qty_series[-6:]
    recent3 = qty_series[-3:]
    recent6_avg = float(np.mean(recent6))
    recent3_avg = float(np.mean(recent3))

    decline_pct = (recent6_avg / peak_avg * 100) if peak_avg > 0 else 100.0
    result["decline_pct"] = round(decline_pct, 1)

    # 최근 12개월 선형 기울기
    win12 = qty_series[-12:] if n >= 12 else qty_series
    x = np.arange(len(win12)).reshape(-1, 1)
    try:
        lr = LinearRegression().fit(x, win12)
        slope = float(lr.coef_[0])
    except Exception:
        slope = 0.0
    result["slope"] = round(slope, 3)

    # 단계 판정
    if peak_avg > 0 and (recent3_avg / peak_avg) < 0.30 and slope < 0:
        result["stage"] = "단종임박"
        note = f"최근3개월평균 {recent3_avg:.1f}개 ({decline_pct:.0f}% of 피크), 기울기 {slope:+.2f}/월"
    elif peak_avg > 0 and decline_pct < 60 and slope < 0:
        result["stage"] = "쇠퇴기"
        note = f"최근6개월평균 {recent6_avg:.1f}개 ({decline_pct:.0f}% of 피크), 기울기 {slope:+.2f}/월"
    elif recent6_avg > peak_avg * 1.10 and slope > 0:
        result["stage"] = "성장기"
        note = f"최근6개월평균 {recent6_avg:.1f}개 ({decline_pct:.0f}% of 피크)"
        result["note"] = note
        return result
    else:
        result["stage"] = "성숙기"
        note = f"최근6개월평균 {recent6_avg:.1f}개 ({decline_pct:.0f}% of 피크)"
        result["note"] = note
        return result

    result["note"] = note

    # EOL 예상 시점 — 선형 외삽으로 수요=0 도달 월 추정
    if slope < 0 and base_ym is not None:
        last_val = float(qty_series[-1])
        # months_to_zero = -last_val / slope (음수 기울기이므로 양수)
        months_to_zero = int(np.ceil(-last_val / slope)) if slope < 0 else None
        if months_to_zero is not None and 1 <= months_to_zero <= 60:
            result["eol_ym"] = base_ym + months_to_zero

    return result


# ════════════════════════════════════════════════════════════════════════════
# 4. 예측 모델
# ════════════════════════════════════════════════════════════════════════════
def _mape(actual, pred):
    actual, pred = np.array(actual, dtype=float), np.array(pred, dtype=float)
    mask = np.abs(actual) > 1e-6   # 극소값 제외 (0원에 가까운 실적은 % 오차 무의미)
    if mask.sum() == 0:
        return np.nan
    pct_errors = np.abs((actual[mask] - pred[mask]) / actual[mask]) * 100
    pct_errors = np.clip(pct_errors, 0, 300)   # 개별 오차 최대 300% 캡핑
    return float(np.mean(pct_errors))

def _mae(actual, pred):
    return np.mean(np.abs(np.array(actual) - np.array(pred)))

def _rmse(actual, pred):
    return np.sqrt(np.mean((np.array(actual) - np.array(pred))**2))

def _tracking_signal(actual, pred):
    """
    Tracking Signal = 대수합(오차) / MAD
    CPSM Module 2 Ch.3 기준:
      TS > +4 : 지속 과소예측 (수요 > 예측) → 재고 부족 위험
      TS < -4 : 지속 과대예측 (수요 < 예측) → 재고 과잉 위험
      -4 ≤ TS ≤ +4 : 예측 편향 없음 (정상)
    """
    actual, pred = np.array(actual, dtype=float), np.array(pred, dtype=float)
    errors = actual - pred          # A - F (양수 = 과소예측)
    algebraic_sum = np.sum(errors)
    mad = np.mean(np.abs(errors))
    if mad == 0:
        return np.nan
    return algebraic_sum / mad


def model_holt_winters(train: np.ndarray, h: int):
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    n = len(train)
    sp = 12 if n >= 24 else (6 if n >= 12 else None)
    if sp is None or n < sp * 2:
        raise ValueError("HW: 데이터 부족")
    try:
        # statsmodels 일부 버전에서 ndarray에 .iloc 호출 오류 → Series로 래핑
        train_s = pd.Series(train.astype(float))
        m = ExponentialSmoothing(
            train_s, trend="add", seasonal="add", seasonal_periods=sp,
            initialization_method="estimated",
        ).fit(optimized=True, use_brute=False)
        fc = np.array(m.forecast(h), dtype=float)
        resid_std = float(np.std(np.array(m.resid, dtype=float)))
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
        # statsmodels 호환성: ndarray → Series 래핑
        train_s = pd.Series(train.astype(float))
        m = SARIMAX(train_s, order=(1,1,1), seasonal_order=(1,1,0,12),
                    enforce_stationarity=False, enforce_invertibility=False).fit(
            disp=False, maxiter=200)
        fc_obj = m.get_forecast(steps=h)
        fc = np.array(fc_obj.predicted_mean, dtype=float)
        ci = fc_obj.conf_int(alpha=0.05).values   # DataFrame → ndarray
        resid_std = float(np.std(np.array(m.resid, dtype=float)))
        return fc, resid_std, ci
    except Exception as e:
        raise ValueError(f"SARIMA 실패: {e}")


def model_ma3(train: np.ndarray, h: int):
    val = float(np.mean(train[-3:])) if len(train) >= 3 else float(np.mean(train))
    fc = np.full(h, val)
    resid_std = float(np.std(train[-6:])) if len(train) >= 6 else float(np.std(train))
    return fc, resid_std


def model_prophet(train: np.ndarray, h: int):
    """Meta Prophet — 연간 계절성 + 변동점 자동 감지. 반환: (fc, resid_std, ci_lo, ci_hi)"""
    from prophet import Prophet
    n = len(train)
    if n < 12:
        raise ValueError("Prophet: 12개월 미만")
    dates = pd.date_range(
        end=pd.Timestamp.now().replace(day=1) - pd.DateOffset(months=1),
        periods=n, freq="MS",
    )
    df_p = pd.DataFrame({"ds": dates, "y": train.astype(float)})
    m = Prophet(
        yearly_seasonality=True,
        weekly_seasonality=False,
        daily_seasonality=False,
        interval_width=0.95,
        changepoint_prior_scale=0.05,
    )
    import logging as _lg
    _lg.getLogger("prophet").setLevel(_lg.WARNING)
    _lg.getLogger("cmdstanpy").setLevel(_lg.WARNING)
    m.fit(df_p)
    future = m.make_future_dataframe(periods=h, freq="MS")
    forecast = m.predict(future)
    fc     = np.maximum(forecast["yhat"].values[-h:], 0)
    ci_lo  = np.maximum(forecast["yhat_lower"].values[-h:], 0)
    ci_hi  = forecast["yhat_upper"].values[-h:]
    resid_std = float(np.std(train - forecast["yhat"].values[:n]))
    return fc, resid_std, ci_lo, ci_hi


def model_lgbm(train: np.ndarray, h: int):
    """LightGBM 래그 피처 — 반복 예측 방식. 반환: (fc, resid_std)"""
    import lightgbm as lgb
    n = len(train)
    if n < 13:
        raise ValueError("LightGBM: 13개월 미만")

    LAGS = [l for l in [1, 2, 3, 6, 12] if l < n]

    def make_features(arr, start_i):
        rows = []
        for i in range(start_i, len(arr)):
            feat = [arr[i - l] for l in LAGS]
            feat.append(i % 12)
            feat.append(float(np.mean(arr[max(0, i - 3):i])))
            rows.append(feat)
        return np.array(rows)

    max_lag  = max(LAGS)
    X_train  = make_features(train, max_lag)
    y_train  = train[max_lag:]
    if len(X_train) < 5:
        raise ValueError("LightGBM: 학습 샘플 부족")

    model = lgb.LGBMRegressor(
        n_estimators=100, learning_rate=0.05,
        num_leaves=15, min_child_samples=3,
        verbose=-1, random_state=42,
    )
    model.fit(X_train, y_train)

    history = list(train)
    preds = []
    for step in range(h):
        feat = [history[-l] for l in LAGS]
        feat.append((n + step) % 12)
        feat.append(float(np.mean(history[-3:])))
        pred = max(float(model.predict([feat])[0]), 0)
        preds.append(pred)
        history.append(pred)

    fc = np.array(preds)
    resid_std = float(np.std(y_train - model.predict(X_train)))
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

    # Prophet
    try:
        fc_pr, _, _, _ = model_prophet(train, h)
        results["Prophet"] = {
            "mape": _mape(test, fc_pr),
            "mae":  _mae(test, fc_pr),
            "rmse": _rmse(test, fc_pr),
        }
    except Exception as e:
        log.debug(f"Prophet 검증 실패: {e}")
        results["Prophet"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

    # LightGBM
    try:
        fc_lgb, _ = model_lgbm(train, h)
        results["LightGBM"] = {
            "mape": _mape(test, fc_lgb),
            "mae":  _mae(test, fc_lgb),
            "rmse": _rmse(test, fc_lgb),
        }
    except Exception as e:
        log.debug(f"LightGBM 검증 실패: {e}")
        results["LightGBM"] = {"mape": np.nan, "mae": np.nan, "rmse": np.nan}

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
            ci = res[2]   # model_sarima가 ndarray로 반환
            ci_lo = ci[:, 0]
            ci_hi = ci[:, 1]
        elif best_model == "Prophet":
            fc, rs, ci_lo, ci_hi = model_prophet(series, h)
        elif best_model == "LightGBM":
            fc, rs = model_lgbm(series, h)
        else:
            fc, rs = model_ma3(series, h)
    except Exception as e:
        log.warning(f"예측 모델 폴백(MA3): {e}")
        fc, rs = model_ma3(series, h)

    fc = np.maximum(np.array(fc, dtype=float), 0)
    if ci_lo is None:
        ci_lo = np.maximum(fc - Z_CI * rs, 0)
        ci_hi = fc + Z_CI * rs
    return fc, np.array(ci_lo, dtype=float), np.array(ci_hi, dtype=float)


# ════════════════════════════════════════════════════════════════════════════
# 5-B. 신제품 콜드스타트 예측
# ════════════════════════════════════════════════════════════════════════════
COLDSTART_THRESHOLD = 6   # 이 개월 수 미만이면 콜드스타트로 처리

def forecast_coldstart(arr: np.ndarray, category_avg: np.ndarray, h: int = 6):
    """
    데이터가 COLDSTART_THRESHOLD개월 미만인 신제품용 예측.
    동일 카테고리 평균 성장 패턴을 스케일해 적용.
    반환: (fc, ci_lo, ci_hi)
    """
    n = len(arr)
    if n == 0:
        return np.zeros(h), np.zeros(h), np.zeros(h)

    last_val = float(arr[-1]) if arr[-1] > 0 else float(np.mean(arr[arr > 0]) if np.any(arr > 0) else 1)

    if len(category_avg) >= h + 1:
        ref_level = float(np.mean(category_avg[-n:])) if np.mean(category_avg[-n:]) > 0 else 1
        scale = last_val / ref_level
        fc = np.maximum(category_avg[-h:] * scale, 0)
    else:
        fc = np.full(h, last_val)

    std = float(np.std(arr)) if n > 1 else last_val * 0.3
    ci_lo = np.maximum(fc - Z_CI * std, 0)
    ci_hi = fc + Z_CI * std
    return fc, ci_lo, ci_hi


# ════════════════════════════════════════════════════════════════════════════
# 5-C. 한국은행 ECOS 외부 지표 연동 (소비자심리지수 등)
# ════════════════════════════════════════════════════════════════════════════
ECOS_API_KEY_FILE = BASE_DIR / "ecos_api_key.txt"
ECOS_SERIES = {
    "소비자심리지수(CSI)": ("521Y001", "I22A"),
    "경기동행지수":        ("901Y067", "I_CLI"),
}

def fetch_ecos_data(n_months: int = 36) -> dict:
    """
    한국은행 ECOS Open API에서 거시경제 지표를 가져옵니다.
    API 키가 없으면 빈 dict 반환 (선택 기능).

    API 키 발급: https://ecos.bok.or.kr/api/#/DevGuide/APIKey
    발급 후 ecos_api_key.txt 파일에 키를 저장하세요.
    """
    if not ECOS_API_KEY_FILE.exists():
        log.info("ECOS API 키 없음 → 외부지표 시트 생략 (ecos_api_key.txt 생성 시 활성화)")
        return {}

    api_key = ECOS_API_KEY_FILE.read_text(encoding="utf-8").strip()
    if not api_key:
        return {}

    try:
        import requests
    except ImportError:
        log.warning("requests 미설치 → ECOS 데이터 생략")
        return {}

    end_dt   = pd.Timestamp.now()
    start_dt = end_dt - pd.DateOffset(months=n_months + 1)
    start_str = start_dt.strftime("%Y%m")
    end_str   = end_dt.strftime("%Y%m")

    results = {}
    for label, (stat_code, item_code) in ECOS_SERIES.items():
        try:
            url = (
                f"https://ecos.bok.or.kr/api/StatisticSearch/{api_key}/json/kr/1/100/"
                f"{stat_code}/MM/{start_str}/{end_str}/{item_code}"
            )
            resp = requests.get(url, timeout=10)
            data = resp.json()
            rows = data.get("StatisticSearch", {}).get("row", [])
            if not rows:
                continue
            df_e = pd.DataFrame(rows)[["TIME", "DATA_VALUE"]].copy()
            df_e["ym"]  = df_e["TIME"].apply(lambda x: pd.Period(x, freq="M"))
            df_e["val"] = pd.to_numeric(df_e["DATA_VALUE"], errors="coerce")
            df_e = df_e.dropna(subset=["val"]).sort_values("ym")
            results[label] = df_e[["ym", "val"]].reset_index(drop=True)
            log.info(f"ECOS [{label}]: {len(df_e)}개월 로드 완료")
        except Exception as e:
            log.warning(f"ECOS [{label}] 로드 실패: {e}")

    return results


# ════════════════════════════════════════════════════════════════════════════
# 6. CPSM 지표
# ════════════════════════════════════════════════════════════════════════════
DEFAULT_LEAD_TIME  = 30   # 리드타임 파일에 없는 품목의 기본값 (일)
OFR_TARGET_DAYS    = 7    # 주문충족률(OFR) 기준 납기일 — 고객 기대납기 (케이티커머스 drop-ship 기준)

# ── 주문상태 그룹 분류 ────────────────────────────────────────────────────
# 수요 집계 시 상태별 처리 기준
ORDER_STATUS_CANCEL   = {"주문취소", "주문승인반려", "운영사결재반려"}          # 수요=0 (제외)
ORDER_STATUS_RETURN   = {"반품완료"}                                          # 수요 차감 (음수)
ORDER_STATUS_RETURN_REQ = {"반품요청", "반송요청(상품권)", "반송완료(상품권)"}  # 반품 진행중 (유지)
ORDER_STATUS_COMPLETE = {"전량입고", "부분입고", "배송완료"}                   # 완료
ORDER_STATUS_INPROG   = {"주문승인요청", "주문승인중", "입금대기", "오더볼력",
                          "주문접수", "운영사결재중", "출하준비중", "배송중",
                          "픽업요청", "픽업완료", "출하대기", "입고대기"}       # 진행중
SETTLE_LAG_MONTHS  = 1    # 정산 지연 개월 수 (당월+N개월 전까지만 정산 완성으로 간주)
                          # 예: 1 → 5월말 기준으로 4월까지만 완성, 5·6월은 예측 대상
                          # 매월 정산이 익월 말까지 완료되면 1, 2개월 후면 2로 설정
ORDER_LAG_MONTHS   = 1    # 발주일 기준 데이터 완성 지연 개월 수
                          # 월말 마감 전 입력 지연으로 전월도 불완전할 수 있음
                          # 예: 1 → 6월 실행 시 5월 데이터는 훈련 제외, 예측 대상으로 전환
                          # 0으로 설정하면 당월만 제외 (기존 동작)

def build_lt_stats(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    협력사×상품코드별 리드타임 통계 + OFR(주문충족률) 산출.
    반환: (lt_grp DataFrame, lt_order_lookup dict)
      lt_grp: vendor, prod_key, lt_mean(일), lt_std(일), lt_n(건수), ofr(%)
      lt_order_lookup: prod_key → 주문시배송리드타임 중앙값 (실측 없는 품목 fallback용)
    OFR = 실제납기 ≤ OFR_TARGET_DAYS 인 건수 비율 (케이티커머스 drop-ship 기준)
    """
    empty = pd.DataFrame(columns=["vendor", "prod_key", "lt_mean", "lt_std", "lt_n", "ofr"])

    # 주문시배송리드타임 품목별 중앙값 lookup (실측 없는 품목 fallback)
    lt_order_lookup = {}
    if "_lt_order" in df.columns:
        _lo = df[["_prod_key", "_lt_order"]].copy()
        _lo["_lt_order"] = pd.to_numeric(_lo["_lt_order"], errors="coerce")
        _lo = _lo[_lo["_lt_order"].between(1, 365)]
        if len(_lo) > 0:
            lt_order_lookup = _lo.groupby("_prod_key")["_lt_order"].median().to_dict()
            log.info(f"주문시배송리드타임 lookup: {len(lt_order_lookup)}개 품목")

    if "_lt" not in df.columns or df["_lt"].notna().sum() == 0:
        return empty, lt_order_lookup

    # 1일 이하는 당일출고/기록오류로 간주해 제외 (드롭십 기준 최소 2일)
    lt_valid = df[df["_lt"].notna() & (df["_lt"] >= 2)]

    lt_grp = (
        lt_valid
        .groupby(["_vendor", "_prod_key"])["_lt"]
        .agg(lt_mean="mean", lt_std="std", lt_n="count")
        .reset_index()
        .rename(columns={"_vendor": "vendor", "_prod_key": "prod_key"})
    )
    lt_grp["lt_std"] = lt_grp["lt_std"].fillna(0)

    # OFR: 기준 납기일(OFR_TARGET_DAYS) 이내 이행 비율
    ofr_s = (
        lt_valid
        .groupby(["_vendor", "_prod_key"])["_lt"]
        .apply(lambda x: (x <= OFR_TARGET_DAYS).mean() * 100)
        .reset_index()
        .rename(columns={"_vendor": "vendor", "_prod_key": "prod_key", "_lt": "ofr"})
    )
    lt_grp = lt_grp.merge(ofr_s, on=["vendor", "prod_key"], how="left")
    lt_grp["ofr"] = lt_grp["ofr"].fillna(np.nan)

    n_low_ofr = (lt_grp["ofr"] < 80).sum()
    log.info(f"리드타임 통계: {len(lt_grp)}개 품목 산출 "
             f"(전체 평균 {lt_grp['lt_mean'].mean():.1f}일, "
             f"OFR<80% 품목 {n_low_ofr}개)")
    return lt_grp, lt_order_lookup


def get_lt_stats(lt_stats: pd.DataFrame, lt_df: pd.DataFrame,
                 vendor: str, prod_key: str,
                 lt_order_lookup: dict | None = None) -> tuple[float, float]:
    """
    품목별 리드타임 (평균, 표준편차) 반환.
    우선순위: ① 실측 통계(품목) ② 실측 통계(협력사 평균) ③ 주문시배송리드타임(표준값)
              ④ 리드타임 파일 ⑤ 실측 전체 평균 ⑥ 기본값
    """
    LT_MIN = 3.0   # 드롭십 기준 최소 보장 리드타임(일) — 1~2일 실측은 이상값 처리 후 여기서 한번 더 보정

    def _floor(mean: float, std: float) -> tuple[float, float]:
        return max(mean, LT_MIN), std

    # ① 실측 통계 — 품목 단위
    if not lt_stats.empty:
        row = lt_stats[(lt_stats["vendor"] == vendor) &
                       (lt_stats["prod_key"] == prod_key)]
        if not row.empty:
            return _floor(float(row.iloc[0]["lt_mean"]), float(row.iloc[0]["lt_std"]))

        # ② 실측 통계 — 협력사 평균
        row_v = lt_stats[lt_stats["vendor"] == vendor]
        if not row_v.empty:
            return _floor(float(row_v["lt_mean"].mean()), float(row_v["lt_std"].mean()))

        # ③ 주문시배송리드타임 (품목별 표준값)
        if lt_order_lookup:
            lo_val = lt_order_lookup.get(prod_key)
            if lo_val and lo_val > 0:
                return _floor(float(lo_val), 0.0)

        # ⑤ 실측 전체 평균
        return _floor(float(lt_stats["lt_mean"].mean()), float(lt_stats["lt_std"].mean()))

    # ③ 주문시배송리드타임 (실측 통계 자체가 없는 경우)
    if lt_order_lookup:
        lo_val = lt_order_lookup.get(prod_key)
        if lo_val and lo_val > 0:
            return _floor(float(lo_val), 0.0)

    # ④ 리드타임 파일 (std=0 가정)
    lt_file_val = get_lead_time(lt_df, vendor, prod_key)
    return _floor(float(lt_file_val), 0.0)


def get_ofr(lt_stats: pd.DataFrame, vendor: str, prod_key: str) -> float:
    """
    품목별 실측 OFR(주문충족률, %) 반환.
    실측 데이터 없으면 np.nan.
    협력사 단위로 fallback.
    """
    if lt_stats.empty or "ofr" not in lt_stats.columns:
        return np.nan
    row = lt_stats[(lt_stats["vendor"] == vendor) & (lt_stats["prod_key"] == prod_key)]
    if not row.empty and not np.isnan(row.iloc[0]["ofr"]):
        return float(row.iloc[0]["ofr"])
    row_v = lt_stats[lt_stats["vendor"] == vendor]
    if not row_v.empty:
        v = row_v["ofr"].dropna()
        if len(v) > 0:
            return float(v.mean())
    return np.nan

def calc_olfr(df: pd.DataFrame) -> pd.DataFrame:
    """
    Order Line Fill Rate (OLFR) 계산.
    주문 라인(주문번호+품목번호) 단위로 OFR_TARGET_DAYS 이내 납기 충족 비율.
    OFR(주문 단위)보다 정밀: 하나의 주문에서 일부 라인만 지연돼도 감지.

    반환: vendor × prod_key 별 olfr(%) DataFrame
    """
    if "_line_key" not in df.columns or "_lt" not in df.columns:
        return pd.DataFrame()

    lt_valid = df[df["_lt"].notna() & df["_line_key"].notna()].copy()
    if lt_valid.empty:
        return pd.DataFrame()

    # 라인 단위로 납기 충족 여부 판정 후 vendor×prod_key 집계
    lt_valid["_on_time"] = lt_valid["_lt"] <= OFR_TARGET_DAYS
    grp = (
        lt_valid.groupby(["_vendor", "_prod_key"])
        .agg(
            total_lines=("_line_key", "count"),
            on_time_lines=("_on_time", "sum"),
        )
        .reset_index()
    )
    grp["olfr"] = grp["on_time_lines"] / grp["total_lines"] * 100
    grp.rename(columns={"_vendor": "vendor", "_prod_key": "prod_key"}, inplace=True)
    return grp[["vendor", "prod_key", "olfr", "total_lines"]]


def get_olfr(olfr_df: pd.DataFrame, vendor: str, prod_key: str) -> float:
    """품목별 OLFR(%) 반환. 없으면 협력사 평균 → np.nan."""
    if olfr_df.empty or "olfr" not in olfr_df.columns:
        return np.nan
    row = olfr_df[(olfr_df["vendor"] == vendor) & (olfr_df["prod_key"] == prod_key)]
    if not row.empty:
        return float(row.iloc[0]["olfr"])
    row_v = olfr_df[olfr_df["vendor"] == vendor]
    if not row_v.empty:
        v = row_v["olfr"].dropna()
        if len(v) > 0:
            return float(v.mean())
    return np.nan


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


# ── 협력사 재고가이드: 단가 인하 구간 ─────────────────────────────────────
SUPPLIER_TIERS = [
    (10_000_000_000, 0.020, "Tier 4 (10억↑)"),
    ( 5_000_000_000, 0.015, "Tier 3 (5억~10억)"),
    ( 1_000_000_000, 0.010, "Tier 2 (1억~5억)"),
    (             0, 0.005, "Tier 1 (1억 미만)"),
]

def _get_tier(annual_amount: float) -> tuple[float, str]:
    for threshold, rate, label in SUPPLIER_TIERS:
        if annual_amount >= threshold:
            return rate, label
    return 0.005, "Tier 1"


def load_purchase_plan(base_dir: Path) -> pd.DataFrame:
    """
    구매계획_입력.xlsx 로드.
    컬럼: 협력사명, 상품코드, YYYY-MM(월별 구매예정수량 열 복수)
    """
    candidates = sorted(base_dir.glob("구매계획_입력*.xlsx"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        return pd.DataFrame()
    path = candidates[0]
    try:
        raw = pd.read_excel(path, engine="openpyxl")
        print(f"  [구매계획] {path.name} 로드: {len(raw):,}행")
        return raw
    except Exception as e:
        print(f"  [구매계획] 로드 실패 ({e})")
        return pd.DataFrame()


def make_purchase_plan_template(abc_df: pd.DataFrame, future_periods, out_path: Path):
    """담당자 구매 예정 수량 입력 템플릿 생성"""
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "구매계획입력"

    yellow = PatternFill("solid", fgColor="FFFACD")
    header_fill = PatternFill("solid", fgColor="4472C4")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    green = PatternFill("solid", fgColor="C8F0C8")  # 현재재고 입력란
    periods = [str(p) for p in future_periods]
    # 헤더: 고정 6열 + 현재재고(초록) + 월별 구매계획(노란)
    headers = [
        "협력사명", "상품코드", "상품명", "ABC등급",
        "권장SS\n(개/월)", "권장ROP\n(개)",
        "현재재고\n(개) ◀입력",
    ] + [f"구매예정\n{p}\n(개) ◀입력" for p in periods]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = header_fill
        cell.font = hdr_font
        cell.alignment = ctr
        cell.border = thin
    ws.row_dimensions[1].height = 40

    ab_rows = abc_df[abc_df["ABC"].isin(["A", "B"])].copy()
    for r, (_, row) in enumerate(ab_rows.iterrows(), 2):
        ws.cell(r, 1, str(row["_vendor"])).border = thin
        ws.cell(r, 2, str(row["_prod_key"])).border = thin
        ws.cell(r, 3, str(row.get("_prod_name", ""))).border = thin
        ws.cell(r, 4, str(row["ABC"])).border = thin
        ws.cell(r, 5, round(row.get("_ss",  0), 1)).border = thin
        ws.cell(r, 6, round(row.get("_rop", 0), 1)).border = thin
        # 현재재고 입력란 (초록)
        cur_cell = ws.cell(r, 7, "")
        cur_cell.fill = green
        cur_cell.border = thin
        cur_cell.alignment = ctr
        # 월별 구매계획 입력란 (노란)
        for c_off in range(len(periods)):
            cell = ws.cell(r, 8 + c_off, 0)
            cell.fill = yellow
            cell.border = thin
            cell.alignment = ctr

    col_widths = [24, 18, 30, 8, 12, 12, 14] + [13] * len(periods)
    for c, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    ws.freeze_panes = "H2"
    note_r = len(ab_rows) + 3
    notes = [
        "※ 초록색 셀 [현재재고]: 현재 협력사 보유 재고 수량(개)을 입력하면 결품위험·발주추천이 활성화됩니다.",
        "※ 노란색 셀 [구매예정]: 향후 월별 추가 구매 예정 수량(개)을 입력하세요. 0은 미입력으로 처리됩니다.",
        "※ 권장SS(안전재고)·ROP(재주문점) 단위는 개/월이며, 현재재고와 비교해 발주 시점을 판단하세요.",
    ]
    for ni, note_txt in enumerate(notes):
        nc = ws.cell(note_r + ni, 1, note_txt)
        nc.font = Font(italic=True, color="555555", size=9)
    wb.save(out_path)


def _build_alert_lookup(detail_rows: list, ss_rop: dict,
                        current_stock_lookup: dict | None = None) -> dict:
    """
    품목별 알림 딕셔너리 생성.
    current_stock_lookup: {(vendor, prod_key): 현재재고수량} — 협력사 입력값.
                          None이면 재고 기반 판단 생략.
    반환: {(vendor, prod_key): {"surge": bool, "surge_pct": float,
                                "stockout_risk": bool, "stockout_month": str,
                                "order_rec": str,
                                "fc6": list, "avg12": float}}
    """
    alerts = {}
    for d in detail_rows:
        vendor   = d["vendor"]
        prod_key = d["prod_key"]
        avg12    = float(d.get("avg12", 0.0) or 0.0)
        _fc_raw  = d.get("fc6")
        fc6 = list(_fc_raw) if _fc_raw is not None and len(_fc_raw) > 0 else []
        key    = (vendor, prod_key)
        sr     = ss_rop.get(key, {})
        ss_val = sr.get("ss", 0.0)
        rop_val= sr.get("rop", 0.0)
        lt_days= sr.get("lead_time", 30.0)

        # ── ① 수요 급등: 향후 3개월 최대값이 최근 12M 평균 대비 30%↑ ──
        surge = False
        surge_pct = 0.0
        if avg12 > 0 and len(fc6) > 0:
            peak = max(fc6[:3])
            surge_pct = (peak - avg12) / avg12 * 100
            surge = surge_pct >= 30.0

        # ── ② 결품위험·발주타이밍: 협력사 현재재고 입력 시에만 판단 ──
        stockout_risk  = False
        stockout_month = ""
        order_rec      = "현재재고 미입력"   # 협력사가 입력해야 판단 가능

        cur_stock = (current_stock_lookup or {}).get(key)
        if cur_stock is not None:
            cur_stock = float(cur_stock)
            # 현재재고가 ROP 이하면 즉시 발주 필요
            if cur_stock <= rop_val:
                order_rec = "🔴 즉시 발주 필요"
            elif cur_stock <= rop_val * 1.5:
                order_rec = "🟡 이번 달 발주 권장"
            else:
                order_rec = "✅ 당분간 발주 불필요"

            # 현재재고로 몇 개월 버티는지 계산
            monthly_demand = avg12 if avg12 > 0 else 1
            months_left = cur_stock / monthly_demand
            for i, fv in enumerate(fc6):
                remaining = cur_stock - sum(fc6[:i+1])
                if remaining < ss_val:
                    stockout_risk  = True
                    stockout_month = f"+{i+1}개월 후 SS 이하"
                    break

        alerts[key] = {
            "surge":          surge,
            "surge_pct":      surge_pct,
            "stockout_risk":  stockout_risk,
            "stockout_month": stockout_month,
            "order_rec":      order_rec,
            "fc6":            fc6,
            "avg12":          avg12,
        }
    return alerts


def write_sheet_supplier_alerts(wb, abc_df: pd.DataFrame, ss_rop: dict,
                                 detail_rows: list, future_periods,
                                 df_raw: pd.DataFrame | None = None,
                                 current_stock_lookup: dict | None = None):
    """
    협력사 알림 보드 시트.
    ① 수요 급등 품목 사전 알림 (재고 없이도 판단 가능)
    ② 결품 위험·발주 추천 (협력사 현재재고 입력 시 활성화)
    - 관리회계 컬럼 포함 (유가증권 등 직접 식별)
    - 모든 수량 단위: 개/월, 금액 단위: 원
    """
    from openpyxl.styles import PatternFill as PF, Font as FT, Alignment as AL
    ws = wb.create_sheet("협력사_알림보드")
    ws.tab_color = "375623"   # 녹색 — 핵심 시트
    sheet_title(ws, "협력사 공급망 알림 보드 (Supplier SCM Alert)",
                "※ 수량 단위: 개/월  |  금액 단위: 원  |  결품위험·발주추천은 협력사 현재재고 입력 후 활성화")

    alerts = _build_alert_lookup(detail_rows, ss_rop, current_stock_lookup)

    periods = [str(p) for p in future_periods]
    SR = 4
    headers = [
        "협력사명", "상품코드", "상품명", "관리회계\n(유가증권 식별)", "ABC등급",
        "📈 수요급등\n알림", "급등률\n(%)",
        "⚠️ 결품위험\n(재고입력 필요)", "위험 예상\n시점",
        "🛒 발주 추천\n(재고입력 필요)",
        "ML월평균수요\n(개/월, 최근12M)",
        "권장 안전재고\n[SS, 개]",
        "권장 재주문점\n[ROP, 개]",
        "평균 리드타임\n(일)",
        "연간구매금액\n(원)",
    ] + [f"수요예측\n{p}\n(개/월)" for p in periods]

    fill_surge    = PF("solid", fgColor="FFE0B2")
    fill_stockout = PF("solid", fgColor="FFCCCC")
    fill_order_r  = PF("solid", fgColor="FFCCCC")
    fill_order_y  = PF("solid", fgColor="FFF9C4")
    fill_no_input = PF("solid", fgColor="F0F0F0")

    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 48

    # 관리회계 룩업 — 품목별 비어있지 않은 값 중 최다 빈도값 사용
    # drop_duplicates(첫 행)은 그 행이 NaN이면 그대로 누락되므로 mode 방식으로 전환
    mgmt_lookup: dict = {}
    if df_raw is not None and "_mgmt_acct" in df_raw.columns and "_prod_key" in df_raw.columns:
        _ma = df_raw[["_prod_key", "_mgmt_acct"]].copy()
        _ma = _ma[~_ma["_mgmt_acct"].isin(["—", "nan", "None", ""])]
        if not _ma.empty:
            mgmt_lookup = (
                _ma.groupby("_prod_key")["_mgmt_acct"]
                .agg(lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else "—")
                .to_dict()
            )
    if not mgmt_lookup:
        mgmt_col = next((c for c in abc_df.columns if "관리회계" in str(c)), None)
        if mgmt_col and "_prod_key" in abc_df.columns:
            mgmt_lookup = abc_df.set_index("_prod_key")[mgmt_col].to_dict()

    ab_df = abc_df[abc_df["ABC"].isin(["A", "B"])].copy()

    # 우선순위 정렬: 결품위험 → 수요급등 → 나머지
    def _priority(row):
        key = (str(row["_vendor"]), str(row["_prod_key"]))
        a = alerts.get(key, {})
        if a.get("stockout_risk"): return 0
        if a.get("surge"):         return 1
        return 2

    ab_df["_priority"] = ab_df.apply(_priority, axis=1)
    ab_df = ab_df.sort_values("_priority")

    r = SR + 1
    for _, row in ab_df.iterrows():
        vendor   = str(row["_vendor"])
        prod_key = str(row["_prod_key"])
        abc      = row["ABC"]
        key      = (vendor, prod_key)
        a        = alerts.get(key, {})
        sr       = ss_rop.get(key, {})

        surge         = a.get("surge", False)
        surge_pct     = a.get("surge_pct", 0.0)
        stockout_risk = a.get("stockout_risk", False)
        stockout_month= a.get("stockout_month", "")
        order_rec     = a.get("order_rec", "현재재고 미입력")
        fc6           = a.get("fc6", [])
        avg12         = a.get("avg12", row.get("avg_qty", 0.0))
        annual_amt    = row.get("total_amount", 0)
        mgmt_val      = mgmt_lookup.get(prod_key, "—")

        row_bg = "FFEEEE" if stockout_risk else ("FFF5E6" if surge else None)

        col = 1
        dcell(ws.cell(r, col), vendor,   align="left", bg=row_bg); col += 1
        dcell(ws.cell(r, col), prod_key, align="left", bg=row_bg); col += 1
        dcell(ws.cell(r, col), str(row.get("_prod_name", "")), align="left", bg=row_bg); col += 1

        # 관리회계 (유가증권 강조)
        mgmt_str = str(mgmt_val) if mgmt_val not in (None, "—", "nan") else "—"
        is_gift = any(kw in mgmt_str for kw in ("유가증권", "상품권", "쿠폰"))
        # 관리회계 미입력이어도 상품명에 유가증권성 키워드가 있으면 추정 표시
        _GIFT_KWS = ("기프티쇼", "기프티콘", "상품권", "쿠폰", "바우처", "포인트권")
        prod_name_str = str(row.get("_prod_name", ""))
        if mgmt_str == "—" and any(kw in prod_name_str for kw in _GIFT_KWS):
            mgmt_str = "유가증권(상품명 추정)"
            is_gift  = True
        dcell(ws.cell(r, col), mgmt_str, align="left",
              bg="FFE0E0" if is_gift else row_bg,
              bold=is_gift); col += 1

        dcell(ws.cell(r, col), abc, align="center", bg=row_bg, bold=True); col += 1

        # 수요 급등
        surge_cell = ws.cell(r, col, "🔺 급등 주의" if surge else "—")
        surge_cell.alignment = AL(horizontal="center")
        if surge:
            surge_cell.fill = fill_surge
            surge_cell.font = FT(bold=True, color="E65100")
        col += 1
        dcell(ws.cell(r, col), f"+{surge_pct:.1f}%" if surge else "—",
              align="center", bg="FFE0B2" if surge else None); col += 1

        # 결품 위험 (재고 입력 여부에 따라 다른 표시)
        has_stock_input = current_stock_lookup and key in current_stock_lookup
        stock_cell = ws.cell(r, col,
            "⚠️ 결품 위험" if stockout_risk else ("— (재고 미입력)" if not has_stock_input else "—"))
        stock_cell.alignment = AL(horizontal="center")
        if stockout_risk:
            stock_cell.fill = fill_stockout
            stock_cell.font = FT(bold=True, color="C00000")
        elif not has_stock_input:
            stock_cell.fill = fill_no_input
            stock_cell.font = FT(color="999999", italic=True)
        col += 1
        dcell(ws.cell(r, col), stockout_month if stockout_month else "—",
              align="center", bg="FFCCCC" if stockout_risk else None); col += 1

        # 발주 추천
        ot_cell = ws.cell(r, col, order_rec)
        ot_cell.alignment = AL(horizontal="center")
        if "즉시" in order_rec:
            ot_cell.fill = fill_order_r
            ot_cell.font = FT(bold=True, color="C00000")
        elif "이번 달" in order_rec:
            ot_cell.fill = fill_order_y
            ot_cell.font = FT(bold=True, color="7B4F00")
        elif "불필요" in order_rec:
            ot_cell.fill = PF("solid", fgColor="E8F5E9")
        elif "미입력" in order_rec:
            ot_cell.fill = fill_no_input
            ot_cell.font = FT(color="999999", italic=True)
        col += 1

        dcell(ws.cell(r, col), round(avg12),              "0",   bg=row_bg); col += 1
        dcell(ws.cell(r, col), round(sr.get("ss",  0)),   "0",   bg=row_bg); col += 1
        dcell(ws.cell(r, col), round(sr.get("rop", 0)),   "0",   bg=row_bg); col += 1
        dcell(ws.cell(r, col), round(sr.get("lead_time", 30), 1), "0.0", bg=row_bg); col += 1
        dcell(ws.cell(r, col), annual_amt, NUM_FMT,        bg=row_bg); col += 1

        for i in range(len(periods)):
            fv = fc6[i] if i < len(fc6) else avg12
            fc_bg = "FFE0B2" if surge and i < 3 and avg12 > 0 and fv >= avg12 * 1.3 else row_bg
            dcell(ws.cell(r, col), round(fv), "0", bg=fc_bg); col += 1

        r += 1

    col_widths = [24, 16, 28, 18, 7, 13, 9, 18, 12, 18, 16, 14, 14, 12, 18]
    col_widths += [13] * len(periods)
    for c, w in enumerate(col_widths, 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    n_cols_alert = len(headers)
    from openpyxl.utils import get_column_letter as _gcl
    ws.auto_filter.ref = f"A{SR}:{_gcl(n_cols_alert)}{r-1}"

    # 범례
    note_r = r + 1
    legends = [
        ("🔺 수요급등: 향후 3개월 예측이 최근 12개월 평균 대비 30% 이상 증가 예상", "FFE0B2"),
        ("⚠️ 결품위험·🛒발주추천: 협력사가 현재재고를 입력해야 활성화됩니다", "F0F0F0"),
        ("🔴 유가증권(빨간셀): 관리회계 기준 상품권·유가증권 해당 품목", "FFE0E0"),
        ("SS=안전재고(개), ROP=재주문점(개): 협력사 권장 재고 수준 가이드", "E8F5E9"),
    ]
    from openpyxl.styles import PatternFill as _PF2, Font as _FT2
    for leg_i, (txt, clr) in enumerate(legends):
        col_start = leg_i * 4 + 1
        cell = ws.cell(note_r, col_start, txt)
        cell.fill = _PF2("solid", fgColor=clr)
        cell.font = _FT2(size=9, italic=True)
        ws.merge_cells(start_row=note_r, start_column=col_start,
                       end_row=note_r, end_column=col_start + 3)

    ws.sheet_view.showGridLines = False


def write_sheet_supplier_guide(wb, abc_df: pd.DataFrame, ss_rop: dict,
                                df_raw: pd.DataFrame, future_periods,
                                purchase_plan: pd.DataFrame):
    """
    협력사 재고 가이드 시트.
    - ML 예측 기반 안전재고·ROP
    - 담당자 구매 예정 수량 반영 시 보정 수요 표시
    - 연간 구매금액 기반 단계별 단가 인하율(Tier)
    """
    ws = wb.create_sheet("협력사_재고가이드")
    ws.tab_color = "375623"   # 녹색 — 핵심 시트
    sheet_title(ws, "협력사 재고 수준 가이드 (Safety Stock Recommendation)")

    periods = [str(p) for p in future_periods]
    n_periods = len(periods)

    SR = 4  # 헤더 행
    fixed_headers = [
        "협력사명", "상품코드", "상품명", "규격",
        "ABC등급", "연간구매금액(원)", "단가인하율(Tier)",
        "LT평균(일)", "안전재고(SS)", "ROP",
        "ML월평균수요(최근12M)",
    ]
    period_headers_ml   = [f"ML예측\n{p}" for p in periods]
    period_headers_plan = [f"구매계획\n{p}" for p in periods]
    period_headers_adj  = [f"보정수요\n{p}" for p in periods]
    headers = fixed_headers + period_headers_ml + period_headers_plan + period_headers_adj + ["비고"]

    # 헤더 색상 구분
    ml_fill   = PatternFill("solid", fgColor="DDEEFF") if 'PatternFill' in dir() else None
    plan_fill = PatternFill("solid", fgColor="E8F5E9") if 'PatternFill' in dir() else None
    adj_fill  = PatternFill("solid", fgColor="FFF9C4") if 'PatternFill' in dir() else None

    try:
        from openpyxl.styles import PatternFill as PF
        ml_fill   = PF("solid", fgColor="DDEEFF")
        plan_fill = PF("solid", fgColor="D5F5E3")
        adj_fill  = PF("solid", fgColor="FFF9C4")
    except Exception:
        ml_fill = plan_fill = adj_fill = None

    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
        wcell = ws.cell(SR, c)
        if ml_fill and len(fixed_headers) < c <= len(fixed_headers) + n_periods:
            wcell.fill = ml_fill
        elif plan_fill and len(fixed_headers) + n_periods < c <= len(fixed_headers) + 2 * n_periods:
            wcell.fill = plan_fill
        elif adj_fill and len(fixed_headers) + 2 * n_periods < c <= len(fixed_headers) + 3 * n_periods:
            wcell.fill = adj_fill
    ws.row_dimensions[SR].height = 36

    # 연간 구매금액 룩업 (협력사별)
    vendor_annual: dict = {}
    if "total_amount" in abc_df.columns and "_vendor" in abc_df.columns:
        vendor_annual = abc_df.groupby("_vendor")["total_amount"].sum().to_dict()

    # 구매계획 룩업: (협력사명, 상품코드) → {period: qty}
    plan_lookup: dict = {}
    if not purchase_plan.empty:
        # 컬럼 탐색: 협력사명·상품코드 외 나머지가 기간 컬럼
        v_col = next((c for c in purchase_plan.columns if "협력사" in str(c)), None)
        p_col = next((c for c in purchase_plan.columns if "상품코드" in str(c) or "품목" in str(c)), None)
        if v_col and p_col:
            date_cols = [c for c in purchase_plan.columns
                         if c not in (v_col, p_col) and str(c)[:4].isdigit()]
            for _, pr in purchase_plan.iterrows():
                key = (str(pr[v_col]).strip(), str(pr[p_col]).strip())
                plan_lookup[key] = {str(dc): float(pr[dc]) if pd.notna(pr[dc]) else 0.0
                                    for dc in date_cols}

    # 월별 ML 예측값 룩업: detail_rows 없이 group_grp 기반 평균으로 근사
    # abc_df에 _fc_next 컬럼이 있으면 사용, 없으면 최근12M 평균으로 대체
    grp_qty_lookup: dict = {}
    if "_channel" in abc_df.columns:
        grp_col = ["_channel", "_vendor", "_prod_key"]
    else:
        grp_col = ["_vendor", "_prod_key"]

    ab_df = abc_df[abc_df["ABC"].isin(["A", "B"])].copy()

    r = SR + 1
    for _, row in ab_df.iterrows():
        vendor   = str(row["_vendor"])
        prod_key = str(row["_prod_key"])
        prod_name = str(row.get("_prod_name", ""))
        prod_spec = str(row.get("_prod_spec", ""))
        abc = row["ABC"]
        bg = "FFF2CC" if abc == "A" else "E8F5E9"

        annual_amt = vendor_annual.get(vendor, 0)
        disc_rate, tier_label = _get_tier(annual_amt)

        d = ss_rop.get((vendor, prod_key), {})
        ss_val  = d.get("ss", 0)
        rop_val = d.get("rop", 0)
        lt_val  = d.get("lead_time", 30)

        # ML 월평균 수요 (최근 12M 근사)
        ml_avg = row.get("_avg_qty", row.get("avg_qty",
                  annual_amt / row.get("total_amount", annual_amt or 1) * 12
                  if annual_amt else 0))
        # total_amount 기반이면 금액이므로, 수량 컬럼 우선
        if "avg_qty" in row.index:
            ml_avg = float(row["avg_qty"])
        elif "total_qty" in row.index:
            ml_avg = float(row["total_qty"]) / max(row.get("months", 12), 1)
        else:
            ml_avg = 0.0

        # 구매계획 값
        plan_key = (vendor, prod_key)
        plan_vals = plan_lookup.get(plan_key, {})

        col = 1
        dcell(ws.cell(r, col), vendor,    align="left",   bg=bg); col += 1
        dcell(ws.cell(r, col), prod_key,  align="left",   bg=bg); col += 1
        dcell(ws.cell(r, col), prod_name, align="left",   bg=bg); col += 1
        dcell(ws.cell(r, col), prod_spec, align="left",   bg=bg); col += 1
        dcell(ws.cell(r, col), abc,       align="center", bg=bg, bold=True); col += 1
        dcell(ws.cell(r, col), annual_amt, NUM_FMT, bg=bg); col += 1
        dcell(ws.cell(r, col), f"{disc_rate:.1%}  ({tier_label})", align="center",
              bg="D5E8D4", bold=True); col += 1
        dcell(ws.cell(r, col), round(lt_val, 1), "0.0", bg=bg); col += 1
        dcell(ws.cell(r, col), round(ss_val, 1),  "0.0", bg=bg); col += 1
        dcell(ws.cell(r, col), round(rop_val, 1), "0.0", bg=bg); col += 1
        dcell(ws.cell(r, col), round(ml_avg, 1),  "0.0", bg=bg); col += 1

        # ML 예측 (단순히 ml_avg 반복 — 실제 예측값이 있으면 교체)
        for p in periods:
            dcell(ws.cell(r, col), round(ml_avg, 1), "0.0", bg="DDEEFF"); col += 1

        # 구매계획
        has_plan = bool(plan_vals)
        for p in periods:
            pv = plan_vals.get(p, 0.0)
            dcell(ws.cell(r, col), round(pv, 1), "0.0",
                  bg="D5F5E3" if pv > 0 else "F5F5F5"); col += 1

        # 보정 수요 = ML예측 + 구매계획
        for p in periods:
            pv = plan_vals.get(p, 0.0)
            adj = ml_avg + pv
            dcell(ws.cell(r, col), round(adj, 1), "0.0",
                  bg="FFF9C4", bold=(pv > 0)); col += 1

        note = "구매계획 반영" if has_plan else ""
        dcell(ws.cell(r, col), note, align="center", bg=bg)
        r += 1

    # 컬럼 너비
    col_widths = [24, 18, 28, 16, 7, 20, 20, 10, 10, 10, 14]
    col_widths += [12] * n_periods  # ML
    col_widths += [12] * n_periods  # 계획
    col_widths += [12] * n_periods  # 보정
    col_widths += [14]
    for c, w in enumerate(col_widths, 1):
        cw(ws, c, w)

    ws.freeze_panes = f"A{SR+1}"

    # 범례
    note_row = r + 1
    legend = [
        ("■ ML예측(파란색)", "DDEEFF"),
        ("■ 구매계획(초록색)", "D5F5E3"),
        ("■ 보정수요=ML+계획(노란색)", "FFF9C4"),
        ("■ Tier 단가인하율: 연간구매액 기준", "D5E8D4"),
    ]
    from openpyxl.styles import PatternFill as _PF, Font as _FT
    for ci, (txt, clr) in enumerate(legend, 1):
        cell = ws.cell(note_row, ci * 3 - 2, txt)
        cell.fill = _PF("solid", fgColor=clr)
        cell.font = _FT(size=9, italic=True)


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
                      lt_mean: float, lt_std: float = 0.0,
                      abc_grade: str = "B") -> tuple[float, str]:
    """
    리드타임 변동성을 반영한 완전한 안전재고 공식 (단위: 수량/월 기준).
    서비스수준 98% (Z=2.05) — MRO 가용성 기준.

    수요 변동만 있을 때:
        SS = Z × σ_D × √LT_months

    리드타임 변동성도 있을 때:
        SS = Z × √(LT_months × σ_D² + D_avg² × σ_LT_months²)

    - D_avg       : 최근 12개월 월평균 수요
    - σ_D         : 최근 12개월 수요 표준편차 (월)
    - LT_months   : 평균 리드타임 (일 → 월 환산: /30)
    - σ_LT_months : 리드타임 표준편차 (일 → 월 환산: /30)
    - abc_grade   : ABC등급 — C등급은 최소 1개 보유 룰 적용
    """
    window = qty_series[-12:] if len(qty_series) >= 12 else qty_series
    d_avg  = float(np.mean(window))
    sigma_d = float(np.std(window))

    lt_m     = lt_mean / 30          # 일 → 월
    sigma_lt = lt_std  / 30

    if sigma_lt > 0:
        ss = Z_SS * np.sqrt(lt_m * sigma_d**2 + d_avg**2 * sigma_lt**2)
        formula = "완전공식(수요+LT변동,SL98%)"
    else:
        ss = Z_SS * sigma_d * np.sqrt(lt_m)
        formula = "기본공식(수요변동만,SL98%)"

    # 수요 변동이 0(매달 동일 수량)이어도 리드타임 기간 중 최소 버퍼 확보
    # 최소 SS = 월평균 수요 × LT개월 × 50% (반달치 수요 수준)
    ss_min = d_avg * lt_m * 0.5
    if ss < ss_min and d_avg > 0:
        ss = ss_min
        formula += "+최소버퍼(σ=0보정)"

    # C등급 최소 1개 보유 룰 — MRO 특성상 단 1건의 결품도 생산 차질 유발
    if abc_grade == "C" and ss < 1.0:
        ss = 1.0
        formula += "+C등급최소1개"

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
    last_actual_amount = None
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
        last_actual_amount = row["amount"]
        r += 1

    # 예측값 행
    for i, period in enumerate(future_periods):
        bg = C_FCST
        warn_bg = C_WARN if model_mape > 30 else C_FCST
        dcell(ws.cell(r, 1), str(period),   align="center", bg=bg)
        dcell(ws.cell(r, 2), None, bg=bg)   # 실제 없음 — 차트 브릿지는 예측 컬럼으로 대체
        dcell(ws.cell(r, 3), None, bg=bg)
        # 첫 번째 예측 행: 예측 컬럼에 마지막 실적값을 시작점으로 기입 (차트 선 연결용)
        bridge_fc = last_actual_amount if (i == 0 and last_actual_amount is not None) else float(fc_vals[i])
        dcell(ws.cell(r, 4), bridge_fc, NUM_FMT, bg=bg)
        dcell(ws.cell(r, 5), float(ci_lo[i]),    NUM_FMT, bg=bg)
        dcell(ws.cell(r, 6), float(ci_hi[i]),    NUM_FMT, bg=bg)
        dcell(ws.cell(r, 7), best_model, align="center", bg=bg)
        mape_disp = f"{model_mape:.1f}%" if not np.isnan(model_mape) else "N/A"
        dcell(ws.cell(r, 8), mape_disp, align="center", bg=warn_bg)
        r += 1

    for c, w in enumerate([14,20,16,20,20,20,16,12], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    ws.auto_filter.ref = f"A{SR}:H{r-1}"

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
    ws.tab_color = "2E75B6"   # 파란색 — 검증 시트
    sheet_title(ws, "ABC × CV 분류표 (CPSM 공급망 분석)")
    SR = 4
    has_channel = "_channel" in abc_df.columns
    headers = (["채널"] if has_channel else []) + \
              ["협력사명", "상품코드", "상품명", "규격", "ABC등급", "CV값", "CV구분",
               "연간정산금액(원)", "누적기여율(%)", "예측전략",
               "LT평균(일)", "LT표준편차(일)", "안전재고공식",
               f"수요버퍼(수량)\n[SS, Z=1.65]",
               f"발주기준량(수량)\n[ROP]",
               f"OFR(%)\n[주문단위]",
               f"OLFR(%)\n[라인단위]"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 30

    ch_colors = {"KT": "DDEEFF", "그룹사": "DDF0DD", "외부사": "FFF0CC", "지입자재": "FFE0E0"}

    for r, (_, row) in enumerate(abc_df.iterrows(), SR+1):
        abc = row["ABC"]
        bg = C_A if abc == "A" else (C_B if abc == "B" else None)
        key = (row["_vendor"], row["_prod_key"])
        d       = ss_rop.get(key, {})
        ss_val  = d.get("ss", 0)
        rop_val = d.get("rop", 0)
        lt_val  = d.get("lead_time", DEFAULT_LEAD_TIME)
        lt_s    = d.get("lt_std", 0)
        formula = d.get("formula", "-")
        ofr_val = d.get("ofr", np.nan)

        col = 1
        if has_channel:
            ch = str(row.get("_channel", "기타"))
            ch_bg = ch_colors.get(ch, "F0F0F0") if bg is None else bg
            dcell(ws.cell(r, col), ch, align="center", bg=ch_bg, bold=True)
            col += 1

        dcell(ws.cell(r, col),   str(row["_vendor"]),   align="left", bg=bg); col += 1
        dcell(ws.cell(r, col),   str(row["_prod_key"]), align="left", bg=bg); col += 1
        # 상품명 / 규격
        _pi = ss_rop.get(key, {})   # prod_info는 write 시점에 없으므로 prod_key로 직접 조회 불가
        # → abc_df에 prod_name/spec 컬럼이 있으면 사용, 없으면 공백
        dcell(ws.cell(r, col), str(row.get("_prod_name", "")), align="left", bg=bg); col += 1
        dcell(ws.cell(r, col), str(row.get("_prod_spec", "")), align="left", bg=bg); col += 1
        dcell(ws.cell(r, col),   abc, align="center", bg=bg, bold=True);      col += 1
        dcell(ws.cell(r, col),   round(row["cv"], 3), "0.000", bg=bg);        col += 1
        dcell(ws.cell(r, col),   row["CV_class"], align="center", bg=bg);     col += 1
        dcell(ws.cell(r, col),   row["total_amount"], NUM_FMT, bg=bg);        col += 1
        dcell(ws.cell(r, col),   round(row["cum_pct"], 2), "0.00%", bg=bg);   col += 1
        dcell(ws.cell(r, col),   row["strategy"], align="center", bg=bg);     col += 1
        lt_bg = "FFF9C4" if abs(lt_val - DEFAULT_LEAD_TIME) < 0.1 and lt_s == 0 else bg
        dcell(ws.cell(r, col),   round(lt_val, 1), "0.0", bg=lt_bg);              col += 1
        dcell(ws.cell(r, col),   round(lt_s, 1),   "0.0",
              bg="E8F5E9" if lt_s > 0 else bg);                                   col += 1
        dcell(ws.cell(r, col),   formula, align="center",
              bg="E8F5E9" if "완전" in formula else bg);                           col += 1
        dcell(ws.cell(r, col),   round(ss_val, 1),  "0.0", bg=bg);               col += 1
        dcell(ws.cell(r, col),   round(rop_val, 1), "0.0", bg=bg);               col += 1
        # OFR 컬럼 (주문 단위) — 낮을수록 빨간 배경
        ofr_val = d.get("ofr", np.nan)
        if not (isinstance(ofr_val, float) and np.isnan(ofr_val)):
            ofr_bg = ("FFCCCC" if ofr_val < 80 else ("FFF2CC" if ofr_val < 95 else "C6EFCE"))
            dcell(ws.cell(r, col), f"{ofr_val:.1f}%", align="center", bg=ofr_bg, bold=(ofr_val < 80))
        else:
            dcell(ws.cell(r, col), "-", align="center", bg=bg)
        col += 1
        # OLFR 컬럼 (라인 단위) — 동일 색상 기준
        olfr_val = d.get("olfr", np.nan)
        if not (isinstance(olfr_val, float) and np.isnan(olfr_val)):
            olfr_bg = ("FFCCCC" if olfr_val < 80 else ("FFF2CC" if olfr_val < 95 else "C6EFCE"))
            dcell(ws.cell(r, col), f"{olfr_val:.1f}%", align="center", bg=olfr_bg, bold=(olfr_val < 80))
        else:
            dcell(ws.cell(r, col), "-", align="center", bg=bg)

    ch_w = [10] if has_channel else []
    n_cols = len(ch_w) + 17
    for c, w in enumerate(ch_w + [22, 22, 30, 18, 8, 8, 10, 18, 12, 12, 10, 10, 16, 12, 12, 10, 10], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A{SR}:{get_column_letter(n_cols)}{r-1}"


def write_sheet3_detail(wb, detail_rows: list):
    ws = wb.create_sheet("협력사별_상품코드_예측")
    ws.tab_color = "2E75B6"   # 파란색 — 검증 시트
    sheet_title(ws, "협력사 × 상품코드별 향후 6개월 예측 (A등급 전체 + B등급 상위 20개)")
    SR = 4
    headers = ["협력사명", "상품코드", "상품명", "규격", "ABC등급",
               "최근12개월평균(개)",
               "예측+1M(개)", "예측+2M(개)", "예측+3M(개)",
               "예측+4M(개)", "예측+5M(개)", "예측+6M(개)",
               "사용모델", "MAPE(%)", "신뢰도",
               "수명주기", "EOL예상", "하락률(%)"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 30

    # 수명주기 단계별 배경색
    _LC_COLOR = {"단종임박": "FF4444", "쇠퇴기": "FFB347", "성숙기": None, "성장기": "E2EFDA"}

    for r, dr in enumerate(detail_rows, SR+1):
        abc = dr["ABC"]
        mape_val = dr.get("mape", np.nan)
        bg_base = C_A if abc == "A" else (C_B if abc == "B" else None)
        bg_warn = C_WARN if (not np.isnan(mape_val) and mape_val > 30) else bg_base
        lc_stage = dr.get("lifecycle_stage", "")
        bg_lc = _LC_COLOR.get(lc_stage, None)

        dcell(ws.cell(r, 1), dr["vendor"],              align="left", bg=bg_base)
        dcell(ws.cell(r, 2), dr["prod_key"],            align="left", bg=bg_base)
        dcell(ws.cell(r, 3), dr.get("prod_name", ""),  align="left", bg=bg_base)
        dcell(ws.cell(r, 4), dr.get("prod_spec", ""),  align="left", bg=bg_base)
        dcell(ws.cell(r, 5), abc, align="center", bold=True, bg=bg_base)
        dcell(ws.cell(r, 6), round(dr["avg12"]), "0", bg=bg_base)
        for i, fc_v in enumerate(dr["fc6"], 1):
            dcell(ws.cell(r, 6+i), round(float(fc_v)), "0", bg=C_FCST)
        dcell(ws.cell(r, 13), dr["model"], align="center", bg=bg_base)
        mape_str = f"{mape_val:.1f}%" if not np.isnan(mape_val) else "N/A"
        dcell(ws.cell(r, 14), mape_str, align="center", bg=bg_warn)
        flag = "⚠ 낮음" if (not np.isnan(mape_val) and mape_val > 30) else "✓ 양호"
        dcell(ws.cell(r, 15), flag, align="center", bg=bg_warn)
        dcell(ws.cell(r, 16), lc_stage, align="center", bold=(lc_stage in ("쇠퇴기", "단종임박")), bg=bg_lc)
        dcell(ws.cell(r, 17), dr.get("eol_ym", ""), align="center", bg=bg_lc)
        dcell(ws.cell(r, 18), dr.get("decline_pct", ""), "0.0", align="center", bg=bg_lc)

    for c, w in enumerate([22, 22, 32, 18, 8, 18, 14, 14, 14, 14, 14, 14, 14, 12, 12, 12, 12, 10], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    ws.auto_filter.ref = f"A{SR}:R{r-1}"


def write_sheet_forecast_detail(wb, detail_rows: list, ss_rop: dict):
    """
    Sheet: 예측근거_상품분석
    상품별 예측값, 모델, 정확도, 안전재고 종합 분석
    """
    ws = wb.create_sheet("예측근거_상품분석")
    ws.tab_color = "2E75B6"   # 파란색 — 검증 시트
    sheet_title(ws, "상품별 예측 분석 근거",
                "A등급 전체 + B등급 상위 20개 품목의 예측 모델 및 안전재고 정보")

    # 헤더 설정 (21개 컬럼 — 상품명·규격·OFR 추가)
    SR = 3
    headers = [
        "협력사", "상품코드", "상품명", "규격",
        "ABC", "통신/일반", "최근12개월\n월평균(원)",
        "6개월누적예측(원)",
        "예측+1M", "예측+2M", "예측+3M", "예측+4M", "예측+5M", "예측+6M",
        "예측모델", "MAPE(%)", "신뢰도",
        f"수요버퍼(개)\n[SS]", f"발주기준량(개)\n[ROP]", "LT평균(일)",
        f"OFR(%)\n[{OFR_TARGET_DAYS}일이내]",
    ]

    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, bg=C_HEADER)
    ws.row_dimensions[SR].height = 28

    # 데이터 행
    r = SR + 1
    for dr in detail_rows:
        vendor = dr["vendor"]
        prod_key = dr["prod_key"]

        # ss_rop에서 해당 상품의 값 조회
        ss_rop_key = (vendor, prod_key)
        ss_rop_info = ss_rop.get(ss_rop_key, {})

        # ABC 등급별 배경색
        abc = dr["ABC"]
        abc_bg = C_A if abc == "A" else (C_B if abc == "B" else C_C)

        # MAPE 기반 신뢰도 플래그 (3단계)
        mape_val = dr.get("mape", np.nan)
        if np.isnan(mape_val):
            trust_flag = "불확정"
            trust_color = "F0F0F0"
        elif mape_val <= 15:
            trust_flag = "●"  # 높음
            trust_color = "BDD7EE"
        elif mape_val <= 30:
            trust_flag = "◐"  # 중간
            trust_color = "FFE699"
        else:
            trust_flag = "◯"  # 낮음
            trust_color = "F8CBAD"

        # col 1: 협력사
        dcell(ws.cell(r, 1), vendor)
        # col 2: 상품코드
        dcell(ws.cell(r, 2), prod_key)
        # col 3: 상품명
        dcell(ws.cell(r, 3), dr.get("prod_name", ""), align="left")
        # col 4: 규격
        dcell(ws.cell(r, 4), dr.get("prod_spec", ""), align="left")
        # col 5: ABC
        dcell(ws.cell(r, 5), abc, align="center", bold=True, bg=abc_bg)
        # col 6: 통신/일반
        prod_type = dr.get("prod_type", "미분류")
        pt_bg = "D6E4F7" if prod_type == "통신" else ("E8F5E9" if prod_type == "일반" else "F5F5F5")
        dcell(ws.cell(r, 6), prod_type, align="center", bold=True, bg=pt_bg)
        # col 7: 최근12개월 평균
        dcell(ws.cell(r, 7), round(dr["avg12"]), "0")
        # col 8: 6개월 누적
        fc6_sum = float(np.sum(dr["fc6"]))
        dcell(ws.cell(r, 8), round(fc6_sum), "0", bold=True, bg=C_FCST)
        # col 9~14: 6개월 개별 예측값
        for i, fc_v in enumerate(dr["fc6"], 9):
            dcell(ws.cell(r, i), round(float(fc_v)), "0", bg=C_FCST)
        # col 15: 예측모델
        dcell(ws.cell(r, 15), dr["model"], align="center")
        # col 16: MAPE
        mape_str = f"{mape_val:.1f}" if not np.isnan(mape_val) else "N/A"
        dcell(ws.cell(r, 16), mape_str, align="center")
        # col 17: 신뢰도
        dcell(ws.cell(r, 17), trust_flag, align="center", bold=True, bg=trust_color)
        # col 18: 수요버퍼(SS)
        ss_val = ss_rop_info.get("ss", 0)
        dcell(ws.cell(r, 18), round(ss_val) if ss_val > 0 else "", NUM_FMT)
        # col 19: 발주기준량(ROP)
        rop_val = ss_rop_info.get("rop", 0)
        dcell(ws.cell(r, 19), round(rop_val) if rop_val > 0 else "", NUM_FMT)
        # col 20: LT평균
        lt_val = ss_rop_info.get("lead_time", DEFAULT_LEAD_TIME)
        dcell(ws.cell(r, 20), f"{lt_val:.0f}", align="center")
        # col 21: OFR
        ofr_val = ss_rop_info.get("ofr", np.nan)
        if not np.isnan(ofr_val):
            ofr_bg = ("FFCCCC" if ofr_val < 80 else
                      ("FFF2CC" if ofr_val < 95 else "C6EFCE"))
            dcell(ws.cell(r, 21), f"{ofr_val:.1f}%", align="center",
                  bg=ofr_bg, bold=(ofr_val < 80))
        else:
            dcell(ws.cell(r, 21), "-", align="center")

        r += 1

    # 컬럼 너비 설정 (21개 컬럼)
    widths = [16, 20, 32, 18, 8, 10, 18, 18, 12, 12, 12, 12, 12, 12, 16, 12, 10, 12, 12, 10, 10]
    for c, w in enumerate(widths, 1):
        cw(ws, c, w)

    # 고정행 및 포맷팅
    ws.freeze_panes = f"A{SR+1}"
    ws.auto_filter.ref = f"A{SR}:U{r-1}"
    ws.sheet_view.showGridLines = False
    # drop-ship 안내 주석
    note = ws.cell(r + 1, 1,
        f"※ [케이티커머스 drop-ship 기준] 수요버퍼(SS)·발주기준량(ROP)은 선매입 재고가 아닌 "
        f"리드타임 내 수요 변동 완충 목적의 발주 참고값입니다.")
    note.font = Font(name=FONT_NAME, size=9, italic=True, color="595959")
    note2 = ws.cell(r + 2, 1,
        f"※ OFR = 실제납기 {OFR_TARGET_DAYS}일 이내 이행 비율 (빨강<80% / 노랑<95% / 초록≥95%)")
    note2.font = Font(name=FONT_NAME, size=9, italic=True, color="595959")


def write_sheet_tracking_signal(wb, detail_rows: list):
    """
    Tracking Signal 분석 시트 (CPSM Module 2 Ch.3 기준)
    TS > +4 : 지속 과소예측 → 재고부족 위험
    TS < -4 : 지속 과대예측 → 재고과잉 위험
    """
    ws = wb.create_sheet("예측편향_TrackingSignal")
    ws.tab_color = "2E75B6"   # 파란색 — 검증 시트
    sheet_title(ws, "Tracking Signal 분석 — 예측 편향 감지",
                "CPSM 기준: |TS| > 4 이면 예측 편향 경보 | TS > 0: 과소예측(재고부족) | TS < 0: 과대예측(재고과잉)")

    SR = 3
    headers = ["협력사", "상품코드", "상품명", "통신/일반", "ABC",
               "Tracking Signal", "판정", "MAPE(%)", "예측모델",
               "최근12개월평균(원)", "6개월누적예측(원)"]
    h_colors = {"Tracking Signal": "1F4E79", "판정": "1F4E79"}
    for c, h in enumerate(headers, 1):
        bg = h_colors.get(h, C_HEADER)
        hcell(ws.cell(SR, c), h, bg=bg)
    ws.row_dimensions[SR].height = 28

    # Tracking Signal 값으로 정렬 (절댓값 큰 순 — 위험 품목 상단)
    rows_sorted = sorted(
        [dr for dr in detail_rows if not np.isnan(dr.get("tracking_signal", np.nan))],
        key=lambda d: abs(d.get("tracking_signal", 0)), reverse=True
    ) + [dr for dr in detail_rows if np.isnan(dr.get("tracking_signal", np.nan))]

    r = SR + 1
    for dr in rows_sorted:
        ts = dr.get("tracking_signal", np.nan)
        mape = dr.get("mape", np.nan)
        abc  = dr["ABC"]

        # 판정 및 색상
        if np.isnan(ts):
            verdict, ts_bg, row_bg = "데이터부족", "F0F0F0", None
        elif ts > 4:
            verdict, ts_bg, row_bg = "⚠ 과소예측(재고부족)", "FF9999", "FFF0F0"
        elif ts < -4:
            verdict, ts_bg, row_bg = "⚠ 과대예측(재고과잉)", "FFD966", "FFFBEA"
        else:
            verdict, ts_bg, row_bg = "✓ 정상", "C6EFCE", None

        abc_bg = C_A if abc == "A" else (C_B if abc == "B" else "F0F0F0")
        pt     = dr.get("prod_type", "미분류")
        pt_bg  = "D6E4F7" if pt == "통신" else ("E8F5E9" if pt == "일반" else "F5F5F5")

        if row_bg:
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).fill = PatternFill("solid", start_color=row_bg)

        dcell(ws.cell(r, 1), dr["vendor"])
        dcell(ws.cell(r, 2), dr["prod_key"])
        dcell(ws.cell(r, 3), dr.get("prod_name", ""), align="left")
        dcell(ws.cell(r, 4), pt, align="center", bold=True, bg=pt_bg)
        dcell(ws.cell(r, 5), abc, align="center", bold=True, bg=abc_bg)
        ts_str = f"{ts:+.2f}" if not np.isnan(ts) else "N/A"
        dcell(ws.cell(r, 6), ts_str, align="center", bold=True, bg=ts_bg)
        dcell(ws.cell(r, 7), verdict, align="center", bg=ts_bg)
        mape_str = f"{mape:.1f}%" if not np.isnan(mape) else "N/A"
        dcell(ws.cell(r, 8), mape_str, align="center")
        dcell(ws.cell(r, 9), dr["model"], align="center")
        dcell(ws.cell(r, 10), round(dr["avg12"]), "0")
        dcell(ws.cell(r, 11), round(float(np.sum(dr["fc6"]))), "0")
        r += 1

    # 범례
    ws.cell(r + 1, 1, "【Tracking Signal 해석】")
    ws.cell(r + 1, 1).font = Font(name=FONT_NAME, bold=True, size=9)
    ws.cell(r + 2, 1, "TS > +4 : 과소예측 지속 — 실제수요가 예측보다 높음 → 발주량/안전재고 상향 검토")
    ws.cell(r + 3, 1, "TS < -4 : 과대예측 지속 — 실제수요가 예측보다 낮음 → 발주량 축소, 재고 소진 검토")
    ws.cell(r + 4, 1, "-4 ≤ TS ≤ +4 : 정상 범위 — 예측 편향 없음")
    for ri in range(r + 2, r + 5):
        ws.cell(ri, 1).font = Font(name=FONT_NAME, size=9, color="595959")

    # 통계 요약
    all_ts = [dr.get("tracking_signal", np.nan) for dr in rows_sorted]
    valid_ts = [v for v in all_ts if not np.isnan(v)]
    n_under  = sum(1 for v in valid_ts if v > 4)
    n_over   = sum(1 for v in valid_ts if v < -4)
    n_normal = sum(1 for v in valid_ts if -4 <= v <= 4)
    ws.cell(r + 6, 1, f"전체 {len(valid_ts)}개 품목 분석 결과: "
                      f"과소예측 {n_under}개 / 과대예측 {n_over}개 / 정상 {n_normal}개")
    ws.cell(r + 6, 1).font = Font(name=FONT_NAME, bold=True, size=10)

    for c, w in enumerate([20, 22, 32, 10, 8, 16, 22, 12, 16, 20, 20], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR + 1}"
    ws.auto_filter.ref = f"A{SR}:K{SR + len(rows_sorted)}"
    ws.sheet_view.showGridLines = False


def write_sheet_supplier_risk(wb, df: pd.DataFrame, abc_df: pd.DataFrame):
    """
    공급업체 Risk Score 시트 (CPSM Module 1 Ch.1 Risk Analysis 기반)
    위험 = 납기준수율 × 가격변동성 × 공급집중도 × 품목중요도(ABC)
    """
    ws = wb.create_sheet("공급업체_리스크")
    ws.tab_color = "2E75B6"   # 파란색 — 검증 시트
    sheet_title(ws, "공급업체 리스크 스코어링",
                "CPSM Module 1 Risk Analysis — 납기·가격·집중도·품목중요도 종합 평가")

    # ── 납기준수율 계산 (주문일 → 입고일 실제 vs 기준 LT) ─────────────────
    has_lt_cols = "_lt" in df.columns and "_date" in df.columns
    if has_lt_cols:
        lt_df_calc = (
            df[df["_lt"].notna() & (df["_lt"] > 0)]
            .groupby("_vendor")["_lt"]
            .agg(lt_mean="mean", lt_std="std", lt_n="count")
            .reset_index()
        )
    else:
        lt_df_calc = pd.DataFrame(columns=["_vendor", "lt_mean", "lt_std", "lt_n"])

    # ── 공급업체별 금액 집계 (최근 12개월) ────────────────────────────────
    if "_ym" in df.columns:
        max_ym = df["_ym"].max()
        recent_mask = df["_ym"] >= (max_ym - 11)
        recent_df = df[recent_mask].copy()
    else:
        recent_df = df.copy()

    vendor_spend = (
        recent_df.groupby("_vendor")["_amount"]
        .agg(total_amount="sum", order_count="count")
        .reset_index()
    )
    total_spend = vendor_spend["total_amount"].sum()
    if total_spend > 0:
        vendor_spend["spend_share"] = vendor_spend["total_amount"] / total_spend * 100
    else:
        vendor_spend["spend_share"] = 0

    # ── 가격변동성 (금액 CV) ──────────────────────────────────────────────
    if "_ym" in df.columns:
        price_cv = (
            recent_df.groupby(["_vendor", "_ym"])["_amount"].sum()
            .reset_index()
            .groupby("_vendor")["_amount"]
            .agg(price_mean="mean", price_std="std")
            .reset_index()
        )
        price_cv["price_cv"] = np.where(
            price_cv["price_mean"] > 0,
            price_cv["price_std"] / price_cv["price_mean"] * 100, 0
        )
        price_cv["price_cv"] = price_cv["price_cv"].fillna(0)
    else:
        price_cv = pd.DataFrame(columns=["_vendor", "price_cv"])

    # ── ABC 등급 분포 (공급업체별 A등급 품목 수) ──────────────────────────
    a_count = (
        abc_df[abc_df["ABC"] == "A"]
        .groupby("_vendor")
        .size().reset_index(name="a_item_count")
    )

    # ── 데이터 통합 ────────────────────────────────────────────────────────
    risk_df = vendor_spend.copy()
    if not lt_df_calc.empty:
        risk_df = risk_df.merge(lt_df_calc, on="_vendor", how="left")
    else:
        risk_df["lt_mean"] = np.nan
        risk_df["lt_std"]  = np.nan
        risk_df["lt_n"]    = 0

    if not price_cv.empty:
        risk_df = risk_df.merge(price_cv[["_vendor", "price_cv"]], on="_vendor", how="left")
    else:
        risk_df["price_cv"] = 0

    risk_df = risk_df.merge(a_count, on="_vendor", how="left")
    risk_df["a_item_count"] = risk_df["a_item_count"].fillna(0)
    risk_df["price_cv"]     = risk_df["price_cv"].fillna(0)
    risk_df["lt_std"]       = risk_df["lt_std"].fillna(0)

    # ── Risk Score 계산 (100점 만점) ──────────────────────────────────────
    # 납기안정성 점수 (낮을수록 위험)
    max_lt_std = risk_df["lt_std"].max() or 1
    risk_df["score_lt"]    = (1 - risk_df["lt_std"] / max_lt_std) * 30

    # 가격안정성 점수 (변동성 낮을수록 안전)
    max_cv = risk_df["price_cv"].max() or 1
    risk_df["score_price"] = (1 - risk_df["price_cv"] / max_cv) * 25

    # 집중도 점수 (의존도 높을수록 위험)
    risk_df["score_conc"]  = (1 - risk_df["spend_share"] / 100) * 25

    # 품목중요도 점수 (A등급 품목 적을수록 위험 낮음)
    max_a = risk_df["a_item_count"].max() or 1
    risk_df["score_item"]  = (1 - risk_df["a_item_count"] / max_a) * 20

    risk_df["risk_score"]  = (risk_df["score_lt"] + risk_df["score_price"]
                               + risk_df["score_conc"] + risk_df["score_item"])
    # 위험도 = 100 - 안전점수
    risk_df["danger_score"] = 100 - risk_df["risk_score"]
    risk_df = risk_df.sort_values("danger_score", ascending=False)

    # ── 시트 작성 ─────────────────────────────────────────────────────────
    SR = 3
    headers = ["협력사명", "위험등급", "위험점수\n(100점)", "최근12개월\n구매금액(원)",
               "지출비중(%)", "A등급품목수", "납기변동성\n(σ일)",
               "가격변동성\n(CV%)", "납기안정성\n(30점)", "가격안정성\n(25점)",
               "집중도\n(25점)", "품목중요도\n(20점)"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 36

    r = SR + 1
    for _, row in risk_df.iterrows():
        ds = row["danger_score"]
        # 위험등급 판정
        if ds >= 70:
            grade, grade_bg, row_bg = "🔴 고위험", "FF9999", "FFF0F0"
        elif ds >= 45:
            grade, grade_bg, row_bg = "🟡 중위험", "FFD966", "FFFBEA"
        else:
            grade, grade_bg, row_bg = "🟢 저위험", "C6EFCE", None

        if row_bg:
            for c in range(1, len(headers) + 1):
                ws.cell(r, c).fill = PatternFill("solid", start_color=row_bg)

        dcell(ws.cell(r, 1),  row["_vendor"])
        dcell(ws.cell(r, 2),  grade, align="center", bold=True, bg=grade_bg)
        dcell(ws.cell(r, 3),  round(ds, 1), align="center", bold=True, bg=grade_bg)
        dcell(ws.cell(r, 4),  round(row["total_amount"]), NUM_FMT)
        dcell(ws.cell(r, 5),  f"{row['spend_share']:.1f}%", align="center")
        dcell(ws.cell(r, 6),  int(row["a_item_count"]), align="center")
        lt_s = row["lt_std"]
        dcell(ws.cell(r, 7),  f"{lt_s:.1f}" if lt_s > 0 else "데이터없음", align="center")
        dcell(ws.cell(r, 8),  f"{row['price_cv']:.1f}%", align="center")
        dcell(ws.cell(r, 9),  round(row["score_lt"], 1), align="center")
        dcell(ws.cell(r, 10), round(row["score_price"], 1), align="center")
        dcell(ws.cell(r, 11), round(row["score_conc"], 1), align="center")
        dcell(ws.cell(r, 12), round(row["score_item"], 1), align="center")
        r += 1

    # 범례
    ws.cell(r + 1, 1, "【위험 점수 구성】 납기안정성(30) + 가격안정성(25) + 집중도(25) + 품목중요도(20) = 100점 | 위험점수 = 100 - 안전점수")
    ws.cell(r + 1, 1).font = Font(name=FONT_NAME, size=9, color="595959")

    n_high = (risk_df["danger_score"] >= 70).sum()
    n_mid  = ((risk_df["danger_score"] >= 45) & (risk_df["danger_score"] < 70)).sum()
    ws.cell(r + 2, 1, f"전체 {len(risk_df)}개 협력사 — 🔴고위험 {n_high}개 / 🟡중위험 {n_mid}개 / 🟢저위험 {len(risk_df)-n_high-n_mid}개")
    ws.cell(r + 2, 1).font = Font(name=FONT_NAME, bold=True, size=10)

    for c, w in enumerate([22, 14, 12, 18, 12, 12, 14, 12, 12, 12, 10, 12], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR + 1}"
    ws.auto_filter.ref = f"A{SR}:L{SR + len(risk_df)}"
    ws.sheet_view.showGridLines = False


def _extract_cat_levels(series: pd.Series):
    """
    '가구 > 실내가구 > 사무용가구 > 거울' 형태의 카테고리를
    대분류 / 중분류 / 소분류 / 세분류로 분리
    """
    split = series.astype(str).str.split(">")
    def _get(lst, i):
        try:
            v = lst[i].strip()
            return v if v not in ("nan", "None", "") else "미분류"
        except (IndexError, AttributeError):
            return "미분류"
    return (
        split.apply(lambda x: _get(x, 0)),   # 대분류
        split.apply(lambda x: _get(x, 1)),   # 중분류
        split.apply(lambda x: _get(x, 2)),   # 소분류
        split.apply(lambda x: _get(x, 3)),   # 세분류
    )


def _spend_pareto_section(ws, r, recent, group_col, section_title,
                          col_label, prod_type_col):
    """
    금액 내림차순 정렬 + Pareto(ABC) 자동 표시 공통 로직
    """
    ws.cell(r, 1, section_title)
    ws.cell(r, 1).font = Font(name=FONT_NAME, bold=True, size=11, color="1F4E79")
    ws.row_dimensions[r].height = 22
    r += 1

    headers = [col_label, "구매금액(원)", "비중(%)", "누적비중(%)", "ABC", "통신/일반", "건수"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(r, c), h)
    r += 1

    cat_spend = (
        recent.groupby(group_col)
        .agg(total=("_amount", "sum"), cnt=("_amount", "count"))
        .reset_index()
        .sort_values("total", ascending=False)
    )

    if prod_type_col:
        cat_type = (
            recent.groupby(group_col)[prod_type_col]
            .agg(lambda x: x.value_counts().index[0] if len(x) > 0 else "미분류")
            .reset_index().rename(columns={prod_type_col: "prod_type"})
        )
        cat_spend = cat_spend.merge(cat_type, on=group_col, how="left")
    else:
        cat_spend["prod_type"] = "미분류"

    grand = cat_spend["total"].sum()
    cat_spend["pct"]     = cat_spend["total"] / grand * 100 if grand > 0 else 0
    cat_spend["cum_pct"] = cat_spend["pct"].cumsum()
    cat_spend["abc"]     = cat_spend["cum_pct"].apply(
        lambda x: "A" if x <= 80 else ("B" if x <= 95 else "C")
    )

    for _, row in cat_spend.iterrows():
        abc   = row["abc"]
        abc_bg = C_A if abc == "A" else (C_B if abc == "B" else "F0F0F0")
        pt    = row.get("prod_type", "미분류")
        pt_bg = "D6E4F7" if pt == "통신" else ("E8F5E9" if pt == "일반" else "F5F5F5")
        dcell(ws.cell(r, 1), row[group_col])
        dcell(ws.cell(r, 2), round(row["total"]), NUM_FMT, bold=(abc == "A"))
        dcell(ws.cell(r, 3), f"{row['pct']:.1f}%",     align="center")
        dcell(ws.cell(r, 4), f"{row['cum_pct']:.1f}%", align="center")
        dcell(ws.cell(r, 5), abc, align="center", bold=True, bg=abc_bg)
        dcell(ws.cell(r, 6), pt,  align="center", bg=pt_bg)
        dcell(ws.cell(r, 7), int(row["cnt"]), align="center")
        r += 1

    dcell(ws.cell(r, 1), "합계", bold=True)
    dcell(ws.cell(r, 2), round(grand), NUM_FMT, bold=True, bg=C_TOTAL)
    dcell(ws.cell(r, 7), int(cat_spend["cnt"].sum()), align="center", bold=True)
    return r + 2


def write_sheet_spend_analysis(wb, df: pd.DataFrame, current_ym):
    """
    카테고리별 Spend 분석 시트 (CPSM Module 1 Ch.2 Category Management 기반)
    대분류 / 중분류 / 소분류 3단계 Pareto + 통신·일반×채널 매트릭스
    """
    ws = wb.create_sheet("카테고리_Spend분석")
    sheet_title(ws, "카테고리별 지출(Spend) 분석",
                "CPSM Module 1 Category Management — 대/중/소분류 Pareto × 통신·일반·채널 매트릭스")

    # ── 최근 12개월 필터 ─────────────────────────────────────────────────
    if "_ym" in df.columns:
        max_ym = df["_ym"].max()
        recent = df[df["_ym"] >= (max_ym - 11)].copy()
    else:
        recent = df.copy()

    # ── 카테고리 컬럼 탐지 ───────────────────────────────────────────────
    cat_col = next((c for c in df.columns
                    if "서비스카테고리" in c or ("카테고리" in c and "_" not in c)), None)
    if cat_col is None:
        ws.cell(4, 1, "서비스카테고리 컬럼을 찾을 수 없습니다.")
        return

    # 대/중/소/세분류 추출
    recent["_cat_L1"], recent["_cat_L2"], recent["_cat_L3"], recent["_cat_L4"] = \
        _extract_cat_levels(recent[cat_col])

    prod_type_col = "_prod_type" if "_prod_type" in recent.columns else None

    # ── Section 1 · 2 · 3: 대/중/소분류 Pareto ─────────────────────────
    SR = 3
    r = SR
    r = _spend_pareto_section(ws, r, recent, "_cat_L1",
                               "▶ 1. 서비스 대분류별 지출 현황 (최근 12개월)",
                               "서비스 대분류", prod_type_col)
    r = _spend_pareto_section(ws, r, recent, "_cat_L2",
                               "▶ 2. 서비스 중분류별 지출 현황 (최근 12개월)",
                               "서비스 중분류", prod_type_col)
    r = _spend_pareto_section(ws, r, recent, "_cat_L3",
                               "▶ 3. 서비스 소분류별 지출 현황 (최근 12개월)",
                               "서비스 소분류", prod_type_col)
    r += 1

    # ── Section 2: 통신/일반 × 채널별 Spend Cube ─────────────────────────
    ws.cell(r, 1, "▶ 4. 통신/일반 × 채널별 지출 매트릭스 (최근 12개월)")
    ws.cell(r, 1).font = Font(name=FONT_NAME, bold=True, size=11, color="1F4E79")
    ws.row_dimensions[r].height = 22
    r += 1

    if prod_type_col and "_channel" in recent.columns:
        CHANNELS_USE = ["KT", "그룹사", "외부사", "지입자재"]
        prod_types = ["통신", "일반", "미분류"]

        # 헤더
        ws.cell(r, 1, "구분").font = Font(name=FONT_NAME, bold=True, size=10)
        for ci, ch in enumerate(CHANNELS_USE, 2):
            hcell(ws.cell(r, ci), ch)
        hcell(ws.cell(r, len(CHANNELS_USE) + 2), "합계", bg=C_TOTAL)
        r += 1

        for pt in prod_types:
            pt_bg = "D6E4F7" if pt == "통신" else ("E8F5E9" if pt == "일반" else "F5F5F5")
            dcell(ws.cell(r, 1), pt, bold=True, bg=pt_bg)
            row_total = 0
            for ci, ch in enumerate(CHANNELS_USE, 2):
                mask = (recent[prod_type_col] == pt) & (recent["_channel"] == ch)
                val = round(recent[mask]["_amount"].sum())
                row_total += val
                dcell(ws.cell(r, ci), val if val > 0 else "", NUM_FMT, bg=pt_bg)
            dcell(ws.cell(r, len(CHANNELS_USE) + 2), round(row_total),
                  NUM_FMT, bold=True, bg=C_TOTAL)
            r += 1

        # 채널 합계 행
        dcell(ws.cell(r, 1), "합계", bold=True, bg=C_TOTAL)
        grand = 0
        for ci, ch in enumerate(CHANNELS_USE, 2):
            val = round(recent[recent["_channel"] == ch]["_amount"].sum())
            grand += val
            dcell(ws.cell(r, ci), val, NUM_FMT, bold=True, bg=C_TOTAL)
        dcell(ws.cell(r, len(CHANNELS_USE) + 2), round(grand), NUM_FMT, bold=True, bg=C_TOTAL)

    for c, w in enumerate([28, 18, 18, 12, 12, 10], 1):
        cw(ws, c, w)
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False


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
    ws.auto_filter.ref = f"A{SR}:G{r-1}"


def log_channel_verification(df: pd.DataFrame, settle_cutoff, n_months: int = 6):
    """
    KT 정합도 검증 로그 — 시트 출력 없이 콘솔에 월별 집계를 항상 출력.
    채널별_분석 시트가 제거돼도 이 함수는 항상 호출.
    """
    if "_settle_ym" not in df.columns or "_channel" not in df.columns:
        return

    _settle_raw_cols = ["정산확정일", "정산일자", "일일정산월", "settlement_date"]
    _is_kt = df["_channel"] == "KT"
    _kt_has_settle = pd.Series(False, index=df.index)
    for _sc in _settle_raw_cols:
        if _sc in df.columns:
            _kt_has_settle |= df[_sc].notna()
    if "_order_status" in df.columns:
        _kt_mask = _is_kt & (df["_order_status"] == "전량입고") & _kt_has_settle & df["_settle_ym"].notna()
    else:
        _kt_mask = _is_kt & _kt_has_settle & df["_settle_ym"].notna()

    _settle_raw_col = next((c for c in _settle_raw_cols if c in df.columns), None)
    if _settle_raw_col:
        _has_settle = pd.Series(False, index=df.index)
        for _sc in _settle_raw_cols:
            if _sc in df.columns:
                _has_settle |= df[_sc].notna()
        _from_settle = df["_source_file"].str.contains("정산현황", na=False) \
            if "_source_file" in df.columns else pd.Series(False, index=df.index)
        _other_mask = ~_is_kt & df["_settle_ym"].notna() & (_has_settle | _from_settle)
    else:
        _other_mask = ~_is_kt & df["_settle_ym"].notna()

    log.info(f"채널별 집계: KT 전량입고 {_kt_mask.sum():,}행, 기타채널 정산확정 {_other_mask.sum():,}행")

    # 최근 N개월 KT 월별 집계 출력 (정합도 검증용)
    settle_df = df[(_kt_mask | _other_mask) & df["_settle_ym"].notna()].copy()
    cutoff_start = settle_cutoff - n_months
    recent = settle_df[(settle_df["_settle_ym"] >= cutoff_start) &
                       (settle_df["_settle_ym"] < settle_cutoff)]
    if len(recent) == 0:
        return

    pivot = (recent.groupby(["_channel", "_settle_ym"])["_amount"]
             .sum().unstack("_channel", fill_value=0))
    yms = sorted(recent["_settle_ym"].unique())
    print("\n  [KT 정합도 검증] 채널별 월별 집계 (최근 %d개월)" % n_months)
    print(f"  {'연월':<10}", end="")
    channels_present = [c for c in CHANNELS if c in pivot.columns]
    for ch in channels_present:
        print(f"  {ch:>12}", end="")
    print(f"  {'합계':>14}")
    print("  " + "-" * (10 + 14 * (len(channels_present) + 1)))
    for ym in yms:
        row_total = 0
        print(f"  {str(ym):<10}", end="")
        for ch in channels_present:
            val = int(pivot.loc[ym, ch]) if ym in pivot.index and ch in pivot.columns else 0
            row_total += val
            print(f"  {val/1e8:>10.1f}억", end="")
        print(f"  {row_total/1e8:>12.1f}억")
    print()


def write_sheet_channel(wb, df: pd.DataFrame, current_ym,
                        settle_cutoff, future_periods_list, rev_fc_vals=None,
                        col_map: dict | None = None, exclude_sv: bool = False):
    """
    Sheet: 채널별_분석
    KT / 그룹사 / 외부사 / 지입자재 별 월별 실적 + 6개월 예측
    """
    ws = wb.create_sheet("채널별_분석")
    _sv_note = " | KT: 유가증권 카테고리 제외 (상품권·배송료)" if exclude_sv else ""
    sheet_title(ws, "채널별 발주 분석 (정산일 기준)",
                f"정산확정일 기준 채널별 월별 실적{_sv_note}")

    # ── 채널별 정산일 기준 월별 집계 ──────────────────────────────────────
    # _settle_ym(정산확정일) 기준: 입고/정산 완료월 귀속 (발주월 ≠ 정산월)
    ch_colors = {"KT": "DDEEFF", "그룹사": "DDF0DD", "외부사": "FFF0CC", "지입자재": "FFE0E0"}
    ch_header_colors = {"KT": "1F4E79", "그룹사": "375623", "외부사": "843C0C", "지입자재": "6B2737"}

    SR = 3
    # 헤더: 연월 | KT실적 | KT예측 | 그룹사실적 | 그룹사예측 | ...
    col_headers = ["연월"]
    for ch in CHANNELS:
        col_headers += [f"{ch} 실적(원)", f"{ch} 예측(원)"]
    col_headers.append("합계 실적(원)")
    col_headers.append("합계 예측(원)")

    for c, h in enumerate(col_headers, 1):
        ch_name = None
        for ch in CHANNELS:
            if ch in h:
                ch_name = ch
                break
        bg = ch_header_colors.get(ch_name, C_HEADER)
        hcell(ws.cell(SR, c), h, bg=bg)
    ws.row_dimensions[SR].height = 28

    # ── 과거 실적 (채널별 정산일 기준 월별 집계) ────────────────────────────
    # KT/비KT 모두 정산확정일(_settle_ym) 기준 귀속
    # 귀속 근거: 입고승인일 기준이 정확하나 데이터 미제공 → 정산확정일 년월로 근사
    # (정산확정일 단순 년월 = KT 4월 실적 +0.16% 오차, 가장 정확한 근사치)
    if "_settle_ym" in df.columns:
        _settle_raw_cols = ["정산확정일", "정산일자", "일일정산월", "settlement_date"]
        _settle_raw_col = next((c for c in _settle_raw_cols if c in df.columns), None)

        _from_settle_file = df["_source_file"].str.contains("정산현황", na=False) \
            if "_source_file" in df.columns else pd.Series(False, index=df.index)

        _is_kt = df["_channel"] == "KT" if "_channel" in df.columns else pd.Series(False, index=df.index)

        # KT: 전량입고 + 정산확정일 있는 건 = 정산완료
        _kt_has_settle = pd.Series(False, index=df.index)
        for _sc in _settle_raw_cols:
            if _sc in df.columns:
                _kt_has_settle = _kt_has_settle | df[_sc].notna()
        if "_order_status" in df.columns:
            _kt_mask = _is_kt & (df["_order_status"] == "전량입고") & _kt_has_settle & df["_settle_ym"].notna()
        else:
            _kt_mask = _is_kt & _kt_has_settle & df["_settle_ym"].notna()

        # 그룹사/외부사/지입자재: 정산확정일 기준
        if _settle_raw_col:
            _has_settle = pd.Series(False, index=df.index)
            for _sc in _settle_raw_cols:
                if _sc in df.columns:
                    _has_settle = _has_settle | df[_sc].notna()
            _other_mask = ~_is_kt & df["_settle_ym"].notna() & (_has_settle | _from_settle_file)
        else:
            _other_mask = ~_is_kt & (df["_settle_ym"].notna() | _from_settle_file)

        _settled_mask = _kt_mask | _other_mask

        _settle_start = pd.Period("2022-01", "M")
        settle_df = df[_settled_mask & (df["_settle_ym"] >= _settle_start) & (df["_settle_ym"] < settle_cutoff)].copy()
        if exclude_sv and "서비스카테고리" in settle_df.columns:
            _kt_excl = (settle_df["_channel"] == "KT") & \
                       settle_df["서비스카테고리"].str.startswith("서비스/유가증권 > 유가증권", na=False)
            n_excl = _kt_excl.sum()
            settle_df = settle_df[~_kt_excl]
            log.info(f"채널별_분석 KT 유가증권 제외: {n_excl:,}행")
        ch_grp = settle_df.groupby(["_channel", "_settle_ym"])["_amount"].sum().reset_index()
        ch_grp.rename(columns={"_settle_ym": "_ym", "_amount": "amount"}, inplace=True)
        log.info(f"채널별_분석 집계: KT 전량입고 {_kt_mask.sum():,}행, "
                 f"기타채널 정산확정 {_other_mask.sum():,}행")
    else:
        ch_grp = pd.DataFrame()

    all_yms = sorted(ch_grp["_ym"].unique()) if len(ch_grp) > 0 else []

    r = SR + 1
    for ym in all_yms:
        ws.cell(r, 1, str(ym)).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 1).border = tborder()
        total_actual = 0
        for ci, ch in enumerate(CHANNELS, 0):
            col_actual = 2 + ci * 2
            col_fcst   = 3 + ci * 2
            val = ch_grp[(ch_grp["_channel"] == ch) & (ch_grp["_ym"] == ym)]["amount"]
            v = float(val.iloc[0]) if len(val) > 0 else 0
            total_actual += v
            bg = ch_colors.get(ch)
            dcell(ws.cell(r, col_actual), round(v) if v > 0 else "", NUM_FMT, bg=bg)
            dcell(ws.cell(r, col_fcst),   "", bg=None)
        dcell(ws.cell(r, 2 + len(CHANNELS)*2), round(total_actual) if total_actual > 0 else "",
              NUM_FMT, bold=True, bg=C_TOTAL)
        dcell(ws.cell(r, 3 + len(CHANNELS)*2), "", bg=None)
        r += 1

    # ── 예측 구간 (미완성 + 당월 + 미래) ──────────────────────────────────
    for ym in future_periods_list:
        is_incomplete = ym < current_ym
        is_cur = ym == current_ym
        bg_row = "FFF9C4" if (is_incomplete or is_cur) else None
        lbl = "정산지연" if is_incomplete else ("당월" if is_cur else "")

        ws.cell(r, 1, f"{ym} [{lbl}]" if lbl else str(ym))
        ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 1).border = tborder()
        if bg_row:
            ws.cell(r, 1).fill = PatternFill("solid", start_color=bg_row)

        # 채널별 예측: 채널별 과거 비율 기반 배분
        # 최근 6개월 채널별 비중으로 예측 배분
        recent_yms = [ym2 for ym2 in all_yms if ym2 >= (settle_cutoff - 6)]
        ch_ratios = {}
        for ch in CHANNELS:
            ch_tot = ch_grp[(ch_grp["_channel"] == ch) &
                            (ch_grp["_ym"].isin(recent_yms))]["amount"].sum()
            ch_ratios[ch] = max(ch_tot, 0)
        total_ratio = sum(ch_ratios.values())
        if total_ratio == 0:
            total_ratio = 1

        # 매출 예측값 조회
        fc_idx = future_periods_list.index(ym) if ym in future_periods_list else -1
        if rev_fc_vals is not None and 0 <= fc_idx < len(rev_fc_vals):
            total_fc = rev_fc_vals[fc_idx]
        else:
            total_fc = 0

        for ci, ch in enumerate(CHANNELS, 0):
            col_actual = 2 + ci * 2
            col_fcst   = 3 + ci * 2
            dcell(ws.cell(r, col_actual), "", bg=None)

            # 채널별 예측값 = 전체 예측 × (채널비중/전체비중)
            ch_fc = (total_fc * ch_ratios[ch] / total_ratio) if total_ratio > 0 else 0
            bg = ch_colors.get(ch, bg_row)
            dcell(ws.cell(r, col_fcst), round(ch_fc) if ch_fc > 0 else "", NUM_FMT, bg=bg)

        # 합계 예측값도 채우기
        dcell(ws.cell(r, 2 + len(CHANNELS)*2), "", bg=None)
        dcell(ws.cell(r, 3 + len(CHANNELS)*2), round(total_fc) if total_fc > 0 else "",
              NUM_FMT, bold=True, bg=C_TOTAL)
        r += 1

    _sv_footer = " / 서비스/유가증권 카테고리 제외" if exclude_sv else ""
    ws.cell(r + 1, 1, f"※ 집계 기준: KT — 전량입고 시점 / 그룹사·외부사·지입자재 — 정산확정일 기준{_sv_footer}")
    ws.cell(r + 2, 1, "※ 예측 컬럼은 매출 예측 배분 기반 채널별 비중 적용")

    for c, w in enumerate([14] + [18, 18] * len(CHANNELS) + [18, 18], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    from openpyxl.utils import get_column_letter as _gcl2
    ws.auto_filter.ref = f"A{SR}:{_gcl2(len(col_headers))}{r-1}"
    ws.sheet_view.showGridLines = False


def write_sheet_revenue(wb, revenue_train: pd.DataFrame,
                         rev_future_periods, rev_fc_vals, rev_ci_lo, rev_ci_hi,
                         rev_best: str, rev_mape: float, current_ym=None):
    """Sheet 6: 정산일 기준 매출 예측"""
    if current_ym is None:
        current_ym = pd.Period(datetime.now(), freq="M")
    ws = wb.create_sheet("매출예측_정산일기준")
    sheet_title(ws, "매출 예측 (정산일 기준)",
                f"실제 매출(정산완료 기준) 과거 추이 + 향후 6개월 예측")

    SR = 3
    headers = ["연월", "실제 정산금액(원)", "예측 정산금액(원)", "하한(95%)", "상한(95%)", "구분"]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)
    ws.row_dimensions[SR].height = 28

    r = SR + 1
    # 과거 실적
    for _, row in revenue_train.iterrows():
        dcell(ws.cell(r, 1), str(row["_ym"]), align="center")
        dcell(ws.cell(r, 2), row["amount"], NUM_FMT)
        dcell(ws.cell(r, 3), "", bg=None)
        dcell(ws.cell(r, 4), "", bg=None)
        dcell(ws.cell(r, 5), "", bg=None)
        dcell(ws.cell(r, 6), "실적", align="center", bg=C_ACTUAL)
        r += 1

    # 예측
    for i, p in enumerate(rev_future_periods):
        is_lag = p < current_ym          # 지연 미완성 구간 (예: 5월)
        is_cur = p == current_ym         # 당월 (6월)
        bg  = "FFF9C4" if (is_lag or is_cur) else C_FCST
        lbl = "정산지연(미완성)" if is_lag else ("이번 달 예측" if is_cur else "예측")
        dcell(ws.cell(r, 1), str(p), align="center", bg=bg, bold=(i==0))
        dcell(ws.cell(r, 2), "", bg=None)
        dcell(ws.cell(r, 3), round(rev_fc_vals[i]),  NUM_FMT, bg=bg, bold=(i==0))
        dcell(ws.cell(r, 4), round(rev_ci_lo[i] if i < len(rev_ci_lo) else 0),
              NUM_FMT, bg=bg)
        dcell(ws.cell(r, 5), round(rev_ci_hi[i] if i < len(rev_ci_hi) else 0),
              NUM_FMT, bg=bg)
        dcell(ws.cell(r, 6), lbl, align="center", bg=bg, bold=(i==0))
        r += 1

    # 모델 정보
    ws.cell(r + 1, 1, f"※ 예측 모델: {rev_best}  |  MAPE: {rev_mape:.1f}%")
    ws.cell(r + 2, 1, "※ 정산일 기준: 해당 월에 정산이 완료된 주문 건의 합계금액")
    ws.cell(r + 3, 1, f"※ 노란색: 정산 미완성 예측 구간 (정산 지연 {SETTLE_LAG_MONTHS}개월 설정)")

    for c, w in enumerate([14, 22, 22, 18, 18, 14], 1):
        cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"
    ws.auto_filter.ref = f"A{SR}:F{r-1}"
    ws.sheet_view.showGridLines = False


def write_sheet_glossary(wb):
    """
    Sheet: 용어설명
    기술 용어를 초보자도 이해할 수 있도록 풀어서 설명하는 가이드 시트.
    항상 첫 번째 시트로 배치.
    """
    from openpyxl.styles import PatternFill as PF, Font as FT, Alignment as AL, Border, Side
    ws = wb.create_sheet("📖 용어설명", 0)
    ws.sheet_view.showGridLines = False
    ws.tab_color = "1F4E79"

    # 제목
    ws.merge_cells("A1:F1")
    tc = ws.cell(1, 1, "📖 수요예측 시스템 — 용어 가이드")
    tc.font      = FT(name=FONT_NAME, bold=True, size=16, color="FFFFFF")
    tc.fill      = PF("solid", fgColor="1F4E79")
    tc.alignment = AL(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:F2")
    sc = ws.cell(2, 1, "이 시트는 보고서에 등장하는 용어를 쉽게 풀어 설명합니다. 처음 보시는 분도 바로 이해할 수 있도록 실무 예시를 함께 제공합니다.")
    sc.font      = FT(name=FONT_NAME, size=10, color="595959", italic=True)
    sc.alignment = AL(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28

    terms = [
        # (섹션색, 용어, 한 줄 정의, 실무 예시)
        ("1F4E79", "📦 재고 관련 용어", None, None),
        ("2E75B6", "안전재고 (SS, Safety Stock)",
         "수요가 갑자기 늘거나 납품이 늦어질 때를 대비해 미리 쌓아두는 최소한의 여유 재고입니다.",
         "예) SS = 50개 → 재고가 50개 아래로 떨어지지 않도록 유지"),
        ("2E75B6", "재주문점 (ROP, Re-Order Point)",
         "재고가 이 수준까지 줄어들면 지금 바로 발주를 넣어야 하는 신호선입니다.",
         "예) ROP = 120개 → 재고 120개 남았을 때 발주 → 납품 받는 동안 안전재고(50개) 소진 안 됨"),
        ("2E75B6", "리드타임 (Lead Time)",
         "발주를 넣은 날부터 실제 물건이 입고되기까지 걸리는 날수입니다.",
         "예) 리드타임 14일 → 오늘 발주 → 14일 후 입고"),
        ("375623", "📊 품목 분류 관련 용어", None, None),
        ("4D7C2E", "ABC 등급",
         "구매금액 기준으로 품목을 중요도에 따라 3단계로 나눈 분류입니다. 소수의 A등급 품목이 전체 매출의 대부분을 차지합니다.",
         "A등급: 구매금액 상위 70% 품목 (핵심 관리 대상)\nB등급: 누적 70~90% 품목 (일반 관리)\nC등급: 나머지 품목 (간소 관리)"),
        ("4D7C2E", "CV (변동계수, Coefficient of Variation)",
         "수요가 얼마나 들쭉날쭉한지 나타내는 지표입니다. 숫자가 클수록 수요 예측이 어렵습니다.",
         "CV 낮음 → 매달 비슷한 수량 나감 → 예측 쉬움\nCV 높음 → 어떤 달은 많고 어떤 달은 적음 → 예측 어려움"),
        ("843C0C", "📈 예측 모델 관련 용어", None, None),
        ("C55A11", "Holt-Winters (홀트-윈터스)",
         "계절성(봄/여름/가을/겨울 같은 반복 패턴)과 추세(오름세·내림세)를 동시에 반영하는 예측 방법입니다. 쉽게 말해 '작년 이맘때 수요 + 최근 추세'를 합산해 예측합니다.",
         "예) 에어컨은 여름에 많이 팔리는 계절성이 있는데, 이 패턴을 자동으로 학습해 예측"),
        ("C55A11", "선형회귀 (Linear Regression)",
         "과거 수요의 증가/감소 추세를 직선으로 그어 미래를 예측하는 방법입니다.",
         "예) 지난 12개월 동안 매달 10개씩 늘었다면 → 다음 달도 10개 더 늘 것으로 예측"),
        ("C55A11", "SARIMA",
         "계절성과 불규칙 변동을 동시에 분석하는 고급 예측 모델입니다. 데이터가 24개월 이상 쌓였을 때 사용합니다.",
         "Holt-Winters보다 복잡하지만 변동성이 큰 품목에 유리"),
        ("C55A11", "MA3 (3개월 이동평균)",
         "최근 3개월 평균을 그대로 다음 달 예측으로 사용하는 가장 단순한 방법입니다.",
         "예) 3월 100개, 4월 120개, 5월 110개 → 6월 예측 = (100+120+110)÷3 = 110개"),
        ("7B2D8B", "🎯 정확도 관련 용어", None, None),
        ("8B4AAF", "MAPE (예측 오차율, Mean Absolute Percentage Error)",
         "예측이 실제와 얼마나 다른지를 퍼센트(%)로 나타낸 지표입니다. 낮을수록 예측이 정확합니다.",
         "MAPE 10% → 실제 100개일 때 예측이 90~110개 수준\nMAPE 30% 이상 → 주의 필요 (수요 변동이 너무 크거나 데이터 부족)"),
        ("8B4AAF", "Tracking Signal (예측 편향 지수)",
         "예측이 지속적으로 높거나 낮은 방향으로 치우쳐 있는지 감지하는 지표입니다.",
         "TS > +4 → 예측이 계속 실제보다 낮음 → 발주량 상향 검토\nTS < -4 → 예측이 계속 실제보다 높음 → 발주량 하향 검토\n-4 ~ +4 → 정상 범위"),
        ("C00000", "⚠️ 알림 관련 용어", None, None),
        ("C00000", "수요 급등 알림",
         "향후 3개월 예측 수요가 최근 12개월 평균보다 30% 이상 높을 때 발생하는 알림입니다.",
         "예) 평균 100개/월 → 향후 예측 140개/월 이상이면 급등 알림 발동"),
        ("C00000", "결품 위험",
         "현재 재고가 예상 수요를 충당하지 못해 품절이 발생할 위험이 있는 상태입니다.",
         "재고 입력 후 활성화됨 — 구매계획_입력.xlsx에 현재재고를 기재해 주세요"),
        ("1F4E79", "🏭 MRO 이관 품목 관련 용어", None, None),
        ("2E75B6", "수명주기 단계 (Lifecycle Stage)",
         "품목이 현재 성장/성숙/쇠퇴 중 어느 단계에 있는지를 수요 추세로 자동 판정한 결과입니다.",
         "성장기: 수요 지속 증가 중\n성숙기: 수요 안정적 (정상)\n쇠퇴기: 최근 6개월 수요가 피크 대비 60% 미만으로 감소\n단종임박: 최근 3개월 수요가 피크 대비 30% 미만"),
        ("2E75B6", "EOL (End of Life, 단종 예상 시점)",
         "수요 감소 추세를 직선으로 연장했을 때 수요가 0에 도달하는 예상 월입니다. 미리 파악해 대체품 검토 또는 재고 소진 계획을 세울 수 있습니다.",
         "예) EOL 2026-10 → 이 품목은 2026년 10월쯤 수요 소멸 예상 → 9월까지 재고 소진 권장"),
        ("2E75B6", "서비스수준 98% (SL98%)",
         "발주~입고 기간(리드타임) 중 수요를 100% 충족할 확률을 98%로 설정한 기준입니다. KT SCM 이관 MRO 품목 기본값.",
         "SL 95% → 안전재고 1.65배\nSL 98% → 안전재고 2.05배 (더 두껍게 유지)"),
        ("2E75B6", "이관 시점 필터 (TRANSFER_DATE)",
         "KT SCM에서 이관받은 날짜 이전 데이터를 예측 학습에서 제외하는 기능입니다. 이관 전 운영 환경이 달라 과거 패턴이 현재와 맞지 않을 때 사용합니다.",
         "예) TRANSFER_DATE = '2023-04-01' → 2023년 4월 이후 데이터만 학습에 사용"),
    ]

    r = 4
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 42

    thin = Side(style="thin", color="DDDDDD")
    box  = Border(left=thin, right=thin, top=thin, bottom=thin)

    for item in terms:
        color, term, definition, example = item
        if definition is None:
            # 섹션 헤더
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            hc = ws.cell(r, 1, f"  {term}")
            hc.font      = FT(name=FONT_NAME, bold=True, size=11, color="FFFFFF")
            hc.fill      = PF("solid", fgColor=color)
            hc.alignment = AL(vertical="center")
            ws.row_dimensions[r].height = 24
            r += 1
            continue

        # 용어 셀
        tc2 = ws.cell(r, 2, term)
        tc2.font      = FT(name=FONT_NAME, bold=True, size=10, color=color)
        tc2.fill      = PF("solid", fgColor="F7F7F7")
        tc2.alignment = AL(vertical="top", wrap_text=True)
        tc2.border    = box

        # 설명 셀
        dc = ws.cell(r, 3, definition)
        dc.font      = FT(name=FONT_NAME, size=10)
        dc.alignment = AL(vertical="top", wrap_text=True)
        dc.border    = box

        # 예시 셀
        ec = ws.cell(r, 4, example or "")
        ec.font      = FT(name=FONT_NAME, size=9, color="595959", italic=True)
        ec.fill      = PF("solid", fgColor="FFFDE7")
        ec.alignment = AL(vertical="top", wrap_text=True)
        ec.border    = box

        ws.row_dimensions[r].height = max(48, definition.count("\n") * 18 + 30)
        r += 1

    # 헤더 행
    for c, label in enumerate(["", "용어", "설명", "실무 예시 / 판단 기준"], 1):
        hc2 = ws.cell(3, c, label)
        hc2.font      = FT(name=FONT_NAME, bold=True, size=10, color="FFFFFF")
        hc2.fill      = PF("solid", fgColor="404040")
        hc2.alignment = AL(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 22

    # 하단 안내
    ws.merge_cells(start_row=r+1, start_column=1, end_row=r+1, end_column=4)
    nc = ws.cell(r+1, 1, "※ 이 보고서는 AI 기반 수요예측 모델(Holt-Winters / 선형회귀 / SARIMA)로 생성됩니다. 예측은 과거 데이터 기반이며 실제 수요와 차이가 발생할 수 있습니다.")
    nc.font      = FT(name=FONT_NAME, size=9, color="888888", italic=True)
    nc.alignment = AL(wrap_text=True)
    ws.row_dimensions[r+1].height = 28


def write_sheet_xai(wb, acc_rows: list):
    """
    XAI(설명 가능한 AI) 시트 — 시계열별 모델 경쟁 결과 및 선택 근거 표시.
    각 행: 시계열명 / 선택모델 / HW·SARIMA·선형회귀·MA3·Prophet·LightGBM MAPE / 선택 이유
    """
    ws = wb.create_sheet("XAI_모델선택근거")
    sheet_title(ws, "XAI — 예측 모델 선택 근거",
                "각 시계열별 전체 후보 모델 MAPE와 최종 선택 이유를 확인할 수 있습니다.")

    MODELS = ["Holt-Winters", "SARIMA", "선형회귀", "MA3", "Prophet", "LightGBM"]
    headers = ["시계열명", "등급", "선택모델"] + [f"{m}\nMAPE(%)" for m in MODELS] + ["선택 근거"]
    r = 3
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(r, c), h, size=9)

    def _reason(row):
        best = row.get("best", "")
        mapes = {m: row.get(m, np.nan) for m in MODELS}
        valid  = {m: v for m, v in mapes.items() if not np.isnan(v)}
        if not valid:
            return "데이터 부족 — 검증 불가"
        sorted_m = sorted(valid, key=lambda x: valid[x])
        rank1, rank2 = sorted_m[0], sorted_m[1] if len(sorted_m) > 1 else None
        gap = valid[rank1] if rank2 is None else valid[rank2] - valid[rank1]
        reason = f"MAPE 최소({valid[rank1]:.1f}%)"
        if rank2 and gap < 2:
            reason += f" — {rank2}({valid[rank2]:.1f}%)와 근소 차이"
        if best != rank1:
            reason += f" ※ ABC 전략({best}) 유지"
        return reason

    alt = False
    for row in acc_rows:
        r += 1
        bg = C_ALT if alt else None
        alt = not alt
        name  = row.get("name", "")
        grade = row.get("note", "")
        best  = row.get("best", "")
        dcell(ws.cell(r, 1), name,  align="left",   bg=bg)
        dcell(ws.cell(r, 2), grade, align="center",  bg=bg)
        best_bg = "E2EFDA" if not np.isnan(row.get(best, np.nan)) and row.get(best, 999) < 20 else \
                  C_WARN   if not np.isnan(row.get(best, np.nan)) and row.get(best, 999) > 30 else bg
        dcell(ws.cell(r, 3), best, align="center", bg=best_bg)
        for ci, m in enumerate(MODELS, 4):
            val = row.get(m, np.nan)
            if np.isnan(val):
                dcell(ws.cell(r, ci), "—", align="center", bg=bg)
            else:
                cell_bg = "E2EFDA" if m == best else (C_WARN if val > 30 else bg)
                dcell(ws.cell(r, ci), round(val, 1), fmt="0.0", align="center", bg=cell_bg)
        dcell(ws.cell(r, len(MODELS) + 4), _reason(row), align="left", bg=bg)

    widths = [35, 6, 14] + [12] * len(MODELS) + [50]
    for ci, w in enumerate(widths, 1):
        cw(ws, ci, w)
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False


def write_sheet_ecos(wb, ecos_data: dict):
    """한국은행 ECOS 외부 지표 시트 — 소비자심리지수 등 거시지표 추이"""
    if not ecos_data:
        return
    ws = wb.create_sheet("외부지표_ECOS")
    sheet_title(ws, "한국은행 ECOS 거시경제 지표",
                "수요 예측 보조 지표 — 소비자심리지수(CSI) 등 외부 환경 모니터링")

    col = 1
    for label, df_e in ecos_data.items():
        r = 3
        hcell(ws.cell(r, col),   "연월")
        hcell(ws.cell(r, col+1), label)
        for _, erow in df_e.iterrows():
            r += 1
            dcell(ws.cell(r, col),   str(erow["ym"]),  align="center")
            dcell(ws.cell(r, col+1), round(float(erow["val"]), 2), fmt="0.00")
        cw(ws, col,   12)
        cw(ws, col+1, 18)
        col += 3

    ws.sheet_view.showGridLines = False


def write_sheet_qbr_savings(wb, detail_rows: list, ss_rop: dict, df_raw: pd.DataFrame):
    """
    QBR용 안전재고 절감액 자동 산출 시트.

    고객사가 KT Commerce 없이 직접 소싱할 경우 필요한 안전재고(SL 95%)와
    KT Commerce 이용 시 제거 가능한 보유비용을 품목별로 산출.

    전제:
      - 고객사 자체 SS = Z_CUST(1.65) × σ_D × √LT  (SL 95% 자체 관리 기준)
      - KTC 납기 보장 시 고객사 보유 SS = 0 (KTC가 납기 확률 98%로 흡수)
      - 연간 보유비용 = SS × 단가 × HOLDING_COST_RATE(25%)
    """
    ws = wb.create_sheet("QBR_안전재고절감분석")
    ws.tab_color = "375623"
    sheet_title(
        ws, "QBR — KT Commerce 이용 시 고객사 안전재고 절감 분석",
        f"전제: 고객사 자체관리 SL 95%(Z=1.65) / 연간 보유비용률 {HOLDING_COST_RATE:.0%} / "
        f"KTC 납기보장 SL 98%(Z={Z_SS})"
    )

    # ── 품목별 단가 산출 (최근 12개월 실적 기준 가중평균 단가) ──────────────
    unit_price_map: dict = {}
    if not df_raw.empty and "_qty" in df_raw.columns and "_amount" in df_raw.columns:
        grp_cols = ["_vendor", "_prod_key"] if "_vendor" in df_raw.columns else ["_prod_key"]
        try:
            up = (
                df_raw.groupby(grp_cols)
                .apply(lambda g: g["_amount"].sum() / g["_qty"].sum()
                       if g["_qty"].sum() > 0 else 0.0)
                .reset_index(name="_unit_price")
            )
            for _, row in up.iterrows():
                if len(grp_cols) == 2:
                    key = (str(row["_vendor"]), str(row["_prod_key"]))
                else:
                    key = ("", str(row["_prod_key"]))
                unit_price_map[key] = float(row["_unit_price"])
        except Exception:
            pass

    # ── 헤더 ──────────────────────────────────────────────────────────────
    SR = 4
    headers = [
        "협력사명", "상품코드", "상품명", "ABC등급",
        "월평균수요(개)", "LT평균(일)", "수요σ(개/월)",
        "고객사 자체SS\n(SL95%, 개)", "KTC이용 후\n고객SS(개)",
        "SS절감\n(개)", "단가(원)",
        "연간 보유비용\n절감 추정액(원)",
        "재고일수\n절감(일)", "비고",
    ]
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h, size=9)
    ws.row_dimensions[SR].height = 32

    # ── 데이터 행 ─────────────────────────────────────────────────────────
    r = SR
    total_saving = 0.0
    alt = False

    for row in detail_rows:
        vendor   = row.get("vendor", "")
        prod_key = row.get("prod_key", "")
        abc      = row.get("ABC", "")
        avg12    = row.get("avg12", 0.0)
        key      = (vendor, prod_key)

        sr_data  = ss_rop.get(key, {})
        ss_ktc   = sr_data.get("ss", 0.0)           # KTC 내부 SS (SL 98%)
        lt_days  = sr_data.get("lead_time", 30.0) * 30   # 월 → 일
        lt_m     = lt_days / 30

        # σ_D 역산: ss_ktc = Z_SS × σ_D × √LT_m (기본공식 기준)
        sigma_d = (ss_ktc / (Z_SS * np.sqrt(lt_m))) if (lt_m > 0 and ss_ktc > 0) else 0.0

        # 고객사 자체 SS (SL 95%)
        cust_ss = Z_CUST * sigma_d * np.sqrt(lt_m)

        # KTC 이용 시 고객 보유 SS = 0 (납기 보장으로 완충재고 불필요)
        ktc_cust_ss = 0.0

        ss_reduction = max(cust_ss - ktc_cust_ss, 0.0)

        # 단가
        unit_price = unit_price_map.get(key, unit_price_map.get(("", prod_key), 0.0))

        # 연간 보유비용 절감 = SS절감량 × 단가 × HOLDING_COST_RATE
        annual_saving = ss_reduction * unit_price * HOLDING_COST_RATE

        # 재고일수 절감 = SS절감 / 일평균수요
        daily_avg = avg12 / 30.0
        days_saved = (ss_reduction / daily_avg) if daily_avg > 0 else 0.0

        total_saving += annual_saving

        r += 1
        bg = C_ALT if alt else None
        alt = not alt

        note = ""
        if abc == "C":
            note = "C등급: 최소1개 유지"
        elif row.get("lifecycle_stage", "") in ("쇠퇴기", "단종임박"):
            note = f"⚠ {row.get('lifecycle_stage', '')} — 재고 검토 권고"

        dcell(ws.cell(r, 1),  vendor,             align="left",   bg=bg)
        dcell(ws.cell(r, 2),  prod_key,            align="center", bg=bg)
        dcell(ws.cell(r, 3),  row.get("prod_name",""), align="left", bg=bg)
        dcell(ws.cell(r, 4),  abc,                 align="center", bg=bg)
        dcell(ws.cell(r, 5),  round(avg12, 1),     fmt="0.0",      bg=bg)
        dcell(ws.cell(r, 6),  round(lt_days, 0),   fmt="0",        bg=bg)
        dcell(ws.cell(r, 7),  round(sigma_d, 2),   fmt="0.00",     bg=bg)
        dcell(ws.cell(r, 8),  round(cust_ss, 1),   fmt="0.0",      bg=bg)
        dcell(ws.cell(r, 9),  round(ktc_cust_ss,1),fmt="0.0",  align="center", bg="E2EFDA")
        dcell(ws.cell(r, 10), round(ss_reduction,1),fmt="0.0",     bg=bg)
        dcell(ws.cell(r, 11), round(unit_price, 0), fmt="#,##0",   bg=bg)
        # 절감액 — 의미 있는 금액만 색상 강조
        saving_bg = "E2EFDA" if annual_saving >= 100_000 else (C_WARN if annual_saving == 0 and unit_price > 0 else bg)
        dcell(ws.cell(r, 12), round(annual_saving, 0), fmt="#,##0", bg=saving_bg)
        dcell(ws.cell(r, 13), round(days_saved, 1),    fmt="0.0",   bg=bg)
        dcell(ws.cell(r, 14), note, align="left", bg=bg)

    # ── 합계 행 ───────────────────────────────────────────────────────────
    r += 1
    hcell(ws.cell(r, 1), "합  계", bg="1F4E79")
    for c in range(2, 12):
        ws.cell(r, c).value = None
    dcell(ws.cell(r, 12), round(total_saving, 0), fmt="#,##0", bold=True, bg="375623")
    ws.cell(r, 12).font = __import__("openpyxl").styles.Font(
        bold=True, color="FFFFFF", size=11)

    # ── 안내 문구 (QBR 발표용) ────────────────────────────────────────────
    r += 2
    note_lines = [
        "【QBR 활용 안내】",
        f"  · 고객사 자체관리 SS: 직접 소싱 시 SL 95% 유지를 위해 보유해야 할 안전재고 추정치입니다.",
        f"  · KTC 이용 후 SS = 0: KT Commerce 납기 보장(SL {Z_SS*100-100:.0f}+%) 활용 시 고객사 완충재고 불필요.",
        f"  · 연간 절감 추정액 = SS절감(개) × 단가 × {HOLDING_COST_RATE:.0%} (자본비용·창고비·진부화 합산 보유비용률).",
        f"  · 단가 미입력 품목(0원)은 절감액이 산출되지 않습니다. 별도 입력 후 재실행하세요.",
    ]
    for line in note_lines:
        dcell(ws.cell(r, 1), line, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        r += 1

    # ── 열 너비 ───────────────────────────────────────────────────────────
    widths = [20, 16, 24, 6, 10, 8, 10, 12, 12, 10, 14, 18, 10, 24]
    for ci, w in enumerate(widths, 1):
        cw(ws, ci, w)
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False


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
    purchase_plan = pd.DataFrame()   # 구매계획 미입력 시 기본값
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

        # 존재 확인 (디렉터리 입력 시 해당 폴더 내 파일 자동 탐색)
        expanded = []
        for f in filepaths:
            if f.is_dir():
                dir_files = [p for p in f.iterdir() if p.is_file() and _is_data_file(p)]
                if dir_files:
                    print(f"  [안내] 폴더를 입력하셨습니다. 내부 파일 {len(dir_files)}개를 사용합니다.")
                    expanded.extend(dir_files)
                else:
                    print(f"  [경고] '{f.name}' 폴더에서 데이터 파일을 찾지 못했습니다.")
            elif f.exists():
                expanded.append(f)
        filepaths = expanded
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

        # ── 구매계획 로드 ──────────────────────────────────────────────────
        print("\n[구매계획 설정]")
        purchase_plan = load_purchase_plan(BASE_DIR)
        if purchase_plan.empty:
            print("  → 구매계획_입력.xlsx 없음 (ML 예측만 사용)")
            print("     파일럿 완료 후 담당자 입력 템플릿이 자동 생성됩니다.")
        else:
            print(f"  → 구매계획 반영: {len(purchase_plan):,}개 품목-월")

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
            print("     파일 경로 입력 / 숫자 입력(전체 단일값) / Enter(주문일자~배송완료일자 실측값 사용)")
            lt_ans = input("  선택 > ").strip()
            if lt_ans == "":
                lt_df = pd.DataFrame()
                lead_time = DEFAULT_LEAD_TIME   # 실측 없는 품목 fallback
                print("  → 주문일자~배송완료일자 실측값 사용 (품목별 평균·표준편차 자동 산출)")
                print(f"     실측 없는 품목은 기본값 {DEFAULT_LEAD_TIME}일 적용")
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
        save_config(filepaths, lead_time, lt_file_used, col_map, processed_files=[])
        print(f"\n[설정 저장] 다음 실행 시 자동으로 로드됩니다. ({CONFIG_FILE.name})")

    # ── 한국은행 ECOS 외부 지표 (API 키 있을 때만) ───────────────────────
    print("\n[외부지표] 한국은행 ECOS 데이터 조회 중...")
    ecos_data = fetch_ecos_data(n_months=36)

    # ── 데이터 로드 (증분: 새 파일만 처리, 기존은 parquet 재사용) ──────────
    processed_files = saved_cfg.get("processed_files", []) if saved_cfg else []
    df, processed_files = load_incremental(filepaths, col_map, lead_time, processed_files)

    # parquet 재사용 시 _prod_type 누락 보완 (매핑 파일 기준 통신/일반 구분)
    df = ensure_prod_type(df, col_map)

    # 처리된 파일 목록을 config에 갱신 저장
    save_config(filepaths, lead_time, lt_file_used, col_map, processed_files)

    # ── 분석 범위 선택 (유가증권 제외 여부) ─────────────────────────────
    # 유가증권 제외 조건 (OR):
    #   ① 마스터카테고리(CMS) 가 "서비스/유가증권 > 유가증권" 으로 시작
    #   ② 서비스카테고리 에 "캠페인용 상품권" 포함
    #   ③ 서비스카테고리 가 "서비스/유가증권 > 유가증권" 으로 시작
    #   ④ 서비스카테고리 가 "서비스 > 유가증권" 으로 시작
    # ※ "서비스/유가증권 > 서비스" 하위(여행·캔틴·프로모션 등)는 실물/서비스이므로 포함
    # 원본 컬럼명: 마스터카테고리(CMS), 서비스카테고리 (parquet 그대로 접근)

    _amt_col   = "_amount" if "_amount" in df.columns else "amount"
    _scope_tag = "전체"

    # ── ① 마스터카테고리(CMS) ─────────────────────────────────────────
    _CMS_COL = "마스터카테고리(CMS)"
    if _CMS_COL in df.columns:
        mask_cms = df[_CMS_COL].astype(str).str.startswith("서비스/유가증권 > 유가증권")
    else:
        mask_cms = pd.Series(False, index=df.index)

    # ── ②③④ 서비스카테고리 ──────────────────────────────────────────
    # col_map["category"] 가 실제 컬럼이면 우선 사용, 없으면 "서비스카테고리" 직접 접근
    _SVC_COL = (col_map.get("category")
                if col_map.get("category") and col_map["category"] in df.columns
                else ("서비스카테고리" if "서비스카테고리" in df.columns else None))

    if _SVC_COL:
        _svc = df[_SVC_COL].astype(str)
        mask_svc = (
            _svc.str.contains("캠페인용 상품권", regex=False, na=False)   # ②
            | _svc.str.startswith("서비스/유가증권 > 유가증권")             # ③
            | _svc.str.startswith("서비스 > 유가증권")                    # ④
        )
    else:
        mask_svc = pd.Series(False, index=df.index)

    # ── ⑥ 상품명 키워드 기반 (기프티쇼·상품권·포인트 등 유가증권성 상품) ──
    # 카테고리가 다르게 분류돼도 상품명으로 잡아냄
    _SEC_NAME_KEYWORDS = [
        "기프티쇼", "기프티콘", "상품권", "포인트권", "네이버페이",
        "카카오페이", "쿠폰", "바우처", "voucher", "gift card",
        "모바일상품권", "모바일쿠폰",
    ]
    _NAME_COL = (col_map.get("prod_name")
                 if col_map.get("prod_name") and col_map["prod_name"] in df.columns
                 else ("상품명" if "상품명" in df.columns else None))
    if _NAME_COL:
        _name = df[_NAME_COL].astype(str)
        _kw_pattern = "|".join(_SEC_NAME_KEYWORDS)
        mask_name = _name.str.contains(_kw_pattern, case=False, regex=True, na=False)
    else:
        mask_name = pd.Series(False, index=df.index)

    # ── ⑦ 과세상태 비과세 기반 ──────────────────────────────────────────
    # ※ 정산용 코드 필터는 KT 정산 실적을 손상시키므로 여기서 제거.
    #    수요예측 시계열 단계(group_grp)에서만 제외 처리함.
    # 유가증권은 부가세 비과세 대상 — 과세상태 컬럼이 있으면 보조 필터로 활용
    _TAX_COL = (col_map.get("tax")
                if col_map.get("tax") and col_map["tax"] in df.columns
                else ("과세상태" if "과세상태" in df.columns else None))
    if _TAX_COL:
        _tax = df[_TAX_COL].astype(str).str.strip()
        mask_tax_exempt = _tax.isin(["비과세", "면세", "영세", "비과세(유가증권)"])
    else:
        mask_tax_exempt = pd.Series(False, index=df.index)

    # ── 통합 마스크 ────────────────────────────────────────────────────
    # 카테고리 OR 상품명키워드 OR (비과세 AND 상품명키워드)
    _sec_mask = mask_cms | mask_svc | mask_name | (mask_tax_exempt & mask_name)

    # ── 감지 기준 문자열 구성 ─────────────────────────────────────────
    _basis_parts = []
    if mask_cms.any():
        _basis_parts.append(_CMS_COL)
    if mask_svc.any() and _SVC_COL:
        _basis_parts.append(_SVC_COL)
    if mask_name.any() and _NAME_COL:
        _basis_parts.append(f"상품명키워드({mask_name.sum():,}건)")
    if mask_tax_exempt.any() and _TAX_COL:
        _basis_parts.append(f"비과세({mask_tax_exempt.sum():,}건)")
    if not _basis_parts:
        # 카테고리 컬럼 없음 → 협력사명 fallback (데이터 없는 환경 대비)
        if "_vendor" in df.columns:
            _sec_mask = (df["_vendor"].astype(str)
                         .str.replace(" ", "", regex=False)
                         .str.lower()
                         .str.contains("케이티커머스|kt커머스", regex=True, na=False))
            _basis_parts.append("협력사명 fallback")
    _detect_basis = " + ".join(_basis_parts) if _basis_parts else "감지 불가"

    _sec_count = int(_sec_mask.sum())

    if _sec_count > 0:
        _sec_amt   = df.loc[_sec_mask, _amt_col].sum()
        _total_amt = df[_amt_col].sum()
        _sec_pct   = _sec_amt / _total_amt * 100 if _total_amt else 0

        # 감지 내역 요약: CMS → 서비스카테고리 순으로 상위 5개
        _ref_col = _CMS_COL if (_CMS_COL in df.columns and mask_cms.any()) else _SVC_COL
        if _ref_col and _ref_col in df.columns:
            _top5    = df.loc[_sec_mask, _ref_col].value_counts().head(5).index.tolist()
            _summary = " / ".join(str(v) for v in _top5)
        else:
            _summary = "(카테고리 컬럼 없음)"

        print("\n" + "─" * 65)
        print("  [분석 범위 선택]")
        print(f"  유가증권 감지 : {_sec_count:,}행  ({_sec_pct:.1f}%,  {_sec_amt:,.0f}원)")
        print(f"  감지 기준     : {_detect_basis}")
        print(f"  주요 항목     : {_summary}")
        print()
        print("  1. 전체 포함  — 유가증권 포함 전체 분석 (현행 방식)")
        print("  2. 유가증권 제외 — 순수 자재/서비스만 분석")
        print("  3. 통신자재만 — _prod_type='통신' 품목만 분석")
        print("─" * 65)
        _scope_sel = input("  선택 (기본값=1): ").strip()

        if _scope_sel == "2":
            df = df.loc[~_sec_mask].reset_index(drop=True)
            _scope_tag = "유가증권제외"
            print(f"  → 유가증권 {_sec_count:,}행 제외 완료 ({len(df):,}행 잔류)")
            log.info(f"분석범위: 유가증권 제외 {_sec_count:,}건 ({_detect_basis} 기준)")
        elif _scope_sel == "3":
            if "_prod_type" in df.columns:
                df = df.loc[df["_prod_type"] == "통신"].reset_index(drop=True)
                _scope_tag = "통신자재"
                print(f"  → 통신 품목만 필터링 완료 ({len(df):,}행 잔류)")
                log.info("분석범위: _prod_type='통신' 필터 적용")
            else:
                print("  → 통신/일반 구분 정보 없음, 전체로 진행합니다.")
        else:
            print("  → 전체 포함으로 진행합니다.")

    # ── 이관 시점 필터 ────────────────────────────────────────────────────
    # TRANSFER_DATE 설정 시 이관 후 데이터만 예측 학습에 사용
    # (KT SCM 이관 품목: 이관 전 수요 패턴은 다른 운영환경 기준이므로 제외)
    _df_forecast = df.copy()
    if TRANSFER_DATE:
        _transfer_dt = pd.to_datetime(TRANSFER_DATE)
        _date_col = col_map.get("order_date") or col_map.get("date") or None
        if _date_col and _date_col in df.columns:
            _before = (pd.to_datetime(df[_date_col], errors="coerce") < _transfer_dt)
            n_before = int(_before.sum())
            _df_forecast = df.loc[~_before].reset_index(drop=True)
            print(f"  [이관필터] {TRANSFER_DATE} 이전 {n_before:,}건 예측 학습에서 제외")
            log.info(f"이관 시점 필터: {TRANSFER_DATE} 이전 {n_before:,}건 제외 (실적 집계는 유지)")
        else:
            log.warning(f"이관 시점 필터: 날짜 컬럼 매핑 없음 — 전체 기간 학습")
    else:
        log.info("이관 시점 필터: 미설정 (전체 기간 학습)")

    # ── 발주일 기준 수요 집계 (수량 예측용) ─────────────────────────────
    current_ym     = pd.Period(datetime.now(), freq="M")
    order_cutoff   = current_ym - ORDER_LAG_MONTHS   # 완성된 발주 데이터 상한

    overall_grp = build_monthly_series(_df_forecast, ym_col="_ym")
    overall_grp = overall_grp.loc[overall_grp["_ym"] < order_cutoff].reset_index(drop=True)  # 미완성 월 제외
    overall_grp = trim_to_36months(overall_grp, "_ym")
    overall_grp.sort_values("_ym", inplace=True)

    if ORDER_LAG_MONTHS > 0:
        log.info(f"발주 완성 기준월: ~{order_cutoff - 1}  "
                 f"(당월 {current_ym} 포함 {ORDER_LAG_MONTHS + 1}개월은 예측 대상)")

    min_ym = overall_grp["_ym"].min()
    max_ym = overall_grp["_ym"].max()
    full_idx = pd.period_range(min_ym, max_ym, freq="M")
    overall_grp = overall_grp.set_index("_ym").reindex(full_idx, fill_value=0).reset_index()
    overall_grp.rename(columns={"index": "_ym"}, inplace=True)

    log.info(f"전체 분석기간(발주일): {min_ym} ~ {max_ym} ({len(overall_grp)}개월)")

    # ── 정산일 기준 매출 집계 (금액 예측용) ─────────────────────────────
    # current_ym은 위 발주일 블록에서 이미 정의됨
    revenue_grp = build_monthly_series(df, ym_col="_settle_ym")
    # 실제 정산 데이터의 최신월을 확인해서 지연 여부를 동적으로 판단
    # 예: 6월 실행 시 5월 정산 데이터가 이미 있으면 settle_cutoff = 6월 (지연 0개월)
    #     5월 데이터가 없으면 SETTLE_LAG_MONTHS 설정값(1개월) 적용
    _settle_yms = df["_settle_ym"].dropna()
    # 이상값 제거: 현재 연도+2년 초과하는 정산일은 데이터 오류로 간주 (예: 2085-04)
    _valid_cutoff_ym = pd.Period(datetime.now().year + 2, freq="M")
    _settle_yms_valid = _settle_yms[_settle_yms <= _valid_cutoff_ym]
    if len(_settle_yms_valid) < len(_settle_yms):
        _n_bad = len(_settle_yms) - len(_settle_yms_valid)
        log.warning(f"정산일 이상값 {_n_bad:,}건 제외 (>{_valid_cutoff_ym}): {_settle_yms[_settle_yms > _valid_cutoff_ym].unique()[:5].tolist()}")
    _max_settle_ym = _settle_yms_valid.max() if len(_settle_yms_valid) > 0 else (current_ym - SETTLE_LAG_MONTHS - 1)
    _expected_last = current_ym - 1   # 당월 직전월
    if _max_settle_ym >= _expected_last:
        # 전월 정산 완료 — 지연 없음
        settle_cutoff = current_ym
        _actual_lag = 0
    else:
        # 아직 지연 중 — 상수 적용
        settle_cutoff = current_ym - SETTLE_LAG_MONTHS
        _actual_lag = SETTLE_LAG_MONTHS
    revenue_grp_train = revenue_grp.loc[
        (revenue_grp["_ym"] < settle_cutoff) & (revenue_grp["_ym"] <= _valid_cutoff_ym)
    ].reset_index(drop=True)
    revenue_grp_train = trim_to_36months(revenue_grp_train, "_ym")
    log.info(f"정산 데이터 최신월: {_max_settle_ym}  실제 지연: {_actual_lag}개월  "
             f"정산 완성 기준월: ~{settle_cutoff - 1}")
    revenue_grp_train.sort_values("_ym", inplace=True)

    if len(revenue_grp_train) >= 3:
        rev_min_ym = revenue_grp_train["_ym"].min()
        rev_max_ym = revenue_grp_train["_ym"].max()
        rev_idx = pd.period_range(rev_min_ym, rev_max_ym, freq="M")
        revenue_grp_train = (revenue_grp_train.set_index("_ym")
                             .reindex(rev_idx, fill_value=0).reset_index())
        revenue_grp_train.rename(columns={"index": "_ym"}, inplace=True)
        log.info(f"전체 분석기간(정산일): {rev_min_ym} ~ {rev_max_ym} "
                 f"({len(revenue_grp_train)}개월)")
    else:
        revenue_grp_train = overall_grp.copy()   # fallback
        log.warning("정산일 기준 데이터 부족 → 발주일 기준으로 대체")

    # ── 전체 시계열 모델 평가 & 예측 (발주일 기준 수요) ──────────────────
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
        "name": "전체 수요(발주일)",
        "best": best_overall,
        "note": f"MAPE {best_mape:.1f}%" if not np.isnan(best_mape) else "",
    }
    for k, v in acc_results.items():
        overall_acc_row[k] = v["mape"]

    # ── 정산일 기준 매출 예측 ─────────────────────────────────────────────
    print("\n[매출예측] 정산일 기준 매출 예측 중...")
    rev_arr = revenue_grp_train["amount"].values.astype(float)
    rev_max_ym = revenue_grp_train["_ym"].max()
    # 실제 지연 개월 포함해서 예측 (미완성 달부터 6개월)
    rev_future_periods = pd.period_range(settle_cutoff, periods=6 + _actual_lag,
                                         freq="M")

    rev_acc, rev_best = walk_forward_eval(rev_arr, "Holt-Winters")
    rev_mape = rev_acc.get(rev_best, {}).get("mape", np.nan)
    if not np.isnan(rev_mape) and rev_mape > 50:
        rev_arr_s = rev_arr[-24:] if len(rev_arr) >= 24 else rev_arr
        rev_acc_s, rev_best_s = walk_forward_eval(rev_arr_s, "Holt-Winters")
        rev_mape_s = rev_acc_s.get(rev_best_s, {}).get("mape", np.nan)
        if not np.isnan(rev_mape_s) and rev_mape_s < rev_mape:
            rev_arr  = rev_arr_s
            rev_best = rev_best_s
            rev_mape = rev_mape_s
            rev_acc  = rev_acc_s

    # 예측: rev_future_periods 전체를 커버하도록 h_rev 계산
    # rev_future_periods[-1] 까지 포함하려면 그 기간만큼 예측 필요
    _last_need = rev_future_periods[-1]
    h_rev = (_last_need.ordinal - rev_max_ym.ordinal)   # 학습마지막 다음달~필요마지막달
    h_rev = max(h_rev, 6)
    log.info(f"  [매출예측 진단] rev_max_ym={rev_max_ym}  settle_cutoff={settle_cutoff}  "
             f"future_periods={rev_future_periods[0]}~{rev_future_periods[-1]}  h_rev={h_rev}")
    rev_fc, rev_ci_lo, rev_ci_hi = forecast_series(rev_arr, rev_best, h=int(h_rev))
    # rev_future_periods 에 해당하는 인덱스 추출
    rev_fc_periods = pd.period_range(rev_max_ym + 1, periods=int(h_rev), freq="M")
    rev_fc_map = {p: (rev_fc[i], rev_ci_lo[i], rev_ci_hi[i])
                  for i, p in enumerate(rev_fc_periods)}
    # 매칭 누락 경고
    _missed = [p for p in rev_future_periods if p not in rev_fc_map]
    if _missed:
        log.warning(f"  [매출예측] rev_fc_map 미매칭 기간 {_missed} → 0원 처리됨")

    rev_fc_vals  = np.array([rev_fc_map.get(p, (0, 0, 0))[0] for p in rev_future_periods])
    rev_ci_lo_v  = np.array([rev_fc_map.get(p, (0, 0, 0))[1] for p in rev_future_periods])
    rev_ci_hi_v  = np.array([rev_fc_map.get(p, (0, 0, 0))[2] for p in rev_future_periods])

    log.info(f"매출예측(정산일): {rev_best} MAPE={rev_mape:.1f}%")
    log.info(f"  예측기간: {rev_future_periods[0]} ~ {rev_future_periods[-1]}")
    for p, v in zip(rev_future_periods, rev_fc_vals):
        log.info(f"  {p}: {v:,.0f}원")

    # ── 협력사×상품코드 집계 ──────────────────────────────────────────────
    print("[STEP 1] 협력사×상품코드 월별 집계 중...")

    # 정산용 코드는 수요예측 시계열에서만 제외 (KT 정산 실적 집계에는 영향 없음)
    # 이관 필터도 적용된 _df_forecast 기반으로 시작
    _df_fc = _df_forecast.copy()
    _name_col_fc = col_map.get("prod_name") if col_map.get("prod_name") and col_map["prod_name"] in df.columns \
                   else ("상품명" if "상품명" in df.columns else None)
    if _name_col_fc:
        _settle_mask = _df_fc[_name_col_fc].astype(str).str.contains("정산용", na=False)
        n_settle = int(_settle_mask.sum())
        if n_settle > 0:
            _df_fc = _df_fc.loc[~_settle_mask].reset_index(drop=True)
            print(f"  → 정산용 코드 {n_settle:,}건 수요예측에서 제외 (KT 실적 집계는 유지)")
            log.info(f"정산용 코드 수요예측 제외: {n_settle:,}건")

    group_grp = build_monthly_series(_df_fc, group_cols=["_channel", "_vendor", "_prod_key"])
    group_grp = group_grp.loc[group_grp["_ym"] < order_cutoff].reset_index(drop=True)  # 미완성 월 제외
    group_grp = trim_to_36months(group_grp, "_ym")

    # ── ABC × CV ─────────────────────────────────────────────────────────
    print("[STEP 2] ABC × CV 분류 중...")
    abc_df = classify_abc_cv(group_grp)
    log.info(f"ABC 분류: A={len(abc_df[abc_df.ABC=='A'])} B={len(abc_df[abc_df.ABC=='B'])} C={len(abc_df[abc_df.ABC=='C'])}")

    # ── 리드타임 통계 (원본 데이터에서 실측) ────────────────────────────
    lt_stats, lt_order_lookup = build_lt_stats(df)
    olfr_df  = calc_olfr(df)
    if not olfr_df.empty:
        log.info(f"OLFR 계산 완료: {len(olfr_df)}개 vendor×품목 "
                 f"(평균 {olfr_df['olfr'].mean():.1f}%)")

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
    # 품목별 수량 배열 사전 구성 (O(1) 조회로 속도 개선)
    print("[SS/ROP] 안전재고·ROP 계산 중...")
    grp_qty = {
        (ch, v, p): g.sort_values("_ym")["qty"].values.astype(float)
        for (ch, v, p), g in group_grp.groupby(["_channel", "_vendor", "_prod_key"], observed=True)
    }

    ss_rop = {}
    # C등급도 최소 1개 보유 룰 적용을 위해 전체 포함
    for _, row in abc_df.iterrows():
        key = (row["_vendor"], row["_prod_key"])
        ch  = row.get("_channel", "기타")
        abc_g = row.get("ABC", "B")
        lt_m, lt_s = get_lt_stats(lt_stats, lt_df, row["_vendor"], row["_prod_key"],
                                   lt_order_lookup=lt_order_lookup)
        ofr_val    = get_ofr(lt_stats,  row["_vendor"], row["_prod_key"])
        olfr_val   = get_olfr(olfr_df,  row["_vendor"], row["_prod_key"])
        qty_arr = grp_qty.get((ch, row["_vendor"], row["_prod_key"]), np.array([]))
        if len(qty_arr) >= 2:
            ss, formula = calc_safety_stock(qty_arr, lt_m, lt_s, abc_grade=abc_g)
            rop = calc_rop(qty_arr, lt_m, ss)
        else:
            ss, rop, formula = 0.0, 0.0, "데이터부족"
        ss_rop[key] = {"ss": ss, "rop": rop, "lead_time": lt_m,
                       "lt_std": lt_s, "formula": formula,
                       "ofr": ofr_val, "olfr": olfr_val}
    print(f"  → SS/ROP 완료: {len(ss_rop):,}개 품목")

    # SS/ROP 값을 abc_df에 병합 (협력사_재고가이드 시트용)
    abc_df["_ss"]  = abc_df.apply(lambda r: ss_rop.get((r["_vendor"], r["_prod_key"]), {}).get("ss",  0.0), axis=1)
    abc_df["_rop"] = abc_df.apply(lambda r: ss_rop.get((r["_vendor"], r["_prod_key"]), {}).get("rop", 0.0), axis=1)

    # 수량 기반 월평균 (협력사_재고가이드의 ML월평균수요 열)
    _qty_avg = {}
    for (ch, v, p), g in group_grp.groupby(["_channel", "_vendor", "_prod_key"], observed=True):
        arr = g.sort_values("_ym")["qty"].values.astype(float)
        window = arr[-12:] if len(arr) >= 12 else arr
        _qty_avg[(v, p)] = float(np.mean(window)) if len(window) > 0 else 0.0
    abc_df["avg_qty"] = abc_df.apply(lambda r: _qty_avg.get((r["_vendor"], r["_prod_key"]), 0.0), axis=1)

    # ── 상품코드별 통신/일반 구분 사전 구성 ──────────────────────────────
    if "_prod_type" in df.columns:
        prod_type_lookup = (
            df[["_prod_key", "_prod_type"]]
            .drop_duplicates(subset=["_prod_key"])
            .set_index("_prod_key")["_prod_type"]
            .to_dict()
        )
    else:
        prod_type_lookup = {}

    # ── 상품명 + 규격 룩업 (Excel 출력 시 상품코드 옆에 표시) ──────────
    _name_col = col_map.get("prod_name")
    _spec_col  = col_map.get("spec")
    _info_cols = ["_prod_key"]
    if _name_col and _name_col in df.columns:
        _info_cols.append(_name_col)
    if _spec_col and _spec_col in df.columns:
        _info_cols.append(_spec_col)

    prod_info_lookup: dict = {}   # prod_key → {"name": str, "spec": str}
    if len(_info_cols) > 1:
        for _, row in df[_info_cols].drop_duplicates("_prod_key").iterrows():
            nm = str(row[_name_col]).strip() if _name_col in _info_cols else ""
            sp = str(row[_spec_col]).strip() if _spec_col in _info_cols else ""
            prod_info_lookup[row["_prod_key"]] = {
                "name": "" if nm in ("nan", "None") else nm,
                "spec": "" if sp in ("nan", "None") else sp,
            }
    log.info(f"상품명 룩업: {len(prod_info_lookup)}개 품목")

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
        channel  = row.get("_channel", None)
        abc      = row["ABC"]
        strategy = row["strategy"]

        mask = (group_grp["_vendor"] == vendor) & (group_grp["_prod_key"] == prod_key)
        if channel is not None and "_channel" in group_grp.columns:
            mask &= (group_grp["_channel"] == channel)
        sub = group_grp[mask].sort_values("_ym")
        if len(sub) < 3:
            log.debug(f"건너뜀(데이터 부족): {vendor} / {prod_key}")
            continue

        # 결측 월 보간 (수치 컬럼만 0 채움, 문자열 컬럼은 ffill)
        sub_idx = pd.period_range(sub["_ym"].min(), sub["_ym"].max(), freq="M")
        sub = sub.set_index("_ym").reindex(sub_idx)
        num_cols = sub.select_dtypes(include="number").columns
        sub[num_cols] = sub[num_cols].fillna(0)
        str_cols = [c for c in sub.columns if c not in num_cols]
        if str_cols:
            sub[str_cols] = sub[str_cols].ffill().bfill()
        sub = sub.reset_index()
        sub.rename(columns={"index": "_ym"}, inplace=True)

        if check_consecutive_zeros(sub["qty"].values, threshold=3):
            log.info(f"제외(연속0): {vendor} / {prod_key}")
            continue

        arr = sub["qty"].values.astype(float)

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

            # Tracking Signal 계산 (walk-forward 검증 구간 사용)
            n_test_ts = min(6, max(1, len(arr) // 4))
            ts_train, ts_test = arr[:-n_test_ts], arr[-n_test_ts:]
            try:
                ts_fc, _ = forecast_series(ts_train, best_item, h=n_test_ts)[:2], None
                ts_fc = ts_fc[0]
                ts_val = _tracking_signal(ts_test, ts_fc)
            except Exception:
                ts_val = np.nan

            avg12 = float(np.mean(arr[-12:])) if len(arr) >= 12 else float(np.mean(arr))
            _pinfo = prod_info_lookup.get(prod_key, {})

            # 수명주기 감지 (MRO 이관 품목용)
            _base_ym = sub["_ym"].max() if "_ym" in sub.columns else None
            lc = detect_lifecycle(arr, base_ym=_base_ym)

            detail_rows.append({
                "vendor":          vendor,
                "prod_key":        prod_key,
                "prod_name":       _pinfo.get("name", ""),
                "prod_spec":       _pinfo.get("spec", ""),
                "ABC":             abc,
                "avg12":           avg12,
                "fc6":             fc6,
                "model":           best_item,
                "mape":            item_mape,
                "prod_type":       prod_type_lookup.get(prod_key, "미분류"),
                "tracking_signal": ts_val,
                "lifecycle_stage": lc["stage"],
                "eol_ym":          str(lc["eol_ym"]) if lc["eol_ym"] else "",
                "decline_pct":     lc["decline_pct"],
                "lc_note":         lc["note"],
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
    # 극단 이상치 제외 후 중위값 사용 (단순 평균은 MAPE 300% 품목에 왜곡됨)
    valid_mapes = [v for v in all_mapes if v <= 300]
    avg_mape = float(np.median(valid_mapes)) if valid_mapes else np.nan

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
    _scope_suffix = f"_{_scope_tag}" if _scope_tag != "전체" else ""
    _base_name = f"수요예측_{datetime.now().strftime('%Y%m%d')}{_scope_suffix}"
    out_path = BASE_DIR / f"{_base_name}.xlsx"
    # 파일이 열려 있거나 OneDrive 잠금 시 _1, _2 ... 순번 파일명으로 우회
    if out_path.exists():
        _suffix = 1
        while out_path.exists():
            try:
                out_path.rename(out_path)   # 잠금 여부 테스트 (이름 변경 없이 시도)
                break                       # 잠금 없음 → 덮어쓰기 가능
            except PermissionError:
                out_path = BASE_DIR / f"{_base_name}_{_suffix}.xlsx"
                _suffix += 1
    out_name = out_path.name
    print(f"\n[출력] {out_name} 작성 중...")

    import openpyxl
    wb = openpyxl.Workbook()
    del wb["Sheet"]

    # ── abc_df 상품명/규격 보완 ──────────────────────────────────────────
    if "_prod_name" not in abc_df.columns:
        abc_df["_prod_name"] = abc_df["_prod_key"].map(
            lambda k: prod_info_lookup.get(k, {}).get("name", ""))
    if "_prod_spec" not in abc_df.columns:
        abc_df["_prod_spec"] = abc_df["_prod_key"].map(
            lambda k: prod_info_lookup.get(k, {}).get("spec", ""))

    # ── KT 정합도 검증 로그 (항상 출력 — 시트 없어도 콘솔에서 확인) ────────
    log_channel_verification(df, settle_cutoff, n_months=6)

    # ── 시트 출력 순서 ────────────────────────────────────────────────────
    # [핵심] 협력사 대상 시트 (앞쪽 배치, 탭 녹색)
    write_sheet_glossary(wb)                                                # 1. 용어설명 (첫 번째)
    write_sheet_supplier_guide(wb, abc_df, ss_rop, df,
                               future_periods, purchase_plan)               # 2. 협력사_재고가이드
    write_sheet_supplier_alerts(wb, abc_df, ss_rop, detail_rows,
                                future_periods, df_raw=df)                  # 3. 협력사_알림보드

    # [검증] 신뢰성 근거 시트 (중간 배치, 탭 파란색)
    write_sheet3_detail(wb, detail_rows)                                    # 4. 협력사별_상품코드_예측
    write_sheet_forecast_detail(wb, detail_rows, ss_rop)                   # 5. 예측근거_상품분석
    write_sheet_tracking_signal(wb, detail_rows)                           # 6. 예측편향_TrackingSignal
    write_sheet_supplier_risk(wb, df, abc_df)                              # 7. 공급업체_리스크
    write_sheet2_abc(wb, abc_df, ss_rop)                                   # 8. ABC_CV_분류표
    write_sheet_xai(wb, acc_rows)                                          # 9. XAI_모델선택근거
    write_sheet_ecos(wb, ecos_data)                                        # 10. 외부지표_ECOS (API 키 있을 때만)
    write_sheet_qbr_savings(wb, detail_rows, ss_rop, df)                   # 11. QBR_안전재고절감분석

    # [내부용] 제거된 시트 목록 (호출 안 함):
    #   write_sheet5_dashboard  → 기술 KPI 대시보드 (초보자 무의미)
    #   write_sheet1_overall    → 전체_월별예측 (내부 분석용)
    #   write_sheet_revenue     → 매출예측_정산일기준 (내부 정산용)
    #   write_sheet_channel     → 채널별_분석 (내부 집계용)
    #   write_sheet_spend_analysis → 카테고리_Spend분석 (무관)
    #   write_sheet4_accuracy   → 모델_정확도_비교 (기술 지표)

    # 저장 시 PermissionError(파일 열림/OneDrive 잠금) 처리
    _save_suffix = 1
    while True:
        try:
            wb.save(out_path)
            log.info(f"저장 완료: {out_path}")
            break
        except PermissionError:
            _base = f"수요예측_{datetime.now().strftime('%Y%m')}"
            out_path = BASE_DIR / f"{_base}_{_save_suffix}.xlsx"
            out_name = out_path.name
            print(f"  [주의] 파일 잠김 → {out_name} 으로 저장 시도...")
            _save_suffix += 1
            if _save_suffix > 9:
                print("  [오류] 저장 파일명을 확보할 수 없습니다. Excel을 닫고 재시도하세요.")
                raise

    # ── 구매계획 입력 템플릿 자동 생성 (최초 1회) ────────────────────────
    tpl_plan_path = BASE_DIR / "구매계획_입력.xlsx"
    if not tpl_plan_path.exists():
        try:
            # abc_df에 _prod_name/_prod_spec 보장
            if "_prod_name" not in abc_df.columns:
                abc_df["_prod_name"] = abc_df["_prod_key"].map(
                    lambda k: prod_info_lookup.get(k, {}).get("name", ""))
            if "_prod_spec" not in abc_df.columns:
                abc_df["_prod_spec"] = abc_df["_prod_key"].map(
                    lambda k: prod_info_lookup.get(k, {}).get("spec", ""))
            make_purchase_plan_template(abc_df, future_periods, tpl_plan_path)
            print(f"\n[안내] 구매계획 입력 템플릿 생성: {tpl_plan_path.name}")
            print("       담당자가 노란색 셀에 품목별 구매 예정 수량을 입력하면")
            print("       다음 실행 시 '협력사_재고가이드' 시트에 자동 반영됩니다.")
        except Exception as _e:
            log.warning(f"구매계획 템플릿 생성 실패: {_e}")

    # ── 최종 요약 출력 ────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print(f"  [분석 범위: {_scope_tag}]")
    print(f"  [발주일 기준 수요]")
    print(f"  분석 기간          : {summary['data_range']}")
    print(f"  예측 모델          : {best_overall}  (MAPE {best_mape:.1f}%)")
    print(f"  A등급 품목 수      : {summary['n_A']}개")
    print(f"  주의 시계열 수     : {len(warn_list)}개 (MAPE > 30%)")
    print()
    print(f"  [정산일 기준 매출]")
    print(f"  예측 모델          : {rev_best}  (MAPE {rev_mape:.1f}%)")
    for p, v in zip(rev_future_periods, rev_fc_vals):
        if p < current_ym:
            flag = " ← 정산지연(미완성 예측)"
        elif p == current_ym:
            flag = " ← 이번 달"
        else:
            flag = ""
        print(f"  {p}  : {v:>20,.0f} 원{flag}")
    print()
    print(f"  출력 파일          : {out_path.name}")
    print("=" * 65)

    if sys.platform == "win32":
        try:
            os.startfile(out_path)
        except OSError:
            pass


if __name__ == "__main__":
    main()
