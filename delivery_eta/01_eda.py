# -*- coding: utf-8 -*-
"""
Phase 0: 배송 예측 데이터 적합성 EDA
원본 KT raw data(주문현황 엑셀, 109개 컬럼)를 대상으로
1) 배송 관련 핵심 필드의 결측률/값 형태 점검
2) 발주->출하, 출하->배송완료, 주문->배송완료 리드타임 분포 (협력사/상품군별)
3) 배송완료일자 vs 입고일자 vs 정산확정일 비교 (수령일 proxy 검증)
4) 상품배송리드타임(선언값) vs 실측 리드타임 비교
를 계산해 결과를 delivery_eta/eda_report.txt 로 저장한다.
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
import numpy as np
from pathlib import Path

DATA_DIR = Path(r"C:\Users\USER\OneDrive - KT Corporation\문서\수요예측")
OUT_DIR = Path(r"C:\Users\USER\OneDrive - KT Corporation\문서\배송예측\delivery_eta")

TARGET_FILES = {
    "KT_2026": "2026_주문현황_KT(요청자_이돈현).xlsx",
    "그룹사_2026": "2026_주문현황_그룹사(요청자_이돈현).xlsx",
    "외부사_2026": "2026_주문현황_외부사(요청자_이돈현).xlsx",
    "지입자재_2026": "2026_주문현황_지입자재(요청자_이돈현).xlsx",
}

USECOLS = [
    "주문번호", "주문일자", "발주일자", "출하일자", "배송완료일자",
    "배송예정일", "배송희망일", "표준납기일", "입고일자", "입고승인일자",
    "정산확정일", "배송지", "배송업체", "송장번호", "상품배송리드타임",
    "주문시배송리드타임", "배송상태", "배송형태", "배송유형", "상품타입",
    "상품유형", "협력사명", "서비스카테고리", "주문상태", "취소수량", "반품수량",
]

DATE_COLS = [
    "주문일자", "발주일자", "출하일자", "배송완료일자", "배송예정일",
    "배송희망일", "표준납기일", "입고일자", "입고승인일자", "정산확정일",
]


def parse_date_col(s: pd.Series) -> pd.Series:
    """str('YYYY-MM-DD') / int or float(YYYYMMDD) 혼재 컬럼을 datetime으로 통일 파싱."""
    s = s.copy()
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    is_num = pd.to_numeric(s, errors="coerce")
    num_mask = is_num.notna()
    if num_mask.any():
        yyyymmdd = is_num[num_mask].astype("Int64")
        valid = (yyyymmdd >= 19000101) & (yyyymmdd <= 20991231)
        parsed = pd.to_datetime(
            yyyymmdd[valid].astype(str), format="%Y%m%d", errors="coerce"
        )
        out.loc[parsed.index] = parsed
    str_mask = ~num_mask & s.notna()
    if str_mask.any():
        parsed = pd.to_datetime(s[str_mask], errors="coerce")
        out.loc[parsed.index] = parsed
    return out


def load_file(path: Path) -> pd.DataFrame:
    print(f"  reading {path.name} ...", flush=True)
    df = pd.read_excel(path, usecols=USECOLS, engine="openpyxl")
    for c in DATE_COLS:
        df[c] = parse_date_col(df[c])
    return df


def report_missing(df: pd.DataFrame, label: str, lines: list):
    lines.append(f"\n=== [{label}] 결측률 (n={len(df):,}) ===")
    for c in USECOLS:
        n_missing = df[c].isna().sum()
        pct = n_missing / len(df) * 100 if len(df) else 0
        lines.append(f"  {c:16s} 결측 {n_missing:>8,} / {len(df):>8,}  ({pct:5.1f}%)")


def report_leadtime(df: pd.DataFrame, label: str, lines: list):
    lines.append(f"\n=== [{label}] 리드타임 분포 (일 단위) ===")
    d = df.copy()
    d["lt_order_to_ship"] = (d["출하일자"] - d["발주일자"]).dt.days
    d["lt_ship_to_deliver"] = (d["배송완료일자"] - d["출하일자"]).dt.days
    d["lt_order_to_deliver"] = (d["배송완료일자"] - d["주문일자"]).dt.days
    d["lt_order_to_recv"] = (d["입고일자"] - d["주문일자"]).dt.days
    d["gap_deliver_vs_recv"] = (d["입고일자"] - d["배송완료일자"]).dt.days
    d["gap_deliver_vs_settle"] = (d["정산확정일"] - d["배송완료일자"]).dt.days

    for col in ["lt_order_to_ship", "lt_ship_to_deliver", "lt_order_to_deliver",
                "lt_order_to_recv", "gap_deliver_vs_recv", "gap_deliver_vs_settle"]:
        s = d[col].dropna()
        s_valid = s[(s >= -5) & (s <= 120)]  # 이상치 제거 후 요약(참고용)
        n_neg = (s < 0).sum()
        n_extreme = (s > 120).sum()
        lines.append(
            f"  {col:24s} n={len(s):>8,} (음수 {n_neg:>6,}, >120일 {n_extreme:>5,})  "
            f"median={s_valid.median():>5.1f}  p90={s_valid.quantile(0.9):>5.1f}  "
            f"mean={s_valid.mean():>5.1f}"
        )
    return d


def report_declared_vs_actual(d: pd.DataFrame, label: str, lines: list):
    lines.append(f"\n=== [{label}] 상품배송리드타임(선언값) vs 실측(출하->배송완료) ===")
    sub = d.dropna(subset=["상품배송리드타임", "lt_ship_to_deliver"])
    sub = sub[(sub["lt_ship_to_deliver"] >= -5) & (sub["lt_ship_to_deliver"] <= 120)]
    if len(sub) == 0:
        lines.append("  비교 가능한 행 없음 (상품배송리드타임 결측 다수)")
        return
    diff = sub["lt_ship_to_deliver"] - sub["상품배송리드타임"]
    lines.append(f"  비교가능 n={len(sub):,}")
    lines.append(f"  선언값 median={sub['상품배송리드타임'].median():.1f}  실측 median={sub['lt_ship_to_deliver'].median():.1f}")
    lines.append(f"  차이(실측-선언) median={diff.median():.1f}  mean={diff.mean():.1f}  std={diff.std():.1f}")


def report_categorical(df: pd.DataFrame, label: str, lines: list):
    lines.append(f"\n=== [{label}] 배송업체 분포 (top 10) ===")
    vc = df["배송업체"].value_counts(dropna=False).head(10)
    for k, v in vc.items():
        lines.append(f"  {str(k):20s} {v:>8,} ({v/len(df)*100:4.1f}%)")

    lines.append(f"\n=== [{label}] 배송형태/배송유형 분포 (top 10) ===")
    for col in ["배송형태", "배송유형", "상품타입"]:
        lines.append(f"  -- {col} --")
        vc = df[col].value_counts(dropna=False).head(10)
        for k, v in vc.items():
            lines.append(f"    {str(k):20s} {v:>8,} ({v/len(df)*100:4.1f}%)")

    lines.append(f"\n=== [{label}] 배송지 샘플 (형태 확인용, 20건) ===")
    for v in df["배송지"].dropna().head(20):
        lines.append(f"  {v}")

    lines.append(f"\n=== [{label}] 송장번호 샘플 & 유효성 (배송업체별 20건) ===")
    for carrier, grp in df.groupby("배송업체"):
        sample = grp["송장번호"].dropna().head(3).tolist()
        lines.append(f"  {str(carrier):20s} 송장번호 샘플: {sample}")


def report_vendor_category(d: pd.DataFrame, label: str, lines: list):
    lines.append(f"\n=== [{label}] 협력사 x 상품 대분류별 lt_ship_to_deliver 중앙값 (건수>=30) ===")
    d = d.copy()
    d["대분류"] = d["서비스카테고리"].fillna("").str.split(">").str[0].str.strip()
    valid = d.dropna(subset=["lt_ship_to_deliver"])
    valid = valid[(valid["lt_ship_to_deliver"] >= -5) & (valid["lt_ship_to_deliver"] <= 120)]
    g = valid.groupby(["협력사명", "대분류"])["lt_ship_to_deliver"].agg(["median", "count"])
    g = g[g["count"] >= 30].sort_values("count", ascending=False).head(30)
    for (vendor, cat), row in g.iterrows():
        lines.append(f"  {str(vendor):20s} | {str(cat):20s} median={row['median']:5.1f}  n={int(row['count']):>6,}")


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    lines = []
    lines.append("배송 예측 정보 과제 - Phase 0 데이터 적합성 EDA")
    lines.append(f"대상 파일: {list(TARGET_FILES.values())}")

    for label, fname in TARGET_FILES.items():
        path = DATA_DIR / fname
        if not path.exists():
            lines.append(f"\n[!] 파일 없음: {path}")
            continue
        df = load_file(path)
        report_missing(df, label, lines)
        d = report_leadtime(df, label, lines)
        report_declared_vs_actual(d, label, lines)
        report_categorical(df, label, lines)
        report_vendor_category(d, label, lines)

    report_path = OUT_DIR / "eda_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nDone. Report saved to {report_path}")


if __name__ == "__main__":
    main()
