# -*- coding: utf-8 -*-
"""
Phase 1 재학습 전 진단: 2~5일 구간 과소신(underconfidence) 원인 규명
1. 연도별 CDF 비교 (2023/2024/2025/2026H1)
2. 월별 "3일 이내 도달률" 트렌드 + 변곡점
3. 우측 절단(censoring) 검증: 2026 데이터 추출일(6/25) 근접 주문의 결측 편향 여부
4. 세그먼트별 개선폭 top/bottom (2023 vs 2025, vendor_cat 레벨)
"""
import pandas as pd
import numpy as np
from pathlib import Path
from data_loader import load_many, SEGMENTS
import importlib.util
spec = importlib.util.spec_from_file_location("phase1", Path(__file__).parent / "03_distribution_fit.py")
phase1 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(phase1)

OUT_DIR = Path(__file__).parent
DAY_MARKS = [1, 2, 3, 5, 7, 10, 14]


def main():
    lines = ["배송 예측 - 재학습 전 진단: 2~5일 과소신 원인 규명"]

    print("전체 연도 데이터 로딩 (캐시 활용, 빠름)...", flush=True)
    df = load_many(SEGMENTS, [2023, 2024, 2025, 2026])
    d = phase1.prep(df)
    d["연도"] = d["주문일자"].dt.year
    d["연월"] = d["주문일자"].dt.to_period("M")
    print(f"  전체 valid rows={len(d):,}", flush=True)

    # 1. 연도별 CDF 비교
    lines.append("\n" + "=" * 70)
    lines.append("[1] 연도별 CDF 비교 (전체)")
    lines.append("=" * 70)
    lines.append(f"{'연도':>6s}" + "".join(f"{d_:>8}일" for d_ in DAY_MARKS) + f"{'n':>12s}")
    for yr, g in d.groupby("연도"):
        row = f"{yr:>6d}"
        for dm in DAY_MARKS:
            row += f"{(g['lt']<=dm).mean()*100:>8.1f}%"
        row += f"{len(g):>12,}"
        lines.append(row)

    # 2. 월별 3일 이내 도달률 트렌드
    lines.append("\n" + "=" * 70)
    lines.append("[2] 월별 '3일 이내 도달률' 트렌드")
    lines.append("=" * 70)
    monthly = d.groupby("연월").agg(
        rate_3d=("lt", lambda s: (s <= 3).mean()),
        rate_5d=("lt", lambda s: (s <= 5).mean()),
        n=("lt", "size"),
    )
    for ym, row in monthly.iterrows():
        lines.append(f"  {str(ym)}  3일도달률={row['rate_3d']*100:5.1f}%  5일도달률={row['rate_5d']*100:5.1f}%  n={int(row['n']):>7,}")

    # 변곡점 탐지: 3일도달률 전월 대비 변화폭 top5
    monthly["diff_3d"] = monthly["rate_3d"].diff() * 100
    lines.append("\n-- 전월 대비 3일도달률 변화폭 top5 (변곡점 후보) --")
    top_diff = monthly["diff_3d"].abs().sort_values(ascending=False).head(5)
    for ym in top_diff.index:
        lines.append(f"  {ym}: {monthly.loc[ym,'diff_3d']:+.1f}%p")

    # 3. 우측 절단(censoring) 검증
    lines.append("\n" + "=" * 70)
    lines.append("[3] 우측 절단(censoring) 검증 - 2026H1 파일 추출일(6/25) 근접 편향")
    lines.append("=" * 70)
    h2026_raw = df[(df["_year"] == 2026)]
    max_order_date = h2026_raw["주문일자"].max()
    lines.append(f"2026 파일 내 최대 주문일자: {max_order_date.date()} (파일 mtime 6/25와 근접 확인)")

    # 원본(필터 전) 기준으로 결측률 비교: 최근 14일 vs 그 이전
    recent_mask = h2026_raw["주문일자"] >= max_order_date - pd.Timedelta(days=14)
    older_mask = (h2026_raw["주문일자"] < max_order_date - pd.Timedelta(days=14)) & (h2026_raw["주문일자"] >= max_order_date - pd.Timedelta(days=60))
    recent_missing = h2026_raw[recent_mask]["배송완료일자"].isna().mean() * 100
    older_missing = h2026_raw[older_mask]["배송완료일자"].isna().mean() * 100
    lines.append(f"  최근 14일 주문의 배송완료일자 결측률: {recent_missing:.1f}%")
    lines.append(f"  15~60일 전 주문의 배송완료일자 결측률: {older_missing:.1f}%")
    lines.append("  (해석: 최근 결측률이 훨씬 높으면 censoring 존재 -> 최근 몇 주는 '완료된 건만' 남아 편향)")

    # censoring 영향 제외 버전: 2026 H1에서 최근 21일 제외하고 재계산
    d2026 = d[d["연도"] == 2026]
    cutoff = max_order_date - pd.Timedelta(days=21)
    d2026_safe = d2026[d2026["주문일자"] <= cutoff]
    d2026_recent = d2026[d2026["주문일자"] > cutoff]
    lines.append(f"\n  censoring 제외(주문일자<={cutoff.date()}) 2026 CDF vs 최근 21일 포함분 비교:")
    lines.append(f"  {'구간':>10s}" + "".join(f"{dm:>8}일" for dm in DAY_MARKS) + f"{'n':>10s}")
    row = f"{'안전구간':>10s}"
    for dm in DAY_MARKS:
        row += f"{(d2026_safe['lt']<=dm).mean()*100:>8.1f}%"
    row += f"{len(d2026_safe):>10,}"
    lines.append(row)
    row = f"{'최근21일':>10s}"
    for dm in DAY_MARKS:
        row += f"{(d2026_recent['lt']<=dm).mean()*100:>8.1f}%"
    row += f"{len(d2026_recent):>10,}"
    lines.append(row)

    # 2023 vs "안전구간만 사용한 2026" 비교 (censoring 보정 후에도 격차가 남는지)
    d2023 = d[d["연도"] == 2023]
    lines.append(f"\n  2023 전체 vs 2026-안전구간 CDF (censoring 보정 후 진짜 격차 확인):")
    row = f"{'2023':>10s}"
    for dm in DAY_MARKS:
        row += f"{(d2023['lt']<=dm).mean()*100:>8.1f}%"
    row += f"{len(d2023):>10,}"
    lines.append(row)
    row = f"{'2026안전':>10s}"
    for dm in DAY_MARKS:
        row += f"{(d2026_safe['lt']<=dm).mean()*100:>8.1f}%"
    row += f"{len(d2026_safe):>10,}"
    lines.append(row)

    # 4. 세그먼트별 개선폭 top/bottom (2023 vs 2025, vendor_cat, 5일 기준)
    lines.append("\n" + "=" * 70)
    lines.append("[4] 세그먼트별 개선폭 (2023 -> 2025, vendor_cat, 5일 도달률 기준, n>=50 각 연도)")
    lines.append("=" * 70)
    g23 = d[d["연도"] == 2023].groupby(["협력사명", "대분류"])["lt"].agg(rate5=lambda s: (s <= 5).mean(), n="count")
    g25 = d[d["연도"] == 2025].groupby(["협력사명", "대분류"])["lt"].agg(rate5=lambda s: (s <= 5).mean(), n="count")
    joined = g23.join(g25, lsuffix="_23", rsuffix="_25", how="inner")
    joined = joined[(joined["n_23"] >= 50) & (joined["n_25"] >= 50)]
    joined["improve"] = (joined["rate5_25"] - joined["rate5_23"]) * 100

    lines.append(f"\n비교 가능 세그먼트 수: {len(joined):,}")
    lines.append("\n-- 개선폭 top 15 (5일도달률 상승) --")
    for (vendor, cat), row in joined.sort_values("improve", ascending=False).head(15).iterrows():
        lines.append(f"  {str(vendor):20s} | {str(cat):16s} 2023={row['rate5_23']*100:5.1f}% -> 2025={row['rate5_25']*100:5.1f}%  (+{row['improve']:.1f}%p)  n23={int(row['n_23'])} n25={int(row['n_25'])}")

    lines.append("\n-- 개선폭 bottom 15 (악화/정체) --")
    for (vendor, cat), row in joined.sort_values("improve", ascending=True).head(15).iterrows():
        lines.append(f"  {str(vendor):20s} | {str(cat):16s} 2023={row['rate5_23']*100:5.1f}% -> 2025={row['rate5_25']*100:5.1f}%  ({row['improve']:+.1f}%p)  n23={int(row['n_23'])} n25={int(row['n_25'])}")

    report_path = OUT_DIR / "recency_diagnosis.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nDone. Report saved to {report_path}", flush=True)


if __name__ == "__main__":
    main()
