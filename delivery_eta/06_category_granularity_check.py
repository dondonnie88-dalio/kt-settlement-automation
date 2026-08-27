# -*- coding: utf-8 -*-
"""
중분류 세분화 가치 검증 + 관리회계 vs 서비스카테고리 중복도 확인
"""
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent


def main():
    df = load_many(SEGMENTS, [2023, 2024, 2025, 2026])
    df["lt"] = (df["배송완료일자"] - df["주문일자"]).dt.days
    df = df[df["lt"].notna() & (df["lt"] >= 0) & (df["lt"] <= 90)]
    parts = df["서비스카테고리"].fillna("").str.split(">")
    df["대분류"] = parts.str[0].str.strip()
    df["중분류"] = parts.str[1].str.strip()

    lines = ["중분류 세분화 검증 리포트"]

    # 1. 물량 top 15 협력사에서 대분류 vs 중분류 비교
    top_vendors = df.groupby("협력사명").size().sort_values(ascending=False).head(15).index
    lines.append("\n" + "=" * 90)
    lines.append("[1] 물량 top15 협력사: 대분류 단일 vs 중분류로 쪼갰을 때 (5일도달률, n>=1000만 표시)")
    lines.append("=" * 90)
    for vendor in top_vendors:
        sub = df[df["협력사명"] == vendor]
        total_n = len(sub)
        overall_5d = (sub["lt"] <= 5).mean() * 100
        lines.append(f"\n{vendor} (전체 n={total_n:,}, 대분류 통합 5일도달률={overall_5d:.1f}%)")
        g = sub.groupby("중분류")["lt"].agg(n="count", rate5=lambda s: (s <= 5).mean() * 100)
        g_usable = g[g["n"] >= 1000].sort_values("n", ascending=False)
        if len(g_usable) <= 1:
            lines.append("  -> n>=1000 중분류 그룹이 1개 이하: 쪼갤 실익 없음")
        else:
            spread = g_usable["rate5"].max() - g_usable["rate5"].min()
            lines.append(f"  -> n>=1000 중분류 {len(g_usable)}개, 5일도달률 범위 {spread:.1f}%p")
            for cat, row in g_usable.head(6).iterrows():
                lines.append(f"     {cat:20s} n={int(row['n']):>7,}  5일도달률={row['rate5']:5.1f}%")

    # 2. 전체 관점: 대분류 단일 vs 중분류 단일 각각의 '설명력' 근사 비교
    # 협력사x대분류 vs 협력사x중분류 세그먼트에서, n>=1000 기준 커버리지와 세그먼트 내 분산 비교
    lines.append("\n" + "=" * 90)
    lines.append("[2] 전체 데이터: 협력사x대분류 vs 협력사x중분류 세그먼트 (n>=1000 기준)")
    lines.append("=" * 90)
    for label, keys in [("협력사x대분류", ["협력사명", "대분류"]), ("협력사x중분류", ["협력사명", "중분류"])]:
        g = df.groupby(keys)["lt"].agg(n="count", rate5=lambda s: (s <= 5).mean())
        g_usable = g[g["n"] >= 1000]
        coverage = g_usable["n"].sum() / len(df) * 100
        lines.append(f"  {label}: 세그먼트 수={len(g_usable):,}  (n>=1000 세그먼트가 커버하는 주문 비중)={coverage:.1f}%")

    report_path = OUT_DIR / "category_granularity_check.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines[:40]))
    print(f"\n... (전체는 {report_path} 참고)")


if __name__ == "__main__":
    main()
