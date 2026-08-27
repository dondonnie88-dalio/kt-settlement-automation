# -*- coding: utf-8 -*-
"""
Phase 3 재고연동 후보 협력사 탐지: '배치성 발주' 패턴(발주->출하 대기가 길고,
특정 월에만 주문이 몰리는) 협력사를 데이터로 찾는다. 향수(10534198, 영풍기획)
사례에서 나온 패턴을 전체 협력사로 일반화해서 확인.

주의(2026-07-08 정정): 원래 이 패턴을 "배치성 수입"이라 불렀으나, 원본 데이터에
원산지/수입여부 필드가 없어 실제로 해외 수입인지 확인할 방법이 없다(제조사 컬럼만
있고 이것만으로는 판별 불가). "배치성 발주"(발주가 몰아서 들어오는 패턴)까지는
데이터로 검증됐지만, 그 원인이 수입 통관 때문인지 국내 정기 대량발주 때문인지는
추정일 뿐 확정된 사실이 아니다 - phase3_backlog.md 참고.
"""
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent
MIN_N = 500  # 이 정도는 있어야 패턴 판단이 의미있음


def main():
    df = load_many(SEGMENTS, [2023, 2024, 2025, 2026])
    df["발주출하"] = (df["출하일자"] - df["발주일자"]).dt.days
    df = df[df["발주출하"].notna() & (df["발주출하"] >= 0) & (df["발주출하"] <= 90)]
    df["주문월"] = df["주문일자"].dt.to_period("M")

    rows = []
    for vendor, g in df.groupby("협력사명"):
        if len(g) < MIN_N:
            continue
        p90 = g["발주출하"].quantile(0.9)
        median = g["발주출하"].median()
        monthly = g.groupby("주문월").size()
        active_months = (monthly > 0).sum()
        total_months = 42  # 2023-01 ~ 2026-06
        month_coverage = active_months / total_months
        rows.append({
            "협력사명": vendor, "n": len(g), "발주출하_median": median, "발주출하_p90": p90,
            "활성월수": active_months, "월커버리지": month_coverage,
        })
    result = pd.DataFrame(rows)
    # 배치성 신호: 발주->출하 대기가 길다(p90>=14) + 특정 달에만 몰린다(월커버리지<=0.5)
    candidates = result[(result["발주출하_p90"] >= 14) & (result["월커버리지"] <= 0.5)].sort_values(
        "발주출하_p90", ascending=False)

    lines = ["Phase 3 재고연동 후보 협력사 탐지 (배치성 발주 패턴 - 원인은 추정, 수입 확정 아님)"]
    lines.append(f"기준: 발주->출하 p90 >= 14일 AND 월커버리지(42개월 중 주문 있었던 달 비율) <= 50%")
    lines.append(f"분석 대상 협력사 수(n>={MIN_N}): {len(result)}, 후보 발견: {len(candidates)}\n")
    lines.append(f"{'협력사명':22s}{'n':>8s}{'발주출하 median':>14s}{'발주출하 p90':>12s}{'활성월수/42':>12s}{'월커버리지':>10s}")
    for _, row in candidates.head(30).iterrows():
        lines.append(f"{row['협력사명']:22s}{int(row['n']):>8,}{row['발주출하_median']:>14.1f}"
                     f"{row['발주출하_p90']:>12.1f}{int(row['활성월수']):>9d}/42{row['월커버리지']*100:>9.0f}%")

    report_path = OUT_DIR / "phase3_batch_vendor_candidates.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved: {report_path}")


if __name__ == "__main__":
    main()
