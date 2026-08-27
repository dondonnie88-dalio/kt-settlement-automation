# -*- coding: utf-8 -*-
"""
세그먼트 안정성 재발 방지: 최근 물량이 과거 대비 급증한 협력사 자동 탐지.
(원창유통/알파/제이와이코퍼레이션 사례 - 03_distribution_fit.py MIN_N=1000 근거)
정기적으로(예: 매 분기) 재실행해서, 새로 폭증한 협력사가 있으면
세그먼트 재검증 없이 그대로 학습에 포함되지 않도록 확인하는 용도.
"""
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent
GROWTH_RATIO_THRESHOLD = 5.0
MIN_RECENT_VOLUME = 500  # 이 정도는 돼야 "무시 못 할 신규 대량 세그먼트"로 판단


def main():
    hist = load_many(SEGMENTS, [2023, 2024, 2025])
    recent = load_many(SEGMENTS, [2026])
    for df in (hist, recent):
        df["대분류"] = df["서비스카테고리"].fillna("").str.split(">").str[0].str.strip()

    # 원창유통/알파 사례는 '협력사' 단위가 아니라 '협력사x대분류' 조합에서만 신규였음
    # (해당 협력사 자체는 기존에도 큰 거래처) -> 반드시 이 레벨에서 체크
    hist_vol = hist.groupby(["협력사명", "대분류"]).size().rename("hist_n")
    recent_vol = recent.groupby(["협력사명", "대분류"]).size().rename("recent_n")

    joined = pd.concat([hist_vol, recent_vol], axis=1).fillna(0)
    joined["recent_annualized"] = joined["recent_n"] * 2
    joined["hist_annual_avg"] = joined["hist_n"] / 3
    joined["growth_ratio"] = joined["recent_annualized"] / joined["hist_annual_avg"].replace(0, 0.5)

    flagged = joined[
        (joined["recent_n"] >= MIN_RECENT_VOLUME)
        & (joined["growth_ratio"] >= GROWTH_RATIO_THRESHOLD)
    ].sort_values("recent_n", ascending=False)

    lines = ["급성장 협력사x상품군 세그먼트 탐지 리포트 (재발 방지용)"]
    lines.append(f"기준: 2026 상반기 물량 >= {MIN_RECENT_VOLUME}건 AND 연환산 성장률 >= {GROWTH_RATIO_THRESHOLD}배")
    lines.append(f"\n탐지된 세그먼트 수: {len(flagged)}")
    lines.append(f"\n{'협력사명':20s}{'대분류':16s}{'2023-25 합계':>12s}{'2026H1':>10s}{'연환산 성장률':>14s}")
    for (vendor, cat), row in flagged.iterrows():
        lines.append(f"{str(vendor):20s}{str(cat):16s}{int(row['hist_n']):>12,}{int(row['recent_n']):>10,}{row['growth_ratio']:>13.1f}x")

    lines.append("\n권고: 이 협력사x상품군 세그먼트는 vendor_mid 레벨로 세분화하기 전에 "
                 "'대량 운영 전환 후 최소 1000건' 누적 여부를 확인할 것. "
                 "누적 전까지는 vendor 레벨(또는 그보다 상위)로 강제 폴백 권장 "
                 "(03_distribution_fit.py의 MIN_N=1000이 이미 이 케이스를 자동 처리함).")

    report_path = OUT_DIR / "growth_vendor_flags.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved: {report_path}")


if __name__ == "__main__":
    main()
