# -*- coding: utf-8 -*-
"""
Phase 3: 배송 리스크 Kraljic 매트릭스 분류 - demand_forecast.py의 classify_kraljic() 로직을
배송예측 데이터(주문 단위, 훨씬 세밀함)에 맞게 적용한다.

09_batch_import_vendor_check.py(협력사 단위, 발주->출하 p90 + 월커버리지라는 임시 기준)를
정식 리스크 점수 체계로 대체 - 그 스크립트가 놓쳤던 "혼합 포트폴리오 협력사의 특정 SKU"
문제를 상품코드 단위로 분석해서 해결한다.

공급 리스크(高) 판정 - 아래 중 2개 이상 충족 (demand_forecast.py의 3개 기준과 동일 철학,
다만 "벤더의존도"는 이 도메인에 맞게 재정의함):
  - 발주->출하 변동성(CV = std/mean) > 0.3  (배치성 발주 신호 - 정기 대량발주 등으로 추정되나
    원인 미확정, 향수 사례에서 확인된 패턴. "수입"이라 부르지 않음 - 원산지 데이터 없어 미확인)
  - 카테고리(중분류) 내 대체 공급사 부족 (해당 중분류를 취급하는 협력사가 2곳 이하)
    ** 주의: 처음엔 "상품코드 단일소싱 여부"(product_lookup.csv의 n_협력사_고유값==1)로
    시도했으나 전체 상품의 97%가 해당되어 구분력이 없었음(대부분 상품이 애초에 SKU 단위로는
    한 협력사만 취급 - 이 도메인의 정상 패턴). 실제 "대체 불가" 리스크는 상품코드가 아니라
    "같은 종류 상품을 파는 협력사가 시장에 몇 곳이나 있는지"(중분류 단위)가 맞는 지표라
    2026-07-06 교체함 - 중분류당 협력사 수는 1~576개로 넓게 분포(median 6, 40개 중분류는
    1곳뿐)해서 훨씬 잘 구분됨.
  - OFR(7일 이내 도착률) < 85%  (segment_distributions.csv의 sku 레벨 p_7d 재사용)

구매 영향력(高): 주문량(n) 상위 80% 누적 기여 구간 (ABC 파레토와 동일 철학, 매출액 대신
주문건수를 프록시로 사용 - 배송예측 맥락에서는 "몇 명의 고객 경험에 영향을 주는지"가
매입금액보다 더 직접적인 지표).
"""
import numpy as np
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

HERE = Path(__file__).parent
LT_CV_THRESHOLD = 0.3
OFR_THRESHOLD = 0.85
MIN_N_FOR_CV = 30  # 발주->출하 변동성 계산에 필요한 최소 표본
MAX_CATEGORY_VENDORS = 2  # 이 이하면 "대체 공급사 부족(시장 병목)"으로 판정


def main():
    print("데이터 로딩 중...", flush=True)
    df = load_many(SEGMENTS, [2023, 2024, 2025, 2026])
    df["발주출하"] = (df["출하일자"] - df["발주일자"]).dt.days
    valid = df[df["발주출하"].notna() & (df["발주출하"] >= 0) & (df["발주출하"] <= 90)]

    print("상품코드별 발주->출하 변동성 계산 중...", flush=True)
    lt_cv_stats = valid.groupby("상품코드")["발주출하"].agg(n_lt="count", lt_mean="mean", lt_std="std")
    lt_cv_stats["lt_std"] = lt_cv_stats["lt_std"].fillna(0)
    lt_cv_stats["lt_cv"] = np.where(
        (lt_cv_stats["n_lt"] >= MIN_N_FOR_CV) & (lt_cv_stats["lt_mean"] > 0),
        lt_cv_stats["lt_std"] / lt_cv_stats["lt_mean"], 0.0
    )

    product_lookup = pd.read_csv(HERE / "product_lookup.csv", dtype={"상품코드": "string"})
    seg = pd.read_csv(HERE / "segment_distributions.csv", dtype={"상품코드": "string"})
    sku_seg = seg[seg["level"] == "sku"][["상품코드", "n", "median", "p_7d"]]

    lt_cv_stats = lt_cv_stats.reset_index()
    lt_cv_stats["상품코드"] = lt_cv_stats["상품코드"].astype("string")

    # 중분류당 대체 공급사 수 (product_lookup 전체 122,061개 SKU 기준 - sku_seg보다 훨씬
    # 큰 모집단으로 계산해야 "시장에 실제 몇 개 협력사가 있는지"가 정확함)
    category_vendor_count = product_lookup.groupby("중분류")["협력사명"].nunique()

    result = sku_seg.merge(product_lookup[["상품코드", "상품명", "협력사명", "대분류", "중분류"]],
                            on="상품코드", how="left")
    result = result.merge(lt_cv_stats[["상품코드", "lt_cv", "n_lt"]], on="상품코드", how="left")
    result["lt_cv"] = result["lt_cv"].fillna(0.0)
    result["n_lt"] = result["n_lt"].fillna(0)
    result["category_vendor_count"] = result["중분류"].map(category_vendor_count)

    # ── 공급 리스크 ──
    result["category_bottleneck"] = result["category_vendor_count"] <= MAX_CATEGORY_VENDORS
    risk_score = (
        (result["lt_cv"] > LT_CV_THRESHOLD).astype(int)
        + result["category_bottleneck"].astype(int)
        + (result["p_7d"] < OFR_THRESHOLD).astype(int)
    )
    result["supply_risk"] = np.where(risk_score >= 2, "高", "低")

    # ── 구매 영향력 (주문량 파레토, ABC와 동일 철학) ──
    result = result.sort_values("n", ascending=False).reset_index(drop=True)
    result["cum_pct"] = result["n"].cumsum() / result["n"].sum() * 100
    result["profit_impact"] = np.where(result["cum_pct"] <= 80, "高", "低")

    def _quadrant(row):
        if row["supply_risk"] == "高":
            return "전략재" if row["profit_impact"] == "高" else "병목재"
        return "레버리지재" if row["profit_impact"] == "高" else "일반재"
    result["kraljic"] = result.apply(_quadrant, axis=1)

    lines = ["배송 리스크 Kraljic 매트릭스 분류 (상품코드 단위, sku 레벨 세그먼트만 대상)"]
    lines.append(f"대상 상품코드 수: {len(result):,}개 (segment_distributions.csv의 level=sku 전체)")
    lines.append(f"기준: 공급리스크 = [발주출하 CV>{LT_CV_THRESHOLD}] + [중분류 내 협력사<={MAX_CATEGORY_VENDORS}곳] + [7일도달률<{OFR_THRESHOLD:.0%}] 중 2개+")
    lines.append(f"      구매영향력 = 주문량 누적기여 80% 이내\n")

    counts = result["kraljic"].value_counts()
    for q in ["전략재", "병목재", "레버리지재", "일반재"]:
        lines.append(f"  {q}: {counts.get(q, 0):,}개")

    lines.append("\n[전략재] 리스크 높음 + 물량 많음 - 최우선 관리 대상 (상위 20개, 물량순)")
    strategic = result[result["kraljic"] == "전략재"].sort_values("n", ascending=False).head(20)
    lines.append(f"{'상품코드':>10s} {'상품명':20s} {'협력사명':18s} {'n':>8s} {'중앙값':>6s} {'7일도달률':>8s} {'발주출하CV':>10s} {'중분류내협력사수':>12s}")
    for _, row in strategic.iterrows():
        lines.append(f"{row['상품코드']:>10s} {str(row['상품명'])[:20]:20s} {str(row['협력사명'])[:18]:18s} "
                     f"{int(row['n']):>8,} {row['median']:>6.1f} {row['p_7d']*100:>7.1f}% {row['lt_cv']:>10.2f} "
                     f"{int(row['category_vendor_count']):>12,d}")

    lines.append("\n[병목재] 리스크 높음 + 물량 적음 - 대체 소싱 검토 대상 (상위 20개, CV순)")
    bottleneck = result[result["kraljic"] == "병목재"].sort_values("lt_cv", ascending=False).head(20)
    lines.append(f"{'상품코드':>10s} {'상품명':20s} {'협력사명':18s} {'n':>8s} {'중앙값':>6s} {'7일도달률':>8s} {'발주출하CV':>10s} {'중분류내협력사수':>12s}")
    for _, row in bottleneck.iterrows():
        lines.append(f"{row['상품코드']:>10s} {str(row['상품명'])[:20]:20s} {str(row['협력사명'])[:18]:18s} "
                     f"{int(row['n']):>8,} {row['median']:>6.1f} {row['p_7d']*100:>7.1f}% {row['lt_cv']:>10.2f} "
                     f"{int(row['category_vendor_count']):>12,d}")

    report_path = HERE / "kraljic_delivery_risk_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines[:40]))
    print(f"\n... (전체는 {report_path} 참고)")

    out_cols = ["상품코드", "상품명", "협력사명", "대분류", "중분류", "n", "median", "p_7d",
                "lt_cv", "category_vendor_count", "category_bottleneck", "supply_risk",
                "profit_impact", "kraljic"]
    result[out_cols].to_csv(HERE / "kraljic_delivery_risk.csv", index=False, encoding="utf-8-sig")
    print(f"Saved: kraljic_delivery_risk_report.txt, kraljic_delivery_risk.csv ({len(result):,} rows)")


if __name__ == "__main__":
    main()
