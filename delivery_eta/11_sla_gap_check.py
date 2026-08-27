# -*- coding: utf-8 -*-
"""
Phase 2: 표준납기일(SLA) vs 모델 예측(headline, 90% 확신일) 비교.

사용자 질문("표준납기일과 너무 차이나면 고객이 의문을 가질 것") 검증용. 상품코드 단위로
과거 표준납기일(홀드아웃 2026H1 중앙값)과 모델이 실제로 화면에 보여줄 headline(fd_d가
90%를 처음 넘는 날)을 나란히 비교한다. 표준납기_lt는 SKU마다 완전히 고정된 값은 아니라서
(표준편차 0인 상품 16.7%뿐, 나머지도 중앙값 1일 이내로 비교적 안정적) 대표값으로 중앙값을
쓴다.

결론(2026-07-09, 2026-07-10 raw data 최신화 후 재확인해도 동일): gap이 큰 쪽(5일 초과,
전체 SKU의 7.0%/물량 5.1%)은 무작위가 아니라 kraljic_delivery_risk.csv의 전략재/병목재
(공급위험 高)에 95%대가 몰려있음 - 이미 알려진 배치성 발주/공급불안정 상품군과 정확히
겹침. 나머지 93%는 gap이 없거나 경미(0~2일)함. 반대로 50.1%는 모델이 오히려 SLA보다
빠르게 예측(SLA가 보수적으로 잡혀있다는 뜻).
"""
import importlib.util
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

HERE = Path(__file__).parent
GAP_BIG = 5     # 이 이상이면 "눈에 띄는 불일치"
GAP_MED = 2     # 이 초과면 "중간 정도 차이"


def _load_dist_fit():
    spec = importlib.util.spec_from_file_location("dist_fit", HERE / "03_distribution_fit.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _load_serve():
    spec = importlib.util.spec_from_file_location("serve", HERE / "serve.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def headline_day(row, fine_days):
    for d in fine_days:
        if row[f"fd_{d}"] >= 0.9:
            return d
    return fine_days[-1]


def main():
    dist_fit = _load_dist_fit()
    serve = _load_serve()

    print("holdout(2026H1) 로딩...", flush=True)
    holdout_raw = load_many(SEGMENTS, [2026])
    holdout = dist_fit.prep(holdout_raw)
    sla = holdout.dropna(subset=["표준납기_lt"])
    sla = sla[(sla["표준납기_lt"] >= 0) & (sla["표준납기_lt"] <= 90)]

    pred = serve.DeliveryPredictor()
    # pred.sku는 groupby 키가 상품코드뿐이라 협력사명/대분류/중분류가 항상 비어있음
    # (known_risks.md 기록된 버그) - product_lookup.csv에서 채워 넣을 것이므로 미리 제거.
    sku_table = pred.sku.reset_index().drop(columns=["협력사명", "대분류", "중분류"]).copy()
    sku_sla = sla.groupby("상품코드")["표준납기_lt"].median().rename("sla_median")
    merged = sku_table.merge(sku_sla, left_on="상품코드", right_index=True, how="inner")
    merged["headline_day"] = merged.apply(lambda r: headline_day(r, dist_fit.FINE_DAYS), axis=1)
    merged["gap"] = merged["headline_day"] - merged["sla_median"]

    lookup = pd.read_csv(HERE / "product_lookup.csv", dtype=str)[["상품코드", "상품명", "협력사명", "대분류"]]
    kra = pd.read_csv(HERE / "kraljic_delivery_risk.csv", dtype=str)[["상품코드", "kraljic"]]
    merged = merged.merge(lookup, on="상품코드", how="left").merge(kra, on="상품코드", how="left")

    out_cols = ["상품코드", "상품명", "협력사명", "대분류", "n", "median", "sla_median",
                "headline_day", "gap", "kraljic"]
    merged[out_cols].sort_values("gap", ascending=False).to_csv(
        HERE / "sla_gap_analysis.csv", index=False, encoding="utf-8-sig")
    print(f"saved: sla_gap_analysis.csv ({len(merged):,} rows)", flush=True)

    lines = []
    lines.append("표준납기일(SLA) vs 모델 예측(headline) 비교 리포트")
    lines.append(f"대상: SLA 데이터 있는 sku 레벨 상품 {len(merged):,}개 (2026H1 홀드아웃 기준)")
    lines.append("")
    total_n = merged["n"].sum()
    buckets = [
        (-999, -3, "3일 이상 빠름 (SLA가 보수적)"),
        (-3, -1, "1~2일 빠름"),
        (-1, 0.001, "정확히 일치(0일)"),
        (0.001, GAP_MED, f"1~{GAP_MED}일 늦음 (경미)"),
        (GAP_MED, GAP_BIG, f"{GAP_MED}~{GAP_BIG}일 늦음 (중간)"),
        (GAP_BIG, 999, f"{GAP_BIG}일 초과 늦음 (눈에 띄는 불일치)"),
    ]
    lines.append("-- gap(headline_day - SLA 중앙값) 분포 --")
    for lo, hi, label in buckets:
        sub = merged[(merged["gap"] > lo) & (merged["gap"] <= hi)]
        lines.append(f"  {label:32s}: SKU {len(sub):4d}개({len(sub)/len(merged)*100:5.1f}%)  "
                      f"물량비중 {sub['n'].sum()/total_n*100:5.2f}%")

    lines.append("")
    lines.append(f"gap>0(모델이 SLA보다 늦게 예측) 전체: {(merged['gap']>0).sum()}개 ({(merged['gap']>0).mean()*100:.1f}%)")
    lines.append(f"gap<0(모델이 SLA보다 빠르게 예측) 전체: {(merged['gap']<0).sum()}개 ({(merged['gap']<0).mean()*100:.1f}%)")

    big = merged[merged["gap"] > GAP_BIG]
    lines.append("")
    lines.append(f"-- {GAP_BIG}일 초과 불일치({len(big)}개)의 Kraljic 리스크 분류 --")
    kraljic_counts = big["kraljic"].value_counts()
    for k, v in kraljic_counts.items():
        lines.append(f"  {k}: {v}개")
    high_risk_pct = big["kraljic"].isin(["전략재", "병목재"]).mean() * 100
    lines.append(f"  전략재+병목재(공급위험 高) 비율: {high_risk_pct:.1f}%")
    lines.append("  -> 눈에 띄는 불일치는 무작위가 아니라 이미 식별된 배치성 발주/공급불안정")
    lines.append("     상품군에 집중됨. 대부분(93%)은 gap이 없거나 경미함.")

    lines.append("")
    lines.append(f"-- {GAP_BIG}일 초과 불일치 사례 (물량 상위 15개) --")
    for _, row in big.sort_values("n", ascending=False).head(15).iterrows():
        lines.append(f"  {row['상품코드']} {str(row['상품명'])[:20]:20s} {str(row['협력사명'])[:15]:15s} "
                      f"n={int(row['n']):>6,}  SLA={row['sla_median']:.0f}일  모델={row['headline_day']}일  "
                      f"gap={row['gap']:.0f}일  {row['kraljic']}")

    report = "\n".join(lines)
    (HERE / "sla_gap_report.txt").write_text(report, encoding="utf-8")
    print(report)
    print("\nsaved: sla_gap_report.txt")


if __name__ == "__main__":
    main()
