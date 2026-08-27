# -*- coding: utf-8 -*-
"""
Phase 1 착수 전 리스크 3종 점검
리스크1: 주문시배송리드타임 = 사전 약속값인가, 사후 실측값인가 (데이터 리키지 여부)
리스크2: 외부사 median 21일이 어느 구간에서 발생하는가
리스크3: 직접배송 vs 택배 리드타임/예측가능성 비교
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent

pd.set_option("display.width", 200)


def add_leadtimes(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["lt_order_to_order2"] = (d["발주일자"] - d["주문일자"]).dt.days
    d["lt_order2_to_ship"] = (d["출하일자"] - d["발주일자"]).dt.days
    d["lt_ship_to_deliver"] = (d["배송완료일자"] - d["출하일자"]).dt.days
    d["lt_deliver_to_recv"] = (d["입고일자"] - d["배송완료일자"]).dt.days
    d["lt_order_to_deliver"] = (d["배송완료일자"] - d["주문일자"]).dt.days
    d["lt_order_to_recv"] = (d["입고일자"] - d["주문일자"]).dt.days
    d["대분류"] = d["서비스카테고리"].fillna("").str.split(">").str[0].str.strip()
    d["is_direct"] = d["배송업체"] == "직접배송"
    return d


def risk1_leadtime_identity(d: pd.DataFrame, lines: list):
    lines.append("\n" + "=" * 70)
    lines.append("[리스크1] 주문시배송리드타임 정체 규명")
    lines.append("=" * 70)

    lt = pd.to_numeric(d["주문시배송리드타임"], errors="coerce")
    lines.append(f"n={len(d):,}, 결측={lt.isna().sum():,} ({lt.isna().mean()*100:.1f}%)")
    lines.append(f"분포: min={lt.min()}, p25={lt.quantile(.25)}, median={lt.median()}, "
                 f"p75={lt.quantile(.75)}, p90={lt.quantile(.9)}, max={lt.max()}")
    lines.append(f"정수 비율: {(lt.dropna() == lt.dropna().round()).mean()*100:.1f}%")
    lines.append(f"최빈값 top5:\n{lt.value_counts().head(5).to_string()}")

    for label, calc_col in [
        ("A. (배송완료일자-주문일자)", "lt_order_to_deliver"),
        ("B. (입고일자-주문일자)", "lt_order_to_recv"),
        ("C. (표준납기일-주문일자)", None),
    ]:
        if calc_col is None:
            calc = (d["표준납기일"] - d["주문일자"]).dt.days
        else:
            calc = d[calc_col]
        sub = pd.DataFrame({"lt": lt, "calc": calc}).dropna()
        exact = (sub["lt"] == sub["calc"]).mean() * 100 if len(sub) else float("nan")
        within1 = (sub["lt"] - sub["calc"]).abs().le(1).mean() * 100 if len(sub) else float("nan")
        lines.append(f"  {label}: n비교={len(sub):,}  정확히 일치={exact:.1f}%  |오차|<=1일={within1:.1f}%")

    lines.append("\n-- 핵심 테스트: 아직 배송완료 안 된(배송완료일자 NaT) 주문도 주문시배송리드타임 값이 채워져 있는가? --")
    max_order_date = d["주문일자"].max()
    recent = d[d["주문일자"] >= max_order_date - pd.Timedelta(days=14)]
    not_delivered = recent[recent["배송완료일자"].isna()]
    delivered = recent[recent["배송완료일자"].notna()]
    lines.append(f"최근 14일 주문 n={len(recent):,} (기준일 {max_order_date.date()})")
    lines.append(f"  배송완료 안됨(NaT) n={len(not_delivered):,}  중 주문시배송리드타임 값 있음 비율="
                 f"{pd.to_numeric(not_delivered['주문시배송리드타임'], errors='coerce').notna().mean()*100:.1f}%")
    lines.append(f"  배송완료 됨       n={len(delivered):,}  중 주문시배송리드타임 값 있음 비율="
                 f"{pd.to_numeric(delivered['주문시배송리드타임'], errors='coerce').notna().mean()*100:.1f}%")
    lines.append("(해석: 배송완료 안 된 주문도 값이 채워져 있으면 -> 주문시점 '약속/예정' 값(피처로 안전),")
    lines.append(" 배송완료 안 된 주문은 값이 비어있으면 -> 배송완료와 함께 사후 기록되는 값(리키지 위험, 라벨 후보에 가까움)")


def risk2_external_breakdown(d: pd.DataFrame, lines: list):
    lines.append("\n" + "=" * 70)
    lines.append("[리스크2] 외부사 median 21일 구간 분해")
    lines.append("=" * 70)

    for seg in SEGMENTS:
        sub = d[d["_segment"] == seg]
        lines.append(f"\n-- {seg} (n={len(sub):,}) --")
        for col, name in [
            ("lt_order_to_order2", "주문->발주"),
            ("lt_order2_to_ship", "발주->출하"),
            ("lt_ship_to_deliver", "출하->배송완료"),
            ("lt_deliver_to_recv", "배송완료->입고"),
            ("lt_order_to_recv", "주문->입고(합계)"),
        ]:
            s = sub[col].dropna()
            s = s[(s >= -5) & (s <= 150)]
            lines.append(f"    {name:16s} median={s.median():6.1f}  p90={s.quantile(.9):6.1f}  n={len(s):,}")
        direct_pct = sub["is_direct"].mean() * 100
        lines.append(f"    직접배송 비중: {direct_pct:.1f}%")

    lines.append("\n-- 외부사 주문->입고 상위 지연 상품 대분류 top 10 (median 기준, n>=30) --")
    ext = d[d["_segment"] == "외부사"]
    g = ext.dropna(subset=["lt_order_to_recv"])
    g = g[(g["lt_order_to_recv"] >= -5) & (g["lt_order_to_recv"] <= 150)]
    cat_stats = g.groupby("대분류")["lt_order_to_recv"].agg(["median", "count"])
    cat_stats = cat_stats[cat_stats["count"] >= 30].sort_values("median", ascending=False).head(10)
    for cat, row in cat_stats.iterrows():
        lines.append(f"    {cat:24s} median={row['median']:6.1f}  n={int(row['count']):,}")

    lines.append("\n-- 외부사: 입고일자 없이 배송완료일자만 있는 비중 (창고 미경유=협력사 직배송 추정) --")
    has_deliver = ext["배송완료일자"].notna()
    no_recv_but_delivered = ext[has_deliver & ext["입고일자"].isna()]
    lines.append(f"    배송완료 있음 n={has_deliver.sum():,} 중 입고일자 없음 n={len(no_recv_but_delivered):,} "
                 f"({len(no_recv_but_delivered)/max(has_deliver.sum(),1)*100:.1f}%)")


def risk3_direct_vs_courier(d: pd.DataFrame, lines: list):
    lines.append("\n" + "=" * 70)
    lines.append("[리스크3] 직접배송 vs 택배 비교 및 사전 예측 가능성")
    lines.append("=" * 70)

    for seg in SEGMENTS:
        sub = d[d["_segment"] == seg]
        lines.append(f"\n-- {seg} --")
        for is_direct, name in [(True, "직접배송"), (False, "택배(그외)")]:
            s = sub[sub["is_direct"] == is_direct]["lt_ship_to_deliver"].dropna()
            s = s[(s >= -5) & (s <= 60)]
            lines.append(f"    {name:10s} n={len(s):>8,}  median={s.median() if len(s) else float('nan'):5.1f}  "
                         f"p90={s.quantile(.9) if len(s) else float('nan'):5.1f}")

    lines.append("\n-- 상품 대분류별 직접배송 비중 top 15 (n>=100) --")
    g = d.groupby("대분류")["is_direct"].agg(["mean", "count"])
    g = g[g["count"] >= 100].sort_values("mean", ascending=False).head(15)
    for cat, row in g.iterrows():
        lines.append(f"    {cat:24s} 직접배송비중={row['mean']*100:5.1f}%  n={int(row['count']):,}")

    lines.append("\n-- 협력사별 배송방식 일관성 (주문시점 예측 가능성 테스트, n>=50) --")
    g = d.groupby("협력사명")["is_direct"].agg(["mean", "count"])
    g = g[g["count"] >= 50]
    consistent = ((g["mean"] >= 0.95) | (g["mean"] <= 0.05))
    lines.append(f"    협력사 수(n>=50건)={len(g):,}")
    lines.append(f"    배송방식 95%+ 일관된 협력사 비율={consistent.mean()*100:.1f}% "
                 f"(즉 {consistent.mean()*100:.0f}%는 협력사명만 알아도 배송방식 예측 가능)")
    lines.append(f"    혼합(5~95%) 협력사 비율={(~consistent).mean()*100:.1f}% -> 이 경우는 협력사만으론 예측 불가, 별도 피처 필요")


def main():
    print("데이터 로딩 (캐시 있으면 재사용, 없으면 raw 파싱+캐싱)...")
    df = load_many(SEGMENTS, [2026])
    d = add_leadtimes(df)

    lines = ["배송 예측 정보 과제 - Phase 1 착수 전 리스크 점검 (2026년 1~6월 기준)"]
    risk1_leadtime_identity(d, lines)
    risk2_external_breakdown(d, lines)
    risk3_direct_vs_courier(d, lines)

    report_path = OUT_DIR / "risk_check_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nDone. Report saved to {report_path}")


if __name__ == "__main__":
    main()
