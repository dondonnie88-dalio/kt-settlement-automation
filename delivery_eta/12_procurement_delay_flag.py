# -*- coding: utf-8 -*-
"""
발주->출하 지연 협력사x상품군x월 자동 탐지 (조기경보용).

배경 (2026-07-15, 협력사 배송리드타임 관리 목적으로 신설): vendor_mid 레벨의 21~30일
잔존 캘리브레이션 오차를 끝까지 추적한 결과, 오차의 대부분(약 93%)이 소수 협력사x상품군
조합에서 "특정 한 달만 극단적으로 느려지고 나머지 달은 정상"인 패턴으로 확인됨(케이눅/
생활용품·기타 4월, 현웅디자인/KT그룹 일체복 4월, 예일토탈싸인/출판·인쇄 5월 등).
배송완료일자/입고일자/발주일자/출하일자를 전부 대조한 결과 "내부 기록 지연 착시"가
아니라 실제 지연이었고, 위치도 정확히 특정됨: 발주는 거의 즉시 처리되고(0~1.4일),
출하 이후 배송완료까지도 정상(~1일)인데, "발주 이후 실제 출고되기까지"(출하 준비 단계)
에서만 지연이 집중됨 - 라스트마일 배송이 아니라 협력사의 조달/재고/포장 단계 이슈.
known_risks.md의 "vendor_mid day21/30 오차의 정체" 항목 참고.

05_growth_vendor_flag.py(물량 급증 감시)와 같은 철학 - 다만 감시 대상은 "물량"이 아니라
"발주->출하 소요일". 정기적으로(예: 매월) 재실행해서 협력사 배송리드타임 관리(구매/운영)
실무에 조기 알림으로 활용하는 용도.
"""
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent
# 기준선(BASELINE_YEARS)과 감시대상(MONITOR_YEARS)을 반드시 분리한다 - 자기 자신의 최근
# 데이터만으로 기준선을 잡으면, 그 협력사의 감시 기간 전체가 이미 이상 상태일 경우(예:
# 케이눅/생활용품·기타는 2026년 물량 전체가 4~5월 이상 구간에만 존재) 이상을 자기 자신과
# 비교하는 꼴이 되어 못 잡아낸다(2026-07-15, 초기 설계 시행착오로 발견). 05_growth_vendor_
# flag.py와 동일하게 과거 3개년을 "평시" 기준으로, 최근 데이터를 "감시 대상"으로 분리한다.
BASELINE_YEARS = [2023, 2024, 2025]
MONITOR_YEARS = [2026]  # 정기 재실행 시 그 시점의 "최근 데이터" 연도로 갱신할 것
RATIO_THRESHOLD = 3.0  # 그 협력사x상품군의 과거(BASELINE_YEARS) 중앙값 대비 몇 배 이상이면 이상으로 볼지
MIN_ABS_DAYS = 14.0  # 배 기준을 넘어도 절대 지연일수가 이 미만이면 무시(사소한 변동 배제)
MIN_MONTHLY_VOLUME = 30  # 이 정도는 돼야 그 달의 평균이 우연 몇 건에 흔들리지 않음
MIN_BASELINE_VOLUME = 30  # 과거 기준선 자체도 이 정도는 있어야 신뢰(너무 적으면 기준선 제외)


def prep_lt(df):
    df = df.copy()
    df["중분류"] = df["서비스카테고리"].fillna("").str.split(">").str[1].str.strip()
    df["발주출하_lt"] = (df["출하일자"] - df["발주일자"]).dt.days
    d = df.dropna(subset=["발주출하_lt"])
    return d[(d["발주출하_lt"] >= 0) & (d["발주출하_lt"] <= 180)]


def main():
    hist = prep_lt(load_many(SEGMENTS, BASELINE_YEARS))
    recent = prep_lt(load_many(SEGMENTS, MONITOR_YEARS))
    recent["월"] = recent["주문일자"].dt.month

    monthly = recent.groupby(["협력사명", "중분류", "월"]).agg(
        n=("발주출하_lt", "size"),
        발주출하_평균=("발주출하_lt", "mean"),
    ).reset_index()
    monthly = monthly[monthly["n"] >= MIN_MONTHLY_VOLUME]

    baseline_grp = hist.groupby(["협력사명", "중분류"])["발주출하_lt"]
    baseline = baseline_grp.median().rename("기준_중앙값")
    baseline_n = baseline_grp.size().rename("기준_n")
    monthly = monthly.merge(baseline, on=["협력사명", "중분류"], how="left")
    monthly = monthly.merge(baseline_n, on=["협력사명", "중분류"], how="left")
    # 과거 기준선이 아예 없거나(신규 협력사x상품군) 너무 적으면 비교 불가 - 제외
    # (신규 조합의 초기 불안정은 05_growth_vendor_flag.py/MIN_N=1000이 이미 별도로 처리함)
    monthly = monthly[monthly["기준_n"] >= MIN_BASELINE_VOLUME]
    monthly["기준_중앙값"] = monthly["기준_중앙값"].replace(0, 0.5)
    monthly["배율"] = monthly["발주출하_평균"] / monthly["기준_중앙값"]

    flagged = monthly[
        (monthly["배율"] >= RATIO_THRESHOLD)
        & (monthly["발주출하_평균"] >= MIN_ABS_DAYS)
    ].sort_values("n", ascending=False)

    lines = ["발주->출하 지연 협력사x상품군x월 탐지 리포트 (조기경보용)"]
    lines.append(f"기준선 기간: {BASELINE_YEARS} (평시)  /  감시 대상 기간: {MONITOR_YEARS} (최근)")
    lines.append(f"기준: 그 달 물량 >= {MIN_MONTHLY_VOLUME}건 AND 발주->출하 평균이 "
                 f"과거({BASELINE_YEARS[0]}~{BASELINE_YEARS[-1]}) 중앙값 대비 "
                 f"{RATIO_THRESHOLD}배 이상 AND 절대 소요일 {MIN_ABS_DAYS:.0f}일 이상 "
                 f"(과거 데이터 {MIN_BASELINE_VOLUME}건 미만인 신규 조합은 비교 불가로 제외)")
    lines.append(f"\n탐지된 협력사x상품군x월: {len(flagged)}건")
    lines.append(f"\n{'협력사명':22s}{'상품군(중분류)':18s}{'월':>4s}{'물량':>8s}"
                 f"{'평시(과거중앙값)':>14s}{'해당월 평균':>12s}{'배율':>8s}")
    for _, row in flagged.iterrows():
        lines.append(f"{str(row['협력사명']):22s}{str(row['중분류']):18s}{int(row['월']):>4d}"
                     f"{int(row['n']):>8,}{row['기준_중앙값']:>11.1f}일{row['발주출하_평균']:>11.1f}일"
                     f"{row['배율']:>7.1f}x")

    lines.append("\n해석: \"발주는 정상 처리됐는데 그 협력사에서 실제 출고되기까지가 평소보다 "
                 "훨씬 오래 걸린\" 달입니다. 배송(라스트마일) 문제가 아니라 그 협력사의 조달/"
                 "재고/포장 등 출하 준비 단계 이슈일 가능성이 높습니다 - 협력사 배송리드타임 "
                 "관리(구매/운영 실무) 관점에서 그 달에 해당 협력사에 무슨 일이 있었는지 확인할 "
                 "가치가 있습니다.")
    lines.append("\n권고: 정기적으로(예: 매월) 재실행해서 새로 발생한 이상을 조기에 파악할 것. "
                 "05_growth_vendor_flag.py(물량 급증 감시)와 함께 실행하는 걸 권장.")

    report_path = OUT_DIR / "procurement_delay_flags.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved: {report_path}")


if __name__ == "__main__":
    main()
