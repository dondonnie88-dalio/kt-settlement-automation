# -*- coding: utf-8 -*-
"""
배송 예측에 영향을 줄 수 있는 '특정 날짜 전사적 이상 이벤트'(명절/기상/택배사 파업 등) 탐지.
정확한 공휴일/파업 날짜를 하드코딩해서 끼워맞추는 대신, 데이터에서 "그 날 주문한 건들이
평소보다 훨씬 느렸다"는 신호를 직접 찾아낸다(협력사 전반에 걸친 동시다발적 지연만 잡음 -
특정 협력사 하나만 느린 건 이 분석의 대상이 아님).
"""
import pandas as pd
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent
MIN_DAILY_N = 500       # 이 정도는 돼야 '그날 전체적으로' 느렸다고 볼 수 있음
ROLL_WINDOW = 21        # 앞뒤 평상시 대비 기준선(일)
DROP_THRESHOLD = 15.0   # 기준선 대비 %p 이상 하락하면 이상일로 플래그


def main():
    df = load_many(SEGMENTS, [2023, 2024, 2025, 2026])
    df["lt"] = (df["배송완료일자"] - df["주문일자"]).dt.days
    df = df[
        df["lt"].notna() & (df["lt"] >= 0) & (df["lt"] <= 90)
        & (df["취소수량"].fillna(0) == 0) & (df["반품수량"].fillna(0) == 0)
    ]

    daily = df.groupby(df["주문일자"].dt.date).agg(
        n=("lt", "size"), rate3=("lt", lambda s: (s <= 3).mean() * 100)
    )
    daily.index = pd.to_datetime(daily.index)
    daily = daily.sort_index()
    full_range = pd.date_range(daily.index.min(), daily.index.max(), freq="D")
    daily = daily.reindex(full_range)

    # 기준선: 같은 요일끼리만 비교 (요일 자체 효과 - 금/목요일은 원래 3일도달률이 낮음 -
    # 를 '이상치'로 오탐하지 않도록, 같은 요일의 앞뒤 값들로만 롤링 기준선을 만든다)
    valid_mask = daily["n"] >= MIN_DAILY_N
    daily["baseline"] = pd.NA
    for wd in range(7):
        wd_mask = daily.index.dayofweek == wd
        series = daily.loc[wd_mask & valid_mask, "rate3"]
        # 요일별 시계열은 7일 간격이므로 window 개수는 그대로 두되 실제로는 몇 주치 창
        base = series.rolling(window=ROLL_WINDOW, center=True, min_periods=5).median()
        daily.loc[base.index, "baseline"] = base
    daily["baseline"] = pd.to_numeric(daily["baseline"])
    daily["deviation"] = daily["rate3"] - daily["baseline"]

    anomalies = daily[valid_mask & (daily["deviation"] <= -DROP_THRESHOLD)].copy()
    anomalies["weekday"] = anomalies.index.day_name()
    anomalies = anomalies.sort_values("deviation")

    lines = ["배송 지연 이상일(Anomaly Day) 탐지 - 명절/기상/파업 등 전사적 이벤트 후보"]
    lines.append(f"기준: 하루 주문 {MIN_DAILY_N}건 이상 & 앞뒤 {ROLL_WINDOW}일 평상시 대비 3일도달률 {DROP_THRESHOLD}%p 이상 하락")
    lines.append(f"\n전체 분석 대상 일수: {valid_mask.sum():,}일 (2023-01 ~ 2026-06)")
    lines.append(f"이상일로 탐지된 날: {len(anomalies)}건 (상위 30개 표시)\n")
    lines.append(f"{'날짜':>12s} {'요일':>6s} {'주문건수':>10s} {'3일도달률':>10s} {'평상시기준':>10s} {'하락폭':>8s}")
    for date, row in anomalies.head(30).iterrows():
        lines.append(f"{date.strftime('%Y-%m-%d'):>12s} {row['weekday']:>6s} {int(row['n']):>10,} "
                     f"{row['rate3']:>9.1f}% {row['baseline']:>9.1f}% {row['deviation']:>7.1f}%p")

    # 참고용 - 잘 알려진 한국 명절 대략 시기(정확한 날짜는 아니며 근사치, 실제 확정 공휴일표로 재확인 필요)
    lines.append("\n\n[참고, 근사치] 알려진 한국 명절 대략 시기 (정확한 날짜는 별도 확인 필요):")
    approx_holidays = [
        ("2023-01-21", "2023-01-24", "설날 연휴(대략)"),
        ("2023-09-28", "2023-09-30", "추석 연휴(대략)"),
        ("2024-02-09", "2024-02-12", "설날 연휴(대략)"),
        ("2024-09-16", "2024-09-18", "추석 연휴(대략)"),
        ("2025-01-28", "2025-01-30", "설날 연휴(대략)"),
        ("2025-10-03", "2025-10-09", "개천절~한글날 추석 황금연휴(대략)"),
        ("2026-02-16", "2026-02-18", "설날 연휴(대략)"),
    ]
    for start, end, label in approx_holidays:
        lines.append(f"  {start} ~ {end}: {label}")
    lines.append("\n위 명절 시기와 이상일 탐지 결과가 겹치는지 아래에서 육안으로 대조 가능.")

    report_path = OUT_DIR / "calendar_anomaly_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved: {report_path}")


if __name__ == "__main__":
    main()
