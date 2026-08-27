# -*- coding: utf-8 -*-
"""
Phase 3: 공휴일 연휴 직전 주문 보정.
07_calendar_effects_check.py 에서 발견한 패턴(2일 이상 연휴 시작 1~3일 전 주문은
3일도달률이 최대 60%p까지 떨어짐)을 실제 보정 계수로 만든다.

방법: 세그먼트(협력사x상품군)를 요일처럼 다시 쪼개지 않는다 - 그건 이미
03_distribution_fit.py에서 시도했다가 표본 파편화로 실패한 접근이다(문서 참고).
대신 전체 데이터로 "공휴일 연휴 직전 vs 평소"의 day-mark별 도착확률 '비율'만 구해서,
기존 세그먼트 예측치에 곱으로 얹는다(세그먼트 자체는 그대로, 후보정 레이어만 추가).

HOLIDAYS는 `holidays` 라이브러리(한국, 인터넷 불필요)로 매 실행마다 동적 생성한다
(2026-07-08 전면 교체 - 이전엔 수동 리스트였고 갱신을 잊어 2026년 하반기 전체가
빠진 사고가 있었음, known_risks.md 참고). 음력 공휴일도 근사치가 아니라 정확한
날짜이고, 제헌절처럼 연도별로 법정공휴일 여부가 바뀌는 것도 정확히 반영한다(실측
확인됨).

[2026-07-06 검증 결과] short/long 버킷은 day7~14 구간이 2026H1 홀드아웃에서 오히려
오차가 커짐(과잉보정) - day1~5만 보정 적용.
[2026-07-08 검증 결과, holidays 라이브러리 교체 전] single 버킷(하루짜리 공휴일,
수동 리스트 22건 기준)은 day7,10까지도 개선됨.
[2026-07-08 재검증, holidays 라이브러리 교체 후] single 버킷 표본이 25,734건으로
늘면서(임시공휴일/선거일 등 다양한 공휴일이 섞임) day5부터 오히려 악화로 바뀜 -
day1~3만 안전. 표본이 다양해질수록(효과 크기가 제각각인 공휴일이 섞일수록) 안전
구간이 좁아지는 경향으로 보임 - HOLIDAYS 소스가 바뀔 때마다 ADJUSTABLE_DAYS_BY_BUCKET도
재검증 필요.
"""
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import date

HERE = Path(__file__).parent
DAY_MARKS = [1, 2, 3, 5, 7, 10, 14, 21, 30]
PRE_HOLIDAY_WINDOW = 3   # 연휴 시작 며칠 전까지를 '직전'으로 볼지
BUCKETS = ["single", "short", "long"]  # 1일 / 2~3일 / 4일+ (2026-07-08: single 추가, 아래 참고)
# 버킷마다 "보정이 실제로 개선되는 구간"이 다름 - 단일 ADJUSTABLE_DAYS로는 이 차이를 못 담아
# 버킷별로 안전 구간을 따로 관리한다. single 버킷의 안전 구간은 HOLIDAYS 소스가 바뀔 때마다
# 재검증 필요(2026-07-08: 수동 리스트 22건 기준일 땐 day7~10까지도 개선됐으나, holidays
# 라이브러리로 교체해 임시공휴일/선거일 등이 추가로 섞이며(25,734건) day5부터 오히려
# 악화로 바뀜 - 표본이 다양해질수록 짧은 구간만 안전해지는 경향, day1~3만 유지).
ADJUSTABLE_DAYS_BY_BUCKET = {
    "single": {1, 2, 3},
    "short": {1, 2, 3, 5},
    "long": {1, 2, 3, 5},
}

# [2026-07-08 추가] 원래 MIN_HOLIDAY_LEN=2로 하루짜리 공휴일(신정/삼일절/광복절 등)을
# "물류센터가 안 멈출 만큼 짧다"고 보고 보정 대상에서 제외했었다. 그런데 사용자가
# 계절성(연말 주문 몰림 등) 질문을 하면서 07_calendar_effects_check.py의 이상일 목록을
# 다시 보니, 광복절(8/15) 직전이 2023~2025년 3년 연속, 신정(1/1) 직전(=12/31 포함)이
# 반복적으로 "이상일"로 잡히고 있었음 - 실측 확인 결과 광복절 직전 3일은 day2 도달률이
# 평소의 40.8%로, 신정 직전은 46.7%로 떨어짐(train 2023-2025, n=19,490/9,293건으로
# 표본도 충분함) - 기존에 보정 중이던 short/long 연휴와 비슷하거나 더 큰 효과. 그래서
# "하루짜리는 무조건 제외"가 틀린 기준이었다고 보고 single 버킷을 신설해 모든 하루짜리
# 공휴일을 포함시키고, 아래 홀드아웃 검증으로 실제 개선 여부를 확인한다.

# [2026-07-08 전면 교체] 원래 수동 관리하던 HOLIDAYS 리스트가 2026-06-06(현충일)까지만
# 있어서 2026년 하반기 전체가 통째로 빠져 있었음(사용자의 "12/1 주문에 연말 계절성
# 반영되나" 질문으로 발견 - known_risks.md 참고). 수동 리스트는 근본적으로: ① 매년
# 갱신을 잊기 쉽고 ② 음력 공휴일(설날/추석/부처님오신날)이 근사치였고 ③ 임시공휴일
# (2023-10-02, 2025-01-27 등)·선거일을 아예 놓쳤다. `holidays` 라이브러리(사내망에
# 이미 설치돼 있어 인터넷 불필요)로 교체 - 검증 결과 연도별 법령 변경까지 정확함
# (예: 제헌절은 2008년 폐지 후 2026년에 재지정됐는데, 라이브러리가 2023~2025엔 제헌절을
# 안 잡고 2026부터만 정확히 잡음 - 실측 확인, 2026-07-08). 라이브러리가 혹시 실제
# 물류 지연과 무관한 날짜를 포함하더라도, 아래 홀드아웃 검증이 "효과 없으면 자동
# 무보정(1.0)" 처리하므로 안전하다 - 임의로 미리 걸러낼 필요 없음.
import holidays as _holidays_lib


def _build_holidays_dynamic(years):
    """`holidays.KR(years=...)`가 반환하는 개별 날짜를, 연속된 날짜끼리 하나의 연휴로
    묶어서 기존 (시작일, 종료일, 이름, 신뢰도) 튜플 형식으로 변환한다."""
    kr = _holidays_lib.KR(years=years)
    dates = sorted(kr.keys())
    ranges = []
    i = 0
    while i < len(dates):
        j = i
        names = [kr[dates[i]]]
        while j + 1 < len(dates) and (dates[j + 1] - dates[j]).days == 1:
            j += 1
            names.append(kr[dates[j]])
        seen = list(dict.fromkeys(names))  # 순서 유지하며 중복 이름 제거
        ranges.append((dates[i].isoformat(), dates[j].isoformat(), "/".join(seen), "holidays 라이브러리"))
        i = j + 1
    return ranges


# train(2023) 이전 여유분 없이 2023부터, 현재 시점 기준 2년 뒤까지 - 재실행할 때마다
# 자동으로 범위가 넓어져서 예전처럼 "리스트가 짧아서 하반기가 빠지는" 사고가 구조적으로
# 재발하지 않는다.
HOLIDAYS = _build_holidays_dynamic(list(range(2023, date.today().year + 3)))


def build_holiday_starts():
    """연휴 길이로 single(1일)/short(2~3일)/long(4일+) 세 묶음으로 나눈다 - 연휴 길이마다
    배송 지연/회복 양상이 다르다(추석 같은 7일 연휴는 회복이 훨씬 느림, 2026-07-06 확인;
    하루짜리도 무시 못할 효과가 있음, 2026-07-08 확인 - 위 BUCKETS 주석 참고)."""
    starts = []
    for start, end, name, conf in HOLIDAYS:
        s, e = pd.Timestamp(start), pd.Timestamp(end)
        length = (e - s).days + 1
        if length >= 4:
            bucket = "long"
        elif length >= 2:
            bucket = "short"
        else:
            bucket = "single"
        starts.append((s, name, conf, bucket))
    return starts


def is_pre_holiday(order_dates: pd.Series, holiday_starts, bucket=None) -> pd.Series:
    flag = pd.Series(False, index=order_dates.index)
    for start, name, conf, b in holiday_starts:
        if bucket is not None and b != bucket:
            continue
        window = (order_dates >= start - pd.Timedelta(days=PRE_HOLIDAY_WINDOW)) & (order_dates < start)
        flag = flag | window
    return flag


def prep(df):
    d = df.copy()
    d["lt"] = (d["배송완료일자"] - d["주문일자"]).dt.days
    valid = (d["lt"].notna() & (d["lt"] >= 0) & (d["lt"] <= 90)
             & (d["취소수량"].fillna(0) == 0) & (d["반품수량"].fillna(0) == 0))
    return d[valid].reset_index(drop=True)


def compute_ratios(train, holiday_starts, bucket, lines):
    train_flag = is_pre_holiday(train["주문일자"], holiday_starts, bucket=bucket)
    normal = train[~is_pre_holiday(train["주문일자"], holiday_starts)]["lt"]  # 어떤 연휴에도 안 걸리는 순수 평소
    holiday = train[train_flag]["lt"]
    lines.append(f"\n-- [{bucket}] 연휴 직전 (n={len(holiday):,}) --")
    lines.append(f"{'day':>4s} {'평소 CDF':>10s} {'연휴직전 CDF':>12s} {'비율':>8s}")
    ratios = {}
    for d in DAY_MARKS:
        p_normal = (normal <= d).mean()
        p_holiday = (holiday <= d).mean()
        ratio = p_holiday / p_normal if p_normal > 0 else 1.0
        ratios[d] = ratio
        lines.append(f"{d:>4d} {p_normal*100:>9.1f}% {p_holiday*100:>11.1f}% {ratio:>7.3f}x")
    return ratios, normal


def validate_bucket(holdout, holiday_starts, bucket, ratios, global_normal_cdf, lines):
    flag = is_pre_holiday(holdout["주문일자"], holiday_starts, bucket=bucket)
    ho = holdout[flag]
    if len(ho) == 0:
        lines.append(f"\n[{bucket}] 홀드아웃에 해당 사례 없음")
        return
    lines.append(f"\n검증 [{bucket}] (2026H1, n={len(ho):,}):")
    lines.append(f"{'day':>4s} {'실제도착률':>10s} {'미보정':>10s} {'보정후':>10s} {'미보정오차':>10s} {'보정후오차':>10s}")
    for d in DAY_MARKS:
        actual = (ho["lt"] <= d).mean()
        unadjusted = global_normal_cdf[d]
        adjusted = min(1.0, unadjusted * ratios[d])
        lines.append(f"{d:>4d} {actual*100:>9.1f}% {unadjusted*100:>9.1f}% {adjusted*100:>9.1f}% "
                     f"{abs(unadjusted-actual)*100:>9.1f}%p {abs(adjusted-actual)*100:>9.1f}%p")


def main():
    # serve.py가 build_holiday_starts()만 가져다 쓰려고 이 파일 전체를 동적 임포트하는데
    # (importlib), 이 파일 최상단에 data_loader를 import해두면 그 경로 하드코딩까지
    # 서빙 배포에 딸려간다 - main() 전용 로직이니 여기서만 지역 임포트(2026-07-27).
    from data_loader import load_many, SEGMENTS

    holiday_starts = build_holiday_starts()

    print("Train 데이터 로딩 (2023~2025, 보정계수 학습용)...", flush=True)
    train = prep(load_many(SEGMENTS, [2023, 2024, 2025]))

    lines = ["공휴일 연휴 직전 보정계수 산출 (train: 2023~2025) - 연휴 길이별(single=1일, short=2~3일, long=4일+) 분리"]
    any_flag = is_pre_holiday(train["주문일자"], holiday_starts)
    lines.append(f"pre_holiday 전체 대상 건수: {any_flag.sum():,} / {len(train):,} ({any_flag.mean()*100:.2f}%)")

    ratios_by_bucket = {}
    normal_by_bucket = {}
    for bucket in BUCKETS:
        ratios_by_bucket[bucket], normal_by_bucket[bucket] = compute_ratios(train, holiday_starts, bucket, lines)

    print("Holdout(2026H1)으로 보정 효과 검증 중...", flush=True)
    holdout = prep(load_many(SEGMENTS, [2026]))
    global_normal = train[~any_flag]["lt"]
    global_normal_cdf = {d: (global_normal <= d).mean() for d in DAY_MARKS}

    for bucket in BUCKETS:
        validate_bucket(holdout, holiday_starts, bucket, ratios_by_bucket[bucket], global_normal_cdf, lines)

    lines.append("\n※ '보정후오차'가 '미보정오차'보다 작아야 이 보정이 실제로 유효한 것.")
    lines.append("※ HOLIDAYS는 holidays 라이브러리로 매 실행마다 동적 생성됨(2026-07-08부터) - 수동 갱신 불필요.")

    report_path = HERE / "holiday_adjustment_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print(f"\nSaved: {report_path}")

    rows = []
    for bucket in BUCKETS:
        adjustable = ADJUSTABLE_DAYS_BY_BUCKET[bucket]
        for d, r in ratios_by_bucket[bucket].items():
            applied = r if d in adjustable else 1.0  # 검증 안 된 구간은 무보정(1.0)으로 저장
            rows.append({"bucket": bucket, "day": d, "raw_ratio": r, "applied_ratio": applied})
    ratio_df = pd.DataFrame(rows)
    ratio_df.to_csv(HERE / "holiday_adjustment_ratios.csv", index=False)
    print("Ratio table saved: holiday_adjustment_ratios.csv (버킷별 적용 구간: "
          + ", ".join(f"{b}={sorted(ADJUSTABLE_DAYS_BY_BUCKET[b])}" for b in BUCKETS) + ")")


if __name__ == "__main__":
    main()
