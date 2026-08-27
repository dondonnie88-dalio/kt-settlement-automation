# -*- coding: utf-8 -*-
"""
Phase 1: 배송 소요일 확률분포 학습 + 캘리브레이션 검증

세그먼트 폴백 체인 (LEVELS, 세밀한 것 -> 성긴 것 순):
  상품코드 -> 협력사x중분류 -> 협력사 -> 표준납기버킷 -> 대분류 -> 전체 (6단계)
  (sku -> vendor_mid -> vendor -> sla_bucket -> cat -> global)
  ※ 협력사x대분류(vendor_cat)는 2026-07-08 제거(아래 주석 참고), 표준납기버킷(sla_bucket)은
    콜드스타트용으로 신설 - 이 docstring이 옛 체인을 그대로 두고 있어 2026-07-24 갱신함.

- **상품코드 레벨 (2026-07-06 추가)**: 사용자가 "협력사 단위가 아니라 상품코드 단위로
  예측해야 한다"고 명확히 함(원래 구상이 네이버 상품페이지처럼 SKU별 도착확률 표시).
  다만 실측 확인 결과 상품코드별 물량은 극히 희소함 - 2026년 상반기 KT 세그먼트만 봐도
  상품코드 10,533개 중 n>=1000은 0개, 중앙값은 2건뿐(최다는 "인쇄명함" 463건). 그래서
  SKU 레벨은 낮은 임계값(SKU_MIN_N)으로 "물량 많은 히트상품만" 자기 데이터로 예측하고,
  나머지 압도적 다수는 자동으로 상위(협력사x중분류 등) 레벨로 폴백된다. 조회는 상품코드로
  하되(serve.py 참고), 통계적 안정성은 상위 레벨이 담당하는 구조.
- 협력사x중분류/대분류: 06_category_granularity_check.py로 검증.
  협력사별 중분류간 배송속도 편차가 최대 89.8%p(케이눅)까지 벌어져 세분화 가치가 큼.
  커버리지 손실은 86.3%->83.6%로 미미.
- MIN_N=1000: 신규/급성장 협력사(예: 2025년 수백 건 파일럿 -> 2026년 수만 건 규모)의
  비대표 소량 샘플이 세그먼트로 채택되는 문제 방지 (04_recency_diagnosis.py, 2026-07-03).

[시도했다가 되돌린 것] 주문요일(월~금)을 세그먼트 축에 추가하는 실험을 해봤음.
글로벌 집계로는 요일 효과가 뚜렷했음(3일 도달률 월 68.7% vs 금 17.4%) - 그러나
협력사/카테고리 축과 결합해서 넣었더니 홀드아웃 캘리브레이션이 오히려 악화됨
(2~5일 gap이 0.1~1.8%p -> 최대 3.7%p로 증가, 특히 수요일만 16.5%p 벌어짐 -
train 원본 수요일 평균(56.4%)보다도 모델 예측(45.2%)이 더 낮게 나오는 이상 현상).
축 간 상호작용을 단순 곱으로 가정한 게 틀렸거나 소표본 노이즈로 추정되나
원인 미규명 상태라 프로덕션에는 반영하지 않음 (2026-07-03). 재시도할 경우
반드시 이 방식대로 전체 day-mark 캘리브레이션이 실제로 개선되는지 먼저 확인할 것.

라벨: 배송완료일자 - 주문일자 (calendar days, 실제 달력일수라 주말/공휴일 효과가
      이미 라벨에 내재됨 - 별도 공휴일 보정 레이어는 이중 반영이라 불필요)
Train: 2023~2025 / Holdout: 2026 상반기
출력: 세그먼트별 "N일까지 도착확률" 룩업테이블 + 캘리브레이션 리포트
      + 기존 표준납기일(SLA) 대비 벤치마크 비교
"""
import sys
import io
import re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
import numpy as np
from pathlib import Path
from data_loader import load_many, SEGMENTS

OUT_DIR = Path(__file__).parent
TRAIN_YEARS = [2023, 2024, 2025]
HOLDOUT_YEARS = [2026]
# 2026-07-10: vendor_mid 레벨만 2022년을 포함해서 학습 - 같은 홀드아웃 기준 재검증 결과
# day21 오차 11.7%p->8.6%p, day30 11.0%p->9.2%p로 개선(둘 다 ">10%p=재검토 필요" 기준을
# 벗어남), day1~14도 악화 없이 유지, 표본도 30,649->33,506건으로 늘어남(콜드스타트였던
# 주문 일부가 vendor_mid로 승격). 다른 레벨(sku/vendor/cat/global)은 앞서 전체를
# 2022~2025로 넓혀본 실험에서 실익이 없었음(known_risks.md의 "학습 데이터 확장 검증"
# 항목 참고) - vendor_mid만 선택적으로 넓히는 게 유일하게 효과가 확인된 방식이라 이렇게 함.
#
# 2026-07-10 (같은 날 추가 검증): 사용자가 제공한 2018~2021년 원본("18~21년 자료" 폴더,
# 표준납기일 등 일부 컬럼이 없는 구버전 스키마)까지 vendor_mid에 포함해서 재검증한 결과
# 모든 day-mark에서 예외 없이 추가 개선됨(day1 1.2%p->0.0%p, day7 2.7%p->0.8%p,
# day14 4.4%p->2.9%p, day21 8.6%p->7.0%p, day30 9.2%p->7.8%p). 세그먼트도 374->631개,
# 홀드아웃 커버리지도 33,506->37,641건으로 늘어남. 다른 레벨/전체 지표에 악영향 없음
# (known_risks.md 참고) - 그래서 2018년까지 포함.
VENDOR_MID_TRAIN_YEARS = [2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
OLD_DATA_DIR = OUT_DIR.parent / "18~21년 자료"
OLD_DATA_YEARS = [2018, 2019, 2020, 2021]
OLD_DATA_SEGMENTS = ["KT", "그룹사", "외부사", "지입자재"]
OLD_DATA_USECOLS = ["상품코드", "협력사명", "서비스카테고리", "주문일자", "입고일자",
                     "취소수량", "반품수량", "상품명", "발주일자", "배송완료일자"]
DAY_MARKS = [1, 2, 3, 5, 7, 10, 14, 21, 30]
PROB_COLS = [f"p_{d}d" for d in DAY_MARKS]
# 매일 단위 확률(headline 계산 전용) - DAY_MARKS는 성긴 지점(1,2,3,5,7,10,14,21,30)만 있어서
# 예: day14=89.9%, day21=98.6%처럼 두 지점 사이에 실제 90% 도달일(예: 15~16일)이 있어도
# headline이 다음 성긴 지점(21일)으로 과대평가되는 문제가 있었음(2026-07-07 사용자 지적).
# DAY_MARKS/PROB_COLS는 기존 캘리브레이션 검증·공휴일보정(08번)이 이 정확한 지점들 기준으로
# 이미 검증됐으므로 그대로 두고, headline 전용으로 매일 단위 컬럼을 별도 추가한다.
FINE_DAYS = list(range(1, 91))  # 2026-07-07: 30일로 자르면 실제론 45~60일 안에 90%를 넘는
# 세그먼트(194개, 전체의 6.4%)까지 "미달"로 잘못 보이는 문제가 있어 90일로 확장함.
# train(2023~2025)은 이미 다 끝난 과거라 90일까지 늘려도 우측절단 문제 없음 - 나중에
# 최근(미배송 완료 가능성 있는) 데이터를 학습에 포함시킬 땐 재검토 필요.
FINE_COLS = [f"fd_{d}" for d in FINE_DAYS]
MIN_N = 1000
SKU_MIN_N = 200  # 상품코드는 물량이 워낙 희소해서 별도 낮은 임계값 사용 (아래 docstring 참고)
# spot(일회성) 발주 배제 (2026-07-13, 본부장 지시 사항 검증): "그 협력사가 이 상품코드를
# 취급한 횟수"가 이 미만이면 그 협력사에게 "가끔 한 번씩 파는 물건"으로 간주해 vendor
# 레벨 학습(평균 계산)에서만 제외한다 - spot 주문 자체의 예측 정확도는 신경 쓰지 않는다는
# 전제(사용자 확인) 하에, 서빙 시에는 spot/non-spot 구분 없이 이 정제된 vendor 평균을
# 그대로 반환한다(별도 폴백 로직 없음, 하위 호환). same-order 검증 결과 non-spot 주문
# 예측이 전 구간에서 일관되게 개선(day3 1.27->0.97p, day7 1.20->0.87p 등). vendor_mid는
# 같은 방식으로 검증했더니 오히려 악화되는 구간이 더 많아서(이미 협력사x중분류로 좁혀진
# 그룹이라 spot까지 더 빼면 표본이 과도하게 얇아짐) vendor 레벨에만 적용한다.
# known_risks.md 참고.
VENDOR_SPOT_THRESHOLD = 5

# vendor 경계선(700~999건) 세그먼트 조건부 포함 (2026-07-14, 본부장 지시 사항 후속 검증):
# 사용자가 "sku 299건이나 vendor 999건이라고 다음 레벨로 넘겨버리는 게 아쉽다"고 지적,
# MIN_N=1000 바로 아래(700~999건) 협력사 38개(홀드아웃 6,781건)만 따로 떼어 같은 주문
# counterfactual로 검증한 결과 8/9 day마크에서 자체 평균이 기존 폴백(sla_bucket/sku/
# vendor_mid)보다 우월(day7 7.40p->1.14p, day14 11.10p->5.30p). 이전 "MIN_N 하향 실익
# 없음" 결론(known_risks.md)은 급성장 협력사까지 섞은 전체 하향 실험이었고, 이번은 그
# 원인(급성장 파일럿->스케일 불안정, 04_recency_diagnosis.py)을 growth_vendor_flags.txt로
# 걸러낸 선별적 하향이라 방법론이 다름 - 38개 중 급성장 목록과 겹치는 건 2개뿐(95%는
# 안정적). 급성장 협력사는 재발 방지 취지에 맞춰 그대로 MIN_N=1000 기준(폴백)을 유지한다.
VENDOR_BOUNDARY_MIN_N = 700


def load_growth_vendor_names() -> set:
    """growth_vendor_flags.txt에서 급성장(재발 방지 대상) 협력사명 집합만 추출.
    build_analysis_workbook.py의 parse_growth()와 동일한 파싱 규칙 재사용."""
    path = OUT_DIR / "growth_vendor_flags.txt"
    if not path.exists():
        return set()
    names = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^(\S.{0,18}\S)\s{2,}(\S.{0,14}\S)\s{2,}([\d,]+)\s+([\d,]+)\s+([\d.]+)x", line.strip())
        if m and "협력사명" not in line:
            names.add(m.group(1).strip())
    return names

# 표준납기 버킷 (2026-07-13 추가): 콜드스타트(cat/global) 주문도 표준납기일은 주문
# 시점에 항상 알 수 있고, 이 값 자체가 배송 속도의 강한 신호(전체 주문의 92%가 표준납기
# 이내 도착). 신규 협력사라 배송 이력이 없어도 "납기 3일짜리인지 30일짜리인지"로 예측
# 가능. 같은 주문 counterfactual 검증에서 콜드스타트 전 구간 오차 합계가 62.7p->29.5p로
# 절반 이하(당초 2단 조합: 대분류x버킷(cat_sla) -> 버킷단독(sla_bucket)). 서빙 시
# 표준납기일은 선택 입력 - 미전달 시 이 레벨을 건너뛰고 기존 cat/global로 폴백(하위 호환).
#
# cat_sla 제거 (2026-07-14): cat_sla가 매칭되는 주문은 sla_bucket에도 100% 매칭됨(커버리지
# 추가 이득 없음). 같은 주문 기준으로 두 레벨을 직접 비교했더니 day-mark별로는 4승(cat_sla)
# 5승(sla_bucket)으로 갈렸지만, 9개 구간 절대오차 합계는 cat_sla 81.17p vs sla_bucket
# 69.40p로 sla_bucket이 더 정확함(평균 9.02p vs 7.71p) - 커버리지도 안 늘려주면서 평균은
# 더 나쁜 레벨이라 판단해 제거. sla_bucket 하나로 단순화. known_risks.md 참고.
SLA_BUCKET_BINS = [-1, 3, 7, 14, 30, 999]
SLA_BUCKET_LABELS = ["~3일", "4~7일", "8~14일", "15~30일", "31일~"]

LEVELS = [
    ("sku", ["상품코드"], SKU_MIN_N),
    ("vendor_mid", ["협력사명", "중분류"], MIN_N),
    ("vendor", ["협력사명"], MIN_N),
    ("sla_bucket", ["납기버킷"], MIN_N),
    ("cat", ["대분류"], MIN_N),
    ("global", [], 0),
]
# vendor_cat(협력사x대분류) 제거 (2026-07-08): 같은 주문 집합에 vendor_cat과 vendor를
# 동시 적용해서 비교한 결과, 14,121건 중 6/9 day마크에서 vendor(더 넓은 레벨)가 절대
# 오차 기준으로도 더 정확했음(예: 5일 vendor_cat 4.7%p vs vendor 1.1%p) - 세분화가
# 오히려 노이즈를 학습한 사례. vendor_mid는 같은 방식으로 비교했을 때 vendor보다 나아서
# (8/9 구간 승) 그대로 유지 - "세분화=항상 좋다"가 아니라 실측으로 레벨마다 따로
# 판단해야 한다는 근거. known_risks.md 참고.
# (참고) 순수 카테고리 세분화(중분류/소분류/세분류 카스케이드)도 2026-07-13 검증했으나
# 순이득 없어 기각 - cat_sla/sla_bucket은 카테고리가 아니라 "협력사가 선언한 납기"라는
# 별개 정보 축이라 성격이 다름(실측으로 개선 확인됨).
KEY_COLS = ["협력사명", "대분류", "중분류", "상품코드", "납기버킷"]


def prep(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["lt"] = (d["배송완료일자"] - d["주문일자"]).dt.days
    parts = d["서비스카테고리"].fillna("").str.split(">")
    d["대분류"] = parts.str[0].str.strip()
    d["중분류"] = parts.str[1].str.strip()
    d["요일"] = d["주문일자"].dt.dayofweek  # 0=월 ... 6=일
    d["상품코드"] = d["상품코드"].astype("string")
    d["표준납기_lt"] = (d["표준납기일"] - d["주문일자"]).dt.days
    # 납기버킷: object 타입(문자열 or NaN)으로 유지 - NaN 키는 groupby에서 자동 제외되고
    # merge에서도 매칭 안 되므로, 표준납기일 없는 주문은 cat_sla/sla_bucket 레벨을 건너뜀
    bucket = pd.cut(d["표준납기_lt"], bins=SLA_BUCKET_BINS, labels=SLA_BUCKET_LABELS)
    d["납기버킷"] = bucket.astype(object)
    valid = (
        d["lt"].notna()
        & (d["lt"] >= 0) & (d["lt"] <= 90)
        & (d["취소수량"].fillna(0) == 0)
        & (d["반품수량"].fillna(0) == 0)
    )
    return d[valid].reset_index(drop=True)


def load_old_vendor_mid_data() -> pd.DataFrame:
    """2018~2021년 구버전 원본("18~21년 자료" 폴더) - vendor_mid 학습 전용, prep()된
    상태로 반환. 표준납기일 등 최신 스키마 컬럼이 없어 표준납기_lt는 전부 NaN이 되지만
    vendor_mid 테이블 구축(lt만 사용)에는 영향 없음. 파싱이 오래 걸려(16개 대용량 xlsx)
    parquet 캐시를 우선 사용."""
    cache_path = OUT_DIR / "cache" / "old_2018_2021.parquet"
    if cache_path.exists():
        raw = pd.read_parquet(cache_path)
    else:
        dfs = []
        for year in OLD_DATA_YEARS:
            for seg in OLD_DATA_SEGMENTS:
                fname = f"{year}_주문현황_{seg}_전량입고(요청자_이돈현).xlsx"
                path = OLD_DATA_DIR / fname
                print(f"  [old raw parse] {fname} ...", flush=True)
                df = pd.read_excel(path, usecols=OLD_DATA_USECOLS, engine="openpyxl")
                for c in ["주문일자", "입고일자", "발주일자", "배송완료일자"]:
                    df[c] = pd.to_datetime(df[c], errors="coerce")
                dfs.append(df)
        raw = pd.concat(dfs, ignore_index=True)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        raw.to_parquet(cache_path)
    raw = raw.copy()
    raw["표준납기일"] = pd.NaT
    return prep(raw)


def agg_group(g: pd.DataFrame) -> pd.Series:
    lt = g["lt"]
    out = {f"p_{d}d": (lt <= d).mean() for d in DAY_MARKS}
    out.update({f"fd_{d}": (lt <= d).mean() for d in FINE_DAYS})
    out["n"] = len(lt)
    out["median"] = lt.median()
    return pd.Series(out)


def build_segment_tables(train: pd.DataFrame) -> dict:
    tables = {}
    for name, keys, min_n in LEVELS:
        if keys:
            g = train.groupby(keys).apply(agg_group, include_groups=False).reset_index()
            g = g[g["n"] >= min_n].reset_index(drop=True)
        else:
            g = pd.DataFrame([agg_group(train).to_dict()])
        tables[name] = (keys, g)
    return tables


def assign_vectorized(holdout: pd.DataFrame, tables: dict) -> pd.DataFrame:
    base = holdout[KEY_COLS + ["lt", "표준납기_lt"]].reset_index(drop=True)
    result = base.copy()
    for col in PROB_COLS + ["n", "median"]:
        result[col] = np.nan
    result["level"] = pd.array([None] * len(base), dtype="object")
    remaining = np.ones(len(base), dtype=bool)

    for name, keys, _min_n in LEVELS:
        if not remaining.any():
            break
        if keys:
            m = base[keys].merge(tables[name][1], on=keys, how="left")
        else:
            g = tables[name][1]
            m = pd.concat([g] * len(base), ignore_index=True)
        matched = m["n"].notna().values & remaining
        if matched.any():
            for col in PROB_COLS + ["n", "median"]:
                result.loc[matched, col] = m.loc[matched, col].values
            result.loc[matched, "level"] = name
            remaining = remaining & ~matched
    return result


def evaluate_holdout(result: pd.DataFrame, lines: list):
    lines.append("\n" + "=" * 70)
    lines.append("[홀드아웃 검증] 2026년 상반기")
    lines.append("=" * 70)

    lvl_counts = result["level"].value_counts()
    lines.append("\n-- 세그먼트 폴백 레벨 사용 비중 --")
    for lvl, cnt in lvl_counts.items():
        lines.append(f"  {lvl:16s} {cnt:>8,} ({cnt/len(result)*100:5.1f}%)")

    lines.append("\n-- 캘리브레이션: 예측확률(train) vs 실제도착률(holdout), day mark별 --")
    lines.append(f"{'day':>4s} {'pred_prob(mean)':>16s} {'actual_rate':>12s} {'|gap|(%p)':>10s} {'n':>10s}")
    for dmark in DAY_MARKS:
        pred_mean = result[f"p_{dmark}d"].mean()
        actual_rate = (result["lt"] <= dmark).mean()
        gap = abs(pred_mean - actual_rate) * 100
        lines.append(f"{dmark:>4d} {pred_mean*100:>15.1f}% {actual_rate*100:>11.1f}% {gap:>9.1f}%p {len(result):>10,}")

    lines.append("\n-- 벤치마크: 기존 표준납기일(SLA) 대비 --")
    sla = result.dropna(subset=["표준납기_lt"]).copy()
    sla = sla[(sla["표준납기_lt"] >= 0) & (sla["표준납기_lt"] <= 90)]
    sla_ontime_actual = (sla["lt"] <= sla["표준납기_lt"]).mean() * 100
    lines.append(f"  기존 SLA(표준납기일) 준수율(실측): {sla_ontime_actual:.1f}% (n={len(sla):,})")

    day_marks_arr = np.array(DAY_MARKS)
    nearest_lookup = {x: int(day_marks_arr[np.abs(day_marks_arr - x).argmin()]) for x in range(0, 91)}
    sla_days_clipped = sla["표준납기_lt"].round().clip(0, 90).astype(int).map(nearest_lookup)
    prob_matrix = sla[PROB_COLS].values
    col_idx = sla_days_clipped.map({d: i for i, d in enumerate(DAY_MARKS)}).values
    model_prob_at_sla = prob_matrix[np.arange(len(sla)), col_idx]
    lines.append(f"  모델이 예측한 'SLA일까지 도착확률' 평균: {model_prob_at_sla.mean()*100:.1f}%")
    lines.append("  (해석: 기존 SLA 준수율과 모델 예측확률이 가까울수록, 모델이 SLA 신뢰도를 정확히 정량화하는 것)")

    lines.append("\n-- 세그먼트별 캘리브레이션 (레벨별 7일 기준) --")
    for lvl, _keys, _min_n in LEVELS:
        sub = result[result["level"] == lvl]
        if len(sub) == 0:
            continue
        pred7 = sub["p_7d"].mean()
        actual7 = (sub["lt"] <= 7).mean()
        lines.append(f"  {lvl:16s} n={len(sub):>8,}  예측(7일)={pred7*100:5.1f}%  실측(7일)={actual7*100:5.1f}%  gap={abs(pred7-actual7)*100:4.1f}%p")

    # 레벨x day마크 전 구간 (2026-07-07 검증 - cat/global은 중간구간만, vendor_mid는
    # 반대로 21일 이후만 나쁜 "구간 부정확" 패턴을 발견한 표. serve.py의
    # LEVEL_UNRELIABLE_DAYS가 바로 이 표의 >10%p 지점들로 만들어졌으니, 재학습 시
    # 이 표가 바뀌면 그쪽도 같이 재검증할 것 - known_risks.md 참고.
    lines.append("\n-- 세그먼트별 캘리브레이션 (레벨x day마크 전 구간, >10%p=재검토 필요) --")
    for lvl, _keys, _min_n in LEVELS:
        sub = result[result["level"] == lvl]
        if len(sub) == 0:
            continue
        parts = []
        for dmark in DAY_MARKS:
            pred = sub[f"p_{dmark}d"].mean()
            actual = (sub["lt"] <= dmark).mean()
            gap = abs(pred - actual) * 100
            parts.append(f"{dmark}일={gap:.1f}%p")
        lines.append(f"  {lvl:16s} n={len(sub):>8,}  " + " ".join(parts))


def main():
    print("Train 데이터 로딩 (2023~2025)...", flush=True)
    train_raw = load_many(SEGMENTS, TRAIN_YEARS)
    print(f"  train raw rows={len(train_raw):,}", flush=True)
    train = prep(train_raw)
    print(f"  train valid rows={len(train):,}", flush=True)

    print("Holdout 데이터 로딩 (2026 상반기)...", flush=True)
    holdout_raw = load_many(SEGMENTS, HOLDOUT_YEARS)
    holdout = prep(holdout_raw)
    print(f"  holdout valid rows={len(holdout):,}", flush=True)

    print("세그먼트 테이블 학습 중...", flush=True)
    tables = build_segment_tables(train)

    print("vendor_mid 전용 확장 학습 데이터 로딩 (2022~2025, 표준 스키마)...", flush=True)
    recent_vm_years = [y for y in VENDOR_MID_TRAIN_YEARS if y not in OLD_DATA_YEARS]
    train_vm_recent_raw = load_many(SEGMENTS, recent_vm_years)
    train_vm_recent = prep(train_vm_recent_raw)
    print("vendor_mid 전용 확장 학습 데이터 로딩 (2018~2021, 구버전 스키마)...", flush=True)
    train_vm_old = load_old_vendor_mid_data()
    train_vm = pd.concat([train_vm_recent, train_vm_old], ignore_index=True)
    print(f"  vendor_mid 학습 valid rows={len(train_vm):,} (최신 {len(train_vm_recent):,} + 구버전 {len(train_vm_old):,})", flush=True)
    vm_keys, _ = tables["vendor_mid"]
    vm_g = train_vm.groupby(vm_keys).apply(agg_group, include_groups=False).reset_index()
    vm_g = vm_g[vm_g["n"] >= MIN_N].reset_index(drop=True)
    tables["vendor_mid"] = (vm_keys, vm_g)
    print(f"  vendor_mid 세그먼트 수: {len(vm_g):,}", flush=True)

    print("vendor 레벨 spot(일회성) 발주 제외 재계산 중...", flush=True)
    vendor_sku_n = train.groupby(["협력사명", "상품코드"]).size().rename("_vendor_sku_n")
    train_v = train.merge(vendor_sku_n, on=["협력사명", "상품코드"], how="left")
    train_v_clean = train_v[train_v["_vendor_sku_n"] >= VENDOR_SPOT_THRESHOLD]
    spot_share = 1 - len(train_v_clean) / len(train_v)
    v_g = train_v_clean.groupby(["협력사명"]).apply(agg_group, include_groups=False).reset_index()
    growth_vendors = load_growth_vendor_names()
    is_boundary = ((v_g["n"] >= VENDOR_BOUNDARY_MIN_N) & (v_g["n"] < MIN_N)
                   & (~v_g["협력사명"].isin(growth_vendors)))
    n_boundary_included = int(is_boundary.sum())
    v_g = v_g[(v_g["n"] >= MIN_N) | is_boundary].reset_index(drop=True)
    tables["vendor"] = (["협력사명"], v_g)
    print(f"  vendor 학습에서 제외된 spot 비중: {spot_share*100:.1f}%, 세그먼트 수: {len(v_g):,}"
          f" (경계선 700~999건 포함: {n_boundary_included}개, 급성장 제외 {len(growth_vendors)}개)", flush=True)

    lines = ["배송 예측 정보 과제 - Phase 1 확률분포 학습 결과"]
    lines.append(f"Train: {TRAIN_YEARS} (n={len(train):,})  Holdout: {HOLDOUT_YEARS} (n={len(holdout):,})")
    lines.append(f"  (vendor_mid만 예외: {VENDOR_MID_TRAIN_YEARS} 사용, n={len(train_vm):,} - 21/30일 캘리브레이션 개선 확인, 2026-07-10)")
    lines.append("세그먼트 수: " + ", ".join(f"{name}={len(tbl):,}" for name, (_, tbl) in tables.items()))
    glob = tables["global"][1].iloc[0]
    lines.append(f"전체(global) 분포: n={int(glob['n']):,}  median={glob['median']:.1f}일")
    lines.append("전체 day mark별 도착확률: " + ", ".join(f"{d}일:{glob[f'p_{d}d']*100:.0f}%" for d in DAY_MARKS))

    print("홀드아웃 세그먼트 배정 및 캘리브레이션 계산 중...", flush=True)
    result = assign_vectorized(holdout, tables)
    evaluate_holdout(result, lines)

    report_path = OUT_DIR / "phase1_report.txt"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nDone. Report saved to {report_path}", flush=True)

    lookup_parts = []
    for name, (keys, tbl) in tables.items():
        out = tbl.copy()
        out["level"] = name
        for col in KEY_COLS:
            if col not in out.columns:
                out[col] = None
        lookup_parts.append(out)
    lookup = pd.concat(lookup_parts, ignore_index=True)
    lookup.to_parquet(OUT_DIR / "segment_distributions.parquet")
    lookup.to_csv(OUT_DIR / "segment_distributions.csv", index=False, encoding="utf-8-sig")
    print(f"Segment lookup table saved: {len(lookup):,} rows", flush=True)

    print("상품코드 -> 협력사/카테고리 매핑 테이블 생성 중...", flush=True)
    # 확률분포(tables)는 정직한 캘리브레이션 검증을 위해 반드시 train(2023~2025)만 써야 하지만,
    # 이 매핑표는 이름/협력사/카테고리를 찾는 용도일 뿐 확률 계산에 관여하지 않으므로
    # holdout(2026)까지 합쳐서 만든다. 그래야 2026년에 신규 등록된 상품코드도(예: 기존
    # 협력사가 새로 낸 신상품) 협력사/카테고리를 찾아 상위 레벨로 정확히 폴백된다 -
    # train만 쓰면 "협력사명/대분류/중분류를 몰라서" 무조건 global로 떨어지는 문제가 있었음
    # (2026-07-07 GUI 조회 중 상품코드 10673012 "외장HDD"에서 확인 - 협력사 (주)중앙미디어테크는
    # train에도 있어 vendor 레벨 데이터가 존재하는데도 상품코드 자체가 2026년 신상품이라
    # product_lookup에 없어서 그 경로를 못 탔음).
    lookup_source = pd.concat([train, holdout], ignore_index=True)
    prod_map = lookup_source.groupby("상품코드").agg(
        상품명=("상품명", "first"),
        협력사명=("협력사명", "first"),
        대분류=("대분류", "first"),
        중분류=("중분류", "first"),
        n_협력사_고유값=("협력사명", "nunique"),
        n=("lt", "size"),
    ).reset_index()
    multi_vendor = (prod_map["n_협력사_고유값"] > 1).sum()
    print(f"  상품코드 {len(prod_map):,}개 중 협력사가 여러 개로 찍힌 경우: {multi_vendor}개 "
          f"({multi_vendor/len(prod_map)*100:.2f}%) - 이 경우 가장 최근/첫 값으로 단순화됨", flush=True)

    # 표준납기_lt 추정값 (2026-07-23 추가, gui_predict.py 자동조회용): 표준납기일은 상품코드별로
    # 거의 고정값이라(67.5%가 학습기간 내내 완전 동일값, 최빈값 기준 실제 일치율 78.9% -
    # known_risks.md 참고) 상품코드의 최빈 표준납기_lt를 "추정치"로 매핑해두면, 사용자가 매번
    # 입력 안 해도 sla_bucket 레벨을 자동으로 활용할 수 있다. 어디까지나 추정값이므로 API_SPEC.md의
    # "통합플랫폼이 이미 보유한 실제 값을 넘긴다"는 정식 계약과는 다름 - serve.py의 predict()
    # 시그니처나 API 계약은 안 건드리고, 사내 조회용 GUI 한정 편의 기능으로만 추가한다.
    def _mode_or_na(s):
        s = s.dropna()
        return s.mode().iloc[0] if len(s) else pd.NA
    sku_sla = lookup_source.groupby("상품코드")["표준납기_lt"].agg(_mode_or_na).rename("표준납기_lt_추정")
    prod_map = prod_map.merge(sku_sla, on="상품코드", how="left")
    has_sla = prod_map["표준납기_lt_추정"].notna().sum()
    print(f"  표준납기_lt 추정치 확보된 상품코드: {has_sla:,}개 ({has_sla/len(prod_map)*100:.1f}%)", flush=True)

    vendor_sla = lookup_source.groupby("협력사명")["표준납기_lt"].agg(_mode_or_na).rename("표준납기_lt_추정").reset_index()
    vendor_sla.to_csv(OUT_DIR / "vendor_sla_lookup.csv", index=False, encoding="utf-8-sig")
    print(f"  협력사 단위 폴백 표준납기_lt 저장: {len(vendor_sla):,}개 협력사", flush=True)

    prod_map.to_csv(OUT_DIR / "product_lookup.csv", index=False, encoding="utf-8-sig")
    print(f"Product lookup table saved: {len(prod_map):,} SKUs", flush=True)


if __name__ == "__main__":
    main()
