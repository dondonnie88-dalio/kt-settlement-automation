# -*- coding: utf-8 -*-
"""
Phase 2: 서빙 로직 (통합플랫폼 API 뒤에 붙일 실제 조회 함수)
segment_distributions.csv를 읽어서, 주문 시점에 알 수 있는 정보만으로 "N일까지 도착확률"과
"실제 달력 날짜별 확률"을 반환한다.

**조회 기준은 상품코드다** (2026-07-06 재설계 - 원래 요구사항이 협력사 단위가 아니라
네이버 상품페이지처럼 SKU별 도착확률이었음). 상품코드 하나만 넘기면 product_lookup.csv로
협력사/대분류/중분류를 자동으로 찾아서 폴백 체인을 탄다. 다만 상품코드별 주문 이력은
대부분 극히 적어서(중앙값 2건/반년), 실제로 자기 데이터로 예측되는 건 물량 많은 히트상품
뿐이고 나머지 압도적 다수는 자동으로 협력사x중분류 이상 레벨로 폴백된다 - 이게 의도된
동작이다(상품코드 자체 데이터로 예측하기엔 표본이 항상 부족하므로).

폴백 순서(03_distribution_fit.py의 LEVELS와 동일): 상품코드 -> 협력사x중분류 -> 협력사
-> 납기버킷 -> 대분류 -> 전체 (2026-07-08: 협력사x대분류 레벨 제거 -
같은 주문 집합 기준 비교에서 협력사 단독 레벨이 절대 오차로도 더 정확해서 세분화 가치가
없었음. 2026-07-13: 표준납기일(standard_delivery_date)을 넘긴 경우에만 납기버킷 레벨을
거침(당초 대분류x납기버킷도 같이 뒀으나 2026-07-14 제거 - 커버리지 이득 없이 평균 오차만
더 나빴음) - 협력사 이력이 없는 콜드스타트 주문도 "며칠짜리 납기인지"는 항상
알 수 있다는 점을 활용, 콜드스타트 오차를 cat/global 대비 절반 이하로 줄임. 안 넘기면
곧바로 대분류/전체로 폴백(하위 호환). known_risks.md 참고).

공휴일 보정: 08_holiday_adjustment.py에서 검증한 결과, 연휴 시작 1~3일 전 주문은
day1/2/3/5 도착확률이 실제로 크게 떨어진다(2026H1 검증 완료). day7 이후 구간은
보정 시 오히려 오차가 커져(과잉보정) 적용하지 않는다 - holiday_adjustment_ratios.csv의
applied_ratio가 이미 이 기준으로 저장돼 있음(day7+는 1.0).
"""
from __future__ import annotations
import importlib.util
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).parent
DAY_MARKS = [1, 2, 3, 5, 7, 10, 14, 21, 30]
EXTRA_DAY_MARKS = [45, 60, 90]  # by_day 표에서 headline이 30일을 넘긴 느린 상품에만 추가로 보여줄 지점
FINE_DAYS = list(range(1, 91))  # headline 전용 매일 단위 (03_distribution_fit.py와 동일, 2026-07-07 30->90일 확장)
WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]
PRE_HOLIDAY_WINDOW = 3

# 레벨별 "불안정 구간"(2026-07-07 검증, 2026-07-10 vendor_mid 재검증 후 갱신) - 2026H1
# 홀드아웃 레벨x day마크 캘리브레이션에서 |예측-실측| 오차가 10%p 초과("재검토 필요", 이
# 프로젝트가 이미 쓰는 기준과 동일)한 DAY_MARKS 지점들. 레벨마다 문제 구간의 위치가
# 다르다(cat/global은 중간 구간만 나쁘고 21일 이후 회복) - 그래서 단일 "reliable_from_day"
# 하나로는 표현이 안 되고 레벨별 구간 집합으로 관리한다.
# vendor_mid는 원래 21/30일이 11.6%p/10.7%p로 불안정했으나, 2026-07-10 vendor_mid 레벨만
# 학습 기간을 2018~2025로 넓혀서(03_distribution_fit.py의 VENDOR_MID_TRAIN_YEARS) 재학습한
# 결과 7.0%p/7.8%p로 개선돼 기준(10%p) 아래로 내려와 불안정 구간에서 제외함
# (known_risks.md의 "vendor_mid 21/30일 개선" 항목 참고).
# 재학습(03_distribution_fit.py) 시 이 표도 같이 재검증해서 갱신할 것.
LEVEL_UNRELIABLE_DAYS = {
    "sku": set(),
    "vendor_mid": set(),
    "vendor": set(),
    "sla_bucket": set(),       # 2026-07-13 검증: 전 구간 10%p 이내(최대 6.1%p). cat_sla는
    # 2026-07-14 제거(known_risks.md 참고) - sla_bucket에 커버리지 이득 없이 평균만 더 나빴음.
    "cat": {2, 3, 5, 7, 10},
    "global": {5, 7},
}
# by_day 확장 지점(45/60/90)은 레벨별 day-mark 캘리브레이션 검증 대상이 아니라(30일까지만
# 검증됨) 불안정 여부를 알 수 없음 - 다만 cat/global 모두 21~30일에 이미 회복되는 패턴이라
# 그 이후도 안정적일 것으로 보고 기본 reliable=True로 둔다.
LEVEL_CONFIDENCE = {
    "sku": "high", "vendor_mid": "high", "vendor": "high",
    "sla_bucket": "medium", # 2026-07-13: 전 구간 10%p 이내(최대 6.1%p) - cat/global보다 뚜렷이 안정적
    "cat": "low", "global": "low",
}


def _load_holiday_module():
    spec = importlib.util.spec_from_file_location("holiday_adj", HERE / "08_holiday_adjustment.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class DeliveryPredictor:
    def __init__(self, seg_path: Path = HERE / "segment_distributions.csv",
                 holiday_ratio_path: Path = HERE / "holiday_adjustment_ratios.csv",
                 product_lookup_path: Path = HERE / "product_lookup.csv",
                 shortage_flags_path: Path = HERE / "shortage_flags.csv",
                 vendor_sla_lookup_path: Path = HERE / "vendor_sla_lookup.csv"):
        df = pd.read_csv(seg_path, dtype={"상품코드": "string"})
        self.sku = df[df["level"] == "sku"].set_index("상품코드")
        self.vendor_mid = df[df["level"] == "vendor_mid"].set_index(["협력사명", "중분류"])
        self.vendor = df[df["level"] == "vendor"].set_index("협력사명")
        # 표준납기 버킷 레벨 (2026-07-13, 콜드스타트 개선): 표준납기일이 입력된 경우에만
        # cat/global 앞에서 사용 - 03_distribution_fit.py의 SLA_BUCKET_* 참고. 원래
        # cat_sla(대분류x납기버킷)도 같이 있었으나 2026-07-14 제거(known_risks.md 참고).
        self.sla_bucket = df[df["level"] == "sla_bucket"].set_index("납기버킷")
        self.cat = df[df["level"] == "cat"].set_index("대분류")
        self.glob = df[df["level"] == "global"].iloc[0]

        self.product_lookup = pd.read_csv(product_lookup_path, dtype={"상품코드": "string"}).set_index("상품코드")

        # 표준납기일 자동 추정용 협력사 최빈값 (2026-07-27, gui_predict.py에 있던 로직을
        # 서빙 라이브러리 본체로 이전 - 호출 측이 표준납기일을 못 넘겨도 predict()가 자체
        # 추정치로 sla_bucket까지 활용하게 해서, 이 폴더 데이터 4종만으로 바로 구동 가능하게
        # 함(연동 초기/테스트 단계 대비). 파일 없으면 조용히 빈 딕셔너리로 대체(하위 호환).
        if vendor_sla_lookup_path.exists():
            vdf = pd.read_csv(vendor_sla_lookup_path)
            self.vendor_sla_lookup = dict(zip(vdf["협력사명"], vdf["표준납기_lt_추정"]))
        else:
            self.vendor_sla_lookup = {}

        ratio_df = pd.read_csv(holiday_ratio_path)
        self.holiday_ratios = {
            bucket: dict(zip(g["day"], g["applied_ratio"]))
            for bucket, g in ratio_df.groupby("bucket")
        }
        holiday_mod = _load_holiday_module()
        self.holiday_starts = holiday_mod.build_holiday_starts()

        # 쇼티지 경고 플래그 (2026-07-09 예방적 설계, C안: 파일구조+연동 로직만 - 실제 GUI
        # 노출 방식/API_SPEC 반영은 박형근 차장 미팅 후 확정. 담당자가 협력사로부터 직접
        # 전달받은 "수입 지연 예상" 같은 정보를 상품코드 단위로 기록해두면, 그 기간에 걸친
        # 주문 조회 시 확률분포는 그대로 두고 경고만 얹어서 반환한다(확률을 억지로 재계산
        # 하지 않음 - 검증 안 된 숫자로 곡선을 왜곡하는 걸 피하기 위함, known_risks.md의
        # "임의 보정 지양" 원칙과 동일).
        if shortage_flags_path.exists():
            sf = pd.read_csv(shortage_flags_path, dtype={"상품코드": "string"})
            sf["시작일"] = pd.to_datetime(sf["시작일"])
            sf["종료일"] = pd.to_datetime(sf["종료일"])
            self.shortage_flags = sf
        else:
            self.shortage_flags = pd.DataFrame(columns=["상품코드", "시작일", "종료일", "사유", "입력자", "입력일"])

    def _active_shortage(self, product_code: str | None, order_date: date) -> dict | None:
        """order_date가 이 상품코드에 걸린 쇼티지 플래그 기간(시작일~종료일) 안이면 그 행을 반환."""
        if not product_code or self.shortage_flags.empty:
            return None
        od = pd.Timestamp(order_date)
        hit = self.shortage_flags[
            (self.shortage_flags["상품코드"] == product_code)
            & (self.shortage_flags["시작일"] <= od)
            & (od <= self.shortage_flags["종료일"])
        ]
        if hit.empty:
            return None
        row = hit.iloc[0]
        return {"reason": row["사유"], "since": row["시작일"].date().isoformat(), "until": row["종료일"].date().isoformat()}

    def resolve_product(self, product_code: str) -> dict | None:
        """상품코드 -> 상품명/협력사명/대분류/중분류. product_lookup.csv에 없으면 None(신규/미등록 상품)."""
        if product_code in self.product_lookup.index:
            row = self.product_lookup.loc[product_code]
            return {"상품명": row["상품명"], "협력사명": row["협력사명"], "대분류": row["대분류"], "중분류": row["중분류"]}
        return None

    def _estimate_sla_lt(self, product_code: str) -> int | None:
        """상품코드 -> 표준납기_lt(주문일로부터 표준납기일까지 며칠) 추정치. 상품코드 자체의
        과거 최빈값을 우선 쓰고, 없으면 협력사 최빈값으로 대체, 둘 다 없으면 None(추정 불가 -
        호출 측이 실제 값을 안 준 경우 기존 cat/global 폴백으로 동작). 어디까지나 과거 이력
        기반 추정치이며 실제 계약상 표준납기일과 다를 수 있음 - 호출 측이 실제 값을 알고 있으면
        predict()에 표준납기일을 직접 넘겨서 이 추정을 건너뛰는 것이 항상 더 정확하다."""
        if product_code not in self.product_lookup.index:
            return None
        row = self.product_lookup.loc[product_code]
        val = row.get("표준납기_lt_추정")
        if pd.notna(val):
            return int(val)
        val = self.vendor_sla_lookup.get(row.get("협력사명"))
        return int(val) if pd.notna(val) else None

    @staticmethod
    def _sla_bucket_of(order_date: date, 표준납기일: date | None) -> str | None:
        """표준납기일(달력 날짜) -> 납기버킷 라벨. 03_distribution_fit.py의 SLA_BUCKET_BINS/
        LABELS와 반드시 일치해야 함(학습 때와 같은 구간 정의)."""
        if 표준납기일 is None:
            return None
        days = (표준납기일 - order_date).days
        if days < 0 or days > 90:
            return None  # 학습 때 prep()이 거르는 범위 밖 - 신뢰 불가
        if days <= 3: return "~3일"
        if days <= 7: return "4~7일"
        if days <= 14: return "8~14일"
        if days <= 30: return "15~30일"
        return "31일~"

    def _lookup(self, product_code: str | None, vendor: str | None, 대분류: str | None,
                중분류: str | None, 납기버킷: str | None = None):
        if product_code and product_code in self.sku.index:
            return "sku", self.sku.loc[product_code]
        if 중분류 and vendor and (vendor, 중분류) in self.vendor_mid.index:
            return "vendor_mid", self.vendor_mid.loc[(vendor, 중분류)]
        if vendor and vendor in self.vendor.index:
            return "vendor", self.vendor.loc[vendor]
        # 표준납기 버킷 레벨(2026-07-13): 협력사 이력이 없는 콜드스타트 주문 구제용.
        # 표준납기일이 입력 안 되면(납기버킷=None) 건너뛰고 기존 cat/global로 폴백(하위 호환).
        if 납기버킷 and 납기버킷 in self.sla_bucket.index:
            return "sla_bucket", self.sla_bucket.loc[납기버킷]
        if 대분류 and 대분류 in self.cat.index:
            return "cat", self.cat.loc[대분류]
        return "global", self.glob

    def _pre_holiday_bucket(self, order_date: date) -> str | None:
        """order_date가 어떤 연휴 시작일 1~PRE_HOLIDAY_WINDOW일 전이면 그 연휴의 bucket(short/long) 반환, 아니면 None"""
        od = pd.Timestamp(order_date)
        for start, name, conf, bucket in self.holiday_starts:
            if start - pd.Timedelta(days=PRE_HOLIDAY_WINDOW) <= od < start:
                return bucket
        return None

    @staticmethod
    def _holiday_ratio_for_day(d: int, ratios: dict | None) -> float:
        """day1~30 사이 임의 날짜(headline의 매일 단위 스캔용)에 적용할 보정 비율을 찾는다.
        DAY_MARKS에 정확히 있으면 그 값을 쓰고, 사이값(day4,6,8,9 등)은 바로 아래 DAY_MARK의
        비율을 그대로 쓴다(day4 -> day3 비율). 검증 안 된 날은 holiday_adjustment_ratios.csv
        저장 시 이미 applied_ratio=1.0으로 박아뒀으므로(08_holiday_adjustment.py의
        ADJUSTABLE_DAYS_BY_BUCKET) 여기서 또 day<=5 같은 하드코딩 컷오프를 걸 필요가 없다 -
        버킷마다 안전 구간이 달라서(single은 day10까지, short/long은 day5까지) 예전의
        고정 컷오프는 오히려 single의 개선 여지를 놓치는 오류였다(2026-07-08 확인)."""
        if ratios is None:
            return 1.0
        if d in ratios:
            return ratios[d]
        lower = max((k for k in ratios if k <= d), default=None)
        return ratios.get(lower, 1.0) if lower is not None else 1.0

    def _headline(self, row, order_date: date, ratios: dict | None, threshold: float) -> dict:
        """'확률이 threshold를 처음 넘는 날'을 매일 단위(fd_1~fd_30)로 정확히 찾는다.
        DAY_MARKS(1,2,3,5,7,10,14,21,30)처럼 성긴 지점만 보면, 예: day14=89.9%->day21=98.6%인
        경우 실제 90% 도달일(예: 15~16일)을 건너뛰고 21일로 과대평가된다(2026-07-07 확인) -
        그래서 headline만은 매일 단위 fd_ 컬럼으로 따로 계산한다.

        일요일은 후보에서 제외한다(2026-07-07 확인): 실측 결과 전체 배송완료의 0.43%만
        일요일에 찍힘(월~금 10~21%, 토 9.8%와 비교하면 사실상 배송이 없는 날). fd_d는
        모든 주문요일을 섞어서 평균 낸 값이라, 특정 주문(예: 화요일 주문 -> 5일째=일요일)
        입장에선 "다른 요일에 주문했으면 평일에 해당했을 5일째"의 증가분까지 섞여 들어가
        실제보다 부풀려진다. 그래서 threshold를 넘는 날이 일요일이면 건너뛰고, 다음 날
        (반드시 월요일 - CDF가 단조증가라 threshold를 이미 넘은 상태라 무조건 넘음)의
        실제 fd 값을 그대로 채택한다 - 토요일(9.8%)은 정상적인 배송일이라 건너뛰지 않는다."""
        SUNDAY = 6
        prev_prob = 0.0
        for d in FINE_DAYS:
            prob = float(row[f"fd_{d}"]) * self._holiday_ratio_for_day(d, ratios)
            prob = max(prob, prev_prob)
            prev_prob = prob
            target = order_date + timedelta(days=d)
            if prob >= threshold and target.weekday() != SUNDAY:
                return {"day": d, "date": target.isoformat(), "weekday": WEEKDAY_KO[target.weekday()],
                        "prob": round(prob, 4), "threshold": threshold, "threshold_met": True}
        target = order_date + timedelta(days=FINE_DAYS[-1])
        return {"day": FINE_DAYS[-1], "date": target.isoformat(), "weekday": WEEKDAY_KO[target.weekday()],
                "prob": round(prev_prob, 4), "threshold": threshold, "threshold_met": False}

    def predict(self, order_date: date, product_code: str | None = None, vendor: str | None = None,
                대분류: str | None = None, 중분류: str | None = None,
                표준납기일: date | None = None, headline_threshold: float = 0.9) -> dict:
        """
        입력: 주문일자, (권장) 상품코드 — 협력사/대분류/중분류는 상품코드로 자동 조회되므로
        생략 가능. 상품코드가 product_lookup에 없거나 안 넘겼으면 vendor/대분류/중분류를
        직접 넘겨서 폴백시킬 수 있음(둘 다 없으면 전체 평균). headline_threshold: 아래
        headline 계산에 쓸 확률 기준(기본 90%).
        표준납기일(선택, 2026-07-13 추가): 통합플랫폼이 이미 보유·표시 중인 그 값이 있으면
        넘길 것 - 실제 값이 항상 추정치보다 정확함. 협력사 이력이 없는 콜드스타트 주문에서
        cat/global 대신 납기버킷 레벨(sla_bucket)로 답할 수 있게 해줌 - 같은 주문 검증에서
        콜드스타트 전 구간 오차 절반 이하 확인.
        안 넘기고 product_code만 주면(2026-07-27 변경): product_lookup.csv/
        vendor_sla_lookup.csv에 있는 상품코드/협력사 최빈값으로 자동 추정해서 채운다 -
        연동 전이거나 실제 값을 아직 못 받는 상황에서도 이 폴더 데이터만으로 바로 동작하게
        하기 위함. 상품코드도 없거나 추정치도 없으면 기존처럼 cat/global로 폴백(완전한
        하위 호환은 아님 - 과거엔 표준납기일 생략 시 무조건 cat/global이었으나 이제 가능하면
        추정치를 씀; 출력의 sla_estimated로 실제값/추정값 여부를 구분할 것).
        출력: {
          level: 어느 세그먼트 단계로 답했는지 (sku가 가장 세밀함),
          resolved_vendor/resolved_category: 상품코드로 조회된 협력사/카테고리(참고용),
          product_recognized: 상품코드를 줬는데 product_lookup.csv에도 없고 vendor/대분류/
            중분류도 안 준 경우 False (2026-07-07 추가). "신규라 자기 데이터가 없는 상품"과
            "애초에 존재하지 않는 코드(오타 등)"를 구분할 방법이 없어서, 이 경우엔 그냥
            global 평균을 자신 있게 보여주는 대신 이 플래그로 "확인이 필요할 수 있음"을
            알린다. 다만 이것도 완전한 판별은 아니다 - 상품코드가 실제 신상품이라
            아직 우리 주문 이력이 0건인 경우도 똑같이 False가 뜬다(구분할 근거 자체가 없음).
          sample_size: 그 세그먼트 학습에 쓰인 과거 주문 건수,
          sla_estimated: 위 표준납기일_사용값이 호출 측이 준 실제 값이 아니라 이 함수가
            자동 추정한 값인지 (True=추정치, False=실제 값을 받았거나 애초에 안 씀).
            화면에 노출할 경우 "추정치 - 실제 값과 다를 수 있음" 같은 문구를 권장.
          표준납기일_사용값: 이번 계산에 실제로 적용된 표준납기일(ISO 날짜, 없으면 None) -
            호출 측이 넘긴 값 또는 자동 추정값.
          holiday_adjusted: 공휴일 연휴 직전 보정이 적용됐는지,
          confidence_level: "high"/"medium"/"low" - 레벨 전체의 신뢰도 요약(LEVEL_CONFIDENCE).
            어느 레벨의 어느 day가 구체적으로 못 미더운지는 아래 by_day[].reliable /
            headline.reliable을 보고 판단할 것 - 레벨마다 불안정한 구간의 위치가 다르다
            (cat/global은 중간 구간만 나쁘고 vendor_mid는 반대로 21일 이후가 나쁨 -
            2026-07-07 검증, LEVEL_UNRELIABLE_DAYS 참고). API는 사실(어디가 못 미더운지)만
            제공하고, 이걸 화면에서 숨길지/흐리게 보일지/문구를 바꿀지는 UI 정책 - serve.py는
            판단하지 않는다.
          headline: {day, date, weekday, prob, threshold_met, reliable} - 확률이
            headline_threshold를 처음 넘는 날 (고정된 날짜의 확률이 아니라, 확신할 수 있는
            날짜를 찾은 것). threshold_met=False면 day90까지도 기준을 못 넘겼다는 뜻(장기
            지연 가능 상품 신호). reliable=False면 이 날짜의 확률이 해당 레벨의 검증된
            불안정 구간에 속한다는 뜻(그래도 계산은 그대로 반환 - 표시 여부는 UI가 결정).
          by_day: [{day, date, weekday, prob, reliable}, ...]  # day=주문일로부터 경과일,
            prob=누적 도착확률, reliable=이 (레벨, day) 조합이 홀드아웃 검증에서 오차
            10%p 이하였는지(45/60/90일 확장 지점은 검증 대상 밖이라 기본 True)
          shortage_alert / shortage_reason: shortage_flags.csv에 이 상품코드로 등록된
            쇼티지 기간(시작일~종료일)에 order_date가 걸리면 True + 사유. 확률/by_day는
            그대로 반환하고(과거 이력 기반 확률을 임의로 재계산하지 않음) 경고만 얹는다 -
            표시 방식은 UI/통합플랫폼 정책(미확정, 2026-07-09 예방적 설계).
        }
        """
        resolved = self.resolve_product(product_code) if product_code else None
        # 상품코드를 줬는데 product_lookup에도 없고 vendor/대분류/중분류도 안 준 경우 -
        # "신규 상품이라 자기 데이터가 없는 것"과 "애초에 존재하지 않는 코드(오타 등)"를
        # 구분할 방법이 없다(품목 마스터가 아니라 주문 이력 기반이라서). 이 경우까지
        # global 평균을 자신 있게 보여주면 오해를 준다(2026-07-07 사용자 지적) - 그래서
        # 이 상태를 product_recognized=False로 명시적으로 노출한다. UI는 이 필드가
        # False면 확률 대신 "상품코드 확인" 안내를 보여줄 것을 권장(강제하지 않음).
        product_recognized = not (product_code and resolved is None and vendor is None and 대분류 is None)
        if resolved:
            vendor = vendor or resolved["협력사명"]
            대분류 = 대분류 or resolved["대분류"]
            중분류 = 중분류 or resolved["중분류"]

        # 표준납기일 자동 추정(2026-07-27): 호출 측이 표준납기일을 안 넘겼으면(실제 값을 모르거나
        # 아직 연동 전이면) product_code로 추정치를 채워본다 - 실제 값이 있으면 그게 항상 우선이고
        # 이 블록은 건드리지 않는다(아래 조건이 표준납기일 is None일 때만 진입).
        sla_estimated = False
        if 표준납기일 is None and product_code:
            sla_lt = self._estimate_sla_lt(product_code)
            if sla_lt is not None:
                표준납기일 = order_date + timedelta(days=sla_lt)
                sla_estimated = True

        sla_bucket = self._sla_bucket_of(order_date, 표준납기일)
        level, row = self._lookup(product_code, vendor, 대분류, 중분류, sla_bucket)
        bucket = self._pre_holiday_bucket(order_date)
        ratios = self.holiday_ratios.get(bucket) if bucket else None
        unreliable_days = LEVEL_UNRELIABLE_DAYS.get(level, set())

        by_day = []
        prev_prob = 0.0
        for d in DAY_MARKS:
            target = order_date + timedelta(days=d)
            prob = float(row[f"p_{d}d"])
            if ratios is not None:
                prob = prob * ratios.get(d, 1.0)
            prob = max(prob, prev_prob)  # 보정 후에도 누적확률 단조증가 보장
            prev_prob = prob
            by_day.append({
                "day": d,
                "date": target.isoformat(),
                "weekday": WEEKDAY_KO[target.weekday()],
                "prob": round(prob, 4),
                "reliable": d not in unreliable_days,
            })

        headline = self._headline(row, order_date, ratios, headline_threshold)
        if headline["day"] > DAY_MARKS[-1]:
            # 배송이 느려서 headline이 30일을 넘긴 경우에만 45/60/90일 지점을 추가로 보여준다
            # (2026-07-07 - 표가 30일에서 끊겨 있으면 76.8%->91.0%처럼 갑자기 뛰는 것처럼
            # 보여서, 30일 이후 흐름도 볼 수 있도록 확장. fd_ 컬럼 사용 - 공휴일 보정은
            # day1~5에만 검증돼 있어 여기엔 적용하지 않음).
            for d in EXTRA_DAY_MARKS:
                if d <= FINE_DAYS[-1]:
                    target = order_date + timedelta(days=d)
                    prob = max(float(row[f"fd_{d}"]), prev_prob)
                    prev_prob = prob
                    by_day.append({
                        "day": d,
                        "date": target.isoformat(),
                        "weekday": WEEKDAY_KO[target.weekday()],
                        "prob": round(prob, 4),
                        "reliable": True,  # 45/60/90일은 레벨별 검증 대상 밖(30일까지만 검증) - 위 주석 참고
                    })
        headline["reliable"] = headline["day"] not in unreliable_days
        shortage = self._active_shortage(product_code, order_date)
        return {
            "level": level,
            "shortage_alert": shortage is not None,
            "shortage_reason": shortage["reason"] if shortage else None,
            "resolved_vendor": vendor,
            "resolved_category": f"{대분류 or ''} > {중분류 or ''}".strip(" >"),
            "product_name": resolved["상품명"] if resolved else None,
            "product_recognized": product_recognized,
            "sample_size": int(row["n"]),
            "median_days": float(row["median"]),
            "holiday_adjusted": bucket is not None,
            "confidence_level": LEVEL_CONFIDENCE.get(level, "high"),
            "sla_estimated": sla_estimated,
            "표준납기일_사용값": 표준납기일.isoformat() if 표준납기일 else None,
            "headline": headline,
            "by_day": by_day,
        }


def _demo():
    pred = DeliveryPredictor()

    # 물량 많은 히트상품 top 5 (자기 데이터로 예측되는 sku 레벨 사례)
    top_skus = pred.sku.sort_values("n", ascending=False).head(5).index.tolist()

    print("=" * 70)
    print("사례 1~5: 상품코드만 넘겼을 때 (히트상품 - sku 레벨로 예측되는 경우)")
    print("=" * 70)
    for code in top_skus:
        r = pred.predict(order_date=date(2026, 7, 2), product_code=code)
        h = r["headline"]
        print(f"\n상품코드={code} ({r['product_name']})  협력사={r['resolved_vendor']}  "
              f"->  level={r['level']}  n={r['sample_size']:,}  median={r['median_days']}일")
        print(f"  headline: {h['date']}({h['weekday']}) 도착확률 {h['prob']*100:.1f}% "
              f"(90% 기준 {'충족' if h['threshold_met'] else '미충족 - day90까지도 90% 못 넘김'})")
        for row in r["by_day"][:5]:
            print(f"  {row['day']:>2}일후 ({row['date']} {row['weekday']}요일): 도착확률 {row['prob']*100:.1f}%")

    print("\n" + "=" * 70)
    print("사례 5.5: 느린 상품 - 고정된 날짜(3일) 확률 대신 headline이 왜 더 의미있는지")
    print("=" * 70)
    slow_code = "10274648"  # UTP케이블 - 1일 2%, 3일 18.5%처럼 낮은 숫자만 보면 오해하기 쉬움
    r = pred.predict(order_date=date(2026, 7, 2), product_code=slow_code)
    h = r["headline"]
    print(f"상품코드={slow_code} ({r['product_name']})")
    print(f"  고정 3일 확률만 보면: {[d for d in r['by_day'] if d['day']==3][0]['prob']*100:.1f}% -> '이상한 상품인가?' 오해 소지")
    print(f"  headline(90% 넘는 첫 날): {h['date']}({h['weekday']}), {h['day']}일 후, 확률 {h['prob']*100:.1f}% -> 훨씬 명확한 정보")

    print("\n" + "=" * 70)
    print("사례 6: 물량이 적은 희귀 상품코드 (자동으로 상위 레벨 폴백)")
    print("=" * 70)
    rare_code = pred.product_lookup[pred.product_lookup["n"] < 30].index[0]
    r = pred.predict(order_date=date(2026, 7, 2), product_code=rare_code)
    print(f"상품코드={rare_code} ({r['product_name']})  협력사={r['resolved_vendor']}  "
          f"->  level={r['level']} (sku 아님 = 폴백됨)  n={r['sample_size']:,}")

    print("\n" + "=" * 70)
    print("사례 7: 상품코드 없이 협력사만으로 조회 (기존 방식, 여전히 지원됨)")
    print("=" * 70)
    r = pred.predict(order_date=date(2026, 7, 2), vendor="존재하지않는협력사")
    print(f"협력사=존재하지않는협력사  ->  level={r['level']}  n={r['sample_size']:,}")

    print("\n" + "=" * 70)
    print("사례 8: 공휴일 보정 예시 (설날 하루 전)")
    print("=" * 70)
    code = top_skus[0]
    r = pred.predict(order_date=date(2026, 2, 15), product_code=code)
    print(f"상품코드={code}  주문일=2026-02-15(설날 하루전)  공휴일보정={r['holiday_adjusted']}")
    for row in r["by_day"]:
        print(f"  {row['day']:>2}일후 ({row['date']} {row['weekday']}요일): 도착확률 {row['prob']*100:.1f}%")


if __name__ == "__main__":
    _demo()
