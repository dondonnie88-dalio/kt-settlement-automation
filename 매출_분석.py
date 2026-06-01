"""
매출.xlsx 자동 분석 스크립트 (KT 정산 데이터 형식)
- 시트별 데이터 읽기
- 사업장별 / 월별 정산금액 집계
- 전월 대비 증감율 계산
- 향후 3개월 매출 추정 (최근 3개월 평균)
- 결과 저장: 매출_분석결과.xlsx
"""

import os, sys
from pathlib import Path

REQUIRED = {"pandas": "pandas", "openpyxl": "openpyxl"}
missing = [p for p, m in REQUIRED.items() if not __import__(m, globals(), locals(), [], 0) if True]
try:
    for p, m in REQUIRED.items():
        __import__(m)
except ImportError as e:
    missing_pkg = str(e).split("'")[1]
    print(f"[설치 필요] pip install {missing_pkg}")
    sys.exit(1)

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from datetime import datetime, date
import random, calendar
import warnings
warnings.filterwarnings("ignore")

BASE_DIR    = Path(__file__).parent
INPUT_FILE  = BASE_DIR / "매출.xlsx"
OUTPUT_FILE = BASE_DIR / "매출_분석결과.xlsx"

# ════════════════════════════════════════════════════════════════════════════
# 0. 샘플 데이터 생성 (KT 정산 형식)
# ════════════════════════════════════════════════════════════════════════════
def create_sample():
    random.seed(42)

    사업장_목록 = [
        ("1000", "KT본사(분당)",    "1028142945"),
        ("1100", "KT서울사옥",      "1028142946"),
        ("1200", "KT광화문빌딩",    "1028142947"),
        ("1300", "KT대전센터",      "1028142948"),
        ("1400", "KT부산센터",      "1028142949"),
        ("1500", "KT대구센터",      "1028142950"),
    ]
    부서_목록 = [
        ("540445", "ICT융합사업부", "성과분석팀"),
        ("540446", "ICT융합사업부", "플랫폼기획팀"),
        ("540447", "네트워크부문",  "망운용팀"),
        ("540448", "네트워크부문",  "인프라팀"),
        ("540449", "기업사업부문",  "B2B영업팀"),
        ("540450", "기업사업부문",  "솔루션팀"),
    ]
    협력사_목록 = [
        ("17477", "오렌지스펙트럼(주)"),
        ("17478", "(주)케이티링커스"),
        ("17479", "엔씨소프트(주)"),
        ("17480", "(주)다우기술"),
        ("17481", "삼성SDS(주)"),
        ("17482", "LG CNS(주)"),
    ]
    상품_목록 = [
        ("10532552", "649", "청소/생활/의료용품 > 생활용품 > 휴대용 통신장치용품 > 휴대전화 케이블",
         "전자제품 > 가전/전자제품 > 통신기기 > 핸드폰액세서리",
         "가전/전자제품", "앤커 플로우 USB C to C 100W PD 고속충전 케이블", "C to C, 90cm", 14000),
        ("10532553", "650", "사무용품 > 전산/사무기기 > 복합기/프린터 > 토너/잉크",
         "사무용품 > 복합기/프린터소모품",
         "사무용품", "삼성 정품 토너 MLT-D111S", "흑백, 1000매", 52000),
        ("10532554", "651", "IT서비스 > 클라우드 > SaaS > 협업툴",
         "IT서비스 > 협업솔루션",
         "IT서비스", "KT 기업용 협업툴 라이선스(월)", "사용자당/월", 8800),
        ("10532555", "652", "사무용품 > 컴퓨터 > 노트북 > 비즈니스용",
         "전자제품 > PC/노트북",
         "전자제품", "LG그램 17인치 비즈니스(i7)", "17Z90Q, Win11 Pro", 1890000),
        ("10532556", "653", "통신서비스 > 인터넷 > 전용회선",
         "통신서비스 > 네트워크",
         "통신서비스", "기가인터넷 전용회선 100M", "월정액, 1년약정", 220000),
    ]
    주문형태_목록 = ["IP", "DIP"]
    매출구분_목록 = ["소액_자율소싱(일반)", "대액_지정구매", "자체상품", "서비스"]
    담당자_목록 = ["최한나", "김민준", "이서연", "박지훈", "정다은", "오승우"]

    # 거래처 규모 매핑 (어음 발행 판별에 사용)
    거래처규모_map = {
        "오렌지스펙트럼(주)": "중소기업",
        "(주)케이티링커스":   "중견기업",
        "엔씨소프트(주)":    "대기업",
        "(주)다우기술":      "중소기업",
        "삼성SDS(주)":      "대기업",
        "LG CNS(주)":       "대기업",
    }

    wb = openpyxl.Workbook()
    columns = [
        "일정산번호", "정산확정일", "일일정산월", "주문번호", "품목번호", "주문&품목",
        "인보이스(KT이동유형명 값)", "사업장코드", "사업장", "사업자번호",
        "상위부서코드", "상위부서명", "부서명", "부서코드",
        "상품ID", "카테고리ID", "마스터카테고리(CMS)", "서비스카테고리", "중분류(IP만변경적용)",
        "담당자", "일반/통신구분", "매출구분", "상품코드", "상품명", "대표규격", "상태",
        "선정산여부", "단위", "주문수량", "주문단가", "주문금액",
        "정산수량", "정산단가", "정산금액", "매입단가", "매입금액", "매출총이익",
        "계정ID", "계정명", "주문자", "주문자ID", "주문형태",
        "코스트센터코드", "코스트센터명", "WBS코드", "WBS명", "세금코드",
        "SA_ID", "입고자ID", "입고자명", "협력사코드", "협력사명",
        "상품타입", "매출과세구분", "매입과세구분", "주문유형",
        "주문일시", "발주일", "출하지시일", "입고일", "승인일", "배송완료일",
        "일일정산일", "정산시점", "계약번호", "공사명", "발주처",
        "공사시작일", "공사종료일", "법인", "그룹", "사이트", "VAT포함여부",
        "담당자부서", "배송형태", "VMI배송구분", "픽업배송구분", "IP/DIP",
        "주문자이메일", "고객사상품코드", "상품관리자", "배송지", "배송지상세주소",
        "송장번호", "배송메모", "공정명", "지급상태", "발행요청상태", "발행요청일",
        "결제유형", "거래처규모"   # 거래처규모 추가
    ]

    def make_row(seq, 정산확정일_str, settle_date, order_date,
                 sp_code, sp_name, sp_biz,
                 dept_code, dept_upper, dept_name,
                 prod, 담당자, coop_code, coop_name,
                 qty, unit_price, 세금구분, introduce_error=False):
        공급가액 = qty * unit_price
        purchase_price = int(unit_price * random.uniform(0.88, 0.96))
        purchase_amt  = qty * purchase_price
        gross_profit  = 공급가액 - purchase_amt

        if 세금구분 == "과세":
            정산금액 = round(공급가액 * 1.1)
            주문금액 = 정산금액
            if introduce_error:           # 의도적 오차 (+3 ~ +5 원)
                정산금액 += random.choice([3, 4, 5])
        else:
            정산금액 = 공급가액
            주문금액 = 공급가액

        return [
            seq, 정산확정일_str, settle_date,
            random.randint(4500000000, 4599999999), random.randint(1, 5), "",
            "입고", sp_code, sp_name, sp_biz,
            "", dept_upper, dept_name, dept_code,
            prod[0], prod[1], prod[2], prod[3], prod[4],
            담당자, "일반", random.choice(매출구분_목록),
            prod[0], prod[5], prod[6], "일마감", "", "EA",
            qty, unit_price, 주문금액,      # 주문수량/단가/금액
            qty, unit_price, 정산금액,      # 정산수량/단가/금액
            purchase_price, purchase_amt, gross_profit,
            "", "부서운영비", "홍길동",
            f"10{random.randint(100000,999999)}",
            random.choice(주문형태_목록),
            dept_code, dept_name, "", "", "", "", "", "홍길동",
            coop_code, coop_name,
            "일반상품", 세금구분, 세금구분, "일반주문",
            order_date, order_date,
            settle_date, settle_date, settle_date, settle_date, settle_date,
            "검수(입고)", "", "", "", "", "",
            "KT", "KT", "KTIP", "별도", "통합SCM사업팀", "협력사배송", "", "",
            random.choice(주문형태_목록),
            f"user{random.randint(100,999)}@kt.com",
            f"K{random.randint(6000000,6999999)}",
            담당자,
            f"서울특별시 {random.choice(['종로구','강남구','송파구','마포구'])}",
            "", "", "", "", "", "", "후결제",
            거래처규모_map.get(coop_name, "중소기업"),   # 거래처규모
        ]

    seq_no = 69451477
    # 세금구분 비율: 과세 70%, 면세 15%, 비과세 15%
    세금구분_가중치 = ["과세"] * 14 + ["면세"] * 3 + ["비과세"] * 3
    error_counter = 0   # 과세 행 중 약 8%에 오차 삽입

    for year in [2024, 2025]:
        ws = wb.create_sheet(f"{year}년")
        ws.append(columns)

        months = range(1, 13) if year == 2024 else range(1, 4)
        for month in months:
            days_in_m = calendar.monthrange(year, month)[1]
            for sp_code, sp_name, sp_biz in 사업장_목록:
                dept_code, dept_upper, dept_name = random.choice(부서_목록)
                coop_code, coop_name = random.choice(협력사_목록)
                prod    = random.choice(상품_목록)
                담당자   = random.choice(담당자_목록)
                n_orders = random.randint(1, 5)

                for _ in range(n_orders):
                    day         = random.randint(1, days_in_m)
                    order_date  = date(year, month, day)
                    settle_date = date(year, month, min(day + random.randint(3, 10), days_in_m))
                    차수         = "1차" if day <= 15 else "2차"
                    정산확정일_str = f"{month}월 {차수}"

                    qty      = random.randint(1, 30)
                    세금구분  = random.choice(세금구분_가중치)

                    # 과세 행 중 약 8%에 의도적 오차 삽입
                    introduce_error = (세금구분 == "과세" and
                                       (error_counter % 13 == 0))
                    if 세금구분 == "과세":
                        error_counter += 1

                    row = make_row(
                        seq_no, 정산확정일_str, settle_date, order_date,
                        sp_code, sp_name, sp_biz,
                        dept_code, dept_upper, dept_name,
                        prod, 담당자, coop_code, coop_name,
                        qty, prod[7], 세금구분, introduce_error
                    )
                    ws.append(row)
                    seq_no += 1

        # ── 대규모 비과세 주문 (어음 발행 테스트용) ───────────────────────
        # 주문1: 삼성SDS(주)/대기업, 비과세 합계 240M → 어음 발행 대상
        for item_qty, unit_p, 확정일, d in [
            (1, 80_000_000, "3월 1차", date(year, 3 if year==2025 else 12, 5)),
            (1, 80_000_000, "3월 1차", date(year, 3 if year==2025 else 12, 7)),
            (1, 80_000_000, "3월 1차", date(year, 3 if year==2025 else 12, 9)),
        ]:
            ws.append(make_row(
                seq_no, 확정일, d, d,
                "1000", "KT본사(분당)", "1028142945",
                "540449", "기업사업부문", "B2B영업팀",
                ("10599001","660","IT서비스>SI개발","IT서비스>SI","IT서비스",
                 "SI 개발 용역(비과세)","일식",unit_p),
                "박지훈", "17481", "삼성SDS(주)",
                item_qty, unit_p, "비과세", False
            ))
            seq_no += 1

        # 주문2: LG CNS(주)/대기업, 비과세 합계 240M, 오차 있음 → 오차 확인 필요
        for idx, (item_qty, unit_p, has_err) in enumerate([
            (1, 120_000_000, False),
            (1, 120_000_000, True),   # 두 번째 품목에 오차 삽입
        ]):
            ws.append(make_row(
                seq_no, "3월 2차",
                date(year, 3 if year==2025 else 12, 20),
                date(year, 3 if year==2025 else 12, 18),
                "1100", "KT서울사옥", "1028142946",
                "540447", "네트워크부문", "망운용팀",
                ("10599002","661","IT서비스>컨설팅","IT서비스>컨설팅","IT서비스",
                 "네트워크 컨설팅 용역(비과세)","일식",unit_p),
                "김민준", "17482", "LG CNS(주)",
                item_qty, unit_p, "비과세", has_err
            ))
            seq_no += 1

        # 주문3: (주)케이티링커스/중견기업, 비과세 합계 160M → 금액 미달 제외
        for item_qty, unit_p in [(1, 80_000_000), (1, 80_000_000)]:
            ws.append(make_row(
                seq_no, "2월 2차",
                date(year, 2 if year==2024 else 2, 25),
                date(year, 2 if year==2024 else 2, 23),
                "1200", "KT광화문빌딩", "1028142947",
                "540446", "ICT융합사업부", "플랫폼기획팀",
                ("10599003","662","IT서비스>유지보수","IT서비스>유지보수","IT서비스",
                 "시스템 유지보수 용역(비과세)","월",unit_p),
                "이서연", "17478", "(주)케이티링커스",
                item_qty, unit_p, "비과세", False
            ))
            seq_no += 1

    del wb["Sheet"]
    wb.save(INPUT_FILE)
    print(f"[샘플 생성] {INPUT_FILE}")


# ════════════════════════════════════════════════════════════════════════════
# 1. 데이터 읽기 & 전처리
# ════════════════════════════════════════════════════════════════════════════
def load_data() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        create_sample()

    all_sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
    print(f"\n[읽기] 시트 {len(all_sheets)}개: {list(all_sheets.keys())}")

    frames = []
    for sheet_name, df in all_sheets.items():
        df.columns = df.columns.str.strip()

        # ── 날짜 컬럼 우선순위: 일일정산월 > 주문일시 > 일일정산일 ─────────
        date_candidates = ["일일정산월", "주문일시", "일일정산일", "정산확정일",
                           "날짜", "일자", "date", "Date", "월", "기간"]
        date_col = next((c for c in date_candidates if c in df.columns), None)

        # ── 거래처 컬럼: 사업장 우선 ──────────────────────────────────────
        client_candidates = ["사업장", "거래처", "고객", "업체", "client", "Client"]
        client_col = next((c for c in client_candidates if c in df.columns), None)

        # ── 매출 컬럼: 정산금액 우선 ─────────────────────────────────────
        revenue_candidates = ["정산금액", "주문금액", "매출액", "매출", "금액",
                               "revenue", "Revenue"]
        revenue_col = next((c for c in revenue_candidates if c in df.columns), None)

        if not all([date_col, client_col, revenue_col]):
            print(f"  [경고] '{sheet_name}' - 필수 컬럼 미발견, 스킵 (발견: 날짜={date_col}, "
                  f"거래처={client_col}, 매출={revenue_col})")
            continue

        # ── 추가 컬럼 선택적 포함 ─────────────────────────────────────────
        extra_cols = {}
        for alias, candidates in {
            "부서명": ["부서명", "부서", "department"],
            "협력사명": ["협력사명", "협력사", "vendor"],
            "상품명": ["상품명", "품목명", "product"],
            "매출총이익": ["매출총이익", "이익", "gross_profit"],
            "정산확정일": ["정산확정일"],
            "매출구분": ["매출구분"],
            "서비스카테고리": ["서비스카테고리"],
        }.items():
            found = next((c for c in candidates if c in df.columns), None)
            if found:
                extra_cols[alias] = found

        select_cols = [date_col, client_col, revenue_col] + list(extra_cols.values())
        rename_map = {date_col: "날짜", client_col: "거래처", revenue_col: "매출액"}
        rename_map.update({v: k for k, v in extra_cols.items()})

        sub = df[select_cols].rename(columns=rename_map).copy()

        # ── 날짜 파싱 ─────────────────────────────────────────────────────
        sub["날짜"] = pd.to_datetime(sub["날짜"], errors="coerce")
        # 정산확정일("1월 1차") 형식 fallback
        if sub["날짜"].isna().all() and "정산확정일" in sub.columns:
            def parse_정산월(s):
                try:
                    m = int(str(s).replace("월", "").split()[0])
                    return pd.Timestamp(f"2024-{m:02d}-01")
                except Exception:
                    return pd.NaT
            sub["날짜"] = sub["날짜_orig"] if "날짜_orig" in sub else sub["날짜"].fillna(
                sub["정산확정일"].apply(parse_정산월) if "정산확정일" in sub.columns else pd.NaT
            )

        sub["매출액"] = pd.to_numeric(sub["매출액"], errors="coerce")
        if "매출총이익" in sub.columns:
            sub["매출총이익"] = pd.to_numeric(sub["매출총이익"], errors="coerce")

        sub.dropna(subset=["날짜", "거래처", "매출액"], inplace=True)
        sub["시트"] = sheet_name
        frames.append(sub)
        print(f"  [OK] '{sheet_name}' → {len(sub):,}행  (날짜={date_col}, "
              f"거래처={client_col}, 매출={revenue_col})")

    if not frames:
        sys.exit("[오류] 분석 가능한 데이터가 없습니다.")

    data = pd.concat(frames, ignore_index=True)
    data["년월"] = data["날짜"].dt.to_period("M")
    data["년"] = data["날짜"].dt.year
    data["월"] = data["날짜"].dt.month
    print(f"\n[전체] {len(data):,}행, 사업장 {data['거래처'].nunique()}개, "
          f"기간 {data['년월'].min()} ~ {data['년월'].max()}")
    return data


# ════════════════════════════════════════════════════════════════════════════
# 2. 집계 및 분석
# ════════════════════════════════════════════════════════════════════════════
def analyze(data: pd.DataFrame):
    # ── 월별 전체 ──────────────────────────────────────────────────────────
    monthly = (
        data.groupby("년월", sort=True)
        .agg(정산금액=("매출액", "sum"),
             건수=("매출액", "count"))
        .reset_index()
    )
    monthly["년월_str"] = monthly["년월"].astype(str)
    monthly["전월금액"] = monthly["정산금액"].shift(1)
    monthly["증감액"] = monthly["정산금액"] - monthly["전월금액"]
    monthly["증감율"] = (monthly["증감액"] / monthly["전월금액"] * 100).round(2)

    # 매출총이익 포함 여부
    has_profit = "매출총이익" in data.columns
    if has_profit:
        profit_monthly = (
            data.groupby("년월", sort=True)["매출총이익"].sum().reset_index()
        )
        monthly = monthly.merge(profit_monthly, on="년월", how="left")
        monthly["이익률"] = (monthly["매출총이익"] / monthly["정산금액"] * 100).round(2)

    # ── 사업장별 월별 피벗 ─────────────────────────────────────────────────
    client_monthly = (
        data.groupby(["거래처", "년월"], sort=True)["매출액"].sum().reset_index()
    )
    client_monthly["년월_str"] = client_monthly["년월"].astype(str)
    pivot = client_monthly.pivot(
        index="거래처", columns="년월_str", values="매출액"
    ).fillna(0)
    pivot["합계"] = pivot.sum(axis=1)
    pivot["비중(%)"] = (pivot["합계"] / pivot["합계"].sum() * 100).round(2)
    pivot.sort_values("합계", ascending=False, inplace=True)
    month_cols = [c for c in pivot.columns if c not in ("합계", "비중(%)")]

    # ── 부서별 집계 ────────────────────────────────────────────────────────
    dept_pivot = None
    if "부서명" in data.columns:
        dept_monthly = (
            data.groupby(["부서명", "년월"], sort=True)["매출액"].sum().reset_index()
        )
        dept_monthly["년월_str"] = dept_monthly["년월"].astype(str)
        dept_pivot = dept_monthly.pivot(
            index="부서명", columns="년월_str", values="매출액"
        ).fillna(0)
        dept_pivot["합계"] = dept_pivot.sum(axis=1)
        dept_pivot.sort_values("합계", ascending=False, inplace=True)

    # ── 협력사별 집계 ──────────────────────────────────────────────────────
    coop_pivot = None
    if "협력사명" in data.columns:
        coop_monthly = (
            data.groupby(["협력사명", "년월"], sort=True)["매출액"].sum().reset_index()
        )
        coop_monthly["년월_str"] = coop_monthly["년월"].astype(str)
        coop_pivot = coop_monthly.pivot(
            index="협력사명", columns="년월_str", values="매출액"
        ).fillna(0)
        coop_pivot["합계"] = coop_pivot.sum(axis=1)
        coop_pivot.sort_values("합계", ascending=False, inplace=True)

    # ── 카테고리별 집계 ────────────────────────────────────────────────────
    cat_summary = None
    cat_col = next((c for c in ["서비스카테고리", "마스터카테고리(CMS)", "매출구분"] if c in data.columns), None)
    if cat_col:
        cat_summary = (
            data.groupby(cat_col)["매출액"]
            .agg(["sum", "count"])
            .rename(columns={"sum": "정산금액", "count": "건수"})
            .sort_values("정산금액", ascending=False)
        )

    # ── 향후 3개월 추정 ────────────────────────────────────────────────────
    last3 = month_cols[-3:] if len(month_cols) >= 3 else month_cols
    last_period = data["년월"].max()
    forecast_periods = pd.period_range(last_period + 1, periods=3, freq="M")
    forecast_labels = [str(p) for p in forecast_periods]

    forecast_df = pd.DataFrame(index=pivot.index)
    for label in forecast_labels:
        forecast_df[label] = pivot[last3].mean(axis=1).round(0)

    return monthly, pivot, month_cols, dept_pivot, coop_pivot, cat_summary, forecast_df, forecast_labels, has_profit


# ════════════════════════════════════════════════════════════════════════════
# 3. 스타일 헬퍼
# ════════════════════════════════════════════════════════════════════════════
FONT      = "맑은 고딕"
C_H_BG    = "1F4E79"; C_H_FG = "FFFFFF"
C_TITLE   = "2E75B6"
C_ALT     = "DEEAF1"
C_FCST    = "E2EFDA"; C_TOTAL = "FFF2CC"
C_POS     = "C6EFCE"; C_NEG   = "FFCCCC"
C_DEPT    = "4472C4"; C_COOP  = "70AD47"

def _side(): return Side(style="thin", color="BFBFBF")
def tborder(): return Border(left=_side(), right=_side(), top=_side(), bottom=_side())

def hcell(cell, text, bg=C_H_BG, fg=C_H_FG, bold=True, wrap=False):
    cell.value = text
    cell.font  = Font(name=FONT, bold=bold, color=fg, size=10)
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = tborder()

def dcell(cell, value, fmt=None, bold=False, bg=None, align="right"):
    cell.value = value
    cell.font  = Font(name=FONT, bold=bold, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = tborder()
    if fmt: cell.number_format = fmt
    if bg:  cell.fill = PatternFill("solid", start_color=bg)

def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

def sheet_title(ws, title, sub=""):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=25)
    c = ws.cell(1, 1, title)
    c.font  = Font(name=FONT, bold=True, size=14, color="FFFFFF")
    c.fill  = PatternFill("solid", start_color=C_TITLE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=25)
        c2 = ws.cell(2, 1, sub)
        c2.font  = Font(name=FONT, size=10, color="595959")
        c2.fill  = PatternFill("solid", start_color="D6E4F0")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 18

NUM = "#,##0"; PCT = "0.0%;[Red]-0.0%;\"−\""
NUM_DELTA = "#,##0;[Red]-#,##0;\"-\""


# ════════════════════════════════════════════════════════════════════════════
# 4. 시트 작성
# ════════════════════════════════════════════════════════════════════════════

def write_monthly(wb, monthly, has_profit):
    ws = wb.create_sheet("📅 월별 정산 요약")
    sheet_title(ws, "월별 정산금액 요약",
                f"기준일: {datetime.today().strftime('%Y-%m-%d')}")

    headers = ["연월", "정산건수", "정산금액 (원)", "전월 정산금액", "증감액 (원)", "증감율 (%)"]
    if has_profit:
        headers += ["매출총이익 (원)", "이익률 (%)"]
    SR = 4
    for c, h in enumerate(headers, 1):
        hcell(ws.cell(SR, c), h)

    for r, row in enumerate(monthly.itertuples(), SR + 1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), row.년월_str, align="center", bg=bg)
        dcell(ws.cell(r, 2), row.건수, NUM, bg=bg)
        dcell(ws.cell(r, 3), row.정산금액, NUM, bg=bg)
        prev = row.전월금액 if pd.notna(row.전월금액) else None
        dcell(ws.cell(r, 4), prev, NUM, bg=bg)
        delta = row.증감액 if pd.notna(row.증감액) else None
        dbg = (C_POS if (delta or 0) >= 0 else C_NEG) if delta is not None else bg
        dcell(ws.cell(r, 5), delta, NUM_DELTA, bg=dbg)
        pct = row.증감율 if pd.notna(row.증감율) else None
        dcell(ws.cell(r, 6), (pct / 100) if pct else None, PCT, bg=dbg)
        if has_profit:
            dcell(ws.cell(r, 7), row.매출총이익 if pd.notna(row.매출총이익) else None, NUM, bg=bg)
            ir = row.이익률 if pd.notna(row.이익률) else None
            dcell(ws.cell(r, 8), (ir / 100) if ir else None, "0.0%", bg=bg)

    last = SR + len(monthly)
    tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for ci, col in enumerate(["B", "C", "G"] if has_profit else ["B", "C"], 2):
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = NUM
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, ci).border = tborder()

    for c, w in enumerate([14,10,20,20,18,12,18,12], 1): cw(ws, c, w)
    ws.freeze_panes = f"A{SR+1}"

    chart = LineChart()
    chart.title = "월별 정산금액 추이"
    chart.style = 10; chart.width = 24; chart.height = 13
    n = len(monthly)
    chart.add_data(Reference(ws, min_col=3, min_row=SR, max_row=SR+n), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=SR+1, max_row=SR+n))
    ws.add_chart(chart, f"I{SR}")


def write_client_pivot(wb, pivot, month_cols):
    ws = wb.create_sheet("🏢 사업장별 집계")
    sheet_title(ws, "사업장별 월별 정산금액 집계")
    SR = 4
    all_cols = month_cols + ["합계", "비중(%)"]
    for c, h in enumerate(["사업장"] + all_cols, 1):
        hcell(ws.cell(SR, c), h, wrap=True)
    ws.row_dimensions[SR].height = 30

    for r, (client, row) in enumerate(pivot.iterrows(), SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), client, align="left", bg=bg)
        for c, col in enumerate(all_cols, 2):
            val = row[col]
            dcell(ws.cell(r, c), val/100 if col=="비중(%)" else val,
                  "0.00%" if col=="비중(%)" else NUM, bg=bg)

    last = SR + len(pivot); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for c in range(2, len(all_cols)+2):
        cl = get_column_letter(c)
        col_name = all_cols[c-2]
        ws.cell(tot, c).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, c).number_format = "0.00%" if col_name=="비중(%)" else NUM
        ws.cell(tot, c).font = Font(name=FONT, bold=True)
        ws.cell(tot, c).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, c).border = tborder()

    cw(ws, 1, 20)
    for c in range(2, len(all_cols)+2): cw(ws, c, 14)
    ws.freeze_panes = f"B{SR+1}"


def write_growth(wb, pivot, month_cols):
    ws = wb.create_sheet("📈 전월 대비 증감율")
    sheet_title(ws, "사업장별 전월 대비 매출 증감율")
    if len(month_cols) < 2:
        ws.cell(4, 1, "비교 가능한 월이 2개월 미만입니다.")
        return
    SR = 4
    compare = month_cols[1:]
    for c, h in enumerate(["사업장"] + compare, 1):
        hcell(ws.cell(SR, c), h, wrap=True)
    ws.row_dimensions[SR].height = 30

    for r, (client, row) in enumerate(pivot.iterrows(), SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), client, align="left", bg=bg)
        for ci, col in enumerate(compare, 2):
            prev = month_cols[month_cols.index(col)-1]
            curr_v, prev_v = row[col], row[prev]
            if prev_v != 0:
                rate = (curr_v - prev_v) / abs(prev_v)
                dcell(ws.cell(r, ci), rate, "0.0%;[Red]-0.0%;\"−\"",
                      bg=C_POS if rate >= 0 else C_NEG)
            else:
                dcell(ws.cell(r, ci), None, bg=bg)

    cw(ws, 1, 20)
    for c in range(2, len(compare)+2): cw(ws, c, 14)
    ws.freeze_panes = f"B{SR+1}"


def write_dept(wb, dept_pivot, month_cols):
    if dept_pivot is None:
        return
    ws = wb.create_sheet("🏛 부서별 집계")
    sheet_title(ws, "부서별 월별 정산금액 집계")
    SR = 4
    mc = [c for c in dept_pivot.columns if c != "합계"]
    all_cols = mc + ["합계"]
    for c, h in enumerate(["부서명"] + all_cols, 1):
        hcell(ws.cell(SR, c), h, bg=C_DEPT, wrap=True)
    ws.row_dimensions[SR].height = 30

    for r, (dept, row) in enumerate(dept_pivot.iterrows(), SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), dept, align="left", bg=bg)
        for c, col in enumerate(all_cols, 2):
            dcell(ws.cell(r, c), row[col], NUM, bold=(col=="합계"), bg=bg)

    last = SR + len(dept_pivot); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for c in range(2, len(all_cols)+2):
        cl = get_column_letter(c)
        ws.cell(tot, c).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, c).number_format = NUM
        ws.cell(tot, c).font = Font(name=FONT, bold=True)
        ws.cell(tot, c).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, c).border = tborder()

    cw(ws, 1, 20)
    for c in range(2, len(all_cols)+2): cw(ws, c, 14)
    ws.freeze_panes = f"B{SR+1}"


def write_coop(wb, coop_pivot):
    if coop_pivot is None:
        return
    ws = wb.create_sheet("🤝 협력사별 집계")
    sheet_title(ws, "협력사별 월별 정산금액 집계")
    SR = 4
    mc = [c for c in coop_pivot.columns if c != "합계"]
    all_cols = mc + ["합계"]
    for c, h in enumerate(["협력사명"] + all_cols, 1):
        hcell(ws.cell(SR, c), h, bg=C_COOP, wrap=True)
    ws.row_dimensions[SR].height = 30

    for r, (coop, row) in enumerate(coop_pivot.iterrows(), SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), coop, align="left", bg=bg)
        for c, col in enumerate(all_cols, 2):
            dcell(ws.cell(r, c), row[col], NUM, bold=(col=="합계"), bg=bg)

    last = SR + len(coop_pivot); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for c in range(2, len(all_cols)+2):
        cl = get_column_letter(c)
        ws.cell(tot, c).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, c).number_format = NUM
        ws.cell(tot, c).font = Font(name=FONT, bold=True)
        ws.cell(tot, c).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, c).border = tborder()

    cw(ws, 1, 22)
    for c in range(2, len(all_cols)+2): cw(ws, c, 14)
    ws.freeze_panes = f"B{SR+1}"


def write_category(wb, cat_summary):
    if cat_summary is None:
        return
    ws = wb.create_sheet("📦 카테고리별 집계")
    sheet_title(ws, "카테고리별 정산금액 집계")
    SR = 4
    for c, h in enumerate(["카테고리", "정산금액 (원)", "정산건수", "비중 (%)"], 1):
        hcell(ws.cell(SR, c), h)

    total = cat_summary["정산금액"].sum()
    for r, (cat, row) in enumerate(cat_summary.iterrows(), SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), cat, align="left", bg=bg)
        dcell(ws.cell(r, 2), row["정산금액"], NUM, bg=bg)
        dcell(ws.cell(r, 3), row["건수"], NUM, bg=bg)
        dcell(ws.cell(r, 4), row["정산금액"]/total, "0.00%", bg=bg)

    last = SR + len(cat_summary); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for ci, fmt in [(2, NUM), (3, NUM)]:
        cl = get_column_letter(ci)
        ws.cell(tot, ci).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, ci).number_format = fmt
        ws.cell(tot, ci).font = Font(name=FONT, bold=True)
        ws.cell(tot, ci).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, ci).border = tborder()
    dcell(ws.cell(tot, 4), 1.0, "0.00%", bold=True, bg=C_TOTAL)

    for c, w in [(1,40),(2,20),(3,12),(4,12)]: cw(ws, c, w)

    chart = BarChart()
    chart.type = "col"; chart.title = "카테고리별 정산금액"
    chart.style = 10; chart.width = 24; chart.height = 14
    n = len(cat_summary)
    chart.add_data(Reference(ws, min_col=2, min_row=SR, max_row=SR+n), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=SR+1, max_row=SR+n))
    ws.add_chart(chart, f"F{SR}")


def write_forecast(wb, pivot, forecast_df, month_cols, forecast_labels):
    ws = wb.create_sheet("🔮 향후 3개월 추정")
    sheet_title(ws, "향후 3개월 매출 추정",
                "※ 최근 3개월 단순 평균 기반 추정값 (참고용)")
    SR = 4
    last3 = month_cols[-3:] if len(month_cols) >= 3 else month_cols
    headers = (["사업장"] + last3 + ["│"] +
               [f"{l}\n(추정)" for l in forecast_labels] + ["추정 3개월 합계"])
    for c, h in enumerate(headers, 1):
        if h == "│":
            hcell(ws.cell(SR, c), "", bg="7F7F7F")
        elif "(추정)" in h:
            hcell(ws.cell(SR, c), h, bg="375623", wrap=True)
        else:
            hcell(ws.cell(SR, c), h, wrap=True)
    ws.row_dimensions[SR].height = 32

    sep_col = len(last3) + 2
    fc_start = sep_col + 1
    sum_col  = fc_start + len(forecast_labels)

    for r, client in enumerate(pivot.index, SR+1):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), client, align="left", bg=bg)
        for ci, col in enumerate(last3, 2):
            dcell(ws.cell(r, ci), pivot.loc[client, col], NUM, bg=bg)
        ws.cell(r, sep_col).border = tborder()
        for ci, label in enumerate(forecast_labels, fc_start):
            dcell(ws.cell(r, ci), forecast_df.loc[client, label], NUM, bg=C_FCST)
        fc_s = get_column_letter(fc_start)
        fc_e = get_column_letter(fc_start + len(forecast_labels) - 1)
        ws.cell(r, sum_col).value = f"=SUM({fc_s}{r}:{fc_e}{r})"
        ws.cell(r, sum_col).number_format = NUM
        ws.cell(r, sum_col).font = Font(name=FONT, bold=True)
        ws.cell(r, sum_col).fill = PatternFill("solid", start_color="D5E8D4")
        ws.cell(r, sum_col).border = tborder()
        ws.cell(r, sum_col).alignment = Alignment(horizontal="right", vertical="center")

    last = SR + len(pivot); tot = last + 1
    dcell(ws.cell(tot, 1), "합  계", bold=True, bg=C_TOTAL, align="center")
    for c in range(2, sum_col+1):
        if c == sep_col: continue
        cl = get_column_letter(c)
        ws.cell(tot, c).value = f"=SUM({cl}{SR+1}:{cl}{last})"
        ws.cell(tot, c).number_format = NUM
        ws.cell(tot, c).font = Font(name=FONT, bold=True)
        ws.cell(tot, c).fill = PatternFill("solid", start_color=C_TOTAL)
        ws.cell(tot, c).border = tborder()
        ws.cell(tot, c).alignment = Alignment(horizontal="right", vertical="center")

    # 범례
    lr = tot + 2
    ws.cell(lr, 1, "※ 연두색 셀 = 추정값 (최근 3개월 평균 적용)").font = Font(name=FONT, size=9, italic=True, color="595959")

    cw(ws, 1, 20)
    for c in range(2, sum_col+1): cw(ws, c, 16)
    ws.freeze_panes = f"B{SR+1}"


def write_dashboard(wb, monthly, pivot, coop_pivot, dept_pivot, forecast_df,
                    forecast_labels, has_profit):
    ws = wb.create_sheet("📊 대시보드", 0)
    sheet_title(ws, "KT 정산 매출 분석 대시보드",
                f"작성일: {datetime.today().strftime('%Y년 %m월 %d일')}")

    # KPI 카드
    total_amt  = monthly["정산금액"].sum()
    avg_amt    = monthly["정산금액"].mean()
    total_cnt  = monthly["건수"].sum()
    last_gr    = monthly["증감율"].dropna().iloc[-1] if monthly["증감율"].dropna().shape[0] else 0

    kpis = [
        ("누적 정산금액",  total_amt, NUM,   "1F4E79"),
        ("월 평균 정산금액", avg_amt,  NUM,   "2E75B6"),
        ("총 정산건수",    total_cnt, "#,##0", "375623"),
        ("최근월 증감율",  last_gr/100 if last_gr else 0, "0.0%", "843C0C"),
    ]
    for i, (label, val, fmt, bg) in enumerate(kpis):
        col = i * 3 + 1
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col+1)
        lc = ws.cell(4, col, label)
        lc.font  = Font(name=FONT, bold=True, color="FFFFFF", size=9)
        lc.fill  = PatternFill("solid", start_color=bg)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        lc.border = tborder()
        ws.cell(4, col+1).fill = PatternFill("solid", start_color=bg)
        ws.cell(4, col+1).border = tborder()
        vc = ws.cell(5, col, val)
        vc.number_format = fmt
        vc.font  = Font(name=FONT, bold=True, color=bg, size=14)
        vc.fill  = PatternFill("solid", start_color="F2F2F2")
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = tborder()
        ws.cell(5, col+1).fill = PatternFill("solid", start_color="F2F2F2")
        ws.cell(5, col+1).border = tborder()
    ws.row_dimensions[4].height = 20; ws.row_dimensions[5].height = 34

    # 최근 6개월 추이
    rs = 7
    ws.cell(rs, 1, "▶ 최근 월별 정산 추이").font = Font(name=FONT, bold=True, size=11, color=C_TITLE)
    hdrs = ["연월", "정산금액", "건수", "전월대비", "증감율"]
    if has_profit: hdrs += ["매출총이익", "이익률"]
    for c, h in enumerate(hdrs, 1): hcell(ws.cell(rs+1, c), h)
    for r, row in enumerate(monthly.tail(6).itertuples(), rs+2):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, 1), row.년월_str, align="center", bg=bg)
        dcell(ws.cell(r, 2), row.정산금액, NUM, bg=bg)
        dcell(ws.cell(r, 3), row.건수, "#,##0", bg=bg)
        delta = row.증감액 if pd.notna(row.증감액) else None
        dbg = (C_POS if (delta or 0) >= 0 else C_NEG) if delta else bg
        dcell(ws.cell(r, 4), delta, NUM_DELTA, bg=dbg)
        pct = row.증감율 if pd.notna(row.증감율) else None
        dcell(ws.cell(r, 5), (pct/100) if pct else None, PCT, bg=dbg)
        if has_profit:
            dcell(ws.cell(r, 6), row.매출총이익 if pd.notna(row.매출총이익) else None, NUM, bg=bg)
            ir = row.이익률 if pd.notna(row.이익률) else None
            dcell(ws.cell(r, 7), (ir/100) if ir else None, "0.0%", bg=bg)

    # 사업장 TOP5
    col_off = 9
    ws.cell(rs, col_off, "▶ 사업장별 매출 TOP 5").font = Font(name=FONT, bold=True, size=11, color=C_TITLE)
    for c, h in enumerate(["사업장", "정산금액", "비중"], col_off): hcell(ws.cell(rs+1, c), h)
    for r, (client, row) in enumerate(pivot.nlargest(5, "합계").iterrows(), rs+2):
        bg = C_ALT if r % 2 == 0 else None
        dcell(ws.cell(r, col_off),   client,            align="left", bg=bg)
        dcell(ws.cell(r, col_off+1), row["합계"],        NUM, bg=bg)
        dcell(ws.cell(r, col_off+2), row["비중(%)"]/100, "0.00%", bg=bg)

    # 협력사 TOP5
    if coop_pivot is not None:
        col_cp = 13
        ws.cell(rs, col_cp, "▶ 협력사별 매출 TOP 5").font = Font(name=FONT, bold=True, size=11, color=C_TITLE)
        for c, h in enumerate(["협력사명", "정산금액"], col_cp): hcell(ws.cell(rs+1, c), h)
        for r, (coop, row) in enumerate(coop_pivot.nlargest(5, "합계").iterrows(), rs+2):
            bg = C_ALT if r % 2 == 0 else None
            dcell(ws.cell(r, col_cp),   coop,        align="left", bg=bg)
            dcell(ws.cell(r, col_cp+1), row["합계"],  NUM, bg=bg)

    # 추정 합계
    col_fc = 16
    ws.cell(rs, col_fc, "▶ 향후 3개월 추정").font = Font(name=FONT, bold=True, size=11, color=C_TITLE)
    for c, h in enumerate(["추정 월", "추정 정산금액"], col_fc): hcell(ws.cell(rs+1, c), h)
    for r, label in enumerate(forecast_labels, rs+2):
        dcell(ws.cell(r, col_fc),   label+"(추정)", align="center", bg=C_FCST)
        dcell(ws.cell(r, col_fc+1), forecast_df[label].sum(), NUM, bg=C_FCST)

    for c in range(1, 20): cw(ws, c, 16)
    ws.sheet_view.showGridLines = False


# ════════════════════════════════════════════════════════════════════════════
# 5. 메인
# ════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("   KT 정산 매출 자동 분석 스크립트")
    print("=" * 60)

    data = load_data()
    (monthly, pivot, month_cols, dept_pivot, coop_pivot,
     cat_summary, forecast_df, forecast_labels, has_profit) = analyze(data)

    wb = openpyxl.Workbook()
    del wb["Sheet"]

    write_dashboard(wb, monthly, pivot, coop_pivot, dept_pivot, forecast_df,
                    forecast_labels, has_profit)
    write_monthly(wb, monthly, has_profit)
    write_client_pivot(wb, pivot, month_cols)
    write_growth(wb, pivot, month_cols)
    write_dept(wb, dept_pivot, month_cols)
    write_coop(wb, coop_pivot)
    write_category(wb, cat_summary)
    write_forecast(wb, pivot, forecast_df, month_cols, forecast_labels)

    wb.save(OUTPUT_FILE)

    print(f"\n[완료] {OUTPUT_FILE}")
    print("-" * 60)
    print(f"  분석 기간       : {monthly['년월_str'].min()} ~ {monthly['년월_str'].max()}")
    print(f"  사업장 수       : {len(pivot)}개")
    if dept_pivot is not None:
        print(f"  부서 수         : {len(dept_pivot)}개")
    if coop_pivot is not None:
        print(f"  협력사 수       : {len(coop_pivot)}개")
    print(f"  카테고리 수     : {len(cat_summary) if cat_summary is not None else '없음'}")
    print(f"  추정 기간       : {', '.join(forecast_labels)}")
    print("=" * 60)

    if sys.platform == "win32":
        try:
            os.startfile(OUTPUT_FILE)
        except OSError:
            print(f"  (파일을 열 수 없습니다. 직접 열어주세요: {OUTPUT_FILE})")


if __name__ == "__main__":
    main()
