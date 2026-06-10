"""
run_tax_verify.py
통합플랫폼 자료만으로 과세구분 검증을 수행하고 결과 Excel을 출력합니다.

사용법:
  python run_tax_verify.py                      # 대화형 (파일 경로 직접 입력)
  python run_tax_verify.py 플랫폼.xlsx           # 플랫폼 파일만 지정
  python run_tax_verify.py 플랫폼.xlsx 마스터.xlsx  # 마스터(담당자 조회) 함께 지정
"""

# ── 0. 라이브러리 자동 설치 ────────────────────────────────────────
import subprocess, sys
for _p in ["pandas", "openpyxl"]:
    try:
        __import__(_p)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", _p, "-q"])

import os, logging, re
from pathlib import Path
from typing import Optional, Dict, List

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 1. 상수 ───────────────────────────────────────────────────────
FONT     = "Arial"
HDR_BLUE = "1F4E79"

# ── 2. 과세구분 검증 키워드 ────────────────────────────────────────
# 비과세 키워드
_TAX_NONTAX_KW = [
    "상품권", "기프티콘", "기프트카드", "gift card", "voucher",
    "교환권", "이용권", "모바일쿠폰", "선불카드", "충전권",
]

# 면세 키워드 (부가가치세법 시행규칙 제24조·별표1 기준)
# 면세 키워드 (부가가치세법 제26조 기준)
_TAX_EXEMPT_KW = [
    # ── 제1호: 미가공식료품 ─────────────────────────────────────────
    # 곡류 (정미·제분 등 1차 가공 이하)
    "쌀", "현미", "찹쌀", "잡곡", "밀", "밀가루", "보리", "귀리",
    "콩", "팥", "녹두", "율무", "수수", "기장", "메밀",
    "옥수수",
    # 유제품·난류
    "우유", "계란",
    # 육류 — 원물·신선·냉장·냉동 상태 면세
    # ※ 가공육(햄·소시지·베이컨 등)은 과세 → _TAX_EXCLUDE_KW 참조
    "육류", "소고기", "돼지고기", "닭고기", "오리고기", "한우",
    # 수산물 — 신선·냉장·냉동·염장·염수장·건조 상태만 면세
    # ※ 기름·조미료 첨가 가공품은 과세 → _TAX_EXCLUDE_KW 참조
    "생선", "수산물", "해물", "어패류",
    "참치", "연어", "고등어", "갈치", "조기", "광어", "우럭",
    "방어", "삼치", "가자미", "명태", "대구", "꽁치",
    "꽃게", "대게", "킹크랩",
    "굴", "홍합", "바지락", "전복", "낙지", "오징어", "멸치", "새우",
    "다시마", "미역", "톳", "파래", "해조류",
    # 채소류 — 신선·염장 등 단순가공 수준 면세
    "채소", "배추", "대파", "양파", "마늘", "감자", "고구마", "당근",
    "오이", "호박", "가지", "브로콜리", "양배추", "상추", "시금치", "깻잎",
    "버섯",
    # 과일류 — 신선 면세
    # ※ 통조림·주스·청 등 가공품은 과세 → _TAX_EXCLUDE_KW 참조
    "과일", "사과", "포도", "딸기", "복숭아", "참외", "수박", "감", "귤",
    "오렌지", "자몽", "레몬", "망고", "파인애플", "바나나", "키위",
    "블루베리", "체리",
    # ※ "배"(pear)는 오탐 위험으로 _KNOWN_EXEMPT_NAMES에서 복합어로만 관리
    # 단순가공식료품 (데치기·절임·염장 수준)
    "김치", "젓갈", "소금",
    # 견과류 — 가열·첨가물 없는 생 상태만 면세
    "견과",
    # ── 제3호: 연탄·무연탄 ──────────────────────────────────────────
    "연탄", "무연탄",
    # ── 제4호: 여성용 생리처리 위생용품 (2022.12.31. 면세 확대) ──────
    "생리대", "탐폰", "생리컵", "월경컵",
    # ── 제8호: 도서·신문·잡지 ──────────────────────────────────────
    "도서", "책", "교재", "서적", "신문", "잡지", "학습지",
    "전자책",    # 전자적 형태 도서 포함 (제26조 제8호)
    # ── 농업용품 ────────────────────────────────────────────────────
    "비료", "농약", "사료", "씨앗", "종자",
]

# 이 회사 과거 면세 처리 이력 기반 상품명 패턴
_KNOWN_EXEMPT_NAMES = [
    # 화훼·식물류 (생화·화분·관엽식물 = 미가공 농산물 면세)
    "화분", "꽃바구니", "화환",
    "동양란", "서양란",
    "뱅갈고무나무", "홍콩야자", "플랜트박스", "플랜트 박스",
    # 식품류 — 단독 키워드 경계 매칭 실패 보완
    "마른김", "건조김", "재래김", "돌김", "도시락김",
    "꿀", "허니",
    "허니스틱",
    "멸균우유",
    "락토프리",
    "과일세트",
    "배세트", "배선물", "배선물세트",   # '배'(pear) 단독 키워드 제외로 복합어 명시 등록
    "견과류",
    # 육류 복합어 — 경계 매칭 실패 케이스 보완
    "소고기세트", "소고기선물세트",
    "한우세트", "한우갈비", "한우등심", "한우선물세트",
    "삼겹살", "목살",
    "닭가슴살", "닭다리", "닭볶음용",
    "오리로스", "오리훈제",
    # 수산물 복합어
    "굴비", "굴비세트",
    "황태", "황태채", "북어채",
    # 제4호 여성 생리처리 위생용품 복합어
    "생리용품", "생리위생용품",
    # 도서류 — 경계 매칭 실패 보완
    "도서구매",
    "해외도서",
    "성경", "성경전서",
    "이북", "e-book", "ebook",
]

# 선물세트 등 컨테이너 키워드: 규격 필드도 검색할 트리거
_EXEMPT_SPEC_TRIGGER_KW = ["선물세트"]

# 제외 키워드: 면세/비과세 키워드 탐지되어도 이 단어 포함 시 제외
_TAX_EXCLUDE_KW = [
    # 배송·서비스 비용
    "배송비", "배송료", "운임", "운송료", "택배비",
    "도서산간", "제주도", "산간", "도서지역",
    "수수료", "설치비", "철거비", "인건비",
    # 가공식품·공산품
    "과자", "세제",
    # 가공유 (과세)
    "딸기우유", "딸기 우유",
    "초코우유", "초코 우유", "초콜릿우유",
    "바나나우유", "바나나 우유",
    "영양강화우유", "단백질우유",
    "가공유",
    # 가공육 (과세) — 염지·훈연·가열 등 2차 가공
    "햄", "소시지", "베이컨", "스팸", "핫도그", "살라미",
    "육가공", "가공육",
    # 소금 가공품 (과세)
    "암염", "핑크소금", "히말라야소금", "히말라야 소금",
    "죽염",
    "구운소금",
    # 식용유·유지류 (과세)
    "식용유", "식용유세트",
    # 조미 해조류 (과세)
    "조미김", "구운김", "볶음김", "김부각",
    "조미미역", "미역줄기볶음",
    "파래김", "조미파래",
    # 가공 견과류 (과세)
    "볶음땅콩", "구운아몬드", "볶은아몬드", "허니버터아몬드",
    "허니버터", "캔디드", "시즈닝견과", "믹스넛",
    "볶음견과", "구운견과", "가공견과",
    # 통조림류 (과세) — 참치통조림·꽁치통조림 등 포함
    "통조림", "캔참치", "참치캔",
    # 가공 과일류 (과세)
    "설탕절임", "시럽과일", "과일잼", "잼",
    "과일통조림", "통조림과일",
    "과일주스", "착즙주스",
    "주스",       # 오렌지주스·포도주스 등 가공음료 광범위 차단
    "에이드", "스무디",
    "과일청",
]

# ── 3. 헬퍼 함수 ──────────────────────────────────────────────────
def _bdr():
    s = Side(style="thin", color="C0C0C0")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_cell(cell):
    cell.font      = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", start_color=HDR_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _bdr()

def _read_excel_smart(path, search_cols=None, max_skip=10, **kwargs):
    """상단 제목/날짜 행을 건너뛰고 실제 헤더 행을 자동 탐지하여 DataFrame 반환."""
    if search_cols is None:
        search_cols = ["주문번호", "일정산번호"]
    for skip in range(max_skip):
        try:
            probe = pd.read_excel(path, header=skip, nrows=0, engine="openpyxl")
            if any(c in probe.columns for c in search_cols):
                return pd.read_excel(path, header=skip, engine="openpyxl", **kwargs)
        except Exception:
            pass
    return pd.read_excel(path, engine="openpyxl", **kwargs)

def _tax_kw_match(text: str, keywords: list):
    """상품명 필드용 — 한글 양쪽 경계 엄격 매칭.
    키워드 앞뒤에 한글(가-힣)이 오면 제외 (복합어 오탐 방지).
    """
    if not isinstance(text, str):
        return None, False
    t = text.strip()
    for kw in keywords:
        pattern = rf"(?<![가-힣]){re.escape(kw)}(?![가-힣])"
        if re.search(pattern, t, re.IGNORECASE):
            return kw, True
    return None, False

def _tax_kw_match_spec(text: str, keywords: list):
    """규격 필드용 — 오른쪽 한글 경계만 체크 (왼쪽은 허용).
    복합어 식품명에서 성분 키워드 탐지.
    """
    if not isinstance(text, str):
        return None, False
    t = text.strip()
    for kw in keywords:
        pattern = rf"{re.escape(kw)}(?![가-힣])"
        if re.search(pattern, t, re.IGNORECASE):
            return kw, True
    return None, False

def _norm_svc_key(raw: str) -> str:
    """서비스카테고리 키 정규화: '>' 주변 공백 통일."""
    return re.sub(r'\s*>\s*', ' > ', str(raw).strip())

# ── 4a. 과세확인목록 로드 (화이트리스트) ─────────────────────────────
def load_whitelist(whitelist_path: str, log) -> set:
    """과세확인목록 Excel에서 확인 완료 상품코드 집합을 로드한다.
    상품코드 컬럼만 필수. 나머지 컬럼(상품명·확인일·비고)은 관리용.
    """
    codes: set = set()
    try:
        df = pd.read_excel(whitelist_path, dtype=str, engine="openpyxl")
        code_c = next(
            (c for c in df.columns if "상품코드" in c),
            next((c for c in df.columns if "코드" in c), None),
        )
        if code_c is None:
            log.warning(f"  과세확인목록: '상품코드' 컬럼 없음 ({Path(whitelist_path).name})")
            return codes
        for v in df[code_c].dropna():
            code = re.sub(r'\.0$', '', str(v).strip())
            if code and code.lower() not in ("nan", ""):
                codes.add(code)
        log.info(f"  과세확인목록 로드: {len(codes)}건 ({Path(whitelist_path).name})")
    except Exception as e:
        log.warning(f"  과세확인목록 로드 실패: {e}")
    return codes


# ── 4b. 마스터 담당자 로드 (선택) ─────────────────────────────────
def load_svccat(master_path: str, log) -> Dict[str, tuple]:
    """매핑 마스터 Excel에서 서비스카테고리 → (관리회계, 담당자, 담당부서) 로드."""
    svccat: Dict[str, tuple] = {}
    try:
        xl = pd.ExcelFile(master_path, engine="openpyxl")
        log.info(f"  마스터 파일 시트 목록: {xl.sheet_names}")
        for sh in xl.sheet_names:
            df = xl.parse(sh)
            cols = df.columns.tolist()
            log.info(f"  [시트 '{sh}'] 컬럼: {cols}")

            # 카테고리 키 컬럼: 플랫폼 '서비스카테고리'(경로 문자열)와 조인할 컬럼
            # 우선순위: "서비스" + "카테고리" > "카테고리 분류" > 기타 "카테고리"
            # UID·명 등 숫자/단순명칭 컬럼은 제외
            _EXCLUDE = {"UID", "uid", "명", "코드", "ID", "id"}
            cat_c = (
                next((c for c in cols if "서비스" in c and "카테고리" in c), None)
                or next((c for c in cols if "카테고리" in c and "분류" in c), None)
                or next((c for c in cols
                         if "카테고리" in c
                         and "마스터" not in c
                         and not any(x in c for x in _EXCLUDE)), None)
            )
            if cat_c is None:
                log.info(f"  [시트 '{sh}'] 카테고리 컬럼 없음 → 건너뜀")
                continue

            # 담당자: "담당자" 또는 "상품담당자" 포함, 부서 제외
            # 가장 최근 날짜 컬럼 우선 (컬럼 순서상 앞에 있는 것)
            mgr_c  = next((c for c in cols if ("담당자" in c or "관리자" in c)
                           and "부서" not in c), None)
            dept_c = next((c for c in cols if "담당부서" in c), None)
            acc_c  = next((c for c in cols if "관리회계" in c), None)
            log.info(f"  [시트 '{sh}'] cat={cat_c!r}, mgr={mgr_c!r}, "
                     f"dept={dept_c!r}, acc={acc_c!r}")

            before = len(svccat)
            for _, r in df.iterrows():
                k = _norm_svc_key(str(r[cat_c]).strip())
                if k and k.lower() not in ("nan", ""):
                    svccat.setdefault(k, (
                        str(r[acc_c]).strip()  if acc_c  and pd.notna(r[acc_c])  else "",
                        str(r[mgr_c]).strip()  if mgr_c  and pd.notna(r[mgr_c])  else "",
                        str(r[dept_c]).strip() if dept_c and pd.notna(r[dept_c]) else "",
                    ))
            added = len(svccat) - before
            log.info(f"  [시트 '{sh}'] {added}건 추가 (누적 {len(svccat)}건)")

        # 샘플 출력 (최대 3건)
        for i, (k, v) in enumerate(list(svccat.items())[:3]):
            log.info(f"  [마스터 샘플] {k!r} → 관리회계={v[0]!r} 담당자={v[1]!r}")

        log.info(f"  마스터 담당자 로드 완료: 총 {len(svccat)}건 ({Path(master_path).name})")
    except Exception as e:
        log.warning(f"  마스터 로드 실패 (담당자 열 비움): {e}")
    return svccat

# ── 5. 과세구분 검증 핵심 로직 ────────────────────────────────────
def validate_tax(pl_path: str, svccat: Dict[str, tuple], log,
                 start_date=None, end_date=None,
                 whitelist: set = None) -> list:
    """플랫폼 파일에 대해 과세구분 검증 수행. tax_issues 리스트 반환.

    start_date / end_date: datetime.date 또는 None (None이면 전체 대상)
    날짜 필터 우선순위: 승인일 > 일일정산일 > 입고일
    """
    import datetime as _dt
    try:
        df_p = _read_excel_smart(pl_path, search_cols=["주문번호", "일정산번호"], dtype=str)
    except Exception as e:
        log.error(f"플랫폼 파일 로드 실패: {e}")
        return []

    log.info(f"  플랫폼 파일 로드: {len(df_p)}행  ({Path(pl_path).name})")

    # ── 승인일 기준 기간 필터 ─────────────────────────────────────
    if start_date or end_date:
        _DATE_COLS = [  # 우선순위 순
            "승인일", "일일정산일", "입고일",
            "입고승인일", "납품일자", "주문일", "결제일",
        ]
        date_col = next((c for c in _DATE_COLS if c in df_p.columns), None)
        if date_col:
            df_p[date_col] = pd.to_datetime(df_p[date_col], errors="coerce").dt.date
            before = len(df_p)
            if start_date:
                df_p = df_p[df_p[date_col] >= start_date]
            if end_date:
                df_p = df_p[df_p[date_col] <= end_date]
            df_p = df_p.reset_index(drop=True)
            sd_str = start_date.strftime("%Y-%m-%d") if start_date else "처음"
            ed_str = end_date.strftime("%Y-%m-%d")   if end_date   else "끝"
            log.info(f"  기간 필터 ({date_col}): {sd_str} ~ {ed_str}  "
                     f"{before}행 → {len(df_p)}행")
        else:
            log.warning(f"  승인일 컬럼 없음 → 기간 필터 생략 (전체 {len(df_p)}행 대상)")

    # 코드성 컬럼 소수점 제거
    for col in ["상품코드", "품목번호", "주문번호", "상품ID", "카테고리ID"]:
        if col in df_p.columns:
            df_p[col] = (df_p[col].astype(str)
                         .str.replace(r'\.0$', '', regex=True)
                         .str.strip())

    ord_c  = next((c for c in ["주문번호", "일정산번호"] if c in df_p.columns), None)
    tax_c  = next((c for c in ["매출과세구분", "과세구분"] if c in df_p.columns), None)
    code_c = next((c for c in ["상품코드", "자재번호", "품목코드"] if c in df_p.columns), None)
    name_c = next((c for c in ["상품명", "자재명", "품목명"] if c in df_p.columns), None)
    spec_c = next((c for c in ["대표규격", "규격", "사양"] if c in df_p.columns), None)
    cat_c  = next((c for c in df_p.columns if "서비스카테고리" in c), None)

    log.info(f"  컬럼 탐지: ord={ord_c!r}, tax={tax_c!r}, code={code_c!r}, "
             f"name={name_c!r}, spec={spec_c!r}, cat={cat_c!r}")

    if not tax_c:
        log.error("  매출과세구분 컬럼 없음 → 검증 불가")
        return []

    # 플랫폼 과세구분 값 정규화 (영세/영세율 → 면세)
    _PL_TAX_NORM = {"영세": "면세", "영세율": "면세"}
    df_p[tax_c] = (df_p[tax_c].astype(str).str.strip()
                   .map(lambda v: _PL_TAX_NORM.get(v, v)))

    # 유효한 과세구분 값 현황
    tax_dist = df_p[tax_c].value_counts().to_dict()
    log.info(f"  과세구분 현황: {tax_dist}")

    # 서비스카테고리 → 담당자 조회 샘플 검증
    if svccat and cat_c and cat_c in df_p.columns:
        sample_cats = (df_p[cat_c].dropna()
                       .astype(str).str.strip()
                       .loc[lambda s: s.str.len() > 0]
                       .unique()[:5])
        log.info(f"  [담당자조회 진단] 플랫폼 서비스카테고리 샘플: {list(sample_cats)}")
        for sc in sample_cats:
            nk = _norm_svc_key(sc)
            hit = svccat.get(nk)
            log.info(f"  [담당자조회 진단] {sc!r} → 정규화={nk!r} → "
                     f"{'HIT: ' + str(hit[1]) if hit else 'MISS'}")
    elif not cat_c:
        log.warning("  서비스카테고리 컬럼 없음 → 담당자 조회 불가")

    def _s(val) -> str:
        v = str(val).strip()
        return "" if v in ("nan", "None", "NaN") else v

    def _has_exclude(text: str) -> bool:
        if not isinstance(text, str):
            return False
        t = text.strip()
        return any(kw in t for kw in _TAX_EXCLUDE_KW)

    def _mgr(r) -> str:
        if not svccat or not cat_c:
            return ""
        cat_val = _s(r.get(cat_c, ""))
        if not cat_val:
            return ""
        return svccat.get(_norm_svc_key(cat_val), ("", "", ""))[1]

    whitelist = whitelist or set()
    tax_issues = []
    cnt_a = cnt_b = cnt_c = cnt_wl = 0

    # ── A·B: 키워드 vs 현재 과세구분 ──
    for _, r in df_p.iterrows():
        name     = _s(r.get(name_c, "")) if name_c else ""
        spec     = _s(r.get(spec_c, "")) if spec_c else ""
        cur_tax  = _s(r.get(tax_c,  ""))
        combined = f"{name} {spec}"
        row_code = re.sub(r'\.0$', '', _s(r.get(code_c, ""))) if code_c else ""

        # 과세확인목록(화이트리스트) 상품코드 → 검증 제외
        if whitelist and row_code and row_code in whitelist:
            cnt_wl += 1
            continue

        if _has_exclude(combined):
            continue

        # A: 비과세 키워드
        matched_kw, hit = _tax_kw_match(combined, _TAX_NONTAX_KW)
        if hit and cur_tax != "비과세":
            tax_issues.append({
                "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                "상품명"      : name,
                "대표규격"    : spec,
                "담당자"      : _mgr(r),
                "현재과세구분": cur_tax,
                "검증결과"    : "비과세 검토 필요",
                "탐지키워드"  : matched_kw,
                "사유"        : (f"상품명/규격에 '{matched_kw}' 포함. "
                                 "부가세법상 비과세 대상일 수 있습니다. 담당자 확인 필요."),
                "비고"        : "",
            })
            cnt_a += 1
            continue

        # B-1: 이력 기반 면세 패턴
        known_kw, _ = _tax_kw_match(name, _KNOWN_EXEMPT_NAMES)
        if not known_kw and any(t in name for t in _EXEMPT_SPEC_TRIGGER_KW):
            known_kw, _ = _tax_kw_match_spec(spec, _KNOWN_EXEMPT_NAMES)
        if known_kw and cur_tax != "면세":
            _tax_note = ("비과세로 등록되어 있으나" if cur_tax == "비과세"
                         else "과세로 등록되어 있으나")
            tax_issues.append({
                "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                "상품명"      : name,
                "대표규격"    : spec,
                "담당자"      : _mgr(r),
                "현재과세구분": cur_tax,
                "검증결과"    : "면세 검토 필요",
                "탐지키워드"  : known_kw,
                "사유"        : (f"상품명/규격에 '{known_kw}' 포함 (과거 면세 처리 이력 패턴). "
                                 f"{_tax_note} 면세 대상일 수 있습니다. 담당자 확인 필요."),
                "비고"        : "",
            })
            cnt_b += 1
            continue

        # B-2: 부가세법 일반 면세 키워드
        matched_kw, _ = _tax_kw_match(name, _TAX_EXEMPT_KW)
        if not matched_kw and any(t in name for t in _EXEMPT_SPEC_TRIGGER_KW):
            matched_kw, _ = _tax_kw_match_spec(spec, _TAX_EXEMPT_KW)
        if matched_kw and cur_tax != "면세":
            _tax_note = ("비과세로 등록되어 있으나" if cur_tax == "비과세"
                         else "과세로 등록되어 있으나")
            tax_issues.append({
                "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                "상품명"      : name,
                "대표규격"    : spec,
                "담당자"      : _mgr(r),
                "현재과세구분": cur_tax,
                "검증결과"    : "면세 검토 필요",
                "탐지키워드"  : matched_kw,
                "사유"        : (f"상품명/규격에 '{matched_kw}' 포함. "
                                 f"{_tax_note} 부가세법상 면세 대상일 수 있습니다. 담당자 확인 필요."),
                "비고"        : "",
            })
            cnt_b += 1

    # ── C: 동일 상품코드에 과세구분 혼재 ──
    if code_c and code_c in df_p.columns:
        valid = df_p[df_p[code_c].str.strip().ne("") & df_p[code_c].notna()].copy()
        grp = (valid.groupby(valid[code_c].str.strip())[tax_c]
               .apply(lambda s: sorted(s.dropna().unique().tolist())))
        mixed = grp[grp.apply(lambda v: len(v) > 1)]
        already_flagged = {r["상품코드"] for r in tax_issues}
        for code, tax_vals in mixed.items():
            code = re.sub(r'\.0$', '', code)
            if code in already_flagged:
                continue
            if whitelist and code in whitelist:
                cnt_wl += 1
                continue
            rows_c = df_p[df_p[code_c].str.strip() == code]
            tax_str = " / ".join(tax_vals)
            r0 = rows_c.iloc[0]
            n_rows   = len(rows_c)
            ord_val  = _s(r0.get(ord_c,  "")) if ord_c  else ""
            name_val = _s(r0.get(name_c, "")) if name_c else ""
            spec_val = _s(r0.get(spec_c, "")) if spec_c else ""
            ord_label = f"{ord_val} 외 {n_rows-1}건" if n_rows > 1 else ord_val
            tax_issues.append({
                "주문번호"    : ord_label,
                "상품코드"    : code,
                "상품명"      : name_val,
                "대표규격"    : spec_val,
                "담당자"      : _mgr(r0),
                "현재과세구분": tax_str,
                "검증결과"    : "과세구분 혼재",
                "탐지키워드"  : "",
                "사유"        : (f"동일 상품코드 {code}에 {tax_str} 혼재. "
                                 "상품 마스터 데이터 확인 필요."),
                "비고"        : "",
            })
            cnt_c += 1

    # ── 상품코드 기준 중복 제거 (A·B 섹션) ──
    if tax_issues:
        deduped: list = []
        code_idx: dict = {}
        for issue in tax_issues:
            code = str(issue.get("상품코드", "")).strip()
            if not code:
                deduped.append(issue.copy())
                continue
            if code not in code_idx:
                entry = issue.copy()
                entry["_ord_count"] = 1
                code_idx[code] = len(deduped)
                deduped.append(entry)
            else:
                idx = code_idx[code]
                deduped[idx]["_ord_count"] = deduped[idx].get("_ord_count", 1) + 1
                existing = deduped[idx].get("담당자", "")
                new_mgr  = issue.get("담당자", "")
                if new_mgr and new_mgr not in existing.split(", "):
                    deduped[idx]["담당자"] = (
                        f"{existing}, {new_mgr}".strip(", ") if existing else new_mgr
                    )
        for entry in deduped:
            cnt = entry.pop("_ord_count", 1)
            if cnt > 1:
                base_ord = str(entry.get("주문번호", "")).split(" 외 ")[0]
                entry["주문번호"] = f"{base_ord} 외 {cnt - 1}건"
        tax_issues = deduped

    log.info("=" * 48)
    log.info("[과세구분 검증 결과]")
    log.info(f"  비과세 검토 필요: {cnt_a}건  (상품권·기프티콘류 의심)")
    log.info(f"  면세 검토 필요:   {cnt_b}건  (면세 품목 의심)")
    log.info(f"  과세구분 혼재:    {cnt_c}건  (마스터 오류 가능)")
    if cnt_wl:
        log.info(f"  과세확인목록 제외: {cnt_wl}건  (담당자 확인 완료 상품코드)")
    log.info("=" * 48)
    if cnt_a + cnt_b + cnt_c > 0:
        log.warning("  계산서 발행 전 결과 파일을 반드시 확인하세요!")

    return tax_issues

# ── 6. Excel 시트 작성 ────────────────────────────────────────────
def write_tax_sheet(wb: openpyxl.Workbook, tax_issues: list,
                    start_date=None, end_date=None):
    ws = wb.active
    ws.title = "⓪과세구분검증"

    cnt_a   = sum(1 for i in tax_issues if i["검증결과"] == "비과세 검토 필요")
    cnt_b   = sum(1 for i in tax_issues if i["검증결과"] == "면세 검토 필요")
    cnt_c   = sum(1 for i in tax_issues if i["검증결과"] == "과세구분 혼재")
    cnt_red = cnt_a + cnt_c

    COLS  = ["상품코드", "주문번호", "상품명", "대표규격", "담당자",
             "현재과세구분", "검증결과", "탐지키워드", "사유", "담당자의견", "비고"]
    COL_W = [14, 18, 28, 24, 12, 12, 16, 12, 50, 30, 20]

    for i, w in enumerate(COL_W, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── 행 1: 검증 기간 표시 ──
    period_txt = ""
    if start_date or end_date:
        sd = start_date.strftime("%Y-%m-%d") if start_date else "처음"
        ed = end_date.strftime("%Y-%m-%d")   if end_date   else "끝"
        period_txt = f"승인일 기준 검증 기간: {sd} ~ {ed}"
    ws.merge_cells(f"A1:{get_column_letter(len(COLS))}1")
    c1 = ws.cell(1, 1, period_txt)
    c1.font      = Font(name=FONT, size=10, color="44546A", italic=True)
    c1.fill      = PatternFill("solid", start_color="D6E4F0")
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 18

    # ── 행 2: 결과 배너 ──
    ws.merge_cells(f"A2:{get_column_letter(len(COLS))}2")
    if not tax_issues:
        banner_txt = "✓ 과세구분 이상 항목 없음"
        b_fg, b_bg = "375623", "CCFFCC"
    elif cnt_red > 0:
        banner_txt = (f"과세구분 검토 필요: 비과세 {cnt_a}건 / 면세 {cnt_b}건 / 혼재 {cnt_c}건"
                      "   ※ 확정 오류 아님 — 담당자 검토 후 판단")
        b_fg, b_bg = "FFFFFF", "C00000"
    else:
        banner_txt = (f"과세구분 검토 필요: 면세 검토 {cnt_b}건"
                      "   ※ 확정 오류 아님 — 담당자 검토 후 판단")
        b_fg, b_bg = "7B3F00", "FFD966"

    c2 = ws.cell(2, 1, banner_txt)
    c2.font      = Font(name=FONT, bold=True, size=12, color=b_fg)
    c2.fill      = PatternFill("solid", start_color=b_bg)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    data_row = 4

    if not tax_issues:
        ws.cell(data_row, 1, "(키워드 탐지 건 없음)").font = Font(
            name=FONT, size=10, color="808080", italic=True)
        return

    # 헤더
    for ci, h in enumerate(COLS, 1):
        hdr_cell(ws.cell(data_row, ci, h))
    ws.row_dimensions[data_row].height = 18
    ws.freeze_panes = f"A{data_row + 1}"
    data_row += 1

    # 데이터
    fill_map = {
        "비과세 검토 필요": "FFCCCC",
        "과세구분 혼재"   : "FFCCCC",
        "면세 검토 필요"  : "FFFF99",
    }
    code_ci = COLS.index("상품코드") + 1

    for issue in tax_issues:
        f = fill_map.get(issue.get("검증결과", ""), None)
        for ci, col in enumerate(COLS, 1):
            val  = issue.get(col, "")
            cell = ws.cell(data_row, ci, val)
            cell.font      = Font(name=FONT, size=10)
            cell.border    = _bdr()
            cell.alignment = Alignment(
                horizontal="left", vertical="center",
                wrap_text=(col in ("사유", "비고", "담당자의견")))
            if ci == code_ci:
                cell.number_format = "@"
            if f:
                cell.fill = PatternFill("solid", start_color=f)
        ws.row_dimensions[data_row].height = 16
        data_row += 1

# ── 7. 날짜 자동 판단 ─────────────────────────────────────────────
def auto_period():
    """오늘 날짜 기준으로 해당 정산 기간의 시작일·종료일(=오늘) 반환.
    1~15일: 당월 1일 ~ 오늘
    16~말일: 당월 16일 ~ 오늘
    """
    import datetime as _dt
    today = _dt.date.today()
    if today.day <= 15:
        start = today.replace(day=1)
    else:
        start = today.replace(day=16)
    return start, today


# ── 8. 메인 ───────────────────────────────────────────────────────
def main():
    import datetime as _dt
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    log = logging.getLogger("tax_verify")

    # ── 파일 경로 결정 ──
    args = sys.argv[1:]

    if args:
        pl_path = args[0].strip('"').strip("'")
        master_path = args[1].strip('"').strip("'") if len(args) > 1 else None
    else:
        print("\n[과세구분 검증 도구]")
        print("통합플랫폼 자료만으로 과세구분 검증을 수행합니다.\n")
        pl_path = input("통합플랫폼 Excel 파일 경로: ").strip().strip('"').strip("'")
        master_path_input = input(
            "매핑마스터 파일 경로 (담당자 조회용, 없으면 Enter): "
        ).strip().strip('"').strip("'")
        master_path = master_path_input if master_path_input else None

    if not pl_path or not Path(pl_path).exists():
        log.error(f"파일을 찾을 수 없습니다: {pl_path!r}")
        sys.exit(1)

    # ── 정산 기간 자동 판단 ──
    start_date, end_date = auto_period()
    log.info(f"  검증 기간 (승인일): {start_date} ~ {end_date}")

    # ── 담당자 마스터 로드 (선택) ──
    svccat: Dict[str, tuple] = {}
    if master_path:
        if Path(master_path).exists():
            svccat = load_svccat(master_path, log)
        else:
            log.warning(f"마스터 파일 없음, 담당자 열 비움: {master_path!r}")

    # ── 과세확인목록 자동 탐지 ──
    whitelist: set = set()
    _wl_auto = Path(pl_path).parent / "과세확인목록.xlsx"
    if _wl_auto.exists():
        whitelist = load_whitelist(str(_wl_auto), log)

    # ── 검증 수행 ──
    log.info(f"\n[STEP 1] 플랫폼 파일 로드 및 과세구분 검증")
    tax_issues = validate_tax(pl_path, svccat, log,
                              start_date=start_date, end_date=end_date,
                              whitelist=whitelist)

    # ── 출력 파일 경로 ──
    pl_stem = Path(pl_path).stem
    out_dir  = Path(pl_path).parent
    out_name = f"과세구분검증_{pl_stem}.xlsx"
    out_path = out_dir / out_name

    # ── Excel 작성 ──
    log.info(f"\n[STEP 2] Excel 작성: {out_path}")
    wb = openpyxl.Workbook()
    write_tax_sheet(wb, tax_issues, start_date=start_date, end_date=end_date)
    wb.save(str(out_path))

    log.info(f"\n완료 → {out_path}")
    total = len(tax_issues)
    if total == 0:
        log.info("  검토 필요 항목 없음 ✓")
    else:
        log.info(f"  총 {total}건 검토 필요 — 결과 파일을 확인하세요.")


if __name__ == "__main__":
    main()
