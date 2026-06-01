"""
run_settlement.py  ─  KT 정산 자동화 시스템 v2.0
=====================================================
실행: python run_settlement.py
로그: settlement_log.txt
출력: 정산결과_YYYYMMDD.xlsx
"""

# ── 0. 라이브러리 자동 설치 ───────────────────────────────────────
import subprocess, sys
for _p in ["pandas", "openpyxl"]:
    try:
        __import__(_p)
    except ImportError:
        print(f"[설치중] {_p}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", _p, "-q"])

import os, logging, datetime, calendar
from pathlib import Path
from typing import Optional, Dict, Tuple, List

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 1. 상수 ──────────────────────────────────────────────────────

FONT = "Arial"

# 행 하이라이트 색상
ROW_MISSING  = "FFCCCC"   # 플랫폼 누락  → 빨강
ROW_RETURN   = "DDEEFF"   # 반품          → 파랑
ROW_AMOUNT   = "FFFF99"   # 금액 불일치   → 노랑
ROW_EVEN     = "F2F2F2"   # 짝수 행 기본

# 헤더 색상
HDR_BLUE     = "1F4E79"
HDR_GREEN    = "375623"
HDR_RED      = "C00000"
HDR_GRAY     = "808080"
ROW_TOTAL    = "D9D9D9"   # 합계 행 배경

# KT 핵심 컬럼 후보 목록  (실제 KT raw 파일 기준)
_KT_CANDIDATES = {
    "주문번호"    : ["구매문서번호"],                          # KT 구매문서번호
    "구매품목"    : ["구매품목"],                              # 라인 아이템 번호 (복합 키 후반부)
    "요청번호"    : ["인수증"],                                # KT 인수증 (플랫폼 요청번호 대응)
    "입고일"      : ["입고승인일", "납품일자"],                  # 입고승인일 우선
    "정산금액"    : ["공급가액"],                              # KT 공급가액 (세전)
    "국책과제여부": ["국책과제여부"],                          # 국책과제 분리용
    "협력사명"    : ["공급자명", "협력사명"],                  # KT 공급자명
    "매출과세구분": ["세금코드명", "매출과세구분"],            # KT 세금코드명 → 정규화
    "이동유형"    : ["이동유형명", "이동유형"],                # 이동유형명 우선
    "관리회계"    : ["관리회계", "관리회계(IP만변경적용)"],
    "담당자"      : ["담당자"],
    "일반통신구분": ["일반/통신구분", "일반통신구분"],
}
# 플랫폼 핵심 컬럼 후보 목록  (플랫폼이 인보이스 기준)
_PL_CANDIDATES = {
    "주문번호"      : ["주문번호"],
    "품목번호"      : ["품목번호"],
    "주문품목키"    : ["주문&품목"],                                 # 플랫폼 사전계산 복합 키
    "입고일"        : ["입고일"],
    "승인일"        : ["승인일"],                                      # 플랫폼 입고승인일 (= KT 입고승인일)
    "일일정산일"    : ["일일정산일"],
    "정산금액"      : ["정산금액"],
    "매입금액"      : ["매입금액"],
    "매출총이익"    : ["매출총이익"],
    "협력사명"      : ["협력사명"],
    "매출과세구분"  : ["매출과세구분"],
    "서비스카테고리": ["서비스카테고리"],
    "중분류"        : ["관리회계\n(IP만변경적용)", "관리회계(IP만변경적용)",   # 실제 파일 컬럼명 우선
                       "중분류(IP만변경적용)", "중분류"],
    "담당자부서"    : ["담당자부서"],
    "결제유형"      : ["결제유형"],
    "이동유형"      : ["인보이스\n(KT이동유형명 값)", "인보이스(KT이동유형명 값)", "이동유형"],
    "담당자"        : ["담당자"],
    "일반통신구분"  : ["일반/통신구분", "일반통신구분"],
}
_PL_OPT = {"요청번호": ["요청번호"]}  # optional — 없으면 경로 B

MONEY_FMT = "#,##0"

# ── 과세구분 검증 키워드 (부가가치세법 기준) ──────────────────────────────────
# TODO: mapping_master.xlsx에 [과세구분키워드] 시트 추가 시
#        비과세/면세/제외 키워드를 Excel에서 직접 관리 가능
#        (코드 수정 없이 담당자가 키워드 추가/삭제)
# TODO: 취급 품목 특성에 맞게 키워드 조정 필요
#        현재 키워드는 일반적인 부가세법 기준이므로
#        실제 오탐/미탐 발생 시 제외키워드 또는 키워드 목록 수정

# 비과세 키워드 ('쿠폰','포인트','캐시','마일리지'는 오탐 가능성 높으므로 제외)
_TAX_NONTAX_KW = [
    "상품권", "기프티콘", "기프트카드", "gift card", "voucher",
    "교환권", "이용권", "모바일쿠폰", "선불카드", "충전권",
]
# 면세 키워드 (부가가치세법 시행규칙 제24조·별표1 기준)
# ※ 미가공식료품: 원생산물 본래 성질이 변하지 않는 1차 가공 이하만 면세
# ※ 제외된 항목:
#    설탕·식용유·버터·치즈 → 원생산물 성질이 변하는 가공품 (과세 대상)
#    두부·된장·간장·고추장 → 2026.1.1부터 제조시설 갖춰 독립 포장 판매분은 과세 대상
# ※ '치료','진단' 단독은 오탐 가능성 높으므로 복합 명사 형태로만 등록
_TAX_EXEMPT_KW = [
    # 미가공 곡류 (별표1 — 정미·제분 등 1차 가공 이하)
    "쌀", "밀", "보리", "콩", "옥수수", "밀가루", "소금",
    # 미가공 신선 식품
    "우유", "계란", "육류", "생선", "채소", "과일",
    # 단순가공식료품 (별표1 제12호 ⑤ — 데치기·절임·염장 수준)
    "김치", "젓갈",
    # 해조류·수산물 — 신선·냉장·냉동·염장·염수장·건조 상태만 면세
    # 근거: 부가가치세법 제12조 제1항 제1호, 시행령 제28조 제1항,
    #       시행규칙 제24조 [별표 1]
    # ※ 기름·소금·조미료 등 첨가하여 굽거나 가공한 제품은 과세 → _TAX_EXCLUDE_KW 참조
    "다시마",
    "미역", "톳", "파래", "해조류",
    "굴", "홍합", "바지락", "전복", "낙지", "오징어", "멸치", "새우",
    # 견과류 — 가열·첨가물 없는 생 상태만 면세
    # ※ 볶음·구운 것·허니버터 등 가공품은 과세 → _TAX_EXCLUDE_KW 참조
    "견과",
    # ※ 의료기기·의약품 제거 이유:
    #   부가세법상 '의료보건용역'(진료·수술 등 서비스)은 면세이나,
    #   '의료기기 판매'는 재화의 공급에 해당하여 과세 대상임.
    #   KT 구매 맥락(재화 구입)에서 의료기기류는 과세 → 면세 오탐 방지를 위해 제외.
    # 도서·교육자료
    "도서", "책", "교재", "서적", "신문", "잡지", "학습지",
    # 농업용품
    "비료", "농약", "사료", "씨앗", "종자",
]
# ── 이 회사 과거 면세 처리 이력 기반 상품명 패턴 ──────────────────────────────────
# 아래 키워드가 상품명·규격에 포함되면 현재 과세구분과 무관하게 면세 검토 대상 추출
# ※ 부가세법 일반 면세 키워드(_TAX_EXEMPT_KW)와 별도 관리
# ※ 추가·삭제는 실무 담당자가 직접 관리 — 이 회사 구매 이력 기반
# ※ 부분 문자열 매칭 사용 (복합어·고유명사 포함)
_KNOWN_EXEMPT_NAMES = [
    # 화훼·식물류 (생화·화분·관엽식물 = 미가공 농산물 면세)
    "화분", "꽃바구니", "화환",
    "동양란", "서양란",
    "뱅갈고무나무", "홍콩야자", "플랜트박스", "플랜트 박스",

    # 식품류 — 단독 키워드(_TAX_EXEMPT_KW)로 탐지 안 되는 복합어 보완
    # 건조 김 — 시행규칙 제24조 [별표 1] 면세 대상 (마른김·재래김·돌김 등)
    # ※ '조미김','구운김' 등 가공 제품은 _TAX_EXCLUDE_KW에서 별도 차단
    "마른김", "건조김", "재래김", "돌김", "도시락김",
    "꿀", "허니",
    "허니스틱",        # '허니' 뒤에 한글 '스틱' → 양쪽경계 실패 → 명시 등록
    "멸균우유",        # '우유' 앞에 한글 '균' → 양쪽경계 실패 → 명시 등록 (면세)
    "락토프리",        # 유당분해 우유 = 면세 (흰 우유와 동일 기준)
    "과일세트",        # '과일' 뒤에 한글 '세' → 양쪽경계 실패 → 명시 등록
    "견과류",          # '견과' 뒤에 한글 '류' → 양쪽경계 실패 → 명시 등록 (생 견과 기준 면세)

    # 도서류 — '도서' 키워드로 탐지 안 되는 복합어 보완
    "도서구매",        # '도서' 뒤에 한글 '구' → 양쪽경계 실패 → 명시 등록
    "해외도서",
    "성경", "성경전서",  # 종교 출판물 (부가세법상 면세 도서) — 전서 붙으면 경계 실패 → 전서도 명시
]
# ※ 개별 도서 서명: 상품명·규격에 '도서' 표기가 없으면 자동 탐지 불가
#   → 구매 시 상품명 또는 규격에 '도서' 포함 표기 권장

# ── 규격(spec) 필드 추가 검색 트리거 ────────────────────────────────────────
# 이 키워드가 상품명(name)에 포함될 때만 규격 필드를 추가 검색
# ※ 상품 내용물을 규격에서만 알 수 있는 '포장 용도' 상품명에 한정
# ※ 예: '선물세트' → 규격의 '밥다시마300g + 쌀톳200g' 검색
#        '랙', '우산' 등 일반 상품 → 규격 검색 안 함
_EXEMPT_SPEC_TRIGGER_KW = ["선물세트"]

# 제외 키워드: 아래 단어 포함 시 면세/비과세 키워드 매칭되어도 검증 대상에서 제외
# (가공식품·공산품 등 명백히 과세 대상인 품목명 오탐 방지)
_TAX_EXCLUDE_KW = [
    "배송비", "배송료", "운임", "운송료", "택배비",
    "도서산간", "제주도", "산간", "도서지역",
    "수수료", "설치비", "철거비", "인건비",
    # 가공식품·공산품 (과세 대상 — 면세 오탐 방지)
    "과자", "세제",
    # 가공유 (과세) — '우유' 키워드 오탐 방지
    # 근거: 향료·당분·영양강화 성분 첨가 시 가공식품으로 분류 → 과세
    # ※ '딸기우유'·'초코우유'는 경계 매칭으로 이미 차단되나,
    #    '딸기 우유'처럼 공백 표기 시 통과 가능 → 명시 차단
    "딸기우유", "딸기 우유",
    "초코우유", "초코 우유", "초콜릿우유",
    "바나나우유", "바나나 우유",
    "영양강화우유", "단백질우유",
    "가공유",
    # 소금 가공품 (과세) — 면세 소금 = 천일염·재제소금에 한정
    # 근거: 국세청 사전-2015-법령해석부가-22498 (2015.3.17.)
    #       부가세법 시행령 제34조 제1항 제13호
    # ① 암염(히말라야 핑크소금 등) — 불순물 제거·가공 공정 포함 → 과세
    "암염", "핑크소금", "히말라야소금", "히말라야 소금",
    # ② 죽염 — 대나무 소성(燒成) 공정, 원생산물 성질이 변하는 2차 가공 → 과세
    "죽염",
    # ③ 구운소금 — 고온 가열 공정 → 과세
    "구운소금",
    # 식용유·유지류 (과세) — 압착·정제 공정으로 원생산물 성질이 변하는 가공품
    # ※ 부가세법 시행령 제34조 면세 미가공식료품 목록에 포함되지 않음
    # ※ 소금+식용유 합포장 선물세트: 과세(식용유)+면세(소금) 혼합 → 세트 전체 과세
    #    → 상품명에 '선물세트' 포함 시 규격 필드도 검색(_EXEMPT_SPEC_TRIGGER_KW)하므로
    #       규격에 '식용유'가 있으면 자동으로 과세 처리됨
    "식용유",
    "식용유세트",
    # 조미 해조류 (과세) — 기름·소금·조미료 등 첨가하여 굽거나 가공한 제품
    # 근거: 부가가치세법 제12조 제1항 제1호, 시행령 제28조 제1항,
    #       시행규칙 제24조 [별표 1] — 면세 김 = 신선·냉장·냉동·염장·염수장·건조에 한정
    #       국세청 부가가치세과-1109 (2013.11.28.)
    "조미김", "구운김", "볶음김", "김부각",
    "조미미역", "미역줄기볶음",
    "파래김", "조미파래",
    # 가공 견과류 (과세) — 볶음·구이·첨가물 가미 시 과세
    "볶음땅콩", "구운아몬드", "볶은아몬드", "허니버터아몬드",
    "허니버터", "캔디드", "시즈닝견과", "믹스넛",
    "볶음견과", "구운견과", "가공견과",
    # 가공 과일류 (과세) — 설탕·감미료·조리 가미 시 과세
    # 근거: 국세청 사전-2015-법령해석부가-0192 (2015.7.3.)
    #       물에 삶거나 찐 것, 설탕·감미료 첨가한 냉동과실 = 면세 제외
    "설탕절임", "시럽과일", "과일잼", "잼",
    "과일통조림", "통조림과일",
    "과일주스", "착즙주스",
    "과일청",                  # 설탕 첨가 → 과세
]

# ── 기업명 정규화 (법인 형태 접미사·접두사 제거) ───────────────────────────────
import re as _re_corp
def _normalize_corp_name(name: str) -> str:
    """법인 형태 표기 제거 후 공백 없는 대문자 문자열 반환.
    예: '(주)케이티커머스' → '케이티커머스', '삼성전자 주식회사' → '삼성전자'
    """
    s = str(name).strip()
    # 접두사 제거: (주), ㈜, 주식회사
    s = _re_corp.sub(r'^\(주\)\s*|^㈜\s*|^주식회사\s+', '', s)
    # 접미사 제거
    s = _re_corp.sub(
        r'\s*\(주\)$|㈜$'
        r'|\s+주식회사$|\s+유한회사$|\s+유한책임회사$'
        r'|\s+합자회사$|\s+합명회사$|\s+농업회사법인$'
        r'|\s+Inc\.$|\s+Ltd\.$|\s+Co\.,?\s*Ltd\.$',
        '', s, flags=_re_corp.IGNORECASE
    )
    return _re_corp.sub(r'\s+', '', s).upper()   # 공백 제거 + 대문자


def _norm_svc_key(s: str) -> str:
    """서비스카테고리 키 정규화: '>' 전후 공백 제거 후 전체 strip.
    예: '통신장비 > 유무선전송 > 광네트워크장치' → '통신장비>유무선전송>광네트워크장치'
    매핑파일과 플랫폼 데이터 간 공백 표기 차이를 흡수한다.
    """
    return _re_corp.sub(r'\s*>\s*', '>', str(s).strip())


# ── 마스터카테고리 기준 매핑 예외 목록 ───────────────────────────────────────
# 이 목록에 있는 마스터카테고리(CMS) 값은 서비스카테고리 대신 마스터카테고리로
# 관리회계·담당자·일반통신구분을 조회합니다.
# 값은 _norm_svc_key() 정규화 후 비교하므로 '>' 전후 공백은 무관합니다.
_MCAT_OVERRIDE_SET: set = {
    _norm_svc_key("통신장비 > 접속/선로자재 > 선로자재 > 전선용랩/전선용테이프"),
    # 추가 예외가 필요하면 아래에 계속 작성
    # _norm_svc_key("..."),
}


# ── 대기업 계열사 목록 (공정거래위원회 공시대상기업집단 기준) ──────────────────
# ※ 출처: 공정거래위원회 대기업집단 포털 + 재계 서열 기준 주요 계열사
# ※ 법인 형태 제거 후 대문자 정규화된 기업명 집합
# ※ 추가/삭제: 기업명을 그대로 넣으면 _normalize_corp_name() 로 자동 정규화
_LARGE_CORP_RAW = [
    # ── 삼성그룹 (재계 1위) ─────────────────────────────────
    "삼성전자", "삼성SDS", "삼성물산", "삼성SDI", "삼성디스플레이",
    "삼성전기", "삼성생명", "삼성화재", "삼성증권", "삼성카드",
    "삼성바이오로직스", "삼성바이오에피스", "에스원", "삼성웰스토리",
    "삼성서울병원", "호텔신라", "제일기획",
    # ── SK그룹 (재계 2위) ────────────────────────────────────
    "SK텔레콤", "SK하이닉스", "SK이노베이션", "SK에너지", "SK지오센트릭",
    "SK E&S", "SK네트웍스", "SK브로드밴드", "SK실트론", "SKC",
    "SK에코플랜트", "SK바이오사이언스", "SK바이오팜", "SK스퀘어",
    "SK렌터카", "SK매직", "SK플래닛", "11번가", "원스토어",
    # ── 현대자동차그룹 (재계 3위) ────────────────────────────
    "현대자동차", "기아", "현대모비스", "현대글로비스", "현대건설",
    "현대위아", "현대트랜시스", "현대오토에버", "현대캐피탈", "현대카드",
    "현대커머셜", "현대로템", "현대제철", "이노션",
    # ── LG그룹 (재계 4위) ───────────────────────────────────
    "LG전자", "LG화학", "LG디스플레이", "LG이노텍", "LG유플러스",
    "LG CNS", "LG에너지솔루션", "LG헬로비전", "LG생활건강", "LG경영개발원",
    "지투알", "시너지모티브", "LG경제연구원",
    # ── 롯데그룹 (재계 5위) ─────────────────────────────────
    "롯데쇼핑", "롯데케미칼", "롯데칠성음료", "롯데정보통신", "롯데건설",
    "롯데하이마트", "롯데홈쇼핑", "롯데렌탈", "롯데멤버스", "롯데지주",
    "코리아세븐", "롯데카드", "롯데손해보험",
    # ── 포스코그룹 (재계 6위) ───────────────────────────────
    "POSCO홀딩스", "포스코", "포스코DX", "포스코ICT", "포스코인터내셔널",
    "포스코건설", "포스코퓨처엠", "포스코실리콘솔루션",
    # ── 한화그룹 (재계 7위) ─────────────────────────────────
    "한화솔루션", "한화에어로스페이스", "한화시스템", "한화생명",
    "한화손해보험", "한화건설", "한화투자증권", "한화갤러리아", "한화오션",
    "한화비전", "한화정밀기계",
    # ── GS그룹 (재계 8위) ───────────────────────────────────
    "GS칼텍스", "GS리테일", "GS건설", "GS ITM", "GS홈쇼핑",
    "GS에너지", "GS파워", "GS글로벌", "GS이피에스",
    # ── HD현대그룹 (재계 9위) ─────────────────────────────
    "HD현대", "현대중공업", "현대미포조선", "현대삼호중공업",
    "HD현대일렉트릭", "HD현대인프라코어", "HD현대마린솔루션",
    # ── 농협그룹 (재계 10위) ────────────────────────────────
    "농협은행", "농협생명", "농협손해보험", "NH투자증권", "농협중앙회",
    "NH농협은행",
    # ── 신세계그룹 (재계 11위) ──────────────────────────────
    "신세계", "이마트", "SSG닷컴", "신세계인터내셔날", "신세계건설",
    "신세계I&C", "스타벅스코리아", "신세계푸드", "이마트에브리데이",
    "G마켓", "옥션",
    # ── CJ그룹 (재계 12위) ──────────────────────────────────
    "CJ대한통운", "CJ제일제당", "CJ ENM", "CJ올리브네트웍스", "CJ CGV",
    "CJ씨푸드", "CJ프레시웨이", "CJ푸드빌", "CJ지엘에스",
    # ── 카카오그룹 (재계 13위) ──────────────────────────────
    "카카오", "카카오뱅크", "카카오페이", "카카오모빌리티",
    "카카오엔터테인먼트", "카카오게임즈", "카카오VX",
    # ── 네이버그룹 (재계 14위) ──────────────────────────────
    "네이버", "네이버파이낸셜", "네이버클라우드", "네이버웹툰", "라인플러스",
    # ── 두산그룹 (재계 15위) ────────────────────────────────
    "두산에너빌리티", "두산밥캣", "두산퓨얼셀", "두산로보틱스",
    "두산테스나", "두산중공업", "두산인프라코어",
    # ── LS그룹 (재계 16위) ──────────────────────────────────
    "LS일렉트릭", "LS전선", "LS MnM", "LS니꼬동제련", "LS엠트론",
    "가온전선", "E1", "예스코홀딩스",
    # ── KT그룹 (재계 17위) ──────────────────────────────────
    "KT", "KT M&S", "KT sat", "KT Cloud", "kt ds",
    "kt is", "KT링커스", "KT스카이라이프", "KT에스테이트",
    "KT텔레캅", "케이티텔레캅",                    # 보안·경비 서비스
    "KT스포츠", "케이티스포츠",                    # 스포츠 사업
    "KT희망지움", "케이티희망지움",                # 사회공헌 재단 (구 표기)
    "KT희망지음", "케이티희망지음",                # 사회공헌 재단 (2026 공시대상기업집단 표기)
    "KT커머스", "케이티커머스",                    # 커머스 사업
    "KT알파", "케이티알파",                        # 콘텐츠·커머스 (구 KT하이텔)
    "KTCS", "케이티씨에스",                        # 고객서비스
    "KT파워텔", "케이티파워텔",                    # 무선통신
    "케이티엠앤에스", "케이티클라우드",
    "케이티링커스", "케이티스카이라이프", "케이티에스테이트",
    # ── 효성그룹 (재계 18위) ────────────────────────────────
    "효성", "효성ITX", "효성인포메이션시스템", "효성첨단소재",
    "효성중공업", "효성화학", "갤럭시아머니트리",
    # ── 영풍그룹 (재계 19위) ────────────────────────────────
    "영풍", "고려아연", "코리아니켈",
    # ── 코오롱그룹 (재계 20위) ──────────────────────────────
    "코오롱인더스트리", "코오롱글로텍", "코오롱베니트", "코오롱생명과학",
    "코오롱플라스틱", "코오롱모빌리티그룹",
    # ── 한진그룹 (재계 21위) ────────────────────────────────
    "대한항공", "한진", "진에어", "한국공항", "한진칼",
    # ── 현대백화점그룹 (재계 22위) ──────────────────────────
    "현대백화점", "현대홈쇼핑", "현대그린푸드", "현대리바트",
    "현대이지웰", "한섬",
    # ── DB그룹 (재계 23위) ──────────────────────────────────
    "DB하이텍", "DB인슈어런스", "DB금융투자", "DB손해보험",
    # ── 하림그룹 (재계 24위) ────────────────────────────────
    "하림", "하림지주", "팬오션",
    # ── 세아그룹 (재계 25위) ────────────────────────────────
    "세아제강", "세아홀딩스", "세아베스틸",
    # ── DL그룹 (재계 26위) ──────────────────────────────────
    "DL건설", "DL이앤씨", "여천NCC",
    # ── OCI그룹 (재계 27위) ─────────────────────────────────
    "OCI홀딩스", "OCI",
    # ── KCC그룹 (재계 28위) ─────────────────────────────────
    "KCC", "KCC글라스",
    # ── 태광그룹 (재계 29위) ────────────────────────────────
    "태광산업", "흥국화재", "흥국생명",
    # ── 금호아시아나그룹 (재계 30위) ────────────────────────
    "아시아나항공", "금호타이어", "금호산업", "에어서울", "에어부산",
    # ── 미래에셋그룹 ─────────────────────────────────────────
    "미래에셋증권", "미래에셋자산운용", "미래에셋생명",
    # ── 한국투자금융그룹 ─────────────────────────────────────
    "한국투자증권", "한국투자저축은행",
    # ── 교보생명그룹 ─────────────────────────────────────────
    "교보생명", "교보증권", "교보문고",
    # ── 메리츠그룹 ───────────────────────────────────────────
    "메리츠화재", "메리츠증권",
    # ── 부영그룹 ─────────────────────────────────────────────
    "부영",
    # ── 기타 주요 대기업 ─────────────────────────────────────
    "KT&G", "한국전력", "한국전력공사", "한국도로공사", "한국수력원자력",
    "한국가스공사", "한국토지주택공사", "LH공사", "인천국제공항공사",
    "한국철도공사", "코레일",
]
# 정규화 집합으로 변환 (법인 형태 제거 + 대문자)
_LARGE_CORP_SET: set = {_normalize_corp_name(n) for n in _LARGE_CORP_RAW}


def _load_ftc_large_corps(path: "Path") -> set:
    """대기업 계열회사 목록 파일에서 기업명을 읽어 정규화된 집합을 반환한다.

    지원 형식:
      1. 대기업목록.xlsx  — A열(또는 첫 번째 열)에 회사명을 한 줄씩 나열한 단순 목록
      2. (별첨5) xlsx     — 공정위 배포 형식 (컬럼5=비금융회사, 컬럼7=금융·보험회사)
      3. .zip             — 공정위 배포 ZIP (내부에서 별첨5 xlsx를 자동 탐색)

    Args:
        path: 위 형식 중 하나의 파일 경로

    Returns:
        정규화된 기업명 집합 (빈 집합이면 로드 실패)
    """
    import io as _io, zipfile as _zf
    _pd = __import__("pandas")

    SKIP_HEADERS = {
        "비금융·비보험회사 명단", "금융·보험회사 명단",
        "비금융\n회사수", "금융\n회사수", "기업집단명", "순위", "전체소속회사수",
        "회사명", "기업명",
    }

    def _names_from_df(df) -> set:
        names: set = set()
        ncols = len(df.columns)
        # 별첨5 형식 판별: 9개 컬럼 이상이면 공정위 표 구조로 간주
        col_indices = [5, 7] if ncols >= 8 else [0]
        for _, row in df.iterrows():
            for c_idx in col_indices:
                if c_idx >= ncols:
                    continue
                v = str(row.iloc[c_idx]).strip()
                if v and v.lower() not in ("nan", "") and v not in SKIP_HEADERS:
                    norm = _normalize_corp_name(v)
                    if len(norm) >= 2:
                        names.add(norm)
        return names

    def _extract_names_from_xlsx(raw_bytes: bytes) -> set:
        try:
            df = _pd.read_excel(
                _io.BytesIO(raw_bytes), engine="openpyxl",
                dtype=str, header=None
            ).fillna("")
        except Exception:
            return set()
        return _names_from_df(df)

    result: set = set()
    path = Path(path)
    if not path.exists():
        return result

    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        result = _extract_names_from_xlsx(path.read_bytes())

    elif suffix == ".zip":
        try:
            with _zf.ZipFile(path, "r") as z:
                for raw_name in z.namelist():
                    try:
                        decoded = raw_name.encode("cp437").decode("euc-kr")
                    except Exception:
                        decoded = raw_name
                    if "5" in decoded and "계열회사" in decoded and decoded.endswith(".xlsx"):
                        result = _extract_names_from_xlsx(z.read(raw_name))
                        break
        except Exception:
            pass

    return result


# KT 세금코드명 → 플랫폼 과세구분 명시적 매핑
# (부분문자열 매칭으로 처리되지 않는 코드값 직접 지정)
# ※ KT에서 '영세율' 표현 미사용 — 별도 매핑 없음
_KT_TAX_NORM_MAP: dict = {
    "매입-계산서-일반(0%)": "면세",   # 플랫폼 '면세'에 해당 (영세율과 혼동 주의)
    "기타-일반":            "비과세",  # 플랫폼 '비과세'에 해당
}
# ※ 면세 키워드 추가 시 주의사항
#    '과자' → 가공식품(부가세 과세 대상) — 면세 키워드 추가 금지
#    '세제' → 공업용품(부가세 과세 대상) — 면세 키워드 추가 금지

def _tax_kw_match(text: str, keywords: list):
    """상품명(name) 필드용 — 한글 양쪽 경계 엄격 매칭.
    키워드 앞뒤에 한글(가-힣)이 오면 제외.

    예) '보리' 검색: '아이보리' → 앞이 한글 '이' → ✗  /  '보리 선물세트' → ✓
    예) '밀' 검색 : '밀폐용기'  → 뒤가 한글 '폐' → ✗  /  '밀 1kg'       → ✓
    예) '전복' 검색: '광전복합분전함' → 앞이 한글 '광' → ✗  / '전복 구이' → ✓
    """
    import re
    if not isinstance(text, str):
        return None, False
    for kw in keywords:
        pattern = r'(?<![가-힣])' + re.escape(kw) + r'(?![가-힣])'
        if re.search(pattern, text.strip()):
            return kw, True
    return None, False

def _tax_kw_match_spec(text: str, keywords: list):
    """규격(spec) 필드용 — 오른쪽 한글 경계만 체크 (왼쪽은 허용).
    복합어 식품명에서 성분 키워드 탐지:
      '밥다시마300g' → '다시마' 뒤가 숫자 → ✓
      '쌀톳200g'     → '톳' 뒤가 숫자   → ✓
      '밀폐용기'     → '밀' 뒤가 '폐'(한글) → ✗ (오른쪽 차단)
      '전복구이'     → '전복' 뒤가 '구'(한글) → ✗ (오른쪽 차단)
    """
    import re
    if not isinstance(text, str):
        return None, False
    for kw in keywords:
        pattern = re.escape(kw) + r'(?![가-힣])'
        if re.search(pattern, text.strip()):
            return kw, True
    return None, False

# ── 2. 스타일 헬퍼 ───────────────────────────────────────────────

def _bdr():
    s = Side(style="thin", color="C0C0C0")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr_cell(cell, color=HDR_BLUE):
    cell.font      = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", start_color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _bdr()

def data_cell(cell, fill=None, money=False, bold=False):
    cell.font      = Font(name=FONT, size=10, bold=bold)
    cell.border    = _bdr()
    cell.alignment = Alignment(horizontal="right" if money else "center",
                               vertical="center")
    if fill:
        cell.fill  = PatternFill("solid", start_color=fill)
    if money:
        cell.number_format = MONEY_FMT

def total_cell(cell, money=False):
    data_cell(cell, fill=ROW_TOTAL, money=money, bold=True)

def col_letter(ws, col_name):
    """워크시트에서 헤더명으로 열 문자 반환"""
    for c in ws.iter_cols(1, ws.max_column, 1, 1):
        if c[0].value == col_name:
            return get_column_letter(c[0].column)
    return None

def auto_col_width(ws, min_w=10, max_w=30):
    for col in ws.columns:
        length = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(length + 2, min_w), max_w)

# ── 3. 컬럼 탐지 헬퍼 ───────────────────────────────────────────

def _read_excel_smart(path, search_cols=None, max_skip=10, **kwargs):
    """상단 제목/날짜 행을 건너뛰고 실제 헤더 행을 자동 탐지하여 DataFrame 반환.
    search_cols 중 하나라도 있는 행을 헤더로 사용한다.
    """
    if search_cols is None:
        search_cols = ["주문번호", "구매문서번호", "일정산번호"]
    for skip in range(max_skip):
        try:
            probe = pd.read_excel(path, header=skip, nrows=0, engine="openpyxl")
            if any(c in probe.columns for c in search_cols):
                return pd.read_excel(path, header=skip, engine="openpyxl", **kwargs)
        except Exception:
            pass
    return pd.read_excel(path, engine="openpyxl", **kwargs)

def detect_col(df: pd.DataFrame, candidates: List[str],
               required=True) -> Optional[str]:
    # Phase 1: 정확히 일치
    for c in candidates:
        if c in df.columns:
            return c
    # Phase 2: 포함 매칭 (fallback)
    # ※ 후보 문자열 바로 뒤에 한글 자모(가-힣)가 이어지면 더 긴 복합어의 앞부분이므로 제외.
    #   예) "담당자" 후보 → "담당자부서" 컬럼: 뒤에 '부'(한글) → 매칭 안 함
    #       "관리회계" 후보 → "관리회계(IP만변경적용)": 뒤에 '(' → 매칭 허용
    for col in df.columns:
        for cand in candidates:
            idx = col.find(cand)
            if idx < 0:
                continue
            after = col[idx + len(cand):]
            if after and '가' <= after[0] <= '힣':
                continue   # 뒤에 한글이 이어짐 → 복합어 오매핑 방지
            return col
    if required:
        raise ValueError(f"필수 컬럼 미발견: {candidates}")
    return None

def build_col_map(df: pd.DataFrame, spec: dict) -> Dict[str, Optional[str]]:
    result = {}
    _optional = {"관리회계", "담당자", "일반통신구분", "서비스카테고리", "요청번호",
                 "구매품목", "품목번호", "주문품목키",
                 "매입금액", "매출총이익", "중분류", "담당자부서", "결제유형", "이동유형",
                 "담당자", "일반통신구분"}
    for key, cands in spec.items():
        result[key] = detect_col(df, cands, required=(key not in _optional))
    return result

# ── 4. 전용 예외 ──────────────────────────────────────────────────

class ReturnMappingNeeded(Exception):
    """반품 매핑 템플릿이 생성됐으며 사용자 입력이 필요함을 알리는 예외.
    GUI 워커의 except Exception 에 잡혀 메시지를 표시한다.
    sys.exit() 대신 사용하여 GUI 스레드가 정상 종료되도록 한다.
    """
    pass

# ── 5. MappingMaster ─────────────────────────────────────────────

class MappingMaster:
    def __init__(self, path: Optional[Path], midcorp_path: Optional[Path] = None):
        self.path         = path
        self.midcorp_path = midcorp_path
        self.svccat    : Dict[str, Tuple[str, str, str]] = {}  # 통합 svccat (하위 호환)
        self.svccat_new: Dict[str, Tuple[str, str, str]] = {}  # 현재 시트 (우선 참조)
        self.svccat_old: Dict[str, Tuple[str, str, str]] = {}  # 과거 시트 (폴백)
        self.person_type: Dict[str, str]             = {}  # {담당자: 일반/통신}
        self.corp_size  : Dict[str, str]             = {}  # {협력사명: 규모}  ← 대기업 수동 등록
        self.midcorp_set: set                        = set()  # 정규화된 중견기업 기업명 집합
        self.tax_rule   : Dict[str, str]             = {}  # {서비스카테고리: 기대과세구분}
        self.loaded     = False

    def get_corp_size(self, corp_name: str) -> str:
        """협력사명 → 기업규모 판별.
        우선순위: ① mapping_master 수동 등록 → ② 대기업 집합 자동 매칭
                  → ③ 중견기업목록 자동 매칭 → ④ 확인필요
        """
        # ① 수동 등록 (mapping_master.xlsx 기업규모 시트 직접 지정)
        v = self.corp_size.get(corp_name, "")
        if v:
            return v
        norm = _normalize_corp_name(corp_name)
        # ② 대기업 집합 자동 매칭 (공정거래위원회 공시대상기업집단 기준)
        if norm in _LARGE_CORP_SET:
            return "대기업"
        # ③ 중견기업목록 정규화 매칭 (한국중견기업연합회 확인서 발급 목록)
        if self.midcorp_set and norm in self.midcorp_set:
            return "중견기업"
        return "확인필요"

    def load(self, log):
        # ── 중견기업목록.xlsx 로드 — mapping_master 유무와 독립적으로 항상 실행 ──
        if self.midcorp_path and self.midcorp_path.exists():
            try:
                df_mid = pd.read_excel(self.midcorp_path, dtype=str, engine="openpyxl").fillna("")
                name_col_m = next((c for c in df_mid.columns if "기업명" in c), None)
                if name_col_m:
                    self.midcorp_set = {
                        _normalize_corp_name(n)
                        for n in df_mid[name_col_m].dropna()
                        if str(n).strip()
                    }
                    log.info(f"  중견기업목록 로드: {len(self.midcorp_set)}개 기업 "
                             f"(파일: {self.midcorp_path.name})")
                else:
                    log.warning("  중견기업목록: '기업명' 컬럼 미발견 → 로드 건너뜀")
            except Exception as _e:
                log.warning(f"  중견기업목록 로드 실패: {_e}")

        if not self.path or not self.path.exists():
            log.warning("  매핑파일 없음 → 관리회계·담당자·담당자부서·일반통신구분·기업규모(수동) 공란 처리")
            return
        try:
            # .xls 파일이 .xlsx 확장자로 저장된 경우 자동 감지
            _eng = SettlementRunner._xls_engine(self.path)
            try:
                all_sheets = pd.read_excel(self.path, sheet_name=None, dtype=str, engine=_eng)
            except Exception:
                # 확장자와 실제 형식이 다를 경우 반대 엔진으로 재시도
                _fallback = "xlrd" if _eng == "openpyxl" else "openpyxl"
                try:
                    all_sheets = pd.read_excel(self.path, sheet_name=None, dtype=str, engine=_fallback)
                    log.warning(f"  매핑파일 엔진 재시도({_fallback}) 성공: {self.path.name}")
                except Exception as _e2:
                    raise _e2
            sheet_dict = {k: df.fillna("") for k, df in all_sheets.items()}
            log.info(f"  매핑파일 시트: {list(sheet_dict.keys())}")

            # ── 내부 헬퍼: 카테고리 시트 → svccat dict ─────────────
            def _load_svccat(df_s, label):
                c = df_s.columns.tolist()
                # 키 컬럼 탐지 우선순위:
                #   1순위: "서비스" + "카테고리" 포함 (서비스카테고리, 서비스 카테고리 등)
                #   2순위: "카테고리" 포함 컬럼 중 "마스터" 제외
                #          ("카테고리 분류" 등 개편전 시트용 컬럼 허용)
                #   ※ "마스터카테고리(CMS)"는 서비스카테고리와 값 체계가 달라 명시적으로 제외
                cat_c = next((col for col in c if "서비스" in col and "카테고리" in col), None)
                if cat_c is None:
                    cat_c = next(
                        (col for col in c if "카테고리" in col and "마스터" not in col), None
                    )
                if cat_c is None:
                    log.warning(f"  [{label}] 카테고리 컬럼 미발견 → 매핑 건너뜀 (컬럼 목록: {c})")
                    return {}
                acc_c  = next((col for col in c if "관리회계" in col), None)
                dept_c = next((col for col in c if "담당부서" in col), None)
                mgr_c  = next((col for col in c if "담당자" in col and "부서" not in col), None)
                log.info(f"  [{label}] cat={cat_c}, acc={acc_c}, dept={dept_c}, mgr={mgr_c}")
                result = {}
                for _, r in df_s.iterrows():
                    k_raw = str(r[cat_c]).strip()
                    if k_raw and k_raw.lower() not in ("nan", ""):
                        k = _norm_svc_key(k_raw)   # '>' 주변 공백 정규화
                        result[k] = (
                            str(r[acc_c]).strip()  if acc_c  else "",
                            str(r[mgr_c]).strip()  if mgr_c  else "",
                            str(r[dept_c]).strip() if dept_c else "",
                        )
                return result

            # ── 시트: 서비스카테고리 → 관리회계/담당자/담당자부서 ──
            # 시트명으로 현재/과거 시트를 구분하여 로드.
            # 현재 시트: "이후", "현재", "신규", "251229" 등 키워드
            # 과거 시트: "이전", "개편전", "과거", "구" 등 키워드
            # 두 시트를 합산하되 현재 시트가 우선 → 현재에 없는 구 카테고리는 과거 시트로 폴백
            _KW_NEW = {"이후", "현재", "신규"}
            _KW_OLD = {"이전", "개편전", "과거", "구"}
            _new_name = next(
                (n for n in sheet_dict
                 if any(k in n for k in _KW_NEW) or "251229" in n), None
            )
            _old_name = next(
                (n for n in sheet_dict
                 if any(k in n for k in _KW_OLD) and n != _new_name), None
            )

            if _new_name or _old_name:
                svccat_old = _load_svccat(sheet_dict[_old_name], f"카테고리(과거: '{_old_name}')") \
                             if _old_name else {}
                svccat_new = _load_svccat(sheet_dict[_new_name], f"카테고리(현재: '{_new_name}')") \
                             if _new_name else {}
                # 인스턴스 속성으로 개별 보관 → apply() 에서 2-tier 조회 (현재 우선, 과거 폴백)
                self.svccat_new = svccat_new
                self.svccat_old = svccat_old
                # 통합 dict (하위 호환): 과거 먼저, 현재로 덮어씀 → 현재 시트가 항상 우선
                self.svccat = {**svccat_old, **svccat_new}
                # 과거 시트에는 있지만 현재 시트에 없는 키 → 과거 폴백으로만 사용됨
                _only_old = set(svccat_old) - set(svccat_new)
                _only_new = set(svccat_new) - set(svccat_old)
                _both     = set(svccat_old) & set(svccat_new)
                log.info(
                    f"  서비스카테고리: 현재 {len(svccat_new)}건 + 과거(폴백) {len(svccat_old)}건"
                    f" = 통합 {len(self.svccat)}건"
                    f"  (현재전용 {len(_only_new)}, 과거폴백전용 {len(_only_old)}, 공통 {len(_both)})"
                )
                # ── 진단: "유무선전송" 포함 키 추적 (관리회계 오분류 원인 파악용) ──
                _diag_kw = "유무선전송"
                _diag_found = False
                for _dk in sorted(k for k in svccat_new if _diag_kw in k):
                    _diag_found = True
                    log.info(f"  [진단] svccat_new 키: {repr(_dk)} → 관리회계={svccat_new[_dk][0]}")
                for _dk in sorted(k for k in svccat_old if _diag_kw in k):
                    _diag_found = True
                    _new_val = svccat_new.get(_dk)
                    _src = "현재시트덮어씀" if _new_val else "과거폴백(활성)"
                    log.info(f"  [진단] svccat_old 키: {repr(_dk)} → 관리회계={svccat_old[_dk][0]}  [{_src}]")
                if not _diag_found:
                    log.warning(f"  [진단] '{_diag_kw}' 포함 키가 현재·과거 시트 어디에도 없음")
                # 정규화된 문제 키 직접 조회
                _probe = _norm_svc_key("통신장비 > 유무선전송 > 광네트워크장치 > 집선스위치")
                log.info(f"  [진단] 문제 키 정규화 결과: {repr(_probe)}")
                log.info(f"  [진단] svccat_new 직접조회: {svccat_new.get(_probe, '(없음)')}")
                log.info(f"  [진단] svccat_old 직접조회: {svccat_old.get(_probe, '(없음)')}")

                # ── person_type 자동 구축 (현재 시트 mgr+dept → 일반/통신구분) ──
                # 과거 시트에 담당부서 컬럼이 없어도, 현재 시트에 등록된
                # 담당자명 → 담당부서 정보로 person_type 딕셔너리를 채운다.
                # → 과거 시트 카테고리 행(담당부서 미기재)의 일반/통신구분 판별에 사용
                def _dept_to_type(dept_str):
                    t = str(dept_str).strip()
                    if not t or t in ("nan", "None"): return ""
                    return "통신" if "네트워크" in t else "일반"

                for _sv in svccat_new.values():
                    _mgr_v, _dept_v = _sv[1], _sv[2]
                    if _mgr_v and _mgr_v not in ("nan", "None", "") and _dept_v:
                        _tp = _dept_to_type(_dept_v)
                        if _tp:
                            self.person_type.setdefault(_mgr_v, _tp)
                log.info(
                    f"  담당자→일반통신구분 자동구축: {len(self.person_type)}건 "
                    f"(현재 시트 담당자+담당부서 기반)"
                )
            else:
                # 시트명 미탐지 → 전체 시트에서 서비스카테고리/관리회계 컬럼 있는 것 모두 로드
                log.info("  서비스카테고리: 시트명 키워드 미탐지 → 컬럼 기반 자동 탐지")
                self.svccat = {}
                for _sname, _sdf in sheet_dict.items():
                    _cols = _sdf.columns.tolist()
                    if any("서비스" in c and "카테고리" in c for c in _cols) \
                            or any("관리회계" in c for c in _cols):
                        _loaded = _load_svccat(_sdf, f"카테고리(자동: '{_sname}')")
                        # 나중에 로드한 시트가 덮어쓰므로, 먼저 나온 시트(=현재) 우선 유지를 위해
                        # 이미 있는 키는 덮지 않음
                        for k, v in _loaded.items():
                            self.svccat.setdefault(k, v)
                        log.info(f"  카테고리 시트 '{_sname}': {len(_loaded)}건")
                self.svccat_new = self.svccat   # 자동탐지 시: 전체를 현재로 취급
                self.svccat_old = {}
                log.info(f"  서비스카테고리 통합(자동): {len(self.svccat)}건")

            # ── 시트: 일반/통신구분 (담당자 → 일반/통신) ─────────────
            _cat_sheet_names = {_new_name, _old_name} - {None}
            _type_name = next((n for n in sheet_dict
                               if "통신" in n or ("담당자" in n and "카테고리" not in n)), None)
            df_type = sheet_dict.get(_type_name) if _type_name else None
            if df_type is None:
                # 폴백: 카테고리 시트 및 과세 시트를 제외한 첫 번째 시트
                _non_cat = [v for k, v in sheet_dict.items()
                            if k not in _cat_sheet_names and "과세" not in k]
                df_type = _non_cat[0] if _non_cat else None
            if df_type is not None:
                c2 = df_type.columns.tolist()
                per_c  = next((c for c in c2 if "담당자" in c), c2[0] if c2 else None)
                type_c = next((c for c in c2 if "구분" in c or "통신" in c),
                              c2[1] if len(c2) > 1 else None)
                if per_c:
                    for _, r in df_type.iterrows():
                        k = str(r[per_c]).strip()
                        if k and type_c:
                            self.person_type[k] = str(r[type_c]).strip()
                    log.info(f"  일반통신구분 시트 '{_type_name}': {len(self.person_type)}건")

            # ── 시트: 기업규모 ─────────────────────────────────────
            _corp_name = next((n for n in sheet_dict if "규모" in n or "기업" in n), None)
            df_corp = sheet_dict.get(_corp_name) if _corp_name else None
            if df_corp is not None:
                c3 = df_corp.columns.tolist()
                corp_c = next((c for c in c3 if "협력사" in c or "기업명" in c or "업체" in c),
                              c3[0] if c3 else None)
                size_c = next((c for c in c3 if "규모" in c), c3[1] if len(c3) > 1 else None)
                if corp_c:
                    for _, r in df_corp.iterrows():
                        k = str(r[corp_c]).strip()
                        if k and size_c:
                            self.corp_size[k] = str(r[size_c]).strip()
                    log.info(f"  기업규모 시트 '{_corp_name}': {len(self.corp_size)}건")

            # ── 시트: 과세구분기준 (이름에 '과세' 포함, 선택적) ───────
            for sh_name, df_tax in sheet_dict.items():
                if "과세" in sh_name:
                    tc = df_tax.columns.tolist()
                    svc_c2 = next((c for c in tc if "카테고리" in c or "서비스" in c),
                                  tc[0] if tc else None)
                    exp_c  = next((c for c in tc if "과세" in c or "구분" in c or "기대" in c),
                                  tc[1] if len(tc) > 1 else None)
                    if svc_c2 and exp_c:
                        for _, r in df_tax.iterrows():
                            k = str(r[svc_c2]).strip()
                            v = str(r[exp_c]).strip()
                            if k and v:
                                self.tax_rule[k] = v
                    log.info(f"  과세구분기준 시트 '{sh_name}': {len(self.tax_rule)}건")
                    break

            self.loaded = True
            log.info(f"  매핑파일 로드 완료: 서비스카테고리 {len(self.svccat)}건 / "
                     f"일반통신구분 {len(self.person_type)}건 / 기업규모(수동) {len(self.corp_size)}건 / "
                     f"중견기업(자동) {len(self.midcorp_set)}개")
        except Exception as e:
            log.warning(f"  매핑파일 로드 실패: {e}")

    def apply(self, df: pd.DataFrame, merged_cols: Dict[str, Optional[str]], log) -> pd.DataFrame:
        """서비스카테고리 → 관리회계/담당자/담당자부서/중분류/일반통신구분 채우기 (기존값 우선, 벡터 연산)"""
        svc_col  = merged_cols.get("서비스카테고리")
        acc_col  = merged_cols.get("관리회계")     or "관리회계"
        mgr_col  = merged_cols.get("담당자")       or "담당자"
        type_col = merged_cols.get("일반통신구분") or "일반/통신구분"
        dept_col = merged_cols.get("담당자부서")   or "담당자부서"
        mdiv_col = merged_cols.get("중분류")       or "중분류(IP만변경적용)"

        # 컬럼 생성 + 문자열 dtype 강제
        for col in [acc_col, mgr_col, type_col, dept_col, mdiv_col]:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].fillna("").astype(str)

        def _blank(series):
            return series.astype(str).str.strip().isin(["", "None", "nan"])

        fail_cats = set()
        stats = {"관리회계": [0, 0], "담당자": [0, 0], "일반통신구분": [0, 0]}

        # ── 서비스카테고리 → 관리회계 / 담당자 / 담당자부서 (벡터) ──
        # 2-tier 조회: 현재 시트(svccat_new) 우선, 미등록 시 과거 시트(svccat_old) 폴백
        # 키는 _norm_svc_key 로 정규화 ('>' 전후 공백 흡수) — 매핑파일·플랫폼 간 표기 차이 극복
        _svc_diag_logged: set = set()   # 중복 진단 로그 방지
        _DIAG_KEY = _norm_svc_key("통신장비 > 유무선전송 > 광네트워크장치 > 집선스위치")
        def _svc_get(x: str, idx: int) -> str:
            norm = _norm_svc_key(x)
            v_new = self.svccat_new.get(norm)
            v_old = self.svccat_old.get(norm)
            v = v_new or v_old
            # 진단: ① 문제 키는 무조건 기록 ② 기타 처음 등장 키 최대 20건
            _force = (norm == _DIAG_KEY)
            if norm and (_force or (norm not in _svc_diag_logged and len(_svc_diag_logged) < 20)):
                if not _force:
                    _svc_diag_logged.add(norm)
                _src = ("현재시트" if v_new else ("과거폴백" if v_old else "미등록"))
                log.debug(
                    f"  [svc조회] {repr(norm)} → "
                    f"new={v_new[0] if v_new else '-'}  old={v_old[0] if v_old else '-'}  "
                    f"결과={v[0] if v else '(없음)'}  출처={_src}"
                )
            return v[idx] if v else ""

        if svc_col and svc_col in df.columns:
            df["_svc"] = df[svc_col].astype(str).str.strip()

            # ── 마스터카테고리 예외 처리 ─────────────────────────────────────
            # _MCAT_OVERRIDE_SET 에 등록된 마스터카테고리 행은
            # 서비스카테고리 대신 마스터카테고리(CMS) 값으로 조회한다.
            _mcat_col = next(
                (c for c in df.columns if "마스터카테고리" in c), None
            )
            if _mcat_col and _MCAT_OVERRIDE_SET:
                _mcat_norm = df[_mcat_col].astype(str).str.strip().map(_norm_svc_key)
                _exc_mask  = _mcat_norm.isin(_MCAT_OVERRIDE_SET)
                if _exc_mask.any():
                    # 예외 행의 조회 키를 마스터카테고리 원본값으로 교체
                    df.loc[_exc_mask, "_svc"] = (
                        df.loc[_exc_mask, _mcat_col].astype(str).str.strip()
                    )
                    log.info(
                        f"  마스터카테고리 예외 적용: {int(_exc_mask.sum())}건 "
                        f"({_mcat_col} 기준 매핑)"
                    )
            df["_acc_mapped"]  = df["_svc"].map(lambda x: _svc_get(x, 0))
            df["_mgr_mapped"]  = df["_svc"].map(lambda x: _svc_get(x, 1))
            df["_dept_mapped"] = df["_svc"].map(lambda x: _svc_get(x, 2))

            # 서비스카테고리 매핑값이 있으면 기존값보다 우선 적용
            # (서비스카테고리가 가장 정확한 기준이므로 기존 KT/플랫폼 원본값 덮어쓰기)
            _has_mapped = df["_acc_mapped"].str.strip().ne("")
            df.loc[_has_mapped, acc_col]  = df.loc[_has_mapped, "_acc_mapped"]
            df.loc[_has_mapped, mgr_col]  = df.loc[_has_mapped, "_mgr_mapped"]
            df.loc[_has_mapped, dept_col] = df.loc[_has_mapped, "_dept_mapped"]

            # 매핑값이 없는 행(=서비스카테고리 미등록)은 기존값 유지 → 빈 경우에만 오류 표기
            _has_dept_mapped = df["_dept_mapped"].str.strip().ne("")
            dept_blank = _blank(df[dept_col])
            df.loc[dept_blank & _has_dept_mapped, dept_col] = df.loc[dept_blank & _has_dept_mapped, "_dept_mapped"]

            # 중분류(IP만변경적용) = 관리회계와 동일값 (서비스카테고리 매핑 적용 후)
            mdiv_blank = _blank(df[mdiv_col])
            df.loc[mdiv_blank, mdiv_col] = df.loc[mdiv_blank, acc_col]

            # 통계
            mapped   = _has_mapped
            unmapped = ~mapped
            stats["관리회계"] = [int(mapped.sum()), int(unmapped.sum())]
            stats["담당자"]   = [int(mapped.sum()), int(unmapped.sum())]
            fail_cats = set(df.loc[unmapped & ~_blank(df["_svc"]), "_svc"].tolist())

            # 관리회계 빈칸(매핑 실패) → 오류 사유 기재 (추적 가능)
            acc_blank_after = _blank(df[acc_col])
            if acc_blank_after.any():
                def _acc_reason(svc):
                    s = str(svc).strip()
                    if not s or s in ("nan", "None"):
                        return "[오류] 서비스카테고리 미기재"
                    return f"[오류] 매핑 미등록: {s}"
                df.loc[acc_blank_after, acc_col] = df.loc[acc_blank_after, "_svc"].map(_acc_reason)

            df.drop(columns=["_svc", "_acc_mapped", "_mgr_mapped", "_dept_mapped"], inplace=True)

        # ── 일반/통신구분 결정 ─────────────────────────────────────
        # 우선순위: ① 기존 값 유지 → ② 서비스카테고리별 매핑파일 담당부서(안)
        #           → ③ 담당자부서 컬럼 직접 참조
        # 네트워크사업1팀 계열 → 통신 / 전략구매사업1팀 계열 → 일반
        def _team_to_type(team: str) -> str:
            t = str(team).strip()
            if not t or t in ("nan", "None"):
                return ""
            return "통신" if "네트워크" in t else "일반"

        type_blank = _blank(df[type_col])

        # ② 서비스카테고리 → 매핑파일 담당부서(안) → 통신/일반
        if svc_col and svc_col in df.columns and type_blank.any():
            mapped_dept = df.loc[type_blank, svc_col].astype(str).str.strip().map(
                lambda x: _svc_get(x, 2)
            )
            df.loc[type_blank, type_col] = mapped_dept.map(_team_to_type)

        # ③ 아직 빈 경우 담당자부서 컬럼 직접 참조
        still_blank = _blank(df[type_col])
        if dept_col in df.columns and still_blank.any():
            df.loc[still_blank, type_col] = df.loc[still_blank, dept_col].map(_team_to_type)

        # ④ 담당자 이름 → person_type 딕셔너리 직접 조회
        # 과거 시트에 담당부서 컬럼이 없어 ②③에서 해결 못한 행(주로 과거 카테고리 건)에 적용
        still_blank2 = _blank(df[type_col])
        if self.person_type and mgr_col in df.columns and still_blank2.any():
            df.loc[still_blank2, type_col] = (
                df.loc[still_blank2, mgr_col]
                .astype(str).str.strip()
                .map(lambda m: self.person_type.get(m, ""))
            )
            _fixed4 = int(still_blank2.sum()) - int(_blank(df[type_col]).sum())
            if _fixed4 > 0:
                log.info(f"  일반통신구분 ④ person_type 보정: {_fixed4}건")

        ok_type = ~_blank(df[type_col])
        stats["일반통신구분"] = [int(ok_type.sum()), int((~ok_type).sum())]

        log.info(f"  관리회계 매핑  : {stats['관리회계'][0]}건 성공 / {stats['관리회계'][1]}건 실패")
        log.info(f"  담당자 매핑    : {stats['담당자'][0]}건 성공 / {stats['담당자'][1]}건 실패")
        log.info(f"  일반통신구분   : {stats['일반통신구분'][0]}건 성공 / {stats['일반통신구분'][1]}건 실패")
        if fail_cats:
            log.warning(f"  매핑 실패 서비스카테고리: {sorted(fail_cats)}")
        return df, stats

# ── 5. SettlementRunner ──────────────────────────────────────────

class SettlementRunner:

    def __init__(self, base_dir="."):
        self.base       = Path(base_dir)
        self.today      = datetime.date.today()
        self.log        = self._setup_logger()
        self.kt_path    : Optional[Path] = None
        self.pl_path    : Optional[Path] = None
        self.rm_path    : Optional[Path] = None   # return_mapping.xlsx
        self.master     : Optional[MappingMaster] = None
        # 정산 기간
        self.ps         : Optional[datetime.date] = None
        self.pe         : Optional[datetime.date] = None
        # 컬럼 매핑
        self.kt_cols    : Dict[str, Optional[str]] = {}
        self.pl_cols    : Dict[str, Optional[str]] = {}
        # 처리 결과
        self.df_kt      = None   # KT 기간 필터 전체
        self.df_pl      = None
        self.df_normal  = None   # KT 일반
        self.df_ret_kt  = None   # KT 반품
        self.df_matched = None
        self.df_missing = None   # 플랫폼 누락
        self.df_pl_only = None
        self.df_amtdiff = None
        self.df_ret_ok  = None   # 반품 매칭 성공
        self.df_ret_ng  = None   # 반품 미확인
        self.ret_route  = None   # 'A' or 'B'
        self.df_invoice    = None
        self.map_stats     = {}
        self.composite_key = False   # True = 주문번호+품목번호 복합 키 사용

    def _setup_logger(self):
        log = logging.getLogger("settlement")
        log.setLevel(logging.DEBUG)
        if not log.handlers:
            fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                    datefmt="%H:%M:%S")
            sh = logging.StreamHandler(sys.stdout); sh.setFormatter(fmt)
            log_path = self.base / "settlement_log.txt"
            fh = logging.FileHandler(str(log_path), encoding="utf-8")
            fh.setFormatter(fmt)
            log.addHandler(sh); log.addHandler(fh)
        return log

    # ── 파일 탐지 ───────────────────────────────────────────────
    def detect_files(self):
        self.log.info("=" * 60)
        self.log.info("  KT 정산 자동화 시스템  v2.0")
        self.log.info("=" * 60)

        def _find(*names):
            for n in names:
                p = self.base / n
                if p.exists():
                    return p
            return None

        self.kt_path = _find("kt_raw.xlsx", "KT_raw.xlsx", "kt_raw.xlsm")
        self.pl_path = _find("platform.xlsx", "플랫폼.xlsx", "통합플랫폼.xlsx")
        self.rm_path = _find("return_mapping.xlsx")
        mm_path      = _find("mapping_master.xlsx", "매핑마스터.xlsx")
        midcorp_path = _find("중견기업목록.xlsx")

        for label, path, req in [
            ("KT 파일        ", self.kt_path,  True),
            ("플랫폼 파일    ", self.pl_path,  True),
            ("mapping_master ", mm_path,        False),
            ("return_mapping ", self.rm_path,   False),
            ("중견기업목록   ", midcorp_path,   False),
        ]:
            self.log.info(f"  {label}: {path or '없음'}")
            if req and not path:
                fname = input(f"\n  '{label.strip()}' 파일명 입력: ").strip()
                if label.strip().startswith("KT"):
                    self.kt_path = self.base / fname
                else:
                    self.pl_path = self.base / fname

        self.master = MappingMaster(mm_path, midcorp_path)
        self.master.load(self.log)

        # ── 공정거래위원회 공시대상기업집단 계열회사 파일 자동 로드 ──
        # ① 정산 폴더의 대기업목록.xlsx 우선 (중견기업목록.xlsx 과 동일한 방식)
        # ② 없으면 정산 폴더 → C:\FTC_downloads\정산도구 → Downloads 에서
        #    "공시대상기업집단" ZIP 또는 "별첨5"/"계열회사현황" xlsx 탐색
        _ftc_candidate: Optional[Path] = _find("대기업목록.xlsx")   # ← 간편 파일명 우선

        if _ftc_candidate is None:
            _ftc_search_dirs = [
                self.base,
                Path(r"C:\FTC_downloads\정산도구"),
                Path.home() / "Downloads",
            ]
            for _d in _ftc_search_dirs:
                if not _d.exists():
                    continue
                for _f in sorted(_d.iterdir()):
                    _fn = _f.name
                    _is_zip = _f.suffix.lower() == ".zip" and "공시대상기업집단" in _fn
                    _is_xl5 = _f.suffix.lower() == ".xlsx" and (
                        "별첨5" in _fn or "계열회사현황" in _fn
                    )
                    if _is_zip or _is_xl5:
                        _ftc_candidate = _f
                        break
                if _ftc_candidate:
                    break

        self.log.info(f"  {'대기업목록      '}: {_ftc_candidate or '없음 (내장 목록 사용)'}")

        if _ftc_candidate:
            _ftc_new = _load_ftc_large_corps(_ftc_candidate)
            if _ftc_new:
                _before = len(_LARGE_CORP_SET)
                _LARGE_CORP_SET.update(_ftc_new)
                _added = len(_LARGE_CORP_SET) - _before
                self.log.info(
                    f"  [FTC] 계열회사 로드 완료: "
                    f"{len(_ftc_new)}개 사 ({_added}개 신규 추가)"
                )
            else:
                self.log.warning(f"  [FTC] 파일 파싱 실패: {_ftc_candidate.name}")

    # ── STEP 0: 컬럼 분석 + 확인 ──────────────────────────────
    def step0_init(self):
        self.log.info("\n[STEP 0] 파일 컬럼 구조 분석")
        df_kt = _read_excel_smart(self.kt_path, search_cols=["구매문서번호"], nrows=3, dtype=str)
        df_pl = _read_excel_smart(self.pl_path, search_cols=["주문번호", "일정산번호"], nrows=3, dtype=str)

        self.kt_cols = build_col_map(df_kt, _KT_CANDIDATES)
        self.pl_cols = build_col_map(df_pl, _PL_CANDIDATES)
        # 플랫폼 요청번호(optional) 별도 탐지
        self.pl_cols["요청번호"] = detect_col(df_pl, _PL_OPT["요청번호"], required=False)

        LINE = "-" * 55
        print(f"\n{LINE}")
        print("[KT 파일 컬럼 탐지 결과]")
        for key, found in self.kt_cols.items():
            mark = "O" if found else "-"
            print(f"  [{mark}] {key:<14} : {found or '(미탐지)'}")

        print(f"\n[플랫폼 파일 컬럼 탐지 결과]")
        for key, found in self.pl_cols.items():
            mark = "O" if found else "-"
            print(f"  [{mark}] {key:<14} : {found or '(없음)'}")

        req_col  = self.pl_cols.get("요청번호")
        kt_item  = self.kt_cols.get("구매품목")
        pl_item  = self.pl_cols.get("품목번호")
        pl_combo = self.pl_cols.get("주문품목키")
        if kt_item and pl_combo:
            key_mode = (f"복합 키  ({self.kt_cols['주문번호']}+{kt_item} (KT)  ↔  "
                        f"{pl_combo} (플랫폼, 사전계산)")
        elif kt_item and pl_item:
            key_mode = (f"복합 키  ({self.kt_cols['주문번호']}+{kt_item}  ↔  "
                        f"{self.pl_cols['주문번호']}+{pl_item})")
        else:
            key_mode = f"단순 키  ({self.kt_cols['주문번호']} ↔ {self.pl_cols['주문번호']})"
        print(f"\n▶ 매칭 키 모드  : {key_mode}")
        print(f"▶ 반품 처리 경로: {'A (요청번호 직접 매칭)' if req_col else 'B (return_mapping.xlsx 사용)'}")
        print(LINE)

        ans = input("위 설정으로 진행하시겠습니까? (Y/n): ").strip().lower()
        if ans == "n":
            sys.exit(0)

    # ── 정산 기간 ──────────────────────────────────────────────
    def get_period(self):
        d, y, m = self.today.day, self.today.year, self.today.month
        if d == 16:
            s, e = datetime.date(y, m, 1), datetime.date(y, m, 15)
        elif d == 1:
            py, pm = (y-1, 12) if m == 1 else (y, m-1)
            s = datetime.date(py, pm, 16)
            e = datetime.date(py, pm, calendar.monthrange(py, pm)[1])
        else:
            self.log.warning(f"  오늘({self.today})은 자동 판단 기준일이 아닙니다.")
            s = datetime.date.fromisoformat(input("  시작일 (YYYY-MM-DD): ").strip())
            e = datetime.date.fromisoformat(input("  종료일 (YYYY-MM-DD): ").strip())
        self.log.info(f"  정산 기간: {s} ~ {e}")
        return s, e

    # ── KT 전처리 ──────────────────────────────────────────────
    def _preprocess_kt(self, df: pd.DataFrame) -> pd.DataFrame:
        """① 합계 행 제거(구매문서번호 공란)  ② 세금코드명 → 과세/면세/비과세 정규화"""
        ord_col = self.kt_cols.get("주문번호")  # 구매문서번호
        if ord_col and ord_col in df.columns:
            before = len(df)
            df = df[df[ord_col].fillna("").astype(str).str.strip() != ""].copy()
            removed = before - len(df)
            if removed > 0:
                self.log.info(f"  합계/빈값 행 {removed}건 제거 ({ord_col} 공란)")

        tax_col = self.kt_cols.get("매출과세구분")  # 세금코드명
        if tax_col and tax_col in df.columns:
            def _norm(v):
                v = str(v).strip()
                if v in _KT_TAX_NORM_MAP:           # 명시적 코드 우선 매핑
                    return _KT_TAX_NORM_MAP[v]
                if "비과세" in v: return "비과세"
                if "면세"   in v: return "면세"
                if "과세"   in v: return "과세"
                if "세금계산서" in v: return "과세"  # 세금계산서 = 과세 거래 (10% VAT)
                return v
            before_vals = df[tax_col].dropna().unique().tolist()
            df[tax_col] = df[tax_col].map(_norm)
            after_vals  = df[tax_col].dropna().unique().tolist()
            self.log.info(f"  세금코드명 정규화 완료: {before_vals} → {after_vals}")
        return df

    # ── 복합 키 생성 ────────────────────────────────────────────
    @staticmethod
    def _item_str(series: pd.Series) -> pd.Series:
        """품목번호(Series)를 정수 문자열로 정규화: 10.0 → '10', '010' → '10', '' → ''"""
        def _conv(v):
            s = str(v).strip()
            if s in ("", "nan", "None"): return ""
            try:
                return str(int(float(s)))
            except (ValueError, TypeError):
                return s
        return series.apply(_conv)

    @staticmethod
    def _key_str(v) -> str:
        """주문&품목 단일 값(Excel float 또는 문자열) → 정수 or 원본 문자열 변환.
        예: 4.50201e+11 → '450200980415'   'ORD-2025-00110' → 'ORD-2025-00110'
        """
        s = str(v).strip()
        if s in ("", "nan", "None"):
            return ""
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            return s

    def _assign_match_keys(self):
        """KT와 플랫폼 DataFrame에 _match_key 컬럼 부여.
        - 플랫폼에 '주문&품목' 사전계산 키가 있으면 직접 사용 (_key_str 변환)
        - 없으면 주문번호+품목번호 연결 (구분자 없음)
        - KT는 항상 구매문서번호+구매품목 연결
        """
        kt_ord      = self.kt_cols["주문번호"]    # 구매문서번호
        pl_ord      = self.pl_cols["주문번호"]    # 주문번호
        kt_item     = self.kt_cols.get("구매품목")
        pl_combo_col = self.pl_cols.get("주문품목키")   # 주문&품목 (사전계산)
        pl_item     = self.pl_cols.get("품목번호")

        if kt_item and kt_item in self.df_kt.columns:
            # KT: 구매문서번호 + 구매품목
            kt_key = (self.df_kt[kt_ord].fillna("").astype(str)
                      + self._item_str(self.df_kt[kt_item]))
            self.composite_key = True

            # 플랫폼 키 생성
            # 주문&품목 = 주문번호+품목번호 연결값이므로,
            # 품목번호 컬럼이 있으면 직접 정규화하여 연결 (leading-zero 제거 통일)
            # 품목번호 컬럼 없으면 주문&품목 컬럼을 그대로 사용
            if pl_item and pl_item in self.df_pl.columns:
                pl_key = (self.df_pl[pl_ord].fillna("").astype(str)
                          + self._item_str(self.df_pl[pl_item]))
                self.log.info(f"  복합 키: {kt_ord}+{kt_item} (KT)  ↔  {pl_ord}+{pl_item} (플랫폼)")
            elif pl_combo_col and pl_combo_col in self.df_pl.columns:
                # 품목번호 컬럼 미존재 → 주문&품목 컬럼 직접 사용
                pl_key = self.df_pl[pl_combo_col].apply(self._key_str)
                self.log.info(f"  복합 키: {kt_ord}+{kt_item} (KT)  ↔  {pl_combo_col} (플랫폼, 품목번호 컬럼 없음)")
            else:
                # 폴백: 단순 키
                pl_key = self.df_pl[pl_ord].fillna("").astype(str)
                kt_key = self.df_kt[kt_ord].fillna("").astype(str)
                self.composite_key = False
                self.log.info(f"  단순 키 폴백: {kt_ord} ↔ {pl_ord} (품목번호 컬럼 없음)")

            self.df_kt["_match_key"] = kt_key
            self.df_pl["_match_key"] = pl_key
            # 반품 원주문 조회용 전체 플랫폼 데이터에도 동일 방식으로 _match_key 부여
            if pl_item and pl_item in self.df_pl_full.columns:
                self.df_pl_full["_match_key"] = (
                    self.df_pl_full[pl_ord].fillna("").astype(str)
                    + self._item_str(self.df_pl_full[pl_item])
                )
            elif pl_combo_col and pl_combo_col in self.df_pl_full.columns:
                self.df_pl_full["_match_key"] = self.df_pl_full[pl_combo_col].apply(self._key_str)
            else:
                self.df_pl_full["_match_key"] = self.df_pl_full[pl_ord].fillna("").astype(str)
        else:
            self.composite_key = False
            self.df_kt["_match_key"] = self.df_kt[kt_ord].fillna("").astype(str)
            self.df_pl["_match_key"] = self.df_pl[pl_ord].fillna("").astype(str)
            self.df_pl_full["_match_key"] = self.df_pl_full[pl_ord].fillna("").astype(str)
            self.log.info(f"  단순 키 사용: {kt_ord} ↔ {pl_ord}")

    # ── STEP 0b: 과세구분 검증 ────────────────────────────────
    def step0b_validate_tax(self):
        """단어 단위 키워드 기반 과세구분 검토 권고 (확정 오류 아님 — 담당자 검토용)."""
        self.log.info("\n[STEP 0b] 과세구분 검토 (키워드 기반)")
        self.tax_issues = []

        # ── 플랫폼 파일 로드 (코드성 컬럼은 str 강제) ──
        try:
            df_p = _read_excel_smart(self.pl_path,
                                     search_cols=["주문번호", "일정산번호"],
                                     dtype=str)
        except Exception as e:
            self.log.warning(f"  플랫폼 파일 로드 실패 → 검증 건너뜀: {e}")
            return

        # 코드성 컬럼 소수점 제거 (ex. 1234.0 → 1234)
        _code_cols = ["상품코드", "품목번호", "주문번호", "상품ID",
                      "카테고리ID", "부서코드", "사업장코드"]
        for col in _code_cols:
            if col in df_p.columns:
                df_p[col] = (df_p[col].astype(str)
                             .str.replace(r'\.0$', '', regex=True)
                             .str.strip())

        ord_c  = self.pl_cols.get("주문번호")
        tax_c  = self.pl_cols.get("매출과세구분")
        code_c = next((c for c in ["상품코드", "자재번호", "품목코드"] if c in df_p.columns), None)
        name_c = next((c for c in ["상품명", "자재명", "품목명"]       if c in df_p.columns), None)
        spec_c = next((c for c in ["대표규격", "규격", "사양"]          if c in df_p.columns), None)
        cat_c  = next((c for c in df_p.columns if "서비스카테고리" in c), None)

        if not tax_c or tax_c not in df_p.columns:
            self.log.warning("  매출과세구분 컬럼 없음 → 검증 건너뜀")
            return

        # KT 정산자료에 있는 주문번호만 검증 (KT에 없는 주문은 검토 불필요)
        if ord_c and ord_c in df_p.columns:
            try:
                kt_ord_c = self.kt_cols.get("주문번호")
                if kt_ord_c:
                    _df_kt_tmp = _read_excel_smart(self.kt_path,
                                                   search_cols=["구매문서번호"],
                                                   dtype={kt_ord_c: str})
                    _df_kt_tmp = self._preprocess_kt(_df_kt_tmp)
                    kt_ord_set = set(_df_kt_tmp[kt_ord_c].astype(str).str.strip())
                    _before = len(df_p)
                    df_p = df_p[df_p[ord_c].astype(str).str.strip().isin(kt_ord_set)].copy()
                    self.log.info(f"  KT 주문번호 기준 필터: {_before}건 → {len(df_p)}건 (KT에 없는 주문 제외)")
            except Exception as _e:
                self.log.warning(f"  KT 주문번호 필터 실패 → 전체 검증: {_e}")

        # 플랫폼 과세구분 값 정규화 (영세/영세율 → 면세)
        # KT 매입-계산서-일반(0%) = 플랫폼 면세에 해당 (영세율 표기 없음)
        _PL_TAX_NORM = {"영세": "면세", "영세율": "면세"}
        df_p[tax_c] = (df_p[tax_c].astype(str).str.strip()
                       .map(lambda v: _PL_TAX_NORM.get(v, v)))

        def _s(val) -> str:
            v = str(val).strip()
            return "" if v in ("nan", "None", "NaN") else v

        def _has_exclude(text: str) -> bool:
            """제외 키워드 부분 문자열 매칭 (복합어 대응).
            예: '세제' → '주방세제' 포함, '과자' → '과자류' 포함
            단어 경계 없이 단순 포함 여부만 검사하므로 오탐 위험 낮음.
            """
            if not isinstance(text, str):
                return False
            t = text.strip()
            return any(kw in t for kw in _TAX_EXCLUDE_KW)

        # ── 과세확인목록(화이트리스트) 로드 ──────────────────────────
        # KT 파일과 같은 폴더에 과세확인목록.xlsx 가 있으면 자동 로드
        _whitelist: set = set()
        _wl_path = self.base / "과세확인목록.xlsx"
        if _wl_path.exists():
            try:
                _wl_df = pd.read_excel(str(_wl_path), dtype=str, engine="openpyxl")
                _wl_code_c = next(
                    (c for c in _wl_df.columns if "상품코드" in c),
                    next((c for c in _wl_df.columns if "코드" in c), None),
                )
                if _wl_code_c:
                    for _v in _wl_df[_wl_code_c].dropna():
                        _cd = _re_corp.sub(r'\.0$', '', str(_v).strip())
                        if _cd and _cd.lower() not in ("nan", ""):
                            _whitelist.add(_cd)
                self.log.info(f"  과세확인목록 로드: {len(_whitelist)}건 ({_wl_path.name})")
            except Exception as _we:
                self.log.warning(f"  과세확인목록 로드 실패: {_we}")
        elif self.pl_path:
            _wl_path2 = Path(self.pl_path).parent / "과세확인목록.xlsx"
            if _wl_path2.exists() and _wl_path2 != _wl_path:
                try:
                    _wl_df = pd.read_excel(str(_wl_path2), dtype=str, engine="openpyxl")
                    _wl_code_c = next(
                        (c for c in _wl_df.columns if "상품코드" in c),
                        next((c for c in _wl_df.columns if "코드" in c), None),
                    )
                    if _wl_code_c:
                        for _v in _wl_df[_wl_code_c].dropna():
                            _cd = _re_corp.sub(r'\.0$', '', str(_v).strip())
                            if _cd and _cd.lower() not in ("nan", ""):
                                _whitelist.add(_cd)
                    self.log.info(f"  과세확인목록 로드: {len(_whitelist)}건 ({_wl_path2.name})")
                except Exception as _we:
                    self.log.warning(f"  과세확인목록 로드 실패: {_we}")

        cnt_a = cnt_b = cnt_c = cnt_wl = 0

        # ── 검증 A · B: 키워드 vs 현재 과세구분 ──
        for _, r in df_p.iterrows():
            name     = _s(r.get(name_c, "")) if name_c else ""
            spec     = _s(r.get(spec_c, "")) if spec_c else ""
            cur_tax  = _s(r.get(tax_c,  ""))
            combined = f"{name} {spec}"
            row_code = _s(r.get(code_c, "")) if code_c else ""
            # 담당자 조회 (서비스카테고리 → MappingMaster.svccat[1])
            _cat_val = _norm_svc_key(_s(r.get(cat_c, ""))) if cat_c else ""
            _mgr_val = (self.master.svccat.get(_cat_val, ("", "", ""))[1]
                        if (self.master and self.master.loaded and _cat_val) else "")

            # 과세확인목록 상품코드 → 검증 제외
            if _whitelist and row_code and row_code in _whitelist:
                cnt_wl += 1
                continue

            # 제외 키워드 먼저 체크
            if _has_exclude(combined):
                continue

            # A: 비과세 키워드 탐지
            matched_kw, hit = _tax_kw_match(combined, _TAX_NONTAX_KW)
            if hit and cur_tax != "비과세":
                self.tax_issues.append({
                    "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                    "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                    "상품명"      : name,
                    "대표규격"    : spec,
                    "담당자"      : _mgr_val,
                    "현재과세구분": cur_tax,
                    "검증결과"    : "비과세 검토 필요",
                    "탐지키워드"  : matched_kw,
                    "사유"        : (f"상품명/규격에 '{matched_kw}' 포함. "
                                     "부가세법상 비과세 대상일 수 있습니다. 담당자 확인 필요."),
                    "비고"        : "",
                })
                cnt_a += 1
                continue  # A 해당이면 B 중복 체크 불필요

            # B-1: 이력 기반 면세 패턴 탐지
            # 상품명: 양쪽 한글 경계 엄격 매칭
            # 규격: 상품명이 '선물세트' 등 컨테이너 키워드일 때만 오른쪽 경계 매칭
            # ※ 과세·비과세 모두 탐지
            known_kw, _ = _tax_kw_match(name, _KNOWN_EXEMPT_NAMES)
            if not known_kw and any(t in name for t in _EXEMPT_SPEC_TRIGGER_KW):
                known_kw, _ = _tax_kw_match_spec(spec, _KNOWN_EXEMPT_NAMES)
            if known_kw and cur_tax != "면세":
                _tax_note = ("비과세로 등록되어 있으나" if cur_tax == "비과세"
                             else "과세로 등록되어 있으나")
                self.tax_issues.append({
                    "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                    "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                    "상품명"      : name,
                    "대표규격"    : spec,
                    "담당자"      : _mgr_val,
                    "현재과세구분": cur_tax,
                    "검증결과"    : "면세 검토 필요",
                    "탐지키워드"  : known_kw,
                    "사유"        : (f"상품명/규격에 '{known_kw}' 포함 (과거 면세 처리 이력 패턴). "
                                     f"{_tax_note} 면세 대상일 수 있습니다. 담당자 확인 필요."),
                    "비고"        : "",
                })
                cnt_b += 1
                continue  # B-1 해당이면 B-2 중복 체크 불필요

            # B-2: 부가세법 일반 면세 키워드 탐지
            # 상품명: 양쪽 한글 경계 엄격 매칭
            # 규격: 상품명이 '선물세트' 등 컨테이너 키워드일 때만 오른쪽 경계 매칭
            # ※ 과세·비과세 모두 탐지
            matched_kw, _ = _tax_kw_match(name, _TAX_EXEMPT_KW)
            if not matched_kw and any(t in name for t in _EXEMPT_SPEC_TRIGGER_KW):
                matched_kw, _ = _tax_kw_match_spec(spec, _TAX_EXEMPT_KW)
            if matched_kw and cur_tax != "면세":
                _tax_note = ("비과세로 등록되어 있으나" if cur_tax == "비과세"
                             else "과세로 등록되어 있으나")
                self.tax_issues.append({
                    "주문번호"    : _s(r.get(ord_c,  "")) if ord_c  else "",
                    "상품코드"    : _s(r.get(code_c, "")) if code_c else "",
                    "상품명"      : name,
                    "대표규격"    : spec,
                    "담당자"      : _mgr_val,
                    "현재과세구분": cur_tax,
                    "검증결과"    : "면세 검토 필요",
                    "탐지키워드"  : matched_kw,
                    "사유"        : (f"상품명/규격에 '{matched_kw}' 포함. "
                                     f"{_tax_note} 부가세법상 면세 대상일 수 있습니다. 담당자 확인 필요."),
                    "비고"        : "",
                })
                cnt_b += 1

        # ── 검증 C: 동일 상품코드에 과세구분 혼재 ──
        if code_c and code_c in df_p.columns:
            valid = df_p[df_p[code_c].str.strip().ne("") & df_p[code_c].notna()].copy()
            grp = (valid.groupby(valid[code_c].str.strip())[tax_c]
                   .apply(lambda s: sorted(s.dropna().unique().tolist())))
            mixed = grp[grp.apply(lambda v: len(v) > 1)]
            already_flagged = {r["상품코드"] for r in self.tax_issues}
            for code, tax_vals in mixed.items():
                if code in already_flagged:
                    continue
                if _whitelist and code in _whitelist:
                    cnt_wl += 1
                    continue
                rows_c = df_p[df_p[code_c].str.strip() == code]
                tax_str = " / ".join(tax_vals)
                r0 = rows_c.iloc[0]
                ord_val  = _s(r0.get(ord_c,  "")) if ord_c  else ""
                name_val = _s(r0.get(name_c, "")) if name_c else ""
                spec_val = _s(r0.get(spec_c, "")) if spec_c else ""
                cat_val  = _norm_svc_key(_s(r0.get(cat_c, ""))) if cat_c else ""
                mgr_val  = (self.master.svccat.get(cat_val, ("", "", ""))[1]
                            if (self.master and self.master.loaded and cat_val) else "")
                n_rows   = len(rows_c)
                ord_label = f"{ord_val} 외 {n_rows-1}건" if n_rows > 1 else ord_val
                self.tax_issues.append({
                    "주문번호"    : ord_label,
                    "상품코드"    : code,
                    "상품명"      : name_val,
                    "대표규격"    : spec_val,
                    "담당자"      : mgr_val,
                    "현재과세구분": tax_str,
                    "검증결과"    : "과세구분 혼재",
                    "탐지키워드"  : "",
                    "사유"        : (f"동일 상품코드 {code}에 {tax_str} 혼재. "
                                     "상품 마스터 데이터 확인 필요."),
                    "비고"        : "",
                })
                cnt_c += 1

        # ── 상품코드 기준 중복 제거 (A·B 섹션) ──────────────────────
        # 동일 상품코드가 여러 주문번호에 걸쳐 탐지되면 한 행으로 합산
        if self.tax_issues:
            deduped: list = []
            code_idx: dict = {}   # 상품코드 → deduped 내 index
            for issue in self.tax_issues:
                code = str(issue.get("상품코드", "")).strip()
                if not code:                          # 코드 없는 행: 그대로 유지
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
                    # 담당자 병합 (중복 제거)
                    existing = deduped[idx].get("담당자", "")
                    new_mgr  = issue.get("담당자", "")
                    if new_mgr and new_mgr not in existing.split(", "):
                        deduped[idx]["담당자"] = (
                            f"{existing}, {new_mgr}".strip(", ") if existing else new_mgr
                        )
            # 주문번호 표기 업데이트 (N건 이상인 경우 "첫번째주문 외 N-1건")
            for entry in deduped:
                cnt = entry.pop("_ord_count", 1)
                if cnt > 1:
                    first_ord = str(entry.get("주문번호", ""))
                    # "XXX 외 N건" 형식으로 표기 (이미 "외 N건"이 붙어 있으면 갱신)
                    base_ord = first_ord.split(" 외 ")[0]
                    entry["주문번호"] = f"{base_ord} 외 {cnt - 1}건"
            self.tax_issues = deduped

        # ── 콘솔 출력 ──
        self.log.info("================================")
        self.log.info("[과세구분 검증 결과]")
        self.log.info(f"  비과세 검토 필요: {cnt_a}건  (상품권·기프티콘류 의심)")
        self.log.info(f"  면세 검토 필요:   {cnt_b}건  (면세 품목 의심)")
        self.log.info(f"  과세구분 혼재:    {cnt_c}건  (마스터 오류 가능)")
        if cnt_wl:
            self.log.info(f"  과세확인목록 제외: {cnt_wl}건  (담당자 확인 완료 상품코드)")
        self.log.info("================================")
        if cnt_a + cnt_b + cnt_c > 0:
            self.log.warning("  계산서 발행 전 [과세구분검증] 시트를 반드시 확인하세요!")

    # ── 시트0: 과세구분검증 ───────────────────────────────────────
    def write_tax_validation_sheet(self, wb):
        self.log.info("  [시트0] 과세구분검증 작성...")
        ws = wb.active
        ws.title = "⓪과세구분검증"

        issues = getattr(self, "tax_issues", [])

        cnt_a      = sum(1 for i in issues if i["검증결과"] == "비과세 검토 필요")
        cnt_b      = sum(1 for i in issues if i["검증결과"] == "면세 검토 필요")
        cnt_c      = sum(1 for i in issues if i["검증결과"] == "과세구분 혼재")
        cnt_red    = cnt_a + cnt_c
        cnt_yellow = cnt_b

        COLS  = ["상품코드", "주문번호", "상품명", "대표규격", "담당자",
                 "현재과세구분", "검증결과", "탐지키워드", "사유", "담당자의견", "비고"]
        COL_W = [14, 18, 28, 24, 12, 12, 16, 12, 50, 30, 20]

        for i, w in enumerate(COL_W, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # ── 상단 배너 ──
        span = f"A1:{get_column_letter(len(COLS))}1"
        ws.merge_cells(span)
        if not issues:
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

        c = ws.cell(1, 1, banner_txt)
        c.font      = Font(name=FONT, bold=True, size=12, color=b_fg)
        c.fill      = PatternFill("solid", start_color=b_bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        data_row = 3

        if not issues:
            ws.cell(data_row, 1, "(키워드 탐지 건 없음)").font = Font(
                name=FONT, size=10, color="808080", italic=True)
            return

        # ── 헤더 ──
        for ci, h in enumerate(COLS, 1):
            hdr_cell(ws.cell(data_row, ci, h))
        ws.row_dimensions[data_row].height = 18
        ws.freeze_panes = f"A{data_row + 1}"
        data_row += 1

        # ── 데이터 ──
        fill_map = {
            "비과세 검토 필요": "FFCCCC",
            "과세구분 혼재"   : "FFCCCC",
            "면세 검토 필요"  : "FFFF99",
        }
        code_ci = COLS.index("상품코드") + 1

        for issue in issues:
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

    # ── STEP 1: 기간 필터 ──────────────────────────────────────
    def step1_filter(self):
        self.log.info("\n[STEP 1] 정산 기간 필터링")
        self.ps, self.pe = self.get_period()

        str_kt = [self.kt_cols.get("주문번호"), self.kt_cols.get("요청번호")]
        str_pl = [self.pl_cols.get("주문번호"), self.pl_cols.get("요청번호")]

        df_kt = _read_excel_smart(self.kt_path, search_cols=["구매문서번호"],
                                  dtype={c: str for c in str_kt if c})
        df_pl = _read_excel_smart(self.pl_path, search_cols=["주문번호", "일정산번호"],
                                  dtype={c: str for c in str_pl if c})

        # KT 전처리: 합계 행 제거 + 세금코드명 정규화
        df_kt = self._preprocess_kt(df_kt)

        # 매출총이익 컬럼이 없으면 정산금액 - 매입금액으로 계산
        gp_col = self.pl_cols.get("매출총이익")
        if not gp_col or gp_col not in df_pl.columns:
            amt_c = self.pl_cols.get("정산금액")
            buy_c = self.pl_cols.get("매입금액")
            if amt_c and buy_c and amt_c in df_pl.columns and buy_c in df_pl.columns:
                df_pl["매출총이익"] = (pd.to_numeric(df_pl[amt_c], errors="coerce")
                                      - pd.to_numeric(df_pl[buy_c], errors="coerce"))
                self.pl_cols["매출총이익"] = "매출총이익"
                self.log.info("  매출총이익 = 정산금액 - 매입금액 으로 계산 (컬럼 자동 생성)")

        dc_kt = self.kt_cols["입고일"]
        df_kt[dc_kt] = pd.to_datetime(df_kt[dc_kt], errors="coerce").dt.date

        # KT 파일은 정산 대상만 담긴 자료이므로 날짜 필터 없이 전체 사용
        self.df_kt = df_kt.copy().reset_index(drop=True)

        # 플랫폼: 승인일(=입고승인일) 기준 우선 필터
        #   우선순위: 승인일 > 일일정산일 > 입고일
        #   (입고일은 물리적 납품일로 정산 기간과 다를 수 있음)
        dc_pl_apv    = self.pl_cols.get("승인일")     # 입고승인일 (= KT 입고승인일)
        dc_pl_settle = self.pl_cols.get("일일정산일")
        dc_pl_in     = self.pl_cols.get("입고일")
        if dc_pl_apv and dc_pl_apv in df_pl.columns:
            df_pl[dc_pl_apv] = pd.to_datetime(df_pl[dc_pl_apv], errors="coerce").dt.date
            self.df_pl = df_pl[(df_pl[dc_pl_apv] >= self.ps) &
                               (df_pl[dc_pl_apv] <= self.pe)].copy().reset_index(drop=True)
            self.log.info(f"  플랫폼 승인일(입고승인일) 기준 필터: {dc_pl_apv}")
        elif dc_pl_settle and dc_pl_settle in df_pl.columns:
            df_pl[dc_pl_settle] = pd.to_datetime(df_pl[dc_pl_settle], errors="coerce").dt.date
            self.df_pl = df_pl[(df_pl[dc_pl_settle] >= self.ps) &
                               (df_pl[dc_pl_settle] <= self.pe)].copy().reset_index(drop=True)
            self.log.warning(f"  플랫폼 일일정산일 기준 필터 (승인일 컬럼 없음): {dc_pl_settle}")
        elif dc_pl_in and dc_pl_in in df_pl.columns:
            df_pl[dc_pl_in] = pd.to_datetime(df_pl[dc_pl_in], errors="coerce").dt.date
            self.df_pl = df_pl[(df_pl[dc_pl_in] >= self.ps) &
                               (df_pl[dc_pl_in] <= self.pe)].copy().reset_index(drop=True)
            self.log.warning(f"  플랫폼 입고일 기준 필터 (승인일/일일정산일/입고일 컬럼 없음): {dc_pl_in}")
        else:
            self.df_pl = df_pl.copy().reset_index(drop=True)
            self.log.warning("  플랫폼 날짜 필터 적용 불가 (승인일/일일정산일/입고일 컬럼 없음) — 전체 사용")

        # 반품 원(原)주문은 이전 정산 기간 소속일 수 있으므로 필터 전 전체 데이터를 보관
        # _route_b / build_invoice_df 의 반품 플랫폼 조회에 사용
        self.df_pl_full = df_pl.copy().reset_index(drop=True)

        self.log.info(f"  KT 전체: {len(self.df_kt)}건 / 플랫폼 기간 내: {len(self.df_pl)}건")
        # 복합 키 생성 (구매문서번호+구매품목 ↔ 주문번호+품목번호)
        self._assign_match_keys()

    # ── STEP 2: 반품 선분리 ────────────────────────────────────
    def step2_separate_returns(self):
        self.log.info("\n[STEP 2] 반품 선분리 (이동유형 기준)")
        mv_col = self.kt_cols.get("이동유형")
        if not mv_col:
            self.log.warning("  이동유형 컬럼 미발견 → 전체 일반 처리")
            self.df_normal = self.df_kt.copy()
            self.df_ret_kt = self.df_kt.iloc[0:0].copy()
            return

        self.log.info(f"  이동유형 컬럼: '{mv_col}'  고유값: {self.df_kt[mv_col].dropna().unique().tolist()}")
        mask = self.df_kt[mv_col].astype(str).str.contains("반품", na=False)
        self.df_ret_kt = self.df_kt[mask].copy().reset_index(drop=True)
        self.df_normal = self.df_kt[~mask].copy().reset_index(drop=True)
        self.log.info(f"  일반: {len(self.df_normal)}건 / 반품: {len(self.df_ret_kt)}건")

    # ── STEP 3: 일반 주문 대사 ─────────────────────────────────
    def step3_reconcile_normal(self):
        key_label = "복합 키" if self.composite_key else "주문번호"
        self.log.info(f"\n[STEP 3] 일반 주문 대사 ({key_label} 기준)")
        amt_kt = self.kt_cols["정산금액"]
        amt_pl = self.pl_cols["정산금액"]

        # _match_key 기반 집합 비교
        kt_set = set(self.df_normal["_match_key"])
        pl_set = set(self.df_pl["_match_key"])

        matched = kt_set & pl_set
        kt_only = kt_set - pl_set
        pl_only = pl_set - kt_set

        # A: 일치 + 금액 비교
        df_m = self.df_normal[self.df_normal["_match_key"].isin(matched)].copy()
        pl_amt_map = (self.df_pl.set_index("_match_key")[amt_pl]
                      .apply(pd.to_numeric, errors="coerce")
                      .to_dict())
        df_m["_amt_kt"] = pd.to_numeric(df_m[amt_kt], errors="coerce")
        df_m["플랫폼정산금액"] = df_m["_match_key"].map(pl_amt_map)
        df_m["금액차이"] = df_m["_amt_kt"] - df_m["플랫폼정산금액"]
        df_m["플랫폼대사결과"] = "정상"
        df_m.drop(columns=["_amt_kt"], inplace=True)

        self.df_matched  = df_m
        self.df_amtdiff  = df_m[df_m["금액차이"].abs() > 1].copy()
        self.df_missing  = self.df_normal[self.df_normal["_match_key"].isin(kt_only)].copy()
        df_pl_only_all   = self.df_pl[self.df_pl["_match_key"].isin(pl_only)].copy()

        # KT미포함: 플랫폼 '승인일'이 정산기간 내인 건만 → 연동 오류 후보
        승인일_col = next((c for c in ["승인일"] if c in df_pl_only_all.columns), None)
        if 승인일_col and not df_pl_only_all.empty:
            df_pl_only_all[승인일_col] = pd.to_datetime(
                df_pl_only_all[승인일_col], errors="coerce").dt.date
            self.df_pl_only = df_pl_only_all[
                df_pl_only_all[승인일_col].apply(
                    lambda d: (self.ps <= d <= self.pe) if pd.notna(d) else False)
            ].copy().reset_index(drop=True)
        else:
            self.df_pl_only = df_pl_only_all.reset_index(drop=True)

        self.log.info(f"  [A] 일치        : {len(matched)}건")
        self.log.info(f"  [B] 플랫폼 누락 : {len(kt_only)}건  ← KT에만 존재 (연동 오류)")
        self.log.info(f"  [C] KT 미포함 (승인일 기간 내): {len(self.df_pl_only)}건  "
                      f"← 전체 {len(df_pl_only_all)}건 중 승인일 기준 필터")
        self.log.info(f"  금액 불일치     : {len(self.df_amtdiff)}건  (±1원 초과)")
        if kt_only:
            self.log.info(f"  누락 복합키     : {sorted(kt_only)}")

    # ── STEP 4: 반품 대사 ──────────────────────────────────────
    def step4_reconcile_returns(self):
        self.log.info("\n[STEP 4] 반품 처리")
        if len(self.df_ret_kt) == 0:
            self.log.info("  반품 건 없음")
            self.df_ret_ok = self.df_ret_kt.copy()
            self.df_ret_ng = self.df_ret_kt.copy()
            self.ret_route = "없음"
            return

        pl_req_col = self.pl_cols.get("요청번호")
        if pl_req_col:
            self._route_a(pl_req_col)
        else:
            self._route_b()

    def _route_a(self, pl_req_col):
        self.ret_route = "A"
        self.log.info("  경로 A: 요청번호 직접 매칭")
        req_col = self.kt_cols.get("요청번호")
        if not req_col:
            self.log.warning("  KT 요청번호 컬럼 없음 → 전체 미확인")
            self.df_ret_ok = self.df_ret_kt.iloc[0:0].copy()
            self.df_ret_ng = self.df_ret_kt.copy()
            return

        pl_reqs = set(self.df_pl[pl_req_col].dropna().astype(str))
        mask = self.df_ret_kt[req_col].astype(str).isin(pl_reqs)

        ok = self.df_ret_kt[mask].copy()
        ok["플랫폼대사결과"] = "반품"
        # 플랫폼 금액 매칭 (플랫폼은 반품을 음수로 기록 → abs로 정규화)
        pl_amt_map = self.df_pl.set_index(pl_req_col)[self.pl_cols["정산금액"]].apply(
            pd.to_numeric, errors="coerce"
        ).abs().to_dict()
        ok["플랫폼정산금액"] = ok[req_col].astype(str).map(pl_amt_map)
        ok["금액차이"] = pd.to_numeric(ok[self.kt_cols["정산금액"]], errors="coerce") - ok["플랫폼정산금액"]

        self.df_ret_ok = ok
        ng = self.df_ret_kt[~mask].copy()
        ng["플랫폼대사결과"] = "미확인"
        self.df_ret_ng = ng
        self.log.info(f"  반품 매칭 성공: {len(ok)}건 / 미확인: {len(ng)}건")

    @staticmethod
    def _xls_engine(path: Path) -> str:
        """파일 확장자에 맞는 pandas Excel 엔진 반환. .xls → xlrd, 나머지 → openpyxl."""
        if not str(path).lower().endswith(".xls"):
            return "openpyxl"
        try:
            import xlrd  # noqa: F401
            return "xlrd"
        except ImportError:
            raise ImportError(
                f"'.xls' 파일 읽기에 xlrd 패키지가 필요합니다.\n"
                f"  파일: {path.name}\n"
                f"  설치 명령: pip install xlrd\n"
                f"  또는: conda install xlrd\n"
                f"  ※ 또는 해당 파일을 .xlsx로 다시 저장하면 바로 사용 가능합니다."
            )

    def _auto_match_return(self, ret_row: pd.Series):
        """KT 반품 행을 플랫폼 데이터에서 자동 매칭 시도.

        주문번호가 달라도 금액(±1원) + 협력사명 + 품목번호가 같은 플랫폼 행을 찾는다.
        이미 일반 주문에 매칭된 플랫폼 행(df_matched)은 후보에서 제외한다.

        Returns:
            (pl_key: str, pl_amt: float)  — 매칭 성공
            (None, None)                  — 매칭 실패
        """
        if self.df_pl is None or self.df_pl.empty or "_match_key" not in self.df_pl.columns:
            return None, None

        kt_amt_col  = self.kt_cols["정산금액"]
        kt_corp_col = self.kt_cols.get("협력사명")
        kt_item_col = self.kt_cols.get("구매품목")
        pl_amt_col  = self.pl_cols["정산금액"]
        pl_corp_col = self.pl_cols.get("협력사명")
        pl_item_col = self.pl_cols.get("품목번호")

        kt_amt  = pd.to_numeric(ret_row.get(kt_amt_col, None), errors="coerce")
        kt_corp = str(ret_row.get(kt_corp_col, "")).strip() if kt_corp_col else ""
        kt_item = str(ret_row.get(kt_item_col, "")).strip() if kt_item_col else ""

        if pd.isna(kt_amt):
            return None, None

        # 이미 일반 주문에 매칭된 플랫폼 행 제외
        matched_keys: set = set()
        if self.df_matched is not None and not self.df_matched.empty \
                and "_match_key" in self.df_matched.columns:
            matched_keys = set(self.df_matched["_match_key"].astype(str))

        cands = self.df_pl[~self.df_pl["_match_key"].isin(matched_keys)].copy()
        if cands.empty:
            return None, None

        # ① 금액 필터 (필수, ±1원)
        # 플랫폼은 반품을 음수로 기록하므로 절대값 비교
        pl_amt_s = pd.to_numeric(cands[pl_amt_col], errors="coerce")
        cands = cands[(pl_amt_s.abs() - abs(kt_amt)).abs() <= 1]
        if cands.empty:
            return None, None

        # ② 협력사명 필터 (있는 경우)
        if kt_corp and pl_corp_col and pl_corp_col in cands.columns:
            corp_m = cands[pl_corp_col].astype(str).str.strip() == kt_corp
            if corp_m.any():
                cands = cands[corp_m]

        # ③ 품목번호 필터 (있는 경우)
        if kt_item and pl_item_col and pl_item_col in cands.columns:
            item_m = cands[pl_item_col].astype(str).str.strip() == kt_item
            if item_m.any():
                cands = cands[item_m]

        if cands.empty:
            return None, None

        if len(cands) > 1:
            self.log.warning(
                f"  [반품 자동매칭] 후보 {len(cands)}건 → 첫 번째 선택 "
                f"(금액 {kt_amt:,.0f}원 / {kt_corp})"
            )

        best    = cands.iloc[0]
        pl_key  = str(best["_match_key"])
        pl_amt  = pd.to_numeric(best[pl_amt_col], errors="coerce")
        # 플랫폼 반품 음수 → 절대값으로 정규화하여 반환
        return pl_key, (float(abs(pl_amt)) if pd.notna(pl_amt) else None)

    def _route_b(self):
        self.ret_route = "B"
        self.log.info("  경로 B: 자동매칭 + return_mapping.xlsx 사용")
        req_col      = self.kt_cols.get("요청번호")
        ord_col      = self.kt_cols["주문번호"]
        corp_col     = self.kt_cols.get("협력사명")
        amt_col      = self.kt_cols["정산금액"]
        kt_item_col  = self.kt_cols.get("구매품목")

        # ── Phase 1: 자동매칭 (주문번호 불일치 대응) ──────────────────────
        # 금액 + 협력사명 + 품목번호가 같은 플랫폼 행을 자동으로 탐색한다.
        ok_list, ng_rows = [], []
        for _, r in self.df_ret_kt.iterrows():
            pl_key, pl_amt = self._auto_match_return(r)
            if pl_key:
                kt_amt = pd.to_numeric(r.get(amt_col), errors="coerce")
                ok_list.append({
                    **r,
                    "플랫폼대사결과" : "반품",
                    "플랫폼정산금액" : pl_amt,
                    "금액차이"       : (kt_amt - pl_amt)
                                       if (pd.notna(kt_amt) and pl_amt is not None) else None,
                    "_ret_pl_key"    : pl_key,   # 플랫폼 행 조회용 (build_invoice_df에서 사용)
                })
                self.log.info(
                    f"  [자동매칭] KT {r.get(ord_col,'')} → 플랫폼 {pl_key} "
                    f"(금액 {kt_amt:,.0f}원)"
                )
            else:
                ng_rows.append(r)

        if ok_list:
            self.log.info(f"  자동매칭 성공: {len(ok_list)}건")

        # 자동매칭으로 전부 해결된 경우
        if not ng_rows:
            self.df_ret_ok = pd.DataFrame(ok_list)
            self.df_ret_ng = self.df_ret_kt.iloc[0:0].copy()
            self.log.info("  반품 전건 자동매칭 완료 → return_mapping.xlsx 불필요")
            return

        # ng_rows: 자동매칭 실패 건 → return_mapping.xlsx 로 처리
        df_ng = pd.DataFrame(ng_rows)

        # ── 공통: 템플릿 생성 헬퍼 ──────────────────────────────────────
        def _make_rm_template(target_df, reason_label):
            """자동매칭 실패 건 → return_mapping.xlsx 템플릿 생성 후 ReturnMappingNeeded raise."""
            rows = []
            for _, r in target_df.iterrows():
                rows.append({
                    # ★ 사용자 입력 영역 (A·B열) — 통합플랫폼에서 확인 후 입력
                    "★플랫폼반품주문번호": "",   # 플랫폼 반품 주문번호 (예: 4502061937)
                    "★플랫폼품목번호"    : "",   # 플랫폼 품목번호       (예: 10)
                    # 참고용 KT 정보 (자동 입력, 수정 불필요)
                    "KT반품주문번호"     : r.get(ord_col, ""),
                    "KT구매품목"         : (r.get(kt_item_col, "") if kt_item_col else ""),
                    "KT요청번호"         : (r.get(req_col, "")     if req_col     else ""),
                    "협력사명"           : (r.get(corp_col, "")    if corp_col    else ""),
                    "KT정산금액"         : r.get(amt_col, ""),
                    # 내부 처리용 (수정 불필요)
                    "KT반품복합키"       : r.get("_match_key", r.get(ord_col, "")),
                })
            rm_df  = pd.DataFrame(rows)
            rm_out = self.base / "return_mapping.xlsx"
            rm_df.to_excel(rm_out, index=False)
            n        = len(rows)
            auto_msg = (f"  (자동매칭 성공 {len(ok_list)}건은 별도 처리됩니다)\n"
                        if ok_list else "")
            self.log.info(
                f"\n{'='*60}\n"
                f"  반품 {n}건의 매핑 파일이 생성됐습니다. ({reason_label})\n"
                f"  저장 위치: {rm_out}\n"
                f"{auto_msg}\n"
                f"  ★ 작성 방법:\n"
                f"  1. return_mapping.xlsx 열기\n"
                f"  2. A열(★플랫폼반품주문번호): 통합플랫폼에서 대응 반품의 주문번호 입력\n"
                f"     예) 4502061937\n"
                f"  3. B열(★플랫폼품목번호): 해당 품목번호 입력  예) 10\n"
                f"  4. C~H열은 참고용 KT 정보로 수정 불필요\n"
                f"  5. 작성된 return_mapping.xlsx를 반품 매핑 파일로 선택 후 재실행\n"
                f"{'='*60}"
            )
            raise ReturnMappingNeeded(
                f"반품 매핑 템플릿 생성 완료 ({n}건) → {rm_out}\n"
                f"A열(★플랫폼반품주문번호)·B열(★플랫폼품목번호) 작성 후 재실행해주세요."
            )

        if not self.rm_path or not self.rm_path.exists():
            # ① 파일 없음 → 자동매칭 실패 건만 템플릿 생성 후 종료
            _make_rm_template(df_ng, "파일 없음")

        # ② 파일 있음 → 컬럼 검증 후 매칭 실행
        _rm_eng = self._xls_engine(self.rm_path)
        try:
            rm = pd.read_excel(self.rm_path, dtype=str, engine=_rm_eng).fillna("")
        except Exception:
            _rm_fallback = "xlrd" if _rm_eng == "openpyxl" else "openpyxl"
            try:
                rm = pd.read_excel(self.rm_path, dtype=str, engine=_rm_fallback).fillna("")
                self.log.warning(
                    f"  반품매핑 엔진 재시도({_rm_fallback}) 성공: {self.rm_path.name}"
                )
            except Exception:
                self.log.warning("  return_mapping.xlsx 읽기 실패 → 새 템플릿을 생성합니다.")
                _make_rm_template(df_ng, "파일 읽기 실패")
        _expected_new = {"★플랫폼반품주문번호", "★플랫폼품목번호"}
        _expected_old = {"KT반품복합키", "KT반품주문번호"}
        if not ((_expected_new | _expected_old) & set(rm.columns)):
            # 기대 컬럼 없음 → 잘못된 파일 선택 → 템플릿 재생성
            self.log.warning(
                f"  [경고] 반품 매핑 파일에 필수 컬럼이 없습니다 ({self.rm_path.name}). "
                f"새 템플릿을 생성합니다."
            )
            self.rm_path = None
            _make_rm_template(df_ng, "잘못된 파일 감지")

        # ★열(신규) 또는 플랫폼원복합키(구) 모두 비어있으면 → 신규 양식 재생성
        _new_filled = (
            "★플랫폼반품주문번호" in rm.columns
            and rm["★플랫폼반품주문번호"].str.strip().ne("").any()
        )
        _old_filled = (
            "플랫폼원복합키" in rm.columns
            and rm["플랫폼원복합키"].str.strip().ne("").any()
        ) or (
            "플랫폼원주문번호" in rm.columns
            and rm["플랫폼원주문번호"].str.strip().ne("").any()
        )
        if not _new_filled and not _old_filled:
            self.log.warning(
                f"  [경고] return_mapping.xlsx 에 플랫폼 주문번호가 입력되지 않았습니다. "
                f"새 양식(★열)으로 템플릿을 재생성합니다."
            )
            _make_rm_template(df_ng, "매핑값 미입력")

        pl_amt_col = self.pl_cols["정산금액"]
        # 플랫폼 금액 룩업: 반품 원주문은 정산 기간 밖일 수 있으므로 전체 데이터(df_pl_full) 사용
        _pl_for_ret = getattr(self, "df_pl_full", self.df_pl)
        pl_amt_map = (_pl_for_ret.set_index("_match_key")[pl_amt_col]
                      .apply(pd.to_numeric, errors="coerce")
                      .to_dict())

        # Phase 2: 자동매칭 실패 건(df_ng)에 대해 return_mapping.xlsx로 매칭
        rm_ok_list, rm_ng_list = [], []
        for _, ret_row in df_ng.iterrows():
            kt_key  = str(ret_row.get("_match_key", "")).strip()
            kt_ord  = str(ret_row.get(ord_col, "")).strip()
            kt_item = str(ret_row.get(kt_item_col, "")).strip() if kt_item_col else ""

            # ── return_mapping에서 이 반품 건 찾기 ──
            # 행 단위 폴백: ① KT반품복합키(비어있지 않은 경우) → ② KT반품주문번호+KT구매품목 → ③ KT반품주문번호
            rm_row = pd.DataFrame()
            if "KT반품복합키" in rm.columns:
                _cand = rm[rm["KT반품복합키"].str.strip() == kt_key]
                if not _cand.empty:
                    rm_row = _cand
            if rm_row.empty and "KT반품주문번호" in rm.columns and "KT구매품목" in rm.columns and kt_item:
                _cand = rm[
                    (rm["KT반품주문번호"].str.strip() == kt_ord) &
                    (rm["KT구매품목"].astype(str).str.strip() == kt_item)
                ]
                if not _cand.empty:
                    rm_row = _cand
            if rm_row.empty and "KT반품주문번호" in rm.columns:
                _cand = rm[rm["KT반품주문번호"].str.strip() == kt_ord]
                if not _cand.empty:
                    rm_row = _cand

            if rm_row.empty:
                rm_ng_list.append({**ret_row, "플랫폼대사결과": "미확인",
                                    "플랫폼정산금액": None, "금액차이": None})
                continue

            # ── 플랫폼 복합키 조합 ──
            # 신규 양식: ★플랫폼반품주문번호 + ★플랫폼품목번호 → 자동 조합
            # 구 양식: 플랫폼원복합키 → 그대로 사용 (하위 호환)
            r0 = rm_row.iloc[0]
            if "★플랫폼반품주문번호" in rm.columns:
                pl_ord_val = str(r0.get("★플랫폼반품주문번호", "")).strip()
                # 품목번호는 _item_str 과 동일하게 정수 문자열로 정규화
                # (Excel이 dtype=str 이어도 "6.0" 으로 읽힐 수 있으므로)
                _itm_raw   = str(r0.get("★플랫폼품목번호", "")).strip()
                try:
                    pl_itm_val = str(int(float(_itm_raw))) if _itm_raw not in ("", "nan") else ""
                except (ValueError, TypeError):
                    pl_itm_val = _itm_raw
                pl_ord_val = self._key_str(pl_ord_val)   # 주문번호도 동일 정규화
                pl_key     = (pl_ord_val + pl_itm_val) if pl_ord_val else ""
            elif "플랫폼원복합키" in rm.columns:
                pl_key = str(r0.get("플랫폼원복합키", "")).strip()
            else:
                pl_key = str(r0.get("플랫폼원주문번호", "")).strip()

            if not pl_key:
                rm_ng_list.append({**ret_row, "플랫폼대사결과": "미확인",
                                    "플랫폼정산금액": None, "금액차이": None})
                continue

            pl_amt_raw = pl_amt_map.get(pl_key)
            # 플랫폼은 반품을 음수로 기록 → abs로 정규화
            pl_amt = abs(pl_amt_raw) if pl_amt_raw is not None else None
            kt_amt = pd.to_numeric(ret_row.get(amt_col), errors="coerce")
            # 진단: 반품 건별 금액 조회 결과 기록
            self.log.info(
                f"  [반품금액진단] KT키={kt_key}  플랫폼키={pl_key}"
                f"  KT금액={kt_amt}  플랫폼금액(원본)={pl_amt_raw}  플랫폼금액(abs)={pl_amt}"
                f"  pl_key_in_map={'YES' if pl_key in pl_amt_map else 'NO'}"
            )
            rm_ok_list.append({**ret_row,
                                "플랫폼대사결과" : "반품",
                                "플랫폼정산금액" : pl_amt,
                                "금액차이"       : (kt_amt - pl_amt) if (pd.notna(kt_amt) and pl_amt is not None) else None,
                                "_ret_pl_key"    : pl_key,  # 플랫폼 행 조회용
                                })

        # Phase 1(자동매칭) + Phase 2(return_mapping) 결과 합산
        all_ok = ok_list + rm_ok_list
        self.df_ret_ok = pd.DataFrame(all_ok) if all_ok else self.df_ret_kt.iloc[0:0].copy()
        self.df_ret_ng = pd.DataFrame(rm_ng_list) if rm_ng_list else self.df_ret_kt.iloc[0:0].copy()
        self.log.info(
            f"  반품 매칭 결과: 자동매칭 {len(ok_list)}건 + 매핑파일 {len(rm_ok_list)}건 "
            f"= 성공 {len(all_ok)}건 / 미확인: {len(rm_ng_list)}건"
        )

    # ── 인보이스 DataFrame 조립 (KT 원본 기준) ───────────────────
    def build_invoice_df(self) -> pd.DataFrame:
        """KT 원본 데이터를 기준으로 인보이스를 조립한다.
        플랫폼 데이터는 금액 비교용으로만 참조하며 행을 추가하지 않는다.
        인보이스 건수 = KT 기간 필터 건수와 동일해야 한다.
        """
        kt_amt_col = self.kt_cols["정산금액"]   # 공급가액
        pl_amt_col = self.pl_cols["정산금액"]   # 정산금액
        n_kt_orig  = len(self.df_kt)

        # ── KT를 기준(base)으로 사용 ──────────────────────────
        df = self.df_kt.copy()

        # ── 플랫폼 룩업 헬퍼 (_match_key → 각 컬럼) ──────────
        def _pl_lookup(col_key, numeric=False):
            col = self.pl_cols.get(col_key)
            if col and col in self.df_pl.columns:
                s = self.df_pl.set_index("_match_key")[col]
                if numeric:
                    s = s.apply(pd.to_numeric, errors="coerce")
                return s.to_dict()
            return {}

        # 플랫폼 정산금액
        pl_amt_map = _pl_lookup("정산금액", numeric=True)
        df["플랫폼정산금액"] = df["_match_key"].map(pl_amt_map)
        df[kt_amt_col]       = pd.to_numeric(df[kt_amt_col], errors="coerce")
        df["금액차이"]       = df["플랫폼정산금액"] - df[kt_amt_col]

        # 대사결과 초기화
        df["대사결과"] = "정상"

        # 플랫폼 누락 (KT에만 존재)
        if self.df_missing is not None and len(self.df_missing) > 0:
            missing_keys = set(self.df_missing["_match_key"].astype(str))
            m = df["_match_key"].astype(str).isin(missing_keys)
            df.loc[m, "대사결과"]                       = "플랫폼누락"
            df.loc[m, ["플랫폼정산금액", "금액차이"]] = pd.NA

        # 금액 불일치 (±1원 초과)
        amt_diff_mask = df["플랫폼정산금액"].notna() & (df["금액차이"].abs() > 1)
        df.loc[amt_diff_mask & (df["대사결과"] == "정상"), "대사결과"] = "금액불일치"

        # ── 플랫폼 전체 컬럼 일괄 룩업 (_match_key 기준, 미존재 컬럼만 추가) ──
        # ※ 반품 데이터 채우기(아래)보다 먼저 실행해야 한다:
        #   반품 행의 플랫폼 컬럼을 채울 때 df에 해당 컬럼이 존재해야
        #   `_col in df.columns` 조건을 통과할 수 있기 때문
        try:
            _pl_idx = self.df_pl.set_index("_match_key")
            for _col in self.df_pl.columns:
                if str(_col).startswith("_") or _col in df.columns:
                    continue
                try:
                    df[_col] = df["_match_key"].map(_pl_idx[_col].to_dict())
                except Exception:
                    pass
        except Exception as _e:
            self.log.warning(f"  플랫폼 전체 컬럼 룩업 실패: {_e}")

        # 반품 표시 — KT df_ret_kt의 _match_key 기반 직접 표시 (경로 A/B 공통)
        # ※ 경로 B에서 플랫폼 키(pl_from_rm)와 KT _match_key를 비교하는 이전 로직은 버그
        #   (KT 주문번호 ≠ 플랫폼 주문번호이므로 항상 0건) → KT 반품 복합키로 교체
        if self.df_ret_kt is not None and not self.df_ret_kt.empty:
            ret_kt_keys = set(self.df_ret_kt["_match_key"].astype(str))
            df.loc[df["_match_key"].isin(ret_kt_keys), "대사결과"] = "반품"

            # 매칭 성공 건(df_ret_ok): 플랫폼정산금액 + 금액차이 + 플랫폼 전체 컬럼 업데이트
            if self.df_ret_ok is not None and not self.df_ret_ok.empty:
                # 플랫폼 룩업 테이블 (_match_key → 행 dict)
                # 반품 원주문은 정산 기간 밖일 수 있으므로 전체 데이터(df_pl_full) 사용
                _pl_ret_lookup = {}
                _pl_src = getattr(self, "df_pl_full", self.df_pl)
                if _pl_src is not None:
                    try:
                        # _match_key 중복 시 to_dict(orient='index') 오류 방지
                        # → 첫 번째 행 우선 유지 (keep='first')
                        _dup_cnt = int(_pl_src["_match_key"].duplicated().sum())
                        if _dup_cnt:
                            self.log.info(
                                f"  [반품채우기] df_pl_full 중복 키 {_dup_cnt}건 → "
                                f"첫 번째 행 유지 후 룩업 구축"
                            )
                        _pl_ret_lookup = (
                            _pl_src.drop_duplicates(subset=["_match_key"], keep="first")
                            .set_index("_match_key")
                            .to_dict(orient="index")
                        )
                        self.log.info(
                            f"  [반품채우기] 플랫폼 룩업 구축: {len(_pl_ret_lookup)}건 "
                            f"(df_pl_full {len(_pl_src)}행 기반)"
                        )
                    except Exception as _e:
                        self.log.warning(f"  [반품채우기] 플랫폼 룩업 구축 실패: {_e}")
                for _, ret_ok_row in self.df_ret_ok.iterrows():
                    rk = str(ret_ok_row.get("_match_key", ""))
                    rp = ret_ok_row.get("플랫폼정산금액")
                    ret_pl_key = str(ret_ok_row.get("_ret_pl_key", "")).strip()
                    # ── 진단 ─────────────────────────────────────────
                    self.log.info(
                        f"  [반품채우기] KT키={rk!r}  플랫폼키={ret_pl_key!r}  "
                        f"키존재={'YES' if (ret_pl_key and ret_pl_key in _pl_ret_lookup) else 'NO'}"
                    )
                    if ret_pl_key and ret_pl_key in _pl_ret_lookup:
                        _diag = _pl_ret_lookup[ret_pl_key]
                        for _dc in ["서비스카테고리", "담당자부서", "결제유형"]:
                            self.log.info(
                                f"  [반품채우기]   {_dc}="
                                f"{str(_diag.get(_dc, '(컬럼없음)'))!r}"
                            )
                    # ─────────────────────────────────────────────────
                    if not rk:
                        continue
                    m = df["_match_key"].astype(str) == rk
                    if not m.any():
                        continue
                    if pd.notna(rp):
                        rp_f = float(rp)
                        df.loc[m, "플랫폼정산금액"] = rp_f
                        df.loc[m, "금액차이"] = (
                            pd.to_numeric(df.loc[m, kt_amt_col], errors="coerce") - rp_f
                        )
                    # 플랫폼 행 전체 컬럼 채우기 (매칭 키가 있는 경우)
                    if ret_pl_key and ret_pl_key in _pl_ret_lookup:
                        pl_row_dict = _pl_ret_lookup[ret_pl_key]
                        for _col, _val in pl_row_dict.items():
                            if str(_col).startswith("_"):
                                continue
                            if _col == kt_amt_col:
                                continue
                            # 컬럼이 아직 없으면 생성 (플랫폼 전체 컬럼 룩업이 이전에
                            # 실행되지 못했을 경우의 안전장치)
                            if _col not in df.columns:
                                df[_col] = pd.NA
                            # 기존 값이 비어 있는 경우에만 채움 (KT 원본 값 보존)
                            cur = df.loc[m, _col]
                            is_empty = cur.isnull() | cur.astype(str).str.strip().isin(
                                ["", "nan", "None"])
                            if is_empty.any():
                                df.loc[m & is_empty, _col] = _val

            # 미확인 반품(df_ret_ng): 플랫폼 금액 NaN 유지
            if self.df_ret_ng is not None and not self.df_ret_ng.empty:
                ng_keys = set(self.df_ret_ng["_match_key"].astype(str))
                df.loc[df["_match_key"].isin(ng_keys),
                       ["플랫폼정산금액", "금액차이"]] = pd.NA

        # ── 반품 행 플랫폼 금액 절대값 통일 ─────────────────────────────
        # 플랫폼은 반품 금액(정산금액·매입금액·매출총이익)을 음수로 기록함
        # → 인보이스/금액검증에서 KT 양수와 비교 가능하도록 절대값 처리
        # ※ 매출현황(write_sales_sheet)에서는 다시 음수로 변환하여 NET 합계에 반영
        _ret_abs_mask = (df["대사결과"] == "반품") if "대사결과" in df.columns \
                        else pd.Series(False, index=df.index)
        if _ret_abs_mask.any():
            _pl_money_cols = list({c for c in [
                "플랫폼정산금액",
                self.pl_cols.get("매입금액"),
                self.pl_cols.get("매출총이익"),
            ] if c and c in df.columns})
            for _amc in _pl_money_cols:
                df.loc[_ret_abs_mask, _amc] = (
                    pd.to_numeric(df.loc[_ret_abs_mask, _amc], errors="coerce").abs()
                )

        # 반품 표시 — 경로 A 보완 (요청번호 기반 매칭 성공 건 재확인)
        # ※ 경로 B는 _match_key 기반(위)으로 충분하므로 경로 A 전용으로 제한
        kt_req_col = self.kt_cols.get("요청번호")
        if (getattr(self, "ret_route", None) == "A"
                and kt_req_col and kt_req_col in df.columns
                and self.df_ret_ok is not None and len(self.df_ret_ok) > 0):
            if kt_req_col in self.df_ret_ok.columns:
                ok_reqs = {
                    v for v in self.df_ret_ok[kt_req_col].astype(str)
                    if v and v.lower() not in ("nan", "none", "")
                }
                if ok_reqs:
                    df.loc[df[kt_req_col].astype(str).isin(ok_reqs), "대사결과"] = "반품"

        # 플랫폼 과세구분 값 정규화 (영세/영세율 → 면세)
        _PL_TAX_NORM = {"영세": "면세", "영세율": "면세"}
        for _tc in ["매출과세구분", "매입과세구분"]:
            if _tc in df.columns:
                df[_tc] = (df[_tc].astype(str).str.strip()
                           .map(lambda v, _n=_PL_TAX_NORM: _n.get(v, v) if v not in ("nan", "None", "") else ""))

        # 컬럼명 정규화 (줄바꿈 제거 및 약식명으로 통일)
        _COL_RENAME = {
            "관리회계\n(IP만변경적용)": "관리회계",
            "관리회계(IP만변경적용)":   "관리회계",
            "인보이스\n(KT이동유형명 값)": "인보이스(KT이동유형명 값)",
        }
        for _old, _new in _COL_RENAME.items():
            if _old in df.columns:
                if _new not in df.columns:
                    df.rename(columns={_old: _new}, inplace=True)
                else:
                    df.drop(columns=[_old], inplace=True)

        # KT 이동유형명 → 인보이스(KT이동유형명 값) (플랫폼 컬럼 없을 때)
        kt_mv_col = self.kt_cols.get("이동유형")
        if kt_mv_col and kt_mv_col in df.columns and "인보이스(KT이동유형명 값)" not in df.columns:
            df["인보이스(KT이동유형명 값)"] = df[kt_mv_col]

        # 매출총이익 자동 계산 (없는 경우)
        gp_col  = self.pl_cols.get("매출총이익")
        buy_col = self.pl_cols.get("매입금액")
        if (not gp_col or gp_col not in df.columns) and buy_col and buy_col in df.columns:
            df["매출총이익"] = (
                df["플랫폼정산금액"].fillna(0)
                - pd.to_numeric(df[buy_col], errors="coerce").fillna(0))
            self.pl_cols["매출총이익"] = "매출총이익"

        # XLOOKUP 매핑 (서비스카테고리 → 관리회계/담당자/일반통신구분)
        if self.master and self.master.loaded:
            merged_cols = {**self.pl_cols, **self.kt_cols}
            df, self.map_stats = self.master.apply(df, merged_cols, self.log)

        # ── 담당자 단위 일반/통신구분 일관성 보정 ──────────────────────────
        # 원칙: 담당자는 통신(네트워크사업1팀) 또는 일반(전략구매사업1팀) 중 하나에 소속.
        # 인사이동으로 담당자부서가 변경된 경우 → 매출구분 키워드로 보완.
        # 최종적으로 담당자별 다수결(신호 건수 多인 쪽)로 전체 행을 통일.
        #
        # 신호 우선순위 (행별):
        #   ① 팀명 직접 일치: 네트워크사업1팀 → 통신 / 전략구매사업1팀 → 일반
        #   ② 매출구분 키워드: "(통신)" → 통신 / "(일반)" → 일반  (인사이동 보정)
        #   ③ 기타 네트워크 계열 부서명 → 통신
        _tc_col_ov   = next((c for c in ["일반/통신구분", "일반통신구분"] if c in df.columns), None)
        _mgr_col_ov  = next((c for c in ["담당자"] if c in df.columns), None)
        _dept_col_ov = next((c for c in ["담당자부서"] if c in df.columns), None)
        _sale_col_ov = next((c for c in ["매출구분"] if c in df.columns), None)
        if _tc_col_ov and _mgr_col_ov:
            def _row_signal(row):
                dept = str(row.get(_dept_col_ov, "")).strip() if _dept_col_ov else ""
                sale = str(row.get(_sale_col_ov, "")).strip() if _sale_col_ov else ""
                if "네트워크사업1팀" in dept: return "통신"
                if "전략구매사업1팀" in dept: return "일반"
                if "(통신)" in sale:          return "통신"
                if "(일반)" in sale:          return "일반"
                if "네트워크" in dept:        return "통신"
                return ""

            df["_tc_signal"] = df.apply(_row_signal, axis=1)

            def _agg_person(grp):
                cnt_t = (grp == "통신").sum()
                cnt_g = (grp == "일반").sum()
                if cnt_t == 0 and cnt_g == 0:
                    return ""   # 신호 없음 → 기존 값 유지
                return "통신" if cnt_t >= cnt_g else "일반"

            _tc_by_person = (
                df.groupby(df[_mgr_col_ov].astype(str))["_tc_signal"]
                .apply(_agg_person)
            )
            df.drop(columns=["_tc_signal"], inplace=True)
            _before_tc = df[_tc_col_ov].copy()
            _mapped = df[_mgr_col_ov].astype(str).map(_tc_by_person)
            # 신호 없음("") 또는 NaN → 기존 값 유지
            _no_signal = _mapped.isin(["", None]) | _mapped.isna()
            df[_tc_col_ov] = _mapped.where(~_no_signal, _before_tc)
            _changed = int((df[_tc_col_ov] != _before_tc).sum())
            if _changed > 0:
                self.log.info(f"  담당자 단위 일반/통신구분 보정: {_changed}건 수정")
            _cnt_t = int((df[_tc_col_ov] == "통신").sum())
            _cnt_g = int((df[_tc_col_ov] == "일반").sum())
            self.log.info(f"  일반/통신구분 최종: 통신 {_cnt_t}건 / 일반 {_cnt_g}건")

        # ── 수주물자 규칙: 매출구분에 '수주물자' 포함 → 관리회계/담당자/일반통신구분 강제 설정 ──
        _매출구분_col = next((c for c in ["매출구분"] if c in df.columns), None)
        if _매출구분_col:
            _unique_vals = df[_매출구분_col].dropna().unique().tolist()
            self.log.info(f"  매출구분 고유값: {_unique_vals[:15]}")
            _suiju = df[_매출구분_col].astype(str).str.contains("수주물자", na=False)
            if _suiju.any():
                # 관리회계 (컬럼명 후보 순서대로 탐지)
                _ma_col = next((c for c in ["관리회계", "관리회계(IP만변경적용)"]
                                if c in df.columns), None)
                if _ma_col:
                    df.loc[_suiju, _ma_col] = "투자형구축물자"
                # 담당자
                _mgr_col = next((c for c in ["담당자"] if c in df.columns), None)
                if _mgr_col:
                    df.loc[_suiju, _mgr_col] = "확인필요"
                # 일반/통신구분
                _tc_col = next((c for c in ["일반/통신구분", "일반통신구분"]
                                if c in df.columns), None)
                if _tc_col:
                    df.loc[_suiju, _tc_col] = "통신"
                self.log.info(
                    f"  수주물자 {int(_suiju.sum())}건: "
                    f"관리회계=투자형구축물자, 담당자=확인필요, 일반/통신구분=통신")

        # ── 국책과제 규칙: 국책과제여부 = Y/O/YES/국책/TRUE/1/국책과제대상 → 관리회계/담당자/일반통신구분 강제 설정 ──
        _gov_col = self.kt_cols.get("국책과제여부")
        if _gov_col and _gov_col in df.columns:
            _gov_mask = df[_gov_col].astype(str).str.strip().str.upper().isin(
                ["Y", "O", "YES", "국책", "TRUE", "1", "국책과제대상"])
            if _gov_mask.any():
                _ma_col_g = next((c for c in ["관리회계", "관리회계(IP만변경적용)"]
                                  if c in df.columns), None)
                if _ma_col_g:
                    df.loc[_gov_mask, _ma_col_g] = "국책과제"
                _mgr_col_g = next((c for c in ["담당자"] if c in df.columns), None)
                if _mgr_col_g:
                    df.loc[_gov_mask, _mgr_col_g] = "확인필요"
                _tc_col_g = next((c for c in ["일반/통신구분", "일반통신구분"]
                                  if c in df.columns), None)
                if _tc_col_g:
                    df.loc[_gov_mask, _tc_col_g] = "통신"
                # 인보이스(KT이동유형명 값) → "국책" 으로 표시
                _inv_col_g = "인보이스(KT이동유형명 값)"
                if _inv_col_g in df.columns:
                    df.loc[_gov_mask, _inv_col_g] = "국책"
                self.log.info(
                    f"  국책과제 {int(_gov_mask.sum())}건: "
                    f"관리회계=국책과제, 담당자=확인필요, 일반/통신구분=통신, 인보이스=국책")

        # ── 선로자재 → 접속자재 규칙 ───────────────────────────────
        # 관리회계 = '선로자재' AND 품명에 '광단자함' 또는 '접속함' 포함
        # → 관리회계를 '접속자재'로 변경 (서비스카테고리 개편 전 발주 건 수작업 처리)
        _ma_col2 = next((c for c in ["관리회계", "관리회계(IP만변경적용)"]
                         if c in df.columns), None)
        _nm_col  = next((c for c in ["상품명", "품목명", "자재명", "품명"]
                         if c in df.columns), None)
        if _ma_col2 and _nm_col:
            _seonro = df[_ma_col2].astype(str).str.strip() == "선로자재"
            _kwmatch = df[_nm_col].astype(str).str.contains("광단자함|접속함", na=False)
            _접속자재 = _seonro & _kwmatch
            if _접속자재.any():
                df.loc[_접속자재, _ma_col2] = "접속자재"
                self.log.info(
                    f"  선로자재→접속자재 {int(_접속자재.sum())}건 "
                    f"(광단자함·접속함 품명 포함)")

        # ── 건수 검증 ──────────────────────────────────────────
        if len(df) != n_kt_orig:
            raise ValueError(
                f"인보이스 건수 오류: KT 원본 {n_kt_orig}건 ≠ 결과 {len(df)}건. "
                "데이터 소스를 확인하세요.")

        # ── 대사 결과 통계 ─────────────────────────────────────
        cnt_normal  = int((df["대사결과"] == "정상").sum())
        cnt_missing = int((df["대사결과"] == "플랫폼누락").sum())
        cnt_return  = int((df["대사결과"] == "반품").sum())
        cnt_amount  = int((df["대사결과"] == "금액불일치").sum())
        _gov_col_stat = self.kt_cols.get("국책과제여부")
        cnt_gov = int(
            df[_gov_col_stat].astype(str).str.strip().str.upper().isin(
                ["Y", "O", "YES", "국책", "TRUE", "1", "국책과제대상"]).sum()
        ) if (_gov_col_stat and _gov_col_stat in df.columns) else 0
        kt_total    = float(df[kt_amt_col].sum())
        # 플랫폼 금액: 미매칭(반품 미확인 등) 행은 NaN → fillna(0) 후 합산
        pl_total    = float(df["플랫폼정산금액"].fillna(0).sum())
        # 차이: kt_total - pl_total (금액차이.sum()은 NaN 행 제외로 부정확)
        diff_total  = kt_total - pl_total

        self.log.info("================================")
        self.log.info("[인보이스 건수 검증]")
        self.log.info(f"  KT 원본 (기간 필터):  {n_kt_orig}건")
        self.log.info(f"  인보이스 결과:        {len(df)}건")
        self.log.info(f"  일치 여부:            ✓ 정상")
        self.log.info("")
        self.log.info("[금액 검증]")
        self.log.info(f"  KT 정산금액 합계:     {kt_total:,.0f}원")
        self.log.info(f"  플랫폼 금액 합계:     {pl_total:,.0f}원")
        diff_sign = "+" if diff_total > 0 else ""
        self.log.info(f"  차이(KT-플랫폼):      {diff_sign}{diff_total:,.0f}원"
                      + ("  ← 반품/누락 미매칭 포함" if abs(diff_total) > 1 else "  ✓"))
        self.log.info("")
        self.log.info("[대사 결과]")
        self.log.info(f"  정상:        {cnt_normal}건")
        self.log.info(f"  플랫폼누락:  {cnt_missing}건")
        self.log.info(f"  반품:        {cnt_return}건")
        self.log.info(f"  금액불일치:  {cnt_amount}건")
        self.log.info(f"  국책과제:    {cnt_gov}건")
        self.log.info("================================")

        # 누락 10건 이상이면 원인 자동 진단
        if cnt_missing >= 10:
            self._diagnose_kt_only(df)

        # 검증 통계 저장 (write_invoice_sheet에서 헤더 작성용)
        self._inv_stats = {
            "n_kt": n_kt_orig, "n_inv": len(df),
            "kt_total": kt_total, "pl_total": pl_total, "diff_total": diff_total,
            "cnt_normal": cnt_normal, "cnt_missing": cnt_missing,
            "cnt_return": cnt_return, "cnt_amount": cnt_amount, "cnt_gov": cnt_gov,
        }

        # 내부 작업 컬럼 제거
        drop_cols = [c for c in df.columns if str(c).startswith("_")]
        df = df.drop(columns=drop_cols, errors="ignore")
        self.df_invoice = df
        return df

    def _diagnose_kt_only(self, df):
        """KT에만 있는 건의 원인 진단.
        - 주문번호는 플랫폼에 있으나 복합키 불일치 → 품목번호 차이 상세 출력
        - 주문번호 자체가 플랫폼에 없는 경우 → 실제 누락
        """
        self.log.info("\n[누락 원인 진단]")
        kt_ord_col  = self.kt_cols["주문번호"]                # 구매문서번호
        kt_item_col = self.kt_cols.get("구매품목")            # 구매품목
        pl_ord_col  = self.pl_cols.get("주문번호")            # 주문번호
        pl_item_col = self.pl_cols.get("품목번호")            # 품목번호
        if not pl_ord_col or pl_ord_col not in self.df_pl.columns:
            self.log.info("  (플랫폼 주문번호 컬럼 없음 — 진단 생략)")
            return

        # 플랫폼: 주문번호 → [품목번호 목록] 맵 구성
        pl_raw_set = set(self.df_pl[pl_ord_col].astype(str).str.strip())
        pl_items_by_ord: dict = {}
        if pl_item_col and pl_item_col in self.df_pl.columns:
            for _ord, _item in zip(
                self.df_pl[pl_ord_col].astype(str).str.strip(),
                self.df_pl[pl_item_col].astype(str).str.strip()
            ):
                pl_items_by_ord.setdefault(_ord, []).append(_item)

        missing_rows = df[df["대사결과"] == "플랫폼누락"]
        key_mismatch = real_missing = 0
        samples_key = []   # 품목번호 차이 샘플
        samples_mis = []   # 실제 누락 샘플

        for _, row in missing_rows.iterrows():
            raw_ord  = str(row.get(kt_ord_col, "")).strip()
            kt_item  = str(row.get(kt_item_col, "")).strip() if kt_item_col else "-"
            if raw_ord in pl_raw_set:
                key_mismatch += 1
                if len(samples_key) < 8:
                    pl_items = pl_items_by_ord.get(raw_ord, [])
                    samples_key.append(
                        f"    주문번호={raw_ord}  KT구매품목={kt_item!r}"
                        f"  플랫폼품목번호={pl_items[:8]}"
                    )
            else:
                real_missing += 1
                if len(samples_mis) < 3:
                    samples_mis.append(f"    주문번호={raw_ord} (플랫폼에 없음)")

        self.log.info(f"  주문번호 일치·복합키 불일치: {key_mismatch}건  ← 품목번호 형식 차이 의심")
        self.log.info(f"  주문번호 자체 누락:          {real_missing}건  ← 실제 미연동")
        if samples_key:
            self.log.info("  [품목번호 차이 샘플]")
            for s in samples_key:
                self.log.info(s)
        if samples_mis:
            self.log.info("  [실제 누락 샘플]")
            for s in samples_mis:
                self.log.info(s)
        if key_mismatch > 0:
            self.log.warning(
                "  !! 복합키 불일치 원인: KT 구매품목 vs 플랫폼 품목번호 값 차이\n"
                "     → 두 시스템의 품목번호 채번 규칙 확인 필요")

    # ── 시트1: 인보이스 ────────────────────────────────────────
    def write_invoice_sheet(self, wb):
        self.log.info("  [시트1] 인보이스 작성...")
        ws      = wb.create_sheet("①인보이스")
        df      = self.df_invoice.copy()
        kt_amt_col = self.kt_cols["정산금액"]   # 공급가액

        # ── 컬럼 순서 정렬 (레거시 관리파일 연속성 기준) ──────────────
        _DESIRED_ORDER = [
            "일정산번호", "정산확정일", "일일정산월",
            "주문번호", "품목번호", "주문&품목",
            "인보이스(KT이동유형명 값)",
            "사업장코드", "사업장", "사업자번호",
            "상위부서코드", "상위부서명", "부서명", "부서코드",
            "상품ID", "카테고리ID", "마스터카테고리(CMS)", "서비스카테고리",
            "관리회계", "담당자",
            "일반/통신구분", "매출구분",
            "상품코드", "상품명", "대표규격", "상태", "선정산여부", "단위",
            "주문수량", "주문단가", "주문금액",
            "정산수량", "정산단가", "정산금액",
            "매입단가", "매입금액", "매출총이익",
            "계정ID", "계정명", "주문자", "주문자ID", "주문형태",
            "코스트센터코드", "코스트센터명", "WBS코드", "WBS명",
            "세금코드", "SA_ID", "입고자ID", "입고자명",
            "협력사코드", "협력사명", "상품타입",
            "매출과세구분", "매입과세구분",
            "주문유형", "주문일시", "발주일", "출하지시일", "입고일",
            "승인일", "배송완료일", "일일정산일", "정산시점",
            "계약번호", "공사명", "발주처", "공사시작일", "공사종료일",
            "법인", "그룹", "사이트", "VAT포함여부", "담당자부서",
            "배송형태", "VMI배송구분", "픽업배송구분", "IP/DIP",
            "주문자이메일", "고객사상품코드", "상품관리자",
            "배송지", "배송지상세주소", "송장번호", "배송메모",
            # ── 대사 컬럼 (배송메모 바로 뒤) ──────────────────────
            kt_amt_col, "플랫폼정산금액", "금액차이", "대사결과",
            # 공정명·지급상태·발행요청상태·발행요청일·결제유형·중분류 제외
        ]
        # 제외할 컬럼 목록 (배송메모 이후 불필요 컬럼)
        _DROP_COLS = {
            "공정명", "지급상태", "발행요청상태", "발행요청일", "결제유형",
            "중분류", "중분류(IP만변경적용)", "관리회계\n(IP만변경적용)",
            "관리회계(IP만변경적용)",
        }
        _seen = set()
        _ordered = []
        for _c in _DESIRED_ORDER:
            if _c in df.columns and _c not in _seen:
                _ordered.append(_c)
                _seen.add(_c)
        # _remaining: _DESIRED_ORDER에 없고 제외 목록·내부 컬럼(_접두사)도 아닌 것
        _remaining = [
            _c for _c in df.columns
            if _c not in _seen
            and _c not in _DROP_COLS
            and not str(_c).startswith("_")
        ]
        df = df[_ordered + _remaining]

        cols = list(df.columns)
        money_set = {
            kt_amt_col, "플랫폼정산금액", "금액차이",
            "정산금액", "주문금액", "매입금액", "매출총이익",
            "주문단가", "정산단가", "매입단가",
        }
        stats = getattr(self, "_inv_stats", {})

        # ── 헬퍼: 요약 헤더용 셀 서식 ──────────────────────────
        LBL_FILL = PatternFill("solid", start_color="DDEEFF")
        HDR_FILL = PatternFill("solid", start_color="F2F2F2")
        def _lbl(row, col, val):
            c = ws.cell(row, col, val)
            c.font      = Font(name=FONT, bold=True, size=10, color="1F4E79")
            c.fill      = LBL_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = _bdr()
        def _val(row, col, val, red=False, money=False, bg=None):
            c = ws.cell(row, col, val)
            c.font      = Font(name=FONT, bold=True, size=10,
                               color=("FF0000" if red else "000000"))
            c.fill      = PatternFill("solid", start_color=(bg or "F2F2F2"))
            c.alignment = Alignment(horizontal="right" if money else "center",
                                    vertical="center")
            c.border    = _bdr()
            if money:
                c.number_format = "#,##0"

        # ── 행 높이 (1~3행 요약 영역) ──────────────────────────
        for r in (1, 2, 3):
            ws.row_dimensions[r].height = 18

        # ── 1행: 정산 기간 / 실행일 / 건수 ───────────────────────
        n_kt  = stats.get("n_kt",  len(df))
        n_inv = stats.get("n_inv", len(df))
        cnt_mismatch = (n_kt != n_inv)
        _lbl(1, 1, "정산 기간");   _val(1, 2, f"{self.ps} ~ {self.pe}")
        _lbl(1, 3, "실행일");      _val(1, 4, str(datetime.date.today()))
        _lbl(1, 5, "KT 원본 건수"); _val(1, 6, f"{n_kt}건")
        _lbl(1, 7, "인보이스 건수"); _val(1, 8, f"{n_inv}건",
                                              red=cnt_mismatch)

        # ── 2행: 금액 합계 ──────────────────────────────────────
        kt_total   = stats.get("kt_total",   0.0)
        pl_total   = stats.get("pl_total",   0.0)
        diff_total = stats.get("diff_total", 0.0)
        _lbl(2, 1, "KT 정산금액 합계");  _val(2, 2, kt_total,   money=True)
        _lbl(2, 3, "플랫폼 금액 합계");  _val(2, 4, pl_total,   money=True)
        _lbl(2, 5, "금액 차이 합계");    _val(2, 6, diff_total, money=True,
                                               red=(abs(diff_total) > 1))

        # ── 3행: 대사 결과 건수 ─────────────────────────────────
        cnt_normal  = stats.get("cnt_normal",  0)
        cnt_missing = stats.get("cnt_missing", 0)
        cnt_return  = stats.get("cnt_return",  0)
        cnt_amount  = stats.get("cnt_amount",  0)
        cnt_gov     = stats.get("cnt_gov",     0)
        _lbl(3, 1, "정상");       _val(3, 2, f"{cnt_normal}건")
        _lbl(3, 3, "플랫폼누락"); _val(3, 4,
            f"{cnt_missing}건" + (" ← 확인 필요" if cnt_missing >= 1 else ""),
            red=(cnt_missing >= 1),
            bg=("FFCCCC" if cnt_missing >= 1 else "F2F2F2"))
        _lbl(3, 5, "반품");       _val(3, 6, f"{cnt_return}건")
        _lbl(3, 7, "금액불일치"); _val(3, 8,
            f"{cnt_amount}건",
            bg=("FFFF99" if cnt_amount >= 1 else "F2F2F2"))
        _lbl(3, 9, "국책과제");   _val(3, 10, f"{cnt_gov}건",
            bg=("FFF2CC" if cnt_gov >= 1 else "F2F2F2"))

        # ── 특이 주문 최상단 정렬 ────────────────────────────────
        # 플랫폼누락·반품·금액불일치 → 국책과제 → 정상 순으로 정렬
        _gov_col_inv = self.kt_cols.get("국책과제여부")
        _GOV_VALUES  = {"Y", "O", "YES", "국책", "TRUE", "1", "국책과제대상"}
        def _sort_priority(row):
            res = str(row.get("대사결과", ""))
            if res in ("플랫폼누락", "반품", "금액불일치"):
                return 0
            ma_val = str(row.get("관리회계", "")).strip()
            if ma_val == "국책과제":
                return 1
            if (_gov_col_inv and _gov_col_inv in row.index
                    and str(row[_gov_col_inv]).strip().upper() in _GOV_VALUES):
                return 1
            return 2
        df["_sort_key"] = df.apply(_sort_priority, axis=1)
        df = df.sort_values("_sort_key", kind="stable").drop(columns=["_sort_key"])
        df = df.reset_index(drop=True)

        # ── 4행: 컬럼 헤더 ──────────────────────────────────────
        for c, name in enumerate(cols, 1):
            hdr_cell(ws.cell(4, c, name))
        ws.row_dimensions[4].height = 20

        # ── 5행~: 데이터 ─────────────────────────────────────────
        for ri, (_, row) in enumerate(df.iterrows(), 5):
            res = row.get("대사결과", "")
            # 국책과제 여부 판단 (관리회계 컬럼 기준)
            is_gov = (str(row.get("관리회계", "")).strip() == "국책과제") or (
                _gov_col_inv and _gov_col_inv in row.index
                and str(row.get(_gov_col_inv, "")).strip().upper() in _GOV_VALUES
            )
            if res == "플랫폼누락":
                fill = ROW_MISSING
            elif res == "반품":
                fill = ROW_RETURN
            elif res == "금액불일치":
                fill = ROW_AMOUNT
            elif is_gov:
                fill = "FFF2CC"   # 연한 노랑 (국책과제)
            else:
                fill = ROW_EVEN if ri % 2 == 0 else None

            for c, col in enumerate(cols, 1):
                val = row[col]
                if pd.isna(val):
                    val = None
                cell = ws.cell(ri, c, val)
                data_cell(cell, fill=fill, money=(col in money_set))

        ws.freeze_panes = "A5"
        auto_col_width(ws)
        self.log.info(f"    → {len(df)}행")

    # ── 시트2: 매출현황 ────────────────────────────────────────
    def write_sales_sheet(self, wb):
        self.log.info("  [시트2] 매출현황 작성...")
        ws = wb.create_sheet("②매출현황")
        df = self.df_invoice.copy()

        # ── 컬럼 탐지 ──
        # 인보이스가 KT 기반으로 변경됨: 플랫폼 정산금액은 "플랫폼정산금액" 컬럼에 저장
        amt_col  = ("플랫폼정산금액" if "플랫폼정산금액" in df.columns
                    else self.pl_cols["정산금액"])
        buy_col  = self.pl_cols.get("매입금액")
        gp_col   = self.pl_cols.get("매출총이익")
        type_col = next((c for c in ["일반/통신구분", "일반통신구분"]
                         if c in df.columns), None)
        # 관리회계 컬럼 사용 (중분류(IP만변경적용) 미사용)
        mdiv_col = next((c for c in ["관리회계"] if c in df.columns), None)
        dept_col = next((c for c in ["담당자부서"] if c in df.columns), None)

        # ── 반품 행 NaN 플랫폼정산금액 → KT 금액으로 대체 (매출현황 합계 정확도) ──
        # 미확인 반품의 경우 플랫폼 금액이 NaN이므로 fillna(0) 시 합계에서 누락됨.
        # KT 공급가액(음수)을 플랫폼 금액 대신 사용하여 합계를 맞춘다.
        kt_amt_col_s = self.kt_cols["정산금액"]
        if ("대사결과" in df.columns and amt_col in df.columns
                and kt_amt_col_s in df.columns):
            _ret_nan_mask = (
                df["대사결과"].isin(["반품", "미확인"])
                & pd.to_numeric(df[amt_col], errors="coerce").isna()
            )
            if _ret_nan_mask.any():
                df.loc[_ret_nan_mask, amt_col] = pd.to_numeric(
                    df.loc[_ret_nan_mask, kt_amt_col_s], errors="coerce"
                )
                self.log.info(
                    f"  반품 {int(_ret_nan_mask.sum())}건: "
                    f"플랫폼 금액 미확인 → KT 금액(음수)으로 대체 집계"
                )

        # ── 수치 변환 ──
        for col in filter(None, [amt_col, buy_col, gp_col]):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # ── 반품 행 부호 반전: 매출현황은 NET 매출 기준 (반품 = 차감) ─────
        # 인보이스에서 절대값 처리된 반품 금액을 다시 음수로 변환
        if "대사결과" in df.columns:
            _ret_s_mask = df["대사결과"].isin(["반품"])
            if _ret_s_mask.any():
                for _nc in filter(None, [amt_col, buy_col, gp_col]):
                    if _nc in df.columns:
                        df.loc[_ret_s_mask, _nc] = -df.loc[_ret_s_mask, _nc].abs()
                self.log.info(
                    f"  반품 {int(_ret_s_mask.sum())}건: "
                    f"매출현황 음수 처리 (매출금액·매입금액·매출총이익 차감)"
                )

        # ── 통신/일반 분리 ──
        if type_col:
            df_t = df[df[type_col].astype(str).str.strip() == "통신"].copy()
            df_g = df[df[type_col].astype(str).str.strip() == "일반"].copy()
        else:
            df_t = pd.DataFrame(columns=df.columns)
            df_g = df.copy()

        # ── 중분류별 집계 ──
        def _agg_by_mdiv(sub_df):
            if sub_df.empty or not mdiv_col or mdiv_col not in sub_df.columns:
                return pd.DataFrame(columns=["중분류", "정산금액", "매입금액", "매출총이익"])
            agg = {"정산금액": (amt_col, "sum")}
            if buy_col and buy_col in sub_df.columns:
                agg["매입금액"] = (buy_col, "sum")
            if gp_col and gp_col in sub_df.columns:
                agg["매출총이익"] = (gp_col, "sum")
            result = sub_df.groupby(mdiv_col).agg(**agg).reset_index()
            result.rename(columns={mdiv_col: "중분류"}, inplace=True)
            if "매입금액" not in result.columns:
                result["매입금액"] = 0
            if "매출총이익" not in result.columns:
                result["매출총이익"] = 0
            return result.sort_values("정산금액", ascending=False).reset_index(drop=True)

        grp_t = _agg_by_mdiv(df_t)
        grp_g = _agg_by_mdiv(df_g)

        # ── 열 너비 ──
        for col_letter, width in [("A", 26), ("B", 16), ("C", 16), ("D", 16),
                                    ("E", 2),  ("F", 22), ("G", 16), ("H", 16), ("I", 16)]:
            ws.column_dimensions[col_letter].width = width

        # ── 섹션 작성 함수 (통신/일반 공통) ──
        def _write_section(start_row, section_name, grp_df, col_base=1):
            r = start_row
            # 섹션 제목
            title = ws.cell(r, col_base, f"▶ {section_name}")
            title.font = Font(name=FONT, bold=True, size=12, color=HDR_BLUE)
            ws.merge_cells(start_row=r, start_column=col_base,
                           end_row=r, end_column=col_base + 3)
            ws.row_dimensions[r].height = 22
            r += 1
            # 헤더
            for ci, h in enumerate(["관리회계", "매출금액", "매입금액", "매출총이익"], col_base):
                hdr_cell(ws.cell(r, ci, h))
            ws.row_dimensions[r].height = 18
            data_start = r + 1
            r += 1
            # 데이터 행
            if grp_df.empty:
                c0 = ws.cell(r, col_base, "(데이터 없음)")
                c0.font = Font(name=FONT, size=10, color="808080", italic=True)
                c0.border = _bdr()
                r += 1
            else:
                for offset, (_, dr) in enumerate(grp_df.iterrows()):
                    ri = r + offset
                    fill = ROW_EVEN if ri % 2 == 0 else None
                    c0 = ws.cell(ri, col_base, str(dr.get("중분류", "")))
                    c0.font = Font(name=FONT, size=10)
                    c0.border = _bdr()
                    c0.alignment = Alignment(horizontal="left", vertical="center")
                    if fill:
                        c0.fill = PatternFill("solid", start_color=fill)
                    for ci_off, col_name in enumerate(["정산금액", "매입금액", "매출총이익"], 1):
                        v = dr.get(col_name, 0)
                        cell = ws.cell(ri, col_base + ci_off, float(v) if pd.notna(v) else 0.0)
                        data_cell(cell, fill=fill, money=True)
                r += len(grp_df)
            data_end = r - 1
            # 합계 행 (SUM 수식)
            total_cell(ws.cell(r, col_base, f"{section_name} 총액"))
            ws.cell(r, col_base).alignment = Alignment(horizontal="left", vertical="center")
            for ci_off in range(1, 4):
                cl = get_column_letter(col_base + ci_off)
                tc = ws.cell(r, col_base + ci_off, f"=SUM({cl}{data_start}:{cl}{data_end})")
                total_cell(tc, money=True)
            ws.row_dimensions[r].height = 18
            return r + 2   # 빈 행 포함

        # ── 좌측 블록 (A:D): 통신 → 일반 ──
        row = 1
        row = _write_section(row, "통신", grp_t, col_base=1)
        row = _write_section(row, "일반", grp_g, col_base=1)

        # ── 우측 블록 (F:I): 검증 요약 + 담당자부서별 ──
        RC = 6   # F열 시작

        # ── KT 금액 집계: df_invoice 기준 (로그 검증과 동일 소스) ──────
        kt_amt_col = self.kt_cols["정산금액"]
        kt_gov_col = self.kt_cols.get("국책과제여부")
        inv        = self.df_invoice.copy() if self.df_invoice is not None else pd.DataFrame()

        def _inv_kt_sum(mask):
            s = pd.to_numeric(inv.loc[mask, kt_amt_col], errors="coerce").fillna(0)
            return float(s.sum()), int(mask.sum())

        # 반품 행
        ret_mask = inv["대사결과"].isin(["반품", "미확인"]) if "대사결과" in inv.columns \
                   else pd.Series(False, index=inv.index)
        kt_ret_total, kt_ret_cnt = _inv_kt_sum(ret_mask)

        # 국책과제 행 (반품 제외)
        kt_gov_total, kt_gov_cnt = 0.0, 0
        if kt_gov_col and kt_gov_col in inv.columns:
            gov_mask = (~ret_mask) & inv[kt_gov_col].astype(str).str.upper().isin(
                ["Y", "O", "YES", "국책", "TRUE", "1", "국책과제대상"])
            kt_gov_total, kt_gov_cnt = _inv_kt_sum(gov_mask)
        else:
            gov_mask = pd.Series(False, index=inv.index)

        # 일반 행 (반품·국책 제외)
        kt_normal_total, kt_normal_cnt = _inv_kt_sum(~ret_mask & ~gov_mask)
        # NET 기준: 반품은 차감 (KT도 반품을 별도 양수로 기록 → 매출현황에서는 차감 처리)
        kt_total = kt_normal_total + kt_gov_total - kt_ret_total

        # 플랫폼 합계 (미매칭 NaN → 0)
        pl_t_total = float(df_t[amt_col].fillna(0).sum()) if not df_t.empty else 0.0
        pl_g_total = float(df_g[amt_col].fillna(0).sum()) if not df_g.empty else 0.0
        pl_total   = pl_t_total + pl_g_total

        diff     = pl_total - kt_total   # 양수: 플랫폼 > KT, 음수: KT > 플랫폼
        verified = abs(diff) <= 1

        right_row = 1

        # ── 검증 요약 제목 ──
        t = ws.cell(right_row, RC, "▶ 정산금액 합계 검증")
        t.font = Font(name=FONT, bold=True, size=12, color=HDR_BLUE)
        ws.merge_cells(start_row=right_row, start_column=RC,
                       end_row=right_row, end_column=RC + 2)
        ws.row_dimensions[right_row].height = 22
        right_row += 1

        # 헤더
        for ci, h in enumerate(["구분", "금액", "건수"], RC):
            hdr_cell(ws.cell(right_row, ci, h), color=HDR_GREEN)
        right_row += 1

        # 상세 행 작성 헬퍼
        def _vrow(label, amt, cnt=None, bold=False, fill=None, is_result=False, result_ok=True):
            nonlocal right_row
            lc = ws.cell(right_row, RC, label)
            lc.font = Font(name=FONT, size=10, bold=bold)
            lc.border = _bdr()
            lc.alignment = Alignment(horizontal="left", vertical="center")
            if fill:
                lc.fill = PatternFill("solid", start_color=fill)

            if is_result:
                vc = ws.cell(right_row, RC + 1, amt)
                vc.font  = Font(name=FONT, size=10, bold=True,
                                color=("375623" if result_ok else "C00000"))
                vc.fill  = PatternFill("solid",
                                       start_color=("CCFFCC" if result_ok else "FFCCCC"))
                vc.border = _bdr()
                vc.alignment = Alignment(horizontal="center", vertical="center")
                ws.merge_cells(start_row=right_row, start_column=RC+1,
                               end_row=right_row, end_column=RC+2)
            else:
                vc = ws.cell(right_row, RC + 1, float(amt))
                data_cell(vc, money=True, fill=fill)
                if cnt is not None:
                    cc = ws.cell(right_row, RC + 2, f"{cnt:,}건")
                    cc.font = Font(name=FONT, size=10)
                    cc.border = _bdr()
                    cc.alignment = Alignment(horizontal="center", vertical="center")
                    if fill:
                        cc.fill = PatternFill("solid", start_color=fill)
            right_row += 1

        HDR_LIGHT = "EEF4FF"
        _vrow("【KT 정산금액】", kt_total, kt_normal_cnt+kt_gov_cnt+kt_ret_cnt,
              bold=True, fill=HDR_LIGHT)
        _vrow("  일반 정산",    kt_normal_total, kt_normal_cnt)
        if kt_gov_cnt > 0:
            _vrow("  국책과제",  kt_gov_total,   kt_gov_cnt,   fill="FFF2CC")
        _vrow("  반품(차감)",   -kt_ret_total,   kt_ret_cnt)
        right_row += 1  # 빈 행

        _vrow("【플랫폼 정산금액】", pl_total, len(df_t)+len(df_g),
              bold=True, fill=HDR_LIGHT)
        _vrow("  통신",         pl_t_total,      len(df_t))
        _vrow("  일반",         pl_g_total,      len(df_g))
        right_row += 1  # 빈 행

        result_text = "TRUE (일치)" if verified else f"FALSE (차이: {diff:+,.0f}원)"
        _vrow("검증 (KT = 플랫폼)", result_text, bold=True,
              is_result=True, result_ok=verified)
        right_row += 1   # 빈 행

        # ── 담당자부서별 집계 (통신) ──────────────────────────────
        # 구조: 네트워크사업1팀(전체) → 투자형구축물자/국책과제/유지보수(소계) → 합계
        t2 = ws.cell(right_row, RC, "▶ 담당자부서별 (통신)")
        t2.font = Font(name=FONT, bold=True, size=12, color=HDR_BLUE)
        ws.merge_cells(start_row=right_row, start_column=RC,
                       end_row=right_row, end_column=RC + 3)
        ws.row_dimensions[right_row].height = 22
        right_row += 1

        for ci, h in enumerate(["구분", "매출금액", "매입금액", "매출총이익"], RC):
            hdr_cell(ws.cell(right_row, ci, h), color=HDR_GREEN)
        right_row += 1

        # 전체 통신 합계 = 네트워크사업1팀
        _t_amt = float(pd.to_numeric(df_t[amt_col], errors="coerce").sum()) if not df_t.empty else 0.0
        _t_buy = float(pd.to_numeric(df_t[buy_col], errors="coerce").sum()) if (buy_col and not df_t.empty and buy_col in df_t.columns) else 0.0
        _t_gp  = float(pd.to_numeric(df_t[gp_col],  errors="coerce").sum()) if (gp_col  and not df_t.empty and gp_col  in df_t.columns) else 0.0

        # 행1: 네트워크사업1팀 (전체 통신)
        c0 = ws.cell(right_row, RC, "네트워크사업1팀")
        c0.font = Font(name=FONT, size=10, bold=True)
        c0.border = _bdr()
        c0.alignment = Alignment(horizontal="left", vertical="center")
        for ci_off, v in enumerate([_t_amt, _t_buy, _t_gp], 1):
            data_cell(ws.cell(right_row, RC + ci_off, v), money=True, bold=True)
        right_row += 1

        # 행2~4: 투자형구축물자·국책과제·유지보수 (소계, 값 0이어도 항상 표시)
        SPECIAL_MA = ["투자형구축물자", "국책과제", "유지보수"]
        _ma_col = next((c for c in ["관리회계"] if c in df_t.columns), None) if not df_t.empty else None

        for offset, _ma_val in enumerate(SPECIAL_MA):
            if not df_t.empty and _ma_col:
                _sub = df_t[df_t[_ma_col].astype(str).str.strip() == _ma_val]
                _s_amt = float(pd.to_numeric(_sub[amt_col], errors="coerce").sum()) if not _sub.empty else 0.0
                _s_buy = float(pd.to_numeric(_sub[buy_col], errors="coerce").sum()) if (not _sub.empty and buy_col and buy_col in _sub.columns) else 0.0
                _s_gp  = float(pd.to_numeric(_sub[gp_col],  errors="coerce").sum()) if (not _sub.empty and gp_col  and gp_col  in _sub.columns) else 0.0
            else:
                _s_amt, _s_buy, _s_gp = 0.0, 0.0, 0.0
            ri   = right_row + offset
            fill = ROW_EVEN if ri % 2 == 0 else None
            c0   = ws.cell(ri, RC, f"  {_ma_val}")
            c0.font      = Font(name=FONT, size=10, italic=True)
            c0.border    = _bdr()
            c0.alignment = Alignment(horizontal="left", vertical="center")
            if fill:
                c0.fill = PatternFill("solid", start_color=fill)
            for ci_off, v in enumerate([_s_amt, _s_buy, _s_gp], 1):
                data_cell(ws.cell(ri, RC + ci_off, v), fill=fill, money=True)
        right_row += len(SPECIAL_MA)

        # 합계 행 (= 네트워크사업1팀 전체 합계와 동일)
        total_cell(ws.cell(right_row, RC, "합계"))
        ws.cell(right_row, RC).alignment = Alignment(horizontal="left", vertical="center")
        for ci_off, v in enumerate([_t_amt, _t_buy, _t_gp], 1):
            total_cell(ws.cell(right_row, RC + ci_off, v), money=True)

        ws.freeze_panes = "A2"
        self.log.info(f"    → 통신 {len(grp_t)}개 중분류 / 일반 {len(grp_g)}개 중분류")

    # ── 시트3: 어음확인_최종 ───────────────────────────────────
    def write_bill_sheet(self, wb):
        self.log.info("  [시트3] 어음확인_최종 작성...")
        # 기업규모 판별 소스 상태 로그
        _mc_cnt = len(self.master.midcorp_set) if self.master else 0
        _lc_cnt = len(_LARGE_CORP_SET)
        _mn_cnt = len(self.master.corp_size)   if self.master else 0
        self.log.info(f"    기업규모 소스: 대기업(코드) {_lc_cnt}개 / "
                      f"중견기업(목록) {_mc_cnt}개 / 수동등록 {_mn_cnt}개")
        # ── 중견기업 진단: 실제 협력사명과 목록 매칭 확인 ──────────────────────────
        if self.master and _mc_cnt > 0 and self.df_invoice is not None:
            _corp_col_diag = next(
                (c for c in ["협력사명", "공급자명"] if c in self.df_invoice.columns), None
            )
            _corps = (
                self.df_invoice[_corp_col_diag].dropna().unique().tolist()
                if _corp_col_diag else []
            )
            _mc_hits = []
            _mc_miss = []
            for _cn in _corps[:30]:  # 최대 30개만 진단
                _norm = _normalize_corp_name(str(_cn))
                if _norm in self.master.midcorp_set:
                    _mc_hits.append(_cn)
                else:
                    _mc_miss.append(f"{_cn}→{_norm}")
            if _mc_hits:
                self.log.info(f"    중견기업 매칭 성공({len(_mc_hits)}개): {_mc_hits}")
            else:
                self.log.info(f"    중견기업 매칭 없음. 미매칭 예시(최대10): {_mc_miss[:10]}")
                # 목록 샘플 출력
                _sample = list(self.master.midcorp_set)[:5]
                self.log.info(f"    중견기업목록 정규화 샘플: {_sample}")
        ws = wb.create_sheet("③어음확인_최종")
        df = self.df_invoice.copy()

        def _find_col(pl_key, kt_key=None):
            pl_c = self.pl_cols.get(pl_key)
            if pl_c and pl_c in df.columns:
                return pl_c
            kt_c = self.kt_cols.get(kt_key or pl_key)
            if kt_c and kt_c in df.columns:
                return kt_c
            return pl_c  # 원래 오류 동작 유지

        tax_col  = _find_col("매출과세구분")   # KT: 세금코드명
        corp_col = _find_col("협력사명")        # KT: 공급자명
        ord_col  = _find_col("주문번호")        # KT: 구매문서번호
        buy_col  = _find_col("매입금액")        # 어음 기준금액

        if not tax_col or not corp_col:
            ws["A1"] = "매출과세구분 또는 협력사명 컬럼 없음"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT)
            return
        if not buy_col or buy_col not in df.columns:
            ws["A1"] = "⚠ 매입금액 컬럼 없음 — 어음확인 불가"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT, size=11)
            return

        df[buy_col] = pd.to_numeric(df[buy_col], errors="coerce").fillna(0)

        if not self.master:
            ws["A1"] = "⚠ 매핑파일 없음 — 기업규모 판별 불가"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT, size=11)
        elif not self.master.loaded and not self.master.midcorp_set:
            ws["A1"] = "⚠ mapping_master.xlsx 및 중견기업목록.xlsx 없음 — 기업규모 판별 불가"
            ws["A1"].font = Font(color="FF0000", bold=True, name=FONT, size=11)

        # 과세 대상만 추출 (어음 발행 대상 = 과세 거래)
        df_tx = df[df[tax_col].astype(str).str.strip() == "과세"].copy()
        if df_tx.empty:
            ws["A1"] = "과세 주문 없음"
            return

        # (주문번호, 협력사명) 기준 합산 (매입금액 기준)
        grp = df_tx.groupby([ord_col, corp_col]).agg(
            매입금액합계 = (buy_col, "sum"),
            대사결과     = ("대사결과", "first"),
        ).reset_index()
        grp.rename(columns={ord_col: "주문번호", corp_col: "협력사명"}, inplace=True)

        # 어음발행금액 = 매입금액합계 × 1.1 (VAT 포함)
        grp["VAT포함합계"] = grp["매입금액합계"] * 1.1

        # 기업규모 매핑
        # ① mapping_master.xlsx 수동 등록(대기업 등) → ② 중견기업목록.xlsx 자동 매칭 → ③ 확인필요
        grp["기업규모"] = grp["협력사명"].apply(
            lambda n: self.master.get_corp_size(n) if self.master else "확인필요"
        )

        # 3조건 판별: ① 매입금액×VAT 2억 이상  ② 대기업/중견기업  ③ 대사 정상
        grp["cond1"] = grp["VAT포함합계"] >= 200_000_000
        grp["cond2"] = grp["기업규모"].isin(["중견기업", "대기업"])
        grp["cond3"] = grp["대사결과"] == "정상"
        grp["어음대상"] = grp.apply(
            lambda r: "발행대상"  if (r.cond1 and r.cond2 and r.cond3)
                 else ("확인필요" if (r.cond1 and (not r.cond3 or r["기업규모"] == "확인필요"))
                 else "미해당"),
            axis=1
        )
        grp["비고"] = grp.apply(
            lambda r: "" if r["어음대상"] == "발행대상"
                      else ("대사 미확인" if r["대사결과"] != "정상"
                            else ("기업규모 미확인" if r["기업규모"] == "확인필요"
                                  else ("매입금액미달(VAT포함)" if not r.cond1
                                        else "규모조건 미충족"))),
            axis=1
        )
        grp.drop(columns=["cond1","cond2","cond3"], inplace=True)
        grp.sort_values("VAT포함합계", ascending=False, inplace=True)

        # 발행 대상 건수 상단 표시
        n_issue = (grp["어음대상"] == "발행대상").sum()
        ws.merge_cells("A1:G1")
        ws["A1"] = (f"어음 발행 대상: {n_issue}건"
                    f"   ※ 과세 + 매입금액(VAT포함) 2억원 이상 + 중견/대기업 + 대사정상"
                    f"   ※ 기업규모 미확인 시 mapping_master.xlsx 에 등록 필요")
        ws["A1"].font      = Font(name=FONT, bold=True, size=11,
                                  color=("375623" if n_issue > 0 else "808080"))
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        COLS = ["주문번호", "협력사명", "매입금액합계", "VAT포함합계", "기업규모", "어음대상", "비고"]
        for c, h in enumerate(COLS, 1):
            hdr_cell(ws.cell(2, c, h))

        fill_map = {"발행대상": "CCFFCC", "확인필요": "FFFF99"}
        for r, (_, row) in enumerate(grp[COLS].iterrows(), 3):
            f = fill_map.get(row["어음대상"])
            for c, col in enumerate(COLS, 1):
                cell = ws.cell(r, c, row[col])
                data_cell(cell, fill=f,
                          money=(col in ("매입금액합계", "VAT포함합계")))

        set_col_widths = [22, 20, 18, 18, 12, 12, 22]
        for i, w in enumerate(set_col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A3"

    # ── 시트4: 재무회계팀 자금예측보고 ────────────────────────
    def write_forecast_sheet(self, wb):
        self.log.info("  [시트4] 재무회계팀 자금예측보고 작성...")
        ws = wb.create_sheet("④자금예측보고")
        df = self.df_invoice.copy()
        amt_col  = ("플랫폼정산금액" if "플랫폼정산금액" in df.columns
                    else self.pl_cols["정산금액"])

        def _find_col(pl_key, kt_key=None):
            pl_c = self.pl_cols.get(pl_key)
            if pl_c and pl_c in df.columns:
                return pl_c
            kt_c = self.kt_cols.get(kt_key or pl_key)
            if kt_c and kt_c in df.columns:
                return kt_c
            return pl_c

        tax_col = _find_col("매출과세구분")   # KT: 세금코드명
        buy_col = _find_col("매입금액")        # 매입금액 (지출 기준)
        ps_s = self.ps.strftime("%Y.%m.%d") if self.ps else "-"
        pe_s = self.pe.strftime("%Y.%m.%d") if self.pe else "-"

        # 제목
        ws.merge_cells("A1:G1")
        ws["A1"] = f"재무회계팀  자금예측보고    정산기간: {ps_s} ~ {pe_s}"
        ws["A1"].font      = Font(name=FONT, bold=True, size=14, color="1F4E79")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        if not tax_col:
            ws["A3"] = "매출과세구분 컬럼 없음"
            return

        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        if tax_col not in df.columns:
            df[tax_col] = "미분류"
        if buy_col and buy_col in df.columns:
            df[buy_col] = pd.to_numeric(df[buy_col], errors="coerce").fillna(0)

        # 과세구분 값 목록: 과세 → 면세 → 비과세 순으로 정렬
        all_tax_vals = df[tax_col].dropna().unique().tolist()
        preferred = ["과세", "면세", "비과세"]
        tax_vals  = [v for v in preferred if v in all_tax_vals] + \
                    [v for v in all_tax_vals if v not in preferred]
        n_tax = len(tax_vals)

        # 섹션 제목
        ws.cell(3, 1, "▶ 자금 예측 요약").font = Font(name=FONT, bold=True,
                                                       size=11, color="1F4E79")

        # 헤더 행
        hdr_row = 4
        hdr_cell(ws.cell(hdr_row, 1, "구분"))
        for ci, tv in enumerate(tax_vals, 2):
            hdr_cell(ws.cell(hdr_row, ci, f"{tv}(VAT포함)" if tv == "과세" else tv))
        hdr_cell(ws.cell(hdr_row, n_tax + 2, "합계"))
        ws.row_dimensions[hdr_row].height = 18

        data_start = hdr_row + 1   # 5행

        # ── 행1: 현금 입금예상액(매출금액) ───────────────────────
        row1 = data_start
        c1 = ws.cell(row1, 1, "현금 입금예상액(매출금액)")
        c1.font      = Font(name=FONT, size=10, bold=True)
        c1.border    = _bdr()
        c1.alignment = Alignment(horizontal="left", vertical="center")
        for ci, tv in enumerate(tax_vals, 2):
            val = float(df[df[tax_col] == tv][amt_col].sum())
            if tv == "과세":
                val = val * 1.1   # 실제 현금흐름: 공급가액 × 1.1 (VAT 포함)
            data_cell(ws.cell(row1, ci, val), money=True)
        sr1 = f"{get_column_letter(2)}{row1}:{get_column_letter(n_tax + 1)}{row1}"
        total_cell(ws.cell(row1, n_tax + 2, f"=SUM({sr1})"), money=True)

        # ── 행2: 현금 지출예상액(매입금액) ───────────────────────
        row2 = data_start + 1
        c2 = ws.cell(row2, 1, "현금 지출예상액(매입금액)")
        c2.font      = Font(name=FONT, size=10, bold=True)
        c2.border    = _bdr()
        c2.alignment = Alignment(horizontal="left", vertical="center")
        for ci, tv in enumerate(tax_vals, 2):
            if buy_col and buy_col in df.columns:
                val = float(df[df[tax_col] == tv][buy_col].sum())
                if tv == "과세":
                    val = val * 1.1   # 실제 현금흐름: 매입가액 × 1.1 (VAT 포함)
            else:
                val = 0.0
            data_cell(ws.cell(row2, ci, val), money=True)
        sr2 = f"{get_column_letter(2)}{row2}:{get_column_letter(n_tax + 1)}{row2}"
        total_cell(ws.cell(row2, n_tax + 2, f"=SUM({sr2})"), money=True)

        # ── 차액 행 (입금 - 지출) ─────────────────────────────────
        # ※ 합계 행(매출+매입 합산) 삭제 — 불필요
        diff_row = data_start + 2
        DIFF_FILL = "E2EFDA"   # 연한 초록 (이익 강조)
        dc = ws.cell(diff_row, 1, "현금 순이익(입금 - 지출)")
        dc.font      = Font(name=FONT, size=10, bold=True, color="375623")
        dc.fill      = PatternFill("solid", start_color=DIFF_FILL)
        dc.border    = _bdr()
        dc.alignment = Alignment(horizontal="left", vertical="center")
        for ci in range(2, n_tax + 3):
            col_l = get_column_letter(ci)
            cell = ws.cell(diff_row, ci,
                           f"={col_l}{row1}-{col_l}{row2}")
            cell.font         = Font(name=FONT, size=10, bold=True, color="375623")
            cell.fill         = PatternFill("solid", start_color=DIFF_FILL)
            cell.border       = _bdr()
            cell.alignment    = Alignment(horizontal="right", vertical="center")
            cell.number_format = MONEY_FMT
        ws.row_dimensions[diff_row].height = 18

        auto_col_width(ws)
        ws.freeze_panes = "A5"

    # ── 메인 실행 ────────────────────────────────────────────────
    def run(self):
        try:
            self.detect_files()
            self.step0_init()
            self.step0b_validate_tax()   # 과세구분 검증 (대사 전 선행)
            self.step1_filter()
            self.step2_separate_returns()
            self.step3_reconcile_normal()
            self.step4_reconcile_returns()

            self.log.info("\n[XLOOKUP] 매핑 적용 중...")
            self.build_invoice_df()

            today_str   = self.today.strftime("%Y%m%d")
            result_path = str(self.base / f"정산결과_{today_str}.xlsx")

            self.log.info(f"\n[출력] {result_path}")
            wb = openpyxl.Workbook()
            self.write_tax_validation_sheet(wb)   # ⓪ 과세구분검증 — 첫 번째 시트
            self.write_invoice_sheet(wb)
            self.write_sales_sheet(wb)
            self.write_bill_sheet(wb)
            self.write_forecast_sheet(wb)

            # 저장: 파일이 열려 있으면 _1, _2 ... 순서로 대체 파일명 시도
            _saved = False
            for _suffix in [""] + [f"_{i}" for i in range(1, 10)]:
                _try_path = str(self.base / f"정산결과_{today_str}{_suffix}.xlsx")
                try:
                    wb.save(_try_path)
                    result_path = _try_path
                    _saved = True
                    break
                except PermissionError:
                    self.log.warning(f"  저장 실패 (파일 열림): {_try_path}")
            if not _saved:
                raise PermissionError(
                    f"정산결과 파일을 저장할 수 없습니다. "
                    f"Excel에서 '정산결과_{today_str}*.xlsx' 파일을 모두 닫은 후 재실행하세요."
                )

            # ─ 콘솔 최종 요약 ─
            n_miss = len(self.df_missing)  if self.df_missing  is not None else 0
            n_ok   = len(self.df_matched)  if self.df_matched   is not None else 0
            n_diff = len(self.df_amtdiff)  if self.df_amtdiff  is not None else 0
            n_rok  = len(self.df_ret_ok)   if self.df_ret_ok   is not None else 0
            n_rng  = len(self.df_ret_ng)   if self.df_ret_ng   is not None else 0

            miss_flag = "  ← 확인 필요!" if n_miss > 0 else "  (없음)"
            self.log.info("")
            self.log.info("=" * 60)
            self.log.info("  실행 완료")
            n_pl = len(self.df_pl) if self.df_pl is not None else 0
            self.log.info(f"  정산 기간    : {self.ps} ~ {self.pe}")
            self.log.info(f"  KT 입고 건수 : {len(self.df_kt)}건  (구매문서번호 기준)")
            self.log.info(f"  플랫폼 건수  : {n_pl}건  (인보이스 기준)")
            self.log.info(f"  정상 매칭    : {n_ok}건")
            self.log.info(f"  플랫폼 누락  : {n_miss}건{miss_flag}")
            self.log.info(f"  반품         : 매칭 {n_rok}건 / 미확인 {n_rng}건  [경로 {self.ret_route}]")
            self.log.info(f"  금액 불일치  : {n_diff}건")
            if self.map_stats:
                for k, (ok, ng) in self.map_stats.items():
                    self.log.info(f"  {k:<12}: {ok}건 성공 / {ng}건 실패")
            self.log.info(f"  결과 파일    : {result_path}")
            self.log.info("=" * 60)

        except Exception as e:
            self.log.error(f"오류: {e}", exc_info=True)
            raise

# ── main ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    SettlementRunner(base_dir=".").run()
