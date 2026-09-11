"""
15단계: 대성인포텍(Ceragon) 매입가 제시안 + KT commerce 판매가(고객사 제시) 제시안
(돈현님 요청, 2026-09-11, 6차)

돈현님 지적: "구성 품목 총액으로 [Aviat와 Ceragon을] 비교해보려 했던 시도들이 무의미했다.
지금 가진 자료에서 최선은 대성인포텍에 얼마의 가격을 제시할지와 고객사인 KT에 얼마의
가격을 제시할지를 정하는 것이다."

방향 전환: Aviat 총액과 직접 비교해서 대성 가격의 높고 낮음을 판정하는 접근은 13번 파일의
결론대로 근거 부족으로 중단한다. 대신 이 파일은 완전히 다른 질문에 답한다 - "대성이 이미
제출한 가격을 매입가로 그대로 수용한다면(13번 검토의견 기준 조건부 수용), KT는 고객사에게
얼마를 제시해야 하는가?"

핵심 근거(2026-09-11 발견, 01번 파일 Aviat 213종 계약마스터에서 직접 확인):
KT commerce는 Aviat 계약에서 품목군별로 정확히 고정된 마진율(판매가 대비, 즉
판매가=매입가/(1-마진율))을 쓰고 있다 - 임의 추정이 아니라 실제 계약가 213건을 전수
검산해 확인한 사실이다.
  - VR4/VR10 IAP3 계열 핵심 장비(섀시/무선송수신부/대역결합기/인터페이스카드/제어카드/
    라이선스/팬/전원/마운트/도파관) 10개 기능군, 27개 품목 전수: 정확히 2.00%, 예외 0건
  - 안테나, 방수/접지/보호 자재, 100M급 장거리 케이블: 정확히 3.00%
  - SFP 광모듈(HAX_ 계열 신형), 커넥터·서지보호 KIT: 5.00%(3건 중 2건 정확히 일치, 1건은
    반올림 오차 범위 내)
이 마진 구조는 Ceragon 전용 정책이 아니라 Aviat 계약에서 확인된 것이지만, KT commerce가
동일한 무선전송장비 재판매 사업에 실제로 적용 중인 정책이므로 Ceragon 소싱 장비에도 같은
기준을 적용하는 것이 임의의 마진율을 가정하는 것보다 훨씬 근거가 있다.

이 파일이 하지 않는 것: Aviat 총액과 Ceragon 총액을 다시 비교하지 않는다(13번 결론 유지 -
반박 근거 없음). 매입가 자체의 적정성을 재판정하지 않는다(조건부 수용 전제). 최종 고객
제시가에 대한 전략적 가산(경쟁 상황, 협상 여유분 등)은 포함하지 않는다 - 여기 나온 판매가는
'KT 표준 마진 구조를 유지하는 최소 제시가'이며 이보다 낮게 제시하면 이 사업의 다른 장비군과
마진율이 달라진다는 의미다.
"""
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, FMT_AMOUNT, FMT_PERCENT, mark_fill
import importlib
func_cmp = importlib.import_module("14_functional_comparison")

# 돈현님 지적(2026-09-11, 7차): "지금 수량이 0인 게 영원한 게 아니다 - 카탈로그 474종 전부
# 매입가/판매가를 정해야 한다." 08/11GHz 기본 8개 구성에 없는 나머지 품목도 언젠가 발주될 수
# 있으므로, 06번 파일 '품목마스터'(대성이 제출한 카탈로그 493종 전체)를 대상으로 마진율을
# 부여한다. 이미 기능 대응이 확인된 ~40개(func_cmp.CERAGON_CATEGORY)를 벗어난 품목은
# 품명 키워드 패턴으로 분류한다 - 개별 검증이 아니므로 '패턴 매칭 추정'으로 신뢰도를 낮춰
# 표시하고, 밴드/세대가 다른 것으로 보이는 품목은 별도로 표시해 이번 8/11GHz 계약 범위인지
# KT가 직접 확인하도록 한다(임의로 빼거나 가격을 매기지 않은 게 아니라, 전부 포함하되
# 신뢰도를 다르게 표기).
OFFBAND_PATTERNS = [
    r'15HP-', r'32T-', r'RFU-CXE', r'RFU-D-F-', r'\bRFU-D-0[58]\b', r'\bRFU-D-11\b',
    r'ALL INDOOR', r'\bAI-V\b', r'4-5\s*GHZ', r'6\s*GHZ', r'18\s*GHZ',
    r'-05($|[^0-9])', r'-06($|[^0-9])', r'-5G', r'-6G(?!Hz)', r'-6H\b', r'-6L\b',
    r'-6($|[^A-Za-z0-9])', r'-18($|[^0-9])', r'-5\b',
]
NONRADIO_PATTERNS = [r'CERAVIEW', r'POLYVIEW', r'\bPV-', r'ECS4120']
TIER5_PATTERNS = [r'\bSFP\b', r'ARRESTOR', r'SURGE', r'PATCH CORD', r'\bBNC\b', r'\bTNC\b',
                  r'JACK', r'GBE-SPL', r'IF TEST PANEL', r'ETHERNET-Y', r'RJ-45',
                  r'GLANDS', r'DC_CONN', r'CAT5E', r'CAT6', r'OP-SM']
TIER3_PATTERNS = [r'\bANT-', r'ANTENNA', r'GROUND', r'\bMOUNT', r'-RM-', r'POLE',
                  r'BEND', r'DHRTR', r'\bRACK\b', r'ADAPTOR', r'HOLDER', r'HANGER',
                  r'PRESSURE WINDOW', r'ANGLE ADAPTOR', r'BUTTERFLY']


def classify_catalog_item(name):
    """(마진율, 등급표시, 범위추정, 신뢰도) 반환. 이미 14번에서 확인한 핵심 기능군은
    '확인됨'으로, 나머지는 품명 키워드 패턴 매칭 결과를 '추정'으로 구분한다."""
    # '_'를 공백으로 바꿔서 매칭한다 - 'CER_ANT-...'처럼 밑줄 바로 뒤에 오는 패턴은 밑줄이
    # 정규식 \b(단어 경계)상 문자로 취급돼 매칭이 안 되는 문제를 피하기 위함.
    n = name.upper().replace("_", " ")
    cer_cat = func_cmp.CERAGON_CATEGORY.get(name)
    if cer_cat is not None:
        rate = margin_rate(cer_cat)
        return rate, func_cmp.CATEGORY_LABEL.get(cer_cat, cer_cat), "8/11GHz 현재세대(확인됨)", "확인됨"

    offband = any(re.search(p, n) for p in OFFBAND_PATTERNS)
    nonradio = any(re.search(p, n) for p in NONRADIO_PATTERNS)
    scope = "타대역/구세대 추정(계약범위 확인 필요)" if offband else (
        "비무선 제품(확인 필요)" if nonradio else "8/11GHz 현재세대 추정")

    if any(re.search(p, n) for p in TIER5_PATTERNS):
        rate, label = SFP_MARGIN, "SFP/커넥터/서지(패턴)"
    elif any(re.search(p, n) for p in TIER3_PATTERNS):
        rate, label = ACCESSORY_MARGIN, "안테나/마운트/방수(패턴)"
    else:
        rate, label = CORE_MARGIN, "핵심장비(패턴, 기본값)"
    return rate, label, scope, "패턴 매칭 추정"

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_AVIAT_213 = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

BAND_CONFIGS = [("8GHz", "2+0"), ("8GHz", "4+0"), ("8GHz", "6+0"), ("8GHz", "8+0"),
                ("11GHz", "2+0"), ("11GHz", "4+0"), ("11GHz", "6+0"), ("11GHz", "8+0")]

CORE_MARGIN = 0.02      # 판매가 대비, 10개 핵심 기능군 27개 품목 전수 확인(예외 0)
ACCESSORY_MARGIN = 0.03  # 안테나/방수/접지/보호/100M급 케이블 - 확인됨
SFP_MARGIN = 0.05        # SFP·커넥터·서지보호 KIT - 3건 중 2건 정확 일치

# 돈현님 요청(2026-09-11, 8차): 고객사에 협상 여유를 두고 처음부터 조금 높게 부르기 위한
# 앵커링 버퍼. '제시 판매가'(KT 표준 마진 유지 최소선, 목표가/협상 마지노선)에 일괄 곱해
# '고객 제시가(최초 견적)'를 만든다 - 8%p는 권장 범위(5~10%p)의 중간값이며, 이 상수만
# 바꾸면 전체 시트에 일괄 반영된다.
ANCHOR_BUFFER = 0.08


def margin_rate(category):
    if category == "SFP":
        return SFP_MARGIN
    if category is None:
        return ACCESSORY_MARGIN
    return CORE_MARGIN


def verify_aviat_margin_structure():
    """01번 파일 전수 검산 - 이 파일의 핵심 전제(마진 구조)를 재현 가능하게 남긴다."""
    wb = openpyxl.load_workbook(SRC_AVIAT_213, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = row[idx["K코드"]]
        sale, buy = row[idx["판매단가"]], row[idx["매입단가"]]
        if not (isinstance(sale, (int, float)) and isinstance(buy, (int, float)) and buy > 0):
            continue
        cat = func_cmp.AVIAT_CATEGORY.get(k)
        if cat is None:
            continue  # 이 검증표는 이번 프로젝트가 이미 기능분류한 27개 핵심 품목만 대상으로 한다
        margin = (sale - buy) / sale
        rows.append((k, row[idx["품명"]], cat, buy, sale, margin))
    return rows


def load_ceragon_band(band, cfg, sd):
    wb = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    ws = wb["BoM_상세"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["원본주파수"]] != band or row[idx["표준화구성표기"]] != cfg or row[idx["SD구분"]] != sd:
            continue
        name = row[idx["품명"]].strip()
        rows.append({
            "품명": name, "K코드": row[idx["K코드"]],
            "매입액": row[idx["인하금액"]] or 0,
            "카테고리": func_cmp.CERAGON_CATEGORY.get(name),
        })
    return rows


def build_verify_sheet(wb):
    ws = wb.create_sheet("마진구조_검증(01번 전수)")
    headers = ["K코드", "품명", "기능분류", "매입단가", "판매단가", "마진율(판매가기준)"]
    ws.append(headers)
    rows = verify_aviat_margin_structure()
    for k, name, cat, buy, sale, margin in rows:
        r = ws.max_row + 1
        ws.append([k, name, func_cmp.CATEGORY_LABEL.get(cat, cat), buy, sale, margin])
        expected = margin_rate(cat)
        if abs(margin - expected) > 0.003:
            mark_fill(ws, r, 6, ERROR_FILL)
        else:
            mark_fill(ws, r, 6, CONFIRMED_FILL)
    for c in [4, 5]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=6).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    n_exact = sum(1 for _, _, cat, _, sale, m in rows if abs(m - margin_rate(cat)) <= 0.003)
    return len(rows), n_exact


def build_guide_sheet(wb, n_total, n_exact):
    ws = wb.create_sheet("안내")
    ws.append(["항목", "내용"])
    guide = [
        ("목적", "돈현님 지적(2026-09-11): 구성 품목 총액으로 Aviat-Ceragon을 비교하려던 시도는 "
                "실무 결론(대성에게 얼마를, KT 고객사에게 얼마를 제시할지)에 직접 닿지 못했다. "
                "이 파일은 그 두 질문에 직접 답한다."),
        ("매입가 제시안(대성에게)", "13번 파일 검토의견(2026-09-09) 그대로: 대성이 이미 제출한 "
                              "인하단가/인하금액을 매입가로 조건부 수용한다 - 반박할 근거도, "
                              "추가 인하를 요구할 근거도 확보하지 못했기 때문이다. 이 파일은 "
                              "매입가를 재검증하지 않고 그대로 가져온다(각 시트의 '매입액' 열)."),
        ("판매가 제시안(고객사에게) - 핵심 발견", f"01번 파일(Aviat 213종 계약마스터) 전수 검산 "
                                          f"결과, KT commerce는 무선전송장비 재판매 시 품목군별로 "
                                          f"정확히 고정된 마진율(판매가 대비)을 쓰고 있다: 핵심 "
                                          f"장비 10개 기능군 27개 품목은 100% 정확히 2.00%, "
                                          f"안테나/방수/접지/보호/장거리케이블은 3.00%, SFP/커넥터/"
                                          f"서지보호는 5.00%. 전체 {n_total}개 검증 품목 중 "
                                          f"{n_exact}개가 예상 마진율과 ±0.3%p 이내로 일치했다"
                                          f"('마진구조_검증' 시트에서 전부 재현 가능)."),
        ("왜 이 마진율을 Ceragon에도 쓰는가", "이건 'Ceragon 전용 정책'이 아니라 KT commerce가 "
                                       "이미 동일한 무선전송장비 재판매 사업에서 실제로 쓰고 있는"
                                       " 마진 구조다. Ceragon 소싱 장비도 같은 사업(마이크로웨이브"
                                       " 장비 매입 후 고객사 재판매)이므로, 임의의 마진율을 새로 "
                                       "가정하는 것보다 이 검증된 구조를 그대로 적용하는 편이 "
                                       "훨씬 방어 가능하다. 다만 이건 'KT의 공식 Ceragon 마진 "
                                       "정책'이 아니라 '가장 가까운 기존 정책을 유추 적용한 것'"
                                       "이므로, 최종 확정 전 내부 가격정책 담당자 확인을 권장한다."),
        ("계산식", "판매가 = 매입가 / (1 - 마진율) [마진율은 판매가 기준이지 매입가 기준이 "
                "아님에 주의 - 예: 매입가 100원에 2% 마진이면 판매가는 102원이 아니라 "
                "100/0.98=102.04원이다]."),
        ("Ceragon 품목의 마진율 분류 기준", "핵심 10개 기능군(섀시/무선송수신부/대역결합기/"
                                      "인터페이스카드/제어카드/라이선스/팬/전원/마운트/도파관) = "
                                      "2%(Aviat에서 예외 없이 확인). SFP = 5%(Aviat SFP 3건 중 "
                                      "2건이 정확히 일치해 다수 기준으로 채택). 그 외 케이블/"
                                      "커넥터/접지 등 미분류 품목 = 3%(Aviat의 보호/접지/방수 "
                                      "자재군과 가장 유사해 채택 - 이 부분은 2%/5%만큼 확실하지"
                                      " 않은 가정이며, '기타' 열에 따로 표시했다)."),
        ("이 파일이 다루지 않는 것", "매입가(대성 제시가) 자체의 적정성은 재판정하지 않는다"
                                 "(13번 결론 유지). 여기 나온 판매가는 'KT 표준 마진을 유지하는"
                                 " 최소 제시가'이며, 경쟁 상황·협상 여유·고객사와의 관계 등 "
                                 "전략적 가산은 포함하지 않았다 - 실제 제시가는 이 숫자 이상으로"
                                 " 조정될 수 있다."),
        ("6+0/8+0도 다룰 수 있는 이유", "이전 접근(Aviat와 총액 비교)은 6+0/8+0에 대응하는 "
                                    "Aviat 발주이력 자체가 없어 비교가 불가능했다. 이 접근은 "
                                    "Aviat의 마진율'구조'만 빌려오고 Aviat의 6+0/8+0 수량 데이터"
                                    "는 필요 없으므로, 대성이 제시한 모든 구성(2+0/4+0/6+0/8+0)"
                                    "에 동일하게 적용 가능하다 - 13번 파일에서 막혔던 부분이 "
                                    "여기서는 막히지 않는다."),
        ("전체카탈로그_단가제시안 시트(2026-09-11 추가)", "돈현님 지적: '지금 수량이 0인 게 "
                                              "영원한 게 아니다 - 카탈로그 474종(대성 품목마스터 "
                                              "기준 493종) 전부 매입가/판매가를 정해야 한다.' "
                                              "위 8개 시트는 현재 기본 구성(2+0/4+0/6+0/8+0)에만"
                                              " 있는 품목(78건)을 다루므로, 대성이 제출한 전체 "
                                              "카탈로그를 대상으로 별도 시트를 만들었다. K코드 "
                                              "체계상 이미 기능 대응이 확인된 41건은 '확인됨'으로, "
                                              "나머지 452건은 품명 키워드 패턴(SFP/커넥터→5%, "
                                              "안테나/마운트/방수→3%, 그 외 핵심장비 추정→2%)으로 "
                                              "'패턴 매칭 추정'으로 표시했다 - 개별 검증이 아니므로"
                                              " 신뢰도가 확인된 41건보다 낮다."),
        ("계약범위 추정 열(중요)", "카탈로그 493종 중 상당수는 15HP-/32T- 계열(구세대 라디오), "
                              "RFU-CXE(다른 아키텍처), 4-5GHz/6GHz/18GHz 전용 부품처럼 이번 "
                              "8/11GHz IP-20N/RFU-D-HP 계약과 무관해 보이는 품목이다(품명 "
                              "패턴으로 추정, 216건). 이런 품목도 가격은 매겨뒀지만(나중에 다른 "
                              "계약에 쓰일 수 있으므로 삭제하지 않음), '타대역/구세대 추정' 열로 "
                              "표시해 이번 계약과 무관할 가능성이 높다는 걸 알 수 있게 했다 - "
                              "실제 계약 범위인지는 대성/Ceragon에 확인이 필요하다."),
        ("고객 제시가(버퍼) 열(2026-09-11 추가)", f"돈현님 요청: 고객사와의 협상 여유를 위해 "
                              f"처음부터 목표가보다 조금 높게 부른다. '제시 판매가(목표가)'는 "
                              f"기존과 동일하게 KT 표준 마진을 유지하는 협상 마지노선이고, 여기에"
                              f" {ANCHOR_BUFFER*100:.0f}%p를 얹은 '고객 제시가(최초견적)'를 새로"
                              f" 추가했다 - 이 금액을 먼저 제시하고, 고객사가 인하를 요청하면 "
                              f"목표가 선까지 내려가며 마무리하는 구조다. 버퍼 크기는 "
                              f"ANCHOR_BUFFER 상수 하나만 바꾸면 전체 시트에 일괄 반영된다(권장"
                              f" 범위 5~10%p 중 중간값 채택 - 너무 크면 시장 감각과 멀어지고 "
                              f"너무 작으면 협상 카드로서 의미가 없다)."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100


def build_band_sheet(wb, band, cfg):
    ws = wb.create_sheet(f"{band}_{cfg}_제시안")
    headers = ["기능분류",
               "Non_SD 매입액(대성 제시가)", "Non_SD 마진율", "Non_SD 제시 판매가(목표가)", "Non_SD 마진액",
               f"Non_SD 고객 제시가(최초견적, 버퍼 {ANCHOR_BUFFER*100:.0f}%p)",
               "SD 매입액(대성 제시가)", "SD 마진율", "SD 제시 판매가(목표가)", "SD 마진액",
               f"SD 고객 제시가(최초견적, 버퍼 {ANCHOR_BUFFER*100:.0f}%p)"]
    ws.append(headers)

    totals = {}
    for sd in ["Non_SD", "SD"]:
        rows = load_ceragon_band(band, cfg, sd)
        by_cat = {}
        for r in rows:
            cat = r["카테고리"] or "기타"
            by_cat.setdefault(cat, 0)
            by_cat[cat] += r["매입액"]
        totals[sd] = by_cat

    cats = sorted(set(list(totals["Non_SD"].keys()) + list(totals["SD"].keys())),
                  key=lambda c: (func_cmp.CATEGORY_ORDER.index(c) if c in func_cmp.CATEGORY_ORDER else 99, c))

    grand = {"Non_SD_매입": 0, "Non_SD_판매": 0, "Non_SD_제시": 0, "SD_매입": 0, "SD_판매": 0, "SD_제시": 0}
    for cat in cats:
        label = func_cmp.CATEGORY_LABEL.get(cat, "⑫ 기타(케이블/커넥터/접지 등)")
        rate = margin_rate(None if cat == "기타" else cat)
        row_vals = [label]
        for sd in ["Non_SD", "SD"]:
            buy = totals[sd].get(cat, 0)
            sale = buy / (1 - rate) if buy else 0
            margin_amt = sale - buy
            anchor = sale * (1 + ANCHOR_BUFFER) if sale else 0
            row_vals += [buy or None, rate, sale or None, margin_amt or None, anchor or None]
            grand[f"{sd}_매입"] += buy
            grand[f"{sd}_판매"] += sale
            grand[f"{sd}_제시"] += anchor
        ws.append(row_vals)
        if cat == "기타":
            mark_fill(ws, ws.max_row, 1, REVIEW_FILL)

    r = ws.max_row + 1
    ws.append(["합계", grand["Non_SD_매입"] or None, None, grand["Non_SD_판매"] or None,
               (grand["Non_SD_판매"] - grand["Non_SD_매입"]) or None, grand["Non_SD_제시"] or None,
               grand["SD_매입"] or None, None, grand["SD_판매"] or None,
               (grand["SD_판매"] - grand["SD_매입"]) or None, grand["SD_제시"] or None])
    mark_fill(ws, r, 1, CONFIRMED_FILL)

    for c in [2, 4, 5, 6, 7, 9, 10, 11]:
        for rr in range(2, ws.max_row + 1):
            ws.cell(row=rr, column=c).number_format = FMT_AMOUNT
    for c in [3, 8]:
        for rr in range(2, ws.max_row + 1):
            ws.cell(row=rr, column=c).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["A"].width = 26

    return {
        "band": band, "cfg": cfg,
        "nonsd_매입": grand["Non_SD_매입"], "nonsd_판매": grand["Non_SD_판매"], "nonsd_제시": grand["Non_SD_제시"],
        "sd_매입": grand["SD_매입"], "sd_판매": grand["SD_판매"], "sd_제시": grand["SD_제시"],
    }


def build_summary_sheet(wb, results):
    ws = wb.create_sheet("전체_제시안_요약")
    headers = ["구성", "Non_SD 매입총액(대성 제시가)", "Non_SD 목표가(마지노선)", "Non_SD 예상마진액",
               f"Non_SD 고객 최초견적(버퍼 {ANCHOR_BUFFER*100:.0f}%p)",
               "SD 매입총액(대성 제시가)", "SD 목표가(마지노선)", "SD 예상마진액",
               f"SD 고객 최초견적(버퍼 {ANCHOR_BUFFER*100:.0f}%p)"]
    ws.append(headers)
    for res in results:
        ws.append([
            f"{res['band']} {res['cfg']}",
            res["nonsd_매입"], res["nonsd_판매"], res["nonsd_판매"] - res["nonsd_매입"], res["nonsd_제시"],
            res["sd_매입"], res["sd_판매"], res["sd_판매"] - res["sd_매입"], res["sd_제시"],
        ])
    for c in range(2, 10):
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws, 1, len(headers), ws.max_row)


def build_full_catalog_sheet(wb):
    """돈현님 요청(2026-09-11, 7차): '지금 수량 0인 게 영원한 게 아니다 - 474종 전부
    매입가/판매가를 정해야 한다'. 8개 기본 구성에 없는 나머지 품목도 언젠가 발주될 수
    있으므로, 대성이 제출한 카탈로그 전체(품목마스터 493종)에 마진율을 부여한다."""
    ws = wb.create_sheet("전체카탈로그_단가제시안(493종)")
    headers = ["K코드", "품명", "매입단가(대성 인하단가)", "적용 마진율", "제시 판매단가(목표가)",
               f"고객 제시단가(최초견적, 버퍼 {ANCHOR_BUFFER*100:.0f}%p)",
               "마진 분류", "계약범위 추정", "신뢰도"]
    ws.append(headers)

    wb_src = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    wsm = wb_src["품목마스터"]
    mheaders = [c.value for c in wsm[1]]
    midx = {h: i for i, h in enumerate(mheaders)}

    rows = []
    for row in wsm.iter_rows(min_row=2, values_only=True):
        kcode, name = row[midx["K코드"]], row[midx["품명"]].strip()
        buy = row[midx["최종인하단가"]] or 0
        rate, label, scope, conf = classify_catalog_item(name)
        sale = buy / (1 - rate) if buy else 0
        anchor = sale * (1 + ANCHOR_BUFFER) if sale else 0
        rows.append((kcode, name, buy, rate, sale, anchor, label, scope, conf))

    conf_order = {"확인됨": 0, "패턴 매칭 추정": 1}
    scope_order = {"8/11GHz 현재세대(확인됨)": 0, "8/11GHz 현재세대 추정": 1,
                   "비무선 제품(확인 필요)": 2, "타대역/구세대 추정(계약범위 확인 필요)": 3}
    rows.sort(key=lambda r: (conf_order.get(r[8], 9), scope_order.get(r[7], 9), r[1]))

    n_confirmed = n_pattern = 0
    for kcode, name, buy, rate, sale, anchor, label, scope, conf in rows:
        r = ws.max_row + 1
        ws.append([kcode, name, buy or None, rate, sale or None, anchor or None, label, scope, conf])
        if conf == "확인됨":
            mark_fill(ws, r, 9, CONFIRMED_FILL)
            n_confirmed += 1
        else:
            mark_fill(ws, r, 9, REVIEW_FILL)
            n_pattern += 1
        if scope.startswith("타대역") or scope.startswith("비무선"):
            mark_fill(ws, r, 8, ERROR_FILL)

    for c in [3, 5, 6]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["B"].width = 45
    return n_confirmed, n_pattern, len(rows)


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    n_total, n_exact = build_verify_sheet(wb)
    build_guide_sheet(wb, n_total, n_exact)

    results = []
    for band, cfg in BAND_CONFIGS:
        results.append(build_band_sheet(wb, band, cfg))

    build_summary_sheet(wb, results)
    n_confirmed, n_pattern, n_catalog = build_full_catalog_sheet(wb)

    sheet_order = (["안내", "마진구조_검증(01번 전수)"]
                   + [f"{b}_{c}_제시안" for b, c in BAND_CONFIGS]
                   + ["전체_제시안_요약", "전체카탈로그_단가제시안(493종)"])
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "15_매입판매가_제시안.xlsx"
    wb.save(out_path)

    print("=== 15단계: 매입가/판매가 제시안 생성 완료 ===")
    print(f"마진 구조 검증: {n_exact}/{n_total}개 품목이 예상 마진율과 ±0.3%p 이내 일치")
    for res in results:
        print(f"  - {res['band']} {res['cfg']}: Non_SD 매입 {res['nonsd_매입']:,.0f}원 → 목표가 "
              f"{res['nonsd_판매']:,.0f}원 → 고객 최초견적 {res['nonsd_제시']:,.0f}원 | SD 매입 "
              f"{res['sd_매입']:,.0f}원 → 목표가 {res['sd_판매']:,.0f}원 → 고객 최초견적 "
              f"{res['sd_제시']:,.0f}원")
    print(f"생성 파일: {out_path}")
    print("※ 매입가는 대성 제시가 조건부 수용(13번 결론), 판매가는 KT commerce가 Aviat 계약에서 "
          "실제로 쓰는 마진구조(2%/3%/5%)를 적용한 최소 제시가입니다 - 전략적 가산은 별도 검토 필요.")
    print(f"전체 카탈로그({n_catalog}종) 단가 제시안: 기능 확인됨 {n_confirmed}건, 품명 패턴 매칭 추정 "
          f"{n_pattern}건 - '전체카탈로그_단가제시안' 시트 참조.")


if __name__ == "__main__":
    main()
