"""
14단계: Aviat-Ceragon 기능 블록 단위 품목 대응/세트 비교 (돈현님 요청, 2026-09-08, 5차)
입력: output/04_Aviat_구성방식별_표준BoM_추정_v2.xlsx, output/06_Ceragon_구성방식별_BoM_정리.xlsx
출력: output/14_Aviat_Ceragon_기능별_비교.xlsx

돈현님 요청 원문: "품목별로 1대1 매칭을 하던구 아니면 대역별로 세트를 꾸려서 비교를 하던지 해주세요"
- 그동안은 K코드가 서로 다르다는 이유로 두 공급사 품목 간 비교를 보류해왔다. 이번에는 K코드
  문자열이 아니라 각 품목의 데이터시트 수준 기능 설명(06번 파일 '설명' 컬럼 등)에 근거해
  무선전송장비의 표준 기능 블록(섀시/무선송수신부/대역결합기/인터페이스카드/제어카드/
  라이선스/팬/전원/마운트/도파관/SFP) 단위로 Aviat 품목과 Ceragon 품목을 대응시킨다.

중요한 구분: 이 대응은 "K코드 A와 K코드 B가 같은 품목"이라는 뜻이 아니라 "서로 다른 두
공급사의 서로 다른 SKU가 링크 안에서 같은 역할을 한다"는 엔지니어링 판단이다. 04b의
확정/유력 후보/참고 후보 분류(발주이력 기반 수량 신뢰도)와는 별개의 축이며, 공급사 정식
데이터시트 대조 전까지는 이 대응관계 자체도 '판단(추정)' 단계로 취급한다.

수량 이상치 처리: Aviat 참고 후보 품목 중 '구성비 배수'(N당 배수)가 6을 넘는 품목은 허브/
대용량 집중국 자재가 발주이력 관측 창에 함께 섞였을 가능성이 커서(예: VR10 계열 품목이
VR4/11GHz 신호와 같은 발주 이벤트에 우연히 동시 존재), 카테고리 합계에서 제외하고 별도
표기한다 - 09번 검증 8번 항목("서로 다른 계열/주파수 오혼합 방지")과 같은 취지다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL,
                         FMT_AMOUNT, mark_fill)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_AVIAT_V2 = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정_v2.xlsx"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

CONFIGS = [("8GHz", "2+0"), ("8GHz", "4+0"), ("11GHz", "2+0"), ("11GHz", "4+0")]
OUTLIER_RATIO_THRESHOLD = 6

CATEGORY_ORDER = ["IDU_CHASSIS", "RADIO", "COUPLER", "IF_CARD", "CTRL_CARD",
                  "LICENSE", "FAN", "POWER", "MOUNT", "WAVEGUIDE", "SFP"]
CATEGORY_LABEL = {
    "IDU_CHASSIS": "① 본체/섀시(IDU·Chassis)",
    "RADIO": "② 무선송수신부(ODU/RFU)",
    "COUPLER": "③ 대역결합기(Coupler/Diplexer)",
    "IF_CARD": "④ 트래픽 인터페이스카드",
    "CTRL_CARD": "⑤ 제어카드(CPU/Control)",
    "LICENSE": "⑥ 라이선스/설정SW",
    "FAN": "⑦ 냉각팬",
    "POWER": "⑧ 전원공급장치",
    "MOUNT": "⑨ 마운팅/브래킷",
    "WAVEGUIDE": "⑩ 도파관(Waveguide)",
    "SFP": "⑪ SFP/광트랜시버",
}

# K코드 -> 기능분류. 04b의 '품목군' 태그는 일부 혼입(예: 케이블이 ODU/RFU로 분류됨)이 있어
# 참고만 하고, 여기서는 품명·역할을 직접 확인해 별도로 재분류했다.
AVIAT_CATEGORY = {
    "K9208698": "IDU_CHASSIS",   # VR4 1RU CHASSIS
    "K9208700": "IDU_CHASSIS",   # VR10 3RU CHASSIS
    "K9208714": "RADIO",         # 8GHz ODU IAP3 LOW
    "K9208715": "RADIO",         # 8GHz ODU IAP3 HIGH
    "K9208716": "RADIO",         # 11GHz ODU IAP3 LOW
    "K9208717": "RADIO",         # 11GHz ODU IAP3 HIGH
    "K9208726": "COUPLER",       # 7/8GHz FLAT HYBRID
    "K9208727": "COUPLER",       # 10/11GHz FLAT HYBRID
    "K9208706": "IF_CARD",       # GbE4f-AV
    "K9208707": "IF_CARD",       # GbE4e-AV
    "K9208709": "IF_CARD",       # 16E1-A
    "K9208710": "IF_CARD",       # STM1-A
    "K9208701": "CTRL_CARD",     # VR10 MC-MV
    "K9208711": "CTRL_CARD",     # MC-MV Redundancy
    "K9208703": "CTRL_CARD",     # TERM-MV
    "K9208733": "LICENSE",       # VR4 Node License
    "K9208734": "LICENSE",       # VR10 Node License
    "K9208699": "FAN",           # FAN-CV
    "K9208702": "FAN",           # FAN-MV
    "K9208704": "POWER",         # PS-MV
    "K9208728": "MOUNT",         # 7/8GHz IAP3 Mounting Bracket
    "K9208729": "MOUNT",         # 10/11GHz IAP3 Mounting Bracket
    "K9208724": "MOUNT",         # OBC2 Mounting Bracket Up to 4RF
    "K9208725": "MOUNT",         # OBC2 Mounting Bracket expand to 1RF
    "K9178534": "WAVEGUIDE",     # FLEXIBLE WAVEGUIDE 7.125-8.5GHz
    "K9188159": "WAVEGUIDE",     # FLEXIBLE WAVEGUIDE 10.7-11.7GHz
    "K9210279": "SFP",           # HAX_Gigabit Ethernet SFP
    "K9198328": "SFP",           # Gigabit Ethernet SFP_Industrial
    "K9210278": "SFP",           # SFP TSoP STM-1/OC3 over Gig-E Module
}

# Ceragon 품명(strip 처리) -> 기능분류
CERAGON_CATEGORY = {
    "CER_IP-20N-1RU/5-Slot-Base-IDU": "IDU_CHASSIS",
    "CER_IP-20N-2RU/10-Slot-Base-IDU": "IDU_CHASSIS",
    "CER_RFU-D-HP-08": "RADIO",
    "CER_RFU-D-HP-08-Non TX/2RX_SD": "RADIO",
    "CER_RFU-D-HP-11": "RADIO",
    "CER_RFU-D-HP-11-Non TX/2RX_SD": "RADIO",
    "CER_FXDH08-xxxY-BccN-H": "COUPLER",
    "CER_FXDH08-xxxY-BccN-H - SD": "COUPLER",
    "CER_FXDH08-xxxY-BccN-L": "COUPLER",
    "CER_FXDH08-xxxY-BccN-L - SD": "COUPLER",
    "CER_FXDH11-xxxY-BccN-H": "COUPLER",
    "CER_FXDH11-xxxY-BccN-H - SD": "COUPLER",
    "CER_FXDH11-xxxY-BccN-L": "COUPLER",
    "CER_FXDH11-xxxY-BccN-L - SD": "COUPLER",
    "CER_IP-20-RIC-D": "IF_CARD",
    "CER_IP-20 TCC-U": "CTRL_CARD",
    "CER_8GHz_11GHz_2+0_ Non SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_2+0_SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_4+0_Non SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_4+0_SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_6+0_Non SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_6+0_SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_8+0_Non SD_Configuration SW Package": "LICENSE",
    "CER_8GHz_11GHz_8+0_SD_Configuration SW Package": "LICENSE",
    "CER_SL-IDU-Radio-Port-Act": "LICENSE",
    "CER_IP-20-Fans-tray-1RU-IDU": "FAN",
    "CER_IP-20-Fans-tray-2RU-IDU": "FAN",
    "CER_IP-20-PDC-A48-dual_1RU": "POWER",
    "CER_IP-20-PDC-B48_2RU": "POWER",
    "CER_FXDH-RM-MOUNT-kit": "MOUNT",
    "CER_FXDH-RM-Term-7-8": "MOUNT",
    "CER_FXDH-RM-Term-11": "MOUNT",
    "CER_FXDH-RM-LU-Bend-7-8": "MOUNT",
    "CER_FXDH-RM-LU-Bend-11": "MOUNT",
    "CER_FXDH-RM-U-Bend-7-8": "MOUNT",
    "CER_FXDH-RM-U-Bend-11": "MOUNT",
    "CER_WG-Flex-1.2m-11GHz": "WAVEGUIDE",
    "CER_WG-Flex-1.2m-8GHz": "WAVEGUIDE",
    "CER_10G SFP+ SR 850nm 300m, MMF": "SFP",
    "CER_1G SFP 1310nm 10km, SMF": "SFP",
    "CER_2.5G SFP  1310nm 2km, SMF": "SFP",
}


def load_aviat_band(band, cfg):
    wb = openpyxl.load_workbook(SRC_AVIAT_V2, data_only=True)
    ws = wb[f"{band}_IAP3_{cfg}"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["K코드"]]:
            continue
        rows.append({
            "K코드": row[idx["K코드"]], "품명": row[idx["품명"]],
            "판정수준": row[idx["판정수준"]],
            "배수": row[idx["구성비 배수(N당 배수, 절대수량 아님)"]],
            "수량": row[idx["실제 수량(N+0 링크 전체, 양쪽 사이트 합산)"]],
            "판매단가": row[idx["판매단가"]] or 0,
        })
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
        rows.append({
            "품명": row[idx["품명"]].strip(), "수량": row[idx["1개링크_총수량"]],
            "금액": row[idx["인하금액"]] or 0,
        })
    return rows


class CategoryAgg:
    def __init__(self):
        self.strong = []      # 유력 후보 (품명, 수량, 금액)
        self.ref_ok = []      # 참고 후보, 배수 <= 임계값
        self.outlier = []     # 참고 후보, 배수 > 임계값(규모 이상치, 합계 제외)
        self.unstable = []    # 배수 자체가 '변동'(비수치)

    @property
    def strong_amt(self):
        return sum(a for _, q, a in self.strong)

    @property
    def ref_ok_amt(self):
        return sum(a for _, q, a in self.ref_ok)


def aggregate_aviat(rows):
    agg = {c: CategoryAgg() for c in CATEGORY_ORDER}
    others = []
    for r in rows:
        cat = AVIAT_CATEGORY.get(r["K코드"])
        if cat is None:
            others.append(r)
            continue
        a = agg[cat]
        qty, ratio, price = r["수량"], r["배수"], r["판매단가"]
        if not isinstance(ratio, (int, float)):
            a.unstable.append((r["품명"], qty, None))
            continue
        amt = qty * price if isinstance(qty, (int, float)) else 0
        if r["판정수준"] == "유력 후보":
            a.strong.append((r["품명"], qty, amt))
        elif ratio > OUTLIER_RATIO_THRESHOLD:
            a.outlier.append((r["품명"], qty, amt))
        else:
            a.ref_ok.append((r["품명"], qty, amt))
    return agg, others


def aggregate_ceragon(rows):
    agg = {c: {"items": [], "amt": 0} for c in CATEGORY_ORDER}
    others = {"items": [], "amt": 0}
    for r in rows:
        cat = CERAGON_CATEGORY.get(r["품명"])
        bucket = agg[cat] if cat else others
        bucket["items"].append((r["품명"], r["수량"], r["금액"]))
        bucket["amt"] += r["금액"] or 0
    return agg, others


def fmt_items(items, max_show=4):
    if not items:
        return ""
    parts = [f"{name}({qty})" for name, qty, _amt in items[:max_show]]
    if len(items) > max_show:
        parts.append(f"외 {len(items) - max_show}건")
    return "; ".join(parts)


def build_band_sheet(wb, band, cfg):
    aviat_rows = load_aviat_band(band, cfg)
    av_agg, av_others = aggregate_aviat(aviat_rows)
    cer_nonsd_rows = load_ceragon_band(band, cfg, "Non_SD")
    cer_sd_rows = load_ceragon_band(band, cfg, "SD")
    cer_nonsd_agg, cer_nonsd_others = aggregate_ceragon(cer_nonsd_rows)
    cer_sd_agg, cer_sd_others = aggregate_ceragon(cer_sd_rows)

    ws = wb.create_sheet(f"{band}_{cfg}_세트비교")
    headers = ["기능분류", "Aviat 유력후보 품목·수량", "Aviat 유력후보 금액",
               "Aviat 참고후보 품목·수량(안정)", "Aviat 참고후보 금액(안정)",
               "Aviat 제외 품목(규모이상치/변동, 공급사 확인 필요)",
               "Ceragon(Non_SD) 품목·수량", "Ceragon(Non_SD) 금액(인하가)",
               "Ceragon(SD) 품목·수량", "Ceragon(SD) 금액(인하가)",
               "비고(대응근거·차이설명)"]
    ws.append(headers)

    notes = {
        "IDU_CHASSIS": "둘 다 라디오/인터페이스/제어카드를 수납하는 본체 셀프. Ceragon 설명: "
                       "'1RU, 5+0까지 구성 가능한 IDU 셀프' - 슬롯 여유가 있어 2+0/4+0 모두 "
                       "섀시 1대(사이트당)로 커버되는 것으로 추정(공급사 확인 필요).",
        "RADIO": "사실로 확인되는 것: Ceragon 설명문에 MC-ABC(멀티캐리어 적응형 대역폭 제어)가 "
                 "명시돼 있고, RFU 총수량은 항상 N/2로 스케일한다(2+0→2, 4+0→4, 06번 원본 수량 "
                 "직접 확인). Aviat는 ODU가 LOW/HIGH 채널별 별도 유닛이라 채널 수만큼(N개) "
                 "필요하다는 것도 확인된 사실이다. 다만 '이 두 사실이 결합해 무선유닛 1대가 "
                 "반송파 2개를 처리하기 때문에 수량이 2배 차이난다'는 인과관계 자체는 "
                 "설명문·수량 패턴에서 나온 가장 유력한 추정이며, 공급사가 직접 확인해준 것은 "
                 "아니다 - 구조 차이로 추정, 기술 확인 필요.",
        "COUPLER": "사실로 확인되는 것: Ceragon 설명문 '채널필터 브랜칭'(Aviat FLAT HYBRID와 "
                   "같은 대역결합/분리 수동소자 역할)과, 이 부품이 2+0 구성 BoM에는 없고 4+0부터 "
                   "등장한다는 수량 패턴(06번 원본 확인)이다. Aviat는 채널별 별도 ODU 구조라 "
                   "2+0부터 이미 결합기가 필요하다. '두 벤더의 결합기 등장 시점이 다른 이유가 "
                   "RFU의 반송파 내부결합 때문'이라는 설명은 유력한 추정이나, 공급사 기술 확인 "
                   "전까지는 추정 단계로 취급한다.",
        "IF_CARD": "Ceragon 설명 'RFU-D, RFU-D-HP 연결용 라인카드'. Aviat는 트래픽 종류별"
                   "(GbE/E1/STM1)로 카드가 나뉘고, Ceragon은 RIC-D 한 종류에 트래픽 처리가 "
                   "통합돼 있어 카드 세분화 방식이 다르다 - 대수 비교보다 총 처리용량 기준 "
                   "비교가 더 적절할 수 있음(확인 필요).",
        "CTRL_CARD": "Ceragon 설명 'Main CPU 및 Monitor/Control Unit(TCC Card)'. Aviat "
                     "MC-MV/TERM-MV와 동일하게 장비 제어·관리 기능을 담당.",
        "LICENSE": "둘 다 장비 용량/기능 활성화 라이선스. Ceragon 설명에는 Capacity/ACM/"
                   "MC-ABC/XPIC/SD/Edge Node 등 세부 기능이 SW Package 1종에 통합돼 있고, "
                   "4+0부터는 무선포트 활성화 라이선스(SL-IDU-Radio-Port-Act)가 별도 SKU로 "
                   "추가된다. Aviat Node License 1종에 이런 세부 항목이 어떻게 포함되는지는 "
                   "확인 필요.",
        "FAN": "둘 다 장치 냉각용 팬 유닛. Ceragon 설명 '장치 냉각용 Fan Unit'.",
        "POWER": "둘 다 DC 전원공급 유닛. Ceragon 설명 '48VDC 듀얼 전원공급 Unit'.",
        "MOUNT": "둘 다 안테나/철탑에 무선유닛을 고정하는 기구물. Ceragon 설명 'ODU 철탑 "
                 "지지용 키트'.",
        "WAVEGUIDE": "명칭·설명·길이(1.2m)까지 근접한 동일 기능 부품(RF 신호 연결용 도파관).",
        "SFP": "광모듈 규격 자체가 업계 표준(SFP MSA)이라 실질적으로 동일 부품군 - 07번 파일 "
               "'직접대응_품목' 시트에서 이미 코드 매칭 없이 가격 비교를 완료한 항목.",
    }

    for cat in CATEGORY_ORDER:
        a = av_agg[cat]
        cn = cer_nonsd_agg[cat]
        cs = cer_sd_agg[cat]
        excl_parts = []
        if a.outlier:
            excl_parts.append("이상치: " + fmt_items(a.outlier))
        if a.unstable:
            excl_parts.append("변동: " + fmt_items(a.unstable))
        r = ws.max_row + 1
        ws.append([
            CATEGORY_LABEL[cat],
            fmt_items(a.strong), a.strong_amt if a.strong else None,
            fmt_items(a.ref_ok), a.ref_ok_amt if a.ref_ok else None,
            "; ".join(excl_parts),
            fmt_items(cn["items"]), cn["amt"] if cn["items"] else None,
            fmt_items(cs["items"]), cs["amt"] if cs["items"] else None,
            notes[cat],
        ])
        if a.strong:
            mark_fill(ws, r, 2, CONFIRMED_FILL)
            mark_fill(ws, r, 3, CONFIRMED_FILL)
        if a.ref_ok:
            mark_fill(ws, r, 4, REVIEW_FILL)
        if excl_parts:
            mark_fill(ws, r, 6, ERROR_FILL)

    # 핵심 11개 카테고리 소계(기타 제외) - 신뢰도별로 분리해서 더한다(섞지 않음)
    r = ws.max_row + 1
    total_strong = sum(av_agg[c].strong_amt for c in CATEGORY_ORDER)
    total_ref_ok = sum(av_agg[c].ref_ok_amt for c in CATEGORY_ORDER)
    total_cn = sum(cer_nonsd_agg[c]["amt"] for c in CATEGORY_ORDER)
    total_cs = sum(cer_sd_agg[c]["amt"] for c in CATEGORY_ORDER)
    ws.append(["소계(핵심 11개 카테고리, 기타/케이블 제외)", "", total_strong or None,
               "", total_ref_ok or None, "", "", total_cn or None, "", total_cs or None,
               "Aviat 유력후보와 참고후보는 신뢰도가 달라 더하지 않고 분리 표기했다. "
               "Ceragon 값은 같은 11개 카테고리만 뽑은 합계라 위 07/13번 파일의 "
               "'Ceragon 전체 총액'과는 다르다(케이블·필터트레이·접지 등 제외)."])
    mark_fill(ws, r, 1, CONFIRMED_FILL)

    # 기타(미대응) 행 - 여기도 카테고리와 동일한 기준으로 규모 이상치를 제외한다(일관성 유지).
    # 제외 없이 그대로 더치면 MODEM-AV(배수30)·TERMINAL UNIT(배수12) 같은 허브급 혼입 품목이
    # 섞여 실제보다 훨씬 큰 금액으로 보인다(최초 실행에서 8,865만원까지 부풀려짐, 확인 후 수정).
    others_stable = [o for o in av_others if isinstance(o["배수"], (int, float))
                     and o["배수"] <= OUTLIER_RATIO_THRESHOLD]
    others_excluded = [o for o in av_others if not isinstance(o["배수"], (int, float))
                       or o["배수"] > OUTLIER_RATIO_THRESHOLD]
    others_amt = sum((o['판매단가'] or 0) * o['수량'] for o in others_stable
                      if isinstance(o['수량'], (int, float)))
    r = ws.max_row + 1
    excl_note = f" (제외 {len(others_excluded)}건: " + ", ".join(
        f"{o['품명']}({o['수량']})" for o in others_excluded) + ")" if others_excluded else ""
    ws.append(["⑫ 기타(설치자재, 미대응)", "", "",
               f"{len(others_stable)}건{excl_note}", others_amt or None,
               "", f"{len(cer_nonsd_others['items'])}건", cer_nonsd_others["amt"] or None,
               f"{len(cer_sd_others['items'])}건", cer_sd_others["amt"] or None,
               "케이블·커넥터·안테나·접지·방수자재 등 - 양쪽 다 존재하나 개별 품목 간 1:1 "
               "기능대응은 이번 표에서 판단하지 않음(세부사양 확인 필요). 배수 6 초과 품목은 "
               "다른 카테고리와 같은 기준으로 제외해 앞 열에 별도 표기했다."])
    mark_fill(ws, r, 1, REVIEW_FILL)
    if others_excluded:
        mark_fill(ws, r, 4, ERROR_FILL)

    for c in [3, 5, 8, 10]:
        for rr in range(2, ws.max_row + 1):
            ws.cell(row=rr, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["K"].width = 55

    strong_categories = [c for c in CATEGORY_ORDER if av_agg[c].strong]
    return {
        "band": band, "cfg": cfg,
        "aviat_strong": total_strong, "aviat_ref_ok": total_ref_ok,
        "ceragon_nonsd": total_cn, "ceragon_sd": total_cs,
        "ceragon_nonsd_full": sum(x["amt"] for x in cer_nonsd_agg.values()) + cer_nonsd_others["amt"],
        "ceragon_sd_full": sum(x["amt"] for x in cer_sd_agg.values()) + cer_sd_others["amt"],
        "strong_category_count": len(strong_categories),
        "strong_categories": strong_categories,
    }


def build_guide_sheet(wb):
    ws = wb.create_sheet("안내")
    ws.append(["항목", "내용"])
    guide = [
        ("목적", "돈현님 요청(2026-09-08, 5차): '품목별로 1대1 매칭을 하던지 대역별로 세트를 "
                "꾸려서 비교를 하던지 해주세요'를 실제로 수행한다. 그동안은 K코드가 서로 달라 "
                "비교를 보류해왔으나, 이번에는 K코드 문자열이 아니라 각 품목의 데이터시트 수준 "
                "기능 설명(06번 파일 '설명' 컬럼)에 근거해 무선전송장비 표준 기능 블록 11종 "
                "단위로 Aviat·Ceragon 품목을 대응시키고(품목대응표 시트), 대역/구성별로 세트를 "
                "꾸려 수량·금액을 나란히 비교했다(각 세트비교 시트)."),
        ("중요한 원칙", "이 대응은 'K코드 A와 K코드 B가 같은 품목'이라는 뜻이 아니다 - 서로 "
                     "다른 두 회사의 서로 다른 SKU가 링크 안에서 같은 역할(무선 송수신, 채널 "
                     "결합, 전원공급 등)을 한다는 엔지니어링 판단이다. 04b의 확정/유력 후보/"
                     "참고 후보 분류(발주이력 기반 '수량' 신뢰도)와는 별개의 축이며, 이 "
                     "기능대응 판단 자체도 공급사 정식 데이터시트 대조 전까지는 '판단(추정)' "
                     "단계로 취급한다 - '확정'이 아니다."),
        ("핵심 구조적 발견 (1) 무선송수신부 - 확인된 사실 vs 추정 구분", "확인된 사실: Ceragon "
                                        "RFU-D-HP 라이선스 설명문에 MC-ABC(멀티캐리어 적응형 "
                                        "대역폭 제어)가 명시돼 있고, RFU 총수량은 항상 N/2로 "
                                        "스케일한다(2+0→2, 4+0→4, 6+0→6, 8+0→8, 06번 파일 원본 "
                                        "수량으로 직접 확인). Aviat는 ODU가 LOW/HIGH 채널별로 "
                                        "분리된 별도 유닛이라 채널 수만큼(N개) 필요한 것도 "
                                        "확인된 사실이다(11GHz 2+0에서 ODU=4 vs RFU=2). 다만 "
                                        "'RFU 1대가 반송파 2개를 내부 처리하기 때문에 이 2배 "
                                        "차이가 생긴다'는 인과 설명 자체는 이 두 사실로부터 "
                                        "도출한 가장 유력한 추정이며, 공급사가 명시적으로 확인해"
                                        "준 내용은 아니다 - 구조 차이로 추정, 기술 확인 필요."),
        ("핵심 구조적 발견 (2) 대역결합기 등장 시점 - 확인된 사실 vs 추정 구분", "확인된 사실: "
                                             "Ceragon 외부 결합기(FXDH 브랜칭 필터)는 2+0 구성 "
                                             "BoM에는 아예 없고, RFU 2대가 쓰이는 4+0부터 등장한다"
                                             "(06번 원본 수량으로 직접 확인). Aviat FLAT HYBRID는 "
                                             "채널별 ODU 구조상 2+0부터 이미 필요하다는 것도 "
                                             "확인된 사실이다. '두 벤더의 결합기 등장 시점이 다른"
                                             " 이유가 RFU의 반송파 내부결합 때문'이라는 설명은 "
                                             "유력한 추정이나, 공급사 기술 확인 전까지는 추정 "
                                             "단계로 취급한다."),
        ("수량 환산 근거(왜 관측수량을 그대로 링크 전체 수량으로 보는가)", "Aviat 발주이력에는 "
                                             "Ceragon처럼 'A국소/B국소' 구분이 없어, 관측된 원 "
                                             "수량을 사이트당으로 볼지 링크 전체로 볼지가 그 "
                                             "자체로는 불확실하다. 이 표는 Ceragon의 명시적 "
                                             "관례(1개링크_A국소수량 + B국소수량 = 총수량, 즉 두 "
                                             "끝 사이트 합산이 링크 총량)를 기준으로 삼아 Aviat "
                                             "원 관측수량을 '링크 전체 수량'으로 그대로 두었다 "
                                             "(즉 배수×N=원 관측수량이 되도록 정의). 근거는 추정이"
                                             " 아니라 교차검증이다 - 11GHz 2+0에서 Aviat가 실제로"
                                             " 주문한 원 수량(Chassis/License/Fan/Mount 4개 "
                                             "품목 모두 2)이 Ceragon이 공식 제출한 링크 총수량"
                                             "(같은 4개 품목 모두 2)과 정확히 일치한다 - 만약 "
                                             "Aviat의 '2'가 사이트당 수량이었다면 링크 전체는 4가"
                                             " 되어야 하므로 Ceragon과 어긋났을 것이다. 다만 이 "
                                             "관례를 Aviat가 명시적으로 확인해준 것은 아니므로 "
                                             "공급사 확인 항목으로 남겨둔다(07번 파일 "
                                             "공급사확인사항 참조)."),
        ("핵심 구조적 발견 (3) 섀시 용량", "Ceragon IP-20N-1RU/5-Slot IDU는 슬롯 여유가 있어 "
                                    "2+0/4+0 모두 사이트당 섀시 1대로 커버된다(수량 고정 "
                                    "확인). Aviat VR4 1RU CHASSIS는 (참고 후보 수준이라 "
                                    "확정은 아니지만) 8GHz 4+0 관측에서 사이트당 2대로 늘어나는"
                                    " 패턴이 나타나 - 다만 11GHz 4+0 관측에서는 사이트당 1대로 "
                                    "나타나 두 참고 후보 관측이 서로 불일치한다. 이 항목은 "
                                    "패턴 확정이 아니라 공급사 확인이 필요한 항목으로 남긴다."),
        ("수량 이상치 제외 기준", "Aviat 참고 후보 품목 중 '구성비 배수'(N당 배수)가 6을 넘는 "
                              "품목은 세트비교 카테고리 합계에서 제외했다 - 이런 품목은 대부분 "
                              "VR10 계열(대용량 집중국/허브 장비)이 VR4/해당 대역 신호와 같은 "
                              "발주 이벤트 창에 우연히 함께 있었던 것으로 보이며(예: 11GHz 2+0"
                              " 관측 창에 VR10 Node License 14개·FAN-MV 28개 등이 섞여 있음), "
                              "1개 링크 단위로 억지로 환산하면 오히려 왜곡된 비교가 된다. "
                              "제외된 품목은 각 세트비교 시트의 '제외 품목' 열에 그대로 남겨 "
                              "투명하게 표시했다(임의로 숨기지 않음)."),
        ("제외한 품목군", "케이블·커넥터·안테나·접지·방수자재 등은 이번 기능대응표에서 다루지 "
                       "않았다(각 세트비교 시트 맨 아래 '기타' 행에 건수·금액만 합산). 이런 "
                       "품목은 현장 시공 조건(케이블 길이, 철탑 형태 등)에 따라 달라져 코드나 "
                       "설명만으로 1:1 대응을 주장하기 어렵기 때문이다 - 강제로 매칭하면 이 "
                       "프로젝트가 처음부터 지켜온 '문자열 유사도만으로 동일 품목 단정 금지' "
                       "원칙에 어긋난다."),
        ("가장 근거가 뚜렷한 세트 - 그러나 아직 직접 비교는 불가", "11GHz 2+0만 유력 후보"
                                "(독립 2개 이벤트로 반복 확인) 품목이 있다 - 본체/무선송수신부/"
                                "팬/라이선스/마운트 5개 카테고리, 6개 품목, 실제 수량 기준 참고"
                                " 판매총액 19,088,472원. 그러나 이 5개 카테고리는 11개 핵심"
                                " 카테고리 중 일부일 뿐이다 - 대역결합기·인터페이스카드·제어"
                                "카드·전원·SFP는 11GHz 2+0에서도 참고 후보(단일 관측)로만 "
                                "존재하거나 아예 근거가 없다. 즉 Aviat 유력후보 19,088,472원을"
                                " Ceragon 핵심11종 합계(35~47백만원)와 곧바로 비교하면 서로 "
                                "다른 범위를 비교하는 것이다 - 나머지 3개 구성(8GHz 2+0/4+0, "
                                "11GHz 4+0)은 전 카테고리가 참고 후보뿐이라 신뢰도가 더 낮다."),
        ("읽는 법 및 현재 활용 범위", "각 '세트비교' 시트는 11개 기능분류를 행으로 하고, Aviat "
                 "유력후보/참고후보(신뢰도 분리)와 Ceragon Non_SD/SD 금액을 나란히 놓았다. 맨 "
                 "아래 '소계'는 케이블 등 기타를 제외한 핵심 11개 카테고리만의 합계로 비교 "
                 "'범위'를 맞추려는 시도이지만, Aviat 유력후보만으로는 이 11개 카테고리를 다 "
                 "채우지 못한다(위 항목 참조) - 정확히 말하면 '비교 틀은 마련됐으나 유력후보 "
                 "기준 직접 가격비교는 아직 불가'다. 참고후보까지 합쳐야 11개 카테고리가 채워"
                 "지지만, 참고후보는 단일 관측이라 신뢰도가 낮다. 지금 이 파일의 안전한 "
                 "활용 범위는 (1)품목 대응구조 설명 (2)대성 BoM 누락 품목 점검 (3)애니콤"
                 " 확인 질문 작성이며, 세트 금액을 가격 확정·인하 요구의 근거로 쓰는 것은 "
                 "아직 이르다 - 반드시 애니콤 정식 회신으로 유력후보 범위를 넓힌 뒤 다시 "
                 "봐야 한다."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100


def build_mapping_sheet(wb):
    ws = wb.create_sheet("품목대응표(1대1)")
    headers = ["기능분류", "대역/구성", "Aviat 대응 품목", "Ceragon 대응 품목", "판단근거(설명 인용 등)"]
    ws.append(headers)
    rows = [
        ("① IDU/Chassis", "공통", "VR4 1RU CHASSIS / VR10 3RU CHASSIS",
         "IP-20N-1RU/5-Slot-Base-IDU / IP-20N-2RU/10-Slot-Base-IDU",
         "라디오/인터페이스/제어카드를 수납하는 본체 셀프. Ceragon 설명: '1RU, 5+0까지 구성 "
         "가능한 IDU 셀프(4+0 MC-ABC 지원 백보드 포함)'."),
        ("② 무선송수신부", "8GHz", "8GHz ODU IAP3 LOW / 8GHz ODU IAP3 HIGH",
         "RFU-D-HP-08 (-Non TX/2RX_SD)",
         "Ceragon 설명 'High power ODU' - 안테나 직결 송수신부로 기능 동일. Ceragon 1대가 "
         "MC-ABC로 반송파 2개 처리, Aviat는 반송파(채널)마다 개별 ODU 필요 → 수량 2배 차이"
         "(구조 차이, 오류 아님)."),
        ("② 무선송수신부", "11GHz", "11GHz ODU IAP3 LOW / 11GHz ODU IAP3 HIGH",
         "RFU-D-HP-11 (-Non TX/2RX_SD)", "위와 동일 근거(11GHz용)."),
        ("③ 대역결합기", "8GHz", "7/8GHz FLAT HYBRID 7.125, Direct, UMT,3.7dB",
         "FXDH08-xxxY-BccN-H/L (-SD)",
         "Ceragon 설명 '8GHz High/Low용 채널필터 브랜칭' - 서로 다른 채널을 하나의 안테나로 "
         "결합/분리하는 수동소자로 Aviat FLAT HYBRID와 동일 기능. 단, Ceragon은 2+0에서는 이 "
         "부품이 BoM에 없고(RFU 1대가 내부 결합) 4+0부터 등장 - Aviat는 2+0부터 이미 필요."),
        ("③ 대역결합기", "11GHz", "10/11GHz FLAT HYBRID 10.0, Direct, UMT,3.7dB",
         "FXDH11-xxxY-BccN-H/L (-SD)", "위와 동일 근거(11GHz용)."),
        ("④ 인터페이스카드", "공통", "GbE4e-AV / GbE4f-AV / 16E1-A / STM1-A",
         "IP-20-RIC-D",
         "Ceragon 설명 'RFU-D, RFU-D-HP 연결용 라인카드'. Aviat는 트래픽 종류별로 카드가 "
         "나뉘고 Ceragon은 RIC-D 한 종류로 통합 - 카드 세분화 방식 차이(대수 비교보다 총 "
         "처리용량 비교가 더 적절할 수 있음, 확인 필요)."),
        ("⑤ 제어카드", "공통", "MC-MV Redundancy / VR10 MC-MV / TERM-MV", "IP-20 TCC-U",
         "Ceragon 설명 'Main CPU 및 Monitor/Control Unit(TCC Card), 6x10Gb 포트 지원' - "
         "제어·관리 기능 동일."),
        ("⑥ 라이선스", "공통", "VR4 Node License / VR10 Node License",
         "Configuration SW Package (+ SL-IDU-Radio-Port-Act, 4+0부터 추가)",
         "둘 다 용량/기능 활성화 라이선스. Ceragon 설명에 Capacity/ACM/MC-ABC/XPIC/SD/"
         "Edge Node 등이 SW Package 1종에 통합. Ceragon은 4+0부터 무선포트 활성화 라이선스가"
         " 별도 SKU로 추가되는데 Aviat Node License 1종에 이게 포함되는지는 확인 필요."),
        ("⑦ 냉각팬", "공통", "FAN-CV / FAN-MV", "IP-20-Fans-tray-1RU/2RU-IDU",
         "Ceragon 설명 '장치 냉각용 Fan Unit' - 동일 기능."),
        ("⑧ 전원공급장치", "공통", "PS-MV", "IP-20-PDC-A48-dual_1RU / IP-20-PDC-B48_2RU",
         "Ceragon 설명 '48VDC 듀얼 전원공급 Unit(PDC Card)' - 동일 기능."),
        ("⑨ 마운팅/브래킷", "공통", "IAP3 Mounting Bracket / OBC2 Mounting Bracket",
         "FXDH-RM-MOUNT-kit / FXDH-RM-Term / FXDH-RM-U-Bend / FXDH-RM-LU-Bend",
         "Ceragon 설명 'ODU 철탑 지지용 키트' 등 - 안테나/철탑에 무선유닛을 고정하는 기구물로"
         " 동일 기능."),
        ("⑩ 도파관", "8GHz", "FLEXIBLE WAVEGUIDE, 7.125-8.500GHz, 1200MM",
         "WG-Flex-1.2m-8GHz", "명칭·설명·길이(1.2m)까지 근접한 동일 기능 부품."),
        ("⑩ 도파관", "11GHz", "FLEXIBLE WAVEGUIDE, 10.7-11.7GHz, 1200MM",
         "WG-Flex-1.2m-11GHz", "위와 동일."),
        ("⑪ SFP/광트랜시버", "공통", "HAX_Gigabit Ethernet SFP 등",
         "1G/2.5G/10G SFP 시리즈",
         "광모듈 규격이 업계 표준(SFP MSA)이라 실질적으로 동일 부품군 - 07번 파일 "
         "'직접대응_품목' 시트에서 이미 코드 매칭 없이 가격 비교 완료."),
    ]
    for row in rows:
        ws.append(list(row))
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 45
    ws.column_dimensions["E"].width = 70


def build_summary_sheet(wb, results):
    ws = wb.create_sheet("세트비교_요약")
    headers = ["구성", "Aviat 유력후보 카테고리 수(11개 중)", "Aviat 유력후보 합계(해당 카테고리만)",
               "Aviat 참고후보 합계(안정, 핵심11종)",
               "Ceragon(Non_SD) 핵심11종 합계", "Ceragon(SD) 핵심11종 합계",
               "Ceragon(Non_SD) 전체 총액(참고)", "Ceragon(SD) 전체 총액(참고)",
               "직접 가격비교 가능 여부", "비고"]
    ws.append(headers)
    for res in results:
        r = ws.max_row + 1
        n_strong = res["strong_category_count"]
        can_compare = "불가(카테고리 범위 다름)"
        if res["band"] == "11GHz" and res["cfg"] == "2+0":
            note = (f"유력후보는 11개 카테고리 중 {n_strong}개({', '.join(CATEGORY_LABEL[c] for c in res['strong_categories'])})"
                     f"만 채운다 - 대역결합기·인터페이스카드·제어카드·전원·SFP는 유력후보가 "
                     f"없어 Ceragon 핵심11종 합계와 범위가 다르다. 따라서 19,088,472원을 "
                     f"35~47백만원과 직접 비교하면 서로 다른 범위를 비교하는 것이다 - '비교 "
                     f"틀은 마련됐으나 유력후보만으로는 아직 직접 가격비교 불가'가 정확한 "
                     f"표현이다. 참고후보까지 합치면 11개 카테고리가 채워지지만 참고후보는 "
                     f"단일 관측이라 신뢰도가 낮다.")
        else:
            note = "전 카테고리가 참고 후보(단일 관측)뿐이거나 근거가 없다 - 세트 합계도 참고용일 뿐 협상 근거로 쓰지 않는다."
        ws.append([f"{res['band']} {res['cfg']}", f"{n_strong}/11", res["aviat_strong"] or None,
                   res["aviat_ref_ok"] or None, res["ceragon_nonsd"], res["ceragon_sd"],
                   res["ceragon_nonsd_full"], res["ceragon_sd_full"], can_compare, note])
        if res["aviat_strong"]:
            mark_fill(ws, r, 3, CONFIRMED_FILL)
        mark_fill(ws, r, 9, ERROR_FILL)
    for c in [3, 4, 5, 6, 7, 8]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["J"].width = 70


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_guide_sheet(wb)
    build_mapping_sheet(wb)

    results = []
    for band, cfg in CONFIGS:
        results.append(build_band_sheet(wb, band, cfg))

    build_summary_sheet(wb, results)

    sheet_order = ["안내", "품목대응표(1대1)"] + [f"{b}_{c}_세트비교" for b, c in CONFIGS] + ["세트비교_요약"]
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "14_Aviat_Ceragon_기능별_비교.xlsx"
    wb.save(out_path)

    print("=== 14단계: Aviat-Ceragon 기능 블록 단위 비교 완료 ===")
    for res in results:
        print(f"  - {res['band']} {res['cfg']}: Aviat 유력후보 {res['aviat_strong']:,}원 / "
              f"참고후보(안정) {res['aviat_ref_ok']:,}원 | Ceragon 핵심11종 "
              f"Non_SD {res['ceragon_nonsd']:,.0f}원 / SD {res['ceragon_sd']:,.0f}원")
    print(f"생성 파일: {out_path}")
    print("※ 기능 대응은 K코드 동일성 주장이 아니라 데이터시트 설명 근거 엔지니어링 판단이며, "
          "'안내' 시트의 구조적 차이 설명과 함께 읽어야 한다.")


if __name__ == "__main__":
    main()
