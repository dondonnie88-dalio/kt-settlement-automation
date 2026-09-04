"""
5단계 v2: Aviat 8GHz/11GHz IAP3 계열 구성방식별 표준 BoM 재구성 (돈현님 피드백 반영, 2026-09-04)

기존 04_Aviat_구성방식별_표준BoM_추정.xlsx는 세부규격에 '2+0/8+0' 또는 WBX '4/8 채널'이
문자 그대로 적힌 품목만 앵커로 사용해, 발주이력에 훨씬 자주 등장하는 VR4/VR10 + IAP3 ODU
계열(체결식 표기가 없는 계열)이 2+0/4+0 후보에서 사실상 빠져 있었다. 이 스크립트는
03/02/01 산출물을 원자료 수준에서 다시 조인하여 아래 원칙으로 재작성한다.

1. '확정 후보'를 두 종류로 분리한다.
   - 구성표기 확인 품목: 품명/규격에 2+0,4+0,8+0,4 CHANNEL,8 CHANNEL이 직접 기재된 품목
     (품목이 그 구성에 '속한다'는 사실만 확정, 수량까지 확정된 것은 아님)
   - 표준 BoM 확정 품목: 1개 링크 기준 수량까지 공식 자료(계약서 텍스트)로 확인된 품목
2. WBX/WTM 'SD OPTION' 품목은 기본품목과 분리하여 Diversity_옵션 시트에 별도 정리한다.
3. VR4(K9208698)/VR10(K9208700)/MODEM-AV(K9208705)/8·11GHz IAP3 ODU LOW·HIGH
   (K9208714~717)/FLAT HYBRID(K9208726,727)/IAP3 Mounting Bracket(K9208728,729)/
   Node License(K9208733,734) 등의 동시출현·수량비를 발주이력 원자료에서 재분석한다.
4. 8GHz IAP3 2+0 / 8GHz IAP3 4+0 / 11GHz IAP3 2+0 / 11GHz IAP3 4+0 후보를 각각 독립적으로
   작성한다. 판별 신호: 같은 이벤트(또는 인접 발주 묶음) 내 ODU LOW수량=HIGH수량(=N)을
   동반 VR4/VR10 Chassis 수로 나눈 값이 2 또는 4에 근접하는 경우.
5. 그룹 단위는 동일 주문번호, 동일 주문일자+주문자+배송지(2차 이벤트), 그리고 같은
   주문자+배송지 조합에서 1~7일/8~30일 이내 후속 발주까지 포함해 인접 이벤트를 함께 본다.
6. 대량 케이블/SFP/Connector 등 공용자재는 핵심 세트에서 분리해 핵심 필수/설치 필수/
   현장 선택/공용 또는 예비품 4범주로 나눈다.
7. 관측 비율의 (최대-최소)/최소가 30% 이상이면 대표 수량을 매기지 않고
   '변동, 공급사 확인 필요'로 표시한다.
8. 모든 결론에 근거 주문번호·이벤트ID·관측수량·관측횟수·판정수준을 남긴다.

출력: output/04_Aviat_구성방식별_표준BoM_추정_v2.xlsx
"""
import re
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path
from statistics import median

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (
    finalize_sheet, ORIGIN_FONT, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL,
    FMT_AMOUNT, FMT_QTY, mark_fill, mark_font,
)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_ORDERS = OUTPUT_DIR / "02_Aviat_발주이력_정제.xlsx"
SRC_ITEMS = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"

CORE_WATCHLIST = {
    "K9208698": "VR4 1RU CHASSIS", "K9208700": "VR10 3RU CHASSIS",
    "K9208705": "MODEM-AV", "K9208714": "8GHz ODU IAP3 LOW",
    "K9208715": "8GHz ODU IAP3 HIGH", "K9208716": "11GHz ODU IAP3 LOW",
    "K9208717": "11GHz ODU IAP3 HIGH", "K9208726": "7/8GHz FLAT HYBRID",
    "K9208727": "10/11GHz FLAT HYBRID", "K9208728": "7/8GHz IAP3 Mounting Bracket",
    "K9208729": "10/11GHz IAP3 Mounting Bracket", "K9208733": "VR4 Node License",
    "K9208734": "VR10 Node License",
}
CHASSIS_FAMILIES = {"VR4", "VR10"}
BAND_TARGETS = [("8GHz", 2), ("8GHz", 4), ("11GHz", 2), ("11GHz", 4)]
VARIATION_THRESHOLD = 0.30
# IAP3(VR4/VR10 chassis) 구성 표에는 이 계열이 아닌 품목을 넣지 않는다. 장비계열이 명시적으로
# 다른 값(CTR8312/8540/8740, WTM4200/4500/4500XT 등)이면 주파수 표기가 없더라도 '공통성 품목'으로
# 보지 않고 제외한다(WTM4500 NODE LICENSE가 11GHz_IAP3_2+0에 혼입된 사례로 확인된 문제 수정).
# WBX/WTM은 extract_family_key()가 별도 계열 키를 붙이지 않으므로 품목군으로 함께 걸러낸다.
INCOMPATIBLE_FAMILIES = {"CTR8312", "CTR8540", "CTR8740", "WTM4200", "WTM4500", "WTM4500XT",
                          "INUe", "ODU300", "ODU600"}
INCOMPATIBLE_GROUPS = {"WBX", "WTM"}


CATEGORY_MAP = {
    "IDU/Chassis": "핵심 필수품목", "ODU/RFU": "핵심 필수품목", "Modem": "핵심 필수품목",
    "License": "핵심 필수품목", "WBX": "핵심 필수품목", "WTM": "핵심 필수품목",
    "OBC2": "핵심 필수품목", "Duplexer": "핵심 필수품목", "CPU/Control": "핵심 필수품목",
    "Fan": "핵심 필수품목", "Interface Card": "핵심 필수품목",
    "Mount/Bracket": "설치 필수품목", "Coupler/Hybrid": "설치 필수품목",
    "Waveguide": "설치 필수품목", "PDP/Power": "설치 필수품목",
    "Antenna": "현장 선택품목", "Surge/Grounding": "현장 선택품목",
    "Cable": "공용 또는 예비품", "Connector": "공용 또는 예비품",
    "SFP/Optical Module": "공용 또는 예비품", "Rack": "공용 또는 예비품",
    "PC/운영장비": "공용 또는 예비품", "NMS/EMS": "공용 또는 예비품",
    "Tool": "공용 또는 예비품", "기타": "공용 또는 예비품",
}


def gap_bucket(days):
    if days == 0:
        return "당일(동일이벤트)"
    if days <= 7:
        return "1~7일 후속"
    if days <= 30:
        return "8~30일 후속"
    if days <= 90:
        return "31~90일"
    return "91일 이상"


def load_item_master():
    wb = openpyxl.load_workbook(SRC_ITEMS, data_only=True)
    ws = wb["품목정규화"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    master = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        kcode = row[idx["K코드"]]
        if not kcode:
            continue
        master[kcode] = {
            "품명": row[idx["품명"]], "품목군": row[idx["품목군"]], "주파수": row[idx["주파수"]] or "",
            "장비계열": row[idx["장비계열"]] or "", "LowHigh_Band": row[idx["LowHigh_Band"]] or "",
            "구성표기": row[idx["구성표기(2+0등)"]] or "", "채널수": row[idx["채널수"]] or "",
            "SD표기": row[idx["SD_Diversity표기"]] or "",
        }
    ws0 = wb["원본정리"]
    headers0 = [c.value for c in ws0[1]]
    idx0 = {h: i for i, h in enumerate(headers0)}
    for row in ws0.iter_rows(min_row=2, values_only=True):
        kcode = row[idx0["K코드"]]
        if kcode in master:
            master[kcode]["판매단가"] = row[idx0["판매단가"]] or 0
            master[kcode]["매입단가"] = row[idx0["매입단가"]] or 0
    return master


def load_orders():
    wb = openpyxl.load_workbook(SRC_ORDERS, data_only=True)
    ws = wb["정제데이터"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    orders = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["주문번호"]] is None:
            continue
        orders.append({
            "주문번호": row[idx["주문번호"]], "주문일자": row[idx["주문일자"]],
            "주문자": row[idx["주문자"]], "부서": row[idx["부서"]], "배송지": row[idx["배송지"]],
            "K코드": row[idx["K코드(보정반영)"]], "품명": row[idx["품명"]],
            "수량": row[idx["수량"]] or 0,
        })
    return orders


def build_events(orders):
    """03단계와 동일한 키(동일 주문일자+주문자+부서+배송지)로 2차 이벤트를 재구성한다.
    03과 달리 이벤트별 실제 주문번호 목록까지 보존해 근거 인용에 사용한다."""
    groups = defaultdict(list)
    for o in orders:
        key = (o["주문일자"], o["주문자"], o["부서"], o["배송지"])
        groups[key].append(o)
    events = {}
    for i, (key, lines) in enumerate(sorted(groups.items(), key=lambda x: (x[0][0] or date.min)), start=1):
        eid = f"EVT-{i:03d}"
        events[eid] = {
            "이벤트ID": eid, "주문일자": key[0], "주문자": key[1], "부서": key[2], "배송지": key[3],
            "주문번호목록": sorted(set(o["주문번호"] for o in lines)), "라인": lines,
        }
    return events


def build_chains(events):
    """동일 주문자+배송지 조합의 이벤트를 날짜순으로 묶어 인접 이벤트 간 발주간격을 계산."""
    groups = defaultdict(list)
    for ev in events.values():
        groups[(ev["주문자"], ev["배송지"])].append(ev)
    chains = []
    for key, evs in groups.items():
        evs_sorted = sorted(evs, key=lambda e: e["주문일자"] or date.min)
        links = []
        for a, b in zip(evs_sorted, evs_sorted[1:]):
            gap = (b["주문일자"] - a["주문일자"]).days if (a["주문일자"] and b["주문일자"]) else None
            links.append({"from": a["이벤트ID"], "to": b["이벤트ID"], "gap_days": gap,
                          "구간": gap_bucket(gap) if gap is not None else "불명"})
        chains.append({"주문자": key[0], "배송지": key[1], "이벤트순서": [e["이벤트ID"] for e in evs_sorted],
                        "연결": links})
    return chains


def find_band_signals(events, master, chains):
    """이벤트 단위 + 1~30일 인접 이벤트를 합산한 '확장윈도우' 단위로 8/11GHz IAP3 ODU
    LOW=HIGH 쌍수 및 동반 VR4/VR10 Chassis 수를 계산해 N(=Chassis당 채널수) 후보를 만든다."""
    windows = []

    def make_window(win_id, member_events, note):
        lines = []
        order_nos = []
        for eid in member_events:
            lines.extend(events[eid]["라인"])
            order_nos.extend(events[eid]["주문번호목록"])
        freqs_present = {master.get(ln["K코드"], {}).get("주파수", "") for ln in lines} - {""}
        multi_band = len([f for f in freqs_present if f in ("8GHz", "11GHz")]) > 1
        chassis_qty = sum(
            ln["수량"] for ln in lines
            if master.get(ln["K코드"], {}).get("품목군") == "IDU/Chassis"
            and master.get(ln["K코드"], {}).get("장비계열") in CHASSIS_FAMILIES
        )
        for freq in ("8GHz", "11GHz"):
            low = sum(ln["수량"] for ln in lines
                      if master.get(ln["K코드"], {}).get("주파수") == freq
                      and master.get(ln["K코드"], {}).get("품목군") == "ODU/RFU"
                      and master.get(ln["K코드"], {}).get("LowHigh_Band") == "Low Band")
            high = sum(ln["수량"] for ln in lines
                       if master.get(ln["K코드"], {}).get("주파수") == freq
                       and master.get(ln["K코드"], {}).get("품목군") == "ODU/RFU"
                       and master.get(ln["K코드"], {}).get("LowHigh_Band") == "High Band")
            if low == 0 and high == 0:
                continue
            symmetric = low == high
            n_raw = low if symmetric else max(low, high)
            # N(=N+0의 N) 판별은 'ODU LOW=HIGH 쌍수 자체'를 1차 신호로 쓴다. 시행착오 기록:
            # 처음에는 이 값을 동반 Chassis 수로 나눈 값을 N으로 썼으나, Chassis 수를 올바르게
            # 세도록 분류 버그(VR4/VR10 CHASSIS가 'Interface Card'로 오분류되던 문제)를 고친
            # 뒤 재계산하니 비율이 거의 전부 1(=Chassis 1대당 채널 1개)로 나와 2+0/4+0 근거가
            # 사실상 사라졌다. 반면 FLAT HYBRID/IAP3 Mounting Bracket 수량은 Chassis가 아니라
            # 'ODU쌍수'와 거의 1:1로 함께 움직여(예: ODU쌍4→Hybrid4), IAP3 결합 구조에서는
            # 'ODU쌍수 자체'가 N+0의 N에 더 가까운 신호로 판단된다. Chassis 수는 참고용 보조
            # 지표로만 별도 열에 남긴다(1개 Chassis가 몇 채널을 수용하는지는 확인 필요 사항).
            n_candidate = n_raw if n_raw in (2, 4, 6, 8) else None
            basis = f"ODU LOW=HIGH 쌍수={n_raw}(직접 채택)"
            if chassis_qty > 0:
                basis += f", 참고: 동반 Chassis={chassis_qty}대(ODU쌍/Chassis={n_raw/chassis_qty:.2f})"
            else:
                basis += ", 참고: 동반 Chassis 없음"
            if not symmetric:
                basis += f" [비대칭 LOW={low}/HIGH={high}, N은 max값 사용]"
            windows.append({
                "윈도우ID": win_id, "이벤트목록": list(member_events), "주문번호목록": sorted(set(order_nos)),
                "주파수": freq, "ODU_LOW": low, "ODU_HIGH": high, "대칭": symmetric,
                "Chassis수량": chassis_qty, "N후보": n_candidate, "다중대역": multi_band,
                "근거": basis, "비고": note, "lines": lines,
            })

    for eid, ev in events.items():
        make_window(eid, [eid], "단일 이벤트(동일일자+주문자+배송지)")

    for chain in chains:
        for link in chain["연결"]:
            if link["gap_days"] is not None and link["gap_days"] <= 30:
                combo_id = f"{link['from']}+{link['to']}"
                make_window(combo_id, [link["from"], link["to"]],
                            f"인접 발주 결합({link['구간']}, gap={link['gap_days']}일)")
    return windows


def independent_subset(obs):
    """'EVT-010'과 'EVT-010+EVT-011'처럼 한쪽이 다른 쪽의 이벤트를 포함하는 결합윈도우는
    서로 독립된 관측이 아니다(같은 발주를 두 번 세는 것과 같음). 이벤트 집합이 서로 겹치지
    않는 관측만 골라 '독립적으로 반복 확인된 횟수'를 계산한다. 더 작은(단일 이벤트) 윈도우를
    우선 선택해, 여러 이벤트를 인위적으로 합친 결합윈도우보다 원래 단일 이벤트 근거를 우선한다."""
    chosen = []
    used_events = set()
    for o in sorted(obs, key=lambda x: len(x["이벤트집합"])):
        if used_events & o["이벤트집합"]:
            continue
        chosen.append(o)
        used_events |= o["이벤트집합"]
    return chosen


def classify_confidence(obs):
    if not obs:
        return "참고", None, []
    indep = independent_subset(obs)
    ratios = [o["비율"] for o in indep]
    lo, hi = min(ratios), max(ratios)
    variable = (hi - lo) / lo >= VARIATION_THRESHOLD if lo > 0 else True
    rep = None if variable else round(median(ratios), 2)
    level = "유력" if (len(indep) >= 2 and not variable) else "참고"
    return level, rep, indep


def category_of(info):
    return CATEGORY_MAP.get(info.get("품목군"), "공용 또는 예비품")


def is_sd_option_name(name):
    return bool(re.search(r"SD\s*OPTION", name or "", re.IGNORECASE))


def main():
    master = load_item_master()
    orders = load_orders()
    events = build_events(orders)
    chains = build_chains(events)
    signals = find_band_signals(events, master, chains)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ================= 구성표기_확인품목 =================
    ws_named = wb.create_sheet("구성표기_확인품목")
    headers_named = ["K코드", "품명", "품목군", "주파수", "확인된 표기", "판정수준", "설명"]
    ws_named.append(headers_named)
    named_kcodes = set()
    for kcode, info in master.items():
        tag = None
        if info["구성표기"] in ("2+0", "4+0", "6+0", "8+0"):
            tag = info["구성표기"]
        elif info["채널수"] in ("4채널", "8채널") and info["품목군"] == "WBX":
            tag = "4+0" if info["채널수"] == "4채널" else "8+0"
        if tag:
            named_kcodes.add(kcode)
            ws_named.append([kcode, info["품명"], info["품목군"], info["주파수"] or "(미태깅)", tag,
                              "구성표기 확인(멤버십 확정, 수량은 별도 확인 필요)",
                              "품명/세부규격에 해당 구성 표기가 직접 존재함(계약서 텍스트 근거)"])
            mark_fill(ws_named, ws_named.max_row, 6, CONFIRMED_FILL)
    finalize_sheet(ws_named, 1, len(headers_named), ws_named.max_row)

    # ================= 표준 BoM 확정 품목 (수량까지 원문에 명시) =================
    # CTR8312=2+0, CTR8540=8+0은 '패킷형 IDU'로 계약서에 명시되어 있으나 1개 링크 수량
    # 자체가 원문에 없으므로(=발주이력으로만 추정 가능) '표준 BoM 확정 품목'은 현재
    # 해당 사항 없음(확인필요 시트에 사유 기록).

    # ================= WTM_WBX_기본구성 (SD OPTION 제외) =================
    ws_wbx = wb.create_sheet("WTM_WBX_기본구성")
    headers_wbx = ["K코드", "품명", "품목군", "주파수", "채널수/구성표기", "판정수준", "비고"]
    ws_wbx.append(headers_wbx)
    for kcode, info in master.items():
        if info["품목군"] not in ("WBX", "WTM"):
            continue
        if is_sd_option_name(info["품명"]):
            continue
        tag = info["채널수"] or info["구성표기"] or "-"
        level = "구성표기 확인" if (info["채널수"] in ("4채널", "8채널") or info["구성표기"]) else "참고"
        ws_wbx.append([kcode, info["품명"], info["품목군"], info["주파수"] or "(미태깅)", tag, level,
                        "SD OPTION 아닌 기본 품목"])
    finalize_sheet(ws_wbx, 1, len(headers_wbx), ws_wbx.max_row)

    # ================= Diversity_옵션 =================
    ws_sd = wb.create_sheet("Diversity_옵션")
    headers_sd = ["K코드", "품명", "품목군", "주파수", "구분", "설명"]
    ws_sd.append(headers_sd)
    for kcode, info in master.items():
        if is_sd_option_name(info["품명"]):
            ws_sd.append([kcode, info["품명"], info["품목군"], info["주파수"] or "(미태깅)",
                          "조합관계 확인 필요",
                          "기본 WBX/WTM 품목과 별개로 존재하는 SD(Space Diversity) 옵션. 기본 품목에 "
                          "'추가'되는 것인지 특정 부품을 '대체'하는 것인지 원본 자료로 확정되지 않음"
                          "(공급사 확인 필요). 기본 구성 완성세트 앵커로 합산하지 않음."])
            mark_fill(ws_sd, ws_sd.max_row, 5, REVIEW_FILL)
        elif info["SD표기"] == "Y":
            ws_sd.append([kcode, info["품명"], info["품목군"], info["주파수"] or "(미태깅)",
                          "SD 관련 표기(옵션 여부 불명)",
                          "품명/세부규격에 SD 관련 단어가 있으나 'OPTION'으로 명시되지는 않음(확인 필요)."])
    if ws_sd.max_row == 1:
        ws_sd.append(["SD 관련 품목 없음"])
    finalize_sheet(ws_sd, 1, len(headers_sd), ws_sd.max_row)

    # ================= 8GHz/11GHz IAP3 2+0 / 4+0 =================
    band_sheet_rows = {}
    for freq, target_n in BAND_TARGETS:
        sheet_name = f"{freq}_IAP3_{target_n}+0"
        ws = wb.create_sheet(sheet_name)
        headers = ["K코드", "품명", "품목군", "분류", "대표 수량(1개 링크=N ODU쌍 기준)", "판정수준",
                   "관측횟수(윈도우, 중복포함)", "독립관측횟수(중복제외, 판정근거)",
                   "관측치 목록(수량/기준)", "근거 이벤트ID", "근거 주문번호",
                   "판매단가", "매입단가", "비고"]
        ws.append(headers)

        # N(=target_n)은 'ODU LOW=HIGH 쌍수' 자체이며, 다른 모든 품목의 수량은 이 N을
        # 기준으로 나눠 '1개 링크(N+0) 기준 배수'를 관측한다. 동반 Chassis 수량 자체도
        # 참고용으로 같은 표에 한 행으로 함께 노출한다(비율 해석은 확인 필요 시트 참고).
        qualifying = [s for s in signals if s["주파수"] == freq and s["N후보"] == target_n]
        item_obs = defaultdict(list)
        item_windows = defaultdict(set)
        item_orders = defaultdict(set)
        for sig in qualifying:
            denom = target_n
            if denom <= 0:
                continue
            for ln in sig["lines"]:
                info = master.get(ln["K코드"], {})
                ln_freq = info.get("주파수", "")
                # VR4/VR10 Chassis, MODEM-AV, Node License 등 핵심 품목은 자체 주파수 표기가
                # 없어(품명에 대역이 없음) 엄격히 freq만 매칭하면 전부 빠진다(돈현님 피드백으로
                # 확인된 문제). 주파수 태그가 없는 품목은 '어느 대역과도 배타적이지 않음'으로
                # 보고 포함하되, 다른 대역으로 명시 태깅된 품목만 제외한다.
                if ln_freq not in (freq, "") or ln["수량"] <= 0:
                    continue
                # 주파수 태그가 없다고 해서 전부 '공통 품목'은 아니다 - CTR8312/8540/8740,
                # WTM4200/4500/4500XT, WBX처럼 VR4/VR10 IAP3 계열과 무관한 다른 장비군이
                # 같은 이벤트에 섞여 있으면(예: 결합윈도우가 CTR 전용 이벤트를 함께 묶은 경우)
                # 명시적으로 다른 계열/품목군인 품목은 제외한다.
                if info.get("장비계열") in INCOMPATIBLE_FAMILIES or info.get("품목군") in INCOMPATIBLE_GROUPS:
                    continue
                ratio = ln["수량"] / denom
                item_obs[ln["K코드"]].append({
                    "윈도우ID": sig["윈도우ID"], "비율": ratio, "수량": ln["수량"], "기준": denom,
                    "다중대역": sig["다중대역"], "주파수귀속": "명시태깅" if ln_freq == freq else "무태깅(공통추정)",
                    "이벤트집합": frozenset(sig["이벤트목록"]),
                })
                item_windows[ln["K코드"]].add(sig["윈도우ID"])
                # 이 K코드 관측치의 근거는 그 K코드 자신의 라인이 속한 주문번호로 한정한다.
                # sig["주문번호목록"]은 같은 이벤트(동일일자+주문자+배송지)에 속한 '모든' 품목의
                # 주문번호 합집합이라 다른 품목 근거까지 섞여 오인시킬 수 있고, 대형 이벤트에서는
                # 수십~수백 건이라 일부만 잘라내면 특정 윈도우의 근거가 통째로 누락될 수 있다.
                item_orders[ln["K코드"]].add(ln["주문번호"])

        rows = []
        for kcode, obs in item_obs.items():
            info = master.get(kcode, {})
            any_multi = any(o["다중대역"] for o in obs)
            # 다중대역(여러 주파수 혼재) 윈도우의 관측치 1건이 섞였다고 해서 나머지 깨끗한
            # 관측치까지 통째로 '참고'로 끌어내리지 않는다. 오염되지 않은(다중대역 아닌)
            # 관측치만 따로 모아 그것만으로 신뢰도를 판정하고, 대표 수량도 그것을 우선한다.
            clean_obs = [o for o in obs if not o["다중대역"]]
            basis_obs = clean_obs if clean_obs else obs
            level_kr, rep, indep_obs = classify_confidence(basis_obs)
            level = "확정" if kcode in named_kcodes else ("유력 후보" if level_kr == "유력" else "참고 후보")
            rep_display = rep if rep is not None else "변동, 공급사 확인 필요"
            obs_list_str = "; ".join(f"{o['수량']}/{o['기준']}(={o['비율']:.2f})"
                                      + ("[다중대역]" if o["다중대역"] else "") for o in obs[:6])
            notes = []
            if any(o["주파수귀속"] == "무태깅(공통추정)" for o in obs):
                notes.append("품목 자체에 주파수 표기 없음(Chassis/Modem/License 등 공통성 품목) - "
                              "동일 윈도우에 다른 대역이 섞이면 귀속이 불확실할 수 있음")
            if any_multi:
                notes.append("일부 근거 윈도우가 여러 주파수 대역 혼재 - Chassis 귀속 불확실(확인 필요)")
            note = "; ".join(notes)
            rows.append({
                "K코드": kcode, "품명": info.get("품명"), "품목군": info.get("품목군"),
                "분류": category_of(info), "대표수량": rep_display, "판정수준": level,
                "관측횟수": len(item_windows[kcode]), "독립관측횟수": len(indep_obs),
                "관측치": obs_list_str,
                "윈도우ID목록": ", ".join(sorted(item_windows[kcode])),
                "주문번호목록": (", ".join(sorted(item_orders[kcode])[:20])
                          + (f" 외 {len(item_orders[kcode]) - 20}건" if len(item_orders[kcode]) > 20 else "")),
                "판매단가": info.get("판매단가", 0), "매입단가": info.get("매입단가", 0), "비고": note,
            })

        rows.sort(key=lambda x: (x["판정수준"] != "유력 후보", x["분류"], -x["관측횟수"]))
        for row in rows:
            r = ws.max_row + 1
            ws.append([row["K코드"], row["품명"], row["품목군"], row["분류"], row["대표수량"],
                       row["판정수준"], row["관측횟수"], row["독립관측횟수"], row["관측치"],
                       row["윈도우ID목록"], row["주문번호목록"], row["판매단가"], row["매입단가"],
                       row["비고"]])
            mark_font(ws, r, 1, ORIGIN_FONT)
            if row["판정수준"] == "유력 후보":
                mark_fill(ws, r, 6, CONFIRMED_FILL)
            elif row["대표수량"] == "변동, 공급사 확인 필요":
                mark_fill(ws, r, 5, ERROR_FILL)
                mark_fill(ws, r, 6, REVIEW_FILL)
            else:
                mark_fill(ws, r, 6, REVIEW_FILL)
        for c in [12, 13]:
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=c).number_format = FMT_AMOUNT
        finalize_sheet(ws, 1, len(headers), ws.max_row)
        band_sheet_rows[sheet_name] = rows

    # ================= 공용자재_제외 =================
    ws_excl = wb.create_sheet("공용자재_제외")
    headers_excl = ["K코드", "품명", "품목군", "분류", "최대 관측 수량", "제외 사유"]
    ws_excl.append(headers_excl)
    excl_seen = set()
    for sheet_name, rows in band_sheet_rows.items():
        for row in rows:
            if row["분류"] == "공용 또는 예비품" and row["K코드"] not in excl_seen:
                excl_seen.add(row["K코드"])
                info = master.get(row["K코드"], {})
                ws_excl.append([row["K코드"], row["품명"], row["품목군"], row["분류"],
                                 row["관측치"], "케이블/커넥터/SFP 등 공용·소모성 자재로, 대량 일괄구매가 "
                                 "섞여 있어 핵심 링크 BoM 수량 산정에서 제외하고 별도 관리"])
    if ws_excl.max_row == 1:
        ws_excl.append(["제외 대상 없음"])
    finalize_sheet(ws_excl, 1, len(headers_excl), ws_excl.max_row)

    # ================= 확인필요 =================
    ws_check = wb.create_sheet("확인필요")
    headers_check = ["구분", "내용"]
    ws_check.append(headers_check)
    ws_check.append(["6+0/8+0(IAP3 계열)", "8GHz/11GHz IAP3-ODU 계열에서 Chassis당 ODU쌍 비율이 "
                                        "6 또는 8로 관측되는 윈도우를 찾지 못함(발주이력에 해당 배수가 "
                                        "없음) - 이 계열의 6+0/8+0 존재 자체를 공급사에 확인 필요."])
    ws_check.append(["CTR8312=2+0 / CTR8540=8+0", "계약서에 '2+0 패킷형 IDU'/'8+0 패킷형 IDU'로 "
                                              "명시되어 있으나(구성표기_확인품목 시트), 1개 링크 "
                                              "기준 전체 수량표는 원문에 없어 '표준 BoM 확정'까지는 "
                                              "이르지 못함 - 공급사 표준BoM 회신으로 확정 필요."])
    ws_check.append(["WBX SD OPTION 결합관계", "Diversity_옵션 시트의 모든 품목 - 기본 구성에 추가되는지 "
                                          "특정 품목을 대체하는지 불명."])
    for sheet_name, rows in band_sheet_rows.items():
        for row in rows:
            if row["대표수량"] == "변동, 공급사 확인 필요" or row["판정수준"] == "참고 후보":
                ws_check.append([f"{sheet_name} - {row['K코드']}({row['품명']})",
                                  f"판정수준={row['판정수준']}, 대표수량={row['대표수량']}. "
                                  f"관측: {row['관측치']}. {row['비고']}"])
    finalize_sheet(ws_check, 1, 2, ws_check.max_row)
    ws_check.column_dimensions["B"].width = 100

    # ================= 구성요약 =================
    ws_sum = wb.create_sheet("구성요약")
    headers_sum = ["구성", "관련품목수", "확정(구성표기)", "유력후보", "참고후보", "변동(대표수량 보류)",
                   "완성 BoM 여부", "가격비교 가능여부", "비고"]
    ws_sum.append(headers_sum)
    core_groups_needed = {"핵심 필수품목"}
    for sheet_name, rows in band_sheet_rows.items():
        n_total = len(rows)
        n_named = sum(1 for r in rows if r["판정수준"] == "확정")
        n_strong = sum(1 for r in rows if r["판정수준"] == "유력 후보")
        n_ref = sum(1 for r in rows if r["판정수준"] == "참고 후보")
        n_var = sum(1 for r in rows if r["대표수량"] == "변동, 공급사 확인 필요")
        core_present = {r["품목군"] for r in rows if r["분류"] == "핵심 필수품목"}
        has_chassis = any(r["품목군"] == "IDU/Chassis" for r in rows)
        has_odu = any(r["품목군"] == "ODU/RFU" for r in rows)
        has_modem_or_license = any(r["품목군"] in ("Modem", "License") for r in rows)
        complete = has_chassis and has_odu and has_modem_or_license
        completeness = ("부분(핵심 항목 일부만 확인)" if complete and (n_strong == 0)
                         else ("완성 근접(유력후보 존재)" if complete else "미완성(핵심 품목 일부 누락)"))
        price_ok = "가능(참고, 대표수량 있는 품목 한정)" if (n_named + n_strong + n_ref - n_var) > 0 else "불가"
        note = "핵심 품목군: " + ", ".join(sorted(core_present)) if core_present else "핵심 품목 없음"
        # 유력 후보가 아직 없어도(독립반복 미확인), 단일 이벤트 하나에 핵심 품목이 여럿 함께
        # 관측되면 그 자체로 주목할 만한 근거다(돈현님 피드백: "8GHz IAP3가 어느 구성에서
        # 반복 관측되는가" 질문에 대한 1차 답). 이 단일 이벤트가 두 번째로 독립 반복되면
        # '유력 후보'로 자동 승격되는 구조다.
        if n_strong == 0:
            window_hits = Counter()
            for r in rows:
                if r["분류"] == "핵심 필수품목" and r["대표수량"] != "변동, 공급사 확인 필요":
                    for w in r["윈도우ID목록"].split(", "):
                        if w and "+" not in w:
                            window_hits[w] += 1
            if window_hits:
                top_window, top_count = window_hits.most_common(1)[0]
                if top_count >= 5:
                    note += (f" - 단일 이벤트 {top_window} 1건에서 핵심 품목 {top_count}개가 함께 "
                             f"관측됨(해당 시트 참조). 아직 독립 반복 이벤트가 없어 '유력 후보'는 "
                             f"아니지만, 공급사 확인 시 우선순위가 높은 후보.")
        ws_sum.append([sheet_name, n_total, n_named, n_strong, n_ref, n_var, completeness, price_ok, note])
        if not complete:
            mark_fill(ws_sum, ws_sum.max_row, 7, ERROR_FILL)
    finalize_sheet(ws_sum, 1, len(headers_sum), ws_sum.max_row)

    # ================= 산출근거 =================
    ws_basis = wb.create_sheet("산출근거")
    ws_basis.append(["항목", "내용"])
    methodology = [
        ("v1 대비 변경 사유", "돈현님 피드백(2026-09-04): 기존 04번 파일은 WBX 채널수/CTR 텍스트만 "
                          "앵커로 사용해 VR4/VR10/IAP3-ODU 계열(체결식 구성표기가 없는 계열)이 "
                          "2+0/4+0 후보에서 누락됨. 이 파일은 ODU LOW=HIGH 쌍수를 동반 Chassis 수로 "
                          "나눈 값을 'Chassis당 채널수(N)'로 역산하는 2차 신호를 추가해 재구성함."),
        ("분석 단위(그룹화)", "① 동일 주문번호(본 데이터에서는 품목행과 1:1이라 사실상 무의미), "
                          "② 동일 주문일자+주문자+배송지(2차 이벤트), ③ 동일 주문자+배송지 조합에서 "
                          "1~7일 이내 후속 이벤트, ④ 8~30일 이내 후속 이벤트. ③④는 인접 이벤트의 "
                          "수량을 합산한 '결합윈도우'로 별도 계산해 근거를 보강함(91일 이상 격차는 "
                          "결합하지 않음)."),
        ("N(=N+0의 N) 산출", "이벤트(또는 결합윈도우) 내 8GHz 또는 11GHz IAP3 ODU LOW수량과 HIGH수량이 "
                          "같으면 그 값을, 다르면 max(LOW,HIGH)를 N후보로 직접 채택함(2/4/6/8만 인정). "
                          "시행착오: 처음에는 이 값을 동반 VR4/VR10 Chassis 수량으로 나눈 값을 N으로 "
                          "썼으나, Chassis 오분류 버그(아래 '한계' 항목 참조)를 고친 뒤 재계산하면 "
                          "비율이 거의 전부 1로 나와 2+0/4+0 근거가 사라졌다. 반면 FLAT HYBRID/IAP3 "
                          "Mounting Bracket 수량은 'ODU쌍수'와 거의 1:1로 함께 움직여, IAP3 결합 "
                          "구조에서는 ODU쌍수 자체가 N에 더 가까운 신호로 판단해 채택함. 동반 Chassis "
                          "수량은 참고용 보조 지표로 각 행에 별도 표기함(Chassis 1대의 수용 채널수는 "
                          "확인 필요 사항으로 남김)."),
        ("확정 등급 재정의", "'확정'은 품명/세부규격에 해당 구성 표기(2+0/4+0/8+0, 4·8채널)가 문자 "
                         "그대로 있는 품목에만 부여(=구성표기 확인, 멤버십 확정이며 수량 확정은 "
                         "아님). '표준 BoM 확정'(수량까지 원문 확정)은 현재 자료에서 해당 사항 없음. "
                         "'유력 후보'=서로 다른 2개 이상 윈도우에서 비율 변동폭이 30% 미만으로 "
                         "반복 관측(다중대역 혼재 윈도우 제외). '참고 후보'=그 외 전부."),
        ("30% 변동 규칙", "품목별 관측 비율의 (최대-최소)/최소가 30% 이상이면 대표 수량을 매기지 "
                        "않고 '변동, 공급사 확인 필요'로 표시함(요청 반영)."),
        ("공용자재 분리", "품목군이 Cable/Connector/SFP/Rack/PC/NMS/Tool/기타인 품목은 '공용 또는 "
                       "예비품'으로 분류해 핵심 세트 표에서 별도 시트(공용자재_제외)로 이동함."),
        ("가격 비교 가능 범위 정정", "8GHz/11GHz는 Aviat·Ceragon 양쪽에 자료가 존재하므로 '대역이 "
                                 "겹치지 않는다'고 표현하지 않는다. 다만 Aviat 측 표는 아직 "
                                 "'참고/유력 후보' 수준이고 '표준 BoM 확정' 단계에 이르지 못했으므로, "
                                 "'대역은 겹치지만 Aviat 완성 BoM 미확정으로 구성 총액 비교가 "
                                 "보류됨'으로 정정함(07번 파일에도 반영 필요)."),
        ("이번 재작업에서 고친 선행 버그", "01번 정규화 단계의 품목군 분류 로직이 '인터페이스' "
                                  "키워드를 Chassis 판정보다 먼저 검사해, VR4/VR10 CHASSIS가 "
                                  "'Interface Card'로 오분류되어 있었다(스펙 설명에 '패킷 인터페이스' "
                                  "문구 포함). 이 버그 때문에 v1 시점에는 Chassis 동반 여부 자체를 "
                                  "제대로 셀 수 없었다. 순서를 Chassis 판정이 먼저 오도록 고치고 "
                                  "01/03/04/04b를 모두 재실행함."),
        ("한계", "발주이력 245건 대다수가 여러 대역·여러 현장 물량이 혼재된 이벤트여서, N 산출과 "
                "품목별 비율 모두 근사치다. 이 파일의 모든 수치는 공급사 표준BoM 회신으로 "
                "재확인이 필요하다."),
    ]
    for k, v in methodology:
        ws_basis.append([k, v])
    finalize_sheet(ws_basis, 1, 2, ws_basis.max_row)
    ws_basis.column_dimensions["B"].width = 100

    sheet_order = ["구성요약", "구성표기_확인품목", "8GHz_IAP3_2+0", "8GHz_IAP3_4+0",
                   "11GHz_IAP3_2+0", "11GHz_IAP3_4+0", "WTM_WBX_기본구성", "Diversity_옵션",
                   "공용자재_제외", "확인필요", "산출근거"]
    assert set(sheet_order) == set(wb.sheetnames), (sheet_order, wb.sheetnames)
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정_v2.xlsx"
    wb.save(out_path)

    print("=== 5단계 v2: 8GHz/11GHz IAP3 구성 BoM 재구성 완료 ===")
    print(f"입력 파일: {SRC_ORDERS.name}, {SRC_ITEMS.name}")
    print(f"처리: 2차 이벤트 {len(events)}건, 결합윈도우 포함 총 신호 {len(signals)}건, 체인 {len(chains)}개")
    for sheet_name, rows in band_sheet_rows.items():
        n_strong = sum(1 for r in rows if r["판정수준"] == "유력 후보")
        n_ref = sum(1 for r in rows if r["판정수준"] == "참고 후보")
        n_var = sum(1 for r in rows if r["대표수량"] == "변동, 공급사 확인 필요")
        print(f"  - {sheet_name}: {len(rows)}품목 (유력 {n_strong}, 참고 {n_ref}, 변동표시 {n_var})")
    print(f"생성 파일: {out_path}")

    covered = set()
    for rows in band_sheet_rows.values():
        covered.update(r["K코드"] for r in rows)
    missing_watchlist = {k: v for k, v in CORE_WATCHLIST.items() if k not in covered}
    print(f"돈현님 지정 핵심 관심품목 13종 중 4개 후보표에 반영된 품목: {len(CORE_WATCHLIST) - len(missing_watchlist)}/{len(CORE_WATCHLIST)}")
    if missing_watchlist:
        print(f"  미반영 품목(발주이력에서 해당 대역 윈도우 매칭 없음): {missing_watchlist}")


if __name__ == "__main__":
    main()
