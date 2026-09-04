"""
4단계: Aviat 발주 이벤트 및 동시 발주 분석
입력: output/02_Aviat_발주이력_정제.xlsx (정제데이터), output/01_Aviat_213종_정규화.xlsx (품목정규화)
출력: output/03_Aviat_발주이벤트_분석.xlsx
시트: 이벤트요약 / 이벤트별_BoM / 후속발주_후보 / 대량일괄발주 / 단품·예비품 / 판정근거

1차 이벤트(동일 주문번호)는 245건 전부가 1:1 결합키여서 사실상 무의미하므로,
분석의 중심은 2차 이벤트 후보(동일 주문일자+주문자+부서+배송지)에 둔다.
배송지/주문자가 같아도 대량 수량이 확인되면 세트로 확정하지 않는다(후보로만 표시).
"""
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (
    finalize_sheet, ORIGIN_FONT, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL,
    FMT_AMOUNT, FMT_DATE, FMT_QTY, mark_fill, mark_font,
)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_ORDERS = OUTPUT_DIR / "02_Aviat_발주이력_정제.xlsx"
SRC_ITEMS = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"

# 휴리스틱 임계값(확정 규칙이 아니라 후보 산출용, 판정근거 시트에도 명시)
BULK_QTY_THRESHOLD = 15       # 이 값 이상인 라인이 하나라도 있으면 대량일괄발주 후보
BULK_DISTINCT_THRESHOLD = 25  # 이벤트 내 서로 다른 K코드 종류가 이 값 이상이면 복수현장/노드 가능성
FOLLOWUP_GAP_DAYS = 90        # 이 이내면 후속발주 연결 후보로 간주(그 이상은 별개 취급)


def gap_bucket(days):
    if days == 0:
        return "당일"
    if days <= 7:
        return "1~7일"
    if days <= 30:
        return "8~30일"
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
            "품명": row[idx["품명"]],
            "품목군": row[idx["품목군"]],
            "주파수": row[idx["주파수"]],
            "구성표기": row[idx["구성표기(2+0등)"]],
            "채널수": row[idx["채널수"]],
            "SD표기": row[idx["SD_Diversity표기"]],
        }
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
            "결합키": row[idx["결합키(주문번호_품목번호)"]],
            "주문번호": row[idx["주문번호"]],
            "주문일자": row[idx["주문일자"]],
            "주문자": row[idx["주문자"]],
            "부서": row[idx["부서"]],
            "배송지": row[idx["배송지"]],
            "K코드": row[idx["K코드(보정반영)"]],
            "품명": row[idx["품명"]],
            "수량": row[idx["수량"]] or 0,
            "판매단가": row[idx["판매단가"]] or 0,
            "판매금액": row[idx["판매금액"]] or 0,
            "매입단가": row[idx["매입단가"]] or 0,
            "매입금액": row[idx["매입금액"]] or 0,
            "전량입고여부": row[idx["전량입고여부"]],
            "반품교환여부": row[idx["반품교환여부"]],
        })
    return orders


def classify_event(lines, master):
    """lines: 해당 이벤트에 속한 주문 라인 리스트. 품목군 집합/최대수량 등으로 후보유형 판정."""
    groups_present = set()
    freqs_present = set()
    max_qty = 0
    total_qty = 0
    distinct_kcodes = set()
    for ln in lines:
        info = master.get(ln["K코드"], {})
        g = info.get("품목군")
        if g:
            groups_present.add(g)
        f = info.get("주파수")
        if f:
            freqs_present.add(f)
        max_qty = max(max_qty, ln["수량"] or 0)
        total_qty += ln["수량"] or 0
        distinct_kcodes.add(ln["K코드"])

    has_chassis = "IDU/Chassis" in groups_present
    has_odu = "ODU/RFU" in groups_present
    has_license = "License" in groups_present
    has_modem = "Modem" in groups_present
    has_interface = "Interface Card" in groups_present
    has_mount = "Mount/Bracket" in groups_present
    has_cable = "Cable" in groups_present

    is_bulk = max_qty >= BULK_QTY_THRESHOLD or len(distinct_kcodes) >= BULK_DISTINCT_THRESHOLD

    reasons = []
    if is_bulk:
        reasons.append(
            f"최대 단일품목 수량 {max_qty}건(임계값 {BULK_QTY_THRESHOLD} 이상) 또는 "
            f"서로 다른 K코드 {len(distinct_kcodes)}종(임계값 {BULK_DISTINCT_THRESHOLD} 이상) "
            f"→ 단일 링크 세트라기보다 창고성/복수현장·복수노드 일괄발주 가능성이 높음."
        )
        candidate = {
            "신규설치후보": "참고(대량 혼재)", "증설후보": "참고(대량 혼재)", "교체후보": "판단불가",
            "예비품또는공용자재후보": "가능성 있음", "복수현장일괄발주후보": "유력",
        }
        confidence = "참고 후보"
    elif has_chassis and has_odu and (has_license or has_modem):
        reasons.append(
            f"Chassis+ODU/RFU+{'License' if has_license else ''}{'/' if has_license and has_modem else ''}"
            f"{'Modem' if has_modem else ''} 등 핵심 구성군이 함께 나타나 신규 설치 세트 가능성이 있음. "
            f"(포함 품목군: {', '.join(sorted(groups_present))})"
        )
        candidate = {
            "신규설치후보": "유력", "증설후보": "낮음", "교체후보": "낮음",
            "예비품또는공용자재후보": "낮음", "복수현장일괄발주후보": "확인 필요(배송지 단독 근거 불가)",
        }
        confidence = "유력 후보"
    elif not has_chassis and (has_odu or has_license or has_interface):
        reasons.append(
            f"신규 Chassis 없이 ODU/License/Interface Card 등 일부 구성군만 나타남 "
            f"→ 기존 설치에 대한 증설·교체 가능성. (포함 품목군: {', '.join(sorted(groups_present))})"
        )
        candidate = {
            "신규설치후보": "낮음", "증설후보": "유력", "교체후보": "참고",
            "예비품또는공용자재후보": "참고", "복수현장일괄발주후보": "확인 필요",
        }
        confidence = "참고 후보"
    elif len(distinct_kcodes) <= 2:
        reasons.append(f"서로 다른 품목 {len(distinct_kcodes)}종만 존재 → 단품 또는 예비품 보충으로 추정.")
        candidate = {
            "신규설치후보": "낮음", "증설후보": "낮음", "교체후보": "참고",
            "예비품또는공용자재후보": "유력", "복수현장일괄발주후보": "낮음",
        }
        confidence = "참고 후보"
    else:
        reasons.append("품목 구성이 뚜렷한 패턴에 맞지 않아 자동 판정 불가.")
        candidate = {
            "신규설치후보": "판단불가", "증설후보": "판단불가", "교체후보": "판단불가",
            "예비품또는공용자재후보": "판단불가", "복수현장일괄발주후보": "판단불가",
        }
        confidence = "판단 불가"

    return {
        "품목군목록": sorted(groups_present), "주파수목록": sorted(freqs_present),
        "최대수량": max_qty, "총수량": total_qty, "고유K코드수": len(distinct_kcodes),
        "대량여부": is_bulk, "판정": candidate, "판단근거": " ".join(reasons), "신뢰도": confidence,
    }


def main():
    master = load_item_master()
    orders = load_orders()

    # ---- 1차 이벤트(주문번호 단위) 통계만 산출 ----
    order_no_groups = defaultdict(list)
    for o in orders:
        order_no_groups[o["주문번호"]].append(o)
    n_1st = len(order_no_groups)
    n_1st_multi = sum(1 for v in order_no_groups.values() if len(v) > 1)

    # ---- 2차 이벤트 후보: 동일 주문일자+주문자+부서+배송지 ----
    key_fn = lambda o: (o["주문일자"], o["주문자"], o["부서"], o["배송지"])
    groups = defaultdict(list)
    for o in orders:
        groups[key_fn(o)].append(o)

    events = []
    for i, (key, lines) in enumerate(sorted(groups.items(), key=lambda x: (x[0][0] or date.min)), start=1):
        order_date, orderer, dept, addr = key
        cls = classify_event(lines, master)
        events.append({
            "이벤트ID": f"EVT-{i:03d}",
            "최초주문일": order_date, "최종주문일": order_date, "발주간격": "당일",
            "주문자": orderer, "부서": dept, "배송지": addr,
            "주문번호목록": sorted(set(o["주문번호"] for o in lines)),
            "라인": lines,
            **cls,
            "판매총액": sum(o["판매금액"] for o in lines),
            "매입총액": sum(o["매입금액"] for o in lines),
        })

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 이벤트요약 ----------------
    ws1 = wb.create_sheet("이벤트요약")
    ws1.append([f"※ 1차 이벤트(동일 주문번호) 참고: 고유 주문번호 {n_1st}건 중 품목행이 2건 이상인 주문번호는 {n_1st_multi}건"
                "(즉 1차 이벤트는 사실상 245건 모두 1:1 결합키). 아래는 2차 이벤트 후보(동일 주문일자+주문자+부서+배송지) 기준."])
    ws1.append([])
    headers1 = ["이벤트ID", "최초주문일", "최종주문일", "발주간격", "주문자", "부서", "배송지",
                "주문번호수", "포함K코드수", "품목행수", "총수량", "판매총액", "매입총액",
                "포함품목군", "포함주파수", "신규설치후보", "증설후보", "교체후보",
                "예비품또는공용자재후보", "복수현장일괄발주후보", "판단근거", "신뢰도"]
    ws1.append(headers1)
    header_row_idx = 3
    for ev in events:
        r = ws1.max_row + 1
        p = ev["판정"]
        ws1.append([
            ev["이벤트ID"], ev["최초주문일"], ev["최종주문일"], ev["발주간격"], ev["주문자"], ev["부서"],
            ev["배송지"], len(ev["주문번호목록"]), ev["고유K코드수"], len(ev["라인"]), ev["총수량"],
            ev["판매총액"], ev["매입총액"], ", ".join(ev["품목군목록"]), ", ".join(ev["주파수목록"]),
            p["신규설치후보"], p["증설후보"], p["교체후보"], p["예비품또는공용자재후보"],
            p["복수현장일괄발주후보"], ev["판단근거"], ev["신뢰도"],
        ])
        for c in [5, 6, 7]:
            mark_font(ws1, r, c, ORIGIN_FONT)
        ws1.cell(row=r, column=2).number_format = FMT_DATE
        ws1.cell(row=r, column=3).number_format = FMT_DATE
        if ev["대량여부"]:
            mark_fill(ws1, r, 20, REVIEW_FILL)
        if ev["신뢰도"] == "유력 후보":
            mark_fill(ws1, r, 22, CONFIRMED_FILL)
    for c in [11, 12, 13]:
        for r in range(header_row_idx + 1, ws1.max_row + 1):
            ws1.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws1, header_row_idx, len(headers1), ws1.max_row)

    # ---------------- 이벤트별_BoM ----------------
    ws2 = wb.create_sheet("이벤트별_BoM")
    headers2 = ["이벤트ID", "주문일자", "K코드", "품명", "품목군", "주파수", "구성표기(원문)",
                "채널수(원문)", "수량", "판매단가", "판매금액", "매입단가", "매입금액"]
    ws2.append(headers2)
    for ev in events:
        for ln in ev["라인"]:
            info = master.get(ln["K코드"], {})
            r = ws2.max_row + 1
            ws2.append([
                ev["이벤트ID"], ev["최초주문일"], ln["K코드"], ln["품명"], info.get("품목군"),
                info.get("주파수"), info.get("구성표기"), info.get("채널수"), ln["수량"],
                ln["판매단가"], ln["판매금액"], ln["매입단가"], ln["매입금액"],
            ])
            mark_font(ws2, r, 3, ORIGIN_FONT)
            mark_font(ws2, r, 4, ORIGIN_FONT)
            ws2.cell(row=r, column=2).number_format = FMT_DATE
    for c in [9]:
        for r in range(2, ws2.max_row + 1):
            ws2.cell(row=r, column=c).number_format = FMT_QTY
    for c in [10, 11, 12, 13]:
        for r in range(2, ws2.max_row + 1):
            ws2.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws2, 1, len(headers2), ws2.max_row)

    # ---------------- 후속발주_후보 ----------------
    ws3 = wb.create_sheet("후속발주_후보")
    headers3 = ["체인ID", "순번", "이벤트ID", "주문일자", "주문자", "부서", "배송지",
                "직전이벤트대비_간격일", "발주간격구분", "직전이벤트대비_추가품목군", "비고"]
    ws3.append(headers3)
    chain_groups = defaultdict(list)
    for ev in events:
        chain_groups[(ev["주문자"], ev["부서"], ev["배송지"])].append(ev)
    chain_id = 0
    for key, evs in chain_groups.items():
        if len(evs) < 2:
            continue
        chain_id += 1
        evs_sorted = sorted(evs, key=lambda e: e["최초주문일"] or date.min)
        prev_groups = set()
        prev_date = None
        for seq, ev in enumerate(evs_sorted, start=1):
            gap_days = (ev["최초주문일"] - prev_date).days if prev_date else 0
            added = sorted(set(ev["품목군목록"]) - prev_groups) if prev_date else sorted(ev["품목군목록"])
            r = ws3.max_row + 1
            ws3.append([
                f"CHAIN-{chain_id:02d}", seq, ev["이벤트ID"], ev["최초주문일"], ev["주문자"], ev["부서"],
                ev["배송지"], gap_days if prev_date else "-", gap_bucket(gap_days) if prev_date else "최초",
                ", ".join(added),
                "동일 주문자/부서/배송지 반복 → 후속발주 연결 후보(자동추정, 실제 동일 현장 여부는 확인 필요)"
                if prev_date else "체인의 최초 이벤트",
            ])
            for c in [5, 6, 7]:
                mark_font(ws3, r, c, ORIGIN_FONT)
            ws3.cell(row=r, column=4).number_format = FMT_DATE
            prev_groups = set(ev["품목군목록"])
            prev_date = ev["최초주문일"]
    if chain_id == 0:
        ws3.append(["연결 가능한 후속발주 체인 없음(모든 배송지 조합이 단일 이벤트)"])
    finalize_sheet(ws3, 1, len(headers3), max(ws3.max_row, 2))

    # ---------------- 대량일괄발주 ----------------
    ws4 = wb.create_sheet("대량일괄발주")
    headers4 = ["이벤트ID", "주문일자", "배송지", "K코드", "품명", "수량", "사유"]
    ws4.append(headers4)
    n_bulk_events = 0
    for ev in events:
        if not ev["대량여부"]:
            continue
        n_bulk_events += 1
        for ln in ev["라인"]:
            r = ws4.max_row + 1
            reason = []
            if ln["수량"] >= BULK_QTY_THRESHOLD:
                reason.append(f"단일 라인 수량 {ln['수량']}건 ≥ 임계값 {BULK_QTY_THRESHOLD}")
            if ev["고유K코드수"] >= BULK_DISTINCT_THRESHOLD:
                reason.append(f"이벤트 내 품목종류 {ev['고유K코드수']}종 ≥ 임계값 {BULK_DISTINCT_THRESHOLD}")
            ws4.append([ev["이벤트ID"], ev["최초주문일"], ev["배송지"], ln["K코드"], ln["품명"],
                        ln["수량"], "; ".join(reason)])
            ws4.cell(row=r, column=2).number_format = FMT_DATE
            if ln["수량"] >= BULK_QTY_THRESHOLD:
                mark_fill(ws4, r, 6, REVIEW_FILL)
    if n_bulk_events == 0:
        ws4.append(["대량일괄발주로 분류된 이벤트 없음"])
    finalize_sheet(ws4, 1, len(headers4), max(ws4.max_row, 2))

    # ---------------- 단품·예비품 ----------------
    ws5 = wb.create_sheet("단품·예비품")
    headers5 = ["이벤트ID", "주문일자", "주문자", "부서", "K코드", "품명", "품목군", "수량", "판정"]
    ws5.append(headers5)
    n_spare = 0
    for ev in events:
        if ev["판정"]["예비품또는공용자재후보"] not in ("유력", "가능성 있음"):
            continue
        n_spare += 1
        for ln in ev["라인"]:
            info = master.get(ln["K코드"], {})
            r = ws5.max_row + 1
            ws5.append([ev["이벤트ID"], ev["최초주문일"], ev["주문자"], ev["부서"], ln["K코드"],
                        ln["품명"], info.get("품목군"), ln["수량"], ev["판정"]["예비품또는공용자재후보"]])
            ws5.cell(row=r, column=2).number_format = FMT_DATE
    if n_spare == 0:
        ws5.append(["단품/예비품 후보로 분류된 이벤트 없음"])
    finalize_sheet(ws5, 1, len(headers5), max(ws5.max_row, 2))

    # ---------------- 판정근거 ----------------
    ws6 = wb.create_sheet("판정근거")
    ws6.append(["항목", "내용"])
    rules = [
        ("1차 이벤트 정의", "동일 주문번호. 본 데이터에서는 245건 결합키(주문번호_품목번호)가 전부 고유하여 "
                         "주문번호 자체가 이미 품목행 단위로 발급됨(1차 이벤트=행 단위, 별도 집계 의미 낮음)."),
        ("2차 이벤트 후보 정의", "동일 주문일자 + 동일 주문자 + 동일 부서 + 동일 배송지. "
                             f"총 {len(events)}건 도출."),
        ("대량일괄발주 판정 임계값(휴리스틱)", f"단일 라인 수량 ≥ {BULK_QTY_THRESHOLD} 또는 이벤트 내 "
                                     f"서로 다른 K코드 ≥ {BULK_DISTINCT_THRESHOLD}종. 확정 기준이 아니라 "
                                     "후보 산출용 규칙(7.1절 동일)."),
        ("근거 사례: 2025-06-04 이벤트", "'HAX_IDU/ODU간 IF 케이블,1Meter' 수량 7,500개, "
                                    "'HAX_Arrestor KIT' 204개 등 단일 링크가 필요로 하기 어려운 수량이 "
                                    "관측되어 대량일괄발주(창고성 소요 또는 복수현장 물량)로 분류함."),
        ("근거 사례: WBX 채널수 구분", "2025-09-17 이벤트에 'HAX_4GHz WBX 4 CHANNEL'(수량2), "
                                  "2025-09-18 이벤트에 'HAX_4GHz WBX 8 CHANNEL'(수량2)이 각각 별도 "
                                  "날짜로 발주되어, WBX 채널수가 4+0/8+0 구성 구분의 명시적 단서로 확인됨"
                                  "(단, 두 이벤트 모두 다수 품목이 혼재된 대량성 이벤트이므로 절대 수량은 "
                                  "참고용이며 5단계에서 별도 검증)."),
        ("근거 사례: 계약서 명시 2+0/8+0", "Aviat 213종 계약품목 중 'CTR 8312 1RU Chassis'(K9188144) "
                                     "세부규격에 '2+0 패킷형 IDU', 'CTR 8540 1RU Chassis'(K9188145) "
                                     "세부규격에 '8+0 패킷형 IDU'라고 명시되어 있어(확정), 이 두 K코드는 "
                                     "각각 2+0/8+0 구성의 핵심 Chassis로 확정 처리함."),
        ("배송지 신뢰도 한계", "245건 전체에서 빌딩명/국사명/프로젝트명이 100% 공란이며, 상위 배송지 "
                          "1건('경기 안양시 ***')이 188건(전체의 77%)을 차지함 → 배송지는 실제 설치 "
                          "국소가 아닌 창고/집하지일 가능성이 높아, 배송지 일치만으로 세트를 확정하지 않음."),
        ("주문자 표기 불일치", "'이상*'과 '이상헌'은 이메일 패턴상 동일인일 가능성이 있으나 부서가 달라 "
                          "임의 병합하지 않음(확인 필요, 02번 파일 주문요약 참조)."),
    ]
    for k, v in rules:
        ws6.append([k, v])
    finalize_sheet(ws6, 1, 2, ws6.max_row)
    ws6.column_dimensions["B"].width = 100

    out_path = OUTPUT_DIR / "03_Aviat_발주이벤트_분석.xlsx"
    wb.save(out_path)

    n_new = sum(1 for e in events if e["판정"]["신규설치후보"] == "유력")
    n_expand = sum(1 for e in events if e["판정"]["증설후보"] == "유력")
    n_bulk = sum(1 for e in events if e["대량여부"])
    n_spare_ev = sum(1 for e in events if e["판정"]["예비품또는공용자재후보"] in ("유력", "가능성 있음"))

    print("=== 4단계: 발주 이벤트 및 동시발주 분석 완료 ===")
    print(f"입력 파일: {SRC_ORDERS.name}, {SRC_ITEMS.name}")
    print(f"처리 행 수: 주문라인 {len(orders)}건 → 2차 이벤트 후보 {len(events)}건")
    print(f"제외 또는 오류 행 수: 0건")
    print(f"생성 파일: {out_path}")
    print("주요 검증 결과:")
    print(f"  - 1차 이벤트(주문번호 단위): {n_1st}건 (품목행 2건 이상: {n_1st_multi}건, 사실상 전부 1:1)")
    print(f"  - 2차 이벤트 후보: {len(events)}건")
    print(f"  - 신규설치 유력 후보: {n_new}건, 증설 유력 후보: {n_expand}건")
    print(f"  - 대량일괄발주 분류: {n_bulk}건")
    print(f"  - 단품/예비품 후보: {n_spare_ev}건")
    print(f"  - 후속발주 연결 체인: {chain_id}개")


if __name__ == "__main__":
    main()
