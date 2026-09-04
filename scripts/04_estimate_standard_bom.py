"""
5단계: Aviat 구성방식별(2+0/4+0/6+0/8+0) 표준 BoM 추정
입력: output/03_Aviat_발주이벤트_분석.xlsx (이벤트별_BoM, 이벤트요약)
출력: output/04_Aviat_구성방식별_표준BoM_추정.xlsx
시트: 구성요약 / 2+0_BoM / 4+0_BoM / 6+0_BoM / 8+0_BoM / 공통품목 / 선택품목 / 확인필요 / 산출근거

방법론(산출근거 시트에도 동일 내용 기록):
1. Aviat 213종 품목 중 '구성표기'(세부규격에 2+0/8+0이 명시된 CTR8312/8540) 또는 '채널수'
   (WBX류의 4채널/8채널 표기)가 확정적으로 존재하는 품목을 각 구성의 '앵커 품목'으로 삼는다.
2. 발주이벤트 중 앵커 품목이 포함된 이벤트를 찾고, 그 이벤트 내에서 앵커와 동일한 '주파수'
   태그를 가진 품목만 같은 링크의 구성품 후보로 본다(서로 다른 주파수 장비가 한 이벤트에
   섞여 있는 경우가 많으므로, 주파수가 다른 품목을 같은 구성으로 묶지 않기 위함).
3. 후보 품목의 수량을 앵커 수량으로 나눈 비율을 이벤트별로 기록하고, 여러 이벤트에서 동일한
   비율이 반복되면 신뢰도를 높인다(확정 규칙이 아니라 후보 산출용 절차).
4. 주파수 태그가 없는 품목(License, PC, Cable 등 공용성 품목)은 특정 구성 전용으로 확정하지
   않고 '공통품목' 시트에 별도로 정리한다.
5. 6+0에 대한 명시적 앵커/채널 표기가 Aviat 자료에 없으면 임의로 만들지 않고 '확인필요'로 남긴다.
"""
import sys
from collections import defaultdict
from pathlib import Path
from statistics import median

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (
    finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL,
    FMT_AMOUNT, FMT_QTY, mark_fill,
)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_EVENTS = OUTPUT_DIR / "03_Aviat_발주이벤트_분석.xlsx"
SRC_ITEMS = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"

CONFIGS = ["2+0", "4+0", "6+0", "8+0"]
CHANNEL_TO_CONFIG = {"4채널": "4+0", "8채널": "8+0"}


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
            "장비계열": row[idx["장비계열"]] or "",
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


def load_event_lines():
    wb = openpyxl.load_workbook(SRC_EVENTS, data_only=True)
    ws = wb["이벤트별_BoM"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    lines = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["이벤트ID"]] is None:
            continue
        lines.append({
            "이벤트ID": row[idx["이벤트ID"]], "K코드": row[idx["K코드"]], "품명": row[idx["품명"]],
            "품목군": row[idx["품목군"]], "주파수": row[idx["주파수"]] or "", "수량": row[idx["수량"]] or 0,
        })
    ws2 = wb["이벤트요약"]
    headers2 = [c.value for c in ws2[3]]
    idx2 = {h: i for i, h in enumerate(headers2)}
    event_meta = {}
    for row in ws2.iter_rows(min_row=4, values_only=True):
        if row[idx2["이벤트ID"]] is None:
            continue
        event_meta[row[idx2["이벤트ID"]]] = {
            "신뢰도": row[idx2["신뢰도"]], "최초주문일": row[idx2["최초주문일"]],
        }
    return lines, event_meta


def find_anchors(master):
    """구성표기 또는 채널수가 명시적으로 확인되는 품목을 구성별 앵커로 식별."""
    anchors = defaultdict(list)  # config -> [(kcode, 근거)]
    for kcode, info in master.items():
        if info["구성표기"] in CONFIGS:
            anchors[info["구성표기"]].append((kcode, f"세부규격에 '{info['구성표기']}' 명시(확정)"))
        if info["채널수"] in CHANNEL_TO_CONFIG and info["품목군"] == "WBX":
            cfg = CHANNEL_TO_CONFIG[info["채널수"]]
            anchors[cfg].append((kcode, f"WBX {info['채널수']} 표기(확정, {info['품명']})"))
    return anchors


def main():
    master = load_item_master()
    lines, event_meta = load_event_lines()
    anchors = find_anchors(master)

    lines_by_event = defaultdict(list)
    for ln in lines:
        lines_by_event[ln["이벤트ID"]].append(ln)

    # observations[config][kcode] = list of dict(이벤트ID, 비율, 앵커K코드, 앵커수량, 품목수량)
    observations = defaultdict(lambda: defaultdict(list))
    anchor_membership_notes = defaultdict(dict)  # config -> kcode -> 근거설명(앵커 자신)

    for config, anchor_list in anchors.items():
        anchor_kcodes = {k for k, _ in anchor_list}
        for k, note in anchor_list:
            anchor_membership_notes[config][k] = note
        for event_id, evlines in lines_by_event.items():
            present_anchors = [ln for ln in evlines if ln["K코드"] in anchor_kcodes and ln["수량"] > 0]
            for anc in present_anchors:
                anc_freq = anc["주파수"]
                anc_family = master.get(anc["K코드"], {}).get("장비계열", "")
                for ln in evlines:
                    if ln["K코드"] == anc["K코드"]:
                        continue
                    ln_family = master.get(ln["K코드"], {}).get("장비계열", "")
                    if anc_freq != "":
                        # 앵커에 주파수 태그가 있으면 동일 주파수 품목만 후보로 본다.
                        if ln["주파수"] != anc_freq:
                            continue
                    else:
                        # 앵커에 주파수 태그가 없는 경우(CTR8312/8540 등) 상대 품목도 주파수가
                        # 없어야 하고, 장비계열 키가 있다면 반드시 앵커와 같은 계열이어야 한다.
                        # (예: CTR8312 옆에 CTR8540 부속이 섞여 있어도 잘못 묶이지 않도록)
                        if ln["주파수"] != "":
                            continue
                        if anc_family and ln_family and anc_family != ln_family:
                            continue
                        if anc_family and not ln_family:
                            continue
                    if ln["수량"] <= 0:
                        continue
                    ratio = ln["수량"] / anc["수량"]
                    observations[config][ln["K코드"]].append({
                        "이벤트ID": event_id, "비율": ratio, "앵커K코드": anc["K코드"],
                        "앵커수량": anc["수량"], "품목수량": ln["수량"],
                        "이벤트신뢰도": event_meta.get(event_id, {}).get("신뢰도", ""),
                    })

    # ---- 공통품목 후보: 주파수 태그가 없는 품목이 여러 config에서 관측된 경우 ----
    common_candidates = defaultdict(set)  # kcode -> set(config)
    for config, kmap in observations.items():
        for kcode in kmap:
            if master.get(kcode, {}).get("주파수", "") == "":
                common_candidates[kcode].add(config)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def confidence_of(obs_list):
        distinct_events = len(set(o["이벤트ID"] for o in obs_list))
        ratios = [round(o["비율"], 2) for o in obs_list]
        ratio_spread = (max(ratios) / min(ratios)) if min(ratios) > 0 else 999
        bulk_only = all(o["이벤트신뢰도"] == "참고 후보" for o in obs_list)
        if distinct_events >= 2 and ratio_spread <= 1.3:
            return "유력 후보"
        if distinct_events >= 2:
            return "참고 후보"
        if bulk_only:
            return "참고 후보"
        return "판단 불가"

    def write_config_sheet(ws, config):
        headers = ["K코드", "품명", "모델명", "품목군", "1개 링크 기준 추정 수량(대표값)",
                   "관측 비율 범위", "판매단가", "판매금액", "매입단가", "매입금액",
                   "근거 이벤트", "반복 횟수(이벤트수)", "근거 설명", "확정 수준", "공급사 확인 필요사항"]
        ws.append(headers)
        kmap = observations.get(config, {})
        anchor_notes = anchor_membership_notes.get(config, {})
        all_kcodes = set(kmap.keys()) | set(anchor_notes.keys())
        rows = []
        for kcode in all_kcodes:
            info = master.get(kcode, {})
            if kcode in anchor_notes:
                rep_qty = "1(앵커 기준)"
                ratio_range = "-"
                events_str = ", ".join(sorted(set(
                    o["이벤트ID"] for o in kmap.get(kcode, [])
                ))) or "(계약서 텍스트 근거, 발주이력 무관)"
                n_events = len(set(o["이벤트ID"] for o in kmap.get(kcode, [])))
                basis = anchor_notes[kcode]
                level = "확정 후보"
                followup = ""
            else:
                obs = kmap[kcode]
                ratios = [o["비율"] for o in obs]
                rep = median(ratios)
                rep_qty = round(rep, 2)
                ratio_range = f"{min(ratios):.2f} ~ {max(ratios):.2f}" if len(set(round(r,2) for r in ratios)) > 1 else f"{rep:.2f}(고정)"
                events_str = ", ".join(sorted(set(o["이벤트ID"] for o in obs)))
                n_events = len(set(o["이벤트ID"] for o in obs))
                basis = (
                    f"앵커품목 대비 관측 비율(수량/앵커수량) 기준. 근거 이벤트: "
                    + "; ".join(f"{o['이벤트ID']}(품목{o['품목수량']}/앵커{o['앵커수량']})" for o in obs[:4])
                )
                level = confidence_of(obs)
                followup = "" if level == "유력 후보" else "실제 1개 링크 소요량 공급사 확인 필요(대량/복수현장 물량 혼재 가능)"
            rows.append((kcode, info, rep_qty, ratio_range, events_str, n_events, basis, level, followup))

        rows.sort(key=lambda x: (x[7] != "확정 후보", -x[5]))
        for kcode, info, rep_qty, ratio_range, events_str, n_events, basis, level, followup in rows:
            price = info.get("판매단가", 0) or 0
            buy = info.get("매입단가", 0) or 0
            r = ws.max_row + 1
            ws.append([
                kcode, info.get("품명"), info.get("품명"), info.get("품목군"), rep_qty, ratio_range,
                price, None, buy, None, events_str, n_events, basis, level, followup,
            ])
            ws.cell(row=r, column=8).value = f"=E{r}*G{r}" if isinstance(rep_qty, (int, float)) else None
            ws.cell(row=r, column=10).value = f"=E{r}*I{r}" if isinstance(rep_qty, (int, float)) else None
            if level == "확정 후보":
                mark_fill(ws, r, 14, CONFIRMED_FILL)
            elif level == "판단 불가":
                mark_fill(ws, r, 14, ERROR_FILL)
            else:
                mark_fill(ws, r, 14, REVIEW_FILL)
        for c in [7, 8, 9, 10]:
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=c).number_format = FMT_AMOUNT
        finalize_sheet(ws, 1, len(headers), ws.max_row)
        return rows

    config_rows = {}
    for config in CONFIGS:
        ws = wb.create_sheet(f"{config}_BoM")
        config_rows[config] = write_config_sheet(ws, config)

    # ---------------- 구성요약 ----------------
    ws0 = wb.create_sheet("구성요약")
    ws0.append(["구성방식", "앵커품목수", "관련품목수(공통제외)", "확정후보", "유력후보", "참고후보", "판단불가", "비고"])
    for config in CONFIGS:
        rows = config_rows[config]
        n_anchor = len(anchor_membership_notes.get(config, {}))
        n_total = len(rows)
        n_conf = sum(1 for x in rows if x[7] == "확정 후보")
        n_strong = sum(1 for x in rows if x[7] == "유력 후보")
        n_ref = sum(1 for x in rows if x[7] == "참고 후보")
        n_na = sum(1 for x in rows if x[7] == "판단 불가")
        note = "" if n_anchor else "Aviat 자료에서 이 구성을 특정하는 명시적 표기(세부규격 2+0/8+0 텍스트 또는 WBX 채널수)를 찾지 못함 → 확인 필요"
        ws0.append([config, n_anchor, n_total, n_conf, n_strong, n_ref, n_na, note])
        if n_anchor == 0:
            mark_fill(ws0, ws0.max_row, 8, ERROR_FILL)
    wb.move_sheet("구성요약", offset=-len(CONFIGS))
    finalize_sheet(ws0, 1, 8, ws0.max_row)

    # ---------------- 공통품목 ----------------
    wsc = wb.create_sheet("공통품목")
    headers_c = ["K코드", "품명", "품목군", "등장 구성", "구성수", "판매단가", "매입단가", "설명"]
    wsc.append(headers_c)
    for kcode, cfgs in sorted(common_candidates.items(), key=lambda x: -len(x[1])):
        info = master.get(kcode, {})
        r = wsc.max_row + 1
        wsc.append([
            kcode, info.get("품명"), info.get("품목군"), ", ".join(sorted(cfgs)), len(cfgs),
            info.get("판매단가", 0), info.get("매입단가", 0),
            "주파수 태그가 없어 특정 구성 전용으로 단정하지 않음(License/Cable/PC 등 공용 성격 품목일 가능성)"
            if len(cfgs) >= 2 else "단일 구성에서만 관측됨(공통 여부 확인 필요)",
        ])
    if wsc.max_row == 1:
        wsc.append(["공통품목 후보 없음"])
    for c in [6, 7]:
        for r in range(2, wsc.max_row + 1):
            wsc.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(wsc, 1, len(headers_c), wsc.max_row)

    # ---------------- 선택품목 (SD 등 옵션) ----------------
    wso = wb.create_sheet("선택품목")
    headers_o = ["K코드", "품명", "품목군", "주파수", "옵션구분", "설명"]
    wso.append(headers_o)
    for kcode, info in master.items():
        if info.get("SD표기") == "Y":
            wso.append([
                kcode, info.get("품명"), info.get("품목군"), info.get("주파수"), "SD/Diversity 관련",
                "품명 또는 세부규격에 SD(Space Diversity) 관련 표기가 확인됨. Aviat 공식 자료에서 "
                "SD/Non-SD 구분이 별도로 확정되지 않았으므로, 이 목록은 '표기가 발견된 품목'일 뿐 "
                "구성별 필수/선택 여부를 확정하지 않음(확인 필요).",
            ])
    if wso.max_row == 1:
        wso.append(["SD/Diversity 관련 표기가 확인된 품목 없음"])
    finalize_sheet(wso, 1, len(headers_o), wso.max_row)

    # ---------------- 확인필요 ----------------
    wsn = wb.create_sheet("확인필요")
    headers_n = ["구분", "내용"]
    wsn.append(headers_n)
    wsn.append(["6+0 구성", "Aviat 213종 계약품목 및 발주이력 전체에서 '6+0'을 직접 지시하는 세부규격 문구나 "
                          "WBX '6채널' 표기가 발견되지 않음. 6GHz/L6GHz/U6GHz 대역 품목(HAX_L6GHz/U6GHz ODU IAP3, "
                          "6G OBC2 CORE UNIT, 6G DUPLEXER UNIT, 6GHz FLAT HYBRID 등)은 존재하나, 이들이 "
                          "6+0 구성 전용인지 2+0/4+0의 6GHz 대역 버전인지는 현재 자료로 구분 불가 → 공급사 확인 필요."])
    for config in CONFIGS:
        for kcode, info, rep_qty, ratio_range, events_str, n_events, basis, level, followup in config_rows[config]:
            if level in ("참고 후보", "판단 불가"):
                wsn.append([f"{config} - {kcode}({info.get('품명')})",
                            f"확정수준={level}. {followup or basis}"])
    finalize_sheet(wsn, 1, 2, wsn.max_row)
    wsn.column_dimensions["B"].width = 100

    # ---------------- 산출근거 ----------------
    wsm = wb.create_sheet("산출근거")
    wsm.append(["항목", "내용"])
    methodology = [
        ("앵커 품목 식별 기준", "① 세부규격에 '2+0' 또는 '8+0'이 직접 명시된 품목(CTR 8312 1RU Chassis, "
                          "CTR 8540 1RU Chassis) - 확정. ② 품목군='WBX'이면서 품명에 '4 CHNNEL/CHANNEL' "
                          "또는 '8 CHNNEL/CHANNEL'이 명시된 품목 - 확정."),
        ("비율 계산 방법", "앵커 품목이 포함된 발주이벤트를 찾아, 그 이벤트 내에서 앵커와 동일한 '주파수' "
                       "태그를 가진 품목(또는 앵커와 마찬가지로 주파수 태그가 없는 품목)의 수량을 앵커 "
                       "수량으로 나눈 값을 관측치로 기록. 서로 다른 주파수 장비가 한 이벤트에 섞여 있는 "
                       "경우가 많아, 주파수가 다른 품목은 관측 대상에서 제외함."),
        ("확정 수준 부여 기준", "확정 후보=계약서에 구성 표기가 직접 명시된 앵커 품목 자신. 유력 후보=서로 "
                          "다른 2개 이상 이벤트에서 비율 편차가 30% 이내로 일관되게 관측. 참고 후보=1개 "
                          "이벤트에서만 관측되었거나 이벤트가 대량/복수현장 성격. 판단 불가=반복성·일관성 "
                          "모두 낮음."),
        ("한계", "발주이력 245건 대부분이 창고성 대량 또는 복수 대역·복수 현장 물량이 혼재된 이벤트여서 "
                "'1개 링크 기준' 수량은 근사치이며, 표에 기재된 수치는 공급사 회신(⑦ 표준BoM 업체작성양식)"
                "으로 반드시 재확인이 필요함. 특히 6+0 구성은 근거가 전혀 없어 전량 확인 필요로 남김."),
        ("VR4/VR10/IAP3-ODU 계열 처리", "이 계열은 세부규격에 채널수/구성 표기가 없어 앵커로 사용하지 않았음. "
                                    "다만 WBX 앵커가 포함된 이벤트에 동일 주파수로 함께 등장하는 경우에는 "
                                    "해당 구성의 후보 품목으로 포함될 수 있음(개별 근거는 각 구성 시트 "
                                    "'근거 설명' 열 참조)."),
    ]
    for k, v in methodology:
        wsm.append([k, v])
    finalize_sheet(wsm, 1, 2, wsm.max_row)
    wsm.column_dimensions["B"].width = 100

    out_path = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정.xlsx"
    wb.save(out_path)

    print("=== 5단계: 구성방식별 표준 BoM 추정 완료 ===")
    print(f"입력 파일: {SRC_EVENTS.name}, {SRC_ITEMS.name}")
    print(f"처리: 이벤트 {len(lines_by_event)}건, 앵커 품목 {sum(len(v) for v in anchors.values())}건")
    for config in CONFIGS:
        rows = config_rows[config]
        n_anchor = len(anchor_membership_notes.get(config, {}))
        print(f"  - {config}: 앵커 {n_anchor}종, 관련품목 {len(rows)}종 "
              f"(확정 {sum(1 for x in rows if x[7]=='확정 후보')}, 유력 {sum(1 for x in rows if x[7]=='유력 후보')}, "
              f"참고 {sum(1 for x in rows if x[7]=='참고 후보')}, 판단불가 {sum(1 for x in rows if x[7]=='판단 불가')})")
    print(f"생성 파일: {out_path}")


if __name__ == "__main__":
    main()
