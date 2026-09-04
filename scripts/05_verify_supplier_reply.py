"""
6단계: 공급사 작성 표준BoM 양식과 비교 검증
입력: input/7cb45e7e-Aviat_MDM_________BoM_____.xlsx (시트 '표준BoM_213종'),
      output/01_Aviat_213종_정규화.xlsx, output/04_Aviat_구성방식별_표준BoM_추정.xlsx
출력: output/05_Aviat_공급사회신_검증.xlsx

주의: 1단계 조사에서 이미 확인했듯 업체작성양식은 2+0/4+0/6+0/8+0/Diversity/비고 열이
전부 공란인 '빈 템플릿' 상태다(공급사 회신 전). 따라서 본 단계는
(1) K코드 구조 일치 여부, (2) 회신 완료 여부, (3) 발주이력 기반 추정치를 검증받기 위한
확인요청 목록을 만드는 데 집중한다. 실제 수량 비교는 회신 후 재실행이 필요하다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, mark_fill

BASE_DIR = Path("/home/user/kt-settlement-automation")
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
SRC_TEMPLATE = INPUT_DIR / "7cb45e7e-Aviat_MDM_________BoM_____.xlsx"
SRC_ITEMS = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"
SRC_BOM = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정.xlsx"

CONFIGS = ["2+0", "4+0", "6+0", "8+0"]


def load_template():
    wb = openpyxl.load_workbook(SRC_TEMPLATE, data_only=True)
    ws = wb["표준BoM_213종"]
    rows = {}
    for r in range(5, 218):
        kcode = ws.cell(row=r, column=3).value
        if not kcode:
            continue
        kcode = str(kcode).strip()
        rows[kcode] = {
            "순번": ws.cell(row=r, column=1).value,
            "주파수": ws.cell(row=r, column=2).value,
            "품명": ws.cell(row=r, column=4).value,
            "단위": ws.cell(row=r, column=5).value,
            "2+0": ws.cell(row=r, column=6).value,
            "4+0": ws.cell(row=r, column=7).value,
            "6+0": ws.cell(row=r, column=8).value,
            "8+0": ws.cell(row=r, column=9).value,
            "Diversity": ws.cell(row=r, column=10).value,
            "비고": ws.cell(row=r, column=11).value,
        }
    return rows


def load_contract_213():
    wb = openpyxl.load_workbook(SRC_ITEMS, data_only=True)
    ws = wb["원본정리"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    items = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = row[idx["K코드"]]
        if k:
            items[k] = {"순번": row[idx["순번"]], "품명": row[idx["품명"]]}
    return items


def load_estimated_bom():
    wb = openpyxl.load_workbook(SRC_BOM, data_only=True)
    result = {}
    for cfg in CONFIGS:
        ws = wb[f"{cfg}_BoM"]
        headers = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers)}
        items = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[idx["K코드"]]:
                continue
            items.append({
                "K코드": row[idx["K코드"]], "품명": row[idx["품명"]],
                "추정수량": row[idx["1개 링크 기준 추정 수량(대표값)"]],
                "확정수준": row[idx["확정 수준"]],
            })
        result[cfg] = items
    return result


def main():
    template = load_template()
    contract = load_contract_213()
    estimated = load_estimated_bom()

    template_codes = set(template.keys())
    contract_codes = set(contract.keys())
    only_in_template = template_codes - contract_codes
    only_in_contract = contract_codes - template_codes
    common_codes = template_codes & contract_codes

    # 순번 기준으로 짝지어 '동일 위치, 다른 코드' 쌍 찾기(구코드->신코드 재부여 패턴)
    seq_to_template_code = {v["순번"]: k for k, v in template.items() if k in only_in_template}
    seq_to_contract_code = {v["순번"]: k for k, v in contract.items() if k in only_in_contract}
    reassigned_pairs = []
    for seq, old_code in seq_to_template_code.items():
        new_code = seq_to_contract_code.get(seq)
        if new_code:
            reassigned_pairs.append((seq, old_code, template[old_code]["품명"], new_code, contract[new_code]["품명"]))

    n_filled_2 = sum(1 for v in template.values() if v["2+0"] not in (None, ""))
    n_filled_4 = sum(1 for v in template.values() if v["4+0"] not in (None, ""))
    n_filled_6 = sum(1 for v in template.values() if v["6+0"] not in (None, ""))
    n_filled_8 = sum(1 for v in template.values() if v["8+0"] not in (None, ""))
    n_filled_div = sum(1 for v in template.values() if v["Diversity"] not in (None, ""))
    n_filled_freq = sum(1 for v in template.values() if v["주파수"] not in (None, ""))
    reply_is_empty = (n_filled_2 + n_filled_4 + n_filled_6 + n_filled_8) == 0

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 회신현황_요약 ----------------
    ws1 = wb.create_sheet("회신현황_요약")
    ws1.append(["항목", "값", "판정"])
    summary = [
        ("업체작성양식 K코드 총수", len(template_codes), ""),
        ("계약품목(213종) K코드 총수", len(contract_codes), ""),
        ("코드 일치(공통)", len(common_codes), "확정"),
        ("업체양식에만 있는 코드(구코드 추정)", len(only_in_template), "확인 필요" if only_in_template else "-"),
        ("계약품목에만 있는 코드(신코드 추정)", len(only_in_contract), "확인 필요" if only_in_contract else "-"),
        ("주파수 열 기재됨", f"{n_filled_freq}/213 (KT 사전기재분)", "확정"),
        ("2+0 수량 기재됨", f"{n_filled_2}/213", "미회신" if n_filled_2 == 0 else "부분회신"),
        ("4+0 수량 기재됨", f"{n_filled_4}/213", "미회신" if n_filled_4 == 0 else "부분회신"),
        ("6+0 수량 기재됨", f"{n_filled_6}/213", "미회신" if n_filled_6 == 0 else "부분회신"),
        ("8+0 수량 기재됨", f"{n_filled_8}/213", "미회신" if n_filled_8 == 0 else "부분회신"),
        ("Diversity/비고 기재됨", f"{n_filled_div}/213", "미회신" if n_filled_div == 0 else "부분회신"),
        ("종합 판정", "공급사 회신 전(빈 템플릿)" if reply_is_empty else "공급사 회신 일부 반영됨",
         "확인 필요: 공급사 회신 요청" if reply_is_empty else ""),
    ]
    for row in summary:
        ws1.append(list(row))
        if row[2] in ("확인 필요", "미회신") or "확인 필요" in str(row[2]):
            mark_fill(ws1, ws1.max_row, 3, REVIEW_FILL)
    finalize_sheet(ws1, 1, 3, ws1.max_row)

    # ---------------- K코드_일치검증 ----------------
    ws2 = wb.create_sheet("K코드_일치검증")
    headers2 = ["분류", "순번(추정)", "업체양식 K코드", "업체양식 품명", "계약품목 K코드", "계약품목 품명", "판정"]
    ws2.append(headers2)
    for seq, old_code, old_name, new_code, new_name in reassigned_pairs:
        ws2.append(["코드 재부여 의심(순번 동일, 품명 동일)", seq, old_code, old_name, new_code, new_name,
                    "확인 필요(공급사에 정정 요청)"])
        mark_fill(ws2, ws2.max_row, 7, ERROR_FILL)
    matched_seqs = {p[0] for p in reassigned_pairs}
    for code in only_in_template:
        if template[code]["순번"] not in matched_seqs:
            ws2.append(["업체양식에만 존재(대응 계약코드 못찾음)", template[code]["순번"], code,
                        template[code]["품명"], "", "", "확인 필요"])
            mark_fill(ws2, ws2.max_row, 7, ERROR_FILL)
    for code in only_in_contract:
        if contract[code]["순번"] not in matched_seqs:
            ws2.append(["계약품목에만 존재(대응 업체양식코드 못찾음)", contract[code]["순번"], "", "",
                        code, contract[code]["품명"], "확인 필요"])
            mark_fill(ws2, ws2.max_row, 7, ERROR_FILL)
    ws2.append(["-", "-", "-", f"(코드 일치 {len(common_codes)}건은 생략, 요약 시트 참조)", "-", "-", "확정"])
    finalize_sheet(ws2, 1, len(headers2), ws2.max_row)

    # ---------------- 구성별_비교 ----------------
    ws3 = wb.create_sheet("구성별_비교")
    headers3 = ["구성방식", "K코드", "품명", "발주이력 기반 추정수량", "추정 확정수준",
                "업체양식 기재수량", "업체양식 Diversity/비고", "비교분류"]
    ws3.append(headers3)
    category_counts = {}
    for cfg in CONFIGS:
        for item in estimated[cfg]:
            k = item["K코드"]
            tpl = template.get(k)
            if tpl is None:
                cls = "코드 불일치(업체양식에 해당 코드 없음)"
                tpl_qty, tpl_note = "", ""
            else:
                tpl_qty = tpl[cfg]
                tpl_note = tpl["Diversity"] or tpl["비고"] or ""
                if tpl_qty in (None, ""):
                    cls = "발주실적은 있으나 공급사 BoM 미기재(회신 대기)"
                else:
                    try:
                        match = float(tpl_qty) == float(item["추정수량"])
                    except (TypeError, ValueError):
                        match = False
                    cls = "공급사 회신과 발주패턴 일치" if match else "수량 불일치"
            category_counts[cls] = category_counts.get(cls, 0) + 1
            ws3.append([cfg, k, item["품명"], item["추정수량"], item["확정수준"], tpl_qty, tpl_note, cls])
            if "불일치" in cls:
                mark_fill(ws3, ws3.max_row, 8, ERROR_FILL)
            elif "회신 대기" in cls:
                mark_fill(ws3, ws3.max_row, 8, REVIEW_FILL)
            elif "일치" in cls:
                mark_fill(ws3, ws3.max_row, 8, CONFIRMED_FILL)
    finalize_sheet(ws3, 1, len(headers3), ws3.max_row)

    # ---------------- 확인필요사항 ----------------
    ws4 = wb.create_sheet("확인필요사항")
    ws4.append(["번호", "요청 내용", "근거"])
    requests = [
        "표준BoM 업체작성양식(⑦ 파일)의 2+0/4+0/6+0/8+0/Diversity/비고 열이 전부 공란입니다. "
        "1개 링크 기준 수량을 기재하여 회신 요청.",
        f"업체양식과 계약품목(213종) 간 K코드가 다른 품목 {len(reassigned_pairs)}건(동일 순번·동일 품명, "
        "코드만 상이) 확인됨 - 계약 갱신 시 K코드가 재부여된 것으로 추정되므로, 업체양식을 최신 K코드로 "
        "정정하여 재제출 요청.",
        "6+0 구성에 대한 근거(계약서 명시 또는 발주이력)가 전혀 없습니다. 6+0 구성이 실제로 존재하는지, "
        "존재한다면 어떤 K코드로 구성되는지 확인 요청.",
        "4+0/8+0 구성의 WBX 관련 품목(K9210610 HAX_WTM 4500XT 4GHz 등)에 대한 1개 링크 기준 정확한 "
        "수량 확인 요청(발주이력상 관측치는 여러 현장 물량이 혼재되어 있을 수 있음).",
        "SD(Space Diversity) 옵션 적용 시 구성별로 추가/변경되는 품목과 수량 차이를 명시적으로 기재 "
        "요청(현재 SD OPTION 품목 자체는 확인되나 기본 구성 대비 차이는 불명).",
        "2+0(CTR 8312 1RU Chassis) 구성의 발주이력이 거의 없어 실제 1개 링크 BoM을 발주이력만으로 "
        "추정하기 어려움 - 표준 BoM 회신 시 우선적으로 확인 요청.",
    ]
    for i, req in enumerate(requests, start=1):
        ws4.append([i, req, ""])
    finalize_sheet(ws4, 1, 3, ws4.max_row)
    ws4.column_dimensions["B"].width = 100

    out_path = OUTPUT_DIR / "05_Aviat_공급사회신_검증.xlsx"
    wb.save(out_path)

    print("=== 6단계: 공급사 회신 양식 비교 검증 완료 ===")
    print(f"입력 파일: {SRC_TEMPLATE.name}, {SRC_ITEMS.name}, {SRC_BOM.name}")
    print(f"처리: 업체양식 {len(template_codes)}건, 계약품목 {len(contract_codes)}건 비교")
    print(f"생성 파일: {out_path}")
    print("주요 검증 결과:")
    print(f"  - K코드 완전 일치: {len(common_codes)}건")
    print(f"  - 코드 재부여 의심(구코드↔신코드): {len(reassigned_pairs)}건")
    print(f"  - 2+0/4+0/6+0/8+0 수량 기재: {n_filled_2}/{n_filled_4}/{n_filled_6}/{n_filled_8} (전부 0=미회신)")
    print(f"  - 구성별 비교 결과 분류: {category_counts}")
    print("  - 결론: 공급사 표준BoM 회신 대기 중. 수량 비교는 회신 후 재실행 필요.")


if __name__ == "__main__":
    main()
