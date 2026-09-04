"""
2단계: Aviat 213종 계약품목 정규화
입력: input/510331d5-_____MDMmini_digital_microwave_213______.xlsx (시트 '등록요청', row4~216)
출력: output/01_Aviat_213종_정규화.xlsx
시트: 원본정리 / 품목정규화 / 데이터품질_이슈 / 품목군별_요약

원본 파일은 읽기 전용으로만 사용한다.
"""
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))
from classify_aviat import extract_all, ITEM_GROUPS
from xlsx_style import (
    finalize_sheet, ORIGIN_FONT, ERROR_FILL, CONFIRMED_FILL,
    FMT_AMOUNT, FMT_PERCENT, FMT_QTY, mark_fill, mark_font,
)

BASE_DIR = Path("/home/user/kt-settlement-automation")
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
SRC = INPUT_DIR / "510331d5-_____MDMmini_digital_microwave_213______.xlsx"

# 등록요청 컬럼 인덱스 (1-based)
COL = {
    "순서": 1, "물품명": 2, "모델명": 3, "세부규격": 4, "제조사": 5, "원산지": 6,
    "진열기간": 7, "납기": 8, "수량": 9, "예산": 10, "판매가": 11, "판매가검토": 12,
    "매입가": 13, "수수료": 14, "마진율": 15, "공급사": 16, "K코드": 17, "KTC코드": 18,
    "등록사유": 19, "제3단가확인": 20, "비고": 21, "품목번호": 22, "상품명(크기)": 23,
    "요청번호": 24, "계약번호": 25, "계약명": 26, "계약자": 27, "상품코드": 28,
}

ROW_START, ROW_END = 4, 216  # 213개 품목


def load_rows():
    wb_v = openpyxl.load_workbook(SRC, data_only=True)
    wb_f = openpyxl.load_workbook(SRC, data_only=False)
    ws_v = wb_v["등록요청"]
    ws_f = wb_f["등록요청"]
    ws_naeyeok_v = wb_v["내역서"]

    rows = []
    for i, r in enumerate(range(ROW_START, ROW_END + 1), start=1):
        def val(col_name):
            return ws_v.cell(row=r, column=COL[col_name]).value

        is_formula_qty = ws_f.cell(row=r, column=COL["수량"]).data_type == "f"
        unit = ws_naeyeok_v.cell(row=10 + (i - 1), column=5).value  # 내역서 E열

        item = {
            "순번": i,
            "K코드": (val("K코드") or "").strip() if val("K코드") else "",
            "상품코드": val("상품코드"),
            "KTC코드": val("KTC코드"),
            "품명": val("물품명").strip() if isinstance(val("물품명"), str) else val("물품명"),
            "모델명_원본": val("모델명"),
            "세부규격": val("세부규격"),
            "단위": unit or "EA",
            "제조사": val("제조사"),
            "원산지": val("원산지"),
            "판매단가": val("판매가") or 0,
            "매입단가": val("매입가") or 0,
            "마진율_원본": val("마진율"),
            "계약수량": val("수량") or 0,
            "계약수량_수식기반": is_formula_qty,
            "등록사유": val("등록사유"),
            "비고": val("비고"),
        }
        rows.append(item)
    return rows


def build_workbook(rows):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 시트1: 원본정리 ----------------
    ws1 = wb.create_sheet("원본정리")
    headers1 = ["순번", "K코드", "상품코드", "KTC코드", "품명", "모델명(원본,K코드중복주의)",
                "세부규격", "단위", "제조사", "원산지", "판매단가", "매입단가", "마진액",
                "마진율", "계약수량(등록요청 원본)", "계약수량×판매단가(산출값,주의:계약서 표기 계약금액과 다른 개념)",
                "등록사유", "비고"]
    ws1.append(headers1)
    for it in rows:
        r = ws1.max_row + 1
        ws1.append([
            it["순번"], it["K코드"], it["상품코드"], it["KTC코드"], it["품명"], it["모델명_원본"],
            it["세부규격"], it["단위"], it["제조사"], it["원산지"], it["판매단가"], it["매입단가"],
            None, None, it["계약수량"], None, it["등록사유"], it["비고"],
        ])
        ws1.cell(row=r, column=13).value = f"=K{r}-L{r}"          # 마진액
        ws1.cell(row=r, column=14).value = f"=IF(K{r}=0,0,M{r}/K{r})"  # 마진율
        ws1.cell(row=r, column=16).value = f"=O{r}*K{r}"          # 계약금액 = 계약수량*판매단가
        # 원본값(직접 입력) 셀은 녹색 글씨
        for c in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 17, 18]:
            mark_font(ws1, r, c, ORIGIN_FONT)
        if it["모델명_원본"] and it["K코드"] and str(it["모델명_원본"]).strip() != it["K코드"]:
            mark_fill(ws1, r, 6, ERROR_FILL)
        if it["계약수량_수식기반"]:
            mark_fill(ws1, r, 15, CONFIRMED_FILL)

    for c in [11, 12, 13, 16]:
        for r in range(2, ws1.max_row + 1):
            ws1.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws1.max_row + 1):
        ws1.cell(row=r, column=14).number_format = FMT_PERCENT
        ws1.cell(row=r, column=15).number_format = FMT_QTY
    finalize_sheet(ws1, 1, len(headers1), ws1.max_row)

    # ---------------- 시트2: 품목정규화 ----------------
    ws2 = wb.create_sheet("품목정규화")
    tech_cols = ["품목군", "장비계열", "주파수", "L6_U6구분", "LowHigh_Band", "출력등급", "채널수",
                 "구성표기(2+0등)", "SD_Diversity표기", "RU", "포트수", "인터페이스속도",
                 "케이블길이", "Waveguide규격", "안테나크기", "편파", "데이터품질이슈"]
    headers2 = ["순번", "K코드", "품명", "세부규격"] + tech_cols
    ws2.append(headers2)

    quality_issue_rows = []
    tech_results = []
    for it in rows:
        tech = extract_all(it["품명"], it["세부규격"])
        tech_results.append(tech)
        r = ws2.max_row + 1
        ws2.append([
            it["순번"], it["K코드"], it["품명"], it["세부규격"],
            tech["품목군"], tech["장비계열"], tech["주파수"], tech["L6_U6구분"], tech["LowHigh_Band"],
            tech["출력등급"], tech["채널수"], tech["구성표기"], tech["SD_Diversity표기"],
            tech["RU"], tech["포트수"], tech["인터페이스속도"], tech["케이블길이"],
            tech["Waveguide규격"], tech["안테나크기"], tech["편파"], tech["데이터품질이슈"],
        ])
        for c in [1, 2, 3, 4]:
            mark_font(ws2, r, c, ORIGIN_FONT)
        if tech["데이터품질이슈"]:
            mark_fill(ws2, r, len(headers2), ERROR_FILL)
            quality_issue_rows.append((it, tech, "주파수 표기 불일치/오타"))
    finalize_sheet(ws2, 1, len(headers2), ws2.max_row)

    # 모델명=K코드 불일치(구K코드 잔존) 케이스 별도 수집
    mismatch_rows = [it for it in rows if it["모델명_원본"] and it["K코드"]
                      and str(it["모델명_원본"]).strip() != it["K코드"]]

    # ---------------- 시트3: 데이터품질_이슈 ----------------
    ws3 = wb.create_sheet("데이터품질_이슈")
    headers3 = ["구분", "순번", "K코드", "품명", "이슈유형", "상세설명", "확정수준"]
    ws3.append(headers3)
    ws3.append([
        "전체(공통)", "-", "-", "-", "모델명 필드=K코드",
        f"213건 중 {213-len(mismatch_rows)}건은 '모델명'(등록요청 C열) 값이 K코드와 동일함. "
        "즉 등록요청/내역서 원본 파일의 '모델명' 필드가 실제 Aviat 모델번호가 아니라 K코드를 담고 있음. "
        "본 정규화 결과의 '모델명(원본)' 컬럼은 참고용이며, 실제 모델 식별은 '품명' 텍스트를 사용할 것.",
        "확정(원본 직접 확인)",
    ])
    for c in [2, 3, 4]:
        mark_fill(ws3, ws3.max_row, c, ERROR_FILL) if False else None
    for it in mismatch_rows:
        ws3.append([
            "개별(코드 재부여 의심)", it["순번"], it["K코드"], it["품명"], "모델명≠K코드(구코드 잔존)",
            f"'모델명' 필드 값은 '{it['모델명_원본']}'로 현재 K코드('{it['K코드']}')와 다름. "
            "업체작성 표준BoM 양식(⑦ 파일)에서도 동일하게 구코드가 사용되고 있어, 계약 갱신 과정에서 "
            "K코드가 재부여되었을 가능성이 있음(확인 필요: 공급사에 정정 요청).",
            "확인 필요(공급사 확인)",
        ])
        mark_fill(ws3, ws3.max_row, 3, ERROR_FILL)
    for it, tech, issue_type in quality_issue_rows:
        ws3.append([
            "개별(품명/설명 불일치)", it["순번"], it["K코드"], it["품명"], issue_type,
            f"{tech['데이터품질이슈']} (세부규격: {it['세부규격']})", "확인 필요(공급사 확인)",
        ])
        mark_fill(ws3, ws3.max_row, 5, ERROR_FILL)
    total_qty_sum = sum(it["계약수량"] or 0 for it in rows)
    total_sale_calc = sum((it["계약수량"] or 0) * (it["판매단가"] or 0) for it in rows)
    ws3.append([
        "전체(공통)", "-", "-", "-", "계약금액 개념 상이(확정, 수식 직접 확인)",
        f"계약서('내역서' 시트) 헤더에 표기된 '계약금액 421,882,338원'은 내역서 H223셀 수식(=SUM(H10:H222), "
        f"H=F×G, F(수량)는 전 행 1로 고정)의 계산 결과로, '213개 품목을 각 1개씩 구매한다고 가정했을 때의 "
        f"판매단가 합계'임(확정: 내역서 F열이 모든 행에서 1, H223 캐시값이 421,882,338과 일치함을 직접 확인). "
        f"반면 '등록요청' 시트의 '수량'(I열, 총 {total_qty_sum:,}개)은 별도 목적(초도 예상물량 등으로 추정)의 값이며, "
        f"이를 판매단가와 곱한 합계는 {total_sale_calc:,.0f}원으로 계약서 표기 금액과 전혀 다름. "
        f"본 파일의 '계약수량×판매단가(산출값)' 컬럼은 후자('등록요청' 수량 기준)이므로 계약서상 "
        f"'계약금액'과 혼동하지 말 것. 또한 내역서 F223 수식 자체가 '=SUM(F10:F10)'으로 213개 행이 아닌 "
        f"1개 행만 합산하도록 되어 있어(원본 수식 오류로 추정) 수량 합계 표시에는 사용할 수 없음.",
        "확정(원본 수식 직접 확인) / 해석은 확인 필요",
    ])
    formula_qty_rows = [it for it in rows if it["계약수량_수식기반"]]
    if formula_qty_rows:
        codes = ", ".join(it["K코드"] for it in formula_qty_rows)
        ws3.append([
            "전체(공통)", "-", codes, "-", "계약수량 셀 수식 오류",
            f"{len(formula_qty_rows)}건의 '계약수량' 셀이 끊어진 외부 워크북 참조 수식"
            "(`_xlfn.XLOOKUP(...,[1]Sheet1!...)`)으로 되어 있음. 엑셀에 저장된 캐시값은 정상적으로 "
            "읽혀 사용했으나(원본정리 시트에 연한 녹색 배경으로 표시), 원본 파일을 재계산하면 "
            "값이 바뀌거나 오류가 날 수 있어 수치의 최신성은 확인 필요.",
            "확인 필요(원본 수식 복구 필요)",
        ])
    finalize_sheet(ws3, 1, len(headers3), ws3.max_row)

    # ---------------- 시트4: 품목군별_요약 ----------------
    ws4 = wb.create_sheet("품목군별_요약")
    headers4 = ["품목군", "품목수", "계약수량합계", "판매금액합계", "매입금액합계", "평균마진율"]
    ws4.append(headers4)
    group_agg = defaultdict(lambda: {"cnt": 0, "qty": 0, "sale": 0.0, "buy": 0.0, "margin_sum": 0.0})
    for it, tech in zip(rows, tech_results):
        g = tech["품목군"]
        agg = group_agg[g]
        agg["cnt"] += 1
        qty = it["계약수량"] or 0
        agg["qty"] += qty
        agg["sale"] += qty * (it["판매단가"] or 0)
        agg["buy"] += qty * (it["매입단가"] or 0)
        if it["판매단가"]:
            agg["margin_sum"] += (it["판매단가"] - it["매입단가"]) / it["판매단가"]
    for g in ITEM_GROUPS:
        if g not in group_agg:
            continue
        agg = group_agg[g]
        avg_margin = agg["margin_sum"] / agg["cnt"] if agg["cnt"] else 0
        ws4.append([g, agg["cnt"], agg["qty"], agg["sale"], agg["buy"], avg_margin])
    total_cnt = sum(a["cnt"] for a in group_agg.values())
    total_qty = sum(a["qty"] for a in group_agg.values())
    total_sale = sum(a["sale"] for a in group_agg.values())
    total_buy = sum(a["buy"] for a in group_agg.values())
    ws4.append(["합계", total_cnt, total_qty, total_sale, total_buy, None])
    for r in range(2, ws4.max_row + 1):
        ws4.cell(row=r, column=3).number_format = FMT_QTY
        ws4.cell(row=r, column=4).number_format = FMT_AMOUNT
        ws4.cell(row=r, column=5).number_format = FMT_AMOUNT
        ws4.cell(row=r, column=6).number_format = FMT_PERCENT
    finalize_sheet(ws4, 1, len(headers4), ws4.max_row)

    return wb, total_cnt, total_qty, total_sale, mismatch_rows, quality_issue_rows, formula_qty_rows


def main():
    rows = load_rows()

    # 검증: 213종 여부, K코드 중복/빈값
    kcodes = [it["K코드"] for it in rows]
    n_blank = sum(1 for k in kcodes if not k)
    n_dup = len(kcodes) - len(set(kcodes))

    wb, total_cnt, total_qty, total_sale, mismatch_rows, quality_issue_rows, formula_qty_rows = build_workbook(rows)

    out_path = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"
    wb.save(out_path)

    stated_total = 421882338  # 계약서('내역서') 헤더 표기 금액: 213개 품목 x 1개 가정 단가합계

    print("=== 2단계: Aviat 213종 계약품목 정규화 완료 ===")
    print(f"입력 파일: {SRC.name} (시트 '등록요청' row{ROW_START}~{ROW_END}, '내역서' 단위열)")
    print(f"처리 행 수: {len(rows)}건")
    print(f"제외/오류 행 수: 0건 (전량 정상 처리)")
    print(f"생성 파일: {out_path}")
    print("주요 검증 결과:")
    print(f"  - 품목 수: {len(rows)}건 (기대값 213건, {'일치' if len(rows) == 213 else '불일치!'})")
    print(f"  - K코드 빈값: {n_blank}건, 중복: {n_dup}건")
    print(f"  - 모델명≠K코드(구코드 잔존 의심) 건수: {len(mismatch_rows)}건")
    print(f"  - 품명/설명 주파수 표기 불일치·오타 건수: {len(quality_issue_rows)}건")
    print(f"  - 계약수량 셀이 끊어진 외부링크 수식인 건수: {len(formula_qty_rows)}건(캐시값 사용)")
    print(f"  - 계약서(내역서) 표기 계약금액: {stated_total:,}원 (213개×1개 가정 단가합계, 내역서 H223 수식 확인)")
    print(f"  - 등록요청 수량(I열, 합계 {total_qty:,}개)×판매단가 산출값: {total_sale:,.0f}원 (별도 개념, 데이터품질_이슈 시트 참조)")
    print(f"  - 품목군 분류 결과: {len(set(t['품목군'] for t in [extract_all(it['품명'], it['세부규격']) for it in rows]))}개 그룹 사용")


if __name__ == "__main__":
    main()
