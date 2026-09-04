"""
11단계(추가): 애니콤(Aviat) 회신 대조 템플릿 생성
입력: output/04_Aviat_구성방식별_표준BoM_추정_v2.xlsx
출력: output/10_Aviat_공급사회신_대조표.xlsx

돈현님 요청(2026-09-04, 3차): 애니콤 회신을 기다리는 동안 회신 도착 시 바로 쓸 수 있는
대조표를 미리 만들어 둔다. 04v2의 8GHz/11GHz IAP3 4개 후보 구성(2+0/4+0 × 8/11GHz)을
그대로 옮기고, 공급사 회신 수량·단가를 입력하면 차이가 자동 계산되도록 수식을 심는다.
공급사 회신이 아직 없으므로 입력란은 전부 공란으로 남긴다(임의로 채우지 않음).
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, INPUT_FONT,
                         FMT_AMOUNT, FMT_PERCENT, mark_fill, mark_font)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_BOM_V2 = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정_v2.xlsx"

BAND_SHEETS = ["8GHz_IAP3_2+0", "8GHz_IAP3_4+0", "11GHz_IAP3_2+0", "11GHz_IAP3_4+0"]


def load_band_rows(sheet_name):
    wb = openpyxl.load_workbook(SRC_BOM_V2, data_only=True)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["K코드"]]:
            continue
        rows.append({
            "K코드": row[idx["K코드"]], "품명": row[idx["품명"]], "품목군": row[idx["품목군"]],
            "판정수준": row[idx["판정수준"]],
            "대표수량": row[idx["대표 수량(1개 링크=N ODU쌍 기준)"]],
            "근거이벤트ID": row[idx["근거 이벤트ID"]],
            "판매단가": row[idx["판매단가"]], "매입단가": row[idx["매입단가"]],
        })
    # 유력 후보 -> 확정 -> 참고 후보 순, 그 안에서는 원래 순서 유지
    order = {"확정": 0, "유력 후보": 1, "참고 후보": 2}
    rows.sort(key=lambda r: order.get(r["판정수준"], 9))
    return rows


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ================= 안내 =================
    ws0 = wb.create_sheet("안내")
    ws0.append(["항목", "내용"])
    guide = [
        ("목적", "애니콤(Aviat) 공급사 표준BoM 회신이 도착하면, 04v2에서 발주이력으로 추정한 "
                "8GHz/11GHz IAP3 2+0/4+0 후보와 바로 대조하기 위한 템플릿이다. 회신 전에는 "
                "'공급사 회신 수량'과 '공급사 회신단가' 열이 전부 공란이며, 임의로 채우지 않는다."),
        ("사용 방법", "1) 애니콤 회신에서 각 K코드(또는 품명이 동일한 품목)를 찾아 '공급사 회신 "
                   "수량(파란 글씨 입력란)'에 입력한다. 2) 단가도 회신에 있으면 같은 방식으로 "
                   "입력한다. 3) '수량 차이'·'일치 여부'·'단가 차이(%)' 열은 자동 계산된다."),
        ("K코드가 다를 경우", "애니콤 회신의 K코드가 여기 목록과 다르면(구코드/신코드 차이 등) "
                          "품명을 기준으로 대조하고, '비고'에 실제 매칭한 애니콤 코드를 적어둔다."),
        ("회신에만 있는 품목", "이 템플릿에 없는 품목이 애니콤 회신에 있으면, 각 시트 맨 아래에 새 "
                          "행을 추가하고 '판정수준'에 '회신에만 있음'이라고 적는다 — 발주이력으로는 "
                          "찾지 못했던 품목일 수 있으므로 특히 중요하게 봐야 한다."),
        ("판정수준 해석", "확정=계약서 텍스트에 구성 표기가 직접 있는 품목(멤버십만 확정, 수량은 "
                       "이번에 처음 확인). 유력 후보=독립된 2개 이상 이벤트에서 반복 관측(11GHz "
                       "2+0의 6개 품목). 참고 후보=단일 이벤트 관측뿐이거나 다중대역이 섞인 관측"
                       "(8GHz 2+0/4+0, 11GHz 4+0 전체)."),
        ("대표수량 '변동, 공급사 확인 필요'", "발주이력상 관측 비율이 30% 이상 벌어져 대표 수량을 "
                                        "임의로 정하지 않은 품목이다. 이런 행은 수량 차이를 계산하지 "
                                        "않고 '내부 추정 변동(비교 보류)'로 표시한다 - 애니콤 회신 "
                                        "수량을 그대로 신뢰하고 참고용으로만 기록한다."),
        ("근거 파일", "각 행의 세부 근거(관측 이벤트ID·주문번호·관측 비율 목록)는 "
                   "`04_Aviat_구성방식별_표준BoM_추정_v2.xlsx`의 해당 시트에서 K코드로 찾아보면 된다."),
    ]
    for k, v in guide:
        ws0.append([k, v])
    finalize_sheet(ws0, 1, 2, ws0.max_row)
    ws0.column_dimensions["B"].width = 100

    headers = ["K코드", "품명", "품목군", "내부 판정수준", "내부 대표수량(1개 링크 기준)",
               "근거 이벤트ID", "공급사 회신 수량", "수량 차이", "일치 여부",
               "내부 판매단가", "공급사 회신단가", "단가 차이(%)", "비고"]

    summary_rows = []
    for sheet_name in BAND_SHEETS:
        rows = load_band_rows(sheet_name)
        ws = wb.create_sheet(f"{sheet_name}_대조")
        ws.append(headers)
        for row in rows:
            r = ws.max_row + 1
            rep = row["대표수량"]
            ws.append([
                row["K코드"], row["품명"], row["품목군"], row["판정수준"], rep,
                row["근거이벤트ID"], None,
                f'=IF(OR(G{r}="",NOT(ISNUMBER(E{r}))),"",G{r}-E{r})',
                f'=IF(G{r}="","회신 대기",IF(NOT(ISNUMBER(E{r})),"내부 추정 변동(비교 보류)",'
                f'IF(G{r}=E{r},"일치","차이 있음")))',
                row["판매단가"], None,
                f'=IF(OR(K{r}="",J{r}=0),"",(K{r}-J{r})/J{r})',
                "",
            ])
            mark_font(ws, r, 7, INPUT_FONT)
            mark_font(ws, r, 11, INPUT_FONT)
            if row["판정수준"] == "유력 후보":
                mark_fill(ws, r, 4, CONFIRMED_FILL)
            elif row["판정수준"] == "참고 후보":
                mark_fill(ws, r, 4, REVIEW_FILL)
        for c in [10, 11]:
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=c).number_format = FMT_AMOUNT
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=12).number_format = FMT_PERCENT
        finalize_sheet(ws, 1, len(headers), ws.max_row)
        summary_rows.append((f"{sheet_name}_대조", len(rows)))

    # ================= 요약 =================
    ws_sum = wb.create_sheet("요약")
    headers_sum = ["구성", "내부 추정 품목수", "회신 완료", "일치", "차이 있음",
                   "회신 대기", "내부 추정 변동(비교 보류)"]
    ws_sum.append(headers_sum)
    for sheet_name, n in summary_rows:
        last_row = n + 1
        rng = f"'{sheet_name}'!I2:I{last_row}"
        ws_sum.append([
            sheet_name, n,
            f'={n}-COUNTIF({rng},"회신 대기")',
            f'=COUNTIF({rng},"일치")',
            f'=COUNTIF({rng},"차이 있음")',
            f'=COUNTIF({rng},"회신 대기")',
            f'=COUNTIF({rng},"내부 추정 변동(비교 보류)")',
        ])
    finalize_sheet(ws_sum, 1, len(headers_sum), ws_sum.max_row)

    sheet_order = ["안내"] + [f"{s}_대조" for s in BAND_SHEETS] + ["요약"]
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "10_Aviat_공급사회신_대조표.xlsx"
    wb.save(out_path)

    print("=== 11단계(추가): 애니콤 회신 대조 템플릿 생성 완료 ===")
    print(f"입력 파일: {SRC_BOM_V2.name}")
    for sheet_name, n in summary_rows:
        print(f"  - {sheet_name}: {n}행 (공급사 회신 수량/단가 입력란 공란)")
    print(f"생성 파일: {out_path}")
    print("※ 공급사 회신이 도착하면 '공급사 회신 수량'·'공급사 회신단가' 열(파란 글씨)에 값을 "
          "입력하면 차이/일치 여부가 자동 계산됩니다.")


if __name__ == "__main__":
    main()
