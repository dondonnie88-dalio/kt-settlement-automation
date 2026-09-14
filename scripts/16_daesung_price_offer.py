"""
16단계: 대성인포텍(Ceragon) 매입단가 제시 (돈현님 요청, 2026-09-14, 10차)

돈현님 피드백(2026-09-14): "할인 더 해달라고 떼쓰는 게 아니라 단가를 협상하는 입장",
"할인율 낮다고 궁시렁 대는 게 아니라 그냥 매입단가를 제시하는 것" - 이전 버전(할인율
재검토_요청)의 "확인 부탁드립니다/재검토 가능할까요" 식 문의 톤을 전면 폐기하고, KT가
근거를 갖춘 매입단가를 선언적으로 제시하는 톤으로 재작성.

재검토 과정에서 이전 버전의 사실관계 오류도 발견해 수정함:
- 이전 버전은 안테나/마운트류 17개 품목을 "할인율이 통상 2%대인데 낮게 적용됨"으로
  플래그했으나, 실제로는 대성 단가표의 동일 안테나 계열 41개 품목 전부가 정확히
  2.00% 할인율로 일관되어 있음 - 즉 이 17개는 대성 내부 기준으로 전혀 이상치가
  아니다. (15번의 '정가상한 적용' 42건 리스트에 걸린 이유는 대성 할인율이 낮아서가
  아니라, KT 자체 마진율 체계에서 안테나 카테고리를 3%로 잡았기 때문 - 이는 대성에
  요청할 사안이 아니라 KT 내부 마진 결정 사안이다.) 이번 버전은 이 17개를 대성向
  요청에서 제외하고, 실제로 대성 자체 데이터에서 이상치로 확인되는 5개 품목만 다룬다:
  - SFP 3종(K9197304/306/309): 할인율 1.00% - 동일 SFP 계열 16개 중 13개(81%)에
    적용된 표준 할인율 14.69%(=판매가/정가 0.8531)와 명백히 어긋남.
  - RFUC-CPLR-8(K9178851), RFUC-TWIST Kit-6(K9178856): 할인 후 단가가 정가보다
    높음(할인율 음수) + 두 품목 모두 정확히 634,418원으로 동일 - 표기 오류로 판단.
    동일 정가를 가진 형제 품목(CPLR-6, TWIST Kit-8/11)의 실제 단가가 존재하므로
    이를 근거로 정확한 대체 단가를 제시할 수 있음(표준 할인율 0.8531 적용값과도
    사실상 일치 - 이중으로 검증됨).
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import (
    finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, FMT_AMOUNT, FMT_PERCENT,
    mark_fill, WRAP_TOP,
)

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

# 대성 단가표 474개 유효품목(신구가 모두 기재) 중 395개(83%)에 적용된 표준 할인율.
# 판매가/정가 = 0.8531 (여러 계열에서 소수점 4자리까지 정확히 일치 확인됨)
STD_RATIO = 0.8531

# (K코드, 구분) - "표준율미적용": SFP 3종, 정상 데이터지만 표준 할인율 대신 1.00%만
# 적용됨. "표기오류": 할인 후 단가가 정가보다 높은 2건, 형제 품목 값으로 대체 제시.
TARGET_ITEMS = [
    ("K9197304", "표준율미적용"),
    ("K9197306", "표준율미적용"),
    ("K9197309", "표준율미적용"),
    ("K9178851", "표기오류"),
    ("K9178856", "표기오류"),
]

# 표기오류 2건의 대체 단가 근거: 동일 정가를 가진 형제 품목의 실제 적용 단가
SIBLING_REF = {
    "K9178851": ("K9178850", "CER_RFUC-CPLR-6"),       # 정가 548,550원 동일
    "K9178856": ("K9178857", "CER_RFUC-TWIST Kit-8"),  # 정가 123,210원 동일
}


def load_master():
    wb = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    ws = wb["품목마스터"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_code = {row[idx["K코드"]]: row for row in ws.iter_rows(min_row=2, values_only=True)}
    return by_code, idx


def load_items():
    by_code, idx = load_master()
    rows = []
    for kcode, kind in TARGET_ITEMS:
        row = by_code[kcode]
        list_price, cur_price = row[idx["기존단가"]], row[idx["최종인하단가"]]
        cur_rate = (list_price - cur_price) / list_price if list_price else None

        if kind == "표준율미적용":
            offer_price = round(list_price * STD_RATIO)
            basis = f"동일 SFP 계열 16종 중 13종(81%)에 적용된 표준 할인율 14.69% 기준 매입단가 제시"
        else:
            sib_code, sib_name = SIBLING_REF[kcode]
            sib_row = by_code[sib_code]
            offer_price = sib_row[idx["최종인하단가"]]
            basis = (f"정가가 동일한 형제 품목 {sib_name}({sib_code})의 실제 적용 단가를 "
                     f"기준으로 제시 (표준 할인율 0.8531 적용값과도 일치)")

        rows.append({
            "K코드": kcode, "품명": row[idx["품명"]], "정가": list_price,
            "대성_현재제시단가": cur_price, "대성_현재할인율": cur_rate,
            "KT_제시매입단가": offer_price, "차액": cur_price - offer_price,
            "구분": kind, "근거": basis,
        })
    return rows


def build_message_sheet(wb, rows):
    ws = wb.create_sheet("매입단가_제시(초안)")
    ws.append(["항목", "내용"])

    sfp = [r for r in rows if r["구분"] == "표준율미적용"]
    err = [r for r in rows if r["구분"] == "표기오류"]

    def line(r):
        return f"  - {r['품명'].strip()}({r['K코드']}): {r['KT_제시매입단가']:,.0f}원 (정가 {r['정가']:,}원)"

    msg = (
        "안녕하세요, KT commerce 이돈현입니다.\n\n"
        "귀사 8GHz/11GHz 구성 단가표를 검토한 결과, 아래 5개 품목은 매입단가를 다음과 "
        "같이 적용하여 진행하고자 합니다.\n\n"
        "1) SFP 3종\n"
        "현재 제시된 단가는 할인율 1.0% 수준이나, 귀사 단가표 내 동일 SFP 계열 16개 "
        "품목 중 13개(81%)에는 표준 할인율 14.69%가 적용되어 있습니다. 이번 3개 "
        "품목에도 동일한 표준 할인율을 적용한 아래 단가로 매입하겠습니다.\n"
        + "\n".join(line(r) for r in sfp) + "\n\n"
        "2) RFUC-CPLR-8, RFUC-TWIST Kit-6\n"
        "두 품목 모두 할인 후 단가가 정가보다 높게(634,418원, 두 품목 동일값) 기재되어 "
        "있어 표기 오류로 판단됩니다. 동일 정가를 가진 형제 품목의 실제 적용 단가를 "
        "기준으로 아래와 같이 매입단가를 적용하겠습니다.\n"
        + "\n".join(line(r) for r in err) + "\n\n"
        "이견 있으시면 회신 부탁드리며, 별도 회신 없을 시 위 단가로 진행하겠습니다.\n\n"
        "감사합니다."
    )

    guide = [
        ("용도", "대성인포텍에 보낼 이메일/메신저 본문 초안. '제시단가_상세내역' 시트를 "
                "첨부하거나 표를 붙여넣으면 된다."),
        ("메시지 초안", msg),
        ("참고(내부용, 대성에 보내지 않음)", "당초 검토 대상은 21개(안테나/마운트류 17개 "
                "포함)였으나, 그 17개는 대성 단가표 내 동일 안테나 계열 41개 품목 전부와 "
                "정확히 같은 2.00% 할인율로 이미 일관되어 있음을 확인해 제외함. 15번 "
                "파일의 '정가상한 적용 42건'에 이 17개가 걸린 이유는 대성 할인율이 낮아서가 "
                "아니라 KT 자체 마진율(안테나 3%)이 대성 할인율(2%)보다 높기 때문 - 이는 "
                "대성에 요청할 사안이 아니라 KT 내부 마진 결정의 문제."),
        ("주의", "이 초안은 자동 생성된 것이니 보내시기 전에 실제 상황(이미 통화/메일로 "
              "언급한 내용, 호칭, 마감 일정 등)에 맞게 다듬어서 사용하시기 바랍니다."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=2).alignment = WRAP_TOP


def build_list_sheet(wb, rows):
    ws = wb.create_sheet("제시단가_상세내역")
    headers = ["구분", "K코드", "품명", "정가", "대성 현재 제시단가", "대성 현재 할인율",
               "KT 제시 매입단가", "차액(현재-제시)", "근거"]
    ws.append(headers)
    for r in rows:
        rr = ws.max_row + 1
        ws.append([r["구분"], r["K코드"], r["품명"], r["정가"], r["대성_현재제시단가"],
                   r["대성_현재할인율"], r["KT_제시매입단가"], r["차액"], r["근거"]])
        if r["구분"] == "표기오류":
            mark_fill(ws, rr, 1, ERROR_FILL)
            mark_fill(ws, rr, 6, ERROR_FILL)
            mark_fill(ws, rr, 7, CONFIRMED_FILL)
        else:
            mark_fill(ws, rr, 1, REVIEW_FILL)
            mark_fill(ws, rr, 7, REVIEW_FILL)
    for c in [4, 5, 7, 8]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=6).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["I"].width = 55
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=9).alignment = WRAP_TOP


def main():
    rows = load_items()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_message_sheet(wb, rows)
    build_list_sheet(wb, rows)
    wb._sheets = [wb["매입단가_제시(초안)"], wb["제시단가_상세내역"]]

    out_path = OUTPUT_DIR / "16_대성_매입단가_제시.xlsx"
    wb.save(out_path)

    n_sfp = sum(1 for r in rows if r["구분"] == "표준율미적용")
    n_err = sum(1 for r in rows if r["구분"] == "표기오류")
    total_down = sum(r["차액"] for r in rows)
    print("=== 16단계: 대성인포텍 매입단가 제시 목록 생성 완료 ===")
    print(f"표준 할인율 미적용(SFP): {n_sfp}건, 표기오류 추정(형제품목 대체): {n_err}건")
    print(f"현재 제시가 대비 인하 총액: {total_down:,.0f}원")
    print(f"생성 파일: {out_path}")


if __name__ == "__main__":
    main()
