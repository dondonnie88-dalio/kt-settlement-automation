"""
16단계: 대성인포텍(Ceragon) 할인율 재검토 요청 목록 (돈현님 요청, 2026-09-14, 9차)

15번 파일에서 발견한 '할인율<KT 마진율' 42개 품목 중, 이번 8/11GHz 딜과 실제 관련 있어
보이는 20개(SFP 3종 + 안테나/마운트류 17종)와, 할인 후 단가가 오히려 정가보다 비싸진
데이터 이상치 2건(대역 무관하게 포함)을 정리해 대성에 보낼 요청 목록을 만든다.

근거의 성격이 이전(Aviat 비교)과 다르다는 점이 중요: 이건 경쟁사 비교가 아니라 대성
자기 자신의 가격표 안에서의 내부 일관성 문제다("경쟁사보다 비싸다"가 아니라 "당신 가격표
안에서도 이 품목들만 할인율이 유독 낮다/음수다") - 훨씬 방어하기 쉽고 대성도 받아들이기
쉽다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, FMT_AMOUNT, FMT_PERCENT, mark_fill, WRAP_TOP

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

# (K코드, 구분) - 구분: "할인율낮음"(정상 데이터, 할인율만 유독 낮음) / "데이터오류"(할인 후
# 단가가 정가보다 비쌈, 사실상 오류로 추정)
TARGET_ITEMS = [
    ("K9197304", "할인율낮음"), ("K9197306", "할인율낮음"), ("K9197309", "할인율낮음"),
    ("K9178803", "할인율낮음"), ("K9178804", "할인율낮음"),
    ("K9178807", "할인율낮음"), ("K9178808", "할인율낮음"),
    ("K9178811", "할인율낮음"), ("K9178812", "할인율낮음"),
    ("K9178815", "할인율낮음"), ("K9178816", "할인율낮음"),
    ("K9178819", "할인율낮음"), ("K9178820", "할인율낮음"),
    ("K9178823", "할인율낮음"), ("K9178824", "할인율낮음"),
    ("K9178829", "할인율낮음"), ("K9178830", "할인율낮음"),
    ("K9178876", "할인율낮음"), ("K9178877", "할인율낮음"),
    ("K9178851", "데이터오류"),  # 위 '할인율낮음' 그룹과 무관하게 자체적으로도 이상치(할인율 음수)
    ("K9178856", "데이터오류"),
]


def load_items():
    wb = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    ws = wb["품목마스터"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    by_code = {row[idx["K코드"]]: row for row in ws.iter_rows(min_row=2, values_only=True)}
    rows = []
    for kcode, kind in TARGET_ITEMS:
        row = by_code[kcode]
        기존, 인하 = row[idx["기존단가"]], row[idx["최종인하단가"]]
        rate = (기존 - 인하) / 기존 if 기존 else None
        rows.append({
            "K코드": kcode, "품명": row[idx["품명"]], "기존단가": 기존, "최종인하단가": 인하,
            "할인율": rate, "구분": kind,
        })
    return rows


def build_message_sheet(wb):
    ws = wb.create_sheet("요청메시지(초안)")
    ws.append(["항목", "내용"])
    guide = [
        ("용도", "대성인포텍에 보낼 이메일/메신저 본문 초안. 아래 표 그대로 복사해서 쓰거나 "
                "다듬어서 사용하면 된다. '확인요청_품목목록' 시트를 첨부하거나 표를 붙여넣으면 "
                "된다."),
        ("메시지 초안", "안녕하세요, KT commerce 이돈현입니다.\n\n"
                     "제출해 주신 8GHz/11GHz 구성 단가표를 검토하던 중, 첨부 목록의 품목들에서 "
                     "확인이 필요한 부분이 있어 문의드립니다.\n\n"
                     "1) 아래 19개 품목은 귀사 단가표 내 다른 품목들의 일반적인 할인율(약 2%대)"
                     "과 달리, 할인율이 1~2% 수준으로 낮게 적용되어 있습니다. 동일한 기준으로 "
                     "재검토가 가능한지 확인 부탁드립니다.\n\n"
                     "2) 아래 2개 품목(RFUC-CPLR-8, RFUC-TWIST Kit-6)은 할인 후 단가가 오히려 "
                     "정가보다 높게 기재되어 있고, 두 품목의 할인 후 단가가 정확히 동일한 값"
                     "(634,418원)입니다. 단가표 작성 과정에서 다른 행의 값이 잘못 들어간 것으로 "
                     "보이는데, 정확한 단가 확인 부탁드립니다.\n\n"
                     "감사합니다."),
        ("주의", "이 초안은 자동 생성된 것이니 보내시기 전에 실제 상황(이미 통화/메일로 언급한 "
              "내용, 호칭, 마감 일정 등)에 맞게 다듬어서 사용하시기 바랍니다."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=2).alignment = WRAP_TOP


def build_list_sheet(wb, rows):
    ws = wb.create_sheet("확인요청_품목목록")
    headers = ["구분", "K코드", "품명", "정가(기존단가)", "할인 후 단가(현재)", "현재 할인율", "비고"]
    ws.append(headers)
    for r in rows:
        note = ("할인율이 귀사 단가표 내 통상 수준(약 2%)보다 낮음 - 재검토 요청" if r["구분"] == "할인율낮음"
                else "할인 후 단가가 정가보다 높음(할인율 음수) - 데이터 오류로 추정, 확인 요청")
        rr = ws.max_row + 1
        ws.append([r["구분"], r["K코드"], r["품명"], r["기존단가"], r["최종인하단가"], r["할인율"], note])
        if r["구분"] == "데이터오류":
            mark_fill(ws, rr, 1, ERROR_FILL)
            mark_fill(ws, rr, 6, ERROR_FILL)
        else:
            mark_fill(ws, rr, 1, REVIEW_FILL)
    for c in [4, 5]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=6).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["G"].width = 45


def main():
    rows = load_items()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_message_sheet(wb)
    build_list_sheet(wb, rows)
    wb._sheets = [wb["요청메시지(초안)"], wb["확인요청_품목목록"]]

    out_path = OUTPUT_DIR / "16_대성_할인율_재검토_요청.xlsx"
    wb.save(out_path)

    n_low = sum(1 for r in rows if r["구분"] == "할인율낮음")
    n_err = sum(1 for r in rows if r["구분"] == "데이터오류")
    print("=== 16단계: 대성인포텍 할인율 재검토 요청 목록 생성 완료 ===")
    print(f"할인율 낮음(재검토 요청): {n_low}건, 데이터 이상치(오류 확인 요청): {n_err}건")
    print(f"생성 파일: {out_path}")


if __name__ == "__main__":
    main()
