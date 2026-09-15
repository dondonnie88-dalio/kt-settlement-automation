"""
17단계: KT 전송망설계팀向 8GHz/11GHz MW 구성 판매단가 안내 (돈현님 요청, 2026-09-15)

배경: 대성이 16번의 매입단가 제안 42건을 수용해 06번에 반영/재계산을 마쳤다(15135a9).
이제 그 결과를 바탕으로 판매단가를 KT 전송망설계팀에 송부하려 한다.

버퍼(목표가 대비 +3%p) 유지 여부를 논의함:
- 전송망설계팀은 대성 같은 외부 협상 상대가 아니라 KT 내부 조직이라 "높게 불러서
  깎아주는" 앵커링 전략 자체가 어색하고, 받아들여지면 오히려 "왜 이렇게 남기냐"는
  의심을 살 수 있다.
- 그러나 반대로 전송망설계팀이 인하를 요구할 가능성도 배제할 수 없다. 그 경우 버퍼가
  없으면 목표가(마지노선) 밑으로 깎일 위험이 있는 반면, 버퍼가 있으면 목표가까지는
  여유 있게 양보할 수 있다 - 버퍼를 없앴다가 나중에 다시 올려달라고 하기는 어렵지만,
  버퍼에서 목표가로 내려주기는 쉽다는 비대칭성 때문에 버퍼는 유지하기로 함.
- 다만 명분은 "협상 여지(버퍼)"가 아니라 "당사 표준 마진 정책"으로 바꾼다 - 내부
  상대에게 협상용 여지를 뒀다고 밝히는 것보다, 일관된 마진 정책이라고 설명하는 편이
  받아들여지기 쉽고 나중에 깎아줘도 부자연스럽지 않다. 이메일에는 '버퍼'/'최초견적'
  같은 문구를 노출하지 않는다(15번 내부 파일에는 그대로 남겨 추적 가능하게 유지).

숫자는 15번 '전체_제시안_요약' 시트의 '고객 최초견적(버퍼 3%p)' 값을 그대로 쓰되,
전송망설계팀에는 내부 마진/매입가 구조를 노출하지 않고 최종 판매단가만 전달한다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, FMT_AMOUNT, WRAP_TOP

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_15 = OUTPUT_DIR / "15_매입판매가_제시안.xlsx"


def load_sale_prices():
    wb = openpyxl.load_workbook(SRC_15, data_only=True)
    ws = wb["전체_제시안_요약"]
    h = [c.value for c in ws[1]]
    i = {v: k for k, v in enumerate(h)}
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        band, cfg = r[i["구성"]].split(" ")
        rows.append({
            "대역": band, "구성": cfg,
            "Non_SD_판매단가": round(r[i["Non_SD 고객 최초견적(버퍼 3%p)"]]),
            "SD_판매단가": round(r[i["SD 고객 최초견적(버퍼 3%p)"]]),
        })
    return rows


def build_message_sheet(wb, rows):
    ws = wb.create_sheet("판매단가_안내(초안)")
    ws.append(["항목", "내용"])

    def line(r):
        return f"  {r['대역']} {r['구성']}  |  Non_SD {r['Non_SD_판매단가']:,}원  |  SD {r['SD_판매단가']:,}원"

    msg = (
        "안녕하세요, 전송망설계팀 담당자님,\n\n"
        "KT commerce 이돈현입니다.\n\n"
        "금번 8GHz/11GHz 대역 MW 장비(Ceragon) 구성별 판매단가를 안내드립니다. "
        "구성 x SD 여부별 1개 링크(양쪽 사이트 합산) 기준 단가는 아래와 같습니다.\n\n"
        + "\n".join(line(r) for r in rows) + "\n\n"
        "위 단가는 당사 표준 마진 정책(장비군별 2~5%)을 반영해 산정하였습니다. "
        "세부 품목 구성이나 산정 근거가 필요하시면 말씀해 주시면 상세 자료를 "
        "공유드리겠습니다.\n\n"
        "검토 후 문의사항 있으시면 편하게 연락 주세요.\n\n"
        "감사합니다.\n"
        "KT commerce 이돈현 드림"
    )

    guide = [
        ("용도", "전송망설계팀에 보낼 이메일 본문 초안. '구성별_판매단가' 시트를 표로 "
                "첨부하거나 붙여넣으면 된다."),
        ("메시지 초안", msg),
        ("버퍼 관련 판단(내부용, 전송망설계팀에 보내지 않음)", "목표가(마지노선) 대비 "
                "+3%p 버퍼를 유지했다 - 전송망설계팀이 인하를 요구할 가능성이 있어, "
                "버퍼가 있으면 목표가까지 여유롭게 양보 가능한 반면 버퍼 없이 보내면 "
                "깎일 경우 마지노선 밑으로 내려갈 위험이 있다(버퍼->목표가 양보는 "
                "쉽지만 목표가->인상 요청은 어려운 비대칭성). 다만 '협상 여지'라는 "
                "표현 대신 '표준 마진 정책'으로 설명해, 그대로 수용되어도 과도해 "
                "보이지 않고 인하 요청이 와도 자연스럽게 목표가까지 내줄 수 있게 "
                "했다. 실제 비용 구조상 이번 대성 협상 결과, 8/11GHz 구성에 포함된 "
                "SFP 3종은 원래도 마진율(5%) 기준 정가상한에 걸려 있던 품목이라 "
                "매입가가 내려가도 판매단가(목표가)는 거의 그대로이고, 협상 성과는 "
                "판매단가 인하가 아니라 KT 마진 개선으로 귀속되었다."),
        ("주의", "이 초안은 자동 생성된 것이니 보내시기 전에 실제 상황(호칭, 첨부 형식, "
              "일정 등)에 맞게 다듬어서 사용하시기 바랍니다."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=2).alignment = WRAP_TOP


def build_table_sheet(wb, rows):
    ws = wb.create_sheet("구성별_판매단가")
    headers = ["대역", "구성", "Non_SD 판매단가(1개 링크)", "SD 판매단가(1개 링크)"]
    ws.append(headers)
    for r in rows:
        ws.append([r["대역"], r["구성"], r["Non_SD_판매단가"], r["SD_판매단가"]])
    for c in [3, 4]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws, 1, len(headers), ws.max_row)


def main():
    rows = load_sale_prices()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_message_sheet(wb, rows)
    build_table_sheet(wb, rows)
    wb._sheets = [wb["판매단가_안내(초안)"], wb["구성별_판매단가"]]

    out_path = OUTPUT_DIR / "17_전송망설계팀_판매단가_안내.xlsx"
    wb.save(out_path)

    print("=== 17단계: 전송망설계팀向 판매단가 안내 생성 완료 ===")
    for r in rows:
        print(f"  {r['대역']} {r['구성']}: Non_SD {r['Non_SD_판매단가']:,}원 / "
              f"SD {r['SD_판매단가']:,}원")
    print(f"생성 파일: {out_path}")


if __name__ == "__main__":
    main()
