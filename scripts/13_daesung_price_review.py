"""
대성인포텍(Ceragon 총판) 매입단가 확정 여부 결정 지원 자료(돈현님 요청, 2026-09-08).
목적: 대성이 이미 제출한 8GHz/11GHz 2+0/4+0/6+0/8+0 전체 BoM·단가(06번 파일 원본 그대로,
input/60d6ff1e-...xlsx, input/3e5cc5b3-...xlsx)를, 애니콤(Aviat) 발주이력 기반 참고치와
대조해 "이 가격을 확정해도 되는가"를 판단하는 데 쓸 요약을 만든다.

핵심 주의사항(이 파일 전체에 반복 명시):
- Aviat 참고치는 04v2에서 '반복 관측으로 확인된 핵심 품목만' 모은 부분 금액이다(케이블·전원·
  안테나·인터페이스카드 등 다수 품목 제외). 반면 대성 금액은 케이블·커넥터·전원까지 포함한
  완성 링크 총액이다. 그래서 Aviat 참고치가 대성 금액보다 훨씬 작게 나오는 것은 "Aviat가
  훨씬 싸다"는 뜻이 아니라 "비교 범위가 다르다"는 뜻이다. 이 파일의 어떤 숫자도 완성 링크
  단위의 직접 가격비교로 사용하지 않는다 - 유일한 예외는 '직접대응 품목' 시트의 3건뿐이다
  (품명·규격이 사실상 동일해 코드 매칭 없이도 비교 가능).
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, FMT_AMOUNT, mark_fill

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_AVIAT_V2 = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정_v2.xlsx"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"
SRC_COMPARE = OUTPUT_DIR / "07_Aviat_Ceragon_구성가격비교.xlsx"

CONFIGS = [("8GHz", "2+0"), ("8GHz", "4+0"), ("11GHz", "2+0"), ("11GHz", "4+0")]


def load_aviat_ref(band, cfg):
    wb = openpyxl.load_workbook(SRC_AVIAT_V2, data_only=True)
    ws = wb[f"{band}_IAP3_{cfg}"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    # 2026-09-08 정정: '구성비 배수' 열은 N당 배수(비율, 예: 1=채널당 1개)이지 링크 전체
    # 실제 수량이 아니다. 이전 버전은 이 배수를 그대로 단가에 곱해(또는 곱하지 않고 단가만
    # 더해) 2+0은 실제 총액의 1/2, 4+0은 1/4로 축소 계산되어 있었다. '비율1.00 여부' 판정에는
    # 여전히 배수 열을 쓰되, 금액 합산에는 반드시 '실제 수량(링크 전체)' 열을 곱한다.
    n_strong = n_ref = n_clean = 0
    clean_total = 0
    strong_total = 0
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["K코드"]]:
            continue
        level = row[idx["판정수준"]]
        rep = row[idx["구성비 배수(N당 배수, 절대수량 아님)"]]
        qty = row[idx["추정 수량(N+0 링크 전체 가정, 양쪽 사이트 합산)"]]
        price = row[idx["판매단가"]] or 0
        if level == "유력 후보":
            n_strong += 1
        elif level == "참고 후보":
            n_ref += 1
        if rep == 1:
            n_clean += 1
            clean_total += qty * price if isinstance(qty, (int, float)) else price
            items.append((row[idx["품명"]], price))
        if level == "유력 후보" and isinstance(qty, (int, float)):
            strong_total += qty * price
    return {
        "유력후보": n_strong, "참고후보": n_ref,
        "정합품목수(비율1.00)": n_clean, "정합품목_참고총액": clean_total,
        "유력후보_참고총액": strong_total if n_strong > 0 else None,
        "품목상세": items,
    }


def load_ceragon_row(band, cfg, sd):
    wb = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    ws = wb["구성별_총액"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["원본주파수"]] == band and row[idx["표준화구성표기"]] == cfg and row[idx["SD구분"]] == sd:
            return {
                "품목수": row[idx["품목종류수"]], "총수량": row[idx["총수량"]],
                "기존금액": row[idx["기존금액합계"]], "인하금액": row[idx["인하금액합계"]],
            }
    return None


def load_direct_pairs():
    wb = openpyxl.load_workbook(SRC_COMPARE, data_only=True)
    ws = wb["직접대응_품목"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["비교대상"]]:
            continue
        rows.append({
            "비교대상": row[idx["비교대상"]], "등급": row[idx["신뢰도 등급"]],
            "Aviat 판매가": row[idx["Aviat 판매가"]],
            "Ceragon 단가": row[idx["Ceragon 단가(판매가/계약단가)"]], "비고": row[idx["비고"]],
        })
    return rows


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ================= 요약_판단 =================
    ws1 = wb.create_sheet("요약_판단")
    ws1.append(["항목", "내용"])
    ws1.append(["질문", "대성인포텍이 이미 제출한 8GHz/11GHz 2+0/4+0/6+0/8+0 매입단가를 지금 "
                      "확정해도 되는가? (고객사 가격 제시 전 내부 검토용, 2026-09-16 기한)"])
    ws1.append(["결론(핵심)", "'대성 가격이 애니콤보다 몇 배 비싸다/싸다' 식의 직접 비교는 지금 자료로 "
                          "할 수 없다 - Aviat 참고치는 발주이력에서 반복 확인된 핵심 품목 몇 개일 뿐 "
                          "완성 링크가 아니고, 대성 금액은 케이블·커넥터·전원까지 포함한 완성 링크 "
                          "총액이라 범위 자체가 다르다. 아래 '직접대응 품목' 3건(SFP류, 코드 매칭 "
                          "없이도 비교 가능)이 그나마 신뢰할 수 있는 가격 비교점이나, 3건도 스펙 "
                          "일치도가 서로 달라 등급을 나눠서 봐야 한다(신뢰도 등급 열 참조)."])
    ws1.append(["가장 중요한 한 문장(2026-09-08, 3차 검토에서 합의)",
                          "Aviat 11GHz 2+0 유력후보 19,088,472원(링크 환산 가정, 내부 참고값 - "
                          "'사실상 확정'은 아님)은 대성보다 '싸다'는 뜻이 아니라, Aviat 쪽이 아직 "
                          "완성 BoM 금액이 아니라는 뜻이다 - 14번 파일 기준 이 유력후보는 11개 "
                          "핵심 기능군 중 5개(본체/무선송수신부/팬/라이선스/마운트)만 채우고, "
                          "대역결합기·인터페이스카드·제어카드·전원·SFP는 빠져 있다. 그래서 지금 "
                          "단계의 실무적 결론은 '대성 가격이 비싸다/싸다'가 아니라 '아직 판단할 "
                          "근거가 부족하다'이며, 대성 제시가를 반박할 근거도 확정할 근거도 없다."])
    ws1.append(["직접대응 3건 결과", "SFP류 3건 모두 대성(Ceragon) 단가가 Aviat 판매가보다 낮게 "
                                "나타나지만 신뢰도는 다르다: 실내용 1G SFP(준직접 비교, 설명 실질적"
                                " 동일)가 가장 근접하고, 옥외용 1G SFP(조건부 비교, 온도등급 -40~"
                                "+85℃ vs -30~+50℃로 다름)와 10G SFP+(참고 비교, 등급·거리 사양 "
                                "모두 불일치 가능)는 그보다 약하다. 정식 견적 비교가 아니라 가격표상"
                                " 단가 비교라는 한계도 여전하다(상세는 직접대응_품목 시트)."])
    ws1.append(["구성 총액 검토(6+0/8+0)", "대성은 6+0/8+0도 완성 BoM·단가를 제시했으나, Aviat "
                                       "발주이력에는 6+0/8+0(IAP3 계열)을 시사하는 근거가 전혀 없다 "
                                       "- 비교 대상 자체가 없으므로 이 두 구성은 대성 제시가를 검증할 "
                                       "방법이 현재 없다(애니콤 회신 또는 별도 시장가 조사 필요)."])
    ws1.append(["구성 총액 검토(2+0/4+0)", "11GHz 2+0은 Aviat 쪽에 그나마 가장 근거가 있다(독립 2개 "
                                       "이벤트로 확인된 핵심 품목 6종). 8GHz 2+0/4+0, 11GHz 4+0은 "
                                       "단일 이벤트 근거뿐이라 더 약하다. '구성별_비교상세' 시트에 "
                                       "품목수·범위 차이를 병기했으니, 대성 제시 품목 목록에 케이블/"
                                       "커넥터/전원 등이 빠짐없이 들어있는지 체크리스트로 대조하는 "
                                       "용도로 활용을 권한다(가격 수준 비교가 아니라 품목 누락 점검)."])
    ws1.append(["시한 내 권고", "2026-09-16까지 애니콤 공식 회신이 오면 그 수치로 재검토. 회신이 "
                              "늦어지면, 현재로선 대성 가격을 반박할 만한 동일 범위의 Aviat 총액이 "
                              "없다는 점을 감안해 판단해야 한다(확정 반대 근거는 없음 = 확정 찬성 "
                              "근거도 아님 - 별개의 문제)."])
    finalize_sheet(ws1, 1, 2, ws1.max_row)
    ws1.column_dimensions["B"].width = 100
    for r in range(2, ws1.max_row + 1):
        mark_fill(ws1, r, 1, REVIEW_FILL)

    # ================= 구성별_비교상세 =================
    # 유력후보(독립 2개 이상 이벤트로 반복 확인)와 참고후보(단일 관측뿐)를 하나의 '참고총액'으로
    # 섞지 않는다 - 신뢰도가 다른 숫자를 더하면 그 합계가 마치 하나의 근거처럼 보이는 착시가
    # 생긴다. 두 열을 분리해서 각각의 신뢰도를 그대로 남긴다.
    ws2 = wb.create_sheet("구성별_비교상세")
    headers2 = ["구성", "SD구분", "대성 품목수", "대성 기존금액", "대성 인하금액",
                "Aviat 유력후보 품목수", "Aviat 유력후보 총액(독립반복 확인)",
                "Aviat 참고후보 중 비율1.00 품목수", "Aviat 참고후보 총액(단일관측뿐, 약함)", "비고"]
    ws2.append(headers2)
    aviat_cache = {}
    for band, cfg in CONFIGS:
        av = load_aviat_ref(band, cfg)
        aviat_cache[(band, cfg)] = av
        n_ref_clean = av["정합품목수(비율1.00)"] - av["유력후보"]
        ref_clean_total = av["정합품목_참고총액"] - (av["유력후보_참고총액"] or 0)
        for sd in ["Non_SD", "SD"]:
            cer = load_ceragon_row(band, cfg, sd)
            r = ws2.max_row + 1
            note = (f"Aviat 두 총액 모두 대성 {cer['품목수']}개 전체 품목과 범위가 다른 부분합 "
                    f"(케이블·전원·안테나 등 다수 미포함) - 직접 비교 금지, 품목 누락 점검용으로만 "
                    f"사용")
            ws2.append([f"{band} {cfg}", sd, cer["품목수"], cer["기존금액"], cer["인하금액"],
                        av["유력후보"], av["유력후보_참고총액"], n_ref_clean, ref_clean_total, note])
            if av["유력후보"] > 0:
                mark_fill(ws2, r, 7, CONFIRMED_FILL)
            mark_fill(ws2, r, 9, REVIEW_FILL)
    for c in [4, 5, 7, 9]:
        for r in range(2, ws2.max_row + 1):
            ws2.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws2, 1, len(headers2), ws2.max_row)

    # ================= 6+0_8+0_근거없음 =================
    ws3 = wb.create_sheet("6+0_8+0_근거없음")
    headers3 = ["구성", "SD구분", "대성 품목수", "대성 인하금액", "Aviat 근거"]
    ws3.append(headers3)
    for band in ["8GHz", "11GHz"]:
        for cfg in ["6+0", "8+0"]:
            for sd in ["Non_SD", "SD"]:
                cer = load_ceragon_row(band, cfg, sd)
                r = ws3.max_row + 1
                ws3.append([f"{band} {cfg}", sd, cer["품목수"], cer["인하금액"],
                            "없음 - 발주이력에 이 배수를 시사하는 근거 자체가 없음(공급사 확인 필요)"])
                mark_fill(ws3, r, 5, ERROR_FILL)
    for r in range(2, ws3.max_row + 1):
        ws3.cell(row=r, column=4).number_format = FMT_AMOUNT
    finalize_sheet(ws3, 1, len(headers3), ws3.max_row)

    # ================= 직접대응_품목 =================
    ws4 = wb.create_sheet("직접대응_품목")
    headers4 = ["비교대상", "신뢰도 등급", "Aviat 판매가", "대성(Ceragon) 단가", "비고"]
    ws4.append(headers4)
    grade_fill = {"준직접 비교": CONFIRMED_FILL, "조건부 비교": REVIEW_FILL, "참고 비교": ERROR_FILL}
    for row in load_direct_pairs():
        r = ws4.max_row + 1
        ws4.append([row["비교대상"], row["등급"], row["Aviat 판매가"], row["Ceragon 단가"], row["비고"]])
        mark_fill(ws4, r, 2, grade_fill.get(row["등급"], REVIEW_FILL))
    for c in [3, 4]:
        for r in range(2, ws4.max_row + 1):
            ws4.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws4, 1, len(headers4), ws4.max_row)

    sheet_order = ["요약_판단", "구성별_비교상세", "6+0_8+0_근거없음", "직접대응_품목"]
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "13_대성단가_검토요약.xlsx"
    wb.save(out_path)

    print("=== 대성인포텍 단가 검토 요약 생성 완료 ===")
    for band, cfg in CONFIGS:
        av = aviat_cache[(band, cfg)]
        strong_txt = f"{av['유력후보_참고총액']:,}원" if av["유력후보"] > 0 else "없음"
        print(f"  - {band} {cfg}: Aviat 유력후보 {av['유력후보']}건(총액 {strong_txt}), "
              f"참고후보 중 비율1.00 {av['정합품목수(비율1.00)'] - av['유력후보']}건")
    print(f"생성 파일: {out_path}")
    print("※ 6+0/8+0은 Aviat 근거 자체가 없어 비교 불가(별도 시트). 직접대응 3건만 신뢰 가능한 "
          "가격 비교점.")


if __name__ == "__main__":
    main()
