"""
8단계: Aviat와 Ceragon 구성방식별 가격 비교
입력: output/04_Aviat_구성방식별_표준BoM_추정.xlsx, output/06_Ceragon_구성방식별_BoM_정리.xlsx
출력: output/07_Aviat_Ceragon_구성가격비교.xlsx

핵심 제약(요약 시트에도 기록):
- Ceragon 자료는 8GHz/11GHz 대역만 제공되고 4GHz/6GHz는 라이선스 단가 일부를 제외하면 수량
  매트릭스가 없음. 반면 Aviat 발주이력에서 근거가 가장 뚜렷한 구성(4+0/8+0)은 4GHz 대역
  WBX 사례였음 → 동일 주파수 대역에서 양측 모두 신뢰할 만한 구성 총액을 갖는 조합이 없어,
  '구성 총액' 단위의 직접 비교는 대부분 '비교불가'로 처리한다(임의로 다른 대역을 묶지 않음).
- Ceragon 파일에는 단가(공급단가)만 있고 별도의 '매입가'가 없다. 과제 지침에 따라 이 값은
  '판매가 또는 계약단가'로만 사용하고, 매입가 비교에는 사용하지 않는다
  ('Ceragon 매입가 미제공'으로 표시).
- 개별 품목 중 기능·설명이 명확히 동일한 것만 '직접대응_품목'으로 비교한다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, FMT_AMOUNT, mark_fill

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_AVIAT_BOM = OUTPUT_DIR / "04_Aviat_구성방식별_표준BoM_추정.xlsx"
SRC_CERAGON_BOM = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

CONFIGS = ["2+0", "4+0", "6+0", "8+0"]

# 근거가 명확한 개별 품목 직접 대응 쌍(K코드는 Aviat / Ceragon 각각의 원본 K코드)
DIRECT_PAIRS = [
    ("K9198328", "K9183522", "옥외(Outdoor)용 산업용 1G GbE SFP 광모듈",
     "양측 모두 '옥외 Outdoor용 Optic Gbit Module' 설명이 동일. 온도사양은 Aviat -40~+85℃, "
     "Ceragon -30~+50℃로 약간 다름(확인 필요)."),
    ("K9210279", "K9178770", "실내용 1G GbE SFP 광모듈",
     "양측 모두 실내형 1G Gigabit Ethernet SFP. 거리/파장 사양은 Aviat 자료에 명시되어 있지 않음"
     "(확인 필요)."),
    ("K9198347", "K9197307", "10G SFP+ 광모듈(참고용, 완전 동일 사양 아님)",
     "Aviat 품목은 산업용/실외 등급(-40~+85℃)이나 Ceragon 품목은 실외 등급 명시가 없어 "
     "완전히 동일한 사양은 아님 - 참고용 비교(확인 필요)."),
]


def load_aviat_config_totals():
    wb = openpyxl.load_workbook(SRC_AVIAT_BOM, data_only=True)
    result = {}
    for cfg in CONFIGS:
        ws = wb[f"{cfg}_BoM"]
        headers = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers)}
        items = []
        sale_total = 0.0
        buy_total = 0.0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[idx["K코드"]]:
                continue
            qty = row[idx["1개 링크 기준 추정 수량(대표값)"]]
            qty_num = qty if isinstance(qty, (int, float)) else (1 if isinstance(qty, str) and "앵커" in qty else None)
            sale_amt = row[idx["판매금액"]] or 0
            buy_amt = row[idx["매입금액"]] or 0
            items.append({
                "K코드": row[idx["K코드"]], "품명": row[idx["품명"]],
                "수량": qty, "확정수준": row[idx["확정 수준"]],
                "판매금액": sale_amt, "매입금액": buy_amt,
            })
            if isinstance(sale_amt, (int, float)):
                sale_total += sale_amt
            if isinstance(buy_amt, (int, float)):
                buy_total += buy_amt
        result[cfg] = {"items": items, "판매총액": sale_total, "매입총액": buy_total,
                        "주파수": "4GHz(WBX 근거) 또는 미상(계약서 텍스트 근거만 있음, 발주이력 없음)"}
    return result


def load_ceragon_config_totals():
    wb = openpyxl.load_workbook(SRC_CERAGON_BOM, data_only=True)
    ws = wb["구성별_총액"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["원본주파수"]]:
            continue
        key = (row[idx["원본주파수"]], row[idx["표준화구성표기"]], row[idx["SD구분"]])
        result[key] = {
            "품목종류수": row[idx["품목종류수"]], "총수량": row[idx["총수량"]],
            "기존금액합계": row[idx["기존금액합계"]], "인하금액합계": row[idx["인하금액합계"]],
        }
    return result


def load_price_lookup(path, sheet, kcode_col, name_col, sale_col, buy_col=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    lut = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = row[idx[kcode_col]] if kcode_col in idx else None
        if not k:
            continue
        lut[k] = {
            "품명": row[idx[name_col]] if name_col in idx else None,
            "판매가": row[idx[sale_col]] if sale_col in idx else None,
            "매입가": row[idx[buy_col]] if buy_col and buy_col in idx else None,
        }
    return lut


def main():
    aviat_cfg = load_aviat_config_totals()
    ceragon_cfg = load_ceragon_config_totals()
    aviat_price = load_price_lookup(OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx", "원본정리",
                                     "K코드", "품명", "판매단가", "매입단가")
    ceragon_price = load_price_lookup(SRC_CERAGON_BOM, "품목마스터", "K코드", "품명",
                                       "최종인하단가", None)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 구성총액_비교 ----------------
    ws1 = wb.create_sheet("구성총액_비교")
    headers1 = ["구성방식", "Aviat 대역/근거", "Aviat 판매총액(추정)", "Aviat 매입총액(추정)",
                "Ceragon 대역", "Ceragon SD구분", "Ceragon 판매총액(단가기준)", "Ceragon 매입총액",
                "직접 비교 가능 여부", "사유"]
    ws1.append(headers1)
    for cfg in CONFIGS:
        av = aviat_cfg[cfg]
        for band in ["8GHz", "11GHz"]:
            for sd in ["SD", "Non_SD"]:
                cer = ceragon_cfg.get((band, cfg, sd))
                r = ws1.max_row + 1
                comparable = "아니오"
                reason = (
                    f"Aviat {cfg} 구성의 발주이력 근거는 4GHz 대역(WBX) 위주이며 {band} 대역 자체 "
                    f"수량 근거가 없음(계약서 텍스트로 존재만 확인). Ceragon은 {band}/{sd} 데이터가 "
                    f"있으나 대역이 달라 동일 링크 총액으로 합산 비교하지 않음."
                )
                ws1.append([
                    cfg, av["주파수"], av["판매총액"] or None, av["매입총액"] or None,
                    band, sd, cer["인하금액합계"] if cer else None, "Ceragon 매입가 미제공",
                    comparable, reason,
                ])
                mark_fill(ws1, r, 9, ERROR_FILL)
    for c in [3, 4, 7]:
        for r in range(2, ws1.max_row + 1):
            ws1.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws1, 1, len(headers1), ws1.max_row)

    # ---------------- 품목구성_차이 ----------------
    ws2 = wb.create_sheet("품목구성_차이")
    ws2.append(["항목", "Aviat", "Ceragon", "근거/비고"])
    diffs = [
        ("구성 확장 방식", "모듈(카드) 추가형 - IDU 셀프에 모뎀/인터페이스 카드를 추가해 채널 증설",
         "시스템(셀프) 단위 - 1RU(Non-SD)/2RU(SD) 셀프 자체를 교체하며 채널 증설",
         "공급사 미팅 녹취(⑦ 파일) P17, P60 참조. 구두 설명으로 확인 필요 등급."),
        ("SD 셀프 구분", "SD/Non-SD를 구분하는 별도 셀프 코드가 확정되어 있지 않음(확인 필요)",
         "1RU=Non-SD 전용, 2RU=SD 전용으로 원본 파일에 명시(확정)",
         "60d6ff1e 파일 header row6 'IP-20N 1RU/2RU적용' 표기."),
        ("6+0 구성 근거", "계약서/발주이력 어디에도 6+0을 지시하는 표기 없음(확인 필요)",
         "6GHz 대역 및 8GHz_11GHz 대역 모두 'Configuration SW Package' 라이선스가 6+0까지 "
         "명시적으로 존재(확정, 단 하드웨어 수량 매트릭스는 6GHz 대역에서 공란)",
         "60d6ff1e 파일 8G_가격인하 시트 row488~489 등."),
        ("주파수 커버리지", "6/8/11GHz(L6/U6 세분) + 4GHz(WTM4500XT)까지 폭넓게 계약되어 있음",
         "제공 자료 기준 8GHz/11GHz 수량 매트릭스만 존재, 4GHz 자체는 확인 안 됨(확인 필요)",
         "1단계 구조조사 참조."),
        ("가격 구조(공급사 발언)", "-",
         "저구성(1+0/2+0)은 원가 근접/이하로 책정하고 설치자재(Waveguide/Coupler 등)에서 "
         "가격을 보정한다고 발언 - 품목 단위 개별비교 시 왜곡 가능성 있음",
         "공급사 미팅 녹취 P45-46, P132-142 (구두 설명, 확인 필요)."),
    ]
    for row in diffs:
        ws2.append(list(row))
    finalize_sheet(ws2, 1, 4, ws2.max_row)
    ws2.column_dimensions["B"].width = 45
    ws2.column_dimensions["C"].width = 45
    ws2.column_dimensions["D"].width = 45

    # ---------------- 직접대응_품목 ----------------
    ws3 = wb.create_sheet("직접대응_품목")
    headers3 = ["비교대상", "Aviat K코드", "Aviat 품명", "Aviat 판매가", "Aviat 매입가",
                "Ceragon K코드", "Ceragon 품명", "Ceragon 단가(판매가/계약단가)",
                "판매가 차액", "판매가 차이율", "비고"]
    ws3.append(headers3)
    for aviat_k, ceragon_k, label, note in DIRECT_PAIRS:
        av = aviat_price.get(aviat_k, {})
        cer = ceragon_price.get(ceragon_k, {})
        r = ws3.max_row + 1
        ws3.append([
            label, aviat_k, av.get("품명"), av.get("판매가"), av.get("매입가"),
            ceragon_k, cer.get("품명"), cer.get("판매가"), None, None, note,
        ])
        ws3.cell(row=r, column=9).value = f"=D{r}-H{r}"
        ws3.cell(row=r, column=10).value = f"=IF(D{r}=0,0,I{r}/D{r})"
    for c in [4, 5, 8, 9]:
        for r in range(2, ws3.max_row + 1):
            ws3.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws3.max_row + 1):
        ws3.cell(row=r, column=10).number_format = "0.0%"
    finalize_sheet(ws3, 1, len(headers3), ws3.max_row)

    # ---------------- 판매가_비교 ----------------
    ws4 = wb.create_sheet("판매가_비교")
    ws4.append(["비교 유형", "내용", "Aviat", "Ceragon(판매가/계약단가로 사용)", "비고"])
    ws4.append(["개별품목(직접대응)", "SFP 등 직접대응 3건 평균 판매가 차이", "직접대응_품목 시트 참조",
                "직접대응_품목 시트 참조", "3건 모두 Ceragon 단가가 낮게 나타남(확인 필요, 정식 견적 아님)"])
    ws4.append(["구성총액(2+0~8+0)", "동일 대역·동일 구성 기준 총액 비교", "데이터 부족(4GHz 위주)",
                "8GHz/11GHz 매트릭스 존재", "구성총액_비교 시트 전체가 '비교불가' - 근거 부족"])
    finalize_sheet(ws4, 1, 5, ws4.max_row)
    ws4.column_dimensions["B"].width = 40
    ws4.column_dimensions["E"].width = 50

    # ---------------- 매입가_비교 ----------------
    ws5 = wb.create_sheet("매입가_비교")
    ws5.append(["비교 유형", "내용", "Aviat 매입가", "Ceragon 매입가", "비고"])
    for aviat_k, ceragon_k, label, note in DIRECT_PAIRS:
        av = aviat_price.get(aviat_k, {})
        ws5.append([label, f"{aviat_k} vs {ceragon_k}", av.get("매입가"), "Ceragon 매입가 미제공",
                    "Ceragon 파일에는 공급단가만 있고 KT 매입가가 별도로 제공되지 않음"])
        mark_fill(ws5, ws5.max_row, 4, REVIEW_FILL)
    for cfg in CONFIGS:
        ws5.append([f"{cfg} 구성총액", "-", aviat_cfg[cfg]["매입총액"] or None, "Ceragon 매입가 미제공",
                    "동일 사유"])
        mark_fill(ws5, ws5.max_row, 4, REVIEW_FILL)
    for r in range(2, ws5.max_row + 1):
        ws5.cell(row=r, column=3).number_format = FMT_AMOUNT
    finalize_sheet(ws5, 1, 5, ws5.max_row)
    ws5.column_dimensions["B"].width = 30
    ws5.column_dimensions["E"].width = 50

    # ---------------- 비교불가 ----------------
    ws6 = wb.create_sheet("비교불가")
    headers6 = ["구성방식", "대역 조합", "비교불가 사유"]
    ws6.append(headers6)
    for cfg in CONFIGS:
        for band in ["8GHz", "11GHz"]:
            ws6.append([cfg, f"Aviat(4GHz 근거) vs Ceragon({band})",
                        "동일 주파수 대역 기준 데이터가 양측 모두에 존재하지 않아 구성 총액 비교 불가"])
        ws6.append([cfg, "Aviat(자체 4GHz) vs Ceragon(4GHz)",
                    "Ceragon 제공 자료에 4GHz 대역 수량 매트릭스 자체가 없음(라이선스 단가 일부 제외)"])
    finalize_sheet(ws6, 1, len(headers6), ws6.max_row)
    ws6.column_dimensions["C"].width = 60

    # ---------------- 공급사확인사항 ----------------
    ws7 = wb.create_sheet("공급사확인사항")
    ws7.append(["번호", "확인 요청 내용"])
    asks = [
        "Aviat: 8GHz/11GHz 대역 각각의 2+0/4+0/6+0/8+0 구성 표준 BoM 수량 회신(현재 4GHz만 "
        "발주이력 근거가 있어 Ceragon과 동일 대역 비교 불가).",
        "Ceragon: 4GHz 대역 지원 여부 및 지원 시 구성별 BoM/단가 제공.",
        "Ceragon: '매입가'(KT 조달원가) 별도 제공 여부 - 현재 자료는 단가(공급가)만 존재.",
        "Ceragon: 6GHz 대역 'Configuration SW Package' 라이선스가 실제 하드웨어 구성과 어떻게 "
        "연결되는지(현재 수량 매트릭스 공란) 확인.",
        "직접대응_품목 3건(SFP류)의 정확한 온도/거리 사양 일치 여부 재확인.",
    ]
    for i, a in enumerate(asks, start=1):
        ws7.append([i, a])
    finalize_sheet(ws7, 1, 2, ws7.max_row)
    ws7.column_dimensions["B"].width = 100

    # ---------------- 요약 ----------------
    ws8 = wb.create_sheet("요약")
    ws8.append(["항목", "내용"])
    ws8.append(["비교 가능 범위", "개별 품목 단위 3건(SFP류)만 직접 비교 가능. 구성 총액(2+0~8+0) "
                              "단위 비교는 Aviat·Ceragon 양측이 신뢰할 만한 데이터를 갖는 공통 주파수 "
                              "대역이 없어 전량 비교불가로 처리."])
    ws8.append(["직접대응 3건 결과", "SFP류 3건 모두 Ceragon 단가가 Aviat 판매가보다 낮게 나타남(구체 "
                                "수치는 직접대응_품목 시트 참조). 단, 이는 공식 견적이 아니라 참고용 "
                                "단가 비교이므로 협상 근거로 바로 사용하지 말 것."])
    ws8.append(["가격축 원칙 준수", "Ceragon 매입가가 없어 Aviat 매입가는 어떤 Ceragon 수치와도 비교하지 "
                                "않았음('Ceragon 매입가 미제공'으로 명시). 판매가 비교만 수행."])
    ws8.append(["핵심 결론", "협상에 즉시 활용 가능한 것은 개별 SFP 단가 비교뿐이며, 구성 단위 총액 "
                          "비교를 위해서는 5·6단계에서 식별된 확인필요 사항(Aviat 8/11GHz 회신, "
                          "Ceragon 4GHz 지원여부, Ceragon 매입가)이 먼저 해소되어야 함."])
    finalize_sheet(ws8, 1, 2, ws8.max_row)
    ws8.column_dimensions["B"].width = 100

    out_path = OUTPUT_DIR / "07_Aviat_Ceragon_구성가격비교.xlsx"
    wb.save(out_path)

    print("=== 8단계: Aviat-Ceragon 구성가격 비교 완료 ===")
    print(f"입력 파일: {SRC_AVIAT_BOM.name}, {SRC_CERAGON_BOM.name}")
    print(f"처리: 구성 {len(CONFIGS)}종 × 대역 2 × SD구분 2 = {ws1.max_row-1}개 조합 검토, 직접대응 품목 {len(DIRECT_PAIRS)}건")
    print(f"생성 파일: {out_path}")
    print("주요 검증 결과:")
    print(f"  - 구성총액 비교 가능: 0건 / 비교불가: {ws1.max_row-1}건 (주파수 대역 불일치)")
    print(f"  - 판매가 직접비교 가능 품목: {len(DIRECT_PAIRS)}건")
    print(f"  - 매입가 비교 가능 품목: 0건 (Ceragon 매입가 미제공)")
    print(f"  - 공급사 확인 필요 항목: {len(asks)}건")


if __name__ == "__main__":
    main()
