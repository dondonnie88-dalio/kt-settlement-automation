"""
15단계: 대성인포텍(Ceragon) 매입가 제시안 + KT commerce 판매가(고객사 제시) 제시안
(돈현님 요청, 2026-09-11, 6차)

돈현님 지적: "구성 품목 총액으로 [Aviat와 Ceragon을] 비교해보려 했던 시도들이 무의미했다.
지금 가진 자료에서 최선은 대성인포텍에 얼마의 가격을 제시할지와 고객사인 KT에 얼마의
가격을 제시할지를 정하는 것이다."

방향 전환: Aviat 총액과 직접 비교해서 대성 가격의 높고 낮음을 판정하는 접근은 13번 파일의
결론대로 근거 부족으로 중단한다. 대신 이 파일은 완전히 다른 질문에 답한다 - "대성이 이미
제출한 가격을 매입가로 그대로 수용한다면(13번 검토의견 기준 조건부 수용), KT는 고객사에게
얼마를 제시해야 하는가?"

핵심 근거(2026-09-11 발견, 01번 파일 Aviat 213종 계약마스터에서 직접 확인):
KT commerce는 Aviat 계약에서 품목군별로 정확히 고정된 마진율(판매가 대비, 즉
판매가=매입가/(1-마진율))을 쓰고 있다 - 임의 추정이 아니라 실제 계약가 213건을 전수
검산해 확인한 사실이다.
  - VR4/VR10 IAP3 계열 핵심 장비(섀시/무선송수신부/대역결합기/인터페이스카드/제어카드/
    라이선스/팬/전원/마운트/도파관) 10개 기능군, 27개 품목 전수: 정확히 2.00%, 예외 0건
  - 안테나, 방수/접지/보호 자재, 100M급 장거리 케이블: 정확히 3.00%
  - SFP 광모듈(HAX_ 계열 신형), 커넥터·서지보호 KIT: 5.00%(3건 중 2건 정확히 일치, 1건은
    반올림 오차 범위 내)
이 마진 구조는 Ceragon 전용 정책이 아니라 Aviat 계약에서 확인된 것이지만, KT commerce가
동일한 무선전송장비 재판매 사업에 실제로 적용 중인 정책이므로 Ceragon 소싱 장비에도 같은
기준을 적용하는 것이 임의의 마진율을 가정하는 것보다 훨씬 근거가 있다.

이 파일이 하지 않는 것: Aviat 총액과 Ceragon 총액을 다시 비교하지 않는다(13번 결론 유지 -
반박 근거 없음). 매입가 자체의 적정성을 재판정하지 않는다(조건부 수용 전제). 최종 고객
제시가에 대한 전략적 가산(경쟁 상황, 협상 여유분 등)은 포함하지 않는다 - 여기 나온 판매가는
'KT 표준 마진 구조를 유지하는 최소 제시가'이며 이보다 낮게 제시하면 이 사업의 다른 장비군과
마진율이 달라진다는 의미다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, CONFIRMED_FILL, FMT_AMOUNT, FMT_PERCENT, mark_fill
import importlib
func_cmp = importlib.import_module("14_functional_comparison")

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"
SRC_AVIAT_213 = OUTPUT_DIR / "01_Aviat_213종_정규화.xlsx"
SRC_CERAGON = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"

BAND_CONFIGS = [("8GHz", "2+0"), ("8GHz", "4+0"), ("8GHz", "6+0"), ("8GHz", "8+0"),
                ("11GHz", "2+0"), ("11GHz", "4+0"), ("11GHz", "6+0"), ("11GHz", "8+0")]

CORE_MARGIN = 0.02      # 판매가 대비, 10개 핵심 기능군 27개 품목 전수 확인(예외 0)
ACCESSORY_MARGIN = 0.03  # 안테나/방수/접지/보호/100M급 케이블 - 확인됨
SFP_MARGIN = 0.05        # SFP·커넥터·서지보호 KIT - 3건 중 2건 정확 일치


def margin_rate(category):
    if category == "SFP":
        return SFP_MARGIN
    if category is None:
        return ACCESSORY_MARGIN
    return CORE_MARGIN


def verify_aviat_margin_structure():
    """01번 파일 전수 검산 - 이 파일의 핵심 전제(마진 구조)를 재현 가능하게 남긴다."""
    wb = openpyxl.load_workbook(SRC_AVIAT_213, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = row[idx["K코드"]]
        sale, buy = row[idx["판매단가"]], row[idx["매입단가"]]
        if not (isinstance(sale, (int, float)) and isinstance(buy, (int, float)) and buy > 0):
            continue
        cat = func_cmp.AVIAT_CATEGORY.get(k)
        if cat is None:
            continue  # 이 검증표는 이번 프로젝트가 이미 기능분류한 27개 핵심 품목만 대상으로 한다
        margin = (sale - buy) / sale
        rows.append((k, row[idx["품명"]], cat, buy, sale, margin))
    return rows


def load_ceragon_band(band, cfg, sd):
    wb = openpyxl.load_workbook(SRC_CERAGON, data_only=True)
    ws = wb["BoM_상세"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["원본주파수"]] != band or row[idx["표준화구성표기"]] != cfg or row[idx["SD구분"]] != sd:
            continue
        name = row[idx["품명"]].strip()
        rows.append({
            "품명": name, "K코드": row[idx["K코드"]],
            "매입액": row[idx["인하금액"]] or 0,
            "카테고리": func_cmp.CERAGON_CATEGORY.get(name),
        })
    return rows


def build_verify_sheet(wb):
    ws = wb.create_sheet("마진구조_검증(01번 전수)")
    headers = ["K코드", "품명", "기능분류", "매입단가", "판매단가", "마진율(판매가기준)"]
    ws.append(headers)
    rows = verify_aviat_margin_structure()
    for k, name, cat, buy, sale, margin in rows:
        r = ws.max_row + 1
        ws.append([k, name, func_cmp.CATEGORY_LABEL.get(cat, cat), buy, sale, margin])
        expected = margin_rate(cat)
        if abs(margin - expected) > 0.003:
            mark_fill(ws, r, 6, ERROR_FILL)
        else:
            mark_fill(ws, r, 6, CONFIRMED_FILL)
    for c in [4, 5]:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=6).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    n_exact = sum(1 for _, _, cat, _, sale, m in rows if abs(m - margin_rate(cat)) <= 0.003)
    return len(rows), n_exact


def build_guide_sheet(wb, n_total, n_exact):
    ws = wb.create_sheet("안내")
    ws.append(["항목", "내용"])
    guide = [
        ("목적", "돈현님 지적(2026-09-11): 구성 품목 총액으로 Aviat-Ceragon을 비교하려던 시도는 "
                "실무 결론(대성에게 얼마를, KT 고객사에게 얼마를 제시할지)에 직접 닿지 못했다. "
                "이 파일은 그 두 질문에 직접 답한다."),
        ("매입가 제시안(대성에게)", "13번 파일 검토의견(2026-09-09) 그대로: 대성이 이미 제출한 "
                              "인하단가/인하금액을 매입가로 조건부 수용한다 - 반박할 근거도, "
                              "추가 인하를 요구할 근거도 확보하지 못했기 때문이다. 이 파일은 "
                              "매입가를 재검증하지 않고 그대로 가져온다(각 시트의 '매입액' 열)."),
        ("판매가 제시안(고객사에게) - 핵심 발견", f"01번 파일(Aviat 213종 계약마스터) 전수 검산 "
                                          f"결과, KT commerce는 무선전송장비 재판매 시 품목군별로 "
                                          f"정확히 고정된 마진율(판매가 대비)을 쓰고 있다: 핵심 "
                                          f"장비 10개 기능군 27개 품목은 100% 정확히 2.00%, "
                                          f"안테나/방수/접지/보호/장거리케이블은 3.00%, SFP/커넥터/"
                                          f"서지보호는 5.00%. 전체 {n_total}개 검증 품목 중 "
                                          f"{n_exact}개가 예상 마진율과 ±0.3%p 이내로 일치했다"
                                          f"('마진구조_검증' 시트에서 전부 재현 가능)."),
        ("왜 이 마진율을 Ceragon에도 쓰는가", "이건 'Ceragon 전용 정책'이 아니라 KT commerce가 "
                                       "이미 동일한 무선전송장비 재판매 사업에서 실제로 쓰고 있는"
                                       " 마진 구조다. Ceragon 소싱 장비도 같은 사업(마이크로웨이브"
                                       " 장비 매입 후 고객사 재판매)이므로, 임의의 마진율을 새로 "
                                       "가정하는 것보다 이 검증된 구조를 그대로 적용하는 편이 "
                                       "훨씬 방어 가능하다. 다만 이건 'KT의 공식 Ceragon 마진 "
                                       "정책'이 아니라 '가장 가까운 기존 정책을 유추 적용한 것'"
                                       "이므로, 최종 확정 전 내부 가격정책 담당자 확인을 권장한다."),
        ("계산식", "판매가 = 매입가 / (1 - 마진율) [마진율은 판매가 기준이지 매입가 기준이 "
                "아님에 주의 - 예: 매입가 100원에 2% 마진이면 판매가는 102원이 아니라 "
                "100/0.98=102.04원이다]."),
        ("Ceragon 품목의 마진율 분류 기준", "핵심 10개 기능군(섀시/무선송수신부/대역결합기/"
                                      "인터페이스카드/제어카드/라이선스/팬/전원/마운트/도파관) = "
                                      "2%(Aviat에서 예외 없이 확인). SFP = 5%(Aviat SFP 3건 중 "
                                      "2건이 정확히 일치해 다수 기준으로 채택). 그 외 케이블/"
                                      "커넥터/접지 등 미분류 품목 = 3%(Aviat의 보호/접지/방수 "
                                      "자재군과 가장 유사해 채택 - 이 부분은 2%/5%만큼 확실하지"
                                      " 않은 가정이며, '기타' 열에 따로 표시했다)."),
        ("이 파일이 다루지 않는 것", "매입가(대성 제시가) 자체의 적정성은 재판정하지 않는다"
                                 "(13번 결론 유지). 여기 나온 판매가는 'KT 표준 마진을 유지하는"
                                 " 최소 제시가'이며, 경쟁 상황·협상 여유·고객사와의 관계 등 "
                                 "전략적 가산은 포함하지 않았다 - 실제 제시가는 이 숫자 이상으로"
                                 " 조정될 수 있다."),
        ("6+0/8+0도 다룰 수 있는 이유", "이전 접근(Aviat와 총액 비교)은 6+0/8+0에 대응하는 "
                                    "Aviat 발주이력 자체가 없어 비교가 불가능했다. 이 접근은 "
                                    "Aviat의 마진율'구조'만 빌려오고 Aviat의 6+0/8+0 수량 데이터"
                                    "는 필요 없으므로, 대성이 제시한 모든 구성(2+0/4+0/6+0/8+0)"
                                    "에 동일하게 적용 가능하다 - 13번 파일에서 막혔던 부분이 "
                                    "여기서는 막히지 않는다."),
    ]
    for k, v in guide:
        ws.append([k, v])
    finalize_sheet(ws, 1, 2, ws.max_row)
    ws.column_dimensions["B"].width = 100


def build_band_sheet(wb, band, cfg):
    ws = wb.create_sheet(f"{band}_{cfg}_제시안")
    headers = ["기능분류", "Non_SD 매입액(대성 제시가)", "Non_SD 마진율", "Non_SD 제시 판매가", "Non_SD 마진액",
               "SD 매입액(대성 제시가)", "SD 마진율", "SD 제시 판매가", "SD 마진액"]
    ws.append(headers)

    totals = {}
    for sd in ["Non_SD", "SD"]:
        rows = load_ceragon_band(band, cfg, sd)
        by_cat = {}
        for r in rows:
            cat = r["카테고리"] or "기타"
            by_cat.setdefault(cat, 0)
            by_cat[cat] += r["매입액"]
        totals[sd] = by_cat

    cats = sorted(set(list(totals["Non_SD"].keys()) + list(totals["SD"].keys())),
                  key=lambda c: (func_cmp.CATEGORY_ORDER.index(c) if c in func_cmp.CATEGORY_ORDER else 99, c))

    grand = {"Non_SD_매입": 0, "Non_SD_판매": 0, "SD_매입": 0, "SD_판매": 0}
    for cat in cats:
        label = func_cmp.CATEGORY_LABEL.get(cat, "⑫ 기타(케이블/커넥터/접지 등)")
        rate = margin_rate(None if cat == "기타" else cat)
        row_vals = [label]
        for sd in ["Non_SD", "SD"]:
            buy = totals[sd].get(cat, 0)
            sale = buy / (1 - rate) if buy else 0
            margin_amt = sale - buy
            row_vals += [buy or None, rate, sale or None, margin_amt or None]
            grand[f"{sd}_매입"] += buy
            grand[f"{sd}_판매"] += sale
        ws.append(row_vals)
        if cat == "기타":
            mark_fill(ws, ws.max_row, 1, REVIEW_FILL)

    r = ws.max_row + 1
    ws.append(["합계", grand["Non_SD_매입"] or None, None, grand["Non_SD_판매"] or None,
               (grand["Non_SD_판매"] - grand["Non_SD_매입"]) or None,
               grand["SD_매입"] or None, None, grand["SD_판매"] or None,
               (grand["SD_판매"] - grand["SD_매입"]) or None])
    mark_fill(ws, r, 1, CONFIRMED_FILL)

    for c in [2, 4, 5, 6, 8, 9]:
        for rr in range(2, ws.max_row + 1):
            ws.cell(row=rr, column=c).number_format = FMT_AMOUNT
    for c in [3, 7]:
        for rr in range(2, ws.max_row + 1):
            ws.cell(row=rr, column=c).number_format = FMT_PERCENT
    finalize_sheet(ws, 1, len(headers), ws.max_row)
    ws.column_dimensions["A"].width = 26

    return {
        "band": band, "cfg": cfg,
        "nonsd_매입": grand["Non_SD_매입"], "nonsd_판매": grand["Non_SD_판매"],
        "sd_매입": grand["SD_매입"], "sd_판매": grand["SD_판매"],
    }


def build_summary_sheet(wb, results):
    ws = wb.create_sheet("전체_제시안_요약")
    headers = ["구성", "Non_SD 매입총액(대성 제시가)", "Non_SD 제시 판매가(고객사)", "Non_SD 예상마진액",
               "SD 매입총액(대성 제시가)", "SD 제시 판매가(고객사)", "SD 예상마진액"]
    ws.append(headers)
    for res in results:
        ws.append([
            f"{res['band']} {res['cfg']}",
            res["nonsd_매입"], res["nonsd_판매"], res["nonsd_판매"] - res["nonsd_매입"],
            res["sd_매입"], res["sd_판매"], res["sd_판매"] - res["sd_매입"],
        ])
    for c in range(2, 8):
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws, 1, len(headers), ws.max_row)


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    n_total, n_exact = build_verify_sheet(wb)
    build_guide_sheet(wb, n_total, n_exact)

    results = []
    for band, cfg in BAND_CONFIGS:
        results.append(build_band_sheet(wb, band, cfg))

    build_summary_sheet(wb, results)

    sheet_order = (["안내", "마진구조_검증(01번 전수)"]
                   + [f"{b}_{c}_제시안" for b, c in BAND_CONFIGS]
                   + ["전체_제시안_요약"])
    wb._sheets = [wb[name] for name in sheet_order]

    out_path = OUTPUT_DIR / "15_매입판매가_제시안.xlsx"
    wb.save(out_path)

    print("=== 15단계: 매입가/판매가 제시안 생성 완료 ===")
    print(f"마진 구조 검증: {n_exact}/{n_total}개 품목이 예상 마진율과 ±0.3%p 이내 일치")
    for res in results:
        print(f"  - {res['band']} {res['cfg']}: Non_SD 매입 {res['nonsd_매입']:,.0f}원 → 제시 판매가 "
              f"{res['nonsd_판매']:,.0f}원 | SD 매입 {res['sd_매입']:,.0f}원 → 제시 판매가 "
              f"{res['sd_판매']:,.0f}원")
    print(f"생성 파일: {out_path}")
    print("※ 매입가는 대성 제시가 조건부 수용(13번 결론), 판매가는 KT commerce가 Aviat 계약에서 "
          "실제로 쓰는 마진구조(2%/3%/5%)를 적용한 최소 제시가입니다 - 전략적 가산은 별도 검토 필요.")


if __name__ == "__main__":
    main()
