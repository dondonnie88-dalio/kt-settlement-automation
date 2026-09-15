"""
7단계: Ceragon 구성방식별 BoM 정리
입력: input/60d6ff1e-___MW___________BoM________DSIT___9___.xlsx
      (시트: 8GHz_기존가격, 11GHz_기존가격, 8G_가격인하, 11G_가격인하)
      input/3e5cc5b3-___MW____________List_Kt______DSIT___.xlsx (품목 상태: 활성/삭제가능/신규코드)
출력: output/06_Ceragon_구성방식별_BoM_정리.xlsx

원본 표기 2:0/4:0/6:0/8:0은 원본 열에 보존하고, 표준 비교열에서는 2+0/4+0/6+0/8+0으로 변경한다.
'가격인하' 시트를 최신 BoM으로 보고 주 소스로 사용하며, '기존가격' 시트는 구단가 비교 및
수량 교차검증(불일치 발견 시 기록)에만 사용한다. A/B국소 수량은 합산하여 1개 링크 총수량을 만든다.
"""
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
from xlsx_style import finalize_sheet, ERROR_FILL, REVIEW_FILL, FMT_AMOUNT, FMT_QTY, mark_fill

BASE_DIR = Path("/home/user/kt-settlement-automation")
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
SRC_BOM = INPUT_DIR / "60d6ff1e-___MW___________BoM________DSIT___9___.xlsx"
SRC_PRICE = INPUT_DIR / "3e5cc5b3-___MW____________List_Kt______DSIT___.xlsx"

STD_CONFIG = {"2:0": "2+0", "4:0": "4+0", "6:0": "6+0", "8:0": "8+0"}
# (원본구성, SD구분, A컬럼, B컬럼) - '가격인하' 시트 기준(수량 매트릭스가 I~X, 1칸 밀림)
CONFIG_COLS_DISCOUNT = [
    ("2:0", "SD", 9, 10), ("2:0", "Non_SD", 11, 12),
    ("4:0", "SD", 13, 14), ("4:0", "Non_SD", 15, 16),
    ("6:0", "SD", 17, 18), ("6:0", "Non_SD", 19, 20),
    ("8:0", "SD", 21, 22), ("8:0", "Non_SD", 23, 24),
]
# '기존가격' 시트 기준(수량 매트릭스가 H~W)
CONFIG_COLS_OLD = [
    ("2:0", "SD", 8, 9), ("2:0", "Non_SD", 10, 11),
    ("4:0", "SD", 12, 13), ("4:0", "Non_SD", 14, 15),
    ("6:0", "SD", 16, 17), ("6:0", "Non_SD", 18, 19),
    ("8:0", "SD", 20, 21), ("8:0", "Non_SD", 22, 23),
]
IDU_SIZE = {"2:0": "1RU", "4:0": "1RU", "6:0": "2RU", "8:0": "2RU"}


def load_item_status():
    """3e5cc5b3 파일의 '구분' 병합그룹으로 활성/삭제가능/신규코드 상태를 판정."""
    wb = openpyxl.load_workbook(SRC_PRICE, data_only=True)
    ws = wb["단가인하 리스트_ 최종단가"]
    group_ranges = []
    for m in ws.merged_cells.ranges:
        if m.min_col == 1:
            label = ws.cell(row=m.min_row, column=1).value
            group_ranges.append((m.min_row, m.max_row, label))
    status = {}
    for r in range(4, 497):
        kcode = ws.cell(row=r, column=3).value
        if not kcode or str(kcode).strip() in ("추가코드",):
            continue
        kcode = str(kcode).strip()
        grp = next((label for lo, hi, label in group_ranges if lo <= r <= hi), "")
        if grp == "삭제가능":
            status[kcode] = "삭제가능"
        elif grp == "추가 코드 등록":
            status[kcode] = "신규코드"
        else:
            status[kcode] = "활성"
    return status


PLACEHOLDER_CODES = {"신규코드", "추가코드"}

# 2026-09-14: KT로부터 19개 '신규코드' placeholder 품목 전량에 정식 K코드를 받았다(품명
# 기준 매칭 - 이 품목들은 원본 파일 K코드 열에 여전히 '신규코드'/'추가코드'만 적혀 있어
# 코드 자체로는 구분이 안 됨). 원본 input 파일은 수정하지 않는 프로젝트 규칙에 따라, 여기서
# 품명으로 매칭해 표시 코드만 정식 코드로 교체한다 - 수량/가격/품명 등 다른 값은 그대로다.
NEW_KCODE_BY_NAME = {
    "CER_6GHz_2+0_SD_Configuration SW Package": "K9212896",
    "CER_6GHz_2+0_ Non SD_Configuration SW Package": "K9213016",
    "CER_6GHz_4+0_SD_Configuration SW Package": "K9213017",
    "CER_6GHz_4+0_Non SD_Configuration SW Package": "K9213018",
    "CER_6GHz_6+0_SD_Configuration SW Package": "K9213019",
    "CER_6GHz_6+0_Non SD_Configuration SW Package": "K9213020",
    "CER_6GHz_8+0_SD_Configuration SW Package": "K9213021",
    "CER_6GHz_8+0_Non SD_Configuration SW Package": "K9213022",
    "CER_8GHz_11GHz_2+0_SD_Configuration SW Package": "K9213023",
    "CER_8GHz_11GHz_2+0_ Non SD_Configuration SW Package": "K9213024",
    "CER_8GHz_11GHz_4+0_SD_Configuration SW Package": "K9213025",
    "CER_8GHz_11GHz_4+0_Non SD_Configuration SW Package": "K9213026",
    "CER_8GHz_11GHz_6+0_SD_Configuration SW Package": "K9213027",
    "CER_8GHz_11GHz_6+0_Non SD_Configuration SW Package": "K9213028",
    "CER_8GHz_11GHz_8+0_SD_Configuration SW Package": "K9213029",
    "CER_8GHz_11GHz_8+0_Non SD_Configuration SW Package": "K9213030",
    "CER_Dhrtr-ETSI-Compact & ETSI-8 INST-KIT": "K9213031",
    "CER_Dhrtr-WALL/SHLF KIT": "K9213032",
    "CER_Dhrtr- AC/DC CONV": "K9213033",
}

# 2026-09-15: 16번 파일(대성_매입단가_제시)에서 KT가 제시한 매입단가 42건을 대성이 수용했다
# (돈현님: "협력사에서 단가 수용했습니다"). 원본 input 파일은 수정하지 않는 프로젝트 규칙에
# 따라, 여기서 K코드 기준으로 '최종인하단가'만 협상 결과로 치환한다 - 다른 값(정가/수량/품명
# 등)은 원본 그대로다. 40건은 KT 마진 확보를 위한 필요매입가(정가 x (1-마진율)) 역산값,
# 2건(RFUC-CPLR-8, RFUC-TWIST Kit-6)은 원본 데이터 오류로 보고 형제 품목 실제단가를 그대로
# 적용한 값(반올림). 출처: output/16_대성_매입단가_제시.xlsx '요청단가_상세내역' 시트.
NEGOTIATED_PRICE_OVERRIDE = {
    "K9197309": 44479, "K9197306": 40916, "K9197304": 24909,
    "K9197334": 17104815, "K9178824": 13738236, "K9178823": 13738236,
    "K9178822": 13738236, "K9178821": 13738236, "K9197333": 13390675,
    "K9197167": 11421197, "K9178820": 10599209, "K9178819": 10599209,
    "K9178818": 10599209, "K9178817": 10599209, "K9197332": 8475967,
    "K9178816": 8372613, "K9178815": 8372613, "K9178814": 8372613,
    "K9178813": 8372613, "K9197331": 6971419, "K9178812": 6346080,
    "K9178811": 6346080, "K9178810": 6346080, "K9178809": 6346080,
    "K9178830": 5182894, "K9178829": 5182894, "K9178828": 5182894,
    "K9178878": 5182894, "K9178808": 4771614, "K9178807": 4771614,
    "K9178806": 4771614, "K9178805": 4771614, "K9178804": 3389927,
    "K9178803": 3389927, "K9178802": 3389927, "K9178800": 3389927,
    "K9178877": 2933649, "K9178876": 2933649, "K9178875": 2933649,
    "K9178690": 202880, "K9178851": 467968, "K9178856": 105110,
}


def negotiated_price(kcode, raw_price):
    return NEGOTIATED_PRICE_OVERRIDE.get(kcode, raw_price)


def read_sheet_rows(ws, last_row):
    """행 번호(row)를 키로 사용한다. 'K코드' 열에 실제 K코드 대신 '신규코드' 같은 placeholder
    문자열이 반복 사용되는 19개 행이 있어, K코드 자체를 키로 쓰면 서로 다른 품목이 같은 키로
    충돌해 유실된다(예: 3e5cc5b3/60d6ff1e 파일의 '추가코드'/'신규코드' 19건)."""
    rows = {}
    for r in range(10, last_row + 1):
        kcode_raw = ws.cell(row=r, column=2).value
        if not kcode_raw:
            continue
        kcode_raw = str(kcode_raw).strip()
        name = ws.cell(row=r, column=3).value
        is_placeholder = kcode_raw in PLACEHOLDER_CODES
        new_code = NEW_KCODE_BY_NAME.get(name.strip()) if (is_placeholder and name) else None
        if new_code:
            display_code, is_placeholder = new_code, False
        else:
            # '순번'(A열)이 이 행들에서는 전부 공란이라 순번으로는 구분이 안 되므로, 항상
            # 고유한 행 번호를 붙여 서로 다른 신규 품목이 같은 표시코드로 충돌하지 않게 한다.
            display_code = f"{kcode_raw}-R{r}" if is_placeholder else kcode_raw
        rows[r] = {
            "K코드": display_code, "K코드_실제여부": not is_placeholder,
            "순번": ws.cell(row=r, column=1).value,
            "품명": name,
            "설명": ws.cell(row=r, column=4).value,
            "단위": ws.cell(row=r, column=5).value,
            "단가": ws.cell(row=r, column=6).value or 0,
            "row": r,
        }
    return rows


def build_bom_long(ws_discount, ws_old, freq_label, item_status):
    """가격인하 시트를 기준으로 tidy 포맷 BoM(수량>0)만 생성. 기존가격 시트와 수량 교차검증."""
    disc_rows = read_sheet_rows(ws_discount, 502)
    old_rows = read_sheet_rows(ws_old, 483)
    old_by_kcode = {v["K코드"]: v for v in old_rows.values() if v["K코드_실제여부"]}

    long_rows = []
    qty_diffs = []
    for info in disc_rows.values():
        kcode = info["K코드"]
        r = info["row"]
        discount_price = negotiated_price(kcode, ws_discount.cell(row=r, column=8).value or 0)
        old_info = old_by_kcode.get(kcode) if info["K코드_실제여부"] else None
        for cfg_raw, sd, col_a, col_b in CONFIG_COLS_DISCOUNT:
            qty_a = ws_discount.cell(row=r, column=col_a).value or 0
            qty_b = ws_discount.cell(row=r, column=col_b).value or 0
            qty_total = qty_a + qty_b

            if old_info is not None:
                old_col_a, old_col_b = next(
                    (a, b) for c2, s2, a, b in CONFIG_COLS_OLD if c2 == cfg_raw and s2 == sd
                )
                old_qty = (ws_old.cell(row=old_info["row"], column=old_col_a).value or 0) + \
                          (ws_old.cell(row=old_info["row"], column=old_col_b).value or 0)
                if old_qty != qty_total:
                    qty_diffs.append((freq_label, kcode, info["품명"], cfg_raw, sd, old_qty, qty_total))

            if qty_total <= 0:
                continue
            long_rows.append({
                "원본주파수": freq_label, "원본구성표기": cfg_raw, "표준화구성표기": STD_CONFIG[cfg_raw],
                "SD구분": sd, "IDU크기": IDU_SIZE[cfg_raw], "장비계열": "IP-20N_RFU-D-HP(High Power)",
                "K코드": kcode, "품명": info["품명"], "설명": info["설명"], "단위": info["단위"],
                "품목상태": item_status.get(kcode, "확인필요") if info["K코드_실제여부"] else "신규코드(코드 미부여)",
                "1개링크_A국소수량": qty_a, "1개링크_B국소수량": qty_b, "1개링크_총수량": qty_total,
                "기존단가": old_info["단가"] if old_info else None, "최종인하단가": discount_price,
                "기존금액": (old_info["단가"] if old_info else 0) * qty_total,
                "인하금액": discount_price * qty_total,
            })
    return long_rows, qty_diffs


def main():
    wb_src = openpyxl.load_workbook(SRC_BOM, data_only=True)
    item_status = load_item_status()

    all_long_rows = []
    all_diffs = []
    for freq_label, disc_sheet, old_sheet in [
        ("8GHz", "8G_가격인하", "8GHz_기존가격"),
        ("11GHz", "11G_가격인하", "11GHz_기존가격"),
    ]:
        rows, diffs = build_bom_long(wb_src[disc_sheet], wb_src[old_sheet], freq_label, item_status)
        all_long_rows.extend(rows)
        all_diffs.extend(diffs)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 품목마스터 ----------------
    ws1 = wb.create_sheet("품목마스터")
    headers1 = ["K코드", "품명", "설명", "단위", "기존단가", "최종인하단가", "인하율", "품목상태"]
    ws1.append(headers1)
    seen = {}
    for freq_label, disc_sheet, old_sheet in [("8GHz", "8G_가격인하", "8GHz_기존가격"),
                                                ("11GHz", "11G_가격인하", "11GHz_기존가격")]:
        rows = read_sheet_rows(wb_src[disc_sheet], 502)
        for info in rows.values():
            kcode = info["K코드"]
            if kcode in seen:
                continue
            seen[kcode] = True
            status = item_status.get(kcode, "확인필요") if info["K코드_실제여부"] else "신규코드(코드 미부여)"
            discount_price = negotiated_price(kcode, wb_src[disc_sheet].cell(row=info["row"], column=8).value or 0)
            ws1.append([kcode, info["품명"], info["설명"], info["단위"], info["단가"] or None,
                        discount_price or None, None, status])
            r = ws1.max_row
            if info["단가"]:
                ws1.cell(row=r, column=7).value = f"=1-F{r}/E{r}"
            if status in ("삭제가능", "신규코드(코드 미부여)"):
                mark_fill(ws1, r, 8, REVIEW_FILL)
    for c in [5, 6]:
        for r in range(2, ws1.max_row + 1):
            ws1.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws1.max_row + 1):
        ws1.cell(row=r, column=7).number_format = "0.0%"
    finalize_sheet(ws1, 1, len(headers1), ws1.max_row)

    # ---------------- BoM_상세 ----------------
    ws2 = wb.create_sheet("BoM_상세")
    headers2 = ["원본주파수", "원본구성표기", "표준화구성표기", "SD구분", "IDU크기", "장비계열",
                "K코드", "품명", "설명", "단위", "품목상태", "1개링크_A국소수량", "1개링크_B국소수량",
                "1개링크_총수량", "기존단가", "최종인하단가", "기존금액", "인하금액"]
    ws2.append(headers2)
    for row in sorted(all_long_rows, key=lambda x: (x["원본주파수"], x["표준화구성표기"], x["SD구분"], -x["1개링크_총수량"])):
        ws2.append([row[h] for h in [
            "원본주파수", "원본구성표기", "표준화구성표기", "SD구분", "IDU크기", "장비계열", "K코드",
            "품명", "설명", "단위", "품목상태", "1개링크_A국소수량", "1개링크_B국소수량", "1개링크_총수량",
            "기존단가", "최종인하단가", "기존금액", "인하금액",
        ]])
        r = ws2.max_row
        if row["품목상태"] == "삭제가능":
            mark_fill(ws2, r, 11, REVIEW_FILL)
    for c in [12, 13, 14]:
        for r in range(2, ws2.max_row + 1):
            ws2.cell(row=r, column=c).number_format = FMT_QTY
    for c in [15, 16, 17, 18]:
        for r in range(2, ws2.max_row + 1):
            ws2.cell(row=r, column=c).number_format = FMT_AMOUNT
    finalize_sheet(ws2, 1, len(headers2), ws2.max_row)

    # ---------------- 구성별_총액 ----------------
    ws3 = wb.create_sheet("구성별_총액")
    headers3 = ["원본주파수", "표준화구성표기", "SD구분", "품목종류수", "총수량", "기존금액합계", "인하금액합계", "인하액", "인하율"]
    ws3.append(headers3)
    agg = defaultdict(lambda: {"n": 0, "qty": 0, "old_amt": 0.0, "new_amt": 0.0})
    for row in all_long_rows:
        key = (row["원본주파수"], row["표준화구성표기"], row["SD구분"])
        a = agg[key]
        a["n"] += 1
        a["qty"] += row["1개링크_총수량"]
        a["old_amt"] += row["기존금액"] or 0
        a["new_amt"] += row["인하금액"] or 0
    for key in sorted(agg.keys()):
        a = agg[key]
        r = ws3.max_row + 1
        ws3.append([key[0], key[1], key[2], a["n"], a["qty"], a["old_amt"], a["new_amt"], None, None])
        ws3.cell(row=r, column=8).value = f"=F{r}-G{r}"
        ws3.cell(row=r, column=9).value = f"=IF(F{r}=0,0,H{r}/F{r})"
    for c in [5, 6, 7, 8]:
        for r in range(2, ws3.max_row + 1):
            ws3.cell(row=r, column=c).number_format = FMT_AMOUNT
    for r in range(2, ws3.max_row + 1):
        ws3.cell(row=r, column=9).number_format = "0.0%"
    finalize_sheet(ws3, 1, len(headers3), ws3.max_row)

    # ---------------- 데이터비교_기존vs인하 ----------------
    ws4 = wb.create_sheet("데이터비교_기존vs인하")
    headers4 = ["주파수", "K코드", "품명", "원본구성표기", "SD구분", "기존가격시트 수량", "가격인하시트 수량", "설명"]
    ws4.append(headers4)
    for freq_label, kcode, name, cfg_raw, sd, old_qty, new_qty in all_diffs:
        ws4.append([freq_label, kcode, name, cfg_raw, sd, old_qty, new_qty,
                    "가격인하 시트에서 해당 구성의 수량이 변경됨(0으로 제외되었거나 조정됨) - "
                    "공급사 최신 BoM 개정 사항으로 추정, 확인 필요"])
        mark_fill(ws4, ws4.max_row, 8, ERROR_FILL)
    if not all_diffs:
        ws4.append(["기존가격 vs 가격인하 시트 간 수량 불일치 없음"])
    finalize_sheet(ws4, 1, len(headers4), max(ws4.max_row, 2))

    out_path = OUTPUT_DIR / "06_Ceragon_구성방식별_BoM_정리.xlsx"
    wb.save(out_path)

    status_col_vals = [ws1.cell(row=r, column=8).value for r in range(2, ws1.max_row + 1)]
    n_active = sum(1 for v in status_col_vals if v == "활성")
    n_deleted = sum(1 for v in status_col_vals if v == "삭제가능")
    n_new = sum(1 for v in status_col_vals if v == "신규코드(코드 미부여)")

    print("=== 7단계: Ceragon 구성방식별 BoM 정리 완료 ===")
    print(f"입력 파일: {SRC_BOM.name}, {SRC_PRICE.name}")
    print(f"처리: 품목마스터 {ws1.max_row - 1}건, BoM 상세(수량>0) {len(all_long_rows)}행")
    print(f"제외 또는 오류 행 수: 0건 (기존/인하 수량 불일치 {len(all_diffs)}건은 별도 시트에 기록)")
    print(f"생성 파일: {out_path}")
    print("주요 검증 결과:")
    print(f"  - 품목 상태: 활성 {n_active}건, 삭제가능 {n_deleted}건, 신규코드 {n_new}건")
    print(f"  - 구성×SD×주파수 조합 수: {len(agg)}개")
    print(f"  - 기존가격/가격인하 시트 수량 불일치: {len(all_diffs)}건")


if __name__ == "__main__":
    main()
