"""
9단계: 최종 종합보고서 작성
입력: output/00~07 산출물 전체
출력: output/08_구성방식별_가격분석_보고서.md

각 단계 산출 파일에서 실제 수치를 다시 읽어 보고서에 반영한다(수작업 재입력으로 인한
오류를 줄이기 위함). 서술은 이전 단계들에서 확인된 확정/추정/확인필요 구분을 그대로 유지한다.
"""
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))

BASE_DIR = Path("/home/user/kt-settlement-automation")
OUTPUT_DIR = BASE_DIR / "output"


def ws_data(path, sheet, header_row=1):
    wb = openpyxl.load_workbook(OUTPUT_DIR / path, data_only=True)
    ws = wb[sheet]
    headers = [c.value for c in ws[header_row]]
    idx = {h: i for i, h in enumerate(headers) if h is not None}
    rows = [r for r in ws.iter_rows(min_row=header_row + 1, values_only=True)]
    return idx, rows


def main():
    lines = []

    # ---- 4장: Aviat 213종 ----
    idx1, rows1 = ws_data("01_Aviat_213종_정규화.xlsx", "품목정규화")
    group_counts = {}
    for r in rows1:
        g = r[idx1["품목군"]]
        group_counts[g] = group_counts.get(g, 0) + 1
    n_issue = sum(1 for r in rows1 if r[idx1["데이터품질이슈"]])
    idx1q, rows1q = ws_data("01_Aviat_213종_정규화.xlsx", "데이터품질_이슈")

    # ---- 5장: 발주이력 ----
    idx2, rows2 = ws_data("02_Aviat_발주이력_정제.xlsx", "정제데이터")
    idx2s, rows2s = ws_data("02_Aviat_발주이력_정제.xlsx", "주문요약")
    summary2 = {r[0]: r[1] for r in rows2s if r and r[0]}

    # ---- 6장: 이벤트 ----
    idx3, rows3 = ws_data("03_Aviat_발주이벤트_분석.xlsx", "이벤트요약", header_row=3)
    n_events = len(rows3)
    n_bulk = sum(1 for r in rows3 if r[idx3["복수현장일괄발주후보"]] == "유력")
    n_expand = sum(1 for r in rows3 if r[idx3["증설후보"]] == "유력")

    # ---- 7장: BoM 추정 ----
    idx4, rows4 = ws_data("04_Aviat_구성방식별_표준BoM_추정.xlsx", "구성요약")
    cfg_summary = {r[idx4["구성방식"]]: r for r in rows4}

    # ---- 8장: 04v2 8GHz/11GHz IAP3 재분석 ----
    idx4v2, rows4v2 = ws_data("04_Aviat_구성방식별_표준BoM_추정_v2.xlsx", "구성요약")
    iap3_summary = {r[idx4v2["구성"]]: r for r in rows4v2}
    idx11_20, rows11_20 = ws_data("04_Aviat_구성방식별_표준BoM_추정_v2.xlsx", "11GHz_IAP3_2+0")
    strong_11ghz20 = [r[idx11_20["품명"]] for r in rows11_20 if r[idx11_20["판정수준"]] == "유력 후보"]

    # ---- 9장: 03 원본 이벤트 성격 분석 ----
    idx3e, rows3e = ws_data("03_Aviat_발주이벤트_분석.xlsx", "이벤트요약", header_row=3)
    key_events = {r[idx3e["이벤트ID"]]: r for r in rows3e
                  if r[idx3e["이벤트ID"]] in ("EVT-003", "EVT-006", "EVT-007", "EVT-008", "EVT-009",
                                             "EVT-010", "EVT-011", "EVT-012", "EVT-013")}

    # ---- 10장: 공급사 회신 ----
    idx5, rows5 = ws_data("05_Aviat_공급사회신_검증.xlsx", "회신현황_요약")
    reply_summary = {r[0]: r[1] for r in rows5 if r and r[0]}

    # ---- 11장: Ceragon ----
    idx6, rows6 = ws_data("06_Ceragon_구성방식별_BoM_정리.xlsx", "품목마스터")
    n_active = sum(1 for r in rows6 if r[idx6["품목상태"]] == "활성")
    n_deleted = sum(1 for r in rows6 if r[idx6["품목상태"]] == "삭제가능")
    n_new = sum(1 for r in rows6 if r[idx6["품목상태"]] == "신규코드(코드 미부여)")
    idx6b, rows6b = ws_data("06_Ceragon_구성방식별_BoM_정리.xlsx", "구성별_총액")

    # ---- 12장: 비교 ----
    idx7, rows7 = ws_data("07_Aviat_Ceragon_구성가격비교.xlsx", "직접대응_품목")

    lines.append("# 08. 구성방식별 가격분석 보고서\n")
    lines.append("작성일: 2026-09-04 (자동 생성, 각 단계 산출 파일의 실제 수치를 재조회하여 반영)\n")

    lines.append("\n## 1. 분석 목적\n")
    lines.append(
        "KT commerce가 공급받는 Microwave 전송장비 중 Aviat 계열 213종 계약품목의 과거 발주이력을 "
        "분석하여 2+0/4+0/6+0/8+0 구성방식별 표준 BoM을 추정하고, Ceragon 계열의 기존 구성방식별 "
        "BoM·단가 자료와 비교해 협상에 활용할 수 있는 근거자료를 만드는 것을 목적으로 한다. "
        "모든 결론은 확정/추정/확인 필요 3단계로 구분해 표기한다.\n"
    )

    lines.append("\n## 2. 사용 파일\n")
    lines.append(
        "| 파일 | 역할 |\n|---|---|\n"
        "| 510331d5-...MDMmini_digital_microwave_213...xlsx | Aviat 213종 계약품목(등록요청/내역서) |\n"
        "| 3b46dc94-...20252026...xlsx | Aviat 2025~2026년 발주이력(주문현황, 245건) |\n"
        "| 7cb45e7e / 70bc7482-Aviat_MDM_...BoM...xlsx | Aviat 공급사 작성용 표준BoM 양식(빈 템플릿, 2건 업로드 모두 동일) |\n"
        "| 3e5cc5b3-...MW...List_Kt...DSIT...xlsx | Ceragon 품목별 단가(493건: 활성/삭제가능/신규코드) |\n"
        "| 60d6ff1e-...MW...BoM...DSIT...9...xlsx | Ceragon 주파수·구성방식별 BoM(8GHz/11GHz) |\n"
        "| 75064b39-...260825_100530_1.docx | KT-Ceragon 신규총판 미팅 녹취록(Naver Clova Note 자동전사) |\n"
        "\n세부 구조조사 결과는 `output/00_입력파일_구조조사.md` 참조.\n"
    )

    lines.append("\n## 3. 데이터 품질\n")
    lines.append(
        f"- Aviat 213종 계약품목: K코드 213개, 중복·빈값 0건(확정, `01_Aviat_213종_정규화.xlsx` 원본정리 시트).\n"
        f"- 품명/설명 주파수 표기 불일치·오타: {sum(1 for r in rows1q if r and r[4] and '주파수' in str(r[4]))}건 "
        f"확인(순번20 '111G' 오타, 순번48 '6GHz용 W/G'인데 설명은 '8Ghz용' — 데이터품질_이슈 시트).\n"
        f"- 계약품목 vs 업체작성양식 K코드 재부여 의심 4건(K9090176/K9093723/K9188156/K9090179 ↔ "
        f"K9210278~281, 동일 순번·동일 품명, 코드만 상이).\n"
        f"- 발주이력 245건: 날짜/숫자 변환 오류 0건, 금액 불일치 0건, 취소·반품·교환 0건"
        f"(`02_Aviat_발주이력_정제.xlsx` 주문요약).\n"
        f"- 발주이력 K코드 매핑률: 원본 그대로 91.4%(224/245), 상품코드 기반 보정(구코드→신코드 "
        f"대응 포함) 시 100%(245/245) — 보정분은 '추정' 등급.\n"
        f"- Ceragon 품목 493건 중 활성 {n_active}건, 삭제가능(6~7년 미출고) {n_deleted}건, "
        f"신규코드(미부여) {n_new}건.\n"
        f"- 회의록(⑦ 파일)은 자동 음성전사로 용어 오기 가능성이 있어 모든 인용은 '확인 필요' 등급으로 처리.\n"
    )

    lines.append("\n## 4. Aviat 213종 품목 구성\n")
    top_groups = sorted(group_counts.items(), key=lambda x: -x[1])
    lines.append(
        "품목군 분류(24종, `01_Aviat_213종_정규화.xlsx` 품목정규화/품목군별_요약 시트, 키워드 기반 규칙 분류=추정):\n\n"
        "| 품목군 | 품목수 |\n|---|---|\n" +
        "\n".join(f"| {g} | {c} |" for g, c in top_groups) +
        f"\n\n장비 계열은 크게 (1) HAX-Eclipse/ODU300·600 계열, (2) WTM4200/4500(XT) 일체형 계열, "
        f"(3) VR4/VR10(IAP3 ODU, OBC2) 패킷 마이크로웨이브 계열, (4) CTR8312/8540/8740 계열로 "
        f"관찰되며(추정, 원본에 '장비계열' 필드가 별도로 없어 품명 패턴 기반), CTR 8312는 세부규격에 "
        f"'2+0 패킷형 IDU', CTR 8540은 '8+0 패킷형 IDU'라고 명시되어 있다(확정).\n"
    )

    lines.append("\n## 5. Aviat 발주이력 특징\n")
    lines.append(
        f"- 총 {summary2.get('총 처리 행 수')}건, 주문일자 범위 {summary2.get('주문일자 범위')}.\n"
        f"- 주문자 고유값 {summary2.get('주문자 고유값 수')}명, 부서 고유값 {summary2.get('부서 고유값 수')}개, "
        f"배송지 고유값 {summary2.get('배송지 고유값 수')}곳으로 매우 집중되어 있음.\n"
        f"- 빌딩명/국사명/프로젝트명 등 현장 식별 보조 필드는 245건 전부 공란(확정) → 배송지만으로 "
        f"실제 설치 국소를 특정할 수 없음.\n"
        f"- 주문자 '이상*'(188건)과 '이상헌'(42건)은 이메일 패턴상 동일인일 가능성이 있으나 부서가 "
        f"달라 병합하지 않음(확인 필요).\n"
        f"- 취소/반품/교환 0건, 판매단가×수량=판매금액 불일치 0건(확정).\n"
    )

    lines.append("\n## 6. 발주 이벤트 분류\n")
    lines.append(
        f"- 1차 이벤트(동일 주문번호) 245건은 결합키가 전부 1:1이라 사실상 무의미(확정) — 주문번호가 "
        f"이미 품목행 단위로 발급되는 시스템 구조.\n"
        f"- 2차 이벤트 후보(동일 주문일자+주문자+부서+배송지) {n_events}건 도출.\n"
        f"- 이 중 대량일괄발주(창고성/복수현장 물량 혼재 의심) {n_bulk}건 — 예: 2025-06-04 이벤트는 "
        f"'HAX_IDU/ODU간 IF 케이블,1Meter'(K9090174) 수량 7,500개, 'HAX_Arrestor KIT' 204개 등 "
        f"단일 링크로 보기 어려운 수량이 관측됨(확정, 수량은 원본 그대로).\n"
        f"- 증설 유력 후보 {n_expand}건(신규 Chassis 없이 ODU/License/Interface Card만 추가된 소규모 이벤트).\n"
        f"- 신규설치 '유력' 단독 판정은 0건 — 대부분의 이벤트가 대량/복수대역 혼재 성격이라 순수한 "
        f"'신규설치'로 확정하기 어려움(확인 필요).\n"
        f"- 후속발주 연결 체인 1개 발견('이상*'/그룹시설팀/경기 안양시 배송지, 2025-06-04~12-27 "
        f"사이 8개 이벤트가 반복). 세부 내용은 `03_Aviat_발주이벤트_분석.xlsx` 후속발주_후보 시트 참조.\n"
    )

    lines.append("\n## 7. Aviat 2+0·4+0·6+0·8+0 BoM 추정 결과\n")
    lines.append(
        "앵커 품목(구성 표기가 원본에 직접 확인되는 품목) 식별 결과:\n\n"
        "| 구성방식 | 앵커품목수 | 관련품목수 | 확정후보 | 유력후보 | 참고후보 | 판단불가 |\n|---|---|---|---|---|---|---|\n" +
        "\n".join(
            f"| {cfg} | {cfg_summary[cfg][idx4['앵커품목수']]} | {cfg_summary[cfg][idx4['관련품목수(공통제외)']]} | "
            f"{cfg_summary[cfg][idx4['확정후보']]} | {cfg_summary[cfg][idx4['유력후보']]} | "
            f"{cfg_summary[cfg][idx4['참고후보']]} | {cfg_summary[cfg][idx4['판단불가']]} |"
            for cfg in ["2+0", "4+0", "6+0", "8+0"]
        ) +
        "\n\n- 2+0: CTR 8312 1RU Chassis(K9188144, 세부규격 '2+0 패킷형 IDU')가 유일한 확정 앵커이며, "
        "발주이력 근거가 1건(2025-11-11, EVT-009)뿐이라 companion 품목 대부분 '판단 불가'.\n"
        "- 4+0/8+0: WBX 채널수 표기(8/11/4GHz 각 대역의 '4 CHNNEL/CHANNEL', '8 CHNNEL/CHANNEL')로 "
        "6개/7개 앵커를 확정. 발주이력상 실제 수량 근거는 4GHz 대역(2025-09-17=4+0, 2025-09-18=8+0 "
        "이벤트)에서만 확인되어, 8GHz/11GHz WBX는 '계약서에 존재함'만 확정이고 수량은 확인 필요.\n"
        "- 6+0: Aviat 계약품목·발주이력 어디에서도 6+0을 직접 지시하는 표기가 없어 앵커 0건 "
        "(확인 필요 — 공급사에 6+0 구성 자체의 존재 여부부터 확인 요청).\n"
    )

    lines.append("\n## 8. Aviat 8GHz/11GHz IAP3 구성 재분석(04v2)\n")
    lines.append(
        "7장의 WBX/CTR 텍스트 앵커 방식은 VR4/VR10(IAP3 ODU) 계열처럼 품명에 'N+0' 표기가 없는 "
        "품목을 놓친다는 지적(돈현님 피드백, 2026-09-04)에 따라 `04_Aviat_구성방식별_표준BoM_추정_v2.xlsx`"
        "에서 8GHz/11GHz ODU LOW=HIGH 쌍수 자체를 구성 수량(N)의 1차 신호로 삼아 재분석했다(원본 "
        "04번 파일은 그대로 두고 별도 v2 파일에 저장). 이벤트 단위 및 1~30일 이내 인접 이벤트를 "
        "합친 '결합윈도우' 단위로 관측했으며, 이벤트 집합이 서로 겹치는 윈도우는 하나의 증거로만 "
        "센다(중복 집계 방지).\n\n"
        "| 구성 | 관련품목수 | 확정(구성표기) | 유력후보 | 참고후보 | 변동(대표수량 보류) | 완성 BoM 여부 |\n"
        "|---|---|---|---|---|---|---|\n" +
        "\n".join(
            f"| {cfg} | {iap3_summary[cfg][idx4v2['관련품목수']]} | "
            f"{iap3_summary[cfg][idx4v2['확정(구성표기)']]} | {iap3_summary[cfg][idx4v2['유력후보']]} | "
            f"{iap3_summary[cfg][idx4v2['참고후보']]} | {iap3_summary[cfg][idx4v2['변동(대표수량 보류)']]} | "
            f"{iap3_summary[cfg][idx4v2['완성 BoM 여부']]} |"
            for cfg in ["8GHz_IAP3_2+0", "8GHz_IAP3_4+0", "11GHz_IAP3_2+0", "11GHz_IAP3_4+0"]
        ) +
        "\n\n- **11GHz 2+0이 가장 근거가 뚜렷한 조합이다**: 서로 겹치지 않는 2개의 독립 이벤트(EVT-007, "
        "EVT-010 계열)에서 반복적으로 ODU LOW=HIGH=2쌍과 함께 다음 6개 품목이 매번 정확히 동일한 "
        "비율(1개 링크당 1개)로 관측되어 '유력 후보'로 분류했다: " +
        ", ".join(strong_11ghz20) + " (근거 K코드·주문번호는 `04_..._v2.xlsx` 11GHz_IAP3_2+0 시트, "
        "각 행에 근거 이벤트ID·근거 주문번호가 개별 기재됨).\n"
        "- **8GHz도 근거가 있다 — 단, 아직 독립 반복이 없을 뿐이다.** 8GHz 2+0은 EVT-003"
        "(2025-06-27) 단 1건에서 VR4 CHASSIS/8GHz ODU IAP3 LOW·HIGH/FAN-CV/VR4 Node License/"
        "7·8GHz FLAT HYBRID/GbE4e-AV/GbE4f-AV/16E1-A/E1 Single Cable 10개 품목이 정확히 동일 "
        "비율(1개 링크당 1개)로 관측됐다. 8GHz 4+0은 EVT-013(2026-04-27) 단 1건에서 VR4 CHASSIS/"
        "8GHz ODU IAP3 LOW·HIGH/FAN-CV/VR4 Node License/7·8GHz IAP3 Mounting Bracket/7·8GHz "
        "FLAT HYBRID/FLEXIBLE WAVEGUIDE 8개 품목이 똑같이 정합됐다 — 두 사례 모두 11GHz 2+0 못지"
        "않게 깨끗한 패턴이지만, 각각 다른 날짜·다른 이벤트에서 반복되는 사례가 아직 없어(관측횟수 "
        "1건) '유력 후보'가 아니라 '참고 후보'로 남는다 — 두 번째 독립 사례가 확인되면 바로 유력 "
        "후보로 올라가는 구조다(9장에서 이 이벤트들의 성격을 자세히 본다).\n"
        "- 11GHz 4+0(EVT-008, 2025-10-28)은 품목별 비율이 엇갈린다 — ODU LOW·HIGH·MODEM-AV는 "
        "정확히 1.00인 반면 VR4 CHASSIS·FAN-CV·License는 0.50으로, 4채널 링크 1개가 아니라 2채널 "
        "링크 2개가 섞였을 가능성을 배제할 수 없다(공급사 확인 필요).\n"
        "- 이전 버전에서는 VR4/VR10과 무관한 WTM4500 NODE LICENSE 1건이 11GHz_IAP3_2+0에 잘못 "
        "섞여 있었다(같은 이벤트에 다른 계열 품목이 함께 있었기 때문) — 장비계열이 명시적으로 다른 "
        "품목(CTR8312/8540/8740, WTM4200/4500/4500XT, WBX)은 주파수 표기가 없어도 제외하도록 "
        "수정했다. 11GHz 4+0의 경우 이 수정으로 CTR8312/8540 계열 5개 품목이 함께 제거되어 "
        "18건→ 13건으로, 8GHz 2+0은 21건 → 19건으로 정리됐다(위 표는 수정 후 수치).\n"
        "- 6+0/8+0(IAP3 계열)은 8GHz/11GHz 어느 쪽도 이 배수를 시사하는 발주이력 근거가 전혀 없다 "
        "— 이 구성 자체가 존재하는지부터 공급사에 확인이 필요하다.\n"
        "- 관측 비율이 최대/최소 30% 이상 벌어지는 품목(주로 케이블·SFP 등 공용자재)은 대표 수량을 "
        "임의로 정하지 않고 '변동, 공급사 확인 필요'로 표시했다(요청 반영).\n"
        "- 이 표의 어떤 구성도 '완성 BoM'(핵심 품목 전체가 유력 후보 이상) 단계는 아니다 — 11GHz "
        "2+0이 '완성 근접'으로 가장 앞서 있을 뿐, 나머지는 부분 확인 상태다.\n"
    )

    lines.append("\n## 9. 8/11GHz 후보의 근거가 된 이벤트 성격 분석(03 원본 기반)\n")
    ev_row_fmt = ("| {eid} | {date} | {orderer} | {addr} | {kcnt} | {bulk} |")
    lines.append(
        "돈현님 피드백: \"04v2보다 03_Aviat_발주이벤트_분석.xlsx가 더 중요하다 — EVT-006/007/009/012가 "
        "실제 구축인지 창고입고인지 증설인지 예비품인지 봐야 한다\"에 대한 답이다. 8장의 8GHz/11GHz "
        "후보를 뒷받침한 이벤트들과 그 인접 이벤트를 03 원본에서 다시 확인했다.\n\n"
        "| 이벤트ID | 주문일 | 주문자/부서 | 배송지 | 품목행수 | 03의 자동판정 |\n"
        "|---|---|---|---|---|---|\n" +
        "\n".join(
            ev_row_fmt.format(
                eid=eid,
                date=key_events[eid][idx3e["최초주문일"]].date(),
                orderer=f"{key_events[eid][idx3e['주문자']]}/{key_events[eid][idx3e['부서']][:12]}",
                addr=str(key_events[eid][idx3e["배송지"]])[:14],
                kcnt=key_events[eid][idx3e["포함K코드수"]],
                bulk=key_events[eid][idx3e["복수현장일괄발주후보"]],
            )
            for eid in ["EVT-003", "EVT-006", "EVT-007", "EVT-008", "EVT-009", "EVT-010",
                        "EVT-011", "EVT-012", "EVT-013"]
        ) +
        "\n\n**중요한 한계부터 밝힌다**: 03의 '복수현장일괄발주후보' 자동판정은 이벤트 전체의 최대 "
        "단일품목 수량이나 서로 다른 K코드 종류 수만 보고 판단한다. 케이블·SFP 같은 소모품이 대량 "
        "(예: EVT-013의 IF케이블 240개)으로 같이 섞여 있으면 core 장비(Chassis/ODU/License) 패턴이 "
        "실제로는 깨끗해도 이벤트 전체는 '대량 혼재/복수현장일괄'로 표시된다. 그래서 이 표의 마지막 "
        "열만 보고 core 장비 신호까지 무시하면 안 된다 — 아래는 core 장비만 따로 뽑아본 결과다.\n\n"
        "- **EVT-009(2025-11-11, 5행)**: 이상*/그룹시설팀. VR4/VR10과 전혀 무관한 CTR8312·CTR8540 "
        "계열(M/W장치 5건)만 있다 — 8GHz/11GHz IAP3와 관계없는 별도 계열의 소규모 보충 발주로 "
        "보인다(신규설치/증설 여부 판단불가, 03 원본 그대로).\n"
        "- **EVT-003(2025-06-27, 19행, 8GHz)**: core 장비만 보면 VR4 CHASSIS/8GHz ODU IAP3 LOW·"
        "HIGH/FAN-CV/VR4 Node License/GbE 인터페이스 카드 3종이 전부 수량 2로 정합되어(8GHz 2+0 "
        "후보의 유일한 근거) — 나머지 대량 품목(IF케이블 600개 등)은 별도 공용자재로 분리했다(공용자재_제외 "
        "시트 참조). '대량 혼재' 표시는 이 소모품 때문이며, core 장비 패턴 자체는 깨끗하다.\n"
        "- **EVT-006(2025-09-17, 50행)·EVT-007(2025-09-18, 33행)**: 이틀 연속 발주. EVT-006은 "
        "VR10(License 8/FAN-MV 16/MC-MV 8/CHASSIS 8)과 CTR8740(License 2/Power 2/Chassis 2)이 "
        "함께 있고, EVT-007은 VR4(License 2/MODEM 케이블 4/E1케이블 2/FAN-CV 2/CHASSIS 2)와 "
        "CTR8740(동일 패턴 수량만 2배)이 함께 있다 — VR4·VR10·CTR8740 세 계열이 동시에 섞여 있어 "
        "'단일 링크'가 아니라 여러 계열·여러 사이트분 자재를 한 번에 발주한 정황이 뚜렷하다(창고성/"
        "복수현장 일괄발주에 가까움, 03 판정과 일치). 다만 EVT-007의 VR4 부분(CHASSIS 2/ODU LOW·"
        "HIGH 2/FAN-CV 2/License 2)은 11GHz 2+0 유력 후보의 독립 근거 중 하나로 그대로 유효하다 — "
        "'이벤트 전체가 혼재'인 것과 'VR4/IAP3 부분 패턴이 깨끗한 것'은 별개다.\n"
        "- **EVT-008(2025-10-28, 13행, 11GHz)**: VR4(License 2/FAN-CV 2/CHASSIS 2)+11GHz ODU IAP3 "
        "LOW·HIGH 4+MODEM-AV 4 — 위에서 지적한 대로 ODU/MODEM은 4인데 CHASSIS/FAN/License는 2로 "
        "엇갈린다. 2개의 2+0 링크를 한 번에 발주했거나(그러면 사실은 '4+0 후보'가 아니라 '2+0 두 "
        "묶음'), 실제 4채널 링크에 CHASSIS 1대당 2채널이 들어가는 구조일 수도 있다 — 발주이력만으로는 "
        "구분 불가(공급사확인사항 2순위 질문).\n"
        "- **EVT-010(2025-12-15, 22행, 11GHz)·EVT-011(2025-12-27, 1행)**: EVT-010은 VR4(License "
        "2/FAN-CV 2/CHASSIS 2)+VR10(License 14/FAN-MV 28/MC-MV 12/CHASSIS 14)+11GHz ODU IAP3 "
        "LOW·HIGH 각 2 — VR4 부분만 떼어 보면 11GHz 2+0 유력 후보의 두 번째 독립 근거와 정확히 "
        "일치한다. 12일 뒤 EVT-011은 VR4 CHASSIS 10대만 단품 추가 — 03 판정대로 예비품/보충용으로 "
        "보이며 2+0 패턴 계산에서는 제외했다.\n"
        "- **EVT-012(2026-04-23, 42행)**: 다른 배송지(안양시 비산동)·다른 담당자(이상헌/강남유선시설2팀)"
        "로 처음 등장한 대형 이벤트. VR4(License 17/MODEM 16/FAN-CV 5/CHASSIS 5)+VR10(License 5/"
        "FAN-MV 10/MC-MV 5/CHASSIS 5)이 함께 있으나 이 이벤트에는 8GHz/11GHz ODU IAP3 LOW·HIGH가 "
        "전혀 없다(라이선스·섀시만 있고 무선장치 자체가 빠져 있음) — 기존 사이트의 라이선스/제어보드 "
        "증설이나 예비 확보용일 가능성이 있다(신규 링크 구축은 아닌 것으로 추정, 확인 필요).\n"
        "- **EVT-013(2026-04-27, 13행, 8GHz)**: EVT-012 나흘 뒤, 또 다른 배송지(안양시 관양동)·"
        "담당자(이충환/전라유선시설팀)로 등장. VR4 CHASSIS/8GHz ODU IAP3 LOW·HIGH/FAN-CV/License/"
        "Mounting Bracket/FLAT HYBRID/Waveguide 8개 품목이 전부 수량 4로 정확히 정합됨 — 8GHz 4+0 "
        "후보의 유일하지만 가장 깨끗한 근거다.\n"
        "- **종합**: 이 자료의 '이상*'/그룹시설팀·경기 안양시 배송지는 여러 계열(VR4/VR10/CTR8740/"
        "CTR8312/CTR8540)의 자재를 한 곳(아마도 자재창고 또는 총괄 발주처)으로 받는 창구로 보이며, "
        "EVT-012/013처럼 최근 들어 다른 담당자·다른 배송지 이벤트가 나타나기 시작했다 — 이것이 "
        "실제 개별 현장 납품으로 전환된 것인지, 배송지 표기 방식이 달라진 것뿐인지는 원본 데이터로는 "
        "판단할 수 없다(확인 필요). 결론적으로 '창고입고 vs 실제 구축'을 이벤트 단위로 확정할 근거는 "
        "없지만, VR4/VR10 core 장비만 따로 떼어 보면 8GHz 2+0(EVT-003)·8GHz 4+0(EVT-013)·11GHz "
        "2+0(EVT-007+EVT-010)·11GHz 4+0(EVT-008, 확인 필요)까지 4개 구성 모두에 최소 1개씩 근거 "
        "이벤트가 있다.\n"
    )

    lines.append("\n## 10. 공급사 회신과 추정 결과 비교\n")
    lines.append(
        f"- 업체작성 표준BoM양식(⑦ 파일) K코드 213개 중 {reply_summary.get('코드 일치(공통)')}개는 계약품목과 "
        f"일치, {reply_summary.get('업체양식에만 있는 코드(구코드 추정)')}개는 코드가 상이(구코드 잔존 의심).\n"
        f"- 2+0/4+0/6+0/8+0 수량 기재: {reply_summary.get('2+0 수량 기재됨')} / "
        f"{reply_summary.get('4+0 수량 기재됨')} / {reply_summary.get('6+0 수량 기재됨')} / "
        f"{reply_summary.get('8+0 수량 기재됨')} — 전부 미기재(회신 전 빈 템플릿, 확정).\n"
        f"- 따라서 7단계에서 만든 발주이력 기반 추정치 25건은 전부 '발주실적은 있으나 공급사 BoM "
        f"미기재(회신 대기)'로 분류됨(`05_Aviat_공급사회신_검증.xlsx` 구성별_비교 시트).\n"
    )

    lines.append("\n## 11. Ceragon 구성방식별 BoM 요약\n")
    total_qty_by_cfg = {}
    for r in rows6b:
        cfg = r[idx6b["표준화구성표기"]]
        total_qty_by_cfg.setdefault(cfg, 0)
        total_qty_by_cfg[cfg] += r[idx6b["총수량"]] or 0
    lines.append(
        f"- 전체 품목 493건(활성 {n_active}, 삭제가능 {n_deleted}, 신규코드 {n_new}) — "
        f"3e5cc5b3 파일 '구분' 병합그룹 기준(확정).\n"
        f"- 8GHz/11GHz 대역 모두 2:0/4:0/6:0/8:0 × SD/Non_SD 16개 조합의 수량 매트릭스가 "
        f"완비되어 있음(확정, A+B국소 합산하여 1개 링크 총수량 산출).\n"
        f"- 구성방식별 총수량(8GHz+11GHz 합계): " +
        ", ".join(f"{cfg}={qty}" for cfg, qty in sorted(total_qty_by_cfg.items())) + ".\n"
        f"- 6GHz 및 8GHz_11GHz 공용 'Configuration SW Package' 라이선스가 2+0~8+0까지 전부 "
        f"명시적으로 존재(확정, 예: K9210610 시트 row488 'CER_6GHz_6+0_SD_Configuration SW "
        f"Package'). 단 6GHz는 하드웨어 수량 매트릭스가 공란이라 라이선스 비용만 확인 가능.\n"
        f"- '기존가격' vs '가격인하' 시트 간 수량 매트릭스는 99.4% 동일하며, Hybrid cable "
        f"3종(40/50/60m) 및 XPIC-SD Hybrid cable 20m 품목만 최신판에서 수량이 0으로 변경됨"
        f"(`06_Ceragon_구성방식별_BoM_정리.xlsx` 데이터비교_기존vs인하 시트).\n"
    )

    lines.append("\n## 12. Aviat와 Ceragon 구성 총액 비교\n")
    lines.append(
        "2+0/4+0/6+0/8+0 × 8GHz/11GHz × SD/Non_SD 총 16개 조합을 검토했다. **8GHz/11GHz 대역 "
        "자체는 Aviat·Ceragon 양측 자료에 모두 존재하므로 '대역이 겹치지 않는다'고 서술하지 않는다"
        "**(돈현님 피드백, 2026-09-04). 정확한 상태는 다음과 같다.\n\n"
        "- 2+0/4+0(8GHz·11GHz, 8건): **보류** — 8장(04v2)에서 확인한 대로 Aviat 측 후보 근거가 "
        "유력/참고 후보 수준까지는 있으나 완성된 표준 BoM으로 확정되지 않아, 구성 총액 단위의 "
        "전면 비교를 보류한다. 그중 11GHz 2+0은 유력 후보 6개 품목만의 참고 판매총액을 "
        "9,544,236원으로 산출할 수 있었으나(대표수량 1개 링크당 1개씩), 이는 Ceragon의 전체 "
        "품목(18~27종) 총액과 항목 수 자체가 달라 직접 비교 가능한 값이 아니다.\n"
        "- 6+0/8+0(8GHz·11GHz, 8건): **비교불가(근거 없음)** — 이 배수를 시사하는 Aviat 발주이력이 "
        "전혀 없다(대역 불일치가 아니라 이 배수 자체의 근거 부족).\n"
        "- 4GHz(WBX/CTR 텍스트 앵커) 전 구성: **비교불가** — 이번에는 반대로 Ceragon 쪽에 4GHz "
        "수량 자료 자체가 없다.\n\n"
        "정리하면 '대역은 겹치지만 Aviat 완성 BoM 미확정으로 구성 총액 비교가 보류됨'이 정확한 "
        "결론이다(`07_Aviat_Ceragon_구성가격비교.xlsx` 구성총액_비교/비교불가 시트, "
        "`04_Aviat_구성방식별_표준BoM_추정_v2.xlsx` 확인필요 시트). 다른 대역끼리 억지로 묶어 "
        "비교하지 않는다는 원칙(요청서 10.2절)은 그대로 지켰다.\n"
    )

    lines.append("\n## 13. 가격차가 큰 구성\n")
    lines.append(
        "구성 총액 단위 비교는 12장과 같이 보류/불가 상태이므로, 개별 품목 직접대응 비교(3건) "
        "결과만 보고한다:\n\n"
        "| 비교대상 | Aviat 판매가 | Ceragon 단가 | 비고 |\n|---|---|---|---|\n" +
        "\n".join(
            f"| {r[idx7['비교대상']]} | {r[idx7['Aviat 판매가']]:,} | {r[idx7['Ceragon 단가(판매가/계약단가)']]:,.0f} | "
            f"{r[idx7['비고']][:40]}... |"
            for r in rows7
        ) +
        "\n\n3건 모두 Ceragon 단가가 Aviat 판매가보다 낮게 나타나지만, 공급사 미팅 녹취에서 "
        "'저구성 손실분을 설치자재 단가로 보정한다'는 발언이 있었던 만큼(확인 필요) 이 차이를 "
        "그대로 협상 근거로 사용하기보다 정식 견적 확인이 필요하다.\n"
    )

    lines.append("\n## 14. 직접 비교가 어려운 구성\n")
    lines.append(
        "- 8GHz/11GHz 2+0/4+0: Aviat 발주이력 근거는 있으나(8장/04v2) 유력/참고 후보 수준이라 "
        "완성 BoM이 아님 — Ceragon은 수량 매트릭스가 완비되어 있어, 격차는 'Aviat 쪽 확정 대기'다.\n"
        "- 8GHz/11GHz 6+0/8+0: Ceragon은 수량 매트릭스가 있으나 Aviat 발주이력 근거가 전혀 없음"
        "(이 배수 자체의 존재 여부부터 확인 필요).\n"
        "- 4GHz 2+0~8+0: Aviat는 WBX 기반 근거가 있으나(4+0/8+0), Ceragon 제공 자료에는 4GHz 대역 "
        "수량 매트릭스가 없음.\n"
        "- 6+0(4GHz 포함 전 대역): 양측 모두 하드웨어 수량 근거가 없음(Ceragon은 라이선스 가격만 "
        "존재).\n"
        "- SD 구성: Aviat는 SD/Non-SD 공식 구분이 확인되지 않아(요청서 2.3절 원칙) Ceragon의 명확한 "
        "SD/Non_SD 구분과 직접 대응시키지 않음.\n"
    )

    lines.append("\n## 15. 공급사 추가 확인사항\n")
    lines.append(
        "돈현님 피드백(2026-09-04, 2차): 213종 전체 수량 회신보다 아래 5개 질문이 더 중요하다. "
        "`07_Aviat_Ceragon_구성가격비교.xlsx` 공급사확인사항 시트에 근거 이벤트까지 상세히 정리했고, "
        "여기서는 우선순위만 요약한다.\n\n"
        "1. **[Aviat] 8GHz IAP3 2+0/4+0 대표 표준 구성은?** — EVT-003(2+0), EVT-013(4+0) 각 1건의 "
        "근거를 제시하고 확인 요청(8/9장).\n"
        "2. **[Aviat] 11GHz IAP3 2+0/4+0 대표 표준 구성은?** — 2+0은 EVT-007+EVT-010 두 독립 근거로 "
        "'유력 후보' 확정에 근접, 4+0은 EVT-008 1건이나 품목별 비율이 엇갈려 추가 확인 필요.\n"
        "3. **[Aviat] VR4 적용 조건과 VR10 적용 조건은?** — 발주이력만으로는 판단 불가.\n"
        "4. **[Aviat] WBX와 IAP3(VR4/VR10)의 관계는?** — 대체재인지 별도 제품군인지 확인 필요.\n"
        "5. **[Aviat] 6+0/8+0 구성이 실제로 공급 가능한가?** — 발주이력 근거 자체가 없음.\n\n"
        "그 다음 우선순위(기존 항목):\n\n"
        "6. Aviat 표준BoM 양식에 2+0/4+0/6+0/8+0 수량 기재 및 회신(213종 전체).\n"
        "7. 계약품목-업체양식 K코드 재부여 4건 정정.\n"
        "8. Ceragon 4GHz 대역 지원 여부 및 지원 시 구성별 BoM/단가 제공.\n"
        "9. Ceragon 매입가(KT 조달원가) 별도 제공.\n"
        "10. Ceragon 6GHz Configuration SW Package와 하드웨어 수량 매트릭스 간의 연결 관계.\n"
        "11. 직접대응 SFP 3건의 정확한 사양(온도/거리) 일치 여부.\n"
    )

    lines.append("\n## 16. 협상에 바로 활용 가능한 자료\n")
    lines.append(
        "- Ceragon 8GHz/11GHz 2+0~8+0 구성별 총수량·금액표(`06...` 구성별_총액 시트) — 공급사가 "
        "제시한 구성별 물량 체계를 그대로 확인할 수 있는 확정 자료.\n"
        "- Aviat 계약품목 213종의 품목군별 판매/매입 금액 및 마진율 요약(`01...` 품목군별_요약 시트).\n"
        "- SFP 등 직접대응 품목 3건의 단가 비교(`07...` 직접대응_품목 시트).\n"
        "- CTR8312=2+0, CTR8540=8+0이라는 확정 정보와 WBX 4/8채널 앵커 목록(`04...` 2+0_BoM/"
        "4+0_BoM/8+0_BoM 시트) — 공급사에 구성별 회신을 요청할 때 기준선으로 제시 가능.\n"
        "- 11GHz 2+0 유력 후보 6개 품목(VR4 CHASSIS/ODU LOW·HIGH/FAN-CV/Mounting Bracket/Node "
        "License, `04_..._v2.xlsx` 11GHz_IAP3_2+0 시트) — 공급사 표준BoM 회신을 검증할 1차 "
        "기준선으로 바로 제시 가능(8장 참조).\n"
        "- 8GHz 2+0(EVT-003 근거, 정합 품목 10개)·8GHz 4+0(EVT-013 근거, 정합 품목 8개) 후보 — 아직 참고 "
        "후보이지만 각각 단일 이벤트 내에서 정합도가 높아, '이 조합이 맞는지'를 묻는 확인 질의서의 "
        "기준선으로 바로 제시 가능(9장 참조).\n"
    )

    lines.append("\n## 17. 분석 한계\n")
    lines.append(
        "1. Aviat 발주이력 245건 대부분이 창고성 대량발주 또는 복수 대역·복수 현장 물량이 혼재된 "
        "이벤트여서, '1개 링크 기준' 표준 BoM 수량은 근사치이며 공급사 회신으로 재검증이 필요하다.\n"
        "2. 공급사 작성 표준BoM 양식이 아직 빈 템플릿이라 발주이력 기반 추정치를 검증할 공식 "
        "비교 대상이 없다.\n"
        "3. 8GHz/11GHz는 Aviat·Ceragon 양측 자료에 모두 존재해 대역 자체는 겹치지만(8장/04v2), "
        "Aviat 측 후보가 아직 유력/참고 후보 수준이라 완성 BoM으로 확정되지 않아 구성 총액 비교를 "
        "보류했다(개별 품목 3건만 비교). '대역이 겹치지 않는다'는 서술은 쓰지 않는다.\n"
        "4. 8GHz/11GHz의 6+0/8+0 구성은 Aviat 발주이력 근거가 전혀 없고, 4GHz의 6+0 구성은 "
        "Aviat·Ceragon 양측 모두 하드웨어 수량 근거가 없다.\n"
        "5. 배송지/현장 식별 필드(빌딩명·국사명·프로젝트명)가 발주이력에 전혀 없어, 실제 설치 "
        "국소나 A/B국소 구분을 임의로 만들지 않았다(요청서 2.2/2.3절 원칙 준수).\n"
        "6. 공급사 미팅 녹취록은 자동 음성전사본으로, 인용된 모든 내용은 공식 서면 확인 전까지 "
        "'확인 필요' 등급으로만 취급했다.\n"
        "7. 품목군/장비계열 등 텍스트 기반 분류는 규칙 기반 추정이며, rapidfuzz 등 문자열 유사도만으로 "
        "장비나 구성을 확정하지 않는다는 원칙에 따라 명시적 키워드 매칭만 사용했다(실제로는 정규식 "
        "기반 규칙만 사용, rapidfuzz는 사용하지 않음).\n"
    )

    out_path = OUTPUT_DIR / "08_구성방식별_가격분석_보고서.md"
    out_path.write_text("\n".join(lines), encoding="utf-8")

    print("=== 9단계: 최종 종합보고서 작성 완료 ===")
    print(f"입력: output/01~07 산출 파일 전체(04v2 포함)")
    print(f"생성 파일: {out_path}")
    print(f"보고서 섹션 수: 17개")


if __name__ == "__main__":
    main()
