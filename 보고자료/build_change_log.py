# -*- coding: utf-8 -*-
"""검토 미팅(2026-07-10 전후) 이후 변경사항 전체 정리 - 1p 보고서(build_exec_summary.py)의
임원용 압축판과 달리, 실무 확인용으로 빠짐없이 정리한 버전. 이 대화에 없었던 사람도 바로
읽을 수 있도록 내부 코드명(vendor_mid, sla_bucket, MIN_N 등) 대신 쉬운 말로 풀어쓴다.
known_risks.md 원본 검증 기록을 사람이 읽기 쉬운 서술로 옮긴 것 - 재계산 없이 이미
검증된 결과만 정리.

2026-07-20 전면 개편: "글이 너무 많아 한눈에 안 들어온다"는 피드백으로, 반복 구조를
가진 섹션(2. 개선사항 3건, 3-1의 개선 시도 4건)을 표로 압축하고, 표 직후 그 내용을
다시 문장으로 되풀이하던 요약문을 제거했다(표 자체가 요약이므로 중복). 두괄식 결론을
문서 서두와 말미에 각 1회만 배치."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1E, 0x27, 0x61)
DARKGRAY = RGBColor(0x36, 0x45, 0x4F)
MIDGRAY = RGBColor(0x6E, 0x7B, 0x8B)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
AMBER = RGBColor(0xB0, 0x6A, 0x00)


def set_cell_shading(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def add_h1(doc, text):
    h = doc.add_paragraph()
    r = h.add_run(text)
    r.font.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = NAVY
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    return h


def add_h2(doc, text, color=NAVY, size=11.5):
    h = doc.add_paragraph()
    r = h.add_run(text)
    r.font.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(3)
    return h


def add_body(doc, text, size=10, color=DARKGRAY, indent=0.3, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def add_bullet(doc, text, size=10, color=DARKGRAY, indent=0.5, after=4):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run("•  " + text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def make_table(doc, headers, rows, widths, result_col=None, center_from=None):
    """표 공통 빌더. rows의 각 셀은 문자열, 단 result_col 인덱스의 셀은
    (텍스트, 색상) 튜플을 줘서 강조색(굵게+중앙정렬)을 넣을 수 있음.
    center_from을 주면 그 인덱스 이후 컬럼은 평문 셀도 가운데 정렬(숫자 그리드용)."""
    t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for col, w in zip(t.columns, widths):
        col.width = w
    for c, w, lbl in zip(t.rows[0].cells, widths, headers):
        c.width = w
        set_cell_shading(c, "2A3673")
        pc = c.paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rc = pc.add_run(lbl)
        rc.font.bold = True
        rc.font.size = Pt(10)
        rc.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for i, (row, values) in enumerate(zip(t.rows[1:], rows)):
        cells = row.cells
        shade = "F4F6FB" if i % 2 == 0 else "FFFFFF"
        for c, w in zip(cells, widths):
            c.width = w
            set_cell_shading(c, shade)
        for ci, (cell, val) in enumerate(zip(cells, values)):
            color, bold, center, size = DARKGRAY, False, (center_from is not None and ci >= center_from), 9.5
            text = val
            if isinstance(val, tuple):
                text, color = val
                bold, center = True, True
            p = cell.paragraphs[0]
            if center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    return t


doc = Document("minimal_template.docx")

section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
    setattr(section, m, Cm(1.9))

style = doc.styles["Normal"]
style.font.name = "맑은 고딕"
style.font.size = Pt(10.5)

# Title
p = doc.add_paragraph()
run = p.add_run("배송 예측 정보 과제")
run.font.size = Pt(24)
run.font.bold = True
run.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
run = p.add_run("검토 미팅(7/10 전후) 이후 무엇이 바뀌었나 - 전체 정리  ·  작성 2026-07-15(최종 갱신 7/24)  ·  네트워크사업1팀 이돈현")
run.font.size = Pt(10.5)
run.font.color.rgb = MIDGRAY
p.paragraph_format.space_after = Pt(10)

p = doc.add_paragraph()
pPr = p._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "8")
bottom.set(qn("w:space"), "1")
bottom.set(qn("w:color"), "1E2761")
pBdr.append(bottom)
pPr.append(pBdr)
p.paragraph_format.space_after = Pt(10)

# ---------------------------------------------------------------
# 두괄식 결론 (문서 전체에서 이 결론은 여기서 딱 한 번만 서술체로 나옴 - 나머지는 표/근거)
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.CENTER
box.autofit = False
box.columns[0].width = Cm(17.2)
cell = box.rows[0].cells[0]
cell.width = Cm(17.2)
set_cell_shading(cell, "F4F6FB")
cell.paragraphs[0].paragraph_format.space_before = Pt(6)
cell.paragraphs[0].paragraph_format.space_after = Pt(2)
r = cell.paragraphs[0].add_run("결론")
r.font.bold = True
r.font.size = Pt(11)
r.font.color.rgb = NAVY
p2 = cell.add_paragraph()
p2.paragraph_format.space_after = Pt(6)
r2 = p2.add_run(
    "미팅 전 대비 예측 오차가 뚜렷이 줄었습니다(전체 오차 최대 1.8%p → 1.4%p). 주요 안건 4건은 "
    "모두 실측으로 검증해 처리했고, 그 과정에서 개선 3건을 추가로 반영했으며, 남아있는 오차 "
    "2건도 원인이 명확해 서비스 운영에는 지장이 없습니다. 이후 사내 조회 도구에서 발견된 "
    "개선사항 3건도 반영했습니다. 아래 1~5는 이 결론의 근거입니다."
)
r2.font.size = Pt(10.5)
r2.font.color.rgb = DARKGRAY

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ---------------------------------------------------------------
add_h1(doc, "1. 검토 미팅에서 나온 주요 안건 4가지")

directive_rows = [
    ("①", "편차의 평균으로 오차 보정", ("적용 안 함", MIDGRAY),
     "편차 방향이 해마다 바뀌어(작년 기준 보정 시 올해는 더 틀림) 안정적이지 않음"),
    ("②", "중분류·소분류·세분류로 세분화", ("적용 안 함", MIDGRAY),
     "일부 상품(약 20%)만 좋아지고 대다수(약 80%)는 나빠져 전체로는 손해"),
    ("③", "최소 표본 기준(MIN_N) 하향", ("조건부 적용", GREEN),
     "일괄 하향은 과거 사고(급성장 협력사 오판) 재발 위험 → 급성장 이력 없는 협력사 36곳만 "
     "700건까지 완화, 해당 오차 절반 이하로 축소"),
    ("④", "일회성(spot) 발주 배제", ("적용 완료", GREEN),
     "협력사가 어쩌다 한 번만 취급하는 상품 기록을 평균 계산에서 제외 → 정규 주문 예측 전 구간 개선"),
]
make_table(doc, ["", "주요 안건", "결과", "검증 근거"], directive_rows,
           [Cm(0.9), Cm(4.6), Cm(2.7), Cm(9.2)])

add_body(doc,
    "※ \"적용 안 함\"은 결정을 미룬 게 아니라, 실제로 시험해보고 손해라는 걸 확인한 뒤 내린 "
    "결론입니다.", size=9, color=MIDGRAY, after=4)

# ---------------------------------------------------------------
add_h1(doc, "2. 검증 과정에서 추가로 반영한 개선사항")

improve_rows = [
    ("협력사×상품군 학습기간 확장\n(2018~2025)",
     "21~30일 구간만 학습 표본이 부족했던 문제 → 이 레벨만 8년치로 확장",
     "해당 구간 오차 11%대 → 7%대",
     "긴 구간일수록 더 많은 과거 표본 필요 - 부족분을 과거 자료로 보충"),
    ("신규 협력사 예측 레벨 신설",
     "거래 이력 없는 협력사도 계약상 표준납기일은 항상 있음 → 이를 활용한 예측 단계 추가",
     "해당 구간 오차 절반 이하로 축소",
     "\"전혀 모르는 상태\"에서 \"어느 정도 짐작 가능한 상태\"로 전환"),
    ("최소표본 경계선 조건부 완화",
     "700~999건 협력사 38곳 중 급성장 위험 없는 36곳만 자기 데이터 사용 허용\n"
     "(급성장 2곳은 과거 12.7%p 사고와 같은 유형이라 제외, 기존 방식 유지)",
     "해당 협력사 오차 절반 이하로 축소",
     "위험한 곳만 걸러내고 안전한 곳은 기준 완화로 구제"),
]
make_table(doc, ["개선 항목", "내용", "효과", "의미"], improve_rows,
           [Cm(3.4), Cm(6.0), Cm(3.0), Cm(5.0)])

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ---------------------------------------------------------------
add_h1(doc, "3. 남아있는 오차 2가지 - 원인 규명 결과")

add_h2(doc, "3-1. 표준납기일 \"6일\" 구간만 유독 예측이 안 맞는 이유", color=AMBER, size=11)
add_body(doc,
    "신규 협력사 예측(2번)은 표준납기일이 비슷한 주문끼리 구간(예: 4~7일)으로 묶어 평균으로 "
    "계산하는데, 이 4~7일 구간만 오차가 유독 컸습니다. \"담당자들이 5·7일처럼 습관적으로 적어서 "
    "그런가\" 추정했지만 확인 결과는 반대였습니다 - 5·7일 주문은 오히려 정확했고, 드물게 섞인 "
    "\"6일\" 주문만 실제로 느려서 구간 평균과 안 맞았습니다. 전체 정확도에 미치는 영향은 작아 "
    "(안전 기준 10%p 이내) 당장 손보진 않았고, 개선 여지를 아래 방법들로 실측 확인했습니다.")

sla_rows = [
    ("납기일 값별 세부 분리 (4·6·5·7일)", ("기각", MIDGRAY), "오차 오히려 증가 + 10%p 초과 구간 신설"),
    ("카테고리별 분리", ("기각", MIDGRAY), "집계는 좋아 보이나 개별 카테고리는 악화(착시)"),
    ("학습기간 확장 (2022년 추가)", ("기각", MIDGRAY), "표본 이미 충분해 효과 없음"),
    ("학습기간 확장 (2018~2021, 근사값 탐색)", ("기각", MIDGRAY), "원본에 표준납기일 컬럼 누락, 유사 컬럼으로 대체 검증해도 효과 없음"),
    ("\"6일\" 주문만 다음 단계로 넘기기", ("기각", MIDGRAY), "일부(전체의 0.047%)는 효과 있으나 대상이 지나치게 적어 실익 없음"),
]
make_table(doc, ["시도", "판정", "이유"], sla_rows, [Cm(6.5), Cm(2.5), Cm(8.5)])

add_body(doc,
    "현재 확보 데이터로는 더 낮추기 어려운 상태입니다. 2018~2021년 원본에는 표준납기일 "
    "컬럼이 빠져 있어 IT부서 재요청도 검토했으나, 대체 컬럼으로 근사 검증한 결과가 위 "
    "2022년 확장 결과와 세 차례 일관되게 \"효과 없음\"으로 나와 재요청은 진행하지 않기로 "
    "했습니다.", size=9.5, color=MIDGRAY, after=8)

add_h2(doc, "3-2. \"3주~한 달 뒤 도착\" 예측에 남은 오차 - 특정 협력사의 일시적 이슈", color=AMBER, size=11)
add_body(doc,
    "2번 개선 후에도 21~30일 구간에 오차가 남아 추적한 결과, 오차의 약 93%가 소수 협력사에 "
    "몰려 있었고 공통점은 \"특정 한 달만 극단적으로 늦고 다른 달은 정상\"이었습니다(예: "
    "주식회사 케이눅의 생활용품이 4월에만 평균 48일, 다른 달은 정상). 이런 단발성 이슈는 과거 기록에 "
    "사전 신호가 없어 모델로는 예측할 수 없는 유형입니다.")
add_body(doc,
    "확인해보니 서류 처리 지연이 아니라 실제 지연이 맞았습니다 - 발주 접수(0~1일)와 배송(약 "
    "1일)은 정상, \"발주 후 물건을 내보내기까지\"만 수십 일이 걸렸습니다. 지연 위치는 배송이 "
    "아니라 협력사의 조달·재고·포장 단계였습니다.")
add_body(doc,
    "이런 일이 재발할 수 있어 매달 자동 점검하는 도구를 만들었습니다: 협력사·상품군별 평소 "
    "발주→출고 소요일(3년 기록) 대비 최근 한 달이 3배 이상·14일 이상 느려지면 자동으로 찾아냅니다. "
    "첫 실행(2026년 상반기)에서 10건이 잡혔고, 이 중 7건은 수동 분석으로는 못 봤던 것입니다 - "
    "수동 분석은 \"모델 오차가 큰 곳\"만 역추적하는 방식이라 물량이 적은 협력사의 지연은 애초에 "
    "안 보였기 때문입니다. 매월 재실행하면 협력사 배송리드타임 관리 참고자료로 쓸 수 있습니다.")
add_body(doc,
    "(추가 검토, 7/24) 3-1의 \"6일\" 사례처럼 이 구간 오차 큰 주문만 다음 단계로 넘기는 방법도 "
    "검토했으나 여기엔 적용할 수 없습니다 - \"6일\"은 주문 시점에 이미 아는 조건(표준납기일)으로 "
    "미리 골라낼 수 있었지만, 이 오차는 특정 협력사가 특정 한 달에만 일시적으로 늦어지는 것이라 "
    "주문 시점엔 어느 주문이 그럴지 알 방법이 없습니다. 그래서 예측을 미리 고치는 대신, 위에서 "
    "설명한 매월 자동 점검으로 사후에 빨리 잡아내는 현재 방식이 맞는 대응입니다.",
    size=9.5, color=MIDGRAY, after=8)

add_h2(doc, "3-3. (참고) 공휴일 직전 보정의 협력사별 편차 - 검증 완료, 현행 유지", color=AMBER, size=11)
add_body(doc,
    "연휴 직전 배송 지연 보정이 특정 협력사·카테고리엔 오히려 손해인지 확인했습니다. 전체로는 "
    "크게 이득(오차합 52.7p → 7.8p)이지만, 상품권류 등 일부(전체 주문의 0.24%)는 보정이 손해였습니다. "
    "이 카테고리만 보정을 면제하는 규칙을 만들어 검증했지만, 공휴일 민감도가 해마다 달라져(주요 "
    "안건 ①이 기각된 것과 같은 이유) 과거 기준 면제 규칙이 성립하지 않았습니다. 현행 보정 유지가 "
    "검증상 최선입니다.", after=6)

add_h2(doc, "3-4. 레벨별 전 구간 오차 (위 내용의 근거 표)", size=10.5)
add_body(doc,
    "예측 단위(레벨)마다 \"못 미더운 구간의 위치\"가 다릅니다 - 위에서 말로 설명한 vendor_mid "
    "21/30일, sla_bucket 5/7일 수치가 실제로 다른 구간과 비교해 얼마나 튀는지 아래 표로 "
    "확인할 수 있습니다(굵게 표시한 셀이 위 본문에서 다룬 구간).", size=9, color=MIDGRAY, after=4)

level_day_rows = [
    ("sku (235,472건)", "0.5%p", "2.4%p", "3.7%p", "3.0%p", "1.6%p", "0.4%p", "0.5%p", "0.7%p", "0.2%p"),
    ("vendor_mid (37,641건)", "0.0%p", "0.6%p", "0.2%p", "1.1%p", "0.8%p", "1.3%p", "2.9%p", ("7.0%p", AMBER), ("7.8%p", AMBER)),
    ("vendor (154,896건)", "2.6%p", "1.8%p", "1.0%p", "1.2%p", "0.8%p", "0.3%p", "0.2%p", "0.2%p", "0.0%p"),
    ("sla_bucket (18,745건)", "0.4%p", "1.9%p", "4.9%p", ("9.7%p", AMBER), ("9.2%p", AMBER), "5.0%p", "1.6%p", "2.9%p", "1.1%p"),
]
make_table(doc, ["레벨 (n)", "1일", "2일", "3일", "5일", "7일", "10일", "14일", "21일", "30일"],
           level_day_rows, [Cm(4.2)] + [Cm(1.5)] * 9, center_from=1)

add_body(doc,
    "협력사가 상품코드/상품군 정보를 안 넘겨 cat·global(전체 평균)까지 폴백된 주문은 "
    "0건이었습니다 - 표준납기일 활용 레벨(sla_bucket) 덕분에 콜드스타트가 사실상 다 "
    "흡수됐습니다. 매일 단위 등 더 세부적인 수치는 배송예측_분석보고서.xlsx의 "
    "\"캘리브레이션_검증\" 시트 참고.", size=8.5, color=MIDGRAY, after=6)

# ---------------------------------------------------------------
add_h1(doc, "4. 사내 조회 도구(GUI) 개선사항")

gui_rows = [
    ("표준납기일 자동조회",
     "사내 조회 화면이 표준납기일을 전달받지 못해, 실제로는 신규 협력사 예측(레벨 신설)을 "
     "쓸 수 있는 상품도 더 부정확한 \"상품군 평균\"으로만 표시되던 문제를 발견",
     "상품코드만 넣으면 DB에서 표준납기일을 자동으로 찾아 적용 (수기입력 불필요) → 조회 "
     "화면 결과가 실제 서비스 수준과 일치"),
    ("확률 표현 명확화",
     "\"7/25 도착 확률 90%\"라는 문구가 \"그날 도착 확률\"로 오해될 수 있음",
     "\"7/25까지 도착 확률 90%(5일 이내)\"로 누적 개념을 명시 - 조회 화면·보고서 표현 통일"),
    ("화면 표시 정비",
     "핵심 결과·상품 정보 문구가 길어지면서 줄바꿈이 깨지거나 잘리는 문제",
     "창 크기·줄바꿈 폭 조정으로 정상 표기"),
]
make_table(doc, ["항목", "문제", "조치"], gui_rows, [Cm(3.2), Cm(7.2), Cm(6.6)])

add_body(doc,
    "※ 위 세 건은 예측 정확도 개선이 아니라 사내 조회 도구의 사용성·오해 방지 조치입니다 - "
    "3번 항목(오차 원인 규명)과는 성격이 다릅니다.", size=9, color=MIDGRAY, after=6)

# ---------------------------------------------------------------
add_h1(doc, "5. 숫자로 보는 변화")

stat_rows4 = [
    ("전체 캘리브레이션 오차", "최대 1.8%p", "최대 1.4%p", "개선"),
    ("세그먼트(예측 단위) 개수", "3,050개", "3,077개", "27개 증가"),
    ("협력사×상품군 21~30일 오차", "최대 11.6%p", "최대 7.8%p", "크게 개선"),
    ("신규 협력사(콜드스타트) 오차", "해당 레벨 없었음", "최대 9.7%p", "레벨 신설"),
]
make_table(doc, ["항목", "미팅 이전", "지금(7/15)", "비고"], stat_rows4,
           [Cm(6.0), Cm(3.5), Cm(3.5), Cm(4.4)])

add_body(doc,
    "참고: 캘리브레이션 오차는 \"모델이 예측한 확률\"과 \"실제로 그 기간 안에 도착한 비율\"의 "
    "차이입니다(\"5일 이내 80%\" 예측에 실제 82% 도착 시 오차 2%p). 숫자가 작을수록 정확합니다.",
    size=9, color=MIDGRAY, after=6)

doc.save("배송예측_미팅후_변경사항_전체.docx")
print("saved")
