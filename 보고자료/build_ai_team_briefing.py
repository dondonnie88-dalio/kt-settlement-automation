# -*- coding: utf-8 -*-
"""배송 예측 정보 과제 - KT AI팀(에이전트 개발) 공유용 검토 이력 요약.

2026-07-29 신규: KT AI팀이 이 프로젝트를 기반으로 에이전트를 개발하겠다며 코드와
검토 이력을 요청 -> known_risks.md(내부 의사결정 로그, 날짜·시행착오가 그대로 섞여
있어 외부 개발자가 읽기엔 맥락 정리가 안 됨) 원본 대신, 에이전트 개발자가 알아야 할
것 위주로 재정리:
  ① 핵심 설계 결정과 근거 (왜 이렇게 만들었나 - 확장 시 이 전제를 깨면 안 됨)
  ② 이미 검토·검증한 개선 아이디어 (재검증 방지)
  ③ 알려진 한계 - 에이전트 개발 시 유의사항
  ④ 재학습/유지보수 체크리스트 (known_risks.md 상단 원문 재사용)
본부장 승인 하에 공유(2026-07-29, 김정훈 본부장 참조 메일).
"""
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
    r.font.size = Pt(14)
    r.font.color.rgb = NAVY
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(5)
    return h


def add_body(doc, text, size=10, color=DARKGRAY, indent=0.2, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def add_bullet(doc, text, size=10, color=DARKGRAY, indent=0.4, after=3):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run("•  " + text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def make_table(doc, headers, rows, widths):
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
            color, bold, center, size = DARKGRAY, False, False, 9.5
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
    setattr(section, m, Cm(1.8))

style = doc.styles["Normal"]
style.font.name = "맑은 고딕"
style.font.size = Pt(10.5)

# Title
p = doc.add_paragraph()
run = p.add_run("배송 예측 정보 과제")
run.font.size = Pt(22)
run.font.bold = True
run.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
run = p.add_run("검토 이력 요약 - KT AI팀 에이전트 개발 공유용  ·  2026-07-29  ·  네트워크사업1팀 이돈현")
run.font.size = Pt(10.5)
run.font.color.rgb = MIDGRAY
p.paragraph_format.space_after = Pt(4)

p = doc.add_paragraph()
run = p.add_run("작성: 네트워크사업1팀 이돈현  ·  공유: 김정훈 본부장 참조 하 KT AI팀 전달")
run.font.size = Pt(9)
run.font.italic = True
run.font.color.rgb = MIDGRAY
p.paragraph_format.space_after = Pt(8)

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
p.paragraph_format.space_after = Pt(8)

# 오프닝(의도)
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.CENTER
box.autofit = False
box.columns[0].width = Cm(17.4)
cell = box.rows[0].cells[0]
cell.width = Cm(17.4)
set_cell_shading(cell, "F4F6FB")
cell.paragraphs[0].paragraph_format.space_before = Pt(6)
cell.paragraphs[0].paragraph_format.space_after = Pt(2)
r = cell.paragraphs[0].add_run("이 문서의 목적")
r.font.bold = True
r.font.size = Pt(11)
r.font.color.rgb = NAVY
p2 = cell.add_paragraph()
p2.paragraph_format.space_after = Pt(6)
r2 = p2.add_run(
    "코드(serve.py 등 서빙 패키지)와 별도로, 이 시스템을 기반으로 에이전트를 개발하실 때 "
    "알아두시면 좋을 검토 이력을 정리했습니다. 원본 의사결정 로그(known_risks.md)는 "
    "개발 과정의 날짜·시행착오가 그대로 섞여 있어 별도로 첨부하되, 이 문서에서는 "
    "① 왜 이렇게 설계했는지 ② 이미 검토했으나 채택 안 한 것 ③ 아직 남은 한계 "
    "④ 유지보수 체크리스트 순으로 핵심만 정리합니다."
)
r2.font.size = Pt(10.5)
r2.font.color.rgb = DARKGRAY

# ---------------------------------------------------------------
add_h1(doc, "① 핵심 설계 결정과 근거")
add_body(doc,
    "에이전트 개발 시 아래 전제들을 깨지 않도록 유의해 주십시오 - 전부 실측 검증을 거쳐 "
    "확정된 것들입니다.", after=6)

design_rows = [
    ("폴백 체인 6단계 + 레벨별 MIN_N",
     "sku(자기 이력 200건↑) → vendor_mid(협력사×상품군) → vendor → sla_bucket(표준납기일) "
     "→ cat → global, 나머지 레벨은 1000건↑",
     "레벨마다 세분화 이득이 다름을 실측으로 확인 - 일괄 기준이 아니라 레벨별로 따로 정함"),
    ("협력사×대분류(vendor_cat) 레벨 제거",
     "한때 있었으나 폐지",
     "같은 주문 비교 시 세분화가 오히려 노이즈를 키움(vendor 단독이 9구간 중 6구간 더 정확)"),
    ("cat_sla → sla_bucket 단순화",
     "대분류×납기버킷 2단 조합에서 납기버킷 단독으로 변경",
     "2단 조합보다 단독이 커버리지·정확도 모두 우수, 콜드스타트 오차 절반 이하로 축소"),
    ("vendor 경계선(700~999건) 조건부 포함",
     "MIN_N=1000 미달이어도 급성장 이력 없는 협력사만 완화 적용",
     "일괄 완화는 과거 급성장 협력사 오판 사고(12.7%p) 재발 위험 - 위험군만 제외"),
    ("spot(일회성) 발주 배제",
     "협력사가 어쩌다 한 번만 취급한 상품 기록은 학습에서 제외",
     "일회성 거래가 평균을 왜곡 - 제외 후 정규 주문 예측 전 구간 개선"),
    ("공휴일 보정 - 연휴 길이별 차등 적용",
     "single(1일)/short(2~3일)/long(4일+) 버킷별로 보정 적용 구간이 다름",
     "표본이 다양해질수록 안전 구간이 좁아짐 - 버킷 통합 시 오히려 악화 확인"),
    ("레벨별 신뢰도 명시적 노출 (confidence_level/reliable)",
     "레벨마다 오차가 큰 구간의 '위치'가 다름(cat/global은 중간 구간만, vendor_mid는 "
     "반대로 장기 구간)",
     "단일 기준으로 뭉뚱그리면 위험 - day별/레벨별 플래그를 API가 그대로 노출, 숨김·표시는 UI 정책"),
]
make_table(doc, ["설계 결정", "내용", "근거"], design_rows, [Cm(4.2), Cm(5.6), Cm(7.6)])

# ---------------------------------------------------------------
add_h1(doc, "② 이미 검토·검증한 개선 아이디어 (재검증 방지용)")
add_body(doc,
    "아래는 전부 실제 홀드아웃 데이터로 검증까지 마치고 기각된 것들입니다. 에이전트가 "
    "이런 방향을 다시 제안하지 않도록 미리 공유드립니다.", after=6)

tried_rows = [
    ("레벨×day마크 편차 사전보정", ("기각", MIDGRAY), "편차 방향이 해마다 바뀌어 불안정"),
    ("상품군 세분화(중분류→소분류→세분류)", ("기각", MIDGRAY), "일부만 개선, 대다수는 표본 쪼개져 악화"),
    ("최소표본 기준(MIN_N) 일괄 하향", ("기각", MIDGRAY), "급성장 협력사 오판 사고 재발 위험 (조건부 완화만 위 ①에 반영)"),
    ("공휴일 유형별 세부 보정 확대", ("부분 적용", AMBER), "표본이 다양해질수록 안전 구간이 오히려 좁아짐"),
    ("주문 요일(월~일) 반영", ("시도 후 롤백", MIDGRAY), "요일 자체 영향은 크나(최대 51%p) 기존 분류와 섞으면 표본 쪼개져 악화"),
    ("표준납기일 구간 오차 개선(날짜분리·카테고리분리·학습기간 확장 4가지)", ("전부 기각", MIDGRAY), "네 각도 모두 same-order 검증에서 개선 확인 안 됨(세 차례 일관된 결론)"),
    ("낙찰 협력사 전환 시 예측방식 전환(sku→vendor)", ("불필요 확인", GREEN), "협력사 바뀌어도 sku 레벨 정확도 저하 안 됨(2026-07-27 실측)"),
]
make_table(doc, ["시도한 아이디어", "결과", "핵심 이유"], tried_rows, [Cm(6.2), Cm(2.4), Cm(8.8)])

# ---------------------------------------------------------------
add_h1(doc, "③ 알려진 한계 - 에이전트 개발 시 유의사항")

limit_rows = [
    ("vendor_mid 21/30일 구간 잔존 오차(7~8%p)",
     "특정 협력사×월 단위 일회성 이상치가 원인 - 예측 시점엔 알 방법이 없어 모델로 사전 "
     "해소 불가. 월간 자동점검(procurement_delay_flags.txt)으로 사후 감지 중"),
    ("sla_bucket 4~7일 구간 잔존 오차(9%p대)",
     "학습기간 확장 등 4가지 각도 모두 검증했으나 개선 안 됨 - 재고·생산 상태 같은 새로운 "
     "정보 축 없이는 안 풀릴 가능성이 높음"),
    ("낙찰 협력사 전환 자동 감지 미구현",
     "전환 자체는 정확도에 영향 없음을 확인했으나, '전환 시점을 자동으로 알아채는' 기능은 "
     "아직 없음 - 필요하면 추가 개발 여지"),
    ("콜드스타트(cat/global) 폴백",
     "2026H1 실사용 비중은 0%였으나(sla_bucket이 대부분 흡수) 다년치 카탈로그 기준으로는 "
     "약 0.2% 존재 - 완전히 없어진 게 아니라 최종 안전망으로 유지 중"),
]
make_table(doc, ["한계", "설명"], limit_rows, [Cm(5.0), Cm(12.4)])

# ---------------------------------------------------------------
add_h1(doc, "④ 재학습/유지보수 체크리스트")
add_body(doc, "known_risks.md 상단에 정리된 원문입니다 - 재학습(run_all.bat) 시마다 순서대로 확인합니다.", after=4)

checklist = [
    "급성장 협력사 신규 발생 - growth_vendor_flags.txt 갱신분 확인(12.7%p 사고 재발 방지)",
    "발주→출하 지연 신규 발생 - procurement_delay_flags.txt 갱신분 확인",
    "전체 캘리브레이션 - phase1_report.txt에서 전 구간 0.1~1.4%p 수준 유지되는지(3%p 악화 시 롤백 검토)",
    "레벨×day마크 10%p 초과 구간 - 특히 sla_bucket(현재 최대 9.7%p), vendor_mid 21/30일(7~8%p)",
    "vendor 경계선(700~999건) 세그먼트 - 편입된 협력사가 급성장 목록에 새로 걸리지 않았는지",
    "LEVEL_UNRELIABLE_DAYS(serve.py) - 위 오차표가 바뀌면 같이 재산출",
]
for item in checklist:
    add_bullet(doc, item)

add_body(doc,
    "※ 원본 의사결정 로그(known_risks.md)는 별도 첨부합니다 - 특정 항목의 실험 과정·수치를 "
    "더 자세히 보고 싶을 때 참고해 주십시오.", size=9, color=MIDGRAY, after=4)

doc.save("배송예측_KT AI팀_검토이력요약.docx")
print("saved")
