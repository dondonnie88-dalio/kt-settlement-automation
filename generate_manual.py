"""
KT 정산 자동화 사용 안내 PPTX 생성 스크립트
실행: python generate_manual.py
출력: KT정산자동화_사용안내.pptx (같은 폴더에 생성)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import copy, os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "KT정산자동화_사용안내.pptx")

prs = Presentation()
prs.slide_width  = Inches(13.3)
prs.slide_height = Inches(7.5)

# ── Colors ──────────────────────────────────────────────────────
def rgb(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

NAVY   = rgb("0B1F3A")
NAVYM  = rgb("1E3252")
NAVYL  = rgb("2D4A6B")
RED    = rgb("E4001A")
SLATE  = rgb("647A93")
WHITE  = rgb("FFFFFF")
BG     = rgb("F5F7FA")
BORDER = rgb("DDE3EB")
TEXTD  = rgb("0B1F3A")
TEXTM  = rgb("3A506B")
TEXTL  = rgb("8898AA")
CODEBG = rgb("EEF2F7")
GREENBG= rgb("EAF6F0")
GREEN  = rgb("1A7F4B")
AMBBG  = rgb("FEF3E2")
AMBER  = rgb("B45309")
REDBG  = rgb("FEE8EB")

BLANK = prs.slide_layouts[6]  # blank layout


# ── Shape helpers ────────────────────────────────────────────────
def i(v): return Inches(v)

def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_w=0):
    shape = slide.shapes.add_shape(1, i(x), i(y), i(w), i(h))  # 1 = rectangle
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color and line_w > 0:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    return shape

def add_rrect(slide, x, y, w, h, fill_color, line_color=None, line_w=0, radius=0.08):
    shape = slide.shapes.add_shape(5, i(x), i(y), i(w), i(h))  # 5 = rounded rectangle
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color and line_w > 0:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    # Set corner radius via XML (adj attribute)
    adj_val = int(radius * 50000)  # 0-50000 range in OOXML
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is not None:
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = etree.SubElement(prstGeom, qn('a:avLst'))
        else:
            avLst.clear()
        gd = etree.SubElement(avLst, qn('a:gd'))
        gd.set('name', 'adj')
        gd.set('fmla', f'val {adj_val}')
    return shape

def add_ellipse(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(9, i(x), i(y), i(w), i(h))  # 9 = oval
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape

def add_text(slide, x, y, w, h, text, font_size, color, bold=False,
             italic=False, align=PP_ALIGN.LEFT, font_face="Calibri",
             wrap=True, valign=None):
    txBox = slide.shapes.add_textbox(i(x), i(y), i(w), i(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_face
    if valign:
        from pptx.enum.text import MSO_ANCHOR
        tf.auto_size = None
        tf.vertical_anchor = valign
    return txBox

def add_hline(slide, x, y, w, color, lw=0.75):
    from pptx.util import Pt as PPt
    shape = slide.shapes.add_shape(20, i(x), i(y), i(w), i(0.001))  # line
    shape.fill.background()
    shape.line.color.rgb = color
    shape.line.width = PPt(lw)
    return shape


# ── Slide header ─────────────────────────────────────────────────
def add_header(slide, num, title, subtitle=None):
    add_rrect(slide, 0.5, 0.3, 0.38, 0.38, RED, radius=0.15)
    add_text(slide, 0.5, 0.3, 0.38, 0.38, str(num), 13, WHITE, bold=True, align=PP_ALIGN.CENTER, font_face="Calibri")
    add_text(slide, 1.05, 0.27, 11.2, 0.36, title, 22, NAVY, bold=True, font_face="Cambria")
    if subtitle:
        add_text(slide, 1.05, 0.63, 11.2, 0.24, subtitle, 10, TEXTL)
    add_hline(slide, 0.5, 0.93, 12.3, BORDER)

def add_slide_num(slide, n):
    add_text(slide, 12.3, 7.12, 0.8, 0.22, f"{n} / 6", 9, TEXTL, align=PP_ALIGN.RIGHT)

def add_tag(slide, x, y, label, t):
    colors = {
        "req":  (REDBG,  RED),
        "opt":  (CODEBG, SLATE),
        "cond": (AMBBG,  AMBER),
    }
    bg, fg = colors.get(t, (CODEBG, SLATE))
    add_rrect(slide, x, y, 0.85, 0.26, bg, bg, 0, radius=0.12)
    add_text(slide, x, y, 0.85, 0.26, label, 8.5, fg, bold=True, align=PP_ALIGN.CENTER)

def add_file_row(slide, y, icon, name, desc, tag_label, tag_type):
    rh = 0.72
    add_rrect(slide, 0.5, y, 12.3, rh, CODEBG, BORDER, 0.5, radius=0.06)
    add_text(slide, 0.62, y + 0.12, 0.46, 0.46, icon, 17, TEXTD, align=PP_ALIGN.CENTER)
    add_text(slide, 1.18, y + 0.06, 10.5, 0.28, name, 10.5, RED, bold=True, font_face="Courier New")
    add_text(slide, 1.18, y + 0.37, 10.5, 0.28, desc, 9, TEXTM)
    add_tag(slide, 12.0, y + (rh - 0.26) / 2, tag_label, tag_type)

def add_kw_pill(slide, x, y, kw, t):
    is_inc = t == "inc"
    bg   = GREENBG if is_inc else REDBG
    fg   = GREEN   if is_inc else RED
    w = max(0.95, len(kw) * 0.13 + 0.28)
    add_rrect(slide, x, y, w, 0.27, bg, bg, 0, radius=0.1)
    add_text(slide, x, y, w, 0.27, kw, 9, fg, bold=True, font_face="Courier New", align=PP_ALIGN.CENTER)
    return w


# ══════════════════════════════════════════════════════════════════
# SLIDE 1: Title (dark navy)
# ══════════════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(BLANK)
add_rect(slide1, 0, 0, 13.3, 7.5, NAVY)

# Glow circles (simulated with transparency using fill only — no true transparency in python-pptx easily)
# Use a slightly lighter navy for the circle effect
add_ellipse(slide1, 9.0, -1.8, 5.5, 5.5, rgb("1A2E4A"))
add_ellipse(slide1, 10.2, 4.2, 3.2, 3.2, rgb("151F30"))

# KT badge
add_rrect(slide1, 12.05, 0.35, 0.82, 0.82, RED, radius=0.15)
add_text(slide1, 12.05, 0.35, 0.82, 0.82, "KT", 18, WHITE, bold=True, align=PP_ALIGN.CENTER)

# Eyebrow
add_text(slide1, 0.7, 2.05, 8.0, 0.35, "사용자 안내서  ·  2026", 11, RED, bold=True)

# Title
add_text(slide1, 0.7, 2.45, 9.2, 1.0, "KT 정산 자동화", 46, WHITE, bold=True, font_face="Cambria")
add_text(slide1, 0.7, 3.45, 9.2, 0.9, "사용 안내", 46, WHITE, bold=True, font_face="Cambria")

# Subtitle
add_text(slide1, 0.7, 4.55, 9.8, 0.35, "정산자동화.exe 사용을 위한 파일 준비·실행·오류 대처 가이드", 13.5, rgb("8898AA"))

# Separator line
add_hline(slide1, 0.7, 5.1, 8.5, NAVYM)

# Pills
pills = [("정산자동화.exe", True), ("KT 정산 파일", False), ("플랫폼 파일", False), ("반품 매핑", False), ("오류 대처", False)]
px = 0.7
for label, primary in pills:
    pw = max(1.5, len(label) * 0.16 + 0.4)
    bg = RED if primary else NAVYM
    fg = WHITE if primary else rgb("A8BBCC")
    add_rrect(slide1, px, 5.4, pw, 0.36, bg, bg, 0, radius=0.18)
    add_text(slide1, px, 5.4, pw, 0.36, label, 9.5, fg, bold=True, align=PP_ALIGN.CENTER)
    px += pw + 0.14


# ══════════════════════════════════════════════════════════════════
# SLIDE 2: 필요 파일
# ══════════════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(BLANK)
add_header(slide2, 1, "폴더에 넣을 파일", "이번 정산 파일만 넣어두세요 — 이전 달 파일은 다른 폴더로 이동")

RH = 0.60  # 파일 행 높이 (슬라이드 내 7개 행 수용)

def add_file_row2(slide, y, icon, name, desc, tag_label, tag_type):
    add_rrect(slide, 0.5, y, 12.3, RH, CODEBG, BORDER, 0.5, radius=0.06)
    add_text(slide, 0.62, y + 0.08, 0.46, 0.42, icon, 15, TEXTD, align=PP_ALIGN.CENTER)
    add_text(slide, 1.18, y + 0.04, 10.5, 0.24, name, 10, RED, bold=True, font_face="Courier New")
    add_text(slide, 1.18, y + 0.3, 10.5, 0.26, desc, 8.5, TEXTM)
    add_tag(slide, 12.0, y + (RH - 0.26) / 2, tag_label, tag_type)

add_text(slide2, 0.5, 1.02, 6.0, 0.18, "필수 — 없으면 실행 불가", 8.5, TEXTL, bold=True)
add_file_row2(slide2, 1.23, "📄", "KT정산_202507.xlsx  (또는 .xls)",
    '통합구매시스템 다운로드 파일  |  파일명에 "KT"·"정산" 또는 "일일정산관리" 포함 필수', "필수", "req")
add_file_row2(slide2, 1.89, "📄", "통합플랫폼_202507.xlsx  (또는 .xls)",
    '통합플랫폼 다운로드 파일  |  파일명에 "플랫폼" 또는 "통합플랫폼" 포함 필수', "필수", "req")

add_text(slide2, 0.5, 2.63, 6.0, 0.18, "선택 — 있으면 자동으로 활용", 8.5, TEXTL, bold=True)
add_file_row2(slide2, 2.84, "📊", "mapping_master.xlsx",
    '관리회계·서비스카테고리 매핑, 기업규모 수동 등록  |  없으면 기업규모 "확인필요"로 표시', "선택", "opt")
add_file_row2(slide2, 3.5, "📋", "중견기업목록.xlsx",
    "중견기업 자동 판별용 목록  |  없으면 내장 대기업 목록만 사용", "선택", "opt")
add_file_row2(slide2, 4.16, "📋", "대기업목록.xlsx",
    "대기업 판별 커스텀 목록  |  없으면 내장 목록 자동 사용 (삼성전자·SK하이닉스 등 주요 대기업 포함)", "선택", "opt")
add_file_row2(slide2, 4.82, "📋", "과세확인목록.xlsx",
    "면세 의심 상품 중 과세 확정 품목 등록  |  있으면 과세구분 검토 단계에서 자동 제외", "선택", "opt")
add_file_row2(slide2, 5.48, "📋", "return_mapping.xlsx",
    "반품 자동매칭 실패 시 프로그램이 자동 생성  |  작성 후 폴더에 두고 재실행", "필요시", "cond")

add_slide_num(slide2, 2)


# ══════════════════════════════════════════════════════════════════
# SLIDE 3: 파일명 규칙
# ══════════════════════════════════════════════════════════════════
slide3 = prs.slides.add_slide(BLANK)
add_header(slide3, 2, "파일 자동 탐지 규칙", "파일명 키워드로 KT·플랫폼 파일을 구분합니다")

def draw_rule_card(slide, y, label, label_color, inc_kws, exc_kws):
    add_rrect(slide, 0.5, y, 12.3, 1.55, CODEBG, BORDER, 0.5, radius=0.08)
    # Left label area
    add_rrect(slide, 0.5, y, 1.6, 1.55, label_color, label_color, 0, radius=0.08)
    add_rect(slide, 1.1, y, 1.0, 1.55, label_color, label_color)
    add_text(slide, 0.52, y, 1.56, 1.55, label, 13, WHITE, bold=True, align=PP_ALIGN.CENTER)
    # Include
    add_text(slide, 2.3, y + 0.14, 0.65, 0.27, "포함:", 9.5, TEXTL, bold=True)
    kx = 3.0
    for kw in inc_kws:
        kx += add_kw_pill(slide, kx, y + 0.15, kw, "inc") + 0.15
    # Exclude
    add_text(slide, 2.3, y + 0.59, 0.65, 0.27, "제외:", 9.5, TEXTL, bold=True)
    ex = 3.0
    for kw in exc_kws:
        ex += add_kw_pill(slide, ex, y + 0.6, kw, "exc") + 0.12

draw_rule_card(slide3, 1.05, "KT\n정산 파일", RED,
    ["kt  +  정산", "kt  +  자료", "kt"],
    ["보안문서", "ktc", "플랫폼", "mapping", "목록"])

draw_rule_card(slide3, 2.75, "플랫폼\n파일", NAVYL,
    ["통합플랫폼", "일일정산관리", "플랫폼"],
    ["보안문서", "return", "mapping", "kt"])

add_rrect(slide3, 0.5, 4.45, 12.3, 0.52, AMBBG, AMBBG, 0, radius=0.06)
add_text(slide3, 0.65, 4.45, 12.0, 0.52,
    "⚠  KT 보안문서(암호화 파일)는 탐지 제외됩니다. 반드시 일반 xlsx/xls 파일을 사용하세요.  통합플랫폼 다운로드 .xls 파일도 그대로 사용 가능합니다.", 10.5, AMBER)
add_text(slide3, 0.5, 5.12, 12.3, 0.38,
    "💡  파일명에 날짜(202507 등)가 포함되어도 정상 탐지됩니다. 폴더 안에 이전 달 파일이 있으면 잘못 탐지될 수 있으니 이번 정산 파일만 남겨두세요.", 9.5, TEXTL)
add_slide_num(slide3, 3)


# ══════════════════════════════════════════════════════════════════
# SLIDE 4: 실행 방법
# ══════════════════════════════════════════════════════════════════
slide4 = prs.slides.add_slide(BLANK)
add_header(slide4, 3, "실행 방법", "6단계로 완료됩니다")

steps = [
    (1, "정산 파일 폴더 준비",    "이번 정산의 KT 파일·플랫폼 파일을 하나의 폴더에 모아둡니다. 이전 달 파일은 다른 폴더로 이동하세요."),
    (2, "정산자동화.exe 실행",    "정산자동화 폴더 안의 정산자동화.exe를 더블클릭하거나 실행.bat을 실행합니다."),
    (3, "작업 폴더 선택",         '"찾아보기" 버튼으로 1단계에서 준비한 폴더를 선택합니다.'),
    (4, "정산 기간 확인",         '날짜가 자동으로 채워집니다. 틀렸을 경우 직접 수정하거나 "자동 판단" 버튼을 누르세요. (YYYY-MM-DD)'),
    (5, "정산 실행 클릭",         '실행 버튼을 누르면 로그가 실시간으로 출력됩니다. 오류 없이 "완료"가 뜰 때까지 기다립니다.'),
    (6, "결과 파일 열기",         '"결과 파일 열기" 버튼이 활성화되면 클릭합니다. 결과 xlsx가 자동으로 열립니다.'),
]
for i_s, (n, title, desc) in enumerate(steps):
    col = 0 if i_s < 3 else 1
    row = i_s if i_s < 3 else i_s - 3
    x = 0.5 + col * 6.4
    y = 1.08 + row * 1.9
    cw, ch = 6.0, 1.75
    add_rrect(slide4, x, y, cw, ch, BG, BORDER, 0.5, radius=0.08)
    add_ellipse(slide4, x + 0.2, y + 0.14, 0.52, 0.52, RED)
    add_text(slide4, x + 0.2, y + 0.14, 0.52, 0.52, str(n), 14, WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide4, x + 0.86, y + 0.1, cw - 1.05, 0.38, title, 13.5, NAVY, bold=True, font_face="Cambria")
    add_text(slide4, x + 0.86, y + 0.53, cw - 1.05, 1.1, desc, 10, TEXTM)
add_slide_num(slide4, 4)


# ══════════════════════════════════════════════════════════════════
# SLIDE 5: 반품 매핑
# ══════════════════════════════════════════════════════════════════
slide5 = prs.slides.add_slide(BLANK)
add_header(slide5, 4, "반품 매핑 — 자동매칭 실패 시", "반품 건이 있고 자동매칭이 안 될 때만 해당됩니다")

# Left card
add_rrect(slide5, 0.5, 1.05, 5.9, 5.4, BG, BORDER, 0.5, radius=0.1)
add_text(slide5, 0.72, 1.16, 5.4, 0.38, "🔄  처리 순서", 13, NAVY, bold=True, font_face="Cambria")

return_steps = [
    "프로그램이 반품 자동매칭을 시도합니다.",
    "실패하면 return_mapping.xlsx가 자동 생성되며 안내 메시지가 표시됩니다.",
    "파일을 열어 빨간 별(★) 열을 채웁니다.",
    "저장 후 프로그램을 재실행합니다. (폴더 선택은 그대로)",
]
for idx, s in enumerate(return_steps):
    sy = 1.75 + idx * 1.12
    add_ellipse(slide5, 0.74, sy, 0.38, 0.38, NAVYM)
    add_text(slide5, 0.74, sy, 0.38, 0.38, str(idx+1), 11, WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide5, 1.24, sy - 0.02, 4.95, 0.46, s, 10.5, TEXTM)

# Right card
add_rrect(slide5, 6.9, 1.05, 5.9, 5.4, BG, BORDER, 0.5, radius=0.1)
add_text(slide5, 7.12, 1.16, 5.4, 0.38, "📝  작성할 열", 13, NAVY, bold=True, font_face="Cambria")

# Table header
add_rect(slide5, 7.05, 1.68, 5.65, 0.36, NAVYM)
add_text(slide5, 7.1, 1.68, 2.1, 0.36, "열 이름", 9, rgb("C8D8E8"), bold=True)
add_text(slide5, 9.25, 1.68, 3.4, 0.36, "입력 내용", 9, rgb("C8D8E8"), bold=True)

# Row 1
add_rect(slide5, 7.05, 2.04, 5.65, 1.1, CODEBG)
add_hline(slide5, 7.05, 3.14, 5.65, BORDER, 0.3)
add_text(slide5, 7.1, 2.08, 2.1, 0.42, "★플랫폼반품주문번호", 8.5, RED, bold=True, font_face="Courier New")
add_text(slide5, 9.25, 2.06, 3.35, 0.36, "통합플랫폼에서 해당 반품의 주문번호", 9.5, TEXTM)
add_text(slide5, 9.25, 2.45, 3.35, 0.28, "예: 4502061937", 8.5, TEXTL, italic=True)

# Row 2
add_rect(slide5, 7.05, 3.14, 5.65, 1.1, WHITE)
add_text(slide5, 7.1, 3.18, 2.1, 0.42, "★플랫폼품목번호", 8.5, RED, bold=True, font_face="Courier New")
add_text(slide5, 9.25, 3.16, 3.35, 0.36, "해당 반품 건의 플랫폼 품목번호", 9.5, TEXTM)
add_text(slide5, 9.25, 3.55, 3.35, 0.28, "예: 100023456", 8.5, TEXTL, italic=True)

# Note
add_rrect(slide5, 7.05, 4.45, 5.65, 0.52, CODEBG, BORDER, 0.5, radius=0.06)
add_text(slide5, 7.15, 4.45, 5.45, 0.52, "나머지 열(KT반품주문번호 등)은 자동 입력되어 있으니 수정하지 마세요.", 9.5, TEXTL)

add_slide_num(slide5, 5)


# ══════════════════════════════════════════════════════════════════
# SLIDE 6: 오류 대처
# ══════════════════════════════════════════════════════════════════
slide6 = prs.slides.add_slide(BLANK)
add_header(slide6, 5, "오류 대처", "로그 창 빨간 텍스트로 원인이 표시됩니다")

errors = [
    ("🔴", '"KT 파일을 찾을 수 없습니다"',
     '파일명에 "KT"·"정산" 또는 "일일정산관리"가 포함되어 있는지 확인하세요. "보안문서", "KTC"가 포함된 파일명은 탐지에서 제외됩니다.'),
    ("🔴", '"플랫폼 파일을 찾을 수 없습니다"',
     '파일명에 "플랫폼" 또는 "통합플랫폼"이 포함되어 있는지 확인하세요. .xls 파일도 그대로 사용 가능합니다.'),
    ("🟡", '"KT 보안문서(암호화)이거나 손상된 파일"',
     "KT 보안 적용 파일은 읽을 수 없습니다. 보안 해제된 일반 xlsx/xls 파일로 교체하세요."),
    ("🟡", "반품 매핑 템플릿 생성 안내 메시지",
     "오류가 아닙니다. return_mapping.xlsx를 작성 후 재실행하세요. (슬라이드 5 참조)"),
    ("🔵", '기업규모 "확인필요" 표시',
     "mapping_master.xlsx의 기업규모 시트에 해당 협력사를 등록하거나, 중견기업목록.xlsx에 해당 기업이 누락된 경우 목록을 업데이트하세요."),
]
for idx, (icon, title, fix) in enumerate(errors):
    y = 1.08 + idx * 1.15
    add_rrect(slide6, 0.5, y, 12.3, 1.05, CODEBG, BORDER, 0.5, radius=0.07)
    add_text(slide6, 0.62, y + 0.12, 0.46, 0.46, icon, 18, TEXTD, align=PP_ALIGN.CENTER)
    add_text(slide6, 1.22, y + 0.07, 11.4, 0.34, title, 12, NAVY, bold=True, font_face="Cambria")
    add_text(slide6, 1.22, y + 0.45, 11.4, 0.55, "→  " + fix, 9.5, TEXTM)

add_slide_num(slide6, 6)


# ══════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════
prs.save(OUT)
print(f"저장 완료: {OUT}")
