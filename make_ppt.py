"""KT 정산 자동화 PoC 발표자료 생성 스크립트."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── 색상 ───────────────────────────────────────────────────────
CB   = RGBColor(0x1E, 0x70, 0xC0)   # 주 파란색
CDB  = RGBColor(0x15, 0x55, 0x9A)   # 진한 파란색
CLB  = RGBColor(0xD6, 0xE8, 0xF8)   # 연한 파란색
CW   = RGBColor(0xFF, 0xFF, 0xFF)   # 흰색
CBK  = RGBColor(0x20, 0x20, 0x20)   # 거의 검정
CGY  = RGBColor(0x88, 0x88, 0x88)   # 회색
CGB  = RGBColor(0xF2, 0xF2, 0xF2)   # 연회색 배경
CBD  = RGBColor(0xBF, 0xBF, 0xBF)   # 테두리 회색
CGR  = RGBColor(0xC6, 0xEF, 0xCE)   # 연초록
CGF  = RGBColor(0x00, 0x88, 0x30)   # 진초록 글자
CYB  = RGBColor(0xFF, 0xEB, 0x9C)   # 연노랑
CYF  = RGBColor(0x9C, 0x65, 0x00)   # 노랑 글자
CRB  = RGBColor(0xFF, 0xC7, 0xCE)   # 연빨강
CRF  = RGBColor(0x9C, 0x00, 0x06)   # 빨강 글자
FONT = "맑은 고딕"
MRCT = 1  # MSO_AUTO_SHAPE_TYPE.RECTANGLE

def blank():
    return prs.slides.add_slide(prs.slide_layouts[6])

def R(slide, l, t, w, h, fill, brd=None, bw=1.5):
    """색칠된 사각형."""
    s = slide.shapes.add_shape(MRCT, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if brd: s.line.color.rgb = brd; s.line.width = Pt(bw)
    else:   s.line.fill.background()
    return s

def RT(slide, text, l, t, w, h, fill, fg=CW, sz=12, bold=False,
       brd=None, bw=1.5, align=PP_ALIGN.CENTER, wrap=True):
    """텍스트 포함 사각형."""
    s = R(slide, l, t, w, h, fill, brd, bw)
    tf = s.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(sz)
    r.font.bold = bold; r.font.color.rgb = fg
    return s

def T(slide, text, l, t, w, h, sz=12, bold=False, fg=CBK,
      align=PP_ALIGN.LEFT, wrap=True, italic=False):
    """텍스트 박스 (배경 없음)."""
    s = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = s.text_frame; tf.word_wrap = wrap
    s.fill.background(); s.line.fill.background()
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(sz)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = fg
    return s

def TML(slide, lines, l, t, w, h, sz=12, fg=CBK, wrap=True):
    """여러 줄 텍스트 박스 (리스트 입력)."""
    s = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = s.text_frame; tf.word_wrap = wrap
    s.fill.background(); s.line.fill.background()
    for i, (line, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = line
        r.font.name = FONT; r.font.size = Pt(sz)
        r.font.bold = bold; r.font.color.rgb = color if color else fg
    return s

def hdr(slide, title):
    R(slide, 0, 0, 13.33, 0.88, CB)
    T(slide, title, 0.4, 0.12, 12, 0.66, sz=22, bold=True, fg=CW)

# ════════════════════════════════════════════════════════════════
# 슬라이드 1: 타이틀
# ════════════════════════════════════════════════════════════════
slide = blank()
R(slide, 0, 0, 13.33, 7.5, CB)
R(slide, 0.8, 1.6, 11.73, 4.3, CW)
R(slide, 0.8, 1.6, 11.73, 0.08, CDB)  # 상단 진한 줄

T(slide, "KT 정산 자동화 PoC",
  0.8, 2.0, 11.73, 1.4, sz=42, bold=True, fg=CDB, align=PP_ALIGN.CENTER)
T(slide, "정산 업무 자동화  |  과세구분 자동 검증 시스템",
  0.8, 3.5, 11.73, 0.7, sz=18, fg=CB, align=PP_ALIGN.CENTER)
T(slide, "KTC  ·  2026",
  0.8, 4.5, 11.73, 0.5, sz=12, fg=CGY, align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════
# 슬라이드 2: Before / After
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "KT 정산 자동화 개요")
R(slide, 6.55, 0.88, 0.04, 6.62, CBd := CBD)  # 중앙 구분선

# Before 헤더
RT(slide, "Before : 기존 방식", 0.5, 0.95, 3.8, 0.48,
   RGBColor(0x60,0x60,0x60), sz=13, bold=True)
before = [
    ("① 수동 데이터 가공",
     "KT 정산 파일과 통합플랫폼 파일을\n수동으로 비교·대조 후 정리"),
    ("② 담당자별 수동 매핑",
     "서비스카테고리별 관리회계·담당자를\n건별로 직접 입력 (수십~수백 건)"),
    ("③ 파일명 수동 정리",
     "날짜·차수·형식을 매번 직접 수정\n오기입 및 버전 혼선 발생"),
]
for i, (t1, t2) in enumerate(before):
    y = 1.58 + i * 1.66
    R(slide, 0.5, y, 5.7, 1.4, CGB, CBd, 0.8)
    T(slide, t1, 0.65, y+0.1, 5.4, 0.42, sz=12, bold=True, fg=CDB)
    T(slide, t2, 0.65, y+0.5, 5.4, 0.8,  sz=10, fg=CBK, wrap=True)

# After 헤더
RT(slide, "After : 자동화", 7.1, 0.95, 3.8, 0.48, CB, sz=13, bold=True)
after = [
    ("① 파일 3개 선택",
     "정산 데이터, 통합플랫폼 데이터,\n매핑 파일 경로만 지정"),
    ("② 버튼 1회 클릭",
     "과세구분 검증 → 인보이스 생성\n→ 담당자 매핑 → 파일명 자동 지정"),
    ("③ 결과 파일 자동 저장",
     "2026년 5월 2차 정산_KTC_20260602.xlsx\n형태로 자동 저장"),
]
for i, (t1, t2) in enumerate(after):
    y = 1.58 + i * 1.66
    R(slide, 7.1, y, 5.8, 1.4, CLB, CB, 0.8)
    T(slide, t1, 7.25, y+0.1, 5.4, 0.42, sz=12, bold=True, fg=CDB)
    T(slide, t2, 7.25, y+0.5, 5.4, 0.8,  sz=10, fg=CBK, wrap=True)

# ════════════════════════════════════════════════════════════════
# 슬라이드 3: 시스템 플로우
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "시스템 처리 흐름")

# 입력 파일 3개
inputs = ["정산 데이터", "통합플랫폼\n데이터", "매핑 파일"]
for i, lbl in enumerate(inputs):
    y = 1.4 + i * 1.62
    R(slide, 0.5, y, 3.3, 1.2, CLB, CB, 1.5)
    T(slide, lbl, 0.5, y+0.2, 3.3, 0.8, sz=14, bold=True,
      fg=CDB, align=PP_ALIGN.CENTER, wrap=True)
    # 화살표
    T(slide, "->", 3.85, y+0.35, 0.9, 0.4, sz=16, bold=True,
      fg=CB, align=PP_ALIGN.CENTER)

# 처리 박스
RT(slide, "자동 처리 엔진",
   4.8, 2.1, 3.6, 3.3, CB, fg=CW, sz=17, bold=True, brd=CDB, bw=2.0)
T(slide, "(Python)",
  4.8, 3.1, 3.6, 0.5, sz=10, fg=CLB, align=PP_ALIGN.CENTER)
T(slide, "과세구분 검증\n담당자/관리회계 매핑\n칼럼 정렬 및 파일 저장",
  4.9, 3.5, 3.4, 1.6, sz=10, fg=CW, align=PP_ALIGN.CENTER, wrap=True)

# 화살표
T(slide, "->", 8.45, 3.55, 0.9, 0.4, sz=16, bold=True,
  fg=CB, align=PP_ALIGN.CENTER)

# 출력 박스
R(slide, 9.4, 1.5, 3.5, 5.0, CLB, CB, 1.5)
RT(slide, "정산 완료 엑셀", 9.4, 1.5, 3.5, 0.58, CB, sz=13, bold=True)
sheets = [
    "⓪ 과세구분검증",
    "① 인보이스",
    "② 어음발행 대상",
    "③ 세금계산서 현황",
    "④ 과세구분 현황",
]
for i, s in enumerate(sheets):
    y = 2.2 + i * 0.78
    R(slide, 9.55, y, 3.2, 0.6, CW, CBd, 0.6)
    T(slide, s, 9.65, y+0.1, 3.0, 0.4, sz=10, bold=(i==0), fg=CDB)

# ════════════════════════════════════════════════════════════════
# 슬라이드 4: GUI 화면
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "KT 정산 자동화 – 실행 화면")

# GUI 외곽
R(slide, 1.0, 1.05, 11.33, 6.0, CGB, CBd, 1.0)
# GUI 상단 바
RT(slide, "|  KT 정산 자동화", 1.0, 1.05, 11.33, 0.48, CB, sz=13, bold=True,
   align=PP_ALIGN.LEFT)

# 컬럼 레이블
T(slide, "파일 선택",    1.2, 1.7, 3.8, 0.4, sz=13, bold=True, fg=CBK)
T(slide, "구성",         5.5, 1.7, 3.2, 0.4, sz=13, bold=True, fg=CBK)
T(slide, "저장 파일명",  9.2, 1.7, 3.5, 0.4, sz=13, bold=True, fg=CBK)

# 파일 선택 버튼 3개
files = ["정산 데이터 선택", "통합플랫폼 데이터 선택", "매핑 파일 선택"]
for i, f in enumerate(files):
    y = 2.2 + i * 1.1
    R(slide, 1.2, y, 3.9, 0.75, CW, CBd, 1.0)
    T(slide, "[  ]  " + f, 1.35, y+0.2, 3.6, 0.35, sz=11, fg=CBK)

# 구성
T(slide, "정산기간", 5.5, 2.22, 3.0, 0.32, sz=11, fg=CGY)
R(slide, 5.5, 2.55, 3.2, 0.52, CW, CBd)
T(slide, "2026년 5월  ▾", 5.65, 2.63, 2.9, 0.36, sz=11, fg=CBK)
T(slide, "차수 선택", 5.5, 3.25, 3.0, 0.32, sz=11, fg=CGY)
R(slide, 5.5, 3.58, 3.2, 0.52, CW, CBd)
T(slide, "2차  ▾", 5.65, 3.66, 2.9, 0.36, sz=11, fg=CBK)

# 저장 파일명 미리보기
R(slide, 9.2, 2.2, 3.7, 1.7, CW, CBd)
T(slide, "2026년 5월 2차 정산\n_KTC_20260602.xlsx",
  9.35, 2.5, 3.4, 1.2, sz=12, fg=CDB, wrap=True)

# 실행 버튼
RT(slide, "정산 실행", 4.2, 5.5, 5.0, 0.82, CB, sz=17, bold=True)

# ════════════════════════════════════════════════════════════════
# 슬라이드 5: 과세구분 검증 – 3 카드
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "과세구분 자동 검증 – 3단계 탐지 로직")

card_data = [
    ("① 키워드 탐지",
     [("[비과세]", True, CB),
      ("상품권, 기프티콘, 교환권,\n모바일쿠폰, 선불카드\n", False, None),
      ("[면세]", True, CB),
      ("우유, 도서, 생리대, 연어,\n버섯, 소고기, 한우, 딸기...\n(부가세법 제26조 기준)\n", False, None),
      ("[이력 기반 패턴]", True, CB),
      ("한우세트, 굴비, 멸균우유,\n락토프리, 과일세트 등", False, None)]),
    ("② 혼용 코드 탐지",
     [("동일 상품코드에\n과세 + 면세가 혼재하는\n경우를 자동 플래그\n", False, None),
      ("[ 예시 ]", True, CB),
      ("상품코드 45020367504\n  → 3건: 과세\n  → 2건: 면세\n  → 마스터 오류 의심", False, None)]),
    ("③ 화이트리스트 제외",
     [("과세확인목록.xlsx 에\n등록된 상품코드는\n검증 제외 처리\n", False, None),
      ("[ 효과 ]", True, CB),
      ("담당자가 확인 완료한\n상품의 반복 탐지 방지\n\n파일만 업데이트하면\n코드 수정 없이 적용", False, None)]),
]

for i, (title_, lines) in enumerate(card_data):
    x = 0.45 + i * 4.27
    R(slide, x, 1.02, 4.0, 6.2, CW, CBd, 1.0)
    RT(slide, title_, x, 1.02, 4.0, 0.62, CB, sz=14, bold=True)
    TML(slide, lines, x+0.2, 1.75, 3.65, 5.2, sz=11)

# ════════════════════════════════════════════════════════════════
# 슬라이드 6: 검증 결과 시트 목업
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "과세구분검증 시트 – 출력 예시")

T(slide, "⓪ 과세구분검증   (이상 항목만 표시 — 담당자 검토 후 '담당자의견' 란 기재)",
  0.4, 0.95, 12.5, 0.44, sz=11, fg=CDB, bold=True)

# 테이블 설정
COLS  = ["상품코드", "상품명", "담당자", "현재\n과세구분", "검증결과",
         "탐지키워드", "사유", "담당자의견", "비고"]
CWIDTHS = [1.25, 2.0, 1.0, 1.1, 1.45, 1.1, 2.1, 1.55, 0.95]
SX, SY, RH = 0.4, 1.5, 0.6

# 헤더
x = SX
for col, cw in zip(COLS, CWIDTHS):
    RT(slide, col, x, SY, cw, RH+0.1, CB, sz=9, bold=True, brd=CW, bw=0.5)
    x += cw

# 데이터 행
rows = [
    (["A10123", "멸균우유 200mL\n(24개입)", "이영희", "과세", "면세 검토 필요",
      "멸균우유", "상품명에 면세 키워드 포함\n→ 면세 대상일 수 있음", "", ""],
     CGB, CYB, CYF),
    (["B20234", "소고기 등심세트\n1kg", "박민준", "면세", "과세구분 혼재",
      "", "동일 코드에 과세/면세 혼재\n→ 마스터 오류 가능성", "", ""],
     CW, CRB, CRF),
    (["C30345", "도서 구매\n(해외도서)", "최지원", "과세", "면세 검토 필요",
      "도서", "상품명에 면세 키워드 포함\n→ 제26조 제8호 해당 가능", "", ""],
     CGB, CYB, CYF),
    (["D40456", "생리대 세트\n30개입", "김철수", "과세", "면세 검토 필요",
      "생리대", "여성 위생용품 → 면세\n(부가세법 제26조 제4호)", "", ""],
     CW, CYB, CYF),
]

for ri, (row_data, row_bg, res_bg, res_fg) in enumerate(rows):
    y = SY + RH + 0.1 + ri * RH
    x = SX
    for j, (val, cw) in enumerate(zip(row_data, CWIDTHS)):
        bg = res_bg if j == 4 else row_bg
        fg = res_fg if j == 4 else CBK
        bld = (j == 4)
        al = PP_ALIGN.CENTER if j in (0, 2, 3, 4, 5) else PP_ALIGN.LEFT
        RT(slide, val, x, y, cw, RH, bg, fg, sz=8, bold=bld,
           brd=CBd, bw=0.5, align=al, wrap=True)
        x += cw

# 범례
T(slide, "색상 범례:  ", 0.4, 6.72, 1.4, 0.4, sz=9, fg=CGY)
RT(slide, "면세 검토 필요", 1.7, 6.72, 1.7, 0.38, CYB, CYF, sz=9, brd=CBd, bw=0.5)
RT(slide, "과세의심/혼재", 3.5, 6.72, 1.7, 0.38, CRB, CRF, sz=9, brd=CBd, bw=0.5)
RT(slide, "정상",          5.3, 6.72, 0.9, 0.38, CGR, CGF, sz=9, brd=CBd, bw=0.5)

# ════════════════════════════════════════════════════════════════
# 슬라이드 7: 기대효과 & 확장
# ════════════════════════════════════════════════════════════════
slide = blank()
hdr(slide, "기대효과 및 향후 확장 가능성")
R(slide, 6.55, 0.88, 0.04, 6.12, CBd)  # 중앙 구분선

# 왼쪽: 정량 효과
T(slide, "정량적 기대효과", 0.5, 1.0, 5.7, 0.5, sz=16, bold=True, fg=CDB,
  align=PP_ALIGN.CENTER)

effects = [
    ("80%",    "월 정산 작업 시간 단축",    "수작업 → 버튼 1회 클릭으로 대체"),
    ("자동화", "과세구분 오류 탐지",         "키워드·혼용·화이트리스트 3단계 검증"),
    ("표준화", "인수인계 부담 감소",         "파일명·컬럼 구조 완전 표준화"),
]
for i, (big, mid, small) in enumerate(effects):
    y = 1.65 + i * 1.6
    R(slide, 0.5, y, 5.7, 1.38, CLB, CB, 0.8)
    T(slide, big,   0.65, y+0.08, 1.6, 0.7, sz=30, bold=True, fg=CB)
    T(slide, mid,   2.3,  y+0.05, 3.7, 0.45, sz=13, bold=True, fg=CDB)
    T(slide, small, 2.3,  y+0.52, 3.7, 0.6,  sz=10, fg=CBK, wrap=True)

# 오른쪽: 확장
T(slide, "향후 확장 가능성", 7.0, 1.0, 5.9, 0.5, sz=16, bold=True, fg=CDB,
  align=PP_ALIGN.CENTER)

expansions = [
    ("화이트리스트 자동 갱신",   "담당자 확인 완료 시 코드 자동 등록"),
    ("정산 이상 패턴 알림",       "금액 급변·혼재 발생 시 자동 통보"),
    ("타 사업부 적용",             "동일 프레임워크로 타팀 정산에 확장"),
    ("서버 배포",                  "로컬 실행 → 사내 서버 자동화로 확장"),
]
for i, (t1, t2) in enumerate(expansions):
    y = 1.65 + i * 1.22
    R(slide, 6.9, y, 6.0, 1.0, CW, CBd, 0.8)
    R(slide, 6.9, y, 0.18, 1.0, CB)  # 왼쪽 포인트 바
    T(slide, t1, 7.15, y+0.08, 5.6, 0.4, sz=12, bold=True, fg=CDB)
    T(slide, t2, 7.15, y+0.5,  5.6, 0.42, sz=10, fg=CBK)

# 하단 강조 문구
R(slide, 0.4, 6.65, 12.53, 0.62, CB)
T(slide, '"반복 수작업을 코드로 대체하고, 담당자는 판단이 필요한 업무에만 집중"',
  0.4, 6.7, 12.53, 0.55, sz=14, bold=True, fg=CW, align=PP_ALIGN.CENTER)

# ── 저장 ──────────────────────────────────────────────────────
OUT = r"C:\Users\82242282\Documents\KT정산자동화_PoC.pptx"
prs.save(OUT)
print(f"저장 완료: {OUT}")
