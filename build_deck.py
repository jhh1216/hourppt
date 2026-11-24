"""골든아워 발표 자료 생성기. `python3 build_deck.py` → golden-hour.pptx"""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree

FONT = "맑은 고딕"

INK = RGBColor(0x11, 0x18, 0x27)
SUB = RGBColor(0x4B, 0x55, 0x63)
MUTED = RGBColor(0x9C, 0xA3, 0xAF)
LINE = RGBColor(0xD1, 0xD5, 0xDB)
BG = RGBColor(0xF8, 0xFA, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xE1, 0x1D, 0x48)
RED_SOFT = RGBColor(0xFF, 0xE4, 0xE6)
NAVY = RGBColor(0x0F, 0x17, 0x2A)
NAVY_2 = RGBColor(0x1E, 0x29, 0x3B)
BLUE_SOFT = RGBColor(0xE0, 0xE7, 0xFF)
BLUE = RGBColor(0x37, 0x30, 0xA3)
GREEN_SOFT = RGBColor(0xDC, 0xFC, 0xE7)
GREEN = RGBColor(0x16, 0x65, 0x34)
AMBER_SOFT = RGBColor(0xFE, 0xF3, 0xC7)
AMBER = RGBColor(0x92, 0x40, 0x0E)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
TOTAL_HOLDER = []


def style_run(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", FONT)


def fill_text(tf, lines, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT, spacing=1.15, gap=0):
    """lines: str | list of str | list of (text, {size,bold,color}) | list of list-of-runs"""
    if isinstance(lines, str):
        lines = [lines]
    tf.word_wrap = True
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = spacing
        if gap:
            p.space_after = Pt(gap)
        runs = line
        if isinstance(line, str):
            runs = [(line, {})]
        elif isinstance(line, tuple):
            runs = [line]
        for text, opt in runs:
            r = p.add_run()
            r.text = text
            style_run(r, opt.get("size", size), opt.get("bold", bold), opt.get("color", color))


def text(slide, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, **kw):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = anchor
    fill_text(tf, lines, **kw)
    return tb


def box(slide, x, y, w, h, lines=None, fill=WHITE, border=LINE, radius=True, anchor=MSO_ANCHOR.MIDDLE,
        dashed=False, border_w=1.0, pad=0.12, **kw):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        shp.adjustments[0] = 0.08
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if border is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = border
        shp.line.width = Pt(border_w)
        if dashed:
            shp.line.dash_style = MSO_LINE.DASH
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Inches(pad)
    tf.margin_top = tf.margin_bottom = Inches(0.06)
    tf.vertical_anchor = anchor
    if lines is not None:
        kw.setdefault("align", PP_ALIGN.CENTER)
        fill_text(tf, lines, **kw)
    return shp


def arrow(slide, x1, y1, x2, y2, color=SUB, width=1.5, dashed=False, head=True):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    if dashed:
        c.line.dash_style = MSO_LINE.DASH
    if head:
        ln = c.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    return c


def pill(slide, x, y, label, fill=RED_SOFT, color=RED, w=None, size=11):
    w = w or (0.22 + 0.15 * len(label))
    return box(slide, x, y, w, 0.32, label, fill=fill, border=None, size=size, bold=True, color=color, pad=0.06)


def table(slide, x, y, w, rows, col_w, row_h=0.42, size=12, header_fill=NAVY_2, zebra=True, first_bold=True):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w),
                                   Inches(row_h * len(rows)))
    tbl = shape.table
    tbl_pr = shape._element.graphic.graphicData.tbl.tblPr
    style = tbl_pr.find(qn("a:tableStyleId"))
    if style is not None:
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r, row in enumerate(rows):
        tbl.rows[r].height = Inches(row_h)
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = header_fill
            elif zebra and r % 2 == 0:
                cell.fill.fore_color.rgb = RGBColor(0xF1, 0xF5, 0xF9)
            else:
                cell.fill.fore_color.rgb = WHITE
            cell.text_frame.paragraphs[0].text = ""
            if r == 0:
                fill_text(cell.text_frame, val, size=size, bold=True, color=WHITE)
            else:
                fill_text(cell.text_frame, val, size=size, bold=(first_bold and c == 0), color=INK)
    return shape


def new_slide(section=None, title=None, subtitle=None, notes=None, dark=False):
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = NAVY if dark else BG
    if not dark:
        s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.12), prs.slide_height).fill.solid()
        bar = s.shapes[-1]
        bar.fill.fore_color.rgb = RED
        bar.line.fill.background()
    if section:
        text(s, 0.6, 0.42, 8, 0.3, section, size=12, bold=True, color=RED)
    if title:
        text(s, 0.6, 0.7, 12.1, 0.6, title, size=28, bold=True, color=WHITE if dark else INK)
    if subtitle:
        text(s, 0.6, 1.33, 12.1, 0.4, subtitle, size=14, color=MUTED if dark else SUB)
    if not dark:
        idx = len(prs.slides)
        text(s, 0.6, 7.05, 6, 0.25, "골든아워 · 팀 코드블루", size=10, color=MUTED)
        num = text(s, 11.7, 7.05, 1.0, 0.25, str(idx), size=10, color=MUTED, align=PP_ALIGN.RIGHT)
        TOTAL_HOLDER.append(num)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def card(slide, x, y, w, h, tag, title, body, tag_fill=RED_SOFT, tag_color=RED):
    box(slide, x, y, w, h, fill=WHITE, border=LINE)
    if tag:
        pill(slide, x + 0.2, y + 0.2, tag, fill=tag_fill, color=tag_color)
    ty = y + (0.62 if tag else 0.22)
    text(slide, x + 0.2, ty, w - 0.4, 0.4, title, size=16, bold=True)
    text(slide, x + 0.2, ty + 0.45, w - 0.4, h - (ty - y) - 0.55, body, size=12, color=SUB, spacing=1.25)


def entity(slide, x, y, w, name, fields, note=None, dashed=False):
    head_h, row_h = 0.36, 0.235
    h = head_h + row_h * len(fields) + 0.1
    border = AMBER if dashed else NAVY_2
    box(slide, x, y, w, h, fill=WHITE, border=border, radius=False, dashed=dashed, border_w=1.25)
    hdr = box(slide, x, y, w, head_h, name, fill=AMBER if dashed else NAVY_2, border=None, radius=False,
              size=12, bold=True, color=WHITE)
    for i, (key, col) in enumerate(fields):
        fy = y + head_h + 0.05 + i * row_h
        if key:
            text(slide, x + 0.1, fy, 0.45, row_h, key, size=9, bold=True,
                 color=RED if key == "PK" else BLUE, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, x + 0.55, fy, w - 0.6, row_h, col, size=10, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    if note:
        pill(slide, x + w - 1.2, y - 0.36, note, fill=AMBER_SOFT, color=AMBER, w=1.2, size=9)
    return (x, y, w, h)


# ── 1. 표지 ───────────────────────────────────────────────────────
s = new_slide(dark=True, notes=(
    "안녕하세요, 팀 코드블루입니다. 저희가 만든 서비스는 골든아워입니다. "
    "응급 상황에서 가장 가까운 병원이 아니라, 지금 받아줄 수 있고 가장 빨리 도착할 수 있는 병원을 추천하는 웹 서비스입니다."))
box(s, 0.9, 1.2, 0.18, 1.6, fill=RED, border=None, radius=False)
text(s, 1.35, 1.15, 10, 0.4, "TEAM CODE BLUE", size=14, bold=True, color=RED)
text(s, 1.35, 1.6, 11, 1.2, "골든아워", size=66, bold=True, color=WHITE)
text(s, 1.35, 3.05, 11, 0.6, "가장 가까운 병원이 아니라, 지금 가장 빨리 진료받을 수 있는 병원", size=24, color=WHITE)
text(s, 1.35, 3.75, 11, 0.5, "실시간 응급실 병상 × 이동 시간 기반 응급실 추천 웹 서비스", size=16, color=MUTED)
for i, t in enumerate(["React 19", "Spring Boot 4.1", "PostgreSQL 16", "국립중앙의료원 API", "카카오 지도"]):
    box(s, 1.35 + i * 2.15, 5.6, 2.0, 0.45, t, fill=NAVY_2, border=None, size=12, color=WHITE, bold=True)

# ── 2. 문제 ───────────────────────────────────────────────────────
s = new_slide("01 · 문제", "가까운 병원이 가장 빠른 병원은 아니다",
              "응급 상황에서 일반 지도 앱으로 응급실을 찾을 때 생기는 일",
              notes=(
                  "응급 상황에서 보통 지도 앱에 응급실을 검색합니다. 그러면 가까운 순서로 핀이 뜹니다. "
                  "그런데 그 병원에 지금 병상이 있는지는 하나씩 눌러보고 전화해 봐야 압니다. "
                  "막상 도착했는데 병상이 없으면 다른 병원으로 다시 가야 합니다. "
                  "그리고 직선으로 가까워도 길이 막히면 더 먼 병원이 먼저 도착할 수도 있습니다."))
problems = [
    ("01", "병상이 있는지 모른다", "지도에는 위치만 나온다.\n지금 받아줄 수 있는지는 하나씩 눌러보고\n전화해 봐야 안다."),
    ("02", "도착해서 다시 옮긴다", "병상이 없거나 장비가 없으면\n다른 병원으로 전원된다.\n이동 시간이 두 번 든다."),
    ("03", "직선거리 ≠ 도착 시간", "직선으로 가까워도 도로가 막히면\n더 먼 병원이 먼저 도착할 수 있다."),
]
for i, (n, t, b) in enumerate(problems):
    x = 0.6 + i * 4.1
    box(s, x, 2.1, 3.85, 3.6, fill=WHITE, border=LINE)
    text(s, x + 0.35, 2.4, 2, 0.8, n, size=40, bold=True, color=RED)
    text(s, x + 0.35, 3.35, 3.2, 0.5, t, size=20, bold=True)
    text(s, x + 0.35, 4.0, 3.2, 1.6, b, size=14, color=SUB, spacing=1.35)
box(s, 0.6, 6.0, 12.1, 0.7, [[("응급 상황의 사용자는 ", {}), ("비교할 시간이 없다.", {"bold": True, "color": RED}),
                                ("  한 화면에서 바로 판단할 수 있어야 한다.", {})]],
    fill=RED_SOFT, border=None, size=16)

# ── 3. 해결 ───────────────────────────────────────────────────────
s = new_slide("02 · 해결", "골든아워: 실시간 병상과 이동 시간을 함께 본다",
              notes=(
                  "골든아워는 세 가지를 한 번에 봅니다. 첫째, 국립중앙의료원 API로 60초마다 응급실 병상을 받아옵니다. "
                  "둘째, 이동 시간을 계산해서 병상과 합쳐 점수를 매깁니다. "
                  "셋째, 증상을 말하거나 입력하면 필요한 장비가 있는 병원으로 걸러주고, 위험한 증상이면 119를 먼저 안내합니다."))
text(s, 0.6, 1.5, 12, 0.6, [[("지도 앱은 ", {"color": SUB}), ("어디 있는지", {"bold": True}),
                             ("를 알려주고, 골든아워는 ", {"color": SUB}),
                             ("지금 받아주는지, 몇 분 걸리는지", {"bold": True, "color": RED}),
                             ("를 알려준다.", {"color": SUB})]], size=20)
pillars = [
    ("실시간 병상", "국립중앙의료원 응급실 병상 API를\n60초마다 동기화.\n병상 수와 갱신 시각을 함께 표시."),
    ("이동 시간 기반 추천", "예상 소요시간 + 병상 없음 페널티로\n종합 점수 산출.\n거리순과 비교해 순위 역전을 보여준다."),
    ("증상 → 병원 필터", "음성 또는 직접 입력한 증상을\n장비·소아 필터로 변환.\n위험 신호면 119 안내를 먼저."),
]
for i, (t, b) in enumerate(pillars):
    x = 0.6 + i * 4.1
    box(s, x, 2.55, 3.85, 3.3, fill=WHITE, border=LINE)
    box(s, x, 2.55, 3.85, 0.12, fill=RED, border=None, radius=False)
    text(s, x + 0.35, 2.95, 3.2, 0.5, t, size=22, bold=True)
    text(s, x + 0.35, 3.6, 3.3, 2.1, b, size=15, color=SUB, spacing=1.4)
text(s, 0.6, 6.2, 12.1, 0.5, "모든 기능은 웹 브라우저에서 동작한다. 설치 없이 검색이나 링크로 바로 들어온다.",
     size=14, color=SUB, align=PP_ALIGN.CENTER)

# ── 4. 주요 기능 ──────────────────────────────────────────────────
s = new_slide("03 · 주요 기능", "주요 기능 한눈에 보기",
              notes=(
                  "주요 기능입니다. 1번부터 6번은 API 명세의 기능 번호와 같습니다. "
                  "병상 지도, 이동 시간 계산, 종합 추천, 지금 진료 중인 병의원과 약국, 병원 상세와 길안내, 외부 API 상태 확인입니다. "
                  "여기에 증상 음성 입력과 119 안내, 병원에 보여주는 접수 카드가 더해집니다."))
feats = [
    ("①", "실시간 병상 지도", "반경 내 병원과 최신 병상 수,\n직선거리를 지도에 표시"),
    ("②", "이동 시간(ETA)", "후보 병원 최대 30곳의\n예상 도착 시간 일괄 산출"),
    ("③", "종합 추천", "병상 + ETA 점수로 정렬.\n거리순 대비 순위 역전 표시"),
    ("④", "지금 진료 중", "요일별 운영시간·공휴일로\n지금 여는 병·의원과 약국"),
    ("⑤", "상세 · 길안내", "응급실 전화, 길찾기,\n로드뷰, 링크 공유"),
    ("⑥", "외부 API 상태", "마지막 성공 시각으로\n장애 감지와 자동 전환"),
    ("+", "증상 입력 · 119", "음성/직접 입력 → 필터.\n위험 신호면 119 신고 문장"),
    ("+", "접수 카드", "알러지·복용약을 병원에\n보여주는 화면 (기기 저장)"),
]
for i, (n, t, b) in enumerate(feats):
    col, row = i % 4, i // 4
    x, y = 0.6 + col * 3.05, 1.65 + row * 2.6
    box(s, x, y, 2.85, 2.35, fill=WHITE, border=LINE)
    box(s, x + 0.25, y + 0.25, 0.5, 0.5, n, fill=RED if n != "+" else NAVY_2, border=None, size=16, bold=True,
        color=WHITE, pad=0)
    text(s, x + 0.25, y + 0.95, 2.4, 0.4, t, size=17, bold=True)
    text(s, x + 0.25, y + 1.42, 2.45, 0.9, b, size=12, color=SUB, spacing=1.3)

# ── 5. 차별점 ─────────────────────────────────────────────────────
s = new_slide("04 · 차별점", "일반 지도 앱과 무엇이 다른가",
              notes=(
                  "일반 지도 앱과 비교했습니다. 지도 앱은 가까운 순서로 위치를 보여줍니다. "
                  "골든아워는 병상과 이동 시간을 합친 점수로 정렬하고, 갱신 시각과 응급의료기관 등급까지 보여줘서 "
                  "사용자가 데이터를 믿을 수 있는지도 판단할 수 있게 했습니다."))
table(s, 0.6, 1.65, 12.1, [
    ["비교 항목", "일반 지도 앱", "골든아워"],
    ["정렬 기준", "직선거리 · 검색 정확도", "예상 소요시간 + 병상 없음 페널티 (종합 점수)"],
    ["응급실 병상", "병원마다 눌러보거나 전화", "목록·지도에 실시간 병상 수 바로 표시"],
    ["데이터 신뢰도", "표시 없음", "병상 갱신 시각 + 60분 경과 경고 + 기관 등급 배지"],
    ["증상 반영", "없음", "증상 → 장비·소아 필터, 중증이면 센터급 우선"],
    ["위급 상황", "없음", "위험 신호 감지 시 119 안내와 신고 문장 먼저"],
    ["병원 도착 후", "없음", "알러지·복용약 접수 카드 (서버 저장 없음)"],
], col_w=[2.4, 3.9, 5.8], row_h=0.62, size=14)
box(s, 0.6, 6.25, 12.1, 0.55, "핵심: 거리순 1위와 종합순 1위가 다를 때, 그 차이를 화면에 그대로 보여준다",
    fill=RED_SOFT, border=None, size=15, bold=True, color=RED)

# ── 6. 기술 스택 ──────────────────────────────────────────────────
s = new_slide("05 · 기술 스택", "기술 스택",
              notes=(
                  "프론트엔드는 React 19와 TypeScript, Vite로 만들었습니다. 백엔드는 Spring Boot 4.1과 Java 17, JPA를 썼고, "
                  "DB는 PostgreSQL 16을 Docker Compose로 띄웁니다. 외부 데이터는 국립중앙의료원 공공 API와 카카오 API를 씁니다. "
                  "음성 입력, 위치, 건강 정보 저장은 브라우저 기능으로 처리했습니다."))
stack = [
    ("프론트엔드", ["React 19", "TypeScript", "Vite", "React Router 7", "axios"], BLUE_SOFT, BLUE),
    ("백엔드", ["Spring Boot 4.1", "Java 17", "Spring Data JPA", "Validation", "Actuator", "Lombok"], GREEN_SOFT, GREEN),
    ("데이터베이스", ["PostgreSQL 16", "Docker Compose"], AMBER_SOFT, AMBER),
    ("외부 API", ["국립중앙의료원 병상", "공공데이터 약국", "카카오 지도 SDK", "카카오 로컬", "카카오모빌리티"],
     RED_SOFT, RED),
    ("브라우저 기능", ["Web Speech API (음성)", "Geolocation (위치)", "localStorage (건강 정보)"], BLUE_SOFT, BLUE),
    ("협업", ["모노레포 backend / frontend / docs", "Git 브랜치 병합"], RGBColor(0xF1, 0xF5, 0xF9), SUB),
]
for i, (label, items, f, c) in enumerate(stack):
    y = 1.6 + i * 0.86
    text(s, 0.6, y + 0.12, 2.2, 0.45, label, size=16, bold=True)
    cx = 2.9
    for it in items:
        ko = sum(1 for ch in it if ord(ch) > 0x3000)
        w = 0.35 + 0.1 * (len(it) - ko) + 0.19 * ko
        box(s, cx, y, w, 0.55, it, fill=f, border=None, size=12, bold=True, color=c, pad=0.06)
        cx += w + 0.12

# ── 7. 아키텍처 ───────────────────────────────────────────────────
s = new_slide("06 · 시스템 아키텍처", "시스템 아키텍처",
              notes=(
                  "전체 구조입니다. 사용자는 브라우저에서 React 앱을 씁니다. 지도는 카카오 지도 SDK로 그리고, 위치와 음성 입력은 브라우저 기능입니다. "
                  "건강 정보는 서버로 보내지 않고 브라우저에만 저장합니다. "
                  "프론트는 같은 주소의 /api/v1로 백엔드를 호출합니다. 백엔드에는 스케줄러가 있어서 60초마다 국립중앙의료원 병상 API를, "
                  "매일 새벽 4시에 병원 목록을 갱신해 PostgreSQL에 저장합니다. 그래서 사용자 요청은 공공 API를 직접 부르지 않고 DB에서 바로 응답합니다. "
                  "API 키는 서버 환경변수에만 있습니다."))
# 브라우저
box(s, 0.6, 1.6, 3.6, 5.1, fill=WHITE, border=BLUE, border_w=1.5)
text(s, 0.85, 1.75, 3.2, 0.4, "브라우저 (사용자)", size=16, bold=True, color=BLUE)
text(s, 0.85, 2.15, 3.2, 0.3, "React 19 · TypeScript · Vite", size=11, color=SUB)
for i, t in enumerate(["카카오 지도 JS SDK · 로드뷰", "Geolocation · 현재 위치", "Web Speech · 음성 증상 입력",
                       "증상 → 필터 / 119 안내", "localStorage · 건강 정보"]):
    box(s, 0.85, 2.6 + i * 0.75, 3.1, 0.58, t, fill=BLUE_SOFT, border=None, size=12, bold=True, color=BLUE)
# 백엔드
box(s, 4.85, 1.6, 4.4, 5.1, fill=WHITE, border=GREEN, border_w=1.5)
text(s, 5.1, 1.75, 4, 0.4, "백엔드 (Spring Boot 4.1)", size=16, bold=True, color=GREEN)
text(s, 5.1, 2.15, 4, 0.3, "Java 17 · JPA · 키는 .env / 환경변수", size=11, color=SUB)
box(s, 5.1, 2.6, 3.9, 0.7, ["REST Controller", ("/api/v1/hospitals · facilities · system", {"size": 10, "bold": False})],
    fill=GREEN_SOFT, border=None, size=13, bold=True, color=GREEN)
box(s, 5.1, 3.55, 3.9, 0.7, ["추천 · ETA · 병원 Service", ("Haversine 반경 · 점수 계산 · 순위 비교", {"size": 10, "bold": False})],
    fill=GREEN_SOFT, border=None, size=13, bold=True, color=GREEN)
box(s, 5.1, 4.5, 3.9, 0.7, ["JPA Repository", ("hospital · hours · pharmacy …", {"size": 10, "bold": False})],
    fill=GREEN_SOFT, border=None, size=13, bold=True, color=GREEN)
box(s, 5.1, 5.6, 3.9, 0.8, ["@Scheduled 동기화", ("병상 60초 · 병원 목록 매일 04시", {"size": 10, "bold": False})],
    fill=AMBER_SOFT, border=None, size=13, bold=True, color=AMBER)
arrow(s, 7.05, 3.3, 7.05, 3.55)
arrow(s, 7.05, 4.25, 7.05, 4.5)
# DB
box(s, 9.9, 1.6, 2.85, 1.6, ["PostgreSQL 16", ("Docker Compose", {"size": 11, "bold": False, "color": SUB}),
                              ("7개 테이블", {"size": 11, "bold": False, "color": SUB})],
    fill=WHITE, border=AMBER, border_w=1.5, size=15, bold=True, color=AMBER)
# 외부
box(s, 9.9, 3.55, 2.85, 3.15, fill=WHITE, border=RED, border_w=1.5)
text(s, 10.1, 3.68, 2.5, 0.35, "외부 API", size=15, bold=True, color=RED)
for i, t in enumerate(["국립중앙의료원 병상", "공공데이터 약국", "특일정보 (공휴일)", "카카오 로컬 검색", "카카오모빌리티 길찾기"]):
    box(s, 10.1, 4.08 + i * 0.5, 2.45, 0.4, t, fill=RED_SOFT, border=None, size=11, bold=True, color=RED, pad=0.05)
# 연결
arrow(s, 4.2, 2.95, 5.1, 2.95, color=BLUE, width=2)
text(s, 4.22, 2.6, 0.9, 0.3, "HTTP", size=10, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
arrow(s, 9.0, 4.85, 9.9, 2.6, color=AMBER, width=2)
arrow(s, 9.0, 6.0, 10.1, 6.0, color=RED, width=2)
text(s, 8.98, 6.08, 1.1, 0.3, "주기 호출", size=10, bold=True, color=RED, align=PP_ALIGN.CENTER)
arrow(s, 9.0, 3.9, 10.1, 5.85, color=RED, width=1.25, dashed=True)

# ── 8. 외부 API ───────────────────────────────────────────────────
s = new_slide("07 · 외부 API 연동", "외부 API 연동",
              "외부 API 키는 레포에 커밋하지 않고 backend/.env와 환경변수로만 주입한다",
              notes=(
                  "사용한 외부 API입니다. 핵심은 국립중앙의료원 응급의료기관 API로, 병원 정보와 실시간 병상, 중증질환 수용 정보를 줍니다. "
                  "공공데이터 API는 모두 자동승인이라 바로 썼고, 카카오모빌리티 길찾기만 수동승인이라 가장 먼저 신청했습니다. "
                  "승인 전에는 직선거리로 이동 시간을 추정하고 있습니다. 이건 뒤에서 자세히 말씀드리겠습니다."))
table(s, 0.6, 1.95, 12.1, [
    ["API", "제공", "용도", "승인"],
    ["전국 응급의료기관 정보 조회", "공공데이터포털", "병원 정보 · 실시간 가용병상 · 중증질환 수용 (①③)", "자동"],
    ["전국 병·의원 찾기", "공공데이터포털", "지금 진료 중인 병·의원 (④)", "자동"],
    ["전국 약국 정보 조회", "공공데이터포털", "지금 여는 약국 (④) · 실패 시 카카오 검색 대체", "자동"],
    ["특일 정보", "한국천문연구원", "공휴일 판정 → 운영시간 계산 (④)", "자동"],
    ["길찾기 (자동차 · 다중목적지)", "카카오모빌리티", "후보 병원 ETA 일괄 산출 (②)", "수동 · 승인 대기"],
    ["지도 / 로컬", "카카오", "지도 UI · 로드뷰 · 좌표 확인", "즉시"],
], col_w=[3.4, 2.1, 4.9, 1.7], row_h=0.6, size=13)

# ── 9. ERD ────────────────────────────────────────────────────────
s = new_slide("08 · ERD", "ERD (7개 테이블)",
              notes=(
                  "ERD입니다. 가운데 HOSPITAL이 병원 마스터이고 최신 병상 수와 장비 보유 여부를 함께 갖습니다. "
                  "왼쪽 BED_STATUS_HISTORY는 60초마다 바뀌는 병상의 이력이고, HOSPITAL_HOURS는 요일별 운영시간입니다. "
                  "오른쪽 SEARCH_LOG와 SEARCH_RESULT_CANDIDATE는 검색 한 번마다 후보 병원의 거리 순위, 시간 순위, 종합 점수를 저장하도록 설계했습니다. "
                  "약국은 PHARMACY와 PHARMACY_HOURS로 따로 둡니다. "
                  "노란 점선으로 표시한 테이블은 엔티티까지는 만들었지만 아직 저장 로직이 붙지 않았습니다."))
bsh = entity(s, 0.6, 1.55, 3.55, "BED_STATUS_HISTORY", [
    ("PK", "id"), ("FK", "hospital_id"), ("", "general / pediatric / icu_beds"), ("", "source"), ("", "recorded_at")],
    note="저장 미구현", dashed=True)
hh = entity(s, 0.6, 3.95, 3.55, "HOSPITAL_HOURS", [
    ("PK", "id"), ("FK", "hospital_id"), ("", "day_of_week"), ("", "open_time / close_time"), ("", "is_24h")])
hos = entity(s, 4.85, 1.55, 3.65, "HOSPITAL", [
    ("PK", "id"), ("", "name · address · phone"), ("", "lat / lng"), ("", "hospital_type (권역/지역)"),
    ("", "latest_general_beds"), ("", "latest_pediatric_beds"), ("", "latest_icu_beds"),
    ("", "장비: CT · MRI · 혈관촬영 · 인공호흡기"), ("", "beds_updated_at")])
src = entity(s, 9.2, 1.55, 3.55, "SEARCH_RESULT_CANDIDATE", [
    ("PK", "id"), ("FK", "search_log_id"), ("FK", "hospital_id"), ("", "straight_distance_m"), ("", "eta_seconds"),
    ("", "available_beds_snapshot"), ("", "distance_rank / eta_rank"), ("", "combined_score / rank")],
    note="저장 미구현", dashed=True)
sl = entity(s, 9.2, 4.55, 3.55, "SEARCH_LOG", [
    ("PK", "id"), ("", "session_id"), ("", "origin_lat / origin_lng"), ("", "requested_at")],
    note="저장 미구현", dashed=True)
ph = entity(s, 4.85, 4.75, 1.6, "PHARMACY", [("PK", "id"), ("", "name · phone"), ("", "lat / lng")])
phh = entity(s, 6.6, 4.75, 1.9, "PHARMACY_HOURS", [("PK", "id"), ("FK", "pharmacy_id"), ("", "day · open/close")])
for (x1, y1, x2, y2, lx, ly) in [
    (4.85, 2.4, 4.15, 2.4, 4.2, 2.08),
    (4.85, 3.4, 4.15, 4.6, 4.2, 3.75),
    (8.5, 2.6, 9.2, 2.6, 8.55, 2.28),
    (10.95, 4.55, 10.95, 3.88, 10.4, 4.08),
    (6.45, 5.35, 6.6, 5.35, 6.25, 5.95),
]:
    arrow(s, x1, y1, x2, y2, color=NAVY_2, width=1.5)
    text(s, lx, ly, 0.6, 0.28, "1 : N", size=10, bold=True, color=NAVY_2)
text(s, 0.6, 6.55, 12, 0.35, [[("■ ", {"color": AMBER}), ("노란 점선", {"bold": True}),
                               (": 엔티티 설계 완료, 저장 로직 연결 예정     ", {"color": SUB}),
                               ("→ ", {"color": NAVY_2}), ("1 : N 관계", {"color": SUB})]], size=12)

# ── 10. 설계 포인트 ───────────────────────────────────────────────
s = new_slide("09 · 테이블 설계", "왜 이렇게 나눴나",
              notes=(
                  "설계에서 고민한 세 가지입니다. 첫째, 병상은 60초마다 바뀌는 시계열이라 이력 테이블을 따로 두었습니다. "
                  "둘째, 그런데 매번 이력을 훑으면 느리니까 최신 값은 HOSPITAL에 중복 저장했습니다. 목표 응답시간은 P95 1.5초입니다. "
                  "셋째, 거리순과 시간순이 뒤바뀌는 사례를 숫자로 모으려고 검색 스냅샷 테이블을 설계했습니다."))
card(s, 0.6, 1.6, 3.9, 4.0, "이력 분리", "병상은 시계열 데이터",
     "병상은 60초마다 바뀐다.\nBED_STATUS_HISTORY에 시점별로 남겨\n갱신 지연 감시와 통계에 쓴다.")
card(s, 4.7, 1.6, 3.9, 4.0, "비정규화", "최신 병상은 HOSPITAL에",
     "조회마다 이력 테이블을 훑지 않도록\n최신 병상 수와 갱신 시각을\nHOSPITAL에 중복 저장.\n\n목표 응답시간 P95 1.5초.")
card(s, 8.8, 1.6, 3.9, 4.0, "검색 스냅샷", "순위 역전을 숫자로",
     "검색 1회마다 후보 병원의\n거리 순위 · 시간 순위 · 종합 점수 저장.\n\n“거리순 1위 ≠ 시간순 1위” 비율을\n정량으로 수집하는 데이터 소스.")
box(s, 0.6, 5.85, 12.1, 0.85, [[("반경 검색: ", {"bold": True}),
                                 ("지금은 자바에서 Haversine 공식으로 거른다. 병원 수가 늘면 PostGIS ", {"color": SUB}),
                                 ("geography(Point) + GiST 인덱스 + ST_DWithin", {"bold": True}),
                                 ("으로 DB에서 처리할 계획.", {"color": SUB})]],
    fill=WHITE, border=LINE, size=13, align=PP_ALIGN.LEFT, pad=0.3)

# ── 11. 내부 API ──────────────────────────────────────────────────
s = new_slide("10 · 내부 API", "내부 REST API 명세", "프론트엔드 ↔ 백엔드 · 기본 경로 /api/v1 · 같은 출처로 요청",
              notes=(
                  "프론트와 백엔드 사이 API입니다. 반경 내 병원 조회, 이동 시간 일괄 계산, 종합 추천, 지금 진료 중인 시설, 병원 상세, "
                  "그리고 외부 API 상태 확인이 있습니다. 상태 API는 공공 API가 죽었을 때 감지하고 대체 동작으로 바꾸는 판단에 씁니다. "
                  "API 기본 주소를 비워 같은 출처로 요청하게 해서 폰이나 다른 네트워크에서도 동작하게 했습니다."))
table(s, 0.6, 1.95, 12.1, [
    ["Method", "Path", "설명", "기능"],
    ["GET", "/api/v1/hospitals/nearby", "반경 내 병원 + 최신 병상현황 + 직선거리", "①"],
    ["POST", "/api/v1/hospitals/eta", "후보 병원(최대 30곳) 다중목적지 ETA 일괄 산출", "②"],
    ["GET", "/api/v1/hospitals/recommend", "병상 + ETA 결합점수 추천순서, 직선거리 순위와 비교", "③"],
    ["GET", "/api/v1/facilities/open-now", "현재 진료 중인 병·의원 / 약국", "④"],
    ["GET", "/api/v1/hospitals/{id}", "상세정보 · 전화 · 경로 딥링크", "⑤"],
    ["GET", "/api/v1/system/external-status", "외부 API 최종 성공 시각 / 상태 (장애 감지 · 자동전환)", "⑥"],
], col_w=[1.2, 4.0, 5.9, 1.0], row_h=0.62, size=13)

# ── 12. 흐름: 병상 동기화 ────────────────────────────────────────
s = new_slide("11 · 기능 흐름 ①", "병상 데이터 동기화 흐름",
              "사용자 요청과 공공 API 호출을 분리했다. 사용자는 항상 DB에 있는 최신 값을 받는다.",
              notes=(
                  "병상 데이터가 들어오는 흐름입니다. 사용자가 요청할 때마다 공공 API를 부르면 느리고 호출 제한에도 걸립니다. "
                  "그래서 스케줄러가 60초마다 병상 API를 불러 HOSPITAL 테이블의 최신 병상과 갱신 시각을 바꿉니다. "
                  "병원 목록 자체는 매일 새벽 4시에 갱신합니다. API 키가 없거나 호출이 실패해도 건너뛰고 서버는 계속 동작합니다. "
                  "화면에는 갱신 시각을 보여주고, 60분이 넘으면 경고를 띄웁니다."))
steps = [("@Scheduled", "60초마다 실행"), ("병상 API 호출", "국립중앙의료원"), ("응답 파싱", "병원별 병상 · 장비"),
         ("HOSPITAL 갱신", "latest_*_beds\nbeds_updated_at"), ("화면 표시", "병상 수 + 갱신 시각\n60분 경과 시 경고")]
for i, (t, b) in enumerate(steps):
    x = 0.6 + i * 2.5
    box(s, x, 2.3, 2.1, 1.5, [(t, {"size": 15, "bold": True}), (b, {"size": 11, "color": SUB})],
        fill=WHITE if i else AMBER_SOFT, border=AMBER if i == 0 else LINE, border_w=1.25)
    if i < len(steps) - 1:
        arrow(s, x + 2.1, 3.05, x + 2.5, 3.05, color=SUB, width=2)
box(s, 3.1, 4.5, 2.1, 1.0, [("실패 / 키 없음", {"size": 14, "bold": True, "color": RED}),
                            ("건너뛰고 서버 계속 동작", {"size": 11, "color": SUB})],
    fill=RED_SOFT, border=None)
arrow(s, 4.15, 3.8, 4.15, 4.5, color=RED, width=1.5, dashed=True)
box(s, 0.6, 4.5, 2.1, 1.0, [("매일 04시", {"size": 14, "bold": True, "color": AMBER}),
                            ("병원 목록 전체 갱신", {"size": 11, "color": SUB})],
    fill=AMBER_SOFT, border=None)
box(s, 5.85, 4.5, 6.85, 1.85, [
    ("사용자 요청 시", {"size": 15, "bold": True}),
    ("GET /hospitals/nearby → HOSPITAL 테이블에서 바로 응답", {"size": 13, "color": SUB}),
    ("공공 API를 직접 부르지 않는다 → 빠르고, 호출 한도에 안 걸리고,", {"size": 13, "color": SUB}),
    ("공공 API가 잠시 죽어도 마지막 값으로 서비스가 유지된다.", {"size": 13, "color": SUB}),
], fill=WHITE, border=LINE, align=PP_ALIGN.LEFT, pad=0.3, spacing=1.35)

# ── 13. 흐름: 추천 ────────────────────────────────────────────────
s = new_slide("12 · 기능 흐름 ②", "병원 추천 흐름",
              notes=(
                  "추천 흐름입니다. 브라우저가 현재 위치를 받아 추천 API를 부릅니다. 서버는 반경 안의 병원을 고르고, 각 병원까지 예상 소요시간을 구합니다. "
                  "점수는 소요시간 분에 병상이 없으면 30분을 더한 값입니다. 점수가 낮을수록 위로 갑니다. "
                  "서버는 거리순과 종합순 두 목록을 같이 내려주고, 프론트는 두 순위가 다를 때 순위 역전을 표시합니다. "
                  "길찾기 API를 못 쓰는 상황에서는 degraded 상태로 거리순에 고정하고 안내 배너를 띄웁니다."))
row1 = [("위치 확인", "Geolocation\n(브라우저)"), ("GET /recommend", "현재 좌표 전달"),
        ("반경 내 병원", "Haversine\n직선거리 계산"), ("ETA 산출", "카카오모빌리티\n(승인 전: 추정식)")]
row2 = [("점수 계산", "ETA(분)\n+ 병상 0이면 +30"), ("두 목록 응답", "거리순 · 종합순"),
        ("화면 표시", "순위 역전 배지\n병상 · 갱신 시각")]
for i, (t, b) in enumerate(row1):
    x = 0.6 + i * 3.1
    box(s, x, 1.65, 2.6, 1.5, [(t, {"size": 15, "bold": True}), (b, {"size": 11, "color": SUB})],
        fill=BLUE_SOFT if i == 0 else WHITE, border=BLUE if i == 0 else LINE)
    if i < 3:
        arrow(s, x + 2.6, 2.4, x + 3.1, 2.4, width=2)
arrow(s, 10.8, 3.15, 10.8, 3.75, width=2)
for i, (t, b) in enumerate(row2):
    x = 9.5 - i * 3.1
    box(s, x, 3.75, 2.6, 1.5, [(t, {"size": 15, "bold": True}), (b, {"size": 11, "color": SUB})],
        fill=RED_SOFT if i == 2 else WHITE, border=RED if i == 2 else LINE)
    if i < 2:
        arrow(s, x, 4.5, x - 0.5, 4.5, width=2)
box(s, 0.6, 5.6, 5.8, 1.15, [("ETA를 못 받으면 → degraded 상태", {"size": 14, "bold": True, "color": AMBER}),
                             ("거리순으로 고정하고 안내 배너 표시", {"size": 12, "color": SUB})],
    fill=AMBER_SOFT, border=None, align=PP_ALIGN.LEFT, pad=0.3)
box(s, 6.6, 5.6, 6.1, 1.15, [("ETA 추정식 (길찾기 승인 전)", {"size": 14, "bold": True}),
                             ("직선거리 × 우회계수 1.4 ÷ 시속 40km", {"size": 12, "color": SUB})],
    fill=WHITE, border=LINE, align=PP_ALIGN.LEFT, pad=0.3)

# ── 14. 점수 & 역전 예시 ─────────────────────────────────────────
s = new_slide("13 · 추천 점수", "점수 계산과 순위 역전",
              notes=(
                  "점수 계산을 예시로 보여드리겠습니다. 숫자는 설명용 예시입니다. "
                  "A병원은 가장 가까워서 거리순 1위지만 병상이 0개라 30분 페널티를 받아 39점이 됩니다. "
                  "B병원은 조금 멀지만 병상이 있어서 13점으로 종합 1위가 됩니다. "
                  "일반 지도 앱이었다면 A병원으로 갔다가 다시 옮겨야 했을 겁니다. 이 차이를 화면에서 바로 보여주는 게 골든아워의 핵심입니다."))
box(s, 0.6, 1.6, 12.1, 1.2, [[("종합 점수  =  ", {"color": SUB}), ("예상 소요시간(분)", {"bold": True}),
                              ("  +  ", {"color": SUB}), ("병상이 없으면 30분 페널티", {"bold": True, "color": RED})],
                             ("점수가 낮을수록 위에 추천", {"color": SUB, "size": 14})],
    fill=WHITE, border=LINE, size=26)
table(s, 0.6, 3.1, 12.1, [
    ["병원 (예시)", "직선거리", "예상 소요", "잔여 병상", "점수", "거리순", "종합순"],
    ["A병원", "2.1 km", "9분", "0", "9 + 30 = 39", "1위", "3위"],
    ["B병원", "3.4 km", "13분", "4", "13", "2위", "1위"],
    ["C병원", "4.0 km", "15분", "2", "15", "3위", "2위"],
], col_w=[2.2, 1.6, 1.6, 1.6, 2.3, 1.4, 1.4], row_h=0.6, size=15)
box(s, 0.6, 5.75, 12.1, 0.9, [[("가장 가까운 A병원은 병상이 없다. ", {"bold": True}),
                               ("일반 지도 앱이면 A로 갔다가 다시 옮긴다. 골든아워는 B병원을 1위로 올리고 ", {}),
                               ("순위 역전", {"bold": True, "color": RED}), ("을 표시한다.", {})]],
    fill=RED_SOFT, border=None, size=15)
text(s, 0.6, 6.72, 12, 0.3, "* 표의 수치는 설명을 위한 예시", size=10, color=MUTED)

# ── 15. 흐름: 증상 ────────────────────────────────────────────────
s = new_slide("14 · 기능 흐름 ③", "증상 입력 → 병원 필터 흐름",
              notes=(
                  "증상 기능입니다. 음성 입력과 직접 입력은 같은 분석 로직을 거칩니다. "
                  "음성 인식이 틀릴 수 있어서, 인식된 문장을 입력칸에 띄워 사용자가 고칠 수 있게 했고, 버튼을 눌러야만 필터가 적용됩니다. "
                  "분석 결과 위험 신호가 있으면 병원 목록보다 119 안내와 신고 문장을 먼저 보여줍니다. "
                  "그렇지 않으면 증상을 장비 필터나 소아 필터로 바꾸고, 중증이면 권역이나 지역센터급 병원을 우선 보여줍니다. "
                  "이건 의료 진단이 아니라 병원 선택을 돕는 참고 기능입니다."))
box(s, 0.6, 1.7, 2.5, 1.1, [("음성 입력", {"size": 15, "bold": True}), ("Web Speech API", {"size": 11, "color": SUB})],
    fill=BLUE_SOFT, border=BLUE)
box(s, 0.6, 3.15, 2.5, 1.1, [("직접 입력", {"size": 15, "bold": True}), ("텍스트", {"size": 11, "color": SUB})],
    fill=BLUE_SOFT, border=BLUE)
box(s, 3.7, 2.45, 2.4, 1.4, [("입력칸에서 확인", {"size": 15, "bold": True}), ("수정 후 눌러야 적용", {"size": 11, "color": SUB})],
    fill=WHITE, border=LINE)
arrow(s, 3.1, 2.25, 3.7, 2.95, width=2)
arrow(s, 3.1, 3.7, 3.7, 3.35, width=2)
box(s, 6.7, 2.45, 2.4, 1.4, [("증상 분석", {"size": 15, "bold": True}), ("키워드 기반 · 공통 로직", {"size": 11, "color": SUB})],
    fill=WHITE, border=LINE)
arrow(s, 6.1, 3.15, 6.7, 3.15, width=2)
dia = s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(9.6), Inches(2.3), Inches(1.8), Inches(1.7))
dia.fill.solid()
dia.fill.fore_color.rgb = WHITE
dia.line.color.rgb = RED
dia.line.width = Pt(1.5)
dia.shadow.inherit = False
fill_text(dia.text_frame, [("위험", {"bold": True}), ("신호?", {"bold": True})], size=14, align=PP_ALIGN.CENTER)
arrow(s, 9.1, 3.15, 9.6, 3.15, width=2)
box(s, 9.35, 4.55, 3.4, 1.3, [("119 안내 먼저", {"size": 15, "bold": True, "color": WHITE}),
                              ("신고 문장 생성 · 전화 버튼", {"size": 11, "color": WHITE})],
    fill=RED, border=None)
arrow(s, 10.5, 4.0, 10.5, 4.55, color=RED, width=2)
text(s, 10.6, 4.1, 0.6, 0.3, "예", size=12, bold=True, color=RED)
box(s, 5.4, 4.55, 3.4, 1.3, [("필터로 변환", {"size": 15, "bold": True}),
                             ("CT · MRI 등 장비 · 소아 병상", {"size": 11, "color": SUB})], fill=WHITE, border=LINE)
arrow(s, 10.05, 3.58, 8.8, 4.9, color=SUB, width=1.5)
text(s, 8.85, 4.05, 0.9, 0.3, "아니오", size=12, bold=True, color=SUB)
box(s, 0.6, 4.55, 4.2, 1.3, [("추천 목록 갱신", {"size": 15, "bold": True}),
                             ("중증이면 센터급 병원 우선", {"size": 11, "color": SUB})], fill=GREEN_SOFT, border=GREEN)
arrow(s, 5.4, 5.2, 4.8, 5.2, width=2)
text(s, 0.6, 6.3, 12.1, 0.4, "※ 의료적 진단이 아니라 병원 선택을 돕는 참고 기능", size=12, color=MUTED)

# ── 16. 신뢰 · 장애 대응 ─────────────────────────────────────────
s = new_slide("15 · 안정성", "데이터를 믿을 수 있게, 장애에도 멈추지 않게",
              notes=(
                  "응급 서비스라서 데이터를 믿을 수 있는지와 장애에도 멈추지 않는지가 중요했습니다. "
                  "병상 갱신 시각을 항상 보여주고 60분이 넘으면 경고합니다. 응급의료기관 등급 배지도 붙였습니다. "
                  "외부 API가 실패해도 서버는 계속 돌고, 길찾기가 안 되면 거리순으로, 약국 공공 API가 안 되면 카카오 검색으로 대체합니다. "
                  "건강 정보는 서버에 저장하지 않고 사용자 기기에만 둡니다."))
items = [
    ("신뢰", "병상 갱신 시각 표시", "모든 병원 카드에 갱신 시각.\n60분이 넘으면 경고.", BLUE_SOFT, BLUE),
    ("신뢰", "응급의료기관 등급 배지", "권역 · 지역센터 · 지역기관 구분.\n중증이면 센터급 우선.", BLUE_SOFT, BLUE),
    ("장애 대응", "외부 API 실패", "키가 없거나 실패해도 건너뛰고\n서버는 계속 동작.", AMBER_SOFT, AMBER),
    ("장애 대응", "ETA 없음 → degraded", "거리순으로 고정하고\n안내 배너 표시.", AMBER_SOFT, AMBER),
    ("장애 대응", "약국 API 실패", "카카오 로컬 검색으로 대체.\n5분 캐시.", AMBER_SOFT, AMBER),
    ("개인정보", "건강 정보는 기기에만", "알러지 · 복용약은 localStorage.\n서버로 보내지 않는다.", GREEN_SOFT, GREEN),
]
for i, (tag, t, b, f, c) in enumerate(items):
    col, row = i % 3, i // 3
    card(s, 0.6 + col * 4.1, 1.65 + row * 2.6, 3.85, 2.35, tag, t, b, tag_fill=f, tag_color=c)

# ── 17. 어려웠던 점 ──────────────────────────────────────────────
s = new_slide("16 · 트러블슈팅", "어려웠던 점과 해결",
              notes=(
                  "개발하면서 어려웠던 점입니다. 가장 컸던 건 카카오모빌리티 길찾기 승인이 늦어진 것입니다. "
                  "그래서 직선거리에 우회계수 1.4를 곱하고 시속 40킬로로 나눠 임시로 추정했고, 승인 후에는 estimateEtaSeconds 함수 하나만 바꾸면 되도록 분리했습니다. "
                  "또 마이크가 폰이나 일부 브라우저에서 에러 없이 조용히 실패하는 문제가 있어서 오류 코드별 안내와 빈 결과 처리를 넣었습니다."))
table(s, 0.6, 1.6, 12.1, [
    ["어려웠던 점", "해결"],
    ["카카오모빌리티 길찾기 승인 전이라 ETA를 못 구함",
     "직선거리 × 1.4 ÷ 40km/h로 임시 추정. 승인 후 estimateEtaSeconds만 교체하도록 분리"],
    ["ETA를 받지 못하는 상황", "프론트에 degraded 상태 → 거리순 고정 + 안내 배너"],
    ["병상 데이터가 오래된 경우", "갱신 시각 표시 + 60분 경과 경고"],
    ["마이크가 폰 · 일부 브라우저에서 조용히 실패", "오류 코드별 안내 문구 + 오류 없이 끝났을 때 empty-result 처리"],
    ["폰 · 터널 등 여러 환경에서 접속", "API 기본 주소를 비워 같은 출처로 요청"],
    ["약국 공공 API 실패", "카카오 검색으로 대체 + 5분 캐시"],
    ["음성 인식 오류로 잘못된 필터 적용 위험", "입력칸에서 직접 수정 → 버튼을 눌러야만 필터 적용"],
], col_w=[4.6, 7.5], row_h=0.6, size=13)

# ── 18. 한계 ──────────────────────────────────────────────────────
s = new_slide("17 · 한계", "아쉬웠던 점",
              notes=(
                  "솔직하게 아쉬운 점도 말씀드리겠습니다. 가장 큰 건 지금 시간순이 실제 도로 시간이 아니라 추정치라는 점입니다. "
                  "그리고 순위 역전 사례를 모으려고 설계한 검색 로그 테이블과 병상 이력 테이블에 아직 저장 코드가 없습니다. "
                  "반경 검색은 자바에서 계산하고 있고, 마이그레이션 도구와 테스트도 부족합니다. 다음 장에서 이걸 어떻게 개선할지 말씀드리겠습니다."))
lims = [
    ("실제 도로 ETA가 아님", "시간순이 추정치라 핵심 가치인\n‘실제 도착 시간순’이 아직 약하다."),
    ("기록 테이블 미사용", "검색 로그 · 병상 이력 테이블에\n저장 코드가 없어 역전 사례 수집이\n아직 동작하지 않는다."),
    ("반경 검색 성능", "전체 병원을 읽어 자바에서 Haversine.\nPostGIS 적용 전."),
    ("DB 관리", "ddl-auto: update, 마이그레이션 도구 없음.\n엔티티가 FK 대신 Long ID 보유."),
    ("테스트 부족", "백엔드는 컨텍스트 로딩 테스트 하나,\n프론트 테스트 없음."),
    ("증상 분석 한계", "키워드 기반이라 다양한 표현을 놓친다.\n의료 판단이 아닌 참고용."),
]
for i, (t, b) in enumerate(lims):
    col, row = i % 3, i // 3
    card(s, 0.6 + col * 4.1, 1.65 + row * 2.6, 3.85, 2.35, None, t, b)

# ── 19. 발전 방향 ────────────────────────────────────────────────
s = new_slide("18 · 발전 방향", "앞으로의 계획",
              notes=(
                  "앞으로 계획입니다. 가장 먼저 길찾기 API를 붙여 실제 도로 시간을 쓰겠습니다. 반경 안 상위 N개만 호출하고 캐시와 대체 로직을 둡니다. "
                  "다음으로 병상 이력과 검색 스냅샷 저장을 켜서, 거리순 1위와 시간순 1위가 다른 비율을 실제 지표로 뽑겠습니다. "
                  "그 외에 PostGIS와 Flyway, 추천 점수 고도화, 테스트와 CI, 야간과 휴일 진료기관 통합을 계획하고 있습니다."))
plan = [
    ("1", "실시간 ETA 적용", "길찾기 API 연동 · 상위 N개만 호출 · 캐시 + 장애 대체", True),
    ("2", "데이터 기록 활성화", "병상 이력 · 검색 스냅샷 저장 → “거리순 1위 ≠ 시간순 1위” 비율 지표", True),
    ("3", "PostGIS · Flyway", "반경 검색을 DB에서 처리, 스키마 변경 이력 관리", False),
    ("4", "추천 점수 고도화", "병상 수 · 장비 · 소아 · 데이터 신선도 · 중증도 가중치", False),
    ("5", "증상 분석 개선", "한국어 표현 사전 확장 · 의료진 검수 규칙", False),
    ("6", "운영 품질", "동기화 실패 알림 · 모니터링 · 테스트 · CI", False),
    ("7", "기능 확장", "오프라인 119 안내 · 다국어 · 접수 카드 QR · 야간/휴일 진료기관 통합", False),
    ("8", "개인정보와 안전", "기기 저장 사실 명시 · 위치 동의 · 면책 문구 보강", False),
]
for i, (n, t, b, hot) in enumerate(plan):
    col, row = i % 2, i // 2
    x, y = 0.6 + col * 6.15, 1.6 + row * 1.32
    box(s, x, y, 5.95, 1.15, fill=WHITE, border=RED if hot else LINE, border_w=1.5 if hot else 1.0)
    box(s, x + 0.2, y + 0.3, 0.55, 0.55, n, fill=RED if hot else NAVY_2, border=None, size=16, bold=True,
        color=WHITE, pad=0)
    text(s, x + 0.95, y + 0.18, 4.9, 0.4, t, size=16, bold=True)
    text(s, x + 0.95, y + 0.6, 4.9, 0.5, b, size=11, color=SUB)

# ── 20. 시연 ──────────────────────────────────────────────────────
s = new_slide("19 · 시연", "시연 시나리오",
              "“새벽 1시, 같이 있던 친구가 넘어져 머리를 다쳤다”",
              notes=(
                  "이제 시연하겠습니다. 상황은 새벽 1시, 같이 있던 친구가 넘어져 머리를 다친 경우입니다. "
                  "사이트에 들어가면 위치를 받아 근처 응급실과 병상이 바로 뜹니다. 머리를 부딪혔다고 말하면 CT가 되는 병원만 남습니다. "
                  "종합순 1위와 거리순 1위가 다른 걸 확인하고, 병원을 눌러 응급실 전화와 길안내, 공유까지 해보겠습니다. "
                  "(시연이 안 되면 녹화 영상으로 대체)"))
demo = [("접속", "위치 허용\n근처 응급실 · 병상 표시"), ("음성 입력", "“머리를 부딪혔어요”\n입력칸 확인 후 적용"),
        ("필터 적용", "CT 가능한 병원만\n남는다"), ("순위 역전", "거리순 1위 ≠ 종합순 1위\n배지 확인"),
        ("병원 상세", "응급실 전화 · 길안내\n로드뷰"), ("공유 · 접수 카드", "보호자에게 링크 공유\n알러지 · 복용약 화면")]
for i, (t, b) in enumerate(demo):
    col, row = i % 3, i // 3
    x, y = 0.6 + col * 4.1, 2.05 + row * 2.15
    box(s, x, y, 3.85, 1.9, fill=WHITE, border=LINE)
    box(s, x + 0.25, y + 0.25, 0.55, 0.55, str(i + 1), fill=RED, border=None, size=16, bold=True, color=WHITE, pad=0)
    text(s, x + 1.0, y + 0.3, 2.7, 0.45, t, size=17, bold=True)
    text(s, x + 0.25, y + 0.95, 3.4, 0.9, b, size=13, color=SUB, spacing=1.3)
    if col < 2:
        arrow(s, x + 3.85, y + 0.95, x + 4.1, y + 0.95, width=2)
text(s, 0.6, 6.5, 12.1, 0.35, "네트워크 문제에 대비해 시연 녹화 영상을 백업으로 준비", size=12, color=MUTED, align=PP_ALIGN.CENTER)

# ── 21. 마무리 ────────────────────────────────────────────────────
s = new_slide(dark=True, notes=(
    "정리하겠습니다. 지도 앱은 병원이 어디 있는지 알려줍니다. 골든아워는 지금 받아주는지, 몇 분이면 도착하는지를 알려줍니다. "
    "응급 상황에서 비교할 시간 없이 한 화면에서 결정할 수 있게 하는 것이 저희의 목표입니다. 감사합니다. 질문 받겠습니다."))
box(s, 0.9, 1.6, 0.18, 2.4, fill=RED, border=None, radius=False)
text(s, 1.35, 1.5, 11, 0.6, "지도는 어디 있는지 알려준다.", size=32, color=MUTED)
text(s, 1.35, 2.25, 11, 1.6, ["골든아워는 지금 받아주는지,", "몇 분이면 도착하는지 알려준다."],
     size=40, bold=True, color=WHITE, spacing=1.2)
text(s, 1.35, 5.2, 11, 0.6, "감사합니다 · Q&A", size=26, bold=True, color=RED)
text(s, 1.35, 5.85, 11, 0.4, "팀 코드블루", size=16, color=MUTED)

out = __import__("pathlib").Path(__file__).with_name("golden-hour.pptx")
prs.save(out)
print(f"saved {out} ({len(prs.slides)} slides)")
