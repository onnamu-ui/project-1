#!/usr/bin/env python3
"""
업무 스킬 인수인계 자료 – HWP / DOCX / XLSX / PDF
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageFont
import os

FONT_DIR = "/usr/share/fonts/truetype/nanum"
REGULAR  = os.path.join(FONT_DIR, "NanumGothic.ttf")
BOLD     = os.path.join(FONT_DIR, "NanumGothicBold.ttf")
SQUARE_B = os.path.join(FONT_DIR, "NanumSquareB.ttf")

pdfmetrics.registerFont(TTFont("NanumGothic",     REGULAR))
pdfmetrics.registerFont(TTFont("NanumGothicBold", BOLD))
pdfmetrics.registerFont(TTFont("NanumSquareB",    SQUARE_B))

OUT_DIR = "/home/user/project-1/handover_docs"
os.makedirs(OUT_DIR, exist_ok=True)

C_HWP   = colors.HexColor("#005bac")   # 한글 파란색
C_DOCX  = colors.HexColor("#2b579a")   # Word 파란색
C_XLSX  = colors.HexColor("#1e7145")   # Excel 녹색
C_PDF   = colors.HexColor("#c11b1b")   # PDF 빨간색
C_HEAD  = colors.HexColor("#2d2d2d")


# ══════════════════════════════════════════════════════════════
# 예시 이미지 생성
# ══════════════════════════════════════════════════════════════

def make_img(title, lines, filename, width=760, height=380, header_bg=(0, 91, 172)):
    img = Image.new("RGB", (width, height), (248, 249, 252))
    draw = ImageDraw.Draw(img)
    try:
        fnt_t = ImageFont.truetype(BOLD,    20)
        fnt_b = ImageFont.truetype(REGULAR, 16)
        fnt_s = ImageFont.truetype(REGULAR, 13)
    except Exception:
        fnt_t = fnt_b = fnt_s = ImageFont.load_default()

    draw.rectangle([0, 0, width, 52], fill=header_bg)
    draw.text((18, 14), title, font=fnt_t, fill=(255, 255, 255))

    y = 68
    for line in lines:
        if line.startswith("##"):
            draw.rectangle([10, y-2, width-10, y+26], fill=(220, 230, 245))
            draw.text((18, y+2), line[2:].strip(), font=fnt_b, fill=header_bg)
            y += 34
        elif line.startswith(">>"):
            draw.polygon([(22, y+5),(30, y+11),(22, y+17)], fill=header_bg)
            draw.text((38, y+1), line[2:].strip(), font=fnt_b, fill=(40, 40, 40))
            y += 26
        elif line.startswith("--"):
            draw.line([26, y+9, width-26, y+9], fill=(200, 210, 225), width=1)
            y += 18
        else:
            draw.text((28, y), line, font=fnt_s, fill=(80, 80, 80))
            y += 22

    draw.rectangle([0, 0, width-1, height-1], outline=(180, 195, 220), width=2)
    path = os.path.join(OUT_DIR, filename)
    img.save(path)
    return path


img_hwp = make_img(
    "HWP (한글) – 주요 기능",
    [
        "## 한글과컴퓨터 '한글' — 국내 공문서 표준 프로그램",
        ">> 스타일 : 서식 > 스타일 설정으로 제목·본문 서식 일괄 적용",
        ">> 표 : 삽입 > 표 → 셀 합치기/나누기 → 표 속성에서 너비·정렬",
        ">> 쪽 번호 : 삽입 > 쪽 번호 → 위치·시작번호 설정",
        "--",
        "## 자주 쓰는 단축키",
        "  Ctrl+Enter   쪽 나누기       F5            블록 지정",
        "  Alt+T         표 삽입          Ctrl+K       하이퍼링크",
        "  Ctrl+N,T    새 탭 열기      Ctrl+G       찾아 바꾸기",
        "--",
        "## PDF 변환 / 저장",
        ">> 파일 > 다른 이름으로 저장 > 파일 형식 : PDF 선택",
        "  ※ 인쇄 > PDF 프린터 방식도 가능",
    ],
    "img_sk01_hwp.png",
    header_bg=(0, 91, 172)
)

img_docx = make_img(
    "DOCX (MS Word) – 주요 기능",
    [
        "## Microsoft Word — 국제 표준 문서 양식",
        ">> 스타일 : 홈 > 스타일 패널 → 제목1·2, 본문 등 클릭 적용",
        ">> 목차 자동생성 : 참조 > 목차 → 스타일 기반 자동 생성",
        ">> 검토·변경 추적 : 검토 > 변경 추적 ON → 수정 내역 색상 표시",
        "--",
        "## 자주 쓰는 단축키",
        "  Ctrl+Enter   페이지 나누기    Ctrl+L/E/R   정렬",
        "  Ctrl+Shift+N  기본 스타일     Alt+Shift+D  날짜 삽입",
        "  Ctrl+Z/Y      실행취소/재실행",
        "--",
        "## PDF 변환",
        ">> 파일 > 내보내기 > PDF/XPS 만들기",
        "  ※ 인쇄 > Microsoft Print to PDF 방식도 가능",
    ],
    "img_sk02_docx.png",
    header_bg=(43, 87, 154)
)

img_xlsx = make_img(
    "XLSX (MS Excel) – 주요 기능",
    [
        "## Microsoft Excel — 표·데이터·수식 작업 표준",
        ">> 기본 수식 : =SUM() 합계 / =AVERAGE() 평균 / =IF() 조건",
        ">> VLOOKUP : =VLOOKUP(찾는값, 범위, 열번호, 0) → 데이터 매칭",
        ">> 필터 : 데이터 > 필터 → 열 헤더 클릭으로 조건 필터링",
        "--",
        "## 자주 쓰는 단축키",
        "  Ctrl+Shift+L   필터 토글      Ctrl+T   표 서식 적용",
        "  Ctrl+1          셀 서식 창     F4       수식 절대참조 ($)",
        "  Alt+=           SUM 자동합계  Ctrl+;   오늘 날짜 입력",
        "--",
        "## 피벗테이블·PDF 변환",
        ">> 삽입 > 피벗테이블 → 행·열·값 드래그로 집계",
        ">> 파일 > 내보내기 > PDF/XPS (인쇄 영역 미리 설정 권장)",
    ],
    "img_sk03_xlsx.png",
    header_bg=(30, 113, 69)
)

img_pdf = make_img(
    "PDF – 생성 · 편집 · 병합",
    [
        "## PDF 생성 방법",
        ">> HWP : 파일 > 다른 이름으로 저장 > PDF",
        ">> Word/Excel : 파일 > 내보내기 > PDF/XPS",
        ">> 공통 : 인쇄 > PDF 프린터 (Microsoft Print to PDF)",
        "--",
        "## PDF 편집 도구",
        ">> Adobe Acrobat : 전문 편집 (텍스트·이미지 수정, 양식 작성)",
        ">> Smallpdf (웹) : 병합·분할·압축·변환 무료 제공 (smallpdf.com)",
        ">> iLovePDF (웹) : 병합·분할·회전·워터마크 등 (ilovepdf.com)",
        "--",
        "## 자주 쓰는 PDF 작업",
        "  병합 : 여러 PDF를 한 파일로 합치기",
        "  분할 : 특정 페이지만 추출",
        "  압축 : 파일 용량 줄이기 (이메일 첨부 전 권장)",
        "  변환 : PDF → Word/Excel 변환 (편집 필요 시)",
    ],
    "img_sk04_pdf.png",
    height=400,
    header_bg=(180, 30, 30)
)


# ══════════════════════════════════════════════════════════════
# 공통 헬퍼
# ══════════════════════════════════════════════════════════════

def shortcut_table_pdf(rows, color, S):
    """단축키 표 – PDF용"""
    hdr = [Paragraph(c, ParagraphStyle("sh", fontName="NanumGothicBold",
            fontSize=10, textColor=colors.white, alignment=TA_CENTER))
           for c in ["단축키", "기능", "단축키", "기능"]]
    data = [hdr]
    for i in range(0, len(rows)-1, 2):
        r1 = rows[i]; r2 = rows[i+1] if i+1 < len(rows) else ("", "")
        data.append([
            Paragraph(r1[0], S["code"]), Paragraph(r1[1], S["body"]),
            Paragraph(r2[0], S["code"]), Paragraph(r2[1], S["body"]),
        ])
    if len(rows) % 2 == 1:
        r = rows[-1]
        data.append([Paragraph(r[0], S["code"]), Paragraph(r[1], S["body"]),
                     Paragraph("", S["body"]),  Paragraph("", S["body"])])
    t = Table(data, colWidths=[3.5*cm, 4.8*cm, 3.5*cm, 4.8*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), color),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f2f4f8")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#c0cce0")),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(-1,-1), 6),
        ("FONTNAME",      (0,1),(0,-1),  "NanumGothicBold"),
        ("FONTNAME",      (2,1),(2,-1),  "NanumGothicBold"),
    ]))
    return t


# ══════════════════════════════════════════════════════════════
# PDF 생성
# ══════════════════════════════════════════════════════════════

def build_pdf(out_path):
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        topMargin=2*cm, bottomMargin=2*cm,
        leftMargin=2.2*cm, rightMargin=2.2*cm
    )

    def ps(name, **kw):
        base = dict(fontName="NanumGothic", fontSize=10.5, leading=18,
                    textColor=colors.HexColor("#333333"))
        base.update(kw)
        return ParagraphStyle(name, **base)

    S = {
        "cover_title": ParagraphStyle("ct", fontName="NanumSquareB", fontSize=26,
            textColor=C_HEAD, alignment=TA_CENTER, spaceAfter=6),
        "cover_sub":   ParagraphStyle("cs", fontName="NanumGothicBold", fontSize=14,
            textColor=colors.HexColor("#555555"), alignment=TA_CENTER),
        "h1":  ParagraphStyle("h1", fontName="NanumSquareB", fontSize=14,
            textColor=colors.white, leftIndent=8),
        "h2":  ParagraphStyle("h2", fontName="NanumGothicBold", fontSize=12,
            textColor=C_HEAD, spaceBefore=12, spaceAfter=4, leftIndent=4),
        "body":   ps("body",   leftIndent=12, spaceAfter=3),
        "bullet": ps("bullet", leftIndent=26, spaceAfter=2),
        "code":   ParagraphStyle("code", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.HexColor("#1a3a6a"), leftIndent=4),
        "note":   ParagraphStyle("note", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.HexColor("#8a3a00"), leftIndent=16, spaceAfter=3),
        "caption":ParagraphStyle("cap",  fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#666666"), alignment=TA_CENTER, spaceAfter=8),
        "footer": ParagraphStyle("ft",   fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#888888"), alignment=TA_CENTER),
        "tag":    ParagraphStyle("tag",  fontName="NanumGothicBold", fontSize=11,
            textColor=colors.white, alignment=TA_CENTER),
    }

    def sec_hdr(text, color):
        tbl = Table([[Paragraph(text, S["h1"])]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), color),
            ("TOPPADDING",    (0,0),(-1,-1), 8),
            ("BOTTOMPADDING", (0,0),(-1,-1), 8),
            ("LEFTPADDING",   (0,0),(-1,-1), 12),
        ]))
        return tbl

    def img_block(path, cap, w=16*cm):
        from reportlab.platypus import Image as RLImage
        h = w * 380/760
        if "sk04" in path: h = w * 400/760
        im = RLImage(path, width=w, height=h)
        return KeepTogether([im, Paragraph(f"▲ {cap}", S["caption"])])

    def tip_box(text, color):
        tbl = Table([[Paragraph(f"💡 {text}", ParagraphStyle("tip",
            fontName="NanumGothicBold", fontSize=10,
            textColor=color, leftIndent=6))]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), colors.HexColor("#f8f8f8")),
            ("LEFTBORDERPADDING", (0,0),(-1,-1), 0),
            ("BOX",           (0,0),(-1,-1), 0, colors.white),
            ("LINEBEFOREBEFORE",(0,0),(0,-1), 4, color),
            ("TOPPADDING",    (0,0),(-1,-1), 6),
            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
            ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ]))
        return tbl

    story = []

    # ── 표지
    story.append(Spacer(1, 3.5*cm))
    story.append(Paragraph("업무 인수인계 자료", S["cover_title"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width="70%", thickness=2, color=C_HEAD, spaceAfter=10))
    story.append(Paragraph("업무 스킬 – 문서 작성 프로그램", S["cover_sub"]))
    story.append(Spacer(1, 1*cm))

    tags = Table([[
        Paragraph("📝  HWP", S["tag"]),
        Paragraph("📄  DOCX", S["tag"]),
        Paragraph("📊  XLSX", S["tag"]),
        Paragraph("🔴  PDF", S["tag"]),
    ]], colWidths=[4.15*cm]*4)
    tags.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(0,0), C_HWP),
        ("BACKGROUND", (1,0),(1,0), C_DOCX),
        ("BACKGROUND", (2,0),(2,0), C_XLSX),
        ("BACKGROUND", (3,0),(3,0), C_PDF),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ("LEFTPADDING",   (0,0),(-1,-1), 4),
        ("RIGHTPADDING",  (0,0),(-1,-1), 4),
    ]))
    story.append(tags)
    story.append(Spacer(1, 0.8*cm))

    ci = Table([
        ["작성일",    "2026년 06월 16일"],
        ["작성 부서", "경영지원팀"],
        ["대상 프로그램", "HWP(한글) / MS Word(DOCX) / MS Excel(XLSX) / PDF"],
    ], colWidths=[3.5*cm, 12*cm])
    ci.setStyle(TableStyle([
        ("FONTNAME",     (0,0),(-1,-1), "NanumGothic"),
        ("FONTNAME",     (0,0),(0,-1),  "NanumGothicBold"),
        ("FONTSIZE",     (0,0),(-1,-1), 11),
        ("TEXTCOLOR",    (0,0),(0,-1),  C_HEAD),
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",   (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LINEBELOW",    (0,0),(-1,-2), 0.5, colors.HexColor("#cccccc")),
    ]))
    story.append(ci)
    story.append(PageBreak())

    # ══ 1. HWP ════════════════════════════════════════════════
    story.append(sec_hdr("1. HWP (한글)  –  국내 공문서 표준 프로그램", C_HWP))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "한글과컴퓨터의 <b>한글(HWP)</b>은 국내 공공기관·기업에서 공문서 작성에 가장 많이 사용하는 "
        "워드프로세서입니다. 정부 보조금 신청서, 제안서, 계획서 등 대부분의 공식 문서 양식이 HWP로 제공됩니다.",
        S["body"]))

    story.append(Paragraph("■ 자주 사용하는 핵심 기능", S["h2"]))
    hwp_feat = [
        [Paragraph(c, ParagraphStyle("fh", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER)) for c in ["기능", "사용 방법"]],
        [Paragraph("스타일 적용", S["body"]),
         Paragraph("서식 > 스타일 → 제목·본문·캡션 등 클릭 적용. 문서 전체 서식 일관성 유지", S["body"])],
        [Paragraph("표 삽입·편집", S["body"]),
         Paragraph("삽입 > 표 → 셀 합치기(M)/나누기(S) → 표 속성에서 너비·정렬·테두리 설정", S["body"])],
        [Paragraph("쪽 번호", S["body"]),
         Paragraph("삽입 > 쪽 번호 → 위치(하단 가운데 등)·시작번호 설정", S["body"])],
        [Paragraph("머리말/꼬리말", S["body"]),
         Paragraph("삽입 > 머리말/꼬리말 → 회사명·문서번호 등 반복 텍스트 설정", S["body"])],
        [Paragraph("찾아 바꾸기", S["body"]),
         Paragraph("Ctrl+H → 특정 단어·서식 일괄 변경", S["body"])],
        [Paragraph("그림 삽입", S["body"]),
         Paragraph("삽입 > 그림 → 본문 속/자리 차지/글 뒤 등 배치 방식 선택", S["body"])],
    ]
    hwp_tbl = Table(hwp_feat, colWidths=[3.5*cm, 13.1*cm])
    hwp_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), C_HWP),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef3fb")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#90aad0")),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(hwp_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 자주 쓰는 단축키", S["h2"]))
    story.append(shortcut_table_pdf([
        ("Ctrl+S",      "저장"),
        ("Ctrl+Z",      "실행취소"),
        ("Ctrl+Enter",  "쪽 나누기"),
        ("F5",          "블록 지정 시작"),
        ("Alt+T",       "표 삽입"),
        ("Ctrl+H",      "찾아 바꾸기"),
        ("Ctrl+G",      "쪽 이동"),
        ("Ctrl+K",      "하이퍼링크"),
        ("Ctrl+N,T",    "새 탭 열기"),
        ("Ctrl+F3",     "맞춤법 검사"),
    ], C_HWP, S))
    story.append(Spacer(1, 0.3*cm))

    story.append(img_block(img_hwp, "HWP 주요 기능 화면 예시"))

    story.append(Paragraph("■ PDF 변환", S["h2"]))
    for i, s in enumerate([
        "파일 > 다른 이름으로 저장 → 파일 형식에서 <b>PDF</b> 선택 후 저장",
        "또는 인쇄(Ctrl+P) > 프린터 : <b>Microsoft Print to PDF</b> 선택",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Paragraph("※ 공모사업 제출 시 PDF 변환 전 글자·표·이미지 위치 꼭 확인", S["note"]))
    story.append(PageBreak())

    # ══ 2. DOCX ═══════════════════════════════════════════════
    story.append(sec_hdr("2. DOCX (MS Word)  –  국제 표준 문서 양식", C_DOCX))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>MS Word(DOCX)</b>는 국제 표준 문서 포맷으로, 계약서·제안서·보고서 등 "
        "상대방과 주고받는 문서에 주로 사용합니다. "
        "HWP에 없는 변경 추적·공동 편집 기능이 강점입니다.",
        S["body"]))

    story.append(Paragraph("■ 자주 사용하는 핵심 기능", S["h2"]))
    docx_feat = [
        [Paragraph(c, ParagraphStyle("fh2", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER)) for c in ["기능", "사용 방법"]],
        [Paragraph("스타일 적용", S["body"]),
         Paragraph("홈 > 스타일 패널 → 제목1·제목2·본문 클릭 적용. 목차 자동 생성의 기반", S["body"])],
        [Paragraph("목차 자동 생성", S["body"]),
         Paragraph("참조 > 목차 → 자동 목차 선택. 제목 스타일 적용 필수", S["body"])],
        [Paragraph("변경 추적", S["body"]),
         Paragraph("검토 > 변경 추적 ON → 수정 내역 색상으로 표시. 계약서 검토 시 필수", S["body"])],
        [Paragraph("메모(주석)", S["body"]),
         Paragraph("검토 > 새 메모 → 특정 텍스트 선택 후 메모 삽입. 협업 시 활용", S["body"])],
        [Paragraph("페이지 나누기", S["body"]),
         Paragraph("Ctrl+Enter → 현재 커서 위치에서 새 페이지 시작", S["body"])],
        [Paragraph("표 삽입·편집", S["body"]),
         Paragraph("삽입 > 표 → 셀 병합/분할·테두리 색상·정렬 등 설정", S["body"])],
    ]
    docx_tbl = Table(docx_feat, colWidths=[3.5*cm, 13.1*cm])
    docx_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), C_DOCX),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef0f8")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#9aaad0")),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(docx_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 자주 쓰는 단축키", S["h2"]))
    story.append(shortcut_table_pdf([
        ("Ctrl+Enter",     "페이지 나누기"),
        ("Ctrl+L/E/R",     "왼쪽/가운데/오른쪽 정렬"),
        ("Ctrl+Shift+N",   "기본 스타일 적용"),
        ("Ctrl+Alt+1/2/3", "제목1/2/3 스타일"),
        ("Ctrl+Z/Y",       "실행취소/재실행"),
        ("Ctrl+F",         "찾기"),
        ("Ctrl+H",         "찾아 바꾸기"),
        ("Alt+Shift+D",    "날짜 필드 삽입"),
    ], C_DOCX, S))
    story.append(Spacer(1, 0.3*cm))

    story.append(img_block(img_docx, "MS Word(DOCX) 주요 기능 화면 예시"))

    story.append(Paragraph("■ PDF 변환", S["h2"]))
    for i, s in enumerate([
        "파일 > 내보내기 > <b>PDF/XPS 만들기</b> → 저장",
        "또는 인쇄(Ctrl+P) > 프린터 : <b>Microsoft Print to PDF</b>",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Paragraph("※ 변경 추적·메모가 있는 경우 PDF 변환 전 '모두 승인' 또는 숨기기 처리 확인", S["note"]))
    story.append(PageBreak())

    # ══ 3. XLSX ═══════════════════════════════════════════════
    story.append(sec_hdr("3. XLSX (MS Excel)  –  표·데이터·수식 작업", C_XLSX))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>MS Excel(XLSX)</b>은 급여대장·사업소득지급대장·예산표 등 "
        "숫자 데이터가 포함된 모든 문서 작업에 사용합니다. "
        "수식·함수·피벗테이블로 데이터를 빠르게 집계·분석할 수 있습니다.",
        S["body"]))

    story.append(Paragraph("■ 자주 사용하는 핵심 기능", S["h2"]))
    xlsx_feat = [
        [Paragraph(c, ParagraphStyle("fh3", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER)) for c in ["기능", "사용 방법 / 예시"]],
        [Paragraph("기본 수식", S["body"]),
         Paragraph("=SUM(B2:B10) 합계  /  =AVERAGE() 평균  /  =COUNT() 개수  /  =ROUND() 반올림", S["body"])],
        [Paragraph("IF 조건 함수", S["body"]),
         Paragraph("=IF(조건, 참일 때 값, 거짓일 때 값)  예: =IF(C2>=60,\"합격\",\"불합격\")", S["body"])],
        [Paragraph("VLOOKUP", S["body"]),
         Paragraph("=VLOOKUP(찾는값, 범위, 열번호, 0)  → 다른 표에서 데이터 자동 매칭", S["body"])],
        [Paragraph("필터·정렬", S["body"]),
         Paragraph("데이터 > 필터(Ctrl+Shift+L) → 열 헤더 ▼ 클릭으로 조건 필터링", S["body"])],
        [Paragraph("셀 서식", S["body"]),
         Paragraph("Ctrl+1 → 숫자·날짜·통화·텍스트 서식 지정. 천 단위 구분: #,##0", S["body"])],
        [Paragraph("피벗테이블", S["body"]),
         Paragraph("삽입 > 피벗테이블 → 행·열·값 필드 드래그로 대용량 데이터 집계", S["body"])],
        [Paragraph("조건부 서식", S["body"]),
         Paragraph("홈 > 조건부 서식 → 값 범위·중복 등 조건에 따라 자동 색상 표시", S["body"])],
    ]
    xlsx_tbl = Table(xlsx_feat, colWidths=[3.5*cm, 13.1*cm])
    xlsx_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), C_XLSX),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#edf8f2")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#80c0a0")),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(xlsx_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 자주 쓰는 단축키", S["h2"]))
    story.append(shortcut_table_pdf([
        ("Ctrl+Shift+L",  "필터 토글"),
        ("Ctrl+T",        "표(Table) 서식 적용"),
        ("Ctrl+1",        "셀 서식 창 열기"),
        ("F4",            "수식 절대참조($) 전환"),
        ("Alt+=",         "SUM 자동합계"),
        ("Ctrl+;",        "오늘 날짜 입력"),
        ("Ctrl+Shift+1",  "숫자 서식(천 단위)"),
        ("Ctrl+D",        "아래 셀에 복사"),
        ("Ctrl+Home",     "A1 셀로 이동"),
        ("F2",            "셀 편집 모드"),
    ], C_XLSX, S))
    story.append(Spacer(1, 0.3*cm))

    story.append(img_block(img_xlsx, "MS Excel(XLSX) 주요 기능 화면 예시"))

    story.append(Paragraph("■ PDF 변환", S["h2"]))
    for i, s in enumerate([
        "파일 > 내보내기 > <b>PDF/XPS 만들기</b> → 저장",
        "인쇄 영역 사전 설정 권장 : 페이지 레이아웃 > 인쇄 영역 > 인쇄 영역 설정",
        "인쇄 미리보기(Ctrl+P)에서 페이지 레이아웃 확인 후 변환",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Paragraph("※ 열이 많은 표는 '가로 방향'으로 페이지 설정 후 변환 권장", S["note"]))
    story.append(PageBreak())

    # ══ 4. PDF ════════════════════════════════════════════════
    story.append(sec_hdr("4. PDF  –  생성 · 편집 · 병합 · 변환", C_PDF))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>PDF</b>는 어떤 환경에서도 동일하게 보이는 문서 포맷입니다. "
        "제출·납품·공유용 최종 문서는 PDF로 변환하여 보관합니다.",
        S["body"]))

    story.append(Paragraph("■ PDF 생성 방법 요약", S["h2"]))
    gen_data = [
        [Paragraph(c, ParagraphStyle("gh", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER)) for c in ["원본 프로그램", "PDF 변환 방법"]],
        [Paragraph("HWP (한글)", S["body"]),
         Paragraph("파일 > 다른 이름으로 저장 > 파일 형식 : PDF", S["body"])],
        [Paragraph("MS Word (DOCX)", S["body"]),
         Paragraph("파일 > 내보내기 > PDF/XPS 만들기", S["body"])],
        [Paragraph("MS Excel (XLSX)", S["body"]),
         Paragraph("파일 > 내보내기 > PDF/XPS (인쇄 영역 먼저 설정)", S["body"])],
        [Paragraph("공통 방법", S["body"]),
         Paragraph("인쇄(Ctrl+P) > 프린터 : Microsoft Print to PDF 선택", S["body"])],
    ]
    gen_tbl = Table(gen_data, colWidths=[4*cm, 12.6*cm])
    gen_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), C_PDF),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#fdeeed")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#e0a0a0")),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(gen_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ PDF 편집·가공 도구", S["h2"]))
    tool_data = [
        [Paragraph(c, ParagraphStyle("th5", fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER)) for c in ["도구", "주요 기능", "접근 방법"]],
        [Paragraph("Adobe Acrobat", S["body"]),
         Paragraph("텍스트·이미지 직접 편집, 양식 작성, 전자서명", S["body"]),
         Paragraph("PC 설치 프로그램 (유료)", S["body"])],
        [Paragraph("Smallpdf", S["body"]),
         Paragraph("병합·분할·압축·Word/Excel 변환", S["body"]),
         Paragraph("smallpdf.com (무료 제한)", S["body"])],
        [Paragraph("iLovePDF", S["body"]),
         Paragraph("병합·분할·회전·워터마크·페이지 삭제", S["body"]),
         Paragraph("ilovepdf.com (무료)", S["body"])],
        [Paragraph("PDF24", S["body"]),
         Paragraph("병합·압축·변환·OCR 텍스트 인식", S["body"]),
         Paragraph("pdf24.org (무료)", S["body"])],
    ]
    tool_tbl = Table(tool_data, colWidths=[3.5*cm, 7.5*cm, 5.6*cm])
    tool_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), C_PDF),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#fdeeed")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#e0a0a0")),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(tool_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 자주 하는 PDF 작업", S["h2"]))
    for s in [
        "<b>병합</b> : 여러 PDF 파일을 하나로 합치기 → iLovePDF / Smallpdf > [PDF 병합]",
        "<b>분할</b> : 특정 페이지만 추출 → iLovePDF > [PDF 분할] > 페이지 범위 입력",
        "<b>압축</b> : 파일 용량 줄이기 → 이메일 첨부 전 권장 (Smallpdf > [PDF 압축])",
        "<b>변환</b> : PDF → Word/Excel → Smallpdf > [PDF → Word] / [PDF → Excel]",
        "<b>회전·페이지 삭제</b> : iLovePDF > [PDF 구성] → 드래그로 순서 변경·삭제",
    ]:
        story.append(Paragraph(f"• {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_pdf, "PDF 생성·편집·병합 작업 화면 예시"))
    story.append(Paragraph("※ 개인정보가 포함된 문서는 온라인 도구 업로드 전 주의", S["note"]))
    story.append(Paragraph("※ 최종 제출본은 반드시 PDF로 변환 후 드롭박스 해당 폴더에 보관", S["note"]))

    story.append(Spacer(1, 0.6*cm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#cccccc"), spaceAfter=8))
    story.append(Paragraph(
        "본 자료는 인수인계 목적으로 작성되었습니다. 프로그램 버전 업데이트 시 담당자가 내용을 업데이트하십시오.",
        S["footer"]))

    doc.build(story)
    print(f"PDF 생성 완료 : {out_path}")


# ══════════════════════════════════════════════════════════════
# DOCX 생성
# ══════════════════════════════════════════════════════════════

def set_cell_bg(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)

def ko(run):
    run.font.name = "나눔고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")

def add_sec_hdr(doc, text, bg_hex):
    tbl = doc.add_table(rows=1, cols=1); tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]; cell.text = text
    set_cell_bg(cell, bg_hex)
    p = cell.paragraphs[0]; r = p.runs[0]; ko(r)
    r.font.bold = True; r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()

def add_para(doc, text, bold=False, size=10.5, indent=0.3, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); ko(r)
    r.font.size = Pt(size); r.font.bold = bold
    if color: r.font.color.rgb = color
    return p

def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); ko(r); r.font.size = Pt(10.5)

def add_img(doc, path, caption, width_cm=14.5):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(width_cm))
    cap = doc.add_paragraph(f"▲ {caption}")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in cap.runs:
        ko(r); r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()

def note(doc, text):
    add_para(doc, text, color=RGBColor(0x8a, 0x3a, 0x00))

def feat_table(doc, rows, hdr_hex, col_widths=[3.5, 13]):
    tbl = doc.add_table(rows=1+len(rows), cols=2); tbl.style = "Table Grid"
    for ci, h in enumerate(["기능", "사용 방법"]):
        c = tbl.rows[0].cells[ci]; c.text = h
        set_cell_bg(c, hdr_hex)
        r = c.paragraphs[0].runs[0]; ko(r)
        r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.width = Cm(col_widths[ci])
    for ri, (feat, desc) in enumerate(rows):
        cells = tbl.rows[ri+1].cells
        for ci, val in enumerate([feat, desc]):
            c = cells[ci]; c.text = val
            set_cell_bg(c, "FFFFFF" if ri%2==0 else "EEF3FB")
            r = c.paragraphs[0].runs[0]; ko(r); r.font.size = Pt(10)
        cells[0].width = Cm(col_widths[0])
        cells[1].width = Cm(col_widths[1])
    doc.add_paragraph()

def shortcut_table_docx(doc, rows, hdr_hex):
    ncols = 4
    tbl = doc.add_table(rows=1 + (len(rows)+1)//2, cols=ncols)
    tbl.style = "Table Grid"
    for ci, h in enumerate(["단축키", "기능", "단축키", "기능"]):
        c = tbl.rows[0].cells[ci]; c.text = h
        set_cell_bg(c, hdr_hex)
        r = c.paragraphs[0].runs[0]; ko(r)
        r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i in range(0, len(rows), 2):
        ri = i // 2 + 1
        cells = tbl.rows[ri].cells
        for ci, idx in enumerate([i, i+1]):
            if idx < len(rows):
                k, v = rows[idx]
                c0 = cells[ci*2]; c0.text = k
                set_cell_bg(c0, "E8F0FF" if ri%2==0 else "F5F5FF")
                r = c0.paragraphs[0].runs[0]; ko(r)
                r.font.bold = True; r.font.size = Pt(10)
                c1 = cells[ci*2+1]; c1.text = v
                set_cell_bg(c1, "FFFFFF" if ri%2==0 else "F9F9FF")
                r2 = c1.paragraphs[0].runs[0]; ko(r2); r2.font.size = Pt(10)
    for ri in range(tbl.rows.__len__()):
        for ci, w in enumerate([3.2, 4.8, 3.2, 4.8]):
            tbl.rows[ri].cells[ci].width = Cm(w)
    doc.add_paragraph()


def build_docx(out_path):
    doc = Document()
    for sec in doc.sections:
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
        sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)

    # 표지
    for _ in range(5): doc.add_paragraph()
    tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("업무 인수인계 자료"); ko(tr)
    tr.font.size = Pt(26); tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x2d, 0x2d, 0x2d)

    sp = doc.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run("업무 스킬 – 문서 작성 프로그램"); ko(sr)
    sr.font.size = Pt(15); sr.font.bold = True
    sr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    doc.add_paragraph()

    ci_tbl = doc.add_table(rows=3, cols=2); ci_tbl.style = "Table Grid"
    for ri, (k, v) in enumerate([
        ("작성일", "2026년 06월 16일"),
        ("작성 부서", "경영지원팀"),
        ("대상 프로그램", "HWP(한글) / MS Word(DOCX) / MS Excel(XLSX) / PDF"),
    ]):
        cells = ci_tbl.rows[ri].cells
        set_cell_bg(cells[0], "dce6f8"); cells[0].text = k; cells[1].text = v
        for ci_, c in enumerate(cells):
            r = c.paragraphs[0].runs[0]; ko(r); r.font.size = Pt(10.5)
            if ci_ == 0: r.font.bold = True
        cells[0].width = Cm(3.5); cells[1].width = Cm(12)
    doc.add_page_break()

    # ── 1. HWP
    add_sec_hdr(doc, "1. HWP (한글)  –  국내 공문서 표준 프로그램", "005bac")
    add_para(doc,
        "한글과컴퓨터의 한글(HWP)은 국내 공공기관·기업에서 공문서 작성에 가장 많이 사용하는 "
        "워드프로세서입니다. 정부 보조금 신청서·제안서·계획서 등 대부분의 공식 문서 양식이 HWP로 제공됩니다.")
    add_para(doc, "■ 자주 사용하는 핵심 기능", bold=True)
    feat_table(doc, [
        ("스타일 적용",    "서식 > 스타일 → 제목·본문·캡션 클릭 적용. 문서 전체 서식 일관성 유지"),
        ("표 삽입·편집",   "삽입 > 표 → 셀 합치기(M)/나누기(S) → 표 속성에서 너비·정렬·테두리"),
        ("쪽 번호",        "삽입 > 쪽 번호 → 위치(하단 가운데 등)·시작번호 설정"),
        ("머리말/꼬리말",  "삽입 > 머리말/꼬리말 → 회사명·문서번호 등 반복 텍스트 설정"),
        ("찾아 바꾸기",    "Ctrl+H → 특정 단어·서식 일괄 변경"),
        ("그림 삽입",      "삽입 > 그림 → 본문 속/자리 차지/글 뒤 등 배치 방식 선택"),
    ], "005bac")
    add_para(doc, "■ 자주 쓰는 단축키", bold=True)
    shortcut_table_docx(doc, [
        ("Ctrl+S",     "저장"),             ("Ctrl+Z",    "실행취소"),
        ("Ctrl+Enter", "쪽 나누기"),        ("F5",        "블록 지정"),
        ("Alt+T",      "표 삽입"),          ("Ctrl+H",    "찾아 바꾸기"),
        ("Ctrl+G",     "쪽 이동"),          ("Ctrl+K",    "하이퍼링크"),
        ("Ctrl+N,T",   "새 탭"),            ("Ctrl+F3",   "맞춤법 검사"),
    ], "005bac")
    add_img(doc, img_hwp, "HWP 주요 기능 화면 예시")
    add_para(doc, "■ PDF 변환", bold=True)
    for i, s in enumerate([
        "파일 > 다른 이름으로 저장 → 파일 형식 : PDF 선택 후 저장",
        "또는 인쇄(Ctrl+P) > 프린터 : Microsoft Print to PDF",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    note(doc, "※ 공모사업 제출 시 PDF 변환 전 글자·표·이미지 위치 꼭 확인")
    doc.add_page_break()

    # ── 2. DOCX
    add_sec_hdr(doc, "2. DOCX (MS Word)  –  국제 표준 문서 양식", "2b579a")
    add_para(doc,
        "MS Word(DOCX)는 국제 표준 문서 포맷으로, 계약서·제안서·보고서 등 "
        "상대방과 주고받는 문서에 주로 사용합니다. 변경 추적·공동 편집 기능이 강점입니다.")
    add_para(doc, "■ 자주 사용하는 핵심 기능", bold=True)
    feat_table(doc, [
        ("스타일 적용",    "홈 > 스타일 패널 → 제목1·2·본문 클릭 적용. 목차 자동 생성의 기반"),
        ("목차 자동 생성", "참조 > 목차 → 자동 목차 선택. 제목 스타일 적용 필수"),
        ("변경 추적",      "검토 > 변경 추적 ON → 수정 내역 색상 표시. 계약서 검토 시 필수"),
        ("메모(주석)",     "검토 > 새 메모 → 텍스트 선택 후 삽입. 협업 시 활용"),
        ("페이지 나누기",  "Ctrl+Enter → 현재 커서 위치에서 새 페이지 시작"),
        ("표 삽입·편집",   "삽입 > 표 → 셀 병합/분할·테두리 색상·정렬 설정"),
    ], "2b579a")
    add_para(doc, "■ 자주 쓰는 단축키", bold=True)
    shortcut_table_docx(doc, [
        ("Ctrl+Enter",     "페이지 나누기"),    ("Ctrl+L/E/R",  "왼쪽/가운데/오른쪽 정렬"),
        ("Ctrl+Shift+N",   "기본 스타일"),       ("Ctrl+Alt+1",  "제목1 스타일"),
        ("Ctrl+Z/Y",       "취소/재실행"),       ("Ctrl+H",      "찾아 바꾸기"),
        ("Ctrl+F",         "찾기"),              ("Alt+Shift+D", "날짜 삽입"),
    ], "2b579a")
    add_img(doc, img_docx, "MS Word(DOCX) 주요 기능 화면 예시")
    add_para(doc, "■ PDF 변환", bold=True)
    for i, s in enumerate([
        "파일 > 내보내기 > PDF/XPS 만들기 → 저장",
        "또는 인쇄(Ctrl+P) > 프린터 : Microsoft Print to PDF",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    note(doc, "※ 변경 추적·메모가 있는 경우 PDF 변환 전 '모두 승인' 처리 확인")
    doc.add_page_break()

    # ── 3. XLSX
    add_sec_hdr(doc, "3. XLSX (MS Excel)  –  표·데이터·수식 작업", "1e7145")
    add_para(doc,
        "MS Excel(XLSX)은 급여대장·사업소득지급대장·예산표 등 "
        "숫자 데이터가 포함된 모든 문서 작업에 사용합니다. "
        "수식·함수·피벗테이블로 데이터를 빠르게 집계·분석할 수 있습니다.")
    add_para(doc, "■ 자주 사용하는 핵심 기능", bold=True)
    feat_table(doc, [
        ("기본 수식",      "=SUM() 합계 / =AVERAGE() 평균 / =COUNT() 개수 / =ROUND() 반올림"),
        ("IF 조건 함수",   "=IF(조건, 참값, 거짓값)  예: =IF(C2>=60,\"합격\",\"불합격\")"),
        ("VLOOKUP",        "=VLOOKUP(찾는값, 범위, 열번호, 0) → 다른 표에서 데이터 자동 매칭"),
        ("필터·정렬",      "데이터 > 필터(Ctrl+Shift+L) → 열 헤더 ▼ 클릭으로 조건 필터링"),
        ("셀 서식",        "Ctrl+1 → 숫자·날짜·통화·텍스트 서식 지정. 천 단위 구분: #,##0"),
        ("피벗테이블",     "삽입 > 피벗테이블 → 행·열·값 필드 드래그로 대용량 데이터 집계"),
        ("조건부 서식",    "홈 > 조건부 서식 → 값 범위·중복 등 조건에 따라 자동 색상 표시"),
    ], "1e7145")
    add_para(doc, "■ 자주 쓰는 단축키", bold=True)
    shortcut_table_docx(doc, [
        ("Ctrl+Shift+L", "필터 토글"),          ("Ctrl+T",       "표 서식 적용"),
        ("Ctrl+1",       "셀 서식 창"),          ("F4",           "절대참조($)"),
        ("Alt+=",        "SUM 자동합계"),        ("Ctrl+;",       "오늘 날짜"),
        ("Ctrl+D",       "아래 셀에 복사"),      ("F2",           "셀 편집 모드"),
        ("Ctrl+Home",    "A1 셀 이동"),          ("Ctrl+Shift+1", "천 단위 서식"),
    ], "1e7145")
    add_img(doc, img_xlsx, "MS Excel(XLSX) 주요 기능 화면 예시")
    add_para(doc, "■ PDF 변환", bold=True)
    for i, s in enumerate([
        "파일 > 내보내기 > PDF/XPS 만들기 → 저장",
        "인쇄 영역 사전 설정 권장 : 페이지 레이아웃 > 인쇄 영역 > 인쇄 영역 설정",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    note(doc, "※ 열이 많은 표는 '가로 방향' 페이지 설정 후 변환 권장")
    doc.add_page_break()

    # ── 4. PDF
    add_sec_hdr(doc, "4. PDF  –  생성 · 편집 · 병합 · 변환", "c11b1b")
    add_para(doc,
        "PDF는 어떤 환경에서도 동일하게 보이는 문서 포맷입니다. "
        "제출·납품·공유용 최종 문서는 PDF로 변환하여 보관합니다.")
    add_para(doc, "■ PDF 생성 방법 요약", bold=True)

    gen_tbl = doc.add_table(rows=5, cols=2); gen_tbl.style = "Table Grid"
    for ci, h in enumerate(["원본 프로그램", "PDF 변환 방법"]):
        c = gen_tbl.rows[0].cells[ci]; c.text = h
        set_cell_bg(c, "c11b1b")
        r = c.paragraphs[0].runs[0]; ko(r)
        r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for ri, (prog, method) in enumerate([
        ("HWP (한글)",       "파일 > 다른 이름으로 저장 > 파일 형식 : PDF"),
        ("MS Word (DOCX)",   "파일 > 내보내기 > PDF/XPS 만들기"),
        ("MS Excel (XLSX)",  "파일 > 내보내기 > PDF/XPS (인쇄 영역 먼저 설정)"),
        ("공통 방법",        "인쇄(Ctrl+P) > 프린터 : Microsoft Print to PDF"),
    ]):
        cells = gen_tbl.rows[ri+1].cells
        cells[0].text = prog; cells[1].text = method
        set_cell_bg(cells[0], "FFFFFF" if ri%2==0 else "FDEEED")
        set_cell_bg(cells[1], "FFFFFF" if ri%2==0 else "FDEEED")
        for ci_, c in enumerate(cells):
            r = c.paragraphs[0].runs[0]; ko(r); r.font.size = Pt(10)
        cells[0].width = Cm(4); cells[1].width = Cm(12.5)
    doc.add_paragraph()

    add_para(doc, "■ PDF 편집·가공 도구", bold=True)
    tool_tbl = doc.add_table(rows=5, cols=3); tool_tbl.style = "Table Grid"
    for ci, h in enumerate(["도구", "주요 기능", "접근 방법"]):
        c = tool_tbl.rows[0].cells[ci]; c.text = h
        set_cell_bg(c, "c11b1b")
        r = c.paragraphs[0].runs[0]; ko(r)
        r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for ri, (tool, feat, access) in enumerate([
        ("Adobe Acrobat", "텍스트·이미지 편집, 양식 작성, 전자서명", "PC 설치 프로그램 (유료)"),
        ("Smallpdf",      "병합·분할·압축·Word/Excel 변환",          "smallpdf.com (무료 제한)"),
        ("iLovePDF",      "병합·분할·회전·워터마크·페이지 삭제",     "ilovepdf.com (무료)"),
        ("PDF24",         "병합·압축·변환·OCR 텍스트 인식",          "pdf24.org (무료)"),
    ]):
        cells = tool_tbl.rows[ri+1].cells
        for ci_, val in enumerate([tool, feat, access]):
            c = cells[ci_]; c.text = val
            set_cell_bg(c, "FFFFFF" if ri%2==0 else "FDEEED")
            r = c.paragraphs[0].runs[0]; ko(r); r.font.size = Pt(10)
        cells[0].width = Cm(3.5); cells[1].width = Cm(7.5); cells[2].width = Cm(5.5)
    doc.add_paragraph()

    add_para(doc, "■ 자주 하는 PDF 작업", bold=True)
    for s in [
        "병합 : 여러 PDF를 하나로 합치기 → iLovePDF / Smallpdf > [PDF 병합]",
        "분할 : 특정 페이지만 추출 → iLovePDF > [PDF 분할] > 페이지 범위 입력",
        "압축 : 파일 용량 줄이기 → 이메일 첨부 전 권장 (Smallpdf > [PDF 압축])",
        "변환 : PDF → Word/Excel → Smallpdf > [PDF → Word] / [PDF → Excel]",
        "회전·페이지 삭제 : iLovePDF > [PDF 구성] → 드래그로 순서 변경·삭제",
    ]:
        add_bullet(doc, f"• {s}")
    doc.add_paragraph()
    add_img(doc, img_pdf, "PDF 생성·편집·병합 작업 화면 예시")
    note(doc, "※ 개인정보가 포함된 문서는 온라인 도구 업로드 전 주의")
    note(doc, "※ 최종 제출본은 반드시 PDF로 변환 후 드롭박스 해당 폴더에 보관")

    doc.add_paragraph()
    fp = doc.add_paragraph(
        "본 자료는 인수인계 목적으로 작성되었습니다. 프로그램 버전 업데이트 시 담당자가 내용을 업데이트하십시오.")
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in fp.runs:
        ko(r); r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.save(out_path)
    print(f"DOCX 생성 완료 : {out_path}")


if __name__ == "__main__":
    pdf_path  = os.path.join(OUT_DIR, "업무스킬_문서작성_인수인계.pdf")
    docx_path = os.path.join(OUT_DIR, "업무스킬_문서작성_인수인계.docx")
    build_pdf(pdf_path)
    build_docx(docx_path)
    print("모든 파일 생성 완료.")
