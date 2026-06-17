"""
AI 툴 활용 업무 인수인계 자료
tools: 캔바, 마누스, 젠스파크, 제미나이, 챗지피티, 클로드, 어도비
"""

import os
from PIL import Image, ImageDraw, ImageFont

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether, Image as RLImage,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_DIR = "/usr/share/fonts/truetype/nanum"
REG  = os.path.join(FONT_DIR, "NanumGothic.ttf")
BOLD = os.path.join(FONT_DIR, "NanumGothicBold.ttf")
SQB  = os.path.join(FONT_DIR, "NanumSquareB.ttf")

pdfmetrics.registerFont(TTFont("NG",     REG))
pdfmetrics.registerFont(TTFont("NGBold", BOLD))
pdfmetrics.registerFont(TTFont("NSB",    SQB))

OUT_DIR = "/home/user/project-1/handover_docs"
IMG_DIR = OUT_DIR
os.makedirs(OUT_DIR, exist_ok=True)

W, _ = A4

# ── 툴별 브랜드 컬러 ─────────────────────────────────────────────────────────
TOOLS = [
    {
        "id":    "canva",
        "name":  "캔바 (Canva)",
        "color": "#7c3aed",   # 보라
        "emoji": "🎨",
        "when":  [
            "SNS 카드뉴스·썸네일·포스터 등 디자인 결과물이 필요할 때",
            "템플릿을 활용해 빠르게 홍보물을 만들어야 할 때",
            "프레젠테이션 슬라이드를 깔끔하게 제작할 때",
            "디자이너 없이 팀원이 직접 시각 자료를 만들어야 할 때",
        ],
        "tips":  [
            "원하는 결과물 유형(인스타 카드뉴스, A4 포스터 등)을 먼저 선택",
            "회사 브랜드 컬러·폰트를 '브랜드 키트'에 저장해두면 통일감 유지",
            "AI 이미지 생성(Text to Image) 기능으로 배경·소품 이미지 제작 가능",
            "완성 후 PNG·PDF 두 가지로 내보내 보관",
        ],
        "caution": "무료 플랜은 일부 요소·폰트 유료. 사내 유료 계정 사용 확인 필요.",
    },
    {
        "id":    "manus",
        "name":  "마누스 (Manus)",
        "color": "#0d7a5f",   # 초록
        "emoji": "🤖",
        "when":  [
            "복잡한 다단계 업무(리서치 → 정리 → 문서 작성)를 자동화할 때",
            "웹 검색·파일 생성·코드 실행까지 한 번에 처리해야 할 때",
            "반복적인 데이터 수집·정리 작업을 AI에게 위임할 때",
            "긴 보고서·제안서 초안을 자동으로 만들고 싶을 때",
        ],
        "tips":  [
            "원하는 최종 결과물을 명확하게 지시 (예: 'A4 2장 분량의 보고서로 작성해줘')",
            "단계별로 지시하기보다 최종 목표를 한 번에 설명하는 것이 효과적",
            "결과물 검토 후 수정 요청으로 다듬기 — 초안으로 활용",
            "민감한 사내 정보는 입력하지 않도록 주의",
        ],
        "caution": "자율 에이전트이므로 결과물은 반드시 사람이 검토 후 사용.",
    },
    {
        "id":    "genspark",
        "name":  "젠스파크 (Genspark)",
        "color": "#b45309",   # 앰버
        "emoji": "⚡",
        "when":  [
            "특정 주제에 대해 빠르고 깊이 있는 리서치가 필요할 때",
            "시장 동향·경쟁사·트렌드 분석 자료를 빠르게 모아야 할 때",
            "공모사업 배경 조사나 사업계획서 근거 자료 수집 시",
            "신뢰도 높은 출처 기반의 요약 정보가 필요할 때",
        ],
        "tips":  [
            "검색어보다 질문 형태로 입력 (예: '2024년 한국 문화콘텐츠 수출 현황은?')",
            "출처 링크를 확인해 중요한 내용은 원문 검증 후 사용",
            "여러 관점을 비교하고 싶을 때 '찬반 양쪽 의견을 정리해줘' 활용",
            "리서치 결과를 ChatGPT·Claude로 넘겨 문서화하면 효율적",
        ],
        "caution": "AI 요약이므로 사실 관계는 반드시 원문 출처에서 재확인.",
    },
    {
        "id":    "gemini",
        "name":  "제미나이 (Gemini)",
        "color": "#1a73e8",   # 구글 블루
        "emoji": "💎",
        "when":  [
            "구글 워크스페이스(Gmail·Docs·Sheets)와 연동해 작업할 때",
            "이미지·PDF·엑셀 파일을 업로드하고 내용을 분석·요약할 때",
            "긴 문서를 빠르게 요약하거나 번역이 필요할 때",
            "영어 이메일·보고서 작성 및 교정이 필요할 때",
        ],
        "tips":  [
            "파일을 직접 업로드하면 내용 기반으로 질문·요약 가능",
            "구글 계정 로그인 상태에서 사용 시 구글 드라이브 파일 참조 가능",
            "'한국어로 답해줘' 지시를 앞에 붙이면 한국어 결과물 정확도 향상",
            "긴 텍스트 번역은 ChatGPT보다 Gemini가 자연스러운 경우 많음",
        ],
        "caution": "회사 기밀 문서는 업로드 전 정보 보안 정책 확인 필요.",
    },
    {
        "id":    "chatgpt",
        "name":  "챗지피티 (ChatGPT)",
        "color": "#10a37f",   # OpenAI 그린
        "emoji": "💬",
        "when":  [
            "글쓰기·보고서·기획서 초안 작성 및 다듬기가 필요할 때",
            "아이디어 브레인스토밍이나 방향성 검토가 필요할 때",
            "한국어·영어 번역, 이메일·공문서 작성 시",
            "엑셀 수식·함수 작성 방법이나 간단한 코드 도움이 필요할 때",
        ],
        "tips":  [
            "역할 지정 후 요청 (예: '너는 마케팅 전문가야, 아래 내용으로 SNS 문구를 작성해줘')",
            "결과가 마음에 안 들면 '더 간결하게', '공식적인 톤으로' 등 수정 요청",
            "긴 내용은 한꺼번에 붙여넣고 '요약해줘' / '핵심만 추려줘' 활용",
            "문서 작성 시 '초안을 만들어줘 → 검토 → 수정 요청' 반복이 효율적",
        ],
        "caution": "학습 데이터 활용 방지를 위해 설정 > 데이터 제어에서 학습 끄기 권장.",
    },
    {
        "id":    "claude",
        "name":  "클로드 (Claude)",
        "color": "#c4531d",   # 주황
        "emoji": "🧠",
        "when":  [
            "긴 문서·계약서·보고서를 정밀하게 분석하거나 검토할 때",
            "논리적 구조가 중요한 제안서·기획서 작성 시",
            "세밀한 맥락 이해가 필요한 복잡한 업무 처리 시",
            "코드 작성·데이터 분석·문서 자동화 작업 시 (Claude Code)",
        ],
        "tips":  [
            "배경 설명을 충분히 제공할수록 맥락에 맞는 결과 출력",
            "긴 계약서·문서도 통째로 붙여넣고 '이 중 위험한 조항 있어?' 식으로 활용",
            "'단계별로 설명해줘', '표 형태로 정리해줘' 등 출력 형식 지정 효과적",
            "ChatGPT와 함께 사용 — 초안은 ChatGPT, 정교한 검토·수정은 Claude",
        ],
        "caution": "최신 실시간 정보는 학습 데이터 기준이므로 최신 트렌드는 Genspark 병행.",
    },
    {
        "id":    "adobe",
        "name":  "어도비 AI (Adobe)",
        "color": "#ee0000",   # 어도비 레드
        "emoji": "✨",
        "when":  [
            "사진 보정·배경 제거·이미지 편집이 필요할 때 (Photoshop AI)",
            "텍스트로 고품질 이미지를 생성해야 할 때 (Firefly)",
            "PDF 문서를 편집하거나 AI 요약이 필요할 때 (Acrobat AI)",
            "영상 편집에서 자동 자막·배경 제거 등이 필요할 때 (Premiere Pro AI)",
        ],
        "tips":  [
            "Photoshop의 '생성형 채우기'로 이미지 일부를 AI로 채우거나 지우기 가능",
            "Firefly는 저작권 걱정 없는 상업용 이미지 생성에 활용 가능",
            "Acrobat AI Assistant로 긴 PDF를 요약하고 질문 답변 가능",
            "결과물의 해상도와 파일 형식(PNG·SVG·PDF)을 용도에 맞게 선택",
        ],
        "caution": "어도비 크리에이티브 클라우드 구독 계정 필요. 플랜 확인 후 사용.",
    },
]


# ── PIL helpers ──────────────────────────────────────────────────────────────
def pil_font(size, bold=False):
    try:
        return ImageFont.truetype(BOLD if bold else REG, size)
    except Exception:
        return ImageFont.load_default()

def make_tool_img(path, tool):
    W_IMG, H_IMG = 760, 320
    color = tool["color"]
    img = Image.new("RGB", (W_IMG, H_IMG), "#f8f9fc")
    d   = ImageDraw.Draw(img)

    # header
    d.rectangle([0, 0, W_IMG, 56], fill=color)
    d.text((18, 12), f"{tool['emoji']}  {tool['name']}", font=pil_font(22, bold=True), fill="white")

    # two columns
    col_x = [18, W_IMG // 2 + 10]
    titles = ["💡 이럴 때 사용하세요", "✅ 잘 쓰는 법"]
    contents = [tool["when"], tool["tips"]]

    for i, (cx, title, items) in enumerate(zip(col_x, titles, contents)):
        y = 70
        # column title
        d.rectangle([cx, y, cx + W_IMG // 2 - 28, y + 26], fill=color)
        d.text((cx + 8, y + 4), title, font=pil_font(14, bold=True), fill="white")
        y += 34
        for item in items[:4]:
            # bullet
            d.ellipse([cx + 4, y + 6, cx + 12, y + 14], fill=color)
            # wrap text
            words = item
            d.text((cx + 18, y), words, font=pil_font(13), fill="#333333")
            y += 24
            if y > H_IMG - 20:
                break

    # bottom note
    note_text = f"⚠  {tool['caution']}"
    d.rectangle([0, H_IMG - 38, W_IMG, H_IMG], fill="#fff3cd")
    d.text((12, H_IMG - 28), note_text, font=pil_font(12), fill="#856404")
    img.save(path)


def make_overview_img(path):
    """전체 AI 툴 한눈에 보기 이미지"""
    W_IMG, H_IMG = 760, 400
    img = Image.new("RGB", (W_IMG, H_IMG), "#f8f9fc")
    d   = ImageDraw.Draw(img)

    d.rectangle([0, 0, W_IMG, 56], fill="#1a3a6b")
    d.text((18, 14), "사내 AI 툴 한눈에 보기", font=pil_font(22, bold=True), fill="white")

    rows = [
        ("캔바",   "#7c3aed", "디자인·시각 자료"),
        ("마누스",  "#0d7a5f", "복잡한 자동화 업무"),
        ("젠스파크", "#b45309", "리서치·정보 수집"),
        ("제미나이", "#1a73e8", "파일 분석·번역"),
        ("챗지피티", "#10a37f", "글쓰기·브레인스토밍"),
        ("클로드",  "#c4531d", "문서 검토·정밀 분석"),
        ("어도비",  "#ee0000", "이미지·영상·PDF 편집"),
    ]
    y = 72
    for name, color, desc in rows:
        d.rectangle([18, y, 18 + 14, y + 22], fill=color)
        d.text((40, y + 2), name, font=pil_font(15, bold=True), fill="#1a3a6b")
        d.text((160, y + 2), "→", font=pil_font(15), fill="#888888")
        d.text((190, y + 2), desc, font=pil_font(15), fill="#333333")
        y += 32
        if y > H_IMG - 20:
            break

    img.save(path)


# ══════════════════════════════════════════════════════════════════════════════
#  PDF
# ══════════════════════════════════════════════════════════════════════════════
def S(name, **kw):
    base = dict(fontName="NG", fontSize=11, leading=18, textColor=colors.HexColor("#333333"))
    base.update(kw)
    return ParagraphStyle(name, **base)

BRAND   = "#1a3a6b"
SEC_HDR = S("sec_hdr", fontName="NSB", fontSize=15, spaceBefore=14, spaceAfter=4,
            textColor=colors.HexColor(BRAND))
BODY    = S("body",   spaceBefore=2, spaceAfter=2)
BULLET  = S("bul",   leftIndent=14, spaceBefore=1, spaceAfter=1)
NOTE_S  = S("note",  fontName="NG", fontSize=10, textColor=colors.HexColor("#856404"),
            backColor=colors.HexColor("#fff3cd"), leftIndent=10, rightIndent=10,
            spaceBefore=4, spaceAfter=4)

def hr():
    return HRFlowable(width="100%", thickness=1, color=colors.HexColor("#c8d4e8"), spaceAfter=6)

def sec_hdr(txt, color=BRAND):
    return [Paragraph(f'<font color="{color}">{txt}</font>', SEC_HDR), hr()]

def para(txt):
    return Paragraph(txt, BODY)

def bul(txt):
    return Paragraph(f"• &nbsp;{txt}", BULLET)

def note(txt):
    return Paragraph(f"⚠  {txt}", NOTE_S)

def add_img_pdf(path, w_mm=155):
    return RLImage(path, width=w_mm*mm, height=w_mm*mm*(320/760))

def mk_table(data, col_w, hdr_color=BRAND):
    t = Table(data, colWidths=col_w)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), colors.HexColor(hdr_color)),
        ("TEXTCOLOR",     (0,0), (-1,0), colors.white),
        ("FONTNAME",      (0,0), (-1,0), "NGBold"),
        ("FONTSIZE",      (0,0), (-1,-1), 10),
        ("FONTNAME",      (0,1), (-1,-1), "NG"),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4fb")]),
        ("GRID",          (0,0), (-1,-1), 0.4, colors.HexColor("#b0bfd0")),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
    ]))
    return t


def build_pdf(out_path):
    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=20*mm, rightMargin=20*mm,
                            topMargin=18*mm, bottomMargin=18*mm)
    E = []
    CW = W - 40*mm

    # ── cover ──
    E.append(Spacer(1, 28*mm))
    E.append(Paragraph("AI 툴 활용 가이드", S("t1", fontName="NSB", fontSize=28,
                        textColor=colors.HexColor(BRAND), alignment=1)))
    E.append(Spacer(1, 4*mm))
    E.append(Paragraph("인수인계 자료", S("t2", fontName="NSB", fontSize=22,
                        textColor=colors.HexColor("#3a7bd5"), alignment=1)))
    E.append(Spacer(1, 8*mm))
    E.append(HRFlowable(width="60%", thickness=2, color=colors.HexColor(BRAND), hAlign="CENTER"))
    E.append(Spacer(1, 6*mm))
    E.append(Paragraph("ONNAMU", S("co", fontName="NGBold", fontSize=13,
                        textColor=colors.HexColor("#555555"), alignment=1)))
    E.append(Spacer(1, 10*mm))

    # overview image on cover
    ov_path = f"{IMG_DIR}/img_ai00_overview.png"
    make_overview_img(ov_path)
    E.append(add_img_pdf(ov_path))
    E.append(Spacer(1, 8*mm))

    # 목차 표
    toc_data = [
        ["툴", "주요 활용 용도"],
        ["캔바 (Canva)",    "SNS 디자인·홍보물·프레젠테이션 제작"],
        ["마누스 (Manus)",  "복잡한 다단계 업무 자동화·보고서 초안"],
        ["젠스파크 (Genspark)", "리서치·트렌드 분석·출처 기반 정보 수집"],
        ["제미나이 (Gemini)",   "파일 분석·번역·구글 워크스페이스 연동"],
        ["챗지피티 (ChatGPT)",  "글쓰기·브레인스토밍·번역·수식 작성"],
        ["클로드 (Claude)",     "문서 검토·정밀 분석·기획서·코드 작업"],
        ["어도비 AI (Adobe)",   "이미지 편집·Firefly 생성·PDF 요약"],
    ]
    E.append(mk_table(toc_data, [45*mm, 120*mm]))
    E.append(PageBreak())

    # ── 공통 프롬프트 팁 ──
    E += sec_hdr("AI 툴 공통 활용 원칙")
    tip_data = [
        ["원칙", "설명"],
        ["역할 지정",   "AI에게 역할을 먼저 지정하면 결과 품질이 높아집니다\n예) '너는 한국 문화콘텐츠 전문가야, 아래 내용을 분석해줘'"],
        ["목적 명확화", "원하는 결과물의 형태·분량·톤을 구체적으로 지시\n예) 'A4 1장 분량, 공식적인 톤으로, 표 포함해서 작성해줘'"],
        ["단계적 수정", "첫 결과가 마음에 안 들면 '더 간결하게', '예시 추가해줘' 등\n추가 요청으로 다듬기 — 한 번에 완벽할 필요 없음"],
        ["검토 필수",   "AI 결과물은 사실 오류·맥락 오해 가능성 있음\n중요 내용은 반드시 사람이 검토 후 사용"],
        ["보안 주의",   "사내 기밀·개인정보·계약 내용은 외부 AI에 입력 자제\n필요 시 민감 정보 제거 후 입력"],
    ]
    E.append(mk_table(tip_data, [28*mm, 137*mm]))
    E.append(Spacer(1, 6*mm))

    # ── 상황별 추천 툴 ──
    E += sec_hdr("상황별 추천 AI 툴")
    rec_data = [
        ["상황", "추천 툴", "이유"],
        ["SNS 콘텐츠·카드뉴스 제작",  "캔바",          "템플릿·디자인 특화"],
        ["공모사업 배경 리서치",       "젠스파크",       "출처 기반 정보 수집"],
        ["제안서·기획서 초안 작성",    "챗지피티 / 클로드", "글쓰기·구조 설계"],
        ["긴 계약서·문서 검토",        "클로드",         "긴 문맥 정밀 분석"],
        ["이미지·사진 편집",          "어도비 AI / 캔바", "전문 편집 기능"],
        ["다국어 번역·영문 교정",      "제미나이 / 챗지피티", "자연스러운 번역"],
        ["복잡한 자료 자동 정리",      "마누스",         "자율 에이전트 처리"],
        ["파일·PDF 업로드 후 분석",    "제미나이 / 클로드", "파일 기반 QA"],
        ["AI 이미지 생성(저작권 안전)", "어도비 Firefly", "상업용 라이선스"],
    ]
    E.append(mk_table(rec_data, [52*mm, 42*mm, 71*mm]))
    E.append(PageBreak())

    # ── 각 툴 상세 ──
    for tool in TOOLS:
        img_path = f"{IMG_DIR}/img_ai_{tool['id']}.png"
        make_tool_img(img_path, tool)

        E += sec_hdr(f"{tool['emoji']}  {tool['name']}", color=tool["color"])

        # 언제 사용
        when_data = [["💡 이럴 때 사용하세요"]] + [[f"• {w}"] for w in tool["when"]]
        t_when = Table(when_data, colWidths=[CW])
        t_when.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), colors.HexColor(tool["color"])),
            ("TEXTCOLOR",     (0,0), (-1,0), colors.white),
            ("FONTNAME",      (0,0), (-1,0), "NGBold"),
            ("FONTSIZE",      (0,0), (-1,-1), 10),
            ("FONTNAME",      (0,1), (-1,-1), "NG"),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.HexColor("#f8f9fc")]),
            ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#d0d8e8")),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ]))
        E.append(t_when)
        E.append(Spacer(1, 3*mm))

        # 잘 쓰는 법
        tips_data = [["✅ 잘 쓰는 법 (프롬프트 팁)"]] + [[f"• {t}"] for t in tool["tips"]]
        t_tips = Table(tips_data, colWidths=[CW])
        t_tips.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#2d5a9e")),
            ("TEXTCOLOR",     (0,0), (-1,0), colors.white),
            ("FONTNAME",      (0,0), (-1,0), "NGBold"),
            ("FONTSIZE",      (0,0), (-1,-1), 10),
            ("FONTNAME",      (0,1), (-1,-1), "NG"),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.HexColor("#f0f4fb")]),
            ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#d0d8e8")),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ]))
        E.append(t_tips)
        E.append(Spacer(1, 3*mm))

        E.append(note(tool["caution"]))
        E.append(Spacer(1, 3*mm))
        E.append(add_img_pdf(img_path))
        E.append(PageBreak())

    doc.build(E)
    print(f"PDF 생성 완료 : {out_path}")


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX helpers
# ══════════════════════════════════════════════════════════════════════════════
def ko(run):
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), "나눔고딕")
    rPr.insert(0, rFonts)
    return run

def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color.lstrip("#"))
    tcPr.append(shd)

def add_sec_hdr_d(doc, txt, hex_color="1a3a6b"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(txt)
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(*bytes.fromhex(hex_color.lstrip("#")))
    ko(r)
    doc.add_paragraph("─" * 52).paragraph_format.space_after = Pt(4)

def add_table_d(doc, data, col_widths, hdr_hex="1a3a6b"):
    table = doc.add_table(rows=len(data), cols=len(data[0]))
    table.style = "Table Grid"
    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx]
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Cm(col_widths[c_idx])
            p    = cell.paragraphs[0]
            run  = p.add_run(str(cell_text))
            run.font.size = Pt(10)
            if r_idx == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                set_cell_bg(cell, hdr_hex)
            else:
                ko(run)
                if r_idx % 2 == 0:
                    set_cell_bg(cell, "f0f4fb")
    return table

def add_img_d(doc, path, width_cm=16):
    try:
        doc.add_picture(path, width=Cm(width_cm))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass

def add_note_d(doc, txt):
    p   = doc.add_paragraph()
    run = p.add_run(f"⚠  {txt}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x85, 0x64, 0x04)
    ko(run)

def hex_to_rgb(hex_str):
    h = hex_str.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def build_docx(out_path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2.5)

    # cover
    doc.add_paragraph()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("AI 툴 활용 가이드 인수인계 자료")
    r.bold = True; r.font.size = Pt(22)
    r.font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b); ko(r)
    doc.add_paragraph()
    t2 = doc.add_paragraph(); t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = t2.add_run("ONNAMU"); r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55); ko(r2)
    doc.add_paragraph()
    add_img_d(doc, f"{IMG_DIR}/img_ai00_overview.png")
    doc.add_page_break()

    # 공통 원칙
    add_sec_hdr_d(doc, "AI 툴 공통 활용 원칙")
    add_table_d(doc,
        [["원칙", "설명"],
         ["역할 지정",   "AI에게 역할을 먼저 지정하면 결과 품질이 높아집니다\n예) '너는 한국 문화콘텐츠 전문가야, 아래 내용을 분석해줘'"],
         ["목적 명확화", "원하는 결과물의 형태·분량·톤을 구체적으로 지시\n예) 'A4 1장 분량, 공식적인 톤으로, 표 포함해서 작성해줘'"],
         ["단계적 수정", "첫 결과가 마음에 안 들면 '더 간결하게', '예시 추가해줘' 등 추가 요청으로 다듬기"],
         ["검토 필수",   "AI 결과물은 사실 오류·맥락 오해 가능성 있음. 중요 내용은 반드시 사람이 검토 후 사용"],
         ["보안 주의",   "사내 기밀·개인정보·계약 내용은 외부 AI에 입력 자제. 필요 시 민감 정보 제거 후 입력"]],
        [3.0, 13.5])
    doc.add_paragraph()

    # 상황별 추천
    add_sec_hdr_d(doc, "상황별 추천 AI 툴")
    add_table_d(doc,
        [["상황", "추천 툴", "이유"],
         ["SNS 콘텐츠·카드뉴스 제작",  "캔바",          "템플릿·디자인 특화"],
         ["공모사업 배경 리서치",       "젠스파크",       "출처 기반 정보 수집"],
         ["제안서·기획서 초안 작성",    "챗지피티 / 클로드", "글쓰기·구조 설계"],
         ["긴 계약서·문서 검토",        "클로드",         "긴 문맥 정밀 분석"],
         ["이미지·사진 편집",          "어도비 AI / 캔바", "전문 편집 기능"],
         ["다국어 번역·영문 교정",      "제미나이 / 챗지피티", "자연스러운 번역"],
         ["복잡한 자료 자동 정리",      "마누스",         "자율 에이전트 처리"],
         ["파일·PDF 업로드 후 분석",    "제미나이 / 클로드", "파일 기반 QA"],
         ["AI 이미지 생성(저작권 안전)", "어도비 Firefly", "상업용 라이선스"]],
        [5.5, 4.5, 6.5])
    doc.add_page_break()

    # 각 툴 상세
    for tool in TOOLS:
        img_path = f"{IMG_DIR}/img_ai_{tool['id']}.png"
        r, g, b = hex_to_rgb(tool["color"])

        add_sec_hdr_d(doc, f"{tool['emoji']}  {tool['name']}", tool["color"])

        # 이럴 때 사용
        p = doc.add_paragraph()
        rh = p.add_run("💡 이럴 때 사용하세요")
        rh.bold = True; rh.font.size = Pt(11)
        rh.font.color.rgb = RGBColor(r, g, b); ko(rh)

        for w in tool["when"]:
            bp = doc.add_paragraph(style="List Bullet")
            br = bp.add_run(w)
            br.font.size = Pt(10); ko(br)

        doc.add_paragraph()

        # 잘 쓰는 법
        p2 = doc.add_paragraph()
        rh2 = p2.add_run("✅ 잘 쓰는 법 (프롬프트 팁)")
        rh2.bold = True; rh2.font.size = Pt(11)
        rh2.font.color.rgb = RGBColor(0x2d, 0x5a, 0x9e); ko(rh2)

        for tip in tool["tips"]:
            bp2 = doc.add_paragraph(style="List Bullet")
            br2 = bp2.add_run(tip)
            br2.font.size = Pt(10); ko(br2)

        doc.add_paragraph()
        add_note_d(doc, tool["caution"])
        doc.add_paragraph()
        add_img_d(doc, img_path)
        doc.add_page_break()

    doc.save(out_path)
    print(f"DOCX 생성 완료 : {out_path}")


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    build_pdf(f"{OUT_DIR}/AI툴활용_업무인수인계.pdf")
    build_docx(f"{OUT_DIR}/AI툴활용_업무인수인계.docx")
    print("모든 파일 생성 완료.")
