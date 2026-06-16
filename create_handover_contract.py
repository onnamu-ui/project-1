#!/usr/bin/env python3
"""
계약 및 기타 문서 발급 업무 인수인계 자료
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

NAVY   = colors.HexColor("#1e3a6e")
BLUE   = colors.HexColor("#1e4fa0")
GREEN  = colors.HexColor("#1a7a40")
TEAL   = colors.HexColor("#0a6a7a")
ORANGE = colors.HexColor("#c05800")
PURPLE = colors.HexColor("#5a1a8a")
GRAY   = colors.HexColor("#4a4a4a")

# ══════════════════════════════════════════════════════════════
# 예시 이미지 생성
# ══════════════════════════════════════════════════════════════

def make_img(title, lines, filename, width=760, height=360,
             header_bg=(30, 62, 110)):
    img = Image.new("RGB", (width, height), (245, 248, 252))
    draw = ImageDraw.Draw(img)
    try:
        fnt_t = ImageFont.truetype(BOLD,    20)
        fnt_b = ImageFont.truetype(REGULAR, 16)
        fnt_s = ImageFont.truetype(REGULAR, 13)
    except Exception:
        fnt_t = fnt_b = fnt_s = ImageFont.load_default()

    draw.rectangle([0, 0, width, 50], fill=header_bg)
    draw.text((18, 13), title, font=fnt_t, fill=(255, 255, 255))

    y = 68
    for line in lines:
        if line.startswith("##"):
            draw.rectangle([10, y-2, width-10, y+24], fill=(218, 230, 248))
            draw.text((18, y), line[2:].strip(), font=fnt_b, fill=(20, 50, 120))
            y += 32
        elif line.startswith(">>"):
            draw.rectangle([24, y+3, 32, y+16], fill=header_bg)
            draw.text((42, y), line[2:].strip(), font=fnt_b, fill=(40, 40, 40))
            y += 26
        elif line.startswith("--"):
            draw.line([28, y+8, width-28, y+8], fill=(200, 210, 230), width=1)
            y += 16
        else:
            draw.text((24, y), line, font=fnt_s, fill=(80, 80, 80))
            y += 22

    draw.rectangle([0, 0, width-1, height-1], outline=(180, 200, 230), width=2)
    path = os.path.join(OUT_DIR, filename)
    img.save(path)
    return path


img_narajangteo = make_img(
    "나라장터 계약 진행 – G2B (g2b.go.kr)",
    [
        "## 나라장터 전자조달시스템 로그인 (공동인증서)",
        ">> 계약관리 > 계약현황 > 해당 계약건 조회",
        "--",
        "## 주요 업무 흐름",
        ">> 낙찰 통보 확인 → 계약서 작성 → 계약 체결 → 납품/이행",
        ">> 납품 완료 후 : 대가청구 > 기성·준공 검사 > 대금 지급",
        "--",
        "  ※ 계약보증금 납부 : 계약금액의 10% (보증서 또는 현금)",
        "  ※ 계약 체결 후 인감증명서·사업자등록증 등 서류 제출 필요",
        "  ※ 전자계약 가능 시 전자서명으로 체결 (공동인증서 필수)",
    ],
    "img_c01_narajangteo.png",
    header_bg=(20, 60, 110)
)

img_modusign = make_img(
    "모두싸인 전자계약 서명 – modusign.co.kr",
    [
        "## 모두싸인 로그인 > 수신된 서명 요청 확인",
        ">> 이메일 또는 알림으로 서명 요청 수신",
        "--",
        "## 서명 절차",
        ">> [서명하기] 클릭 → 계약서 내용 검토",
        ">> 서명란 클릭 → 서명 입력 (직접 서명 또는 도장 이미지)",
        ">> [완료] 클릭 → 상대방 서명 완료 후 최종 계약서 PDF 발송",
        "--",
        "  ※ 서명 전 계약서 내용(금액·기간·조건) 반드시 확인",
        "  ※ 완료된 계약서 PDF는 '계약서/YYYY' 폴더에 보관",
        "  ※ 서명 요청 발송 : 문서관리 > 새 문서 > 서명 요청",
    ],
    "img_c02_modusign.png",
    header_bg=(60, 30, 110)
)

img_bizreg = make_img(
    "사업자등록증 발급 – 홈택스",
    [
        "## 홈택스 > 사업자등록 > 사업자등록증 재발급",
        ">> 사업장 선택 후 [즉시발급] → PDF 저장",
        "--",
        "  ※ 사업자등록증 변경사항(주소·대표자 등) 있을 시 정정 후 재발급",
        "  ※ 별도 유효기간 없음 (최신 정보 반영본으로 발급 권장)",
        "  ※ 정부24(gov.kr)에서도 발급 가능",
    ],
    "img_c03_bizreg.png",
    header_bg=(20, 100, 60)
)

img_corpregister = make_img(
    "법인등기부등본 발급 – 인터넷등기소",
    [
        "## 인터넷등기소 (iros.go.kr) > 열람/발급 > 법인",
        ">> 상호 또는 등록번호로 검색 → 법인등기부등본 선택",
        "--",
        ">> 발급 종류 선택 : 전부사항 / 현재사항 (제출처 확인 후 선택)",
        ">> [발급] → PDF 저장  |  수수료 : 건당 700원 (전자발급)",
        "--",
        "  ※ 말소사항 포함 여부 : 제출처에 따라 상이 → 사전 확인",
        "  ※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)",
    ],
    "img_c04_corpregister.png",
    header_bg=(100, 50, 20)
)

img_seal = make_img(
    "인감증명서 발급 – 법인인감 (등기소 방문)",
    [
        "## 발급 방법 : 관할 등기소 직접 방문",
        ">> 준비물 : 법인인감도장 + 법인등기사항증명서 + 대표자 신분증",
        "--",
        ">> 창구에서 '법인 인감증명서 발급 신청서' 작성 후 제출",
        ">> 수수료 : 건당 600원",
        "--",
        "  ※ 온라인 발급 불가 (법인인감증명서는 반드시 등기소 방문)",
        "  ※ 위임 발급 가능 : 위임장(인감날인) + 대리인 신분증 지참",
        "  ※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)",
    ],
    "img_c05_seal.png",
    header_bg=(110, 20, 20)
)

img_insured = make_img(
    "4대보험 가입자명부 발급 – 4insure.or.kr",
    [
        "## 4대사회보험 정보연계센터 > 사업장 업무 > 가입자명부 발급",
        ">> 사업장 선택 → [발급] → PDF 저장",
        "--",
        "  ※ 건강·국민연금·고용·산재 가입 직원 전체 목록 확인 가능",
        "  ※ 입찰·계약·금융기관 제출 등에 활용",
        "  ※ 개별 공단(건강보험공단 등)에서도 별도 발급 가능",
        "  ※ 유효기간 : 발급일로부터 30일",
    ],
    "img_c06_insured.png",
    header_bg=(0, 80, 100)
)

img_useseal = make_img(
    "사용인감계 작성",
    [
        "## 사용인감계란?",
        ">> 법인인감 대신 사용할 별도 인감(사용인감)을 신고하는 서류",
        "--",
        "## 작성 방법",
        ">> 사용인감계 양식에 사용인감 날인",
        ">> 법인인감으로 날인 및 법인명·대표자명 기재",
        ">> 법인인감증명서 첨부하여 제출처에 함께 제출",
        "--",
        "  ※ 사용인감계 양식은 사내 서식함 또는 제출처 지정 양식 사용",
        "  ※ 제출 시 법인인감증명서 반드시 함께 첨부",
    ],
    "img_c07_useseal.png",
    header_bg=(60, 60, 20)
)


# ══════════════════════════════════════════════════════════════
# PDF 생성
# ══════════════════════════════════════════════════════════════

def build_pdf(out_path):
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        topMargin=2*cm, bottomMargin=2*cm,
        leftMargin=2.2*cm, rightMargin=2.2*cm
    )

    S = {
        "cover_title": ParagraphStyle("cover_title",
            fontName="NanumSquareB", fontSize=26,
            textColor=NAVY, alignment=TA_CENTER, spaceAfter=6),
        "cover_sub": ParagraphStyle("cover_sub",
            fontName="NanumGothicBold", fontSize=14,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER, spaceAfter=4),
        "h1": ParagraphStyle("h1",
            fontName="NanumSquareB", fontSize=14,
            textColor=colors.white, leftIndent=8),
        "h2": ParagraphStyle("h2",
            fontName="NanumGothicBold", fontSize=12,
            textColor=NAVY, spaceBefore=12, spaceAfter=4, leftIndent=4),
        "body": ParagraphStyle("body",
            fontName="NanumGothic", fontSize=10.5,
            leading=18, textColor=colors.HexColor("#333333"),
            spaceAfter=3, leftIndent=12),
        "bullet": ParagraphStyle("bullet",
            fontName="NanumGothic", fontSize=10.5,
            leading=18, textColor=colors.HexColor("#333333"),
            leftIndent=24, spaceAfter=2),
        "note": ParagraphStyle("note",
            fontName="NanumGothicBold", fontSize=10,
            textColor=colors.HexColor("#b04000"),
            leftIndent=16, spaceAfter=3),
        "caption": ParagraphStyle("caption",
            fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#666666"),
            alignment=TA_CENTER, spaceAfter=8),
        "footer": ParagraphStyle("footer",
            fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#888888"),
            alignment=TA_CENTER),
    }

    def sec_hdr(text, color=NAVY):
        tbl = Table([[Paragraph(text, S["h1"])]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), color),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
            ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ]))
        return tbl

    def divider(text, color=BLUE):
        tbl = Table([[Paragraph(f"◆  {text}", ParagraphStyle("div",
            fontName="NanumGothicBold", fontSize=11,
            textColor=colors.white, leftIndent=6))]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), color),
            ("TOPPADDING",    (0,0),(-1,-1), 5),
            ("BOTTOMPADDING", (0,0),(-1,-1), 5),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ]))
        return tbl

    def img_block(path, caption_text, w=16*cm):
        from reportlab.platypus import Image as RLImage
        im = RLImage(path, width=w, height=w * 360/760)
        cap = Paragraph(f"▲ {caption_text}", S["caption"])
        return KeepTogether([im, cap])

    story = []

    # ── 표지
    story.append(Spacer(1, 4*cm))
    story.append(Paragraph("업무 인수인계 자료", S["cover_title"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width="70%", thickness=2, color=NAVY, spaceAfter=10))
    story.append(Paragraph("계약 업무 및 기타 문서 발급", S["cover_sub"]))
    story.append(Spacer(1, 0.8*cm))

    ci = Table([
        ["작성일", "2026년 06월 16일"],
        ["작성 부서", "경영지원팀"],
        ["인수인계 범위",
         "나라장터 계약 / 모두싸인 전자계약\n"
         "사업자등록증 / 법인등기부등본 / 인감증명서 / 4대보험 가입자명부 / 사용인감계"],
    ], colWidths=[3.5*cm, 12*cm])
    ci.setStyle(TableStyle([
        ("FONTNAME",     (0,0),(-1,-1), "NanumGothic"),
        ("FONTNAME",     (0,0),(0,-1),  "NanumGothicBold"),
        ("FONTSIZE",     (0,0),(-1,-1), 11),
        ("TEXTCOLOR",    (0,0),(0,-1),  NAVY),
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",   (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LINEBELOW",    (0,0),(-1,-2), 0.5, colors.HexColor("#cccccc")),
    ]))
    story.append(ci)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # PART 1 : 계약 업무
    # ══════════════════════════════════════════════════════════
    story.append(divider("PART 1.  계약 업무", color=NAVY))
    story.append(Spacer(1, 0.4*cm))

    # 1. 나라장터
    story.append(sec_hdr("1. 나라장터 계약 진행 (G2B)", color=BLUE))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>나라장터</b>(g2b.go.kr)는 정부·공공기관의 전자조달 시스템입니다. "
        "공동인증서로 로그인하여 계약 체결·납품·대금청구 등을 처리합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("■ 계약 체결 절차", S["h2"]))
    for i, s in enumerate([
        "나라장터(g2b.go.kr) 공동인증서 로그인",
        "계약관리 &gt; 계약현황에서 해당 계약건 조회",
        "계약서 내용 검토 후 전자서명 (공동인증서)",
        "계약보증금 납부 확인 (계약금액의 10%, 보증서 또는 현금)",
        "계약 체결 완료 → 관련 서류(사업자등록증·인감증명서 등) 제출",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Paragraph("■ 납품 및 대금청구 절차", S["h2"]))
    for i, s in enumerate([
        "납품·용역 이행 완료 후 납품서(검수요청서) 등록",
        "발주기관 검사 완료 확인",
        "대가청구 &gt; 기성·준공대가 청구 등록",
        "세금계산서 발행 후 대금 수령 확인",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_narajangteo, "나라장터(G2B) 계약 진행 화면 예시"))
    story.append(Paragraph("※ 계약 체결 후 계약서 PDF 저장 → '나라장터계약/YYYY' 폴더 보관", S["note"]))
    story.append(Paragraph("※ 입찰·계약 관련 공문서는 나라장터 내 문서함에서 확인", S["note"]))
    story.append(PageBreak())

    # 2. 모두싸인
    story.append(sec_hdr("2. 모두싸인 전자계약 서명", color=PURPLE))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>모두싸인</b>(modusign.co.kr)은 전자계약 서명 플랫폼입니다. "
        "이메일 또는 알림으로 서명 요청이 수신되면 내용을 확인하고 서명합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("■ 서명 요청 수신 시 처리 절차", S["h2"]))
    for i, s in enumerate([
        "이메일 또는 모두싸인 알림으로 서명 요청 확인",
        "[서명하기] 클릭 → 계약서 전체 내용 검토 (금액·기간·조건 확인)",
        "서명란 클릭 → 서명 입력 (직접 서명 또는 도장 이미지 적용)",
        "[완료] 클릭 → 상대방 서명 완료 후 최종 계약서 PDF 자동 발송",
        "완료된 계약서 PDF 저장 → '계약서/YYYY' 폴더 보관",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Paragraph("■ 서명 요청 발송 시 절차", S["h2"]))
    for i, s in enumerate([
        "모두싸인 로그인 &gt; 문서관리 &gt; 새 문서 &gt; 파일 업로드",
        "서명자 정보(이름·이메일) 입력 후 서명란 위치 지정",
        "[서명 요청 발송] → 상대방 서명 완료 알림 수신 후 확인",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_modusign, "모두싸인 전자계약 서명 화면 예시"))
    story.append(Paragraph("※ 서명 전 반드시 계약서 내용 전체 검토 후 서명", S["note"]))
    story.append(Paragraph("※ 완료된 계약서는 모두싸인 > 문서함에서도 재다운로드 가능", S["note"]))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # PART 2 : 기타 문서 발급
    # ══════════════════════════════════════════════════════════
    story.append(divider("PART 2.  기타 문서 발급", color=colors.HexColor("#2a5a2a")))
    story.append(Spacer(1, 0.4*cm))

    # 3. 사업자등록증
    story.append(sec_hdr("3. 사업자등록증 발급", color=GREEN))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>홈택스</b>(hometax.go.kr)에서 사업자등록증을 온라인으로 즉시 발급할 수 있습니다.",
        S["body"]))
    for i, s in enumerate([
        "홈택스 로그인",
        "사업자등록 &gt; 사업자등록증 재발급 선택",
        "사업장 선택 후 [즉시발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_bizreg, "홈택스 사업자등록증 발급 화면 예시"))
    story.append(Paragraph("※ 별도 유효기간 없음 (주소·대표자 변경 시 정정 후 최신본 발급)", S["note"]))
    story.append(Paragraph("※ 정부24(gov.kr)에서도 발급 가능", S["note"]))
    story.append(PageBreak())

    # 4. 법인등기부등본
    story.append(sec_hdr("4. 법인등기부등본 발급", color=ORANGE))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>인터넷등기소</b>(iros.go.kr)에서 온라인으로 발급합니다. "
        "수수료는 전자발급 기준 건당 700원입니다.",
        S["body"]))
    for i, s in enumerate([
        "인터넷등기소(iros.go.kr) 접속 &gt; 열람/발급 &gt; 법인 선택",
        "상호 또는 등록번호로 검색",
        "발급 종류 선택 : <b>전부사항</b>(전체) 또는 <b>현재사항</b>(현재 유효 정보만)",
        "[발급] 클릭 → 수수료 결제(700원) → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_corpregister, "인터넷등기소 법인등기부등본 발급 화면 예시"))
    story.append(Paragraph("※ 말소사항 포함 여부는 제출처에 사전 확인 후 선택", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)", S["note"]))
    story.append(PageBreak())

    # 5. 인감증명서
    story.append(sec_hdr("5. 인감증명서 발급", color=colors.HexColor("#8a1a1a")))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "법인인감증명서는 온라인 발급이 불가합니다. <b>관할 등기소를 직접 방문</b>하여 발급합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("■ 준비물", S["h2"]))
    for s in ["법인인감도장", "법인등기사항증명서", "대표자 신분증"]:
        story.append(Paragraph(f"• {s}", S["bullet"]))

    story.append(Paragraph("■ 발급 절차", S["h2"]))
    for i, s in enumerate([
        "관할 등기소 방문",
        "창구에서 '법인 인감증명서 발급 신청서' 작성·제출",
        "법인인감도장 날인",
        "수수료 납부(건당 600원) 후 수령",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_seal, "인감증명서 발급 절차 안내 예시"))
    story.append(Paragraph("※ 대리 발급 가능 : 위임장(법인인감 날인) + 대리인 신분증 지참", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)", S["note"]))
    story.append(PageBreak())

    # 6. 4대보험 가입자명부
    story.append(sec_hdr("6. 4대보험 가입자명부 발급", color=TEAL))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "<b>4대사회보험 정보연계센터</b>(4insure.or.kr)에서 사업장에 가입된 직원 전체의 "
        "4대보험 가입 현황을 한 번에 발급합니다.",
        S["body"]))
    for i, s in enumerate([
        "4대사회보험 정보연계센터(4insure.or.kr) 로그인",
        "사업장 업무 &gt; 가입자명부 발급 메뉴 선택",
        "사업장 선택 후 [발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_insured, "4대보험 가입자명부 발급 화면 예시"))
    story.append(Paragraph("※ 건강·국민·고용·산재 전 직원 가입 현황 통합 확인 가능", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 30일", S["note"]))
    story.append(PageBreak())

    # 7. 사용인감계
    story.append(sec_hdr("7. 사용인감계 작성", color=colors.HexColor("#5a5a10")))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "사용인감계는 법인인감 대신 별도 인감(사용인감)을 사용하겠다고 제출처에 신고하는 서류입니다. "
        "계약·입찰 등에서 법인인감 대신 사용인감을 사용할 때 함께 제출합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("■ 작성 방법", S["h2"]))
    for i, s in enumerate([
        "사용인감계 양식 준비 (사내 서식함 또는 제출처 지정 양식)",
        "사용인감란에 사용인감 날인",
        "법인명·대표자명·사업자등록번호 기재",
        "법인인감으로 날인",
        "법인인감증명서를 첨부하여 제출처에 함께 제출",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_useseal, "사용인감계 작성 예시"))
    story.append(Paragraph("※ 제출 시 법인인감증명서 반드시 함께 첨부", S["note"]))
    story.append(Paragraph("※ 사용인감계 양식은 제출처마다 상이할 수 있으므로 사전 확인", S["note"]))

    story.append(Spacer(1, 0.6*cm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#cccccc"), spaceAfter=8))
    story.append(Paragraph(
        "본 자료는 인수인계 목적으로 작성되었습니다. 법령 개정 또는 사내 규정 변경 시 담당자가 업데이트하십시오.",
        S["footer"]))

    doc.build(story)
    print(f"PDF 생성 완료 : {out_path}")


# ══════════════════════════════════════════════════════════════
# DOCX 생성
# ══════════════════════════════════════════════════════════════

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#"))
    tcPr.append(shd)

def ko(run):
    run.font.name = "나눔고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")

def add_sec_hdr(doc, text, bg="1e3a6e"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    cell.text = text
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    r = p.runs[0]; ko(r)
    r.font.bold = True; r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()

def add_divider(doc, text, bg="1e3a6e"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    cell.text = f"◆  {text}"
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    r = p.runs[0]; ko(r)
    r.font.bold = True; r.font.size = Pt(11)
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
    r = p.add_run(text); ko(r)
    r.font.size = Pt(10.5)

def add_img(doc, path, caption, width_cm=14.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(width_cm))
    cap = doc.add_paragraph(f"▲ {caption}")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in cap.runs:
        ko(r); r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    doc.add_paragraph()

def note(doc, text):
    add_para(doc, text, color=RGBColor(0xB0, 0x40, 0x00))


def build_docx(out_path):
    doc = Document()
    for sec in doc.sections:
        sec.page_width  = Cm(21); sec.page_height = Cm(29.7)
        sec.top_margin    = Cm(2); sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5); sec.right_margin  = Cm(2.5)

    # 표지
    for _ in range(5):
        doc.add_paragraph()
    tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("업무 인수인계 자료"); ko(tr)
    tr.font.size = Pt(26); tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x1E, 0x3A, 0x6E)

    sp = doc.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run("계약 업무 및 기타 문서 발급"); ko(sr)
    sr.font.size = Pt(15); sr.font.bold = True
    sr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()
    ci_tbl = doc.add_table(rows=3, cols=2); ci_tbl.style = "Table Grid"
    ci_data = [
        ("작성일", "2026년 06월 16일"),
        ("작성 부서", "경영지원팀"),
        ("인수인계 범위",
         "나라장터 계약 / 모두싸인 전자계약\n"
         "사업자등록증 / 법인등기부등본 / 인감증명서 / 4대보험 가입자명부 / 사용인감계"),
    ]
    for ri, (k, v) in enumerate(ci_data):
        cells = ci_tbl.rows[ri].cells
        set_cell_bg(cells[0], "dce6f8")
        cells[0].text = k; cells[1].text = v
        for ci_, c in enumerate(cells):
            r = c.paragraphs[0].runs[0]; ko(r)
            r.font.size = Pt(10.5)
            if ci_ == 0: r.font.bold = True
        cells[0].width = Cm(3.5); cells[1].width = Cm(12)

    doc.add_page_break()

    # PART 1
    add_divider(doc, "PART 1.  계약 업무", bg="1e3a6e")

    # 1. 나라장터
    add_sec_hdr(doc, "1. 나라장터 계약 진행 (G2B)", bg="1e4fa0")
    add_para(doc,
        "나라장터(g2b.go.kr)는 정부·공공기관의 전자조달 시스템입니다. "
        "공동인증서로 로그인하여 계약 체결·납품·대금청구 등을 처리합니다.")
    add_para(doc, "■ 계약 체결 절차", bold=True)
    for i, s in enumerate([
        "나라장터(g2b.go.kr) 공동인증서 로그인",
        "계약관리 > 계약현황에서 해당 계약건 조회",
        "계약서 내용 검토 후 전자서명 (공동인증서)",
        "계약보증금 납부 확인 (계약금액의 10%, 보증서 또는 현금)",
        "계약 체결 완료 → 관련 서류(사업자등록증·인감증명서 등) 제출",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    add_para(doc, "■ 납품 및 대금청구 절차", bold=True)
    for i, s in enumerate([
        "납품·용역 이행 완료 후 납품서(검수요청서) 등록",
        "발주기관 검사 완료 확인",
        "대가청구 > 기성·준공대가 청구 등록",
        "세금계산서 발행 후 대금 수령 확인",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_narajangteo, "나라장터(G2B) 계약 진행 화면 예시")
    note(doc, "※ 계약 체결 후 계약서 PDF 저장 → '나라장터계약/YYYY' 폴더 보관")
    note(doc, "※ 입찰·계약 관련 공문서는 나라장터 내 문서함에서 확인")
    doc.add_page_break()

    # 2. 모두싸인
    add_sec_hdr(doc, "2. 모두싸인 전자계약 서명", bg="5a1a8a")
    add_para(doc,
        "모두싸인(modusign.co.kr)은 전자계약 서명 플랫폼입니다. "
        "이메일 또는 알림으로 서명 요청이 수신되면 내용을 확인하고 서명합니다.")
    add_para(doc, "■ 서명 요청 수신 시 처리 절차", bold=True)
    for i, s in enumerate([
        "이메일 또는 모두싸인 알림으로 서명 요청 확인",
        "[서명하기] 클릭 → 계약서 전체 내용 검토 (금액·기간·조건 확인)",
        "서명란 클릭 → 서명 입력 (직접 서명 또는 도장 이미지 적용)",
        "[완료] 클릭 → 상대방 서명 완료 후 최종 계약서 PDF 자동 발송",
        "완료된 계약서 PDF 저장 → '계약서/YYYY' 폴더 보관",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    add_para(doc, "■ 서명 요청 발송 시 절차", bold=True)
    for i, s in enumerate([
        "모두싸인 로그인 > 문서관리 > 새 문서 > 파일 업로드",
        "서명자 정보(이름·이메일) 입력 후 서명란 위치 지정",
        "[서명 요청 발송] → 상대방 서명 완료 알림 수신 후 확인",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_modusign, "모두싸인 전자계약 서명 화면 예시")
    note(doc, "※ 서명 전 반드시 계약서 내용 전체 검토 후 서명")
    note(doc, "※ 완료된 계약서는 모두싸인 > 문서함에서도 재다운로드 가능")
    doc.add_page_break()

    # PART 2
    add_divider(doc, "PART 2.  기타 문서 발급", bg="2a5a2a")

    # 3. 사업자등록증
    add_sec_hdr(doc, "3. 사업자등록증 발급", bg="1a7a40")
    add_para(doc, "홈택스(hometax.go.kr)에서 즉시 발급합니다.")
    for i, s in enumerate([
        "홈택스 로그인",
        "사업자등록 > 사업자등록증 재발급 선택",
        "사업장 선택 후 [즉시발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_bizreg, "홈택스 사업자등록증 발급 화면 예시")
    note(doc, "※ 별도 유효기간 없음 (주소·대표자 변경 시 정정 후 최신본 발급)")
    note(doc, "※ 정부24(gov.kr)에서도 발급 가능")
    doc.add_page_break()

    # 4. 법인등기부등본
    add_sec_hdr(doc, "4. 법인등기부등본 발급", bg="c05800")
    add_para(doc,
        "인터넷등기소(iros.go.kr)에서 온라인으로 발급합니다. "
        "수수료는 전자발급 기준 건당 700원입니다.")
    for i, s in enumerate([
        "인터넷등기소(iros.go.kr) 접속 > 열람/발급 > 법인 선택",
        "상호 또는 등록번호로 검색",
        "발급 종류 선택 : 전부사항(전체) 또는 현재사항(현재 유효 정보만)",
        "[발급] 클릭 → 수수료 결제(700원) → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_corpregister, "인터넷등기소 법인등기부등본 발급 화면 예시")
    note(doc, "※ 말소사항 포함 여부는 제출처에 사전 확인 후 선택")
    note(doc, "※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)")
    doc.add_page_break()

    # 5. 인감증명서
    add_sec_hdr(doc, "5. 인감증명서 발급", bg="8a1a1a")
    add_para(doc,
        "법인인감증명서는 온라인 발급이 불가합니다. 관할 등기소를 직접 방문하여 발급합니다.")
    add_para(doc, "■ 준비물", bold=True)
    for s in ["법인인감도장", "법인등기사항증명서", "대표자 신분증"]:
        add_bullet(doc, f"• {s}")
    add_para(doc, "■ 발급 절차", bold=True)
    for i, s in enumerate([
        "관할 등기소 방문",
        "창구에서 '법인 인감증명서 발급 신청서' 작성·제출",
        "법인인감도장 날인",
        "수수료 납부(건당 600원) 후 수령",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_seal, "인감증명서 발급 절차 안내 예시")
    note(doc, "※ 대리 발급 가능 : 위임장(법인인감 날인) + 대리인 신분증 지참")
    note(doc, "※ 유효기간 : 발급일로부터 3개월 (제출처별 상이)")
    doc.add_page_break()

    # 6. 4대보험 가입자명부
    add_sec_hdr(doc, "6. 4대보험 가입자명부 발급", bg="0a6a7a")
    add_para(doc,
        "4대사회보험 정보연계센터(4insure.or.kr)에서 사업장 전체 직원의 "
        "4대보험 가입 현황을 한 번에 발급합니다.")
    for i, s in enumerate([
        "4대사회보험 정보연계센터(4insure.or.kr) 로그인",
        "사업장 업무 > 가입자명부 발급 메뉴 선택",
        "사업장 선택 후 [발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_insured, "4대보험 가입자명부 발급 화면 예시")
    note(doc, "※ 건강·국민·고용·산재 전 직원 가입 현황 통합 확인 가능")
    note(doc, "※ 유효기간 : 발급일로부터 30일")
    doc.add_page_break()

    # 7. 사용인감계
    add_sec_hdr(doc, "7. 사용인감계 작성", bg="5a5a10")
    add_para(doc,
        "사용인감계는 법인인감 대신 별도 인감(사용인감)을 사용하겠다고 제출처에 신고하는 서류입니다. "
        "계약·입찰 등에서 법인인감 대신 사용인감을 사용할 때 함께 제출합니다.")
    add_para(doc, "■ 작성 방법", bold=True)
    for i, s in enumerate([
        "사용인감계 양식 준비 (사내 서식함 또는 제출처 지정 양식)",
        "사용인감란에 사용인감 날인",
        "법인명·대표자명·사업자등록번호 기재",
        "법인인감으로 날인",
        "법인인감증명서를 첨부하여 제출처에 함께 제출",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_useseal, "사용인감계 작성 예시")
    note(doc, "※ 제출 시 법인인감증명서 반드시 함께 첨부")
    note(doc, "※ 사용인감계 양식은 제출처마다 상이할 수 있으므로 사전 확인")

    doc.add_paragraph()
    fp = doc.add_paragraph(
        "본 자료는 인수인계 목적으로 작성되었습니다. 법령 개정 또는 사내 규정 변경 시 담당자가 업데이트하십시오.")
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in fp.runs:
        ko(r); r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.save(out_path)
    print(f"DOCX 생성 완료 : {out_path}")


if __name__ == "__main__":
    pdf_path  = os.path.join(OUT_DIR, "계약및문서발급_업무인수인계.pdf")
    docx_path = os.path.join(OUT_DIR, "계약및문서발급_업무인수인계.docx")
    build_pdf(pdf_path)
    build_docx(docx_path)
    print("모든 파일 생성 완료.")
