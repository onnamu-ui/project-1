#!/usr/bin/env python3
"""
회계·세무 업무 인수인계 자료 생성 스크립트 (간소화 버전)
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

# ── 폰트 등록
FONT_DIR = "/usr/share/fonts/truetype/nanum"
REGULAR  = os.path.join(FONT_DIR, "NanumGothic.ttf")
BOLD     = os.path.join(FONT_DIR, "NanumGothicBold.ttf")
SQUARE_B = os.path.join(FONT_DIR, "NanumSquareB.ttf")

pdfmetrics.registerFont(TTFont("NanumGothic",     REGULAR))
pdfmetrics.registerFont(TTFont("NanumGothicBold", BOLD))
pdfmetrics.registerFont(TTFont("NanumSquareB",    SQUARE_B))

OUT_DIR = "/home/user/project-1/handover_docs"
os.makedirs(OUT_DIR, exist_ok=True)

BLUE   = colors.HexColor("#1e4fa0")
GREEN  = colors.HexColor("#1a7a40")
ORANGE = colors.HexColor("#e07000")
PURPLE = colors.HexColor("#6a1a8a")
TEAL   = colors.HexColor("#0a6a7a")

# ══════════════════════════════════════════════════════════════
# 예시 이미지 생성
# ══════════════════════════════════════════════════════════════

def make_img(title, lines, filename, width=760, height=360,
             header_bg=(30, 80, 160)):
    img = Image.new("RGB", (width, height), (245, 248, 252))
    draw = ImageDraw.Draw(img)
    try:
        fnt_title = ImageFont.truetype(BOLD,    20)
        fnt_body  = ImageFont.truetype(REGULAR, 16)
        fnt_small = ImageFont.truetype(REGULAR, 13)
    except Exception:
        fnt_title = fnt_body = fnt_small = ImageFont.load_default()

    draw.rectangle([0, 0, width, 50], fill=header_bg)
    draw.text((18, 13), title, font=fnt_title, fill=(255, 255, 255))

    y = 68
    for line in lines:
        if line.startswith("##"):
            draw.rectangle([10, y - 2, width - 10, y + 24], fill=(218, 230, 248))
            draw.text((18, y), line[2:].strip(), font=fnt_body, fill=(20, 60, 140))
            y += 32
        elif line.startswith(">>"):
            draw.rectangle([24, y + 2, 32, y + 16], fill=(30, 80, 160))
            draw.text((42, y), line[2:].strip(), font=fnt_body, fill=(40, 40, 40))
            y += 26
        elif line.startswith("--"):
            draw.line([28, y + 8, width - 28, y + 8], fill=(200, 210, 230), width=1)
            y += 16
        else:
            draw.text((24, y), line, font=fnt_small, fill=(80, 80, 80))
            y += 22

    draw.rectangle([0, 0, width - 1, height - 1], outline=(180, 200, 230), width=2)
    path = os.path.join(OUT_DIR, filename)
    img.save(path)
    return path


img_tax_invoice = make_img(
    "세금계산서 발급 – 홈택스 (hometax.go.kr)",
    [
        "## 전자세금계산서 > 건별발급",
        ">> 공급받는자 사업자등록번호 입력 → 자동조회",
        ">> 작성일자 / 공급가액 / 세액(10%) / 품목 입력",
        ">> [발급] 클릭 → 상대방 이메일 자동 전송",
        "--",
        "  ※ 발급 마감 : 익월 10일까지  |  지연 발급 시 가산세 발생",
        "  ※ 오기재 시 : 전자세금계산서 > 수정발급",
        "  ※ 발급 후 PDF 저장 → '세금계산서/YYYY-MM' 폴더 보관",
    ],
    "img_01_tax_invoice.png"
)

img_insurance_pay = make_img(
    "보험료·세금 납부",
    [
        "## 4대보험 : 4insure.or.kr > 사업장 업무 > 보험료 조회·납부",
        ">> 매월 고지서 확인 후 가상계좌 또는 인터넷뱅킹으로 납부",
        "--",
        "## 국세 : 홈택스 > 납부·고지·환급 > 세금납부",
        ">> 부가세·법인세·종합소득세·원천세 등",
        "--",
        "## 지방세 : 위택스 (wetax.go.kr) > 납부하기",
        ">> 지방소득세 등  |  납부 후 영수증 PDF 저장 보관",
    ],
    "img_02_insurance_pay.png"
)

img_cert_tax = make_img(
    "국세·지방세 납세증명서 발급",
    [
        "## 국세 납세증명서 : 홈택스 > 민원증명 > 납세증명서(국세완납증명)",
        ">> 용도 선택 → [즉시발급] → PDF 저장",
        "--",
        "## 지방세 납세증명서 : 위택스 > 신고납부 > 납세증명서 발급",
        ">> 지자체 선택 → [발급] → PDF 저장",
        "--",
        "  ※ 유효기간 : 발급일로부터 30일  (제출 직전 발급 권장)",
        "  ※ 체납 시 발급 불가 → 납부 후 재발급",
    ],
    "img_03_cert_tax.png"
)

img_cert_insurance = make_img(
    "4대보험 완납증명서 발급 – 4insure.or.kr",
    [
        "## 4대사회보험 정보연계센터 > 사업장 업무 > 완납증명서 발급",
        ">> 발급 사유·제출처 입력 후 [발급] → PDF 저장",
        "--",
        "  ※ 건강·국민·고용·산재보험 4개 기관 통합 1장으로 발급 가능",
        "  ※ 개별 발급 필요 시 각 공단 사이트에서 별도 발급",
        "  ※ 유효기간 : 발급일로부터 30일",
    ],
    "img_04_cert_insurance.png"
)

img_payroll = make_img(
    "급여대장 및 사업소득지급대장 작성 · 회계법인 송부",
    [
        "## 급여대장 작성",
        ">> 기존 양식 파일에 매월 입력 (기본급·수당·공제·실수령액)",
        ">> 은행 거래내역 다운로드 후 함께 첨부",
        "--",
        "## 사업소득지급대장 작성 (프리랜서·용역)",
        ">> 기존 양식 파일에 입력 (지급액·원천징수 3.3%)",
        ">> 은행 거래내역 다운로드 후 함께 첨부",
        "--",
        "## 회계법인 송부 및 급여명세서 전달",
        ">> 급여 지급 후 3 영업일 이내 이메일 발송",
        ">> 회계법인으로부터 급여명세서 수령 → 급여일 3일 전 대표님 전달",
    ],
    "img_05_payroll.png"
)

img_vat_cert = make_img(
    "부가가치세 과세표준증명 발급 – 홈택스",
    [
        "## 홈택스 > 민원증명 > 부가가치세 과세표준증명",
        ">> 사업자등록번호 확인 후 과세기간 선택",
        ">> [즉시발급] → PDF 저장",
        "--",
        "  ※ 과세기간 : 1기(1~6월) / 2기(7~12월) 선택",
        "  ※ 부가세 신고 완료 후 발급 가능",
        "  ※ 주로 금융기관·공공기관 제출용으로 사용",
        "  ※ 유효기간 : 발급일로부터 30일 (제출 직전 발급 권장)",
    ],
    "img_06_vat_cert.png"
)

img_fs_cert = make_img(
    "표준재무제표증명 발급 – 홈택스",
    [
        "## 홈택스 > 민원증명 > 표준재무제표증명(법인)",
        ">> 사업자등록번호 확인 후 사업연도 선택",
        ">> [즉시발급] → PDF 저장",
        "--",
        "  ※ 법인세 신고 완료 후 해당 사업연도 발급 가능",
        "  ※ 재무상태표 / 손익계산서 등 포함",
        "  ※ 금융기관 대출·입찰·계약 등 제출용으로 사용",
        "  ※ 유효기간 : 발급일로부터 30일 (제출 직전 발급 권장)",
    ],
    "img_07_fs_cert.png"
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
            textColor=BLUE, alignment=TA_CENTER, spaceAfter=6),
        "cover_sub": ParagraphStyle("cover_sub",
            fontName="NanumGothicBold", fontSize=14,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER, spaceAfter=4),
        "h1": ParagraphStyle("h1",
            fontName="NanumSquareB", fontSize=14,
            textColor=colors.white, leftIndent=8),
        "h2": ParagraphStyle("h2",
            fontName="NanumGothicBold", fontSize=12,
            textColor=BLUE, spaceBefore=12, spaceAfter=4, leftIndent=4),
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
        "table_hdr": ParagraphStyle("table_hdr",
            fontName="NanumGothicBold", fontSize=10,
            textColor=colors.white, alignment=TA_CENTER),
        "table_cell": ParagraphStyle("table_cell",
            fontName="NanumGothic", fontSize=10,
            textColor=colors.HexColor("#222222"), leading=15),
        "footer": ParagraphStyle("footer",
            fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#888888"), alignment=TA_CENTER),
    }

    def sec_hdr(text, color=BLUE):
        tbl = Table([[Paragraph(text, S["h1"])]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), color),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
            ("LEFTPADDING",   (0,0),(-1,-1), 10),
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
    story.append(HRFlowable(width="70%", thickness=2, color=BLUE, spaceAfter=10))
    story.append(Paragraph("회계 · 세무 업무", S["cover_sub"]))
    story.append(Spacer(1, 0.8*cm))

    ci = Table([
        ["작성일", "2026년 06월 15일"],
        ["작성 부서", "경영지원팀"],
        ["인수인계 범위", "세금계산서 발급 / 보험료·세금 납부 / 납세·완납증명서 발급 / 급여대장 작성·송부"],
    ], colWidths=[3.5*cm, 12*cm])
    ci.setStyle(TableStyle([
        ("FONTNAME",     (0,0),(-1,-1), "NanumGothic"),
        ("FONTNAME",     (0,0),(0,-1),  "NanumGothicBold"),
        ("FONTSIZE",     (0,0),(-1,-1), 11),
        ("TEXTCOLOR",    (0,0),(0,-1),  BLUE),
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",   (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LINEBELOW",    (0,0),(-1,-2), 0.5, colors.HexColor("#cccccc")),
    ]))
    story.append(ci)
    story.append(PageBreak())

    # ══ 1. 세금계산서 발급
    story.append(sec_hdr("1. 세금계산서 발급"))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "거래 상대방에게 재화·용역을 공급한 경우 <b>국세청 홈택스</b>에서 전자세금계산서를 발급합니다. "
        "발급 마감은 <b>익월 10일</b>이며, 지연 시 가산세가 부과됩니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인 (공동인증서 또는 간편인증)",
        "전자세금계산서 &gt; 건별발급 메뉴 선택",
        "공급받는 자 사업자등록번호 입력 후 자동조회",
        "작성일자 / 공급가액 / 세액(공급가액 × 10%) / 품목 입력 후 [발급]",
        "발급 완료 후 PDF 저장 → '세금계산서/YYYY-MM' 폴더 보관",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_tax_invoice, "홈택스 전자세금계산서 발급 화면 예시"))
    story.append(Paragraph("※ 오기재 수정 : 전자세금계산서 &gt; 수정발급 메뉴 이용", S["note"]))
    story.append(Paragraph("※ 면세 거래 : 부가세 없는 전자계산서(별도 메뉴)로 발급", S["note"]))
    story.append(PageBreak())

    # ══ 2. 보험료 및 세금 납부
    story.append(sec_hdr("2. 보험료 및 세금 납부", color=GREEN))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 4대 사회보험료", S["h2"]))
    story.append(Paragraph(
        "<b>4대사회보험 정보연계센터</b>(4insure.or.kr) &gt; 사업장 업무 &gt; 보험료 조회·납부에서 "
        "매월 고지서를 확인하고 가상계좌 또는 인터넷뱅킹으로 납부합니다.",
        S["body"]))

    story.append(Paragraph("■ 국세 납부", S["h2"]))
    story.append(Paragraph(
        "<b>홈택스</b> &gt; 납부·고지·환급 &gt; 세금납부 메뉴에서 부가가치세·법인세·종합소득세·원천세 등을 납부합니다.",
        S["body"]))

    story.append(Paragraph("■ 지방세 납부", S["h2"]))
    story.append(Paragraph(
        "<b>위택스</b>(wetax.go.kr) &gt; 납부하기 메뉴에서 지방소득세 등을 납부합니다. "
        "납부 완료 후 영수증 PDF를 저장해 보관합니다.",
        S["body"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_insurance_pay, "4대보험·국세·지방세 납부 화면 예시"))
    story.append(PageBreak())

    # ══ 3. 납세증명서 발급
    story.append(sec_hdr("3. 납세증명서 발급", color=ORANGE))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "납세증명서는 체납 없이 세금을 납부하고 있음을 증명하는 서류로, "
        "입찰·대출·공공기관 계약 등에 제출합니다. "
        "국세와 지방세를 각각 발급합니다.",
        S["body"]))

    story.append(Paragraph("■ 국세 납세증명서 (홈택스)", S["h2"]))
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 &gt; 납세증명서(국세완납증명) &gt; 발급신청",
        "용도 선택 후 [즉시발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Paragraph("■ 지방세 납세증명서 (위택스)", S["h2"]))
    for i, s in enumerate([
        "위택스(wetax.go.kr) 로그인",
        "신고납부 &gt; 납세증명서 발급",
        "사업장 소재 지자체 선택 후 [발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_cert_tax, "홈택스·위택스 납세증명서 발급 화면 예시"))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장", S["note"]))
    story.append(Paragraph("※ 체납 존재 시 발급 불가 → 납부 후 재발급", S["note"]))
    story.append(PageBreak())

    # ══ 4. 4대보험 완납증명서 발급
    story.append(sec_hdr("4. 4대보험 완납증명서 발급", color=TEAL))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "4대 사회보험료(건강·국민연금·고용·산재)를 완납하였음을 증명하는 서류입니다. "
        "입찰·계약 등에 제출하며 <b>4대사회보험 정보연계센터</b>에서 통합 발급합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))
    for i, s in enumerate([
        "4대사회보험 정보연계센터(4insure.or.kr) 로그인",
        "사업장 업무 &gt; 완납증명서 발급 메뉴 선택",
        "발급 사유 및 제출처 입력 후 [발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_cert_insurance, "4대보험 완납증명서 발급 화면 예시"))
    story.append(Paragraph("※ 4개 기관 통합 1장 발급 가능 (개별 발급 필요 시 각 공단 사이트 이용)", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 30일", S["note"]))
    story.append(PageBreak())

    # ══ 5. 급여대장 작성 및 회계법인 송부
    story.append(sec_hdr("5. 급여대장 작성 및 회계법인 송부", color=PURPLE))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 급여대장 작성", S["h2"]))
    story.append(Paragraph(
        "기존 양식 파일에 매월 급여 내용을 입력합니다. "
        "기본급·수당·4대보험 공제·소득세·실수령액 항목을 양식에 맞춰 작성하면 됩니다. "
        "해당 월 <b>은행 거래내역을 다운로드하여 함께 첨부</b>합니다.",
        S["body"]))

    story.append(Paragraph("■ 사업소득지급대장 작성 (프리랜서·용역)", S["h2"]))
    story.append(Paragraph(
        "기존 양식 파일에 프리랜서·용역 지급 내용을 입력합니다. "
        "성명·지급액·원천징수세액(3.3%) 항목을 양식에 맞춰 작성하면 됩니다. "
        "해당 월 <b>은행 거래내역을 다운로드하여 함께 첨부</b>합니다.",
        S["body"]))

    story.append(Paragraph("■ 회계법인 송부", S["h2"]))
    for i, s in enumerate([
        "급여 지급일 기준 <b>3 영업일 이내</b> 이메일 발송",
        "첨부 파일 : 급여대장 + 사업소득지급대장 + 은행 거래내역(해당월)",
        "수신처 : 담당 세무사 이메일 (주소록 '회계법인' 그룹 참조)",
        "제목 형식 : [회사명] YYYY년 MM월 급여 관련 자료 송부",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Paragraph("■ 급여명세서 대표님 전달", S["h2"]))
    story.append(Paragraph(
        "자료 송부 후 회계법인으로부터 <b>급여명세서</b>를 메일로 수령합니다. "
        "수령 즉시 내용을 확인하고 <b>급여일 3일 전까지 대표님께 전달</b>합니다.",
        S["body"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_payroll, "급여대장 작성 및 회계법인 송부 예시"))
    story.append(PageBreak())

    # ══ 6. 부가가치세 과세표준증명 발급
    DARK_RED = colors.HexColor("#8a1a1a")
    story.append(sec_hdr("6. 부가가치세 과세표준증명 발급", color=DARK_RED))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "사업자의 부가가치세 신고 과세표준 금액을 증명하는 서류로, "
        "금융기관·공공기관 제출 등에 활용합니다. "
        "<b>홈택스</b> &gt; 민원증명 메뉴에서 발급합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 &gt; 부가가치세 과세표준증명 선택",
        "사업자등록번호 확인 후 과세기간 선택 (1기: 1~6월 / 2기: 7~12월)",
        "[즉시발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_vat_cert, "홈택스 부가가치세 과세표준증명 발급 화면 예시"))
    story.append(Paragraph("※ 부가세 신고 완료 후 해당 과세기간 발급 가능", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장", S["note"]))
    story.append(PageBreak())

    # ══ 7. 표준재무제표증명 발급
    DARK_GREEN2 = colors.HexColor("#1a5a2a")
    story.append(sec_hdr("7. 표준재무제표증명 발급", color=DARK_GREEN2))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph(
        "법인세 신고 내용을 바탕으로 재무상태표·손익계산서 등을 증명하는 서류입니다. "
        "금융기관 대출·입찰·계약 등 제출에 활용하며 <b>홈택스</b>에서 발급합니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 &gt; 표준재무제표증명(법인) 선택",
        "사업자등록번호 확인 후 사업연도 선택",
        "[즉시발급] → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))
    story.append(img_block(img_fs_cert, "홈택스 표준재무제표증명 발급 화면 예시"))
    story.append(Paragraph("※ 법인세 신고 완료 후 해당 사업연도 발급 가능", S["note"]))
    story.append(Paragraph("※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장", S["note"]))

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

def add_sec_hdr(doc, text, bg="1e4fa0"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    cell.text = text
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    r = p.runs[0]; ko(r)
    r.font.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()

def add_para(doc, text, bold=False, size=10.5, indent=0.3, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); ko(r)
    r.font.size = Pt(size)
    r.font.bold = bold
    if color:
        r.font.color.rgb = color
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


def build_docx(out_path):
    doc = Document()
    for sec in doc.sections:
        sec.page_width  = Cm(21)
        sec.page_height = Cm(29.7)
        sec.top_margin    = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2.5)

    # 표지
    for _ in range(5):
        doc.add_paragraph()
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("업무 인수인계 자료"); ko(tr)
    tr.font.size = Pt(26); tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x1E, 0x4F, 0xA0)

    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run("회계 · 세무 업무"); ko(sr)
    sr.font.size = Pt(15); sr.font.bold = True
    sr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()
    ci_tbl = doc.add_table(rows=3, cols=2)
    ci_tbl.style = "Table Grid"
    ci_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    ci_data = [
        ("작성일", "2026년 06월 15일"),
        ("작성 부서", "경영지원팀"),
        ("인수인계 범위", "세금계산서 발급 / 보험료·세금 납부 / 납세·완납증명서 발급 / 급여대장 작성·송부"),
    ]
    for ri, (k, v) in enumerate(ci_data):
        cells = ci_tbl.rows[ri].cells
        set_cell_bg(cells[0], "dce6f8")
        cells[0].text = k; cells[1].text = v
        for ci, c in enumerate(cells):
            r = c.paragraphs[0].runs[0]; ko(r)
            r.font.size = Pt(10.5)
            if ci == 0: r.font.bold = True
        cells[0].width = Cm(3.5); cells[1].width = Cm(12)

    doc.add_page_break()

    # 1. 세금계산서
    add_sec_hdr(doc, "1. 세금계산서 발급")
    add_para(doc,
        "거래 상대방에게 재화·용역을 공급한 경우 국세청 홈택스에서 전자세금계산서를 발급합니다. "
        "발급 마감은 익월 10일이며, 지연 시 가산세가 부과됩니다.")
    doc.add_paragraph()
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인 (공동인증서 또는 간편인증)",
        "전자세금계산서 > 건별발급 메뉴 선택",
        "공급받는 자 사업자등록번호 입력 후 자동조회",
        "작성일자 / 공급가액 / 세액(공급가액 × 10%) / 품목 입력 후 [발급]",
        "발급 완료 후 PDF 저장 → '세금계산서/YYYY-MM' 폴더 보관",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_tax_invoice, "홈택스 전자세금계산서 발급 화면 예시")
    add_para(doc, "※ 오기재 수정 : 전자세금계산서 > 수정발급 메뉴 이용",
             color=RGBColor(0xB0, 0x40, 0x00))
    add_para(doc, "※ 면세 거래 : 부가세 없는 전자계산서(별도 메뉴)로 발급",
             color=RGBColor(0xB0, 0x40, 0x00))
    doc.add_page_break()

    # 2. 보험료 및 세금 납부
    add_sec_hdr(doc, "2. 보험료 및 세금 납부", bg="1a7a40")
    add_para(doc, "■ 4대 사회보험료", bold=True)
    add_para(doc,
        "4대사회보험 정보연계센터(4insure.or.kr) > 사업장 업무 > 보험료 조회·납부에서 "
        "매월 고지서를 확인하고 가상계좌 또는 인터넷뱅킹으로 납부합니다.")
    add_para(doc, "■ 국세 납부", bold=True)
    add_para(doc,
        "홈택스 > 납부·고지·환급 > 세금납부 메뉴에서 부가가치세·법인세·종합소득세·원천세 등을 납부합니다.")
    add_para(doc, "■ 지방세 납부", bold=True)
    add_para(doc,
        "위택스(wetax.go.kr) > 납부하기 메뉴에서 지방소득세 등을 납부합니다. "
        "납부 완료 후 영수증 PDF를 저장해 보관합니다.")
    doc.add_paragraph()
    add_img(doc, img_insurance_pay, "4대보험·국세·지방세 납부 화면 예시")
    doc.add_page_break()

    # 3. 납세증명서
    add_sec_hdr(doc, "3. 납세증명서 발급", bg="e07000")
    add_para(doc,
        "납세증명서는 체납 없이 세금을 납부하고 있음을 증명하는 서류로, "
        "입찰·대출·공공기관 계약 등에 제출합니다. 국세와 지방세를 각각 발급합니다.")
    add_para(doc, "■ 국세 납세증명서 (홈택스)", bold=True)
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 > 납세증명서(국세완납증명) > 발급신청",
        "용도 선택 후 [즉시발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    add_para(doc, "■ 지방세 납세증명서 (위택스)", bold=True)
    for i, s in enumerate([
        "위택스(wetax.go.kr) 로그인",
        "신고납부 > 납세증명서 발급",
        "사업장 소재 지자체 선택 후 [발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_cert_tax, "홈택스·위택스 납세증명서 발급 화면 예시")
    add_para(doc, "※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장",
             color=RGBColor(0xB0, 0x40, 0x00))
    add_para(doc, "※ 체납 존재 시 발급 불가 → 납부 후 재발급",
             color=RGBColor(0xB0, 0x40, 0x00))
    doc.add_page_break()

    # 4. 4대보험 완납증명서
    add_sec_hdr(doc, "4. 4대보험 완납증명서 발급", bg="0a6a7a")
    add_para(doc,
        "4대 사회보험료(건강·국민연금·고용·산재)를 완납하였음을 증명하는 서류입니다. "
        "입찰·계약 등에 제출하며 4대사회보험 정보연계센터에서 통합 발급합니다.")
    doc.add_paragraph()
    for i, s in enumerate([
        "4대사회보험 정보연계센터(4insure.or.kr) 로그인",
        "사업장 업무 > 완납증명서 발급 메뉴 선택",
        "발급 사유 및 제출처 입력 후 [발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_cert_insurance, "4대보험 완납증명서 발급 화면 예시")
    add_para(doc, "※ 4개 기관 통합 1장 발급 가능 (개별 발급 필요 시 각 공단 사이트 이용)",
             color=RGBColor(0xB0, 0x40, 0x00))
    add_para(doc, "※ 유효기간 : 발급일로부터 30일",
             color=RGBColor(0xB0, 0x40, 0x00))
    doc.add_page_break()

    # 5. 급여대장
    add_sec_hdr(doc, "5. 급여대장 작성 및 회계법인 송부", bg="6a1a8a")
    add_para(doc, "■ 급여대장 작성", bold=True)
    add_para(doc,
        "기존 양식 파일에 매월 급여 내용을 입력합니다. "
        "기본급·수당·4대보험 공제·소득세·실수령액 항목을 양식에 맞춰 작성하면 됩니다. "
        "해당 월 은행 거래내역을 다운로드하여 함께 첨부합니다.")
    add_para(doc, "■ 사업소득지급대장 작성 (프리랜서·용역)", bold=True)
    add_para(doc,
        "기존 양식 파일에 프리랜서·용역 지급 내용을 입력합니다. "
        "성명·지급액·원천징수세액(3.3%) 항목을 양식에 맞춰 작성하면 됩니다. "
        "해당 월 은행 거래내역을 다운로드하여 함께 첨부합니다.")
    add_para(doc, "■ 회계법인 송부", bold=True)
    for i, s in enumerate([
        "급여 지급일 기준 3 영업일 이내 이메일 발송",
        "첨부 파일 : 급여대장 + 사업소득지급대장 + 은행 거래내역(해당월)",
        "수신처 : 담당 세무사 이메일 (주소록 '회계법인' 그룹 참조)",
        "제목 형식 : [회사명] YYYY년 MM월 급여 관련 자료 송부",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    add_para(doc, "■ 급여명세서 대표님 전달", bold=True)
    add_para(doc,
        "자료 송부 후 회계법인으로부터 급여명세서를 메일로 수령합니다. "
        "수령 즉시 내용을 확인하고 급여일 3일 전까지 대표님께 전달합니다.")
    doc.add_paragraph()
    add_img(doc, img_payroll, "급여대장 작성 및 회계법인 송부 예시")
    doc.add_page_break()

    # 6. 부가가치세 과세표준증명
    add_sec_hdr(doc, "6. 부가가치세 과세표준증명 발급", bg="8a1a1a")
    add_para(doc,
        "사업자의 부가가치세 신고 과세표준 금액을 증명하는 서류로, "
        "금융기관·공공기관 제출 등에 활용합니다. 홈택스 > 민원증명 메뉴에서 발급합니다.")
    doc.add_paragraph()
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 > 부가가치세 과세표준증명 선택",
        "사업자등록번호 확인 후 과세기간 선택 (1기: 1~6월 / 2기: 7~12월)",
        "[즉시발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_vat_cert, "홈택스 부가가치세 과세표준증명 발급 화면 예시")
    add_para(doc, "※ 부가세 신고 완료 후 해당 과세기간 발급 가능",
             color=RGBColor(0xB0, 0x40, 0x00))
    add_para(doc, "※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장",
             color=RGBColor(0xB0, 0x40, 0x00))
    doc.add_page_break()

    # 7. 표준재무제표증명
    add_sec_hdr(doc, "7. 표준재무제표증명 발급", bg="1a5a2a")
    add_para(doc,
        "법인세 신고 내용을 바탕으로 재무상태표·손익계산서 등을 증명하는 서류입니다. "
        "금융기관 대출·입찰·계약 등 제출에 활용하며 홈택스에서 발급합니다.")
    doc.add_paragraph()
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "민원증명 > 표준재무제표증명(법인) 선택",
        "사업자등록번호 확인 후 사업연도 선택",
        "[즉시발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")
    doc.add_paragraph()
    add_img(doc, img_fs_cert, "홈택스 표준재무제표증명 발급 화면 예시")
    add_para(doc, "※ 법인세 신고 완료 후 해당 사업연도 발급 가능",
             color=RGBColor(0xB0, 0x40, 0x00))
    add_para(doc, "※ 유효기간 : 발급일로부터 30일 → 제출 직전 발급 권장",
             color=RGBColor(0xB0, 0x40, 0x00))

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
    pdf_path  = os.path.join(OUT_DIR, "회계세무_업무인수인계.pdf")
    docx_path = os.path.join(OUT_DIR, "회계세무_업무인수인계.docx")
    build_pdf(pdf_path)
    build_docx(docx_path)
    print("모든 파일 생성 완료.")
