#!/usr/bin/env python3
"""
회계·세무 업무 인수인계 자료 생성 스크립트
DOCX + PDF 출력
"""

import io
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx.opc.constants

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageFont
import os

# ── 폰트 등록 ──────────────────────────────────────────────
FONT_DIR = "/usr/share/fonts/truetype/nanum"
REGULAR  = os.path.join(FONT_DIR, "NanumGothic.ttf")
BOLD     = os.path.join(FONT_DIR, "NanumGothicBold.ttf")
EXTRA    = os.path.join(FONT_DIR, "NanumSquareB.ttf")

pdfmetrics.registerFont(TTFont("NanumGothic",      REGULAR))
pdfmetrics.registerFont(TTFont("NanumGothicBold",  BOLD))
pdfmetrics.registerFont(TTFont("NanumGothicExtra", EXTRA))

OUT_DIR = "/home/user/project-1/handover_docs"
os.makedirs(OUT_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════════════
# 공통 예시 이미지 생성 (Pillow)
# ══════════════════════════════════════════════════════════════

def make_img(title, lines, filename, width=760, height=420,
             bg=(245, 248, 252), header_bg=(30, 80, 160)):
    """간단한 예시 화면 이미지를 생성해 반환 (경로)."""
    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)

    try:
        fnt_title = ImageFont.truetype(BOLD,  22)
        fnt_body  = ImageFont.truetype(REGULAR, 17)
        fnt_small = ImageFont.truetype(REGULAR, 14)
    except Exception:
        fnt_title = fnt_body = fnt_small = ImageFont.load_default()

    # 헤더 바
    draw.rectangle([0, 0, width, 54], fill=header_bg)
    draw.text((20, 14), title, font=fnt_title, fill=(255, 255, 255))

    # 내용
    y = 78
    for line in lines:
        if line.startswith("##"):
            draw.rectangle([12, y - 2, width - 12, y + 26], fill=(218, 230, 248))
            draw.text((20, y), line[2:].strip(), font=fnt_body, fill=(20, 60, 140))
            y += 34
        elif line.startswith(">>"):
            draw.rectangle([26, y, 36, y + 18], fill=(30, 80, 160))
            draw.text((46, y), line[2:].strip(), font=fnt_body, fill=(40, 40, 40))
            y += 28
        elif line.startswith("--"):
            draw.line([30, y + 10, width - 30, y + 10], fill=(200, 210, 230), width=1)
            y += 18
        else:
            draw.text((26, y), line, font=fnt_small, fill=(80, 80, 80))
            y += 24

    # 테두리
    draw.rectangle([0, 0, width - 1, height - 1],
                   outline=(180, 200, 230), width=2)

    path = os.path.join(OUT_DIR, filename)
    img.save(path)
    return path


# ──────────────────────────────────────────────────────────────
# 이미지 생성
# ──────────────────────────────────────────────────────────────

img_tax_invoice = make_img(
    "세금계산서 발급 – 홈택스",
    [
        "## 국세청 홈택스 (hometax.go.kr) 로그인 후 진행",
        ">> 메뉴 경로 : 전자세금계산서 > 건별발급",
        "--",
        "  ① 공급받는자 사업자등록번호 입력 → 자동조회",
        "  ② 작성일자 / 공급가액 / 세액 입력",
        "  ③ 품목 (상품·용역명) 및 수량·단가 입력",
        "  ④ [발급] 버튼 클릭 → 상대방 메일 자동 발송",
        "--",
        "  ★ 발급 마감 : 익월 10일까지 (지연 시 가산세 발생)",
        "  ★ 수정세금계산서 필요 시 : 전자세금계산서 > 수정발급",
        "--",
        "  발급 후 PDF 저장 → '세금계산서' 폴더에 월별 보관",
    ],
    "img_01_tax_invoice.png"
)

img_insurance = make_img(
    "4대보험료 납부 – 4대사회보험 정보연계센터",
    [
        "## 4대사회보험 정보연계센터 (4insure.or.kr)",
        ">> 메뉴 : 사업장 업무 > 보험료 조회·납부",
        "--",
        "  ① 매월 15일 전후 고지서 확인 (건강·국민·고용·산재)",
        "  ② 납부 기한 : 매월 말일 (건강보험 10일, 나머지 말일)",
        "  ③ 인터넷뱅킹 또는 가상계좌로 납부",
        "--",
        "## 국세·지방세 납부 경로",
        ">> 국세 : 홈택스 > 납부·고지·환급 > 세금납부",
        ">> 지방세 : 위택스 (wetax.go.kr) > 납부하기",
        "--",
        "  ★ 부가세 : 1월·7월 / 법인세 : 3월 / 종합소득세 : 5월",
        "  ★ 납부 완료 후 영수증 PDF 저장 보관",
    ],
    "img_02_insurance_tax.png"
)

img_certificate = make_img(
    "납세증명서 발급 – 홈택스 / 위택스",
    [
        "## 국세 납세증명서 : 홈택스 (hometax.go.kr)",
        ">> 메뉴 : 민원증명 > 납세증명서(국세완납증명) > 발급",
        "--",
        "  ① 용도 선택 (금융기관 제출 / 관공서 제출 등)",
        "  ② [즉시발급] → PDF 저장 또는 인쇄",
        "--",
        "## 지방세 납세증명서 : 위택스 (wetax.go.kr)",
        ">> 메뉴 : 신고납부 > 납세증명서 발급",
        "--",
        "  ① 지자체 선택 (등록된 사업장 소재지 기준)",
        "  ② [발급] → PDF 다운로드",
        "--",
        "  ★ 증명서 유효기간 : 발급일로부터 30일",
        "  ★ 제출처별 유효기간 상이 → 제출 직전 발급 권장",
    ],
    "img_03_certificate.png"
)

img_payroll = make_img(
    "급여대장 작성 및 회계법인 송부",
    [
        "## 매월 급여대장 작성 (Excel)",
        ">> 파일 : [연도]_급여대장_[월].xlsx",
        "--",
        "  기재 항목 : 성명 / 입사일 / 기본급 / 수당",
        "              4대보험 공제 / 소득세·지방소득세 / 실수령액",
        "--",
        "## 사업소득지급대장 작성 (프리랜서·용역)",
        ">> 파일 : [연도]_사업소득지급대장_[월].xlsx",
        "--",
        "  기재 항목 : 성명 / 주민번호(뒤 가림) / 지급액 / 원천징수(3.3%)",
        "--",
        "## 회계법인 송부",
        "  ① 매월 급여 지급 후 3 영업일 내 이메일 발송",
        "  ② 첨부파일 : 급여대장 + 사업소득지급대장 + 통장사본",
        "  ③ 수신처 : 담당 세무사 이메일 (주소록 참조)",
    ],
    "img_04_payroll.png"
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

    # 스타일
    S = {
        "cover_title": ParagraphStyle("cover_title",
            fontName="NanumGothicExtra", fontSize=26,
            textColor=colors.HexColor("#1e4fa0"),
            alignment=TA_CENTER, spaceAfter=6),
        "cover_sub": ParagraphStyle("cover_sub",
            fontName="NanumGothicBold", fontSize=14,
            textColor=colors.HexColor("#555555"),
            alignment=TA_CENTER, spaceAfter=4),
        "cover_info": ParagraphStyle("cover_info",
            fontName="NanumGothic", fontSize=11,
            textColor=colors.HexColor("#333333"),
            alignment=TA_CENTER, spaceAfter=2),
        "h1": ParagraphStyle("h1",
            fontName="NanumGothicExtra", fontSize=16,
            textColor=colors.white,
            alignment=TA_LEFT, spaceAfter=0, spaceBefore=0,
            leftIndent=8),
        "h2": ParagraphStyle("h2",
            fontName="NanumGothicBold", fontSize=13,
            textColor=colors.HexColor("#1e4fa0"),
            spaceBefore=14, spaceAfter=4, leftIndent=4),
        "body": ParagraphStyle("body",
            fontName="NanumGothic", fontSize=10.5,
            leading=18, textColor=colors.HexColor("#333333"),
            spaceAfter=3, leftIndent=12),
        "bullet": ParagraphStyle("bullet",
            fontName="NanumGothic", fontSize=10.5,
            leading=18, textColor=colors.HexColor("#333333"),
            leftIndent=24, spaceAfter=2,
            bulletIndent=10, bulletFontName="NanumGothic"),
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
            textColor=colors.HexColor("#222222"), alignment=TA_LEFT,
            leading=15),
        "footer": ParagraphStyle("footer",
            fontName="NanumGothic", fontSize=9,
            textColor=colors.HexColor("#888888"),
            alignment=TA_CENTER),
    }

    BLUE   = colors.HexColor("#1e4fa0")
    LBLUE  = colors.HexColor("#dce6f8")
    ORANGE = colors.HexColor("#e07000")
    GREEN  = colors.HexColor("#1a7a40")

    def section_header(text, color=BLUE):
        tbl = Table([[Paragraph(text, S["h1"])]], colWidths=[16.6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), color),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
            ("LEFTPADDING",   (0,0),(-1,-1), 10),
            ("ROUNDEDCORNERS", [4,4,4,4]),
        ]))
        return tbl

    def img_block(path, caption_text, w=16*cm):
        from reportlab.platypus import Image as RLImage
        im = RLImage(path, width=w, height=w * 420/760)
        cap = Paragraph(f"▲ {caption_text}", S["caption"])
        return KeepTogether([im, cap])

    def checklist_table(rows, col_w=None):
        if col_w is None:
            col_w = [2*cm, 14.6*cm]
        data = [[Paragraph(c, S["table_hdr"]) for c in ["체크", "업무 내용"]]]
        for r in rows:
            data.append([Paragraph("☐", S["table_cell"]),
                         Paragraph(r, S["table_cell"])])
        t = Table(data, colWidths=col_w)
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0), (-1,0), BLUE),
            ("BACKGROUND",   (0,1), (-1,-1), colors.HexColor("#f5f8fc")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef2fa")]),
            ("GRID",         (0,0), (-1,-1), 0.4, colors.HexColor("#b0c0e0")),
            ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",   (0,0), (-1,-1), 5),
            ("BOTTOMPADDING",(0,0), (-1,-1), 5),
            ("LEFTPADDING",  (0,0), (-1,-1), 8),
            ("ALIGN",        (0,0), (0,-1), "CENTER"),
        ]))
        return t

    def schedule_table(rows):
        data = [[Paragraph(c, S["table_hdr"]) for c in ["시기","업무","담당 시스템","비고"]]]
        for r in rows:
            data.append([Paragraph(c, S["table_cell"]) for c in r])
        cw = [2.8*cm, 5*cm, 4.8*cm, 4*cm]
        t = Table(data, colWidths=cw)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), BLUE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef2fa")]),
            ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#b0c0e0")),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 5),
            ("BOTTOMPADDING", (0,0),(-1,-1), 5),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ]))
        return t

    # ── 본문 조립 ──────────────────────────────────────────────
    story = []

    # ── 표지 ──────────────────────────────────────────────────
    story.append(Spacer(1, 3.5*cm))
    story.append(Paragraph("업무 인수인계 자료", S["cover_title"]))
    story.append(Spacer(1, 0.4*cm))
    story.append(HRFlowable(width="70%", thickness=2, color=BLUE, spaceAfter=12))
    story.append(Paragraph("회계 · 세무 업무", S["cover_sub"]))
    story.append(Spacer(1, 0.6*cm))

    cover_info = [
        ["작성일", "2026년 06월 15일"],
        ["작성 부서", "경영지원팀"],
        ["인수인계 범위", "세금계산서 발급 / 세금·보험료 납부\n납세증명서 발급 / 급여대장 작성·송부"],
    ]
    ci_tbl = Table(cover_info, colWidths=[3.5*cm, 12*cm])
    ci_tbl.setStyle(TableStyle([
        ("FONTNAME",  (0,0),(-1,-1), "NanumGothic"),
        ("FONTSIZE",  (0,0),(-1,-1), 11),
        ("FONTNAME",  (0,0),(0,-1),  "NanumGothicBold"),
        ("TEXTCOLOR", (0,0),(0,-1),  BLUE),
        ("VALIGN",    (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",(0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LINEBELOW", (0,0),(-1,-2), 0.5, colors.HexColor("#cccccc")),
    ]))
    story.append(ci_tbl)
    story.append(Spacer(1, 1*cm))

    # 목차
    toc_data = [
        [Paragraph("목  차", S["table_hdr"])],
        [Paragraph("1. 세금계산서 발급", S["body"])],
        [Paragraph("2. 보험료 및 세금 납부", S["body"])],
        [Paragraph("3. 납세증명서 발급", S["body"])],
        [Paragraph("4. 급여대장 작성 및 회계법인 송부", S["body"])],
        [Paragraph("5. 월별 업무 일정 요약", S["body"])],
    ]
    toc_tbl = Table(toc_data, colWidths=[16.6*cm])
    toc_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(0,0), BLUE),
        ("BACKGROUND", (0,1),(-1,-1), colors.HexColor("#f0f4fc")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#e8eefa")]),
        ("GRID",  (0,0),(-1,-1), 0.4, colors.HexColor("#b0c0e0")),
        ("TOPPADDING",   (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",  (0,1),(-1,-1), 16),
    ]))
    story.append(toc_tbl)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # 섹션 1 : 세금계산서 발급
    # ══════════════════════════════════════════════════════════
    story.append(section_header("1. 세금계산서 발급"))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 개요", S["h2"]))
    story.append(Paragraph(
        "거래 상대방에게 재화·용역을 공급한 경우 <b>국세청 홈택스</b>를 통해 전자세금계산서를 발급합니다. "
        "발급 마감일(익월 10일)을 반드시 준수해야 하며, 미발급·지연 발급 시 가산세가 부과됩니다.",
        S["body"]))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("■ 발급 절차", S["h2"]))
    steps = [
        "홈택스(hometax.go.kr) 로그인 (공동인증서 또는 간편인증)",
        "메뉴 : <b>전자세금계산서 &gt; 건별발급</b> 클릭",
        "공급받는 자 사업자등록번호 입력 후 자동조회",
        "작성일자 / 공급가액 / 세액 입력 (세액 = 공급가액 × 10%)",
        "품목명·수량·단가 입력 후 <b>[발급]</b> 버튼 클릭",
        "발급 완료 → 상대방 이메일 자동 전송 확인",
        "발급된 세금계산서 PDF 저장 → <b>'세금계산서/YYYY-MM'</b> 폴더 보관",
    ]
    for i, s in enumerate(steps, 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))
    story.append(Spacer(1, 0.3*cm))

    story.append(img_block(img_tax_invoice,
        "홈택스 전자세금계산서 발급 화면 예시"))

    story.append(Paragraph("■ 주요 유의사항", S["h2"]))
    notes = [
        "⚠ 발급 마감 : 공급일 다음달 10일까지 (e.g. 6월 거래 → 7월 10일까지)",
        "⚠ 수정세금계산서 : 오기재 발생 시 메뉴 > 수정발급 → 사유·수정금액 입력",
        "⚠ 매입세금계산서 : 상대방 발급본 홈택스에서 수신 확인 후 보관",
        "⚠ 면세 거래 : 계산서(부가세 없음)로 발급 (전자계산서 메뉴 별도)",
    ]
    for n in notes:
        story.append(Paragraph(n, S["note"]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("■ 발급 체크리스트", S["h2"]))
    story.append(checklist_table([
        "공급받는자 사업자등록번호 및 상호 정확히 입력",
        "공급가액·세액 금액 확인 (세액 = 공급가액 × 10%)",
        "작성일자 = 실제 거래일 기준",
        "발급 완료 후 상대방 수신 여부 확인",
        "PDF 저장 후 지정 폴더에 보관",
        "익월 10일 마감 전 전체 발급 완료 확인",
    ]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # 섹션 2 : 보험료 및 세금 납부
    # ══════════════════════════════════════════════════════════
    story.append(section_header("2. 보험료 및 세금 납부", color=GREEN))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 4대 사회보험료 납부", S["h2"]))
    story.append(Paragraph(
        "<b>4대사회보험 정보연계센터</b>(4insure.or.kr)에서 월별 고지서를 확인하고, "
        "인터넷뱅킹 또는 가상계좌로 납부합니다.",
        S["body"]))

    ins_data = [
        [Paragraph(c, S["table_hdr"]) for c in ["보험 종류", "납부 기한", "납부처", "비고"]],
        [Paragraph("건강보험", S["table_cell"]),
         Paragraph("매월 10일", S["table_cell"]),
         Paragraph("건강보험공단 가상계좌", S["table_cell"]),
         Paragraph("장기요양보험 포함", S["table_cell"])],
        [Paragraph("국민연금", S["table_cell"]),
         Paragraph("매월 말일", S["table_cell"]),
         Paragraph("국민연금공단 가상계좌", S["table_cell"]),
         Paragraph("", S["table_cell"])],
        [Paragraph("고용보험", S["table_cell"]),
         Paragraph("매월 말일", S["table_cell"]),
         Paragraph("근로복지공단 가상계좌", S["table_cell"]),
         Paragraph("산재보험 동시 납부", S["table_cell"])],
        [Paragraph("산재보험", S["table_cell"]),
         Paragraph("매월 말일", S["table_cell"]),
         Paragraph("근로복지공단 가상계좌", S["table_cell"]),
         Paragraph("", S["table_cell"])],
    ]
    ins_tbl = Table(ins_data, colWidths=[3*cm, 3*cm, 5.8*cm, 4.8*cm])
    ins_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), GREEN),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#edfaf2")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#90c0a0")),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(ins_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 국세 납부", S["h2"]))
    story.append(Paragraph(
        "<b>홈택스</b> 로그인 후 <b>납부·고지·환급 &gt; 세금납부</b> 메뉴에서 납부합니다. "
        "주요 세금 납부 일정은 아래와 같습니다.",
        S["body"]))

    tax_rows = [
        ["1월 / 7월", "부가가치세 (법인)", "홈택스", "1월 25일·7월 25일"],
        ["1월 / 5월", "부가가치세 (개인)", "홈택스", "예정·확정 신고"],
        ["3월", "법인세", "홈택스", "3월 31일"],
        ["5월", "종합소득세", "홈택스", "5월 31일"],
        ["매월", "원천세 (소득세)", "홈택스", "다음달 10일"],
        ["매월", "지방소득세", "위택스", "원천세와 동일"],
    ]
    story.append(schedule_table(tax_rows))
    story.append(Spacer(1, 0.3*cm))

    story.append(img_block(img_insurance,
        "4대보험·국세·지방세 납부 화면 예시"))

    story.append(Paragraph("■ 납부 체크리스트", S["h2"]))
    story.append(checklist_table([
        "4대사회보험 고지서 수신 확인 (매월 초)",
        "건강보험료 10일 내 납부 완료",
        "국민연금·고용·산재보험 말일 내 납부 완료",
        "원천세 다음달 10일까지 홈택스 납부",
        "지방소득세 위택스 납부 (원천세 납부 후 즉시)",
        "납부 영수증 PDF 저장 후 '납부영수증/YYYY-MM' 폴더 보관",
    ]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # 섹션 3 : 납세증명서 발급
    # ══════════════════════════════════════════════════════════
    story.append(section_header("3. 납세증명서 발급", color=ORANGE))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 개요", S["h2"]))
    story.append(Paragraph(
        "납세증명서는 체납 없이 세금을 성실히 납부하고 있음을 증명하는 서류로, "
        "입찰·금융기관 대출·공공기관 계약 등에 제출합니다. "
        "<b>국세(홈택스)</b>와 <b>지방세(위택스)</b>를 각각 발급해야 합니다.",
        S["body"]))

    story.append(Paragraph("■ 국세 납세증명서 발급 절차 (홈택스)", S["h2"]))
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "메뉴 : <b>민원증명 &gt; 납세증명서(국세완납증명) &gt; 발급신청</b>",
        "사용 용도 선택 (금융기관 제출 / 관공서 제출 / 기타)",
        "<b>[즉시발급]</b> 선택 → 온라인 발급 (공동인증서 필요)",
        "PDF 저장 또는 출력 → 제출처에 따라 원본·사본 확인",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("■ 지방세 납세증명서 발급 절차 (위택스)", S["h2"]))
    for i, s in enumerate([
        "위택스(wetax.go.kr) 로그인",
        "메뉴 : <b>신고납부 &gt; 납세증명서 발급</b>",
        "사업장 소재 지자체 선택",
        "발급 사유 입력 후 <b>[발급]</b> → PDF 저장",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.2*cm))
    story.append(img_block(img_certificate,
        "홈택스·위택스 납세증명서 발급 화면 예시"))

    story.append(Paragraph("■ 유의사항", S["h2"]))
    for n in [
        "⚠ 유효기간 : 발급일로부터 <b>30일</b> (제출 직전 발급 권장)",
        "⚠ 체납 존재 시 발급 불가 → 즉시 납부 후 재발급",
        "⚠ 지방세는 사업장 소재 지자체 기준으로 발급 (본점·지점 별도 발급 가능)",
        "⚠ 민원24(정부24)에서도 통합 발급 가능 — 급하거나 오프라인 시 활용",
    ]:
        story.append(Paragraph(n, S["note"]))

    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("■ 발급 체크리스트", S["h2"]))
    story.append(checklist_table([
        "발급 목적 및 제출처 사전 확인",
        "국세 납세증명서 (홈택스) 발급 및 저장",
        "지방세 납세증명서 (위택스) 발급 및 저장",
        "유효기간(30일) 이내 제출 여부 확인",
        "원본 필요 여부 확인 (온라인 발급본도 원본 인정)",
    ]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # 섹션 4 : 급여대장 작성 및 회계법인 송부
    # ══════════════════════════════════════════════════════════
    story.append(section_header("4. 급여대장 작성 및 회계법인 송부",
                                color=colors.HexColor("#6a1a8a")))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 급여대장 작성 (Excel)", S["h2"]))
    story.append(Paragraph(
        "매월 급여 지급 후 <b>급여대장</b>을 작성합니다. "
        "파일명 형식 : <b>[연도]_급여대장_[월].xlsx</b> (예: 2026_급여대장_06.xlsx)",
        S["body"]))

    pay_data = [
        [Paragraph(c, S["table_hdr"]) for c in ["항목", "내용", "비고"]],
        [Paragraph("기본 정보", S["table_cell"]),
         Paragraph("성명, 주민번호(뒤 2자리 가림), 입사일, 부서", S["table_cell"]),
         Paragraph("개인정보 처리방침 준수", S["table_cell"])],
        [Paragraph("지급 항목", S["table_cell"]),
         Paragraph("기본급, 직책수당, 식대, 교통비 등", S["table_cell"]),
         Paragraph("비과세 항목 별도 구분", S["table_cell"])],
        [Paragraph("4대보험 공제", S["table_cell"]),
         Paragraph("건강·국민연금·고용·산재보험 근로자 부담분", S["table_cell"]),
         Paragraph("공단 고지 요율 적용", S["table_cell"])],
        [Paragraph("소득세·지방소득세", S["table_cell"]),
         Paragraph("간이세액표 적용 / 원천징수", S["table_cell"]),
         Paragraph("홈택스 간이세액표 참조", S["table_cell"])],
        [Paragraph("실수령액", S["table_cell"]),
         Paragraph("지급 총액 - 4대보험 - 소득세 - 지방소득세", S["table_cell"]),
         Paragraph("", S["table_cell"])],
    ]
    pay_tbl = Table(pay_data, colWidths=[3.2*cm, 8*cm, 5.4*cm])
    pay_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), colors.HexColor("#6a1a8a")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f5eefa")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#c090e0")),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(pay_tbl)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("■ 사업소득지급대장 작성 (프리랜서·용역)", S["h2"]))
    story.append(Paragraph(
        "프리랜서 또는 용역 계약 개인에게 지급 시 <b>사업소득지급대장</b>을 별도 작성합니다. "
        "파일명 : <b>[연도]_사업소득지급대장_[월].xlsx</b>",
        S["body"]))
    for s in [
        "성명 / 주민번호(뒤 6자리 가림) / 지급일 / 지급금액",
        "원천징수세액 : 지급액 × 3% (소득세) + 지급액 × 0.3% (지방소득세) = <b>3.3%</b>",
        "실지급액 = 지급금액 - 원천징수세액",
    ]:
        story.append(Paragraph(f"• {s}", S["bullet"]))

    story.append(Spacer(1, 0.2*cm))
    story.append(img_block(img_payroll,
        "급여대장 및 사업소득지급대장 작성·송부 절차 예시"))

    story.append(Paragraph("■ 회계법인 송부 절차", S["h2"]))
    for i, s in enumerate([
        "급여 지급일 기준 <b>3 영업일 이내</b> 이메일 발송",
        "첨부 파일 : 급여대장.xlsx + 사업소득지급대장.xlsx + 법인통장 사본(해당월)",
        "수신처 : 담당 세무사 이메일 (사내 주소록 '회계법인' 그룹 참조)",
        "제목 형식 : <b>[회사명] 2026년 6월 급여 관련 자료 송부</b>",
        "발송 후 수신 확인 회신 받을 것 — 미회신 시 전화 확인",
    ], 1):
        story.append(Paragraph(f"{i}. {s}", S["bullet"]))

    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("■ 체크리스트", S["h2"]))
    story.append(checklist_table([
        "급여대장 작성 완료 (기본급·수당·공제·실수령액 검토)",
        "사업소득지급대장 작성 완료 (3.3% 원천징수 적용 확인)",
        "법인 통장 사본 준비 (해당월 급여 지급 내역 포함)",
        "이메일 제목·수신처 확인 후 발송",
        "회계법인 수신 확인 회신 수령",
        "원천세 신고·납부 다음달 10일까지 완료",
    ]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════
    # 섹션 5 : 월별 업무 일정 요약
    # ══════════════════════════════════════════════════════════
    story.append(section_header("5. 월별 업무 일정 요약"))
    story.append(Spacer(1, 0.4*cm))

    cal_data = [
        [Paragraph(c, S["table_hdr"]) for c in ["시기", "업무", "마감일", "관련 시스템"]],
        ["매월 1~5일", "4대보험 고지서 확인", "납부 전", "4insure.or.kr"],
        ["매월 10일", "건강보험료 납부", "10일", "건강보험공단"],
        ["매월 10일", "원천세 (소득세) 신고·납부", "10일", "홈택스"],
        ["매월 10일", "지방소득세 신고·납부", "10일", "위택스"],
        ["매월 15일 내외", "전월 급여대장·사업소득지급대장 송부", "급여지급 후 3영업일", "이메일"],
        ["매월 말일", "국민연금·고용·산재보험 납부", "말일", "근로복지공단"],
        ["익월 10일", "세금계산서 전월분 발급 마감", "익월 10일", "홈택스"],
        ["1월 / 7월", "부가가치세 신고·납부", "25일", "홈택스"],
        ["3월", "법인세 신고·납부", "31일", "홈택스"],
        ["5월", "종합소득세 신고·납부", "31일", "홈택스"],
        ["수시", "납세증명서 발급 (요청 시)", "즉시", "홈택스·위택스"],
    ]
    cal_rows = []
    for row in cal_data:
        if isinstance(row[0], Paragraph):
            cal_rows.append(row)
        else:
            cal_rows.append([Paragraph(c, S["table_cell"]) for c in row])

    cal_tbl = Table(cal_rows, colWidths=[3.2*cm, 5.8*cm, 3.6*cm, 4*cm])
    cal_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), BLUE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eef2fa")]),
        ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#b0c0e0")),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
    ]))
    story.append(cal_tbl)

    story.append(Spacer(1, 0.6*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cccccc"), spaceAfter=8))
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

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        props = kwargs.get(side, {"val": "single", "sz": "4", "color": "B0C0E0"})
        el = OxmlElement(f"w:{side}")
        for k, v in props.items():
            el.set(qn(f"w:{k}"), v)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def add_heading_row(doc, text, level=1,
                    bg="1e4fa0", fg="FFFFFF", size=14):
    """배경색 있는 제목 단락 (표로 구현)."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.runs[0]
    run.font.name = "나눔고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    set_cell_bg(cell, bg)
    doc.add_paragraph()

def add_para(doc, text, bold=False, size=10.5, indent=0.3, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.name = "나눔고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return p

def add_bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "나눔고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    run.font.size = Pt(size)
    return p

def add_img(doc, path, caption, width_cm=14.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    cap = doc.add_paragraph(f"▲ {caption}")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in cap.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        run.font.name = "나눔고딕"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    doc.add_paragraph()

def add_simple_table(doc, headers, rows, hdr_bg="1e4fa0",
                     col_widths=None):
    cols = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=cols)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        set_cell_bg(hdr_cells[i], hdr_bg)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.name = "나눔고딕"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)

    # 데이터 행
    for ri, row in enumerate(rows):
        cells = tbl.rows[ri + 1].cells
        for ci, val in enumerate(row):
            cells[ci].text = val
            bg = "F5F8FC" if ri % 2 == 0 else "EEF2FA"
            set_cell_bg(cells[ci], bg)
            p = cells[ci].paragraphs[0]
            run = p.runs[0]
            run.font.name = "나눔고딕"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
            run.font.size = Pt(10)

    if col_widths:
        for ri in range(len(tbl.rows)):
            for ci, w in enumerate(col_widths):
                tbl.rows[ri].cells[ci].width = Cm(w)

    doc.add_paragraph()

def build_docx(out_path):
    doc = Document()

    # 기본 페이지 여백
    for section in doc.sections:
        section.page_width  = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── 표지 ──────────────────────────────────────────────────
    for _ in range(5):
        doc.add_paragraph()

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("업무 인수인계 자료")
    title_run.font.name = "나눔고딕"
    title_run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    title_run.font.size = Pt(26)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x1E, 0x4F, 0xA0)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("회계 · 세무 업무")
    sub_run.font.name = "나눔고딕"
    sub_run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")
    sub_run.font.size = Pt(16)
    sub_run.font.bold = True
    sub_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()
    add_simple_table(doc,
        ["항목", "내용"],
        [["작성일", "2026년 06월 15일"],
         ["작성 부서", "경영지원팀"],
         ["인수인계 범위", "세금계산서 발급 / 세금·보험료 납부 / 납세증명서 발급 / 급여대장 작성·송부"]],
        hdr_bg="1e4fa0",
        col_widths=[3.5, 12]
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════
    # 섹션 1
    # ══════════════════════════════════════════════════════════
    add_heading_row(doc, "1. 세금계산서 발급")
    add_para(doc, "■ 개요", bold=True)
    add_para(doc,
        "거래 상대방에게 재화·용역을 공급한 경우 국세청 홈택스를 통해 전자세금계산서를 발급합니다. "
        "발급 마감일(익월 10일)을 반드시 준수해야 하며, 미발급·지연 발급 시 가산세가 부과됩니다.")

    add_para(doc, "■ 발급 절차", bold=True)
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인 (공동인증서 또는 간편인증)",
        "메뉴 : 전자세금계산서 > 건별발급 클릭",
        "공급받는 자 사업자등록번호 입력 후 자동조회",
        "작성일자 / 공급가액 / 세액 입력 (세액 = 공급가액 × 10%)",
        "품목명·수량·단가 입력 후 [발급] 버튼 클릭",
        "발급 완료 → 상대방 이메일 자동 전송 확인",
        "발급된 세금계산서 PDF 저장 → '세금계산서/YYYY-MM' 폴더 보관",
    ], 1):
        add_bullet(doc, f"{i}. {s}")

    doc.add_paragraph()
    add_img(doc, img_tax_invoice, "홈택스 전자세금계산서 발급 화면 예시")

    add_para(doc, "■ 유의사항", bold=True)
    for n in [
        "⚠ 발급 마감 : 공급일 다음달 10일까지",
        "⚠ 수정세금계산서 : 오기재 발생 시 메뉴 > 수정발급",
        "⚠ 매입세금계산서 : 상대방 발급본 홈택스에서 수신 확인 후 보관",
        "⚠ 면세 거래 : 계산서(부가세 없음)로 발급 (전자계산서 메뉴 별도)",
    ]:
        add_para(doc, n, color=RGBColor(0xB0, 0x40, 0x00))

    add_para(doc, "■ 발급 체크리스트", bold=True)
    add_simple_table(doc, ["체크", "업무 내용"],
        [["☐", "공급받는자 사업자등록번호 및 상호 정확히 입력"],
         ["☐", "공급가액·세액 금액 확인 (세액 = 공급가액 × 10%)"],
         ["☐", "작성일자 = 실제 거래일 기준"],
         ["☐", "발급 완료 후 상대방 수신 여부 확인"],
         ["☐", "PDF 저장 후 지정 폴더에 보관"],
         ["☐", "익월 10일 마감 전 전체 발급 완료 확인"]],
        col_widths=[2, 13])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════
    # 섹션 2
    # ══════════════════════════════════════════════════════════
    add_heading_row(doc, "2. 보험료 및 세금 납부", bg="1a7a40")
    add_para(doc, "■ 4대 사회보험료 납부", bold=True)
    add_simple_table(doc,
        ["보험 종류", "납부 기한", "납부처", "비고"],
        [["건강보험", "매월 10일", "건강보험공단 가상계좌", "장기요양보험 포함"],
         ["국민연금", "매월 말일", "국민연금공단 가상계좌", ""],
         ["고용보험", "매월 말일", "근로복지공단 가상계좌", "산재보험 동시 납부"],
         ["산재보험", "매월 말일", "근로복지공단 가상계좌", ""]],
        hdr_bg="1a7a40", col_widths=[3, 3, 5.5, 4.5])

    add_para(doc, "■ 국세 납부 일정", bold=True)
    add_simple_table(doc,
        ["시기", "업무", "담당 시스템", "비고"],
        [["1월 / 7월", "부가가치세 (법인)", "홈택스", "25일"],
         ["3월", "법인세", "홈택스", "31일"],
         ["5월", "종합소득세", "홈택스", "31일"],
         ["매월", "원천세 (소득세)", "홈택스", "다음달 10일"],
         ["매월", "지방소득세", "위택스", "원천세와 동일"]],
        hdr_bg="1a7a40", col_widths=[3, 5, 5, 3])

    add_img(doc, img_insurance, "4대보험·국세·지방세 납부 화면 예시")
    add_para(doc, "■ 납부 체크리스트", bold=True)
    add_simple_table(doc, ["체크", "업무 내용"],
        [["☐", "4대사회보험 고지서 수신 확인 (매월 초)"],
         ["☐", "건강보험료 10일 내 납부 완료"],
         ["☐", "국민연금·고용·산재보험 말일 내 납부 완료"],
         ["☐", "원천세 다음달 10일까지 홈택스 납부"],
         ["☐", "지방소득세 위택스 납부 (원천세 납부 후 즉시)"],
         ["☐", "납부 영수증 PDF 저장 후 '납부영수증/YYYY-MM' 폴더 보관"]],
        col_widths=[2, 13])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════
    # 섹션 3
    # ══════════════════════════════════════════════════════════
    add_heading_row(doc, "3. 납세증명서 발급", bg="e07000")
    add_para(doc,
        "납세증명서는 체납 없이 세금을 납부하고 있음을 증명하는 서류로, "
        "입찰·금융기관 대출·공공기관 계약 등에 제출합니다. "
        "국세(홈택스)와 지방세(위택스)를 각각 발급합니다.")

    add_para(doc, "■ 국세 납세증명서 발급 (홈택스)", bold=True)
    for i, s in enumerate([
        "홈택스(hometax.go.kr) 로그인",
        "메뉴 : 민원증명 > 납세증명서(국세완납증명) > 발급신청",
        "사용 용도 선택 후 [즉시발급] 클릭 → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")

    add_para(doc, "■ 지방세 납세증명서 발급 (위택스)", bold=True)
    for i, s in enumerate([
        "위택스(wetax.go.kr) 로그인",
        "메뉴 : 신고납부 > 납세증명서 발급",
        "지자체 선택 후 [발급] → PDF 저장",
    ], 1):
        add_bullet(doc, f"{i}. {s}")

    add_img(doc, img_certificate, "홈택스·위택스 납세증명서 발급 화면 예시")

    add_para(doc, "■ 유의사항", bold=True)
    for n in [
        "⚠ 유효기간 : 발급일로부터 30일 (제출 직전 발급 권장)",
        "⚠ 체납 존재 시 발급 불가 → 즉시 납부 후 재발급",
        "⚠ 지방세는 사업장 소재 지자체 기준으로 발급",
    ]:
        add_para(doc, n, color=RGBColor(0xB0, 0x40, 0x00))

    add_para(doc, "■ 발급 체크리스트", bold=True)
    add_simple_table(doc, ["체크", "업무 내용"],
        [["☐", "발급 목적 및 제출처 사전 확인"],
         ["☐", "국세 납세증명서 (홈택스) 발급 및 저장"],
         ["☐", "지방세 납세증명서 (위택스) 발급 및 저장"],
         ["☐", "유효기간(30일) 이내 제출 여부 확인"],
         ["☐", "원본 필요 여부 확인"]],
        col_widths=[2, 13])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════
    # 섹션 4
    # ══════════════════════════════════════════════════════════
    add_heading_row(doc, "4. 급여대장 작성 및 회계법인 송부", bg="6a1a8a")
    add_para(doc, "■ 급여대장 작성 항목", bold=True)
    add_simple_table(doc,
        ["항목", "내용", "비고"],
        [["기본 정보", "성명, 주민번호(뒤 2자리 가림), 입사일, 부서", "개인정보 처리방침 준수"],
         ["지급 항목", "기본급, 직책수당, 식대, 교통비 등", "비과세 항목 별도 구분"],
         ["4대보험 공제", "건강·국민·고용·산재 근로자 부담분", "공단 고지 요율 적용"],
         ["소득세·지방소득세", "간이세액표 적용 / 원천징수", "홈택스 간이세액표 참조"],
         ["실수령액", "지급 총액 - 4대보험 - 소득세 - 지방소득세", ""]],
        hdr_bg="6a1a8a", col_widths=[3.5, 8, 5])

    add_para(doc, "■ 사업소득지급대장 (프리랜서)", bold=True)
    for s in [
        "성명 / 주민번호(뒤 6자리 가림) / 지급일 / 지급금액",
        "원천징수 : 지급액 × 3.3% (소득세 3% + 지방소득세 0.3%)",
        "실지급액 = 지급금액 - 원천징수세액",
    ]:
        add_bullet(doc, s)

    add_img(doc, img_payroll, "급여대장 및 사업소득지급대장 작성·송부 절차 예시")

    add_para(doc, "■ 회계법인 송부 절차", bold=True)
    for i, s in enumerate([
        "급여 지급일 기준 3 영업일 이내 이메일 발송",
        "첨부 파일 : 급여대장.xlsx + 사업소득지급대장.xlsx + 통장사본",
        "수신처 : 담당 세무사 이메일 (주소록 '회계법인' 그룹)",
        "제목 형식 : [회사명] 2026년 6월 급여 관련 자료 송부",
        "발송 후 수신 확인 회신 받을 것",
    ], 1):
        add_bullet(doc, f"{i}. {s}")

    add_para(doc, "■ 체크리스트", bold=True)
    add_simple_table(doc, ["체크", "업무 내용"],
        [["☐", "급여대장 작성 완료 (기본급·수당·공제·실수령액 검토)"],
         ["☐", "사업소득지급대장 작성 완료 (3.3% 원천징수 적용 확인)"],
         ["☐", "법인 통장 사본 준비 (해당월 급여 지급 내역 포함)"],
         ["☐", "이메일 제목·수신처 확인 후 발송"],
         ["☐", "회계법인 수신 확인 회신 수령"],
         ["☐", "원천세 신고·납부 다음달 10일까지 완료"]],
        col_widths=[2, 13])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════
    # 섹션 5 : 월별 업무 일정 요약
    # ══════════════════════════════════════════════════════════
    add_heading_row(doc, "5. 월별 업무 일정 요약")
    add_simple_table(doc,
        ["시기", "업무", "마감일", "관련 시스템"],
        [["매월 1~5일", "4대보험 고지서 확인", "납부 전", "4insure.or.kr"],
         ["매월 10일", "건강보험료 납부", "10일", "건강보험공단"],
         ["매월 10일", "원천세 신고·납부", "10일", "홈택스"],
         ["매월 10일", "지방소득세 신고·납부", "10일", "위택스"],
         ["매월 15일 내외", "급여대장·사업소득지급대장 회계법인 송부", "급여지급 후 3영업일", "이메일"],
         ["매월 말일", "국민연금·고용·산재보험 납부", "말일", "근로복지공단"],
         ["익월 10일", "세금계산서 전월분 발급 마감", "익월 10일", "홈택스"],
         ["1월 / 7월", "부가가치세 신고·납부", "25일", "홈택스"],
         ["3월", "법인세 신고·납부", "31일", "홈택스"],
         ["5월", "종합소득세 신고·납부", "31일", "홈택스"],
         ["수시", "납세증명서 발급 (요청 시)", "즉시", "홈택스·위택스"]],
        col_widths=[3.5, 6, 3.5, 4])

    doc.add_paragraph()
    footer_p = doc.add_paragraph(
        "본 자료는 인수인계 목적으로 작성되었습니다. "
        "법령 개정 또는 사내 규정 변경 시 담당자가 업데이트하십시오.")
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer_p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        run.font.name = "나눔고딕"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "나눔고딕")

    doc.save(out_path)
    print(f"DOCX 생성 완료 : {out_path}")


if __name__ == "__main__":
    pdf_path  = os.path.join(OUT_DIR, "회계세무_업무인수인계.pdf")
    docx_path = os.path.join(OUT_DIR, "회계세무_업무인수인계.docx")
    build_pdf(pdf_path)
    build_docx(docx_path)
    print("모든 파일 생성 완료.")
