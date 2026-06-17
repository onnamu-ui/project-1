"""
인사·총무 업무 인수인계 자료
sections:
  1. 입사 처리
  2. 퇴사 처리
  3. 근태 관리
  4. 휴가 신청·관리
  5. 4대보험 취득·상실 신고
  6. 급여 외 인사 행정
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
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── fonts ──────────────────────────────────────────────────────────────────
FONT_DIR = "/usr/share/fonts/truetype/nanum"
REG  = os.path.join(FONT_DIR, "NanumGothic.ttf")
BOLD = os.path.join(FONT_DIR, "NanumGothicBold.ttf")
SQB  = os.path.join(FONT_DIR, "NanumSquareB.ttf")

pdfmetrics.registerFont(TTFont("NG",    REG))
pdfmetrics.registerFont(TTFont("NGBold", BOLD))
pdfmetrics.registerFont(TTFont("NSB",   SQB))

OUT_DIR  = "/home/user/project-1/handover_docs"
IMG_DIR  = OUT_DIR
os.makedirs(OUT_DIR, exist_ok=True)

BRAND    = "#1a3a6b"        # 주색
ACCENT   = "#e8eef7"        # 연한 배경
W, H_A4  = A4

# ── PIL image helper ────────────────────────────────────────────────────────
def pil_font(size, bold=False):
    try:
        return ImageFont.truetype(BOLD if bold else REG, size)
    except Exception:
        return ImageFont.load_default()

def make_img(path, title, lines, width=760, height=340,
             bar_color="#1a3a6b", line_color="#3a7bd5"):
    img = Image.new("RGB", (width, height), "#f8f9fc")
    d   = ImageDraw.Draw(img)
    # header bar
    d.rectangle([0, 0, width, 52], fill=bar_color)
    d.text((20, 12), title, font=pil_font(22, bold=True), fill="white")
    # divider
    d.rectangle([0, 52, width, 55], fill=line_color)
    y = 72
    for item in lines:
        if item.startswith("##"):          # sub-header
            d.rectangle([6, y, 16, y+22], fill=line_color)
            d.text((24, y), item[2:].strip(), font=pil_font(17, bold=True), fill="#1a3a6b")
            y += 30
        elif item.startswith("--"):        # divider
            d.rectangle([16, y+6, width-16, y+7], fill="#dee2e9")
            y += 18
        else:
            d.ellipse([18, y+7, 26, y+15], fill=line_color)
            d.text((34, y), item, font=pil_font(15), fill="#333333")
            y += 26
        if y > height - 20:
            break
    img.save(path)
    return path

# ── PDF helpers ─────────────────────────────────────────────────────────────
def S(name, **kw):
    base = dict(fontName="NG", fontSize=11, leading=18, textColor=colors.HexColor("#333333"))
    base.update(kw)
    return ParagraphStyle(name, **base)

SEC_HDR = S("sec_hdr", fontName="NSB", fontSize=15, textColor=colors.HexColor(BRAND),
            spaceBefore=14, spaceAfter=4)
BODY    = S("body", spaceBefore=2, spaceAfter=2)
BULLET  = S("bullet", leftIndent=16, spaceBefore=1, spaceAfter=1)
NOTE_S  = S("note", fontName="NG", fontSize=10, textColor=colors.HexColor("#555555"),
            backColor=colors.HexColor("#f0f4fb"), leftIndent=10, rightIndent=10,
            spaceBefore=4, spaceAfter=4)

def hr():
    return HRFlowable(width="100%", thickness=1, color=colors.HexColor("#c8d4e8"), spaceAfter=6)

def sec_hdr(txt):
    return [Paragraph(txt, SEC_HDR), hr()]

def para(txt):
    return Paragraph(txt, BODY)

def bullet(txt):
    return Paragraph(f"• &nbsp;{txt}", BULLET)

def note(txt):
    return Paragraph(f"※ {txt}", NOTE_S)

def add_img(path, w_mm=155):
    return RLImage(path, width=w_mm*mm, height=w_mm*mm*0.42)

def mk_table(data, col_w, hdr_color=BRAND):
    t = Table(data, colWidths=col_w)
    style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(hdr_color)),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "NGBold"),
        ("FONTSIZE",   (0,0), (-1,-1), 10),
        ("FONTNAME",   (0,1), (-1,-1), "NG"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4fb")]),
        ("GRID",       (0,0), (-1,-1), 0.4, colors.HexColor("#b0bfd0")),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ]
    t.setStyle(TableStyle(style))
    return t


# ══════════════════════════════════════════════════════════════════════════════
#  IMAGES
# ══════════════════════════════════════════════════════════════════════════════
def make_images():
    make_img(f"{IMG_DIR}/img_hr01_onboard.png",
             "입사 처리 프로세스",
             ["## 입사 서류 수령",
              "근로계약서 작성 (2부 서명 후 1부 직원 교부)",
              "주민등록증 사본, 통장 사본, 학력증명서 수령",
              "## 시스템 등록",
              "4대보험 취득 신고 (취득일로부터 14일 이내)",
              "급여 계좌 등록 및 세금 정보 확인",
              "## 업무 안내",
              "슬랙·드롭박스·Zoom 계정 생성 및 안내",
              "사내 규정, 복무 규정 전달"],
             height=300)

    make_img(f"{IMG_DIR}/img_hr02_offboard.png",
             "퇴사 처리 프로세스",
             ["## 퇴사 확인",
              "퇴직 의사 확인 및 퇴직일 결정",
              "퇴직 처리 내부 공유 (대표 보고)",
              "## 행정 처리",
              "4대보험 상실 신고 (상실일로부터 14일 이내)",
              "퇴직금 정산 (1년 이상 근무 시)",
              "마지막 급여 정산 (해당 월 일할 계산)",
              "## 계정·자산 반납",
              "노트북·사원증 등 사내 자산 반납",
              "슬랙·드롭박스 등 계정 권한 해제"],
             height=310)

    make_img(f"{IMG_DIR}/img_hr03_attendance.png",
             "근태 관리",
             ["## 출·퇴근 기록",
              "슬랙 #1-근무및휴가 채널에 출근/퇴근 메시지 작성",
              "재택 근무 시 동일하게 슬랙 채널에 기록",
              "## 지각·조퇴·외출",
              "당일 슬랙 채널에 사유 및 예상 시간 공유",
              "대표님 사전 승인 후 처리",
              "## 월말 근태 확인",
              "월말 근태 현황 취합 → 급여 정산에 반영"],
             height=290)

    make_img(f"{IMG_DIR}/img_hr04_leave.png",
             "휴가 신청·관리",
             ["## 연차 발생 기준",
              "입사 1년 미만: 매월 1일씩 발생 (최대 11일)",
              "입사 1년 이상: 15일 기본 + 2년마다 1일 추가",
              "## 휴가 신청 방법",
              "슬랙 #1-근무및휴가 채널에 휴가 신청 메시지 작성",
              "대표님 승인 후 확정",
              "## 반차·특별 휴가",
              "반차(오전/오후) 신청 시 시간 명시",
              "경조사 휴가 등 별도 규정 확인"],
             height=300)

    make_img(f"{IMG_DIR}/img_hr05_insurance.png",
             "4대보험 취득·상실 신고",
             ["## 신고 사이트",
              "4대사회보험 정보연계센터 (4insure.or.kr)",
              "공동인증서(사업자) 로그인 필요",
              "## 취득 신고 (입사)",
              "근로자 취득 신고 메뉴 → 신규 등록",
              "취득일로부터 14일 이내 신고",
              "월 보수액 입력 (국민연금·건강보험 기준)",
              "## 상실 신고 (퇴사)",
              "근로자 상실 신고 메뉴 → 상실 등록",
              "상실일로부터 14일 이내 신고"],
             height=330)

    make_img(f"{IMG_DIR}/img_hr06_admin.png",
             "급여 외 인사 행정",
             ["## 재직증명서 발급",
              "요청 시 회사 양식으로 작성 → 직인 날인 후 교부",
              "## 경력증명서 발급",
              "퇴직자 요청 시 발급 (재직 기간·직위 기재)",
              "## 근로계약서 갱신",
              "매년 또는 계약 만료 시 갱신 작성",
              "## 인사 파일 관리",
              "직원별 서류 드롭박스 계약/행정서류 폴더에 보관",
              "개인정보 보호 주의 (접근 권한 관리)"],
             height=300)


# ══════════════════════════════════════════════════════════════════════════════
#  PDF
# ══════════════════════════════════════════════════════════════════════════════
def build_pdf(out_path):
    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=20*mm, rightMargin=20*mm,
                            topMargin=18*mm, bottomMargin=18*mm)
    E = []

    # ── cover ──
    E.append(Spacer(1, 30*mm))
    E.append(Paragraph("인사·총무 업무", S("t1", fontName="NSB", fontSize=28,
                        textColor=colors.HexColor(BRAND), alignment=1)))
    E.append(Spacer(1, 4*mm))
    E.append(Paragraph("인수인계 자료", S("t2", fontName="NSB", fontSize=22,
                        textColor=colors.HexColor("#3a7bd5"), alignment=1)))
    E.append(Spacer(1, 8*mm))
    E.append(HRFlowable(width="60%", thickness=2, color=colors.HexColor(BRAND),
                         hAlign="CENTER"))
    E.append(Spacer(1, 6*mm))
    E.append(Paragraph("ONNAMU", S("co", fontName="NGBold", fontSize=13,
                        textColor=colors.HexColor("#555555"), alignment=1)))
    E.append(Spacer(1, 40*mm))

    # 목차
    toc_data = [
        ["No.", "섹션", "주요 내용"],
        ["01", "입사 처리", "서류 수령 · 4대보험 취득 · 계정 생성"],
        ["02", "퇴사 처리", "4대보험 상실 · 퇴직금 정산 · 자산 반납"],
        ["03", "근태 관리", "슬랙 기록 · 지각/조퇴 처리 · 월말 집계"],
        ["04", "휴가 신청·관리", "연차 발생 기준 · 신청 방법 · 반차"],
        ["05", "4대보험 취득·상실 신고", "4insure.or.kr · 취득/상실 신고 절차"],
        ["06", "급여 외 인사 행정", "재직·경력증명서 · 근로계약서 갱신 · 파일 관리"],
    ]
    E.append(mk_table(toc_data,
                       [12*mm, 55*mm, 98*mm]))
    E.append(PageBreak())

    CW = W - 40*mm  # content width

    # ── SEC 1 입사 처리 ──────────────────────────────────────────────────────
    E += sec_hdr("01  입사 처리")
    E.append(para("신규 직원 입사 시 아래 순서로 행정 처리를 진행합니다."))
    E.append(Spacer(1, 3*mm))

    step_data = [
        ["단계", "업무", "비고"],
        ["1", "근로계약서 작성", "2부 작성 → 1부 직원 교부, 1부 회사 보관"],
        ["2", "입사 서류 수령", "주민등록증 사본, 통장 사본, 최종학력증명서"],
        ["3", "4대보험 취득 신고", "4insure.or.kr, 취득일로부터 14일 이내"],
        ["4", "급여 정보 등록", "급여 계좌번호, 주민등록번호(세금 신고용)"],
        ["5", "사내 계정 생성", "슬랙, 드롭박스, Zoom 초대 및 채널 안내"],
        ["6", "규정 안내", "복무 규정, 보안 규정, 업무 매뉴얼 전달"],
    ]
    E.append(mk_table(step_data, [15*mm, 55*mm, 95*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr01_onboard.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("근로계약서는 반드시 입사일 전 또는 당일 작성 완료. 미작성 시 노동법 위반."))
    E.append(PageBreak())

    # ── SEC 2 퇴사 처리 ──────────────────────────────────────────────────────
    E += sec_hdr("02  퇴사 처리")
    E.append(para("직원 퇴사 시 아래 순서로 행정 처리를 진행합니다."))
    E.append(Spacer(1, 3*mm))

    step2_data = [
        ["단계", "업무", "비고"],
        ["1", "퇴직 의사 확인", "퇴직일 결정 후 대표님 보고"],
        ["2", "4대보험 상실 신고", "4insure.or.kr, 상실일로부터 14일 이내"],
        ["3", "마지막 급여 정산", "퇴직일 기준 일할 계산, 미사용 연차 수당 포함"],
        ["4", "퇴직금 정산", "1년 이상 근무 시 지급 (평균 임금 × 30일 × 근속연수)"],
        ["5", "자산 반납", "노트북, 사원증, 법인카드 등 회사 자산 회수"],
        ["6", "계정 권한 해제", "슬랙, 드롭박스, 이메일 등 접근 권한 즉시 해제"],
        ["7", "서류 발급", "퇴직 후 경력증명서 요청 시 발급"],
    ]
    E.append(mk_table(step2_data, [15*mm, 55*mm, 95*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr02_offboard.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("퇴직금은 퇴직일로부터 14일 이내 지급 원칙 (합의 시 연장 가능)."))
    E.append(PageBreak())

    # ── SEC 3 근태 관리 ──────────────────────────────────────────────────────
    E += sec_hdr("03  근태 관리")
    E.append(para("출퇴근 및 근무 상황은 슬랙 채널을 통해 기록·관리합니다."))
    E.append(Spacer(1, 3*mm))

    attend_data = [
        ["구분", "방법", "채널/비고"],
        ["출근", "출근 메시지 작성", "#1-근무및휴가 (예: '출근합니다')"],
        ["퇴근", "퇴근 메시지 작성", "#1-근무및휴가 (예: '퇴근합니다')"],
        ["재택 근무", "재택 여부 메시지 작성", "#1-근무및휴가 (예: '오늘 재택 근무합니다')"],
        ["지각·조퇴", "사유 및 예상 시간 공유", "대표님 사전 승인 필요"],
        ["외출", "외출 시간 및 사유 공유", "대표님 사전 승인 필요"],
        ["월말 집계", "월별 근태 현황 취합", "급여 정산 반영용, 매월 말 작성"],
    ]
    E.append(mk_table(attend_data, [25*mm, 55*mm, 85*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr03_attendance.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("근태 기록은 급여 정산 기준이 되므로 빠짐없이 작성 필요."))
    E.append(PageBreak())

    # ── SEC 4 휴가 ───────────────────────────────────────────────────────────
    E += sec_hdr("04  휴가 신청·관리")
    E.append(para("연차 휴가 발생 기준 및 신청 방법을 안내합니다."))
    E.append(Spacer(1, 3*mm))

    leave_data = [
        ["구분", "발생 기준", "비고"],
        ["입사 1년 미만", "매월 1일씩 발생 (최대 11일)", "만근 시 1일 발생"],
        ["입사 1년 이상", "15일 기본 부여", "매년 초 부여"],
        ["3년 이상", "15일 + 2년마다 1일 추가", "최대 25일"],
        ["반차", "오전(09:00~13:00) / 오후(13:00~18:00)", "0.5일 차감"],
        ["경조사 휴가", "결혼 5일, 부모 상 5일 등", "별도 규정 확인"],
    ]
    E.append(mk_table(leave_data, [35*mm, 70*mm, 60*mm]))
    E.append(Spacer(1, 4*mm))

    req_data = [
        ["신청 방법", "내용"],
        ["채널", "슬랙 #1-근무및휴가 채널에 신청 메시지 작성"],
        ["내용 포함사항", "휴가 종류, 날짜(시간), 사유 간략 기재"],
        ["승인", "대표님 확인 후 승인 메시지로 확정"],
        ["사전 신청", "가급적 3일 전 신청 (긴급 시 당일 신청 가능)"],
    ]
    E.append(mk_table(req_data, [35*mm, 130*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr04_leave.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("미사용 연차는 연도 말 소멸 또는 수당으로 정산 (사내 규정 확인)."))
    E.append(PageBreak())

    # ── SEC 5 4대보험 ────────────────────────────────────────────────────────
    E += sec_hdr("05  4대보험 취득·상실 신고")
    E.append(para("4대사회보험 정보연계센터(4insure.or.kr)에서 일괄 신고합니다."))
    E.append(Spacer(1, 3*mm))

    ins_data = [
        ["구분", "신고 기한", "주요 입력 항목", "비고"],
        ["취득 신고\n(입사)", "취득일로부터\n14일 이내",
         "성명, 주민번호, 취득일,\n월 보수액",
         "공동인증서(사업자)\n로그인 필요"],
        ["상실 신고\n(퇴사)", "상실일로부터\n14일 이내",
         "성명, 상실일,\n상실 사유 코드",
         "퇴직 사유 코드\n정확히 입력"],
        ["보수 변경", "변경 월\n다음 달 15일까지",
         "변경된 월 보수액",
         "급여 인상 시 반드시\n신고"],
    ]
    E.append(mk_table(ins_data, [25*mm, 30*mm, 60*mm, 50*mm]))
    E.append(Spacer(1, 3*mm))

    ins_type = [
        ["보험 종류", "사업주 부담률", "근로자 부담률", "비고"],
        ["국민연금", "4.5%", "4.5%", "월 보수액 기준"],
        ["건강보험", "3.545%", "3.545%", "장기요양보험 별도"],
        ["고용보험", "0.9%", "0.9%", "업종별 상이"],
        ["산재보험", "업종별 상이", "-", "전액 사업주 부담"],
    ]
    E.append(mk_table(ins_type, [35*mm, 35*mm, 35*mm, 60*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr05_insurance.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("신고 지연 시 과태료 발생 가능. 취득/상실 즉시 처리 원칙."))
    E.append(PageBreak())

    # ── SEC 6 급여 외 인사 행정 ──────────────────────────────────────────────
    E += sec_hdr("06  급여 외 인사 행정")
    E.append(para("재직증명서, 경력증명서 발급 및 인사 파일 관리 방법입니다."))
    E.append(Spacer(1, 3*mm))

    doc_data = [
        ["문서", "발급 대상", "처리 방법", "비고"],
        ["재직증명서", "현직 직원",
         "회사 양식 작성 →\n대표이사 직인 날인",
         "요청 후 당일\n또는 익일 교부"],
        ["경력증명서", "퇴직자",
         "재직 기간·직위·담당 업무\n기재 후 직인 날인",
         "퇴직 후\n언제든 요청 가능"],
        ["근로계약서\n갱신", "전 직원",
         "계약 만료 전 새 계약서\n작성·서명",
         "매년 또는\n조건 변경 시"],
    ]
    E.append(mk_table(doc_data, [28*mm, 25*mm, 60*mm, 52*mm]))
    E.append(Spacer(1, 4*mm))

    file_data = [
        ["항목", "보관 위치", "내용"],
        ["입사 서류", "드롭박스 > 계약/행정서류", "근로계약서, 주민등록증 사본 등"],
        ["근태·휴가 기록", "드롭박스 > 계약/행정서류", "월별 근태 현황 파일"],
        ["4대보험 신고 내역", "드롭박스 > 계약/행정서류", "취득·상실 신고서 PDF"],
        ["발급 서류 사본", "드롭박스 > 계약/행정서류", "재직·경력증명서 사본"],
    ]
    E.append(mk_table(file_data, [35*mm, 55*mm, 75*mm]))
    E.append(Spacer(1, 4*mm))
    E.append(add_img(f"{IMG_DIR}/img_hr06_admin.png"))
    E.append(Spacer(1, 3*mm))
    E.append(note("직원 개인정보 서류는 드롭박스 내 접근 권한 설정 후 보관. 외부 유출 금지."))

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

def add_sec_hdr(doc, txt):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(txt)
    run.bold      = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)
    ko(run)
    doc.add_paragraph("─" * 52).paragraph_format.space_after = Pt(4)

def add_para(doc, txt, size=11):
    p = doc.add_paragraph()
    run = p.add_run(txt)
    run.font.size = Pt(size)
    ko(run)

def add_bullet(doc, txt):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(txt)
    run.font.size = Pt(11)
    ko(run)

def add_note(doc, txt):
    p   = doc.add_paragraph()
    run = p.add_run(f"※ {txt}")
    run.font.size  = Pt(10)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    ko(run)

def add_table(doc, data, col_widths, hdr_hex="1a3a6b"):
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

def add_img_docx(doc, path, width_cm=16):
    try:
        doc.add_picture(path, width=Cm(width_cm))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  DOCX
# ══════════════════════════════════════════════════════════════════════════════
def build_docx(out_path):
    doc = Document()

    # page margins
    for sec in doc.sections:
        sec.top_margin    = Cm(2)
        sec.bottom_margin = Cm(2)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2.5)

    # cover
    doc.add_paragraph()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("인사·총무 업무 인수인계 자료")
    r.bold = True; r.font.size = Pt(24)
    r.font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)
    ko(r)
    doc.add_paragraph()
    t2 = doc.add_paragraph()
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = t2.add_run("ONNAMU")
    r2.font.size = Pt(14); r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    ko(r2)
    doc.add_page_break()

    # SEC 1
    add_sec_hdr(doc, "01  입사 처리")
    add_para(doc, "신규 직원 입사 시 아래 순서로 행정 처리를 진행합니다.")
    doc.add_paragraph()
    add_table(doc,
        [["단계", "업무", "비고"],
         ["1", "근로계약서 작성", "2부 작성 → 1부 직원 교부, 1부 회사 보관"],
         ["2", "입사 서류 수령", "주민등록증 사본, 통장 사본, 최종학력증명서"],
         ["3", "4대보험 취득 신고", "4insure.or.kr, 취득일로부터 14일 이내"],
         ["4", "급여 정보 등록", "급여 계좌번호, 주민등록번호(세금 신고용)"],
         ["5", "사내 계정 생성", "슬랙, 드롭박스, Zoom 초대 및 채널 안내"],
         ["6", "규정 안내", "복무 규정, 보안 규정, 업무 매뉴얼 전달"]],
        [1.5, 5.5, 9.5])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr01_onboard.png")
    doc.add_paragraph()
    add_note(doc, "근로계약서는 반드시 입사일 전 또는 당일 작성 완료. 미작성 시 노동법 위반.")
    doc.add_page_break()

    # SEC 2
    add_sec_hdr(doc, "02  퇴사 처리")
    add_para(doc, "직원 퇴사 시 아래 순서로 행정 처리를 진행합니다.")
    doc.add_paragraph()
    add_table(doc,
        [["단계", "업무", "비고"],
         ["1", "퇴직 의사 확인", "퇴직일 결정 후 대표님 보고"],
         ["2", "4대보험 상실 신고", "4insure.or.kr, 상실일로부터 14일 이내"],
         ["3", "마지막 급여 정산", "퇴직일 기준 일할 계산, 미사용 연차 수당 포함"],
         ["4", "퇴직금 정산", "1년 이상 근무 시 지급 (평균 임금 × 30일 × 근속연수)"],
         ["5", "자산 반납", "노트북, 사원증, 법인카드 등 회사 자산 회수"],
         ["6", "계정 권한 해제", "슬랙, 드롭박스, 이메일 등 접근 권한 즉시 해제"],
         ["7", "서류 발급", "퇴직 후 경력증명서 요청 시 발급"]],
        [1.5, 5.5, 9.5])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr02_offboard.png")
    doc.add_paragraph()
    add_note(doc, "퇴직금은 퇴직일로부터 14일 이내 지급 원칙 (합의 시 연장 가능).")
    doc.add_page_break()

    # SEC 3
    add_sec_hdr(doc, "03  근태 관리")
    add_para(doc, "출퇴근 및 근무 상황은 슬랙 채널을 통해 기록·관리합니다.")
    doc.add_paragraph()
    add_table(doc,
        [["구분", "방법", "채널/비고"],
         ["출근", "출근 메시지 작성", "#1-근무및휴가 (예: '출근합니다')"],
         ["퇴근", "퇴근 메시지 작성", "#1-근무및휴가 (예: '퇴근합니다')"],
         ["재택 근무", "재택 여부 메시지 작성", "#1-근무및휴가 (예: '오늘 재택 근무합니다')"],
         ["지각·조퇴", "사유 및 예상 시간 공유", "대표님 사전 승인 필요"],
         ["외출", "외출 시간 및 사유 공유", "대표님 사전 승인 필요"],
         ["월말 집계", "월별 근태 현황 취합", "급여 정산 반영용, 매월 말 작성"]],
        [2.5, 5.5, 8.5])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr03_attendance.png")
    doc.add_paragraph()
    add_note(doc, "근태 기록은 급여 정산 기준이 되므로 빠짐없이 작성 필요.")
    doc.add_page_break()

    # SEC 4
    add_sec_hdr(doc, "04  휴가 신청·관리")
    add_para(doc, "연차 휴가 발생 기준 및 신청 방법을 안내합니다.")
    doc.add_paragraph()
    add_table(doc,
        [["구분", "발생 기준", "비고"],
         ["입사 1년 미만", "매월 1일씩 발생 (최대 11일)", "만근 시 1일 발생"],
         ["입사 1년 이상", "15일 기본 부여", "매년 초 부여"],
         ["3년 이상", "15일 + 2년마다 1일 추가", "최대 25일"],
         ["반차", "오전(09:00~13:00) / 오후(13:00~18:00)", "0.5일 차감"],
         ["경조사 휴가", "결혼 5일, 부모 상 5일 등", "별도 규정 확인"]],
        [3.5, 7.0, 6.0])
    doc.add_paragraph()
    add_table(doc,
        [["신청 방법", "내용"],
         ["채널", "슬랙 #1-근무및휴가 채널에 신청 메시지 작성"],
         ["내용 포함사항", "휴가 종류, 날짜(시간), 사유 간략 기재"],
         ["승인", "대표님 확인 후 승인 메시지로 확정"],
         ["사전 신청", "가급적 3일 전 신청 (긴급 시 당일 신청 가능)"]],
        [3.5, 13.0])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr04_leave.png")
    doc.add_paragraph()
    add_note(doc, "미사용 연차는 연도 말 소멸 또는 수당으로 정산 (사내 규정 확인).")
    doc.add_page_break()

    # SEC 5
    add_sec_hdr(doc, "05  4대보험 취득·상실 신고")
    add_para(doc, "4대사회보험 정보연계센터(4insure.or.kr)에서 일괄 신고합니다.")
    doc.add_paragraph()
    add_table(doc,
        [["구분", "신고 기한", "주요 입력 항목", "비고"],
         ["취득 신고(입사)", "취득일로부터 14일 이내", "성명, 주민번호, 취득일, 월 보수액", "공동인증서(사업자) 로그인 필요"],
         ["상실 신고(퇴사)", "상실일로부터 14일 이내", "성명, 상실일, 상실 사유 코드", "퇴직 사유 코드 정확히 입력"],
         ["보수 변경", "변경 월 다음 달 15일까지", "변경된 월 보수액", "급여 인상 시 반드시 신고"]],
        [3.0, 4.0, 6.0, 3.5])
    doc.add_paragraph()
    add_table(doc,
        [["보험 종류", "사업주 부담률", "근로자 부담률", "비고"],
         ["국민연금", "4.5%", "4.5%", "월 보수액 기준"],
         ["건강보험", "3.545%", "3.545%", "장기요양보험 별도"],
         ["고용보험", "0.9%", "0.9%", "업종별 상이"],
         ["산재보험", "업종별 상이", "-", "전액 사업주 부담"]],
        [3.5, 3.5, 3.5, 6.0])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr05_insurance.png")
    doc.add_paragraph()
    add_note(doc, "신고 지연 시 과태료 발생 가능. 취득/상실 즉시 처리 원칙.")
    doc.add_page_break()

    # SEC 6
    add_sec_hdr(doc, "06  급여 외 인사 행정")
    add_para(doc, "재직증명서, 경력증명서 발급 및 인사 파일 관리 방법입니다.")
    doc.add_paragraph()
    add_table(doc,
        [["문서", "발급 대상", "처리 방법", "비고"],
         ["재직증명서", "현직 직원", "회사 양식 작성 → 대표이사 직인 날인", "요청 후 당일 또는 익일 교부"],
         ["경력증명서", "퇴직자", "재직 기간·직위·담당 업무 기재 후 직인 날인", "퇴직 후 언제든 요청 가능"],
         ["근로계약서 갱신", "전 직원", "계약 만료 전 새 계약서 작성·서명", "매년 또는 조건 변경 시"]],
        [3.0, 2.5, 6.5, 4.5])
    doc.add_paragraph()
    add_table(doc,
        [["항목", "보관 위치", "내용"],
         ["입사 서류", "드롭박스 > 계약/행정서류", "근로계약서, 주민등록증 사본 등"],
         ["근태·휴가 기록", "드롭박스 > 계약/행정서류", "월별 근태 현황 파일"],
         ["4대보험 신고 내역", "드롭박스 > 계약/행정서류", "취득·상실 신고서 PDF"],
         ["발급 서류 사본", "드롭박스 > 계약/행정서류", "재직·경력증명서 사본"]],
        [4.0, 5.5, 7.0])
    doc.add_paragraph()
    add_img_docx(doc, f"{IMG_DIR}/img_hr06_admin.png")
    doc.add_paragraph()
    add_note(doc, "직원 개인정보 서류는 드롭박스 내 접근 권한 설정 후 보관. 외부 유출 금지.")

    doc.save(out_path)
    print(f"DOCX 생성 완료 : {out_path}")


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    make_images()
    build_pdf(f"{OUT_DIR}/인사총무_업무인수인계.pdf")
    build_docx(f"{OUT_DIR}/인사총무_업무인수인계.docx")
    print("모든 파일 생성 완료.")
