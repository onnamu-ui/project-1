# -*- coding: utf-8 -*-
"""2026 한복인문학 특강 - 내부용 상세 예산표 (v2) DOCX 생성
- 회색(온나무 직접 수행) 위쪽으로 모음, 흰색(용역사 위탁) 아래로 배치
- 회색 추가: SNS 초안 제작비, 전체 계획 및 발주 획득 관리, 결과보고서 작성 완성
- 흰색 추가: 결과보고서 재료 정리 및 저장
- 흰색 총합 = 7,000,000 (부가세 포함)
"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GRAY = "D9D9D9"
HEADER = "404040"

def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_cell(cell, text, bold=False, align='center', white=False, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'left': WD_ALIGN_PARAGRAPH.LEFT, 'right': WD_ALIGN_PARAGRAPH.RIGHT}[align]
    run = p.add_run(str(text))
    run.font.size = Pt(size); run.font.name = '맑은 고딕'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    run.bold = bold
    if white:
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc = Document()
for s in doc.sections:
    s.left_margin = Cm(1.8); s.right_margin = Cm(1.8)

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("「2026 한복인문학 특강」 상세 예산표 (내부용)")
r.bold = True; r.font.size = Pt(15); r.font.name = '맑은 고딕'
r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

sub = doc.add_paragraph()
sr = sub.add_run("※ 회색 음영 = 온나무 직접 수행 영역(내부 처리, 상단) / 백색 = 용역사 위탁 범위(하단, 합계 7,000,000원 부가세 포함)")
sr.font.size = Pt(9); sr.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
sr.font.name = '맑은 고딕'; sr._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

headers = ["구분", "항목", "세부내용", "단위", "수량", "단가", "합계", "담당(역할)", "비고"]

# 회색(온나무 직접) — 위쪽
gray_rows = [
    ("강사료",  "인문학특강 강사료",       "회", "6", "", "", "기획팀(온나무)",       "교통비 등 포함 / 내부 처리"),
    ("인건비",  "사회진행비",              "회", "6", "", "", "기획팀(온나무)",       "사회자 교통비 포함 / 내부 처리"),
    ("인쇄비",  "강연자료집 제작 및 인쇄", "식", "1", "", "", "기획팀(온나무)",       "내부 처리"),
    ("제작비",  "포스터 초안 제작비",      "식", "1", "", "", "기획·홍보(온나무)",    "초안 제작 내부 수행"),
    ("제작비",  "SNS 초안 제작비",         "식", "1", "", "", "기획·홍보(온나무)",    "초안 제작 내부 수행"),
    ("기획비",  "전체 계획 및 발주 획득 관리", "식", "1", "", "", "기획팀(온나무)",   "총괄 기획·발주 관리"),
    ("보고",    "결과보고서 작성 완성",     "식", "1", "", "", "기획팀(온나무)",       "최종 보고서 완성"),
]

# 흰색(용역사 위탁) — 아래쪽 / 총합 7,000,000 (VAT 포함)
white_rows = [
    ("제작비",  "행사 사진·영상 촬영 및 편집", "식", "1",  "2,500,000", "2,500,000", "행정·운영지원(영상)", "최종 결과물 1EA"),
    ("제작비",  "SNS 제작비",              "식",  "1",  "600,000",   "600,000",   "홍보(용역사)",         "SNS 게시물 제작"),
    ("제작비",  "포스터 수정비",           "식",  "1",  "500,000",   "500,000",   "홍보(용역사)",         "포스터 수정·편집"),
    ("인쇄비",  "강연장 현수막",           "EA",  "6",  "80,000",    "480,000",   "제작·인쇄(용역사)",    ""),
    ("인건비",  "운영스텝비",              "2인", "6",  "130,000",   "1,560,000", "현장운영",              ""),
    ("운영비",  "행사운영 잡비, 다과 및 회의비", "식", "1", "860,000", "860,000",  "운영지원",              ""),
    ("보고",    "결과보고서 재료 정리 및 저장", "식", "1", "500,000", "500,000",  "행정·운영지원(용역사)", "원자료 정리·아카이빙"),
]

table = doc.add_table(rows=1, cols=len(headers))
table.style = 'Table Grid'; table.alignment = WD_TABLE_ALIGNMENT.CENTER

hdr = table.rows[0].cells
for i, h in enumerate(headers):
    set_cell(hdr[i], h, bold=True, white=True); shade(hdr[i], HEADER)

aligns = ['center','center','left','center','center','right','right','center','left']
idx = 1
def add_row(data, gray):
    global idx
    cells = table.add_row().cells
    vals = (str(idx),) + data
    for i, v in enumerate(vals):
        set_cell(cells[i], v, align=aligns[i])
        if gray: shade(cells[i], GRAY)
    idx += 1

for row in gray_rows: add_row(row, True)
for row in white_rows: add_row(row, False)

white_total = sum(int(r[5].replace(",", "")) for r in white_rows)  # 7,000,000

foot = table.add_row().cells
set_cell(foot[0], "흰색 소계", bold=True)
for i in range(1, 6): set_cell(foot[i], "")
set_cell(foot[6], f"{white_total:,}", bold=True, align='right')
set_cell(foot[7], "용역사 위탁분", align='center')
set_cell(foot[8], "부가세 포함", align='left')

note = doc.add_paragraph()
nr = note.add_run("※ 회색 영역(강사료·포스터 초안·SNS 초안·전체 계획 및 발주 관리·결과보고서 작성)은 온나무 내부에서 직접 수행/정산하며, 본 예산표 합계에서는 별도 산정합니다.")
nr.font.size = Pt(8.5); nr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
nr.font.name = '맑은 고딕'; nr._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

doc.save("2026_한복인문학특강_상세예산표_내부용.docx")
print("white_total=", white_total)
