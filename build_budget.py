# -*- coding: utf-8 -*-
"""2026 한복인문학 특강 - 내부용 상세 예산표 (역할 배분) DOCX 생성"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GRAY = "D9D9D9"      # 온나무(우리) 직접 수행 영역
HEADER = "404040"    # 헤더 배경

def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_cell(cell, text, bold=False, align='center', white=False, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER,
                   'left': WD_ALIGN_PARAGRAPH.LEFT,
                   'right': WD_ALIGN_PARAGRAPH.RIGHT}[align]
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.font.name = '맑은 고딕'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    run.bold = bold
    if white:
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc = Document()
for s in doc.sections:
    s.left_margin = Cm(1.8); s.right_margin = Cm(1.8)

# 제목
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("「2026 한복인문학 특강」 상세 예산표 (내부용)")
r.bold = True; r.font.size = Pt(15)
r.font.name = '맑은 고딕'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

sub = doc.add_paragraph()
sr = sub.add_run("※ 회색 음영 = 온나무 직접 수행 영역(내부 처리) / 백색 = 용역사 위탁 범위")
sr.font.size = Pt(9); sr.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
sr.font.name = '맑은 고딕'; sr._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

headers = ["구분", "항목", "세부내용", "단위", "수량", "단가", "합계", "담당(역할)", "비고"]
# rows: (구분,항목,세부내용,단위,수량,단가,합계,담당,비고,gray?)
rows = [
    ("1", "강사료", "인문학특강 강사료", "회", "6", "", "", "기획팀(온나무 직접)", "교통비 등 포함 / 내부 처리", True),
    ("2", "인건비", "사회진행비", "회", "6", "361,000", "2,166,000", "현장운영(로컬 협업 매니저)", "사회자 교통비 포함", False),
    ("3", "제작비", "행사 사진·영상 촬영 및 편집", "식", "1", "2,232,500", "2,232,500", "행정·운영지원(영상제작)", "최종 결과물 1EA", False),
    ("4", "인쇄비", "강연자료집 제작 및 인쇄", "식", "1", "1,425,000", "1,425,000", "제작·인쇄(용역사)", "", False),
    ("5", "제작비", "SNS 제작비", "식", "1", "", "", "홍보(용역사)", "SNS 게시물 제작", False),
    ("6", "제작비", "포스터 수정비", "식", "1", "", "", "홍보(용역사)", "포스터 수정·편집", False),
    ("7", "제작비", "포스터 초안 제작비", "식", "1", "", "", "기획·홍보(온나무 직접)", "초안 제작 내부 처리", True),
    ("8", "인쇄비", "강연장 현수막", "EA", "6", "114,000", "684,000", "제작·인쇄(용역사)", "", False),
    ("9", "인건비", "운영스텝비", "2인", "6", "142,500", "1,710,000", "현장운영", "", False),
    ("10", "운영비", "행사운영 잡비, 다과 및 회의비", "식", "1", "2,850,000", "2,850,000", "운영지원", "", False),
]

table = doc.add_table(rows=1, cols=len(headers))
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

hdr = table.rows[0].cells
for i, h in enumerate(headers):
    set_cell(hdr[i], h, bold=True, white=True)
    shade(hdr[i], HEADER)

for row in rows:
    cells = table.add_row().cells
    gray = row[-1]
    vals = row[:-1]
    aligns = ['center','center','left','center','center','right','right','center','left']
    for i, v in enumerate(vals):
        set_cell(cells[i], v, align=aligns[i])
        if gray:
            shade(cells[i], GRAY)

# 소계 / 총합계 (회색 영역 금액 미산정 상태의 백색=용역사 범위 합계)
total = 2166000 + 2232500 + 1425000 + 684000 + 1710000 + 2850000
foot = table.add_row().cells
set_cell(foot[0], "소 계", bold=True)
for i in range(1, 6):
    set_cell(foot[i], "")
set_cell(foot[6], f"{total:,}", bold=True, align='right')
set_cell(foot[7], "용역사 위탁분(백색)", align='center')
set_cell(foot[8], "회색 영역 별도", align='left')

note = doc.add_paragraph()
nr = note.add_run("※ 회색 영역(강사료·포스터 초안 제작비) 및 SNS 제작비·포스터 수정비 단가는 협의 후 확정 예정입니다.")
nr.font.size = Pt(8.5); nr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
nr.font.name = '맑은 고딕'; nr._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

doc.save("2026_한복인문학특강_상세예산표_내부용.docx")
print("saved")
