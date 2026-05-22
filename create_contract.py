from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

doc = Document()

# Set default font
style = doc.styles['Normal']
style.font.name = 'Malgun Gothic'
style.font.size = Pt(11)

# Title
title = doc.add_heading('상호협약서', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

subtitle = doc.add_paragraph('건물 활용 및 수익 창출에 관한 협력 협약')
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# Preamble
p = doc.add_paragraph()
p.add_run('본 협약서는 정식 계약 체결을 위한 중간 단계의 협의 문서이며, 정식 계약서 체결 전까지 법적 구속력을 갖지 아니한다. ')
p.add_run('★단, 제7조(비밀유지), 제8조(분쟁해결), 제9조(손해배상)는 본 협약 체결 시부터 법적 구속력을 가진다.★').bold = True

doc.add_paragraph()

p = doc.add_paragraph('최동녘(이하 "갑")과 이석민(이하 "을")은 대상 건물의 활용 및 향후 매각을 통한 수익 창출을 위해 상호 협력하기로 하고, 다음과 같이 협약한다.')

doc.add_paragraph()

# Article 1
doc.add_heading('제1조 (협약의 목적)', level=2)
doc.add_paragraph('본 협약은 갑과 을이 대상 건물의 사용·운영 및 향후 매각에 관한 협력 방향을 합의하는 것을 목적으로 한다.')

# Article 2
doc.add_heading('제2조 (협약의 대상)', level=2)
doc.add_paragraph('① 본 협약의 대상 건물은 양 당사자가 상호 합의한 건물로 하며, 구체적인 내용(위치, 면적 등)은 정식 계약 체결 시 특정한다.')
p = doc.add_paragraph()
p.add_run('★② 양 당사자는 본 협약 체결일로부터 30일 이내에 대상 건물을 특정하기 위한 협의를 개시하여야 하며, 협의 개시일로부터 60일 이내에 대상 건물에 관한 합의서를 별도 작성한다.★').bold = True

# Article 3
doc.add_heading('제3조 (건물 사용 방식)', level=2)
doc.add_paragraph('① 을은 대상 건물 소유주와의 협의를 조력하며, 대여 방식은 아래 각 호 중 하나로 정한다.')
doc.add_paragraph('    가. 소유주로부터 갑이 직접 계약하는 방식 (을이 협의 조력)')
doc.add_paragraph('    나. 을(또는 을이 대표로 있는 법인)이 소유주로부터 위탁받아 갑과 재계약하는 방식')
doc.add_paragraph('② 구체적인 방식은 양 당사자의 협의로 결정한다.')
p = doc.add_paragraph()
p.add_run('★③ 을의 협의 조력에 따른 비용(교통비, 통신비 등 실비)은 갑이 부담하며, 그 범위는 월 50만원을 한도로 한다. 한도 초과 비용 발생 시 사전 서면 합의를 요한다.★').bold = True

# Article 4
doc.add_heading('제4조 (수익 창출 목표)', level=2)
doc.add_paragraph('① 양 당사자는 건물 사용 개시 후 약 3년~4년 시점을 기준으로 대상 건물의 매각을 통한 수익 창출을 목표로 한다.')
p = doc.add_paragraph()
p.add_run('★② 매각 시 수익 배분은 갑과 을이 각각 [  ]%와 [  ]%로 하며, 구체적인 비율은 정식 계약 체결 시 확정한다.★').bold = True
p = doc.add_paragraph()
p.add_run("★③ '수익'의 정의는 매각대금에서 취득원가, 운영비용, 세금 및 기타 필수 경비를 공제한 순수익을 의미한다.★").bold = True
doc.add_paragraph('④ 매각 시기·조건·수익 배분 등의 세부 사항은 정식 계약 시 별도 협의한다.')

# Article 5
doc.add_heading('제5조 (법적 효력 및 유효 기간)', level=2)
p = doc.add_paragraph()
p.add_run('① 본 협약서는 정식 계약 체결 전까지 법적 구속력을 갖지 아니하며, 양 당사자의 협력 의사를 확인하는 문서로서의 의미를 가진다. ')
p.add_run('★단, 제7조 내지 제10조는 예외로 한다.★').bold = True
doc.add_paragraph('② 본 협약의 유효 기간은 협약 체결일로부터 정식 계약이 체결되는 날까지로 한다.')
p = doc.add_paragraph()
p.add_run('★③ 본 협약 체결일로부터 6개월 이내에 정식 계약이 체결되지 아니할 경우, 양 당사자는 협약의 연장 또는 종료 여부를 협의하여야 한다.★').bold = True

# Article 6
doc.add_heading('제6조 (추가 협약)', level=2)
doc.add_paragraph('양 당사자는 필요에 따라 본 협약을 보완하거나 구체화하기 위한 추가 협약을 별도로 체결할 수 있으며, 추가 협약은 양 당사자가 서명함으로써 본 협약의 일부를 이룬다.')

# New Articles Header
doc.add_paragraph()
p = doc.add_paragraph()
p.add_run('★ 추가 조항 ★').bold = True
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Article 7
doc.add_heading('★제7조 (비밀유지)★', level=2)
doc.add_paragraph('① 양 당사자는 본 협약의 체결 및 이행 과정에서 알게 된 상대방의 비밀정보(사업계획, 재무정보, 협상내용 등)를 제3자에게 누설하거나 본 협약의 목적 외의 용도로 사용하여서는 아니 된다.')
doc.add_paragraph('② 본 조의 비밀유지 의무는 본 협약의 종료 후 2년간 존속한다.')
doc.add_paragraph('③ 다만, 다음 각 호의 경우에는 비밀유지 의무가 적용되지 아니한다.')
doc.add_paragraph('    가. 공지의 사실이 된 정보')
doc.add_paragraph('    나. 법령에 따라 공개가 요구되는 정보')
doc.add_paragraph('    다. 상대방의 사전 서면 동의를 받은 경우')

# Article 8
doc.add_heading('★제8조 (분쟁해결)★', level=2)
doc.add_paragraph('① 본 협약과 관련하여 분쟁이 발생한 경우, 양 당사자는 먼저 상호 협의를 통해 원만히 해결하도록 노력한다.')
doc.add_paragraph('② 제1항의 협의가 30일 이내에 이루어지지 아니할 경우, 대한민국 법률을 준거법으로 하여 서울중앙지방법원을 제1심 전속관할 법원으로 하여 해결한다.')

# Article 9
doc.add_heading('★제9조 (손해배상)★', level=2)
doc.add_paragraph('① 양 당사자는 고의 또는 중대한 과실로 본 협약을 위반하여 상대방에게 손해를 끼친 경우, 이를 배상할 책임을 진다.')
doc.add_paragraph('② 손해배상의 범위는 통상손해에 한하며, 특별손해는 예견가능성이 있는 경우에 한하여 배상한다.')

# Article 10
doc.add_heading('★제10조 (협약의 종료 및 해지)★', level=2)
doc.add_paragraph('① 양 당사자는 상대방에게 30일 전 서면 통지로 본 협약을 해지할 수 있다.')
doc.add_paragraph('② 다음 각 호의 경우, 일방 당사자는 즉시 본 협약을 해지할 수 있다.')
doc.add_paragraph('    가. 상대방이 본 협약상 의무를 중대하게 위반하고, 시정 요구 후 14일 이내에 시정하지 아니한 경우')
doc.add_paragraph('    나. 상대방이 파산, 회생절차 개시 등 지급불능 상태에 빠진 경우')
doc.add_paragraph('③ 협약 종료 시 양 당사자는 상대방으로부터 제공받은 자료 및 정보를 반환하거나 폐기하여야 한다.')

# Article 11
doc.add_heading('★제11조 (양도 금지)★', level=2)
doc.add_paragraph('양 당사자는 상대방의 사전 서면 동의 없이 본 협약상의 권리 및 의무를 제3자에게 양도하거나 담보로 제공할 수 없다.')

# Article 12
doc.add_heading('★제12조 (통지)★', level=2)
doc.add_paragraph('① 본 협약과 관련한 모든 통지는 서면(이메일 포함)으로 하여야 하며, 상대방이 지정한 주소로 발송한다.')
doc.add_paragraph('② 통지는 상대방에게 도달한 때에 효력이 발생한다.')

doc.add_paragraph()

# Governing Law
p = doc.add_paragraph()
p.add_run('준거법 / 관할: ').bold = True
p.add_run('대한민국')

doc.add_paragraph()
doc.add_paragraph('본 협약의 성립을 증명하기 위하여 본 협약서 2부를 작성하고, 양 당사자가 서명 또는 기명날인 후 각 1부씩 보관한다.')

doc.add_paragraph()
p = doc.add_paragraph()
p.add_run('협약 체결일: ').bold = True
p.add_run('20___년 ___월 ___일')

doc.add_paragraph()

# Signature table
table = doc.add_table(rows=4, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER

table.cell(0, 0).text = '갑 (甲)'
table.cell(0, 1).text = '을 (乙)'
table.cell(1, 0).text = '성명: 최동녘'
table.cell(1, 1).text = '성명: 이석민'
table.cell(2, 0).text = '주소:'
table.cell(2, 1).text = '주소:'
table.cell(3, 0).text = '서명:'
table.cell(3, 1).text = '서명:'

# Make headers bold
for cell in table.rows[0].cells:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.bold = True

doc.add_paragraph()
doc.add_paragraph()

# Summary section
doc.add_heading('수정 내역 요약', level=1)

summary_table = doc.add_table(rows=6, cols=4)
summary_table.style = 'Table Grid'

headers = ['조항', '수정 전', '수정 후', '수정 이유']
for i, header in enumerate(headers):
    summary_table.cell(0, i).text = header
    for paragraph in summary_table.cell(0, i).paragraphs:
        for run in paragraph.runs:
            run.bold = True

data = [
    ['전문', '전체 비구속적', '비밀유지, 분쟁해결, 손해배상 조항은 즉시 구속력 부여', '위험도 상: MOU 단계에서도 핵심 보호 조항은 효력 필요'],
    ['제2조', '대상 건물 특정 시기 불명확', '30일 내 협의 개시, 60일 내 합의서 작성', '위험도 상: 계약 대상 불특정으로 인한 분쟁 방지'],
    ['제3조', '비용 부담 주체 불명확', '실비 부담 주체(갑) 및 한도(월 50만원) 명시', '위험도 중: 비용 분쟁 예방'],
    ['제4조', '수익의 정의 및 배분 기준 없음', '수익 정의 명확화, 배분 비율 기재란 추가', '위험도 상: 핵심 이익 사항 불명확'],
    ['제5조', '유효기간 종료 시점만 규정', '6개월 내 미체결 시 협의 의무 추가', '위험도 중: 장기 표류 방지'],
]

for row_idx, row_data in enumerate(data, start=1):
    for col_idx, cell_data in enumerate(row_data):
        summary_table.cell(row_idx, col_idx).text = cell_data

doc.add_paragraph()

# Added clauses section
doc.add_heading('추가된 조항 목록', level=1)

added_table = doc.add_table(rows=7, cols=3)
added_table.style = 'Table Grid'

headers2 = ['추가 조항', '내용', '추가 이유']
for i, header in enumerate(headers2):
    added_table.cell(0, i).text = header
    for paragraph in added_table.cell(0, i).paragraphs:
        for run in paragraph.runs:
            run.bold = True

added_data = [
    ['★제7조 (비밀유지)★', '비밀정보 보호, 2년간 존속, 예외 사항', '협상 과정 정보 유출 방지 필수'],
    ['★제8조 (분쟁해결)★', '협의 우선, 서울중앙지방법원 전속관할', '분쟁 발생 시 해결 절차 부재 보완'],
    ['★제9조 (손해배상)★', '고의·중과실 위반 시 배상 책임', '계약 위반에 대한 구제 수단 확보'],
    ['★제10조 (협약의 종료 및 해지)★', '30일 사전통지, 즉시해지 사유, 자료 반환', '출구 전략 및 종료 절차 명확화'],
    ['★제11조 (양도 금지)★', '사전 서면 동의 없이 양도 불가', '당사자 변경으로 인한 리스크 방지'],
    ['★제12조 (통지)★', '서면 통지 원칙, 도달주의', '의사소통 분쟁 예방'],
]

for row_idx, row_data in enumerate(added_data, start=1):
    for col_idx, cell_data in enumerate(row_data):
        added_table.cell(row_idx, col_idx).text = cell_data

doc.add_paragraph()

# Notes
doc.add_heading('참고사항', level=2)
doc.add_paragraph('정식 계약 체결 시 아래 사항을 반드시 확정하시기 바랍니다:')
doc.add_paragraph('1. 제2조 ②항의 대상 건물 특정 (주소, 면적, 등기사항)')
doc.add_paragraph('2. 제4조 ②항의 수익 배분 비율 (갑 ___%, 을 ___%로 공란 기재)')
doc.add_paragraph('3. 제12조 ②항의 각 당사자 통지 수령 주소 및 이메일')

# Save
doc.save('/home/user/project-1/상호협약서_수정본.docx')
print('Document created successfully!')
