from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 페이지 여백 설정 (A4 기준 상하좌우 2.5cm)
section = doc.sections[0]
section.page_width  = Cm(21)
section.page_height = Cm(29.7)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

def add_heading(doc, text, size=13, bold=True, color=None, space_before=10, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p

def add_body(doc, text, size=12, indent=False, space_before=2, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(12)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p

# ── 제목 ──────────────────────────────────────────────
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after  = Pt(10)
title_run = title_p.add_run("RFID의 개념과 활용 분야")
title_run.font.size = Pt(16)
title_run.font.bold = True
title_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

# ── 1. RFID란 무엇인가 ────────────────────────────────
add_heading(doc, "1. RFID란 무엇인가?", size=13, color=(0x1F, 0x49, 0x7D), space_before=6, space_after=3)

add_body(doc,
    "RFID(Radio Frequency Identification)란 전파(라디오 주파수)를 이용하여 태그(Tag)에 저장된 정보를 비접촉 방식으로 자동 인식하는 기술이다. "
    "바코드와 달리 직접적인 시야 확보 없이도 데이터를 읽을 수 있으며, 복수의 태그를 동시에 인식할 수 있다는 것이 핵심 특징이다.",
    indent=True)

add_body(doc,
    "RFID 시스템은 크게 세 가지 요소로 구성된다. 첫째, 태그(Tag)는 고유 식별 코드와 데이터를 저장하는 초소형 칩과 안테나로 이루어진 장치이다. "
    "태그는 전원 공급 방식에 따라 자체 배터리를 갖는 능동형(Active)과 리더기의 전파 에너지를 수신하여 구동되는 수동형(Passive)으로 나뉜다. "
    "둘째, 리더기(Reader)는 안테나를 통해 전자기파를 방출하여 태그를 활성화하고 데이터를 수신하는 장치이다. "
    "셋째, 미들웨어·서버 시스템은 리더기로부터 전달된 데이터를 처리·저장·분석한다.",
    indent=True)

add_body(doc,
    "RFID는 사용 주파수 대역에 따라 저주파(LF, 125~134 kHz), 고주파(HF, 13.56 MHz), 극초단파(UHF, 860~960 MHz), 마이크로파(2.45 GHz) 등으로 구분되며, "
    "주파수가 높을수록 인식 거리가 길어지나 투과성은 낮아지는 특성이 있다.",
    indent=True)

# ── 2. RFID 활용 분야 ─────────────────────────────────
add_heading(doc, "2. RFID의 주요 활용 분야", size=13, color=(0x1F, 0x49, 0x7D), space_before=8, space_after=3)

# 2-1
add_heading(doc, "① 물류·공급망 관리(SCM)", size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "RFID는 물류 분야에서 가장 광범위하게 활용된다. 제품 팔레트, 박스, 개별 상품에 태그를 부착하면 창고 입·출고부터 유통 센터 경유, 매장 진열까지의 "
    "전 과정을 실시간으로 추적할 수 있다. 예컨대 글로벌 유통 기업 월마트(Walmart)는 2000년대 초부터 협력사에 RFID 태그 부착을 의무화하여 "
    "재고 오류를 획기적으로 줄이고 품절 발생률을 약 30% 감소시켰다. 또한 스마트 게이트를 통과하는 것만으로 수백 개의 상품 정보를 "
    "일괄 스캔할 수 있어 작업 속도와 정확성이 대폭 향상된다.",
    indent=True)

# 2-2
add_heading(doc, "② 교통·출입 통제(스마트 카드·하이패스)", size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "교통 분야에서 RFID는 고속도로 자동 요금 징수 시스템(ETC)에 핵심 기술로 사용된다. 국내의 하이패스(Hi-pass) 단말기는 차량 내부에 부착된 "
    "RFID 태그와 톨게이트 리더기가 UHF 대역으로 통신하여 차량이 주행 중에도 통행료를 자동 정산한다. "
    "이와 유사하게 지하철·버스 교통 카드(HF 13.56 MHz, ISO 14443 규격)도 RFID 원리를 활용하며, "
    "아파트·사무실의 출입 통제 시스템에서도 직원증·카드키가 리더기에 근접하면 인증이 이루어지는 방식으로 보안 관리가 이루어진다.",
    indent=True)

# 2-3
add_heading(doc, "③ 의료·환자 안전 관리", size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "의료 현장에서 RFID는 환자 안전과 의약품 관리에 효과적으로 활용된다. 환자 손목 밴드에 RFID 태그를 부착하면 간호사가 처치 전 "
    "리더기로 태그를 스캔하여 환자 정보·투약 이력을 즉시 확인할 수 있어 오처치·투약 오류를 방지한다. "
    "또한 의료 장비와 수술 도구에 태그를 달아 실시간 위치를 파악함으로써 분실을 예방하고, "
    "혈액 팩·의약품에 태그를 적용하면 유통 경로와 유효 기간을 자동 관리하여 의약품 위변조 및 오남용을 차단할 수 있다.",
    indent=True)

# ── 결론 ──────────────────────────────────────────────
add_heading(doc, "3. 결론", size=13, color=(0x1F, 0x49, 0x7D), space_before=8, space_after=3)
add_body(doc,
    "RFID는 비접촉·자동 인식이라는 핵심 강점을 바탕으로 물류, 교통, 의료 등 다양한 산업 분야에서 업무 효율과 데이터 정확성을 높이는 "
    "핵심 기술로 자리매김하고 있다. 사물인터넷(IoT)과의 융합이 가속화되면서 RFID의 활용 범위는 더욱 확대될 것으로 전망된다.",
    indent=True)

# ── 참고 문헌 ─────────────────────────────────────────
from docx.oxml.ns import qn as _qn
from docx.oxml import OxmlElement as _OxmlElement

# 페이지 나누기
page_break_p = doc.add_paragraph()
page_break_p.paragraph_format.space_before = Pt(0)
page_break_p.paragraph_format.space_after  = Pt(0)
run_br = page_break_p.add_run()
br = _OxmlElement('w:br')
br.set(_qn('w:type'), 'page')
run_br._r.append(br)

ref_title = doc.add_paragraph()
ref_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
ref_title.paragraph_format.space_before = Pt(0)
ref_title.paragraph_format.space_after  = Pt(8)
rt_run = ref_title.add_run("참고 문헌")
rt_run.font.size = Pt(14)
rt_run.font.bold = True
rt_run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

refs = [
    "Finkenzeller, K. (2010). RFID Handbook: Fundamentals and Applications in Contactless Smart Cards, Radio Frequency Identification and Near-Field Communication (3rd ed.). Wiley.",
    "Want, R. (2006). An introduction to RFID technology. IEEE Pervasive Computing, 5(1), 25-33. https://doi.org/10.1109/MPRV.2006.11",
    "Bhatt, H., & Bhatt, G. (2005). RFID technology and its applications in information systems and supply chain management. Information Systems Management, 22(4), 51-65.",
    "ISO/IEC 18000-6:2013. Information technology — Radio frequency identification for item management — Part 6: Parameters for air interface communications at 860 MHz to 960 MHz General. ISO.",
    "EPCglobal. (2008). EPC Radio-Frequency Identity Protocols Class-1 Generation-2 UHF RFID Protocol for Communications at 860 MHz–960 MHz (Version 1.2.0). GS1.",
    "한국도로공사. (2024). 하이패스 시스템 소개. https://www.ex.co.kr",
    "식품의약품안전처. (2022). 의료기기 고유식별코드(UDI) 제도 안내. 식품의약품안전처.",
    "한국정보통신기술협회(TTA). (2021). RFID 기술 동향 및 표준화 현황. TTA Journal, 196, 42-48.",
]

for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Pt(24)
    p.paragraph_format.first_line_indent = Pt(-24)
    run = p.add_run(f"[{i}] {ref}")
    run.font.size = Pt(10)

doc.save("/workspace/RFID_개념과_활용분야.docx")
print("문서 생성 완료: RFID_개념과_활용분야.docx")
