from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# 페이지 설정 (A4, 여백 2.5cm)
section = doc.sections[0]
section.page_width    = Cm(21)
section.page_height   = Cm(29.7)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

TITLE_COLOR = RGBColor(0x1F, 0x49, 0x7D)

def add_heading(doc, text, size=13, bold=True, color=None, space_before=8, space_after=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p

def add_body(doc, text, size=12, indent=True, space_before=2, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Pt(12)
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p

# ── 제목 ──────────────────────────────────────────────
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after  = Pt(10)
title_run = title_p.add_run("2차원 바코드의 개념과 대표 유형")
title_run.font.size = Pt(16)
title_run.font.bold = True
title_run.font.color.rgb = TITLE_COLOR

# ── 1. 2차원 바코드란 무엇인가 ────────────────────────
add_heading(doc, "1. 2차원 바코드란 무엇인가?",
            size=13, color=(0x1F, 0x49, 0x7D), space_before=4, space_after=3)

add_body(doc,
    "2차원 바코드(2D Barcode)는 기존의 1차원(선형) 바코드가 가로 방향으로만 데이터를 표현하는 것과 달리, "
    "가로(수평)와 세로(수직) 두 방향 모두를 활용하여 정보를 저장하는 코드 체계이다. "
    "1차원 바코드는 수십 자(字) 수준의 숫자·문자만 담을 수 있는 반면, 2차원 바코드는 수백에서 수천 자의 "
    "텍스트, 숫자, 이진 데이터, URL 등 다양한 형태의 정보를 하나의 심벌 안에 압축할 수 있다.")

add_body(doc,
    "2차원 바코드의 핵심 장점은 크게 세 가지이다. 첫째, 높은 정보 밀도로 좁은 면적에 많은 데이터를 저장한다. "
    "둘째, 오류 정정(Error Correction) 기능을 내장하여 코드의 일부가 오염·손상되어도 원본 데이터를 복원할 수 있다. "
    "셋째, 스마트폰 카메라 등 일반 이미지 센서로 판독이 가능하여 별도의 전용 장비 없이도 활용할 수 있다. "
    "이러한 특성 덕분에 유통, 물류, 마케팅, 의료, 공공 서비스 등 광범위한 분야에서 활용되고 있다.")

# ── 2. 대표적인 세 가지 유형 ───────────────────────────
add_heading(doc, "2. 대표적인 2차원 바코드 유형",
            size=13, color=(0x1F, 0x49, 0x7D), space_before=8, space_after=3)

# ① QR 코드
add_heading(doc, "① QR 코드(Quick Response Code)",
            size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "QR 코드는 1994년 일본의 덴소 웨이브(Denso Wave)가 자동차 부품 관리를 위해 개발한 정사각형 격자 무늬 형태의 2차원 바코드이다. "
    "최대 7,089자(숫자 기준)의 데이터를 저장할 수 있으며, 4단계의 오류 정정 레벨(L·M·Q·H)을 제공한다. "
    "세 모서리에 위치한 '파인더 패턴(정렬 마커)'을 통해 어떤 각도에서도 빠르게 인식된다. "
    "현재 스마트폰 보급과 함께 모바일 결제, 식당 메뉴 안내, 전자 탑승권, 개인 명함 등 일상생활 전반에서 "
    "가장 널리 쓰이는 2차원 바코드로 자리 잡았다.")

# ② PDF417
add_heading(doc, "② PDF417(Portable Data File 417)",
            size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "PDF417은 1991년 Symbol Technologies가 개발한 적층형(스택형) 2차원 바코드로, 여러 개의 1차원 바코드 행을 수직으로 쌓은 구조를 가진다. "
    "'417'이라는 명칭은 각 코드워드가 4개의 바(bar)와 공간(space)으로 구성되고 각각 17개 모듈 너비를 가지는 데서 유래했다. "
    "최대 1,108바이트의 이진 데이터 혹은 1,850개의 텍스트 문자를 저장할 수 있으며, "
    "한국 운전면허증, 미국 주(州) 운전면허증, 항공 탑승권(BCBP 규격), 물류 라벨 등 "
    "공공 신분증 및 물류 문서에 표준적으로 활용된다.")

# ③ Data Matrix
add_heading(doc, "③ 데이터 매트릭스(Data Matrix)",
            size=12, bold=True, space_before=4, space_after=2)
add_body(doc,
    "데이터 매트릭스는 1987년 International Data Matrix社(현 Microscan)가 개발한 정사각형 또는 직사각형 격자 형태의 2차원 바코드이다. "
    "코드 경계를 나타내는 'L자형 실선(Finder Pattern)'과 데이터 영역을 구분하는 점선으로 이루어진다. "
    "최소 2mm² 이하의 극소형 크기에도 인쇄·각인이 가능하여, "
    "반도체 칩, PCB 기판, 의약품 낱알 포장, 외과 수술 도구 등 소형 부품 추적에 특화되어 있다. "
    "GS1 및 ISO/IEC 16022 국제 표준으로 채택되어 있으며, 미국 FDA는 의약품 단위 포장에 "
    "데이터 매트릭스 코드 부착을 의무화(UDI 규정)하고 있다.")

# ── 3. 결론 ───────────────────────────────────────────
add_heading(doc, "3. 결론",
            size=13, color=(0x1F, 0x49, 0x7D), space_before=8, space_after=3)
add_body(doc,
    "2차원 바코드는 대용량 데이터 저장, 오류 정정, 범용 판독이라는 장점을 바탕으로 다양한 산업 분야에 깊숙이 침투해 있다. "
    "QR 코드는 일상적인 모바일 서비스에, PDF417은 신분증·항공권 등 공공 문서에, "
    "데이터 매트릭스는 초소형 부품·의약품 추적에 각각 최적화된 형태로 발전해 왔으며, "
    "디지털 전환(DX) 흐름 속에서 그 활용 범위는 더욱 확대될 것으로 예상된다.")

# ── 참고 문헌 ─────────────────────────────────────────
from docx.oxml.ns import qn as _qn
from docx.oxml import OxmlElement as _OxmlElement

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
    "Denso Wave Incorporated. (2015). QR Code Standardization. https://www.qrcode.com/en/about/standards.html",
    "ISO/IEC 18004:2015. Information technology — Automatic identification and data capture techniques — QR Code bar code symbology specification. ISO.",
    "Wang, Y. J., & Liu, W. (2011). Comparison of QR code and data matrix for mobile phone barcode systems. 2011 International Conference on Computer and Management (CAMAN), 1-5. https://doi.org/10.1109/CAMAN.2011.5778834",
    "Symbol Technologies. (1992). PDF417: A New Bar Code Symbology for Industry. Symbology Specification.",
    "ISO/IEC 15438:2015. Information technology — Automatic identification and data capture techniques — PDF417 bar code symbology specification. ISO.",
    "ISO/IEC 16022:2006. Information technology — Automatic identification and data capture techniques — Data Matrix bar code symbology specification. ISO.",
    "GS1. (2023). GS1 DataMatrix Guideline: Overview and technical introduction to the use of GS1 DataMatrix. GS1 Global Office.",
    "U.S. Food and Drug Administration. (2023). Unique Device Identification System (UDI System). https://www.fda.gov/medical-devices/device-advice-comprehensive-regulatory-assistance/unique-device-identification-system-udi-system",
    "한국정보통신기술협회(TTA). (2020). 2차원 바코드 기술 및 응용 표준화 동향. TTA Journal, 189, 55-62.",
]

for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Pt(24)
    p.paragraph_format.first_line_indent = Pt(-24)
    run = p.add_run(f"[{i}] {ref}")
    run.font.size = Pt(10)

doc.save("/workspace/2차원바코드_개념과_유형.docx")
print("문서 생성 완료: 2차원바코드_개념과_유형.docx")
