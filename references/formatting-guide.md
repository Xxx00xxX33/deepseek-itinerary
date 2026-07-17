# Voyage-style formatting guide (itinerary DOCX)

Paragraph-flow layout matching 东北草原 / Voyage client docs.  
**Default is not a multi-column blue table.**

## Colors

```python
from docx.shared import RGBColor

C_TITLE  = RGBColor(0x1A, 0x6B, 0x5E)  # title, 宿, labels
C_SUB    = RGBColor(0x55, 0x55, 0x55)  # city line
C_ORANGE = RGBColor(0xC4, 0x59, 0x11)  # 全程不进店
C_HL     = RGBColor(0x41, 0x9D, 0x99)  # optional highlight
C_FOOD   = RGBColor(0x2E, 0x8B, 0x7A)  # food / gift accent
C_DAY_BG = '2E8B7A'                    # day bar fill (hex string)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_BODY   = RGBColor(0x00, 0x00, 0x00)
C_HOTEL  = RGBColor(0x1A, 0x6B, 0x5E)

FONT_CN  = 'SimSun'
FONT_DAY = 'Arial Unicode MS'
```

## Template shell (required for Word reliability)

```python
from docx import Document
from docx.oxml.ns import qn

TEMPLATE = r'C:\Users\DELL\Downloads\12D11N_东北草原深度游_2026年6月.docx'

def clear_body(doc: Document):
    body = doc.element.body
    for child in list(body):
        if child.tag != qn('w:sectPr'):
            body.remove(child)

doc = Document(TEMPLATE)
clear_body(doc)
# ... add paragraphs ...
doc.save(output_path)
```

Keeps company header/logo, section margins, and theme. Prefer this over `Document()` empty packages when previous pure-python-docx files failed to open.

If the user says「按某某行程排版」, use **that** DOCX as `TEMPLATE` instead.

## Run font (Chinese-safe; never bold=False)

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

def set_run_font(run, name=FONT_CN, size_pt=11, bold=None, color=None):
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:cs'), name)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    # Only set bold when True — never bold=False (Word rejects w:b w:val="0")
    if bold is True:
        run.font.bold = True
        if rPr.find(qn('w:bCs')) is None:
            rPr.append(OxmlElement('w:bCs'))
    if color is not None:
        run.font.color.rgb = color
```

## Paragraph shading (day bars)

```python
def set_paragraph_shading(paragraph, fill_hex: str):
    pPr = paragraph._p.get_or_add_pPr()
    for old in pPr.findall(qn('w:shd')):
        pPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    pPr.append(shd)
```

Day header pattern (no giant space padding):

```python
header = f'{day["day"]}  {day["route"]}    （{day["meals"]}）'
add_para(doc, header, size=11, bold=True, color=C_WHITE, font=FONT_DAY,
         before=10, after=4, left=2, shade=C_DAY_BG, line=1.1)
```

Use **2–4 spaces** as separators only.

## Generic paragraph helper

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH

def set_spacing(paragraph, before=None, after=None, left=None, line=None):
    pf = paragraph.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if left is not None:
        pf.left_indent = Pt(left)
    if line is not None:
        pf.line_spacing = line

def add_para(doc, text, *, size=11, bold=False, color=C_BODY, font=FONT_CN,
             align=None, before=0, after=4, left=None, justify=False,
             shade=None, line=1.15):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align is not None:
        p.alignment = align
    set_spacing(p, before=before, after=after, left=left, line=line)
    if shade:
        set_paragraph_shading(p, shade)
    if text:
        run = p.add_run(text)
        set_run_font(run, name=font, size_pt=size, bold=bold if bold else None, color=color)
    return p
```

When `bold` is False, pass `bold=None` into `set_run_font` so no `w:b` is emitted.

## Sight line and label+value

```python
def add_sight_para(doc, name: str, desc: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_spacing(p, before=2, after=3, line=1.2)
    r1 = p.add_run(f'【{name}】')
    set_run_font(r1, size_pt=10.5, bold=True, color=C_TITLE)
    r2 = p.add_run(desc)
    set_run_font(r2, size_pt=10.5, color=C_BODY)
    return p

def add_label_value(doc, label: str, value: str, *, before=2, after=2):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.15)
    r1 = p.add_run(label)
    set_run_font(r1, size_pt=10.5, bold=True, color=C_TITLE)
    r2 = p.add_run(value)
    set_run_font(r2, size_pt=10.5, color=C_BODY)
    return p
```

## Document structure order

1. Title (center, ~18pt, `C_TITLE`)
2. City line (center, `C_SUB`)
3. Shopping / no-shop note (center, `C_ORANGE`)
4. `独家安排：` / `各地美食：` / `赠送：` label blocks
5. For each day: day bar → summary → `【景点】` lines → `宿：…`
6. `报价与说明` + pricing / 购物 / 小费 / 备注 / 酒店 / 膳食
7. `联系方式`

## Typography / readability

| Bad | Good |
|-----|------|
| `· 集庆楼 / 善庆楼` | `·集庆楼、善庆楼、绳庆楼` |
| Day header with 40 spaces | `Day 2  厦门→永定（3H）    （早：酒店｜午：…）` |
| Duplicate full day text under every sight | Short scan lines under `【景点】` |
| `宿：酒店 — 五星` if quote uses parentheses | `宿：酒店（五星）` |

Prefer:

- `→` for routes
- `｜` between meal slots
- Chinese顿号 `、` in lists
- Line spacing ~1.15–1.25

## Extracting source .doc

`python-docx` cannot read legacy `.doc`.

```bash
# Prefer if available
antiword -m UTF-8 "quote.doc"

# WPS / Windows fallback
python scripts/extract_doc.py "quote.doc"
```

## Legacy table patterns (do not use as default)

Cell shading, 3-column day tables, and blue badges are **legacy**. Only use if the user explicitly requests the old table style. If you must shade a cell:

```python
# Remove existing shd first; pass align as string 'center', not enum
```

Never emit `bold=False` or `WD_ALIGN_VERTICAL.CENTER` enum string into OOXML.

## Encoding (Windows)

```powershell
$env:PYTHONIOENCODING='utf-8'
python scripts/generate.py
```

## Output paths

- Primary: same folder as source `.doc`, Chinese product name
- Also copy to short ASCII name for tools that choke on CJK paths
