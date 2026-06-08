# python-docx Formatting Patterns for Itinerary Generator

A quick reference of code patterns used in the itinerary generator.

## Table Setup

```python
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT

t = doc.add_table(rows=1, cols=3)
t.style = 'Table Grid'  # REQUIRED for borders
t.alignment = WD_TABLE_ALIGNMENT.CENTER
```

## Cell Shading (Background Color)

```python
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def set_cell_shading(cell, color):
    # IMPORTANT: remove existing shd first to avoid duplicates
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}shd' % NS):
        tcPr.remove(existing)
    shading_elm = parse_xml(
        '<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color)
    )
    tcPr.append(shading_elm)
```

## Cell Vertical Alignment

```python
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def set_cell_va(cell, align='center'):
    """Set cell vertical alignment.
    IMPORTANT: pass 'center' as string, NOT WD_ALIGN_VERTICAL.CENTER
    (the enum serializes as 'CENTER (1)' which Word rejects)
    """
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}vAlign' % NS):
        tcPr.remove(existing)
    tcPr.append(
        parse_xml('<w:vAlign {} w:val="{}"/>'.format(nsdecls("w"), align))
    )
```

## Adding Formatted Text to a Cell

```python
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_formatted_paragraph(cell, text, font_name='Arial', size=Pt(10), bold=False,
                            color=RGBColor(0x33, 0x33, 0x33),
                            alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            space_before=0, space_after=0, line_spacing=1.3):
    p = cell.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    return p
```

## Merging Cells

```python
# Merge all 3 cells in a row (for free-day rows, etc.)
row = table.add_row()
row.cells[0].merge(row.cells[2])
```

## Setting Column Widths

```python
cells[0].width = Cm(2.2)
cells[1].width = Cm(11.5)
cells[2].width = Cm(3.3)
```

## Day Header with Badge

```python
# Create a 1-col full-width table above the detail table
t = doc.add_table(rows=1, cols=1)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.style = 'Table Grid'
cell = t.cell(0, 0)
set_cell_shading(cell, 'EDF2F9')

# Badge
rb = p.add_run('  D-01  ')
rb.font.size = Pt(12)
rb.font.bold = True
rb.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
rb._element.get_or_add_rPr().append(
    parse_xml('<w:shd {} w:fill="1A3C6E" w:val="clear"/>'.format(nsdecls("w"))))

# Route text
rr = p.add_run('  Singapore -> Shanghai')
rr.font.size = Pt(12)
rr.font.bold = True
rr.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)
```

## Document Setup

```python
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(1.8)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
```

## Encoding Notes

### Unicode escapes in Python source

When writing Chinese characters as `\uXXXX` escapes, each Chinese character is one escape.

### Emoji handling

- Surrogate pairs will crash with `surrogates not allowed`
- Use full Unicode codepoint or plain text labels

### Reading .doc files

Always use `antiword`:
```bash
antiword -m UTF-8 "input.doc"
```
python-docx and python libraries cannot read old .doc format (only .docx).
