#!/usr/bin/env python3
"""
Template: Generate travel itinerary DOCX from structured data.
Customize for each itinerary: edit DAYS list and run.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

# ── CONFIG ──────────────────────────────────────────────
FONT = 'Arial'       # Use 'Arial' for English, 'Microsoft YaHei' (微软雅黑) for Chinese
OUTPUT_FILE = 'itinerary.docx'
TITLE = 'Your Trip Title'
SUBTITLE = 'City1 \u00b7 City2 \u00b7 City3'
DAYS = []            # Populate with Day objects (see below)

# ── COLOR PALETTE ───────────────────────────────────────
DEEP_BLUE   = RGBColor(0x1A, 0x3C, 0x6E)
DARK_GREEN  = RGBColor(0x2E, 0x6B, 0x4F)
WARM_GOLD   = RGBColor(0xB8, 0x86, 0x2C)
DARK_GRAY   = RGBColor(0x33, 0x33, 0x33)
MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)
LIGHT_GRAY  = RGBColor(0x99, 0x99, 0x99)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DEEP_BLUE_HEX = '1A3C6E'
BG_LIGHT    = 'EDF2F9'
BG_WARM     = 'FEF6E9'
BG_GREEN    = 'E8F5ED'
BG_RED      = 'FFF1F0'

# ── HELPERS ─────────────────────────────────────────────

def set_shading(cell, color):
    cell._tc.get_or_add_tcPr().append(
        parse_xml('<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color)))

def mkp(cell, text, sz=10, bold=False, color=DARK_GRAY, align=WD_ALIGN_PARAGRAPH.LEFT,
        sb=0, sa=0, ls=1.3):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = sb
    p.paragraph_format.space_after = sa
    p.paragraph_format.line_spacing = ls
    r = p.add_run(text)
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.color.rgb = color
    return p

# ═══════════════════════════════════════════════════════
# DATA MODEL - populate DAYS list below
# ═══════════════════════════════════════════════════════

class Sight:
    def __init__(self, time, name, description, meal=None, note=None):
        self.time = time
        self.name = name
        self.description = description  # 100-150 chars, factual, for foreign readers
        self.meal = meal
        self.note = note

class Day:
    def __init__(self, num, route, sights, hotel, meals_summary=None):
        self.num = num
        self.route = route
        self.sights = sights  # list of Sight objects
        self.hotel = hotel
        self.meals_summary = meals_summary

# ═══════════════════════════════════════════════════════
# GENERATION ENGINE
# ═══════════════════════════════════════════════════════

def generate(days, title, subtitle, output_path):
    doc = Document()
    _setup_page(doc)
    _add_title_page(doc, title, subtitle)
    for day in days:
        _add_day(doc, day)
    doc.save(output_path)
    print('OK - Saved:', output_path)

def _setup_page(doc):
    s = doc.sections[0]
    s.page_width = Cm(21)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(2.0)
    s.bottom_margin = Cm(1.8)
    s.left_margin = Cm(2.0)
    s.right_margin = Cm(2.0)

def _add_title_page(doc, title, subtitle):
    for _ in range(4):
        doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('*  *  *  *  *')
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(14)
    r.font.color.rgb = WARM_GOLD
    doc.add_paragraph('')
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(28)
    r.font.bold = True
    r.font.color.rgb = DEEP_BLUE
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(14)
    r.font.color.rgb = MEDIUM_GRAY
    doc.add_paragraph('')
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('*  *  *  *  *')
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(14)
    r.font.color.rgb = WARM_GOLD
    doc.add_page_break()

def _add_day(doc, day):
    # Day header
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    c = t.cell(0, 0)
    set_shading(c, BG_LIGHT)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    rb = p.add_run('  D-' + str(day.num).zfill(2) + '  ')
    rb.font.name = FONT
    rb._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    rb.font.size = Pt(12)
    rb.font.bold = True
    rb.font.color.rgb = WHITE
    rb._element.get_or_add_rPr().append(
        parse_xml('<w:shd {} w:fill="1A3C6E" w:val="clear"/>'.format(nsdecls("w"))))
    rsp = p.add_run('  ')
    rsp.font.size = Pt(10)
    rr = p.add_run(day.route)
    rr.font.name = FONT
    rr._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    rr.font.size = Pt(12)
    rr.font.bold = True
    rr.font.color.rgb = DEEP_BLUE

    # Detail table
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, txt in enumerate(['Time', 'Highlights', 'Meals/Notes']):
        mkp(tbl.rows[0].cells[i], txt, sz=9, bold=True, color=WHITE,
            align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=4)
        set_shading(tbl.rows[0].cells[i], DEEP_BLUE_HEX)
    tbl.rows[0].cells[0].width = Cm(2.2)
    tbl.rows[0].cells[1].width = Cm(11.5)
    tbl.rows[0].cells[2].width = Cm(3.3)

    for sight in day.sights:
        _add_sight(tbl, sight)

    if day.meals_summary:
        _add_meal_row(tbl, day.meals_summary)

    _add_hotel_row(tbl, day.hotel)
    doc.add_paragraph('')

def _add_sight(tbl, sight):
    row = tbl.add_row()
    cs = row.cells
    mkp(cs[0], sight.time, sz=9, bold=True, color=DEEP_BLUE,
        align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=4)
    set_shading(cs[0], BG_LIGHT)
    pn = mkp(cs[1], '', sz=10, sb=2, sa=2)
    rn = pn.add_run(sight.name)
    rn.font.name = FONT
    rn._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    rn.font.size = Pt(11)
    rn.font.bold = True
    rn.font.color.rgb = DARK_GREEN
    if sight.description:
        mkp(cs[1], sight.description, sz=9, color=MEDIUM_GRAY, sb=2, sa=4, ls=1.4)
    if sight.meal:
        mkp(cs[2], sight.meal, sz=9, color=MEDIUM_GRAY,
            align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=4)
    if sight.note:
        mkp(cs[2], sight.note, sz=8, color=LIGHT_GRAY,
            align=WD_ALIGN_PARAGRAPH.CENTER, sb=2, sa=2)
    set_shading(cs[2], BG_WARM)
    cs[0].width = Cm(2.2)
    cs[1].width = Cm(11.5)
    cs[2].width = Cm(3.3)

def _add_meal_row(tbl, txt):
    row = tbl.add_row()
    cs = row.cells
    for c in cs:
        set_shading(c, BG_GREEN)
    mkp(cs[0], '[Meals]', sz=9, bold=True, color=DARK_GREEN,
        align=WD_ALIGN_PARAGRAPH.CENTER, sb=3, sa=3)
    mkp(cs[1], txt, sz=9, sb=3, sa=3)
    cs[0].width = Cm(2.2)
    cs[1].width = Cm(11.5)
    cs[2].width = Cm(3.3)

def _add_hotel_row(tbl, txt):
    row = tbl.add_row()
    cs = row.cells
    for c in cs:
        set_shading(c, BG_RED)
    mkp(cs[0], '[Hotel]', sz=9, bold=True, color=RGBColor(0xCC, 0x33, 0x33),
        align=WD_ALIGN_PARAGRAPH.CENTER, sb=3, sa=3)
    mkp(cs[1], txt, sz=9, sb=3, sa=3)
    cs[0].width = Cm(2.2)
    cs[1].width = Cm(11.5)
    cs[2].width = Cm(3.3)

# ═══════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════

if __name__ == '__main__':
    generate(DAYS, TITLE, SUBTITLE, OUTPUT_FILE)
