#!/usr/bin/env python3
"""
Template: Generate travel itinerary DOCX from structured data.
Customize for each itinerary: edit DAYS list and run.

Run with:
    PYTHONIOENCODING=utf-8 python scripts/generate.py

Before re-running, delete the old output file if Word has it open.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os
import re
import zipfile
import shutil

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
BG_GREEN    = 'E8F5ED'
BG_RED      = 'FFF1F0'

# ── HELPERS ─────────────────────────────────────────────

def set_shading(cell, color):
    """Set cell background. Removes existing shd to avoid duplicates."""
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}shd' % NS):
        tcPr.remove(existing)
    tcPr.append(
        parse_xml('<w:shd {} w:fill="{}"/>'.format(nsdecls("w"), color)))

def set_cell_va(cell, align='center'):
    """Set vertical alignment. Pass 'center' as string, NOT the enum."""
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}vAlign' % NS):
        tcPr.remove(existing)
    tcPr.append(
        parse_xml('<w:vAlign {} w:val="{}"/>'.format(nsdecls("w"), align)))

def add_run_to_cell(cell, text, sz=10, bold=None, color=DARK_GRAY,
                    align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=0, ls=1.3):
    """Add a paragraph with a run to a cell. bold=None = no <w:b/> element emitted."""
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = sb
    p.paragraph_format.space_after = sa
    p.paragraph_format.line_spacing = ls
    r = p.add_run(text)
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    r.font.size = Pt(sz)
    if bold is not None:
        r.font.bold = bold
    r.font.color.rgb = color
    return p

def cleanup_docx(filepath):
    """Remove customXml, thumbnail, stylesWithEffects that can break Word."""
    tmp = filepath + '.tmp'
    with zipfile.ZipFile(filepath, 'r') as zin:
        ct = zin.read('[Content_Types].xml').decode('utf-8')
        ct = re.sub(r'<Override[^>]*?customXml[^>]*?/>', '', ct)
        ct = re.sub(r'<Override[^>]*?stylesWithEffects[^>]*?/>', '', ct)
        ct = re.sub(r'<Default[^>]*?jpeg[^>]*?/>', '', ct)

        doc_rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')
        doc_rels = re.sub(r'<Relationship[^>]*?customXml[^>]*?/>', '', doc_rels)
        doc_rels = re.sub(r'<Relationship[^>]*?stylesWithEffects[^>]*?/>', '', doc_rels)

        main_rels = zin.read('_rels/.rels').decode('utf-8')
        main_rels = re.sub(r'<Relationship[^>]*?thumbnail[^>]*?/>', '', main_rels)

        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                fname = item.filename
                if any(x in fname for x in ['customXml', 'stylesWithEffects',
                                              'thumbnail', 'item1.xml']):
                    continue
                if fname == '[Content_Types].xml':
                    zout.writestr(fname, ct.encode('utf-8'))
                elif fname == 'word/_rels/document.xml.rels':
                    zout.writestr(fname, doc_rels.encode('utf-8'))
                elif fname == '_rels/.rels':
                    zout.writestr(fname, main_rels.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(fname))
    shutil.move(tmp, filepath)

# ═══════════════════════════════════════════════════════
# DATA MODEL - populate DAYS list below
# ═══════════════════════════════════════════════════════

class Sight:
    def __init__(self, time, name, description, note=None, grade=None):
        self.time = time
        self.name = name
        self.description = description  # 100-150 chars, factual, for foreign readers
        self.note = note
        self.grade = grade

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
    cleanup_docx(output_path)
    print('OK - Cleaned up package')

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

    # Detail table — 2 columns only: Time | Highlights
    # Meals go ONLY in the [Meals] row at the bottom, never in sight rows.
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, txt in enumerate(['时间', '观光内容']):
        add_run_to_cell(tbl.rows[0].cells[i], txt, sz=9, bold=True, color=WHITE,
                        align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=4)
        set_shading(tbl.rows[0].cells[i], DEEP_BLUE_HEX)
    tbl.rows[0].cells[0].width = Cm(2.5)
    tbl.rows[0].cells[1].width = Cm(14.5)

    for sight in day.sights:
        _add_sight(tbl, sight)

    if day.meals_summary:
        _add_meal_row(tbl, day.meals_summary)

    _add_hotel_row(tbl, day.hotel)

def _add_sight(tbl, sight):
    row = tbl.add_row()
    cs = row.cells
    # Time cell
    add_run_to_cell(cs[0], sight.time, sz=9, bold=True, color=DEEP_BLUE,
                    align=WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=4)
    set_shading(cs[0], BG_LIGHT)

    # Sight cell — clear default paragraph first
    cs[1].paragraphs[0].clear()

    # Sight name with optional grade badge
    name_text = sight.name
    if sight.grade:
        name_text += '  [' + sight.grade + ']'
    pn = add_run_to_cell(cs[1], '', sz=10, sb=2, sa=1)
    rn = pn.add_run(name_text)
    rn.font.name = FONT
    rn._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    rn.font.size = Pt(10)
    rn.font.bold = True
    rn.font.color.rgb = DARK_GREEN

    # Description
    if sight.description:
        add_run_to_cell(cs[1], sight.description, sz=8.5, color=MEDIUM_GRAY, sb=1, sa=4, ls=1.35)

    # Note (seasonal info, closures, self-paid marker) — appended under description
    if sight.note:
        add_run_to_cell(cs[1], sight.note, sz=8, color=LIGHT_GRAY, sb=0, sa=2)

    cs[0].width = Cm(2.5)
    cs[1].width = Cm(14.5)

def _add_meal_row(tbl, txt):
    """Add a [Meals] summary row — 2 columns, merged horizontally for the meal text."""
    row = tbl.add_row()
    cs = row.cells
    # Merge both cells for the full-width meal text
    cs[0].merge(cs[1])
    for c in row.cells:
        set_shading(c, BG_GREEN)
    add_run_to_cell(row.cells[0], '[Meals]  ' + txt, sz=9, color=DARK_GRAY,
                    align=WD_ALIGN_PARAGRAPH.LEFT, sb=3, sa=3)

def _add_hotel_row(tbl, txt):
    """Add a [Hotel] row — 2 columns, merged horizontally for the hotel text."""
    row = tbl.add_row()
    cs = row.cells
    cs[0].merge(cs[1])
    for c in row.cells:
        set_shading(c, BG_RED)
    add_run_to_cell(row.cells[0], '[Hotel]  ' + txt, sz=9, color=DARK_GRAY,
                    align=WD_ALIGN_PARAGRAPH.LEFT, sb=3, sa=3)

# ═══════════════════════════════════════════════════════
# RUN
# ═══════════════════════════════════════════════════════

if __name__ == '__main__':
    generate(DAYS, TITLE, SUBTITLE, OUTPUT_FILE)
